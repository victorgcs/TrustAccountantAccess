VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCurrencyRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCurrencyRates"

'VGC 09/28/2015: CHANGES!

Private Const strCallingForm As String = "frmCurrency_Rate"

Private lngChkBox_Offset As Long, lngNotes_Top As Long, lngDetail_Height As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
Private frmPar As Access.Form, frmSub As Access.Form
Private lngTpp As Long, blnNotes As Boolean
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String
        Dim lngCnt As Long, lngOption_Width As Long
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

150         Set frmPar = Forms(strCallingForm)
160         Set frmSub = frmPar.frmCurrency_Rate_Sub.Form

170         Select Case frmPar.chkPageOf
            Case True
180           .PageNumOf.Visible = True
190           .PageNum.Visible = False
200         Case False
210           .PageNumOf.Visible = False
220           .PageNum.Visible = True
230         End Select

240         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
250         .AsOf_lbl2.Caption = "As of: " & Format(Date, "mmmm dd, yyyy")
260         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

270         strOrderBy = Forms(strCallingForm).frmCurrency_Rate_Sub.Form.SortNow_Get  ' ** Form Function: frmCurrency_Rate_Sub.
            ' ** [curr_code]
            ' ** [curr_name_sort]
            ' ** [curr_active], [curr_name_sort]
            ' ** [curr_rate1], [curr_word1], [curr_name_sort]
            ' ** [curr_rate2], [curr_word1], [curr_name_sort]
            ' ** [curr_date], [curr_name_sort]
            ' ** [curr_iso]
            ' ** [curr_fund], [curr_name_sort]
            ' ** [curr_metal], [curr_name_sort]
            ' ** [curr_bmu], [curr_name_sort]
            ' ** [curr_alt], [curr_name_sort]
            ' ** [curr_unit], [curr_name_sort]
            ' ** [curr_word1], [curr_name_sort]
            ' ** [curr_notes], [curr_name_sort]
280         .OrderBy = strOrderBy
290         .OrderByOn = True

300         lngChkBox_Offset = (.curr_fund.Left - .curr_fund_lbl.Left)
310         .detail_height_line.Top = 0&
320         lngNotes_Top = .curr_notes_display.Top
330         lngDetail_Height = .Detail.Height
340         lngOneLine_Top = .curr_name_lbl.Top
350         lngTwoLine_Top = .curr_code_lbl.Top  ' ** Zero.
360         lngUnderline_Top = .curr_code_lbl_line.Top
370         lngPageHeaderSection_Height = .Section(acPageHeader).Height

380         Select Case frmPar.chkIncludeNotes
            Case True
390           blnNotes = True
400         Case False
410           blnNotes = False
420           .curr_notes_display.Visible = False
430           .curr_notes_display.Top = 0&
440         End Select

450         Select Case frmPar.opgFund
            Case frmPar.opgFund_optInclude.OptionValue, frmPar.opgFund_optOnly.OptionValue
460           .curr_fund.Visible = True
470           .curr_fund_lbl.Visible = True
480           .curr_fund_lbl_line.Visible = True
490         Case frmPar.opgFund_optExclude.OptionValue
500           .curr_fund.Visible = False
510           .curr_fund_lbl.Visible = False
520           .curr_fund_lbl_line.Visible = False
530         End Select
540         Select Case frmPar.opgMetal
            Case frmPar.opgMetal_optInclude.OptionValue, frmPar.opgMetal_optOnly.OptionValue
550           .curr_metal.Visible = True
560           .curr_metal_lbl.Visible = True
570           .curr_metal_lbl_line.Visible = True
580         Case frmPar.opgMetal_optExclude.OptionValue
590           .curr_metal.Visible = False
600           .curr_metal_lbl.Visible = False
610           .curr_metal_lbl_line.Visible = False
620         End Select
630         Select Case frmPar.opgBMU
            Case frmPar.opgBMU_optInclude.OptionValue, frmPar.opgBMU_optOnly.OptionValue
640           .curr_bmu.Visible = True
650           .curr_bmu_lbl.Visible = True
660           .curr_bmu_lbl_line.Visible = True
670         Case frmPar.opgBMU_optExclude.OptionValue
680           .curr_bmu.Visible = False
690           .curr_bmu_lbl.Visible = False
700           .curr_bmu_lbl_line.Visible = False
710         End Select
720         Select Case frmPar.opgAlt
            Case frmPar.opgAlt_optInclude.OptionValue, frmPar.opgAlt_optOnly.OptionValue
730           .curr_alt.Visible = True
740           .curr_alt_lbl.Visible = True
750           .curr_alt_lbl_line.Visible = True
760         Case frmPar.opgAlt_optExclude.OptionValue
770           .curr_alt.Visible = False
780           .curr_alt_lbl.Visible = False
790           .curr_alt_lbl_line.Visible = False
800         End Select
810         Select Case frmPar.opgUnit
            Case frmPar.opgUnit_optInclude.OptionValue, frmPar.opgUnit_optOnly.OptionValue
820           .curr_unit.Visible = True
830           .curr_unit_lbl.Visible = True
840           .curr_unit_lbl_line.Visible = True
850         Case frmPar.opgUnit_optExclude.OptionValue
860           .curr_unit.Visible = False
870           .curr_unit_lbl.Visible = False
880           .curr_unit_lbl_line.Visible = False
890         End Select

900         lngCnt = Abs(.curr_fund.Visible + .curr_metal.Visible + .curr_bmu.Visible + .curr_alt.Visible + .curr_unit.Visible)
910         lngOption_Width = (.curr_metal_lbl.Left - .curr_fund_lbl.Left)  ' ** One option's width.
920         lngTmp01 = lngOption_Width
930         lngTmp02 = (2& * lngOption_Width)
940         lngTmp03 = (3& * lngOption_Width)
950         lngTmp04 = (4& * lngOption_Width)

            ' ** It looks like I'm pushing these all to the right, rather then pulling them left.
960         Select Case lngCnt
            Case 5&
              ' ** No change.
970         Case 4&
              ' ** Only 1 option is missing.
980           If .curr_fund.Visible = False Then
                ' ** No option movement.
                ' ** Leave the hole where Fund was.
990           ElseIf .curr_metal.Visible = False Then
                ' ** Push Fund right 1 slot.
1000            .curr_fund_lbl.Left = (.curr_fund_lbl.Left + lngTmp01)
1010            .curr_fund_lbl_line.Left = .curr_fund_lbl.Left
1020            .curr_fund.Left = (.curr_fund_lbl.Left + lngChkBox_Offset)
1030          ElseIf .curr_bmu.Visible = False Then
                ' ** Push Fund and Metal 1 slot right.
1040            .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp01)
1050            .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
1060            .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
1070            .curr_fund_lbl.Left = (.curr_fund_lbl.Left + lngTmp01)
1080            .curr_fund_lbl_line.Left = .curr_fund_lbl.Left
1090            .curr_fund.Left = (.curr_fund_lbl.Left + lngChkBox_Offset)
1100          ElseIf .curr_alt.Visible = False Then
                ' ** Push Fund, Metal, and BMU 1 slot right.
1110            .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp01)
1120            .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
1130            .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
1140            .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp01)
1150            .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
1160            .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
1170            .curr_fund_lbl.Left = (.curr_fund_lbl.Left + lngTmp01)
1180            .curr_fund_lbl_line.Left = .curr_fund_lbl.Left
1190            .curr_fund.Left = (.curr_fund_lbl.Left + lngChkBox_Offset)
1200          ElseIf .curr_unit.Visible = False Then
                ' ** Push Fund, Metal, BMU, and Misc to the right.
1210            .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
1220            .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
1230            .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
1240            .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp01)
1250            .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
1260            .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
1270            .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp01)
1280            .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
1290            .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
1300            .curr_fund_lbl.Left = (.curr_fund_lbl.Left + lngTmp01)
1310            .curr_fund_lbl_line.Left = .curr_fund_lbl.Left
1320            .curr_fund.Left = (.curr_fund_lbl.Left + lngChkBox_Offset)
1330          End If
1340        Case 3&
              ' ** 2 options are missing.
1350          If .curr_fund.Visible = True Then
                ' ** Move 2 right. (2 missing among Metal, BMU, Misc, and Unit.)
1360            .curr_fund_lbl.Left = (.curr_fund_lbl.Left + lngTmp02)
1370            .curr_fund_lbl_line.Left = .curr_fund_lbl.Left
1380            .curr_fund.Left = (.curr_fund_lbl.Left + lngChkBox_Offset)
                ' ** Fund now in slot 3.
1390            If .curr_metal.Visible = True Then
                  ' ** Move 2 right. (2 missing among BMU, Misc, and Unit.)
1400              .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp02)
1410              .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
1420              .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
                  ' ** Metal now in slot 4.
1430              If .curr_bmu.Visible = True Then
                    ' ** Move 2 right. (2 missing are Misc and Unit.)
1440                .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp02)
1450                .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
1460                .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
                    ' ** BMU now in slot 5.
1470              Else
1480                If .curr_alt.Visible = True Then
                      ' ** Move 1 right. (2 missing are BMU and Unit.)
1490                  .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
1500                  .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
1510                  .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
                      ' ** Misc now in slot 5.
1520                ElseIf .curr_unit.Visible = True Then
                      ' ** No movement. (2 missing are BMU and Misc.)
                      ' ** Unit remains in slot 5.
1530                End If
1540              End If
1550            ElseIf .curr_bmu.Visible = True Then
                  ' ** Move 1 right. (2 missing are Metal, and either Misc or Unit.)
1560              .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp01)
1570              .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
1580              .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
                  ' ** BMU now in slot 4.
1590              If .curr_alt.Visible = True Then
                    ' ** Move 1 right. (2 missing are Metal and Unit.)
1600                .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
1610                .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
1620                .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
                    ' ** Misc now in slot 5.
1630              ElseIf .curr_unit.Visible = True Then
                    ' ** No movement. (2 missing are Metal and Misc.)
                    ' ** Unit remains in slot 5.
1640              End If
1650            ElseIf .curr_alt.Visible = True Then
                  ' ** No movement. (2 missing are Metal and BMU.)
                  ' ** Misc remains in slot 4.
                  ' ** Unit remains in slot 5.
1660            End If
1670          ElseIf .curr_metal.Visible = True Then
                ' ** Move 1 right. (2 missing are Fund, and among BMU, Misc, and Unit.)
1680            .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp01)
1690            .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
1700            .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
                ' ** Metal now in slot 3.
1710            If .curr_bmu.Visible = True Then
                  ' ** Move 1 right. (2 missing are Fund, and either Misc or Unit.)
1720              .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp01)
1730              .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
1740              .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
                  ' ** BMU now in slot 4.
1750              If .curr_alt.Visible = True Then
                    ' ** Move 1 right. (2 missing are Fund and Unit.)
1760                .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
1770                .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
1780                .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
                    ' ** Misc now in slot 5.
1790              ElseIf .curr_unit.Visible = True Then
                    ' ** No movement. (2 missing are Fund and Misc.)
                    ' ** Unit remains in slot 5.
1800              End If
1810            ElseIf .curr_alt.Visible = True Then
                  ' ** No movement. (2 missing are Fund and BMU.)
                  ' ** Misc remains in slot 4.
                  ' ** Unit remains in slot 5.
1820            End If
1830          ElseIf .curr_bmu.Visible = True Then
                ' ** No movement. (2 missing are Fund and Metal.)
                ' ** BMU remains in slot 3.
                ' ** Misc remains in slot 4.
                ' ** Unit remains in slot 5.
1840          End If
1850        Case 2&
              ' ** 3 options are missing.
1860          If .curr_fund.Visible = True Then
                ' ** Move 3 right. (3 missing are among Metal, BMU, Misc, and Unit.)
1870            .curr_fund_lbl.Left = (.curr_fund_lbl.Left + lngTmp03)
1880            .curr_fund_lbl_line.Left = .curr_fund_lbl.Left
1890            .curr_fund.Left = (.curr_fund_lbl.Left + lngChkBox_Offset)
                ' ** Fund now in slot 4.
1900            If .curr_metal.Visible = True Then
                  ' ** Move 3 right. (3 missing are BMU, Misc, and Unit.)
1910              .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp03)
1920              .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
1930              .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
                  ' ** Metal now in slot 5.
1940            ElseIf .curr_bmu.Visible = True Then
                  ' ** Move 2 right. (3 missing are Metal, Misc, and Unit.)
1950              .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp02)
1960              .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
1970              .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
                  ' ** BMU now in slot 5.
1980            ElseIf .curr_alt.Visible = True Then
                  ' ** Move 1 right. (3 missing are Metal, BMU, and Unit.)
1990              .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
2000              .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
2010              .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
                  ' ** Misc now in slot 5.
2020            ElseIf .curr_unit.Visible = True Then
                  ' ** No movement. (3 missing are Metal, BMU, and Misc.)
                  ' ** Unit remains in slot 5.
2030            End If
2040          ElseIf .curr_metal.Visible = True Then
                ' ** Move 2 right. (3 missing are Fund, and among BMU, Misc, and Unit.)
2050            .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp02)
2060            .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
2070            .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
                ' ** Metal now in slot 4.
2080            If .curr_bmu.Visible = True Then
                  ' ** Move 2 right. (3 missing are Fund, Misc, and Unit.)
2090              .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp02)
2100              .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
2110              .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
                  ' ** BMU now in slot 5.
2120            ElseIf .curr_alt.Visible = True Then
                  ' ** Move 1 right. (3 missing are Fund, BMU, and Unit.)
2130              .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
2140              .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
2150              .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
                  ' ** Misc now in slot 5.
2160            ElseIf .curr_unit.Visible = True Then
                  ' ** No movement. (3 missing are Fund, BMU, Misc.)
                  ' ** Unit remains in slot 5.
2170            End If
2180          ElseIf .curr_bmu.Visible = True Then
                ' ** Move 1 right. (3 missing are Fund, Metal, and among Misc and Unit.)
2190            .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp01)
2200            .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
2210            .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
                ' ** BMU now in slot 4.
2220            If .curr_alt.Visible = True Then
                  ' ** Move 1 right. (3 missing are Fund, Metal, and Unit.)
2230              .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
2240              .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
2250              .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
                  ' ** Misc now in slot 5.
2260            ElseIf .curr_unit.Visible = True Then
                  ' ** No movement. (3 missing are Fund, Metal, and Misc.)
                  ' ** Unit remains in slot 5.
2270            End If
2280          ElseIf .curr_alt.Visible = True Then
                ' ** No movement. (3 missing are Fund, Metal, and BMU.)
                ' ** Misc remains in slot 4.
                ' ** Unit remains in slot 5.
2290          End If
2300        Case 1&
              ' ** 4 options are missing.
2310          If .curr_fund.Visible = True Then
                ' ** Move 4 right. (4 missing are Metal, BMU, Misc, and Unit.)
2320            .curr_fund_lbl.Left = (.curr_fund_lbl.Left + lngTmp04)
2330            .curr_fund_lbl_line.Left = .curr_fund_lbl.Left
2340            .curr_fund.Left = (.curr_fund_lbl.Left + lngChkBox_Offset)
                ' ** Fund now in slot 5.
2350          ElseIf .curr_metal.Visible = True Then
                ' ** Move 3 right. (4 missing are Fund, BMU, Misc, and Unit.)
2360            .curr_metal_lbl.Left = (.curr_metal_lbl.Left + lngTmp03)
2370            .curr_metal_lbl_line.Left = .curr_metal_lbl.Left
2380            .curr_metal.Left = (.curr_metal_lbl.Left + lngChkBox_Offset)
                ' ** Metal now in slot 5.
2390          ElseIf .curr_bmu.Visible = True Then
                ' ** Move 2 right. (4 missing are Fund, Metal, Misc, and Unit.)
2400            .curr_bmu_lbl.Left = (.curr_bmu_lbl.Left + lngTmp02)
2410            .curr_bmu_lbl_line.Left = .curr_bmu_lbl.Left
2420            .curr_bmu.Left = (.curr_bmu_lbl.Left + lngChkBox_Offset)
                ' ** BMU now in slot 5.
2430          ElseIf .curr_alt.Visible = True Then
                ' ** Move 1 right. (4 missing are Fund, Metal, BMU, and Unit.)
2440            .curr_alt_lbl.Left = (.curr_alt_lbl.Left + lngTmp01)
2450            .curr_alt_lbl_line.Left = .curr_alt_lbl.Left
2460            .curr_alt.Left = (.curr_alt_lbl.Left + lngChkBox_Offset)
                ' ** Misc now in slot 5.
2470          ElseIf .curr_unit.Visible = True Then
                ' ** No movement. (4 missing are Fund, Metal, BMU, and Misc.)
                ' ** Unit remains in slot 5.
2480          End If
2490        Case 0&
              ' ** All 5 options missing.
              ' ** No movement.
2500        End Select

            ' ** All extra space going into curr_name and curr_notes.
2510        Select Case lngCnt
            Case 5&
2520          lngTmp01 = 0&
2530        Case 4&
              'lngTmp01 = lngTmp01
2540        Case 3&
2550          lngTmp01 = lngTmp02
2560        Case 2&
2570          lngTmp01 = lngTmp03
2580        Case 1&
2590          lngTmp01 = lngTmp04
2600        Case 0&
2610          lngTmp01 = (lngTmp01 + lngTmp04)
2620        End Select

2630        .curr_date.Left = (.curr_date.Left + lngTmp01)
2640        .curr_date_lbl.Left = .curr_date.Left
2650        .curr_date_lbl_line.Left = .curr_date.Left
2660        .curr_rate2.Left = (.curr_rate2.Left + lngTmp01)
2670        .curr_rate2_lbl.Left = .curr_rate2.Left
2680        .curr_rate2_lbl_line.Left = .curr_rate2.Left
2690        .curr_rate1.Left = (.curr_rate1.Left + lngTmp01)
2700        .curr_rate1_lbl.Left = .curr_rate1.Left
2710        .curr_rate1_lbl_line.Left = .curr_rate1.Left
2720        .curr_active.Left = (.curr_active.Left + lngTmp01)
2730        .curr_active_lbl.Left = (.curr_active_lbl.Left + lngTmp01)
2740        .curr_active_lbl_line.Left = .curr_active_lbl.Left
2750        .curr_name.Width = (.curr_name.Width + lngTmp01)
2760        .curr_name_lbl.Width = (.curr_name_lbl.Width + lngTmp01)
2770        .curr_name_lbl_line.Width = (.curr_name_lbl_line.Width + lngTmp01)
2780        .curr_notes_display.Width = .curr_name.Width

2790        Select Case frmPar.opgActive
            Case frmPar.opgActive_optInclude.OptionValue, frmPar.opgActive_optOnly.OptionValue
              ' ** No change.
2800        Case frmPar.opgActive_optExclude.OptionValue
2810          .curr_active.Visible = False
2820          .curr_active_lbl.Visible = False
2830          .curr_active_lbl_line.Visible = False
2840          lngTmp01 = ((.curr_active_lbl.Left + .curr_active_lbl.Width) - (.curr_name_lbl.Left + .curr_name_lbl.Width))
2850          .curr_name.Width = (.curr_name.Width + lngTmp01)
2860          .curr_name_lbl.Width = .curr_name.Width
2870          .curr_name_lbl_line.Width = .curr_name_lbl.Width
2880          .curr_notes_display.Width = .curr_name.Width
2890        End Select

2900        .Caption = "Currency_Rates_" & Format$(Date, "yymmdd")

2910      Else
2920        Cancel = -1
2930      End If
2940    End With

2950    If Cancel = -1 Then
2960      Beep
2970      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
2980    End If

EXITP:
2990    Exit Sub

ERRH:
3000    Forms(strCallingForm).Visible = True
3010    DoCmd.Restore
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long

3110    With Me
3120      lngTmp01 = .AsOf_lbl2.Height
3130      If .Page >= 2 Then
3140        .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
3150        lngTmp02 = (lngOneLine_Top + lngTmp01)
3160        lngTmp03 = (lngUnderline_Top + lngTmp01)
3170        .curr_code_lbl.Top = lngTmp01
3180        .curr_code_lbl_line.Top = lngTmp03
3190        .curr_iso_lbl.Top = lngTmp02
3200        .curr_iso_lbl_line.Top = lngTmp03
3210        .curr_name_lbl.Top = lngTmp02
3220        .curr_name_lbl_line.Top = lngTmp03
3230        .curr_active_lbl.Top = lngTmp02
3240        .curr_active_lbl_line.Top = lngTmp03
3250        .curr_rate1_lbl.Top = lngTmp01
3260        .curr_rate1_lbl_line.Top = lngTmp03
3270        .curr_rate2_lbl.Top = lngTmp01
3280        .curr_rate2_lbl_line.Top = lngTmp03
3290        .curr_date_lbl.Top = lngTmp01
3300        .curr_date_lbl_line.Top = lngTmp03
3310        .curr_fund_lbl.Top = lngTmp02
3320        .curr_fund_lbl_line.Top = lngTmp03
3330        .curr_metal_lbl.Top = lngTmp02
3340        .curr_metal_lbl_line.Top = lngTmp03
3350        .curr_bmu_lbl.Top = lngTmp02
3360        .curr_bmu_lbl_line.Top = lngTmp03
3370        .curr_alt_lbl.Top = lngTmp02
3380        .curr_alt_lbl_line.Top = lngTmp03
3390        .curr_unit_lbl.Top = lngTmp02
3400        .curr_unit_lbl_line.Top = lngTmp03
3410        .Header_lbl2.Visible = True
3420        .AsOf_lbl2.Visible = True
3430        .Header_hline05.Visible = True
3440      Else
3450        .Header_lbl2.Visible = False
3460        .AsOf_lbl2.Visible = False
3470        .Header_hline05.Visible = False
3480        .curr_code_lbl.Top = lngTwoLine_Top
3490        .curr_code_lbl_line.Top = lngUnderline_Top
3500        .curr_iso_lbl.Top = lngOneLine_Top
3510        .curr_iso_lbl_line.Top = lngUnderline_Top
3520        .curr_name_lbl.Top = lngOneLine_Top
3530        .curr_name_lbl_line.Top = lngUnderline_Top
3540        .curr_active_lbl.Top = lngOneLine_Top
3550        .curr_active_lbl_line.Top = lngUnderline_Top
3560        .curr_rate1_lbl.Top = lngTwoLine_Top
3570        .curr_rate1_lbl_line.Top = lngUnderline_Top
3580        .curr_rate2_lbl.Top = lngTwoLine_Top
3590        .curr_rate2_lbl_line.Top = lngUnderline_Top
3600        .curr_date_lbl.Top = lngTwoLine_Top
3610        .curr_date_lbl_line.Top = lngUnderline_Top
3620        .curr_fund_lbl.Top = lngOneLine_Top
3630        .curr_fund_lbl_line.Top = lngUnderline_Top
3640        .curr_metal_lbl.Top = lngOneLine_Top
3650        .curr_metal_lbl_line.Top = lngUnderline_Top
3660        .curr_bmu_lbl.Top = lngOneLine_Top
3670        .curr_bmu_lbl_line.Top = lngUnderline_Top
3680        .curr_alt_lbl.Top = lngOneLine_Top
3690        .curr_alt_lbl_line.Top = lngUnderline_Top
3700        .curr_unit_lbl.Top = lngOneLine_Top
3710        .curr_unit_lbl_line.Top = lngUnderline_Top
3720        .Section(acPageHeader).Height = lngPageHeaderSection_Height
3730      End If
3740    End With

EXITP:
3750    Exit Sub

ERRH:
3760    Select Case ERR.Number
        Case Else
3770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3780    End Select
3790    Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

3810    With Me
3820      If IsNull(.curr_notes) = True Or blnNotes = False Then
3830        .curr_notes_display.Top = 0&
3840        .Detail.Height = .detail_height_line.Height
3850      ElseIf blnNotes = True Then
3860        .Detail.Height = lngDetail_Height
3870        .curr_notes_display.Top = lngNotes_Top
3880      End If
3890    End With

EXITP:
3900    Exit Sub

ERRH:
3910    Forms(strCallingForm).Visible = True
3920    DoCmd.Restore
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub Report_Close()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

4010    DoCmd.Restore

4020    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

4030    If Forms(strCallingForm).Visible = False Then
4040      Forms(strCallingForm).Visible = True
4050      DoCmd.SelectObject acForm, strCallingForm, False
4060      DoCmd.Restore
4070    End If

EXITP:
4080    Set frmSub = Nothing
4090    Set frmPar = Nothing
4100    Exit Sub

ERRH:
4110    Forms(strCallingForm).Visible = True
4120    DoCmd.Restore
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub
