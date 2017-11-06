VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetHistory"

'VGC 03/29/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_AssetHistory"

Private blnAll As Boolean
Private lngSF_Left As Long, lngSF_Width As Long
Private lngIC_Left As Long, lngIC_Width As Long
Private lngPC_Left As Long, lngPC_Width As Long
Private lngCO_Left As Long, lngCO_Width As Long

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strFilter As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Select Case Forms(strCallingForm).chkIncludeHiddenTrans
            Case True
150           strFilter = vbNullString
160           .Filter = strFilter
170           .FilterOn = False
180         Case False
190           strFilter = "[ledger_HIDDEN] = False"
200           .Filter = strFilter
210           .FilterOn = True
220         End Select

230         Select Case Forms(strCallingForm).chkPageOf
            Case True
240           .PageNumOf.Visible = True
250           .PageNum.Visible = False
260         Case False
270           .PageNum.Visible = True
280           .PageNumOf.Visible = False
290           .PageNumOf.ControlSource = vbNullString
300         End Select

310         lngPageHeaderSection_Height = .Section(acPageHeader).Height
320         lngOneLine_Top = .transdate_lbl.Top
330         lngTwoLine_Top = .SumICash_lbl.Top
340         lngUnderline_Top = .transdate_lbl_line.Top

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
350         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

360         .CoInfoBlock.Caption = gstrCo_InfoBlock
370         .AsOf.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
380         .AsOf2.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

390         Select Case Forms(strCallingForm).chkIncludeArchive
            Case True
400           .RptSubtype_lbl.Visible = True
410         Case False
420           .RptSubtype_lbl.Visible = False
430         End Select

440         Select Case Forms(strCallingForm).cmbAssets.Column(3)
            Case "90"  ' ** Liability.
450           .Sumshareface_txt_Liability.Visible = True
460           .Sumshareface_txt.Visible = False
470           .Sumshareface_txt_wparens.Visible = False
480         Case Else
490           .Sumshareface_txt.Visible = False
500           .Sumshareface_txt_wparens.Visible = True
510           .Sumshareface_txt_Liability.Visible = False
520         End Select

530         Select Case Forms(strCallingForm).opgAccountNumber
            Case Forms(strCallingForm).opgAccountNumber_optSpecified.OptionValue
              ' ** qryAssetHistory_11_cusip.
540           blnAll = False
550           .Footer_hline01.Visible = False
560         Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue  ' ** NO LONGER NEEDED!
570           blnAll = True
580           .RecordSource = "qryAssetHistory_20_cusip"
590         End Select

600         Select Case Forms(strCallingForm).chkShowPurchaseDate
            Case True
610           SetPurchaseDate True  ' ** Procedure: Below.
620         Case False
630           SetPurchaseDate False  ' ** Procedure: Below.
640         End Select

650         lngSF_Left = .PreviousShareface.Left
660         lngSF_Width = .PreviousShareface.Width
670         lngIC_Left = .PreviousIcash.Left
680         lngIC_Width = .PreviousIcash.Width
690         lngPC_Left = .PreviousPcash.Left
700         lngPC_Width = .PreviousPcash.Width
710         lngCO_Left = .PreviousCost.Left
720         lngCO_Width = .PreviousCost.Width

730       Else
740         Cancel = -1
750       End If
760     End With

770     If Cancel = -1 Then
780       Beep
790       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
800     End If

EXITP:
810     Exit Sub

ERRH:
820     Select Case ERR.Number
        Case Else
830       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
840     End Select
850     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim strCusip As String
        Dim lngTmp01 As Long

910     With Me
920   On Error Resume Next
          ' ** 2427  You entered an expression that has no value.
930       strCusip = .cusip
940       .Caption = "Asset_History_" & strCusip & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")
950       lngTmp01 = .AsOf2.Height
960       lngTmp01 = (lngTmp01 * 2&)  ' ** AsOf2 and TotDesc2.
970       If .Page >= 2 Then
980         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
990         .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
1000        .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1010        .journaltype_lbl.Top = (lngOneLine_Top + lngTmp01)
1020        .journaltype_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1030        .assetdate_lbl.Top = (lngTwoLine_Top + lngTmp01)
1040        .assetdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1050        .PurchaseDate_long_lbl.Top = (lngTwoLine_Top + lngTmp01)
1060        .PurchaseDate_long_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1070        .Sumshareface_txt_lbl.Top = (lngOneLine_Top + lngTmp01)
1080        .Sumshareface_txt_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1090        .SumICash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1100        .SumICash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1110        .SumPCash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1120        .SumPCash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1130        .SumCost_lbl.Top = (lngOneLine_Top + lngTmp01)
1140        .SumCost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1150        .ledger_HIDDEN_lbl.Top = (lngOneLine_Top + lngTmp01)
1160        .ledger_HIDDEN_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1170        .Header_lbl2.Visible = True
1180        .AsOf2.Visible = True
1190        .TotDesc2.Top = .AsOf2.Height
1200        .TotDesc2.Visible = True
1210        .TotDesc2_lbl.Top = .AsOf2.Height
1220        .TotDesc2_lbl.Visible = True
1230        .Header_hline05.Top = lngTmp01
1240        .Header_hline05.Visible = True
1250      Else
1260        .Header_lbl2.Visible = False
1270        .AsOf2.Visible = False
1280        .TotDesc2.Visible = False
1290        .TotDesc2.Top = 0&
1300        .TotDesc2_lbl.Visible = False
1310        .TotDesc2_lbl.Top = 0
1320        .Header_hline05.Visible = False
1330        .Header_hline05.Top = .AsOf2.Height
1340        .transdate_lbl.Top = lngOneLine_Top
1350        .transdate_lbl_line.Top = lngUnderline_Top
1360        .journaltype_lbl.Top = lngOneLine_Top
1370        .journaltype_lbl_line.Top = lngUnderline_Top
1380        .assetdate_lbl.Top = lngTwoLine_Top
1390        .assetdate_lbl_line.Top = lngUnderline_Top
1400        .PurchaseDate_long_lbl.Top = lngTwoLine_Top
1410        .PurchaseDate_long_lbl_line.Top = lngUnderline_Top
1420        .Sumshareface_txt_lbl.Top = lngOneLine_Top
1430        .Sumshareface_txt_lbl_line.Top = lngUnderline_Top
1440        .SumICash_lbl.Top = lngTwoLine_Top
1450        .SumICash_lbl_line.Top = lngUnderline_Top
1460        .SumPCash_lbl.Top = lngTwoLine_Top
1470        .SumPCash_lbl_line.Top = lngUnderline_Top
1480        .SumCost_lbl.Top = lngOneLine_Top
1490        .SumCost_lbl_line.Top = lngUnderline_Top
1500        .ledger_HIDDEN_lbl.Top = lngOneLine_Top
1510        .ledger_HIDDEN_lbl_line.Top = lngUnderline_Top
1520        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1530      End If
1540  On Error GoTo ERRH
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim dblTmp01 As Double, dblTmp02 As Double, dblTmp03 As Double, dblTmp04 As Double

1710    With Me
1720      If blnAll = True Then
1730        Forms(strCallingForm).GetBal_Beg .accountno  ' ** Form Procedure: frmRpt_AssetHistory.
1740        Forms(strCallingForm).GetBal_End .accountno  ' ** Form Procedure: frmRpt_AssetHistory.
1750      End If
1760      dblTmp01 = Round(Nz(Forms(strCallingForm).PreviousShareface, 0#), 4)
1770      .PreviousShareface.Caption = Format(dblTmp01, "#,##0.0###;(#,##0.0###);0;0")
1780      If dblTmp01 < 0 Then .PreviousShareface.Left = .PreviousShareface.Left + (6& * lngTpp)
1790      dblTmp02 = Nz(Forms(strCallingForm).PreviousIcash, 0#)
1800      dblTmp03 = Nz(Forms(strCallingForm).PreviousPcash, 0#)
1810      dblTmp04 = Nz(Forms(strCallingForm).PreviousCost, 0#)
1820      .PreviousIcash.Caption = Format(dblTmp02, "Currency")
1830      If dblTmp02 >= 0 Then .PreviousIcash.Width = .PreviousIcash.Width - (6& * lngTpp)
1840      .PreviousPcash.Caption = Format(dblTmp03, "Currency")
1850      If dblTmp03 >= 0 Then .PreviousPcash.Width = .PreviousPcash.Width - (6& * lngTpp)
1860      .PreviousCost.Caption = Format(dblTmp04, "Currency")
1870      If dblTmp04 >= 0 Then .PreviousCost.Width = .PreviousCost.Width - (6& * lngTpp)
1880      .Sumshareface_tot.Left = lngSF_Left
1890      .Sumshareface_tot.Width = lngSF_Width
1900      .EndingShareface.Left = lngSF_Left
1910      .EndingShareface.Width = lngSF_Width
1920      .EndingICash.Left = lngIC_Left - (6& * lngTpp)
1930      .EndingICash.Width = lngIC_Width
1940      .EndingPCash.Left = lngPC_Left - (6& * lngTpp)
1950      .EndingPCash.Width = lngPC_Width
1960      .EndingCost.Left = lngCO_Left + (3& * lngTpp)
1970      .EndingCost.Width = lngCO_Width
1980    End With

EXITP:
1990    Exit Sub

ERRH:
2000    Select Case ERR.Number
        Case Else
2010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2020    End Select
2030    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

        Dim dblTmp01 As Double, dblTmp02 As Double, dblTmp03 As Double, dblTmp04 As Double

2110    With Me

2120  On Error Resume Next
          ' ** 2427  You entered an expression that has no value.
2130      dblTmp01 = Round(Nz(Forms(strCallingForm).EndingShareface, 0#), 4)
2140      .EndingShareface.Caption = Format(dblTmp01, "#,##0.0###;(#,##0.0###);0;0")
2150      If dblTmp01 < 0 Then .EndingShareface.Left = .EndingShareface.Left + (6& * lngTpp)
2160      dblTmp02 = Nz(Forms(strCallingForm).EndingICash, 0#)
2170      dblTmp03 = Nz(Forms(strCallingForm).EndingPCash, 0#)
2180      dblTmp04 = Nz(Forms(strCallingForm).EndingCost, 0#)
2190      .EndingICash.Caption = Format(Forms(strCallingForm).EndingICash, "Currency")
2200      If dblTmp02 >= 0 Then .EndingICash.Width = .EndingICash.Width - 90&
2210      .EndingPCash.Caption = Format(Forms(strCallingForm).EndingPCash, "Currency")
2220      If dblTmp03 >= 0 Then .EndingPCash.Width = .EndingPCash.Width - 90&
2230      .EndingCost.Caption = Format(Forms(strCallingForm).EndingCost, "Currency")
2240      If dblTmp04 >= 0 Then .EndingCost.Width = .EndingCost.Width - 90&

2250      dblTmp01 = .Sumshareface_runsum
2260      If dblTmp01 < 0 Then .Sumshareface_tot.Left = .Sumshareface_tot.Left + (4& * lngTpp)
2270      .PreviousShareface.Left = lngSF_Left
2280      .PreviousShareface.Width = lngSF_Width
2290      .PreviousIcash.Left = lngIC_Left
2300      .PreviousIcash.Width = lngIC_Width
2310      .PreviousPcash.Left = lngPC_Left
2320      .PreviousPcash.Width = lngPC_Width
2330      .PreviousCost.Left = lngCO_Left
2340      .PreviousCost.Width = lngCO_Width

2350      .EndingICash.Left = .EndingICash.Left + (6& * lngTpp)
2360      .EndingPCash.Left = .EndingPCash.Left + (6& * lngTpp)
2370      .EndingCost.Left = .EndingCost.Left - (3& * lngTpp)

2380  On Error GoTo ERRH

2390    End With

EXITP:
2400    Exit Sub

ERRH:
2410    Select Case ERR.Number
        Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub Report_Close()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2510    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub SetPurchaseDate(blnShow As Boolean)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "SetPurchaseDate"

2610    With Me
2620      Select Case blnShow
          Case True

2630        .SumCost.Width = .SumCost.Width - (13& * 15&)
2640        .SumCost_lbl.Width = .SumCost.Width
2650        .SumCost_lbl_line.Width = .SumCost.Width
2660        .SumCost_tot.Width = .SumCost.Width
2670        .SumCost_tot_line1.Width = .SumCost_tot_line1.Width - (13& * 15&)
2680        .PreviousCost.Width = .SumCost.Width
2690        .EndingCost.Width = .SumCost.Width

2700        .SumCost.Left = .SumCost.Left + (13& * 15&)
2710        .SumCost_lbl.Left = .SumCost.Left
2720        .SumCost_lbl_line.Left = .SumCost.Left
2730        .SumCost_tot.Left = .SumCost.Left
2740        .SumCost_tot_line1.Left = .SumCost.Left + 120&
2750        .PreviousCost.Left = .SumCost.Left
2760        .EndingCost.Left = .SumCost.Left

2770        .SumPcash.Width = .SumPcash.Width - (13& * 15&)
2780        .SumPCash_lbl.Width = .SumPcash.Width
2790        .SumPCash_lbl_line.Width = .SumPcash.Width
2800        .SumPCash_tot.Width = .SumPcash.Width
2810        .SumPCash_tot_line1.Width = .SumPCash_tot_line1.Width - (13& * 15&)
2820        .PreviousPcash.Width = .SumPcash.Width
2830        .EndingPCash.Width = .SumPcash.Width

2840        .SumPcash.Left = .SumPcash.Left + (2& * (13& * 15&))
2850        .SumPCash_lbl.Left = .SumPcash.Left
2860        .SumPCash_lbl_line.Left = .SumPcash.Left
2870        .SumPCash_tot.Left = .SumPcash.Left
2880        .SumPCash_tot_line1.Left = .SumPcash.Left + 120&
2890        .PreviousPcash.Left = .SumPcash.Left
2900        .EndingPCash.Left = .SumPcash.Left

2910        .SumIcash.Width = .SumIcash.Width - (13& * 15&)
2920        .SumICash_lbl.Width = .SumIcash.Width
2930        .SumICash_lbl_line.Width = .SumIcash.Width
2940        .SumICash_tot.Width = .SumIcash.Width
2950        .SumICash_tot_line1.Width = .SumICash_tot_line1.Width - (13& * 15&)
2960        .PreviousIcash.Width = .SumIcash.Width
2970        .EndingICash.Width = .SumIcash.Width

2980        .SumIcash.Left = .SumIcash.Left + (3& * (13& * 15&))
2990        .SumICash_lbl.Left = .SumIcash.Left
3000        .SumICash_lbl_line.Left = .SumIcash.Left
3010        .SumICash_tot.Left = .SumIcash.Left
3020        .SumICash_tot_line1.Left = .SumIcash.Left + 120&
3030        .PreviousIcash.Left = .SumIcash.Left
3040        .EndingICash.Left = .SumIcash.Left

3050        .Sumshareface_txt.Width = .Sumshareface_txt.Width - (13& * 15&)
3060        .Sumshareface_txt_wparens.Width = .Sumshareface_txt_wparens.Width - (13& * 15&)
3070        .Sumshareface_txt_lbl.Width = .Sumshareface_txt.Width
3080        .Sumshareface_txt_lbl_line.Width = .Sumshareface_txt.Width
3090        .Sumshareface_txt_Liability.Width = .Sumshareface_txt.Width
3100        .Sumshareface_tot.Width = .Sumshareface_txt.Width
3110        .Sumshareface_tot_line1.Width = .Sumshareface_tot_line1.Width - (13& * 15&)
3120        .PreviousShareface.Width = .Sumshareface_txt.Width
3130        .EndingShareface.Width = .Sumshareface_txt.Width

3140        .Sumshareface_txt.Left = .Sumshareface_txt.Left + (4& * (13& * 15&))
3150        .Sumshareface_txt_wparens.Left = .Sumshareface_txt_wparens.Left + (4& * (13& * 15&))
3160        .Sumshareface_txt_lbl.Left = .Sumshareface_txt.Left
3170        .Sumshareface_txt_lbl_line.Left = .Sumshareface_txt.Left
3180        .Sumshareface_txt_Liability.Left = .Sumshareface_txt.Left
3190        .Sumshareface_tot.Left = .Sumshareface_txt.Left
3200        .Sumshareface_tot_line1.Left = ((.Sumshareface_tot.Left + .Sumshareface_tot.Width) - .Sumshareface_tot_line1.Width) + 30&
3210        .PreviousShareface.Left = .Sumshareface_txt.Left
3220        .EndingShareface.Left = .Sumshareface_txt.Left

3230        .BeginningBalance_lbl.Left = .BeginningBalance_lbl.Left + (4& * (13& * 15&))
3240        .PeriodActivity_lbl.Left = .BeginningBalance_lbl.Left
3250        .EndingBalance_lbl.Left = .BeginningBalance_lbl.Left

3260        .assetdate_dp.Visible = True
3270        .assetdate_ws.Visible = True
3280        .assetdate.Visible = False
3290        .assetdate_lbl.Width = .assetdate_dp.Width
3300        .assetdate_lbl_line.Width = .assetdate_dp.Width
3310        .assetdate_lbl.Caption = vbCrLf & "Trade Date"

3320        .PurchaseDate_long.Left = (.assetdate_dp.Left + .assetdate_dp.Width) + 60&
3330        .PurchaseDate_long.Visible = True
3340        .PurchaseDate_long_lbl.Left = .PurchaseDate_long.Left
3350        .PurchaseDate_long_lbl.Visible = True
3360        .PurchaseDate_long_lbl_line.Left = .PurchaseDate_long.Left
3370        .PurchaseDate_long_lbl_line.Visible = True

3380      Case False
3390        .PurchaseDate_long.Visible = False
3400        .PurchaseDate_long_lbl.Visible = False
3410        .PurchaseDate_long_lbl_line.Visible = False

3420        .assetdate.Left = .assetdate.Left + 720&
3430        .assetdate_lbl.Left = .assetdate.Left
3440        .assetdate_lbl_line.Left = .assetdate.Left

3450        .journaltype.Left = .journaltype.Left + 480&
3460        .journaltype_lbl.Left = .journaltype.Left
3470        .journaltype_lbl_line.Left = .journaltype.Left

3480        .transdate.Left = .transdate.Left + 240&
3490        .transdate_lbl.Left = .transdate.Left
3500        .transdate_lbl_line.Left = .transdate.Left

3510        .BeginningBalance_lbl.Left = .BeginningBalance_lbl.Left - 1200&
3520        .PeriodActivity_lbl.Left = .BeginningBalance_lbl.Left
3530        .EndingBalance_lbl.Left = .BeginningBalance_lbl.Left

3540        .Sumshareface_txt.Left = .Sumshareface_txt.Left - 1200&
3550        .Sumshareface_txt_wparens.Left = .Sumshareface_txt_wparens.Left - 1200&
3560        .Sumshareface_txt_lbl.Left = .Sumshareface_txt.Left
3570        .Sumshareface_txt_lbl_line.Left = .Sumshareface_txt.Left
3580        .Sumshareface_txt_Liability.Left = .Sumshareface_txt.Left
3590        .Sumshareface_tot.Left = .Sumshareface_txt.Left
3600        .Sumshareface_tot_line1.Left = ((.Sumshareface_tot.Left + .Sumshareface_tot.Width) - .Sumshareface_tot_line1.Width) + 30&
3610        .PreviousShareface.Left = .Sumshareface_txt.Left
3620        .EndingShareface.Left = .Sumshareface_txt.Left

3630        .SumIcash.Left = .SumIcash.Left - 960&
3640        .SumICash_lbl.Left = .SumIcash.Left
3650        .SumICash_lbl_line.Left = .SumIcash.Left
3660        .SumICash_tot.Left = .SumIcash.Left
3670        .SumICash_tot_line1.Left = .SumIcash.Left + 120&
3680        .PreviousIcash.Left = .SumIcash.Left
3690        .EndingICash.Left = .SumIcash.Left

3700        .SumPcash.Left = .SumPcash.Left - 720&
3710        .SumPCash_lbl.Left = .SumPcash.Left
3720        .SumPCash_lbl_line.Left = .SumPcash.Left
3730        .SumPCash_tot.Left = .SumPcash.Left
3740        .SumPCash_tot_line1.Left = .SumPcash.Left + 120&
3750        .PreviousPcash.Left = .SumPcash.Left
3760        .EndingPCash.Left = .SumPcash.Left

3770        .SumCost.Left = .SumCost.Left - 480&
3780        .SumCost_lbl.Left = .SumCost.Left
3790        .SumCost_lbl_line.Left = .SumCost.Left
3800        .SumCost_tot.Left = .SumCost.Left
3810        .SumCost_tot_line1.Left = .SumCost.Left + 120&
3820        .PreviousCost.Left = .SumCost.Left
3830        .EndingCost.Left = .SumCost.Left

3840        .ledger_HIDDEN.Left = .ledger_HIDDEN.Left - 240&
3850        .ledger_HIDDEN_lbl.Left = .ledger_HIDDEN_lbl.Left - 240&
3860        .ledger_HIDDEN_lbl_line.Left = .ledger_HIDDEN_lbl.Left

3870      End Select

3880    End With

EXITP:
3890    Exit Sub

ERRH:
3900    Select Case ERR.Number
        Case Else
3910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3920    End Select
3930    Resume EXITP

End Sub
