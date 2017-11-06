VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetHistory_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetHistory_All"

'VGC 03/29/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_AssetHistory"

Private blnAll As Boolean
Private lngSF_Left As Long, lngSF_Width As Long
Private lngIC_Left As Long, lngIC_Width As Long
Private lngPC_Left As Long, lngPC_Width As Long
Private lngCO_Left As Long, lngCO_Width As Long

Private lngTpp As Long
Private dblEndingShareface_RunSum As Double
Private dblEndingICash_RunSum As Double, dblEndingPCash_RunSum As Double, dblEndingCost_RunSum As Double

Private dblEndingShareface_RunSum2 As Double
Private dblEndingICash_RunSum2 As Double, dblEndingPCash_RunSum2 As Double, dblEndingCost_RunSum2 As Double

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
              ' ** Nothing special.
540           blnAll = False
550           .Footer_hline01.Visible = False
560         Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue
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
920       strCusip = .cusip
930       .Caption = "Asset_History_All_" & strCusip & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")
940       lngTmp01 = .AsOf2.Height
950       lngTmp01 = (lngTmp01 * 2&)  ' ** AsOf2 and TotDesc2.
960       If .Page >= 2 Then
970         If .Header_lbl2.Visible = False Then
980           .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
990           .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
1000          .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1010          .journaltype_lbl.Top = (lngOneLine_Top + lngTmp01)
1020          .journaltype_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1030          .assetdate_lbl.Top = (lngTwoLine_Top + lngTmp01)
1040          .assetdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1050          .PurchaseDate_long_lbl.Top = (lngTwoLine_Top + lngTmp01)
1060          .PurchaseDate_long_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1070          .Sumshareface_txt_lbl.Top = (lngOneLine_Top + lngTmp01)
1080          .Sumshareface_txt_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1090          .SumICash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1100          .SumICash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1110          .SumPCash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1120          .SumPCash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1130          .SumCost_lbl.Top = (lngOneLine_Top + lngTmp01)
1140          .SumCost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1150          .ledger_HIDDEN_lbl.Top = (lngOneLine_Top + lngTmp01)
1160          .ledger_HIDDEN_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1170          .Header_lbl2.Visible = True
1180          .AsOf2.Visible = True
1190          .TotDesc2.Top = .AsOf2.Height
1200          .TotDesc2.Visible = True
1210          .TotDesc2_lbl.Top = .AsOf2.Height
1220          .TotDesc2_lbl.Visible = True
1230          .Header_hline05.Top = lngTmp01
1240          .Header_hline05.Visible = True
1250        End If
1260      Else
1270        If .Header_lbl2.Visible = True Then
1280          .Header_lbl2.Visible = False
1290          .AsOf2.Visible = False
1300          .TotDesc2.Visible = False
1310          .TotDesc2.Top = 0&
1320          .TotDesc2_lbl.Visible = False
1330          .TotDesc2_lbl.Top = 0
1340          .Header_hline05.Visible = False
1350          .Header_hline05.Top = .AsOf2.Height
1360          .transdate_lbl.Top = lngOneLine_Top
1370          .transdate_lbl_line.Top = lngUnderline_Top
1380          .journaltype_lbl.Top = lngOneLine_Top
1390          .journaltype_lbl_line.Top = lngUnderline_Top
1400          .assetdate_lbl.Top = lngTwoLine_Top
1410          .assetdate_lbl_line.Top = lngUnderline_Top
1420          .PurchaseDate_long_lbl.Top = lngTwoLine_Top
1430          .PurchaseDate_long_lbl_line.Top = lngUnderline_Top
1440          .Sumshareface_txt_lbl.Top = lngOneLine_Top
1450          .Sumshareface_txt_lbl_line.Top = lngUnderline_Top
1460          .SumICash_lbl.Top = lngTwoLine_Top
1470          .SumICash_lbl_line.Top = lngUnderline_Top
1480          .SumPCash_lbl.Top = lngTwoLine_Top
1490          .SumPCash_lbl_line.Top = lngUnderline_Top
1500          .SumCost_lbl.Top = lngOneLine_Top
1510          .SumCost_lbl_line.Top = lngUnderline_Top
1520          .ledger_HIDDEN_lbl.Top = lngOneLine_Top
1530          .ledger_HIDDEN_lbl_line.Top = lngUnderline_Top
1540          .Section(acPageHeader).Height = lngPageHeaderSection_Height
1550        End If
1560      End If
1570    End With

EXITP:
1580    Exit Sub

ERRH:
1590    Select Case ERR.Number
        Case Else
1600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1610    End Select
1620    Resume EXITP

End Sub

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader2_Format"

        Dim dblTmp01 As Double, dblTmp02 As Double, dblTmp03 As Double, dblTmp04 As Double

1710    With Me
1720      If blnAll = True Then
1730        Forms(strCallingForm).GetBal_Beg .accountno  ' ** Form Procedure: frmRpt_AssetHistory.
1740        Forms(strCallingForm).GetBal_End .accountno  ' ** Form Procedure: frmRpt_AssetHistory.
1750      End If
1760      dblTmp01 = Round(Nz(Forms(strCallingForm).PreviousShareface, 0#), 4)
1770      .PreviousShareface.Caption = Format(dblTmp01, "#,##0.0###;(#,##0.0###);0;0")
1780      If dblTmp01 < 0 Then .PreviousShareface.Left = .PreviousShareface.Left + 90&
1790      dblTmp02 = Nz(Forms(strCallingForm).PreviousIcash, 0#)
1800      dblTmp03 = Nz(Forms(strCallingForm).PreviousPcash, 0#)
1810      dblTmp04 = Nz(Forms(strCallingForm).PreviousCost, 0#)
1820      .PreviousIcash.Caption = Format(dblTmp02, "Currency")
1830      If dblTmp02 >= 0 Then .PreviousIcash.Width = .PreviousIcash.Width - 90&
1840      .PreviousPcash.Caption = Format(dblTmp03, "Currency")
1850      If dblTmp03 >= 0 Then .PreviousPcash.Width = .PreviousPcash.Width - 90&
1860      .PreviousCost.Caption = Format(dblTmp04, "Currency")
1870      If dblTmp04 >= 0 Then .PreviousCost.Width = .PreviousCost.Width - 90&
1880      .Sumshareface_tot.Left = lngSF_Left
1890      .Sumshareface_tot.Width = lngSF_Width
1900      .EndingShareface.Left = lngSF_Left
1910      .EndingShareface.Width = lngSF_Width
1920      .EndingICash.Left = lngIC_Left - 90&
1930      .EndingICash.Width = lngIC_Width
1940      .EndingPCash.Left = lngPC_Left - 90&
1950      .EndingPCash.Width = lngPC_Width
1960      .EndingCost.Left = lngCO_Left + 45&
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

2120      dblTmp01 = Round(Nz(Forms(strCallingForm).EndingShareface, 0#), 4)
2130      dblEndingShareface_RunSum = dblEndingShareface_RunSum + dblTmp01
2140      .EndingShareface.Caption = Format(dblTmp01, "#,##0.0###;(#,##0.0###);0;0")
2150      If dblTmp01 < 0 Then .EndingShareface.Left = .EndingShareface.Left + 90&
2160      dblTmp02 = Nz(Forms(strCallingForm).EndingICash, 0#)
2170      dblEndingICash_RunSum = dblEndingICash_RunSum + dblTmp02
2180      dblTmp03 = Nz(Forms(strCallingForm).EndingPCash, 0#)
2190      dblEndingPCash_RunSum = dblEndingICash_RunSum + dblTmp03
2200      dblTmp04 = Nz(Forms(strCallingForm).EndingCost, 0#)
2210      dblEndingCost_RunSum = dblEndingICash_RunSum + dblTmp04
2220      .EndingICash.Caption = Format(Forms(strCallingForm).EndingICash, "Currency")
2230      If dblTmp02 >= 0 Then .EndingICash.Width = .EndingICash.Width - 90&
2240      .EndingPCash.Caption = Format(Forms(strCallingForm).EndingPCash, "Currency")
2250      If dblTmp03 >= 0 Then .EndingPCash.Width = .EndingPCash.Width - 90&
2260      .EndingCost.Caption = Format(Forms(strCallingForm).EndingCost, "Currency")
2270      If dblTmp04 >= 0 Then .EndingCost.Width = .EndingCost.Width - 90&

2280      dblTmp01 = .Sumshareface_runsum
2290      If dblTmp01 < 0 Then .Sumshareface_tot.Left = .Sumshareface_tot.Left + 60&
2300      .PreviousShareface.Left = lngSF_Left
2310      .PreviousShareface.Width = lngSF_Width
2320      .PreviousIcash.Left = lngIC_Left
2330      .PreviousIcash.Width = lngIC_Width
2340      .PreviousPcash.Left = lngPC_Left
2350      .PreviousPcash.Width = lngPC_Width
2360      .PreviousCost.Left = lngCO_Left
2370      .PreviousCost.Width = lngCO_Width

2380      .EndingICash.Left = .EndingICash.Left + 90&
2390      .EndingPCash.Left = .EndingPCash.Left + 90&
2400      .EndingCost.Left = .EndingCost.Left - 45&

2410    End With

EXITP:
2420    Exit Sub

ERRH:
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportFooter_Format"

2510    With Me
          ' ** These totals were doubling!
2520      If dblEndingShareface_RunSum2 = 0# Then
2530        dblEndingShareface_RunSum2 = dblEndingShareface_RunSum
2540      End If
2550      If dblEndingICash_RunSum2 = 0# Then
2560        dblEndingICash_RunSum2 = dblEndingICash_RunSum
2570      End If
2580      If dblEndingPCash_RunSum2 = 0# Then
2590        dblEndingPCash_RunSum2 = dblEndingPCash_RunSum
2600      End If
2610      If dblEndingCost_RunSum2 = 0# Then
2620        dblEndingCost_RunSum2 = dblEndingCost_RunSum
2630      End If
2640      .EndingShareface_Tot.Caption = Format(dblEndingShareface_RunSum2, "#,##0.0###;(#,##0.0###);0;0")
2650      .EndingShareface_Tot_narrow.Caption = Format(dblEndingShareface_RunSum2, "#,##0.0###;(#,##0.0###);0;0")
2660      .EndingICash_Tot.Caption = Format(dblEndingICash_RunSum2, "Currency")
2670      .EndingICash_Tot_narrow.Caption = Format(dblEndingICash_RunSum2, "Currency")
2680      .EndingPCash_Tot.Caption = Format(dblEndingPCash_RunSum2, "Currency")
2690      .EndingPCash_Tot_narrow.Caption = Format(dblEndingPCash_RunSum2, "Currency")
2700      .EndingCost_Tot.Caption = Format(dblEndingCost_RunSum2, "Currency")
2710      .EndingCost_Tot_narrow.Caption = Format(dblEndingCost_RunSum2, "Currency")
2720    End With

        'dblEndingShareface_RunSum: 0  GroupFooter1_Format()
        'dblEndingShareface_RunSum: 613810.47  ReportFooter_Format()
        'dblEndingShareface_RunSum: 613810.47  GroupFooter1_Format()
        'dblEndingShareface_RunSum: 1227620.94  ReportFooter_Format()

EXITP:
2730    Exit Sub

ERRH:
2740    Select Case ERR.Number
        Case Else
2750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2760    End Select
2770    Resume EXITP

End Sub

Private Sub Report_Close()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2810    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub SetPurchaseDate(blnShow As Boolean)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "SetPurchaseDate"

2910    With Me
2920      Select Case blnShow
          Case True

2930        .SumCost.Width = .SumCost.Width - (13& * 15&)
2940        .SumCost_lbl.Width = .SumCost.Width
2950        .SumCost_lbl_line.Width = .SumCost.Width
2960        .SumCost_tot.Width = .SumCost.Width
2970        .SumCost_tot_line1.Width = .SumCost_tot_line1.Width - (13& * 15&)
2980        .PreviousCost.Width = .SumCost.Width
2990        .EndingCost.Width = .SumCost.Width

3000        .SumCost.Left = .SumCost.Left + (13& * 15&)
3010        .SumCost_lbl.Left = .SumCost.Left
3020        .SumCost_lbl_line.Left = .SumCost.Left
3030        .SumCost_tot.Left = .SumCost.Left
3040        .SumCost_tot_line1.Left = .SumCost.Left + 120&
3050        .PreviousCost.Left = .SumCost.Left
3060        .EndingCost.Left = .SumCost.Left
3070        .EndingCost_Tot.Left = (.SumCost.Left + .SumCost.Width) - .EndingCost_Tot.Width
3080        .EndingCost_Tot_line.Left = (.SumCost_tot_line1.Left + .SumCost_tot_line1.Width) - .EndingCost_Tot_line.Width
3090        .EndingCost_Tot_narrow.Left = .EndingCost_Tot.Left
3100        .EndingCost_Tot_narrow.Visible = True
3110        .EndingCost_Tot.Visible = False

3120        .SumPcash.Width = .SumPcash.Width - (13& * 15&)
3130        .SumPCash_lbl.Width = .SumPcash.Width
3140        .SumPCash_lbl_line.Width = .SumPcash.Width
3150        .SumPCash_tot.Width = .SumPcash.Width
3160        .SumPCash_tot_line1.Width = .SumPCash_tot_line1.Width - (13& * 15&)
3170        .PreviousPcash.Width = .SumPcash.Width
3180        .EndingPCash.Width = .SumPcash.Width

3190        .SumPcash.Left = .SumPcash.Left + (2& * (13& * 15&))
3200        .SumPCash_lbl.Left = .SumPcash.Left
3210        .SumPCash_lbl_line.Left = .SumPcash.Left
3220        .SumPCash_tot.Left = .SumPcash.Left
3230        .SumPCash_tot_line1.Left = .SumPcash.Left + 120&
3240        .PreviousPcash.Left = .SumPcash.Left
3250        .EndingPCash.Left = .SumPcash.Left
3260        .EndingPCash_Tot.Left = (.SumPcash.Left + .SumPcash.Width) - .EndingPCash_Tot.Width
3270        .EndingPCash_Tot_line.Left = (.SumPCash_tot_line1.Left + .SumPCash_tot_line1.Width) - .EndingPCash_Tot_line.Width
3280        .EndingPCash_Tot_narrow.Left = .EndingPCash_Tot.Left
3290        .EndingPCash_Tot_narrow.Visible = True
3300        .EndingPCash_Tot.Visible = False

3310        .SumIcash.Width = .SumIcash.Width - (13& * 15&)
3320        .SumICash_lbl.Width = .SumIcash.Width
3330        .SumICash_lbl_line.Width = .SumIcash.Width
3340        .SumICash_tot.Width = .SumIcash.Width
3350        .SumICash_tot_line1.Width = .SumICash_tot_line1.Width - (13& * 15&)
3360        .PreviousIcash.Width = .SumIcash.Width
3370        .EndingICash.Width = .SumIcash.Width

3380        .SumIcash.Left = .SumIcash.Left + (3& * (13& * 15&))
3390        .SumICash_lbl.Left = .SumIcash.Left
3400        .SumICash_lbl_line.Left = .SumIcash.Left
3410        .SumICash_tot.Left = .SumIcash.Left
3420        .SumICash_tot_line1.Left = .SumIcash.Left + 120&
3430        .PreviousIcash.Left = .SumIcash.Left
3440        .EndingICash.Left = .SumIcash.Left
3450        .EndingICash_Tot.Left = (.SumIcash.Left + .SumIcash.Width) - .EndingICash_Tot.Width
3460        .EndingICash_Tot_line.Left = (.SumICash_tot_line1.Left + .SumICash_tot_line1.Width) - .EndingICash_Tot_line.Width
3470        .EndingICash_Tot_narrow.Left = .EndingICash_Tot.Left
3480        .EndingICash_Tot_narrow.Visible = True
3490        .EndingICash_Tot.Visible = False

3500        .Sumshareface_txt.Width = .Sumshareface_txt.Width - (13& * 15&)
3510        .Sumshareface_txt_wparens.Width = .Sumshareface_txt_wparens.Width - (13& * 15&)
3520        .Sumshareface_txt_lbl.Width = .Sumshareface_txt.Width
3530        .Sumshareface_txt_lbl_line.Width = .Sumshareface_txt.Width
3540        .Sumshareface_txt_Liability.Width = .Sumshareface_txt.Width
3550        .Sumshareface_tot.Width = .Sumshareface_txt.Width
3560        .Sumshareface_tot_line1.Width = .Sumshareface_tot_line1.Width - (13& * 15&)
3570        .PreviousShareface.Width = .Sumshareface_txt.Width
3580        .EndingShareface.Width = .Sumshareface_txt.Width

3590        .Sumshareface_txt.Left = .Sumshareface_txt.Left + (4& * (13& * 15&))
3600        .Sumshareface_txt_wparens.Left = .Sumshareface_txt_wparens.Left + (4& * (13& * 15&))
3610        .Sumshareface_txt_lbl.Left = .Sumshareface_txt.Left
3620        .Sumshareface_txt_lbl_line.Left = .Sumshareface_txt.Left
3630        .Sumshareface_txt_Liability.Left = .Sumshareface_txt.Left
3640        .Sumshareface_tot.Left = .Sumshareface_txt.Left
3650        .Sumshareface_tot_line1.Left = ((.Sumshareface_tot.Left + .Sumshareface_tot.Width) - .Sumshareface_tot_line1.Width) + 30&
3660        .PreviousShareface.Left = .Sumshareface_txt.Left
3670        .EndingShareface.Left = .Sumshareface_txt.Left
3680        .EndingShareface_Tot.Left = (.Sumshareface_txt.Left + .Sumshareface_txt.Width) - .EndingShareface_Tot.Width
3690        .EndingShareface_Tot_line.Left = (.Sumshareface_tot_line1.Left + .Sumshareface_tot_line1.Width) - .EndingShareface_Tot_line.Width
3700        .EndingShareface_Tot_narrow.Left = .EndingShareface_Tot.Left
3710        .EndingShareface_Tot_narrow.Visible = True
3720        .EndingShareface_Tot.Visible = False

3730        .BeginningBalance_lbl.Left = .BeginningBalance_lbl.Left + (4& * (13& * 15&))
3740        .PeriodActivity_lbl.Left = .BeginningBalance_lbl.Left
3750        .EndingBalance_lbl.Left = .BeginningBalance_lbl.Left

3760        .assetdate_dp.Visible = True
3770        .assetdate_ws.Visible = True
3780        .assetdate.Visible = False
3790        .assetdate_lbl.Width = .assetdate_dp.Width
3800        .assetdate_lbl_line.Width = .assetdate_dp.Width
3810        .assetdate_lbl.Caption = vbCrLf & "Trade Date"

3820        .PurchaseDate_long.Left = (.assetdate_dp.Left + .assetdate_dp.Width) + 60&
3830        .PurchaseDate_long.Visible = True
3840        .PurchaseDate_long_lbl.Left = .PurchaseDate_long.Left
3850        .PurchaseDate_long_lbl.Visible = True
3860        .PurchaseDate_long_lbl_line.Left = .PurchaseDate_long.Left
3870        .PurchaseDate_long_lbl_line.Visible = True

3880      Case False
3890        .PurchaseDate_long.Visible = False
3900        .PurchaseDate_long_lbl.Visible = False
3910        .PurchaseDate_long_lbl_line.Visible = False

3920        .assetdate.Left = .assetdate.Left + 720&
3930        .assetdate_lbl.Left = .assetdate.Left
3940        .assetdate_lbl_line.Left = .assetdate.Left

3950        .journaltype.Left = .journaltype.Left + 480&
3960        .journaltype_lbl.Left = .journaltype.Left
3970        .journaltype_lbl_line.Left = .journaltype.Left

3980        .transdate.Left = .transdate.Left + 240&
3990        .transdate_lbl.Left = .transdate.Left
4000        .transdate_lbl_line.Left = .transdate.Left

4010        .BeginningBalance_lbl.Left = .BeginningBalance_lbl.Left - 1200&
4020        .PeriodActivity_lbl.Left = .BeginningBalance_lbl.Left
4030        .EndingBalance_lbl.Left = .BeginningBalance_lbl.Left
4040        .EndingBalance_Tot_lbl.Left = (.BeginningBalance_lbl.Left + .BeginningBalance_lbl.Width) - .EndingBalance_Tot_lbl.Width

4050        .Sumshareface_txt.Left = .Sumshareface_txt.Left - 1200&
4060        .Sumshareface_txt_wparens.Left = .Sumshareface_txt_wparens.Left - 1200&
4070        .Sumshareface_txt_lbl.Left = .Sumshareface_txt.Left
4080        .Sumshareface_txt_lbl_line.Left = .Sumshareface_txt.Left
4090        .Sumshareface_txt_Liability.Left = .Sumshareface_txt.Left
4100        .Sumshareface_tot.Left = .Sumshareface_txt.Left
4110        .Sumshareface_tot_line1.Left = ((.Sumshareface_tot.Left + .Sumshareface_tot.Width) - .Sumshareface_tot_line1.Width) + 30&
4120        .PreviousShareface.Left = .Sumshareface_txt.Left
4130        .EndingShareface.Left = .Sumshareface_txt.Left
4140        .EndingShareface_Tot.Left = (.Sumshareface_txt.Left + .Sumshareface_txt.Width) - .EndingShareface_Tot.Width
4150        .EndingShareface_Tot_line.Left = (.Sumshareface_tot_line1.Left + .Sumshareface_tot_line1.Width) - .EndingShareface_Tot_line.Width
4160        .EndingShareface_Tot_narrow.Left = .EndingShareface_Tot.Left

4170        .SumIcash.Left = .SumIcash.Left - 960&
4180        .SumICash_lbl.Left = .SumIcash.Left
4190        .SumICash_lbl_line.Left = .SumIcash.Left
4200        .SumICash_tot.Left = .SumIcash.Left
4210        .SumICash_tot_line1.Left = .SumIcash.Left + 120&
4220        .PreviousIcash.Left = .SumIcash.Left
4230        .EndingICash.Left = .SumIcash.Left
4240        .EndingICash_Tot.Left = (.SumIcash.Left + .SumIcash.Width) - .EndingICash_Tot.Width
4250        .EndingICash_Tot_line.Left = (.SumICash_tot_line1.Left + .SumICash_tot_line1.Width) - .EndingICash_Tot_line.Width
4260        .EndingICash_Tot_narrow.Left = .EndingICash_Tot.Left

4270        .SumPcash.Left = .SumPcash.Left - 720&
4280        .SumPCash_lbl.Left = .SumPcash.Left
4290        .SumPCash_lbl_line.Left = .SumPcash.Left
4300        .SumPCash_tot.Left = .SumPcash.Left
4310        .SumPCash_tot_line1.Left = .SumPcash.Left + 120&
4320        .PreviousPcash.Left = .SumPcash.Left
4330        .EndingPCash.Left = .SumPcash.Left
4340        .EndingPCash_Tot.Left = (.SumPcash.Left + .SumPcash.Width) - .EndingPCash_Tot.Width
4350        .EndingPCash_Tot_line.Left = (.SumPCash_tot_line1.Left + .SumPCash_tot_line1.Width) - .EndingPCash_Tot_line.Width
4360        .EndingPCash_Tot_narrow.Left = .EndingPCash_Tot.Left

4370        .SumCost.Left = .SumCost.Left - 480&
4380        .SumCost_lbl.Left = .SumCost.Left
4390        .SumCost_lbl_line.Left = .SumCost.Left
4400        .SumCost_tot.Left = .SumCost.Left
4410        .SumCost_tot_line1.Left = .SumCost.Left + 120&
4420        .PreviousCost.Left = .SumCost.Left
4430        .EndingCost.Left = .SumCost.Left
4440        .EndingCost_Tot.Left = (.SumCost.Left + .SumCost.Width) - .EndingCost_Tot.Width
4450        .EndingCost_Tot_line.Left = (.SumCost_tot_line1.Left + .SumCost_tot_line1.Width) - .EndingCost_Tot_line.Width
4460        .EndingCost_Tot_narrow.Left = .EndingCost_Tot.Left

4470        .ledger_HIDDEN.Left = .ledger_HIDDEN.Left - 240&
4480        .ledger_HIDDEN_lbl.Left = .ledger_HIDDEN_lbl.Left - 240&
4490        .ledger_HIDDEN_lbl_line.Left = .ledger_HIDDEN_lbl.Left

4500      End Select

4510    End With

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub
