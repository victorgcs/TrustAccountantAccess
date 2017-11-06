Attribute VB_Name = "modAcctExportFuncs"
Option Compare Database
Option Explicit

'VGC 09/09/2017: CHANGES!

Private Const THIS_NAME As String = "modAcctExportFuncs"
' **

Public Sub Tier1_Enable_AE(blnAble As Boolean, blnNoData As Boolean, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Tier1_Enable_AE"

110     With frm
120       Select Case blnAble
          Case True
130         .ckgTier_chkTier1.Enabled = True
140         .ckgTier_chkTier1_lbl2.ForeColor = CLR_VDKGRY
150         .ckgTier_chkTier1_lbl2_dim_hi.Visible = False
160         Select Case .ckgTier_chkTier1
            Case True
170           .ckgTier_chkTier1_box.BackStyle = acBackStyleNormal
180           .ckgTier_chkTier1_box2.BackStyle = acBackStyleNormal
190           .ckgTier_chkTier1_hline03.BorderColor = MY_CLR_VLTBGE
200           .ckgTier1_chk01_Account.Enabled = True
210           .ckgTier1_chk02_MasterAsset.Enabled = True
220           .ckgTier1_chk03_ActiveAssets.Enabled = True
230           .ckgTier1_chk04_LedgerArchive.Enabled = True
240           .ckgTier1_chk05_Ledger.Enabled = True
250           .ckgTier1_chk06_Balance.Enabled = True
260           .ckgTier1_chk01_Account_lbl_recs.ForeColor = CLR_DKGRY
270           .ckgTier1_chk01_Account_lbl_recs.BorderColor = CLR_VLTBLU2
280           .ckgTier1_chk02_MasterAsset_lbl_recs.ForeColor = CLR_DKGRY
290           .ckgTier1_chk02_MasterAsset_lbl_recs.BorderColor = CLR_VLTBLU2
300           .ckgTier1_chk03_ActiveAssets_lbl_recs.ForeColor = CLR_DKGRY
310           .ckgTier1_chk03_ActiveAssets_lbl_recs.BorderColor = CLR_VLTBLU2
320           .ckgTier1_chk04_LedgerArchive_lbl_recs.ForeColor = CLR_DKGRY
330           .ckgTier1_chk04_LedgerArchive_lbl_recs.BorderColor = CLR_VLTBLU2
340           .ckgTier1_chk05_Ledger_lbl_recs.ForeColor = CLR_DKGRY
350           .ckgTier1_chk05_Ledger_lbl_recs.BorderColor = CLR_VLTBLU2
360           .ckgTier1_chk06_Balance_lbl_recs.ForeColor = CLR_DKGRY
370           .ckgTier1_chk06_Balance_lbl_recs.BorderColor = CLR_VLTBLU2
380           .cmdTier1_Select_box.BackStyle = acBackStyleNormal
390           .cmdTier1_Select_hline03.BorderColor = MY_CLR_VLTBGE
400           .cmdTier1_Select_hline04.BorderColor = MY_CLR_VLTBGE
410           If blnNoData = False Then
420             .cmdTier1_SelectAll.Enabled = True
430             .cmdTier1_SelectAll_raised_img.Visible = True
440             .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
450             .cmdTier1_SelectAll_raised_focus_img.Visible = False
460             .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
470             .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
480             .cmdTier1_SelectAll_raised_img_dis.Visible = False
490             .cmdTier1_SelectNone.Enabled = True
500             .cmdTier1_SelectNone_raised_img.Visible = True
510             .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
520             .cmdTier1_SelectNone_raised_focus_img.Visible = False
530             .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
540             .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
550             .cmdTier1_SelectNone_raised_img_dis.Visible = False
560           End If
570         Case False
580           .ckgTier_chkTier1_box.BackStyle = acBackStyleTransparent
590           .ckgTier_chkTier1_box2.BackStyle = acBackStyleTransparent
600           .ckgTier_chkTier1_hline03.BorderColor = MY_CLR_LTBGE
610           .ckgTier1_chk01_Account.Enabled = False
620           .ckgTier1_chk02_MasterAsset.Enabled = False
630           .ckgTier1_chk03_ActiveAssets.Enabled = False
640           .ckgTier1_chk04_LedgerArchive.Enabled = False
650           .ckgTier1_chk05_Ledger.Enabled = False
660           .ckgTier1_chk06_Balance.Enabled = False
670           .cmdTier1_Select_box.BackStyle = acBackStyleTransparent
680           .ckgTier1_chk01_Account_lbl_recs.ForeColor = WIN_CLR_DISF
690           .ckgTier1_chk01_Account_lbl_recs.BorderColor = WIN_CLR_DISR
700           .ckgTier1_chk02_MasterAsset_lbl_recs.ForeColor = WIN_CLR_DISF
710           .ckgTier1_chk02_MasterAsset_lbl_recs.BorderColor = WIN_CLR_DISR
720           .ckgTier1_chk03_ActiveAssets_lbl_recs.ForeColor = WIN_CLR_DISF
730           .ckgTier1_chk03_ActiveAssets_lbl_recs.BorderColor = WIN_CLR_DISR
740           .ckgTier1_chk04_LedgerArchive_lbl_recs.ForeColor = WIN_CLR_DISF
750           .ckgTier1_chk04_LedgerArchive_lbl_recs.BorderColor = WIN_CLR_DISR
760           .ckgTier1_chk05_Ledger_lbl_recs.ForeColor = WIN_CLR_DISF
770           .ckgTier1_chk05_Ledger_lbl_recs.BorderColor = WIN_CLR_DISR
780           .ckgTier1_chk06_Balance_lbl_recs.ForeColor = WIN_CLR_DISF
790           .ckgTier1_chk06_Balance_lbl_recs.BorderColor = WIN_CLR_DISR
800           .cmdTier1_Select_hline03.BorderColor = MY_CLR_LTBGE
810           .cmdTier1_Select_hline04.BorderColor = MY_CLR_LTBGE
820           If blnNoData = False Then
830             .cmdTier1_SelectAll.Enabled = False
840             .cmdTier1_SelectAll_raised_img_dis.Visible = True
850             .cmdTier1_SelectAll_raised_img.Visible = False
860             .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
870             .cmdTier1_SelectAll_raised_focus_img.Visible = False
880             .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
890             .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
900             .cmdTier1_SelectNone.Enabled = False
910             .cmdTier1_SelectNone_raised_img_dis.Visible = True
920             .cmdTier1_SelectNone_raised_img.Visible = False
930             .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
940             .cmdTier1_SelectNone_raised_focus_img.Visible = False
950             .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
960             .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
970           End If
980         End Select
990       Case False
1000        .ckgTier_chkTier1.Enabled = False
1010        .ckgTier_chkTier1_lbl2.ForeColor = WIN_CLR_DISF
1020        If blnNoData = False Then
1030          .ckgTier_chkTier1_lbl2_dim_hi.Visible = True
1040        End If
1050        .ckgTier_chkTier1_box.BackStyle = acBackStyleTransparent
1060        .ckgTier_chkTier1_box2.BackStyle = acBackStyleTransparent
1070        .ckgTier_chkTier1_hline03.BorderColor = MY_CLR_LTBGE
1080        .ckgTier1_chk01_Account.Enabled = False
1090        .ckgTier1_chk02_MasterAsset.Enabled = False
1100        .ckgTier1_chk03_ActiveAssets.Enabled = False
1110        .ckgTier1_chk04_LedgerArchive.Enabled = False
1120        .ckgTier1_chk05_Ledger.Enabled = False
1130        .ckgTier1_chk06_Balance.Enabled = False
1140        .ckgTier1_chk01_Account_lbl_recs.ForeColor = WIN_CLR_DISF
1150        .ckgTier1_chk01_Account_lbl_recs.BorderColor = WIN_CLR_DISR
1160        .ckgTier1_chk02_MasterAsset_lbl_recs.ForeColor = WIN_CLR_DISF
1170        .ckgTier1_chk02_MasterAsset_lbl_recs.BorderColor = WIN_CLR_DISR
1180        .ckgTier1_chk03_ActiveAssets_lbl_recs.ForeColor = WIN_CLR_DISF
1190        .ckgTier1_chk03_ActiveAssets_lbl_recs.BorderColor = WIN_CLR_DISR
1200        .ckgTier1_chk04_LedgerArchive_lbl_recs.ForeColor = WIN_CLR_DISF
1210        .ckgTier1_chk04_LedgerArchive_lbl_recs.BorderColor = WIN_CLR_DISR
1220        .ckgTier1_chk05_Ledger_lbl_recs.ForeColor = WIN_CLR_DISF
1230        .ckgTier1_chk05_Ledger_lbl_recs.BorderColor = WIN_CLR_DISR
1240        .ckgTier1_chk06_Balance_lbl_recs.ForeColor = WIN_CLR_DISF
1250        .ckgTier1_chk06_Balance_lbl_recs.BorderColor = WIN_CLR_DISR
1260        .cmdTier1_Select_box.BackStyle = acBackStyleTransparent
1270        .cmdTier1_Select_hline03.BorderColor = MY_CLR_LTBGE
1280        .cmdTier1_Select_hline04.BorderColor = MY_CLR_LTBGE
1290        If blnNoData = False Then
1300          .cmdTier1_SelectAll.Enabled = False
1310          .cmdTier1_SelectAll_raised_img_dis.Visible = True
1320          .cmdTier1_SelectAll_raised_img.Visible = False
1330          .cmdTier1_SelectAll_raised_semifocus_dots_img.Visible = False
1340          .cmdTier1_SelectAll_raised_focus_img.Visible = False
1350          .cmdTier1_SelectAll_raised_focus_dots_img.Visible = False
1360          .cmdTier1_SelectAll_sunken_focus_dots_img.Visible = False
1370          .cmdTier1_SelectNone.Enabled = False
1380          .cmdTier1_SelectNone_raised_img_dis.Visible = True
1390          .cmdTier1_SelectNone_raised_img.Visible = False
1400          .cmdTier1_SelectNone_raised_semifocus_dots_img.Visible = False
1410          .cmdTier1_SelectNone_raised_focus_img.Visible = False
1420          .cmdTier1_SelectNone_raised_focus_dots_img.Visible = False
1430          .cmdTier1_SelectNone_sunken_focus_dots_img.Visible = False
1440        End If
1450      End Select
1460    End With

EXITP:
1470    Exit Sub

ERRH:
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Public Sub Tier2_Enable_AE(blnAble As Boolean, blnNoData As Boolean, frm As Access.Form)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Tier2_Enable_AE"

1610    With frm
1620      Select Case blnAble
          Case True
1630        .ckgTier_chkTier2.Enabled = True
1640        .ckgTier_chkTier2_lbl2.ForeColor = CLR_VDKGRY
1650        .ckgTier_chkTier2_lbl2_dim_hi.Visible = False
1660        Select Case .ckgTier_chkTier2
            Case True
1670          .ckgTier_chkTier2_box.BackStyle = acBackStyleNormal
1680          .ckgTier_chkTier2_box2.BackStyle = acBackStyleNormal
1690          .ckgTier_chkTier2_hline03.BorderColor = MY_CLR_VLTBGE
1700          .ckgTier2_chk01_m_REVCODE.Enabled = True
1710          .ckgTier2_chk02_RecurringItems.Enabled = True
1720          .ckgTier2_chk03_Pricing_MasterAsset_History.Enabled = True
1730          .ckgTier2_chk04_Currency_History.Enabled = True
1740          .ckgTier2_chk05_LedgerHidden.Enabled = True
1750          .ckgTier2_chk06_Location.Enabled = True
1760          .ckgTier2_chk01_m_REVCODE_lbl_recs.ForeColor = CLR_DKGRY
1770          .ckgTier2_chk01_m_REVCODE_lbl_recs.BorderColor = CLR_VLTBLU2
1780          .ckgTier2_chk02_RecurringItems_lbl_recs.ForeColor = CLR_DKGRY
1790          .ckgTier2_chk02_RecurringItems_lbl_recs.BorderColor = CLR_VLTBLU2
1800          .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.ForeColor = CLR_DKGRY
1810          .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.BorderColor = CLR_VLTBLU2
1820          .ckgTier2_chk04_Currency_History_lbl_recs.ForeColor = CLR_DKGRY
1830          .ckgTier2_chk04_Currency_History_lbl_recs.BorderColor = CLR_VLTBLU2
1840          .ckgTier2_chk05_LedgerHidden_lbl_recs.ForeColor = CLR_DKGRY
1850          .ckgTier2_chk05_LedgerHidden_lbl_recs.BorderColor = CLR_VLTBLU2
1860          .ckgTier2_chk06_Location_lbl_recs.ForeColor = CLR_DKGRY
1870          .ckgTier2_chk06_Location_lbl_recs.BorderColor = CLR_VLTBLU2
1880          .cmdTier2_Select_box.BackStyle = acBackStyleNormal
1890          .cmdTier2_Select_hline03.BorderColor = MY_CLR_VLTBGE
1900          .cmdTier2_Select_hline04.BorderColor = MY_CLR_VLTBGE
1910          If blnNoData = False Then
1920            .cmdTier2_SelectAll.Enabled = True
1930            .cmdTier2_SelectAll_raised_img.Visible = True
1940            .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
1950            .cmdTier2_SelectAll_raised_focus_img.Visible = False
1960            .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
1970            .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
1980            .cmdTier2_SelectAll_raised_img_dis.Visible = False
1990            .cmdTier2_SelectNone.Enabled = True
2000            .cmdTier2_SelectNone_raised_img.Visible = True
2010            .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
2020            .cmdTier2_SelectNone_raised_focus_img.Visible = False
2030            .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
2040            .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
2050            .cmdTier2_SelectNone_raised_img_dis.Visible = False
2060          End If
2070        Case False
2080          .ckgTier_chkTier2_box.BackStyle = acBackStyleTransparent
2090          .ckgTier_chkTier2_box2.BackStyle = acBackStyleTransparent
2100          .ckgTier_chkTier2_hline03.BorderColor = MY_CLR_LTBGE
2110          .ckgTier2_chk01_m_REVCODE.Enabled = False
2120          .ckgTier2_chk02_RecurringItems.Enabled = False
2130          .ckgTier2_chk03_Pricing_MasterAsset_History.Enabled = False
2140          .ckgTier2_chk04_Currency_History.Enabled = False
2150          .ckgTier2_chk05_LedgerHidden.Enabled = False
2160          .ckgTier2_chk06_Location.Enabled = False
2170          .ckgTier2_chk01_m_REVCODE_lbl_recs.ForeColor = WIN_CLR_DISF
2180          .ckgTier2_chk01_m_REVCODE_lbl_recs.BorderColor = WIN_CLR_DISR
2190          .ckgTier2_chk02_RecurringItems_lbl_recs.ForeColor = WIN_CLR_DISF
2200          .ckgTier2_chk02_RecurringItems_lbl_recs.BorderColor = WIN_CLR_DISR
2210          .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.ForeColor = WIN_CLR_DISF
2220          .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.BorderColor = WIN_CLR_DISR
2230          .ckgTier2_chk04_Currency_History_lbl_recs.ForeColor = WIN_CLR_DISF
2240          .ckgTier2_chk04_Currency_History_lbl_recs.BorderColor = WIN_CLR_DISR
2250          .ckgTier2_chk05_LedgerHidden_lbl_recs.ForeColor = WIN_CLR_DISF
2260          .ckgTier2_chk05_LedgerHidden_lbl_recs.BorderColor = WIN_CLR_DISR
2270          .ckgTier2_chk06_Location_lbl_recs.ForeColor = WIN_CLR_DISF
2280          .ckgTier2_chk06_Location_lbl_recs.BorderColor = WIN_CLR_DISR
2290          .cmdTier2_Select_box.BackStyle = acBackStyleTransparent
2300          .cmdTier2_Select_hline03.BorderColor = MY_CLR_LTBGE
2310          .cmdTier2_Select_hline04.BorderColor = MY_CLR_LTBGE
2320          If blnNoData = False Then
2330            .cmdTier2_SelectAll.Enabled = False
2340            .cmdTier2_SelectAll_raised_img_dis.Visible = True
2350            .cmdTier2_SelectAll_raised_img.Visible = False
2360            .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
2370            .cmdTier2_SelectAll_raised_focus_img.Visible = False
2380            .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
2390            .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
2400            .cmdTier2_SelectNone.Enabled = False
2410            .cmdTier2_SelectNone_raised_img_dis.Visible = True
2420            .cmdTier2_SelectNone_raised_img.Visible = False
2430            .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
2440            .cmdTier2_SelectNone_raised_focus_img.Visible = False
2450            .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
2460            .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
2470          End If
2480        End Select
2490      Case False
2500        .ckgTier_chkTier2.Enabled = False
2510        .ckgTier_chkTier2_lbl2.ForeColor = WIN_CLR_DISF
2520        If blnNoData = False Then
2530          .ckgTier_chkTier2_lbl2_dim_hi.Visible = True
2540        End If
2550        .ckgTier_chkTier2_box.BackStyle = acBackStyleTransparent
2560        .ckgTier_chkTier2_box2.BackStyle = acBackStyleTransparent
2570        .ckgTier_chkTier2_hline03.BorderColor = MY_CLR_LTBGE
2580        .ckgTier2_chk01_m_REVCODE.Enabled = False
2590        .ckgTier2_chk02_RecurringItems.Enabled = False
2600        .ckgTier2_chk03_Pricing_MasterAsset_History.Enabled = False
2610        .ckgTier2_chk04_Currency_History.Enabled = False
2620        .ckgTier2_chk05_LedgerHidden.Enabled = False
2630        .ckgTier2_chk06_Location.Enabled = False
2640        .ckgTier2_chk01_m_REVCODE_lbl_recs.ForeColor = WIN_CLR_DISF
2650        .ckgTier2_chk01_m_REVCODE_lbl_recs.BorderColor = WIN_CLR_DISR
2660        .ckgTier2_chk02_RecurringItems_lbl_recs.ForeColor = WIN_CLR_DISF
2670        .ckgTier2_chk02_RecurringItems_lbl_recs.BorderColor = WIN_CLR_DISR
2680        .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.ForeColor = WIN_CLR_DISF
2690        .ckgTier2_chk03_Pricing_MasterAsset_History_lbl_recs.BorderColor = WIN_CLR_DISR
2700        .ckgTier2_chk04_Currency_History_lbl_recs.ForeColor = WIN_CLR_DISF
2710        .ckgTier2_chk04_Currency_History_lbl_recs.BorderColor = WIN_CLR_DISR
2720        .ckgTier2_chk05_LedgerHidden_lbl_recs.ForeColor = WIN_CLR_DISF
2730        .ckgTier2_chk05_LedgerHidden_lbl_recs.BorderColor = WIN_CLR_DISR
2740        .ckgTier2_chk06_Location_lbl_recs.ForeColor = WIN_CLR_DISF
2750        .ckgTier2_chk06_Location_lbl_recs.BorderColor = WIN_CLR_DISR
2760        .cmdTier2_Select_box.BackStyle = acBackStyleTransparent
2770        .cmdTier2_Select_hline03.BorderColor = MY_CLR_LTBGE
2780        .cmdTier2_Select_hline04.BorderColor = MY_CLR_LTBGE
2790        If blnNoData = False Then
2800          .cmdTier2_SelectAll.Enabled = False
2810          .cmdTier2_SelectAll_raised_img_dis.Visible = True
2820          .cmdTier2_SelectAll_raised_img.Visible = False
2830          .cmdTier2_SelectAll_raised_semifocus_dots_img.Visible = False
2840          .cmdTier2_SelectAll_raised_focus_img.Visible = False
2850          .cmdTier2_SelectAll_raised_focus_dots_img.Visible = False
2860          .cmdTier2_SelectAll_sunken_focus_dots_img.Visible = False
2870          .cmdTier2_SelectNone.Enabled = False
2880          .cmdTier2_SelectNone_raised_img_dis.Visible = True
2890          .cmdTier2_SelectNone_raised_img.Visible = False
2900          .cmdTier2_SelectNone_raised_semifocus_dots_img.Visible = False
2910          .cmdTier2_SelectNone_raised_focus_img.Visible = False
2920          .cmdTier2_SelectNone_raised_focus_dots_img.Visible = False
2930          .cmdTier2_SelectNone_sunken_focus_dots_img.Visible = False
2940        End If
2950      End Select
2960    End With

EXITP:
2970    Exit Sub

ERRH:
2980    Select Case ERR.Number
        Case Else
2990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3000    End Select
3010    Resume EXITP

End Sub

Public Sub Tier3_Enable_AE(blnAble As Boolean, blnNoData As Boolean, frm As Access.Form)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Tier3_Enable_AE"

3110    With frm
3120      Select Case blnAble
          Case True
3130        .ckgTier_chkTier3.Enabled = True
3140        .ckgTier_chkTier3_lbl2.ForeColor = CLR_VDKGRY
3150        .ckgTier_chkTier3_lbl2_dim_hi.Visible = False
3160        Select Case .ckgTier_chkTier3
            Case True
3170          .ckgTier_chkTier3_box.BackStyle = acBackStyleNormal
3180          .ckgTier_chkTier3_box2.BackStyle = acBackStyleNormal
3190          .ckgTier_chkTier3_hline03.BorderColor = MY_CLR_VLTBGE
3200          .ckgTier3_chk01_AdminOfficer.Enabled = True
3210          .ckgTier3_chk02_PortfolioModel.Enabled = True
3220          .ckgTier3_chk03_Schedule.Enabled = True
3230          .ckgTier3_chk04_Schedule_Detail.Enabled = True
3240          .ckgTier3_chk05_CheckPOSPay.Enabled = True
3250          .ckgTier3_chk06_CheckPOSPay_Detail.Enabled = True
3260          .ckgTier3_chk07_CheckReconcile_Amount.Enabled = True
3270          .ckgTier3_chk08_CheckReconcile_Item.Enabled = True
3280          .ckgTier3_chk09_CheckMemo.Enabled = True
3290          .ckgTier3_chk01_AdminOfficer_lbl_recs.ForeColor = CLR_DKGRY
3300          .ckgTier3_chk01_AdminOfficer_lbl_recs.BorderColor = CLR_VLTBLU2
3310          .ckgTier3_chk02_PortfolioModel_lbl_recs.ForeColor = CLR_DKGRY
3320          .ckgTier3_chk02_PortfolioModel_lbl_recs.BorderColor = CLR_VLTBLU2
3330          .ckgTier3_chk03_Schedule_lbl_recs.ForeColor = CLR_DKGRY
3340          .ckgTier3_chk03_Schedule_lbl_recs.BorderColor = CLR_VLTBLU2
3350          .ckgTier3_chk04_Schedule_Detail_lbl_recs.ForeColor = CLR_DKGRY
3360          .ckgTier3_chk04_Schedule_Detail_lbl_recs.BorderColor = CLR_VLTBLU2
3370          .ckgTier3_chk05_CheckPOSPay_lbl_recs.ForeColor = CLR_DKGRY
3380          .ckgTier3_chk05_CheckPOSPay_lbl_recs.BorderColor = CLR_VLTBLU2
3390          .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.ForeColor = CLR_DKGRY
3400          .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.BorderColor = CLR_VLTBLU2
3410          .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.ForeColor = CLR_DKGRY
3420          .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.BorderColor = CLR_VLTBLU2
3430          .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.ForeColor = CLR_DKGRY
3440          .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.BorderColor = CLR_VLTBLU2
3450          .ckgTier3_chk09_CheckMemo_lbl_recs.ForeColor = CLR_DKGRY
3460          .ckgTier3_chk09_CheckMemo_lbl_recs.BorderColor = CLR_VLTBLU2
3470          .cmdTier3_Select_box.BackStyle = acBackStyleNormal
3480          .cmdTier3_Select_hline03.BorderColor = MY_CLR_VLTBGE
3490          .cmdTier3_Select_hline04.BorderColor = MY_CLR_VLTBGE
3500          If blnNoData = False Then
3510            .cmdTier3_SelectAll.Enabled = True
3520            .cmdTier3_SelectAll_raised_img.Visible = True
3530            .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
3540            .cmdTier3_SelectAll_raised_focus_img.Visible = False
3550            .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
3560            .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
3570            .cmdTier3_SelectAll_raised_img_dis.Visible = False
3580            .cmdTier3_SelectNone.Enabled = True
3590            .cmdTier3_SelectNone_raised_img.Visible = True
3600            .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
3610            .cmdTier3_SelectNone_raised_focus_img.Visible = False
3620            .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
3630            .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
3640            .cmdTier3_SelectNone_raised_img_dis.Visible = False
3650          End If
3660        Case False
3670          .ckgTier_chkTier3_box.BackStyle = acBackStyleTransparent
3680          .ckgTier_chkTier3_box2.BackStyle = acBackStyleTransparent
3690          .ckgTier_chkTier3_hline03.BorderColor = MY_CLR_LTBGE
3700          .ckgTier3_chk01_AdminOfficer.Enabled = False
3710          .ckgTier3_chk02_PortfolioModel.Enabled = False
3720          .ckgTier3_chk03_Schedule.Enabled = False
3730          .ckgTier3_chk04_Schedule_Detail.Enabled = False
3740          .ckgTier3_chk05_CheckPOSPay.Enabled = False
3750          .ckgTier3_chk06_CheckPOSPay_Detail.Enabled = False
3760          .ckgTier3_chk07_CheckReconcile_Amount.Enabled = False
3770          .ckgTier3_chk08_CheckReconcile_Item.Enabled = False
3780          .ckgTier3_chk09_CheckMemo.Enabled = False
3790          .ckgTier3_chk01_AdminOfficer_lbl_recs.ForeColor = WIN_CLR_DISF
3800          .ckgTier3_chk01_AdminOfficer_lbl_recs.BorderColor = WIN_CLR_DISR
3810          .ckgTier3_chk02_PortfolioModel_lbl_recs.ForeColor = WIN_CLR_DISF
3820          .ckgTier3_chk02_PortfolioModel_lbl_recs.BorderColor = WIN_CLR_DISR
3830          .ckgTier3_chk03_Schedule_lbl_recs.ForeColor = WIN_CLR_DISF
3840          .ckgTier3_chk03_Schedule_lbl_recs.BorderColor = WIN_CLR_DISR
3850          .ckgTier3_chk04_Schedule_Detail_lbl_recs.ForeColor = WIN_CLR_DISF
3860          .ckgTier3_chk04_Schedule_Detail_lbl_recs.BorderColor = WIN_CLR_DISR
3870          .ckgTier3_chk05_CheckPOSPay_lbl_recs.ForeColor = WIN_CLR_DISF
3880          .ckgTier3_chk05_CheckPOSPay_lbl_recs.BorderColor = WIN_CLR_DISR
3890          .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.ForeColor = WIN_CLR_DISF
3900          .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.BorderColor = WIN_CLR_DISR
3910          .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.ForeColor = WIN_CLR_DISF
3920          .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.BorderColor = WIN_CLR_DISR
3930          .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.ForeColor = WIN_CLR_DISF
3940          .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.BorderColor = WIN_CLR_DISR
3950          .ckgTier3_chk09_CheckMemo_lbl_recs.ForeColor = WIN_CLR_DISF
3960          .ckgTier3_chk09_CheckMemo_lbl_recs.BorderColor = WIN_CLR_DISR
3970          .cmdTier3_Select_box.BackStyle = acBackStyleTransparent
3980          .cmdTier3_Select_hline03.BorderColor = MY_CLR_LTBGE
3990          .cmdTier3_Select_hline04.BorderColor = MY_CLR_LTBGE
4000          If blnNoData = False Then
4010            .cmdTier3_SelectAll.Enabled = False
4020            .cmdTier3_SelectAll_raised_img_dis.Visible = True
4030            .cmdTier3_SelectAll_raised_img.Visible = False
4040            .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
4050            .cmdTier3_SelectAll_raised_focus_img.Visible = False
4060            .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
4070            .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
4080            .cmdTier3_SelectNone.Enabled = False
4090            .cmdTier3_SelectNone_raised_img_dis.Visible = True
4100            .cmdTier3_SelectNone_raised_img.Visible = False
4110            .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
4120            .cmdTier3_SelectNone_raised_focus_img.Visible = False
4130            .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
4140            .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
4150          End If
4160        End Select
4170      Case False
4180        .ckgTier_chkTier3.Enabled = False
4190        .ckgTier_chkTier3_lbl2.ForeColor = WIN_CLR_DISF
4200        If blnNoData = False Then
4210          .ckgTier_chkTier3_lbl2_dim_hi.Visible = True
4220        End If
4230        .ckgTier_chkTier3_box.BackStyle = acBackStyleTransparent
4240        .ckgTier_chkTier3_box2.BackStyle = acBackStyleTransparent
4250        .ckgTier_chkTier3_hline03.BorderColor = MY_CLR_LTBGE
4260        .ckgTier3_chk01_AdminOfficer.Enabled = False
4270        .ckgTier3_chk02_PortfolioModel.Enabled = False
4280        .ckgTier3_chk03_Schedule.Enabled = False
4290        .ckgTier3_chk04_Schedule_Detail.Enabled = False
4300        .ckgTier3_chk05_CheckPOSPay.Enabled = False
4310        .ckgTier3_chk06_CheckPOSPay_Detail.Enabled = False
4320        .ckgTier3_chk07_CheckReconcile_Amount.Enabled = False
4330        .ckgTier3_chk08_CheckReconcile_Item.Enabled = False
4340        .ckgTier3_chk09_CheckMemo.Enabled = False
4350        .ckgTier3_chk01_AdminOfficer_lbl_recs.ForeColor = WIN_CLR_DISF
4360        .ckgTier3_chk01_AdminOfficer_lbl_recs.BorderColor = WIN_CLR_DISR
4370        .ckgTier3_chk02_PortfolioModel_lbl_recs.ForeColor = WIN_CLR_DISF
4380        .ckgTier3_chk02_PortfolioModel_lbl_recs.BorderColor = WIN_CLR_DISR
4390        .ckgTier3_chk03_Schedule_lbl_recs.ForeColor = WIN_CLR_DISF
4400        .ckgTier3_chk03_Schedule_lbl_recs.BorderColor = WIN_CLR_DISR
4410        .ckgTier3_chk04_Schedule_Detail_lbl_recs.ForeColor = WIN_CLR_DISF
4420        .ckgTier3_chk04_Schedule_Detail_lbl_recs.BorderColor = WIN_CLR_DISR
4430        .ckgTier3_chk05_CheckPOSPay_lbl_recs.ForeColor = WIN_CLR_DISF
4440        .ckgTier3_chk05_CheckPOSPay_lbl_recs.BorderColor = WIN_CLR_DISR
4450        .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.ForeColor = WIN_CLR_DISF
4460        .ckgTier3_chk06_CheckPOSPay_Detail_lbl_recs.BorderColor = WIN_CLR_DISR
4470        .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.ForeColor = WIN_CLR_DISF
4480        .ckgTier3_chk07_CheckReconcile_Amount_lbl_recs.BorderColor = WIN_CLR_DISR
4490        .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.ForeColor = WIN_CLR_DISF
4500        .ckgTier3_chk08_CheckReconcile_Item_lbl_recs.BorderColor = WIN_CLR_DISR
4510        .ckgTier3_chk09_CheckMemo_lbl_recs.ForeColor = WIN_CLR_DISF
4520        .ckgTier3_chk09_CheckMemo_lbl_recs.BorderColor = WIN_CLR_DISR
4530        .cmdTier3_Select_box.BackStyle = acBackStyleTransparent
4540        .cmdTier3_Select_hline03.BorderColor = MY_CLR_LTBGE
4550        .cmdTier3_Select_hline04.BorderColor = MY_CLR_LTBGE
4560        If blnNoData = False Then
4570          .cmdTier3_SelectAll.Enabled = False
4580          .cmdTier3_SelectAll_raised_img_dis.Visible = True
4590          .cmdTier3_SelectAll_raised_img.Visible = False
4600          .cmdTier3_SelectAll_raised_semifocus_dots_img.Visible = False
4610          .cmdTier3_SelectAll_raised_focus_img.Visible = False
4620          .cmdTier3_SelectAll_raised_focus_dots_img.Visible = False
4630          .cmdTier3_SelectAll_sunken_focus_dots_img.Visible = False
4640          .cmdTier3_SelectNone.Enabled = False
4650          .cmdTier3_SelectNone_raised_img_dis.Visible = True
4660          .cmdTier3_SelectNone_raised_img.Visible = False
4670          .cmdTier3_SelectNone_raised_semifocus_dots_img.Visible = False
4680          .cmdTier3_SelectNone_raised_focus_img.Visible = False
4690          .cmdTier3_SelectNone_raised_focus_dots_img.Visible = False
4700          .cmdTier3_SelectNone_sunken_focus_dots_img.Visible = False
4710        End If
4720      End Select
4730    End With

EXITP:
4740    Exit Sub

ERRH:
4750    Select Case ERR.Number
        Case Else
4760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4770    End Select
4780    Resume EXITP

End Sub
