Attribute VB_Name = "modJrnlCol_Timer"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Timer"

'VGC 09/09/2017: CHANGES!

Private Const FRM_PAR As String = "frmJournal_Columns"

' ** Combo box column constants: assetno.
Private Const CBX_AST_DESC  As Integer = 1  'totdesc
' **

Public Sub Timer1_JC(blnGoingToReport As Boolean, blnGTR_NoAdd As Boolean, blnGTR_Emblem As Boolean, blnGoneToReport As Boolean, blnGoneToReport2 As Boolean, lngGTR_Stat As Long, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Timer1_JC"

        Dim intPos01 As Integer

        Static strAccountNo As String, lngAssetNo As Long

110     With frm

120       If gblnGoToReport = True Then
130         If IsNull(garr_varGoToReport(GTR_CTL2)) = False Then
140           If garr_varGoToReport(GTR_CTL2) = "cost" Then
                ' ** The emblem on frmJournal_Columns should be still spinning.
150             Select Case blnGoingToReport
                Case True
160               lngGTR_Stat = lngGTR_Stat + 1&
170               Select Case lngGTR_Stat
                  Case 1&
180                 .accountno.SetFocus
190                 .GoToReport_arw_jcol_blu2_img.Visible = True
200                 .GoToReport_arw_jcol_blu1_img.Visible = False
210                 .GoToReport_arw_jcol_bge_img.Visible = False
220                 .accountno = strAccountNo
                    ' ** I can't get the normal highlighter to show, so this is GoToReport's own.
230                 DoEvents
240                 If IsNull(.JrnlCol_ID) = False Then
250                   .GoToReport_fld_jcol_id = .JrnlCol_ID
260                   .GoToReports_fld_bg1.Requery
270                   .GoToReports_fld_bg2.Requery
280                   .GoToReports_fld_bg1.Visible = True
290                   .GoToReports_fld_bg2.Visible = True
300                   DoEvents
310                 End If
320                 DoEvents
330                 .TimerInterval = 50&
340               Case 12&
350                 .accountno.SelLength = 0
360                 .accountno.SelStart = 19
370                 DoEvents
380                 .TimerInterval = 50&
390               Case 24&
400                 blnGTR_NoAdd = True
410                 .GoToReport_arw_jcol_bge_img.Left = .journaltype.Left
420                 .GoToReport_arw_jcol_bge_img.Visible = True
430                 .GoToReport_arw_jcol_blu1_img.Visible = False
440                 .GoToReport_arw_jcol_blu2_img.Visible = False
450                 .GoToReport_arw_jcol_blu1_img.Left = .GoToReport_arw_jcol_bge_img.Left
460                 .GoToReport_arw_jcol_blu2_img.Left = .GoToReport_arw_jcol_bge_img.Left
470                 DoEvents
480                 .TimerInterval = 50&
490               Case 36&
500                 .accountno_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
510                 .GoToReport_arw_jcol_blu1_img.Visible = False
520                 .GoToReport_arw_jcol_blu2_img.Visible = True
530                 .GoToReport_arw_jcol_bge_img.Visible = False
540                 DoEvents
550                 .TimerInterval = 50&
560               Case 48&
570                 .journaltype = "Sold"
580                 DoEvents
590                 .TimerInterval = 50&
600               Case 60&
610                 .journaltype.SetFocus
620                 .journaltype.SelLength = 0
630                 .journaltype.SelStart = 19
640                 DoEvents
650                 .TimerInterval = 50&
660               Case 72&
670                 .GoToReport_arw_jcol_blu1_img.Visible = False
680                 .GoToReport_arw_jcol_blu2_img.Visible = False
690                 .assetno_lbl.Caption = "      " & .assetno_lbl.Caption  ' ** Move the caption aside.
700                 .GoToReport_arw_jcol_bge_img.Left = .assetno.Left
710                 .GoToReport_arw_jcol_bge_img.Visible = True
720                 DoEvents
730                 .GoToReport_arw_jcol_blu1_img.Left = .GoToReport_arw_jcol_bge_img.Left
740                 .GoToReport_arw_jcol_blu2_img.Left = .GoToReport_arw_jcol_bge_img.Left
750                 DoEvents
760                 .TimerInterval = 50&
770               Case 84&
780                 .journaltype_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
790                 .GoToReport_arw_jcol_blu1_img.Visible = True
800                 .GoToReport_arw_jcol_blu2_img.Visible = False
810                 .GoToReport_arw_jcol_bge_img.Visible = False
820                 DoEvents
830                 .TimerInterval = 50&
840               Case 96&
850                 .assetno = lngAssetNo
860                 DoEvents
870                 .TimerInterval = 50&
880               Case 108&
890                 .assetno.SetFocus
900                 .assetno.SelLength = 0
910                 .assetno.SelStart = 0
920                 DoEvents
930                 .TimerInterval = 50
940               Case 120&
950                 .GoToReport_arw_jcol_blu1_img.Visible = False
960                 .GoToReport_arw_jcol_blu2_img.Visible = False
970                 .assetno_lbl.Caption = Trim(.assetno_lbl.Caption)
980                 .GoToReport_arw_jcol_bge_img.Left = .assetdate_display.Left
990                 .GoToReport_arw_jcol_bge_img.Visible = True
1000                .GoToReport_arw_jcol_blu1_img.Left = .GoToReport_arw_jcol_bge_img.Left
1010                .GoToReport_arw_jcol_blu2_img.Left = .GoToReport_arw_jcol_bge_img.Left
1020                DoEvents
1030                .TimerInterval = 50&
1040              Case 132&
1050                .assetno_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
1060                .GoToReport_arw_jcol_blu1_img.Visible = False
1070                .GoToReport_arw_jcol_blu2_img.Visible = True
1080                .GoToReport_arw_jcol_bge_img.Visible = False
1090                DoEvents
1100                .TimerInterval = 50&
1110              Case 144&
1120                .assetdate_display = Date
1130                DoEvents
1140                .TimerInterval = 50&
1150              Case 156&
1160                .assetdate_display.SetFocus
1170                .assetdate_display.SelLength = 0
1180                .assetdate_display.SelStart = 19
1190                If IsNull(.assetno_description) = True Then
1200                  .assetno_description = .assetno.Column(CBX_AST_DESC)
1210                Else
1220                  If Trim(.assetno_description) = vbNullString Then
1230                    .assetno_description = .assetno.Column(CBX_AST_DESC)
1240                  End If
1250                End If
1260                DoEvents
1270                .TimerInterval = 50&
1280              Case 168&
1290                .GoToReport_arw_jcol_blu1_img.Visible = False
1300                .GoToReport_arw_jcol_blu2_img.Visible = False
1310                .GoToReport_arw_jcol_bge_img.Left = .shareface.Left
1320                .GoToReport_arw_jcol_bge_img.Visible = True
1330                .GoToReport_arw_jcol_blu1_img.Left = .GoToReport_arw_jcol_bge_img.Left
1340                .GoToReport_arw_jcol_blu2_img.Left = .GoToReport_arw_jcol_bge_img.Left
1350                DoEvents
1360                .TimerInterval = 50&
1370              Case 180&
1380                .assetdate_display_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
1390                .GoToReport_arw_jcol_blu1_img.Visible = False
1400                .GoToReport_arw_jcol_blu2_img.Visible = True
1410                .GoToReport_arw_jcol_bge_img.Visible = False
1420                DoEvents
1430                .TimerInterval = 50&
1440              Case 192&
1450                .shareface = 10
1460                DoEvents
1470                .TimerInterval = 50&
1480              Case 204&
1490                .shareface.SetFocus
1500                .shareface.SelLength = 0
1510                .shareface.SelStart = 19
1520                DoEvents
1530                .TimerInterval = 50&
1540              Case 216&
1550                .GoToReport_arw_jcol_bge_img.Left = .PCash.Left
1560                .GoToReport_arw_jcol_bge_img.Visible = True
1570                .GoToReport_arw_jcol_blu1_img.Visible = False
1580                .GoToReport_arw_jcol_blu2_img.Visible = False
1590                .GoToReport_arw_jcol_blu1_img.Left = .GoToReport_arw_jcol_bge_img.Left
1600                .GoToReport_arw_jcol_blu2_img.Left = .GoToReport_arw_jcol_bge_img.Left
1610                .shareface_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
1620                DoEvents
1630                .TimerInterval = 50&
1640              Case 240&
1650                .GoToReport_arw_jcol_blu1_img.Visible = False
1660                .GoToReport_arw_jcol_blu2_img.Visible = True
1670                .GoToReport_arw_jcol_bge_img.Visible = False
1680                .PCash.SetFocus
1690                .PCash = 100
1700                DoEvents
1710                .TimerInterval = 50&
1720              Case 252&
1730                .PCash.SetFocus
1740                .PCash.SelLength = 0
1750                .PCash.SelStart = 19
1760                DoEvents
1770                .TimerInterval = 50&
1780              Case 264&
1790                .GoToReport_arw_jcol_bge_img.Left = .Cost.Left
1800                .GoToReport_arw_jcol_bge_img.Visible = True
1810                .GoToReport_arw_jcol_blu1_img.Visible = False
1820                .GoToReport_arw_jcol_blu2_img.Visible = False
1830                .GoToReport_arw_jcol_blu1_img.Left = .GoToReport_arw_jcol_bge_img.Left
1840                .GoToReport_arw_jcol_blu2_img.Left = .GoToReport_arw_jcol_bge_img.Left
1850                .pcash_AfterUpdate  ' ** Form Procedure: frmJournal_Columns_Sub.
1860                DoEvents
1870  On Error Resume Next
1880                .Cost.SetFocus  ' ** This is to set off the pcash_Exit() event.
1890  On Error GoTo ERRH
1900                .TimerInterval = 50&
1910              Case Else
                    ' ** Just let the emblem go.
                    ' ** lngGTR_Stat will NOT match frmJournal_Columns'!
1920                If lngGTR_Stat < 288& Then
1930                  .TimerInterval = 50&
1940                Else
1950                  .TimerInterval = 0&
1960                  DoCmd.Hourglass False
1970                  blnGoingToReport = False
1980                  gblnGoToReport = False
1990                  blnGTR_Emblem = False
2000                  blnGTR_NoAdd = False
2010                  Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
2020                End If
2030              End Select
2040            Case False
2050              If blnGTR_Emblem = False Then
2060                blnGTR_Emblem = .Parent.GTRStuff(2, False)  ' ** Form Function: frmJournal_Columns.
2070              End If
2080              strAccountNo = GetSoldAsset  ' ** Module Function: modGoToReportFuncs.
2090              If strAccountNo <> vbNullString Then
2100                blnGoingToReport = True
2110                intPos01 = InStr(strAccountNo, ";")
2120                lngAssetNo = Val(Mid(strAccountNo, (intPos01 + 1)))
2130                strAccountNo = Left(strAccountNo, (intPos01 - 1))
2140                .TimerInterval = 50&
2150              Else
2160                .TimerInterval = 0&
2170                Forms(FRM_PAR).GTRStuff 1, False  ' ** Form Procedure: frmJournal_Columns.
2180                blnGoingToReport = False
2190                gblnGoToReport = False
2200                blnGTR_Emblem = False
2210                blnGTR_NoAdd = False
2220                .GoToReport_arw_jcol_bge_img.Visible = False
2230                .GoToReport_arw_jcol_blu1_img.Visible = False
2240                .GoToReport_arw_jcol_blu2_img.Visible = False
2250                Beep
2260                DoCmd.Hourglass False
2270                MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There are insufficient asset holdings to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
2280              End If
2290            End Select
2300          End If
2310        End If
2320      ElseIf blnGoneToReport = True Then
            ' ** This has become an unholy mess!
2330      ElseIf blnGoneToReport2 = True Then
            ' ** This has become an unholy mess!
2340        blnGoneToReport2 = False
2350        .Parent.cmdDelete.SetFocus
2360        .DelRec_Send  ' ** Module Function: modGoToReportFuncs.
2370        DoEvents
2380        DoCmd.SelectObject acForm, .Parent.Name, False
2390        .Parent.cmdClose.SetFocus
2400      End If

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

Public Sub Timer2_JC(blnJTypeSet As Boolean, blnReinvestment As Boolean, strSaveMoveCtl As String, frm As Access.Form)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Timer2_JC"

        Dim strThisJType As String
        Dim strTmp01 As String
        Dim lngRetVal As Long

2510    With frm

          ' ** Because these came from JC_Key_Sub_Next(), they've already been vetted.
          ' ** Setting locks is only of concern if this came from journaltype,
          ' ** for any other source, just move there.
2520      Select Case blnJTypeSet
          Case True

2530        DoCmd.Hourglass True
2540        DoEvents

2550        If IsNull(.journaltype) = False Then

2560          blnJTypeSet = False
2570          strThisJType = .journaltype

              ' ** Lock and unlock fields appropriate to JournalType.
2580          JC_Key_JType_Set strThisJType, frm  ' ** Module Function: modJrnlCol_Keys.
2590          DoEvents

              ' ** Initialize default values.
2600          If IsNull(.Journal_ID) = True Then .Journal_ID = 0&
2610          .Recur_Name = Null
2620          .RecurringItem_ID = 0&
2630          .shareface = 0#
2640          Select Case blnReinvestment
              Case True
                'blnReinvestment = False  'WAIT TILL AFTER THE LAST SAVE!
2650          Case False
2660            .ICash = 0@
2670            .PCash = 0@
2680            .Cost = 0@
2690          End Select
2700          .pershare = 0#
2710          .rate = 0#
2720          .IsAverage = False
2730          .PrintCheck = False
2740          .Location_ID = 1&
2750          .Loc_Name = "{no entry}"
2760          .Loc_Name_display = vbNullString  ' ** These do allow zero-length strings.
2770          .revcode_DESC_display = Null
2780          .taxcode_description_display = Null

2790          .chkShowAllAssets.Visible = False
2800          Select Case strThisJType
              Case "Dividend", "Interest"
2810            .chkShowAllAssets.Visible = True
2820            .chkShowAllAssets.Locked = False  ' ** Doesn't stick!
2830          Case "Purchase", "Deposit", "Liability (+)", "Sold", "Withdrawn", "Liability (-)", "Received", "Misc.", "Paid"
                ' ** Leave it off.
2840          End Select

              ' ** Set the AssetNo RowSource.
2850          JC_Msc_Asset_Set frm  ' ** Module Procedure: modJrnlCol_Misc.

2860          Select Case strThisJType
              Case "Dividend", "Interest"
                ' ** REVID_INC  ' ** Unspecified Income.
2870            .revcode_ID = REVID_INC  ' ** Unspecified Income.
2880            .revcode_DESC = "Unspecified Income"
2890            .revcode_TYPE = REVTYP_INC
2900            .taxcode = TAXID_INC  ' ** Unspecified Income.
2910            .taxcode_description = "Unspecified Income"
2920            .taxcode_type = TAXTYP_INC
2930          Case "Purchase", "Deposit"
                ' ** REVID_INC  ' ** Unspecified Income.
2940            .revcode_ID = REVID_INC  ' ** Unspecified Income.
2950            .revcode_DESC = "Unspecified Income"
2960            .revcode_TYPE = REVTYP_INC
2970            .taxcode = TAXID_INC  ' ** Unspecified Income.
2980            .taxcode_description = "Unspecified Income"
2990            .taxcode_type = TAXTYP_INC
3000          Case "Sold"
                ' ** REVID_INC  ' ** Unspecified Income.
3010            .revcode_ID = REVID_INC  ' ** Unspecified Income.
3020            .revcode_DESC = "Unspecified Income"
3030            .revcode_TYPE = REVTYP_INC
3040            .taxcode = TAXID_INC  ' ** Unspecified Income.
3050            .taxcode_description = "Unspecified Income"
3060            .taxcode_type = TAXTYP_INC
3070          Case "Liability (+)", "Liability (-)"
                ' ** REVID_EXP  ' ** Unspecified Expense.
3080            .revcode_ID = REVID_EXP  ' ** Unspecified Expense.
3090            .revcode_DESC = "Unspecified Expense"
3100            .revcode_TYPE = REVTYP_EXP
3110            .taxcode = TAXID_DED  ' ** Unspecified Deduction.
3120            .taxcode_description = "Unspecified Deduction"
3130            .taxcode_type = TAXTYP_DED
3140          Case "Withdrawn", "Cost Adj."
                ' ** EITHER (We'll start with INCOME.)
3150            .revcode_ID = REVID_INC  ' ** Unspecified Income.
3160            .revcode_DESC = "Unspecified Income"
3170            .revcode_TYPE = REVTYP_INC
3180            .taxcode = TAXID_INC  ' ** Unspecified Income.
3190            .taxcode_description = "Unspecified Income"
3200            .taxcode_type = TAXTYP_INC
3210          Case "Paid"
                ' ** REVID_EXP  ' ** Unspecified Expense.
3220            .revcode_ID = REVID_EXP  ' ** Unspecified Expense.
3230            .revcode_DESC = "Unspecified Expense"
3240            .revcode_TYPE = REVTYP_EXP
3250            .taxcode = TAXID_DED  ' ** Unspecified Deduction.
3260            .taxcode_description = "Unspecified Deduction"
3270            .taxcode_type = TAXTYP_DED
3280          Case "Received"
                ' ** REVID_INC  ' ** Unspecified Income.
3290            .revcode_ID = REVID_INC  ' ** Unspecified Income.
3300            .revcode_DESC = "Unspecified Income"
3310            .revcode_TYPE = REVTYP_INC
3320            .taxcode = TAXID_INC  ' ** Unspecified Income.
3330            .taxcode_description = "Unspecified Income"
3340            .taxcode_type = TAXTYP_INC
3350          Case "Misc."
                ' ** EITHER (We'll start with INCOME.)
3360            .revcode_ID = REVID_INC  ' ** Unspecified Income.
3370            .revcode_DESC = "Unspecified Income"
3380            .revcode_TYPE = REVTYP_INC
3390            .taxcode = TAXID_INC  ' ** Unspecified Income.
3400            .taxcode_description = "Unspecified Income"
3410            .taxcode_type = TAXTYP_INC
3420          End Select
              ' ** Linked RevCode/TaxCode covered by defaults, above.

              ' ** Set the RevCode_ID RowSource.
3430          JC_Msc_RevCode_Set frm  ' ** Module Procedure: modJrnlCol_Misc.
              ' ** Set the TaxCode RowSource.
3440          JC_Msc_TaxCode_Set frm  ' ** Module Procedure: modJrnlCol_Misc.

3450          strSaveMoveCtl = vbNullString
3460          .cmdSave_Click  ' ** Form Procedure: frmJournal_Columns_Sub.

              ' ** Set the button status on frmJournal_Columns.
3470          JC_Btn_Set strThisJType, .posted, frm.Parent  ' ** Module Procedure: modJrnlCol_Buttons.

3480  On Error Resume Next
3490          .Controls(strTmp01).SetFocus
3500          If ERR.Number <> 0 Then
3510  On Error GoTo ERRH
3520            Select Case strTmp01
                Case "assetno_description"
                  ' ** 2110:  Microsoft Access can't move the focus to the control revcode_DESC_display.
3530              .assetno.SetFocus  ' ** Let's try this.
3540            Case "Loc_Name_display"
3550              .Location_ID.SetFocus
3560            Case "revcode_DESC_display"
3570              .revcode_ID.SetFocus
3580            Case "taxcode_description_display"
3590              .taxcode.SetFocus
3600            End Select
3610          Else
3620  On Error GoTo ERRH
3630          End If

3640        End If

3650        DoCmd.Hourglass False

3660      Case False
3670        strThisJType = Nz(.journaltype, vbNullString)
3680        Select Case strTmp01
            Case "shareface"
3690          Select Case strThisJType
              Case "Deposit", "Purchase", "Withdrawn", "Sold", "Liability (+)", "Liability (-)", "Received"
3700            lngRetVal = fSetScrollBarPosHZ(frm, 999&)  ' ** Module Function: modScrollBarFuncs.
3710          Case Else
                ' ** Not sure what else.
3720          End Select
3730        Case "icash"
3740          Select Case strThisJType
              Case "Dividend", "Interest", "Misc.", "Received"
3750            lngRetVal = fSetScrollBarPosHZ(frm, 999&)  ' ** Module Function: modScrollBarFuncs.
3760          Case Else
                ' ** Not sure what else.
3770          End Select
3780        Case "pcash"
3790          Select Case strThisJType
              Case "Paid"
3800            lngRetVal = fSetScrollBarPosHZ(frm, 999&)  ' ** Module Function: modScrollBarFuncs.
3810          Case Else
                ' ** Not sure what else.
3820          End Select
3830        Case "cost"
3840          Select Case strThisJType
              Case "Cost Adj."
3850            lngRetVal = fSetScrollBarPosHZ(frm, 999&)  ' ** Module Function: modScrollBarFuncs.
3860          Case Else
                ' ** Not sure what else.
3870          End Select
3880        End Select
3890  On Error Resume Next
3900        .Controls(strTmp01).SetFocus
3910        If ERR.Number <> 0 Then
3920  On Error GoTo ERRH
3930          Select Case strTmp01
              Case "assetno_description"
3940            .assetno.SetFocus
3950          Case "Loc_Name_display"
3960            .Location_ID.SetFocus
3970          Case "revcode_DESC_display"
                ' ** 2110:  Microsoft Access can't move the focus to the control revcode_DESC_display.
3980            .revcode_ID.SetFocus  ' ** Let's try this.
3990          Case "taxcode_description_display"
4000            .taxcode.SetFocus
4010          End Select
4020        Else
4030  On Error GoTo ERRH
4040        End If
4050      End Select  ' ** blnJTypeSet.

4060    End With

EXITP:
4070    Exit Sub

ERRH:
4080    Select Case ERR.Number
        Case Else
4090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4100    End Select
4110    Resume EXITP

End Sub
