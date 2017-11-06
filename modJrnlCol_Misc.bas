Attribute VB_Name = "modJrnlCol_Misc"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Misc"

'VGC 09/09/2017: CHANGES!

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!

' ** Array: arr_varStat().
Private lngStats As Long, arr_varStat() As Variant
Private Const S_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const S_SEC As Integer = 0
Private Const S_NAM As Integer = 1
Private Const S_TXT As Integer = 2

Private lngTpp As Long
' **

Public Sub JC_Msc_Cur_Set(frmSub As Access.Form, frmPar As Access.Form)
' ** Called by:
' **   frmJournal_Columns_Sub:
' **     Form_Current()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Cur_Set"

        Dim strThisJType As String
        Dim blnNewRec As Boolean

110     With frmSub

120       .chkShowAllAssets.Visible = False
130       .MiscXFer_ItoP_lbl.Visible = False
140       .MiscXFer_PtoI_lbl.Visible = False

150       gstrFormQuerySpec = frmPar.Name
160       blnNewRec = False

          ' ** Fields shouldn't be Enabled-False, Locked-False,
          ' ** because then all records will show as disabled field.
170       strThisJType = vbNullString

180       Select Case IsNull(.journaltype)
          Case True

190         If .Parent.JrnlMemo_Memo.Visible = True Then
200           JC_Msc_Memo_Set False, frmPar  ' ** Procedure: Below.
210         End If

220         .transdate.Locked = False
230         .accountno.Locked = False
240         .accountno2.Locked = False
250         .journaltype.Locked = False
260         .assetno.Locked = True
270         .Recur_Name.Locked = True
280         .assetdate.Locked = True
290         .shareface.Locked = True
300         .ICash.Locked = True
310         .PCash.Locked = True
320         .Cost.Locked = True
330         .description.Locked = True
340         .Location_ID.Locked = True
350         .PrintCheck.Locked = True

360         Select Case gblnRevenueExpenseTracking
            Case True
              ' ** revcode_DESC_display should always remain locked.
370           .revcode_ID.Locked = True
380         Case False
              ' ** As-is.
390         End Select

400         Select Case gblnIncomeTaxCoding
            Case True
              ' ** taxcode_description_display should always remain locked.
410           .taxcode.Locked = True
420         Case False
              ' ** As-is.
430         End Select

440         If IsNull(.accountno) = True Then
450           blnNewRec = True
460         End If

470       Case False

480         strThisJType = .journaltype

            ' *************************
            ' ** AssetNo, AssetDate:
            ' *************************
490         .assetdate_display.Enabled = True
500         .assetno.Enabled = True
510         Select Case strThisJType
            Case "Deposit", "Purchase", "Liability (+)", "Liability (-)", "Sold", "Withdrawn", "Cost Adj.", "Dividend", "Interest"
              ' ** AssetNo, AssetDate allowed.
520           .assetdate_display.Locked = False
530           .assetno.Locked = False
540         Case "Received"
              ' ** Allowed: Received (if there's an assetno)
              ' ** Not Allowed: Received (with no assetno)
              ' ** Since both fields must be available to make the
              ' ** choice, checking must be handled by other means.
550           If .assetno > 0& Then
560             .assetdate_display.Locked = False
570             .assetno.Locked = False
580           Else
590             .assetdate_display.Locked = True
600             .assetno.Locked = True
610           End If
620         Case Else
              ' ** AssetNo, AssetDate not allowed.
              ' ** Misc., Paid.
630           .assetdate_display.Locked = True
640           .assetno.Locked = True
650         End Select

660         Select Case strThisJType
            Case "Dividend", "Interest"
670           If .posted = False Then
680             .chkShowAllAssets.Visible = True
690           End If
700         Case Else
              ' ** Leave it off.
710         End Select
720         gstrFormQuerySpec = frmPar.Name  ' ** This gets unset somewhere!
730         JC_Msc_Asset_Set frmSub  ' ** Procedure: Below.

            ' *************************
            ' ** Recurring Item:
            ' *************************
740         .Recur_Name.Enabled = True
750         Select Case strThisJType
            Case "Misc."
              ' ** RecurringItem allowed.
760           If .Recur_Name.RowSource <> "qryJournal_Columns_10_RecurringItems_02" Then
770             .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_02"
780           End If
790           .Recur_Name.Requery
800           .Recur_Name.Locked = False
              '.RecurringItem_ID
810         Case "Paid"
              ' ** RecurringItem allowed.
820           If .Recur_Name.RowSource <> "qryJournal_Columns_10_RecurringItems_03" Then
830             .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_03"
840           End If
850           .Recur_Name.Requery
860           .Recur_Name.Locked = False
              '.RecurringItem_ID
870         Case "Received"
880           If .assetno > 0& Then  ' ** Only created by LTCG Map.
                ' ** RecurringItem not allowed.
890             .Recur_Name.Locked = True
900           Else
                ' ** RecurringItem allowed.
910             If .Recur_Name.RowSource <> "qryJournal_Columns_10_RecurringItems_04" Then
920               .Recur_Name.RowSource = "qryJournal_Columns_10_RecurringItems_04"
930             End If
940             .Recur_Name.Requery
950             .Recur_Name.Locked = False
960           End If
970         Case Else
980           .Recur_Name.Locked = True
              '.RecurringItem_ID
990         End Select

            ' *************************
            ' ** ShareFace:
            ' *************************
1000        .shareface.Enabled = True
1010        Select Case strThisJType
            Case "Deposit", "Purchase", "Liability (+)", "Liability (-)", "Sold", "Withdrawn", "Dividend", "Interest"
              ' ** ShareFace allowed.
1020          .shareface.Locked = False
1030        Case "Received"
              ' ** Allowed: Received (if there's an assetno)
              ' ** Not Allowed: Received (with no assetno)
1040          If .assetno > 0& Then  ' ** Only created by LTCG Map.
1050            .shareface.Locked = False
1060          Else
1070            .shareface.Locked = True
1080          End If
1090        Case Else
              ' ** ShareFace not allowed.
              ' **   Cost Adj., Misc., Paid.
1100          .shareface.Locked = True
1110        End Select

            ' *************************
            ' ** ICash:
            ' *************************
1120        .ICash.Enabled = True
1130        Select Case strThisJType
            Case "Dividend", "Interest"
              ' ** ICash required.
1140          .ICash.Locked = False
1150        Case "Misc.", "Paid", "Received", "Purchase", "Sold"
              ' ** ICash allowed.
1160          .ICash.Locked = False
1170        Case "Cost Adj.", "Deposit", "Withdrawn"
              ' ** ICash not allowed.
1180          .ICash.Locked = True
1190        Case "Liability (+)"
              ' ** ICash not allowed.
1200          .ICash.Locked = True
1210        Case "Liability (-)"
              ' ** ICash allowed.
1220          .ICash.Locked = False
1230        End Select

            ' *************************
            ' ** PCash:
            ' *************************
1240        .PCash.Enabled = True
1250        Select Case strThisJType
            Case "Misc.", "Paid", "Purchase", "Received", "Sold", "Liability (+)", "Liability (-)"
              ' ** PCash allowed.
1260          .PCash.Locked = False
1270        Case "Cost Adj.", "Deposit", "Dividend", "Interest", "Withdrawn"
              ' ** PCash not allowed.
1280          .PCash.Locked = True
1290        Case Else
              ' ** PCash required.
              ' ** None!
1300        End Select

            ' *************************
            ' ** Cost:
            ' *************************
1310        .Cost.Enabled = True
1320        Select Case strThisJType
            Case "Cost Adj.", "Liability (+)", "Liability (-)", "Purchase", "Sold"  ' ** (Though there's 1 Sold in the Demo data without a Cost.)
              ' ** Cost required.
1330          .Cost.Locked = False
1340        Case "Deposit", "Withdrawn"
              ' ** Cost allowed.
1350          .Cost.Locked = False
1360        Case "Dividend", "Interest", "Misc.", "Paid", "Received"
              ' ** Cost not allowed.
1370          .Cost.Locked = True
1380        End Select

            ' *************************
            ' ** Location:
            ' *************************
1390        Select Case strThisJType
            Case "Purchase", "Deposit", "Liability (+)"
1400          .Location_ID.Locked = False
1410        Case Else
1420          .Location_ID.Locked = True
1430        End Select

            ' *************************
            ' ** Print Check:
            ' *************************
1440        Select Case strThisJType
            Case "Paid"
1450          .PrintCheck.Locked = False
1460          Select Case .PrintCheck
              Case True
1470            JC_Msc_Memo_Set .PrintCheck, frmPar  ' ** Procedure: Below.
1480            If IsNull(.JrnlMemo_Memo) = False Then
1490              frmPar.JrnlMemo_Memo = .JrnlMemo_Memo
1500            End If
1510          Case False
1520            frmPar.JrnlMemo_Memo = Null
1530            JC_Msc_Memo_Set .PrintCheck, frmPar  ' ** Procedure: Below.
1540          End Select
1550        Case "Dividend", "Interest", "Purchase", "Deposit", "Sold", "Withdrawn", "Liability (+)", "Liability (-)", _
                "Cost Adj.", "Misc.", "Received"
1560          .PrintCheck.Locked = True
1570          frmPar.JrnlMemo_Memo = Null
1580          JC_Msc_Memo_Set False, frmPar  ' ** Procedure: Below.
1590        End Select

1600        gstrFormQuerySpec = frmPar.Name  ' ** This gets unset somewhere!

            ' *************************
            ' ** Inc./Exp. Code:
            ' *************************
1610        Select Case gblnRevenueExpenseTracking
            Case True
1620          JC_Msc_RevCode_Set frmSub  ' ** Procedure: Below.
1630          .revcode_ID.Locked = False
1640        Case False
              ' ** Disabled.
1650        End Select

1660        gstrFormQuerySpec = frmPar.Name  ' ** This gets unset somewhere!

            ' *************************
            ' ** Tax Code:
            ' *************************
1670        Select Case gblnIncomeTaxCoding
            Case True
1680          JC_Msc_TaxCode_Set frmSub  ' ** Procedure: Below.
1690          .taxcode.Locked = False
1700        Case False
              ' ** Disabled.
1710        End Select

            ' *************************
            ' ** Reinvested:
            ' *************************
1720        Select Case strThisJType
            Case "Dividend", "Interest"
1730          Select Case .posted
              Case True
1740            .Reinvested.Locked = True
1750          Case False
1760            Select Case .Reinvested
                Case True
                  ' ** If they've invoked this, lock it whether they've committed or not.
1770              .Reinvested.Locked = True
1780            Case False
1790              .Reinvested.Locked = False
1800            End Select
1810          End Select
1820        Case "Purchase", "Deposit", "Sold", "Withdrawn", "Liability (+)", "Liability (-)", _
                "Cost Adj.", "Misc.", "Paid", "Received"
1830          .Reinvested.Locked = True
1840        End Select

            ' ** Liability Criteria, Liability Rules:
            ' **   ONLY PCash and Cost define it!
            ' **   Liability with '+' (positive) OR ZERO pcash : Purchase  with  (negative) cost
            ' **   Liability with '-' (negative) pcash         : Sold      with  (positive) cost
            ' ** A Liability 'Sale' may have '-' (negative) icash, which
            ' **   is for interest tacked on top of the principal.
            ' ** A Liability 'Purchase' should never have an icash amount.

1850      End Select

          ' ** On a committed record, lock Posted.
          ' ** Any editing, however, will uncommit and unlock it.
1860      Select Case .posted
          Case True
1870        .posted.Locked = True
1880      Case False
1890        .posted.Locked = False
1900      End Select

1910      JC_Btn_Set strThisJType, .posted, frmPar, blnNewRec  ' ** Module Procedure: modJrnlCol_Buttons.
          ' ** On a completely empty record (save transdate),
          ' ** I want to keep the top buttons active!

1920      gstrFormQuerySpec = frmPar.Name  ' ** This gets unset somewhere!

1930      Select Case IsNull(.journaltype)
          Case True
1940        With frmPar
1950          If .cmdMemoReveal.Enabled = True Then
1960            .cmdMemoReveal.Enabled = False
1970            .cmdMemoReveal_R_raised_img_dis.Visible = True
1980            .cmdMemoReveal_R_raised_img.Visible = False
1990            .cmdMemoReveal_R_raised_semifocus_img.Visible = False
2000            .cmdMemoReveal_R_raised_focus_img.Visible = False
2010            .cmdMemoReveal_R_sunken_focus_img.Visible = False
2020            .cmdMemoReveal_L_raised_img_dis.Visible = False
2030            .cmdMemoReveal_L_raised_img.Visible = False
2040            .cmdMemoReveal_L_raised_semifocus_img.Visible = False
2050            .cmdMemoReveal_L_raised_focus_img.Visible = False
2060            .cmdMemoReveal_L_sunken_focus_img.Visible = False
2070          End If
2080        End With
2090      Case False
2100        If .journaltype <> "Paid" And frmPar.cmdMemoReveal.Enabled = True Then
2110          With frmPar
2120            .cmdMemoReveal.Enabled = False
2130            .cmdMemoReveal_R_raised_img_dis.Visible = True
2140            .cmdMemoReveal_R_raised_img.Visible = False
2150            .cmdMemoReveal_R_raised_semifocus_img.Visible = False
2160            .cmdMemoReveal_R_raised_focus_img.Visible = False
2170            .cmdMemoReveal_R_sunken_focus_img.Visible = False
2180            .cmdMemoReveal_L_raised_img_dis.Visible = False
2190            .cmdMemoReveal_L_raised_img.Visible = False
2200            .cmdMemoReveal_L_raised_semifocus_img.Visible = False
2210            .cmdMemoReveal_L_raised_focus_img.Visible = False
2220            .cmdMemoReveal_L_sunken_focus_img.Visible = False
2230          End With
2240        End If
2250      End Select

2260      Select Case .posted
          Case True
2270        With frmPar
2280          .cmdUnCommitOne.Enabled = True
2290          .cmdUnCommitOne_raised_img.Visible = True
2300          .cmdUnCommitOne_raised_semifocus_dots_img.Visible = False
2310          .cmdUnCommitOne_raised_focus_img.Visible = False
2320          .cmdUnCommitOne_raised_focus_dots_img.Visible = False
2330          .cmdUnCommitOne_sunken_focus_dots_img.Visible = False
2340          .cmdUnCommitOne_raised_img_dis.Visible = False
2350        End With
2360      Case False
2370        With frmPar
2380          .cmdUnCommitOne.Enabled = False
2390          .cmdUnCommitOne_raised_img_dis.Visible = True
2400          .cmdUnCommitOne_raised_img.Visible = False
2410          .cmdUnCommitOne_raised_semifocus_dots_img.Visible = False
2420          .cmdUnCommitOne_raised_focus_img.Visible = False
2430          .cmdUnCommitOne_raised_focus_dots_img.Visible = False
2440          .cmdUnCommitOne_sunken_focus_dots_img.Visible = False
2450        End With
2460      End Select

2470    End With

EXITP:
2480    Exit Sub

ERRH:
2490    Select Case ERR.Number
        Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Sub

Public Sub JC_Msc_Memo_Set(blnShow As Boolean, frmPar As Access.Form)
' ** Called by:
' **   JC_Msc_Cur_Set(), Above
' **   frmJournal_Columns_Sub:
' **     PrintCheck_AfterUpdate()
' **     AddRec()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Memo_Set"

2610    If lngTpp = 0& Then
          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
2620      lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
2630    End If

2640    With frmPar
2650      Select Case blnShow
          Case True
2660        With frmPar
2670          .cmdPreviewReport.Visible = False
2680          .cmdPreviewReport_raised_img.Visible = False
2690          .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
2700          .cmdPreviewReport_raised_focus_img.Visible = False
2710          .cmdPreviewReport_raised_focus_dots_img.Visible = False
2720          .cmdPreviewReport_sunken_focus_dots_img.Visible = False
2730          .cmdPreviewReport_raised_img_dis.Visible = False
2740          .cmdPrintReport.Visible = False
2750          .cmdPrintReport_raised_img.Visible = False
2760          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2770          .cmdPrintReport_raised_focus_img.Visible = False
2780          .cmdPrintReport_raised_focus_dots_img.Visible = False
2790          .cmdPrintReport_sunken_focus_dots_img.Visible = False
2800          .cmdPrintReport_raised_img_dis.Visible = False
2810          .RecsTot_Committed.Visible = False
2820          .RecsTot_Uncommitted.Visible = False
2830          .RecsTot_box.Visible = False
2840          .RecsTot_vline01.Visible = False
2850          .RecsTot_vline02.Visible = False
2860          .RecsTot_vline03.Visible = False
2870          .RecsTot_vline04.Visible = False
2880          .cmdUncom_box.Visible = False
2890          .cmdUncom_vline01.Visible = False
2900          .cmdUncom_vline02.Visible = False
2910          .cmdUncom_vline03.Visible = False
2920          .cmdUncom_vline04.Visible = False
2930          .cmdUncom_lbl.Visible = False
2940          .cmdUncomComAll.Visible = False
2950          .cmdUncomComAll_raised_img.Visible = False
2960          .cmdUncomComAll_raised_img_dis.Visible = False
2970          .cmdUncomDelAll.Visible = False
2980          .cmdUncomDelAll_raised_img.Visible = False
2990          .cmdUncomDelAll_raised_img_dis.Visible = False
3000          .cmdUnCommitOne.Visible = False
3010          .cmdUnCommitOne_raised_img.Visible = False
3020          .cmdUnCommitOne_raised_img_dis.Visible = False
3030          .JrnlMemo_Memo.Enabled = True
3040          .JrnlMemo_Memo.Visible = True
3050          .JrnlMemo_Memo_lbl2.Visible = True
3060          .JrnlMemo_Memo_box.Visible = True
3070          .JrnlMemo_Memo_box2.Visible = True
3080          .JrnlMemo_Memo_box3.Visible = True
3090          .cmdMemoReveal.Enabled = True
3100          .cmdMemoReveal_L_raised_img.Visible = True  ' ** Starts open.
3110          .cmdMemoReveal_L_raised_semifocus_img.Visible = False
3120          .cmdMemoReveal_L_raised_focus_img.Visible = False
3130          .cmdMemoReveal_L_sunken_focus_img.Visible = False
3140          .cmdMemoReveal_L_raised_img_dis.Visible = False
3150          .cmdMemoReveal_R_raised_img.Visible = False
3160          .cmdMemoReveal_R_raised_semifocus_img.Visible = False
3170          .cmdMemoReveal_R_raised_focus_img.Visible = False
3180          .cmdMemoReveal_R_sunken_focus_img.Visible = False
3190          .cmdMemoReveal_R_raised_img_dis.Visible = False
3200        End With
3210      Case False
3220        With frmPar
3230  On Error Resume Next
3240          .JrnlMemo_Memo.Enabled = False
3250          If ERR.Number <> 0 Then
                ' ** 2164:  You can't disable a control while it has the focus.
                ' ** 2165:  You can't hide a control that has the focus.
3260  On Error GoTo ERRH
3270            DoCmd.SelectObject acForm, .Name, False
3280            .frmJournal_Columns_Sub.SetFocus
3290          Else
3300  On Error GoTo ERRH
3310          End If
3320          .cmdMemoReveal.Enabled = False
3330          .cmdMemoReveal_R_raised_img_dis.Visible = True
3340          .cmdMemoReveal_R_raised_img.Visible = False
3350          .cmdMemoReveal_R_raised_semifocus_img.Visible = False
3360          .cmdMemoReveal_R_raised_focus_img.Visible = False
3370          .cmdMemoReveal_R_sunken_focus_img.Visible = False
3380          .cmdMemoReveal_L_raised_img.Visible = False
3390          .cmdMemoReveal_L_raised_semifocus_img.Visible = False
3400          .cmdMemoReveal_L_raised_focus_img.Visible = False
3410          .cmdMemoReveal_L_sunken_focus_img.Visible = False
3420          .cmdMemoReveal_L_raised_img_dis.Visible = False
3430          .JrnlMemo_Memo.Visible = False
3440          .JrnlMemo_Memo_lbl2.Visible = False
3450          .JrnlMemo_Memo_box.Visible = False
3460          .JrnlMemo_Memo_box2.Visible = False
3470          .JrnlMemo_Memo_box3.Visible = False
3480          .cmdPreviewReport.Visible = True
3490          Select Case .cmdPreviewReport.Enabled
              Case True
3500            .cmdPreviewReport_raised_img.Visible = True
3510          Case False
3520            .cmdPreviewReport_raised_img_dis.Visible = True
3530          End Select
3540          .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
3550          .cmdPreviewReport_raised_focus_img.Visible = False
3560          .cmdPreviewReport_raised_focus_dots_img.Visible = False
3570          .cmdPreviewReport_sunken_focus_dots_img.Visible = False
3580          .cmdPrintReport.Visible = True
3590          Select Case .cmdPrintReport.Enabled
              Case True
3600            .cmdPrintReport_raised_img.Visible = True
3610          Case False
3620            .cmdPrintReport_raised_img_dis.Visible = True
3630          End Select
3640          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3650          .cmdPrintReport_raised_focus_img.Visible = False
3660          .cmdPrintReport_raised_focus_dots_img.Visible = False
3670          .cmdPrintReport_sunken_focus_dots_img.Visible = False
3680          .RecsTot_box.Visible = True
3690          .RecsTot_vline01.Visible = True
3700          .RecsTot_vline02.Visible = True
3710          .RecsTot_vline03.Visible = True
3720          .RecsTot_vline04.Visible = True
3730          .RecsTot_Committed.Visible = True
3740          .RecsTot_Uncommitted.Visible = True
3750          .cmdUncom_box.Visible = True
3760          .cmdUncom_vline01.Visible = True
3770          .cmdUncom_vline02.Visible = True
3780          .cmdUncom_vline03.Visible = True
3790          .cmdUncom_vline04.Visible = True
3800          .cmdUncom_lbl.Visible = True
3810          .cmdUncomComAll.Visible = True
3820          Select Case .cmdUncomComAll.Enabled
              Case True
3830            .cmdUncomComAll_raised_img.Visible = True
3840          Case False
3850            .cmdUncomComAll_raised_img_dis.Visible = True
3860          End Select
3870          .cmdUncomDelAll.Visible = True
3880          Select Case .cmdUncomDelAll.Enabled
              Case True
3890            .cmdUncomDelAll_raised_img.Visible = True
3900          Case False
3910            .cmdUncomDelAll_raised_img_dis.Visible = True
3920          End Select
3930          .cmdUnCommitOne.Visible = True
3940          Select Case .cmdUnCommitOne.Enabled
              Case True
3950            .cmdUnCommitOne_raised_img.Visible = True
3960          Case False
3970            .cmdUnCommitOne_raised_img_dis.Visible = True
3980          End Select
3990        End With
4000      End Select
4010    End With

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Public Sub JC_Msc_Asset_Set(frmSub As Access.Form)
' ** Called by:
' **   JC_Msc_Cur_Set(), Above
' **   frmJournal_Columns_Sub:
' **     Form_Timer()
' **     chkShowAllAssets_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Asset_Set"

        Dim strThisJType As String
        Dim strRowSource As String

4110    strThisJType = vbNullString: strRowSource = vbNullString

4120    With frmSub
4130      gstrFormQuerySpec = .Name
4140      If IsNull(.journaltype) = False Then
4150        strThisJType = .journaltype
4160        Select Case strThisJType
            Case "Dividend"
              ' ** Dividend (dividendAssetNo.RowSource)
4170          If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                ' ** INCOME O/U:
                ' **   After-the-fact: qryJournal_Dividend_02f
4180            strRowSource = "qryJournal_Columns_10_MasterAsset_02"
4190          Else
4200            Select Case .chkShowAllAssets
                Case True
                  ' ** chkShowAllAssets = True:
                  ' **   All:            qryJournal_Dividend_02g
4210              strRowSource = "qryJournal_Columns_10_MasterAsset_06"
4220            Case False
                  ' ** chkShowAllAssets = False:
                  ' **   Just theirs:    qryJournal_Dividend_02e
4230              strRowSource = "qryJournal_Columns_10_MasterAsset_04"
4240            End Select
4250          End If
4260        Case "Interest"
              ' ** Interest (interestAssetNo.RowSource):
4270          If .accountno = "INCOME O/U" Or .accountno = "99-INCOME O/U" Then
                ' ** INCOME O/U:
                ' **   After-the-fact: qryJournal_Interest_02e
4280            strRowSource = "qryJournal_Columns_10_MasterAsset_08"
4290          Else
4300            Select Case .chkShowAllAssets
                Case True
                  ' ** chkShowAllAssets = True:
                  ' **   All:            qryJournal_Interest_02f
4310              strRowSource = "qryJournal_Columns_10_MasterAsset_12"
4320            Case False
                  ' ** chkShowAllAssets = False:
                  ' **   Just theirs:    qryJournal_Interest_02d
4330              strRowSource = "qryJournal_Columns_10_MasterAsset_10"
4340            End Select
4350          End If
4360        Case "Purchase", "Deposit"
              ' ** Purchase (purchaseAssetNo.RowSource): qryJournal_Purchase_03
              ' ** Non-Liability:   qryJournal_Purchase_03c
4370          strRowSource = "qryJournal_Columns_10_MasterAsset_15"
4380        Case "Liability (+)"
              ' ** Purchase (purchaseAssetNo.RowSource): qryJournal_Purchase_03
              ' ** Liability:       qryJournal_Purchase_03b
4390          strRowSource = "qryJournal_Columns_10_MasterAsset_14"
4400        Case "Sold", "Withdrawn"
              ' ** Sale (saleAssetno.RowSource):
              ' ** Their others:    qryJournal_Sale_04c
4410          strRowSource = "qryJournal_Columns_10_MasterAsset_19"
4420        Case "Liability (-)"
              ' ** Sale (saleAssetno.RowSource):
              ' ** Their Liability: qryJournal_Sale_04a
4430          strRowSource = "qryJournal_Columns_10_MasterAsset_17"
4440        Case "Cost Adj."
              ' ** Sale (saleAssetno.RowSource):
              ' ** Their Cost Adj.: qryJournal_Sale_04b
4450          strRowSource = "qryJournal_Columns_10_MasterAsset_18"
4460        Case "Received"
              ' ** Misc. (miscAssetNo.RowSource):
              ' ** All:             qryJournal_Misc_04
4470          strRowSource = "qryJournal_Columns_10_MasterAsset_21"
4480        Case "Misc.", "Paid"
              ' ** Nothing else.
4490        End Select
4500        .assetno.RowSource = strRowSource  ' ** Yes, that means Misc. and Paid will have vbNullString RowSources.
4510      End If  ' ** JournalType.
4520    End With
        'qryJournal_Columns_10_MasterAsset_02
        'qryJournal_Columns_10_MasterAsset_04
        'qryJournal_Columns_10_MasterAsset_06
        'qryJournal_Columns_10_MasterAsset_08
        'qryJournal_Columns_10_MasterAsset_10
        'qryJournal_Columns_10_MasterAsset_12
        'qryJournal_Columns_10_MasterAsset_14
        'qryJournal_Columns_10_MasterAsset_15
        'qryJournal_Columns_10_MasterAsset_17
        'qryJournal_Columns_10_MasterAsset_18
        'qryJournal_Columns_10_MasterAsset_19
        'qryJournal_Columns_10_MasterAsset_21

EXITP:
4530    Exit Sub

ERRH:
4540    Select Case ERR.Number
        Case Else
4550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4560    End Select
4570    Resume EXITP

End Sub

Public Sub JC_Msc_RevCode_Set(frmSub As Access.Form)
' ** Called by:
' **   JC_Msc_Cur_Set(), Above
' **   frmJournal_Columns_Sub:
' **     Form_Timer()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_RevCode_Set"

4610    With frmSub
4620      Select Case .journaltype
          Case "Dividend"
            ' ** INCOME.
4630        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_02" Then
4640          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_02"
4650          .revcode_ID.Requery
4660        End If
4670      Case "Interest"
            ' ** INCOME.
4680        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_02" Then
4690          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_02"
4700          .revcode_ID.Requery
4710        End If
4720      Case "Purchase", "Deposit"
            ' ** INCOME.
4730        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_02" Then
4740          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_02"
4750          .revcode_ID.Requery
4760        End If
4770      Case "Sold"
            ' ** INCOME.
4780        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_02" Then
4790          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_02"
4800          .revcode_ID.Requery
4810        End If
4820      Case "Withdrawn"
            ' ** ALL.
4830        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_01" Then
4840          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_01"
4850          .revcode_ID.Requery
4860        End If
4870      Case "Misc."
            ' ** ALL.
4880        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_01" Then
4890          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_01"
4900          .revcode_ID.Requery
4910        End If
4920      Case "Paid"
            ' ** EXPENSE.
4930        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_03" Then
4940          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_03"
4950          .revcode_ID.Requery
4960        End If
4970      Case "Received"
            ' ** INCOME.
4980        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_02" Then
4990          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_02"
5000          .revcode_ID.Requery
5010        End If
5020      Case "Liability (+)", "Liability (-)"
            ' ** EXPENSE.
5030        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_03" Then
5040          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_03"
5050          .revcode_ID.Requery
5060        End If
5070      Case "Cost Adj."
            ' ** ALL.
5080        If .revcode_ID.RowSource <> "qryJournal_Columns_10_M_RevCode_01" Then
5090          .revcode_ID.RowSource = "qryJournal_Columns_10_M_RevCode_01"
5100          .revcode_ID.Requery
5110        End If
5120      End Select
5130    End With

EXITP:
5140    Exit Sub

ERRH:
5150    Select Case ERR.Number
        Case Else
5160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5170    End Select
5180    Resume EXITP

End Sub

Public Sub JC_Msc_TaxCode_Set(frmSub As Access.Form)
' ** Called by:
' **   JC_Msc_Cur_Set(), Above
' **   frmJournal_Columns_Sub:
' **     Form_Timer()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_TaxCode_Set"

5210    With frmSub
5220      Select Case .journaltype
          Case "Dividend"
            ' ** INCOME.
5230        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_02" Then
5240          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_02"
5250          .taxcode.Requery
5260        End If
5270      Case "Interest"
            ' ** INCOME.
5280        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_02" Then
5290          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_02"
5300          .taxcode.Requery
5310        End If
5320      Case "Purchase", "Deposit"
            ' ** INCOME.
5330        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_02" Then
5340          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_02"
5350          .taxcode.Requery
5360        End If
5370      Case "Sold"
            ' ** INCOME.
5380        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_02" Then
5390          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_02"
5400          .taxcode.Requery
5410        End If
5420      Case "Withdrawn"
            ' ** ALL.
5430        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_01" Then
5440          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_01"
5450          .taxcode.Requery
5460        End If
5470      Case "Liability (+)", "Liability (-)"
            ' ** EXPENSE.
5480        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_04" Then
5490          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_04"  ' ** EXPENSE LTD.
5500          .taxcode.Requery
5510        End If
5520      Case "Paid"
            ' ** EXPENSE.
5530        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_03" Then
5540          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_03"
5550          .taxcode.Requery
5560        End If
5570      Case "Received"
            ' ** INCOME.
5580        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_02" Then
5590          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_02"
5600          .taxcode.Requery
5610        End If
5620      Case "Misc."
            ' ** ALL.
5630        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_01" Then
5640          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_01"
5650          .taxcode.Requery
5660        End If
5670      Case "Cost Adj."
            ' ** ALL.
5680        If .taxcode.RowSource <> "qryJournal_Columns_10_TaxCode_01" Then
5690          .taxcode.RowSource = "qryJournal_Columns_10_TaxCode_01"
5700          .taxcode.Requery
5710        End If
5720      End Select
5730    End With

EXITP:
5740    Exit Sub

ERRH:
5750    Select Case ERR.Number
        Case Else
5760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5770    End Select
5780    Resume EXITP

End Sub

Public Function JC_Msc_Loc_Set(strAccountNo As String, lngAssetNo As Long) As Long
' ** Called by:
' **   frmJournal_Columns_Sub:
' **     assetno_AfterUpdate()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Loc_Set"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRetVal As Long

5810    lngRetVal = 1&  ' ** Default to '{Unassigned}', '{no entry}'.

5820    If strAccountNo <> vbNullString And lngAssetNo > 0& Then
5830      Set dbs = CurrentDb
5840      With dbs

            ' ** qryMap_02b (qryMap_02a (ActiveAssets, grouped by accountno, assetno, Location_ID), grouped by
            ' ** accountno, assetno, with cnt, Location_ID_min, Location_ID_max), by specified [actno], [astno].
5850        Set qdf = .QueryDefs("qryMap_03")
5860        With qdf.Parameters
5870          ![actno] = strAccountNo
5880          ![astno] = lngAssetNo
5890        End With
5900        Set rst = qdf.OpenRecordset
5910        With rst
5920          If .BOF = True And .EOF = True Then
                ' ** Check the Ledger.
5930          Else
5940            .MoveFirst
5950            Select Case ![cnt]
                Case 1
5960              lngRetVal = ![Location_ID_max]  ' ** They'll both be the same.
5970            Case 2
5980              If ![Location_ID_min] = 1& Then
5990                lngRetVal = ![Location_ID_max]
6000              Else
                    ' ** The default stands.
6010              End If
6020            Case Else
                  ' ** The default stands.
6030            End Select
6040          End If
6050          .Close
6060        End With

6070        If lngRetVal = 1& Then
              ' ** qryMap_04b (qryMap_04a (Ledger, grouped by accountno, assetno, Location_ID.), grouped by
              ' ** accountno, assetno, with cnt, Location_ID_min, Location_ID_max), by specified [actno], [astno].
6080          Set qdf = .QueryDefs("qryMap_05")
6090          With qdf.Parameters
6100            ![actno] = strAccountNo
6110            ![astno] = lngAssetNo
6120          End With
6130          Set rst = qdf.OpenRecordset
6140          With rst
6150            If .BOF = True And .EOF = True Then
                  ' ** The default stands.
6160            Else
6170              .MoveFirst
6180              Select Case ![cnt]
                  Case 1
6190                lngRetVal = ![Location_ID_max]  ' ** They'll both be the same.
6200              Case 2
6210                If ![Location_ID_min] = 1& Then
6220                  lngRetVal = ![Location_ID_max]
6230                Else
                      ' ** The default stands.
6240                End If
6250              Case Else
                    ' ** The default stands.
6260              End Select
6270            End If
6280            .Close
6290          End With
6300        End If

6310        .Close
6320      End With
6330    End If

EXITP:
6340    Set rst = Nothing
6350    Set qdf = Nothing
6360    Set dbs = Nothing
6370    JC_Msc_Loc_Set = lngRetVal
6380    Exit Function

ERRH:
6390    lngRetVal = 1&
6400    Select Case ERR.Number
        Case Else
6410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6420    End Select
6430    Resume EXITP

End Function

Public Function JC_Msc_Find_JType(strJType As String, blnPosted As Boolean, rst As DAO.Recordset) As Long
' ** Called by:
' **   frmJournal_Columns:
' **     cmdSpecPurp_Sold_PaidTotal_Click()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Find_JType"

        Dim lngRecs As Long
        Dim lngRetVal As Long

6510    lngRetVal = 0&

6520    With rst
6530      If .BOF = True And .EOF = True Then
            ' ** No trans, return zero.
6540      Else
6550        .MoveLast
6560        lngRecs = .RecordCount
6570        .MoveFirst
6580        .FindFirst "[journaltype] = '" & strJType & "' And [posted] = " & IIf(blnPosted = True, "True", "False")
6590        If .NoMatch = False Then
6600          lngRetVal = ![JrnlCol_ID]
6610        End If
6620      End If
6630      .Close
6640    End With

EXITP:
6650    JC_Msc_Find_JType = lngRetVal
6660    Exit Function

ERRH:
6670    lngRetVal = 0&
6680    Select Case ERR.Number
        Case Else
6690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6700    End Select
6710    Resume EXITP

End Function

Public Function JC_Msc_Find_Adjascent(lngJrnlColID As Long, rst As DAO.Recordset) As Long
' ** Called by:
' **   frmJournal_Columns:
' **     cmdDelete_Click()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Find_Adjascent"

        Dim lngRecs As Long
        Dim lngRetVal As Long

6810    lngRetVal = 0&

6820    With rst
6830      If .BOF = True And .EOF = True Then
            ' ** Then this shouldn't have been called!
6840      Else
6850        .MoveLast
6860        lngRecs = .RecordCount
6870        If lngRecs > 1& Then
6880          .MoveFirst
6890          .FindFirst "[JrnlCol_ID] = " & CStr(lngJrnlColID)
6900          If .NoMatch = False Then
6910  On Error Resume Next
6920            .MovePrevious
6930            If ERR.Number <> 0 Then
6940  On Error GoTo ERRH
6950              .MoveFirst
6960              .FindFirst "[JrnlCol_ID] = " & CStr(lngJrnlColID)
6970              If .NoMatch = False Then
6980                .MoveNext
6990                lngRetVal = ![JrnlCol_ID]
7000              End If
7010            Else
7020  On Error GoTo ERRH
7030  On Error Resume Next
7040              lngRetVal = ![JrnlCol_ID]
7050              If ERR.Number <> 0 Then
7060  On Error GoTo ERRH
7070                .MoveFirst
7080                .FindFirst "[JrnlCol_ID] = " & CStr(lngJrnlColID)
7090                If .NoMatch = False Then
7100                  .MoveNext
7110                  lngRetVal = ![JrnlCol_ID]
7120                End If
7130              Else
7140  On Error GoTo ERRH
7150              End If
7160            End If
7170          End If
7180        End If
7190      End If
7200      .Close
7210    End With

EXITP:
7220    JC_Msc_Find_Adjascent = lngRetVal
7230    Exit Function

ERRH:
7240    lngRetVal = 0&
7250    Select Case ERR.Number
        Case Else
7260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7270    End Select
7280    Resume EXITP

End Function

Public Sub JC_Msc_StatusBar_SetSub(strProc As String, blnNotPopup As Boolean, frm As Access.Form)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_StatusBar_SetSub"

        Dim ctl As Access.Control, ctl2 As Access.Control
        Dim strControl As String
        Dim blnShowAssets As Boolean
        Dim intPos01 As Integer
        Dim lngX As Long

        'frmJournal_Columns
        'frmJournal_Columns_Sub
        'modJrnlCol_Misc

7310    frm.StatusBar_Load  ' ** Form Procedure: frmJournal_Columns_Sub.
7320    JC_Msc_StatusBar_Load blnNotPopup, frm  ' ** Procedure: Below.

        ' ********************************************
        ' ** Array: arr_varStat()
        ' **
        ' **   Element  Name              Constant
        ' **   =======  ================  ==========
        ' **      0     Section Number    S_SEC
        ' **      1     Control Name      S_NAM
        ' **      2     StatusBarText     S_TXT
        ' **
        ' ********************************************

7330    If lngTpp = 0& Then
          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7340      lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7350    End If

7360    If blnNotPopup = True And strProc <> vbNullString Then
7370      intPos01 = InStr(strProc, "_GotFocus")
7380      If intPos01 > 0 Then
7390        strControl = Left(strProc, (intPos01 - 1))
7400        For lngX = 0& To (lngStats - 1&)
7410          If arr_varStat(S_NAM, lngX) = strControl Then
7420            If arr_varStat(S_TXT, lngX) <> vbNullString Then
7430              SysCmd acSysCmdSetStatus, arr_varStat(S_TXT, lngX)
7440            Else
7450              SysCmd acSysCmdClearStatus
7460            End If
7470            Exit For
7480          End If
7490        Next
7500        With frm
7510          .HighLight_BG_box.Visible = False
7520          .HighLight_BG_box.Top = 0&
7530          .chkShowAllAssets_lbl.BackStyle = acBackStyleTransparent
7540          .chkShowAllAssets_lbl.ForeColor = CLR_DKGRY2
7550          blnShowAssets = False
7560          For Each ctl In .FormHeader.Controls
7570            Select Case ctl.ControlType
                Case acLabel
7580              If ctl.Visible = True Then
7590                intPos01 = InStr(ctl.Name, "_lbl")
7600                If intPos01 > 0 Then
7610                  If Left(ctl.Name, (intPos01 - 1)) = strControl Then
7620                    If Right(ctl.Name, 4) = "_lbl" Then
7630                      If ctl.Name = "chkShowAllAssets_lbl" Then
                            ' ** The whole column focus should be on assetno.
7640                        blnShowAssets = True
7650                        Set ctl2 = .Controls("assetno_lbl")
7660                        .HighLight_BG_box.BackStyle = acBackStyleNormal
7670                        .HighLight_BG_box.Left = ctl2.Left
7680                        .HighLight_BG_box.Width = ctl2.Width
7690                        .HighLight_BG_box.Height = ctl2.Height
7700                        .HighLight_BG_box.Top = ctl2.Top - lngTpp
7710                        .HighLight_BG_box.Visible = True
7720                        ctl2.ForeColor = CLR_WHT
7730                        ctl.BackStyle = acBackStyleNormal
7740                        ctl.BackColor = CLR_HILITE
7750                        ctl.ForeColor = CLR_WHT
7760                      Else
                            '.HighLight_BG_box.BackColor = CLR_AC07  '13603685
7770                        .HighLight_BG_box.BackStyle = acBackStyleNormal
7780                        .HighLight_BG_box.Left = ctl.Left
7790                        If strControl <> "taxcode" Then
7800                          .HighLight_BG_box.Width = ctl.Width
7810                        Else
7820                          .HighLight_BG_box.Width = ctl.Width - (3& * lngTpp)
7830                        End If
7840                        .HighLight_BG_box.Height = ctl.Height - lngTpp
7850                        .HighLight_BG_box.Top = ctl.Top
7860                        .HighLight_BG_box.Visible = True
7870                        ctl.ForeColor = CLR_WHT
7880                        Select Case ctl.Name
                            Case "assetno_lbl"
7890                          If .chkShowAllAssets.Visible = True Then
7900                            blnShowAssets = True
7910                            .HighLight_BG_box.Top = .HighLight_BG_box.Top - lngTpp
7920                            .HighLight_BG_box.Height = .HighLight_BG_box.Height + lngTpp
7930                            .chkShowAllAssets_lbl.ForeColor = CLR_WHT
7940                          End If
7950                        End Select
7960                      End If
7970                    Else
7980                      ctl.ForeColor = CLR_WHT  ' ** These are .._lbl2's.
7990                    End If
8000                  Else
8010                    If ctl.ForeColor = CLR_WHT Then
8020                      Select Case ctl.Name
                          Case "chkShowAllAssets_lbl"
8030                        If blnShowAssets = False Then
8040                          ctl.ForeColor = CLR_DKGRY2
8050                        End If
8060                      Case "assetno_lbl"
8070                        If blnShowAssets = False Then
8080                          ctl.ForeColor = CLR_DKGRY2
8090                        End If
8100                      Case Else
8110                        ctl.ForeColor = CLR_DKGRY2
8120                      End Select
8130                    End If
8140                  End If
8150                End If
8160              End If
8170            End Select
8180          Next
8190        End With
8200      End If
8210    End If

EXITP:
8220    Set ctl2 = Nothing
8230    Set ctl = Nothing
8240    Exit Sub

ERRH:
8250    Select Case ERR.Number
        Case Else
8260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8270    End Select
8280    Resume EXITP

End Sub

Private Sub JC_Msc_StatusBar_Load(blnNotPopup As Boolean, frm As Access.Form)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_StatusBar_Load"

        Dim arr_varRetVal As Variant

        Const RV_ERR   As Integer = 0
        Const RV_SBARS As Integer = 1
        Const RV_POPUP As Integer = 2

8310  On Error Resume Next
8320    If lngStats = 0& Or IsEmpty(arr_varStat) Then
8330      arr_varRetVal = JC_Msc_SBar_Load(frm, True)  ' ** Function: Below.
8340      If arr_varRetVal(RV_ERR, 0) = vbNullString Then
8350        Select Case IsEmpty(arr_varRetVal(RV_SBARS, 0))
            Case True
8360          arr_varStat = Empty
8370          lngStats = 0&
8380          blnNotPopup = False
8390        Case False
8400          Select Case IsNull(arr_varRetVal(RV_SBARS, 0))
              Case True
8410            arr_varStat = Empty
8420            lngStats = 0&
8430            blnNotPopup = False
8440          Case False
8450            arr_varStat = arr_varRetVal(RV_SBARS, 0)
8460            lngStats = (UBound(arr_varStat, 2) + 1&)
8470            blnNotPopup = arr_varRetVal(RV_POPUP, 0)
8480          End Select
8490        End Select
8500      End If
8510    End If
8520  On Error GoTo ERRH

EXITP:
8530    Exit Sub

ERRH:
8540    Select Case ERR.Number
        Case Else
8550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8560    End Select
8570    Resume EXITP

End Sub

Public Function JC_Msc_SBar_Load(frm As Access.Form, Optional varIsSub As Variant) As Variant
' ** Form may be either parent or sub.

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_SBar_Load"

        Dim ctl As Access.Control
        Dim strSBarText As String
        Dim blnNotPopup As Boolean, blnDoc As Boolean, blnIsSub As Boolean
        Dim lngX As Long, lngE As Long
        Dim arr_varRetVal() As Variant

        ' ** Array: arr_varRetVal().
        Const RV_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const RV_ERR   As Integer = 0
        Const RV_SBARS As Integer = 1
        Const RV_POPUP As Integer = 2

8610    ReDim arr_varRetVal(RV_ELEMS, 0)
8620    arr_varRetVal(RV_ERR, 0) = vbNullString

8630    With frm

8640      Select Case IsMissing(varIsSub)
          Case True
8650        blnIsSub = False
8660      Case False
8670        blnIsSub = CBool(varIsSub)
8680      End Select

8690      Select Case blnIsSub
          Case True
8700        blnNotPopup = Not (.Parent.PopUp)
8710      Case False
8720        blnNotPopup = Not (.PopUp)
8730      End Select

8740      If blnNotPopup = True Then

8750        lngStats = 0&
8760        ReDim arr_varStat(S_ELEMS, 0)

8770        For lngX = 0& To 2&
8780          For Each ctl In .Section(lngX).Controls
8790            strSBarText = vbNullString
8800            With ctl
8810              If Left(.Name, 11) <> "FocusHolder" Then
8820                blnDoc = True
8830                If (.ControlType = acTextBox Or .ControlType = acCheckBox) Then
8840                  If .Visible = False And .Enabled = False And .Locked = True Then
8850                    Select Case .ControlType
                        Case acTextBox
8860                      If (.ForeColor = CLR_BLU) And (.BorderColor = CLR_BLU And .BorderStyle = acBorderStyleSolid) And _
                              (.BackColor = CLR_WHT And .BackStyle = acBackStyleNormal) And (.SpecialEffect = acSpecialEffectFlat) And _
                              (.Width = 360& And .Height = 255&) Then
                            ' ** Alright, so I got a little carried away... (I could have include the FontName and FontSize!)
                            ' ** NO TXT: currow_hilite_forecolor
                            ' ** NO TXT: ToTaxLot
8870                        blnDoc = False
8880                      End If
8890                    Case acCheckBox
8900                      If .Controls.Count = 0 Then  ' ** It's got no label.
8910                        blnDoc = False
8920                      End If
8930                    End Select
8940                  ElseIf .ControlType = acTextBox Then
8950                    If .FontName = "Terminal" Then
8960                      blnDoc = False
8970                    End If
8980                  End If
8990                End If
9000                If blnDoc = True Then
9010                  Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acCommandButton, acOptionButton, acCheckBox, _
                          acToggleButton, acTabCtl, acPage
                        ' ** Generally, the text goes on the buttons for an acOptionGroup.
9020  On Error Resume Next
9030                    strSBarText = .StatusBarText
9040                    If ERR.Number <> 0 Then
9050  On Error GoTo ERRH
                          'Debug.Print "'NO TXT: " & .Name
9060                    Else
9070  On Error GoTo ERRH
9080                      lngStats = lngStats + 1&
9090                      lngE = lngStats - 1&
9100                      ReDim Preserve arr_varStat(S_ELEMS, lngE)
                          ' ********************************************
                          ' ** Array: arr_varStat()
                          ' **
                          ' **   Element  Name              Constant
                          ' **   =======  ================  ==========
                          ' **      0     Section Number    S_SEC
                          ' **      1     Control Name      S_NAM
                          ' **      2     StatusBarText     S_TXT
                          ' **
                          ' ********************************************
9110                      arr_varStat(S_SEC, lngE) = lngX
9120                      arr_varStat(S_NAM, lngE) = .Name
9130                      arr_varStat(S_TXT, lngE) = strSBarText  ' ** May be vbNullString.
9140                    End If
9150                  Case Else
                        ' ** No StatusBarText.
9160                  End Select
9170                End If  ' ** blnDoc.
9180              Else
                    ' ** NO TXT: FocusHolder2
9190              End If  ' ** FocusHolder.
9200            End With
9210          Next
9220        Next

9230      End If  ' ** blnIsPopup.

9240    End With

9250    arr_varRetVal(RV_SBARS, 0) = arr_varStat
9260    arr_varRetVal(RV_POPUP, 0) = blnNotPopup

EXITP:
9270    JC_Msc_SBar_Load = arr_varRetVal
9280    Exit Function

ERRH:
9290    arr_varRetVal(RV_ERR, 0) = RET_ERR
9300    Select Case ERR.Number
        Case Else
9310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9320    End Select
9330    Resume EXITP

End Function

Public Function JC_Msc_Gfx_Load() As Boolean
' ** Generates qryJournal_Columns_18.

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Gfx_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, fld As DAO.Field
        Dim lngQFlds As Long, arr_varQFld() As Variant
        Dim strNum As String, strSQL As String, strLastCtlName As String, strTable As String, strLastNum As String
        Dim lngLastAltNum As Long, lngLastCnt As Long, lngLastCtlID As Long
        Dim lngRecs_FGfx As Long
        Dim lngTmp01 As Long
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQFld().
        Const QF_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const QF_FNAM As Integer = 0
        Const QF_FVAL As Integer = 1
        Const QF_IMG  As Integer = 2
        Const QF_NUM  As Integer = 3
        Const QF_ALT  As Integer = 4

        Const QF_TBL1 As String = "tblForm_Graphics."
        Const QF_TBL2 As String = "tblForm_Graphics_1."

9410    blnRetVal = True

9420    Set dbs = CurrentDb
9430    With dbs
          ' ** tblForm_Graphics, just frmJournal_Colums.
9440      Set qdf = .QueryDefs("qryJournal_Columns_16d")
9450      Set rst = qdf.OpenRecordset
9460      With rst
9470        .MoveLast
9480        lngRecs_FGfx = .RecordCount
9490        .MoveFirst

9500        lngQFlds = 0&
9510        ReDim arr_varQFld(QF_ELEMS, 0)

9520        strNum = vbNullString: strSQL = vbNullString

9530        For lngX = 1& To lngRecs_FGfx
9540          lngLastAltNum = -1&
9550          For lngY = 1& To 4&
9560            Select Case lngY
                Case 1&
9570              Set fld = .Fields("frmgfx_id")
9580            Case 2&
9590              Set fld = .Fields("frmgfx_alt")
9600              lngLastAltNum = fld.Value
9610            Case 3&
9620              Set fld = .Fields("frm_id")
9630            Case 4&
9640              Set fld = .Fields("frm_name")
9650            End Select
9660            With fld
9670              lngQFlds = lngQFlds + 1&
9680              lngE = lngQFlds - 1&
9690              ReDim Preserve arr_varQFld(QF_ELEMS, lngE)
9700              arr_varQFld(QF_FNAM, lngE) = .Name
9710              arr_varQFld(QF_FVAL, lngE) = .Value
9720              arr_varQFld(QF_IMG, lngE) = CBool(False)
9730              arr_varQFld(QF_NUM, lngE) = vbNullString
9740              arr_varQFld(QF_ALT, lngE) = lngLastAltNum
9750            End With
9760          Next  ' ** lngY.

9770          lngLastCtlID = -1&
9780          For lngY = 0& To (.Fields.Count - 1)
9790            Set fld = .Fields(lngY)
9800            With fld
9810              If Left(.Name, 7) = "ctl_id_" Then
9820                lngLastCtlID = .Value
9830              End If
9840              If lngLastCtlID > 0& Then
9850                If Left(.Name, 9) = "ctl_name_" Then
9860                  strNum = Right(.Name, 2)
9870                  If rst.Fields("ctl_id_" & strNum) > 0& Then
9880                    lngQFlds = lngQFlds + 1&
9890                    lngE = lngQFlds - 1&
9900                    ReDim Preserve arr_varQFld(QF_ELEMS, lngE)
9910                    arr_varQFld(QF_FNAM, lngE) = .Name
9920                    arr_varQFld(QF_FVAL, lngE) = .Value
9930                    arr_varQFld(QF_IMG, lngE) = CBool(False)
9940                    arr_varQFld(QF_NUM, lngE) = strNum
9950                    arr_varQFld(QF_ALT, lngE) = lngLastAltNum
9960                  End If
9970                ElseIf Left(.Name, 13) = "xadgfx_image_" Then
9980                  strNum = Right(.Name, 2)
9990                  If rst.Fields("xadgfx_id_" & strNum) > 0& Then
10000                   lngQFlds = lngQFlds + 1&
10010                   lngE = lngQFlds - 1&
10020                   ReDim Preserve arr_varQFld(QF_ELEMS, lngE)
10030                   arr_varQFld(QF_FNAM, lngE) = .Name
10040                   arr_varQFld(QF_FVAL, lngE) = Null
10050                   arr_varQFld(QF_IMG, lngE) = CBool(True)
10060                   arr_varQFld(QF_NUM, lngE) = strNum
10070                   arr_varQFld(QF_ALT, lngE) = lngLastAltNum
10080                 End If
10090               End If
10100             End If
10110           End With
10120         Next  ' ** Fields: lngY.
10130         If lngX < lngRecs_FGfx Then .MoveNext
10140       Next  ' ** lngRecs_FGfx: lngX.

10150       .Close
10160     End With

          ' ** Put the frmgfx_alt number into the elements still missing it.
10170     For lngX = 1& To lngRecs_FGfx
10180       lngLastAltNum = -1&
10190       For lngY = 0& To (lngQFlds - 1&)
10200         If arr_varQFld(QF_ALT, lngY) = -1& Then
                ' ** Skip first, then we'll come back to them.
10210         Else
10220           lngLastAltNum = arr_varQFld(QF_ALT, lngY)
10230           Exit For
10240         End If
10250       Next  ' ** lngY.
10260       For lngY = 0& To (lngQFlds - 1&)
10270         If arr_varQFld(QF_ALT, lngY) = -1& Then
10280           arr_varQFld(QF_ALT, lngY) = lngLastAltNum
10290         End If
10300         If arr_varQFld(QF_NUM, lngY) <> vbNullString And lngX = 1& Then
10310           Exit For
10320         End If
10330       Next  ' ** lngY.
10340     Next  ' ** lngX.

10350     strSQL = "SELECT "
10360     strLastCtlName = vbNullString: strTable = vbNullString: strLastNum = vbNullString
10370     lngLastCnt = 0&: lngLastAltNum = -1&
10380     lngTmp01 = 0&
10390     For lngX = 0& To (lngQFlds - 1&)
10400       If arr_varQFld(QF_NUM, lngX) = vbNullString Then
10410         If lngLastCnt = 0& And lngLastAltNum = -1& Then
                ' ** It'll only hit here once!
10420           lngLastCnt = lngLastCnt + 1&
10430           lngLastAltNum = arr_varQFld(QF_ALT, lngX)
10440           strTable = QF_TBL1
10450         ElseIf arr_varQFld(QF_ALT, lngX) <> lngLastAltNum Then
                ' ** It'll only hit here once per tblForm_Graphics record!
10460           lngLastCnt = lngLastCnt + 1&
10470           lngLastAltNum = arr_varQFld(QF_ALT, lngX)
10480           strTable = QF_TBL2
10490         End If
10500         If arr_varQFld(QF_FNAM, lngX) = "frmgfx_id" Then
10510           If lngTmp01 = 0& Then
10520             strSQL = strSQL & strTable & arr_varQFld(QF_FNAM, lngX) & ", "
10530             lngTmp01 = arr_varQFld(QF_FVAL, lngX)
10540           Else
10550             strSQL = strSQL & strTable & arr_varQFld(QF_FNAM, lngX) & " " & _
                    "AS " & arr_varQFld(QF_FNAM, lngX) & CStr(2) & ", "
10560           End If
10570         Else
10580           If lngLastCnt = 1& Then
10590             If arr_varQFld(QF_FNAM, lngX) = "frm_name" Then
10600               strSQL = strSQL & "tblForm." & arr_varQFld(QF_FNAM, lngX) & ", "
10610             Else
10620               strSQL = strSQL & strTable & arr_varQFld(QF_FNAM, lngX) & ", "
10630             End If
10640           Else
10650             If arr_varQFld(QF_FNAM, lngX) = "frm_name" Then
10660               strSQL = strSQL & "tblForm." & arr_varQFld(QF_FNAM, lngX) & " " & _
                      "AS " & arr_varQFld(QF_FNAM, lngX) & CStr(lngLastCnt) & ", "
10670             Else
10680               strSQL = strSQL & strTable & arr_varQFld(QF_FNAM, lngX) & " " & _
                      "AS " & arr_varQFld(QF_FNAM, lngX) & CStr(lngLastCnt) & ", "
10690             End If
10700           End If
10710         End If
10720       Else
10730         Select Case arr_varQFld(QF_IMG, lngX)
              Case True
10740           strSQL = strSQL & strTable & arr_varQFld(QF_FNAM, lngX) & " AS " & strLastCtlName & ", "
10750         Case False
10760           If strLastNum = vbNullString Then
10770             strLastNum = arr_varQFld(QF_NUM, lngX)
10780           End If
10790           If Val(arr_varQFld(QF_NUM, lngX)) < Val(strLastNum) Then
10800             strLastNum = Right("00" & CStr(Val(strLastNum) + 1), 2)
10810             strSQL = strSQL & strTable & arr_varQFld(QF_FNAM, lngX) & " " & _
                    "AS " & Left(arr_varQFld(QF_FNAM, lngX), (Len(arr_varQFld(QF_FNAM, lngX)) - 2)) & strLastNum & ", "
10820           Else
10830             strLastNum = arr_varQFld(QF_NUM, lngX)
10840             strSQL = strSQL & strTable & arr_varQFld(QF_FNAM, lngX) & ", "
10850           End If
10860           strLastCtlName = arr_varQFld(QF_FVAL, lngX)
10870         End Select
10880       End If
10890     Next  ' ** lngQFlds: lngX.
10900     strSQL = Trim(strSQL)
10910     If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, (Len(strSQL) - 1))
10920     strSQL = strSQL & vbCrLf
10930     strSQL = strSQL & "FROM (tblForm INNER JOIN tblForm_Graphics ON " & _
            "tblForm.frm_id = tblForm_Graphics.frm_id) INNER JOIN " & _
            "tblForm_Graphics AS tblForm_Graphics_1 ON " & _
            "tblForm_Graphics.frm_id = tblForm_Graphics_1.frm_id"  ' ** Presumes only 2 records.
10940     strSQL = strSQL & vbCrLf
10950     strSQL = strSQL & "WHERE (((tblForm_Graphics.frmgfx_alt)=0) AND " & _
            "((tblForm.frm_name)='frmJournal_Columns') AND " & _
            "((tblForm_Graphics_1.frmgfx_alt)=" & CStr(lngLastAltNum) & "));"  ' ** Presumes only 2 records.
          'Debug.Print "'" & strSQL

10960     Set qdf = .QueryDefs("qryJournal_Columns_18")
10970     qdf.SQL = strSQL

10980     .Close
10990   End With

11000   Beep

EXITP:
11010   Set fld = Nothing
11020   Set rst = Nothing
11030   Set qdf = Nothing
11040   Set dbs = Nothing
11050   JC_Msc_Gfx_Load = blnRetVal
11060   Exit Function

ERRH:
11070   blnRetVal = False
11080   Select Case ERR.Number
        Case Else
11090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11100   End Select
11110   Resume EXITP

End Function

Public Function JC_Msc_Gfx_Load2() As Boolean
' ** Reporting only, no changes made anywhere.

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Gfx_Load2"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngQryFlds As Long, arr_varQryFld() As Variant
        Dim lngImgs As Long, arr_varImg() As Variant
        Dim lngFlds As Long, arr_varFld() As Variant
        Dim strQryName As String
        Dim blnListButtons As Boolean, blnListImages As Boolean, blnListFields As Boolean
        Dim lngRecs As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQryFld().
        Const QF_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const QF_NAM As Integer = 0
        'Const QF_TYP As Integer = 1
        Const QF_FND As Integer = 2

        ' ** Array: arr_varImg().
        Const I_ELEMS As Integer = 6  ' ** Array's first-element UBound().
        Const I_CSID   As Integer = 0
        Const I_FRMID  As Integer = 1
        Const I_CTLID  As Integer = 2
        Const I_CTLNAM As Integer = 3
        Const I_GFXID  As Integer = 4
        Const I_GFXNAM As Integer = 5
        Const I_FND    As Integer = 6

        ' ** Array: arr_varFld().
        Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const F_NAM As Integer = 0
        'Const F_TYP As Integer = 1
        'Const F_REQ As Integer = 2

11210   blnRetVal = True

11220   blnListButtons = False  ' ** True: List query button fields; False: Don't list.
11230   blnListImages = False   ' ** True: List image fields; False: Don't list.
11240   blnListFields = True    ' ** True: List all query fields; False: Don't list.

11250   lngQryFlds = 0&
11260   ReDim arr_varQryFld(QF_ELEMS, 0)

11270   strQryName = "qryJournal_Columns_18"
11280   blnRetVal = Tbl_Fld_List(arr_varQryFld, strQryName)  ' ** Module Functions: modXAdminFuncs
11290   lngQryFlds = (UBound(arr_varQryFld, 2) + 1&)

        ' *****************************************************
        ' ** Array: arr_varQryFld()
        ' **
        ' **   Field  Element  Name                Constant
        ' **   =====  =======  ==================  ==========
        ' **     1       0     fld_name            QF_NAM
        ' **     2       1     datatype_db_type    QF_TYP
        ' **     3       2     Found               QF_FND
        ' **
        ' *****************************************************

11300   If lngQryFlds > 0& Then

11310     lngImgs = 0&
11320     ReDim arr_varImg(I_ELEMS, 0)

11330     Set dbs = CurrentDb
11340     With dbs

            ' ** qryJournal_Columns_15 (tblForm_Control, just frmJournal_Columns Bound Object Frames), linked to
            ' ** qryJournal_Columns_16 (tblXAdmin_Graphics, just frmJournal_Columns buttons), by Sort1.
11350       Set qdf = .QueryDefs("qryJournal_Columns_17")
11360       Set rst = qdf.OpenRecordset
11370       With rst
11380         If .BOF = True And .EOF = True Then
                ' ** Something's messed up!
11390         Else
11400           .MoveLast
11410           lngRecs = .RecordCount  '38
11420           .MoveFirst
11430           For lngX = 1& To lngRecs
11440             lngImgs = lngImgs + 1&
11450             lngE = lngImgs - 1&
11460             ReDim Preserve arr_varImg(I_ELEMS, lngE)
                  ' *************************************************
                  ' ** Array: arr_varImg()
                  ' **
                  ' **   Field  Element  Name           Constant
                  ' **   =====  =======  =============  ===========
                  ' **     1       0     ctlspec_id     I_CSID
                  ' **     2       1     frm_id         I_FRMID
                  ' **     3       2     ctl_id         I_CTLID
                  ' **     4       3     ctl_name       I_CTLNAM
                  ' **     5       4     xadgfx_id      I_GFXID
                  ' **     6       5     xadgfx_name    I_GFXNAM
                  ' **     7       6     Found          I_FND
                  ' **
                  ' *************************************************
11470             arr_varImg(I_CSID, lngE) = ![ctlspec_id]
11480             arr_varImg(I_FRMID, lngE) = ![frm_id]
11490             arr_varImg(I_CTLID, lngE) = ![ctl_id]
11500             arr_varImg(I_CTLNAM, lngE) = ![ctl_name]
11510             arr_varImg(I_GFXID, lngE) = ![xadgfx_id]
11520             arr_varImg(I_GFXNAM, lngE) = ![xadgfx_name]
11530             arr_varImg(I_FND, lngE) = CBool(False)
11540             If lngX < lngRecs Then .MoveNext
11550           Next  ' ** lngX.
11560         End If
11570         .Close
11580       End With  ' ** rst.

11590       .Close
11600     End With  ' ** dbs.

          ' ** ctl_name = fld_name
11610     For lngX = 0& To (lngImgs - 1&)
11620       For lngY = 0& To (lngQryFlds - 1&)
11630         If arr_varQryFld(QF_NAM, lngY) = arr_varImg(I_CTLNAM, lngX) Then
11640           If arr_varImg(I_FND, lngX) = True Then
11650             Debug.Print "'FOUND TWICE! " & arr_varQryFld(QF_NAM, lngY)
11660           End If
11670           arr_varImg(I_FND, lngX) = CBool(True)
11680           Exit For
11690         End If
11700       Next  ' ** lngY.
11710     Next  ' ** lngX.

11720     For lngX = 0& To (lngImgs - 1&)
11730       If blnListImages = True Then
11740         Debug.Print "'I" & Left(CStr(lngX + 1&) & "  ", 2) & "  " & arr_varImg(I_CTLNAM, lngX)
11750       End If
11760       If arr_varImg(I_FND, lngX) = False Then
11770         Debug.Print "'NOT FOUND!  " & arr_varImg(I_CTLNAM, lngX)
11780       End If
11790     Next  ' ** lngX.

11800     If blnListImages = True Then
11810       Debug.Print "'IMAGES: " & CStr(lngRecs)
11820     End If
11830     DoEvents

11840     lngY = 0&
11850     For lngX = 0& To (lngQryFlds - 1&)
11860       strTmp01 = vbNullString
11870       intPos01 = InStr(arr_varQryFld(QF_NAM, lngX), "_")
11880       If intPos01 > 0 Then
11890         strTmp01 = Left(arr_varQryFld(QF_NAM, lngX), (intPos01 - 1))
11900       Else
11910         strTmp01 = arr_varQryFld(QF_NAM, lngX)
11920       End If
11930       Select Case strTmp01
            Case "frmgfx", "frm", "ctl", "xadgfx"
              ' ** Skip these.
11940       Case Else
11950         lngY = lngY + 1&
11960         If blnListButtons = True Then
11970           Debug.Print "'B" & Left(CStr(lngY) & "  ", 2) & "  " & arr_varQryFld(QF_NAM, lngX)
11980         End If
11990       End Select
12000     Next  ' ** lngX.

12010     If blnListButtons = True Then
12020       Debug.Print "'BTN NAMES: " & CStr(lngY)
12030     End If
12040     DoEvents

12050     For lngX = 0& To (lngQryFlds - 1&)
12060       strTmp01 = vbNullString
12070       intPos01 = InStr(arr_varQryFld(QF_NAM, lngX), "_")
12080       If intPos01 > 0 Then
12090         strTmp01 = Left(arr_varQryFld(QF_NAM, lngX), (intPos01 - 1))
12100       Else
12110         strTmp01 = arr_varQryFld(QF_NAM, lngX)
12120       End If
12130       Select Case strTmp01
            Case "frmgfx", "frm", "ctl", "xadgfx"
              ' ** Skip these.
12140       Case Else
12150         For lngY = 0& To (lngImgs - 1&)
12160           If arr_varImg(I_CTLNAM, lngY) = arr_varQryFld(QF_NAM, lngX) Then
12170             arr_varQryFld(QF_FND, lngX) = True
12180             Exit For
12190           End If
12200         Next
12210         If arr_varQryFld(QF_FND, lngX) = False Then
12220           If arr_varQryFld(QF_NAM, lngX) <> "frmJournal_Columns_Sub_lbl_raised_img2" Then
12230             Debug.Print "'NOT FOUND!  " & arr_varQryFld(QF_NAM, lngX)
12240           End If
12250         End If
12260       End Select
12270     Next  ' ** lngX.

12280   End If  ' ** lngQryFlds

12290   If blnListFields = True Then
12300     lngFlds = 0&
12310     ReDim arr_varFld(F_ELEMS, 0)
12320     strTmp01 = "qryJournal_Columns_18"
12330     blnRetVal = Tbl_Fld_List(arr_varFld, strTmp01)  ' ** Module Function: modXAdminFuncs.
12340     If blnRetVal = True Then
12350       lngFlds = (UBound(arr_varFld, 2) + 1&)
            ' *****************************************************
            ' ** Array: arr_varFld().
            ' **
            ' **   Field  Element  Name                Constant
            ' **   =====  =======  ==================  ==========
            ' **     1       0     qryfld_name         F_NAM
            ' **     2       1     datatype_db_type    F_TYP
            ' **     3       2     fld_required        F_REQ
            ' **
            ' *****************************************************
12360       Debug.Print "'QRY: " & strTmp01 & "  FLDS: " & CStr(lngFlds)
12370       For lngX = 0& To (lngFlds - 1&)
12380         Debug.Print "'F" & Left(CStr(lngX + 1&) & "   ", 3) & "  " & arr_varFld(F_NAM, lngX)
12390       Next
12400     End If
12410   End If

        'QRY: qryJournal_Columns_18  FLDS: 147
        'F1    frmgfx_id
        'F2    frm_id
        'F3    frm_name
        'F4    ctl_id_01
        'F5    ctl_name_01
        'F6    xadgfx_id_01
        'F7    cmdSpecPurp_Div_Map_raised_img
        'F8    ctl_id_02
        'F9    ctl_name_02
        'F10   xadgfx_id_02
        'F11   cmdSpecPurp_Div_Map_raised_focus_dots_img
        'F12   ctl_id_03
        'F13   ctl_name_03
        'F14   xadgfx_id_03
        'F15   cmdSpecPurp_Div_Map_sunken_focus_dots_img
        'F16   ctl_id_04
        'F17   ctl_name_04
        'F18   xadgfx_id_04
        'F19   cmdSpecPurp_Div_Map_raised_img_dis
        'F20   ctl_id_05
        'F21   ctl_name_05
        'F22   xadgfx_id_05
        'F23   cmdSpecPurp_Int_Map_raised_img
        'F24   ctl_id_06
        'F25   ctl_name_06
        'F26   xadgfx_id_06
        'F27   cmdSpecPurp_Int_Map_raised_focus_dots_img
        'F28   ctl_id_07
        'F29   ctl_name_07
        'F30   xadgfx_id_07
        'F31   cmdSpecPurp_Int_Map_sunken_focus_dots_img
        'F32   ctl_id_08
        'F33   ctl_name_08
        'F34   xadgfx_id_08
        'F35   cmdSpecPurp_Int_Map_raised_img_dis
        'F36   ctl_id_09
        'F37   ctl_name_09
        'F38   xadgfx_id_09
        'F39   cmdSpecPurp_Misc_MapLTCG_raised_img
        'F40   ctl_id_10
        'F41   ctl_name_10
        'F42   xadgfx_id_10
        'F43   cmdSpecPurp_Misc_MapLTCG_raised_focus_dots_img
        'F44   ctl_id_11
        'F45   ctl_name_11
        'F46   xadgfx_id_11
        'F47   cmdSpecPurp_Misc_MapLTCG_sunken_focus_dots_img
        'F48   ctl_id_12
        'F49   ctl_name_12
        'F50   xadgfx_id_12
        'F51   cmdSpecPurp_Misc_MapLTCG_raised_img_dis
        'F52   ctl_id_13
        'F53   ctl_name_13
        'F54   xadgfx_id_13
        'F55   cmdSpecPurp_Purch_MapSplit_raised_img
        'F56   ctl_id_14
        'F57   ctl_name_14
        'F58   xadgfx_id_14
        'F59   cmdSpecPurp_Purch_MapSplit_raised_focus_dots_img
        'F60   ctl_id_15
        'F61   ctl_name_15
        'F62   xadgfx_id_15
        'F63   cmdSpecPurp_Purch_MapSplit_sunken_focus_dots_img
        'F64   ctl_id_16
        'F65   ctl_name_16
        'F66   xadgfx_id_16
        'F67   cmdSpecPurp_Purch_MapSplit_raised_img_dis
        'F68   ctl_id_17
        'F69   ctl_name_17
        'F70   xadgfx_id_17
        'F71   cmdSpecPurp_Sold_PaidTotal_raised_img
        'F72   ctl_id_18
        'F73   ctl_name_18
        'F74   xadgfx_id_18
        'F75   cmdSpecPurp_Sold_PaidTotal_raised_focus_dots_img
        'F76   ctl_id_19
        'F77   ctl_name_19
        'F78   xadgfx_id_19
        'F79   cmdSpecPurp_Sold_PaidTotal_sunken_focus_dots_img
        'F80   ctl_id_20
        'F81   ctl_name_20
        'F82   xadgfx_id_20
        'F83   cmdSpecPurp_Sold_PaidTotal_raised_img_dis
        'F84   ctl_id_21
        'F85   ctl_name_21
        'F86   xadgfx_id_21
        'F87   frmJournal_Columns_Sub_lbl_raised_img
        'F88   ctl_id_22
        'F89   ctl_name_22
        'F90   xadgfx_id_22
        'F91   frmJournal_Columns_Sub_lbl_raised_img_dis
        'F92   ctl_id_23
        'F93   ctl_name_23
        'F94   xadgfx_id_23
        'F95   cmdScroll_lbl_raised_img
        'F96   ctl_id_24
        'F97   ctl_name_24
        'F98   xadgfx_id_24
        'F99   cmdScroll_lbl_raised_focus_dots_img
        'F100  ctl_id_25
        'F101  ctl_name_25
        'F102  xadgfx_id_25
        'F103  cmdScrollLeft_raised_img
        'F104  ctl_id_26
        'F105  ctl_name_26
        'F106  xadgfx_id_26
        'F107  cmdScrollLeft_sunken_img
        'F108  ctl_id_27
        'F109  ctl_name_27
        'F110  xadgfx_id_27
        'F111  cmdScrollRight_raised_img
        'F112  ctl_id_28
        'F113  ctl_name_28
        'F114  xadgfx_id_28
        'F115  cmdScrollRight_sunken_img
        'F116  ctl_id_29
        'F117  ctl_name_29
        'F118  xadgfx_id_29
        'F119  cmdUncomComAll_raised_img
        'cmdUncomComAll_raised_semifocus_dots_img
        'cmdUncomComAll_raised_focus_img
        'F120  ctl_id_30
        'F121  ctl_name_30
        'F122  xadgfx_id_30
        'F123  cmdUncomComAll_raised_focus_dots_img
        'F124  ctl_id_31
        'F125  ctl_name_31
        'F126  xadgfx_id_31
        'F127  cmdUncomComAll_sunken_focus_dots_img
        'F128  ctl_id_32
        'F129  ctl_name_32
        'F130  xadgfx_id_32
        'F131  cmdUncomComAll_raised_img_dis
        'F132  ctl_id_33
        'F133  ctl_name_33
        'F134  xadgfx_id_33
        'F135  cmdUncomDelAll_raised_img
        'cmdUncomDelAll_raised_semifocus_dots_img
        'cmdUncomDelAll_raised_focus_img
        'F136  ctl_id_34
        'F137  ctl_name_34
        'F138  xadgfx_id_34
        'F139  cmdUncomDelAll_raised_focus_dots_img
        'F140  ctl_id_35
        'F141  ctl_name_35
        'F142  xadgfx_id_35
        'F143  cmdUncomDelAll_sunken_focus_dots_img
        'F144  ctl_id_36
        'F145  ctl_name_36
        'F146  xadgfx_id_36
        'F147  cmdUncomDelAll_raised_img_dis

        'IMAGES: 36
        'I1   cmdSpecPurp_Div_Map_raised_img
        'I2   cmdSpecPurp_Div_Map_raised_focus_dots_img
        'I3   cmdSpecPurp_Div_Map_sunken_focus_dots_img
        'I4   cmdSpecPurp_Div_Map_raised_img_dis
        'I5   cmdSpecPurp_Int_Map_raised_img
        'I6   cmdSpecPurp_Int_Map_raised_focus_dots_img
        'I7   cmdSpecPurp_Int_Map_sunken_focus_dots_img
        'I8   cmdSpecPurp_Int_Map_raised_img_dis
        'I9   cmdSpecPurp_Misc_MapLTCG_raised_img
        'I10  cmdSpecPurp_Misc_MapLTCG_raised_focus_dots_img
        'I11  cmdSpecPurp_Misc_MapLTCG_sunken_focus_dots_img
        'I12  cmdSpecPurp_Misc_MapLTCG_raised_img_dis
        'I13  cmdSpecPurp_Purch_MapSplit_raised_img
        'I14  cmdSpecPurp_Purch_MapSplit_raised_focus_dots_img
        'I15  cmdSpecPurp_Purch_MapSplit_sunken_focus_dots_img
        'I16  cmdSpecPurp_Purch_MapSplit_raised_img_dis
        'I17  cmdSpecPurp_Sold_PaidTotal_raised_img
        'I18  cmdSpecPurp_Sold_PaidTotal_raised_focus_dots_img
        'I19  cmdSpecPurp_Sold_PaidTotal_sunken_focus_dots_img
        'I20  cmdSpecPurp_Sold_PaidTotal_raised_img_dis
        'I21  cmdScroll_lbl_raised_img
        'I22  cmdScroll_lbl_raised_focus_dots_img
        'I23  cmdScrollLeft_raised_img
        'I24  cmdScrollLeft_sunken_img
        'I25  cmdScrollRight_raised_img
        'I26  cmdScrollRight_sunken_img
        'I27  frmJournal_Columns_Sub_lbl_raised_img
        'I28  frmJournal_Columns_Sub_lbl_raised_img_dis
        'I29  cmdUncomComAll_raised_img
        'I30  cmdUncomComAll_raised_focus_dots_img
        'I31  cmdUncomComAll_sunken_focus_dots_img
        'I32  cmdUncomComAll_raised_img_dis
        'I33  cmdUncomDelAll_raised_img
        'I34  cmdUncomDelAll_raised_focus_dots_img
        'I35  cmdUncomDelAll_sunken_focus_dots_img
        'I36  cmdUncomDelAll_raised_img_dis

        'BTN NAMES: 36
        'B1   cmdSpecPurp_Div_Map_raised_img
        'B2   cmdSpecPurp_Div_Map_raised_focus_dots_img
        'B3   cmdSpecPurp_Div_Map_sunken_focus_dots_img
        'B4   cmdSpecPurp_Div_Map_raised_img_dis
        'B5   cmdSpecPurp_Int_Map_raised_img
        'B6   cmdSpecPurp_Int_Map_raised_focus_dots_img
        'B7   cmdSpecPurp_Int_Map_sunken_focus_dots_img
        'B8   cmdSpecPurp_Int_Map_raised_img_dis
        'B9   cmdSpecPurp_Misc_MapLTCG_raised_img
        'B10  cmdSpecPurp_Misc_MapLTCG_raised_focus_dots_img
        'B11  cmdSpecPurp_Misc_MapLTCG_sunken_focus_dots_img
        'B12  cmdSpecPurp_Misc_MapLTCG_raised_img_dis
        'B13  cmdSpecPurp_Purch_MapSplit_raised_img
        'B14  cmdSpecPurp_Purch_MapSplit_raised_focus_dots_img
        'B15  cmdSpecPurp_Purch_MapSplit_sunken_focus_dots_img
        'B16  cmdSpecPurp_Purch_MapSplit_raised_img_dis
        'B17  cmdSpecPurp_Sold_PaidTotal_raised_img
        'B18  cmdSpecPurp_Sold_PaidTotal_raised_focus_dots_img
        'B19  cmdSpecPurp_Sold_PaidTotal_sunken_focus_dots_img
        'B20  cmdSpecPurp_Sold_PaidTotal_raised_img_dis
        'B21  frmJournal_Columns_Sub_lbl_raised_img
        'B22  frmJournal_Columns_Sub_lbl_raised_img_dis
        'B23  cmdScroll_lbl_raised_img
        'B24  cmdScroll_lbl_raised_focus_dots_img
        'B25  cmdScrollLeft_raised_img
        'B26  cmdScrollLeft_sunken_img
        'B27  cmdScrollRight_raised_img
        'B28  cmdScrollRight_sunken_img
        'B29  cmdUncomComAll_raised_img
        'B30  cmdUncomComAll_raised_focus_dots_img
        'B31  cmdUncomComAll_sunken_focus_dots_img
        'B32  cmdUncomComAll_raised_img_dis
        'B33  cmdUncomDelAll_raised_img
        'B34  cmdUncomDelAll_raised_focus_dots_img
        'B35  cmdUncomDelAll_sunken_focus_dots_img
        'B36  cmdUncomDelAll_raised_img_dis

12420   Beep

EXITP:
12430   Set rst = Nothing
12440   Set qdf = Nothing
12450   Set dbs = Nothing
12460   JC_Msc_Gfx_Load2 = blnRetVal
12470   Exit Function

ERRH:
12480   blnRetVal = False
12490   Select Case ERR.Number
        Case Else
12500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12510   End Select
12520   Resume EXITP

End Function

Public Sub JC_Msc_ChkAverage(frmSub As Access.Form)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_ChkAverage"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String, lngAssetNo As Long

12610   With frmSub
12620     .IsAverage = False
12630     If IsNull(.assetno) = False Then
12640       Select Case .journaltype
            Case "Deposit", "Purchase", "Withdrawn", "Sold", "Liability", "Liability (+)", "Liability (-)", "Cost Adj."
12650         strAccountNo = .accountno
12660         lngAssetNo = .assetno
12670         Set dbs = CurrentDb
              ' ** ActiveAssets, just IsAverage = True, grouped, by accountno, assetno, IsAverage, with cnt.
12680         Set qdf = dbs.QueryDefs("qryJournal_Columns_10_MasterAsset_22")
12690         Set rst = qdf.OpenRecordset
12700         If rst.BOF = True And rst.EOF = True Then
                ' ** No averaging anywhere.
12710           rst.Close
12720         Else
12730           rst.Close
                ' ** ActiveAssets, grouped, by accountno, assetno, IsAverage,
                ' ** with cnt, by specified [actno], [astno].
12740           Set qdf = dbs.QueryDefs("qryJournal_Columns_10_MasterAsset_23")
12750           With qdf.Parameters
12760             ![actno] = strAccountNo
12770             ![astno] = lngAssetNo
12780           End With
12790           Set rst = qdf.OpenRecordset
12800           If rst.BOF = True And rst.EOF = True Then
                  ' ** This asset not averaged for this account.
12810             rst.Close
12820           Else
12830             rst.MoveFirst
12840             .IsAverage = True
12850             rst.Close
12860           End If
12870         End If
12880         dbs.Close
12890       End Select
12900     End If
12910   End With

EXITP:
12920   Set rst = Nothing
12930   Set qdf = Nothing
12940   Set dbs = Nothing
12950   Exit Sub

ERRH:
12960   Select Case ERR.Number
        Case Else
12970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12980   End Select
12990   Resume EXITP

End Sub

Public Function JC_Msc_Pub_Reset(frmPar As Access.Form, frmSub As Access.Form) As Boolean
' ** Check and reset any variables that may have been lost.

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Pub_Reset"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim grp As DAO.Group, usr As DAO.User
        Dim datPostingDate As Date
        Dim blnAdd As Boolean, blnEdit As Boolean
        Dim varTmp00 As Variant, blnTmp01 As Boolean, lngTmp02 As Long
        Dim blnRetVal As Boolean

13010   blnRetVal = True

13020   With frmSub

          ' ** datPostingDate.
13030     blnAdd = False: blnEdit = False
13040     datPostingDate = .PostDate_GetSet(True)  ' ** Form Function: frmJournal_Columns_Sub.
13050     If datPostingDate = 0 Then  ' ** 12/30/1899.
13060       varTmp00 = DLookup("Posting_Date", "PostingDate", "[Username] = '" & CurrentUser & "'")  ' ** Internal Access Function: Trust Accountant login.
13070       If IsNull(varTmp00) = False Then
13080         datPostingDate = CDate(varTmp00)
13090         If datPostingDate = 0 Then
13100           datPostingDate = Date
13110           blnEdit = True
13120         End If
13130       Else
13140         blnAdd = True
13150       End If
13160       .PostDate_GetSet False, datPostingDate  ' ** Form Procedure: frmJournal_Columns_Sub.
13170       If blnAdd = True Then
13180         Set dbs = CurrentDb
13190         With dbs
                ' ** Append new record to PostingDate table, by specified [usr].
13200           Set qdf = .QueryDefs("qryPostingDate_02")
13210           With qdf.Parameters
13220             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
13230           End With
13240           qdf.Execute dbFailOnError
13250           Set qdf = Nothing
                ' ** PostingDate, by specified [usr].
13260           Set qdf = .QueryDefs("qryPostingDate_07")
13270           With qdf.Parameters
13280             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
13290           End With
13300           Set rst = qdf.OpenRecordset
13310           With rst
13320             .MoveFirst
13330             glngPostingDateID = ![PostingDate_ID]
13340             .Close
13350           End With
13360           Set rst = Nothing
13370           Set qdf = Nothing
                ' ** Update tblCalendar_Staging, for unique_id, by specified [unqid].
13380           Set qdf = .QueryDefs("qryCalendar_02")
13390           With qdf.Parameters
13400             ![unqid] = glngPostingDateID
13410           End With
13420           qdf.Execute
13430           Set qdf = Nothing
13440           .Close
13450         End With
13460       ElseIf blnEdit = True Then
13470         Set dbs = CurrentDb
13480         With dbs
                ' ** Update PostingDate, by specified [usr], [pdat].
13490           Set qdf = .QueryDefs("qryPostingDate_06")
13500           With qdf.Parameters
13510             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
13520             ![pdat] = datPostingDate
13530           End With
13540           qdf.Execute dbFailOnError
13550           Set qdf = Nothing
                ' ** PostingDate, by specified [usr].
13560           Set qdf = .QueryDefs("qryPostingDate_07")
13570           With qdf.Parameters
13580             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
13590           End With
13600           Set rst = qdf.OpenRecordset
13610           With rst
13620             .MoveFirst
13630             glngPostingDateID = ![PostingDate_ID]
13640             .Close
13650           End With
13660           Set rst = Nothing
13670           Set qdf = Nothing
                ' ** Update tblCalendar_Staging, for unique_id, by specified [unqid].
13680           Set qdf = .QueryDefs("qryCalendar_02")
13690           With qdf.Parameters
13700             ![unqid] = glngPostingDateID
13710           End With
13720           qdf.Execute
13730           Set qdf = Nothing
13740           .Close
13750         End With
13760       End If
13770     End If

          ' ** clsMonthClass
13780     .CalendarCheck  ' ** Form Procedure: frmJournal_Columns_Sub.

          ' ** Array: arr_varStat().
13790     lngTmp02 = .StatusBar_Get  ' ** Form Function: frmJournal_Columns_Sub.
13800     If lngTmp02 = 0& Then
13810       .StatusBar_Load  ' ** Form Procedure: frmJournal_Columns_Sub.
13820     End If

          ' ** CLR_HILITE, CLR_DISABLED_FG, CLR_DISABLED_BG
13830     .ColorCheck  ' ** Form Procedure: frmJournal_Columns_Sub.

13840   End With  ' ** frmSub.

13850   With frmPar

          ' ** gstrTrustDataLocation.
13860     If gstrTrustDataLocation = vbNullString Then
13870       IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
13880     End If

          ' ** gblnAdmin.
13890     gblnAdmin = False
13900     For Each grp In DBEngine.Workspaces(0).Groups
13910       If grp.Name = "Admins" Then
13920         For Each usr In grp.Users
13930           If usr.Name = CurrentUser Then  ' ** Internal Access Function: Trust Accountant login.
13940             gblnAdmin = True
13950             Exit For
13960           End If
13970         Next
13980       End If
13990     Next
14000     Set usr = Nothing
14010     Set grp = Nothing

          ' ** CoOptions.
14020     If gstrCo_Name = vbNullString Then
            ' ** Of course, I've seen user's that don't put their name here.
14030       blnTmp01 = CoOptions_Read  ' ** Module Function: modStartupFuncs.
14040     End If

          ' ** gblnSingleUser.
14050     gblnSingleUser = IsSingleUser  ' ** Module Function: modSecurityFunctions.

          ' ** gblnLocalData.
14060     gblnLocalData = IsLocalData  ' ** Module Function: modStartupFuncs.

          ' ** gstrFormQuerySpec.
14070     If gstrFormQuerySpec = vbNullString Then
14080       gstrFormQuerySpec = .Name
14090     End If

14100     blnTmp01 = MouseWheelON  ' ** Module Function: modMouseWheel

          ' ** THIS_NAME_SUB.
14110     .ThisNameSub_Set  ' ** Form Procedure: frmJournal_Columns.

14120   End With  ' ** frmPar.

EXITP:
14130   Set usr = Nothing
14140   Set grp = Nothing
14150   Set rst = Nothing
14160   Set qdf = Nothing
14170   Set dbs = Nothing
14180   JC_Msc_Pub_Reset = blnRetVal
14190   Exit Function

ERRH:
14200   blnRetVal = False
14210   Select Case ERR.Number
        Case Else
14220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14230   End Select
14240   Resume EXITP

End Function

Public Sub JC_Msc_Clear()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_Clear"

14310   lngStats = 0&
14320   ReDim arr_varStat(S_ELEMS, 0)

EXITP:
14330   Exit Sub

ERRH:
14340   Select Case ERR.Number
        Case Else
14350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14360   End Select
14370   Resume EXITP

End Sub

Public Sub JC_Msc_RevCodeDescDisp(blnNoMove As Boolean, blnNotPopup As Boolean, frm As Access.Form)

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_RevCodeDescDisp"

        Dim varTmp00 As Variant

14410   With frm
14420     JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, frm  ' ** Procedure: Above.
14430     blnNoMove = False
14440 On Error Resume Next
14450     varTmp00 = .revcode_DESC
14460     If ERR.Number <> 0 Then
14470 On Error GoTo ERRH
14480     Else
14490 On Error GoTo ERRH
14500       If Nz(.revcode_DESC, vbNullString) = vbNullString Then
14510         .revcode_ID.SetFocus
14520       Else
14530         Select Case .revcode_ID.RowSource
              Case "qryJournal_Columns_10_M_RevCode_02"
                ' ** INCOME.
14540           If .revcode_ID.Column(2) = REVTYP_EXP Then
14550             Select Case .journaltype
                  Case "Dividend"
14560               .revcode_ID = REVID_ORDDIV  ' ** Ordinary Dividend.
14570               .revcode_DESC = "Ordinary Dividend"
14580               .revcode_DESC_display = "Ordinary Dividend"
14590             Case "Interest"
14600               .revcode_ID = REVID_INTINC  ' ** Interest Income.
14610               .revcode_DESC = "Interest Income"
14620               .revcode_DESC_display = "Interest Income"
14630             Case Else
14640               .revcode_ID = REVID_INC  ' ** Unspecified Income.
14650               .revcode_DESC = "Unspecified Income"
14660               .revcode_DESC_display = vbNullString  ' ** This does allow zero-length strings.
14670             End Select
14680             .revcode_ID.SetFocus
14690           Else
14700             .revcode_ID.SetFocus
14710           End If
14720         Case "qryJournal_Columns_10_M_RevCode_03"
                ' ** EXPENSE.
14730           If .revcode_ID.Column(2) = REVTYP_INC Then
14740             .revcode_ID = REVID_EXP  ' ** Unspecified Expense.
14750             .revcode_DESC = "Unspecified Expense"
14760             .revcode_DESC_display = vbNullString  ' ** This does allow zero-length strings.
14770             .revcode_ID.SetFocus
14780           Else
14790             .revcode_ID.SetFocus
14800           End If
14810         Case "qryJournal_Columns_10_M_RevCode_01"
                ' ** ALL.
14820           .revcode_ID.SetFocus
14830         End Select
14840       End If
14850     End If
14860   End With

EXITP:
14870   Exit Sub

ERRH:
14880   Select Case ERR.Number
        Case Else
14890     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14900   End Select
14910   Resume EXITP

End Sub

Public Sub JC_Msc_TaxCodeDescDisp(blnNoMove As Boolean, blnNotPopup As Boolean, frm As Access.Form)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_TaxCodeDescDisp"

        Dim varTmp00 As Variant

15010   With frm
15020     JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, frm  ' ** Module Procedure: modJrnlCol_Misc.
15030     blnNoMove = False
15040 On Error Resume Next
15050     varTmp00 = .taxcode_description
15060     If ERR.Number <> 0 Then
15070 On Error GoTo ERRH
15080     Else
15090 On Error GoTo ERRH
15100       If Nz(.taxcode_description, vbNullString) = vbNullString Then
15110         .taxcode.SetFocus
15120       Else
15130         Select Case .taxcode.RowSource
              Case "qryJournal_Columns_10_TaxCode_02"
                ' ** INCOME.
15140           If .taxcode_type = TAXTYP_DED Then
15150             .taxcode = TAXID_INC
15160             .taxcode_description = "Unspecified Income"
15170             .taxcode_description_display = vbNullString
15180             .taxcode.SetFocus
15190           Else
15200             .taxcode.SetFocus
15210           End If
15220         Case "qryJournal_Columns_10_TaxCode_03", "qryJournal_Columns_10_TaxCode_04"
                ' ** EXPENSE.
15230           If .taxcode_type = TAXTYP_INC Then
15240             .taxcode = TAXID_DED
15250             .taxcode_description = "Unspecified Deduction"
15260             .taxcode_description_display = vbNullString
15270             .taxcode.SetFocus
15280           Else
15290             .taxcode.SetFocus
15300           End If
15310         Case "qryJournal_Columns_10_TaxCode_01"
                ' ** ALL.
15320           .taxcode.SetFocus
15330         End Select
15340       End If
15350     End If
15360   End With

EXITP:
15370   Exit Sub

ERRH:
15380   Select Case ERR.Number
        Case Else
15390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15400   End Select
15410   Resume EXITP

End Sub

Public Sub JC_Msc_LocNameDisp(blnNoMove As Boolean, blnNotPopup As Boolean, frm As Access.Form)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Msc_LocNameDisp"

        Dim varTmp00 As Variant

15510   With frm
15520     JC_Msc_StatusBar_SetSub THIS_PROC, blnNotPopup, frm  ' ** Module Procedure: modJrnlCol_Misc.
15530     blnNoMove = False
15540 On Error Resume Next
15550     varTmp00 = .Loc_Name
15560     If ERR.Number <> 0 Then
15570 On Error GoTo ERRH
15580     Else
15590 On Error GoTo ERRH
15600       If IsNull(.Location_ID) = True Then
15610         .Location_ID = 1&  ' ** {Unassigned}, {no entry}.
15620         .Loc_Name = "{no entry}"
15630         .Loc_Name_display = vbNullString  ' ** This does allow zero-length strings.
15640       End If
15650       .Location_ID.SetFocus
15660     End If
15670   End With

EXITP:
15680   Exit Sub

ERRH:
15690   Select Case ERR.Number
        Case Else
15700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15710   End Select
15720   Resume EXITP

End Sub
