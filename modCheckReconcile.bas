Attribute VB_Name = "modCheckReconcile"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCheckReconcile"

'VGC 10/09/2017: CHANGES!

' ** Array: arr_varAsset_ThisAcct().
'Private lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant
Private Const A_CRACTID As Integer = 0
'Private Const A_ACTNO   As Integer = 1
Private Const A_ASTNO   As Integer = 2
Private Const A_CUSIP   As Integer = 3

' ** Combo box column constants: cmbAccounts.
Private Const CBX_ACT_ACTNO   As Integer = 0
Private Const CBX_ACT_DESC    As Integer = 1
'Private Const CBX_ACT_STMDATE As Integer = 2
'Private Const CBX_ACT_SHORT   As Integer = 3
'Private Const CBX_ACT_LEGAL   As Integer = 4
'Private Const CBX_ACT_BALDATE As Integer = 5
'Private Const CBX_ACT_HASREL  As Integer = 6
'Private Const CBX_ACT_DEFAST  As Integer = 7

Private lngChks As Long, arr_varChk() As Variant
Private Const C_ELEMS As Integer = 12  ' ** Array's first-element UBound().
Private Const C_ACTID  As Integer = 0
Private Const C_STGID  As Integer = 1
Private Const C_JNO    As Integer = 2
Private Const C_ACTNO  As Integer = 3
Private Const C_CHKID  As Integer = 4
Private Const C_CNUM   As Integer = 5
Private Const C_PAID   As Integer = 6
Private Const C_DESC   As Integer = 7
Private Const C_HDESC  As Integer = 8
Private Const C_ASGN   As Integer = 9
Private Const C_ASTNO1 As Integer = 10
Private Const C_ASTNO2 As Integer = 11
Private Const C_FND    As Integer = 12

Private lngTpp As Long, lngRecsCur As Long
Private lngOutChecksBox_Left As Long
' **

Public Sub UpdateStageTbl(lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant, dbs As DAO.Database)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateStageTbl"

        Dim qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnFound As Boolean
        Dim intPos01 As Integer
        Dim varTmp00 As Variant, strTmp02 As String, strTmp03 As String
        Dim lngX As Long, lngY As Long

        ' ** Array: arr_varAsset_ThisAcct().
        'Const A_CRACTID As Integer = 0
        'Const A_ACTNO   As Integer = 1
        Const A_ASTNO   As Integer = 2
        Const A_CUSIP   As Integer = 3

        Const ASSIGN_MSG As String = "Associated with CUSIP:"

110     With dbs
          ' ** tblCheckReconcile_Staging, just description <> Null.  '## OK
120       Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_24")
130       Set rst = qdf.OpenRecordset
140       With rst
150         If .BOF = True And .EOF = True Then
              ' ** No multi-asset assignments. (Or anything else, for that matter.)
160         Else
170           .MoveLast
180           lngRecs = .RecordCount
190           .MoveFirst
200           For lngX = 1& To lngRecs
210             strTmp02 = ![description]
220             strTmp03 = vbNullString
230             intPos01 = InStr(strTmp02, ASSIGN_MSG)
240             If intPos01 > 0 Then
250               If intPos01 > 1 Then strTmp03 = Left(strTmp02, (intPos01 - 1))  ' ** Save this for later.
260               strTmp02 = Mid(strTmp02, intPos01)  ' ** Skips any other preceding, unrelated message.
270               intPos01 = InStr(strTmp02, ";")
280               If intPos01 > 0 Then
290                 If strTmp03 = vbNullString Then
300                   strTmp03 = Mid(strTmp02, (intPos01 + 1))  ' ** Save this for later.
310                 Else
320                   strTmp03 = strTmp03 & ";" & Mid(strTmp02, (intPos01 + 1))  ' ** Save these for later.
330                 End If
340                 strTmp02 = Left(strTmp02, (intPos01 - 1))  ' ** Some other, unrelated message.
350               End If
360               intPos01 = InStr(strTmp02, ":")
370               If intPos01 > 0 Then
380                 If lngAssets_ThisAcct > 1& Then
390                   strTmp02 = Trim(Mid(strTmp02, (intPos01 + 1)))  ' ** Just the CUSIP.
400                   varTmp00 = DLookup("[assetno]", "masterasset", "[cusip] = '" & strTmp02 & "'")
410                   Select Case IsNull(varTmp00)
                      Case True
                        ' ** CUSIP must have changed!
420                     .Edit
430                     If strTmp03 <> vbNullString Then
440                       ![description] = strTmp03  'HOW DO I TELL IT TO UPDATE THIS!
                          '[description_new]
450                     Else
460                       ![description] = Null  'HOW DO I TELL IT TO UPDATE THIS!
                          '[description_new]
470                     End If
480                     ![crstage_assign] = 0
490                     ![crstage_asset1] = 0&
500                     ![crstage_asset2] = 0&
510                     ![crstage_datemodified] = Now()
520                     .Update
530                   Case False
540                     blnFound = False
550                     For lngY = 0& To (lngAssets_ThisAcct - 1&)
560                       If arr_varAsset_ThisAcct(A_CUSIP, lngY) = strTmp02 Then
570                         blnFound = True
580                         .Edit
590                         Select Case lngY
                            Case 0&
600                           ![crstage_assign] = 1&
610                           ![crstage_asset1] = arr_varAsset_ThisAcct(A_ASTNO, lngY)
620                           ![crstage_asset2] = 0&
630                         Case 1&
640                           ![crstage_assign] = 2
650                           ![crstage_asset1] = 0&
660                           ![crstage_asset2] = arr_varAsset_ThisAcct(A_ASTNO, lngY)
670                         End Select
680                         ![crstage_datemodified] = Now()
690                         .Update  ' ** Description remains as-is.
700                         Exit For
710                       End If
720                     Next
730                     If blnFound = False Then
                          ' ** They've chosen a different asset.
740                       .Edit
750                       If strTmp03 <> vbNullString Then
760                         ![description] = strTmp03  'HOW DO I TELL IT TO UPDATE THIS!
                            '[description_new]
770                       Else
780                         ![description] = Null  'HOW DO I TELL IT TO UPDATE THIS!
                            '[description_new]
790                       End If
800                       ![crstage_assign] = 0
810                       ![crstage_asset1] = 0&
820                       ![crstage_asset2] = 0&
830                       ![crstage_datemodified] = Now()
840                       .Update
850                     End If
860                   End Select
870                 End If
880               Else
                    ' ** No longer a multi-asset account.
890                 .Edit
900                 If strTmp03 <> vbNullString Then
910                   ![description] = strTmp03  'HOW DO I TELL IT TO UPDATE THIS!
                      '[description_new]
920                 Else
930                   ![description] = Null  'HOW DO I TELL IT TO UPDATE THIS!
                      '[description_new]
940                 End If
950                 ![crstage_assign] = 0
960                 ![crstage_asset1] = 0&
970                 ![crstage_asset2] = 0&
980                 ![crstage_datemodified] = Now()
990                 .Update
1000              End If
1010            End If
1020            strTmp02 = vbNullString: strTmp03 = vbNullString
1030            If lngX < lngRecs Then .MoveNext
1040          Next
1050        End If
1060        .Close
1070      End With
1080    End With

        ' ** If description is Null, do nothing.
        ' ** If our 'Associated' msg is there, pull CUSIP.
        ' ** Match CUSIP TO one or the other of the two assets.
        ' ** If not a 2-asset account, Null out description, and Zero-out others.
        ' ** If CUSIP not found, Null out description, and Zero-out others.
        ' ** If CUSIP found, but doesn't match either asset, Null out description, and Zero-out others.

EXITP:
1090    Set rst = Nothing
1100    Set qdf = Nothing
1110    Exit Sub

ERRH:
1120    Select Case ERR.Number
        Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

End Sub

Public Sub SetMultiAsset(lngAssets As Long)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "SetMultiAsset"

        Dim frm As Access.Form
        Dim lngOutChksDiff_Width As Long, lngOutChksDiff_Left As Long
        Dim lngTransdate_Left As Long, lngIters As Long, lngOpgDiff_Left As Long, lngOpg_Width As Long
        Dim lngOff1 As Long, lngOff2 As Long, lngOff3 As Long, lngOff4 As Long, lngMod As Long
        Dim strTmp01 As String, dblTmp02 As Double
        Dim lngX As Long

1210    If IsLoaded("frmCheckReconcile", acForm) = True Then  ' ** Module Function: modFileUtilities.
1220      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
1230        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1240      End If
1250      Set frm = Forms("frmCheckReconcile")
1260      With frm
1270        lngOutChecksBox_Left = .OutChksBox_Left
1280        lngOutChksDiff_Left = (.frmCheckReconcile_Sub_OutChecks.Left - .OutChecks_box.Left)
1290        If lngAssets > 1& Then
1300          If .OutChecks_box.Left = lngOutChecksBox_Left Then
1310            lngOutChksDiff_Width = (lngOutChecksBox_Left - (.BSBalance_box.Left + 150&))
1320            .OutChecks_box.Left = .BSBalance_box.Left + 150&
1330            .OutChecks_box.Width = (.OutChecks_box.Width + lngOutChksDiff_Width)
1340            .frmCheckReconcile_Sub_OutChecks.Left = (.OutChecks_box.Left + lngOutChksDiff_Left)
1350            .frmCheckReconcile_Sub_OutChecks.Width = (.frmCheckReconcile_Sub_OutChecks.Width + lngOutChksDiff_Width)
1360            .frmCheckReconcile_Sub_OutChecks_box.Left = (.frmCheckReconcile_Sub_OutChecks.Left - lngTpp)
1370            .frmCheckReconcile_Sub_OutChecks_box.Width = (.frmCheckReconcile_Sub_OutChecks_box.Width + lngOutChksDiff_Width)
1380            .frmCheckReconcile_Sub_OutChecks_lbl.Left = (.frmCheckReconcile_Sub_OutChecks.Left - (2& * lngTpp))
1390            .frmCheckReconcile_Sub_OutChecks_lbl.Width = (.frmCheckReconcile_Sub_OutChecks_lbl.Width + lngOutChksDiff_Width)
1400            .OutChecks_Arrow_lbl.Left = (.OutChecks_Arrow_lbl.Left - lngOutChksDiff_Width)
1410            .OutChecks_Arrow_lbl_line1.Left = (.OutChecks_Arrow_lbl_line1.Left - lngOutChksDiff_Width)
1420            .OutChecks_Arrow_lbl_line2.Left = (.OutChecks_Arrow_lbl_line2.Left - lngOutChksDiff_Width)
1430            .OutChecksCnt_lbl.Left = .frmCheckReconcile_Sub_OutChecks.Left
1440            With .frmCheckReconcile_Sub_OutChecks.Form
1450              lngTransdate_Left = .transdate_left
1460              lngOpg_Width = .opgAssign.Width
                  ' ** The first time, move the Option Group to the left-hand side.
1470              If (.opgAssign.Left <> lngTransdate_Left) And (.opgAssign.Left > lngTransdate_Left) Then
1480                lngOff1 = (.opgAssign_optOne.Left - .opgAssign.Left)
1490                lngOff2 = (.opgAssign_optOne_lbl.Left - .opgAssign_optOne.Left)
1500                lngOff3 = (.opgAssign_optTwo.Left - .opgAssign.Left)
1510                lngOpgDiff_Left = (.opgAssign.Left - (lngTransdate_Left - 15&)) '30&))
1520                dblTmp02 = (lngOpgDiff_Left / lngTpp)
1530                lngMod = 0&
1540                If InStr(CStr(dblTmp02), ".") > 0 Then
                      ' ** Not an even division.
1550                  lngIters = Val(Left(CStr(dblTmp02), (InStr(CStr(dblTmp02), ".") - 1)))
1560                  lngMod = ((lngIters * lngTpp) - lngTransdate_Left)
1570                Else
1580                  lngIters = dblTmp02
1590                End If
1600                For lngX = 1& To lngIters
1610                  .opgAssign.Left = (.opgAssign.Left - lngTpp)
1620                  .opgAssign.Width = (.opgAssign.Width + lngTpp)
1630                Next
1640                If lngMod > 0& Then
1650                  .opgAssign.Left = (.opgAssign.Left - lngMod)
1660                  .opgAssign.Width = (.opgAssign.Width + lngMod)
1670                End If
1680                .opgAssign_optOne.Left = (.opgAssign.Left + lngOff1)
1690                .opgAssign_optOne_lbl.Left = (.opgAssign_optOne.Left + lngOff2)
1700                .opgAssign_optTwo.Left = (.opgAssign.Left + lngOff3)
1710                .opgAssign_optTwo_lbl.Left = (.opgAssign_optTwo.Left + lngOff2)
1720                .opgAssign.Width = lngOpg_Width
1730                .opgAssign_lbl.Left = .opgAssign.Left + lngTpp
1740                .opgAssign_lbl_line.Left = .opgAssign.Left + lngTpp
1750                .croutchk_asset1_Off.Left = .opgAssign_optOne_lbl.Left
1760                .croutchk_asset1_On.Left = .opgAssign_optOne_lbl.Left
1770                .croutchk_asset2_Off.Left = .opgAssign_optTwo_lbl.Left
1780                .croutchk_asset2_On.Left = .opgAssign_optTwo_lbl.Left
1790              End If
1800              If (.transdate.Left = lngTransdate_Left) Or (.transdate.Left < (.opgAssign.Left + .opgAssign.Width)) Then
1810                lngOff4 = (lngOpg_Width + (5& * lngTpp))
1820                .transdate.Left = (.transdate.Left + lngOff4)
1830                .transdate_lbl.Left = (.transdate_lbl.Left + lngOff4)
1840                .transdate_lbl_line.Left = (.transdate_lbl_line.Left + lngOff4)
1850                .CheckNum.Left = (.CheckNum.Left + lngOff4)
1860                .CheckNum_lbl.Left = (.CheckNum_lbl.Left + lngOff4)
1870                .CheckNum_lbl_line.Left = (.CheckNum_lbl_line.Left + lngOff4)
1880                .croutchk_payee.Left = (.croutchk_payee.Left + lngOff4)
1890                .croutchk_payee_lbl.Left = (.croutchk_payee_lbl.Left + lngOff4)
1900                .croutchk_payee_lbl_line.Left = (.croutchk_payee_lbl_line.Left + lngOff4)
1910                .croutchk_amount.Left = (.croutchk_amount.Left + lngOff4)
1920                .croutchk_amount_lbl.Left = (.croutchk_amount_lbl.Left + lngOff4)
1930                .croutchk_amount_lbl_line.Left = (.croutchk_amount_lbl_line.Left + lngOff4)
1940                .CheckPaid.Left = (.CheckPaid.Left + lngOff4)
1950                .CheckPaid_lbl.Left = (.CheckPaid_lbl.Left + lngOff4)
1960                .CheckPaid_lbl_line.Left = (.CheckPaid_lbl_line.Left + lngOff4)
1970                .FocusHolder2.Left = (.FocusHolder2.Left + lngOff4)
1980                For lngX = 1& To 10&
1990                  strTmp01 = Right("00" & CStr(lngX), 2)
2000                  .Controls("Detail_vline" & strTmp01).Left = (.Controls("Detail_vline" & strTmp01).Left + lngOff4)
2010                Next
2020                .Detail_vline11.Left = .opgAssign.Left
2030                .Detail_vline12.Left = ((.opgAssign.Left + .opgAssign.Width) - lngTpp)
2040                .Detail_vline11.Visible = True
2050                .Detail_vline12.Visible = True
2060              End If
2070              .opgAssign.Visible = True
2080              .croutchk_asset1_Off.Visible = True
2090              .croutchk_asset1_On.Visible = True
2100              .croutchk_asset2_Off.Visible = True
2110              .croutchk_asset2_On.Visible = True
2120              .opgAssign_lbl.Visible = True
2130              .opgAssign_lbl_line.Visible = True
2140            End With
2150          Else
2160            lngOutChksDiff_Width = 0&
2170          End If
2180        Else
2190          If .OutChecks_box.Left <> lngOutChecksBox_Left Then
2200            lngOutChksDiff_Width = (lngOutChecksBox_Left - .OutChecks_box.Left)
2210            .OutChecks_box.Width = (.OutChecks_box.Width - lngOutChksDiff_Width)
2220            .OutChecks_box.Left = lngOutChecksBox_Left
2230            .frmCheckReconcile_Sub_OutChecks.Width = (.frmCheckReconcile_Sub_OutChecks.Width - lngOutChksDiff_Width)
2240            .frmCheckReconcile_Sub_OutChecks.Left = (lngOutChecksBox_Left + lngOutChksDiff_Left)
2250            .frmCheckReconcile_Sub_OutChecks_box.Width = (.frmCheckReconcile_Sub_OutChecks_box.Width - lngOutChksDiff_Width)
2260            .frmCheckReconcile_Sub_OutChecks_box.Left = (.frmCheckReconcile_Sub_OutChecks.Left - lngTpp)
2270            .frmCheckReconcile_Sub_OutChecks_lbl.Width = (.frmCheckReconcile_Sub_OutChecks_lbl.Width - lngOutChksDiff_Width)
2280            .frmCheckReconcile_Sub_OutChecks_lbl.Left = (.frmCheckReconcile_Sub_OutChecks.Left - (2& * lngTpp))
2290            .OutChecks_Arrow_lbl.Left = (.OutChecks_Arrow_lbl.Left + lngOutChksDiff_Width)
2300            .OutChecks_Arrow_lbl_line1.Left = (.OutChecks_Arrow_lbl_line1.Left + lngOutChksDiff_Width)
2310            .OutChecks_Arrow_lbl_line2.Left = (.OutChecks_Arrow_lbl_line2.Left + lngOutChksDiff_Width)
2320            .OutChecksCnt_lbl.Left = .frmCheckReconcile_Sub_OutChecks.Left
2330            With .frmCheckReconcile_Sub_OutChecks.Form
2340              lngTransdate_Left = .transdate_left
2350              .opgAssign.Visible = False
2360              .opgAssign_lbl.Visible = False
2370              .opgAssign_lbl_line.Visible = False
2380              .croutchk_asset1_Off.Visible = False
2390              .croutchk_asset1_On.Visible = False
2400              .croutchk_asset2_Off.Visible = False
2410              .croutchk_asset2_On.Visible = False
2420              .Detail_vline11.Visible = False
2430              .Detail_vline12.Visible = False
2440              lngOff4 = (.transdate.Left - lngTransdate_Left)
2450              .transdate.Left = (.transdate.Left - lngOff4)
2460              .transdate_lbl.Left = (.transdate_lbl.Left - lngOff4)
2470              .transdate_lbl_line.Left = (.transdate_lbl_line.Left - lngOff4)
2480              .CheckNum.Left = (.CheckNum.Left - lngOff4)
2490              .CheckNum_lbl.Left = (.CheckNum_lbl.Left - lngOff4)
2500              .CheckNum_lbl_line.Left = (.CheckNum_lbl_line.Left - lngOff4)
2510              .croutchk_payee.Left = (.croutchk_payee.Left - lngOff4)
2520              .croutchk_payee_lbl.Left = (.croutchk_payee_lbl.Left - lngOff4)
2530              .croutchk_payee_lbl_line.Left = (.croutchk_payee_lbl_line.Left - lngOff4)
2540              .croutchk_amount.Left = (.croutchk_amount.Left - lngOff4)
2550              .croutchk_amount_lbl.Left = (.croutchk_amount_lbl.Left - lngOff4)
2560              .croutchk_amount_lbl_line.Left = (.croutchk_amount_lbl_line.Left - lngOff4)
2570              .CheckPaid.Left = (.CheckPaid.Left - lngOff4)
2580              .CheckPaid_lbl.Left = (.CheckPaid_lbl.Left - lngOff4)
2590              .CheckPaid_lbl_line.Left = (.CheckPaid_lbl_line.Left - lngOff4)
2600              .FocusHolder2.Left = (.FocusHolder2.Left - lngOff4)
2610              For lngX = 1& To 10&
2620                strTmp01 = Right("00" & CStr(lngX), 2)
2630                .Controls("Detail_vline" & strTmp01).Left = (.Controls("Detail_vline" & strTmp01).Left - lngOff4)
2640              Next
2650            End With  ' ** Subform.
2660          Else
2670            lngOutChksDiff_Width = 0&
2680          End If
2690        End If
2700        .frmCheckReconcile_Sub_OutChecks.Form.SortNow "Form_Load"
2710      End With  ' ** frm.
2720    End If  ' ** IsLoaded().

EXITP:
2730    Set frm = Nothing
2740    Exit Sub

ERRH:
2750    Select Case ERR.Number
        Case Else
2760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2770    End Select
2780    Resume EXITP

End Sub

Public Sub SetSubformTots()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "SetSubformTots"

        Dim frm As Access.Form
        Dim lngRecsCur As Long

2810    If IsLoaded("frmCheckReconcile", acForm) = True Then  ' ** Module Function: modFileUtilities.
2820      Set frm = Forms("frmCheckReconcile")
2830      With frm
2840        lngRecsCur = .frmCheckReconcile_Sub_BSCredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSCredits.
2850        .BSCreditsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
2860        lngRecsCur = .frmCheckReconcile_Sub_BSDebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_BSDebits.
2870        .BSDebitsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
2880        lngRecsCur = .frmCheckReconcile_Sub_TACredits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TACredits.
2890        .TACreditsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
2900        lngRecsCur = .frmCheckReconcile_Sub_TADebits.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_TADebits.
2910        .TADebitsCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")
2920      End With
2930    End If

EXITP:
2940    Set frm = Nothing
2950    Exit Sub

ERRH:
2960    Select Case ERR.Number
        Case Else
2970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2980    End Select
2990    Resume EXITP

End Sub

Public Sub SetSubForms(blnOutChk As Boolean, blnCrDb As Boolean)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "SetSubForms"

        Dim frm As Access.Form

3010    If IsLoaded("frmCheckReconcile", acForm) = True Then  ' ** Module Functions: modFileUtilities.
3020      Set frm = Forms("frmCheckReconcile")
3030      With frm
3040        .TotalOutstandingChecks.Locked = blnOutChk
3050        .TotalOutstandingChecks2.Locked = blnOutChk
3060        .BSTotal.Locked = blnOutChk
3070        .DifferenceBSTA.Locked = blnOutChk
3080        If .frmCheckReconcile_Sub_OutChecks.Enabled <> blnOutChk Then
3090          .cmdUpdate.Enabled = False  ' ** Only if changes have been made.
3100          .frmCheckReconcile_Sub_OutChecks.Enabled = blnOutChk
3110          .chkShowPaid.Enabled = blnOutChk
3120        End If
3130        Select Case blnOutChk
            Case True
3140          .frmCheckReconcile_Sub_OutChecks_lbl.BackStyle = acBackStyleNormal
3150          .TotalOutstandingChecks.BorderColor = CLR_LTBLU2
3160          .TotalOutstandingChecks.BackStyle = acBackStyleNormal
3170          .TotalOutstandingChecks2.BorderColor = CLR_LTBLU2
3180          .TotalOutstandingChecks2.BackStyle = acBackStyleNormal
              '.BSTotal.ForeColor = CLR_DKGRY
3190          .BSTotal.BorderColor = CLR_LTBLU2
3200          .BSTotal.BackStyle = acBackStyleNormal
              '.DifferenceBSTA.ForeColor = CLR_DKGRY
3210          .DifferenceBSTA.BorderColor = CLR_LTBLU2
3220          .DifferenceBSTA.BackStyle = acBackStyleNormal
3230          With .frmCheckReconcile_Sub_OutChecks.Form
3240            .transdate_lbl.ForeColor = CLR_DKGRY2
3250            .CheckNum_lbl.ForeColor = CLR_DKGRY2
3260            .croutchk_payee_lbl.ForeColor = CLR_DKGRY2
3270            .croutchk_amount_lbl.ForeColor = CLR_DKGRY2
3280            .CheckPaid_lbl.ForeColor = CLR_DKGRY2
3290            .opgAssign_lbl.ForeColor = CLR_DKGRY2
3300            .transdate_lbl_line.BorderColor = 13158600
3310            .CheckNum_lbl_line.BorderColor = 13158600
3320            .croutchk_payee_lbl_line.BorderColor = 13158600
3330            .croutchk_amount_lbl_line.BorderColor = 13158600
3340            .CheckPaid_lbl_line.BorderColor = 13158600
3350            .opgAssign_lbl_line.BorderColor = 13158600
3360            .transdate_lbl_dim_hi.Visible = False
3370            .CheckNum_lbl_dim_hi.Visible = False
3380            .croutchk_payee_lbl_dim_hi.Visible = False
3390            .croutchk_amount_lbl_dim_hi.Visible = False
3400            .CheckPaid_lbl_dim_hi.Visible = False
3410            .opgAssign_lbl_dim_hi.Visible = False
3420            .transdate_lbl_line_dim_hi.Visible = False
3430            .CheckNum_lbl_line_dim_hi.Visible = False
3440            .croutchk_payee_lbl_line_dim_hi.Visible = False
3450            .croutchk_amount_lbl_line_dim_hi.Visible = False
3460            .CheckPaid_lbl_line_dim_hi.Visible = False
3470            .opgAssign_lbl_line_dim_hi.Visible = False
3480          End With
3490        Case False
3500          .frmCheckReconcile_Sub_OutChecks_lbl.BackStyle = acBackStyleTransparent
3510          .TotalOutstandingChecks.BorderColor = WIN_CLR_DISR
3520          .TotalOutstandingChecks.BackStyle = acBackStyleTransparent
3530          .TotalOutstandingChecks2.BorderColor = WIN_CLR_DISR
3540          .TotalOutstandingChecks2.BackStyle = acBackStyleTransparent
              '.BSTotal.ForeColor = WIN_CLR_DISF
3550          .BSTotal.BorderColor = WIN_CLR_DISR
3560          .BSTotal.BackStyle = acBackStyleTransparent
3570          .BSTotal = 0
              '.DifferenceBSTA.ForeColor = WIN_CLR_DISF
3580          .DifferenceBSTA.BorderColor = WIN_CLR_DISR
3590          .DifferenceBSTA.BackStyle = acBackStyleTransparent
3600          .DifferenceBSTA = 0
3610          With .frmCheckReconcile_Sub_OutChecks.Form
3620            .transdate_lbl.ForeColor = WIN_CLR_DISF
3630            .CheckNum_lbl.ForeColor = WIN_CLR_DISF
3640            .croutchk_payee_lbl.ForeColor = WIN_CLR_DISF
3650            .croutchk_amount_lbl.ForeColor = WIN_CLR_DISF
3660            .CheckPaid_lbl.ForeColor = WIN_CLR_DISF
3670            .opgAssign_lbl.ForeColor = WIN_CLR_DISF
3680            .transdate_lbl_line.BorderColor = WIN_CLR_DISR
3690            .CheckNum_lbl_line.BorderColor = WIN_CLR_DISR
3700            .croutchk_payee_lbl_line.BorderColor = WIN_CLR_DISR
3710            .croutchk_amount_lbl_line.BorderColor = WIN_CLR_DISR
3720            .CheckPaid_lbl_line.BorderColor = WIN_CLR_DISR
3730            .opgAssign_lbl_line.BorderColor = WIN_CLR_DISR
3740            .transdate_lbl_dim_hi.Visible = True
3750            .CheckNum_lbl_dim_hi.Visible = True
3760            .croutchk_payee_lbl_dim_hi.Visible = True
3770            .croutchk_amount_lbl_dim_hi.Visible = True
3780            .CheckPaid_lbl_dim_hi.Visible = True
3790            .opgAssign_lbl_dim_hi.Visible = True
3800            .transdate_lbl_line_dim_hi.Visible = True
3810            .CheckNum_lbl_line_dim_hi.Visible = True
3820            .croutchk_payee_lbl_line_dim_hi.Visible = True
3830            .croutchk_amount_lbl_line_dim_hi.Visible = True
3840            .CheckPaid_lbl_line_dim_hi.Visible = True
3850            .opgAssign_lbl_line_dim_hi.Visible = True
3860          End With
3870        End Select
3880        If .frmCheckReconcile_Sub_BSCredits.Enabled <> blnCrDb Then
3890          .frmCheckReconcile_Sub_BSCredits.Enabled = blnCrDb
3900          .frmCheckReconcile_Sub_BSCredits.Form.Header_lbl_dim_hi.Visible = (Not blnCrDb)
3910          .frmCheckReconcile_Sub_BSDebits.Enabled = blnCrDb
3920          .frmCheckReconcile_Sub_BSDebits.Form.Header_lbl_dim_hi.Visible = (Not blnCrDb)
3930          .frmCheckReconcile_Sub_TACredits.Enabled = blnCrDb
3940          .frmCheckReconcile_Sub_TACredits.Form.Header_lbl_dim_hi.Visible = (Not blnCrDb)
3950          .frmCheckReconcile_Sub_TADebits.Enabled = blnCrDb
3960          .frmCheckReconcile_Sub_TADebits.Form.Header_lbl_dim_hi.Visible = (Not blnCrDb)
3970          .BSTotalCredits.Locked = blnCrDb
3980          .BSTotalDebits.Locked = blnCrDb
3990          .TATotalCredits.Locked = blnCrDb
4000          .TATotalDebits.Locked = blnCrDb
4010          .TABalance.Locked = blnCrDb
4020          .TATotal.Locked = blnCrDb
4030          Select Case blnCrDb
              Case True
4040            .frmCheckReconcile_Sub_BSCredits.Form.FormHeader.BackColor = CLR_DKBLU
4050            .frmCheckReconcile_Sub_BSCredits.Form.Header_lbl.ForeColor = CLR_WHT
4060            .frmCheckReconcile_Sub_BSDebits.Form.FormHeader.BackColor = CLR_DKBLU
4070            .frmCheckReconcile_Sub_BSDebits.Form.Header_lbl.ForeColor = CLR_WHT
4080            .frmCheckReconcile_Sub_TACredits.Form.FormHeader.BackColor = CLR_DKBLU
4090            .frmCheckReconcile_Sub_TACredits.Form.Header_lbl.ForeColor = CLR_WHT
4100            .frmCheckReconcile_Sub_TADebits.Form.FormHeader.BackColor = CLR_DKBLU
4110            .frmCheckReconcile_Sub_TADebits.Form.Header_lbl.ForeColor = CLR_WHT
4120            .BSTotalCredits.BorderColor = CLR_LTBLU2
4130            .BSTotalCredits.BackStyle = acBackStyleNormal
4140            .BSTotalDebits.BorderColor = CLR_LTBLU2
4150            .BSTotalDebits.BackStyle = acBackStyleNormal
4160            .TATotalCredits.BorderColor = CLR_LTBLU2
4170            .TATotalCredits.BackStyle = acBackStyleNormal
4180            .TATotalDebits.BorderColor = CLR_LTBLU2
4190            .TATotalDebits.BackStyle = acBackStyleNormal
                '.TABalance.ForeColor = CLR_DKGRY
4200            .TABalance.BorderColor = CLR_LTBLU2
4210            .TABalance.BackStyle = acBackStyleNormal
                '.TATotal.ForeColor = CLR_DKGRY
4220            .TATotal.BorderColor = CLR_LTBLU2
4230            .TATotal.BackStyle = acBackStyleNormal
4240          Case False
4250            .frmCheckReconcile_Sub_BSCredits.Form.FormHeader.BackColor = MY_CLR_LTBGE
4260            .frmCheckReconcile_Sub_BSCredits.Form.Header_lbl.ForeColor = WIN_CLR_DISF
4270            .frmCheckReconcile_Sub_BSDebits.Form.FormHeader.BackColor = MY_CLR_LTBGE
4280            .frmCheckReconcile_Sub_BSDebits.Form.Header_lbl.ForeColor = WIN_CLR_DISF
4290            .frmCheckReconcile_Sub_TACredits.Form.FormHeader.BackColor = MY_CLR_LTBGE
4300            .frmCheckReconcile_Sub_TACredits.Form.Header_lbl.ForeColor = WIN_CLR_DISF
4310            .frmCheckReconcile_Sub_TADebits.Form.FormHeader.BackColor = MY_CLR_LTBGE
4320            .frmCheckReconcile_Sub_TADebits.Form.Header_lbl.ForeColor = WIN_CLR_DISF
4330            .BSTotalCredits.BorderColor = WIN_CLR_DISR
4340            .BSTotalCredits.BackStyle = acBackStyleTransparent
4350            .BSTotalDebits.BorderColor = WIN_CLR_DISR
4360            .BSTotalDebits.BackStyle = acBackStyleTransparent
4370            .TATotalCredits.BorderColor = WIN_CLR_DISR
4380            .TATotalCredits.BackStyle = acBackStyleTransparent
4390            .TATotalDebits.BorderColor = WIN_CLR_DISR
4400            .TATotalDebits.BackStyle = acBackStyleTransparent
                '.TABalance.ForeColor = WIN_CLR_DISF
4410            .TABalance.BorderColor = WIN_CLR_DISR
4420            .TABalance.BackStyle = acBackStyleTransparent
4430            .TABalance = 0
                '.TATotal.ForeColor = WIN_CLR_DISF
4440            .TATotal.BorderColor = WIN_CLR_DISR
4450            .TATotal.BackStyle = acBackStyleTransparent
4460            .TATotal = 0
4470          End Select
4480        End If
4490      End With
4500    End If

EXITP:
4510    Set frm = Nothing
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Public Function TotOutChks(frm As Access.Form, strAccountNo As String) As Boolean

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "TotOutChks"

        Dim strQryName As String
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long
        Dim blnRetVal As Boolean

4610    With frm

4620      blnRetVal = True

4630      varTmp00 = Null
4640      lngTmp01 = .AssetNumGet(1, 1)  ' ** Form Procedure: frmCheckReconcile.
4650      lngTmp02 = .AssetNumGet(1, 2)  ' ** Form Procedure: frmCheckReconcile.
4660      lngTmp03 = .AssetNumGet(3, 0)  ' ** Form Procedure: frmCheckReconcile.

4670      Select Case lngTmp03
          Case 1&
            ' ** For accounts with only 1 asset.
4680        Select Case strAccountNo
            Case "CRTC01"
              ' ** tblCheckReconcile_Check, grouped and summed, forCheckPaid = False,
              ' ** croutchk_assign = 0, accountno = 'CRTC01'.
4690          strQryName = "qryCheckReconcile_OutChecks_09_04"
4700        Case Else
              ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
              ' ** croutchk_assign = 0, by specified GlobalVarGet("gstrAccountNo").
4710          strQryName = "qryCheckReconcile_OutChecks_09_01"
4720        End Select
4730      Case Else
4740        Select Case .cmbAssets
            Case lngTmp01
              ' ** For accounts with 2 assets, and this is the first.
4750          Select Case strAccountNo
              Case "CRTC01"
                ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                ' ** croutchk_assign = 0,1, accountno = 'CRTC01'.
4760            strQryName = "qryCheckReconcile_OutChecks_09_05"
4770          Case Else
                ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                ' ** croutchk_assign = 0,1, by specified GlobalVarGet("gstrAccountNo").
4780            strQryName = "qryCheckReconcile_OutChecks_09_02"
4790          End Select
4800        Case lngTmp02
              ' ** For accounts with 2 assets, and this is the second.
4810          Select Case strAccountNo
              Case "CRTC01"
                ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                ' ** croutchk_assign = 0,2, accountno = 'CRTC01'.
4820            strQryName = "qryCheckReconcile_OutChecks_09_06"
4830          Case Else
                ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                ' ** croutchk_assign = 0,2, by specified GlobalVarGet("gstrAccountNo").
4840            strQryName = "qryCheckReconcile_OutChecks_09_03"
4850          End Select
4860        End Select
4870      End Select

4880      If strQryName <> vbNullString Then
4890        varTmp00 = DLookup("CalcTotal", strQryName)
4900      End If

4910      .TotalOutstandingChecks = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
4920      .TotalOutstandingChecks2 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
4930      DoEvents

4940    End With

        'TotalOutstandingChecks =
        'IIf([cmbAccounts]="CRTC01",DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09b"),DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09a"))
        'TotalOutstandingChecks2 =
        'IIf([cmbAccounts]="CRTC01",DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09b"),DLookUp("CalcTotal","qryCheckReconcile_OutChecks_09a"))

EXITP:
4950    Exit Function

ERRH:
4960    Select Case ERR.Number
        Case Else
4970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4980    End Select
4990    Resume EXITP

End Function

Public Sub CheckingType_Set(lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant, frm As Access.Form)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckingType_Set"

        Dim strTmp01 As String

5010    With frm
5020      Select Case .opgCheckingType
          Case .opgCheckingType_optIndividual.OptionValue
5030        AccountsRowSource_Set frm  ' ** Procedure: Below.
5040        .opgCheckingType_optIndividual_lbl.FontBold = True
5050        .opgCheckingType_optTotalCash_lbl.FontBold = False
5060        If IsNull(.cmbAccounts) = False Then
5070          .cmbAccounts_AfterUpdate  ' ** Form Procedure: CheckReconcile.
5080          .cmdPreviewReport.Enabled = True
5090          .cmdPreviewReport_raised_img.Visible = True
5100          .cmdPrintReport.Enabled = True
5110          .cmdPrintReport_raised_img.Visible = True
5120        Else
5130          SetSubForms False, False  ' ** Procedure: Above.
5140          .cracct_id.Visible = False
5150          .cmbAccounts.Enabled = True
5160          .cmbAccounts.BorderColor = CLR_LTBLU2
5170          .cmbAccounts.BackStyle = acBackStyleNormal
5180          .cmbAccounts_lbl.BackStyle = acBackStyleNormal
5190          .cmbAccounts_lbl_box.Visible = False
5200          .chkRememberMe.Enabled = True
5210          .chkRememberMe_lbl.Visible = True
5220          .chkRememberMe_lbl2_dim.Visible = False
5230          .chkRememberMe_lbl2_dim_hi.Visible = False
5240          .chkRememberMe = False
5250          .chkRememberMe_lbl.FontBold = False
5260          .chkRememberMe_lbl2_dim.FontBold = False
5270          .chkRememberMe_lbl2_dim_hi.FontBold = False
5280          .opgAccountSource.Enabled = True
5290          .cmbAssets.Enabled = False
5300          .cmbAssets.BorderColor = WIN_CLR_DISR
5310          .cmbAssets.BackStyle = acBackStyleTransparent
5320          .cmbAssets_lbl.BackStyle = acBackStyleTransparent
5330          .cmbAssets_lbl_box.Visible = True
5340          .opgAssetSource.Enabled = False
5350          .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
5360          .opgAssetSource_optType_lbl2_dim_hi.Visible = True
5370          .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
5380          .opgAssetSource_optName_lbl2_dim_hi.Visible = True
5390          .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
5400          .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
5410          .cmdAddAsset.Enabled = False
5420          .cmdAddAsset_raised_img_dis.Visible = True
5430          .cmdAddAsset_raised_img.Visible = False
5440          .cmdAssetPrevious.Enabled = False
5450          .cmdAssetNext.Enabled = False
5460          .DateEnd.Enabled = False
5470          .DateEnd = Null
5480          .DateEnd.BorderColor = WIN_CLR_DISR
5490          .DateEnd.BackStyle = acBackStyleTransparent
5500          .DateEnd_lbl.BackStyle = acBackStyleTransparent
5510          .DateEnd_lbl_box.Visible = True
5520          .cmdCalendar.Enabled = False
5530          .cmdCalendar_raised_img_dis.Visible = True
5540          .cmdCalendar_raised_img.Visible = False
5550          .Bank_Name.Enabled = False
5560          .Bank_Name.Locked = False
5570          .Bank_Name.BorderColor = WIN_CLR_DISR
5580          .Bank_Name.BackStyle = acBackStyleTransparent
5590          .Bank_AccountNumber.Enabled = False
5600          .Bank_AccountNumber.Locked = False
5610          .Bank_AccountNumber.BorderColor = WIN_CLR_DISR
5620          .Bank_AccountNumber.BackStyle = acBackStyleTransparent
5630          .cracct_bsbalance_display.Enabled = False
5640          .cracct_bsbalance_display.BorderColor = WIN_CLR_DISR
5650          .cracct_bsbalance_display.BackStyle = acBackStyleTransparent
5660  On Error Resume Next
5670          strTmp01 = Screen.ActiveControl.Name  ' ** If a control doesn't have the focus, it'll error.
5680  On Error GoTo ERRH
5690          If strTmp01 = "cmdDelete" Or strTmp01 = "cmdClear" Then
5700            .cmdClose.SetFocus
5710          End If
5720          .cmdDelete.Enabled = False
5730          .cmdDelete_lbl1.ForeColor = WIN_CLR_DISF
5740          .cmdDelete_lbl2.ForeColor = WIN_CLR_DISF
5750          If .UnpostedJournalMsg_lbl.Visible = False Then
5760            .cmdDelete_lbl1_dim_hi.Visible = True
5770            .cmdDelete_lbl2_dim_hi.Visible = True
5780          End If
5790          .cmdClear.Enabled = False
5800          .cmdClear_lbl1.ForeColor = WIN_CLR_DISF
5810          .cmdClear_lbl2.ForeColor = WIN_CLR_DISF
5820          If .UnpostedJournalMsg_lbl.Visible = False Then
5830            .cmdClear_lbl1_dim_hi.Visible = True
5840            .cmdClear_lbl2_dim_hi.Visible = True
5850          End If
5860          .cmdCheckAll.Enabled = False
5870          .cmdCheckAll_raised_img_dis.Visible = True
5880          .cmdCheckAll_raised_img.Visible = False
5890          .cmdCheckNone.Enabled = False
5900          .cmdCheckNone_raised_img_dis.Visible = True
5910          .cmdCheckNone_raised_img.Visible = False
5920          .cmdPreviewReport.Enabled = False
5930          .cmdPreviewReport_raised_img_dis.Visible = True
5940          .cmdPreviewReport_raised_img.Visible = False
5950          .cmdPrintReport.Enabled = False
5960          .cmdPrintReport_raised_img_dis.Visible = True
5970          .cmdPrintReport_raised_img.Visible = False
5980        End If
5990      Case .opgCheckingType_optTotalCash.OptionValue
6000        AccountsRowSource_Set frm  ' ** Procedure: Below.
6010        .opgCheckingType_optTotalCash_lbl.FontBold = True
6020        .opgCheckingType_optIndividual_lbl.FontBold = False
6030        .DateEnd.Enabled = True
6040        .DateEnd = Null
6050        .DateEnd.BorderColor = CLR_LTBLU2
6060        .DateEnd.BackStyle = acBackStyleNormal
6070        .DateEnd_lbl.BackStyle = acBackStyleNormal
6080        .DateEnd_lbl_box.Visible = False
6090        .cmdCalendar.Enabled = True
6100        .cmdCalendar_raised_img.Visible = True
6110        .cmdCalendar_raised_semifocus_dots_img.Visible = False
6120        .cmdCalendar_raised_focus_img.Visible = False
6130        .cmdCalendar_raised_focus_dots_img.Visible = False
6140        .cmdCalendar_sunken_focus_dots_img.Visible = False
6150        .cmdCalendar_raised_img_dis.Visible = False
6160        .DateEnd.SetFocus
6170        .cmbAccounts.Enabled = False
6180        .cmbAccounts.BorderColor = WIN_CLR_DISR
6190        .cmbAccounts.BackStyle = acBackStyleTransparent
6200        .cmbAccounts_lbl.BackStyle = acBackStyleTransparent
6210        .cmbAccounts_lbl_box.Visible = True
6220        .chkRememberMe.Enabled = False
6230        .chkRememberMe_lbl.Visible = False
6240        .chkRememberMe_lbl2_dim.Visible = True
6250        .chkRememberMe_lbl2_dim_hi.Visible = True
6260        .chkRememberMe = True
6270        .chkRememberMe_lbl.FontBold = True
6280        .chkRememberMe_lbl2_dim.FontBold = True
6290        .chkRememberMe_lbl2_dim_hi.FontBold = True
6300        .opgAccountSource.Enabled = False
6310        .cmbAssets = Null
6320        .cmbAssets.Enabled = False
6330        .cmbAssets.BorderColor = WIN_CLR_DISR
6340        .cmbAssets.BackStyle = acBackStyleTransparent
6350        .cmbAssets_lbl.BackStyle = acBackStyleTransparent
6360        .cmbAssets_lbl_box.Visible = True
6370        .opgAssetSource.Enabled = False
6380        .opgAssetSource_optType_lbl2.ForeColor = WIN_CLR_DISF
6390        .opgAssetSource_optType_lbl2_dim_hi.Visible = True
6400        .opgAssetSource_optName_lbl2.ForeColor = WIN_CLR_DISF
6410        .opgAssetSource_optName_lbl2_dim_hi.Visible = True
6420        .opgAssetSource_optCusip_lbl2.ForeColor = WIN_CLR_DISF
6430        .opgAssetSource_optCusip_lbl2_dim_hi.Visible = True
6440        .cmdAddAsset.Enabled = False
6450        .cmdAddAsset_raised_img_dis.Visible = True
6460        .cmdAddAsset_raised_img.Visible = False
6470        .cmdAssetPrevious.Enabled = False
6480        .cmdAssetNext.Enabled = False
6490        .Bank_Name.Enabled = True
6500        .Bank_Name.Locked = False
6510        .Bank_Name.BorderColor = CLR_LTBLU2
6520        .Bank_Name.BackStyle = acBackStyleNormal
6530        .Bank_AccountNumber.Enabled = True
6540        .Bank_AccountNumber.Locked = False
6550        .Bank_AccountNumber.BorderColor = CLR_LTBLU2
6560        .Bank_AccountNumber.BackStyle = acBackStyleNormal
6570        lngAssets_ThisAcct = 0&
6580        arr_varAsset_ThisAcct = Empty
6590      End Select
6600    End With

EXITP:
6610    Exit Sub

ERRH:
6620    DoCmd.Hourglass False
6630    Select Case ERR.Number
        Case Else
6640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6650    End Select
6660    Resume EXITP

End Sub

Public Function ResetCheckSub(lngTmp01 As Long, arr_varTmp02 As Variant, lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant, lngAssetNo_Move As Long, frm As Access.Form) As Boolean

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "ResetCheckSub"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String
        Dim lngCRAcctID As Long, datDateEnd As Date
        Dim blnOutChk As Boolean, blnCrDb As Boolean
        Dim lngRecs As Long, lngRecsCur As Long
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

6710    blnRetVal = True

6720    With frm
6730      If IsNull(.cmbAccounts) = False And IsNull(.cracct_id) = False And IsNull(.DateEnd) = False Then
6740        If Trim(.cmbAccounts) <> vbNullString And IsDate(.DateEnd) = True Then

6750          DoCmd.Hourglass True
6760          DoEvents

6770          lngChks = 0&
6780          ReDim arr_varChk(C_ELEMS, 0)

6790          lngRecs = lngTmp01
6800          For lngX = 0& To (lngRecs - 1&)
6810            lngChks = lngChks + 1&
6820            lngE = lngChks - 1&
6830            ReDim Preserve arr_varChk(C_ELEMS, lngE)
6840            For lngY = 0& To C_ELEMS
6850              arr_varChk(lngY, lngE) = arr_varTmp02(lngY, lngX)
6860            Next
6870          Next

6880          strAccountNo = .cmbAccounts.Column(CBX_ACT_ACTNO)
6890          lngCRAcctID = .cracct_id
6900          datDateEnd = .DateEnd

              ' ** This is used by FormRef() to enable appending the tblCheckReconcile_Check records.
6910          .cracct_id_tmp = lngCRAcctID

6920          Set dbs = CurrentDb

              ' ** Empty tblCheckReconcile_Staging.  '## OK
6930          Set qdf = dbs.QueryDefs("qryCheckReconcile_OutChecks_20")
6940          qdf.Execute

6950          Select Case .opgCheckingType
              Case .opgCheckingType_optIndividual.OptionValue
                ' ** Append qryCheckReconcile_OutChecks_21a (Ledger, just 'Paid', by specified [actno])
                ' ** to tblCheckReconcile_Staging, by specified [actno], [cractid], [datend].
6960            Set qdf = dbs.QueryDefs("qryCheckReconcile_OutChecks_22a")
6970            With qdf.Parameters
6980              ![actno] = strAccountNo
6990              ![cractid] = lngCRAcctID
7000              ![datEnd] = datDateEnd
7010            End With
7020          Case .opgCheckingType_optTotalCash.OptionValue
                ' ** Append qryCheckReconcile_OutChecks_21b (Ledger, just 'Paid', for accountno = 'CRTC01')
                ' ** to tblCheckReconcile_Staging, for accountno = 'CRTC01', by specified [cractid], [datend].
7030            Set qdf = dbs.QueryDefs("qryCheckReconcile_OutChecks_22b")
7040            With qdf.Parameters
7050              ![cractid] = lngCRAcctID
7060              ![datEnd] = Nz(frm.DateEnd, Date)
7070            End With
7080          End Select
7090          qdf.Execute dbFailOnError

7100          UpdateStageTbl lngAssets_ThisAcct, arr_varAsset_ThisAcct, dbs  ' ** Procedure: Above.

7110          lngChks = 0&
7120          ReDim arr_varChk(C_ELEMS, 0)

7130          With dbs

                ' ** tblCheckReconcile_Staging, just needed fields.  '## OK
7140            Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_23")
7150            Set rst = qdf.OpenRecordset
7160            With rst
7170              If .BOF = True And .EOF = True Then
                    ' ** No unpaid checks.
7180              Else
7190                .MoveLast
7200                lngRecs = .RecordCount
7210                .MoveFirst
7220                For lngX = 1& To lngRecs
7230                  lngChks = lngChks + 1&
7240                  lngE = lngChks - 1&
7250                  ReDim Preserve arr_varChk(C_ELEMS, lngE)
7260                  arr_varChk(C_ACTID, lngE) = ![cracct_id]
7270                  arr_varChk(C_STGID, lngE) = ![crstage_id]
7280                  arr_varChk(C_JNO, lngE) = ![journalno]
7290                  arr_varChk(C_ACTNO, lngE) = ![accountno]
7300                  arr_varChk(C_CHKID, lngE) = CLng(0)  ' ** croutchk_id.
7310                  arr_varChk(C_CNUM, lngE) = ![CheckNumx]
7320                  arr_varChk(C_PAID, lngE) = ![CheckPaidx]
7330                  arr_varChk(C_DESC, lngE) = ![Descriptionx]
7340                  arr_varChk(C_HDESC, lngE) = ![crstage_hasdesc]
7350                  arr_varChk(C_ASGN, lngE) = ![crstage_assign]
7360                  arr_varChk(C_ASTNO1, lngE) = ![crstage_asset1]
7370                  arr_varChk(C_ASTNO2, lngE) = ![crstage_asset2]
7380                  arr_varChk(C_FND, lngE) = CBool(True)
7390                  If lngX < lngRecs Then .MoveNext
7400                Next
7410              End If
7420              .Close
7430            End With  ' ** rst.

                ' ** Empty tblCheckReconcile_Check.  '## OK
7440            Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_02")
7450            qdf.Execute

7460            Select Case frm.chkShowPaid
                Case True
                  ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check, by specified [datend]; All.  '## OK
7470              Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_25")
7480              With qdf.Parameters
7490                ![datEnd] = datDateEnd
7500              End With
7510              qdf.Execute
7520            Case False
                  ' ** For multi-asset, we need just Zeroes, or their asset!
7530              If lngAssets_ThisAcct > 1& Then
7540                If lngAssetNo_Move >= 0& Then
7550                  If lngAssetNo_Move > 0& Then
                        ' ** Zeroes and this association.
7560                    If lngAssets_ThisAcct > 1& Then
                          ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check, for
                          ' ** CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/ChangedMatch.
7570                      Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_31")  '####  UNKNOWN  ####
7580                    Else
                          ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                          ' ** for CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/Changed.
7590                      Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_29")  '####  UNKNOWN  ####
7600                    End If
7610                    With qdf.Parameters
7620                      ![datEnd] = datDateEnd
7630                      ![astno] = lngAssetNo_Move
7640                    End With
7650                  Else
                        ' ** Zeroes only.
                        ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check, for
                        ' ** CheckPaid = False, by specified [datend]; Zeroes/Changed.
7660                    Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_30")  '####  UNKNOWN  ####
7670                    With qdf.Parameters
7680                      ![datEnd] = datDateEnd
7690                    End With
7700                  End If
7710                Else
                      ' ** Zeroes, and the first association.
7720                  If lngAssets_ThisAcct > 1& Then
                        ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check, for
                        ' ** CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/ChangedMatch.
7730                    Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_31")  '####  UNKNOWN  ####
7740                  Else
                        ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check, for
                        ' ** CheckPaid = False, by specified [datend], [astno]; Zeroes/Match/Changed.
7750                    Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_29")  '####  UNKNOWN  ####
7760                  End If
7770                  With qdf.Parameters
7780                    ![datEnd] = datDateEnd
7790                    ![astno] = arr_varAsset_ThisAcct(A_ASTNO, 0)
7800                  End With
7810                End If
7820              Else
                    ' ** Append tblCheckReconcile_Staging to tblCheckReconcile_Check,
                    ' ** for CheckPaid = False, by specified [datend]; No Criteria.  '## OK
7830                Set qdf = .QueryDefs("qryCheckReconcile_OutChecks_28")
7840                With qdf.Parameters
7850                  ![datEnd] = datDateEnd
7860                End With
7870              End If
7880              qdf.Execute dbFailOnError
7890            End Select

7900            .Close
7910          End With  ' ** dbs.

7920          DoEvents
7930          .Requery

7940          .frmCheckReconcile_Sub_OutChecks.Form.Requery
7950          DoEvents
7960          lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
7970          .OutChecksCnt_lbl.Caption = IIf(lngRecsCur = 1, "1 Item", CStr(lngRecsCur) & " Items")

7980          blnOutChk = True
7990          If IsNull(.cmbAssets) = False Then
8000            If .cmbAssets > 0 Or .opgCheckingType = .opgCheckingType_optTotalCash.OptionValue Then
8010              blnCrDb = True
8020            End If
8030          End If

8040          SetSubForms blnOutChk, blnCrDb  ' ** Module Procedure: modCheckReconcile.

8050          .FormValidate  ' ** Form Procedure: frmCheckReconcile.

8060          If .Bank_Name.Enabled = True And .Bank_Name.Locked = False Then
8070            .Bank_Name.SetFocus
8080          End If

8090          DoCmd.Hourglass False

8100          lngTmp01 = lngChks
8110          arr_varTmp02 = arr_varChk

8120        End If  ' ** vbNullstring.
8130      End If  ' ** IsNull().
8140    End With

EXITP:
8150    Set rst = Nothing
8160    Set qdf = Nothing
8170    Set dbs = Nothing
8180    Exit Function

ERRH:
8190    DoCmd.Hourglass False
8200    Select Case ERR.Number
        Case Else
8210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8220    End Select
8230    Resume EXITP

End Function

Public Sub AccountsRowSource_Set(frm As Access.Form)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "AccountsRowSource_Set"

8310    With frm
8320      Select Case .opgCheckingType
          Case .opgCheckingType_optIndividual.OptionValue
8330        Select Case .opgAccountSource
            Case .opgAccountSource_optNumber.OptionValue
8340          If .cmbAccounts.RowSource <> "qryCheckReconcile_03a" Then
                ' ** Account, linked to qryCheckReconcile_02 (Balance, grouped by accountno, with
                ' ** Min(Balance_Date)), accountno, AcctName (accountno,shortname), with cash criteria.
8350            .cmbAccounts.RowSource = "qryCheckReconcile_03a"
8360          End If
8370        Case .opgAccountSource_optName.OptionValue
8380          If .cmbAccounts.RowSource <> "qryCheckReconcile_03b" Then
                ' ** Account, linked to qryCheckReconcile_02 (Balance, grouped by accountno, with
                ' ** Min(Balance_Date)), accountno, AcctName (shortname), with cash criteria.
8390            .cmbAccounts.RowSource = "qryCheckReconcile_03b"
8400          End If
8410        End Select
8420      Case .opgCheckingType_optTotalCash.OptionValue
8430        Select Case .opgAccountSource
            Case .opgAccountSource_optNumber.OptionValue
8440          If .cmbAccounts.RowSource <> "qryCheckReconcile_03c" Then
                ' ** Account, linked to qryCheckReconcile_02 (Balance, grouped by accountno, with
                ' ** Min(Balance_Date)), accountno, AcctName (accountno,shortname), with cash criteria, with 'CRTC01'.
8450            .cmbAccounts.RowSource = "qryCheckReconcile_03c"
8460          End If
8470        Case .opgAccountSource_optName.OptionValue
8480          If .cmbAccounts.RowSource <> "qryCheckReconcile_03d" Then
                ' ** Account, linked to qryCheckReconcile_02 (Balance, grouped by accountno, with
                ' ** Min(Balance_Date)), accountno, AcctName (shortname), with cash criteria, with 'CRTC01'.
8490            .cmbAccounts.RowSource = "qryCheckReconcile_03d"
8500          End If
8510        End Select
8520      End Select
8530    End With

EXITP:
8540    Exit Sub

ERRH:
8550    DoCmd.Hourglass False
8560    Select Case ERR.Number
        Case Else
8570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8580    End Select
8590    Resume EXITP

End Sub

Public Sub GetBal_End_CR(strAccountNo As String, frm As Access.Form)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "GetBal_End_CR"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef

8610    gstrFormQuerySpec = frm.Name
8620    With frm
8630      gstrAccountNo = strAccountNo
8640      Set dbs = CurrentDb
          ' ** Get an ending balance.
8650      Select Case .opgCheckingType
          Case .opgCheckingType_optIndividual.OptionValue
            ' ** qryCheckReconcile_Balance_03 (qryCheckReconcile_Balance_02 (qryCheckReconcile_Balance_01
            ' ** (Ledger, with add'l fields, by specified FormRef('AccountNo', 'AssetNo', 'PStartDate',
            ' ** 'PEndDate')), grouped and summed), with .._dp, .._ws totals broken out, and
            ' ** Sumshareface_txt_Liability), grouped and summed by accountno.
8660        Set qdf = dbs.QueryDefs("qryCheckReconcile_Balance_04")
8670        Set rst = qdf.OpenRecordset
8680        If rst.BOF = True And rst.EOF = True Then
              ' ** No transactions listed in Ledger.
8690          .TABalance = 0@
8700        Else
8710          rst.MoveFirst
8720          If IsNull(rst![Sumshareface_dp]) = False Then
8730            .TABalance = rst![SumCost]
8740          Else
8750            .TABalance = 0@
8760          End If
8770        End If
8780      Case .opgCheckingType_optTotalCash.OptionValue
            ' ** qryCheckReconcile_Balance_13 (qryCheckReconcile_Balance_12 (qryCheckReconcile_Balance_11
            ' ** (qryCheckReconcile_Balance_10 (Ledger, with add'l fields, by specified FormRef('PStartDate',
            ' ** 'PEndDate')), grouped and summed), with .._dp, .._ws totals broken out,
            ' ** and Sumshareface_txt_Liability), grouped and summed by accountno),
            ' ** grouped and summed, for accountno = 'CRTC01'.
8790        Set qdf = dbs.QueryDefs("qryCheckReconcile_Balance_14")
8800        Set rst = qdf.OpenRecordset
8810        If rst.BOF = True And rst.EOF = True Then
              ' ** No transactions listed in Ledger.
8820          .TABalance = 0@
8830        Else
8840          rst.MoveFirst
8850          .TABalance = ((rst![SumICash_dp] + rst![SumICash_ws]) + (rst![SumPCash_dp] + rst![SumPCash_ws]))
8860        End If
8870      End Select
8880      rst.Close
8890      Select Case .opgCheckingType
          Case .opgCheckingType_optIndividual.OptionValue
            ' ** qryCheckReconcile_Balance_07 (qryCheckReconcile_Balance_06 (qryCheckReconcile_Balance_05
            ' ** (LedgerArchive, with add'l fields, by specified FormRef('AccountNo', 'AssetNo', 'PStartDate',
            ' ** 'PEndDate')), grouped and summed), with .._dp, .._ws totals broken out, and
            ' ** Sumshareface_txt_Liability), grouped and summed by accountno.
8900        Set qdf = dbs.QueryDefs("qryCheckReconcile_Balance_08")
8910  On Error Resume Next
8920        Set rst = qdf.OpenRecordset
8930        If ERR.Number = 0 Then
8940  On Error GoTo ERRH
8950          If rst.BOF = True And rst.EOF = True Then
                ' ** No transactions listed in LedgerArchive.
8960          Else
8970            rst.MoveFirst
8980            If IsNull(rst![Sumshareface_dp]) = False Then
8990              .TABalance = (.TABalance + rst![SumCost])
9000            End If
9010          End If
9020        Else
9030  On Error GoTo ERRH
9040        End If
9050      Case .opgCheckingType_optTotalCash.OptionValue
            ' ** qryCheckReconcile_Balance_18 (qryCheckReconcile_Balance_17 (qryCheckReconcile_Balance_16
            ' ** (qryCheckReconcile_Balance_15 (LedgerArchive, with add'l fields, by specified FormRef('PStartDate',
            ' ** 'PEndDate')), grouped and summed), with .._dp, .._ws totals broken out,
            ' ** and Sumshareface_txt_Liability), grouped and summed by accountno), grouped and summed,
            ' ** for accountno = 'CRTC01'.
9060        Set qdf = dbs.QueryDefs("qryCheckReconcile_Balance_19")
9070  On Error Resume Next
9080        Set rst = qdf.OpenRecordset
9090        If ERR.Number = 0 Then
9100  On Error GoTo ERRH
9110          If rst.BOF = True And rst.EOF = True Then
                ' ** No transactions listed in LedgerArchive.
9120          Else
9130            rst.MoveFirst
9140            .TABalance = (.TABalance + ((rst![SumICash_dp] + rst![SumICash_ws]) + (rst![SumPCash_dp] + rst![SumPCash_ws])))
9150          End If
9160        Else
9170  On Error GoTo ERRH
9180        End If
9190      End Select
9200      rst.Close
9210      dbs.Close
9220    End With

EXITP:
9230    Set rst = Nothing
9240    Set qdf = Nothing
9250    Set dbs = Nothing
9260    Exit Sub

ERRH:
9270    DoCmd.Hourglass False
9280    Select Case ERR.Number
        Case Else
9290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9300    End Select
9310    Resume EXITP

End Sub

Public Sub UpdateBSTotal_CR(frm As Access.Form, lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant, Optional varProc As Variant)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateBSTotal_CR"

        Dim varTmp00 As Variant, varTmp01 As Variant, varTmp02 As Variant, varTmp03 As Variant, varTmp04 As Variant
        Dim strAccountNo As String, strQryName As String

9410    gstrFormQuerySpec = frm.Name
9420    With frm
9430      If IsNull(.cmbAccounts) = False Then
9440        If Trim(.cmbAccounts) <> vbNullString Then
9450          strAccountNo = .cmbAccounts
9460          If IsNull(.cmbAssets) = False Then
9470            If .cmbAssets > 0 Then
9480              varTmp00 = .cracct_bsbalance_display
                  ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'BS',
                  ' ** EntryType = 'Credit', by specified FormRef('AccountNo', 'AssetNo').
9490              varTmp01 = DLookup("CalcAmount", "qryCheckReconcile_BSCreditTotal_01a")
                  ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'BS',
                  ' ** EntryType = 'Debit', by specified FormRef('AccountNo', 'AssetNo').
9500              varTmp02 = DLookup("CalcAmount", "qryCheckReconcile_BSDebitTotal_01a")
                  ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                  ' ** by specified FormRef('AccountNo').
9510              Select Case lngAssets_ThisAcct
                  Case 1&
                    ' ** For accounts with only 1 asset.
                    ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                    ' ** croutchk_assign = 0, by specified GlobalVarGet("gstrAccountNo").
9520                strQryName = "qryCheckReconcile_OutChecks_09_01"
9530              Case Else
9540                If IsNull(.cmbAssets) = True Then
                      ' ** No asset, or not finished populating fields.
9550                  strQryName = vbNullString
9560                ElseIf lngAssets_ThisAcct = 0& Or IsEmpty(arr_varAsset_ThisAcct) = True Then
                      ' ** Form not finished loading yet.
9570                Else
9580                  Select Case .cmbAssets
                      Case arr_varAsset_ThisAcct(A_ASTNO, 0)
                        ' ** For accounts with 2 assets, and this is the first.
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,1, by specified GlobalVarGet("gstrAccountNo").
9590                    strQryName = "qryCheckReconcile_OutChecks_09_02"
9600                  Case arr_varAsset_ThisAcct(A_ASTNO, 1)
                        ' ** For accounts with 2 assets, and this is the second.
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,2, by specified GlobalVarGet("gstrAccountNo").
9610                    strQryName = "qryCheckReconcile_OutChecks_09_03"
9620                  End Select
9630                End If
9640              End Select
9650              If strQryName = vbNullString Then
9660                varTmp03 = 0
9670              Else
9680                varTmp03 = DLookup("CalcTotal", strQryName)
9690              End If
9700              varTmp00 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
9710              varTmp01 = ZeroIfNull(varTmp01)  ' ** Module Function: modStringFuncs.
9720              varTmp02 = ZeroIfNull(varTmp02)  ' ** Module Function: modStringFuncs.
9730              varTmp03 = ZeroIfNull(varTmp03)  ' ** Module Function: modStringFuncs.
9740              .BSTotal = (((varTmp00 + varTmp01) - varTmp02) - Abs(varTmp03))
9750              varTmp04 = .TATotal
9760              varTmp04 = ZeroIfNull(varTmp04)  ' ** Module Function: modStringFuncs.
9770              .DifferenceBSTA = ((((varTmp00 + varTmp01) - varTmp02) - Abs(varTmp03)) - varTmp04)
9780            Else
9790              If strAccountNo = "CRTC01" Then
9800                varTmp00 = .cracct_bsbalance_display
                    ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'BS',
                    ' ** EntryType = 'Credit', for accountno = 'CRTC01'.
9810                varTmp01 = DLookup("CalcAmount", "qryCheckReconcile_BSCreditTotal_01b")
                    ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'BS',
                    ' ** EntryType = 'Debit', for accountno = 'CRTC01'.
9820                varTmp02 = DLookup("CalcAmount", "qryCheckReconcile_BSDebitTotal_01b")
                    ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                    ' ** accountno = 'CRTC01'.
9830                Select Case lngAssets_ThisAcct
                    Case 1&
                      ' ** For accounts with only 1 asset.
                      ' ** tblCheckReconcile_Check, grouped and summed, forCheckPaid = False,
                      ' ** croutchk_assign = 0, accountno = 'CRTC01'.
9840                  strQryName = "qryCheckReconcile_OutChecks_09_04"
9850                Case Else
9860                  Select Case .cmbAssets
                      Case arr_varAsset_ThisAcct(A_ASTNO, 0)
                        ' ** For accounts with 2 assets, and this is the first.
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,1, accountno = 'CRTC01'.
9870                    strQryName = "qryCheckReconcile_OutChecks_09_05"
9880                  Case arr_varAsset_ThisAcct(A_ASTNO, 1)
                        ' ** For accounts with 2 assets, and this is the second.
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,2, accountno = 'CRTC01'.
9890                    strQryName = "qryCheckReconcile_OutChecks_09_06"
9900                  End Select
9910                End Select
9920                varTmp03 = DLookup("CalcTotal", strQryName)
9930                varTmp00 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
9940                varTmp01 = ZeroIfNull(varTmp01)  ' ** Module Function: modStringFuncs.
9950                varTmp02 = ZeroIfNull(varTmp02)  ' ** Module Function: modStringFuncs.
9960                varTmp03 = ZeroIfNull(varTmp03)  ' ** Module Function: modStringFuncs.
9970                .BSTotal = (((varTmp00 + varTmp01) - varTmp02) - Abs(varTmp03))
9980                varTmp04 = .TATotal
9990                varTmp04 = ZeroIfNull(varTmp04)  ' ** Module Function: modStringFuncs.
10000               .DifferenceBSTA = ((((varTmp00 + varTmp01) - varTmp02) - Abs(varTmp03)) - varTmp04)
10010             Else
10020               .BSTotal = 0
10030               .DifferenceBSTA = 0
10040             End If
10050           End If
10060         Else
10070           .BSTotal = 0
10080           .DifferenceBSTA = 0
10090         End If
10100       Else
10110         .BSTotal = 0
10120         .DifferenceBSTA = 0
10130       End If
10140     Else
10150       .BSTotal = 0
10160       .DifferenceBSTA = 0
10170     End If
10180   End With
10190   gstrFormQuerySpec = frm.Name

EXITP:
10200   gstrFormQuerySpec = frm.Name
10210   Exit Sub

ERRH:
10220   DoCmd.Hourglass False
10230   Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Form loading, ignore.
10240   Case Else
10250     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10260   End Select
10270   Resume EXITP

End Sub

Public Sub UpdateTATotal_CR(frm As Access.Form)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "UpdateTATotal_CR"

        Dim strAccountNo As String
        Dim varTmp00 As Variant, varTmp01 As Variant, varTmp02 As Variant, varTmp03 As Variant

10310   gstrFormQuerySpec = frm.Name
10320   With frm
10330     If IsNull(.cmbAccounts) = False Then
10340       If Trim(.cmbAccounts) <> vbNullString Then
10350         strAccountNo = .cmbAccounts
10360         If IsNull(.cmbAssets) = False Then
10370           If .cmbAssets > 0 Then
10380             varTmp00 = .TABalance
                  ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'TA',
                  ' ** EntryType = 'Credit', by specified FormRef('AccountNo', 'AssetNo').
10390             varTmp01 = DLookup("CalcAmount", "qryCheckReconcile_TACreditTotal_01a")
                  ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'TA',
                  ' ** EntryType = 'Debit', by specified FormRef('AccountNo', 'AssetNo').
10400             varTmp02 = DLookup("CalcAmount", "qryCheckReconcile_TADebitTotal_01a")
10410             varTmp00 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
10420             varTmp01 = ZeroIfNull(varTmp01)  ' ** Module Function: modStringFuncs.
10430             varTmp02 = ZeroIfNull(varTmp02)  ' ** Module Function: modStringFuncs.
10440             .TATotal = ((varTmp00 + varTmp01) - varTmp02)
10450             varTmp03 = .BSTotal
10460             varTmp03 = ZeroIfNull(varTmp03)  ' ** Module Function: modStringFuncs.
10470             .DifferenceBSTA = (varTmp03 - ((varTmp00 + varTmp01) - varTmp02))
10480           Else
10490             Select Case strAccountNo
                  Case "CRTC01"
10500               varTmp00 = .TABalance
                    ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'TA',
                    ' ** EntryType = 'Credit', for accountno = 'CRTC01'.
10510               varTmp01 = DLookup("CalcAmount", "qryCheckReconcile_TACreditTotal_01b")
                    ' ** tblCheckReconcile_Item, grouped and summed, just Source = 'TA',
                    ' ** EntryType = 'Debit', for accountno = 'CRTC01'.
10520               varTmp02 = DLookup("CalcAmount", "qryCheckReconcile_TADebitTotal_01b")
10530               varTmp00 = ZeroIfNull(varTmp00)  ' ** Module Function: modStringFuncs.
10540               varTmp01 = ZeroIfNull(varTmp01)  ' ** Module Function: modStringFuncs.
10550               varTmp02 = ZeroIfNull(varTmp02)  ' ** Module Function: modStringFuncs.
10560               .TATotal = ((varTmp00 + varTmp01) - varTmp02)
10570               varTmp03 = .BSTotal
10580               varTmp03 = ZeroIfNull(varTmp03)  ' ** Module Function: modStringFuncs.
10590               .DifferenceBSTA = (varTmp03 - ((varTmp00 + varTmp01) - varTmp02))
10600             Case Else
10610               .TATotal = 0
10620               .DifferenceBSTA = 0
10630             End Select
10640           End If
10650         Else
10660           .TATotal = 0
10670           .DifferenceBSTA = 0
10680         End If
10690       Else
10700         .TATotal = 0
10710         .DifferenceBSTA = 0
10720       End If
10730     Else
10740       .TATotal = 0
10750       .DifferenceBSTA = 0
10760     End If
10770   End With
10780   gstrFormQuerySpec = frm.Name

EXITP:
10790   Exit Sub

ERRH:
10800   DoCmd.Hourglass False
10810   Select Case ERR.Number
        Case Else
10820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10830   End Select
10840   Resume EXITP

End Sub

Public Sub Detail_Mouse_CR(blnAddAsset_Focus As Boolean, blnCalendar1_Focus As Boolean, blnPreviewReport_Focus As Boolean, blnPrintReport_Focus As Boolean, blnCheckAll_Focus As Boolean, blnCheckNone_Focus As Boolean, frm As Access.Form)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Mouse_CR"

10910   With frm
10920     If .cmdAddAsset_raised_focus_dots_img.Visible = True Or .cmdAddAsset_raised_focus_img.Visible = True Then
10930       Select Case blnAddAsset_Focus
            Case True
10940         .cmdAddAsset_raised_semifocus_dots_img.Visible = True
10950         .cmdAddAsset_raised_img.Visible = False
10960       Case False
10970         .cmdAddAsset_raised_img.Visible = True
10980         .cmdAddAsset_raised_semifocus_dots_img.Visible = False
10990       End Select
11000       .cmdAddAsset_raised_focus_img.Visible = False
11010       .cmdAddAsset_raised_focus_dots_img.Visible = False
11020       .cmdAddAsset_sunken_focus_dots_img.Visible = False
11030       .cmdAddAsset_raised_img_dis.Visible = False
11040     End If
11050     If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
11060       Select Case blnCalendar1_Focus
            Case True
11070         .cmdCalendar_raised_semifocus_dots_img.Visible = True
11080         .cmdCalendar_raised_img.Visible = False
11090       Case False
11100         .cmdCalendar_raised_img.Visible = True
11110         .cmdCalendar_raised_semifocus_dots_img.Visible = False
11120       End Select
11130       .cmdCalendar_raised_focus_img.Visible = False
11140       .cmdCalendar_raised_focus_dots_img.Visible = False
11150       .cmdCalendar_sunken_focus_dots_img.Visible = False
11160       .cmdCalendar_raised_img_dis.Visible = False
11170     End If
11180     If .cmdPreviewReport_raised_focus_dots_img.Visible = True Or .cmdPreviewReport_raised_focus_img.Visible = True Then
11190       Select Case blnPreviewReport_Focus
            Case True
11200         .cmdPreviewReport_raised_semifocus_dots_img.Visible = True
11210         .cmdPreviewReport_raised_img.Visible = False
11220       Case False
11230         .cmdPreviewReport_raised_img.Visible = True
11240         .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
11250       End Select
11260       .cmdPreviewReport_raised_focus_img.Visible = False
11270       .cmdPreviewReport_raised_focus_dots_img.Visible = False
11280       .cmdPreviewReport_sunken_focus_dots_img.Visible = False
11290       .cmdPreviewReport_raised_img_dis.Visible = False
11300     End If
11310     If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
11320       Select Case blnPrintReport_Focus
            Case True
11330         .cmdPrintReport_raised_semifocus_dots_img.Visible = True
11340         .cmdPrintReport_raised_img.Visible = False
11350       Case False
11360         .cmdPrintReport_raised_img.Visible = True
11370         .cmdPrintReport_raised_semifocus_dots_img.Visible = False
11380       End Select
11390       .cmdPrintReport_raised_focus_img.Visible = False
11400       .cmdPrintReport_raised_focus_dots_img.Visible = False
11410       .cmdPrintReport_sunken_focus_dots_img.Visible = False
11420       .cmdPrintReport_raised_img_dis.Visible = False
11430     End If
11440     If .cmdCheckAll_raised_focus_dots_img.Visible = True Or .cmdCheckAll_raised_focus_img.Visible = True Then
11450       Select Case blnCheckAll_Focus
            Case True
11460         .cmdCheckAll_raised_semifocus_dots_img.Visible = True
11470         .cmdCheckAll_raised_img.Visible = False
11480       Case False
11490         .cmdCheckAll_raised_img.Visible = True
11500         .cmdCheckAll_raised_semifocus_dots_img.Visible = False
11510       End Select
11520       .cmdCheckAll_raised_focus_img.Visible = False
11530       .cmdCheckAll_raised_focus_dots_img.Visible = False
11540       .cmdCheckAll_sunken_focus_dots_img.Visible = False
11550       .cmdCheckAll_raised_img_dis.Visible = False
11560     End If
11570     If .cmdCheckNone_raised_focus_dots_img.Visible = True Or .cmdCheckNone_raised_focus_img.Visible = True Then
11580       Select Case blnCheckNone_Focus
            Case True
11590         .cmdCheckNone_raised_semifocus_dots_img.Visible = True
11600         .cmdCheckNone_raised_img.Visible = False
11610       Case False
11620         .cmdCheckNone_raised_img.Visible = True
11630         .cmdCheckNone_raised_semifocus_dots_img.Visible = False
11640       End Select
11650       .cmdCheckNone_raised_focus_img.Visible = False
11660       .cmdCheckNone_raised_focus_dots_img.Visible = False
11670       .cmdCheckNone_sunken_focus_dots_img.Visible = False
11680       .cmdCheckNone_raised_img_dis.Visible = False
11690     End If
11700   End With

EXITP:
11710   Exit Sub

ERRH:
11720   Select Case ERR.Number
        Case Else
11730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11740   End Select
11750   Resume EXITP

End Sub

Public Sub PrevPrint_Handler_CR(strProc As String, blnPreviewReport_Focus As Boolean, blnPreviewReport_MouseDown As Boolean, blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean, lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant, frm As Access.Form)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "PrevPrint_Handler_CR"

        Dim dblTABalance As Double, dblOCTot As Double, dblBSTot As Double, dblTATot As Double, dblDiff As Double
        Dim strDocName As String, strAccountNo As String, strQryName As String
        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim varTmp00 As Variant

11810   With frm

11820     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
11830     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
11840     strEvent = Mid(strProc, (intPos01 + 1))
11850     strCtlName = Left(strProc, (intPos01 - 1))

11860     Select Case strEvent
          Case "GotFocus"
11870       Select Case strCtlName
            Case "cmdPreviewReport"
11880         blnPreviewReport_Focus = True
11890         .cmdPreviewReport_raised_semifocus_dots_img.Visible = True
11900         .cmdPreviewReport_raised_img.Visible = False
11910         .cmdPreviewReport_raised_focus_img.Visible = False
11920         .cmdPreviewReport_raised_focus_dots_img.Visible = False
11930         .cmdPreviewReport_sunken_focus_dots_img.Visible = False
11940         .cmdPreviewReport_raised_img_dis.Visible = False
11950       Case "cmdPrintReport"
11960         blnPrintReport_Focus = True
11970         .cmdPrintReport_raised_semifocus_dots_img.Visible = True
11980         .cmdPrintReport_raised_img.Visible = False
11990         .cmdPrintReport_raised_focus_img.Visible = False
12000         .cmdPrintReport_raised_focus_dots_img.Visible = False
12010         .cmdPrintReport_sunken_focus_dots_img.Visible = False
12020         .cmdPrintReport_raised_img_dis.Visible = False
12030       End Select
12040     Case "MouseDown"
12050       Select Case strCtlName
            Case "cmdPreviewReport"
12060         blnPreviewReport_MouseDown = True
12070         .cmdPreviewReport_sunken_focus_dots_img.Visible = True
12080         .cmdPreviewReport_raised_img.Visible = False
12090         .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
12100         .cmdPreviewReport_raised_focus_img.Visible = False
12110         .cmdPreviewReport_raised_focus_dots_img.Visible = False
12120         .cmdPreviewReport_raised_img_dis.Visible = False
12130       Case "cmdPrintReport"
12140         blnPrintReport_MouseDown = True
12150         .cmdPrintReport_sunken_focus_dots_img.Visible = True
12160         .cmdPrintReport_raised_img.Visible = False
12170         .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12180         .cmdPrintReport_raised_focus_img.Visible = False
12190         .cmdPrintReport_raised_focus_dots_img.Visible = False
12200         .cmdPrintReport_raised_img_dis.Visible = False
12210       End Select
12220     Case "MouseMove"
12230       Select Case strCtlName
            Case "cmdPreviewReport"
12240         If blnPreviewReport_MouseDown = False Then
12250           Select Case blnPreviewReport_Focus
                Case True
12260             .cmdPreviewReport_raised_focus_dots_img.Visible = True
12270             .cmdPreviewReport_raised_focus_img.Visible = False
12280           Case False
12290             .cmdPreviewReport_raised_focus_img.Visible = True
12300             .cmdPreviewReport_raised_focus_dots_img.Visible = False
12310           End Select
12320           .cmdPreviewReport_raised_img.Visible = False
12330           .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
12340           .cmdPreviewReport_sunken_focus_dots_img.Visible = False
12350           .cmdPreviewReport_raised_img_dis.Visible = False
12360         End If
12370       Case "cmdPrintReport"
12380         If blnPrintReport_MouseDown = False Then
12390           Select Case blnPrintReport_Focus
                Case True
12400             .cmdPrintReport_raised_focus_dots_img.Visible = True
12410             .cmdPrintReport_raised_focus_img.Visible = False
12420           Case False
12430             .cmdPrintReport_raised_focus_img.Visible = True
12440             .cmdPrintReport_raised_focus_dots_img.Visible = False
12450           End Select
12460           .cmdPrintReport_raised_img.Visible = False
12470           .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12480           .cmdPrintReport_sunken_focus_dots_img.Visible = False
12490           .cmdPrintReport_raised_img_dis.Visible = False
12500         End If
12510       End Select
12520     Case "MouseUp"
12530       Select Case strCtlName
            Case "cmdPreviewReport"
12540         .cmdPreviewReport_raised_focus_dots_img.Visible = True
12550         .cmdPreviewReport_raised_img.Visible = False
12560         .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
12570         .cmdPreviewReport_raised_focus_img.Visible = False
12580         .cmdPreviewReport_sunken_focus_dots_img.Visible = False
12590         .cmdPreviewReport_raised_img_dis.Visible = False
12600         blnPreviewReport_MouseDown = False
12610       Case "cmdPrintReport"
12620         .cmdPrintReport_raised_focus_dots_img.Visible = True
12630         .cmdPrintReport_raised_img.Visible = False
12640         .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12650         .cmdPrintReport_raised_focus_img.Visible = False
12660         .cmdPrintReport_sunken_focus_dots_img.Visible = False
12670         .cmdPrintReport_raised_img_dis.Visible = False
12680         blnPrintReport_MouseDown = False
12690       End Select
12700     Case "LostFocus"
12710       Select Case strCtlName
            Case "cmdPreviewReport"
12720         .cmdPreviewReport_raised_img.Visible = True
12730         .cmdPreviewReport_raised_semifocus_dots_img.Visible = False
12740         .cmdPreviewReport_raised_focus_img.Visible = False
12750         .cmdPreviewReport_raised_focus_dots_img.Visible = False
12760         .cmdPreviewReport_sunken_focus_dots_img.Visible = False
12770         .cmdPreviewReport_raised_img_dis.Visible = False
12780         blnPreviewReport_Focus = False
12790       Case "cmdPrintReport"
12800         .cmdPrintReport_raised_img.Visible = True
12810         .cmdPrintReport_raised_semifocus_dots_img.Visible = False
12820         .cmdPrintReport_raised_focus_img.Visible = False
12830         .cmdPrintReport_raised_focus_dots_img.Visible = False
12840         .cmdPrintReport_sunken_focus_dots_img.Visible = False
12850         .cmdPrintReport_raised_img_dis.Visible = False
12860         blnPrintReport_Focus = False
12870       End Select
12880     Case "Click"
12890       Select Case strCtlName
            Case "cmdPreviewReport"
12900         If IsNull(.cmbAccounts) = False Then
12910           If Trim(.cmbAccounts) <> vbNullString Then
12920             If .cmdUpdate.Enabled = True Then
12930               .cmdUpdate_Click  ' ** Form Procedure: frmCheckReconcile.
12940               DoEvents
12950             End If
12960             strAccountNo = .cmbAccounts
12970             gstrAccountNo = strAccountNo
12980             dblTABalance = Nz(.TABalance, 0)
12990             dblBSTot = Nz(.BSTotal, 0)
13000             dblTATot = Nz(.TATotal, 0)
13010             Select Case lngAssets_ThisAcct
                  Case 1&
                    ' ** For accounts with only 1 asset.
13020               Select Case strAccountNo
                    Case "CRTC01"
                      ' ** tblCheckReconcile_Check, grouped and summed, forCheckPaid = False,
                      ' ** croutchk_assign = 0, accountno = 'CRTC01'.
13030                 strQryName = "qryCheckReconcile_OutChecks_09_04"
13040               Case Else
                      ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                      ' ** croutchk_assign = 0, by specified GlobalVarGet("gstrAccountNo").
13050                 strQryName = "qryCheckReconcile_OutChecks_09_01"
13060               End Select
13070             Case Else
13080               Select Case .cmbAssets
                    Case arr_varAsset_ThisAcct(A_ASTNO, 0)
                      ' ** For accounts with 2 assets, and this is the first.
13090                 Select Case strAccountNo
                      Case "CRTC01"
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,1, accountno = 'CRTC01'.
13100                   strQryName = "qryCheckReconcile_OutChecks_09_05"
13110                 Case Else
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,1, by specified GlobalVarGet("gstrAccountNo").
13120                   strQryName = "qryCheckReconcile_OutChecks_09_02"
13130                 End Select
13140               Case arr_varAsset_ThisAcct(A_ASTNO, 1)
                      ' ** For accounts with 2 assets, and this is the second.
13150                 Select Case strAccountNo
                      Case "CRTC01"
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,2, accountno = 'CRTC01'.
13160                   strQryName = "qryCheckReconcile_OutChecks_09_06"
13170                 Case Else
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,2, by specified GlobalVarGet("gstrAccountNo").
13180                   strQryName = "qryCheckReconcile_OutChecks_09_03"
13190                 End Select
13200               End Select
13210             End Select
13220             varTmp00 = DLookup("CalcTotal", strQryName)
13230             Select Case IsNull(varTmp00)
                  Case True
13240               dblOCTot = 0#
13250             Case False
13260               dblOCTot = varTmp00
13270             End Select
13280             dblDiff = Nz(.DifferenceBSTA, 0)
13290             gstrFormQuerySpec = frm.Name
13300             gstrReportCallingForm = frm.Name  ' ** When not vbNullString, sets this form .Visible = False.
13310             strDocName = "rptCheckReconcile"
13320             DoCmd.OpenReport strDocName, acViewPreview, , , , CStr(dblTABalance) & "~" & CStr(dblBSTot) & "~" & _
                    CStr(dblTATot) & "~" & CStr(dblOCTot) & "~" & CStr(dblDiff)
13330             DoCmd.Maximize
13340             DoCmd.RunCommand acCmdFitToWindow
13350           Else
                  ' ** Nothing happening.
13360           End If
13370         Else
                ' ** Nothing happening.
13380         End If
13390       Case "cmdPrintReport"
13400         If IsNull(.cmbAccounts) = False Then
13410           If Trim(.cmbAccounts) <> vbNullString Then
13420             If .cmdUpdate.Enabled = True Then
13430               .cmdUpdate_Click  ' ** Form Procedure: frmCheckReconcile.
13440               DoEvents
13450             End If
13460             strAccountNo = .cmbAccounts
13470             gstrAccountNo = strAccountNo
13480             dblTABalance = Nz(.TABalance, 0)
13490             dblBSTot = Nz(.BSTotal, 0)
13500             dblTATot = Nz(.TATotal, 0)
13510             Select Case lngAssets_ThisAcct
                  Case 1&
                    ' ** For accounts with only 1 asset.
13520               Select Case strAccountNo
                    Case "CRTC01"
                      ' ** tblCheckReconcile_Check, grouped and summed, forCheckPaid = False,
                      ' ** croutchk_assign = 0, accountno = 'CRTC01'.
13530                 strQryName = "qryCheckReconcile_OutChecks_09_04"
13540               Case Else
                      ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                      ' ** croutchk_assign = 0, by specified GlobalVarGet("gstrAccountNo").
13550                 strQryName = "qryCheckReconcile_OutChecks_09_01"
13560               End Select
13570             Case Else
13580               Select Case .cmbAssets
                    Case arr_varAsset_ThisAcct(A_ASTNO, 0)
                      ' ** For accounts with 2 assets, and this is the first.
13590                 Select Case strAccountNo
                      Case "CRTC01"
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,1, accountno = 'CRTC01'.
13600                   strQryName = "qryCheckReconcile_OutChecks_09_05"
13610                 Case Else
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,1, by specified GlobalVarGet("gstrAccountNo").
13620                   strQryName = "qryCheckReconcile_OutChecks_09_02"
13630                 End Select
13640               Case arr_varAsset_ThisAcct(A_ASTNO, 1)
                      ' ** For accounts with 2 assets, and this is the second.
13650                 Select Case strAccountNo
                      Case "CRTC01"
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,2, accountno = 'CRTC01'.
13660                   strQryName = "qryCheckReconcile_OutChecks_09_06"
13670                 Case Else
                        ' ** tblCheckReconcile_Check, grouped and summed, for CheckPaid = False,
                        ' ** croutchk_assign = 0,2, by specified GlobalVarGet("gstrAccountNo").
13680                   strQryName = "qryCheckReconcile_OutChecks_09_03"
13690                 End Select
13700               End Select
13710             End Select
13720             varTmp00 = DLookup("CalcTotal", strQryName)
13730             Select Case IsNull(varTmp00)
                  Case True
13740               dblOCTot = 0#
13750             Case False
13760               dblOCTot = varTmp00
13770             End Select
13780             dblDiff = Nz(.DifferenceBSTA, 0)
13790             gstrFormQuerySpec = frm.Name
13800             gstrReportCallingForm = vbNullString  ' ** When vbNullString, leaves this form .Visible = True.
13810             strDocName = "rptCheckReconcile"
                  '##GTR_Ref: rptCheckReconcile
13820             DoCmd.OpenReport strDocName, acViewNormal, , , , CStr(dblTABalance) & "~" & CStr(dblBSTot) & "~" & _
                    CStr(dblTATot) & "~" & CStr(dblOCTot) & "~" & CStr(dblDiff)
13830           Else
                  ' ** Nothing happening.
13840           End If
13850         Else
                ' ** Nothing happening.
13860         End If
13870       End Select
13880     End Select

13890   End With

EXITP:
13900   Exit Sub

ERRH:
13910   DoCmd.Hourglass False
13920   DoCmd.Restore
13930   If frm.Visible = False Then
13940     frm.Visible = True
13950   End If
13960   Select Case ERR.Number
        Case Else
13970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13980   End Select
13990   Resume EXITP

End Sub

Public Sub Check_Handler_CR(strProc As String, blnCheckAll_Focus As Boolean, blnCheckAll_MouseDown As Boolean, blnCheckNone_Focus As Boolean, blnCheckNone_MouseDown As Boolean, frm As Access.Form)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "Check_Handler_CR"

        Dim rst As DAO.Recordset
        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim lngX As Long

14010   With frm

14020     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
14030     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
14040     strEvent = Mid(strProc, (intPos01 + 1))
14050     strCtlName = Left(strProc, (intPos01 - 1))

14060     Select Case strEvent
          Case "GotFocus"
14070       Select Case strCtlName
            Case "cmdCheckAll"
14080         blnCheckAll_Focus = True
14090         .cmdCheckAll_raised_semifocus_dots_img.Visible = True
14100         .cmdCheckAll_raised_img.Visible = False
14110         .cmdCheckAll_raised_focus_img.Visible = False
14120         .cmdCheckAll_raised_focus_dots_img.Visible = False
14130         .cmdCheckAll_sunken_focus_dots_img.Visible = False
14140         .cmdCheckAll_raised_img_dis.Visible = False
14150       Case "cmdCheckNone"
14160         blnCheckNone_Focus = True
14170         .cmdCheckNone_raised_semifocus_dots_img.Visible = True
14180         .cmdCheckNone_raised_img.Visible = False
14190         .cmdCheckNone_raised_focus_img.Visible = False
14200         .cmdCheckNone_raised_focus_dots_img.Visible = False
14210         .cmdCheckNone_sunken_focus_dots_img.Visible = False
14220         .cmdCheckNone_raised_img_dis.Visible = False
14230       End Select
14240     Case "MouseDown"
14250       Select Case strCtlName
            Case "cmdCheckAll"
14260         blnCheckAll_MouseDown = True
14270         .cmdCheckAll_sunken_focus_dots_img.Visible = True
14280         .cmdCheckAll_raised_img.Visible = False
14290         .cmdCheckAll_raised_semifocus_dots_img.Visible = False
14300         .cmdCheckAll_raised_focus_img.Visible = False
14310         .cmdCheckAll_raised_focus_dots_img.Visible = False
14320         .cmdCheckAll_raised_img_dis.Visible = False
14330       Case "cmdCheckNone"
14340         blnCheckNone_MouseDown = True
14350         .cmdCheckNone_sunken_focus_dots_img.Visible = True
14360         .cmdCheckNone_raised_img.Visible = False
14370         .cmdCheckNone_raised_semifocus_dots_img.Visible = False
14380         .cmdCheckNone_raised_focus_img.Visible = False
14390         .cmdCheckNone_raised_focus_dots_img.Visible = False
14400         .cmdCheckNone_raised_img_dis.Visible = False
14410       End Select
14420     Case "MouseMove"
14430       Select Case strCtlName
            Case "cmdCheckAll"
14440         If blnCheckAll_MouseDown = False Then
14450           Select Case blnCheckAll_Focus
                Case True
14460             .cmdCheckAll_raised_focus_dots_img.Visible = True
14470             .cmdCheckAll_raised_focus_img.Visible = False
14480           Case False
14490             .cmdCheckAll_raised_focus_img.Visible = True
14500             .cmdCheckAll_raised_focus_dots_img.Visible = False
14510           End Select
14520           .cmdCheckAll_raised_img.Visible = False
14530           .cmdCheckAll_raised_semifocus_dots_img.Visible = False
14540           .cmdCheckAll_sunken_focus_dots_img.Visible = False
14550           .cmdCheckAll_raised_img_dis.Visible = False
14560         End If
14570         If .cmdCheckNone_raised_focus_dots_img.Visible = True Or .cmdCheckNone_raised_focus_img.Visible = True Then
14580           Select Case blnCheckNone_Focus
                Case True
14590             .cmdCheckNone_raised_semifocus_dots_img.Visible = True
14600             .cmdCheckNone_raised_img.Visible = False
14610           Case False
14620             .cmdCheckNone_raised_img.Visible = True
14630             .cmdCheckNone_raised_semifocus_dots_img.Visible = False
14640           End Select
14650           .cmdCheckNone_raised_focus_img.Visible = False
14660           .cmdCheckNone_raised_focus_dots_img.Visible = False
14670           .cmdCheckNone_sunken_focus_dots_img.Visible = False
14680           .cmdCheckNone_raised_img_dis.Visible = False
14690         End If
14700       Case "cmdCheckNone"
14710         If blnCheckNone_MouseDown = False Then
14720           Select Case blnCheckNone_Focus
                Case True
14730             .cmdCheckNone_raised_focus_dots_img.Visible = True
14740             .cmdCheckNone_raised_focus_img.Visible = False
14750           Case False
14760             .cmdCheckNone_raised_focus_img.Visible = True
14770             .cmdCheckNone_raised_focus_dots_img.Visible = False
14780           End Select
14790           .cmdCheckNone_raised_img.Visible = False
14800           .cmdCheckNone_raised_semifocus_dots_img.Visible = False
14810           .cmdCheckNone_sunken_focus_dots_img.Visible = False
14820           .cmdCheckNone_raised_img_dis.Visible = False
14830         End If
14840         If .cmdCheckAll_raised_focus_dots_img.Visible = True Or .cmdCheckAll_raised_focus_img.Visible = True Then
14850           Select Case blnCheckAll_Focus
                Case True
14860             .cmdCheckAll_raised_semifocus_dots_img.Visible = True
14870             .cmdCheckAll_raised_img.Visible = False
14880           Case False
14890             .cmdCheckAll_raised_img.Visible = True
14900             .cmdCheckAll_raised_semifocus_dots_img.Visible = False
14910           End Select
14920           .cmdCheckAll_raised_focus_img.Visible = False
14930           .cmdCheckAll_raised_focus_dots_img.Visible = False
14940           .cmdCheckAll_sunken_focus_dots_img.Visible = False
14950           .cmdCheckAll_raised_img_dis.Visible = False
14960         End If
14970       End Select
14980     Case "MouseUp"
14990       Select Case strCtlName
            Case "cmdCheckAll"
15000         .cmdCheckAll_raised_focus_dots_img.Visible = True
15010         .cmdCheckAll_raised_img.Visible = False
15020         .cmdCheckAll_raised_semifocus_dots_img.Visible = False
15030         .cmdCheckAll_raised_focus_img.Visible = False
15040         .cmdCheckAll_sunken_focus_dots_img.Visible = False
15050         .cmdCheckAll_raised_img_dis.Visible = False
15060         blnCheckAll_MouseDown = False
15070       Case "cmdCheckNone"
15080         .cmdCheckNone_raised_focus_dots_img.Visible = True
15090         .cmdCheckNone_raised_img.Visible = False
15100         .cmdCheckNone_raised_semifocus_dots_img.Visible = False
15110         .cmdCheckNone_raised_focus_img.Visible = False
15120         .cmdCheckNone_sunken_focus_dots_img.Visible = False
15130         .cmdCheckNone_raised_img_dis.Visible = False
15140         blnCheckNone_MouseDown = False
15150       End Select
15160     Case "LostFocus"
15170       Select Case strCtlName
            Case "cmdCheckAll"
15180         .cmdCheckAll_raised_img.Visible = True
15190         .cmdCheckAll_raised_semifocus_dots_img.Visible = False
15200         .cmdCheckAll_raised_focus_img.Visible = False
15210         .cmdCheckAll_raised_focus_dots_img.Visible = False
15220         .cmdCheckAll_sunken_focus_dots_img.Visible = False
15230         .cmdCheckAll_raised_img_dis.Visible = False
15240         blnCheckAll_Focus = False
15250       Case "cmdCheckNone"
15260         .cmdCheckNone_raised_img.Visible = True
15270         .cmdCheckNone_raised_semifocus_dots_img.Visible = False
15280         .cmdCheckNone_raised_focus_img.Visible = False
15290         .cmdCheckNone_raised_focus_dots_img.Visible = False
15300         .cmdCheckNone_sunken_focus_dots_img.Visible = False
15310         .cmdCheckNone_raised_img_dis.Visible = False
15320         blnCheckNone_Focus = False
15330       End Select
15340     Case "Click"
15350       Select Case strCtlName
            Case "cmdCheckAll"
15360         lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
15370         If lngRecsCur > 0& Then
15380           Set rst = .frmCheckReconcile_Sub_OutChecks.Form.RecordsetClone
15390           With rst
15400             .MoveFirst
15410             For lngX = 1& To lngRecsCur
15420               If ![CheckPaid] = False Then
15430                 .Edit
15440                 ![CheckPaid] = True
15450                 .Update
15460                 frm.UpdateCheckArray ![journalno], ![croutchk_id], ![CheckPaid], ![CheckNum], ![description]  ' ** Form Procedure: frmCheckReconcile.
15470               End If
15480               If lngX < lngRecsCur Then .MoveNext
15490             Next
15500             .Close
15510           End With
15520           .frmCheckReconcile_Sub_OutChecks.Form.Requery
15530           If .cmdUpdate.Enabled = False Then
15540             .cmdUpdate.Enabled = True
15550           End If
15560         End If
15570       Case "cmdCheckNone"
15580         lngRecsCur = .frmCheckReconcile_Sub_OutChecks.Form.RecCnt  ' ** Form Function: frmCheckReconcile_Sub_OutChecks.
15590         If lngRecsCur > 0& Then
15600           Set rst = .frmCheckReconcile_Sub_OutChecks.Form.RecordsetClone
15610           With rst
15620             .MoveFirst
15630             For lngX = 1& To lngRecsCur
15640               If ![CheckPaid] = True Then
15650                 .Edit
15660                 ![CheckPaid] = False
15670                 .Update
15680                 frm.UpdateCheckArray ![journalno], ![croutchk_id], ![CheckPaid], ![CheckNum], ![description]  ' ** Form Procedure: frmCheckReconcile.
15690               End If
15700               If lngX < lngRecsCur Then .MoveNext
15710             Next
15720             .Close
15730           End With
15740           .frmCheckReconcile_Sub_OutChecks.Form.Requery
15750           If .cmdUpdate.Enabled = False Then
15760             .cmdUpdate.Enabled = True
15770           End If
15780         End If
15790       End Select
15800     End Select
15810   End With

EXITP:
15820   Set rst = Nothing
15830   Exit Sub

ERRH:
15840   Select Case ERR.Number
        Case Else
15850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15860   End Select
15870   Resume EXITP

End Sub

Public Sub RecalcTots_CR(strAccountNo As String, lngAssets_ThisAcct As Long, arr_varAsset_ThisAcct As Variant, frm As Access.Form)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcTots_CR"

15910   UpdateBSTotal_CR frm, lngAssets_ThisAcct, arr_varAsset_ThisAcct  ' ** Procedure: Above.
15920   UpdateTATotal_CR frm  ' ** Procedure: Above.
15930   TotOutChks frm, strAccountNo  ' ** Procedure: Above.
15940   frm.CalcTABalance  ' ** Form Procedure: frmCheckReconcile.
        ' ** CalcTABalance includes calss to TotOutChks(), UpdateTATotal_CR().

EXITP:
15950   Exit Sub

ERRH:
15960   Select Case ERR.Number
        Case Else
15970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15980   End Select
15990   Resume EXITP

End Sub
