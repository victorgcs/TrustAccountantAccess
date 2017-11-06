Attribute VB_Name = "modJrnlCol_Buttons"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modJrnlCol_Buttons"

'VGC 09/02/2017: CHANGES!

'tblJOURNAL_FIELD DOESN'T GET AUTOMATICALLY UPDATED!!!

' ** Array: arr_varSpecPurp().
Private lngSpecPurps As Long, arr_varSpecPurp() As Variant
Private Const SP_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const SP_NAM As Integer = 0
Private Const SP_ABL As Integer = 1

' ** Array: arr_varImgVar().
Private lngImgVars As Long, arr_varImgVar() As Variant
Private Const IV_ELEMS As Integer = 3  ' ** Array's first-element UBound().
Private Const IV_SFX As Integer = 0
Private Const IV_FOC As Integer = 1
Private Const IV_DWN As Integer = 2
Private Const IV_DIS As Integer = 3

Private blnAssetNew_Focus As Boolean, blnAssetNew_MouseDown As Boolean
Private blnLocNew_Focus As Boolean, blnLocNew_MouseDown As Boolean
Private blnRecurNew_Focus As Boolean, blnRecurNew_MouseDown As Boolean
Private blnRefresh_Focus As Boolean, blnRefresh_MouseDown As Boolean
Private blnPreviewReport_Focus As Boolean, blnPreviewReport_MouseDown As Boolean
Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean
Private blnSwitch_Focus As Boolean, blnSwitch_MouseDown As Boolean
Private blnMemoReveal_Focus As Boolean, blnMemoReveal_MouseDown As Boolean
Private blnUncomCom_Focus As Boolean, blnUncomCom_MouseDown As Boolean
Private blnUncomDel_Focus As Boolean, blnUncomDel_MouseDown As Boolean
Private blnUnCommitOne_Focus As Boolean, blnUnCommitOne_MouseDown As Boolean
Private blnMapDiv_Focus As Boolean, blnMapDiv_MouseDown As Boolean
Private blnMapInt_Focus As Boolean, blnMapInt_MouseDown As Boolean
Private blnPaid_Focus As Boolean, blnPaid_MouseDown As Boolean
Private blnMapSplit_Focus As Boolean, blnMapSplit_MouseDown As Boolean
Private blnMapLTCG_Focus As Boolean, blnMapLTCG_MouseDown As Boolean
Private blnMapLTCL_Focus As Boolean, blnMapLTCL_MouseDown As Boolean
Private blnMapSTCGL_Focus As Boolean, blnMapSTCGL_MouseDown As Boolean

Private lngTpp As Long
Private lngOpGrp_Height As Long, lngOpt_Offset As Long, lngOptLbl_Offset As Long
Private lngWideBtn_Width As Long, lngVLine_Left As Long, lngSizable_Offset As Long
Private lngAdd_Left As Long, lngEdit_Left As Long, lngDelete_Left As Long, lngClose_Left As Long
Private lngAssets_Left As Long, lngLoc_Left As Long, lngRecur_Left As Long, lngRefresh_Left As Long
' **

Public Sub JC_Btn_Set(strThisJType As String, blnPosted As Boolean, frmPar As Access.Form, Optional varNewRec As Variant)
' ** Special-Purpose Buttons.
' ** Called by:
' **   JC_Msc_Cur_Set(), Above
' **   frmJournal_Columns:
' **     cmdDelete_Click()
' **     cmdUncomDelAll_Click()
' **   frmJournal_Columns_Sub:
' **     Form_Timer()
' **     posted_AfterUpdate()
' **     accountno_AfterUpdate()
' **     accountno2_AfterUpdate()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Set"

        Dim blnNewRec As Boolean
        Dim lngRecsCur As Long, lngUncoms As Long
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngY As Long

110     With frmPar

120       JC_Btn_Load frmPar  ' ** Function: Below.

130       Select Case IsMissing(varNewRec)
          Case True
140         blnNewRec = False
150       Case False
160         blnNewRec = CBool(varNewRec)
170       End Select

180       Select Case blnPosted
          Case True
            ' ** Committed.

190         For lngX = 0& To (lngSpecPurps - 1&)
200           strTmp01 = arr_varSpecPurp(SP_NAM, lngX)
210           If InStr(arr_varSpecPurp(SP_NAM, lngX), "_Sold_") = 0 Then
                ' ** Enabled.
220             .Controls(strTmp01).Enabled = True
230             For lngY = 0& To (lngImgVars - 1&)
240               strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
250               If arr_varImgVar(IV_FOC, lngY) = False And _
                      arr_varImgVar(IV_DWN, lngY) = False And _
                      arr_varImgVar(IV_DIS, lngY) = False Then
260                 .Controls(strTmp02).Visible = True
270               Else
280                 .Controls(strTmp02).Visible = False
290               End If
300             Next
310             .cmdSpecPurp_Div_lbl.ForeColor = CLR_BLUGRY
320             .cmdSpecPurp_Div_lbl_dim_hi.Visible = False
330             .cmdSpecPurp_Int_lbl.ForeColor = CLR_BLUGRY
340             .cmdSpecPurp_Int_lbl_dim_hi.Visible = False
350             .cmdSpecPurp_Rec_lbl.ForeColor = CLR_BLUGRY
360             .cmdSpecPurp_Rec_lbl_dim_hi.Visible = False
370             .cmdSpecPurp_Misc_lbl.ForeColor = CLR_BLUGRY
380             .cmdSpecPurp_Misc_lbl_dim_hi.Visible = False
390             .cmdSpecPurp_Purch_lbl.ForeColor = CLR_BLUGRY
400             .cmdSpecPurp_Purch_lbl_dim_hi.Visible = False
410           Else
                ' ** Disabled.
420   On Error Resume Next
430             .Controls(strTmp01).Enabled = False
440             If ERR.Number <> 0 Then
                  ' ** 2164:  You can't disable a control while it has the focus.
                  ' ** 2165:  You can't hide a control that has the focus.
450   On Error GoTo ERRH
460               DoCmd.SelectObject acForm, frmPar.Name, False
470               If frmPar.cmdAdd.Enabled = True Then
480                 frmPar.cmdAdd.SetFocus
490               Else
500                 frmPar.cmdClose.SetFocus
510               End If
520               .Controls(strTmp01).Enabled = False
530             Else
540   On Error GoTo ERRH
550             End If
560             For lngY = 0& To (lngImgVars - 1&)
570               strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
580               If arr_varImgVar(IV_DIS, lngY) = True Then
590                 .Controls(strTmp02).Visible = True
600               Else
610                 .Controls(strTmp02).Visible = False
620               End If
630             Next
640             .cmdSpecPurp_Sold_lbl.ForeColor = WIN_CLR_DISF
650             .cmdSpecPurp_Sold_lbl_dim_hi.Visible = True
660           End If
670         Next

680       Case False
            ' ** Uncommitted.

690         Select Case strThisJType

            Case "Sold"

700           For lngX = 0& To (lngSpecPurps - 1&)
710             strTmp01 = arr_varSpecPurp(SP_NAM, lngX)
720             If InStr(arr_varSpecPurp(SP_NAM, lngX), "_Sold_") <> 0 Then
                  ' ** Enabled.
730               .Controls(strTmp01).Enabled = True
740               For lngY = 0& To (lngImgVars - 1&)
750                 strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
760                 If arr_varImgVar(IV_FOC, lngY) = False And _
                        arr_varImgVar(IV_DWN, lngY) = False And _
                        arr_varImgVar(IV_DIS, lngY) = False Then
770                   .Controls(strTmp02).Visible = True
780                 Else
790                   .Controls(strTmp02).Visible = False
800                 End If
810               Next
820               .cmdSpecPurp_Sold_lbl.ForeColor = CLR_BLUGRY
830               .cmdSpecPurp_Sold_lbl_dim_hi.Visible = False
840             Else
                  ' ** Disabled.
850   On Error Resume Next
860               .Controls(strTmp01).Enabled = False
870               If ERR.Number <> 0 Then
                    ' ** 2164:  You can't disable a control while it has the focus.
                    ' ** 2165:  You can't hide a control that has the focus.
880   On Error GoTo ERRH
890                 DoCmd.SelectObject acForm, frmPar.Name, False
900                 If frmPar.cmdAdd.Enabled = True Then
910                   frmPar.cmdAdd.SetFocus
920                 Else
930                   frmPar.cmdClose.SetFocus
940                 End If
950                 .Controls(strTmp01).Enabled = False
960               Else
970   On Error GoTo ERRH
980               End If
990               For lngY = 0& To (lngImgVars - 1&)
1000                strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
1010                If arr_varImgVar(IV_DIS, lngY) = True Then
1020                  .Controls(strTmp02).Visible = True
1030                Else
1040                  .Controls(strTmp02).Visible = False
1050                End If
1060              Next
1070              .cmdSpecPurp_Div_lbl.ForeColor = WIN_CLR_DISF
1080              .cmdSpecPurp_Div_lbl_dim_hi.Visible = True
1090              .cmdSpecPurp_Int_lbl.ForeColor = WIN_CLR_DISF
1100              .cmdSpecPurp_Int_lbl_dim_hi.Visible = True
1110              .cmdSpecPurp_Rec_lbl.ForeColor = WIN_CLR_DISF
1120              .cmdSpecPurp_Rec_lbl_dim_hi.Visible = True
1130              .cmdSpecPurp_Misc_lbl.ForeColor = WIN_CLR_DISF
1140              .cmdSpecPurp_Misc_lbl_dim_hi.Visible = True
1150              .cmdSpecPurp_Purch_lbl.ForeColor = WIN_CLR_DISF
1160              .cmdSpecPurp_Purch_lbl_dim_hi.Visible = True
1170            End If
1180          Next

1190        Case "Dividend", "Interest", "Purchase", "Deposit", "Withdrawn", "Liability (+)", "Liability (-)", _
                "Misc.", "Paid", "Received"

1200          For lngX = 0& To (lngSpecPurps - 1&)
1210            strTmp01 = arr_varSpecPurp(SP_NAM, lngX)
                ' ** Disabled.
1220  On Error Resume Next
1230            .Controls(strTmp01).Enabled = False
1240            If ERR.Number <> 0 Then
                  ' ** 2164:  You can't disable a control while it has the focus.
                  ' ** 2165:  You can't hide a control that has the focus.
1250  On Error GoTo ERRH
1260              DoCmd.SelectObject acForm, frmPar.Name, False
1270              If frmPar.cmdAdd.Enabled = True Then
1280                frmPar.cmdAdd.SetFocus
1290              Else
1300                frmPar.cmdClose.SetFocus
1310              End If
1320              .Controls(strTmp01).Enabled = False
1330            Else
1340  On Error GoTo ERRH
1350            End If
1360            For lngY = 0& To (lngImgVars - 1&)
1370              strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
1380              If arr_varImgVar(IV_DIS, lngY) = True Then
1390                .Controls(strTmp02).Visible = True
1400              Else
1410                .Controls(strTmp02).Visible = False
1420              End If
1430            Next
1440          Next
1450          .cmdSpecPurp_Div_lbl.ForeColor = WIN_CLR_DISF
1460          .cmdSpecPurp_Div_lbl_dim_hi.Visible = True
1470          .cmdSpecPurp_Int_lbl.ForeColor = WIN_CLR_DISF
1480          .cmdSpecPurp_Int_lbl_dim_hi.Visible = True
1490          .cmdSpecPurp_Sold_lbl.ForeColor = WIN_CLR_DISF
1500          .cmdSpecPurp_Sold_lbl_dim_hi.Visible = True
1510          .cmdSpecPurp_Rec_lbl.ForeColor = WIN_CLR_DISF
1520          .cmdSpecPurp_Rec_lbl_dim_hi.Visible = True
1530          .cmdSpecPurp_Misc_lbl.ForeColor = WIN_CLR_DISF
1540          .cmdSpecPurp_Misc_lbl_dim_hi.Visible = True
1550          .cmdSpecPurp_Purch_lbl.ForeColor = WIN_CLR_DISF
1560          .cmdSpecPurp_Purch_lbl_dim_hi.Visible = True

1570        Case Else  ' ** vbNullString.
              ' ** On a completely empty record (save transdate),
              ' ** I want to keep the top buttons active!

1580          Select Case blnNewRec
              Case True
1590            For lngX = 0& To (lngSpecPurps - 1&)
1600              strTmp01 = arr_varSpecPurp(SP_NAM, lngX)
1610              If InStr(arr_varSpecPurp(SP_NAM, lngX), "_Sold_") = 0 Then
                    ' ** Enabled.
1620                .Controls(strTmp01).Enabled = True
1630                For lngY = 0& To (lngImgVars - 1&)
1640                  strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
1650                  If arr_varImgVar(IV_FOC, lngY) = False And _
                          arr_varImgVar(IV_DWN, lngY) = False And _
                          arr_varImgVar(IV_DIS, lngY) = False Then
1660                    If gblnGoToReport = True And gblnSpecialCapGainLoss = True Then
1670                      If (.GoToReport_arw_mapsplit_img.Visible = True And InStr(strTmp02, "cmdSpecPurp_Misc_MapSTCGL") > 0) Or _
                              (.GoToReport_arw_mapstcgl_img.Visible = True And InStr(strTmp02, "cmdSpecPurp_Misc_MapLTCL") > 0) Then
                            ' ** No, don't turn them on!
1680                      Else
1690                        .Controls(strTmp02).Visible = True
1700                      End If
1710                    Else
1720                      .Controls(strTmp02).Visible = True
1730                    End If
1740                  Else
1750                    .Controls(strTmp02).Visible = False
1760                  End If
1770                Next
1780                .cmdSpecPurp_Div_lbl.ForeColor = CLR_BLUGRY
1790                .cmdSpecPurp_Div_lbl_dim_hi.Visible = False
1800                .cmdSpecPurp_Int_lbl.ForeColor = CLR_BLUGRY
1810                .cmdSpecPurp_Int_lbl_dim_hi.Visible = False
1820                .cmdSpecPurp_Rec_lbl.ForeColor = CLR_BLUGRY
1830                .cmdSpecPurp_Rec_lbl_dim_hi.Visible = False
1840                .cmdSpecPurp_Misc_lbl.ForeColor = CLR_BLUGRY
1850                .cmdSpecPurp_Misc_lbl_dim_hi.Visible = False
1860                .cmdSpecPurp_Purch_lbl.ForeColor = CLR_BLUGRY
1870                .cmdSpecPurp_Purch_lbl_dim_hi.Visible = False
1880              Else
                    ' ** Disabled.
1890  On Error Resume Next
1900                .Controls(strTmp01).Enabled = False
1910                If ERR.Number <> 0 Then
                      ' ** 2164:  You can't disable a control while it has the focus.
                      ' ** 2165:  You can't hide a control that has the focus.
1920  On Error GoTo ERRH
1930                  DoCmd.SelectObject acForm, frmPar.Name, False
1940                  If frmPar.cmdAdd.Enabled = True Then
1950                    frmPar.cmdAdd.SetFocus
1960                  Else
1970                    frmPar.cmdClose.SetFocus
1980                  End If
1990                  .Controls(strTmp01).Enabled = False
2000                Else
2010  On Error GoTo ERRH
2020                End If
2030                For lngY = 0& To (lngImgVars - 1&)
2040                  strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
2050                  If arr_varImgVar(IV_DIS, lngY) = True Then
2060                    .Controls(strTmp02).Visible = True
2070                  Else
2080                    .Controls(strTmp02).Visible = False
2090                  End If
2100                Next
2110                .cmdSpecPurp_Sold_lbl.ForeColor = WIN_CLR_DISF
2120                .cmdSpecPurp_Sold_lbl_dim_hi.Visible = True
2130              End If
2140            Next
2150          Case False
2160            For lngX = 0& To (lngSpecPurps - 1&)
2170              strTmp01 = arr_varSpecPurp(SP_NAM, lngX)
                  ' ** Disabled.
2180  On Error Resume Next
2190              .Controls(strTmp01).Enabled = False
2200              If ERR.Number <> 0 Then
                    ' ** 2164:  You can't disable a control while it has the focus.
                    ' ** 2165:  You can't hide a control that has the focus.
2210  On Error GoTo ERRH
2220                DoCmd.SelectObject acForm, frmPar.Name, False
2230                If frmPar.cmdAdd.Enabled = True Then
2240                  frmPar.cmdAdd.SetFocus
2250                Else
2260                  frmPar.cmdClose.SetFocus
2270                End If
2280                .Controls(strTmp01).Enabled = False
2290              Else
2300  On Error GoTo ERRH
2310              End If
2320              For lngY = 0& To (lngImgVars - 1&)
2330                strTmp02 = strTmp01 & arr_varImgVar(IV_SFX, lngY)
2340                If arr_varImgVar(IV_DIS, lngY) = True Then
2350                  .Controls(strTmp02).Visible = True
2360                Else
2370                  .Controls(strTmp02).Visible = False
2380                End If
2390              Next
2400            Next
2410            .cmdSpecPurp_Div_lbl.ForeColor = WIN_CLR_DISF
2420            .cmdSpecPurp_Div_lbl_dim_hi.Visible = True
2430            .cmdSpecPurp_Int_lbl.ForeColor = WIN_CLR_DISF
2440            .cmdSpecPurp_Int_lbl_dim_hi.Visible = True
2450            .cmdSpecPurp_Sold_lbl.ForeColor = WIN_CLR_DISF
2460            .cmdSpecPurp_Sold_lbl_dim_hi.Visible = True
2470            .cmdSpecPurp_Rec_lbl.ForeColor = WIN_CLR_DISF
2480            .cmdSpecPurp_Rec_lbl_dim_hi.Visible = True
2490            .cmdSpecPurp_Misc_lbl.ForeColor = WIN_CLR_DISF
2500            .cmdSpecPurp_Misc_lbl_dim_hi.Visible = True
2510            .cmdSpecPurp_Purch_lbl.ForeColor = WIN_CLR_DISF
2520            .cmdSpecPurp_Purch_lbl_dim_hi.Visible = True
2530          End Select

2540        End Select  ' ** strThisJType.
2550      End Select  ' ** blnPosted.

2560      lngRecsCur = frmPar.frmJournal_Columns_Sub.Form.RecCnt  ' ** Form Function: frmJournal_Columns_Sub.
2570  On Error Resume Next
2580      .cmdUncomComAll.Enabled = False
2590      .cmdUncomDelAll.Enabled = False
2600      .cmdUnCommitOne.Enabled = False
2610      If ERR.Number <> 0 Then
2620  On Error GoTo ERRH
2630        DoCmd.SelectObject acForm, .Name, False
2640        .cmdAdd.SetFocus
2650        DoEvents
2660      Else
2670  On Error GoTo ERRH
2680      End If
2690      If lngRecsCur = 0& Then
2700        .cmdUncomComAll.Enabled = False
2710        .cmdUncomDelAll.Enabled = False
2720        .cmdUnCommitOne.Enabled = False
2730        .cmdUncomComAll_raised_img.Visible = False
2740        .cmdUncomComAll_raised_focus_dots_img.Visible = False
2750        .cmdUncomComAll_sunken_focus_dots_img.Visible = False
2760        .cmdUncomDelAll_raised_img.Visible = False
2770        .cmdUncomDelAll_raised_focus_dots_img.Visible = False
2780        .cmdUncomDelAll_sunken_focus_dots_img.Visible = False
2790        .cmdUnCommitOne_raised_img.Visible = False
2800        .cmdUnCommitOne_raised_focus_dots_img.Visible = False
2810        .cmdUnCommitOne_sunken_focus_dots_img.Visible = False
2820        If strThisJType = "Paid" And .JrnlMemo_Memo.Visible = True Then
2830          .cmdUncom_lbl.Visible = False
2840          .cmdUncom_lbl_dim_hi.Visible = False
2850          .cmdUncomComAll_raised_img_dis.Visible = False
2860          .cmdUncomDelAll_raised_img_dis.Visible = False
2870          .cmdUnCommitOne_raised_img_dis.Visible = False
2880        Else
2890          .cmdUncom_lbl.ForeColor = WIN_CLR_DISF
2900          .cmdUncom_lbl.Visible = True
2910          .cmdUncom_lbl_dim_hi.Visible = True
2920          .cmdUncomComAll_raised_img_dis.Visible = True
2930          .cmdUncomDelAll_raised_img_dis.Visible = True
2940          .cmdUnCommitOne_raised_img_dis.Visible = True
2950        End If
2960      Else
2970        lngUncoms = .RecsTot_Uncommitted
2980        If lngUncoms = 0& Then
2990          .cmdUncomComAll.Enabled = False
3000          .cmdUncomDelAll.Enabled = False
3010          .cmdUnCommitOne.Enabled = False
3020          If strThisJType = "Paid" And .JrnlMemo_Memo.Visible = True Then
3030            .cmdUncom_lbl.Visible = False
3040            .cmdUncom_lbl_dim_hi.Visible = False
3050            .cmdUncomComAll_raised_img_dis.Visible = False
3060            .cmdUncomDelAll_raised_img_dis.Visible = False
3070            .cmdUnCommitOne_raised_img_dis.Visible = False
3080          Else
3090            .cmdUncom_lbl.ForeColor = WIN_CLR_DISF
3100            .cmdUncom_lbl.Visible = True
3110            .cmdUncom_lbl_dim_hi.Visible = True
3120            .cmdUncomComAll_raised_img_dis.Visible = True
3130            .cmdUncomDelAll_raised_img_dis.Visible = True
3140            .cmdUnCommitOne_raised_img_dis.Visible = True
3150          End If
3160          .cmdUncomComAll_raised_img.Visible = False
3170          .cmdUncomComAll_raised_focus_dots_img.Visible = False
3180          .cmdUncomComAll_sunken_focus_dots_img.Visible = False
3190          .cmdUncomDelAll_raised_img.Visible = False
3200          .cmdUncomDelAll_raised_focus_dots_img.Visible = False
3210          .cmdUncomDelAll_sunken_focus_dots_img.Visible = False
3220          .cmdUnCommitOne_raised_img.Visible = False
3230          .cmdUnCommitOne_raised_focus_dots_img.Visible = False
3240          .cmdUnCommitOne_sunken_focus_dots_img.Visible = False
3250        Else
3260          If strThisJType = "Paid" And .JrnlMemo_Memo.Visible = True Then
3270            .cmdUncom_lbl.Visible = False
3280            .cmdUncom_lbl_dim_hi.Visible = False
3290            .cmdUncomComAll.Enabled = False
3300            .cmdUncomDelAll.Enabled = False
3310            .cmdUnCommitOne.Enabled = False
3320            .cmdUncomComAll_raised_img.Visible = False
3330            .cmdUncomDelAll_raised_img.Visible = False
3340            .cmdUnCommitOne_raised_img.Visible = False
3350          Else
3360            .cmdUncom_lbl.ForeColor = CLR_BLUGRY
3370            .cmdUncom_lbl.Visible = True
3380            .cmdUncom_lbl_dim_hi.Visible = False
3390            .cmdUncomComAll.Enabled = True
3400            .cmdUncomDelAll.Enabled = True
3410            .cmdUnCommitOne.Enabled = True
3420            .cmdUncomDelAll_raised_img.Visible = True
3430            .cmdUncomComAll_raised_img.Visible = True
3440            .cmdUnCommitOne_raised_img.Visible = True
3450          End If
3460          .cmdUncomComAll_raised_focus_dots_img.Visible = False
3470          .cmdUncomComAll_sunken_focus_dots_img.Visible = False
3480          .cmdUncomComAll_raised_img_dis.Visible = False
3490          .cmdUncomDelAll_raised_focus_dots_img.Visible = False
3500          .cmdUncomDelAll_sunken_focus_dots_img.Visible = False
3510          .cmdUncomDelAll_raised_img_dis.Visible = False
3520          .cmdUnCommitOne_raised_focus_dots_img.Visible = False
3530          .cmdUnCommitOne_sunken_focus_dots_img.Visible = False
3540          .cmdUnCommitOne_raised_img_dis.Visible = False
3550        End If
3560      End If

3570    End With

EXITP:
3580    Exit Sub

ERRH:
3590    Select Case ERR.Number
        Case Else
3600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3610    End Select
3620    Resume EXITP

End Sub

Private Function JC_Btn_Load(frmPar As Access.Form) As Boolean
' ** Called by:
' **   JC_Btn_Set(), Above

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Load"

        Dim ctl As Access.Control
        Dim lngE As Long
        Dim blnRetVal As Boolean

3710    blnRetVal = True

3720    If lngSpecPurps = 0& Or IsEmpty(arr_varSpecPurp) Then

3730      lngSpecPurps = 0&
3740      ReDim arr_varSpecPurp(SP_ELEMS, 0)

          ' ** Get the names of the special purpose buttons.
3750      For Each ctl In frmPar.FormHeader.Controls
3760        With ctl
3770          If .ControlType = acCommandButton Then
3780            If Left(.Name, 12) = "cmdSpecPurp_" And InStr(.Name, "CostAdj") = 0 Then
3790              lngSpecPurps = lngSpecPurps + 1&
3800              lngE = lngSpecPurps - 1&
3810              ReDim Preserve arr_varSpecPurp(SP_ELEMS, lngE)
3820              arr_varSpecPurp(SP_NAM, lngE) = .Name
3830              arr_varSpecPurp(SP_ABL, lngE) = CBool(True)
3840            End If
3850          End If
3860        End With
3870      Next

3880      lngImgVars = 0&
3890      ReDim arr_varImgVar(IV_ELEMS, 0)

          ' ** Set the 6 image name variants.
3900      lngImgVars = lngImgVars + 1&
3910      lngE = lngImgVars - 1&
3920      ReDim Preserve arr_varImgVar(IV_ELEMS, lngE)
3930      arr_varImgVar(IV_SFX, lngE) = "_raised_img"
3940      arr_varImgVar(IV_FOC, lngE) = CBool(False)
3950      arr_varImgVar(IV_DWN, lngE) = CBool(False)
3960      arr_varImgVar(IV_DIS, lngE) = CBool(False)

3970      lngImgVars = lngImgVars + 1&
3980      lngE = lngImgVars - 1&
3990      ReDim Preserve arr_varImgVar(IV_ELEMS, lngE)
4000      arr_varImgVar(IV_SFX, lngE) = "_raised_semifocus_dots_img"
4010      arr_varImgVar(IV_FOC, lngE) = CBool(False)
4020      arr_varImgVar(IV_DWN, lngE) = CBool(False)
4030      arr_varImgVar(IV_DIS, lngE) = CBool(False)

4040      lngImgVars = lngImgVars + 1&
4050      lngE = lngImgVars - 1&
4060      ReDim Preserve arr_varImgVar(IV_ELEMS, lngE)
4070      arr_varImgVar(IV_SFX, lngE) = "_raised_focus_img"
4080      arr_varImgVar(IV_FOC, lngE) = CBool(False)
4090      arr_varImgVar(IV_DWN, lngE) = CBool(False)
4100      arr_varImgVar(IV_DIS, lngE) = CBool(False)

4110      lngImgVars = lngImgVars + 1&
4120      lngE = lngImgVars - 1&
4130      ReDim Preserve arr_varImgVar(IV_ELEMS, lngE)
4140      arr_varImgVar(IV_SFX, lngE) = "_raised_focus_dots_img"
4150      arr_varImgVar(IV_FOC, lngE) = CBool(True)
4160      arr_varImgVar(IV_DWN, lngE) = CBool(False)
4170      arr_varImgVar(IV_DIS, lngE) = CBool(False)

4180      lngImgVars = lngImgVars + 1&
4190      lngE = lngImgVars - 1&
4200      ReDim Preserve arr_varImgVar(IV_ELEMS, lngE)
4210      arr_varImgVar(IV_SFX, lngE) = "_sunken_focus_dots_img"
4220      arr_varImgVar(IV_FOC, lngE) = CBool(False)
4230      arr_varImgVar(IV_DWN, lngE) = CBool(True)
4240      arr_varImgVar(IV_DIS, lngE) = CBool(False)

4250      lngImgVars = lngImgVars + 1&
4260      lngE = lngImgVars - 1&
4270      ReDim Preserve arr_varImgVar(IV_ELEMS, lngE)
4280      arr_varImgVar(IV_SFX, lngE) = "_raised_img_dis"
4290      arr_varImgVar(IV_FOC, lngE) = CBool(False)
4300      arr_varImgVar(IV_DWN, lngE) = CBool(False)
4310      arr_varImgVar(IV_DIS, lngE) = CBool(True)

4320    End If

EXITP:
4330    Set ctl = Nothing
4340    JC_Btn_Load = blnRetVal
4350    Exit Function

ERRH:
4360    blnRetVal = False
4370    Select Case ERR.Number
        Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Function

Public Sub JC_Btn_Resize_Btn(frm As Access.Form, blnToNorm As Boolean, lngVLine_Left_New As Long, lngVLine_Left_Diff As Long)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Resize_Btn"

        Dim lngTweak As Long

4510    With frm
4520      Select Case blnToNorm

          Case True
            ' ** Buttons back to normal.

4530        If JC_Btn_Resize_Norm(frm) = False Then  ' ** Function: Below.
4540          .JrnlCol_FocusSet blnAssetNew_Focus, blnLocNew_Focus, blnRecurNew_Focus, blnRefresh_Focus  ' ** Form Procedure: frmJournal_Columns.
4550          .cmdAssetNew_raised_img.Visible = False
4560          .cmdAssetNew_raised_semifocus_dots_img.Visible = False
4570          .cmdAssetNew_raised_focus_img.Visible = False
4580          .cmdAssetNew_raised_focus_dots_img.Visible = False
4590          .cmdAssetNew_sunken_focus_dots_img.Visible = False
4600          .cmdAssetNew_raised_img_dis.Visible = False
4610          .cmdRefresh_raised_img.Visible = False
4620          .cmdRefresh_raised_semifocus_dots_img.Visible = False
4630          .cmdRefresh_raised_focus_img.Visible = False
4640          .cmdRefresh_raised_focus_dots_img.Visible = False
4650          .cmdRefresh_sunken_focus_dots_img.Visible = False
4660          .cmdRefresh_raised_img_dis.Visible = False
4670          .cmdLocNew_raised_img.Visible = False
4680          .cmdLocNew_raised_semifocus_dots_img.Visible = False
4690          .cmdLocNew_raised_focus_img.Visible = False
4700          .cmdLocNew_raised_focus_dots_img.Visible = False
4710          .cmdLocNew_sunken_focus_dots_img.Visible = False
4720          .cmdLocNew_raised_img_dis.Visible = False
4730          .cmdRecurNew_raised_img.Visible = False
4740          .cmdRecurNew_raised_semifocus_dots_img.Visible = False
4750          .cmdRecurNew_raised_focus_img.Visible = False
4760          .cmdRecurNew_raised_focus_dots_img.Visible = False
4770          .cmdRecurNew_sunken_focus_dots_img.Visible = False
4780          .cmdRecurNew_raised_img_dis.Visible = False
4790          .cmdAssetNew.Transparent = False
4800          .cmdAssetNew.Top = .cmdClose.Top  ' ** Moving down.
4810          .cmdAssetNew.Left = lngAssets_Left
4820          .cmdAssetNew.Width = lngWideBtn_Width
4830          .cmdAssetNew.Height = .cmdClose.Height
4840          .cmdLocNew.Transparent = False
4850          .cmdLocNew.Top = .cmdClose.Top  ' ** Moving down.
4860          .cmdLocNew.Left = lngLoc_Left
4870          .cmdLocNew.Width = lngWideBtn_Width
4880          .cmdLocNew.Height = .cmdClose.Height
4890          .cmdRecurNew.Transparent = False
4900          .cmdRecurNew.Top = .cmdClose.Top  ' ** Moving up.
4910          .cmdRecurNew.Left = lngRecur_Left
4920          .cmdRecurNew.Width = lngWideBtn_Width
4930          .cmdRecurNew.Height = .cmdClose.Height
4940          .Footer_vline03.Visible = True
4950          .Footer_vline04.Visible = True
4960          .cmdRefresh.Transparent = False
4970          .cmdRefresh.Top = .cmdClose.Top  ' ** Moving up.
4980          .cmdRefresh.Left = lngRefresh_Left
4990          .cmdRefresh.Width = .cmdClose.Width
5000          .cmdRefresh.Height = .cmdClose.Height
              'JC_Btn_BoolReset True  ' ** Procedure: Below.
5010        End If

5020        .Footer_vline05.Left = lngVLine_Left
5030        .Footer_vline06.Left = (lngVLine_Left + lngTpp)
5040        .cmdAdd_box.Left = (lngVLine_Left + (2& * lngTpp))
5050        .cmdAdd.Left = lngAdd_Left
5060        .cmdEdit.Left = lngEdit_Left
5070        .cmdDelete.Left = lngDelete_Left
5080        .Footer_vline07.Left = ((.cmdAdd_box.Left + .cmdAdd_box.Width) + lngTpp)
5090        .Footer_vline08.Left = (.Footer_vline07.Left + lngTpp)
5100        .cmdClose.Left = lngClose_Left
5110        .Sizable_lbl1.Left = (((.frmJournal_Columns_Sub_box.Left + .frmJournal_Columns_Sub_box.Width) + _
              lngSizable_Offset) - .Sizable_lbl1.Width)
5120        .Sizable_lbl2.Left = .Sizable_lbl1.Left

5130      Case False
            ' ** Smaller buttons take over.

5140        If JC_Btn_Resize_Norm(frm) = True Then  ' ** Function: Below.
              ' ** Adjust Height before Top.
              'JC_Btn_BoolReset True  ' ** Procedure: Below.
5150          .cmdAssetNew.Transparent = True
5160          .cmdAssetNew.Left = .cmdAssetNew_raised_img.Left
5170          .cmdAssetNew.Width = .cmdAssetNew_raised_img.Width
5180          .cmdAssetNew.Height = .cmdAssetNew_raised_img.Height
5190          .cmdAssetNew.Top = .cmdAssetNew_raised_img.Top
5200          .cmdRefresh.Transparent = True
5210          .cmdRefresh.Left = .cmdRefresh_raised_img.Left
5220          .cmdRefresh.Width = .cmdRefresh_raised_img.Width
5230          .cmdRefresh.Height = .cmdRefresh_raised_img.Height
5240          .cmdRefresh.Top = .cmdRefresh_raised_img.Top
5250          .cmdLocNew.Transparent = True
5260          .cmdLocNew.Left = .cmdLocNew_raised_img.Left
5270          .cmdLocNew.Width = .cmdLocNew_raised_img.Width
5280          .cmdLocNew.Height = .cmdLocNew_raised_img.Height
5290          .cmdLocNew.Top = .cmdLocNew_raised_img.Top
5300          .cmdRecurNew.Transparent = True
5310          .cmdRecurNew.Left = .cmdRecurNew_raised_img.Left
5320          .cmdRecurNew.Width = .cmdRecurNew_raised_img.Width
5330          .cmdRecurNew.Height = .cmdRecurNew_raised_img.Height
5340          .cmdRecurNew.Top = .cmdRecurNew_raised_img.Top
5350          .Footer_vline03.Visible = False
5360          .Footer_vline04.Visible = False
5370          Select Case .cmdAssetNew.Enabled
              Case True
5380            Select Case blnAssetNew_Focus
                Case True
5390              .cmdAssetNew_raised_semifocus_dots_img.Visible = True
5400            Case False
5410              .cmdAssetNew_raised_img.Visible = True
5420            End Select
5430          Case False
5440            .cmdAssetNew_raised_img_dis.Visible = True
5450          End Select
5460          Select Case .cmdLocNew.Enabled
              Case True
5470            Select Case blnLocNew_Focus
                Case True
5480              .cmdLocNew_raised_semifocus_dots_img.Visible = True
5490            Case False
5500              .cmdLocNew_raised_img.Visible = True
5510            End Select
5520          Case False
5530            .cmdLocNew_raised_img_dis.Visible = True
5540          End Select
5550          Select Case .cmdRecurNew.Enabled
              Case True
5560            Select Case blnRecurNew_Focus
                Case True
5570              .cmdRecurNew_raised_semifocus_dots_img.Visible = True
5580            Case False
5590              .cmdRecurNew_raised_img.Visible = True
5600            End Select
5610          Case False
5620            .cmdRecurNew_raised_img_dis.Visible = True
5630          End Select
5640          Select Case .cmdRefresh.Enabled
              Case True
5650            Select Case blnRefresh_Focus
                Case True
5660              .cmdRefresh_raised_semifocus_dots_img.Visible = True
5670            Case False
5680              .cmdRefresh_raised_img.Visible = True
5690            End Select
5700          Case False
5710            .cmdRefresh_raised_img_dis.Visible = True
5720          End Select
5730        End If

5740        lngTweak = 120&  ' ** Additional adjustment 09/24/2011.
5750        .Footer_vline05.Left = lngVLine_Left_New - lngTweak
5760        .Footer_vline06.Left = (lngVLine_Left_New + lngTpp) - lngTweak
5770        .cmdAdd_box.Left = (lngVLine_Left_New + (2& * lngTpp)) - lngTweak
5780        .cmdAdd.Left = (lngAdd_Left - lngVLine_Left_Diff) - lngTweak
5790        .cmdEdit.Left = (lngEdit_Left - lngVLine_Left_Diff) - lngTweak
5800        .cmdDelete.Left = (lngDelete_Left - lngVLine_Left_Diff) - lngTweak
5810        .Footer_vline07.Left = ((.cmdAdd_box.Left + .cmdAdd_box.Width) + lngTpp)
5820        .Footer_vline08.Left = (.Footer_vline07.Left + lngTpp)
5830        .cmdClose.Left = (lngClose_Left - lngVLine_Left_Diff) - lngTweak
5840        .Sizable_lbl1.Left = (((.frmJournal_Columns_Sub_box.Left + .frmJournal_Columns_Sub_box.Width) + _
              lngSizable_Offset) - .Sizable_lbl1.Width)
5850        .Sizable_lbl2.Left = .Sizable_lbl1.Left

5860      End Select
5870    End With

EXITP:
5880    Exit Sub

ERRH:
5890    Select Case ERR.Number
        Case Else
5900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5910    End Select
5920    Resume EXITP

End Sub

Public Sub JC_Btn_Resize_Opg(frm As Access.Form, blnShrinking As Boolean, lngSub_Bottom As Long)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Resize_Opg"

        Dim lngTmp01 As Long

6010    With frm
6020      Select Case blnShrinking
          Case True
6030        .opgEnterKey.Top = lngSub_Bottom + (5& * lngTpp)
6040        .opgEnterKey_optRight.Top = (.opgEnterKey.Top + lngOpt_Offset)
6050        .opgEnterKey_optRight_lbl.Top = (.opgEnterKey.Top + lngOptLbl_Offset)
6060        .opgEnterKey_optDown.Top = .opgEnterKey_optRight.Top
6070        .opgEnterKey_optDown_lbl.Top = .opgEnterKey_optRight_lbl.Top
            ' ** .opgEnterKey_lbl.Top  ' ** Has Tag.
6080        .opgEnterKey.Height = lngOpGrp_Height
            ' ** .opgEnterKey_vline01.Top  ' ** Has Tag.
            ' ** .opgEnterKey_vline02.Top  ' ** Has Tag.
            ' ** .opgEnterKey_vline03.Top  ' ** Has Tag.
            ' ** .opgEnterKey_vline04.Top  ' ** Has Tag.
6090        .opgFilter.Top = lngSub_Bottom + (5& * lngTpp)
6100        .opgFilter_optAll.Top = (.opgFilter.Top + lngOpt_Offset)
6110        .opgFilter_optAll_lbl.Top = (.opgFilter.Top + lngOptLbl_Offset)
6120        .opgFilter_optCommitted.Top = .opgFilter_optAll.Top
6130        .opgFilter_optCommitted_lbl.Top = .opgFilter_optAll_lbl.Top
6140        .opgFilter_optUncommitted.Top = .opgFilter_optAll.Top
6150        .opgFilter_optUncommitted_lbl.Top = .opgFilter_optAll_lbl.Top
            ' ** .opgFilter_lbl.Top  ' ** Has Tag.
6160        .opgFilter.Height = lngOpGrp_Height
            ' ** .opgFilter_vline01.Top  ' ** Has Tag.
            ' ** .opgFilter_vline02.Top  ' ** Has Tag.
            ' ** .opgFilter_vline03.Top  ' ** Has Tag.
            ' ** .opgFilter_vline04.Top  ' ** Has Tag.
6170      Case False
            ' ** (lngSubBottom - (.opgEnterKey.Top - 75&))                              ' ** Distance option group top is to be moved.
            ' ** ((.opgEnterKey.Top + .opgEnterKey.Height) - (.opgEnterKey.Top - 75&))  ' ** Distance between bottom of sub and bottom of opgrp.
            ' ** ((lngSubBottom + 75&) + .opgEnterKey.Height)                           ' ** New bottom of option group.
            ' ** (.opgEnterKey.Top + .opgEnterKey.Height)                               ' ** Old bottom of option group.
            ' ** (((lngSubBottom + 75&) + .opgEnterKey.Height) - (.opgEnterKey.Top + .opgEnterKey.Height))  ' ** Additional height needed.
6180        lngTmp01 = (((lngSub_Bottom + (5& * lngTpp)) + .opgEnterKey.Height) - (.opgEnterKey.Top + .opgEnterKey.Height))
6190        .opgEnterKey.Height = .opgEnterKey.Height + lngTmp01
6200        .opgEnterKey_optRight.Top = (.opgEnterKey_optRight.Top + (lngSub_Bottom - (.opgEnterKey.Top - (5& * lngTpp))))
6210        .opgEnterKey_optRight_lbl.Top = (.opgEnterKey_optRight_lbl.Top + (lngSub_Bottom - (.opgEnterKey.Top - (5& * lngTpp))))
6220        .opgEnterKey_optDown.Top = .opgEnterKey_optRight.Top
6230        .opgEnterKey_optDown_lbl.Top = .opgEnterKey_optRight_lbl.Top
            '.opgEnterKey_lbl.Top  ' ** Has Tag.
6240        .opgFilter.Height = .opgFilter.Height + lngTmp01
6250        .opgFilter_optAll.Top = (.opgFilter_optAll.Top + (lngSub_Bottom - (.opgFilter.Top - (5& * lngTpp))))
6260        .opgFilter_optAll_lbl.Top = (.opgFilter_optAll_lbl.Top + (lngSub_Bottom - (.opgFilter.Top - (5& * lngTpp))))
6270        .opgFilter_optCommitted.Top = .opgFilter_optAll.Top
6280        .opgFilter_optCommitted_lbl.Top = .opgFilter_optAll_lbl.Top
6290        .opgFilter_optUncommitted.Top = .opgFilter_optAll.Top
6300        .opgFilter_optUncommitted_lbl.Top = .opgFilter_optAll_lbl.Top
            ' ** .opgFilter_lbl.Top  ' ** Has Tag.
6310        .Detail.Height = (.Detail.Height + .opgEnterKey.Height)
6320        .opgEnterKey.Top = (.opgEnterKey.Top + (lngSub_Bottom - (.opgEnterKey.Top - (5& * lngTpp))))
6330        .opgEnterKey.Height = lngOpGrp_Height
            ' ** .opgEnterKey_vline01.Top  ' ** Has Tag.
            ' ** .opgEnterKey_vline02.Top  ' ** Has Tag.
            ' ** .opgEnterKey_vline03.Top  ' ** Has Tag.
            ' ** .opgEnterKey_vline04.Top  ' ** Has Tag.
6340        .opgFilter.Top = (.opgFilter.Top + (lngSub_Bottom - (.opgFilter.Top - (5& * lngTpp))))
6350        .opgFilter.Height = lngOpGrp_Height
            ' ** .opgFilter_vline01.Top  ' ** Has Tag.
            ' ** .opgFilter_vline02.Top  ' ** Has Tag.
            ' ** .opgFilter_vline03.Top  ' ** Has Tag.
            ' ** .opgFilter_vline04.Top  ' ** Has Tag.
6360        .Detail.Height = (.Detail.Height - lngTmp01)
6370      End Select
6380    End With

EXITP:
6390    Exit Sub

ERRH:
6400    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** May happen after Print Preview.
6410    Case Else
6420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6430    End Select
6440    Resume EXITP

End Sub

Public Sub JC_Btn_Resize_Set(lngWideBtn_WidthX As Long, lngVLine_LeftX As Long, lngAssets_LeftX As Long, lngLoc_LeftX As Long, lngRecur_LeftX As Long, lngRefresh_LeftX As Long, lngAdd_LeftX As Long, lngEdit_LeftX As Long, lngDelete_LeftX As Long, lngClose_LeftX As Long, lngOpGrp_HeightX As Long, lngOpt_OffsetX As Long, lngOptLbl_OffsetX As Long, lngSizable_OffsetX As Long)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Resize_Set"

        'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
6510    lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

6520    lngWideBtn_Width = lngWideBtn_WidthX
6530    lngVLine_Left = lngVLine_LeftX
6540    lngAssets_Left = lngAssets_LeftX
6550    lngLoc_Left = lngLoc_LeftX
6560    lngRecur_Left = lngRecur_LeftX
6570    lngRefresh_Left = lngRefresh_LeftX
6580    lngAdd_Left = lngAdd_LeftX
6590    lngEdit_Left = lngEdit_LeftX
6600    lngDelete_Left = lngDelete_LeftX
6610    lngClose_Left = lngClose_LeftX
6620    lngSizable_Offset = lngSizable_OffsetX
6630    lngOpGrp_Height = lngOpGrp_HeightX
6640    lngOpt_Offset = lngOpt_OffsetX
6650    lngOptLbl_Offset = lngOptLbl_OffsetX

EXITP:
6660    Exit Sub

ERRH:
6670    Select Case ERR.Number
        Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Sub

Public Function JC_Btn_Resize_Norm(frm As Access.Form) As Boolean

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Resize_Norm"

        Dim blnRetVal As Boolean

6810    blnRetVal = True

6820    With frm
6830      If .cmdAssetNew_raised_img.Visible = True Or .cmdAssetNew_raised_semifocus_dots_img.Visible = True Or _
              .cmdAssetNew_raised_focus_img.Visible = True Or .cmdAssetNew_raised_focus_dots_img.Visible = True Or _
              .cmdAssetNew_sunken_focus_dots_img.Visible = True Or .cmdAssetNew_raised_img_dis.Visible = True Then
            ' ** Arbitrary representative.
6840        blnRetVal = False
6850      End If
6860    End With

EXITP:
6870    JC_Btn_Resize_Norm = blnRetVal
6880    Exit Function

ERRH:
6890    blnRetVal = False
6900    Select Case ERR.Number
        Case Else
6910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6920    End Select
6930    Resume EXITP

End Function

Public Sub JC_Btn_Image(strProc As String, frmPar As Access.Form)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Image"

        Dim strCmd As String, strEvent As String
        Dim strCat As String
        Dim intLen As Integer
        Dim strTmp01 As String, strTmp02 As String, intTmp03 As Integer, blnTmp04 As Boolean, blnTmp05 As Boolean
        Dim intX As Integer

        Const IMG_R   As String = "_raised_img"
        Const IMG_RS  As String = "_raised_semifocus_img"       'SHOULDN'T BE ANY!
        Const IMG_RSD As String = "_raised_semifocus_dots_img"
        Const IMG_RF  As String = "_raised_focus_img"
        Const IMG_RFD As String = "_raised_focus_dots_img"
        'Const IMG_S   As String = "_sunken_img"                 'SHOULDN'T BE ANY!
        Const IMG_SF  As String = "_sunken_focus_img"           'SHOULDN'T BE ANY!
        Const IMG_SFD As String = "_sunken_focus_dots_img"
        Const IMG_D   As String = "_raised_img_dis"
        Const IMG_LBL As String = "_lbl"

7010    If strProc <> vbNullString Then

7020      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
7030        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
7040      End If

7050      strCmd = vbNullString: strEvent = vbNullString
7060      intLen = Len(strProc)
7070      For intX = intLen To 1 Step -1
7080        If Mid(strProc, intX, 1) = "_" Then
7090          strCmd = Left(strProc, (intX - 1))
7100          strEvent = Mid(strProc, (intX + 1))
7110          Exit For
7120        End If
7130      Next

7140      strCat = vbNullString
7150      strTmp01 = Left(strCmd, 8)
7160      Select Case strTmp01
          Case "cmdScrol"
7170        strCat = "cmdScroll"
7180      Case "cmdUncom"
7190        strCat = "cmdUncom"
7200      Case "cmdSpecP"
7210        strCat = "cmdSpecPurp"
7220      Case "cmdAsset", "cmdLocNe", "cmdRecur", "cmdRefre"
7230        strCat = "cmdBtnGroup"
7240      Case "cmdSwitc"
7250        strCat = "cmdSwitch"
7260      Case "cmdPrevi"
7270        strCat = "cmdPreviewReport"
7280      Case "cmdPrint"
7290        strCat = "cmdPrintReport"
7300      Case "cmdMemoR"
7310        strCat = "cmdMemoReveal"
7320      Case "Detail"
7330        strCat = "Detail"
7340      End Select

7350      With frmPar
7360        Select Case strEvent
            Case "GotFocus"
7370          Select Case strCat
              Case "cmdScroll"
                ' ** cmdScrollLeft, cmdScrollRight.
                ' ** Label gets focus dots for both Left and Right.
7380            .Controls(strCat & IMG_LBL & IMG_RFD).Visible = True
7390            .Controls(strCat & IMG_LBL & IMG_R).Visible = False
7400          Case "cmdUncom"
                ' ** cmdUncomComAll, cmdUncomDelAll, cmdUnCommitOne.
7410            Select Case strCmd
                Case "cmdUncomComAll"
7420              blnUncomCom_Focus = True
7430            Case "cmdUncomDelAll"
7440              blnUncomDel_Focus = True
7450            Case "cmdUnCommitOne"
7460              blnUnCommitOne_Focus = True
7470            End Select
7480            .Controls(strCmd & IMG_R).Visible = False    '_raised_img
7490            .Controls(strCmd & IMG_RSD).Visible = True   '_raised_semifocus_dots_img
7500            .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
7510            .Controls(strCmd & IMG_RFD).Visible = False  '_raised_focus_dots_img
7520            .Controls(strCmd & IMG_SFD).Visible = False  '_sunken_focus_dots_img
7530            .Controls(strCmd & IMG_D).Visible = False    '_raised_img_dis
7540          Case "cmdSpecPurp"
                ' ** cmdSpecPurp_Div_Map, cmdSpecPurp_Int_Map, cmdSpecPurp_Sold_PaidTotal, cmdSpecPurp_Purch_MapSplit,
                ' ** cmdSpecPurp_Misc_MapLTCG, cmdSpecPurp_Misc_MapLTCL, cmdSpecPurp_Misc_MapSTCGL.
7550            Select Case strCmd
                Case "cmdSpecPurp_Div_Map"
7560              blnMapDiv_Focus = True
7570            Case "cmdSpecPurp_Int_Map"
7580              blnMapInt_Focus = True
7590            Case "cmdSpecPurp_Sold_PaidTotal"
7600              blnPaid_Focus = True
7610            Case "cmdSpecPurp_Purch_MapSplit"
7620              blnMapSplit_Focus = True
7630            Case "cmdSpecPurp_Misc_MapLTCG"
7640              blnMapLTCG_Focus = True
7650            Case "cmdSpecPurp_Misc_MapLTCL"
7660              blnMapLTCL_Focus = True
7670            Case "cmdSpecPurp_Misc_MapSTCGL"
7680              blnMapSTCGL_Focus = True
7690            End Select
7700            .Controls(strCmd & IMG_R).Visible = False    '_raised_img
7710            .Controls(strCmd & IMG_RSD).Visible = True   '_raised_semifocus_dots_img
7720            .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
7730            .Controls(strCmd & IMG_RFD).Visible = False  '_raised_focus_dots_img
7740            .Controls(strCmd & IMG_SFD).Visible = False  '_sunken_focus_dots_img
7750            .Controls(strCmd & IMG_D).Visible = False    '_raised_img_dis
7760          Case "cmdBtnGroup"
                ' ** cmdAssetNew, cmdLocNew, cmdRecurNew, cmdRefresh
7770            Select Case strCmd
                Case "cmdAssetNew"
7780              blnAssetNew_Focus = True
7790            Case "cmdLocNew"
7800              blnLocNew_Focus = True
7810            Case "cmdRecurNew"
7820              blnRecurNew_Focus = True
7830            Case "cmdRefresh"
7840              blnRefresh_Focus = True
7850            End Select
7860            .Controls(strCmd & IMG_R).Visible = False
7870            .Controls(strCmd & IMG_RSD).Visible = True
7880            .Controls(strCmd & IMG_RF).Visible = False
7890            .Controls(strCmd & IMG_RFD).Visible = False
7900            .Controls(strCmd & IMG_SFD).Visible = False
7910            .Controls(strCmd & IMG_D).Visible = False
7920          Case "cmdPreviewReport", "cmdPrintReport", "cmdSwitch"
7930            Select Case strCmd
                Case "cmdPreviewReport"
7940              blnPreviewReport_Focus = True
7950            Case "cmdPrintReport"
7960              blnPrintReport_Focus = True
7970            Case "cmdSwitch"
7980              blnSwitch_Focus = True
7990            End Select
8000            .Controls(strCmd & IMG_R).Visible = False
8010            .Controls(strCmd & IMG_RSD).Visible = True
8020            .Controls(strCmd & IMG_RF).Visible = False
8030            .Controls(strCmd & IMG_RFD).Visible = False
8040            .Controls(strCmd & IMG_SFD).Visible = False
8050            .Controls(strCmd & IMG_D).Visible = False
8060          Case "cmdMemoReveal"
8070            blnMemoReveal_Focus = True
8080            Select Case .JrnlMemo_Memo.Visible
                Case True  ' ** So button points left, to close.
8090              .Controls(strCmd & "_L" & IMG_RS).Visible = True
8100              .Controls(strCmd & "_R" & IMG_RS).Visible = False
8110            Case False  ' ** So button points right, to open.
8120              .Controls(strCmd & "_R" & IMG_RS).Visible = True
8130              .Controls(strCmd & "_L" & IMG_RS).Visible = False
8140            End Select
8150            For intX = 1 To 2
8160              Select Case intX
                  Case 1
8170                strTmp01 = "_L"
8180              Case 2
8190                strTmp01 = "_R"
8200              End Select
8210              .Controls(strCmd & strTmp01 & IMG_R).Visible = False
8220              .Controls(strCmd & strTmp01 & IMG_RF).Visible = False
8230              .Controls(strCmd & strTmp01 & IMG_SF).Visible = False
8240              .Controls(strCmd & strTmp01 & IMG_D).Visible = False
8250            Next
                '.cmdMemoReveal_R_raised_imgX.Visible = False
                '.cmdMemoReveal_R_raised_semifocus_imgX.Visible = false
                '.cmdMemoReveal_R_raised_focus_imgX.Visible = False
                '.cmdMemoReveal_R_sunken_focus_imgX.Visible = False
                '.cmdMemoReveal_R_raised_imgX_dis.Visible = False
                '.cmdMemoReveal_L_raised_imgX.Visible = False
                '.cmdMemoReveal_L_raised_semifocus_imgX.Visible = False
                '.cmdMemoReveal_L_raised_focus_imgX.Visible = False
                '.cmdMemoReveal_L_sunken_focus_imgX.Visible = False
                '.cmdMemoReveal_L_raised_imgX_dis.Visible = False
8260          End Select  ' ** strCat.
8270        Case "LostFocus"
8280          Select Case strCat
              Case "cmdScroll"
                ' ** cmdScrollLeft, cmdScrollRight.
                ' ** Label gets focus dots for both Left and Right.
8290            .Controls(strCat & IMG_LBL & IMG_R).Visible = True
8300            .Controls(strCat & IMG_LBL & IMG_RFD).Visible = False
8310          Case "cmdUncom"
                ' ** cmdUncomComAll, cmdUncomDelAll.
8320            Select Case strCmd
                Case "cmdUncomComAll"
8330              blnUncomCom_Focus = False
8340            Case "cmdUncomDelAll"
8350              blnUncomDel_Focus = False
8360            Case "cmdUnCommitOne"
8370              blnUnCommitOne_Focus = False
8380            End Select
8390            .Controls(strCmd & IMG_R).Visible = True     '_raised_img
8400            .Controls(strCmd & IMG_RSD).Visible = False  '_raised_semifocus_dots_img
8410            .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
8420            .Controls(strCmd & IMG_RFD).Visible = False  '_raised_focus_dots_img
8430            .Controls(strCmd & IMG_SFD).Visible = False  '_sunken_focus_dots_img
8440            .Controls(strCmd & IMG_D).Visible = False    '_raised_img_dis
8450          Case "cmdSpecPurp"
                ' ** cmdSpecPurp_Div_Map, cmdSpecPurp_Int_Map, cmdSpecPurp_Sold_PaidTotal, cmdSpecPurp_Purch_MapSplit,
                ' ** cmdSpecPurp_Misc_MapLTCG, cmdSpecPurp_Misc_MapLTCL, cmdSpecPurp_Misc_MapSTCGL.
8460            Select Case strCmd
                Case "cmdSpecPurp_Div_Map"
8470              blnMapDiv_Focus = False
8480            Case "cmdSpecPurp_Int_Map"
8490              blnMapInt_Focus = False
8500            Case "cmdSpecPurp_Sold_PaidTotal"
8510              blnPaid_Focus = False
8520            Case "cmdSpecPurp_Purch_MapSplit"
8530              blnMapSplit_Focus = False
8540            Case "cmdSpecPurp_Misc_MapLTCG"
8550              blnMapLTCG_Focus = False
8560            Case "cmdSpecPurp_Misc_MapLTCL"
8570              blnMapLTCL_Focus = False
8580            Case "cmdSpecPurp_Misc_MapSTCGL"
8590              blnMapSTCGL_Focus = False
8600            End Select
8610            .Controls(strCmd & IMG_R).Visible = True
8620            .Controls(strCmd & IMG_RSD).Visible = False  '_raised_semifocus_dots_img
8630            .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
8640            .Controls(strCmd & IMG_RFD).Visible = False
8650            .Controls(strCmd & IMG_SFD).Visible = False
8660            .Controls(strCmd & IMG_D).Visible = False
8670          Case "cmdBtnGroup"
                ' ** cmdAssetNew, cmdLocNew, cmdRecurNew, cmdRefresh
8680            .Controls(strCmd & IMG_R).Visible = True
8690            .Controls(strCmd & IMG_RSD).Visible = False
8700            .Controls(strCmd & IMG_RF).Visible = False
8710            .Controls(strCmd & IMG_RFD).Visible = False
8720            .Controls(strCmd & IMG_SFD).Visible = False
8730            .Controls(strCmd & IMG_D).Visible = False
8740            Select Case strCmd
                Case "cmdAssetNew"
8750              blnAssetNew_Focus = False
8760            Case "cmdLocNew"
8770              blnLocNew_Focus = False
8780            Case "cmdRecurNew"
8790              blnRecurNew_Focus = False
8800            Case "cmdRefresh"
8810              blnRefresh_Focus = False
8820            End Select
8830          Case "cmdPreviewReport", "cmdPrintReport", "cmdSwitch"
8840            .Controls(strCmd & IMG_R).Visible = True
8850            .Controls(strCmd & IMG_RSD).Visible = False
8860            .Controls(strCmd & IMG_RF).Visible = False
8870            .Controls(strCmd & IMG_RFD).Visible = False
8880            .Controls(strCmd & IMG_SFD).Visible = False
8890            .Controls(strCmd & IMG_D).Visible = False
8900            Select Case strCmd
                Case "cmdPreviewReport"
8910              blnPreviewReport_Focus = False
8920            Case "cmdPrintReport"
8930              blnPrintReport_Focus = False
8940            Case "cmdSwitch"
8950              blnSwitch_Focus = False
8960            End Select
8970          Case "cmdMemoReveal"
8980            Select Case .JrnlMemo_Memo.Visible
                Case True  ' ** So button points left, to close.
8990              .Controls(strCmd & "_L" & IMG_R).Visible = True
9000              .Controls(strCmd & "_R" & IMG_R).Visible = False
9010            Case False  ' ** So button points right, to open.
9020              .Controls(strCmd & "_R" & IMG_R).Visible = True
9030              .Controls(strCmd & "_L" & IMG_R).Visible = False
9040            End Select
9050            For intX = 1 To 2
9060              Select Case intX
                  Case 1
9070                strTmp01 = "_L"
9080              Case 2
9090                strTmp01 = "_R"
9100              End Select
9110              .Controls(strCmd & strTmp01 & IMG_RS).Visible = False
9120              .Controls(strCmd & strTmp01 & IMG_RF).Visible = False
9130              .Controls(strCmd & strTmp01 & IMG_SF).Visible = False
9140              .Controls(strCmd & strTmp01 & IMG_D).Visible = False
9150            Next
9160            blnMemoReveal_Focus = False
9170          End Select  ' ** strCat.
9180        Case "MouseDown"
9190          Select Case strCat
              Case "cmdScroll"
                ' ** cmdScrollLeft, cmdScrollRight.
9200            .Controls(strCmd & IMG_SFD).Visible = True
9210            .Controls(strCmd & IMG_R).Visible = False
9220          Case "cmdUncom"
                ' ** cmdUncomComAll, cmdUncomDelAll.
9230            Select Case strCmd
                Case "cmdUncomComAll"
9240              blnUncomCom_MouseDown = True
9250            Case "cmdUncomDelAll"
9260              blnUncomDel_MouseDown = True
9270            Case "cmdUnCommitOne"
9280              blnUnCommitOne_MouseDown = True
9290            End Select
9300            .Controls(strCmd & IMG_R).Visible = False    '_raised_img
9310            .Controls(strCmd & IMG_RSD).Visible = False  '_raised_semifocus_dots_img
9320            .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
9330            .Controls(strCmd & IMG_RFD).Visible = False  '_raised_focus_dots_img
9340            .Controls(strCmd & IMG_SFD).Visible = True   '_sunken_focus_dots_img
9350            .Controls(strCmd & IMG_D).Visible = False    '_raised_img_dis
9360          Case "cmdSpecPurp"
                ' ** cmdSpecPurp_Div_Map, cmdSpecPurp_Int_Map, cmdSpecPurp_Sold_PaidTotal, cmdSpecPurp_Purch_MapSplit,
                ' ** cmdSpecPurp_Misc_MapLTCG, cmdSpecPurp_Misc_MapLTCL, cmdSpecPurp_Misc_MapSTCGL.
9370            Select Case strCmd
                Case "cmdSpecPurp_Div_Map"
9380              blnMapDiv_MouseDown = True
9390            Case "cmdSpecPurp_Int_Map"
9400              blnMapInt_MouseDown = True
9410            Case "cmdSpecPurp_Sold_PaidTotal"
9420              blnPaid_MouseDown = True
9430            Case "cmdSpecPurp_Purch_MapSplit"
9440              blnMapSplit_MouseDown = True
9450            Case "cmdSpecPurp_Misc_MapLTCG"
9460              blnMapLTCG_MouseDown = True
9470            Case "cmdSpecPurp_Misc_MapLTCL"
9480              blnMapLTCL_MouseDown = True
9490            Case "cmdSpecPurp_Misc_MapSTCGL"
9500              blnMapSTCGL_MouseDown = True
9510            End Select
9520            .Controls(strCmd & IMG_R).Visible = False
9530            .Controls(strCmd & IMG_RSD).Visible = False  '_raised_semifocus_dots_img
9540            .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
9550            .Controls(strCmd & IMG_RFD).Visible = False
9560            .Controls(strCmd & IMG_SFD).Visible = True
9570            .Controls(strCmd & IMG_D).Visible = False
9580          Case "cmdBtnGroup"
                ' ** cmdAssetNew, cmdLocNew, cmdRecurNew, cmdRefresh
9590            Select Case strCmd
                Case "cmdAssetNew"
9600              blnAssetNew_MouseDown = True
9610            Case "cmdLocNew"
9620              blnLocNew_MouseDown = True
9630            Case "cmdRecurNew"
9640              blnRecurNew_MouseDown = True
9650            Case "cmdRefresh"
9660              blnRefresh_MouseDown = True
9670            End Select
9680            .Controls(strCmd & IMG_R).Visible = False
9690            .Controls(strCmd & IMG_RSD).Visible = False
9700            .Controls(strCmd & IMG_RF).Visible = False
9710            .Controls(strCmd & IMG_RFD).Visible = False
9720            .Controls(strCmd & IMG_SFD).Visible = True
9730            .Controls(strCmd & IMG_D).Visible = False
9740          Case "cmdPreviewReport", "cmdPrintReport", "cmdSwitch"
9750            Select Case strCmd
                Case "cmdPreviewReport"
9760              blnPreviewReport_MouseDown = True
9770            Case "cmdPrintReport"
9780              blnPrintReport_MouseDown = True
9790            Case "cmdSwitch"
9800              blnSwitch_MouseDown = True
9810            End Select
9820            .Controls(strCmd & IMG_R).Visible = False
9830            .Controls(strCmd & IMG_RSD).Visible = False
9840            .Controls(strCmd & IMG_RF).Visible = False
9850            .Controls(strCmd & IMG_RFD).Visible = False
9860            .Controls(strCmd & IMG_SFD).Visible = True
9870            .Controls(strCmd & IMG_D).Visible = False
9880          Case "cmdMemoReveal"
9890            blnMemoReveal_MouseDown = True
9900            Select Case .JrnlMemo_Memo.Visible
                Case True  ' ** So button points left, to close.
9910              .Controls(strCmd & "_L" & IMG_SF).Visible = True
9920              .Controls(strCmd & "_R" & IMG_SF).Visible = False
9930            Case False  ' ** So button points right, to open.
9940              .Controls(strCmd & "_R" & IMG_SF).Visible = True
9950              .Controls(strCmd & "_L" & IMG_SF).Visible = False
9960            End Select
9970            For intX = 1 To 2
9980              Select Case intX
                  Case 1
9990                strTmp01 = "_L"
10000             Case 2
10010               strTmp01 = "_R"
10020             End Select
10030             .Controls(strCmd & strTmp01 & IMG_R).Visible = False
10040             .Controls(strCmd & strTmp01 & IMG_RS).Visible = False
10050             .Controls(strCmd & strTmp01 & IMG_RF).Visible = False
10060             .Controls(strCmd & strTmp01 & IMG_D).Visible = False
10070           Next
10080         End Select  ' ** strCat.
10090       Case "MouseUp"
10100         Select Case strCat
              Case "cmdScroll"
                ' ** cmdScrollLeft, cmdScrollRight.
10110           .Controls(strCmd & IMG_R).Visible = True
10120           .Controls(strCmd & IMG_SFD).Visible = False
10130         Case "cmdUncom"
                ' ** cmdUncomComAll, cmdUncomDelAll.
10140           Select Case strCmd
                Case "cmdUncomComAll"
10150             blnUncomCom_MouseDown = False
10160           Case "cmdUncomDelAll"
10170             blnUncomDel_MouseDown = False
10180           Case "cmdUnCommitOne"
10190             blnUnCommitOne_MouseDown = False
10200           End Select
10210           .Controls(strCmd & IMG_R).Visible = False    '_raised_img
10220           .Controls(strCmd & IMG_RSD).Visible = False  '_raised_semifocus_dots_img
10230           .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
10240           .Controls(strCmd & IMG_RFD).Visible = True   '_raised_focus_dots_img
10250           .Controls(strCmd & IMG_SFD).Visible = False  '_sunken_focus_dots_img
10260           .Controls(strCmd & IMG_D).Visible = False    '_raised_img_dis
10270         Case "cmdSpecPurp"
                ' ** cmdSpecPurp_Div_Map, cmdSpecPurp_Int_Map, cmdSpecPurp_Sold_PaidTotal, cmdSpecPurp_Purch_MapSplit,
                ' ** cmdSpecPurp_Misc_MapLTCG, cmdSpecPurp_Misc_MapLTCL, cmdSpecPurp_Misc_MapSTCGL.
10280           Select Case strCmd
                Case "cmdSpecPurp_Div_Map"
10290             blnMapDiv_MouseDown = False
10300           Case "cmdSpecPurp_Int_Map"
10310             blnMapInt_MouseDown = False
10320           Case "cmdSpecPurp_Sold_PaidTotal"
10330             blnPaid_MouseDown = False
10340           Case "cmdSpecPurp_Purch_MapSplit"
10350             blnMapSplit_MouseDown = False
10360           Case "cmdSpecPurp_Misc_MapLTCG"
10370             blnMapLTCG_MouseDown = False
10380           Case "cmdSpecPurp_Misc_MapLTCL"
10390             blnMapLTCL_MouseDown = False
10400           Case "cmdSpecPurp_Misc_MapSTCGL"
10410             blnMapSTCGL_MouseDown = False
10420           End Select
10430           .Controls(strCmd & IMG_R).Visible = False
10440           .Controls(strCmd & IMG_RSD).Visible = False  '_raised_semifocus_dots_img
10450           .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
10460           .Controls(strCmd & IMG_RFD).Visible = True
10470           .Controls(strCmd & IMG_SFD).Visible = False
10480           .Controls(strCmd & IMG_D).Visible = False
10490         Case "cmdBtnGroup"
                ' ** cmdAssetNew, cmdLocNew, cmdRecurNew, cmdRefresh
10500           .Controls(strCmd & IMG_R).Visible = False
10510           .Controls(strCmd & IMG_RSD).Visible = False
10520           .Controls(strCmd & IMG_RF).Visible = False
10530           .Controls(strCmd & IMG_RFD).Visible = True
10540           .Controls(strCmd & IMG_SFD).Visible = False
10550           .Controls(strCmd & IMG_D).Visible = False
10560           Select Case strCmd
                Case "cmdAssetNew"
10570             blnAssetNew_MouseDown = False
10580           Case "cmdLocNew"
10590             blnLocNew_MouseDown = False
10600           Case "cmdRecurNew"
10610             blnRecurNew_MouseDown = False
10620           Case "cmdRefresh"
10630             blnRefresh_MouseDown = False
10640           End Select
10650         Case "cmdPreviewReport", "cmdPrintReport", "cmdSwitch"
10660           .Controls(strCmd & IMG_R).Visible = False
10670           .Controls(strCmd & IMG_RSD).Visible = False
10680           .Controls(strCmd & IMG_RF).Visible = False
10690           .Controls(strCmd & IMG_RFD).Visible = True
10700           .Controls(strCmd & IMG_SFD).Visible = False
10710           .Controls(strCmd & IMG_D).Visible = False
10720           Select Case strCmd
                Case "cmdPreviewReport"
10730             blnPreviewReport_MouseDown = False
10740           Case "cmdPrintReport"
10750             blnPrintReport_MouseDown = False
10760           Case "cmdSwitch"
10770             blnSwitch_MouseDown = False
10780           End Select
10790         Case "cmdMemoReveal"
10800           Select Case .JrnlMemo_Memo.Visible
                Case True  ' ** So button points left, to close.
10810             .Controls(strCmd & "_L" & IMG_RF).Visible = True
10820             .Controls(strCmd & "_R" & IMG_RF).Visible = False
10830           Case False  ' ** So button points right, to open.
10840             .Controls(strCmd & "_R" & IMG_RF).Visible = True
10850             .Controls(strCmd & "_L" & IMG_RF).Visible = False
10860           End Select
10870           For intX = 1 To 2
10880             Select Case intX
                  Case 1
10890               strTmp01 = "_L"
10900             Case 2
10910               strTmp01 = "_R"
10920             End Select
10930             .Controls(strCmd & strTmp01 & IMG_R).Visible = False
10940             .Controls(strCmd & strTmp01 & IMG_RS).Visible = False
10950             .Controls(strCmd & strTmp01 & IMG_SF).Visible = False
10960             .Controls(strCmd & strTmp01 & IMG_D).Visible = False
10970           Next
10980           blnMemoReveal_MouseDown = False
10990         End Select  ' ** strCat.
11000       Case "MouseMove"
11010         Select Case strCat
              Case "cmdUncom"
11020           blnTmp04 = False: blnTmp05 = False
11030           Select Case strCmd
                Case "cmdUncomComAll"
11040             If blnUncomCom_MouseDown = False Then
11050               blnTmp04 = True
11060             End If
11070           Case "cmdUncomDelAll"
11080             If blnUncomDel_MouseDown = False Then
11090               blnTmp04 = True
11100             End If
11110           Case "cmdUnCommitOne"
11120             If blnUnCommitOne_MouseDown = False Then
11130               blnTmp04 = True
11140             End If
11150           End Select
11160           If blnTmp04 = True Then
11170             Select Case strCmd
                  Case "cmdUncomComAll"
11180               blnTmp04 = blnUncomCom_Focus
11190             Case "cmdUncomDelAll"
11200               blnTmp04 = blnUncomDel_Focus
11210             Case "cmdUnCommitOne"
11220               blnTmp04 = blnUnCommitOne_Focus
11230             End Select
11240             Select Case blnTmp04
                  Case True
11250               .Controls(strCmd & IMG_RFD).Visible = True   '_raised_focus_dots_img
11260               .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
11270             Case False
11280               .Controls(strCmd & IMG_RF).Visible = True    '_raised_focus_img
11290               .Controls(strCmd & IMG_RFD).Visible = False  '_raised_focus_dots_img
11300             End Select
11310             .Controls(strCmd & IMG_R).Visible = False      '_raised_img
11320             .Controls(strCmd & IMG_RSD).Visible = False    '_raised_semifocus_dots_img
11330             .Controls(strCmd & IMG_SFD).Visible = False    '_sunken_focus_dots_img
11340             .Controls(strCmd & IMG_D).Visible = False      '_raised_img_dis
11350           End If
11360           Select Case strCmd
                Case "cmdUncomComAll"
11370             strTmp01 = "cmdUncomDelAll"
11380             strTmp02 = "cmdUnCommitOne"
11390             blnTmp04 = blnUncomDel_Focus
11400             blnTmp05 = blnUnCommitOne_Focus
11410           Case "cmdUncomDelAll"
11420             strTmp01 = "cmdUncomComAll"
11430             strTmp02 = "cmdUnCommitOne"
11440             blnTmp04 = blnUncomCom_Focus
11450             blnTmp05 = blnUnCommitOne_Focus
11460           Case "cmdUnCommitOne"
11470             strTmp01 = "cmdUncomComAll"
11480             strTmp02 = "cmdUncomDelAll"
11490             blnTmp04 = blnUncomCom_Focus
11500             blnTmp05 = blnUncomDel_Focus
11510           End Select
11520           If .Controls(strTmp01 & IMG_RFD).Visible = True Or .Controls(strTmp01 & IMG_RF).Visible = True Then
11530             Select Case blnTmp04
                  Case True
11540               .Controls(strTmp01 & IMG_RSD).Visible = True
11550               .Controls(strTmp01 & IMG_R).Visible = False
11560             Case False
11570               .Controls(strTmp01 & IMG_R).Visible = True
11580               .Controls(strTmp01 & IMG_RSD).Visible = False
11590             End Select
11600             .Controls(strTmp01 & IMG_RF).Visible = False
11610             .Controls(strTmp01 & IMG_RFD).Visible = False
11620             .Controls(strTmp01 & IMG_SFD).Visible = False
11630             .Controls(strTmp01 & IMG_D).Visible = False
11640           End If
11650           If .Controls(strTmp02 & IMG_RFD).Visible = True Or .Controls(strTmp02 & IMG_RF).Visible = True Then
11660             Select Case blnTmp05
                  Case True
11670               .Controls(strTmp02 & IMG_RSD).Visible = True
11680               .Controls(strTmp02 & IMG_R).Visible = False
11690             Case False
11700               .Controls(strTmp02 & IMG_R).Visible = True
11710               .Controls(strTmp02 & IMG_RSD).Visible = False
11720             End Select
11730             .Controls(strTmp02 & IMG_RF).Visible = False
11740             .Controls(strTmp02 & IMG_RFD).Visible = False
11750             .Controls(strTmp02 & IMG_SFD).Visible = False
11760             .Controls(strTmp02 & IMG_D).Visible = False
11770           End If
11780         Case "cmdSpecPurp"
                ' ** cmdSpecPurp_Div_Map, cmdSpecPurp_Int_Map, cmdSpecPurp_Sold_PaidTotal, cmdSpecPurp_Purch_MapSplit,
                ' ** cmdSpecPurp_Misc_MapLTCG, cmdSpecPurp_Misc_MapLTCL, cmdSpecPurp_Misc_MapSTCGL.
11790           blnTmp04 = False
11800           Select Case strCmd
                Case "cmdSpecPurp_Div_Map"
11810             If blnMapDiv_MouseDown = False Then
11820               blnTmp04 = True
11830             End If
11840           Case "cmdSpecPurp_Int_Map"
11850             If blnMapInt_MouseDown = False Then
11860               blnTmp04 = True
11870             End If
11880           Case "cmdSpecPurp_Sold_PaidTotal"
11890             If blnPaid_MouseDown = False Then
11900               blnTmp04 = True
11910             End If
11920           Case "cmdSpecPurp_Purch_MapSplit"
11930             If blnMapSplit_MouseDown = False Then
11940               blnTmp04 = True
11950             End If
11960           Case "cmdSpecPurp_Misc_MapLTCG"
11970             If blnMapLTCG_MouseDown = False Then
11980               blnTmp04 = True
11990             End If
12000           Case "cmdSpecPurp_Misc_MapLTCL"
12010             If blnMapLTCL_MouseDown = False Then
12020               blnTmp04 = True
12030             End If
12040           Case "cmdSpecPurp_Misc_MapSTCGL"
12050             If blnMapSTCGL_MouseDown = False Then
12060               blnTmp04 = True
12070             End If
12080           End Select
12090           If blnTmp04 = True Then
12100             Select Case strCmd
                  Case "cmdSpecPurp_Div_Map"
12110               blnTmp04 = blnMapDiv_Focus
12120             Case "cmdSpecPurp_Int_Map"
12130               blnTmp04 = blnMapInt_Focus
12140             Case "cmdSpecPurp_Sold_PaidTotal"
12150               blnTmp04 = blnPaid_Focus
12160             Case "cmdSpecPurp_Purch_MapSplit"
12170               blnTmp04 = blnMapSplit_Focus
12180             Case "cmdSpecPurp_Misc_MapLTCG"
12190               blnTmp04 = blnMapLTCG_Focus
12200             Case "cmdSpecPurp_Misc_MapLTCL"
12210               blnTmp04 = blnMapLTCL_Focus
12220             Case "cmdSpecPurp_Misc_MapSTCGL"
12230               blnTmp04 = blnMapSTCGL_Focus
12240             End Select
12250             Select Case blnTmp04
                  Case True
12260               .Controls(strCmd & IMG_RFD).Visible = True   '_raised_focus_dots_img
12270               .Controls(strCmd & IMG_RF).Visible = False   '_raised_focus_img
12280             Case False
12290               .Controls(strCmd & IMG_RF).Visible = True    '_raised_focus_img
12300               .Controls(strCmd & IMG_RFD).Visible = False  '_raised_focus_dots_img
12310             End Select
12320             .Controls(strCmd & IMG_R).Visible = False      '_raised_img
12330             .Controls(strCmd & IMG_RSD).Visible = False    '_raised_semifocus_dots_img
12340             .Controls(strCmd & IMG_SFD).Visible = False    '_sunken_focus_dots_img
12350             .Controls(strCmd & IMG_D).Visible = False      '_raised_img_dis
12360           End If
12370           Select Case gblnSpecialCapGainLoss
                Case True
12380             intTmp03 = 6
12390           Case False
12400             intTmp03 = 4
12410           End Select
12420           For intX = 1 To intTmp03  ' ** For the 4/6 other buttons in the group.
12430             strTmp02 = vbNullString
12440             blnTmp04 = False
12450             Select Case strCmd
                  Case "cmdSpecPurp_Div_Map"
12460               Select Case intX
                    Case 1
12470                 strTmp02 = "cmdSpecPurp_Int_Map"
12480                 blnTmp04 = blnMapInt_Focus
12490               Case 2
12500                 strTmp02 = "cmdSpecPurp_Sold_PaidTotal"
12510                 blnTmp04 = blnPaid_Focus
12520               Case 3
12530                 strTmp02 = "cmdSpecPurp_Purch_MapSplit"
12540                 blnTmp04 = blnMapSplit_Focus
12550               Case 4
12560                 strTmp02 = "cmdSpecPurp_Misc_MapLTCG"
12570                 blnTmp04 = blnMapLTCG_Focus
12580               Case 5
12590                 strTmp02 = "cmdSpecPurp_Misc_MapLTCL"
12600                 blnTmp04 = blnMapLTCL_Focus
12610               Case 6
12620                 strTmp02 = "cmdSpecPurp_Misc_MapSTCGL"
12630                 blnTmp04 = blnMapSTCGL_Focus
12640               End Select
12650             Case "cmdSpecPurp_Int_Map"
12660               Select Case intX
                    Case 1
12670                 strTmp02 = "cmdSpecPurp_Div_Map"
12680                 blnTmp04 = blnMapDiv_Focus
12690               Case 2
12700                 strTmp02 = "cmdSpecPurp_Sold_PaidTotal"
12710                 blnTmp04 = blnPaid_Focus
12720               Case 3
12730                 strTmp02 = "cmdSpecPurp_Purch_MapSplit"
12740                 blnTmp04 = blnMapSplit_Focus
12750               Case 4
12760                 strTmp02 = "cmdSpecPurp_Misc_MapLTCG"
12770                 blnTmp04 = blnMapLTCG_Focus
12780               Case 5
12790                 strTmp02 = "cmdSpecPurp_Misc_MapLTCL"
12800                 blnTmp04 = blnMapLTCL_Focus
12810               Case 6
12820                 strTmp02 = "cmdSpecPurp_Misc_MapSTCGL"
12830                 blnTmp04 = blnMapSTCGL_Focus
12840               End Select
12850             Case "cmdSpecPurp_Sold_PaidTotal"
12860               Select Case intX
                    Case 1
12870                 strTmp02 = "cmdSpecPurp_Div_Map"
12880                 blnTmp04 = blnMapDiv_Focus
12890               Case 2
12900                 strTmp02 = "cmdSpecPurp_Int_Map"
12910                 blnTmp04 = blnMapInt_Focus
12920               Case 3
12930                 strTmp02 = "cmdSpecPurp_Purch_MapSplit"
12940                 blnTmp04 = blnMapSplit_Focus
12950               Case 4
12960                 strTmp02 = "cmdSpecPurp_Misc_MapLTCG"
12970                 blnTmp04 = blnMapLTCG_Focus
12980               Case 5
12990                 strTmp02 = "cmdSpecPurp_Misc_MapLTCL"
13000                 blnTmp04 = blnMapLTCL_Focus
13010               Case 6
13020                 strTmp02 = "cmdSpecPurp_Misc_MapSTCGL"
13030                 blnTmp04 = blnMapSTCGL_Focus
13040               End Select
13050             Case "cmdSpecPurp_Purch_MapSplit"
13060               Select Case intX
                    Case 1
13070                 strTmp02 = "cmdSpecPurp_Div_Map"
13080                 blnTmp04 = blnMapDiv_Focus
13090               Case 2
13100                 strTmp02 = "cmdSpecPurp_Int_Map"
13110                 blnTmp04 = blnMapInt_Focus
13120               Case 3
13130                 strTmp02 = "cmdSpecPurp_Sold_PaidTotal"
13140                 blnTmp04 = blnPaid_Focus
13150               Case 4
13160                 strTmp02 = "cmdSpecPurp_Misc_MapLTCG"
13170                 blnTmp04 = blnMapLTCG_Focus
13180               Case 5
13190                 strTmp02 = "cmdSpecPurp_Misc_MapLTCL"
13200                 blnTmp04 = blnMapLTCL_Focus
13210               Case 6
13220                 strTmp02 = "cmdSpecPurp_Misc_MapSTCGL"
13230                 blnTmp04 = blnMapSTCGL_Focus
13240               End Select
13250             Case "cmdSpecPurp_Misc_MapLTCG"
13260               Select Case intX
                    Case 1
13270                 strTmp02 = "cmdSpecPurp_Div_Map"
13280                 blnTmp04 = blnMapDiv_Focus
13290               Case 2
13300                 strTmp02 = "cmdSpecPurp_Int_Map"
13310                 blnTmp04 = blnMapInt_Focus
13320               Case 3
13330                 strTmp02 = "cmdSpecPurp_Sold_PaidTotal"
13340                 blnTmp04 = blnPaid_Focus
13350               Case 4
13360                 strTmp02 = "cmdSpecPurp_Purch_MapSplit"
13370                 blnTmp04 = blnMapSplit_Focus
13380               Case 5
13390                 strTmp02 = "cmdSpecPurp_Misc_MapLTCL"
13400                 blnTmp04 = blnMapLTCL_Focus
13410               Case 6
13420                 strTmp02 = "cmdSpecPurp_Misc_MapSTCGL"
13430                 blnTmp04 = blnMapSTCGL_Focus
13440               End Select
13450             Case "cmdSpecPurp_Misc_MapLTCL"
13460               Select Case intX
                    Case 1
13470                 strTmp02 = "cmdSpecPurp_Div_Map"
13480                 blnTmp04 = blnMapDiv_Focus
13490               Case 2
13500                 strTmp02 = "cmdSpecPurp_Int_Map"
13510                 blnTmp04 = blnMapInt_Focus
13520               Case 3
13530                 strTmp02 = "cmdSpecPurp_Sold_PaidTotal"
13540                 blnTmp04 = blnPaid_Focus
13550               Case 4
13560                 strTmp02 = "cmdSpecPurp_Purch_MapSplit"
13570                 blnTmp04 = blnMapSplit_Focus
13580               Case 5
13590                 strTmp02 = "cmdSpecPurp_Misc_MapLTCG"
13600                 blnTmp04 = blnMapLTCG_Focus
13610               Case 6
13620                 strTmp02 = "cmdSpecPurp_Misc_MapSTCGL"
13630                 blnTmp04 = blnMapSTCGL_Focus
13640               End Select
13650             Case "cmdSpecPurp_Misc_MapSTCGL"
13660               Select Case intX
                    Case 1
13670                 strTmp02 = "cmdSpecPurp_Div_Map"
13680                 blnTmp04 = blnMapDiv_Focus
13690               Case 2
13700                 strTmp02 = "cmdSpecPurp_Int_Map"
13710                 blnTmp04 = blnMapInt_Focus
13720               Case 3
13730                 strTmp02 = "cmdSpecPurp_Sold_PaidTotal"
13740                 blnTmp04 = blnPaid_Focus
13750               Case 4
13760                 strTmp02 = "cmdSpecPurp_Purch_MapSplit"
13770                 blnTmp04 = blnMapSplit_Focus
13780               Case 5
13790                 strTmp02 = "cmdSpecPurp_Misc_MapLTCG"
13800                 blnTmp04 = blnMapLTCG_Focus
13810               Case 6
13820                 strTmp02 = "cmdSpecPurp_Misc_MapLTCL"
13830                 blnTmp04 = blnMapLTCL_Focus
13840               End Select
13850             End Select
13860             If .Controls(strTmp02 & IMG_RFD).Visible = True Or .Controls(strTmp02 & IMG_RF).Visible = True Then
13870               Select Case blnTmp04
                    Case True
13880                 .Controls(strTmp02 & IMG_RSD).Visible = True
13890                 .Controls(strTmp02 & IMG_R).Visible = False
13900               Case False
13910                 .Controls(strTmp02 & IMG_R).Visible = True
13920                 .Controls(strTmp02 & IMG_RSD).Visible = False
13930               End Select
13940               .Controls(strTmp02 & IMG_RF).Visible = False
13950               .Controls(strTmp02 & IMG_RFD).Visible = False
13960               .Controls(strTmp02 & IMG_SFD).Visible = False
13970               .Controls(strTmp02 & IMG_D).Visible = False
13980             End If
13990           Next
14000         Case "cmdBtnGroup"
14010           blnTmp04 = False
14020           Select Case strCmd
                Case "cmdAssetNew"
14030             If blnAssetNew_MouseDown = False Then
14040               blnTmp04 = True
14050             End If
14060           Case "cmdLocNew"
14070             If blnLocNew_MouseDown = False Then
14080               blnTmp04 = True
14090             End If
14100           Case "cmdRecurNew"
14110             If blnRecurNew_MouseDown = False Then
14120               blnTmp04 = True
14130             End If
14140           Case "cmdRefresh"
14150             If blnRefresh_MouseDown = False Then
14160               blnTmp04 = True
14170             End If
14180           End Select
14190           If blnTmp04 = True Then
14200             Select Case strCmd
                  Case "cmdAssetNew"
14210               blnTmp04 = blnAssetNew_Focus
14220             Case "cmdLocNew"
14230               blnTmp04 = blnLocNew_Focus
14240             Case "cmdRecurNew"
14250               blnTmp04 = blnRecurNew_Focus
14260             Case "cmdRefresh"
14270               blnTmp04 = blnRefresh_Focus
14280             End Select
14290             Select Case blnTmp04
                  Case True
14300               .Controls(strCmd & IMG_RFD).Visible = True
14310               .Controls(strCmd & IMG_RF).Visible = False
14320             Case False
14330               .Controls(strCmd & IMG_RF).Visible = True
14340               .Controls(strCmd & IMG_RFD).Visible = False
14350             End Select
14360             .Controls(strCmd & IMG_R).Visible = False
14370             .Controls(strCmd & IMG_RSD).Visible = False
14380             .Controls(strCmd & IMG_SFD).Visible = False
14390             .Controls(strCmd & IMG_D).Visible = False
14400           End If
14410         Case "cmdPreviewReport", "cmdPrintReport", "cmdSwitch"
14420           Select Case strCmd
                Case "cmdPreviewReport"
14430             blnTmp04 = blnPreviewReport_MouseDown
14440           Case "cmdPrintReport"
14450             blnTmp04 = blnPrintReport_MouseDown
14460           Case "cmdSwitch"
14470             blnTmp04 = blnSwitch_MouseDown
14480           End Select
14490           If blnTmp04 = False Then
14500             Select Case strCmd
                  Case "cmdPreviewReport"
14510               blnTmp04 = blnPreviewReport_Focus
14520             Case "cmdPrintReport"
14530               blnTmp04 = blnPrintReport_Focus
14540             Case "cmdSwitch"
14550               blnTmp04 = blnSwitch_Focus
14560             End Select
14570             Select Case blnTmp04
                  Case True
14580               .Controls(strCmd & IMG_RFD).Visible = True
14590               .Controls(strCmd & IMG_RF).Visible = False
14600             Case False
14610               .Controls(strCmd & IMG_RF).Visible = True
14620               .Controls(strCmd & IMG_RFD).Visible = False
14630             End Select
14640             .Controls(strCmd & IMG_R).Visible = False
14650             .Controls(strCmd & IMG_RSD).Visible = False
14660             .Controls(strCmd & IMG_SFD).Visible = False
14670             .Controls(strCmd & IMG_D).Visible = False
14680           End If
14690         Case "cmdMemoReveal"
14700           If blnMemoReveal_MouseDown = False Then
14710             Select Case .JrnlMemo_Memo.Visible
                  Case True  ' ** So button points left, to close.
14720               Select Case blnMemoReveal_Focus
                    Case True
14730                 .Controls(strCmd & "_L" & IMG_RF).Visible = True
14740                 .Controls(strCmd & "_L" & IMG_R).Visible = False
14750                 .Controls(strCmd & "_R" & IMG_R).Visible = False
14760                 .Controls(strCmd & "_R" & IMG_RF).Visible = False
14770               Case False
14780                 .Controls(strCmd & "_L" & IMG_RF).Visible = True
14790                 .Controls(strCmd & "_R" & IMG_R).Visible = False
14800                 .Controls(strCmd & "_L" & IMG_R).Visible = False
14810                 .Controls(strCmd & "_R" & IMG_RF).Visible = False
14820               End Select
14830             Case False  ' ** So button points right, to open.
14840               Select Case blnMemoReveal_Focus
                    Case True
14850                 .Controls(strCmd & "_R" & IMG_RF).Visible = True
14860                 .Controls(strCmd & "_R" & IMG_R).Visible = False
14870                 .Controls(strCmd & "_L" & IMG_R).Visible = False
14880                 .Controls(strCmd & "_L" & IMG_RF).Visible = False
14890               Case False
14900                 .Controls(strCmd & "_R" & IMG_RF).Visible = True
14910                 .Controls(strCmd & "_R" & IMG_R).Visible = False
14920                 .Controls(strCmd & "_L" & IMG_R).Visible = False
14930                 .Controls(strCmd & "_L" & IMG_RF).Visible = False
14940               End Select
14950             End Select
14960             For intX = 1 To 2
14970               Select Case intX
                    Case 1
14980                 strTmp01 = "_L"
14990               Case 2
15000                 strTmp01 = "_R"
15010               End Select
15020               .Controls(strCmd & strTmp01 & IMG_RS).Visible = False
15030               .Controls(strCmd & strTmp01 & IMG_SF).Visible = False
15040               .Controls(strCmd & strTmp01 & IMG_D).Visible = False
15050             Next
15060           End If
15070         Case "Detail"
15080           For intX = 1& To 17&
15090             strCmd = vbNullString
15100             blnTmp04 = False
15110             Select Case intX
                  Case 1
15120               strCmd = "cmdAssetNew"
15130               blnTmp04 = blnAssetNew_Focus
15140             Case 2
15150               strCmd = "cmdLocNew"
15160               blnTmp04 = blnLocNew_Focus
15170             Case 3
15180               strCmd = "cmdRecurNew"
15190               blnTmp04 = blnRecurNew_Focus
15200             Case 4
15210               strCmd = "cmdRefresh"
15220               blnTmp04 = blnRefresh_Focus
15230             Case 5
15240               strCmd = "cmdPreviewReport"
15250               blnTmp04 = blnPreviewReport_Focus
15260             Case 6
15270               strCmd = "cmdPrintReport"
15280               blnTmp04 = blnPrintReport_Focus
15290             Case 7
15300               strCmd = "cmdSwitch"
15310               blnTmp04 = blnSwitch_Focus
15320             Case 8
15330               strCmd = "cmdUncomComAll"
15340               blnTmp04 = blnUncomCom_Focus
15350             Case 9
15360               strCmd = "cmdUncomDelAll"
15370               blnTmp04 = blnUncomDel_Focus
15380             Case 10
15390               strCmd = "cmdUnCommitOne"
15400               blnTmp04 = blnUnCommitOne_Focus
15410             Case 11
15420               strCmd = "cmdSpecPurp_Div_Map"
15430               blnTmp04 = blnMapDiv_Focus
15440             Case 12
15450               strCmd = "cmdSpecPurp_Int_Map"
15460               blnTmp04 = blnMapInt_Focus
15470             Case 13
15480               strCmd = "cmdSpecPurp_Sold_PaidTotal"
15490               blnTmp04 = blnPaid_Focus
15500             Case 14
15510               strCmd = "cmdSpecPurp_Purch_MapSplit"
15520               blnTmp04 = blnMapSplit_Focus
15530             Case 15
15540               strCmd = "cmdSpecPurp_Misc_MapLTCG"
15550               blnTmp04 = blnMapLTCG_Focus
15560             Case 16
15570               strCmd = "cmdSpecPurp_Misc_MapLTCL"
15580               blnTmp04 = blnMapLTCL_Focus
15590             Case 17
15600               strCmd = "cmdSpecPurp_Misc_MapSTCGL"
15610               blnTmp04 = blnMapSTCGL_Focus
15620             End Select
15630             If .Controls(strCmd & IMG_RFD).Visible = True Or .Controls(strCmd & IMG_RF).Visible = True Then
15640               Select Case blnTmp04
                    Case True
15650                 .Controls(strCmd & IMG_RSD).Visible = True
15660                 .Controls(strCmd & IMG_R).Visible = False
15670               Case False
15680                 .Controls(strCmd & IMG_R).Visible = True
15690                 .Controls(strCmd & IMG_RSD).Visible = False
15700               End Select
15710               .Controls(strCmd & IMG_RF).Visible = False
15720               .Controls(strCmd & IMG_RFD).Visible = False
15730               .Controls(strCmd & IMG_SFD).Visible = False
15740               .Controls(strCmd & IMG_D).Visible = False
15750             End If
15760           Next  ' ** intX.
15770           strCmd = "cmdMemoReveal"
15780           If .Controls(strCmd & "_R" & IMG_RF).Visible = True Or .Controls(strCmd & "_L" & IMG_RF).Visible = True Then
15790             Select Case .JrnlMemo_Memo.Visible
                  Case True  ' ** So button points left, to close.
15800               Select Case blnMemoReveal_Focus
                    Case True
15810                 .Controls(strCmd & "_L" & IMG_RS).Visible = True
15820                 .Controls(strCmd & "_L" & IMG_R).Visible = False
15830                 .Controls(strCmd & "_R" & IMG_R).Visible = False
15840                 .Controls(strCmd & "_R" & IMG_RS).Visible = False
15850               Case False
15860                 .Controls(strCmd & "_L" & IMG_R).Visible = True
15870                 .Controls(strCmd & "_L" & IMG_RS).Visible = False
15880                 .Controls(strCmd & "_R" & IMG_R).Visible = False
15890                 .Controls(strCmd & "_R" & IMG_RS).Visible = False
15900               End Select
15910             Case False  ' ** So button points right, to open.
15920               Select Case blnMemoReveal_Focus
                    Case True
15930                 .Controls(strCmd & "_R" & IMG_RS).Visible = True
15940                 .Controls(strCmd & "_R" & IMG_R).Visible = False
15950                 .Controls(strCmd & "_L" & IMG_R).Visible = False
15960                 .Controls(strCmd & "_L" & IMG_RS).Visible = False
15970               Case False
15980                 .Controls(strCmd & "_R" & IMG_R).Visible = True
15990                 .Controls(strCmd & "_R" & IMG_RS).Visible = False
16000                 .Controls(strCmd & "_L" & IMG_R).Visible = False
16010                 .Controls(strCmd & "_L" & IMG_RS).Visible = False
16020               End Select
16030             End Select
16040             .Controls(strCmd & "_L" & IMG_RF).Visible = False
16050             .Controls(strCmd & "_L" & IMG_SF).Visible = False
16060             .Controls(strCmd & "_L" & IMG_D).Visible = False
16070             .Controls(strCmd & "_R" & IMG_RF).Visible = False
16080             .Controls(strCmd & "_R" & IMG_SF).Visible = False
16090             .Controls(strCmd & "_R" & IMG_D).Visible = False
16100           End If
16110         End Select
16120       End Select ' ** strEvent.
16130     End With  ' ** frmPar.

16140   End If  ' ** strProc.

EXITP:
16150   Exit Sub

ERRH:
16160   Select Case ERR.Number
        Case Else
16170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16180   End Select
16190   Resume EXITP

End Sub

Public Sub JC_Btn_BoolReset(Optional varBtnGroupOnly As Variant)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_BoolReset"

        Dim blnBtnGroupOnly As Boolean

16210   Select Case IsMissing(varBtnGroupOnly)
        Case True
16220     blnBtnGroupOnly = False
16230   Case False
16240     blnBtnGroupOnly = CBool(varBtnGroupOnly)
16250   End Select
16260   blnAssetNew_Focus = False: blnAssetNew_MouseDown = False
16270   blnLocNew_Focus = False: blnLocNew_MouseDown = False
16280   blnRecurNew_Focus = False: blnRecurNew_MouseDown = False
16290   blnRefresh_Focus = False: blnRefresh_MouseDown = False
16300   If blnBtnGroupOnly = False Then
16310     blnPreviewReport_Focus = False: blnPreviewReport_MouseDown = False
16320     blnPrintReport_Focus = False: blnPrintReport_MouseDown = False
16330     blnSwitch_Focus = False: blnSwitch_MouseDown = False
16340     blnMemoReveal_Focus = False: blnMemoReveal_MouseDown = False
16350   End If

EXITP:
16360   Exit Sub

ERRH:
16370   Select Case ERR.Number
        Case Else
16380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16390   End Select
16400   Resume EXITP

End Sub

Public Sub JC_Btn_FocusSet(strProc As String, blnGotFocus As Boolean)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_FocusSet"

16510   If Trim(strProc) <> vbNullString Then
16520     strProc = Left(strProc, (InStr(strProc, "_") - 1))
16530     Select Case strProc
          Case "cmdAssetNew"
16540       blnAssetNew_Focus = blnGotFocus
16550     Case "cmdLocNew"
16560       blnLocNew_Focus = blnGotFocus
16570     Case "cmdRecurNew"
16580       blnRecurNew_Focus = blnGotFocus
16590     Case "cmdRefresh"
16600       blnRefresh_Focus = blnGotFocus
16610     End Select
16620   End If

EXITP:
16630   Exit Sub

ERRH:
16640   Select Case ERR.Number
        Case Else
16650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16660   End Select
16670   Resume EXITP

End Sub

Public Sub JC_Btn_Clear()

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "JC_Btn_Clear"

16710   lngSpecPurps = 0&
16720   ReDim arr_varSpecPurp(SP_ELEMS, 0)
16730   lngImgVars = 0&
16740   ReDim arr_varImgVar(IV_ELEMS, 0)

EXITP:
16750   Exit Sub

ERRH:
16760   Select Case ERR.Number
        Case Else
16770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16780   End Select
16790   Resume EXITP

End Sub
