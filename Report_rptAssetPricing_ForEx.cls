VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetPricing_ForEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetPricing_ForEx"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmAssetPricing"

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long, lngUnderlineTwo_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         lngPageHeaderSection_Height = .Section(acPageHeader).Height
170         lngOneLine_Top = .cusip_lbl.Top
180         lngTwoLine_Top = .marketvaluecurrent_lbl.Top
190         lngUnderline_Top = .cusip_lbl_line.Top
200         lngUnderlineTwo_Top = .marketvaluecurrent_lbl_line.Top

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
210         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

220         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
230         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

240         strOrderBy = Forms(strCallingForm).frmAssetPricing_Sub.Form.SortNow_Get  ' ** Form Function: frmAssetPricing_Sub.
250         .OrderBy = strOrderBy
260         .OrderByOn = True

270         .Caption = "Asset_Pricing_" & Format(Date, "yymmdd")

280       Else
290         Cancel = -1
300       End If
310     End With

320     If Cancel = -1 Then
330       Beep
340       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
350     End If

EXITP:
360     Exit Sub

ERRH:
370     DoCmd.Restore
380     Forms(strCallingForm).Visible = True
390     Select Case ERR.Number
        Case Else
400       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
410     End Select
420     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

510     With Me
520       .MarketAsOf_lbl.Caption = "Market as of " & Format(.currentDate, "mm/dd/yyyy")
530     End With

EXITP:
540     Exit Sub

ERRH:
550     DoCmd.Restore
560     Forms(strCallingForm).Visible = True
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

710     With Me
720       lngTmp01 = .AsOf2.Height
730       If .Page >= 2 Then
740         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
750         .cusip_lbl.Top = (lngOneLine_Top + lngTmp01)
760         .cusip_lbl_line.Top = (lngUnderline_Top + lngTmp01)
770         .totdesc_lbl.Top = (lngOneLine_Top + lngTmp01)
780         .totdesc_lbl_line.Top = (lngUnderline_Top + lngTmp01)
790         .yield_lbl.Top = (lngOneLine_Top + lngTmp01)
800         .yield_lbl_line.Top = (lngUnderline_Top + lngTmp01)
810         .curr_code_lbl.Top = (lngOneLine_Top + lngTmp01)
820         .curr_code_lbl_line.Top = (lngUnderline_Top + lngTmp01)
830         .marketvaluecurrent_lbl.Top = (lngTwoLine_Top + lngTmp01)
840         .marketvaluecurrent_lbl_line.Top = (lngUnderlineTwo_Top + lngTmp01)
850         .mktvalcurr_str_lbl.Top = (lngOneLine_Top + lngTmp01)
860         .mktvalcurr_str_lbl_line.Top = (lngUnderline_Top + lngTmp01)
870         .mktvalcurr_usd_str_lbl.Top = (lngOneLine_Top + lngTmp01)
880         .mktvalcurr_usd_str_lbl_line.Top = (lngUnderline_Top + lngTmp01)
890         .Header_lbl2.Visible = True
900         .AsOf2.Visible = True
910         .Header_hline05.Top = lngTmp01
920         .Header_hline05.Visible = True
930       Else
940         .Header_lbl2.Visible = False
950         .AsOf2.Visible = False
960         .Header_hline05.Visible = False
970         .Header_hline05.Top = .AsOf2.Height
980         .cusip_lbl.Top = lngOneLine_Top
990         .cusip_lbl_line.Top = lngUnderline_Top
1000        .totdesc_lbl.Top = lngOneLine_Top
1010        .totdesc_lbl_line.Top = lngUnderline_Top
1020        .yield_lbl.Top = lngOneLine_Top
1030        .yield_lbl_line.Top = lngUnderline_Top
1040        .curr_code_lbl.Top = lngOneLine_Top
1050        .curr_code_lbl_line.Top = lngUnderline_Top
1060        .marketvaluecurrent_lbl.Top = lngTwoLine_Top
1070        .marketvaluecurrent_lbl_line.Top = lngUnderlineTwo_Top
1080        .mktvalcurr_str_lbl.Top = lngOneLine_Top
1090        .mktvalcurr_str_lbl_line.Top = lngUnderline_Top
1100        .mktvalcurr_usd_str_lbl.Top = lngOneLine_Top
1110        .mktvalcurr_usd_str_lbl_line.Top = lngUnderline_Top
1120        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1130      End If
1140    End With

EXITP:
1150    Exit Sub

ERRH:
1160    Select Case ERR.Number
        Case Else
1170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1180    End Select
1190    Resume EXITP

End Sub

Private Sub Report_Close()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1210    DoCmd.Restore

1220    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1230    If Forms(strCallingForm).Visible = False Then
1240      Forms(strCallingForm).Visible = True
1250    End If

EXITP:
1260    Exit Sub

ERRH:
1270    DoCmd.Restore
1280    Forms(strCallingForm).Visible = True
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub
