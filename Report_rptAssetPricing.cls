VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetPricing"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmAssetPricing"

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
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

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
220         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

230         strOrderBy = Forms(strCallingForm).frmAssetPricing_Sub.Form.SortNow_Get  ' ** Form Function: frmAssetPricing_Sub.
            ' ** Potential sorts:
            ' **   [assettype], [description_masterasset_sort], [cusip]
            ' **   [assetno]
            ' **   [cusip], [description_masterasset_sort], [assettype]
            ' **   [description_masterasset_sort], [cusip], [assettype]
            ' **   [curr_code], [description_masterasset_sort], [assettype]
            ' **   [yield], [description_masterasset_sort], [assettype]
            ' **   [marketvaluecurrent], [description_masterasset_sort], [assettype]
            ' **   [marketvaluecurrent_usd], [description_masterasset_sort], [assettype]
            ' **   [marketvalue], [description_masterasset_sort], [assettype]
            ' **   [marketvalue_usd], [description_masterasset_sort], [assettype]
240         .OrderBy = strOrderBy
250         .OrderByOn = True

260         .Caption = "Asset_Pricing_" & Format(Date, "yymmdd")

270       Else
280         Cancel = -1
290       End If
300     End With

310     If Cancel = -1 Then
320       Beep
330       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
340     End If

EXITP:
350     Exit Sub

ERRH:
360     DoCmd.Restore
370     Forms(strCallingForm).Visible = True
380     Select Case ERR.Number
        Case Else
390       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
400     End Select
410     Resume EXITP

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
810         .marketvaluecurrent_lbl.Top = (lngTwoLine_Top + lngTmp01)
820         .marketvaluecurrent_lbl_line.Top = (lngUnderline_Top + lngTmp01)
830         .Header_lbl2.Visible = True
840         .AsOf2.Visible = True
850         .Header_hline05.Top = lngTmp01
860         .Header_hline05.Visible = True
870       Else
880         .Header_lbl2.Visible = False
890         .AsOf2.Visible = False
900         .Header_hline05.Visible = False
910         .Header_hline05.Top = .AsOf2.Height
920         .cusip_lbl.Top = lngOneLine_Top
930         .cusip_lbl_line.Top = lngUnderline_Top
940         .totdesc_lbl.Top = lngOneLine_Top
950         .totdesc_lbl_line.Top = lngUnderline_Top
960         .yield_lbl.Top = lngOneLine_Top
970         .yield_lbl_line.Top = lngUnderline_Top
980         .marketvaluecurrent_lbl.Top = lngTwoLine_Top
990         .marketvaluecurrent_lbl_line.Top = lngUnderline_Top
1000        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1010      End If
1020    End With

EXITP:
1030    Exit Sub

ERRH:
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub Report_Close()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1110    DoCmd.Restore

1120    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1130    If Forms(strCallingForm).Visible = False Then
1140      Forms(strCallingForm).Visible = True
1150    End If

EXITP:
1160    Exit Sub

ERRH:
1170    DoCmd.Restore
1180    Forms(strCallingForm).Visible = True
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1210    End Select
1220    Resume EXITP

End Sub
