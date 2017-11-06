VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAssetPricing_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAssetPricing_History"

'VGC 07/15/2016: CHANGES!

' **********************************************************************
'NOTE: Filter is inluded in OpenReport statement for 'WhereCondition'.
' **********************************************************************

Private datCurrentDate As Date
Private strCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String
        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           intPos01 = InStr(Trim(.OpenArgs), "~")
150           If intPos01 > 0 Then

160             TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

                ' ** Default to PageNumOf if calling form has no check box.
170             .PageNumOf.Visible = True
180             .PageNum.Visible = False

190             strCallingForm = Trim(.OpenArgs)
200             datCurrentDate = CDate(Mid(strCallingForm, (intPos01 + 1)))
210             strCallingForm = Left(strCallingForm, (intPos01 - 1))

220             .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
230             .AsOf.Caption = Format(datCurrentDate, "mmmm dd, yyyy")
240             .MarketAsOf_lbl.Caption = "Printed " & Format(Date, "mm/dd/yyyy")

250             strOrderBy = Forms(strCallingForm).frmAssetPricing_History_Sub.Form.SortNow_Get  ' ** Form Function: frmAssetPricing_History_Sub.
                ' ** Potential Sorts:
                ' **   [assettype], [totdesc], [cusip]
                ' **   [assetno]
                ' **   [cusip], [totdesc], [assettype]
                ' **   [totdesc], [cusip], [assettype]
                ' **   [assettype_description], [totdesc], [cusip]
                ' **   [curr_code], [totdesc], [assettype]
                ' **   [yield], [totdesc], [assettype]
                ' **   [marketvaluecurrent], [totdesc], [assettype]
                ' **   [marketvaluecurrent_usd], [totdesc], [assettype]
260             .OrderBy = strOrderBy
270             .OrderByOn = True

280             .Caption = "Asset_Pricing_History_For_" & Format(datCurrentDate, "yymmdd")

290           Else
300             Cancel = -1
310           End If
320         Else
330           Cancel = -1
340         End If
350       Else
360         Cancel = -1
370       End If
380     End With

390     If Cancel = -1 Then
400       Beep
410       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
420     End If

EXITP:
430     Exit Sub

ERRH:
440     Forms("frmAssetPricing").Visible = True
450     Forms(strCallingForm).Visible = True
460     DoCmd.Restore
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub Report_Close()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

610     DoCmd.Restore

620     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

630     If Forms("frmAssetPricing").Visible = False Then
640       Forms("frmAssetPricing").Visible = True
650     End If

660     If Forms(strCallingForm).Visible = False Then
670       Forms(strCallingForm).Visible = True
680       DoCmd.SelectObject acForm, strCallingForm, False
690       DoCmd.Restore
700     End If

EXITP:
710     Exit Sub

ERRH:
720     Forms("frmAssetPricing").Visible = True
730     Forms(strCallingForm).Visible = True
740     DoCmd.Restore
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Sub
