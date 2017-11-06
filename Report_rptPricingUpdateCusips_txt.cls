VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptPricingUpdateCusips_txt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptPricingUpdateCusips_txt"

'VGC 02/03/2013: CHANGES!

Private Const strCallingForm As String = "frmAssetPricing_Import"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Select Case Forms(strCallingForm).chkPageOf
            Case True
150           .PageNumOf.Visible = True
160           .PageNum.Visible = False
170         Case False
180           .PageNum.Visible = True
190           .PageNumOf.Visible = False
200           .PageNumOf.ControlSource = vbNullString
210         End Select

220         .AsOf.Caption = Format(Date, "mmmm d, yyyy")
230         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

240         .RptSubtype_lbl.Caption = Forms(strCallingForm).priceimport_id_lbx.Column(1)

250         .Caption = "Asset_Pricing_Update_" & Format(Date, "yymmdd")

260         If gstrReportCallingForm <> vbNullString Then
270           Forms(gstrReportCallingForm).Visible = False
280         End If

290       Else
300         Cancel = -1
310       End If
320     End With

330     If Cancel = -1 Then
340       Beep
350       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
360     End If

EXITP:
370     Exit Sub

ERRH:
380     If gstrReportCallingForm <> vbNullString Then
390       DoCmd.Restore
400       Forms(gstrReportCallingForm).Visible = True
410       DoCmd.SelectObject acForm, gstrReportCallingForm, False
420     End If
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Report_Close()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

510     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

520     DoCmd.Restore

530     If gstrReportCallingForm <> vbNullString Then
540       Forms(gstrReportCallingForm).Visible = True
550     End If

EXITP:
560     Exit Sub

ERRH:
570     If gstrReportCallingForm <> vbNullString Then
580       DoCmd.Restore
590       Forms(gstrReportCallingForm).Visible = True
600       DoCmd.SelectObject acForm, gstrReportCallingForm, False
610     End If
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub
