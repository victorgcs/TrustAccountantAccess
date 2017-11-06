VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfTaxCodes_Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfTaxCodes_Sort"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmTaxCode"
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

160         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
170         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

180         strOrderBy = Forms(strCallingForm).frmTaxCode_Sub.Form.SortNow_Get  ' ** Form Function: frmTaxCode_Sub.
            ' ** Potential Sorts:
            ' **   [taxcode_type], [taxcode_order]
            ' **   [taxcode]
            ' **   [taxcode_description]
            ' **   [taxcode_order], [taxcode_type]
            ' **   [taxcode_active], [taxcode_type], [taxcode_order]
            ' **   [revcode_TYPE], [taxcode_type], [taxcode_order]
190         .OrderBy = strOrderBy
200         .OrderByOn = True

210         .Caption = "List_of_Tax_Codes_" & Format(Date, "yymmdd")

220         Select Case gblnLinkRevTaxCodes
            Case True
230           .revcode_TYPE_Description.Visible = True
240           .revcode_TYPE_Description_lbl.Visible = True
250           .revcode_TYPE_Description_lbl_line.Visible = True
260         Case False
270           .revcode_TYPE_Description.Visible = False
280           .revcode_TYPE_Description_lbl.Visible = False
290           .revcode_TYPE_Description_lbl_line.Visible = False
300         End Select

310       Else
320         Cancel = -1
330       End If
340     End With

350     If Cancel = -1 Then
360       Beep
370       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
380     End If

EXITP:
390     Exit Sub

ERRH:
400     Forms(strCallingForm).Visible = True
410     DoCmd.Restore
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub Report_Close()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

510     DoCmd.Restore

520     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

530     If Forms(strCallingForm).Visible = False Then
540       Forms(strCallingForm).Visible = True
550       DoCmd.SelectObject acForm, strCallingForm, False
560       DoCmd.Restore
570     End If

EXITP:
580     Exit Sub

ERRH:
590     Forms(strCallingForm).Visible = True
600     DoCmd.Restore
610     Select Case ERR.Number
        Case Else
620       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
630     End Select
640     Resume EXITP

End Sub
