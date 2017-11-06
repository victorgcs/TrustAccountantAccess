VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfIncExpCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfIncExpCodes"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmIncomeExpenseCodes"

Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         lngPageHeaderSection_Height = .Section(acPageHeader).Height
170         lngOneLine_Top = .revcode_DESC_lbl.Top
180         lngTwoLine_Top = 0&
190         lngUnderline_Top = .revcode_DESC_lbl_line.Top

200         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
210         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
220         .AsOf2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

230         .Caption = "List_of_Income_Expense_Codes_" & Format(Date, "yymmdd")

240         Select Case Forms(strCallingForm).frmIncomeExpenseCodes_Sub.Form.taxcode_type.Visible
            Case True
250           .TaxCode_Type_Description.Visible = True
260         Case False
270           .TaxCode_Type_Description.Visible = False
280         End Select

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
380     Forms(strCallingForm).Visible = True
390     DoCmd.Restore
400     Select Case ERR.Number
        Case Else
410       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
420     End Select
430     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

510     With Me
520       lngTmp01 = .AsOf2.Height
530       If .Page >= 2 Then
540         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
550         .revcode_DESC_lbl.Top = (lngOneLine_Top + lngTmp01)
560         .revcode_DESC_lbl_line.Top = (lngUnderline_Top + lngTmp01)
570         .revcode_SORTORDER_lbl.Top = (lngOneLine_Top + lngTmp01)
580         .revcode_SORTORDER_lbl_line.Top = (lngUnderline_Top + lngTmp01)
590         .revcode_ACTIVE_lbl.Top = (lngOneLine_Top + lngTmp01)
600         .revcode_ACTIVE_lbl_line.Top = (lngUnderline_Top + lngTmp01)
610         .AsOf2.Visible = True
620         .Header_lbl2.Visible = True
630         .Header_hline05.Visible = True
640       Else
650         .AsOf2.Visible = False
660         .Header_lbl2.Visible = False
670         .Header_hline05.Visible = False
680         .revcode_DESC_lbl.Top = lngOneLine_Top
690         .revcode_DESC_lbl_line.Top = lngUnderline_Top
700         .revcode_SORTORDER_lbl.Top = lngOneLine_Top
710         .revcode_SORTORDER_lbl_line.Top = lngUnderline_Top
720         .revcode_ACTIVE_lbl.Top = lngOneLine_Top
730         .revcode_ACTIVE_lbl_line.Top = lngUnderline_Top
740         .Section(acPageHeader).Height = lngPageHeaderSection_Height
750       End If
760     End With

EXITP:
770     Exit Sub

ERRH:
780     Forms("frmRecurringItems").Visible = True
790     DoCmd.Restore
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub Report_Close()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

910     DoCmd.Restore

920     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

930     If Forms(strCallingForm).Visible = False Then
940       Forms(strCallingForm).Visible = True
950       DoCmd.SelectObject acForm, strCallingForm, False
960       DoCmd.Restore
970     End If

EXITP:
980     Exit Sub

ERRH:
990     Forms(strCallingForm).Visible = True
1000    DoCmd.Restore
1010    Select Case ERR.Number
        Case Else
1020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1030    End Select
1040    Resume EXITP

End Sub
