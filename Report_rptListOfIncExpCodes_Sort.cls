VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfIncExpCodes_Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfIncExpCodes_Sort"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmIncomeExpenseCodes"

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
170         lngOneLine_Top = .revcode_DESC_lbl.Top
180         lngTwoLine_Top = 0&
190         lngUnderline_Top = .revcode_DESC_lbl_line.Top

200         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
210         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
220         .AsOf2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

230         strOrderBy = Forms(strCallingForm).frmIncomeExpenseCodes_Sub.Form.SortNow_Get  ' ** Form Function: frmIncomeExpenseCodes_Sub.
            ' ** Potential Sorts:
            ' **   [revcode_TYPE], [revcode_SORTORDER]
            ' **   [revcode_DESC], [revcode_TYPE], [revcode_SORTORDER]
            ' **   [revcode_SORTORDER], [revcode_TYPE]
            ' **   [revcode_ACTIVE], [revcode_TYPE], [revcode_SORTORDER]
            ' **   [TaxCode_Type_Description], [revcode_TYPE], [revcode_SORTORDER]
240         .OrderBy = strOrderBy
250         .OrderByOn = True

260         .Caption = "List_of_Income_Expense_Codes_" & Format(Date, "yymmdd")

270         Select Case Forms(strCallingForm).frmIncomeExpenseCodes_Sub.Form.taxcode_type.Visible
            Case True
280           .TaxCode_Type_Description.Visible = True
290           .TaxCode_Type_Description_lbl.Visible = True
300           .TaxCode_Type_Description_lbl_line.Visible = True
310         Case False
320           .TaxCode_Type_Description.Visible = False
330           .TaxCode_Type_Description_lbl.Visible = False
340           .TaxCode_Type_Description_lbl_line.Visible = False
350         End Select

360       Else
370         Cancel = -1
380       End If
390     End With

400     If Cancel = -1 Then
410       Beep
420       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
430     End If

EXITP:
440     Exit Sub

ERRH:
450     Forms(strCallingForm).Visible = True
460     DoCmd.Restore
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

610     With Me
620       lngTmp01 = .AsOf2.Height
630       If .Page >= 2 Then
640         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
650         .revcode_TYPE_Description_lbl.Top = (lngTwoLine_Top + lngTmp01)
660         .revcode_TYPE_Description_lbl_line.Top = (lngUnderline_Top + lngTmp01)
670         .revcode_DESC_lbl.Top = (lngOneLine_Top + lngTmp01)
680         .revcode_DESC_lbl_line.Top = (lngUnderline_Top + lngTmp01)
690         .revcode_SORTORDER_lbl.Top = (lngOneLine_Top + lngTmp01)
700         .revcode_SORTORDER_lbl_line.Top = (lngUnderline_Top + lngTmp01)
710         .revcode_ACTIVE_lbl.Top = (lngOneLine_Top + lngTmp01)
720         .revcode_ACTIVE_lbl_line.Top = (lngUnderline_Top + lngTmp01)
730         .TaxCode_Type_Description_lbl.Top = (lngOneLine_Top + lngTmp01)
740         .TaxCode_Type_Description_lbl_line.Top = (lngUnderline_Top + lngTmp01)
750         .AsOf2.Visible = True
760         .Header_lbl2.Visible = True
770         .Header_hline05.Visible = True
780       Else
790         .AsOf2.Visible = False
800         .Header_lbl2.Visible = False
810         .Header_hline05.Visible = False
820         .revcode_TYPE_Description_lbl.Top = lngTwoLine_Top
830         .revcode_TYPE_Description_lbl_line.Top = lngUnderline_Top
840         .revcode_DESC_lbl.Top = lngOneLine_Top
850         .revcode_DESC_lbl_line.Top = lngUnderline_Top
860         .revcode_SORTORDER_lbl.Top = lngOneLine_Top
870         .revcode_SORTORDER_lbl_line.Top = lngUnderline_Top
880         .revcode_ACTIVE_lbl.Top = lngOneLine_Top
890         .revcode_ACTIVE_lbl_line.Top = lngUnderline_Top
900         .TaxCode_Type_Description_lbl.Top = lngOneLine_Top
910         .TaxCode_Type_Description_lbl_line.Top = lngUnderline_Top
920         .Section(acPageHeader).Height = lngPageHeaderSection_Height
930       End If
940     End With

EXITP:
950     Exit Sub

ERRH:
960     Forms("frmRecurringItems").Visible = True
970     DoCmd.Restore
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub Report_Close()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1110    DoCmd.Restore

1120    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1130    If Forms(strCallingForm).Visible = False Then
1140      Forms(strCallingForm).Visible = True
1150      DoCmd.SelectObject acForm, strCallingForm, False
1160      DoCmd.Restore
1170    End If

EXITP:
1180    Exit Sub

ERRH:
1190    Forms(strCallingForm).Visible = True
1200    DoCmd.Restore
1210    Select Case ERR.Number
        Case Else
1220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1230    End Select
1240    Resume EXITP

End Sub
