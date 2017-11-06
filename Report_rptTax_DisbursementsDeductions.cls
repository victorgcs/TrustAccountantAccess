VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTax_DisbursementsDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTax_DisbursementsDeductions"

'VGC 03/14/2013: CHANGES!

Private Const strCallingForm As String = "frmRpt_TaxIncomeDeductions"

Private lngPageHeaderSection_Height As Long, lngAccountNo_Top As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strAccountNo As String

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

220         lngPageHeaderSection_Height = .Section(acPageHeader).Height
230         lngAccountNo_Top = .accountno.Top
240         lngOneLine_Top = .transdate_lbl.Top
250         lngTwoLine_Top = .icash_lbl.Top
260         lngUnderline_Top = .transdate_lbl_line.Top

270         .AsOf_lbl.Caption = "From " & Format(Forms(strCallingForm).DateStart, "mm/dd/yyyy") & " To " & _
              Format(Forms(strCallingForm).DateEnd, "mm/dd/yyyy")
280         .AsOf_lbl2.Caption = "From " & Format(Forms(strCallingForm).DateStart, "mm/dd/yyyy") & " To " & _
              Format(Forms(strCallingForm).DateEnd, "mm/dd/yyyy")
290         .AsOf_lbl3.Caption = "From " & Format(Forms(strCallingForm).DateStart, "mm/dd/yyyy") & " To " & _
              Format(Forms(strCallingForm).DateEnd, "mm/dd/yyyy")
300         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

310         Select Case Forms(strCallingForm).opgAccountNumber
            Case Forms(strCallingForm).opgAccountNumber_optSpecified.OptionValue
320           strAccountNo = Forms(strCallingForm).cmbAccounts
330           .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
340         Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue
350           strAccountNo = "All"
360           .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
370         End Select

380         Select Case Forms(strCallingForm).chkShowTrans
            Case True
390           .AllTrans_lbl.Visible = True
400           .Caption = "Deductions_Tax_Report_All_Transactions_" & strAccountNo & "_" & _
                Format(Forms(strCallingForm).DateStart, "yymmdd") & "_To_" & _
                Format(Forms(strCallingForm).DateEnd, "yymmdd")
410         Case False
420           .AllTrans_lbl.Visible = False
430           .Caption = "Deductions_Tax_Report_" & strAccountNo & "_" & _
                Format(Forms(strCallingForm).DateStart, "yymmdd") & "_To_" & _
                Format(Forms(strCallingForm).DateEnd, "yymmdd")
440         End Select

450       Else
460         Cancel = -1
470       End If
480     End With

490     If Cancel = -1 Then
500       Beep
510       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
520     End If

EXITP:
530     Exit Sub

ERRH:
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

610     With Me
620       lngTmp01 = .AsOf_lbl2.Height
630       If .Page >= 2 Then
640         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
650         .accountno.Top = (lngAccountNo_Top + lngTmp01)
660         .shortname.Top = .accountno.Top
670         .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
680         .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
690         .TaxCode_Description_lbl.Top = (lngOneLine_Top + lngTmp01)
700         .TaxCode_Description_lbl_line.Top = (lngUnderline_Top + lngTmp01)
710         .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
720         .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
730         .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
740         .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
750         .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
760         .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
770         Select Case Forms(strCallingForm).chkShowTrans
            Case True
780           .Header_lbl3.Visible = True
790           .AsOf_lbl3.Visible = True
800         Case False
810           .Header_lbl2.Visible = True
820           .AsOf_lbl2.Visible = True
830         End Select
840         .Header_hline05.Visible = True
850       Else
860         .accountno.Top = lngAccountNo_Top
870         .shortname.Top = .accountno.Top
880         .transdate_lbl.Top = lngOneLine_Top
890         .transdate_lbl_line.Top = lngUnderline_Top
900         .TaxCode_Description_lbl.Top = lngOneLine_Top
910         .TaxCode_Description_lbl_line.Top = lngUnderline_Top
920         .cost_lbl.Top = lngOneLine_Top
930         .cost_lbl_line.Top = lngUnderline_Top
940         .icash_lbl.Top = lngTwoLine_Top
950         .icash_lbl_line.Top = lngUnderline_Top
960         .pcash_lbl.Top = lngTwoLine_Top
970         .pcash_lbl_line.Top = lngUnderline_Top
980         .Header_lbl3.Visible = False
990         .AsOf_lbl3.Visible = False
1000        .Header_lbl2.Visible = False
1010        .AsOf_lbl2.Visible = False
1020        .Header_hline05.Visible = False
1030        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1040      End If
1050    End With

EXITP:
1060    Exit Sub

ERRH:
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Private Sub Report_Close()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1210    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1220    Exit Sub

ERRH:
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub
