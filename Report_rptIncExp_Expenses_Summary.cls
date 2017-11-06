VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptIncExp_Expenses_Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptIncExp_Expenses_Summary"

'VGC 08/17/2012: CHANGES!

Private Const strCallingForm As String = "frmRpt_IncomeExpense"

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngSubHeaderLbl_Top As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         lngPageHeaderSection_Height = .Section(acPageHeader).Height
150         lngOneLine_Top = .revcode_DESC_lbl.Top
160         lngTwoLine_Top = .icash_lbl.Top
170         lngUnderline_Top = .revcode_DESC_lbl_line.Top
180         lngSubHeaderLbl_Top = .SubHeader_lbl.Top

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         Set frm = Forms(strCallingForm)

210         Select Case frm.chkPageOf
            Case True
220           .PageNumOf.Visible = True
230           .PageNum.Visible = False
240         Case False
250           .PageNum.Visible = True
260           .PageNumOf.Visible = False
270           .PageNumOf.ControlSource = vbNullString
280         End Select

290         Select Case frm.opgUnspecified
            Case frm.opgUnspecified_optInclude.OptionValue
300           Select Case frm.opgPrincipalCash
              Case frm.opgPrincipalCash_optInclude.OptionValue
                ' ** .RecordSource = "qryIncomeExpenseReports_44_all"  ' ** Default.
310           Case frm.opgPrincipalCash_optExclude.OptionValue
320             RecordSource = "qryIncomeExpenseReports_49d_all"
330           End Select
340         Case frm.opgUnspecified_optOnly.OptionValue
350           Select Case frm.opgPrincipalCash
              Case frm.opgPrincipalCash_optInclude.OptionValue
360             .RecordSource = "qryIncomeExpenseReports_49f_all"
370           Case frm.opgPrincipalCash_optExclude.OptionValue
380             .RecordSource = "qryIncomeExpenseReports_49g_all"
390           End Select
400         Case frm.opgUnspecified_optExclude.OptionValue
410           Select Case frm.opgPrincipalCash
              Case frm.opgPrincipalCash_optInclude.OptionValue
420             .RecordSource = "qryIncomeExpenseReports_49c_all"
430           Case frm.opgPrincipalCash_optExclude.OptionValue
440             .RecordSource = "qryIncomeExpenseReports_49e_all"
450           End Select
460         End Select

470         .AsOf_lbl.Caption = "From " & Format(CDate(gdatStartDate), "mm/dd/yyyy") & _
              " To " & Format(CDate(gdatEndDate), "mm/dd/yyyy")
480         .CoInfoBlock.Caption = gstrCo_InfoBlock

490         .Caption = "Expenses_Summary_All_" & Format(gdatStartDate, "yymmdd") & "_To_" & _
              Format(gdatEndDate, "yymmdd")

500         If frm.chkSweepOnly = True Then
510           .AsOf.Caption = .AsOf.Caption & " - Sweep Asset Only"
520         End If

530         .RptSubtype_lbl.Caption = gstrCrtRpt_Ordinal  ' ** Borrowing this variable from Court Reports.

540       Else
550         Cancel = -1
560       End If
570     End With

580     If Cancel = -1 Then
590       Beep
600       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
610     End If

EXITP:
620     Set frm = Nothing
630     Exit Sub

ERRH:
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

710     With Me
720       lngTmp01 = .AsOf_lbl2.Height
730       If .Page >= 2 Then
740         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
750         .revcode_DESC_lbl.Top = (lngOneLine_Top + lngTmp01)
760         .revcode_DESC_lbl_line.Top = (lngUnderline_Top + lngTmp01)
770         .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
780         .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
790         .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
800         .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
810         .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
820         .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
830         .SubHeader_lbl.Top = (lngSubHeaderLbl_Top + lngTmp01)
840         .Header_lbl2.Visible = True
850         .AsOf_lbl2.Visible = True
860         .Header_hline05.Visible = True
870       Else
880         .SubHeader_lbl.Top = lngSubHeaderLbl_Top
890         .revcode_DESC_lbl.Top = lngOneLine_Top
900         .revcode_DESC_lbl_line.Top = lngUnderline_Top
910         .icash_lbl.Top = lngTwoLine_Top
920         .icash_lbl_line.Top = lngUnderline_Top
930         .pcash_lbl.Top = lngTwoLine_Top
940         .pcash_lbl_line.Top = lngUnderline_Top
950         .cost_lbl.Top = lngOneLine_Top
960         .cost_lbl_line.Top = lngUnderline_Top
970         .Header_lbl2.Visible = False
980         .AsOf_lbl2.Visible = False
990         .Header_hline05.Visible = False
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

Private Sub Report_NoData(Cancel As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1110    With Me
1120      .NoData_lbl.BackColor = CLR_WHT
1130      .NoData_lbl.Visible = True
1140      .icash_tot.Visible = False
1150      .icash_tot_line.Visible = False
1160      .pcash_tot.Visible = False
1170      .pcash_tot_line.Visible = False
1180      .cost_tot.Visible = False
1190      .cost_tot_line.Visible = False
1200    End With

EXITP:
1210    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub Report_Close()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1310    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1320    Exit Sub

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub
