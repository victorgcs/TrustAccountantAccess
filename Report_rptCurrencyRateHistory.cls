VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCurrencyRateHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCurrencyRateHistory"

'VGC 05/30/2016: CHANGES!

Private Const strCallingForm As String = "frmCurrency_History"

Private strFilter As String, lngTpp As Long
Private frmPar As Access.Form, frmSub As Access.Form
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

150         lngPageHeaderSection_Height = .Section(acPageHeader).Height
160         lngOneLine_Top = .curr_name_lbl.Top
170         lngTwoLine_Top = .curr_date_lbl.Top
180         lngUnderline_Top = .curr_date_lbl_line.Top

190         Set frmPar = Forms(strCallingForm)
200         Set frmSub = frmPar.frmCurrency_History_Sub.Form

210         .PageNumOf.Visible = True
220         .PageNum.Visible = False

230         .AsOf_lbl.Caption = "As of " & Format(Date, "mmmm dd, yyyy")
240         .AsOf_lbl2.Caption = "As of " & Format$(Date, "mmmm dd, yyyy")
250         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

260         strFilter = frmSub.Filter
270         .Filter = strFilter
280         .FilterOn = True

290         strOrderBy = frmSub.SortNow_Get  ' ** Form Function: frmCurrency_History_Sub.
300         .OrderBy = strOrderBy
310         .OrderByOn = True

320         .Caption = "Currency_Rate_History_" & Format$(Date, "yymmdd")

330       Else
340         Cancel = -1
350       End If
360     End With

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     Forms(strCallingForm).Visible = True
430     DoCmd.Restore
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

510     With Me
520       If InStr(strFilter, "curr_date") > 0 Then
530         .AsOf.Caption = "Rates as of: " & Format(.curr_date, "mm/dd/yyyy")
540       Else
550         .AsOf.Caption = .curr_code & "  " & .curr_name
560       End If
570     End With

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

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim lngTmp01 As Long

710     With Me
720       lngTmp01 = .AsOf_lbl2.Height
730       If .Page >= 2 Then
740         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
750         .curr_date_lbl.Top = (lngTwoLine_Top + lngTmp01)
760         .curr_date_lbl_line.Top = (lngUnderline_Top + lngTmp01)
770         .curr_code_lbl.Top = (lngTwoLine_Top + lngTmp01)
780         .curr_code_lbl_line.Top = (lngUnderline_Top + lngTmp01)
790         .curr_name_lbl.Top = (lngOneLine_Top + lngTmp01)
800         .curr_name_lbl_line.Top = (lngUnderline_Top + lngTmp01)
810         .curr_rate1_lbl.Top = (lngTwoLine_Top + lngTmp01)
820         .curr_rate1_lbl_line.Top = (lngUnderline_Top + lngTmp01)
830         .curr_rate2_lbl.Top = (lngTwoLine_Top + lngTmp01)
840         .curr_rate2_lbl_line.Top = (lngUnderline_Top + lngTmp01)
850         .currsym_symbol_lbl.Top = (lngOneLine_Top + lngTmp01)
860         .currsym_symbol_lbl_line.Top = (lngUnderline_Top + lngTmp01)
870         .curr_active_lbl.Top = (lngOneLine_Top + lngTmp01)
880         .curr_active_lbl_line.Top = (lngUnderline_Top + lngTmp01)
890         .Header_lbl2.Visible = True
900         .AsOf_lbl2.Visible = True
910         .Header_hline05.Visible = True
920       Else
930         .Header_lbl2.Visible = False
940         .AsOf_lbl2.Visible = False
950         .Header_hline05.Visible = False
960         .curr_date_lbl.Top = lngTwoLine_Top
970         .curr_date_lbl_line.Top = lngUnderline_Top
980         .curr_code_lbl.Top = lngTwoLine_Top
990         .curr_code_lbl_line.Top = lngUnderline_Top
1000        .curr_name_lbl.Top = lngOneLine_Top
1010        .curr_name_lbl_line.Top = lngUnderline_Top
1020        .curr_rate1_lbl.Top = lngTwoLine_Top
1030        .curr_rate1_lbl_line.Top = lngUnderline_Top
1040        .curr_rate2_lbl.Top = lngTwoLine_Top
1050        .curr_rate2_lbl_line.Top = lngUnderline_Top
1060        .currsym_symbol_lbl.Top = lngOneLine_Top
1070        .currsym_symbol_lbl_line.Top = lngUnderline_Top
1080        .curr_active_lbl.Top = lngOneLine_Top
1090        .curr_active_lbl_line.Top = lngUnderline_Top
1100        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1110      End If
1120    End With

EXITP:
1130    Exit Sub

ERRH:
1140    Forms(strCallingForm).Visible = True
1150    DoCmd.Restore
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

1230    If frmPar.Visible = False Then
1240      frmPar.Visible = True
1250      DoCmd.SelectObject acForm, strCallingForm, False
1260      DoCmd.Restore
1270    End If

EXITP:
1280    Set frmSub = Nothing
1290    Set frmPar = Nothing
1300    Exit Sub

ERRH:
1310    Forms(strCallingForm).Visible = True
1320    DoCmd.Restore
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub
