VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfCurrencySymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfCurrencySymbols"

'VGC 10/16/2015: CHANGES!

Private Const strCallingForm As String = "frmCurrency_Symbol"

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
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

150         lngPageHeaderSection_Height = .Section(acPageHeader).Height
160         lngOneLine_Top = .curr_word_lbl.Top
170         lngTwoLine_Top = .curr_code_lbl.Top
180         lngUnderline_Top = .curr_code_lbl_line.Top

            ' ** Default to PageNumOf if calling form has no check box.
190         .PageNumOf.Visible = True
200         .PageNum.Visible = False

210         .AsOf.Caption = Format$(Date, "mmmm dd, yyyy")
220         .AsOf_lbl2.Caption = "As of " & Format$(Date, "mmmm dd, yyyy")
230         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

240         strOrderBy = Forms(strCallingForm).frmCurrency_Symbol_Sub.Form.SortNow_Get  ' ** Form Function: frmCurrency_Symbol_Sub.
            'strOrderBy = StringReplace(strOrderBy, "country_namex", "country_name")  ' ** Module Function: modStringFuncs.
250         strOrderBy = StringReplace(strOrderBy, "currsym_symbol_display", "currsym_symbol")  ' ** Module Function: modStringFuncs.
260         strOrderBy = StringReplace(strOrderBy, "curr_word1", "curr_word")  ' ** Module Function: modStringFuncs.
270         .OrderBy = strOrderBy
280         .OrderByOn = True
 
290         .Caption = "List_of_Currency_Symbols_" & Format$(Date, "yymmdd")

300       Else
310         Cancel = -1
320       End If
330     End With

340     If Cancel = -1 Then
350       Beep
360       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
370     End If

EXITP:
380     Exit Sub

ERRH:
390     Forms(strCallingForm).Visible = True
400     DoCmd.Restore
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

510     With Me
520       lngTmp01 = .AsOf_lbl2.Height
530       If .Page >= 2 Then
540         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
550         lngTmp02 = (lngOneLine_Top + lngTmp01)
560         lngTmp03 = (lngTwoLine_Top + lngTmp01)
570         lngTmp04 = (lngUnderline_Top + lngTmp01)
580         .curr_code_lbl.Top = lngTmp03
590         .curr_code_lbl_line.Top = lngTmp04
600         .currsym_symbol_lbl.Top = lngTmp02
610         .currsym_symbol_lbl_line.Top = lngTmp04
620         .curr_word_lbl.Top = lngTmp02
630         .curr_word_lbl_line.Top = lngTmp04
640         .curr_active_lbl.Top = lngTmp02
650         .curr_active_lbl_line.Top = lngTmp04
660         .country_code3_lbl.Top = lngTmp03
670         .country_code3_lbl_line.Top = lngTmp04
680         .country_name1_lbl.Top = lngTmp02
690         .country_name1_lbl_line.Top = lngTmp04
700         .Header_lbl2.Visible = True
710         .AsOf_lbl2.Visible = True
720         .Header_hline05.Visible = True
730       Else
740         .Header_lbl2.Visible = False
750         .AsOf_lbl2.Visible = False
760         .Header_hline05.Visible = False
770         .curr_code_lbl.Top = lngTwoLine_Top
780         .curr_code_lbl_line.Top = lngUnderline_Top
790         .currsym_symbol_lbl.Top = lngOneLine_Top
800         .currsym_symbol_lbl_line.Top = lngUnderline_Top
810         .curr_word_lbl.Top = lngOneLine_Top
820         .curr_word_lbl_line.Top = lngUnderline_Top
830         .curr_active_lbl.Top = lngOneLine_Top
840         .curr_active_lbl_line.Top = lngUnderline_Top
850         .country_code3_lbl.Top = lngTwoLine_Top
860         .country_code3_lbl_line.Top = lngUnderline_Top
870         .country_name1_lbl.Top = lngOneLine_Top
880         .country_name1_lbl_line.Top = lngUnderline_Top
890         .Section(acPageHeader).Height = lngPageHeaderSection_Height
900       End If
910     End With

EXITP:
920     Exit Sub

ERRH:
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub

Private Sub Report_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1010    DoCmd.Restore

1020    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1030    If Forms(strCallingForm).Visible = False Then
1040      Forms(strCallingForm).Visible = True
1050      DoCmd.SelectObject acForm, strCallingForm, False
1060      DoCmd.Restore
1070    End If

EXITP:
1080    Exit Sub

ERRH:
1090    Forms(strCallingForm).Visible = True
1100    DoCmd.Restore
1110    Select Case ERR.Number
        Case Else
1120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1130    End Select
1140    Resume EXITP

End Sub
