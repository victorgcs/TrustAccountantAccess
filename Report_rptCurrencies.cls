VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCurrencies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCurrencies"

'VGC 09/26/2015: CHANGES!

Private Const strCallingForm As String = "frmCurrency"

Private frmPar As Access.Form, frmSub As Access.Form
Private lngTpp As Long
Private lngDetail_Height As Long, lngNotes_Top As Long
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

            '.detail_height_line.Top = 0&  'NO, IT'S USED BELOW!
            '.detail_height_line.Height = lngTpp

150         lngPageHeaderSection_Height = .Section(acPageHeader).Height
160         lngOneLine_Top = .curr_word1_lbl.Top
170         lngTwoLine_Top = .curr_code_lbl.Top
180         lngUnderline_Top = .curr_code_lbl_line.Top
190         lngDetail_Height = .Detail.Height
200         lngNotes_Top = .curr_notes.Top

210         Set frmPar = Forms(strCallingForm)
220         Set frmSub = frmPar.frmCurrency_Sub.Form

230         Select Case frmPar.chkPageOf
            Case True
240           .PageNumOf.Visible = True
250           .PageNum.Visible = False
260         Case False
270           .PageNumOf.Visible = False
280           .PageNum.Visible = True
290         End Select

300         If frmPar.chkIncludeNotes = False Then
310           .curr_notes.Visible = False
320           .curr_notes.Top = 0&
330           .curr_notes_lbl.Top = 0&
340           .Detail.Height = .detail_height_line.Height
350         End If

360         .AsOf.Caption = Format$(Date, "mmmm dd, yyyy")
370         .AsOf_lbl2.Caption = "As of " & Format$(Date, "mmmm dd, yyyy")
380         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

390         strOrderBy = Forms(strCallingForm).frmCurrency_Sub.Form.SortNow_Get  ' ** Form Function: frmCurrency_Sub.
400         .OrderBy = strOrderBy
410         .OrderByOn = True

420         .Caption = "Currencies_" & Format$(Date, "yymmdd")

430       Else
440         Cancel = -1
450       End If
460     End With

470     If Cancel = -1 Then
480       Beep
490       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
500     End If

EXITP:
510     Exit Sub

ERRH:
520     Forms(strCallingForm).Visible = True
530     DoCmd.Restore
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
650         .curr_code_lbl.Top = (lngTwoLine_Top + lngTmp01)
660         .curr_code_lbl_line.Top = (lngUnderline_Top + lngTmp01)
670         .curr_word1_lbl.Top = (lngOneLine_Top + lngTmp01)
680         .curr_word1_lbl_line.Top = (lngUnderline_Top + lngTmp01)
690         .curr_name_lbl.Top = (lngOneLine_Top + lngTmp01)
700         .curr_name_lbl_line.Top = (lngUnderline_Top + lngTmp01)
710         .curr_iso_lbl.Top = (lngTwoLine_Top + lngTmp01)
720         .curr_iso_lbl_line.Top = (lngUnderline_Top + lngTmp01)
730         .curr_rate1_lbl.Top = (lngTwoLine_Top + lngTmp01)
740         .curr_rate1_lbl_line.Top = (lngUnderline_Top + lngTmp01)
750         .curr_rate2_lbl.Top = (lngTwoLine_Top + lngTmp01)
760         .curr_rate2_lbl_line.Top = (lngUnderline_Top + lngTmp01)
770         .curr_date_lbl.Top = (lngOneLine_Top + lngTmp01)
780         .curr_date_lbl_line.Top = (lngUnderline_Top + lngTmp01)
790         .curr_active_lbl.Top = (lngTwoLine_Top + lngTmp01)
800         .curr_active_lbl_line.Top = (lngUnderline_Top + lngTmp01)
810         .country_code3_lbl.Top = (lngTwoLine_Top + lngTmp01)
820         .country_code3_lbl_line.Top = (lngUnderline_Top + lngTmp01)
830         .country_name_lbl.Top = (lngOneLine_Top + lngTmp01)
840         .country_name_lbl_line.Top = (lngUnderline_Top + lngTmp01)
850         .country_iso_lbl.Top = (lngTwoLine_Top + lngTmp01)
860         .country_iso_lbl_line.Top = (lngUnderline_Top + lngTmp01)
870         .country_active_lbl.Top = (lngTwoLine_Top + lngTmp01)
880         .country_active_lbl_line.Top = (lngUnderline_Top + lngTmp01)
890         .Header_lbl2.Visible = True
900         .AsOf_lbl2.Visible = True
910         .Header_hline05.Visible = True
920       Else
930         .Header_lbl2.Visible = False
940         .AsOf_lbl2.Visible = False
950         .Header_hline05.Visible = False
960         .curr_code_lbl.Top = lngTwoLine_Top
970         .curr_code_lbl_line.Top = lngUnderline_Top
980         .curr_word1_lbl.Top = lngOneLine_Top
990         .curr_word1_lbl_line.Top = lngUnderline_Top
1000        .curr_name_lbl.Top = lngOneLine_Top
1010        .curr_name_lbl_line.Top = lngUnderline_Top
1020        .curr_iso_lbl.Top = lngTwoLine_Top
1030        .curr_iso_lbl_line.Top = lngUnderline_Top
1040        .curr_rate1_lbl.Top = lngTwoLine_Top
1050        .curr_rate1_lbl_line.Top = lngUnderline_Top
1060        .curr_rate2_lbl.Top = lngTwoLine_Top
1070        .curr_rate2_lbl_line.Top = lngUnderline_Top
1080        .curr_date_lbl.Top = lngOneLine_Top
1090        .curr_date_lbl_line.Top = lngUnderline_Top
1100        .curr_active_lbl.Top = lngTwoLine_Top
1110        .curr_active_lbl_line.Top = lngUnderline_Top
1120        .country_code3_lbl.Top = lngTwoLine_Top
1130        .country_code3_lbl_line.Top = lngUnderline_Top
1140        .country_name_lbl.Top = lngOneLine_Top
1150        .country_name_lbl_line.Top = lngUnderline_Top
1160        .country_iso_lbl.Top = lngTwoLine_Top
1170        .country_iso_lbl_line.Top = lngUnderline_Top
1180        .country_active_lbl.Top = lngTwoLine_Top
1190        .country_active_lbl_line.Top = lngUnderline_Top
1200        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1210      End If
1220    End With

EXITP:
1230    Exit Sub

ERRH:
1240    Select Case ERR.Number
        Case Else
1250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1260    End Select
1270    Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

1310    With Me
1320      If frmPar.chkIncludeNotes = True Then
1330        Select Case IsNull(.curr_notes)
            Case True
1340          .curr_notes.Visible = False
1350          .curr_notes.Top = 0&
1360          .curr_notes_lbl.Top = 0&
1370          .Detail.Height = .detail_height_line.Height
1380        Case False
1390          .Detail.Height = lngDetail_Height
1400          .curr_notes.Top = lngNotes_Top
1410          .curr_notes_lbl.Top = lngNotes_Top
1420          .curr_notes.Visible = True
1430        End Select
1440      End If
1450    End With

EXITP:
1460    Exit Sub

ERRH:
1470    Forms(strCallingForm).Visible = True
1480    DoCmd.Restore
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Report_Close()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1610    DoCmd.Restore

1620    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1630    If Forms(strCallingForm).Visible = False Then
1640      Forms(strCallingForm).Visible = True
1650      DoCmd.SelectObject acForm, strCallingForm, False
1660      DoCmd.Restore
1670    End If

EXITP:
1680    Set frmSub = Nothing
1690    Set frmPar = Nothing
1700    Exit Sub

ERRH:
1710    Forms(strCallingForm).Visible = True
1720    DoCmd.Restore
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub
