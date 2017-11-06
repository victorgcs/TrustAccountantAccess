VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfCountryCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfCountryCodes"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmCountryCode"

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim lngTmp01 As Long, lngTmp02 As Long
        Dim strOrderBy As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

150         lngPageHeaderSection_Height = .Section(acPageHeader).Height
160         lngOneLine_Top = .country_iso_lbl.Top
170         lngTwoLine_Top = .country_code3_lbl.Top
180         lngUnderline_Top = .country_iso_lbl_line.Top

            ' ** Default to PageNumOf if calling form has no check box.
190         .PageNumOf.Visible = True
200         .PageNum.Visible = False

210         .AsOf.Caption = Format$(Date, "mmmm dd, yyyy")
220         .AsOf_lbl2.Caption = "As of " & Format$(Date, "mmmm dd, yyyy")
230         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

240         If Forms(strCallingForm).chkShowCode2 = True Then
250           .country_code3_lbl.Left = (.country_code2_lbl.Left - (.country_code3_lbl.Width + (4& * lngTpp)))
260           .country_code3_lbl_line.Left = .country_code3_lbl.Left
270           .country_code3.Left = (.country_code3_lbl.Left + (.country_code2.Left - .country_code2_lbl.Left))
280           .country_name_lbl.Left = ((.country_code2_lbl.Left + .country_code2_lbl.Width) + (4& * lngTpp))
290           .country_name_lbl_line.Left = .country_name_lbl.Left
300           .country_name.Left = .country_name_lbl.Left
310           lngTmp01 = (.country_iso.Left - .country_iso_lbl.Left)
320           .country_iso_lbl.Left = ((.country_name_lbl.Left + .country_name_lbl.Width) + (4& * lngTpp))
330           .country_iso_lbl_line.Left = .country_iso_lbl.Left
340           .country_iso.Left = (.country_iso_lbl.Left + lngTmp01)
350           lngTmp01 = (.country_active.Left - .country_active_lbl.Left)
360           .country_active_lbl.Left = ((.country_iso_lbl.Left + .country_iso_lbl.Width) + (4& * lngTpp))
370           .country_active_lbl_line.Left = .country_active_lbl.Left
380           .country_active.Left = (.country_active_lbl.Left + lngTmp01)
390           lngTmp02 = ((.country_active_lbl.Left + .country_active_lbl.Width) + (4& * lngTpp))
400           .country_na_lbl.Left = lngTmp02
410           .country_na_lbl_line.Left = .country_na_lbl.Left
420           .country_na.Left = (.country_na_lbl.Left + lngTmp01)
430           .country_code2.Visible = True
440           .country_code2_lbl.Visible = True
450           .country_code2_lbl_line.Visible = True
460         End If

470         strOrderBy = Forms(strCallingForm).frmCountryCode_Sub.Form.SortNow_Get  ' ** Form Function: frmCountryCode_Sub.
480         .OrderBy = strOrderBy
490         .OrderByOn = True

500         .Caption = "List_of_Country_Codes_" & Format$(Date, "yymmdd")

510       Else
520         Cancel = -1
530       End If
540     End With

550     If Cancel = -1 Then
560       Beep
570       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
580     End If

EXITP:
590     Exit Sub

ERRH:
600     Forms(strCallingForm).Visible = True
610     DoCmd.Restore
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long

710     With Me
720       lngTmp01 = .AsOf_lbl2.Height
730       If .Page >= 2 Then
740         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
750         lngTmp02 = (lngOneLine_Top + lngTmp01)
760         lngTmp03 = (lngTwoLine_Top + lngTmp01)
770         lngTmp04 = (lngUnderline_Top + lngTmp01)
780         .country_code3_lbl.Top = lngTmp03
790         .country_code3_lbl_line.Top = lngTmp04
800         .country_code2_lbl.Top = lngTmp03
810         .country_code2_lbl_line.Top = lngTmp04
820         .country_name_lbl.Top = lngTmp02
830         .country_name_lbl_line.Top = lngTmp04
840         .country_iso_lbl.Top = lngTmp02
850         .country_iso_lbl_line.Top = lngTmp04
860         .country_active_lbl.Top = lngTmp02
870         .country_active_lbl_line.Top = lngTmp04
880         .country_na_lbl.Top = lngTmp02
890         .country_na_lbl_line.Top = lngTmp04
900         .Header_lbl2.Visible = True
910         .AsOf_lbl2.Visible = True
920         .Header_hline05.Visible = True
930       Else
940         .Header_lbl2.Visible = False
950         .AsOf_lbl2.Visible = False
960         .Header_hline05.Visible = False
970         .country_code3_lbl.Top = lngTwoLine_Top
980         .country_code3_lbl_line.Top = lngUnderline_Top
990         .country_code2_lbl.Top = lngTwoLine_Top
1000        .country_code2_lbl_line.Top = lngUnderline_Top
1010        .country_name_lbl.Top = lngOneLine_Top
1020        .country_name_lbl_line.Top = lngUnderline_Top
1030        .country_iso_lbl.Top = lngOneLine_Top
1040        .country_iso_lbl_line.Top = lngUnderline_Top
1050        .country_active_lbl.Top = lngOneLine_Top
1060        .country_active_lbl_line.Top = lngUnderline_Top
1070        .country_na_lbl.Top = lngOneLine_Top
1080        .country_na_lbl_line.Top = lngUnderline_Top
1090        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1100      End If
1110    End With

EXITP:
1120    Exit Sub

ERRH:
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub

Private Sub Report_Close()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1210    DoCmd.Restore

1220    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1230    If Forms(strCallingForm).Visible = False Then
1240      Forms(strCallingForm).Visible = True
1250      DoCmd.SelectObject acForm, strCallingForm, False
1260      DoCmd.Restore
1270    End If

EXITP:
1280    Exit Sub

ERRH:
1290    Forms(strCallingForm).Visible = True
1300    DoCmd.Restore
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub
