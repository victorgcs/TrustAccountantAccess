VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAccountContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAccountContacts"

'VGC 02/03/2013: CHANGES!

Private Const strCallingForm As String = "frmAccountContacts"

Private lngRptWidth As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim strSortNow As String
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         strSortNow = vbNullString
150         If IsNull(.OpenArgs) = False Then
160           If Trim(.OpenArgs) <> vbNullString Then
170             strTmp01 = Trim(.OpenArgs)
180             intPos01 = InStr(strTmp01, "~")
190             If intPos01 > 0 Then
200               strSortNow = Mid(strTmp01, (intPos01 + 1))
210             End If
220           End If
230         End If

240         lngRptWidth = .Width
250         lngPageHeaderSection_Height = .Section(acPageHeader).Height
260         lngOneLine_Top = .accountno_lbl.Top
270         lngUnderline_Top = .accountno_lbl_line.Top

280         Set frm = Forms(strCallingForm)
290         Select Case frm.chkPageOf
            Case True
300           .PageNumOf.Visible = True
310           .PageNum.Visible = False
320         Case False
330           .PageNum.Visible = True
340           .PageNumOf.Visible = False
350           .PageNumOf.ControlSource = vbNullString
360         End Select
370         If strSortNow = vbNullString Then  ' ** If this is heading to Word.
380           strSortNow = frm.frmAccountContacts_Sub.Form.SortNow_Get  ' ** Form Function: frmAccountContacts_Sub.
390         End If
400         Set frm = Nothing

410         SetupRptFlds strCallingForm, THIS_NAME  ' ** Module Function: modAccountContacts.

420         .OrderBy = strSortNow
430         .OrderByOn = True

440         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
450         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
460         .AsOf_lbl2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

470         If gstrReportCallingForm <> vbNullString Then
480           Forms(gstrReportCallingForm).Visible = False
490         End If

500       Else
510         Cancel = -1
520       End If
530     End With

540     If Cancel = -1 Then
550       Beep
560       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
570     End If

EXITP:
580     Set frm = Nothing
590     Exit Sub

ERRH:
600     If gstrReportCallingForm <> vbNullString Then
610       DoCmd.Restore
620       Forms(gstrReportCallingForm).Visible = True
630       DoCmd.SelectObject acForm, gstrReportCallingForm, False
640     End If
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

710     With Me
720       lngTmp01 = .AsOf_lbl2.Height
730       If .Page >= 2 Then
740         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
750         .accountno_lbl.Top = (lngOneLine_Top + lngTmp01)
760         .accountno_lbl_line.Top = (lngUnderline_Top + lngTmp01)
770         .Contact_Number_lbl.Top = (lngOneLine_Top + lngTmp01)
780         .Contact_Number_lbl_line.Top = (lngUnderline_Top + lngTmp01)
790         .Contact_Name_lbl.Top = (lngOneLine_Top + lngTmp01)
800         .Contact_Name_lbl_line.Top = (lngUnderline_Top + lngTmp01)
810         .Contact_Name_LastFirst_lbl.Top = (lngOneLine_Top + lngTmp01)
820         .Contact_Name_LastFirst_lbl_line.Top = (lngUnderline_Top + lngTmp01)
830         .Contact_Address1_lbl.Top = (lngOneLine_Top + lngTmp01)
840         .Contact_Address1_lbl_line.Top = (lngUnderline_Top + lngTmp01)
850         .Contact_Address2_lbl.Top = (lngOneLine_Top + lngTmp01)
860         .Contact_Address2_lbl_line.Top = (lngUnderline_Top + lngTmp01)
870         .Contact_Address_Combined_lbl.Top = (lngOneLine_Top + lngTmp01)
880         .Contact_Address_Combined_lbl_line.Top = (lngUnderline_Top + lngTmp01)
890         .Contact_City_lbl.Top = (lngOneLine_Top + lngTmp01)
900         .Contact_City_lbl_line.Top = (lngUnderline_Top + lngTmp01)
910         .Contact_State_lbl.Top = (lngOneLine_Top + lngTmp01)
920         .Contact_State_lbl_line.Top = (lngUnderline_Top + lngTmp01)
930         .Contact_Zip_lbl.Top = (lngOneLine_Top + lngTmp01)
940         .Contact_Zip_lbl_line.Top = (lngUnderline_Top + lngTmp01)
950         .Contact_Zip_Format_lbl.Top = (lngOneLine_Top + lngTmp01)
960         .Contact_Zip_Format_lbl_line.Top = (lngUnderline_Top + lngTmp01)
970         .Contact_CSZ_lbl.Top = (lngOneLine_Top + lngTmp01)
980         .Contact_CSZ_lbl_line.Top = (lngUnderline_Top + lngTmp01)
990         .Contact_CSZ_Format_lbl.Top = (lngOneLine_Top + lngTmp01)
1000        .Contact_CSZ_Format_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1010        .Contact_Phone1_lbl.Top = (lngOneLine_Top + lngTmp01)
1020        .Contact_Phone1_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1030        .Contact_Phone1_Format_lbl.Top = (lngOneLine_Top + lngTmp01)
1040        .Contact_Phone1_Format_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1050        .Contact_Phone2_lbl.Top = (lngOneLine_Top + lngTmp01)
1060        .Contact_Phone2_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1070        .Contact_Phone2_Format_lbl.Top = (lngOneLine_Top + lngTmp01)
1080        .Contact_Phone2_Format_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1090        .Contact_Phone_Combined_lbl.Top = (lngOneLine_Top + lngTmp01)
1100        .Contact_Phone_Combined_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1110        .Contact_Phone_Format_Combined_lbl.Top = (lngOneLine_Top + lngTmp01)
1120        .Contact_Phone_Format_Combined_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1130        .Contact_Fax_lbl.Top = (lngOneLine_Top + lngTmp01)
1140        .Contact_Fax_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1150        .Contact_Fax_Format_lbl.Top = (lngOneLine_Top + lngTmp01)
1160        .Contact_Fax_Format_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1170        .Contact_Email_lbl.Top = (lngOneLine_Top + lngTmp01)
1180        .Contact_Email_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1190        .Header_lbl2.Visible = True
1200        .AsOf_lbl2.Visible = True
1210        .Header_hline05.Visible = True
1220      Else
1230        .Header_lbl2.Visible = False
1240        .AsOf_lbl2.Visible = False
1250        .Header_hline05.Visible = False
1260        .accountno_lbl.Top = lngOneLine_Top
1270        .accountno_lbl_line.Top = lngUnderline_Top
1280        .Contact_Number_lbl.Top = lngOneLine_Top
1290        .Contact_Number_lbl_line.Top = lngUnderline_Top
1300        .Contact_Name_lbl.Top = lngOneLine_Top
1310        .Contact_Name_lbl_line.Top = lngUnderline_Top
1320        .Contact_Name_LastFirst_lbl.Top = lngOneLine_Top
1330        .Contact_Name_LastFirst_lbl_line.Top = lngUnderline_Top
1340        .Contact_Address1_lbl.Top = lngOneLine_Top
1350        .Contact_Address1_lbl_line.Top = lngUnderline_Top
1360        .Contact_Address2_lbl.Top = lngOneLine_Top
1370        .Contact_Address2_lbl_line.Top = lngUnderline_Top
1380        .Contact_Address_Combined_lbl.Top = lngOneLine_Top
1390        .Contact_Address_Combined_lbl_line.Top = lngUnderline_Top
1400        .Contact_City_lbl.Top = lngOneLine_Top
1410        .Contact_City_lbl_line.Top = lngUnderline_Top
1420        .Contact_State_lbl.Top = lngOneLine_Top
1430        .Contact_State_lbl_line.Top = lngUnderline_Top
1440        .Contact_Zip_lbl.Top = lngOneLine_Top
1450        .Contact_Zip_lbl_line.Top = lngUnderline_Top
1460        .Contact_Zip_Format_lbl.Top = lngOneLine_Top
1470        .Contact_Zip_Format_lbl_line.Top = lngUnderline_Top
1480        .Contact_CSZ_lbl.Top = lngOneLine_Top
1490        .Contact_CSZ_lbl_line.Top = lngUnderline_Top
1500        .Contact_CSZ_Format_lbl.Top = lngOneLine_Top
1510        .Contact_CSZ_Format_lbl_line.Top = lngUnderline_Top
1520        .Contact_Phone1_lbl.Top = lngOneLine_Top
1530        .Contact_Phone1_lbl_line.Top = lngUnderline_Top
1540        .Contact_Phone1_Format_lbl.Top = lngOneLine_Top
1550        .Contact_Phone1_Format_lbl_line.Top = lngUnderline_Top
1560        .Contact_Phone2_lbl.Top = lngOneLine_Top
1570        .Contact_Phone2_lbl_line.Top = lngUnderline_Top
1580        .Contact_Phone2_Format_lbl.Top = lngOneLine_Top
1590        .Contact_Phone2_Format_lbl_line.Top = lngUnderline_Top
1600        .Contact_Phone_Combined_lbl.Top = lngOneLine_Top
1610        .Contact_Phone_Combined_lbl_line.Top = lngUnderline_Top
1620        .Contact_Phone_Format_Combined_lbl.Top = lngOneLine_Top
1630        .Contact_Phone_Format_Combined_lbl_line.Top = lngUnderline_Top
1640        .Contact_Fax_lbl.Top = lngOneLine_Top
1650        .Contact_Fax_lbl_line.Top = lngUnderline_Top
1660        .Contact_Fax_Format_lbl.Top = lngOneLine_Top
1670        .Contact_Fax_Format_lbl_line.Top = lngUnderline_Top
1680        .Contact_Email_lbl.Top = lngOneLine_Top
1690        .Contact_Email_lbl_line.Top = lngUnderline_Top
1700        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1710      End If
1720    End With

EXITP:
1730    Exit Sub

ERRH:
1740    If gstrReportCallingForm <> vbNullString Then
1750      DoCmd.Restore
1760      Forms(gstrReportCallingForm).Visible = True
1770      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1780    End If
1790    Select Case ERR.Number
        Case Else
1800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1810    End Select
1820    Resume EXITP

End Sub

Private Sub Report_Close()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1910    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1920    DoCmd.Restore

1930    If gstrReportCallingForm <> vbNullString Then
1940      Forms(gstrReportCallingForm).Visible = True
1950    End If

EXITP:
1960    Exit Sub

ERRH:
1970    If gstrReportCallingForm <> vbNullString Then
1980      DoCmd.Restore
1990      Forms(gstrReportCallingForm).Visible = True
2000      DoCmd.SelectObject acForm, gstrReportCallingForm, False
2010    End If
2020    Select Case ERR.Number
        Case Else
2030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2040    End Select
2050    Resume EXITP

End Sub
