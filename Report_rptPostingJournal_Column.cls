VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptPostingJournal_Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptPostingJournal_Column"

'VGC 10/23/2017: CHANGES!

Private Const strCallingForm As String = "frmJournal_Columns"

Private lngTransCnt As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strTmp01 As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         lngTransCnt = CLng(.OpenArgs)

140         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
150         .PageNumOf.Visible = True
160         .PageNum.Visible = False

170         Select Case gblnAdmin
            Case True
180           If .RecordSource <> "qryJournal_Columns_53c" Then .RecordSource = "qryJournal_Columns_53c"
190         Case False
200           .RecordSource = "qryJournal_Columns_53d"
210         End Select

220         lngPageHeaderSection_Height = .Section(acPageHeader).Height
230         lngOneLine_Top = .transdate_lbl.Top
240         lngTwoLine_Top = .icash_lbl.Top
250         lngUnderline_Top = .transdate_lbl_line.Top

260         If gstrCo_Name = vbNullString Then
270           CoOptions_Read  ' ** Module Function: modStartupFuncs.
280         End If

290         .AsOf.Caption = Format(Now(), "mmmm dd, yyyy hh:nn:ss AM/PM")
300         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
310         .AsOf2.Caption = "As of " & Format(Now(), "mmmm dd, yyyy hh:nn:ss AM/PM")

320         .Caption = "Post_Journal_Including_Uncommitted_Entries_" & Format(Date, "yymmdd")

330         strTmp01 = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
340         .PrintedBy_lbl.Caption = "Printed by " & IIf(GetUserName = gstrDevUserName, "TAAdmin", strTmp01)  ' ** Module Procedure: modFileUtilities.

350       Else
360         Cancel = -1
370       End If
380     End With

390     If Cancel = -1 Then
400       Beep
410       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
420     End If

EXITP:
430     Exit Sub

ERRH:
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim varTmp00 As Variant

510     With Me
520   On Error Resume Next
530       varTmp00 = .HasUncom
540       If ERR.Number = 0 Then
550   On Error GoTo ERRH
560         Select Case .HasUncom
            Case True
570           .AsOf_lbl.Caption = "Including Unommitted Entries"
580           .AsOf_lbl.FontItalic = True
590           .AsOf_lbl.ForeColor = CLR_DKBLU
600           .transdate_u.Visible = True
610           .journaltype_u.Visible = True
620           .totdesc_u.Visible = True
630           .totdesc2_u.Visible = True
640           .icash_u.Visible = True
650           .pcash_u.Visible = True
660           .cost_u.Visible = True
670           .PrintCheck_display_u.Visible = True
680           .HasUncom_lbl1_u.Visible = True
690           .HasUncom_lbl2_u.Visible = True
700           .shareface_change.Visible = False
710           .shareface_change_u.Visible = True
720           .icash_tot.Visible = False
730           .icash_tot_u.Visible = True
740           .pcash_tot.Visible = False
750           .pcash_tot_u.Visible = True
760           .cost_tot.Visible = False
770           .cost_tot_u.Visible = True
780         Case False
790           .AsOf_lbl.Caption = "as of"
800           .AsOf_lbl.FontItalic = False
810           .AsOf_lbl.ForeColor = CLR_BLK
820           .transdate_u.Visible = False
830           .journaltype_u.Visible = False
840           .totdesc_u.Visible = False
850           .totdesc2_u.Visible = False
860           .icash_u.Visible = False
870           .pcash_u.Visible = False
880           .cost_u.Visible = False
890           .PrintCheck_display_u.Visible = False
900           .HasUncom_lbl1_u.Visible = False
910           .HasUncom_lbl2_u.Visible = False
920           .shareface_change.Visible = True
930           .shareface_change_u.Visible = False
940           .icash_tot.Visible = True
950           .icash_tot_u.Visible = False
960           .pcash_tot.Visible = True
970           .pcash_tot_u.Visible = False
980           .cost_tot.Visible = True
990           .cost_tot_u.Visible = False
1000        End Select
1010      Else
            ' ** No data!
1020  On Error GoTo ERRH
1030      End If
1040    End With

EXITP:
1050    Exit Sub

ERRH:
1060    Select Case ERR.Number
        Case Else
1070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1080    End Select
1090    Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

1110    With Me
1120      .SingleUser = gblnSingleUser
1130      lngTmp01 = .AsOf2.Height
1140      If .Page >= 2 Then
1150        .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
1160        .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
1170        .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1180        .totdesc_lbl.Top = (lngOneLine_Top + lngTmp01)
1190        .totdesc_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1200        .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1210        .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1220        .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1230        .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1240        .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
1250        .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1260        .Header_lbl2.Visible = True
1270        .AsOf2.Visible = True
1280        .Header_hline05.Visible = True
1290      Else
1300        .Header_lbl2.Visible = False
1310        .AsOf2.Visible = False
1320        .Header_hline05.Visible = False
1330        .transdate_lbl.Top = lngOneLine_Top
1340        .transdate_lbl_line.Top = lngUnderline_Top
1350        .totdesc_lbl.Top = lngOneLine_Top
1360        .totdesc_lbl_line.Top = lngUnderline_Top
1370        .icash_lbl.Top = lngTwoLine_Top
1380        .icash_lbl_line.Top = lngUnderline_Top
1390        .pcash_lbl.Top = lngTwoLine_Top
1400        .pcash_lbl_line.Top = lngUnderline_Top
1410        .cost_lbl.Top = lngOneLine_Top
1420        .cost_lbl_line.Top = lngUnderline_Top
1430        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1440      End If
1450    End With

EXITP:
1460    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim varTmp00 As Variant

1610    With Me
1620  On Error Resume Next
1630      varTmp00 = .HasUncomAcct
1640      If ERR.Number = 0 Then
1650  On Error GoTo ERRH
1660        Select Case .HasUncomAcct
            Case True
1670          .HasUncomAcct_lbl1_u.Visible = True
1680          .HasUncomAcct_lbl2_u.Visible = True
1690          .icash_subtot.Visible = False
1700          .icash_subtot_u.Visible = True
1710          .pcash_subtot.Visible = False
1720          .pcash_subtot_u.Visible = True
1730          .cost_subtot.Visible = False
1740          .cost_subtot_u.Visible = True
1750        Case False
1760          .HasUncomAcct_lbl1_u.Visible = False
1770          .HasUncomAcct_lbl2_u.Visible = False
1780          .icash_subtot.Visible = True
1790          .icash_subtot_u.Visible = False
1800          .pcash_subtot.Visible = True
1810          .pcash_subtot_u.Visible = False
1820          .cost_subtot.Visible = True
1830          .cost_subtot_u.Visible = False
1840        End Select
1850      Else
            ' ** No data!
1860  On Error GoTo ERRH
1870      End If
1880    End With

EXITP:
1890    Exit Sub

ERRH:
1900    Select Case ERR.Number
        Case Else
1910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920    End Select
1930    Resume EXITP

End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportFooter_Format"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strQryName As String, strFilter As String

2010    With Me
2020      If lngTransCnt = 0& Then
2030        strQryName = .RecordSource
2040        strFilter = .Filter
2050        Set dbs = CurrentDb
2060        Set qdf = dbs.QueryDefs(strQryName)
2070        Set rst1 = qdf.OpenRecordset
2080        rst1.Filter = strFilter
2090        Set rst2 = rst1.OpenRecordset
2100        rst2.MoveLast
2110        lngTransCnt = rst2.RecordCount
2120        rst2.Close
2130        rst1.Close
2140        dbs.Close
2150      End If
2160      .TransCount = lngTransCnt
2170    End With

EXITP:
2180    Set rst1 = Nothing
2190    Set rst2 = Nothing
2200    Set qdf = Nothing
2210    Set dbs = Nothing
2220    Exit Sub

ERRH:
2230    If gstrReportCallingForm <> vbNullString Then
2240      DoCmd.Restore
2250      Forms(gstrReportCallingForm).Visible = True
2260      DoCmd.SelectObject acForm, gstrReportCallingForm, False
2270    End If
2280    Select Case ERR.Number
        Case Else
2290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2300    End Select
2310    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

2410    With Me
2420      .NoData_lbl.BackColor = CLR_WHT
2430      .NoData_lbl.Visible = True
2440      .GroupHeader1.Visible = False
2450      .GroupFooter1.Visible = False
2460      .ReportFooter.Visible = False
2470      .ICash.Visible = False
2480      .PCash.Visible = False
2490      .Cost.Visible = False
2500      .PrintCheck_display.Visible = False
2510      .totdesc.Visible = False
2520      .TotDesc2.Visible = False
2530      .journaltype.Visible = False
2540      .journaltype_u.Visible = False
2550      .transdate.Visible = False
2560    End With

EXITP:
2570    Exit Sub

ERRH:
2580    Select Case ERR.Number
        Case Else
2590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2600    End Select
2610    Resume EXITP

End Sub

Private Sub Report_Close()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2710    If gstrReportCallingForm <> vbNullString Then
2720      gblnSetFocus = True
2730      DoCmd.SelectObject acForm, gstrReportCallingForm, False
2740      Forms(gstrReportCallingForm).TimerInterval = 250&
2750    End If

2760    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2770    Exit Sub

ERRH:
2780    Select Case ERR.Number
        Case Else
2790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2800    End Select
2810    Resume EXITP

End Sub
