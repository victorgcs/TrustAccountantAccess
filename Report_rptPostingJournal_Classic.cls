VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptPostingJournal_Classic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptPostingJournal_Classic"

'VGC 10/23/2017: CHANGES!

'shareface_change:
'=Sum(IIf([journaltype] In ("Purchase","Deposit"),[shareface],
'     IIf([journaltype] In ("Sold","Withdrawn"),-1*[shareface],
'     IIf(([journaltype])="liability",IIf([journaltypex]="Liability (-)",(-1*[shareface]),[shareface]),0))))

Private Const strCallingForm As String = "frmMenu_Post"

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

170         If gblnAdmin = False Then
180           .RecordSource = "qryPost_Journal_03f"  '#jnox
190         End If

200         lngPageHeaderSection_Height = .Section(acPageHeader).Height
210         lngOneLine_Top = .transdate_lbl.Top
220         lngTwoLine_Top = .icash_lbl.Top
230         lngUnderline_Top = .transdate_lbl_line.Top

240         .AsOf.Caption = Format(Now(), "mmmm dd, yyyy hh:nn:ss AM/PM")
250         .AsOf2.Caption = "As of " & Format(Now(), "mmmm dd, yyyy hh:nn:ss AM/PM")
260         If gstrCo_Name = vbNullString Then
270           CoOptions_Read  ' ** Module Function: modStartupFuncs.
280         End If
290         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

300         .Caption = "Post_Journal_" & Format(Date, "yymmdd")

310         strTmp01 = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
320         .PrintedBy_lbl.Caption = "Printed by " & IIf(GetUserName = gstrDevUserName, "TAAdmin", strTmp01)  ' ** Module Procedure: modFileUtilities.

330         If gstrReportCallingForm <> vbNullString Then
340           Forms(gstrReportCallingForm).Visible = False
350         End If

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
450     If gstrReportCallingForm <> vbNullString Then
460       DoCmd.Restore
470       Forms(gstrReportCallingForm).Visible = True
480       DoCmd.SelectObject acForm, gstrReportCallingForm, False
490     End If
500     Select Case ERR.Number
        Case Else
510       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
520     End Select
530     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

610     With Me
620       .SingleUser = gblnSingleUser
630       lngTmp01 = .AsOf2.Height
640       If .Page >= 2 Then
650         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
660         .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
670         .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
680         .totdesc_lbl.Top = (lngOneLine_Top + lngTmp01)
690         .totdesc_lbl_line.Top = (lngUnderline_Top + lngTmp01)
700         .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
710         .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
720         .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
730         .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
740         .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
750         .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
760         .Header_lbl2.Visible = True
770         .AsOf2.Visible = True
780         .Header_hline05.Visible = True
790       Else
800         .Header_lbl2.Visible = False
810         .AsOf2.Visible = False
820         .Header_hline05.Visible = False
830         .transdate_lbl.Top = lngOneLine_Top
840         .transdate_lbl_line.Top = lngUnderline_Top
850         .totdesc_lbl.Top = lngOneLine_Top
860         .totdesc_lbl_line.Top = lngUnderline_Top
870         .icash_lbl.Top = lngTwoLine_Top
880         .icash_lbl_line.Top = lngUnderline_Top
890         .pcash_lbl.Top = lngTwoLine_Top
900         .pcash_lbl_line.Top = lngUnderline_Top
910         .cost_lbl.Top = lngOneLine_Top
920         .cost_lbl_line.Top = lngUnderline_Top
930         .Section(acPageHeader).Height = lngPageHeaderSection_Height
940       End If
950     End With

EXITP:
960     Exit Sub

ERRH:
970     If gstrReportCallingForm <> vbNullString Then
980       DoCmd.Restore
990       Forms(gstrReportCallingForm).Visible = True
1000      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1010    End If
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportFooter_Format"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strQryName As String, strFilter As String

1110    With Me
1120      If lngTransCnt = 0& Then
1130        strQryName = .RecordSource
1140        strFilter = .Filter
1150        Set dbs = CurrentDb
1160        Set qdf = dbs.QueryDefs(strQryName)
1170        Set rst1 = qdf.OpenRecordset
1180        rst1.Filter = strFilter
1190        Set rst2 = rst1.OpenRecordset
1200        rst2.MoveLast
1210        lngTransCnt = rst2.RecordCount
1220        rst2.Close
1230        rst1.Close
1240        dbs.Close
1250      End If
1260      .TransCount = lngTransCnt
1270    End With

EXITP:
1280    Set rst1 = Nothing
1290    Set rst2 = Nothing
1300    Set qdf = Nothing
1310    Set dbs = Nothing
1320    Exit Sub

ERRH:
1330    If gstrReportCallingForm <> vbNullString Then
1340      DoCmd.Restore
1350      Forms(gstrReportCallingForm).Visible = True
1360      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1370    End If
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1510    With Me
1520      .NoData_lbl.BackColor = CLR_WHT
1530      .NoData_lbl.Visible = True
1540      .GroupHeader1.Visible = False
1550      .GroupFooter1.Visible = False
1560      .ReportFooter.Visible = False
1570      .ICash.Visible = False
1580      .PCash.Visible = False
1590      .Cost.Visible = False
1600      .PrintCheck_display.Visible = False
1610      .totdesc.Visible = False
1620      .TotDesc2.Visible = False
1630      .journaltypex.Visible = False
1640      .transdate.Visible = False
1650    End With

EXITP:
1660    Exit Sub

ERRH:
1670    If gstrReportCallingForm <> vbNullString Then
1680      DoCmd.Restore
1690      Forms(gstrReportCallingForm).Visible = True
1700      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1710    End If
1720    Select Case ERR.Number
        Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub

Private Sub Report_Close()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1810    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1820    DoCmd.Restore

1830    If gstrReportCallingForm <> vbNullString Then
1840      Forms(gstrReportCallingForm).Visible = True
1850    End If

EXITP:
1860    Exit Sub

ERRH:
1870    If gstrReportCallingForm <> vbNullString Then
1880      DoCmd.Restore
1890      Forms(gstrReportCallingForm).Visible = True
1900      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1910    End If
1920    Select Case ERR.Number
        Case Else
1930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1940    End Select
1950    Resume EXITP

End Sub
