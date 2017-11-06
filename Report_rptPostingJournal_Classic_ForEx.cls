VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptPostingJournal_Classic_ForEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptPostingJournal_Classic_ForEx"

'VGC 10/23/2017: CHANGES!

'shareface_change:
'=Sum(IIf([journaltype] In ("Purchase","Deposit"),[shareface],
'     IIf([journaltype] In ("Sold","Withdrawn"),-1*[shareface],
'     IIf(([journaltype])="liability",IIf([journaltypex]="Liability (-)",(-1*[shareface]),[shareface]),0))))

Private Const strCallingForm As String = "frmMenu_Post"

Private lngTransCnt As Long
Private lngTpp As Long, lngPageHeaderSection_Height As Long
Private lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
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

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
200         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

210         lngPageHeaderSection_Height = .Section(acPageHeader).Height
220         lngOneLine_Top = .transdate_lbl.Top
230         lngTwoLine_Top = .icash_usd_lbl.Top
240         lngUnderline_Top = .transdate_lbl_line.Top

250         .AsOf.Caption = Format(Now(), "mmmm dd, yyyy hh:nn:ss AM/PM")
260         .AsOf2.Caption = "As of " & Format(Now(), "mmmm dd, yyyy hh:nn:ss AM/PM")
270         If gstrCo_Name = vbNullString Then
280           CoOptions_Read  ' ** Module Function: modStartupFuncs.
290         End If
300         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

310         .Caption = "Post_Journal_" & Format(Date, "yymmdd")

320         strTmp01 = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
330         .PrintedBy_lbl.Caption = "Printed by " & IIf(GetUserName = gstrDevUserName, "TAAdmin", strTmp01)  ' ** Module Procedure: modFileUtilities.

340         If gstrReportCallingForm <> vbNullString Then
350           Forms(gstrReportCallingForm).Visible = False
360         End If

370       Else
380         Cancel = -1
390       End If
400     End With

410     If Cancel = -1 Then
420       Beep
430       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
440     End If

EXITP:
450     Exit Sub

ERRH:
460     If gstrReportCallingForm <> vbNullString Then
470       DoCmd.Restore
480       Forms(gstrReportCallingForm).Visible = True
490       DoCmd.SelectObject acForm, gstrReportCallingForm, False
500     End If
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

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
700         .curr_code_lbl.Top = ((lngOneLine_Top + lngTmp01) + lngTpp)
710         .curr_code_lbl_line.Top = (lngUnderline_Top + lngTmp01)
720         .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
730         .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
740         .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
750         .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
760         .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
770         .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
780         .icash_usd_lbl.Top = (lngTwoLine_Top + lngTmp01)
790         .icash_usd_lbl_line.Top = (lngUnderline_Top + lngTmp01)
800         .pcash_usd_lbl.Top = (lngTwoLine_Top + lngTmp01)
810         .pcash_usd_lbl_line.Top = (lngUnderline_Top + lngTmp01)
820         .cost_usd_lbl.Top = (lngOneLine_Top + lngTmp01)
830         .cost_usd_lbl_line.Top = (lngUnderline_Top + lngTmp01)
840         .Header_lbl2.Visible = True
850         .AsOf2.Visible = True
860         .Header_hline05.Visible = True
870       Else
880         .Header_lbl2.Visible = False
890         .AsOf2.Visible = False
900         .Header_hline05.Visible = False
910         .transdate_lbl.Top = lngOneLine_Top
920         .transdate_lbl_line.Top = lngUnderline_Top
930         .totdesc_lbl.Top = lngOneLine_Top
940         .totdesc_lbl_line.Top = lngUnderline_Top
950         .curr_code_lbl.Top = (lngOneLine_Top + lngTpp)
960         .curr_code_lbl_line.Top = lngUnderline_Top
970         .icash_lbl.Top = lngTwoLine_Top
980         .icash_lbl_line.Top = lngUnderline_Top
990         .pcash_lbl.Top = lngTwoLine_Top
1000        .pcash_lbl_line.Top = lngUnderline_Top
1010        .cost_lbl.Top = lngOneLine_Top
1020        .cost_lbl_line.Top = lngUnderline_Top
1030        .icash_usd_lbl.Top = lngTwoLine_Top
1040        .icash_usd_lbl_line.Top = lngUnderline_Top
1050        .pcash_usd_lbl.Top = lngTwoLine_Top
1060        .pcash_usd_lbl_line.Top = lngUnderline_Top
1070        .cost_usd_lbl.Top = lngOneLine_Top
1080        .cost_usd_lbl_line.Top = lngUnderline_Top
1090        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1100      End If
1110    End With

EXITP:
1120    Exit Sub

ERRH:
1130    If gstrReportCallingForm <> vbNullString Then
1140      DoCmd.Restore
1150      Forms(gstrReportCallingForm).Visible = True
1160      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1170    End If
1180    Select Case ERR.Number
        Case Else
1190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1200    End Select
1210    Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

        Dim intLen As Integer
        Dim strTmp01 As String, strTmp03 As String, intTmp04 As Integer

1310    With Me
1320      If .curr_code = "USD" Then
1330        strTmp01 = "Currency"
1340        intTmp04 = 2
1350      Else
1360        strTmp03 = .currsym_symbol
1370        If .curr_decimal = "." Then
1380          intTmp04 = 4
1390        Else
1400          intTmp04 = .curr_decimal
1410        End If
1420        intLen = Len(strTmp03)
1430        Select Case intLen
            Case 1
              'strTmp03 = strTmp03
1440        Case 2
1450          strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2)
1460        Case 3
1470          strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3)
1480        Case 4
1490          strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & "\" & Mid(strTmp03, 4)
1500        Case 5
1510          strTmp03 = "\" & Left(strTmp03, 1) & "\" & Mid(strTmp03, 2, 1) & "\" & Mid(strTmp03, 3, 1) & _
                "\" & Mid(strTmp03, 4, 1) & "\" & Mid(strTmp03, 5)
1520        End Select
1530        Select Case intTmp04
            Case 0
1540          strTmp01 = strTmp03 & "#,##0;(" & strTmp03 & "#,##0);" & strTmp03 & "0;" & strTmp03 & "0"
1550        Case 1  ' ** None have this currently.
1560          strTmp01 = strTmp03 & "#,##0.0;(" & strTmp03 & "#,##0.0);" & strTmp03 & "0.0;" & strTmp03 & "0.0"
1570        Case 2
1580          strTmp01 = strTmp03 & "#,##0.00;(" & strTmp03 & "#,##0.00);" & strTmp03 & "0.00;" & strTmp03 & "0.00"
1590        Case 3
1600          strTmp01 = strTmp03 & "#,##0.000;(" & strTmp03 & "#,##0.000);" & strTmp03 & "0.000;" & strTmp03 & "0.000"
1610        End Select
1620      End If
1630      .ICash.Format = strTmp01
1640      .ICash.DecimalPlaces = intTmp04
1650      .PCash.Format = strTmp01
1660      .PCash.DecimalPlaces = intTmp04
1670      .Cost.Format = strTmp01
1680      .Cost.DecimalPlaces = intTmp04
1690    End With

EXITP:
1700    Exit Sub

ERRH:
1710    If gstrReportCallingForm <> vbNullString Then
1720      DoCmd.Restore
1730      Forms(gstrReportCallingForm).Visible = True
1740      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1750    End If
1760    Select Case ERR.Number
        Case Else
1770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1780    End Select
1790    Resume EXITP

End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportFooter_Format"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strQryName As String, strFilter As String

1810    With Me
1820      If lngTransCnt = 0& Then
1830        strQryName = .RecordSource
1840        strFilter = .Filter
1850        Set dbs = CurrentDb
1860        Set qdf = dbs.QueryDefs(strQryName)
1870        Set rst1 = qdf.OpenRecordset
1880        rst1.Filter = strFilter
1890        Set rst2 = rst1.OpenRecordset
1900        rst2.MoveLast
1910        lngTransCnt = rst2.RecordCount
1920        rst2.Close
1930        rst1.Close
1940        dbs.Close
1950      End If
1960      .TransCount = lngTransCnt
1970    End With

EXITP:
1980    Set rst1 = Nothing
1990    Set rst2 = Nothing
2000    Set qdf = Nothing
2010    Set dbs = Nothing
2020    Exit Sub

ERRH:
2030    If gstrReportCallingForm <> vbNullString Then
2040      DoCmd.Restore
2050      Forms(gstrReportCallingForm).Visible = True
2060      DoCmd.SelectObject acForm, gstrReportCallingForm, False
2070    End If
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

2210    With Me
2220      .NoData_lbl.BackColor = CLR_WHT
2230      .NoData_lbl.Visible = True
2240      .GroupHeader1.Visible = False
2250      .GroupFooter1.Visible = False
2260      .ReportFooter.Visible = False
2270      .curr_code.Visible = False
2280      .ICash.Visible = False
2290      .PCash.Visible = False
2300      .Cost.Visible = False
2310      .icash_usd.Visible = False
2320      .pcash_usd.Visible = False
2330      .cost_usd.Visible = False
2340      .PrintCheck_display.Visible = False
2350      .totdesc.Visible = False
2360      .TotDesc2.Visible = False
2370      .journaltypex.Visible = False
2380      .transdate.Visible = False
2390    End With

EXITP:
2400    Exit Sub

ERRH:
2410    If gstrReportCallingForm <> vbNullString Then
2420      DoCmd.Restore
2430      Forms(gstrReportCallingForm).Visible = True
2440      DoCmd.SelectObject acForm, gstrReportCallingForm, False
2450    End If
2460    Select Case ERR.Number
        Case Else
2470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2480    End Select
2490    Resume EXITP

End Sub

Private Sub Report_Close()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2510    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

2520    DoCmd.Restore

2530    If gstrReportCallingForm <> vbNullString Then
2540      Forms(gstrReportCallingForm).Visible = True
2550    End If

EXITP:
2560    Exit Sub

ERRH:
2570    If gstrReportCallingForm <> vbNullString Then
2580      DoCmd.Restore
2590      Forms(gstrReportCallingForm).Visible = True
2600      DoCmd.SelectObject acForm, gstrReportCallingForm, False
2610    End If
2620    Select Case ERR.Number
        Case Else
2630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2640    End Select
2650    Resume EXITP

End Sub
