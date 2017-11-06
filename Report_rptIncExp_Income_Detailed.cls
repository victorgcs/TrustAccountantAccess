VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptIncExp_Income_Detailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptIncExp_Income_Detailed"

'VGC 01/14/2017: CHANGES!

'totdesc =
'Trim(
'  IIf(
'      IsNull([RecurringItem]),"",
'      IIf(
'          [journaltype]="Received",[RecurringItem],
'          IIf(
'              [journaltype]="Paid",[RecurringItem],[RecurringItem]
'             )
'         )
'     ) &
'  IIf(
'      [assetno] Is Not Null,
'      IIf(
'          [assetdate] Is Not Null,Format([assetdate],"mm/dd/yyyy") & " "
'         ) &
'      IIf(
'          [shareface]-CLng([shareface])=0,Format([shareface],"#,##0 "),Format([shareface],"#,##0.000 ")
'         ) & " " & CStr([Description]) &
'      IIf(
'          [rate]>0," " & Format([rate],"#,##0.000%")),""
'         ) & IIf([due] Is Not Null,"  Due " & Format([due],"mm/dd/yyyy")
'     ) & "  " & [Jcomment]
')

'Original:
'=Trim(IIf(IsNull([RecurringItem]),"",IIf([journaltype]="Received",[RecurringItem],IIf([journaltype]="Paid",[RecurringItem],[RecurringItem]))) & IIf([assetno] Is Not Null,IIf([assetdate] Is Not Null,Format([assetdate],"mm/dd/yyyy") & " ") & IIf([shareface]-CLng([shareface])=0,Format([shareface],"#,##0 "),Format([shareface],"#,##0.000 ")) & " " & CStr([Description]) & IIf([rate]>0," " & Format([rate],"#,##0.000%")),"") & IIf([due] Is Not Null,"  Due " & Format([due],"mm/dd/yyyy")) & "  " & [Jcomment])
'New:
'=Trim(IIf(IsNull([RecurringItem])=True,"",IIf([journaltype]="Received",[RecurringItem],IIf([journaltype]="Paid",[RecurringItem],[RecurringItem]))) & " " & IIf(IsNull([assetno])=True,"",IIf(IsNull([assetdate])=True,"",Format([assetdate],"mm/dd/yyyy") & " ") & IIf(IsNull([shareface])=True,"",IIf([shareface]=0,"",IIf([shareface]-CLng([shareface])=0,Format([shareface],"#,##0 "),Format([shareface],"#,##0.000 ")))) & " " & IIf(IsNull([RecurringItem])=True,CStr([Description]),IIf([RecurringItem]=[Description],"",CStr([Description]))) & IIf([rate]>0," " & Format([rate],"#,##0.000%"),"") & IIf(IsNull([due])=True,"","  Due " & Format([due],"mm/dd/yyyy")) & "  " & Nz([Jcomment],"")))

Private Const strCallingForm As String = "frmRpt_IncomeExpense"

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngAccountNo_Top As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim strAccountNo As String
        Dim strFilter As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Set frm = Forms(strCallingForm)

150         Select Case frm.chkPageOf
            Case True
160           .PageNumOf.Visible = True
170           .PageNum.Visible = False
180         Case False
190           .PageNum.Visible = True
200           .PageNumOf.Visible = False
210           .PageNumOf.ControlSource = vbNullString
220         End Select

230         lngPageHeaderSection_Height = .Section(acPageHeader).Height
240         lngAccountNo_Top = .accountno.Top
250         lngOneLine_Top = .transdate_lbl.Top
260         lngTwoLine_Top = .icash_lbl.Top
270         lngUnderline_Top = .transdate_lbl_line.Top

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
280         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

290         .AsOf_lbl.Caption = "From " & Format(CDate(gdatStartDate), "mm/dd/yyyy") & _
              " To " & Format(CDate(gdatEndDate), "mm/dd/yyyy")
300         .AsOf_lbl2.Caption = "From " & Format(CDate(gdatStartDate), "mm/dd/yyyy") & " To " & _
              Format(CDate(gdatEndDate), "mm/dd/yyyy")
310         .CoInfoBlock.Caption = gstrCo_InfoBlock

320         Select Case frm.opgAccountNumber
            Case frm.opgAccountNumber_optSpecified.OptionValue
330           strAccountNo = frm.cmbAccounts
340           .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
350         Case frm.opgAccountNumber_optAll.OptionValue
360           strAccountNo = "All"
370           .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
380         End Select

390         .Caption = "Income_Detailed_" & strAccountNo & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & _
              Format(gdatEndDate, "yymmdd")

400         Select Case frm.opgUnspecified
            Case frm.opgUnspecified_optInclude.OptionValue
410           Select Case frm.opgPrincipalCash
              Case frm.opgPrincipalCash_optInclude.OptionValue
420             strFilter = "(([transdate] Between #" & Format(CDate(gdatStartDate), "mm/dd/yyyy") & "# And " & _
                  "#" & Format(CDate(gdatEndDate), "mm/dd/yyyy") & "#) And [revcode_TYPE] = 1)"
430           Case frm.opgPrincipalCash_optExclude.OptionValue
440             strFilter = "(([transdate] Between #" & Format(CDate(gdatStartDate), "mm/dd/yyyy") & "# And " & _
                  "#" & Format(CDate(gdatEndDate), "mm/dd/yyyy") & "#) And [revcode_TYPE] = 1) And " & _
                  "([icash] <> 0)"
450           End Select
460         Case frm.opgUnspecified_optExclude.OptionValue
470           Select Case frm.opgPrincipalCash
              Case frm.opgPrincipalCash_optInclude.OptionValue
480             strFilter = "(([transdate] Between #" & Format(CDate(gdatStartDate), "mm/dd/yyyy") & "# And " & _
                  "#" & Format(CDate(gdatEndDate), "mm/dd/yyyy") & "#) And [revcode_TYPE] = 1) And " & _
                  "([revcode_DESC] <> 'Unspecified Income' And [revcode_DESC] <> 'Unspecified Expense')"
490           Case frm.opgPrincipalCash_optExclude.OptionValue
500             strFilter = "(([transdate] Between #" & Format(CDate(gdatStartDate), "mm/dd/yyyy") & "# And " & _
                  "#" & Format(CDate(gdatEndDate), "mm/dd/yyyy") & "#) And [revcode_TYPE] = 1) And " & _
                  "([revcode_DESC] <> 'Unspecified Income' And [revcode_DESC] <> 'Unspecified Expense') And " & _
                  "([icash] <> 0)"
510           End Select
520         Case frm.opgUnspecified_optOnly.OptionValue
530           Select Case frm.opgPrincipalCash
              Case frm.opgPrincipalCash_optInclude.OptionValue
540             strFilter = "(([transdate] Between #" & Format(CDate(gdatStartDate), "mm/dd/yyyy") & "# And " & _
                  "#" & Format(CDate(gdatEndDate), "mm/dd/yyyy") & "#) And [revcode_TYPE] = 1)" & " And " & _
                  "[revcode_DESC] = 'Unspecified Income'"
550           Case frm.opgPrincipalCash_optExclude.OptionValue
560             strFilter = "(([transdate] Between #" & Format(CDate(gdatStartDate), "mm/dd/yyyy") & "# And " & _
                  "#" & Format(CDate(gdatEndDate), "mm/dd/yyyy") & "#) And [revcode_TYPE] = 1) And " & _
                  "([icash] <> 0)" & " And [revcode_DESC] = 'Unspecified Income'"
570           End Select

580         End Select

            ' ** Get all accounts if we were passed ALL.
590         If gstrAccountNo <> "ALL" Then
600           If Trim(strFilter) <> vbNullString Then strFilter = strFilter & " And "
610           strFilter = strFilter & "[accountno] = '" & gstrAccountNo & "'"
620         End If

630         If frm.chkSweepOnly = True Then
640           If Trim(strFilter) <> vbNullString Then strFilter = strFilter & " And "
650           strFilter = strFilter & "[assetno] <= 1"
660           .AsOf.Caption = .AsOf.Caption & " - Sweep Asset Only"
670         End If

680         If frm.chkDontCombineMulti = True Then
690           .MultiLotMsg_lbl.Visible = True
700         End If

710         .RptSubtype_lbl.Caption = gstrCrtRpt_Ordinal  ' ** Borrowing this variable from Court Reports.

720         .Filter = strFilter
730         .FilterOn = True

740         .OrderBy = "revcode_SORTORDER"
750         .OrderByOn = True

760       Else
770         Cancel = -1
780       End If
790     End With

800     If Cancel = -1 Then
810       Beep
820       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
830     End If

EXITP:
840     Set frm = Nothing
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

910     With Me
920       lngTmp01 = .AsOf_lbl2.Height
930       If .Page >= 2 Then
940         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
950         .accountno.Top = ((lngAccountNo_Top + lngTmp01) + (4& * lngTpp))
960         .shortname.Top = .accountno.Top
970         .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
980         .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
990         .totdesc_lbl.Top = (lngOneLine_Top + lngTmp01)
1000        .totdesc_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1010        .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
1020        .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1030        .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1040        .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1050        .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1060        .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1070        .Header_lbl2.Visible = True
1080        .AsOf_lbl2.Visible = True
1090        .Header_hline05.Visible = True
1100      Else
1110        .accountno.Top = lngAccountNo_Top
1120        .shortname.Top = .accountno.Top
1130        .transdate_lbl.Top = lngOneLine_Top
1140        .transdate_lbl_line.Top = lngUnderline_Top
1150        .totdesc_lbl.Top = lngOneLine_Top
1160        .totdesc_lbl_line.Top = lngUnderline_Top
1170        .cost_lbl.Top = lngOneLine_Top
1180        .cost_lbl_line.Top = lngUnderline_Top
1190        .icash_lbl.Top = lngTwoLine_Top
1200        .icash_lbl_line.Top = lngUnderline_Top
1210        .pcash_lbl.Top = lngTwoLine_Top
1220        .pcash_lbl_line.Top = lngUnderline_Top
1230        .Header_lbl2.Visible = False
1240        .AsOf_lbl2.Visible = False
1250        .Header_hline05.Visible = False
1260        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1270      End If
1280    End With

EXITP:
1290    Exit Sub

ERRH:
1300    Select Case ERR.Number
        Case Else
1310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1320    End Select
1330    Resume EXITP

End Sub

Private Sub GroupHeader4_Print(Cancel As Integer, PrintCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader4_Print"

        Dim strTmp01 As String

1410    With Me
1420  On Error Resume Next
1430      strTmp01 = .totdesc
1440      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1450  On Error GoTo ERRH
1460        Report_NoData Cancel  ' ** Though this may be triggered anyway.
1470      Else
1480  On Error GoTo ERRH
1490      End If
1500    End With

EXITP:
1510    Exit Sub

ERRH:
1520    Select Case ERR.Number
        Case Else
1530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1540    End Select
1550    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1610    With Me
1620  On Error Resume Next
1630      strTmp01 = .totdesc
1640      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1650  On Error GoTo ERRH
            ' ** Do nothing.
1660      Else
1670  On Error GoTo ERRH
1680      End If
1690    End With

EXITP:
1700    Exit Sub

ERRH:
1710    Select Case ERR.Number
        Case Else
1720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1730    End Select
1740    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1810    With Me
1820      .NoData_lbl.BackColor = CLR_WHT
1830      .NoData_lbl.Visible = True
1840      .GroupHeader4.Visible = False
1850      .GroupFooter1.Visible = False
1860      .GroupFooter4.Visible = False
1870      .transdate.Visible = False
1880      .journaltype.Visible = False
1890      .totdesc.Visible = False
1900      .ICash.Visible = False
1910      .PCash.Visible = False
1920      .Cost.Visible = False
1930      .icash_tot.Visible = False
1940      .icash_tot_line.Visible = False
1950      .pcash_tot.Visible = False
1960      .pcash_tot_line.Visible = False
1970      .cost_tot.Visible = False
1980      .cost_tot_line.Visible = False
1990    End With

EXITP:
2000    Exit Sub

ERRH:
2010    Select Case ERR.Number
        Case Else
2020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2030    End Select
2040    Resume EXITP

End Sub

Private Sub Report_Close()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2110    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub
