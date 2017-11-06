VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_00A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_00A"

'VGC 03/11/2016: CHANGES!

' ** No data produces a report with zeroes.

Private Const strCallingForm As String = "frmRpt_CourtReports_CA"

' ** cmbAccounts combo box constants:
Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname

Private lngDevLeft As Long, lngDevCnt As Long

Private frm As Access.Form
Private lngTotLeft_1_Orig As Long, lngTotLeft_1_Moved As Long
Private lngTotLeftLine_1_Orig As Long, lngTotLeftLine_1_Moved As Long
Private lngTotLeft_2_Orig As Long, lngTotLeft_2_Moved As Long
Private lngTotLeftLine_2_Orig As Long, lngTotLeftLine_2_Moved As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strFilter As String
        Dim lngLegalNameMove_Height As Long
        Dim intLen As Integer
        Dim varFootnote As Variant

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Procedure: modFileUtilities.

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

            ' ** Because the PageFooterSection does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
230         If frm.chkIncludeFootnote = True Then
240           varFootnote = frm.CourtReports_Footnote
250           If IsNull(varFootnote) = False Then
260             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
270               intLen = Len(Trim(varFootnote))
280               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
290                 .Section("PageFooterSection").Height = 480&
300               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
310                 .Section("PageFooterSection").Height = 690&
320               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
330                 .Section("PageFooterSection").Height = 900&
340               End If
350             End If
360           End If
370         End If

380         lngLegalNameMove_Height = (.AcctName_Short.Height - 120&)

            ' ** Set the account name heading.
390         If gblnLegalName = True Then
400           If IsNull(frm.cmbAccounts.Column(CBX_A_LEGAL)) = True Then
                ' ** Use shortname.
410           Else
420             If Trim(frm.cmbAccounts.Column(CBX_A_LEGAL)) = vbNullString Then
                  ' ** Use shortname.
430             Else
440               .Section(acPageHeader).Height = (.Section(acPageHeader).Height + lngLegalNameMove_Height)
450               .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
460               .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
470               .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
480               .AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height
490               .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
500               .CaseNum.Top = .CaseNum.Top + lngLegalNameMove_Height
510               .CaseNum_lbl.Top = .CaseNum_lbl.Top + lngLegalNameMove_Height
520               .AcctName_Legal.Visible = True
530               .AcctName_Short.Visible = False
540             End If
550           End If
560         End If

            ' ** Set the date range heading.
570         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the ordinal and version heading.
580         .AccountHead_lbl.Caption = frm.Ordinal & " And " & _
              frm.Version & " Account"

            ' ** Set the bottom-left report version footer.
590         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
600         strFilter = "[date] >= #" & Format(gdatStartDate, _
              "mm/dd/yyyy") & "# AND [date] < #" & Format(gdatEndDate, "mm/dd/yyyy") & "# + 1"
610         strFilter = strFilter & " AND [accountno] = '" & gstrAccountNo & "'"
620         .Filter = strFilter
630         .FilterOn = True

            ' ** Turn on the group footer for the reveune/expense code lines.
640         If gblnUseReveuneExpenseCodes = True Then
650           .GroupFooter3.Visible = True
660           .GroupHeader3.Visible = True
670         Else
680           .GroupFooter3.Visible = False
690           .GroupHeader3.Visible = False
700         End If

710         lngTotLeft_1_Orig = .txtReportCategorySum.Left
720         lngTotLeft_1_Moved = lngTotLeft_1_Orig * 0.8   ' ** Move left about 20% of the text box's width.
730         lngTotLeftLine_1_Orig = .txtReportCategorySum_line.Left
740         lngTotLeftLine_1_Moved = lngTotLeft_1_Moved + (.txtReportCategorySum.Width - .txtReportCategorySum_line.Width)

750         lngTotLeft_2_Orig = .Amount_Subtot2.Left
760         lngTotLeft_2_Moved = lngTotLeft_2_Orig * 0.8   ' ** Move left about 20% of the text box's width.
770         lngTotLeftLine_2_Orig = .Amount_Subtot2_line.Left
780         lngTotLeftLine_2_Moved = lngTotLeft_2_Moved + (.Amount_Subtot2.Width - .Amount_Subtot2_line.Width)

790         .RptNumPOH_Beg.ControlSource = "=" & CStr(CRPT_ON_HAND_BEG)
800         .RptNumPOH_End.ControlSource = "=" & CStr(CRPT_ON_HAND_END)
810         .PropOnHand_Beg.ControlSource = "=" & CStr(gdblCrtRpt_CA_POHBeg)
820         .PropOnHand_End.ControlSource = "=" & CStr(gdblCrtRpt_CA_POHEnd)

830         .RptNumNonCash_Beg.ControlSource = "=" & CStr(CRPT_NON_CASH_BEG)
840         .RptNumNonCash_End.ControlSource = "=" & CStr(CRPT_NON_CASH_END)
850         .CashAssets_Beg.ControlSource = "=" & CStr(gdblCrtRpt_CA_COHBeg)
860         .CashAssets_End.ControlSource = "=" & CStr(gdblCrtRpt_CA_COHEnd)

870         .RptNumInvestInfo.ControlSource = "=" & CStr(CRPT_INVEST_INFO)
880         .RptNumInvestChange.ControlSource = "=" & CStr(CRPT_CHANGES)
890         .InvestInfo.ControlSource = "=" & CStr(gdblCrtRpt_CA_InvestInfo)
900         .InvestChange.ControlSource = "=" & CStr(gdblCrtRpt_CA_InvestChange)

910         lngDevLeft = .Dev1.Left

920         .Caption = "CourtReport_CA_Summary_Grouped_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

930       Else
940         Cancel = -1
950       End If
960     End With

970     If Cancel = -1 Then
980       Beep
990       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1000    End If

EXITP:
1010    Exit Sub

ERRH:
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub Report_Activate()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Activate"

1110    TAReports_SetZero True  ' ** Module Function: modReportFunctions.

EXITP:
1120    Exit Sub

ERRH:
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

1210    With Me
1220      If .ReportDivisionTitle = "CREDITS" Then
1230        .Section("GroupFooter1").ForceNewPage = 2
1240      Else
1250        .Section("GroupFooter1").ForceNewPage = 0
1260      End If
1270    End With

        ' ** The ForceNewPage property uses the following settings.
        ' **   0  None            The current section (the section for which you're setting the property) is printed on the
        ' **                      current page. (Default)
        ' **   1  Before Section  The current section is printed at the top of a new page.
        ' **   2  After Section   The section immediately following the current section is printed at the top of a new page.
        ' **   3  Before & After  The current section is printed at the top of a new page, and the next section is printed
        ' **                      at the top of a new page.

EXITP:
1280    Exit Sub

ERRH:
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub

Private Sub GroupHeader5_Format(Cancel As Integer, FormatCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader5_Format"

1410    With Me
1420      Select Case .ReportNumber
          Case CRPT_RECEIPTS
1430        If .txtReportDivision = CRPT_DIV_CHARGES Then
1440          Cancel = 0
1450          If .revcode_DESC = "Dummy entry" Then
1460            Cancel = -1  ' ** Do not print dummy entries.
1470          End If
1480        Else
1490          Cancel = -1
1500        End If
1510      Case CRPT_DISBURSEMENTS
1520        If .txtReportDivision = CRPT_DIV_CREDITS Then ' And .txtReportGroup = 30 Then  ' ** CALIFORNIA HARD-CODED NUMBERS!
1530          Cancel = 0
1540          If .revcode_DESC = "Dummy entry" Then
1550            Cancel = -1  ' ** Do not print dummy entries.
1560          End If
1570        Else
1580          Cancel = -1
1590        End If
1600      Case Else
1610        Cancel = -1
1620      End Select
1630    End With

EXITP:
1640    Exit Sub

ERRH:
1650    Select Case ERR.Number
        Case Else
1660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1670    End Select
1680    Resume EXITP

End Sub

Private Sub GroupHeader3_Format(Cancel As Integer, FormatCount As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader3_Format"

1710    With Me
1720      Select Case .ReportNumber
          Case CRPT_RECEIPTS
1730        If .txtReportDivision = CRPT_DIV_CHARGES Then
1740          .txtReportCategorySum.Visible = False
1750          .GroupFooter3.Visible = True
1760        Else
1770          .txtReportCategorySum.Visible = True
1780          .GroupFooter3.Visible = False
1790        End If
1800      Case CRPT_DISBURSEMENTS
1810        If .txtReportDivision = CRPT_DIV_CREDITS Then ' And .txtReportGroup = 30 Then  ' ** CALIFORNIA HARD-CODED NUMBERS!
1820          .txtReportCategorySum.Visible = False
1830          .GroupFooter3.Visible = True
1840        Else
1850          .txtReportCategorySum.Visible = True
1860          .GroupFooter3.Visible = False
1870        End If
1880      Case Else
1890        .txtReportCategorySum.Visible = True
1900        .GroupFooter3.Visible = False
1910      End Select
1920      If .ReportNumber = CRPT_ON_HAND_BEGL Or .ReportNumber = CRPT_ON_HAND_ENDL Then
1930        .txtReportCategorySum.Visible = False
1940      Else
1950        .txtReportCategorySum.Visible = True
1960      End If
1970      If (.ReportNumber = CRPT_CASH_BEG Or .ReportNumber = CRPT_NON_CASH_BEG) Or _
              (.ReportNumber = CRPT_CASH_END Or .ReportNumber = CRPT_NON_CASH_END) Then
1980        .ReportSchedule.Visible = False
1990        .txtReportCategorySum.Left = lngTotLeft_1_Moved
2000        .Dev1.Left = lngDevLeft - (lngTotLeft_1_Orig - lngTotLeft_1_Moved)
2010        .txtReportCategorySum_line.Left = lngTotLeftLine_1_Orig
2020        .txtReportCategorySum_line.Visible = False
2030      ElseIf .ReportNumber = CRPT_ON_HAND_BEG Or .ReportNumber = CRPT_ON_HAND_END Then
2040        .ReportSchedule.Visible = False
2050        .txtReportCategorySum.Left = lngTotLeft_1_Orig
2060        .Dev1.Left = lngDevLeft
2070        .txtReportCategorySum_line.Left = lngTotLeftLine_1_Moved
2080        .txtReportCategorySum_line.Visible = True
2090      Else
2100        .ReportSchedule.Visible = True
2110        .txtReportCategorySum.Left = lngTotLeft_1_Orig
2120        .Dev1.Left = lngDevLeft
2130        .txtReportCategorySum_line.Left = lngTotLeftLine_1_Orig
2140        .txtReportCategorySum_line.Visible = False
2150      End If
2160      If .ReportNumber = CRPT_INVEST_INFO Or .ReportNumber = CRPT_CHANGES Or .ReportNumber = CRPT_ON_HAND_BEGL Then
2170        .ReportSchedule.Visible = False
2180      End If
2190      lngDevCnt = lngDevCnt + 1&
2200    End With

EXITP:
2210    Exit Sub

ERRH:
2220    Select Case ERR.Number
        Case Else
2230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2240    End Select
2250    Resume EXITP

End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter3_Format"

2310    With Me
2320      If .ReportNumber2 = CRPT_ON_HAND_BEGL Or .ReportNumber2 = CRPT_ON_HAND_ENDL Then
2330        .Amount_Subtot2.Visible = False
2340      Else
2350        .Amount_Subtot2.Visible = True
2360      End If
2370      If (.ReportNumber2 = CRPT_CASH_BEG Or .ReportNumber2 = CRPT_NON_CASH_BEG) Or _
              (.ReportNumber2 = CRPT_CASH_END Or .ReportNumber2 = CRPT_NON_CASH_END) Then
2380        .ReportSchedule.Visible = False
2390        .Amount_Subtot2.Left = lngTotLeft_2_Moved
2400        .Amount_Subtot2_line.Left = lngTotLeftLine_2_Orig
2410        .Amount_Subtot2_line.Visible = False
2420      ElseIf .ReportNumber2 = CRPT_ON_HAND_BEG Or .ReportNumber2 = CRPT_ON_HAND_END Then
2430        .ReportSchedule.Visible = False
2440        .Amount_Subtot2.Left = lngTotLeft_2_Orig
2450        .Amount_Subtot2_line.Left = lngTotLeftLine_2_Moved
2460        .Amount_Subtot2_line.Visible = True
2470      ElseIf .ReportNumber2 = CRPT_RECEIPTS Then
2480        .TotalDisbursements_lbl.Caption = "Total Receipts During Period of Account"
2490        .Amount_Subtot2_line.Left = lngTotLeftLine_2_Moved
2500        .Amount_Subtot2_line.Visible = True
2510      ElseIf .ReportNumber2 = CRPT_DISBURSEMENTS Then
2520        .TotalDisbursements_lbl.Caption = "Total Disbursements During Period of Account"
2530        .Amount_Subtot2_line.Left = lngTotLeftLine_2_Moved
2540        .Amount_Subtot2_line.Visible = True
2550      Else
2560        .ReportSchedule.Visible = True
2570        .Amount_Subtot2.Left = lngTotLeft_2_Orig
2580        .Amount_Subtot2_line.Left = lngTotLeftLine_2_Orig
2590        .Amount_Subtot2_line.Visible = False
2600      End If
2610      If .ReportNumber2 = CRPT_INVEST_INFO Or .ReportNumber2 = CRPT_CHANGES Or .ReportNumber2 = CRPT_ON_HAND_BEGL Then
2620        .ReportSchedule.Visible = False
2630      End If
2640    End With

EXITP:
2650    Exit Sub

ERRH:
2660    Select Case ERR.Number
        Case Else
2670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2680    End Select
2690    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
' ** Stop the Additional Information group total from printing.

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

2710    With Me
2720      If .txtReportDivision = CRPT_DIV_ADDL Then
2730        Cancel = -1
2740      End If
2750    End With

EXITP:
2760    Exit Sub

ERRH:
2770    Select Case ERR.Number
        Case Else
2780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2790    End Select
2800    Resume EXITP

End Sub

Private Sub Report_Close()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2910    TAReports_SetZero False  ' ** Module Function: modReportFunctions.

2920    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2930    Set frm = Nothing
2940    Exit Sub

ERRH:
2950    Select Case ERR.Number
        Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

3010    varRetVal = Null
3020    With Me
3030      If frm.chkIncludeFootnote = True Then
3040        varRetVal = frm.CourtReports_Footnote
3050      End If
3060    End With

EXITP:
3070    ThisFootnote = varRetVal
3080    Exit Function

ERRH:
3090    varRetVal = RET_ERR
3100    Select Case ERR.Number
        Case Else
3110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3120    End Select
3130    Resume EXITP

End Function
