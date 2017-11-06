VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_00"

'VGC 11/27/2016: CHANGES!

' ** 'No data' produces a report with zeroes.

'Amount_Subtot2: =IIf([ReportNumber]=[RptNumPOH_Beg],[PropOnHand_Beg],IIf([ReportNumber]=[RptNumPOH_End],[PropOnHand_End],IIf([ReportNumber]=[RptNumNonCash_Beg],([PropOnHand_Beg]-[CashAssets_Beg]),IIf([ReportNumber]=[RptNumNonCash_End],([PropOnHand_End]-[CashAssets_End]),IIf([ReportNumber]=[RptNumInvestInfo],[InvestInfo],IIf([ReportNumber]=[RptNumInvestChange],[InvestChange],Sum([Amount])))))))
'=IIf(
'     [ReportNumber]=[RptNumPOH_Beg],[PropOnHand_Beg],
'     IIf(
'         [ReportNumber]=[RptNumPOH_End],[PropOnHand_End],
'         IIf(
'             [ReportNumber]=[RptNumNonCash_Beg],
'               (
'                [PropOnHand_Beg]-[CashAssets_Beg]   THIS APPEARS TO BE WHERE IT SHOULD SUBTRACT CASH ASSETS AT BEG (THE ENTERED AMOUNT)
'               ),                                   FROM THE NON-CASH ASSETS AT BEG, KEEPING THE TOTAL THE UNCHANGED!
'             IIf(                                   THIS FIGURE APPEARS TO GIVE GOOD RESULTS!
'                 [ReportNumber]=[RptNumNonCash_End],  IT'S THE TOTAL CHARGES/CREDITS THAT'S GETTING CHANGED!
'                   (
'                    [PropOnHand_End]-[CashAssets_End]
'                   ),
'                 IIf(
'                     [ReportNumber]=[RptNumInvestInfo],[InvestInfo],
'                     IIf(
'                         [ReportNumber]=[RptNumInvestChange],[InvestChange],
'                         Sum([Amount])
'                        )
'                    )
'                )
'            )
'        )
'    )

'Amount_Subtot2_runsum: =IIf([ReportNumber]=[RptNumPOH_Beg],[PropOnHand_Beg],IIf([ReportNumber]=[RptNumPOH_End],[PropOnHand_End],IIf([ReportNumber]=[RptNumNonCash_Beg],0,IIf([ReportNumber]=[RptNumNonCash_End],0,Sum([Amount])))))
'=IIf(
'     [ReportNumber]=[RptNumPOH_Beg],[PropOnHand_Beg],
'     IIf(
'         [ReportNumber]=[RptNumPOH_End],[PropOnHand_End],
'         IIf(
'             [ReportNumber]=[RptNumNonCash_Beg],0,  THE RUNNING SUM DOES NOT MAKE THE SUBTRACTION!
'             IIf(
'                 [ReportNumber]=[RptNumNonCash_End],0,
'                 Sum([Amount])
'                )
'            )
'        )
'    )

'txtAmountTotalSumCost_charges: =[Amount_Subtot3_runsum]-[CashAssets_Beg]

'txtAmountTotalSumCost_credits: =[Amount_Subtot3_runsum]-[CashAssets_End]

Private Const strCallingForm As String = "frmRpt_CourtReports_CA"

' ** cmbAccounts combo box constants:
Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname

Private frm As Access.Form
Private lngTotLeft_Orig As Long, lngTotLeft_Moved As Long
Private lngTotLeftLine_Orig As Long, lngTotLeftLine_Moved As Long
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
650           .Section("GroupFooter5").Visible = True
660         Else
670           .Section("GroupFooter5").Visible = False
680         End If

690         lngTotLeft_Orig = .Amount_Subtot2.Left
700         lngTotLeft_Moved = lngTotLeft_Orig * 0.8   ' ** Move left about 20% of the text box's width.
710         lngTotLeftLine_Orig = .Amount_Subtot2_line.Left
720         lngTotLeftLine_Moved = lngTotLeft_Moved + (.Amount_Subtot2.Width - .Amount_Subtot2_line.Width)

730         .RptNumPOH_Beg.ControlSource = "=" & CStr(CRPT_ON_HAND_BEG)
740         .RptNumPOH_End.ControlSource = "=" & CStr(CRPT_ON_HAND_END)
750         .PropOnHand_Beg.ControlSource = "=" & CStr(gdblCrtRpt_CA_POHBeg)
760         .PropOnHand_End.ControlSource = "=" & CStr(gdblCrtRpt_CA_POHEnd)

770         .RptNumNonCash_Beg.ControlSource = "=" & CStr(CRPT_NON_CASH_BEG)
780         .RptNumNonCash_End.ControlSource = "=" & CStr(CRPT_NON_CASH_END)
790         .CashAssets_Beg.ControlSource = "=" & CStr(gdblCrtRpt_CA_COHBeg)
800         .CashAssets_End.ControlSource = "=" & CStr(gdblCrtRpt_CA_COHEnd)

810         .RptNumInvestInfo.ControlSource = "=" & CStr(CRPT_INVEST_INFO)
820         .RptNumInvestChange.ControlSource = "=" & CStr(CRPT_CHANGES)
830         .InvestInfo.ControlSource = "=" & CStr(gdblCrtRpt_CA_InvestInfo)
840         .InvestChange.ControlSource = "=" & CStr(gdblCrtRpt_CA_InvestChange)

850         .Caption = "CourtReport_CA_Summary_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

860       Else
870         Cancel = -1
880       End If
890     End With

900     If Cancel = -1 Then
910       Beep
920       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
930     End If

EXITP:
940     Exit Sub

ERRH:
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Sub

Private Sub Report_Activate()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Activate"

1010    TAReports_SetZero True  ' ** Module Function: modReportFunctions.

EXITP:
1020    Exit Sub

ERRH:
1030    Select Case ERR.Number
        Case Else
1040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1050    End Select
1060    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

1110    With Me
1120      If .ReportDivisionTitle = "CREDITS" Then
1130        .Section("GroupFooter1").ForceNewPage = 2
            ' ** 07/21/2008 NOTE: The 'Entered' amount (Cash Assets at Beginning of Account Period),
            ' ** gets added twice into the "CHARGES" total. Likewise, the 'Entered' amount (Cash Assets
            ' ** at End of Account Period), gets added twice into the "CREDITS" total.
            ' ** These extra fields subtract them out.
1140        .txtAmountTotalSumCost_credits.Visible = True
1150        .txtAmountTotalSumCost_charges.Visible = False
1160      Else
1170        .Section("GroupFooter1").ForceNewPage = 0
1180        .txtAmountTotalSumCost_credits.Visible = False
1190        .txtAmountTotalSumCost_charges.Visible = True
1200      End If
1210    End With

        ' ** The ForceNewPage property uses the following settings.
        ' **   0  None            The current section (the section for which you're setting the property) is printed on the
        ' **                      current page. (Default)
        ' **   1  Before Section  The current section is printed at the top of a new page.
        ' **   2  After Section   The section immediately following the current section is printed at the top of a new page.
        ' **   3  Before & After  The current section is printed at the top of a new page, and the next section is printed
        ' **                      at the top of a new page.

EXITP:
1220    Exit Sub

ERRH:
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter3_Format"

1310    With Me
1320      If .ReportNumber = CRPT_ON_HAND_BEGL Or .ReportNumber = CRPT_ON_HAND_ENDL Then
1330        .Amount_Subtot2.Visible = False
1340      Else
1350        .Amount_Subtot2.Visible = True
1360      End If
1370      If (.ReportNumber = CRPT_CASH_BEG Or .ReportNumber = CRPT_NON_CASH_BEG) Or _
              (.ReportNumber = CRPT_CASH_END Or .ReportNumber = CRPT_NON_CASH_END) Then
1380        .ReportSchedule.Visible = False
1390        .Amount_Subtot2.Left = lngTotLeft_Moved
1400        .Amount_Subtot2_line.Left = lngTotLeftLine_Orig
1410        .Amount_Subtot2_line.Visible = False
1420      ElseIf .ReportNumber = CRPT_ON_HAND_BEG Or .ReportNumber = CRPT_ON_HAND_END Then
1430        .ReportSchedule.Visible = False
1440        .Amount_Subtot2.Left = lngTotLeft_Orig
1450        .Amount_Subtot2_line.Left = lngTotLeftLine_Moved
1460        .Amount_Subtot2_line.Visible = True
1470      Else
1480        .ReportSchedule.Visible = True
1490        .Amount_Subtot2.Left = lngTotLeft_Orig
1500        .Amount_Subtot2_line.Left = lngTotLeftLine_Orig
1510        .Amount_Subtot2_line.Visible = False
1520      End If
1530      If .ReportNumber = CRPT_INVEST_INFO Or .ReportNumber = CRPT_CHANGES Or .ReportNumber = CRPT_ON_HAND_BEGL Then
1540        .ReportSchedule.Visible = False
1550      End If
1560    End With

EXITP:
1570    Exit Sub

ERRH:
1580    Select Case ERR.Number
        Case Else
1590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1600    End Select
1610    Resume EXITP

End Sub

Private Sub GroupFooter5_Format(Cancel As Integer, FormatCount As Integer)
' ** Select only details for the disbursements areas.

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter5_Format"

1710    With Me
1720      Select Case .ReportNumber
          Case CRPT_DISBURSEMENTS
1730        If .txtReportDivision = CRPT_DIV_CREDITS Then ' And .txtReportGroup = 30 Then  ' ** CALIFORNIA HARD-CODED NUMBERS!
1740          Cancel = 0
1750          If .revcode_DESC = "Dummy entry" Then
1760            Cancel = -1  ' ** don't print dummy entries.
1770          End If
1780        Else
1790          Cancel = -1
1800        End If
1810      Case Else
1820        Cancel = -1
1830      End Select
1840    End With

EXITP:
1850    Exit Sub

ERRH:
1860    Select Case ERR.Number
        Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
' ** Stop the Additional Information group total from printing.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

1910    With Me
1920      If .txtReportDivision = CRPT_DIV_ADDL Then
1930        Cancel = -1
1940      End If
1950    End With

EXITP:
1960    Exit Sub

ERRH:
1970    Select Case ERR.Number
        Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub Report_Close()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2110    TAReports_SetZero False  ' ** Module Function: modReportFunctions.

2120    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2130    Set frm = Nothing
2140    Exit Sub

ERRH:
2150    Select Case ERR.Number
        Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2210    varRetVal = Null
2220    With Me
2230      If frm.chkIncludeFootnote = True Then
2240        varRetVal = frm.CourtReports_Footnote
2250      End If
2260    End With

EXITP:
2270    ThisFootnote = varRetVal
2280    Exit Function

ERRH:
2290    varRetVal = RET_ERR
2300    Select Case ERR.Number
        Case Else
2310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2320    End Select
2330    Resume EXITP

End Function
