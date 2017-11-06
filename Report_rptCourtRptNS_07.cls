VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_07"

'VGC 12/16/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_NS"

Private blnNoData As Boolean
Private lngLegalNameMove_Height As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim ctl As Access.Control
        Dim strFilter As String
        Dim lngLegalNameMove_Height As Long
        Dim intLen As Integer
        Dim varFootnote As Variant

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Procedure: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Select Case Forms(strCallingForm).chkPageOf
            Case True
150           .PageNumOf.Visible = True
160           .PageNum.Visible = False
170         Case False
180           .PageNum.Visible = True
190           .PageNumOf.Visible = False
200           .PageNumOf.ControlSource = vbNullString
210         End Select

            ' ** Because the PageFooterSection does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
220         If Forms(strCallingForm).chkIncludeFootnote = True Then
230           varFootnote = Forms(strCallingForm).CourtReports_Footnote
240           If IsNull(varFootnote) = False Then
250             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
260               intLen = Len(Trim(varFootnote))
270               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
280                 .Section("PageFooterSection").Height = 480&
290               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
300                 .Section("PageFooterSection").Height = 690&
310               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
320                 .Section("PageFooterSection").Height = 900&
330               End If
340             End If
350           End If
360         End If

370         lngLegalNameMove_Height = (.AcctName_Short.Height - 120&)

            ' ** Set the account name heading.
380         If gblnLegalName = True Then
390           .Section(acPageHeader).Height = (.Section(acPageHeader).Height + lngLegalNameMove_Height)
400           .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
410           .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
420           .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
430           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
440           For Each ctl In .Section(acPageHeader).Controls
450             With ctl
460               Select Case .ControlType
                  Case acLabel
470                 Select Case .Name
                    Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                      ' ** Ignore.
480                 Case Else
490                   .Top = .Top + lngLegalNameMove_Height
500                 End Select
510               Case acLine
520                 If .Width <> Me.Width Then
530                   .Top = .Top + lngLegalNameMove_Height
540                 End If
550               End Select
560             End With
570           Next
580           .AcctName_Legal.Visible = True
590           .AcctName_Short.Visible = False
600         End If

            ' ** Set the date range heading.
610         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
620         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
630         strFilter = "transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & _
              Format(gdatEndDate + 1, "mm/dd/yyyy") & "# And "
640         strFilter = strFilter & "[accountno] = '" & gstrAccountNo & "' And "
650         strFilter = strFilter & "("
660         strFilter = strFilter & "([journaltype] <> 'Purchase' And [journaltype] <> 'Sold' And [icash] > 0) Or "
670         strFilter = strFilter & "([journaltype] = 'Purchase' And [icash] < 0 And [pcash] = (-[cost]))"
680         strFilter = strFilter & ")"
690         .Filter = strFilter
700         .FilterOn = True

            ' ** Assume there is data to be printed.
710         blnNoData = False

720         .Caption = "CourtReport_NS_Receipts_of_Income_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

730       Else
740         Cancel = -1
750       End If
760     End With

770     If Cancel = -1 Then
780       Beep
790       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
800     End If

EXITP:
810     Set ctl = Nothing
820     Exit Sub

ERRH:
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader2_Format"

910     With Me
920       If blnNoData Then
930         .journaltype_lbl.Visible = False
940         .AmountJTypeSum_lbl.Visible = False
950       Else
960         Select Case .journaltype
            Case "Received"
970           .journaltype_lbl.Caption = "Other Income"
980           .AmountJTypeSum_lbl.Caption = "Other Income"
990         Case "Dividend"
1000          .journaltype_lbl.Caption = "Dividends"
1010          .AmountJTypeSum_lbl.Caption = "Dividends"
1020        Case "Misc."
1030          .journaltype_lbl.Caption = "Miscellaneous"
1040          .AmountJTypeSum_lbl.Caption = "Miscellaneous"
1050        Case Else
1060          .journaltype_lbl.Caption = .journaltype
1070          .AmountJTypeSum_lbl.Caption = .journaltype
1080        End Select
1090      End If
1100    End With

EXITP:
1110    Exit Sub

ERRH:
1120    Select Case ERR.Number
        Case Else
1130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1140    End Select
1150    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1210    With Me
1220  On Error Resume Next
1230      strTmp01 = .totdesc
1240      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1250  On Error GoTo ERRH
1260        Report_NoData Cancel  ' ** Though this may be triggered anyway.
1270      Else
1280  On Error GoTo ERRH
1290      End If
1300    End With

EXITP:
1310    Exit Sub

ERRH:
1320    Select Case ERR.Number
        Case Else
1330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1340    End Select
1350    Resume EXITP

End Sub

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter2_Format"

1410    With Me
1420      If blnNoData Then
1430        .AmountJTypeSum_lbl.Visible = False
1440      Else
1450        Select Case .journaltype
            Case "Received"
1460          .AmountJTypeSum_lbl.Caption = "Total Other Income"
1470        Case "Dividend"
1480          .AmountJTypeSum_lbl.Caption = "Total Dividends"
1490        Case "Misc."
1500          .AmountJTypeSum_lbl.Caption = "Total Miscellaneous"
1510        Case Else
1520          .AmountJTypeSum_lbl.Caption = "Total " & .journaltype
1530        End Select
1540      End If
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1710    blnNoData = True

1720    With Me
1730      .NoData_lbl.BackColor = CLR_WHT
1740      .NoData_lbl.Visible = True  ' ** Turn on the message.
1750      .journaltype_lbl.Visible = False
1760      .transdate.Visible = False  ' ** Make everything invisible that might show an error.
1770      .shareface_display.Visible = False
1780      .totdesc.Visible = False
1790      .amount.Visible = False
1800      .AmountSum.Visible = False
1810      .AmountSum_line.Visible = False
1820      .AmountJTypeSum.Visible = False
1830      .AmountJTypeSum_lbl.Visible = False
1840      .AmountTotalSumICash.Visible = False
1850      .AmountTotalSumICash_line.Visible = False
1860      .GroupHeader2.Visible = False
1870      If gblnLegalName = True Then
            ' ** Column 4: legalname
1880        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1890      Else
            ' ** Column 3: shorname
1900        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1910      End If
1920      .AcctName_Legal.Visible = False
1930      .AcctName_Short.Visible = False
1940      .AcctName_lbl.Visible = True
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

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2210    varRetVal = Null
2220    With Me
2230      If Forms(strCallingForm).chkIncludeFootnote = True Then
2240        varRetVal = Forms(strCallingForm).CourtReports_Footnote
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
