VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_02"

'VGC 04/04/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_CA"

' ** cmbAccounts combo box constants:
Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname

Private frm As Access.Form
Private blnNoData As Boolean
Private lngLegalNameMove_Height As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim ctl As Access.Control
        Dim strFilter As String
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
500               For Each ctl In .Section(acPageHeader).Controls
510                 With ctl
520                   Select Case .ControlType
                      Case acLabel
530                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Ignore.
540                     Case Else
550                       .Top = .Top + lngLegalNameMove_Height
560                     End Select
570                   Case acLine
580                     If .Width <> Me.Width Then
590                       .Top = .Top + lngLegalNameMove_Height
600                     End If
610                   End Select
620                 End With
630               Next
640               .AcctName_Legal.Visible = True
650               .AcctName_Short.Visible = False
660             End If
670           End If
680         End If

            ' ** Set the date range heading.
690         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
700         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
710         strFilter = "([date] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And " & _
              "[date] <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "# + 1) And "
720         strFilter = strFilter & "[accountno] = '" & gstrAccountNo & "' And "
730         strFilter = strFilter & "([reportnumber] = " & CStr(CRPT_RECEIPTS) & " And "
740         strFilter = strFilter & "[amount] <> 0)"
750         .Filter = strFilter
760         .FilterOn = True

            ' ** Assume there is data to be printed.
770         blnNoData = False

780         .Caption = "CourtReport_CA_Receipts_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

790       Else
800         Cancel = -1
810       End If
820     End With

830     If Cancel = -1 Then
840       Beep
850       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
860     End If

EXITP:
870     Set ctl = Nothing
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
' ** journaltype Header.

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

1010    With Me
1020      If blnNoData Then
1030        .journaltype_lbl.Visible = False
1040      Else
1050        Select Case .txtjournaltype
            Case "Dividend"
1060          .journaltype_lbl.Caption = "Dividends"
1070        Case "Misc."
1080          .journaltype_lbl.Caption = "Miscellaneous"
1090        Case "Received"
1100          .journaltype_lbl.Caption = "Other Receipts"
1110        Case Else
1120          .journaltype_lbl.Caption = .txtjournaltype
1130        End Select
1140      End If
1150    End With

EXITP:
1160    Exit Sub

ERRH:
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1310    With Me
1320  On Error Resume Next
1330      strTmp01 = .totdesc
1340      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1350  On Error GoTo ERRH
1360        Report_NoData Cancel  ' ** Though this may be triggered anyway.
1370      Else
1380  On Error GoTo ERRH
1390      End If
1400    End With

EXITP:
1410    Exit Sub

ERRH:
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
' ** journaltype Footer.

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

1510    With Me
1520      If blnNoData Then
1530        .txtJournalTypeTotal_lbl.Visible = False
1540      Else
1550        Select Case [txtjournaltype]
            Case "Dividend"
1560          .txtJournalTypeTotal_lbl.Caption = "Total Dividends"
1570        Case "Misc."
1580          .txtJournalTypeTotal_lbl.Caption = "Total Miscellaneous"
1590        Case "Received"
1600          .txtJournalTypeTotal_lbl.Caption = "Total Other Receipts"
1610        Case Else
1620          .txtJournalTypeTotal_lbl.Caption = "Total " & .txtjournaltype
1630        End Select
1640      End If
1650    End With

EXITP:
1660    Exit Sub

ERRH:
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1810    blnNoData = True

1820    With Me
1830      .NoData_lbl.BackColor = CLR_WHT
1840      .NoData_lbl.Visible = True  ' ** Turn on the message.
1850      .transdate.Visible = False
1860      .TotalShareFace.Visible = False
1870      .totdesc.Visible = False
1880      .Gain.Visible = False
1890      .txtTotalAmount_line.Visible = False
1900      .txtTotalAmount.Visible = False
1910      .txtJournalTypeTotal_lbl.Visible = False
1920      .txtJournalTypeTotal.Visible = False
1930      .txtGrandTotal_lbl.Visible = False
1940      .txtGrandTotal_line.Visible = False
1950      .txtGrandTotal.Visible = False
1960      .GroupHeader1.Visible = False
1970      If gblnLegalName = True Then
            ' ** Column 4: legalname
1980        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1990      Else
            ' ** Column 3: shorname
2000        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
2010      End If
2020      .AcctName_Legal.Visible = False
2030      .AcctName_Short.Visible = False
2040      .AcctName_lbl.Visible = True
2050    End With

EXITP:
2060    Exit Sub

ERRH:
2070    Select Case ERR.Number
        Case Else
2080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2090    End Select
2100    Resume EXITP

End Sub

Private Sub Report_Close()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2210    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2220    Set frm = Nothing
2230    Exit Sub

ERRH:
2240    Select Case ERR.Number
        Case Else
2250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2260    End Select
2270    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2310    varRetVal = Null
2320    With Me
2330      If frm.chkIncludeFootnote = True Then
2340        varRetVal = frm.CourtReports_Footnote
2350      End If
2360    End With

EXITP:
2370    ThisFootnote = varRetVal
2380    Exit Function

ERRH:
2390    varRetVal = RET_ERR
2400    Select Case ERR.Number
        Case Else
2410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2420    End Select
2430    Resume EXITP

End Function
