VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_10"

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
710         strFilter = "([date] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# " & _
              "AND [date] <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "# + 1) " & _
              "AND [accountno] = '" & gstrAccountNo & "' " & _
              "AND ([revcode_ID] = " & CStr(REVID_OCHG) & " AND [amount] <> 0)"
720         .Filter = strFilter
730         .FilterOn = True

            ' ** Assume there is data to be printed.
740         blnNoData = False

750         .Caption = "CourtReport_CA_Other_Charges_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

760       Else
770         Cancel = -1
780       End If
790     End With

800     If Cancel = -1 Then
810       Beep
820       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
830     End If

EXITP:
840     Set ctl = Nothing
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

910     With Me
920       If blnNoData Then
930         .journaltype_lbl.Visible = False
940       Else
950         Select Case .txtjournaltype
            Case "Received"
960           .journaltype_lbl.Caption = "Other Receipts"
970         Case "Dividend"
980           .journaltype_lbl.Caption = "Dividends"
990         Case Else
1000          .journaltype_lbl.Caption = .txtjournaltype
1010        End Select
1020      End If
1030    End With

EXITP:
1040    Exit Sub

ERRH:
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1110    With Me
1120  On Error Resume Next
1130      strTmp01 = .totdesc
1140      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1150  On Error GoTo ERRH
1160        Report_NoData Cancel  ' ** Though this may be triggered anyway.
1170      Else
1180  On Error GoTo ERRH
1190      End If
1200    End With

EXITP:
1210    Exit Sub

ERRH:
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

1310    With Me
1320      If blnNoData Then
1330        .txtJournalTypeTotal_lbl.Visible = False
1340      Else
1350        Select Case [txtjournaltype]
            Case "Received"
1360          .txtJournalTypeTotal_lbl.Caption = "Total Other Receipts"
1370        Case "Dividend"
1380          .txtJournalTypeTotal_lbl.Caption = "Total Dividends"
1390        Case Else
1400          .txtJournalTypeTotal_lbl.Caption = "Total " & .txtjournaltype
1410        End Select
1420      End If
1430    End With

EXITP:
1440    Exit Sub

ERRH:
1450    Select Case ERR.Number
        Case Else
1460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1470    End Select
1480    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1510    blnNoData = True

1520    With Me
1530      .NoData_lbl.BackColor = CLR_WHT
1540      .NoData_lbl.Visible = True  ' ** Turn on the message.
1550      .transdate.Visible = False
1560      .TotalShareFace.Visible = False
1570      .totdesc.Visible = False
1580      .Gain.Visible = False
1590      .txtTotalAmount_line.Visible = False
1600      .txtTotalAmount.Visible = False
1610      .txtJournalTypeTotal_lbl.Visible = False
1620      .txtJournalTypeTotal.Visible = False
1630      .txtGrandTotal_lbl.Visible = False
1640      .txtGrandTotal_line.Visible = False
1650      .txtGrandTotal.Visible = False
1660      .GroupHeader1.Visible = False
1670      If gblnLegalName = True Then
            ' ** Column 4: legalname
1680        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1690      Else
            ' ** Column 3: shorname
1700        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1710      End If
1720      .AcctName_Legal.Visible = False
1730      .AcctName_Short.Visible = False
1740      .AcctName_lbl.Visible = True
1750    End With

EXITP:
1760    Exit Sub

ERRH:
1770    Select Case ERR.Number
        Case Else
1780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1790    End Select
1800    Resume EXITP

End Sub

Private Sub Report_Close()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1910    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1920    Set frm = Nothing
1930    Exit Sub

ERRH:
1940    Select Case ERR.Number
        Case Else
1950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1960    End Select
1970    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2010    varRetVal = Null
2020    With Me
2030      If frm.chkIncludeFootnote = True Then
2040        varRetVal = frm.CourtReports_Footnote
2050      End If
2060    End With

EXITP:
2070    ThisFootnote = varRetVal
2080    Exit Function

ERRH:
2090    varRetVal = RET_ERR
2100    Select Case ERR.Number
        Case Else
2110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2120    End Select
2130    Resume EXITP

End Function
