VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_03A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_03A"

'VGC 04/04/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_NS"

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

140         If glngTaxCode_Distribution = 0& Then
150           glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
160         End If

170         Select Case Forms(strCallingForm).chkPageOf
            Case True
180           .PageNumOf.Visible = True
190           .PageNum.Visible = False
200         Case False
210           .PageNum.Visible = True
220           .PageNumOf.Visible = False
230           .PageNumOf.ControlSource = vbNullString
240         End Select

            ' ** Because the PageFooterSection does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
250         If Forms(strCallingForm).chkIncludeFootnote = True Then
260           varFootnote = Forms(strCallingForm).CourtReports_Footnote
270           If IsNull(varFootnote) = False Then
280             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
290               intLen = Len(Trim(varFootnote))
300               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
310                 .Section("PageFooterSection").Height = 480&
320               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
330                 .Section("PageFooterSection").Height = 690&
340               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
350                 .Section("PageFooterSection").Height = 900&
360               End If
370             End If
380           End If
390         End If

400         lngLegalNameMove_Height = (.AcctName_Short.Height - 120&)

            ' ** Set the account name heading.
410         If gblnLegalName = True Then
420           .Section(acPageHeader).Height = (.Section(acPageHeader).Height + lngLegalNameMove_Height)
430           .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
440           .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
450           .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
460           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
470           For Each ctl In .Section(acPageHeader).Controls
480             With ctl
490               Select Case .ControlType
                  Case acLabel
500                 Select Case .Name
                    Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                      ' ** Ignore.
510                 Case Else
520                   .Top = .Top + lngLegalNameMove_Height
530                 End Select
540               Case acLine
550                 If .Width <> Me.Width Then
560                   .Top = .Top + lngLegalNameMove_Height
570                 End If
580               End Select
590             End With
600           Next
610           .AcctName_Legal.Visible = True
620           .AcctName_Short.Visible = False
630         End If

            ' ** Set the date range heading.
640         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
650         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
660         strFilter = "transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & _
              Format(gdatEndDate, "mm/dd/yyyy") & "# And "
670         strFilter = strFilter & "accountno = '" & gstrAccountNo & "' And "
680         strFilter = strFilter & "((journaltype = 'Paid' And pcash <> 0 And taxcode <> " & CStr(glngTaxCode_Distribution) & ") Or "  '<> "Distribution"
690         strFilter = strFilter & "(journaltype = 'Misc.' And pcash < 0 ) Or "
700         strFilter = strFilter & "(journaltype = 'Cost Adj.' And cost < 0) Or "
710         strFilter = strFilter & "(journaltype = 'Withdrawn' And taxcode <> " & CStr(glngTaxCode_Distribution) & "))"  '<> "Distribution"
720         .Filter = strFilter    '####  TAXCODE  ####
730         .FilterOn = True

            ' ** Turn on the group header for the reveune / expense code lines.
740         If gblnUseReveuneExpenseCodes = True Then
750           .GroupHeader3.Visible = True
760         Else
770           .GroupHeader3.Visible = False
780         End If

790         .Caption = "CourtReport_NS_Disbursements_of_Principal_Grouped_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

800       Else
810         Cancel = -1
820       End If
830     End With

840     If Cancel = -1 Then
850       Beep
860       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
870     End If

EXITP:
880     Set ctl = Nothing
890     Exit Sub

ERRH:
900     Select Case ERR.Number
        Case Else
910       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
920     End Select
930     Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1010    With Me
1020  On Error Resume Next
1030      strTmp01 = .totdesc
1040      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1050  On Error GoTo ERRH
1060        Report_NoData Cancel  ' ** Though this may be triggered anyway.
1070      Else
1080  On Error GoTo ERRH
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

Private Sub Report_NoData(Cancel As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1210    With Me
1220      .NoData_lbl.BackColor = CLR_WHT
1230      .NoData_lbl.Visible = True  ' ** Turn on the message.
1240      .totdesc.Visible = False  ' ** Make everything invisible that might show an error.
1250      .amount.Visible = False
1260      .AmountTotalSumCost_lbl.Visible = False
1270      .AmountTotalSumCost.Visible = False
1280      .AmountTotalSumCost_line.Visible = False
1290      .totdesc.Visible = False
1300      .amount.Visible = False
1310      .NoData_lbl.BackColor = CLR_WHT
1320      .NoData_lbl.Visible = True
1330      .GroupHeader3.Visible = False
1340      .GroupFooter1.Visible = False
1350      If gblnLegalName = True Then
            ' ** Column 4: legalname
1360        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1370      Else
            ' ** Column 3: shorname
1380        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1390      End If
1400      .AcctName_Legal.Visible = False
1410      .AcctName_Short.Visible = False
1420      .AcctName_lbl.Visible = True
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

Private Sub Report_Close()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1510    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

1610    varRetVal = Null
1620    With Me
1630      If Forms(strCallingForm).chkIncludeFootnote = True Then
1640        varRetVal = Forms(strCallingForm).CourtReports_Footnote
1650      End If
1660    End With

EXITP:
1670    ThisFootnote = varRetVal
1680    Exit Function

ERRH:
1690    varRetVal = RET_ERR
1700    Select Case ERR.Number
        Case Else
1710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1720    End Select
1730    Resume EXITP

End Function
