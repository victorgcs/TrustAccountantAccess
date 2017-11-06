VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_06H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_06H"

'VGC 09/08/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_CA"

' ** cmbAccounts combo box constants:
Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname

Private frm As Access.Form
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

230         If glngTaxCode_Distribution = 0& Then
240           glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
250         End If

            ' ** Because the PageFooterSection does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
260         If frm.chkIncludeFootnote = True Then
270           varFootnote = frm.CourtReports_Footnote
280           If IsNull(varFootnote) = False Then
290             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
300               intLen = Len(Trim(varFootnote))
310               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
320                 .Section("PageFooterSection").Height = 480&
330               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
340                 .Section("PageFooterSection").Height = 690&
350               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
360                 .Section("PageFooterSection").Height = 900&
370               End If
380             End If
390           End If
400         End If

410         lngLegalNameMove_Height = (.AcctName_Short.Height - 120&)

            ' ** Set the account name heading.
420         If gblnLegalName = True Then
430           If IsNull(frm.cmbAccounts.Column(CBX_A_LEGAL)) = True Then
                ' ** Use shortname.
440           Else
450             If Trim(frm.cmbAccounts.Column(CBX_A_LEGAL)) = vbNullString Then
                  ' ** Use shortname.
460             Else
470               .Section(acPageHeader).Height = (.Section(acPageHeader).Height + lngLegalNameMove_Height)
480               .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
490               .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
500               .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
510               .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
520               For Each ctl In .Section(acPageHeader).Controls
530                 With ctl
540                   Select Case .ControlType
                      Case acLabel
550                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Ignore.
560                     Case Else
570                       .Top = .Top + lngLegalNameMove_Height
580                     End Select
590                   Case acLine
600                     If .Width <> Me.Width Then
610                       .Top = .Top + lngLegalNameMove_Height
620                     End If
630                   End Select
640                 End With
650               Next
660               .AcctName_Legal.Visible = True
670               .AcctName_Short.Visible = False
680             End If
690           End If
700         End If

            ' ** Set the date range heading.
710         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
720         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
730         strFilter = "([transdate] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# AND " & _
              "[transdate] <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) "
740         strFilter = strFilter & "AND [accountno] = '" & gstrAccountNo & "' "
750         strFilter = strFilter & "AND (([journaltype] = 'Paid' AND [pcash] <> 0 AND [taxcode] = " & CStr(glngTaxCode_Distribution) & ") "  '= "Distribution"
760         strFilter = strFilter & "OR ([journaltype] = 'Paid' AND [icash] <> 0 AND [taxcode] = " & CStr(glngTaxCode_Distribution) & ") "  '= "Distribution"
770         strFilter = strFilter & "OR ([journaltype] = 'Withdrawn' AND [taxcode] = " & CStr(glngTaxCode_Distribution) & "))"  '= "Distribution"
780         .Filter = strFilter    '####  TAXCODE  ####
790         .FilterOn = True

800         .OrderBy = "transdate"
810         .OrderByOn = True

820         .Caption = "CourtReport_CA_Distributions_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

830       Else
840         Cancel = -1
850       End If
860     End With

870     If Cancel = -1 Then
880       Beep
890       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
900     End If

EXITP:
910     Exit Sub

ERRH:
920     Set ctl = Nothing
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

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
1240      .totdesc.Visible = False
1250      .amount.Visible = False
1260      .txtAmountTotalSumCost_lbl.Visible = False
1270      .txtAmountTotalSumCost.Visible = False
1280      .txtAmountTotalSumCost_line.Visible = False
1290      If gblnLegalName = True Then
            ' ** Column 4: legalname
1300        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1310      Else
            ' ** Column 3: shorname
1320        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1330      End If
1340      .AcctName_Legal.Visible = False
1350      .AcctName_Short.Visible = False
1360      .AcctName_lbl.Visible = True
1370    End With

EXITP:
1380    Exit Sub

ERRH:
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub Report_Close()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1510    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1520    Set frm = Nothing
1530    Exit Sub

ERRH:
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

1610    varRetVal = Null
1620    With Me
1630      If frm.chkIncludeFootnote = True Then
1640        varRetVal = frm.CourtReports_Footnote
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
