VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_03"

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

740         .Caption = "CourtReport_NS_Disbursements_of_Principal_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

750       Else
760         Cancel = -1
770       End If
780     End With

790     If Cancel = -1 Then
800       Beep
810       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
820     End If

EXITP:
830     Set ctl = Nothing
840     Exit Sub

ERRH:
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
870     End Select
880     Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

910     With Me
920   On Error Resume Next
930       strTmp01 = .totdesc
940       If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
950   On Error GoTo ERRH
960         Report_NoData Cancel  ' ** Though this may be triggered anyway.
970       Else
980   On Error GoTo ERRH
990       End If
1000    End With

EXITP:
1010    Exit Sub

ERRH:
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1110    With Me
1120      .NoData_lbl.BackColor = CLR_WHT
1130      .NoData_lbl.Visible = True  ' ** Turn on the message.
1140      .totdesc.Visible = False  ' ** Make everything invisible that might show an error.
1150      .amount.Visible = False
1160      .AmountTotalSumCost_lbl.Visible = False
1170      .AmountTotalSumCost.Visible = False
1180      .AmountTotalSumCost_line.Visible = False
1190      If gblnLegalName = True Then
            ' ** Column 4: legalname
1200        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1210      Else
            ' ** Column 3: shorname
1220        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1230      End If
1240      .AcctName_Legal.Visible = False
1250      .AcctName_Short.Visible = False
1260      .AcctName_lbl.Visible = True
1270    End With

EXITP:
1280    Exit Sub

ERRH:
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub

Private Sub Report_Close()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1410    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1420    Exit Sub

ERRH:
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

1510    varRetVal = Null
1520    With Me
1530      If Forms(strCallingForm).chkIncludeFootnote = True Then
1540        varRetVal = Forms(strCallingForm).CourtReports_Footnote
1550      End If
1560    End With

EXITP:
1570    ThisFootnote = varRetVal
1580    Exit Function

ERRH:
1590    varRetVal = RET_ERR
1600    Select Case ERR.Number
        Case Else
1610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1620    End Select
1630    Resume EXITP

End Function
