VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_08"

'VGC 04/04/2016: CHANGES!

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
480               .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
490               For Each ctl In .Section(acPageHeader).Controls
500                 With ctl
510                   Select Case .ControlType
                      Case acLabel
520                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Ignore.
530                     Case Else
540                       .Top = .Top + lngLegalNameMove_Height
550                     End Select
560                   Case acLine
570                     If .Width <> Me.Width Then
580                       .Top = .Top + lngLegalNameMove_Height
590                     End If
600                   End Select
610                 End With
620               Next
630               .AcctName_Legal.Visible = True
640               .AcctName_Short.Visible = False
650             End If
660           End If
670         End If

            ' ** Set the date range heading.
680         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
690         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
700         strFilter = "([transdate] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And " & _
              "[transdate] < #" & Format(gdatEndDate + 1, "mm/dd/yyyy") & "#) And "
710         strFilter = strFilter & "([accountno] = '" & gstrAccountNo & "') And "
720         strFilter = strFilter & "("
730         strFilter = strFilter & "([journaltype] = 'Purchase') Or "
740         strFilter = strFilter & "([journaltype] = 'Liability' And [cost] < 0 And [cost] = (-[pcash]))"  'VGC 01/29/2013: CHANGES!
750         strFilter = strFilter & ")"
760         .Filter = strFilter
770         .FilterOn = True

            ' ** 01/24/2009: New formula for Amount: =(IIf([icash]=0,[pcash],IIf([pcash]=0,[icash],[pcash])))*-1
            ' ** If the original (icash + pcash) was intended to simplify choosing whichever column has the
            ' ** value (since most Purchases only have a value in one or the other), no provision was given
            ' ** to the case when both have a value. This new formula, then, chooses pcash over both.
            ' ** 2 examples: 10/18/2007, ($17.18) ($9,845.28); 01/22/2009, ($100.00) ($1,000.00)
            ' ** Do taxcode's matter?            4 (Non-Taxable)                    2 (Fed. Tax Only)
            ' ** Both have a revcode_ID of 1, Unspecified Income  ?
            ' ** Not yet checked with Rich.
780         .OrderBy = "transdate"
790         .OrderByOn = True

800         .Caption = "CourtReport_CA_Information_For_Investments_Made_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

810       Else
820         Cancel = -1
830       End If
840     End With

850     If Cancel = -1 Then
860       Beep
870       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
880     End If

EXITP:
890     Set ctl = Nothing
900     Exit Sub

ERRH:
910     Select Case ERR.Number
        Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1010    With Me
1020  On Error Resume Next
1030      strTmp01 = .txtDescription
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
1240      .txtDescription.Visible = False
1250      .txtAmount.Visible = False
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
