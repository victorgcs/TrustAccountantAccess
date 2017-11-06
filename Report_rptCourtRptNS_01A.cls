VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_01A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_01A"

'VGC 04/04/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_NS"

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
630         strFilter = "[transdate] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And [transdate] <= #" & _
              Format(gdatEndDate, "mm/dd/yyyy") & "# And "
640         strFilter = strFilter & "[accountno] = '" & gstrAccountNo & "' And "
650         strFilter = strFilter & "("
660         strFilter = strFilter & "([journaltype] = 'Received' And [pcash] > 0) Or "
670         strFilter = strFilter & "([journaltype] = 'Misc.' And [pcash] > 0) Or "
680         strFilter = strFilter & "([journaltype] = 'Cost Adj.' And [cost] > 0) Or "
690         strFilter = strFilter & "([journaltype] = 'Deposit' And Not([jcomment] Like '*stock split*')) Or "
700         strFilter = strFilter & "(([journaltype] = 'Liability' Or [journaltype] = 'Liability (+)') And [cost] < 0 And [pcash] = 0) Or "
710         strFilter = strFilter & "(([journaltype] = 'Liability' Or [journaltype] = 'Liability (-)') And [cost] > 0 And [pcash] = 0)"
720         strFilter = strFilter & ")"
730         .Filter = strFilter
740         .FilterOn = True

750         .OrderBy = "transdate"
760         .OrderByOn = True

770         .Caption = "CourtReport_NS_Receipts_Of_Principal_Grouped_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

780       Else
790         Cancel = -1
800       End If
810     End With

820     If Cancel = -1 Then
830       Beep
840       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
850     End If

EXITP:
860     Set ctl = Nothing
870     Exit Sub

ERRH:
880     Select Case ERR.Number
        Case Else
890       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
900     End Select
910     Resume EXITP

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
1260      .shareface_display.Visible = False
1270      .AmountTotalSumCost_lbl.Visible = False
1280      .AmountTotalSumCost.Visible = False
1290      .AmountTotalSumCost_line.Visible = False
1300      If gblnLegalName = True Then
            ' ** Column 4: legalname
1310        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1320      Else
            ' ** Column 3: shorname
1330        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1340      End If
1350      .AcctName_Legal.Visible = False
1360      .AcctName_Short.Visible = False
1370      .AcctName_lbl.Visible = True
1380    End With

EXITP:
1390    Exit Sub

ERRH:
1400    Select Case ERR.Number
        Case Else
1410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1420    End Select
1430    Resume EXITP

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
