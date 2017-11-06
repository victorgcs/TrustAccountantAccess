VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_06"

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
630         strFilter = "([transdate] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And [transdate] < #" & _
              Format(gdatEndDate + 1, "mm/dd/yyyy") & "#) And "
640         strFilter = strFilter & "([accountno] = '" & gstrAccountNo & "') And "
650         strFilter = strFilter & "("
660         strFilter = strFilter & "([journaltype] = 'Sold' And [GainLoss] = 0) Or "
670         strFilter = strFilter & "([journaltype] = 'Deposit' And [jcomment] Like '*stock split*') Or "
680         strFilter = strFilter & "([journaltype] = 'Liability' And [cost] > 0 And [cost] = (-[pcash]))"  'VGC 01/29/2013: CHANGES!
            'strFilter = strFilter & "([journaltype] = 'Liability' And [cost] < 0)"  'VGC 08/20/2010: CHANGES!
690         strFilter = strFilter & ")"
700         .Filter = strFilter
710         .FilterOn = True

720         .Caption = "CourtReport_NS_Change_in_Investment_Holdings_" & gstrAccountNo & "_" & _
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
