VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptFL_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptFL_02"

'VGC 04/04/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_FL"

Private blnNoData As Boolean
Private lngLegalNameMove_Height As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim ctl As Access.Control
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
430           .AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height
440           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
450           .CaseNum.Top = .CaseNum.Top + lngLegalNameMove_Height
              '.CaseNum_lbl.Top = .CaseNum_lbl.Top + lngLegalNameMove_Height
460           For Each ctl In .Section(acPageHeader).Controls
470             With ctl
480               Select Case .ControlType
                  Case acLabel
490                 Select Case .Name
                    Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                      ' ** Ignore.
500                 Case Else
510                   .Top = .Top + lngLegalNameMove_Height
520                 End Select
530               Case acLine
540                 If .Width <> Me.Width Then
550                   .Top = .Top + lngLegalNameMove_Height
560                 End If
570               End Select
580             End With
590           Next
600           .AcctName_Legal.Visible = True
610           .AcctName_Short.Visible = False
620         End If

            ' ** Set the date range heading.
630         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the ordinal and version heading.
640         .AccountHead_lbl.Caption = "Schedule A by " & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Guardian of Property", "Personal Representative")

            ' ** Set the bottom-left report version footer.
650         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Assume there is data to be printed.
660         blnNoData = False

670         .Caption = "CourtReport_FL_Receipts_" & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Grdn", "Rep") & "_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

680       Else
690         Cancel = -1
700       End If
710     End With

720     If Cancel = -1 Then
730       Beep
740       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
750     End If

EXITP:
760     Set ctl = Nothing
770     Exit Sub

ERRH:
780     Select Case ERR.Number
        Case Else
790       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
800     End Select
810     Resume EXITP

End Sub

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader2_Format"

        ' ** This may need to change with new, combined report.
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
1130      strTmp01 = Me.totdesc
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

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter2_Format"

        ' ** This may need to change with new, combined report.
1310    With Me
1320      If blnNoData Then
1330        .txtJournalTypeTotal_Inc_lbl.Visible = False
1340      Else
1350        Select Case [txtjournaltype]
            Case "Received"
1360          .txtJournalTypeTotal_Inc_lbl.Caption = "Total Other Receipts"
1370        Case "Dividend"
1380          .txtJournalTypeTotal_Inc_lbl.Caption = "Total Dividends"
1390        Case Else
1400          .txtJournalTypeTotal_Inc_lbl.Caption = "Total " & .txtjournaltype
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
1550      .Date.Visible = False
1560      .TotalShareFace.Visible = False
1570      .totdesc.Visible = False
1580      .Amount_Inc.Visible = False
1590      .txtTotalAmount_Inc_line.Visible = False
1600      .txtTotalAmount_Prin_line.Visible = False
1610      .txtTotalAmount_Inc.Visible = False
1620      .txtTotalAmount_Prin.Visible = False
1630      .txtJournalTypeTotal_Inc_lbl.Visible = False
1640      .txtJournalTypeTotal_Inc.Visible = False
1650      .txtJournalTypeTotal_Prin.Visible = False
1660      .txtGrandTotal_Inc_lbl.Visible = False
1670      .txtGrandTotal_Inc_line.Visible = False
1680      .txtGrandTotal_Prin_line.Visible = False
1690      .txtGrandTotal_Inc.Visible = False
1700      .txtGrandTotal_Prin.Visible = False
1710      .GroupHeader2.Visible = False
1720      If gblnLegalName = True Then
            ' ** Column 4: legalname
1730        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1740      Else
            ' ** Column 3: shorname
1750        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1760      End If
1770      .AcctName_Legal.Visible = False
1780      .AcctName_Short.Visible = False
1790      .AcctName_lbl.Visible = True
1800    End With

EXITP:
1810    Exit Sub

ERRH:
1820    Select Case ERR.Number
        Case Else
1830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1840    End Select
1850    Resume EXITP

End Sub

Private Sub Report_Close()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1910    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2010    varRetVal = Null
2020    With Me
2030      If Forms(strCallingForm).chkIncludeFootnote = True Then
2040        varRetVal = Forms(strCallingForm).CourtReports_Footnote
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
