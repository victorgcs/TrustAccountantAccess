VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptFL_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptFL_03"

'VGC 09/19/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_FL"

Private lngTotDesc_Width As Long, lngCost_Top As Long
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
640         .AccountHead_lbl.Caption = Forms(strCallingForm).cmdPreview04_Sch_lbl.Caption & " by " & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Guardian of Property", "Personal Representative")

            ' ** Set the bottom-left report version footer.
650         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

660         lngTotDesc_Width = .totdesc.Width
670         lngCost_Top = .Cost2.Top

680         .Caption = "CourtReport_FL_Capital_Transactions_and_Adjustments_" & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Grdn", "Rep") & "_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

690       Else
700         Cancel = -1
710       End If
720     End With

730     If Cancel = -1 Then
740       Beep
750       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
760     End If

EXITP:
770     Set ctl = Nothing
780     Exit Sub

ERRH:
790     Select Case ERR.Number
        Case Else
800       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
810     End Select
820     Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim strTmp01 As String

910     With Me
920   On Error Resume Next
930       strTmp01 = .journaltypex
940       If ERR.Number = 0 Then
950   On Error GoTo ERRH
960         Select Case .journaltypex
            Case "Sales"
970           .GroupFooter3.Visible = True  ' ** journaltypey Footer
980           .txtTotalGain_lbl.Caption = "Net Gain/Loss on Sales"
990           .GainLoss_Gain.Visible = True
1000          .GainLoss_Loss.Visible = True
1010          .totdesc.Width = lngTotDesc_Width
1020          .Cost2.Visible = False
1030          .GroupFooter1.Visible = True  ' ** uniqueid Footer.
1040        Case "Purchases", "Stock Splits"
1050          .GroupFooter3.Visible = False  ' ** journaltypey Footer
1060          .GainLoss_Gain.Visible = False
1070          .GainLoss_Loss.Visible = False
1080          .totdesc.Width = lngTotDesc_Width - 840&
1090          .Cost2.Visible = True
1100          .GroupFooter1.Visible = False  ' ** uniqueid Footer.
1110        Case "Cost Adjustments"
1120          .GroupFooter3.Visible = True  ' ** journaltypey Footer
1130          .txtTotalGain_lbl.Caption = "Net Gain/Loss on Cost Adjustments"
1140          .GainLoss_Gain.Visible = True
1150          .GainLoss_Loss.Visible = True
1160          .totdesc.Width = lngTotDesc_Width
1170          .Cost2.Visible = False
1180          .GroupFooter1.Visible = False  ' ** uniqueid Footer.
1190        Case "Withdrawals"
1200          .GroupFooter3.Visible = True  ' ** journaltypey Footer
1210          .txtTotalGain_lbl.Caption = "Net Gain/Loss on Withdrawals"
1220          .GainLoss_Gain.Visible = True
1230          .GainLoss_Loss.Visible = True
1240          .totdesc.Width = lngTotDesc_Width
1250          .Cost2.Visible = False
1260          .GroupFooter1.Visible = False  ' ** uniqueid Footer.
1270        End Select
1280      Else
1290  On Error GoTo ERRH
1300      End If
1310    End With

EXITP:
1320    Exit Sub

ERRH:
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1410    With Me
1420  On Error Resume Next
1430      strTmp01 = .totdesc
1440      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1450  On Error GoTo ERRH
1460        Report_NoData Cancel  ' ** Though this may be triggered anyway.
1470      Else
1480  On Error GoTo ERRH
1490      End If
1500    End With

EXITP:
1510    Exit Sub

ERRH:
1520    Select Case ERR.Number
        Case Else
1530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1540    End Select
1550    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1610    With Me
1620      .NoData_lbl.BackColor = CLR_WHT
1630      .NoData_lbl.Visible = True  ' ** Turn on the message.
1640      .totdesc.Visible = False  ' ** Make everything invisible that might show an error.
1650      .GainLoss_Gain.Visible = False
1660      .GainLoss_Loss.Visible = False
1670      .Date.Visible = False
1680      .TotalShareFace.Visible = False
1690      .txtTotalGain.Visible = False
1700      .txtTotalLoss.Visible = False
1710      .txtTotalGain_lbl.Visible = False
1720      .txtTotalGain_line.Visible = False
1730      .txtTotalLoss_line.Visible = False
1740      .txtTotalLoss_Tot.Visible = False
1750      .txtTotalGain_Tot.Visible = False
1760      .txtTotalGainLoss_Tot_line.Visible = False
1770      .Cost2.Visible = False
1780      .pcash3.Visible = False
1790      .cost3.Visible = False
1800      .uniqueid_line.Visible = False
1810      .GroupHeader1.Visible = False
1820      .GroupFooter3.Visible = False
1830      .GroupFooter1.Visible = False
1840      .ReportFooter.Visible = False
1850      If gblnLegalName = True Then
            ' ** Column 4: legalname
1860        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1870      Else
            ' ** Column 3: shorname
1880        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1890      End If
1900      .AcctName_Legal.Visible = False
1910      .AcctName_Short.Visible = False
1920      .AcctName_lbl.Visible = True
1930    End With

EXITP:
1940    Exit Sub

ERRH:
1950    Select Case ERR.Number
        Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Sub

Private Sub Report_Close()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2010    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2110    varRetVal = Null
2120    With Me
2130      If Forms(strCallingForm).chkIncludeFootnote = True Then
2140        varRetVal = Forms(strCallingForm).CourtReports_Footnote
2150      End If
2160    End With

EXITP:
2170    ThisFootnote = varRetVal
2180    Exit Function

ERRH:
2190    varRetVal = RET_ERR
2200    Select Case ERR.Number
        Case Else
2210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2220    End Select
2230    Resume EXITP

End Function
