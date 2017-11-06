VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_04A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_04A"

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
              "AND [date] <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) " & _
              "AND [accountno] = '" & gstrAccountNo & "' " & _
              "AND ([reportnumber] = " & CStr(CRPT_DISBURSEMENTS) & " AND [amount] <> 0)"
720         .Filter = strFilter
730         .FilterOn = True

740         .Caption = "CourtReport_CA_Disbursements_Grouped_" & gstrAccountNo & "_" & _
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
830     Exit Sub

ERRH:
840     Set ctl = Nothing
850     Select Case ERR.Number
        Case Else
860       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
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
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1040    End Select
1050    Resume EXITP

End Sub

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter2_Format"

1110    With Me
1120      If blnNoData Then
1130        .txtRevcodeTotal_lbl.Visible = False
1140      Else
1150        .txtRevcodeTotal_lbl.Caption = "Total " & .txtRevcode_desc
1160      End If
1170    End With

EXITP:
1180    Exit Sub

ERRH:
1190    Select Case ERR.Number
        Case Else
1200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1210    End Select
1220    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1310    blnNoData = True

1320    With Me
1330      .NoData_lbl.BackColor = CLR_WHT
1340      .NoData_lbl.Visible = True  ' ** Turn on the message.
1350      .transdate.Visible = False
1360      .totdesc.Visible = False
1370      .Gain.Visible = False
1380      .txtAccountTotal_lbl.Visible = False
1390      .txtAccountTotal.Visible = False
1400      .txtAccountTotal_line.Visible = False
1410      .txtRevcodeTotal_lbl.Visible = False
1420      .GroupHeader2.Visible = False
1430      .GroupFooter1.Visible = False
1440      .GroupFooter2.Visible = False
1450      .txtRevcodeTotal.Visible = False
1460      .txtRevcodeTotal_line.Visible = False
1470      If gblnLegalName = True Then
            ' ** Column 4: legalname
1480        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1490      Else
            ' ** Column 3: shorname
1500        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1510      End If
1520      .AcctName_Legal.Visible = False
1530      .AcctName_Short.Visible = False
1540      .AcctName_lbl.Visible = True
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub Report_Close()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1710    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1720    Set frm = Nothing
1730    Exit Sub

ERRH:
1740    Select Case ERR.Number
        Case Else
1750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1760    End Select
1770    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

1810    varRetVal = Null
1820    With Me
1830      If frm.chkIncludeFootnote = True Then
1840        varRetVal = frm.CourtReports_Footnote
1850      End If
1860    End With

EXITP:
1870    ThisFootnote = varRetVal
1880    Exit Function

ERRH:
1890    varRetVal = RET_ERR
1900    Select Case ERR.Number
        Case Else
1910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1920    End Select
1930    Resume EXITP

End Function
