VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptFL_06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptFL_06"

'VGC 04/04/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_FL"

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
460           .AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height
470           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
480           .CaseNum.Top = .CaseNum.Top + lngLegalNameMove_Height
              '.CaseNum_lbl.Top = .CaseNum_lbl.Top + lngLegalNameMove_Height
490           For Each ctl In .Section(acPageHeader).Controls
500             With ctl
510               Select Case .ControlType
                  Case acLabel
520                 Select Case .Name
                    Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                      ' ** Ignore.
530                 Case Else
540                   .Top = .Top + lngLegalNameMove_Height
550                 End Select
560               Case acLine
570                 If .Width <> Me.Width Then
580                   .Top = .Top + lngLegalNameMove_Height
590                 End If
600               End Select
610             End With
620           Next
630           .AcctName_Legal.Visible = True
640           .AcctName_Short.Visible = False
650         End If

            ' ** Set the date range heading.
660         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & _
              Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the ordinal and version heading.
670         .AccountHead_lbl.Caption = "Schedule C by " & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Guardian of Property", "Personal Representative")

            ' ** Set the bottom-left report version footer.
680         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

690         strFilter = "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# AND " & _
              "transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) "
700         strFilter = strFilter & "AND accountno = '" & gstrAccountNo & "' "
710         strFilter = strFilter & "AND ((journaltype = 'Paid' AND pcash <> 0 AND taxcode = " & CStr(glngTaxCode_Distribution) & ") "  '= "Distribution"
720         strFilter = strFilter & "OR (journaltype = 'Paid' AND icash <> 0 AND taxcode = " & CStr(glngTaxCode_Distribution) & ") "  '= "Distribution"
730         strFilter = strFilter & "OR (journaltype = 'Withdrawn' AND taxcode = " & CStr(glngTaxCode_Distribution) & "))"  '= "Distribution"
740         .Filter = strFilter    '####  TAXCODE  ####
750         .FilterOn = True

760         .Caption = "CourtReport_FL_Distributions_" & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, "Grdn", "Rep") & "_" & _
              gstrAccountNo & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

770       Else
780         Cancel = -1
790       End If
800     End With

810     If Cancel = -1 Then
820       Beep
830       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
840     End If

EXITP:
850     Set ctl = Nothing
860     Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

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
1250      .Amount_Inc.Visible = False
1260      .Amount_Prin.Visible = False
1270      .txtAmountTotalSum_Inc_lbl.Visible = False
1280      .txtAmountTotalSum_Inc.Visible = False
1290      .txtAmountTotalSum_Prin.Visible = False
1300      .txtAmountTotalSum_Inc_line.Visible = False
1310      .txtAmountTotalSum_Prin_line.Visible = False
1320      If gblnLegalName = True Then
            ' ** Column 4: legalname
1330        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1340      Else
            ' ** Column 3: shorname
1350        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1360      End If
1370      .AcctName_Legal.Visible = False
1380      .AcctName_Short.Visible = False
1390      .AcctName_lbl.Visible = True
1400    End With

EXITP:
1410    Exit Sub

ERRH:
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

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
