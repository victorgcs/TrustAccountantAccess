VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptFL_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptFL_07"

'VGC 01/19/2013: CHANGES!

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
640         .AccountHead_lbl.Caption = Forms(strCallingForm).cmdPreview05_Sch_lbl.Caption & " by " & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Guardian of Property", "Personal Representative")

            ' ** Set the bottom-left report version footer.
650         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

660         strFilter = "accountno = '" & gstrAccountNo & "'"
670         .Filter = strFilter
680         .FilterOn = True

690         .Caption = "CourtReport_FL_Assets_on_Hand_at_Close_of_Accounting_Period_" & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Grdn", "Rep") & "_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

700       Else
710         Cancel = -1
720       End If
730     End With

740     If Cancel = -1 Then
750       Beep
760       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
770     End If

EXITP:
780     Set ctl = Nothing
790     Exit Sub

ERRH:
800     Select Case ERR.Number
        Case Else
810       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
820     End Select
830     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim varCaseNum As Variant

910     With Me
920       varCaseNum = DLookup("[CaseNum]", "account", "[accountno] = '" & gstrAccountNo & "'")
930       If IsNull(varCaseNum) = False Then
940         .CaseNum = varCaseNum
950       End If
960     End With

EXITP:
970     Exit Sub

ERRH:
980     Select Case ERR.Number
        Case Else
990       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1000    End Select
1010    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

        Dim intPos01 As Integer
        Dim strTmp01 As String

1110    With Me
1120      .NoData_lbl.BackColor = CLR_WHT
1130      .NoData_lbl.Visible = True
1140      .GroupHeader2.Visible = False
1150      .GroupFooter1.Visible = False
1160      .GroupFooter2.Visible = False
1170      If gblnLegalName = True Then
            ' ** Column 4: legalname
1180        strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1190      Else
            ' ** Column 3: shorname
1200        strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1210      End If
1220      intPos01 = InStr(strTmp01, "&")
1230      Do While intPos01 > 0
1240        strTmp01 = Left(strTmp01, intPos01) & Mid(strTmp01, intPos01)  ' ** Double the ampersand.
1250        intPos01 = InStr((intPos01 + 2), strTmp01, "&")
1260      Loop
1270      .AcctName_lbl.Caption = strTmp01
1280      .AcctName_Legal.Visible = False
1290      .AcctName_Short.Visible = False
1300      .AcctName_lbl.Visible = True
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
