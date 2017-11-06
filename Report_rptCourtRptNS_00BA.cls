VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_00BA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_00BA"

'VGC 01/19/2013: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_NS"
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
                    Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl"
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
630         strFilter = "accountno = '" & gstrAccountNo & "'"
640         .Filter = strFilter
650         .FilterOn = True

660         .Caption = "CourtReport_NS_Property_on_Hand_at_Beginning_of_Account_Period_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

670       Else
680         Cancel = -1
690       End If
700     End With

710     If Cancel = -1 Then
720       Beep
730       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
740     End If

EXITP:
750     Set ctl = Nothing
760     Exit Sub

ERRH:
770     Select Case ERR.Number
        Case Else
780       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
790     End Select
800     Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

        Dim intPos01 As Integer
        Dim strTmp01 As String

910     With Me
920       .NoData_lbl.BackColor = CLR_WHT
930       .NoData_lbl.Visible = True
940       .GroupHeader2.Visible = False
950       .GroupFooter1.Visible = False
960       .GroupFooter2.Visible = False
970       If gblnLegalName = True Then
            ' ** Column 4: legalname
980         strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
990       Else
            ' ** Column 3: shorname
1000        strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1010      End If
1020      intPos01 = InStr(strTmp01, "&")
1030      Do While intPos01 > 0
1040        strTmp01 = Left(strTmp01, intPos01) & Mid(strTmp01, intPos01)  ' ** Double the ampersand.
1050        intPos01 = InStr((intPos01 + 2), strTmp01, "&")
1060      Loop
1070      .AcctName_lbl.Caption = strTmp01
1080      .AcctName_Legal.Visible = False
1090      .AcctName_Short.Visible = False
1100      .AcctName_lbl.Visible = True
1110    End With

EXITP:
1120    Exit Sub

ERRH:
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub

Private Sub Report_Close()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1210    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1220    Exit Sub

ERRH:
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

1310    varRetVal = Null
1320    With Me
1330      If Forms(strCallingForm).chkIncludeFootnote = True Then
1340        varRetVal = Forms(strCallingForm).CourtReports_Footnote
1350      End If
1360    End With

EXITP:
1370    ThisFootnote = varRetVal
1380    Exit Function

ERRH:
1390    varRetVal = RET_ERR
1400    Select Case ERR.Number
        Case Else
1410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1420    End Select
1430    Resume EXITP

End Function
