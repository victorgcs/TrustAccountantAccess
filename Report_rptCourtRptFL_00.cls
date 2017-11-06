VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptFL_00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptFL_00"

'VGC 01/19/2013: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_FL"
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
640         .AccountHead_lbl.Caption = Forms(strCallingForm).Ordinal & " And " & _
              Forms(strCallingForm).Version & " Account by " & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Guardian of Property", "Personal Representative")

            ' ** Set the bottom-left report version footer.
650         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

660         .Caption = "CourtReport_FL_Summary_" & _
              IIf(Forms(strCallingForm).opgType = Forms(strCallingForm).opgType_optGuard.OptionValue, _
              "Grdn", "Rep") & "_" & gstrAccountNo & "_" & _
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

Private Sub Report_Activate()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Activate"

910     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

EXITP:
920     Exit Sub

ERRH:
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1010    With Me
1020      .NoData_lbl.BackColor = CLR_WHT
1030      .NoData_lbl.Visible = True
1040      If gblnLegalName = True Then
            ' ** Column 4: legalname
1050        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1060      Else
            ' ** Column 3: shorname
1070        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1080      End If
1090      .AcctName_Legal.Visible = False
1100      .AcctName_Short.Visible = False
1110      .AcctName_lbl.Visible = True
1120    End With

EXITP:
1130    Exit Sub

ERRH:
1140    Select Case ERR.Number
        Case Else
1150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1160    End Select
1170    Resume EXITP

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
