VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_00B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_00B"

'VGC 03/11/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_CA"

' ** cmbAccounts combo box constants:
Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname

Private frm As Access.Form
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
630               .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
640               .AcctName_Legal.Visible = True
650               .AcctName_Short.Visible = False
660             End If
670           End If
680         End If

            ' ** Set the date range heading.
690         .AsOf_lbl.Caption = "As of " & Format(gdatStartDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
700         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Leave off the Schedule.
710         .AccountHead_lbl.Caption = vbNullString

            ' ** Set the date and account filter.
720         strFilter = "[accountno] = '" & gstrAccountNo & "'"
730         .Filter = strFilter
740         .FilterOn = True

750         .Caption = "CourtReport_CA_Property_on_Hand_at_Beginning_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

760       Else
770         Cancel = -1
780       End If
790     End With

800     If Cancel = -1 Then
810       Beep
820       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
830     End If

EXITP:
840     Set ctl = Nothing
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

        Dim intPos01 As Integer
        Dim strTmp01 As String

910     With Me
920       .NoData_lbl.BackColor = CLR_WHT
930       .NoData_lbl.Visible = True  ' ** Turn on the message.
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
1220    Set frm = Nothing
1230    Exit Sub

ERRH:
1240    Select Case ERR.Number
        Case Else
1250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1260    End Select
1270    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

1310    varRetVal = Null
1320    With Me
1330      If frm.chkIncludeFootnote = True Then
1340        varRetVal = frm.CourtReports_Footnote
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
