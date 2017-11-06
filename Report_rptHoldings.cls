VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptHoldings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptHoldings"

'VGC 09/19/2015: CHANGES!

Private dblTotalMarketValue As Double
Private Const strCallingForm As String = "frmRpt_Holdings"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim strFilter As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Set frm = Forms(strCallingForm)

150         strFilter = vbNullString

            ' ** Output to Word doesn't have a way to send a 'WHERE' clause.
160         If frm.IsWord = True Then
170           If frm.opgAssetNumber = frm.opgAssetNumber_optSpecified.OptionValue Then
180             If IsNull(frm.cmbAssets) = False Then
190               strFilter = "[assetno] = " & CStr(frm.cmbAssets)
200               .Filter = strFilter
210               .FilterOn = True
220             End If
230           End If
240         End If

250         Select Case frm.chkPageOf
            Case True
260           .PageNumOf.Visible = True
270           .PageNum.Visible = False
280         Case False
290           .PageNum.Visible = True
300           .PageNumOf.Visible = False
310           .PageNumOf.ControlSource = vbNullString
320         End Select

330         dblTotalMarketValue = 0#
340         Select Case frm.chkTotalMarketValue
            Case True
350           Select Case frm.chkTotalMarketValue.Enabled
              Case True
360             Select Case IsNull(frm.TotalMarketValue)
                Case True
370               .TotalMarketValue.ControlSource = "=0"
380               .NetAssetValue.Visible = False
390               .Section("ReportFooter").Visible = False
400             Case False
410               dblTotalMarketValue = frm.TotalMarketValue
420               If dblTotalMarketValue <> 0# Then
430                 .TotalMarketValue.ControlSource = "=" & CStr(dblTotalMarketValue)
440                 .NetAssetValue.Visible = True
450                 .Section("ReportFooter").Visible = True
460               Else
470                 .TotalMarketValue.ControlSource = "=0"
480                 .NetAssetValue.Visible = False
490                 .Section("ReportFooter").Visible = False
500               End If
510             End Select
520           Case False
530             .TotalMarketValue.ControlSource = "=0"
540             .NetAssetValue.Visible = False
550             .Section("ReportFooter").Visible = False
560           End Select
570         Case False
580           .TotalMarketValue.ControlSource = "=0"
590           .NetAssetValue.Visible = False
600           .Section("ReportFooter").Visible = False
610         End Select

            'TotalShareFace_display
            '=CStr(IIf([TotalShareFace]=CLng([TotalShareFace]),Format([TotalShareFace],'#,##0'),Format([TotalShareFace],'#,##0.####')))
            'TotalShareFace_Tot
            '=CStr(IIf(Sum([Totalshareface])=CLng(Sum([Totalshareface])),Format(Sum([Totalshareface]),'#,##0'),Format(Sum([Totalshareface]),'#,##0.####')))

620         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
630         .AsOf2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")
640         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

650         Select Case frm.opgAssetNumber
            Case frm.opgAssetNumber_optSpecified.OptionValue
660           .Caption = "Holdings_" & Format(Date, "yymmdd")
670         Case frm.opgAssetNumber_optAll.OptionValue
680           .Caption = "Holdings_All_" & Format(Date, "yymmdd")
690         End Select

700       Else
710         Cancel = -1
720       End If
730     End With

740     If Cancel = -1 Then
750       Beep
760       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
770     End If

EXITP:
780     Set frm = Nothing
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

910     With Me
920       If .Page >= 2 Then
930         .Section(acPageHeader).Visible = True
940       Else
950         .Section(acPageHeader).Visible = False
960       End If
970     End With

EXITP:
980     Exit Sub

ERRH:
990     Select Case ERR.Number
        Case Else
1000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1010    End Select
1020    Resume EXITP

End Sub

Private Sub Report_Close()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1110    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1120    Exit Sub

ERRH:
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub
