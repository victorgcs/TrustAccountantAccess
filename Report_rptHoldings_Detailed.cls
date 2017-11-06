VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptHoldings_Detailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptHoldings_Detailed"

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
160         Select Case frm.IsWord
            Case True
170           If frm.opgAssetNumber = frm.opgAssetNumber_optSpecified.OptionValue Then
180             If IsNull(frm.cmbAssets) = False Then
190               strFilter = "[assetno] = " & CStr(frm.cmbAssets)
200               .Filter = strFilter
210               .FilterOn = True
220             Else
230               .Filter = strFilter
240               .FilterOn = False
250             End If
260           Else
270             .Filter = strFilter
280             .FilterOn = False
290           End If
300         Case False
310           If frm.opgAssetNumber = frm.opgAssetNumber_optAll.OptionValue Then
320             .Filter = strFilter
330             .FilterOn = False
340           End If
350         End Select

360         Select Case frm.chkPageOf
            Case True
370           .PageNumOf.Visible = True
380           .PageNum.Visible = False
390         Case False
400           .PageNum.Visible = True
410           .PageNumOf.Visible = False
420           .PageNumOf.ControlSource = vbNullString
430         End Select

440         dblTotalMarketValue = 0#
450         Select Case frm.chkTotalMarketValue
            Case True
460           Select Case frm.chkTotalMarketValue.Enabled
              Case True
470             Select Case IsNull(frm.TotalMarketValue)
                Case True
480               .TotalMarketValue.ControlSource = "=0"
490               .NetAssetValue.Visible = False
500               .Section("ReportFooter").Visible = False
510             Case False
520               dblTotalMarketValue = frm.TotalMarketValue
530               If dblTotalMarketValue <> 0# Then
540                 .TotalMarketValue.ControlSource = "=" & CStr(dblTotalMarketValue)
550                 .NetAssetValue.Visible = True
560                 .Section("ReportFooter").Visible = True
570               Else
580                 .TotalMarketValue.ControlSource = "=0"
590                 .NetAssetValue.Visible = False
600                 .Section("ReportFooter").Visible = False
610               End If
620             End Select
630           Case False
640             .TotalMarketValue.ControlSource = "=0"
650             .NetAssetValue.Visible = False
660             .Section("ReportFooter").Visible = False
670           End Select
680         Case False
690           .TotalMarketValue.ControlSource = "=0"
700           .NetAssetValue.Visible = False
710           .Section("ReportFooter").Visible = False
720         End Select

            'TotalShareFace_display
            '=CStr(IIf([TotalShareFace]=CLng([TotalShareFace]),Format([TotalShareFace],'#,##0'),Format([TotalShareFace],'#,##0.####')))
            'TotalShareFace_SubTot
            '=CStr(IIf(Sum([Totalshareface])=CLng(Sum([Totalshareface])),Format(Sum([Totalshareface]),'#,##0'),Format(Sum([Totalshareface]),'#,##0.####')))
            'TotalShareFace_Tot
            '=CStr(IIf(Sum([Totalshareface])=CLng(Sum([Totalshareface])),Format(Sum([Totalshareface]),'#,##0'),Format(Sum([Totalshareface]),'#,##0.####')))

730         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
740         .AsOf2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")
750         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

760         Select Case frm.opgAssetNumber
            Case frm.opgAssetNumber_optSpecified.OptionValue
770           .Caption = "Holdings_Detailed_" & Format(Date, "yymmdd")
780         Case frm.opgAssetNumber_optAll.OptionValue
790           .Caption = "Holdings_Detailed_All_" & Format(Date, "yymmdd")
800         End Select

810       Else
820         Cancel = -1
830       End If
840     End With

850     If Cancel = -1 Then
860       Beep
870       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
880     End If

EXITP:
890     Set frm = Nothing
900     Exit Sub

ERRH:
910     Select Case ERR.Number
        Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

1010    With Me
1020      If .Page >= 2 Then
1030        .Section(acPageHeader).Visible = True
1040      Else
1050        .Section(acPageHeader).Visible = False
1060      End If
1070    End With

EXITP:
1080    Exit Sub

ERRH:
1090    Select Case ERR.Number
        Case Else
1100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1110    End Select
1120    Resume EXITP

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
