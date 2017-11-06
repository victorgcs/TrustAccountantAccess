VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptUnrealizedGainAndLoss_SortName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptUnrealizedGainAndLoss_SortName"

'VGC 10/08/2017: CHANGES!

Private Const strCallingForm As String = "frmRpt_UnrealizedGainAndLoss"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strAccountNo As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

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

220         Select Case Forms(strCallingForm).opgAccountNumber
            Case Forms(strCallingForm).opgAccountNumber_optSpecified.OptionValue
230           If .RecordSource <> "qryRpt_UnrealizedGainAndLoss_01" Then
240             .RecordSource = "qryRpt_UnrealizedGainAndLoss_01"
250           End If
260           strAccountNo = Forms(strCallingForm).cmbAccounts
270           .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
280         Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue
290           .RecordSource = "qryRpt_UnrealizedGainAndLoss_02"
300           strAccountNo = "All"
310           .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
320         End Select

330         .AsOf_lbl.Caption = "As of " & Format(CDate(Forms(strCallingForm).DateAsOf), "mmmm dd"", ""yyyy")

340         .AsOf.Caption = "Printed " & Format(Date, "mmmm dd"", ""yyyy")
350         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
360         .AsOf_lbl2.Caption = "As of " & Format(CDate(Forms(strCallingForm).DateAsOf), "mmmm dd"", ""yyyy")
370         .Caption = "Unrealized_Gains_and_Losses_" & strAccountNo & "_" & Format(Forms(strCallingForm).DateAsOf, "yymmdd")

380       Else
390         Cancel = -1
400       End If
410     End With

420     If Cancel = -1 Then
430       Beep
440       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
450     End If

EXITP:
460     Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

610     With Me

620       Select Case IsNull(.currentDate)
          Case True
630         If gdatStartDate = 0 Then
640           .MarketAsOf_lbl.Caption = vbNullString
650           .MarketAsOf_lbl2.Caption = vbNullString
660         ElseIf CDate(Forms(strCallingForm).DateAsOf) = gdatStartDate Then
670           .MarketAsOf_lbl.Caption = vbNullString
680           .MarketAsOf_lbl2.Caption = vbNullString
690         Else
700           .MarketAsOf_lbl.Caption = "Market as of " & Format(gdatStartDate, "mm/dd/yyyy")
710           .MarketAsOf_lbl2.Caption = "Market as of " & Format(gdatStartDate, "mm/dd/yyyy")
720         End If
730       Case False
740         .MarketAsOf_lbl.Caption = "Market as of " & Format(.currentDate, "mm/dd/yyyy")
750         .MarketAsOf_lbl2.Caption = "Market as of " & Format(.currentDate, "mm/dd/yyyy")
760       End Select

770     End With

EXITP:
780     Exit Sub

ERRH:
790     Select Case ERR.Number
        Case Else
800       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
810     End Select
820     Resume EXITP

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
