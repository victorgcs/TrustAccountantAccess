VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCapitalGainAndLoss_SortNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCapitalGainAndLoss_SortNumber"

'VGC 03/14/2014: CHANGES!

' ** Note: Because the HoldingPeriod_lbl caption is calculated,
' ** it must repeat within the Holding Period Header,
' ** rather than in the Page Header.

Private Const strCallingForm As String = "frmRpt_CapitalGainAndLoss"
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

220         Select Case Forms(strCallingForm).chkShowAllSales
            Case True
230           Select Case Forms(strCallingForm).opgAccountNumber
              Case Forms(strCallingForm).opgAccountNumber_optSpecified.OptionValue
240             .RecordSource = "qryCapitalGainsAndLoss_02c"
250           Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue
260             .RecordSource = "qryCapitalGainsAndLoss_04c"
270           End Select
280         Case False
290           Select Case Forms(strCallingForm).opgAccountNumber
              Case Forms(strCallingForm).opgAccountNumber_optSpecified.OptionValue
                ' ** qryCapitalGainsAndLoss_02a
300           Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue
310             .RecordSource = "qryCapitalGainsAndLoss_04a"
320           End Select
330         End Select

340         Select Case Forms(strCallingForm).opgAccountNumber
            Case Forms(strCallingForm).opgAccountNumber_optSpecified.OptionValue
350           strAccountNo = Forms(strCallingForm).cmbAccounts
360         Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue
370           strAccountNo = "All"
380         End Select

390         .AsOf_lbl.Caption = "From " & Format(CDate(Forms(strCallingForm).DateStart), "mm/dd/yyyy") & " To " & _
              Format(CDate(Forms(strCallingForm).DateEnd), "mm/dd/yyyy")
400         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
410         .AsOf_lbl2.Caption = "From " & Format(CDate(Forms(strCallingForm).DateStart), "mm/dd/yyyy") & " To " & _
              Format(CDate(Forms(strCallingForm).DateEnd), "mm/dd/yyyy")
420         .Caption = "Capital_Gains_and_Losses_" & strAccountNo & "_" & _
              Format(Forms(strCallingForm).DateStart, "yymmdd") & "_To_" & _
              Format(Forms(strCallingForm).DateEnd, "yymmdd")

430       Else
440         Cancel = -1
450       End If
460     End With

470     If Cancel = -1 Then
480       Beep
490       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
500     End If

EXITP:
510     Exit Sub

ERRH:
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

610     With Me
620       If .Page >= 2 Then
630         .Section(acPageHeader).Visible = True
640       Else
650         .Section(acPageHeader).Visible = False
660       End If
670     End With

EXITP:
680     Exit Sub

ERRH:
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Report_Close()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

810     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
820     Exit Sub

ERRH:
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub
