VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptStatementOfCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptStatementOfCondition"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_StatementOfCondition"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strHeader As String, strScope As String

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

220         strHeader = "Statement of Condition"

230         .CoName.Caption = gstrCo_Name
240         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
250         .AsOf_lbl.Caption = "As of " & Format(Forms(strCallingForm).DateAsOf, "mmmm dd, yyyy")
260         .AsOf.Caption = "Printed " & Format(Date, "mmmm dd"", ""yyyy")

270         Select Case Forms(strCallingForm).opgAccountType
            Case Forms(strCallingForm).opgAccountType_optAll.OptionValue
              ' ** Leave the header as-is.
280           strScope = vbNullString
290           .AccountScope_lbl.Caption = "All Accounts"
300         Case Forms(strCallingForm).opgAccountType_optDisc.OptionValue
310           strScope = "Discretionary"
320           strHeader = strHeader & " - " & strScope
330           .AccountScope_lbl.Caption = strScope
340         Case Forms(strCallingForm).opgAccountType_optNonDisc.OptionValue
350           strScope = "Non-Discretionary"
360           strHeader = strHeader & " - " & strScope
370           .AccountScope_lbl.Caption = strScope
380           strScope = "NonDiscretionary"
390         End Select
400         .Header_lbl.Caption = strHeader

410         If strScope = vbNullString Then
420           .Caption = "Statement_of_Condition_" & Format(Forms(strCallingForm).DateAsOf, "yymmdd")
430         Else
440           .Caption = "Statement_of_Condition_" & strScope & "_" & Format(Forms(strCallingForm).DateAsOf, "yymmdd")
450         End If

460       Else
470         Cancel = -1
480       End If
490     End With

500     If Cancel = -1 Then
510       Beep
520       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
530     End If

EXITP:
540     Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Report_Close()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

610     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
620     Exit Sub

ERRH:
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub
