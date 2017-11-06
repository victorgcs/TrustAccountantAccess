VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptStatementOfCondition_Sub_Liability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptStatementOfCondition_Sub_Liability"

'VGC 04/24/2014: CHANGES!

Private strCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "rptStatementOfCondition" Then
160         If IsLoaded("frmRpt_StatementOfCondition", acForm) = True Then  ' ** Module Function: modFileUtilities.

170           Set frm = Forms("frmRpt_StatementOfCondition")
180           Select Case frm.opgAccountType
              Case frm.opgAccountType_optAll.OptionValue      ' ** For All.
190             Select Case frm.chkIncludeClosed
                Case True
200               If .RecordSource <> "qryStatementOfCondition_11c" Then
210                 .RecordSource = "qryStatementOfCondition_11c"
220               End If
230             Case False
240               If .RecordSource <> "qryStatementOfCondition_11f" Then
250                 .RecordSource = "qryStatementOfCondition_11f"
260               End If
270             End Select
280           Case frm.opgAccountType_optDisc.OptionValue     ' ** For descretion = Yes.
290             Select Case frm.chkIncludeClosed
                Case True
300               If .RecordSource <> "qryStatementOfCondition_11b" Then
310                 .RecordSource = "qryStatementOfCondition_11b"
320               End If
330             Case False
340               If .RecordSource <> "qryStatementOfCondition_11e" Then
350                 .RecordSource = "qryStatementOfCondition_11e"
360               End If
370             End Select
380           Case frm.opgAccountType_optNonDisc.OptionValue  ' ** For descretion = No.
390             Select Case frm.chkIncludeClosed
                Case True
400               If .RecordSource <> "qryStatementOfCondition_11a" Then
410                 .RecordSource = "qryStatementOfCondition_11a"
420               End If
430             Case False
440               If .RecordSource <> "qryStatementOfCondition_11d" Then
450                 .RecordSource = "qryStatementOfCondition_11d"
460               End If
470             End Select
480           End Select

490         Else
500           Cancel = -1
510         End If
520       Else
530         Cancel = -1
540       End If
550     End With

560     If Cancel = -1 Then
570       Beep
580       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
590     End If

EXITP:
600     Set frm = Nothing
610     Exit Sub

ERRH:
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub
