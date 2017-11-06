VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCashControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCashControl"

'VGC 08/17/2012: CHANGES!

Private Const strCallingForm As String = "frmRpt_CashControl"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim lngDelta As Long

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

            '="As of " & [Forms]![Cash Control].[DateAsOf].Value
220         .AsOf_lbl.Caption = "as of " & Format(Forms(strCallingForm).DateAsOf, "mmmm dd, yyyy")
230         .AsOf.Caption = "Printed " & Format(Date, "mmmm dd, yyyy")
240         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
250         .Caption = "Cash_Control_" & Format(Forms(strCallingForm).DateAsOf, "yymmdd")

260         lngDelta = 2000&

270         If gblnSeparateCheckingAccounts = True Then
280           .Bank_Name_lbl.Visible = True
290           .Bank_Name_lbl_line.Visible = True
300           .Bank_Name.Visible = True
310           .Bank_AccountNumber_lbl.Visible = True
320           .Bank_AccountNumber_lbl_line.Visible = True
330           .Bank_AccountNumber.Visible = True
340         Else
350           .Bank_Name_lbl.Visible = False
360           .Bank_Name_lbl_line.Visible = False
370           .Bank_Name.Visible = False
380           .Bank_AccountNumber_lbl.Visible = False
390           .Bank_AccountNumber_lbl_line.Visible = False
400           .Bank_AccountNumber.Visible = False
410           .accountno_lbl.Left = .accountno_lbl.Left + lngDelta
420           .accountno_lbl_line.Left = .accountno_lbl_line.Left + lngDelta
430           .accountno.Left = .accountno.Left + lngDelta
440           .shortname_lbl.Left = .shortname_lbl.Left + lngDelta
450           .shortname_lbl_line.Left = .shortname_lbl_line.Left + lngDelta
460           .shortname.Left = .shortname.Left + lngDelta
470           .account_SWEEP_disp_lbl.Left = .account_SWEEP_disp_lbl.Left + lngDelta
480           .account_SWEEP_disp_lbl_line.Left = .account_SWEEP_disp_lbl_line.Left + lngDelta
490           .account_SWEEP_disp.Left = .account_SWEEP_disp.Left + lngDelta
500           .IncomeCash_lbl.Left = .IncomeCash_lbl.Left + lngDelta
510           .IncomeCash_lbl_line.Left = .IncomeCash_lbl_line.Left + lngDelta
520           .IncomeCash.Left = .IncomeCash.Left + lngDelta
530           .PrincipalCash_lbl.Left = .PrincipalCash_lbl.Left + lngDelta
540           .PrincipalCash_lbl_line.Left = .PrincipalCash_lbl_line.Left + lngDelta
550           .PrincipalCash.Left = .PrincipalCash.Left + lngDelta
560           .TotalCash_lbl.Left = .TotalCash_lbl.Left + lngDelta
570           .TotalCash_lbl_line.Left = .TotalCash_lbl_line.Left + lngDelta
580           .TotalCash.Left = .TotalCash.Left + lngDelta
590           .IncomeCashSum_line.Left = .IncomeCashSum_line.Left + lngDelta
600           .IncomeCashSum.Left = .IncomeCashSum.Left + lngDelta
610           .PrincipalCashSum_line.Left = .PrincipalCashSum_line.Left + lngDelta
620           .PrincipalCashSum.Left = .PrincipalCashSum.Left + lngDelta
630           .TotalCashSum_line.Left = .TotalCashSum_line.Left + lngDelta
640           .TotalCashSum.Left = .TotalCashSum.Left + lngDelta
650           .IncomeCashSum_lbl.Left = .IncomeCashSum_lbl.Left + lngDelta
660         End If

670       Else
680         Cancel = -1
690       End If
700     End With

710     If Cancel = -1 Then
720       Beep
730       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
740     End If

EXITP:
750     Exit Sub

ERRH:
760     Select Case ERR.Number
        Case Else
770       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
780     End Select
790     Resume EXITP

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
