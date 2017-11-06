VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCheckReconcile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCheckReconcile"

'VGC 03/22/2017: CHANGES!

'Private Const strCallingForm As String = "frmCheckReconcile"

Private dblTABalance As Double, dblBSTot As Double, dblTATot As Double, dblOCTot As Double, dblDiff As Double
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           If InStr(Trim(.OpenArgs), "~") > 0 Then
                ' ** DoCmd.OpenReport strDocName, acViewPreview, , , , _
                ' **   CStr(dblTABalance) & "~" & CStr(dblBSTot) & "~" & _
                ' **   CStr(dblTATot) & "~" & CStr(dblOCTot) & "~" & CStr(dblDiff)

150             TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

                ' ** Default to PageNumOf if calling form has no check box.
160             .PageNumOf.Visible = True
170             .PageNum.Visible = False

180             strTmp01 = Trim(.OpenArgs)
190             intPos01 = InStr(strTmp01, "~")
200             dblTABalance = CDbl(Left(strTmp01, (intPos01 - 1)))
210             .TABalance.ControlSource = "=" & CStr(dblTABalance)
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             dblBSTot = CDbl(Left(strTmp01, (intPos01 - 1)))
250             .BSTotal.ControlSource = "=" & CStr(dblBSTot)
260             strTmp01 = Mid(strTmp01, (intPos01 + 1))
270             intPos01 = InStr(strTmp01, "~")
280             dblTATot = CDbl(Left(strTmp01, (intPos01 - 1)))
290             .TATotal.ControlSource = "=" & CStr(dblTATot)
300             strTmp01 = Mid(strTmp01, (intPos01 + 1))
310             intPos01 = InStr(strTmp01, "~")
320             dblOCTot = CDbl(Left(strTmp01, (intPos01 - 1)))
330             .OCTotal.ControlSource = "=" & CStr(dblOCTot)
340             dblDiff = CDbl(Mid(strTmp01, (intPos01 + 1)))
350             .DifferenceBSTA.ControlSource = "=" & CStr(dblDiff)

360             .AsOf.Caption = "Printed " & Format(Date, "mm/dd/yyyy")
370             .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
380             .AsOf_lbl2.Caption = Format(Now(), "mm/dd/yyyy hh:nn AM/PM")

390             If gstrReportCallingForm <> vbNullString Then
400               Forms(gstrReportCallingForm).Visible = False
410             End If

420           Else
430             Cancel = -1
440           End If
450         Else
460           Cancel = -1
470         End If
480       Else
490         Cancel = -1
500       End If
510     End With

520     If Cancel = -1 Then
530       Beep
540       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
550     End If

EXITP:
560     Exit Sub

ERRH:
570     If gstrReportCallingForm <> vbNullString Then
580       DoCmd.Restore
590       Forms(gstrReportCallingForm).Visible = True
600       DoCmd.SelectObject acForm, gstrReportCallingForm, False
610     End If
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

710     With Me
720       If .Page >= 2 Then
730         .Section(acPageHeader).Visible = True
740       Else
750         .Section(acPageHeader).Visible = False
760       End If
770     End With

EXITP:
780     Exit Sub

ERRH:
790     If gstrReportCallingForm <> vbNullString Then
800       DoCmd.Restore
810       Forms(gstrReportCallingForm).Visible = True
820       DoCmd.SelectObject acForm, gstrReportCallingForm, False
830     End If
840     Select Case ERR.Number
        Case Else
850       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
860     End Select
870     Resume EXITP

End Sub

Private Sub Report_Close()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

910     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

920     DoCmd.Restore

930     If gstrReportCallingForm <> vbNullString Then
940       Forms(gstrReportCallingForm).Visible = True
950     End If

EXITP:
960     Exit Sub

ERRH:
970     If gstrReportCallingForm <> vbNullString Then
980       DoCmd.Restore
990       Forms(gstrReportCallingForm).Visible = True
1000      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1010    End If
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub
