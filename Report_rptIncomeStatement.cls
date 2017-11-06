VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptIncomeStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptIncomeStatement"

'VGC 08/17/2012: CHANGES!

Private Const strCallingForm As String = "frmRpt_IncomeStatement"
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

220         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
230         .CoInfoBlock.Caption = gstrCo_InfoBlock

240         strAccountNo = Forms(strCallingForm).cmbAccounts
250         .Caption = "Income_Statement_" & strAccountNo & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & _
              Format(gdatEndDate, "yymmdd")

260         .Filter = "([transdate] Between #" & Format(CDate(gdatStartDate), "mm/dd/yyyy") & "# And " & _
              "#" & Format(CDate(gdatEndDate), "mm/dd/yyyy") & "#)"
            '.Filter = .Filter & " and (Not(Journaltype=""Purchase"" and icash < 0))"

            ' ** Get all accounts if we were passed ALL
270         If gstrAccountNo <> "ALL" Then
280           .Filter = .Filter & " and accountno = '" & gstrAccountNo & "'"
290         End If
300         .FilterOn = True

310         .OrderBy = "revcode_SORTORDER"
320         .OrderByOn = True

330       Else
340         Cancel = -1
350       End If
360     End With

370     If Cancel = -1 Then
380       Beep
390       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
400     End If

EXITP:
410     Exit Sub

ERRH:
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub GroupHeader1_Print(Cancel As Integer, PrintCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Print"

510   On Error Resume Next  ' ** Allow processing of errors inline.

520     With Me
530       If .totdesc = vbNullString Then  ' ** If there is no data in the report force an error.
540         If ERR.Number = 2427 Then  ' ** This means the report field is empty.
550           .transdate.Visible = False
560           .journaltype.Visible = False
570           .totdesc.Visible = False
580           .IPCash.Visible = False
590           .GroupHeader2.Visible = False
600           .GroupFooter1.Visible = False
610           .GroupFooter2.Visible = False
620           .GroupFooter3.Visible = False
630           .NoData_lbl.Visible = True
640           .NoData_lbl.BackColor = CLR_WHT
650         End If
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

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

810   On Error Resume Next  ' ** Allow processing of errors inline.

820     With Me
830       If .totdesc = vbNullString Then  ' ** If there is no data in the report force an error.
840         If ERR.Number = 2427 Then  ' ** This means the report field is empty.
              ' ** Ignore
850         End If
860       End If
870     End With

EXITP:
880     Exit Sub

ERRH:
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub

Private Sub Report_Close()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1010    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1020    Exit Sub

ERRH:
1030    Select Case ERR.Number
        Case Else
1040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1050    End Select
1060    Resume EXITP

End Sub
