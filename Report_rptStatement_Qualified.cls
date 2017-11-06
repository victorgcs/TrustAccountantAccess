VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptStatement_Qualified"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptStatement_Qualified"

'VGC 08/16/2012: CHANGES!

Private Const strCallingForm As String = "frmStatementParameters"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strMonth As String, strYear As String
        Dim strTmp01 As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         strMonth = Forms(strCallingForm).cmbMonth
170         strYear = Forms(strCallingForm).StatementsYear
180         .AsOf_lbl_txt.Caption = strMonth & ", " & strYear
190         .ThisMonth_lbl.Left = .Controls("smt" & Left(strMonth, 3) & "_lbl").Left
200         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
210         strTmp01 = LCase$(Left(strMonth, 3)) & Right(strYear, 2)
220         .Caption = "Statement_Accounts_" & strTmp01 & "_" & Format(Date, "yymmdd")

230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

End Sub

Private Sub Report_Close()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

410     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

420     DoCmd.Hourglass False  ' ** This seems to remain on!

EXITP:
430     Exit Sub

ERRH:
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Sub
