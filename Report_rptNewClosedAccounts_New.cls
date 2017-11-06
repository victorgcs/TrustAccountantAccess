VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptNewClosedAccounts_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptNewClosedAccounts_New"

'VGC 05/03/2015: CHANGES!

Private Const strCallingForm As String = "frmRpt_NewClosedAccounts"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

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

220         .AsOf_lbl.Caption = "From " & Format(CDate(Forms(strCallingForm).DateStart), "mm/dd/yyyy") & " To " & _
              Format(CDate(Forms(strCallingForm).DateEnd), "mm/dd/yyyy")
230         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

240         .Caption = "New_Accounts_" & Format(Forms(strCallingForm).DateStart, "yymmdd") & "_To_" & _
              Format(Forms(strCallingForm).DateEnd, "yymmdd")

250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
320     End If

EXITP:
330     Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Report_Close()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

410     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub
