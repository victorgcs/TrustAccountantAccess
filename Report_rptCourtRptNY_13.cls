VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_13"

'VGC 01/19/2013: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_NY"
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

220         .AsOf_lbl.Caption = Format(Date, "mmmm dd, yyyy")
230         .AsOf_lbl2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

240       Else
250         Cancel = -1
260       End If
270     End With

280     If Cancel = -1 Then
          ' ** Just for the record.
290       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
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

EXITP:
420     Exit Sub

ERRH:
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub
