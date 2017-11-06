VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_06_Sub1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_06_Sub1"

'VGC 01/18/2017: CHANGES!
'FC 11/27/2011: CHANGES!

Private strCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "rptCourtRptNY_06" Then

160       Else
170         Cancel = -1
180       End If
190     End With

200     If Cancel = -1 Then
210       Beep
220       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
230     End If

EXITP:
240     Exit Sub

ERRH:
250     Select Case ERR.Number
        Case Else
260       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
270     End Select
280     Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

310     With Me

320   On Error Resume Next  ' ** Allow processing of errors inline.
330       If .totdesc = vbNullString Then  ' ** If there is no data in the report force an error.
340         If ERR.Number = 2427 Then  ' ** This means the report field is empty.
350           .transdate.Visible = False  ' ** Make everything invisible that might show an error.
360           .shareface.Visible = False
370           .totdesc.Visible = False
380           .amount.Visible = False
390           .AmountTotal.Visible = False
400           .AmountTotal_line.Visible = False
410           .NoData_lbl.Visible = True  ' ** Turn on the message.
420           .NoData_lbl.BackColor = CLR_WHT
430         End If
440       End If

450     End With

EXITP:
460     Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

EXITP:
610     Exit Sub

ERRH:
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub
