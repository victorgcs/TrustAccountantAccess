VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAccountSummary_ForEx_Sub_Decreases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAccountSummary_ForEx_Sub_Decreases"

'VGC 03/22/2017: CHANGES!

'Private Const strCallingForm As String = "frmStatementParameters"

Private strCallingReport As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120   On Error Resume Next
130       strCallingReport = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingReport = "rptAccountSummary_ForEx" Then

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

Private Sub Report_NoData(Cancel As Integer)
' ** THIS DOESN'T TRIGGER!

300   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

310     With Me
320       .Parent.rptAccountSummary_ForEx_Sub_Decreases_lbl.Visible = True
330       .Parent.NoData_Decreases_lbl.Visible = True
340       .Parent.TotalDecreases.Visible = True
350     End With

EXITP:
360     Exit Sub

ERRH:
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Sub
