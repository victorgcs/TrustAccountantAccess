VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptPOSPay_Export_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptPOSPay_Export_02"

'VGC 12/24/2016: CHANGES!

Private strCallingForm As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me

120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then
150           intPos01 = InStr(strTmp01, "~")
160           strCallingForm = Left(strTmp01, (intPos01 - 1))
170           strOrderBy = Mid(strTmp01, (intPos01 + 1))
180         End If
190       End If

200       TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

          ' ** Default to PageNumOf if calling form has no check box.
210       .PageNumOf.Visible = True
220       .PageNum.Visible = False

230       .Caption = "POSPay_Export_File_Current_" & Format(Date, "yymmdd")

240       If strCallingForm <> vbNullString Then
250         If strOrderBy = vbNullString Then
260           strOrderBy = Forms(strCallingForm).frmCheckPOSPay_Sub3.Form.SortNow_Get  ' ** Form Function: frmCheckPOSPay_Sub1.
270         End If
            ' ** Potential Sorts:
            ' **   [Check_Number], [Issue_Date], [Account]
            ' **   [Issue_Type], [Issue_Date], [Account]
            ' **   [Issue_Date], [Issue_Date], [Account]
            ' **   [Amount], [Issue_Date], [Account]
            ' **   [Payee_Name], [Issue_Date], [Account]
            ' **   [ABA_TRC], [Issue_Date], [Account]
            ' **   [Account], [Issue_Date], [Account]
280         .OrderBy = strOrderBy
290         .OrderByOn = True
300       End If

310     End With

EXITP:
320     Exit Sub

ERRH:
330     Select Case ERR.Number
        Case Else
340       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
350     End Select
360     Resume EXITP

End Sub

Private Sub Report_Close()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

410     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

420     If strCallingForm <> vbNullString Then
430       gblnSetFocus = True
440       DoCmd.SelectObject acForm, strCallingForm, False
450       Forms(strCallingForm).TimerInterval = 100&
460     End If

EXITP:
470     Exit Sub

ERRH:
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub
