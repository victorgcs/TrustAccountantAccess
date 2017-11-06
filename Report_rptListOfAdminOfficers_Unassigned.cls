VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfAdminOfficers_Unassigned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfAdminOfficers_Unassigned"

'VGC 11/27/2016: CHANGES!

Private Const strCallingForm As String = "frmAdminOfficer"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         .AsOf.Caption = "As of " & Format(Date, "mmmm dd, yyyy")
170         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

180         strOrderBy = Forms(strCallingForm).frmAdminOfficer_Sub.Form.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
            ' ** Potential Sorts:
            ' **   [adminno]
            ' **   [officer], [accountno]
            ' **   [assigned], [officer]
            ' **   [accountno]
            ' **   [shortname]
            ' **   [alphasort]
190         .OrderBy = strOrderBy
200         .OrderByOn = True

210         .Caption = "List_of_Administrative_Officers_Unassigned_" & Format(Date, "yymmdd")

220       Else
230         Cancel = -1
240       End If
250     End With

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     Forms(strCallingForm).Visible = True
320     DoCmd.Restore
330     Select Case ERR.Number
        Case Else
340       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
350     End Select
360     Resume EXITP

End Sub

Private Sub Report_Close()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

410     DoCmd.Restore

420     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

430     If Forms(strCallingForm).Visible = False Then
440       Forms(strCallingForm).Visible = True
450       DoCmd.SelectObject acForm, strCallingForm, False
460       DoCmd.Restore
470     End If

EXITP:
480     Exit Sub

ERRH:
490     Forms(strCallingForm).Visible = True
500     DoCmd.Restore
510     Select Case ERR.Number
        Case Else
520       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
530     End Select
540     Resume EXITP

End Sub
