VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfAdminOfficers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfAdminOfficers"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmAdminOfficer"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String
        Dim varTmp00 As Variant

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
170         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

180         strOrderBy = Forms(strCallingForm).frmAdminOfficer_Sub.Form.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
            ' ** Potential Sorts:
            ' **   [adminno]
            ' **   [officer], [accountno]
            ' **   [assigned], [officer]
            ' **   [accountno]
            ' **   [shortname]
190         .OrderBy = strOrderBy
200         .OrderByOn = True

210         .Caption = "List_of_Administrative_Officers_" & Format(Date, "yymmdd")

220         If Forms(strCallingForm).Unassigned_lbl.Visible = True Then
230           varTmp00 = DLookup("[cnt]", "qryAdminOfficer_05", "[adminno] = 1")
240           If varTmp00 = 1 Then
250             .Unassigned.ControlSource = "='1 Account Is Unassigned'"
260           Else
270             .Unassigned.ControlSource = "='" & CStr(varTmp00) & " Accounts Are Unassigned'"
280           End If
290         End If

300       Else
310         Cancel = -1
320       End If
330     End With

340     If Cancel = -1 Then
350       Beep
360       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
370     End If

EXITP:
380     Exit Sub

ERRH:
390     Forms(strCallingForm).Visible = True
400     DoCmd.Restore
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Sub

Private Sub Report_Close()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

510     DoCmd.Restore

520     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

530     If Forms(strCallingForm).Visible = False Then
540       Forms(strCallingForm).Visible = True
550       DoCmd.SelectObject acForm, strCallingForm, False
560       DoCmd.Restore
570     End If

EXITP:
580     Exit Sub

ERRH:
590     Forms(strCallingForm).Visible = True
600     DoCmd.Restore
610     Select Case ERR.Number
        Case Else
620       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
630     End Select
640     Resume EXITP

End Sub
