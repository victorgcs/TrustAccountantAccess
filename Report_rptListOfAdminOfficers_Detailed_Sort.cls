VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfAdminOfficers_Detailed_Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfAdminOfficers_Detailed_Sort"

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
180         .AsOf_lbl2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

190         strOrderBy = Forms(strCallingForm).frmAdminOfficer_Sub.Form.SortNow_Get  ' ** Form Function: frmAdminOfficer_Sub.
            ' ** Potential Sorts:
            ' **   [adminno]
            ' **   [officer], [accountno]
            ' **   [assigned], [officer]
            ' **   [accountno]
            ' **   [shortname]
200         .OrderBy = strOrderBy
210         .OrderByOn = True

220         If strOrderBy = "[accountno]" Or strOrderBy = "[shortname]" Then
230           .admin_count.Visible = False
240         End If

250         .Caption = "List_of_Administrative_Officers_Detailed_" & Format(Date, "yymmdd")

260         If Forms(strCallingForm).Unassigned_lbl.Visible = True Then
270           varTmp00 = DLookup("[cnt]", "qryAdminOfficer_05", "[adminno] = 1")
280           If varTmp00 = 1 Then
290             .Unassigned.ControlSource = "='1 Account Is Unassigned'"
300           Else
310             .Unassigned.ControlSource = "='" & CStr(varTmp00) & " Accounts Are Unassigned'"
320           End If
330         End If

340       Else
350         Cancel = -1
360       End If
370     End With

380     If Cancel = -1 Then
390       Beep
400       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
410     End If

EXITP:
420     Exit Sub

ERRH:
430     Forms(strCallingForm).Visible = True
440     DoCmd.Restore
450     Select Case ERR.Number
        Case Else
460       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
470     End Select
480     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

510     With Me
520       If .Page >= 2 Then
530         .Header_lbl2.Visible = True
540         .AsOf_lbl2.Visible = True
550         .Header_hline05.Visible = True
560       Else
570         .Header_lbl2.Visible = False
580         .AsOf_lbl2.Visible = False
590         .Header_hline05.Visible = False
600       End If
610     End With

EXITP:
620     Exit Sub

ERRH:
630     Forms(strCallingForm).Visible = True
640     DoCmd.Restore
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub Report_Close()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

710     DoCmd.Restore

720     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

730     If Forms(strCallingForm).Visible = False Then
740       Forms(strCallingForm).Visible = True
750       DoCmd.SelectObject acForm, strCallingForm, False
760       DoCmd.Restore
770     End If

EXITP:
780     Exit Sub

ERRH:
790     Forms(strCallingForm).Visible = True
800     DoCmd.Restore
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub
