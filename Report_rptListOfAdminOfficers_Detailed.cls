VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfAdminOfficers_Detailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfAdminOfficers_Detailed"

'VGC 05/29/2016: CHANGES!

Private Const strCallingForm As String = "frmAdminOfficer"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

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

190         .Caption = "List_of_Administrative_Officers_Detailed_" & Format(Date, "yymmdd")

200         If Forms(strCallingForm).Unassigned_lbl.Visible = True Then
210           varTmp00 = DLookup("[cnt]", "qryAdminOfficer_05", "[adminno] = 1")
220           If varTmp00 = 1 Then
230             .Unassigned.ControlSource = "='1 Account Is Unassigned'"
240           Else
250             .Unassigned.ControlSource = "='" & CStr(varTmp00) & " Accounts Are Unassigned'"
260           End If
270         End If

280       Else
290         Cancel = -1
300       End If
310     End With

320     If Cancel = -1 Then
330       Beep
340       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
350     End If

EXITP:
360     Exit Sub

ERRH:
370     Forms(strCallingForm).Visible = True
380     DoCmd.Restore
390     Select Case ERR.Number
        Case Else
400       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
410     End Select
420     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

510     With Me
520       If .Page >= 2 Then
530         .Section(acPageHeader).Visible = True
540       Else
550         .Section(acPageHeader).Visible = False
560       End If
570     End With

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
