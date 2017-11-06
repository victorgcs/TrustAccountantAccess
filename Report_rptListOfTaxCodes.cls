VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfTaxCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfTaxCodes"

'VGC 11/27/2016: CHANGES!

Private Const strCallingForm As String = "frmTaxCode"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
140         .PageNumOf.Visible = True
150         .PageNum.Visible = False

160         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
170         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

180         .Caption = "List_of_Tax_Codes_" & Format(Date, "yymmdd")

190         Select Case gblnLinkRevTaxCodes
            Case True
200           .revcode_TYPE_Description.Visible = True
210         Case False
220           .revcode_TYPE_Description.Visible = False
230         End Select

240       Else
250         Cancel = -1
260       End If
270     End With

280     If Cancel = -1 Then
290       Beep
300       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
310     End If

EXITP:
320     Exit Sub

ERRH:
330     Forms(strCallingForm).Visible = True
340     DoCmd.Restore
350     Select Case ERR.Number
        Case Else
360       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
370     End Select
380     Resume EXITP

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
