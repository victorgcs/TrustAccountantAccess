VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfJournalTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfJournalTypes"

'VGC 03/19/2016: CHANGES!

Private Const strCallingForm As String = "frmJournalType"
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

160         .detail_height_line.Top = 0&  ' ** Get it out of the way so the section can shrink.

170         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
180         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

190         strOrderBy = Forms(strCallingForm).frmJournalType_Sub.Form.SortNow_Get  ' ** Form Function: frmJournalType_Sub.
            ' ** Potential Sorts:
            ' **   [JournalType_Order]
            ' **   [JournalType]
            ' **   [JournalType_Description]
200         .OrderBy = strOrderBy
210         .OrderByOn = True

220         .Caption = "List_of_Journal_Types_" & Format(Date, "yymmdd")

230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Forms(strCallingForm).Visible = True
330     DoCmd.Restore
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

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
