VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptErrorLog"

'VGC 03/21/2016: CHANGES!

Private strVersion As String
Private Const strCallingForm As String = "frmErrorLog"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
          'If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

120       TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

130       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
140         Select Case Forms(strCallingForm).chkPageOf
            Case True
150           .PageNumOf.Visible = True
160           .PageNum.Visible = False
170         Case False
180           .PageNum.Visible = True
190           .PageNumOf.Visible = False
200           .PageNumOf.ControlSource = vbNullString
210         End Select
220       Else
230         .PageNumOf.Visible = True
240         .PageNum.Visible = False
250       End If

260       strVersion = AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
270       strVersion = "Ver: " & strVersion & ", Rel: "
280       strVersion = strVersion & Format(AppDate_Get(True), "mm/dd/yyyy hh:nn:ss")  ' ** Module Function: modAppVersionFuncs.
290       .Ver_lbl.Caption = strVersion

300       .AsOf_lbl.Caption = "As of " & Format(Date, "mm/dd/yyyy")
310       .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

320         strOrderBy = Forms(strCallingForm).frmErrorLog_Sub.Form.SortNow_Get  ' ** Form Function: frmErrorLog_Sub.
            ' ** Potential Sorts:
            ' **   [ErrLog_Date] DESC
            ' **   [ErrLog_ErrNum], [ErrLog_Date] DESC
            ' **   [ErrLog_Message_Disp], [ErrLog_Date] DESC
            ' **   [ErrLog_LineNum], [ErrLog_Date] DESC
            ' **   [ErrLog_Form], [ErrLog_FuncSub], [ErrLog_ErrNum], [ErrLog_Date] DESC
            ' **   [ErrLog_FuncSub], [ErrLog_Form], [ErrLog_ErrNum], [ErrLog_Date] DESC
330         .OrderBy = strOrderBy
340         .OrderByOn = True

350       .Caption = "Error_Log_" & Format(Date, "yymmdd")

360       If gstrReportCallingForm <> vbNullString Then
370         Forms(gstrReportCallingForm).Visible = False
380       End If

          'Else
          '  Cancel = -1
          'End If
390     End With

400     If Cancel = -1 Then
          ' ** Just for the record.
410       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
420     End If

EXITP:
430     Exit Sub

ERRH:
440     If gstrReportCallingForm <> vbNullString Then
450       DoCmd.Restore
460       Forms(gstrReportCallingForm).Visible = True
470       DoCmd.SelectObject acForm, gstrReportCallingForm, False
480     End If
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Sub

Private Sub Report_Close()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

610     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

620     DoCmd.Restore

630     If gstrReportCallingForm <> vbNullString Then
640       Forms(gstrReportCallingForm).Visible = True
650     End If

EXITP:
660     Exit Sub

ERRH:
670     If gstrReportCallingForm <> vbNullString Then
680       DoCmd.Restore
690       Forms(gstrReportCallingForm).Visible = True
700       DoCmd.SelectObject acForm, gstrReportCallingForm, False
710     End If
720     Select Case ERR.Number
        Case Else
730       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
740     End Select
750     Resume EXITP

End Sub
