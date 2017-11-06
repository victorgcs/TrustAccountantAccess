VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptMap_Dividend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptMap_Dividend"

'VGC 07/24/2017: CHANGES!

Private Const strCallingForm As String = "frmMap_Div_Detail"

Private intMode As Integer
Private strCallersCallingForm2 As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         If IsLoaded("frmJournal", acForm) = True Then
140           strCallersCallingForm2 = "frmJournal"
150         ElseIf IsLoaded("frmJournal_Columns", acForm) = True Then
160           strCallersCallingForm2 = "frmJournal_Columns"
170         End If

180         intMode = CInt(.OpenArgs)
190         If intMode = acViewPreview Then
200           Forms(strCallersCallingForm2).Visible = False
210         End If

220         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
230         .PageNumOf.Visible = True
240         .PageNum.Visible = False

250         .AsOf.Caption = Format(Now(), "mmmm dd, yyyy  hh:nn:ss AM/PM")
260         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

270         strOrderBy = Forms(strCallingForm).frmMap_Div_Detail_Sub.Form.SortNow_Get  ' ** Form Function: frmMap_Div_Detail_Sub.
            ' ** Potential Sorts:
            ' **   [alphasort]
            ' **   [shortname]
            ' **   [total_shareface], [alphasort]
            ' **   [total_icash], [alphasort]
            ' **   [total_icash_usd], [alphasort]
280         .OrderBy = strOrderBy
290         .OrderByOn = True

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
390     Forms(strCallersCallingForm2).Visible = True
400     Forms(strCallingForm).Visible = True
410     DoCmd.Restore
420     Select Case ERR.Number
        Case Else
430       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
440     End Select
450     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim strCusip As String

510     With Me
520       strCusip = .cusip
530       .Caption = "Dividend_Map_Detail_" & strCusip & "_" & Format(Date, "yymmdd")
540     End With

EXITP:
550     Exit Sub

ERRH:
560     Forms(strCallersCallingForm2).Visible = True
570     Forms(strCallingForm).Visible = True
580     DoCmd.Restore
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub Report_Close()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

710     If Forms(strCallingForm).Visible = False Then
720       DoCmd.Restore
730       Forms(strCallersCallingForm2).Visible = True
740       Forms(strCallingForm).Visible = True
750       DoCmd.SelectObject acForm, strCallingForm, False
760       DoCmd.Restore
770     End If

780     If gstrReportCallingForm <> vbNullString Then
790       If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
800         Forms(gstrReportCallingForm).Modal = True
810       End If
820       DoCmd.SelectObject acForm, gstrReportCallingForm, False
830     End If

840     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
850     Exit Sub

ERRH:
860     Forms(strCallersCallingForm2).Visible = True
870     Forms(strCallingForm).Visible = True
880     DoCmd.SelectObject acForm, strCallingForm, False
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub
