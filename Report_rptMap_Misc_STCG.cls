VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptMap_Misc_STCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptMap_Misc_STCG"

'VGC 07/24/2017: CHANGES!

Private Const strCallingForm As String = "frmMap_Misc_STCGL_Detail"

Private intMode As Integer
Private strCallersCallingForm2 As String
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strOrderBy As String
        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         strTmp01 = Trim(.OpenArgs)
140         intPos01 = InStr(strTmp01, "~")
150         strCallersCallingForm2 = Left(strTmp01, (intPos01 - 1))
160         intMode = CInt(Mid(strTmp01, (intPos01 + 1)))

170         If intMode = acViewPreview Then
180           Forms(strCallersCallingForm2).Visible = False
190         End If

200         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

            ' ** Default to PageNumOf if calling form has no check box.
210         .PageNumOf.Visible = True
220         .PageNum.Visible = False

230         .AsOf.Caption = Format(Now(), "mmmm dd, yyyy  hh:nn:ss AM/PM")
240         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

250         strOrderBy = Forms(strCallingForm).frmMap_Misc_STCGL_Detail_Sub.Form.SortNow_Get  ' ** Form Function: frmMap_Misc_STCGL_Detail_Sub.
            ' ** Potential Sorts:
            ' **   [alphasort]
            ' **   [shortname]
            ' **   [total_shareface], [alphasort]
            ' **   [total_icash], [alphasort]
            ' **   [total_icash_usd], [alphasort]
260         .OrderBy = strOrderBy
270         .OrderByOn = True

280         .transdate.ControlSource = "=#" & Format(Forms(strCallingForm).transdate, "mm/dd/yyyy") & "#"
290         .pershare.ControlSource = "=" & Chr(34) & CStr(Forms(strCallingForm).pershare) & Chr(34)
300         .amount.ControlSource = "=" & Chr(34) & Format(Forms(strCallingForm).amount, "Currency") & Chr(34)

310       Else
320         Cancel = -1
330       End If
340     End With

350     If Cancel = -1 Then
360       Beep
370       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
380     End If

EXITP:
390     Exit Sub

ERRH:
400     Forms(strCallersCallingForm2).Visible = True
410     Forms(strCallingForm).Visible = True
420     DoCmd.Restore
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim strCusip As String

510     With Me
520       strCusip = .cusip
530       .Caption = "STCG_Misc_Map_Detail_" & strCusip & "_" & Format(Date, "yymmdd")
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
880     DoCmd.Restore
890     Select Case ERR.Number
        Case Else
900       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
910     End Select
920     Resume EXITP

End Sub
