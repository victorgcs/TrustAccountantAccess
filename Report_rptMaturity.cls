VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptMaturity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptMaturity"

'VGC 08/28/2015: CHANGES!

Private Const strCallingForm As String = "frmRpt_Maturity"

Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Select Case Forms(strCallingForm).chkPageOf
            Case True
150           .PageNumOf.Visible = True
160           .PageNum.Visible = False
170         Case False
180           .PageNum.Visible = True
190           .PageNumOf.Visible = False
200           .PageNumOf.ControlSource = vbNullString
210         End Select

220         lngPageHeaderSection_Height = .Section(acPageHeader).Height
230         lngOneLine_Top = .due_lbl.Top
240         lngTwoLine_Top = .accountno_lbl.Top
250         lngUnderline_Top = .due_lbl_line.Top

260         .AsOf_lbl.Caption = "From " & Format(CDate(Forms(strCallingForm).DateStart), "mm/dd/yyyy") & " To " & _
              Format(CDate(Forms(strCallingForm).DateEnd), "mm/dd/yyyy")
270         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
280         .AsOf_lbl2.Caption = "From " & Format(CDate(Forms(strCallingForm).DateStart), "mm/dd/yyyy") & _
              " To " & Format(CDate(Forms(strCallingForm).DateEnd), "mm/dd/yyyy")  ' ** Module Function: modQueryFunctions1.

290         .Caption = "Security_Maturity_" & Format(Forms(strCallingForm).DateStart, "yymmdd") & "_To_" & _
              Format(Forms(strCallingForm).DateEnd, "yymmdd")

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
390     Select Case ERR.Number
        Case Else
400       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
410     End Select
420     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

510     With Me
520       lngTmp01 = .AsOf_lbl2.Height
530       If .Page >= 2 Then
540         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
550         .due_lbl.Top = (lngOneLine_Top + lngTmp01)
560         .due_lbl_line.Top = (lngUnderline_Top + lngTmp01)
570         .shareface_lbl.Top = (lngOneLine_Top + lngTmp01)
580         .shareface_lbl_line.Top = (lngUnderline_Top + lngTmp01)
590         .totdesc_lbl.Top = (lngOneLine_Top + lngTmp01)
600         .totdesc_lbl_line.Top = (lngUnderline_Top + lngTmp01)
610         .accountno_lbl.Top = (lngTwoLine_Top + lngTmp01)
620         .accountno_lbl_line.Top = (lngUnderline_Top + lngTmp01)
630         .shortname_lbl.Top = (lngOneLine_Top + lngTmp01)
640         .shortname_lbl_line.Top = (lngUnderline_Top + lngTmp01)
650         .Header_lbl2.Visible = True
660         .AsOf_lbl2.Visible = True
670         .Header_hline05.Visible = True
680       Else
690         .Header_lbl2.Visible = False
700         .AsOf_lbl2.Visible = False
710         .Header_hline05.Visible = False
720         .due_lbl.Top = lngOneLine_Top
730         .due_lbl_line.Top = lngUnderline_Top
740         .shareface_lbl.Top = lngOneLine_Top
750         .shareface_lbl_line.Top = lngUnderline_Top
760         .totdesc_lbl.Top = lngOneLine_Top
770         .totdesc_lbl_line.Top = lngUnderline_Top
780         .accountno_lbl.Top = lngTwoLine_Top
790         .accountno_lbl_line.Top = lngUnderline_Top
800         .shortname_lbl.Top = lngOneLine_Top
810         .shortname_lbl_line.Top = lngUnderline_Top
820         .Section(acPageHeader).Height = lngPageHeaderSection_Height
830       End If
840     End With

EXITP:
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub Report_Close()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

910     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
920     Exit Sub

ERRH:
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub
