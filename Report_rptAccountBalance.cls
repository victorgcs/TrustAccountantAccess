VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAccountBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAccountBalance"

'VGC 05/03/2015: CHANGES!

Private Const strCallingForm As String = "frmRpt_AccountBalance"

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
230         lngOneLine_Top = .accountno_lbl.Top
240         lngTwoLine_Top = .Assets_lbl.Top
250         lngUnderline_Top = .accountno_lbl_line.Top

260         .AsOf_lbl.Caption = "as of " & Format(Forms(strCallingForm).DateAsOf, "mmmm dd, yyyy")
270         .AsOf_lbl2.Caption = "as of " & Format(Date, "mm/dd/yyyy")
280         .AsOf.Caption = "Printed " & Format(Date, "mmmm dd, yyyy")
290         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
300         .Caption = "Account_Balance_" & Format(Forms(strCallingForm).DateAsOf, "yymmdd")

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
400     Select Case ERR.Number
        Case Else
410       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
420     End Select
430     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

510     With Me
520       lngTmp01 = .AsOf_lbl2.Height
530       If .Page >= 2 Then
540         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
550         .accountno_lbl.Top = (lngOneLine_Top + lngTmp01)
560         .accountno_lbl_line.Top = (lngUnderline_Top + lngTmp01)
570         .shortname_lbl.Top = (lngOneLine_Top + lngTmp01)
580         .shortname_lbl_line.Top = (lngUnderline_Top + lngTmp01)
590         .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
600         .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
610         .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
620         .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
630         .Assets_lbl.Top = (lngTwoLine_Top + lngTmp01)
640         .Assets_lbl_line.Top = (lngUnderline_Top + lngTmp01)
650         .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
660         .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
670         .MV_lbl.Top = (lngOneLine_Top + lngTmp01)
680         .MV_lbl_line.Top = (lngUnderline_Top + lngTmp01)
690         .Header_lbl2.Visible = True
700         .AsOf_lbl2.Visible = True
710         .Header_hline05.Visible = True
720       Else
730         .Header_lbl2.Visible = False
740         .AsOf_lbl2.Visible = False
750         .Header_hline05.Visible = False
760         .accountno_lbl.Top = lngOneLine_Top
770         .accountno_lbl_line.Top = lngUnderline_Top
780         .shortname_lbl.Top = lngOneLine_Top
790         .shortname_lbl_line.Top = lngUnderline_Top
800         .icash_lbl.Top = lngTwoLine_Top
810         .icash_lbl_line.Top = lngUnderline_Top
820         .pcash_lbl.Top = lngTwoLine_Top
830         .pcash_lbl_line.Top = lngUnderline_Top
840         .Assets_lbl.Top = lngTwoLine_Top
850         .Assets_lbl_line.Top = lngUnderline_Top
860         .cost_lbl.Top = lngOneLine_Top
870         .cost_lbl_line.Top = lngUnderline_Top
880         .MV_lbl.Top = lngOneLine_Top
890         .MV_lbl_line.Top = lngUnderline_Top
900         .Section(acPageHeader).Height = lngPageHeaderSection_Height
910       End If
920     End With

EXITP:
930     Exit Sub

ERRH:
940     Select Case ERR.Number
        Case Else
950       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
960     End Select
970     Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1010    MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
1020    Cancel = -1

EXITP:
1030    Exit Sub

ERRH:
1040    Select Case ERR.Number
        Case Else
1050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1060    End Select
1070    Resume EXITP

End Sub

Private Sub Report_Close()

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1110    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1120    Exit Sub

ERRH:
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub
