VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAccountReviews_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAccountReviews_Admin"

'VGC 03/22/2017: CHANGES!

Private Const strCallingForm As String = "frmRpt_AccountReviews"

' ** cmbMonth combo box constants:
'Private Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
Private Const CBX_MON_NAME  As Integer = 1  ' ** month_name
Private Const CBX_MON_SHORT As Integer = 2  ' ** month_short
'Private Const CBX_MON_CNT   As Integer = 3  ' ** cnt

Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngUnderline_Top As Long
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
240         lngUnderline_Top = .accountno_lbl_line.Top

250         .AsOf_lbl.Caption = "For Month of " & Forms(strCallingForm).cmbMonth.Column(CBX_MON_NAME)
260         .AsOf.Caption = "As of " & Format(Date, "mmmm dd, yyyy")
270         .AsOf2.Caption = "As of " & Format(Date, "mm/dd/yyyy")
280         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

290         .Caption = "Account_Reviews_By_Administrative_Officer_" & FormatProperCase(Forms(strCallingForm).cmbMonth.Column(CBX_MON_SHORT))  ' ** Module Function: modStringFuncs.

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
520       lngTmp01 = .AsOf2.Height
530       If .Page >= 2 Then
540         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
550         .accountno_lbl.Top = (lngOneLine_Top + lngTmp01)
560         .accountno_lbl_line.Top = (lngUnderline_Top + lngTmp01)
570         .shortname_lbl.Top = (lngOneLine_Top + lngTmp01)
580         .shortname_lbl_line.Top = (lngUnderline_Top + lngTmp01)
590         .Header_lbl2.Visible = True
600         .AsOf2.Visible = True
610         .Header_hline05.Visible = True
620       Else
630         .Header_lbl2.Visible = False
640         .AsOf2.Visible = False
650         .Header_hline05.Visible = False
660         .accountno_lbl.Top = lngOneLine_Top
670         .accountno_lbl_line.Top = lngUnderline_Top
680         .shortname_lbl.Top = lngOneLine_Top
690         .shortname_lbl_line.Top = lngUnderline_Top
700         .Section(acPageHeader).Height = lngPageHeaderSection_Height
710       End If
720     End With

EXITP:
730     Exit Sub

ERRH:
740     Select Case ERR.Number
        Case Else
750       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
760     End Select
770     Resume EXITP

End Sub

Private Sub Report_Close()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

810     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
820     Exit Sub

ERRH:
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub
