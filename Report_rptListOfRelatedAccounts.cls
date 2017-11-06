VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfRelatedAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfRelatedAccounts"

'VGC 04/04/2016: CHANGES!

Private Const strCallingForm As String = "frmAccountProfile_RelAccts"
Private Const strCallersCallingForm As String = "frmAccountProfile"

Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngUnderline_Top As Long
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

160         lngPageHeaderSection_Height = .Section(acPageHeader).Height
170         lngOneLine_Top = .accountno_lbl.Top
180         lngUnderline_Top = .accountno_lbl_line.Top

190         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
200         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
210         .AsOf_lbl2.Caption = "As of " & Format(Date, "mmmm dd, yyyy")

220       Else
230         Cancel = -1
240       End If
250     End With

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     DoCmd.Restore
320     Forms(strCallersCallingForm).Visible = True
330     Forms(strCallingForm).Visible = True
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

410     With Me
420       lngTmp01 = .AsOf_lbl2.Height
430       If .Page >= 2 Then
440         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
450         .accountno_lbl.Top = (lngOneLine_Top + lngTmp01)
460         .accountno_lbl_line.Top = (lngUnderline_Top + lngTmp01)
470         .shortname_lbl.Top = (lngOneLine_Top + lngTmp01)
480         .shortname_lbl_line.Top = (lngUnderline_Top + lngTmp01)
490         .AccountType_Description_lbl.Top = (lngOneLine_Top + lngTmp01)
500         .AccountType_Description_lbl_line.Top = (lngUnderline_Top + lngTmp01)
510         .tin_lbl.Top = (lngOneLine_Top + lngTmp01)
520         .tin_lbl_line.Top = (lngUnderline_Top + lngTmp01)
530         .CaseNum_lbl.Top = (lngOneLine_Top + lngTmp01)
540         .CaseNum_lbl_line.Top = (lngUnderline_Top + lngTmp01)
550         .admin_officer_lbl.Top = (lngOneLine_Top + lngTmp01)
560         .admin_officer_lbl_line.Top = (lngUnderline_Top + lngTmp01)
570         .Header_lbl2.Visible = True
580         .AsOf_lbl2.Visible = True
590         .Header_hline05.Visible = True
600       Else
610         .Header_lbl2.Visible = False
620         .AsOf_lbl2.Visible = False
630         .Header_hline05.Visible = False
640         .accountno_lbl.Top = lngOneLine_Top
650         .accountno_lbl_line.Top = lngUnderline_Top
660         .shortname_lbl.Top = lngOneLine_Top
670         .shortname_lbl_line.Top = lngUnderline_Top
680         .AccountType_Description_lbl.Top = lngOneLine_Top
690         .AccountType_Description_lbl_line.Top = lngUnderline_Top
700         .tin_lbl.Top = lngOneLine_Top
710         .tin_lbl_line.Top = lngUnderline_Top
720         .CaseNum_lbl.Top = lngOneLine_Top
730         .CaseNum_lbl_line.Top = lngUnderline_Top
740         .admin_officer_lbl.Top = lngOneLine_Top
750         .admin_officer_lbl_line.Top = lngUnderline_Top
760         .Section(acPageHeader).Height = lngPageHeaderSection_Height
770       End If
780     End With

EXITP:
790     Exit Sub

ERRH:
800     DoCmd.Restore
810     Forms(strCallersCallingForm).Visible = True
820     Forms(strCallingForm).Visible = True
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Sub

Private Sub Report_Close()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

910     DoCmd.Restore

920     TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

930     If Forms(strCallingForm).Visible = False Then
940       DoCmd.Restore
950       Forms(strCallersCallingForm).Visible = True
960       Forms(strCallingForm).Visible = True
          'DoCmd.SelectObject acForm, strCallingForm, False
970     End If

EXITP:
980     Exit Sub

ERRH:
990     DoCmd.Restore
1000    Forms(strCallersCallingForm).Visible = True
1010    Forms(strCallingForm).Visible = True
1020    Select Case ERR.Number
        Case Else
1030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1040    End Select
1050    Resume EXITP

End Sub
