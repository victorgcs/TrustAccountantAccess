VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptListOfVoidedChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptListOfVoidedChecks"

'VGC 07/13/2017: CHANGES!

Private Const strCallingForm As String = "frmRpt_Checks_Void_Mgmt"
Private Const strCallersCallingForm As String = "frmRpt_Checks"

Private lngTpp As Long
Private lngPageHeaderSection_Height As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

150         lngPageHeaderSection_Height = .Section(acPageHeader).Height
160         lngOneLine_Top = .chkvoid_payee_lbl.Top
170         lngTwoLine_Top = .chkvoid_chknum_lbl.Top
180         lngUnderline_Top = .chkvoid_chknum_lbl_line.Top

            ' ** Default to PageNumOf if calling form has no check box.
190         .PageNumOf.Visible = True
200         .PageNum.Visible = False

210         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
220         .AsOf2.Caption = "As of " & Format(Now(), "mm/dd/yyyy")
230         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
240         .Caption = "List_of_Voided_Checks_" & Format(Date, "yymmdd")

250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
320     End If

EXITP:
330     Exit Sub

ERRH:
340     DoCmd.Restore
350     Forms(strCallersCallingForm).Visible = True
360     Forms(strCallingForm).Visible = True
370     Select Case ERR.Number
        Case Else
380       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
390     End Select
400     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

510     With Me
520       lngTmp01 = (.AsOf2.Height + (8& * lngTpp))
530       If .Page >= 2 Then
540         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
550         .chkvoid_chknum_lbl.Top = (lngTwoLine_Top + lngTmp01)
560         .chkvoid_chknum_lbl2.Top = (lngOneLine_Top + lngTmp01)
570         .chkvoid_chknum_lbl_line.Top = (lngUnderline_Top + lngTmp01)
580         .chkvoid_date_lbl.Top = (lngTwoLine_Top + lngTmp01)
590         .chkvoid_date_lbl2.Top = (lngOneLine_Top + lngTmp01)
600         .chkvoid_date_lbl_line.Top = (lngUnderline_Top + lngTmp01)
610         .accountno_lbl.Top = (lngTwoLine_Top + lngTmp01)
620         .accountno_lbl2.Top = (lngOneLine_Top + lngTmp01)
630         .accountno_lbl_line.Top = (lngUnderline_Top + lngTmp01)
640         .shortname_lbl.Top = (lngOneLine_Top + lngTmp01)
650         .shortname_lbl_line.Top = (lngUnderline_Top + lngTmp01)
660         .transdate_lbl.Top = (lngTwoLine_Top + lngTmp01)
670         .transdate_lbl2.Top = (lngOneLine_Top + lngTmp01)
680         .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
690         .chkvoid_payee_lbl.Top = (lngOneLine_Top + lngTmp01)
700         .chkvoid_payee_lbl_line.Top = (lngUnderline_Top + lngTmp01)
710         .chkvoid_amount_lbl.Top = (lngOneLine_Top + lngTmp01)
720         .chkvoid_amount_lbl_line.Top = (lngUnderline_Top + lngTmp01)
730         .Header_lbl2.Visible = True
740         .AsOf2.Visible = True
750         .Header_hline05.Visible = True
760       Else
770         .Header_lbl2.Visible = False
780         .AsOf2.Visible = False
790         .Header_hline05.Visible = False
800         .chkvoid_chknum_lbl.Top = lngTwoLine_Top
810         .chkvoid_chknum_lbl2.Top = lngOneLine_Top
820         .chkvoid_chknum_lbl_line.Top = lngUnderline_Top
830         .chkvoid_date_lbl.Top = lngTwoLine_Top
840         .chkvoid_date_lbl2.Top = lngOneLine_Top
850         .chkvoid_date_lbl_line.Top = lngUnderline_Top
860         .accountno_lbl.Top = lngTwoLine_Top
870         .accountno_lbl2.Top = lngOneLine_Top
880         .accountno_lbl_line.Top = lngUnderline_Top
890         .shortname_lbl.Top = lngOneLine_Top
900         .shortname_lbl_line.Top = lngUnderline_Top
910         .transdate_lbl.Top = lngTwoLine_Top
920         .transdate_lbl2.Top = lngOneLine_Top
930         .transdate_lbl_line.Top = lngUnderline_Top
940         .chkvoid_payee_lbl.Top = lngOneLine_Top
950         .chkvoid_payee_lbl_line.Top = lngUnderline_Top
960         .chkvoid_amount_lbl.Top = lngOneLine_Top
970         .chkvoid_amount_lbl_line.Top = lngUnderline_Top
980         .Section(acPageHeader).Height = lngPageHeaderSection_Height
990       End If
1000    End With

EXITP:
1010    Exit Sub

ERRH:
1020    DoCmd.Restore
1030    Forms(strCallersCallingForm).Visible = True
1040    Forms(strCallingForm).Visible = True
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1110    MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
1120    Cancel = -1

EXITP:
1130    Exit Sub

ERRH:
1140    DoCmd.Restore
1150    Forms(strCallersCallingForm).Visible = True
1160    Forms(strCallingForm).Visible = True
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
1200    Resume EXITP

End Sub

Private Sub Report_Close()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1310    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1320    DoCmd.Restore
1330    Forms(strCallersCallingForm).Visible = True
1340    Forms(strCallingForm).Visible = True

EXITP:
1350    Exit Sub

ERRH:
1360    DoCmd.Restore
1370    Forms(strCallersCallingForm).Visible = True
1380    Forms(strCallingForm).Visible = True
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub
