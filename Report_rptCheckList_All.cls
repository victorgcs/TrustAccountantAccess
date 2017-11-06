VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCheckList_All"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCheckList_All"

'VGC 07/06/2017: CHANGES!

' ** Without User subdivisions; for all or single user.

' ** tblMark populated in frmMenu_Post.

Private Const strCallingForm As String = "frmMenu_Post"

Private frm As Access.Form
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

150         Set frm = Forms(strCallingForm)

160         Select Case frm.WithVoids
            Case True
              ' ** qryPrintChecks_20_03 (qryPrintChecks_20_02 (Union of qryPrintChecks_20 (qryPrintChecks_20c
              ' ** (qryPrintChecks_20a (Journal, linked to tblCheckPrint, with add'l fields, all journal_USER's),
              ' ** linked to tblJournal_Memo, including journal_USER, just 'Paid', for PrintCheck = True), linked
              ' ** to tblMark, with Memo_Print, Rpt_Ord), qryPrintChecks_20_01 (tblCheckVoid, linked to Journal)),
              ' ** sorted by accountno, transdate, sortOrder, ID), linked to tblCheckPrint_Report, and qryPrintChecks_20_10
              ' ** (tblCheckPrint_Report, grouped, with Min(chkprintrpt_id)), with Rpt_Ord; Cartesian.
170           .RecordSource = "qryPrintChecks_20_11"
180           .chkvoid_void_lbl.Visible = True
190           .chkvoid_void_lbl_line.Visible = True
200           .chkvoid_void.Visible = True
210           .chkvoid_voidtot.Visible = True
220           .chkvoid_voidtot_line.Visible = True
230         Case False
              ' ** qryPrintChecks_20c (qryPrintChecks_20a (Journal, linked to tblCheckPrint, with add'l fields,
              ' ** all journal_USER's), linked to tblJournal_Memo, including journal_USER, just 'Paid', for
              ' ** PrintCheck = True), linked to tblMark, with Memo_Print, Rpt_Ord.
240           If .RecordSource <> "qryPrintChecks_20" Then
250             .RecordSource = "qryPrintChecks_20"
260           End If
270           .chkvoid_void_lbl.Visible = False
280           .chkvoid_void_lbl_line.Visible = False
290           .chkvoid_void.Visible = False
300           .chkvoid_voidtot.Visible = False
310           .chkvoid_voidtot_line.Visible = False
320         End Select

330         lngPageHeaderSection_Height = .Section(acPageHeader).Height
340         lngOneLine_Top = .CheckNum_lbl.Top
350         lngTwoLine_Top = .pcash_lbl.Top
360         lngUnderline_Top = .CheckNum_lbl_line.Top

            ' ** Default to PageNumOf if calling form has no check box.
370         .PageNumOf.Visible = True
380         .PageNum.Visible = False

390         If gstrCo_Name = vbNullString Then
400           CoOptions_Read  ' ** Module Function: modStartupFuncs.
410         End If

420         .AsOf.Caption = Format(Date, "mmmm dd, yyyy")
430         .AsOf_lbl2.Caption = Format$(Date, "mmmm dd, yyyy")
440         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

450         .Caption = "Check_List_All_" & Format(Date, "yymmdd")

460         If gstrReportCallingForm <> vbNullString Then
470           Forms(gstrReportCallingForm).Visible = False
480         End If

490       Else
500         Cancel = -1
510       End If
520     End With

530     If Cancel = -1 Then
540       Beep
550       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
560     End If

EXITP:
570     Exit Sub

ERRH:
580     If gstrReportCallingForm <> vbNullString Then
590       DoCmd.Restore
600       Forms(gstrReportCallingForm).Visible = True
610       DoCmd.SelectObject acForm, gstrReportCallingForm, False
620     End If
630     Select Case ERR.Number
        Case Else
640       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
650     End Select
660     Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

710     With Me
720       If .Filter = vbNullString Then
730         .AsOf_lbl.Caption = "All Journal Entries Marked To Receive A Check"
740       Else
750         .AsOf_lbl.Caption = "Journal Entries Marked To Receive A Check For User: " & gstrJournalUser
760       End If
770     End With

EXITP:
780     Exit Sub

ERRH:
790     If gstrReportCallingForm <> vbNullString Then
800       DoCmd.Restore
810       Forms(gstrReportCallingForm).Visible = True
820       DoCmd.SelectObject acForm, gstrReportCallingForm, False
830     End If
840     Select Case ERR.Number
        Case Else
850       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
860     End Select
870     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

910     With Me
920       lngTmp01 = (.AsOf_lbl2.Height + (8& * lngTpp))
930       If .Page >= 2 Then
940         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
950         .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
960         .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
970         .totdesc_lbl.Top = (lngOneLine_Top + lngTmp01)
980         .totdesc_lbl_line.Top = (lngUnderline_Top + lngTmp01)
990         .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1000        .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1010        .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1020        .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1030        .CheckNum_lbl.Top = (lngOneLine_Top + lngTmp01)
1040        .CheckNum_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1050        .chkvoid_void_lbl.Top = (lngOneLine_Top + lngTmp01)
1060        .chkvoid_void_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1070        .Header_lbl2.Visible = True
1080        .AsOf_lbl2.Visible = True
1090        .Header_hline05.Visible = True
1100      Else
1110        .Header_lbl2.Visible = False
1120        .AsOf_lbl2.Visible = False
1130        .Header_hline05.Visible = False
1140        .transdate_lbl.Top = lngOneLine_Top
1150        .transdate_lbl_line.Top = lngUnderline_Top
1160        .totdesc_lbl.Top = lngOneLine_Top
1170        .totdesc_lbl_line.Top = lngUnderline_Top
1180        .icash_lbl.Top = lngTwoLine_Top
1190        .icash_lbl_line.Top = lngUnderline_Top
1200        .pcash_lbl.Top = lngTwoLine_Top
1210        .pcash_lbl_line.Top = lngUnderline_Top
1220        .CheckNum_lbl.Top = lngOneLine_Top
1230        .CheckNum_lbl_line.Top = lngUnderline_Top
1240        .chkvoid_void_lbl.Top = lngOneLine_Top
1250        .chkvoid_void_lbl_line.Top = lngUnderline_Top
1260        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1270      End If
1280    End With

EXITP:
1290    Exit Sub

ERRH:
1300    Select Case ERR.Number
        Case Else
1310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1320    End Select
1330    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1410    MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
1420    Cancel = -1

EXITP:
1430    Exit Sub

ERRH:
1440    If gstrReportCallingForm <> vbNullString Then
1450      DoCmd.Restore
1460      Forms(gstrReportCallingForm).Visible = True
1470      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1480    End If
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Report_Close()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1610    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1620    DoCmd.Restore

1630    If gstrReportCallingForm <> vbNullString Then
1640      Forms(gstrReportCallingForm).Visible = True
1650    End If

EXITP:
1660    Exit Sub

ERRH:
1670    If gstrReportCallingForm <> vbNullString Then
1680      DoCmd.Restore
1690      Forms(gstrReportCallingForm).Visible = True
1700      DoCmd.SelectObject acForm, gstrReportCallingForm, False
1710    End If
1720    Select Case ERR.Number
        Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub

'totdesc:
'=IIf(
'     IsNull([RecurringItem]),
'     "",
'     IIf(
'         [journaltype]="Received",
'         [RecurringItem],
'         IIf(
'             [journaltype]="Paid",
'             [RecurringItem],
'             [RecurringItem]
'            )
'        )
'    ) &
' IIf(
'     [assetno] Is Not Null,
'     IIf(
'         [assetdate] Is Not Null,
'         Format([assetdate],"mm/dd/yyyy") & " ",
'         ""
'        ) &
'     IIf(
'         [shareface]-CLng([shareface])=0,
'         Format([shareface],"#,##0 "),
'         Format([shareface],"#,##0.0000 ")
'        ) &
'     CStr([Description]) & " " &
'     IIf(
'         [rate]>0,
'         " " & Format([rate],"#,##0.000%"),
'         ""
'        ) &
'     IIf(
'         [due] Is Not Null,
'         "  Due " & Format([due],"mm/dd/yyyy")
'        ),
'     " " & [jcomment]
'    )
