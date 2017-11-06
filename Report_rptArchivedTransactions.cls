VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptArchivedTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptArchivedTransactions"

'VGC 04/25/2015: CHANGES!

Private Const strCallingForm As String = "frmRpt_ArchivedTransactions"

Private lngPageHeaderSection_Height As Long, lngAccountNo_Top As Long, lngOneLine_Top As Long, lngTwoLine_Top As Long, lngUnderline_Top As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim strAccountNo As String, strSort As String

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

220         Select Case Forms(strCallingForm).opgAccountNumber
            Case Forms(strCallingForm).opgAccountNumber_optSpecified.OptionValue
230           If .RecordSource <> "qryRpt_ArchivedTransactions_02" Then
240             .RecordSource = "qryRpt_ArchivedTransactions_02"
250           End If
260           strAccountNo = Forms(strCallingForm).cmbAccounts
270           .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
280         Case Forms(strCallingForm).opgAccountNumber_optAll.OptionValue
290           .RecordSource = "qryRpt_ArchivedTransactions_03"
300           strAccountNo = "All"
310           .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
320         End Select

330         lngPageHeaderSection_Height = .Section(acPageHeader).Height
340         lngOneLine_Top = .transdate_lbl.Top
350         lngTwoLine_Top = .icash_lbl.Top
360         lngUnderline_Top = .transdate_lbl_line.Top
370         lngAccountNo_Top = .accountno.Top

380         .AsOf_lbl.Caption = "From " & Format(Forms(strCallingForm).DateStart, "mm/dd/yyyy") & " To " & _
              Format(Forms(strCallingForm).DateEnd, "mm/dd/yyyy")
390         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
400         .AsOf_lbl2.Caption = "From " & Format(Forms(strCallingForm).DateStart, "mm/dd/yyyy") & " To " & _
              Format(Forms(strCallingForm).DateEnd, "mm/dd/yyyy")

410         .Caption = "Archived_Transactions_" & strAccountNo & "_" & Format(Forms(strCallingForm).DateStart, "yymmdd") & "_To_" & _
              Format(Forms(strCallingForm).DateEnd, "yymmdd")

420         With Forms(strCallingForm)
430           Select Case .opgOrderBy
              Case .opgOrderBy_optDate.OptionValue
                ' ** Order by date.
440             strSort = "[journalno]"
450           Case .opgOrderBy_optType.OptionValue
                ' ** Order by journal type.
460             strSort = "[journaltype]"
470           End Select
480         End With
490         .OrderBy = strSort
500         .OrderByOn = True

510         Select Case Forms(strCallingForm).chkCombineCash
            Case True
520           .icash_lbl.Visible = False
530           .icash_lbl_line.Visible = False
540           .PreviousIcash.Visible = False
550           .ICash.Visible = False
560           .SumIcash.Visible = False
570           .SumIcash_line.Visible = False
580           .pcash_lbl.Visible = False
590           .pcash_lbl_line.Visible = False
600           .PreviousPcash.Visible = False
610           .PCash.Visible = False
620           .SumPcash.Visible = False
630           .SumPcash_line.Visible = False
640           .cash_lbl.Visible = True
650           .cash_lbl_line.Visible = True
660           .PreviousCash.Visible = True
670           .cash.Visible = True
680           .SumCash.Visible = True
690           .SumCash_line.Visible = True
700         Case False
              ' ** Leave as is.
710         End Select

720       Else
730         Cancel = -1
740       End If
750     End With

760     If Cancel = -1 Then
770       Beep
780       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
790     End If

EXITP:
800     Exit Sub

ERRH:
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim lngTmp01 As Long

910     With Me
920       lngTmp01 = .AsOf_lbl2.Height
930       If .Page >= 2 Then
940         .Section(acPageHeader).Height = (lngPageHeaderSection_Height + lngTmp01)
950         .transdate_lbl.Top = (lngOneLine_Top + lngTmp01)
960         .transdate_lbl_line.Top = (lngUnderline_Top + lngTmp01)
970         .totdesc_lbl.Top = (lngOneLine_Top + lngTmp01)
980         .totdesc_lbl_line.Top = (lngUnderline_Top + lngTmp01)
990         .cash_lbl.Top = (lngOneLine_Top + lngTmp01)
1000        .cash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1010        .icash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1020        .icash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1030        .pcash_lbl.Top = (lngTwoLine_Top + lngTmp01)
1040        .pcash_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1050        .cost_lbl.Top = (lngOneLine_Top + lngTmp01)
1060        .cost_lbl_line.Top = (lngUnderline_Top + lngTmp01)
1070        .accountno.Top = (lngAccountNo_Top + lngTmp01)
1080        .shortname.Top = (lngAccountNo_Top + lngTmp01)
1090        .Header_lbl2.Visible = True
1100        .AsOf_lbl2.Visible = True
1110        .Header_hline05.Visible = True
1120      Else
1130        .Header_lbl2.Visible = False
1140        .AsOf_lbl2.Visible = False
1150        .Header_hline05.Visible = False
1160        .accountno.Top = lngAccountNo_Top
1170        .shortname.Top = lngAccountNo_Top
1180        .transdate_lbl.Top = lngOneLine_Top
1190        .transdate_lbl_line.Top = lngUnderline_Top
1200        .totdesc_lbl.Top = lngOneLine_Top
1210        .totdesc_lbl_line.Top = lngUnderline_Top
1220        .cash_lbl.Top = lngOneLine_Top
1230        .cash_lbl_line.Top = lngUnderline_Top
1240        .icash_lbl.Top = lngTwoLine_Top
1250        .icash_lbl_line.Top = lngUnderline_Top
1260        .pcash_lbl.Top = lngTwoLine_Top
1270        .pcash_lbl_line.Top = lngUnderline_Top
1280        .cost_lbl.Top = lngOneLine_Top
1290        .cost_lbl_line.Top = lngUnderline_Top
1300        .Section(acPageHeader).Height = lngPageHeaderSection_Height
1310      End If
1320    End With

EXITP:
1330    Exit Sub

ERRH:
1340    Select Case ERR.Number
        Case Else
1350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1360    End Select
1370    Resume EXITP

End Sub

Private Sub Report_Close()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1410    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1420    Exit Sub

ERRH:
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub
