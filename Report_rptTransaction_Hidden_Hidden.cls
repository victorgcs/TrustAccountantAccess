VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Hidden_Hidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Hidden_Hidden"

'VGC 04/04/2016: CHANGES!

' ** Original grouped sort order:
' **   [accountno], [shortname], [ledghid_grpnum], [ledghid_ord], [transdate]

Private lngICash_Left As Long, lngCurrCode_Left As Long, lngCurrCodeLbl_Width As Long, lngCurrCodeLbl_Left As Long, lngTotDesc_Width As Long

Private lngJournalNo_Width As Long
Private blnFromHidden As Boolean
Private strAccountNo As String, lngTpp As Long
Private Const strCallingForm1 As String = "frmAccountHideTrans2"
Private Const strCallingForm2 As String = "frmAccountHideTrans2_Hidden"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim datEndDate As Date
        Dim strFilter As String, strOrderBy As String
        Dim varTmp00 As Variant

110     With Me
120       If IsLoaded(strCallingForm1, acForm) = True Or IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
130         If IsNull(.OpenArgs) = False Then
140           If Trim(.OpenArgs) <> vbNullString Then

150             strAccountNo = Trim(.OpenArgs)

160             TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180             blnFromHidden = IsLoaded(strCallingForm2, acForm)  ' ** Module Functions: modFileUtilities.

                ' ** Default to PageNumOf if calling form has no check box.
190             .PageNumOf.Visible = True
200             .PageNum.Visible = False

210             lngJournalNo_Width = (.transdate.Left - .journalno.Left)
220             ShowJournalNo blnFromHidden  ' ** Procedure: Below.

230             .AsOf.Visible = True
240             .AsOf2.Visible = True
250             If .RecordSource <> "qryAccountHide_Rpt_02_02" Then
260               .RecordSource = "qryAccountHide_Rpt_02_02"
270             End If

280             lngICash_Left = .ICash.Left
290             lngCurrCode_Left = .curr_code.Left
300             lngCurrCodeLbl_Left = .curr_code_lbl.Left
310             lngCurrCodeLbl_Width = .curr_code_lbl.Width
320             lngTotDesc_Width = .totdesc.Width

330             Select Case Forms(strCallingForm1).chkIncludeCurrency
                Case True
340               .icash_str.Visible = True
350               .ICash.Visible = False
360               .pcash_str.Visible = True
370               .PCash.Visible = False
380               .cost_str.Visible = True
390               .Cost.Visible = False
400               varTmp00 = (.journaltype_lbl.Left - (.transdate_lbl.Left + .transdate_lbl.Width))
410               varTmp00 = (lngTotDesc_Width - (lngCurrCodeLbl_Width + varTmp00))
420               .totdesc.Width = varTmp00
430               .totdesc_lbl.Width = .totdesc.Width
440               .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
450               .totdesc_lbl2.Width = .totdesc_lbl.Width
460               .totdesc_lbl2_line.Width = .totdesc_lbl_line.Width
470               .curr_code.Visible = True
480               .curr_code_lbl.Visible = True
490               .curr_code_lbl_line.Visible = True
500               .curr_code_lbl2.Visible = True
510               .curr_code_lbl2_line.Visible = True
520               varTmp00 = Empty
530             Case False
540               .ICash.Visible = True
550               .icash_str.Visible = False
560               .PCash.Visible = True
570               .pcash_str.Visible = False
580               .Cost.Visible = True
590               .cost_str.Visible = False
600               .curr_code.Visible = False
610               .curr_code_lbl.Visible = False
620               .curr_code_lbl_line.Visible = False
630               .curr_code_lbl2.Visible = False
640               .curr_code_lbl2_line.Visible = False
650             End Select

660             .Name_lbl.Caption = vbNullString

670             strFilter = "[accountno] = '" & strAccountNo & "'"
680             .Filter = strFilter
690             .FilterOn = True

700             datEndDate = Date
710             .AsOf_lbl.Caption = "As of " & Format(datEndDate, "mm/dd/yyyy")
720             .AsOf_lbl2.Caption = "As of " & Format(datEndDate, "mm/dd/yyyy")

730             .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
740             .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

750             If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
760               strOrderBy = Forms(strCallingForm2).frmAccountHideTrans2_Hidden_Sub.Form.SortNow_Get  ' ** Form Function: frmAccountHideTrans2_Hidden_Sub.
                  ' ** Potential Sorts:
                  ' **   [ledghid_grpnum], [ledghid_ord]
                  ' **   [PriorPeriod], [ledghid_grpnum], [ledghid_ord]
                  ' **   [IsArch], [ledghid_grpnum], [ledghid_ord]
                  ' **   [journalno]
                  ' **   [transdate], [ledghid_grpnum], [ledghid_ord]
                  ' **   [journaltype], [ledghid_grpnum], [ledghid_ord]
                  ' **   [shareface], [ledghid_grpnum], [ledghid_ord]
                  ' **   [totdesc], [ledghid_grpnum], [ledghid_ord]
                  ' **   [icash], [ledghid_grpnum], [ledghid_ord]
                  ' **   [pcash], [ledghid_grpnum], [ledghid_ord]
                  ' **   [cost], [ledghid_grpnum], [ledghid_ord]
                  ' **   [ledger_HIDDEN], [ledghid_grpnum], [ledghid_ord]
770             Else
780               strOrderBy = Forms(strCallingForm1).frmAccountHideTrans2_Sub_List.Form.SortNow_Get  ' ** Form Function: frmAccountHideTrans2_Sub_List.
                  ' ** Potential Sorts:
                  ' **   [transdate] DESC, [JournalType_Order], [journalno]
                  ' **   [journalno]
                  ' **   [journaltype], [transdate], [journalno]
                  ' **   [shareface], [transdate], [JournalType_Order], [journalno]
                  ' **   [totdesc], [transdate], [JournalType_Order], [journalno]
                  ' **   [icash], [transdate], [JournalType_Order], [journalno]
                  ' **   [pcash], [transdate], [JournalType_Order], [journalno]
                  ' **   [cost], [transdate], [JournalType_Order], [journalno]
                  ' **   [curr_code], [transdate], [JournalType_Order], [journalno]
                  ' **   [ledger_HIDDEN], [transdate], [JournalType_Order], [journalno]
790             End If
800             .OrderBy = strOrderBy
810             .OrderByOn = True

820             varTmp00 = Forms(strCallingForm1).accountno
830             If IsNull(varTmp00) = False Then
840               varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
850               varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
860               .Caption = "Hidden_Transactions_" & varTmp00 & "_" & Format(datEndDate, "yymmdd")
870             End If

880             .icash_lbl.Visible = True
890             .icash_lbl_line.Visible = True
900             .pcash_lbl.Visible = True
910             .pcash_lbl_line.Visible = True
920             .cost_lbl.Visible = True
930             .cost_lbl_line.Visible = True

940             .icash_lbl2.Visible = True
950             .icash_lbl2_line.Visible = True
960             .pcash_lbl2.Visible = True
970             .pcash_lbl2_line.Visible = True
980             .cost_lbl2.Visible = True
990             .cost_lbl2_line.Visible = True

1000          Else
1010            Cancel = -1
1020          End If
1030        Else
1040          Cancel = -1
1050        End If
1060      Else
1070        Cancel = -1
1080      End If
1090    End With

1100    If Cancel = -1 Then
1110      Beep
1120      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1130    End If

EXITP:
1140    Exit Sub

ERRH:
1150    DoCmd.Restore
1160    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
1170    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1180      Forms(strCallingForm1).Visible = True
1190      Forms(strCallingForm2).Visible = True
1200      Forms(strCallingForm2).TimerInterval = 250&
1210    Else
1220      Forms(strCallingForm1).Visible = True
1230      Forms(strCallingForm1).TimerInterval = 250&
1240    End If
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim varTmp00 As Variant

1310    With Me
1320      varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1330      If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1340      .Name_lbl.Caption = varTmp00
1350    End With

EXITP:
1360    Exit Sub

ERRH:
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim varTmp00 As Variant

1510    With Me
1520      If .Page >= 2 Then
1530        .Section(acPageHeader).Visible = True
1540        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1550        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1560        .Name_lbl.Caption = varTmp00
1570      Else
1580        .Section(acPageHeader).Visible = False
1590      End If
1600    End With

EXITP:
1610    Exit Sub

ERRH:
1620    Select Case ERR.Number
        Case Else
1630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1640    End Select
1650    Resume EXITP

End Sub

Private Sub Report_Close()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1710    gblnReportClose = True
1720    DoCmd.Restore  ' ** Doesn't matter whether they were visible or not visible.
1730    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1740      Forms(strCallingForm1).Visible = True
1750      Forms(strCallingForm2).Visible = True
1760      Forms(strCallingForm2).TimerInterval = 250&
1770    Else
1780      Forms(strCallingForm1).Visible = True
1790      Forms(strCallingForm1).TimerInterval = 250&
1800    End If

EXITP:
1810    Exit Sub

ERRH:
1820    DoCmd.Restore
1830    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
1840    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1850      Forms(strCallingForm1).Visible = True
1860      Forms(strCallingForm2).Visible = True
1870      Forms(strCallingForm2).TimerInterval = 250&
1880    Else
1890      Forms(strCallingForm1).Visible = True
1900      Forms(strCallingForm1).TimerInterval = 250&
1910    End If
1920    Select Case ERR.Number
        Case Else
1930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1940    End Select
1950    Resume EXITP

End Sub

Private Sub ShowJournalNo(blnFromHid As Boolean)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo"

        Dim blnShow As Boolean

2010    With Me

2020      Select Case blnFromHid
          Case True
2030        blnShow = Forms(strCallingForm2).chkShowJournalNo
2040      Case False
2050        blnShow = Forms(strCallingForm1).chkShowJournalNo
2060      End Select

2070      Select Case blnShow
          Case True
            ' ** Leave everything as it is.
2080      Case False
2090        .journalno.Visible = False
2100        .journalno_lbla.Visible = False
2110        .journalno_lblb.Visible = False
2120        .journalno_lbl_line.Visible = False
2130        .journalno_lbl2a.Visible = False
2140        .journalno_lbl2b.Visible = False
2150        .journalno_lbl2_line.Visible = False
2160        .transdate.Left = (.transdate.Left - lngJournalNo_Width)
2170        .transdate_lbl.Left = .transdate.Left
2180        .transdate_lbl_line.Left = .transdate_lbl.Left
2190        .transdate_lbl2.Left = .transdate_lbl.Left
2200        .transdate_lbl2_line.Left = .transdate_lbl2.Left
2210        .journaltype.Left = (.journaltype.Left - lngJournalNo_Width)
2220        .journaltype_lbl.Left = .journaltype.Left
2230        .journaltype_lbl_line.Left = .journaltype_lbl.Left
2240        .journaltype_lbl2.Left = .journaltype_lbl.Left
2250        .journaltype_lbl2_line.Left = .journaltype_lbl2.Left
2260        .shareface.Left = (.shareface.Left - lngJournalNo_Width)
2270        .shareface_lbl.Left = .shareface.Left
2280        .shareface_lbl_line.Left = .shareface_lbl.Left
2290        .shareface_lbl2.Left = .shareface_lbl.Left
2300        .shareface_lbl2_line.Left = .shareface_lbl2.Left
2310        .totdesc.Left = (.totdesc.Left - lngJournalNo_Width)
2320        .totdesc_lbl.Left = .totdesc.Left
2330        .totdesc_lbl_line.Left = .totdesc_lbl.Left
2340        .totdesc_lbl2.Left = .totdesc_lbl.Left
2350        .totdesc_lbl2_line.Left = .totdesc_lbl2.Left
2360        .totdesc.Width = (.totdesc.Width + lngJournalNo_Width)
2370        .totdesc_lbl.Width = .totdesc.Width
2380        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
2390        .totdesc_lbl2.Width = .totdesc_lbl.Width
2400        .totdesc_lbl2_line.Width = (.totdesc_lbl2.Width + lngTpp)
2410      End Select

2420    End With

EXITP:
2430    Exit Sub

ERRH:
2440    DoCmd.Restore
2450    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
2460    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
2470      Forms(strCallingForm1).Visible = True
2480      Forms(strCallingForm2).Visible = True
2490      Forms(strCallingForm2).TimerInterval = 250&
2500    Else
2510      Forms(strCallingForm1).Visible = True
2520      Forms(strCallingForm1).TimerInterval = 250&
2530    End If
2540    Select Case ERR.Number
        Case Else
2550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2560    End Select
2570    Resume EXITP

End Sub
