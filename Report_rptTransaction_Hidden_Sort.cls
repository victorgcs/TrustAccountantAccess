VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Hidden_Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Hidden_Sort"

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

230             Select Case Forms(strCallingForm1).chkIncludeArchive
                Case True
240               .AsOf.Visible = True
250               .AsOf2.Visible = True
260               If .RecordSource <> "qryAccountHide_Rpt_02_02" Then
270                 .RecordSource = "qryAccountHide_Rpt_02_02"
280               End If
290             Case False
300               .AsOf.Visible = False
310               .AsOf2.Visible = False
320               If .RecordSource <> "qryAccountHide_Rpt_02_01" Then
330                 .RecordSource = "qryAccountHide_Rpt_02_01"
340               End If
350             End Select  ' ** chkIncludeArchive.

360             lngICash_Left = .ICash.Left
370             lngCurrCode_Left = .curr_code.Left
380             lngCurrCodeLbl_Left = .curr_code_lbl.Left
390             lngCurrCodeLbl_Width = .curr_code_lbl.Width
400             lngTotDesc_Width = .totdesc.Width

410             Select Case Forms(strCallingForm1).chkIncludeCurrency
                Case True
420               .icash_str.Visible = True
430               .ICash.Visible = False
440               .pcash_str.Visible = True
450               .PCash.Visible = False
460               .cost_str.Visible = True
470               .Cost.Visible = False
480               varTmp00 = (.journaltype_lbl.Left - (.transdate_lbl.Left + .transdate_lbl.Width))
490               varTmp00 = (lngTotDesc_Width - (lngCurrCodeLbl_Width + varTmp00))
500               .totdesc.Width = varTmp00
510               .totdesc_lbl.Width = .totdesc.Width
520               .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
530               .totdesc_lbl2.Width = .totdesc_lbl.Width
540               .totdesc_lbl2_line.Width = .totdesc_lbl_line.Width
550               .curr_code.Visible = True
560               .curr_code_lbl.Visible = True
570               .curr_code_lbl_line.Visible = True
580               .curr_code_lbl2.Visible = True
590               .curr_code_lbl2_line.Visible = True
600               varTmp00 = Empty
610             Case False
620               .ICash.Visible = True
630               .icash_str.Visible = False
640               .PCash.Visible = True
650               .pcash_str.Visible = False
660               .Cost.Visible = True
670               .cost_str.Visible = False
680               .curr_code.Visible = False
690               .curr_code_lbl.Visible = False
700               .curr_code_lbl_line.Visible = False
710               .curr_code_lbl2.Visible = False
720               .curr_code_lbl2_line.Visible = False
730             End Select

740             .Name_lbl.Caption = vbNullString

750             strFilter = "[accountno] = '" & strAccountNo & "'"
760             .Filter = strFilter
770             .FilterOn = True

780             datEndDate = Date
790             .AsOf_lbl.Caption = "As of " & Format(datEndDate, "mm/dd/yyyy")
800             .AsOf_lbl2.Caption = "As of " & Format(datEndDate, "mm/dd/yyyy")

810             .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
820             .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

830             strOrderBy = Forms(strCallingForm1).frmAccountHideTrans2_Sub_List.Form.SortNow_Get  ' ** Form Function: frmAccountHideTrans2_Sub_List.
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
840             .OrderBy = strOrderBy
850             .OrderByOn = True

860             varTmp00 = Forms(strCallingForm1).accountno
870             If IsNull(varTmp00) = False Then
880               varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
890               varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
900               .Caption = "Hidden_Transactions_" & varTmp00 & "_" & Format(datEndDate, "yymmdd")
910             End If

920             .icash_lbl.Visible = True
930             .icash_lbl_line.Visible = True
940             .pcash_lbl.Visible = True
950             .pcash_lbl_line.Visible = True
960             .cost_lbl.Visible = True
970             .cost_lbl_line.Visible = True

980             .icash_lbl2.Visible = True
990             .icash_lbl2_line.Visible = True
1000            .pcash_lbl2.Visible = True
1010            .pcash_lbl2_line.Visible = True
1020            .cost_lbl2.Visible = True
1030            .cost_lbl2_line.Visible = True

1040          Else
1050            Cancel = -1
1060          End If
1070        Else
1080          Cancel = -1
1090        End If
1100      Else
1110        Cancel = -1
1120      End If
1130    End With

1140    If Cancel = -1 Then
1150      Beep
1160      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1170    End If

EXITP:
1180    Exit Sub

ERRH:
1190    DoCmd.Restore
1200    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
1210    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1220      Forms(strCallingForm1).Visible = True
1230      Forms(strCallingForm2).Visible = True
1240      Forms(strCallingForm2).TimerInterval = 250&
1250    Else
1260      Forms(strCallingForm1).Visible = True
1270      Forms(strCallingForm1).TimerInterval = 250&
1280    End If
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub

Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportHeader_Format"

        Dim varTmp00 As Variant

1410    With Me
1420      varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1430      If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1440      .Name_lbl.Caption = varTmp00
1450    End With

EXITP:
1460    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

        Dim varTmp00 As Variant

1610    With Me
1620      If .Page >= 2 Then
1630        .Section(acPageHeader).Visible = True
1640        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1650        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1660        .Name_lbl.Caption = varTmp00
1670      Else
1680        .Section(acPageHeader).Visible = False
1690      End If
1700    End With

EXITP:
1710    Exit Sub

ERRH:
1720    Select Case ERR.Number
        Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub

Private Sub Report_Close()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1810    gblnReportClose = True
1820    DoCmd.Restore  ' ** Doesn't matter whether they were visible or not visible.
1830    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1840      Forms(strCallingForm1).Visible = True
1850      Forms(strCallingForm2).Visible = True
1860      Forms(strCallingForm2).TimerInterval = 250&
1870    Else
1880      Forms(strCallingForm1).Visible = True
1890      Forms(strCallingForm1).TimerInterval = 250&
1900    End If

EXITP:
1910    Exit Sub

ERRH:
1920    DoCmd.Restore
1930    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
1940    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1950      Forms(strCallingForm1).Visible = True
1960      Forms(strCallingForm2).Visible = True
1970      Forms(strCallingForm2).TimerInterval = 250&
1980    Else
1990      Forms(strCallingForm1).Visible = True
2000      Forms(strCallingForm1).TimerInterval = 250&
2010    End If
2020    Select Case ERR.Number
        Case Else
2030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2040    End Select
2050    Resume EXITP

End Sub

Private Sub ShowJournalNo(blnFromHid As Boolean)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowJournalNo"

        Dim blnShow As Boolean

2110    With Me

2120      Select Case blnFromHid
          Case True
2130        blnShow = Forms(strCallingForm2).chkShowJournalNo
2140      Case False
2150        blnShow = Forms(strCallingForm1).chkShowJournalNo
2160      End Select

2170      Select Case blnShow
          Case True
            ' ** Leave everything as it is.
2180      Case False
2190        .journalno.Visible = False
2200        .journalno_lbla.Visible = False
2210        .journalno_lblb.Visible = False
2220        .journalno_lbl_line.Visible = False
2230        .journalno_lbl2a.Visible = False
2240        .journalno_lbl2b.Visible = False
2250        .journalno_lbl2_line.Visible = False
2260        .transdate.Left = (.transdate.Left - lngJournalNo_Width)
2270        .transdate_lbl.Left = .transdate.Left
2280        .transdate_lbl_line.Left = .transdate_lbl.Left
2290        .transdate_lbl2.Left = .transdate_lbl.Left
2300        .transdate_lbl2_line.Left = .transdate_lbl2.Left
2310        .journaltype.Left = (.journaltype.Left - lngJournalNo_Width)
2320        .journaltype_lbl.Left = .journaltype.Left
2330        .journaltype_lbl_line.Left = .journaltype_lbl.Left
2340        .journaltype_lbl2.Left = .journaltype_lbl.Left
2350        .journaltype_lbl2_line.Left = .journaltype_lbl2.Left
2360        .shareface.Left = (.shareface.Left - lngJournalNo_Width)
2370        .shareface_lbl.Left = .shareface.Left
2380        .shareface_lbl_line.Left = .shareface_lbl.Left
2390        .shareface_lbl2.Left = .shareface_lbl.Left
2400        .shareface_lbl2_line.Left = .shareface_lbl2.Left
2410        .totdesc.Left = (.totdesc.Left - lngJournalNo_Width)
2420        .totdesc_lbl.Left = .totdesc.Left
2430        .totdesc_lbl_line.Left = .totdesc_lbl.Left
2440        .totdesc_lbl2.Left = .totdesc_lbl.Left
2450        .totdesc_lbl2_line.Left = .totdesc_lbl2.Left
2460        .totdesc.Width = (.totdesc.Width + lngJournalNo_Width)
2470        .totdesc_lbl.Width = .totdesc.Width
2480        .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
2490        .totdesc_lbl2.Width = .totdesc_lbl.Width
2500        .totdesc_lbl2_line.Width = (.totdesc_lbl2.Width + lngTpp)
2510      End Select

2520    End With

EXITP:
2530    Exit Sub

ERRH:
2540    DoCmd.Restore
2550    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
2560    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
2570      Forms(strCallingForm1).Visible = True
2580      Forms(strCallingForm2).Visible = True
2590      Forms(strCallingForm2).TimerInterval = 250&
2600    Else
2610      Forms(strCallingForm1).Visible = True
2620      Forms(strCallingForm1).TimerInterval = 250&
2630    End If
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub
