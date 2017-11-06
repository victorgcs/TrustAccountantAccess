VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Hidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Hidden"

'VGC 03/25/2016: CHANGES!

Private lngICash_Left As Long, lngCurrCode_Left As Long, lngCurrCodeLbl_Width As Long, lngCurrCodeLbl_Left As Long, lngTotDesc_Width As Long

Private strAccountNo As String, lngTpp As Long
Private Const strCallingForm1 As String = "frmAccountHideTrans2"
Private Const strCallingForm2 As String = "frmAccountHideTrans2_Hidden"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim datEndDate As Date
        Dim strFilter As String
        Dim varTmp00 As Variant

110     With Me
120       If IsLoaded(strCallingForm1, acForm) = True Or IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
130         If IsNull(.OpenArgs) = False Then
140           If Trim(.OpenArgs) <> vbNullString Then

150             strAccountNo = Trim(.OpenArgs)

160             TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

                ' ** Default to PageNumOf if calling form has no check box.
180             .PageNumOf.Visible = True
190             .PageNum.Visible = False

200             Select Case Forms(strCallingForm1).chkIncludeArchive
                Case True
210               .AsOf.Visible = True
220               .AsOf2.Visible = True
230               If .RecordSource <> "qryAccountHide_Rpt_02_02" Then
240                 .RecordSource = "qryAccountHide_Rpt_02_02"
250               End If
260             Case False
270               .AsOf.Visible = False
280               .AsOf2.Visible = False
290               If .RecordSource <> "qryAccountHide_Rpt_02_01" Then
300                 .RecordSource = "qryAccountHide_Rpt_02_01"
310               End If
320             End Select  ' ** chkIncludeArchive.

330             lngICash_Left = .ICash.Left
340             lngCurrCode_Left = .curr_code.Left
350             lngCurrCodeLbl_Left = .curr_code_lbl.Left
360             lngCurrCodeLbl_Width = .curr_code_lbl.Width
370             lngTotDesc_Width = .totdesc.Width

380             Select Case Forms(strCallingForm1).chkIncludeCurrency
                Case True
390               .icash_str.Visible = True
400               .ICash.Visible = False
410               .pcash_str.Visible = True
420               .PCash.Visible = False
430               .cost_str.Visible = True
440               .Cost.Visible = False
450               varTmp00 = (.journaltype_lbl.Left - (.transdate_lbl.Left + .transdate_lbl.Width))
460               varTmp00 = (lngTotDesc_Width - (lngCurrCodeLbl_Width + varTmp00))
470               .totdesc.Width = varTmp00
480               .totdesc_lbl.Width = .totdesc.Width
490               .totdesc_lbl_line.Width = (.totdesc_lbl.Width + lngTpp)
500               .curr_code.Visible = True
510               .curr_code_lbl.Visible = True
520               .curr_code_lbl_line.Visible = True
530               varTmp00 = Empty
540             Case False
550               .ICash.Visible = True
560               .icash_str.Visible = False
570               .PCash.Visible = True
580               .pcash_str.Visible = False
590               .Cost.Visible = True
600               .cost_str.Visible = False
610               .curr_code.Visible = False
620               .curr_code_lbl.Visible = False
630               .curr_code_lbl_line.Visible = False
640             End Select

650             .Name_lbl.Caption = vbNullString

660             strFilter = "[accountno] = '" & strAccountNo & "'"
670             .Filter = strFilter
680             .FilterOn = True

690             datEndDate = Date
700             .AsOf_lbl.Caption = "As of " & Format(datEndDate, "mm/dd/yyyy")
710             .AsOf_lbl2.Caption = "As of " & Format(datEndDate, "mm/dd/yyyy")

720             .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
730             .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

740             varTmp00 = Forms(strCallingForm1).accountno
750             If IsNull(varTmp00) = False Then
760               varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
770               varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
780               .Caption = "Hidden_Transactions_" & varTmp00 & "_" & Format(datEndDate, "yymmdd")
790             End If

800             .icash_lbl.Visible = True
810             .icash_lbl_line.Visible = True
820             .pcash_lbl.Visible = True
830             .pcash_lbl_line.Visible = True
840             .cost_lbl.Visible = True
850             .cost_lbl_line.Visible = True

860           Else
870             Cancel = -1
880           End If
890         Else
900           Cancel = -1
910         End If
920       Else
930         Cancel = -1
940       End If
950     End With

960     If Cancel = -1 Then
970       Beep
980       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
990     End If

EXITP:
1000    Exit Sub

ERRH:
1010    DoCmd.Restore
1020    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
1030    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1040      Forms(strCallingForm1).Visible = True
1050      Forms(strCallingForm2).Visible = True
1060      Forms(strCallingForm2).TimerInterval = 250&
1070    Else
1080      Forms(strCallingForm1).Visible = True
1090      Forms(strCallingForm1).TimerInterval = 250&
1100    End If
1110    Select Case ERR.Number
        Case Else
1120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1130    End Select
1140    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim varTmp00 As Variant

1210    With Me
1220      varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1230      If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1240      .Name_lbl.Caption = varTmp00
1250    End With

EXITP:
1260    Exit Sub

ERRH:
1270    DoCmd.Restore
1280    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
1290    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1300      Forms(strCallingForm1).Visible = True
1310      Forms(strCallingForm2).Visible = True
1320      Forms(strCallingForm2).TimerInterval = 250&
1330    Else
1340      Forms(strCallingForm1).Visible = True
1350      Forms(strCallingForm1).TimerInterval = 250&
1360    End If
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub

Private Sub Report_Close()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1510    gblnReportClose = True
1520    DoCmd.Restore  ' ** Doesn't matter whether they were visible or not visible.
1530    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1540      Forms(strCallingForm1).Visible = True
1550      Forms(strCallingForm2).Visible = True
1560      Forms(strCallingForm2).TimerInterval = 250&
1570    Else
1580      Forms(strCallingForm1).Visible = True
1590      Forms(strCallingForm1).TimerInterval = 250&
1600    End If

EXITP:
1610    Exit Sub

ERRH:
1620    DoCmd.Restore
1630    gblnReportClose = True  ' ** Doesn't matter whether they were visible or not visible.
1640    If IsLoaded(strCallingForm2, acForm) = True Then  ' ** Module Functions: modFileUtilities.
1650      Forms(strCallingForm1).Visible = True
1660      Forms(strCallingForm2).Visible = True
1670      Forms(strCallingForm2).TimerInterval = 250&
1680    Else
1690      Forms(strCallingForm1).Visible = True
1700      Forms(strCallingForm1).TimerInterval = 250&
1710    End If
1720    Select Case ERR.Number
        Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub
