VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptCA_02A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptCA_02A"

'VGC 04/04/2016: CHANGES!

Private Const strCallingForm As String = "frmRpt_CourtReports_CA"

' ** cmbAccounts combo box constants:
Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname

Private frm As Access.Form
Private lngRevLabel_Left As Long
Private blnNoData As Boolean
Private lngLegalNameMove_Height As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim ctl As Access.Control
        Dim strFilter As String
        Dim intLen As Integer
        Dim varFootnote As Variant

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Procedure: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

140         Set frm = Forms(strCallingForm)

150         Select Case frm.chkPageOf
            Case True
160           .PageNumOf.Visible = True
170           .PageNum.Visible = False
180         Case False
190           .PageNum.Visible = True
200           .PageNumOf.Visible = False
210           .PageNumOf.ControlSource = vbNullString
220         End Select

            ' ** Because the PageFooterSection does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
230         If frm.chkIncludeFootnote = True Then
240           varFootnote = frm.CourtReports_Footnote
250           If IsNull(varFootnote) = False Then
260             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
270               intLen = Len(Trim(varFootnote))
280               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
290                 .Section("PageFooterSection").Height = 480&
300               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
310                 .Section("PageFooterSection").Height = 690&
320               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
330                 .Section("PageFooterSection").Height = 900&
340               End If
350             End If
360           End If
370         End If

380         lngLegalNameMove_Height = (.AcctName_Short.Height - 120&)

            ' ** Set the account name heading.
390         If gblnLegalName = True Then
400           If IsNull(frm.cmbAccounts.Column(CBX_A_LEGAL)) = True Then
                ' ** Use shortname.
410           Else
420             If Trim(frm.cmbAccounts.Column(CBX_A_LEGAL)) = vbNullString Then
                  ' ** Use shortname.
430             Else
440               .Section(acPageHeader).Height = (.Section(acPageHeader).Height + lngLegalNameMove_Height)
450               .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
460               .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
470               .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
480               .AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height
490               .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
500               For Each ctl In .Section(acPageHeader).Controls
510                 With ctl
520                   Select Case .ControlType
                      Case acLabel
530                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Ignore.
540                     Case Else
550                       .Top = .Top + lngLegalNameMove_Height
560                     End Select
570                   Case acLine
580                     If .Width <> Me.Width Then
590                       .Top = .Top + lngLegalNameMove_Height
600                     End If
610                   End Select
620                 End With
630               Next
640               .AcctName_Legal.Visible = True
650               .AcctName_Short.Visible = False
660             End If
670           End If
680         End If

            ' ** Set the date range heading.
690         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
700         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
710         strFilter = "([date] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And " & _
              "[date] <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "# + 1) And "
720         strFilter = strFilter & "[accountno] = '" & gstrAccountNo & "' And "
730         strFilter = strFilter & "([reportnumber] = " & CStr(CRPT_RECEIPTS) & " And "
740         strFilter = strFilter & "[amount] <> 0)"
750         .Filter = strFilter
760         .FilterOn = True

770         lngRevLabel_Left = .revcode_DESC_lbl.Left

            ' ** Assume there is data to be printed.
780         blnNoData = False

790         .Caption = "CourtReport_CA_Receipts_Grouped_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

800       Else
810         Cancel = -1
820       End If
830     End With

840     If Cancel = -1 Then
850       Beep
860       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
870     End If

EXITP:
880     Set ctl = Nothing
890     Exit Sub

ERRH:
900     Select Case ERR.Number
        Case Else
910       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
920     End Select
930     Resume EXITP

End Sub

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)
' ** journaltype Header.

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader2_Format"

1010    With Me
1020      If blnNoData Then
1030        .journaltype_lbl.Visible = False
1040      Else
1050        Cancel = 0
1060        Select Case .txtjournaltype
            Case "Dividend"
1070          If .SpecialSort = 3 Then
1080            .journaltype_lbl.Caption = vbNullString
1090            Cancel = -1
1100          Else
1110            .journaltype_lbl.Caption = "Dividends"
1120          End If
1130        Case "Interest"
1140          If .SpecialSort = 3 Then
1150            .journaltype_lbl.Caption = vbNullString
1160            Cancel = -1
1170          Else
1180            .journaltype_lbl.Caption = "Interest"
1190          End If
1200        Case "Misc."
1210          .journaltype_lbl.Caption = "Miscellaneous"
1220        Case "Received"
              ' ** Received entries may have a variety of revcode_ID's, but just lump
              ' ** them all into Other Receipts, without showing their revcode_DESC.
              ' ** VGC 08/01/2009: This applies only to Received not given a specific revcode_ID.
1230          If .revcode_ID <= 2& Then
1240            If .journaltype_lbl.Caption <> "Other Receipts" Then
1250              .journaltype_lbl.Caption = "Other Receipts"
1260            End If
1270          Else
1280            .journaltype_lbl.Caption = Me.txtjournaltype
1290          End If
1300        Case Else
1310          .journaltype_lbl.Caption = Me.txtjournaltype
1320        End Select
1330      End If
1340    End With

EXITP:
1350    Exit Sub

ERRH:
1360    Select Case ERR.Number
        Case Else
1370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1380    End Select
1390    Resume EXITP

End Sub

Private Sub GroupHeader4_Format(Cancel As Integer, FormatCount As Integer)
' ** revcode_SORTORDER Header.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader4_Format"

1410    If blnNoData = True Then
1420      Cancel = -1
1430    Else
1440      With Me
1450        Cancel = 0
1460        Select Case .txtRevcode_desc
            Case "Unspecified Income"
1470          .revcode_DESC_lbl.Left = lngRevLabel_Left
1480          Select Case .txtjournaltype
              Case "Dividend"
1490            .revcode_DESC_lbl.Caption = vbNullString
1500            Cancel = -1
1510          Case "Interest"
1520            .revcode_DESC_lbl.Caption = vbNullString
1530            Cancel = -1
1540          Case "Misc."
1550            .revcode_DESC_lbl.Caption = vbNullString
1560            Cancel = -1
1570          Case "Received"
1580            If .revcode_ID <= 2& Then
1590              .revcode_DESC_lbl.Caption = vbNullString
1600              Cancel = -1
1610            End If
1620          End Select
1630        Case Else
1640          Select Case .txtjournaltype
              Case "Received"
                ' ** Received entries may have a variety of revcode_ID's, but just lump
                ' ** them all into Other Receipts, without showing their revcode_DESC.
1650            If .revcode_ID <= 2& Then
1660              .revcode_DESC_lbl.Caption = vbNullString
1670              Cancel = -1
1680            Else
1690              .revcode_DESC_lbl.Left = 0&
1700              .revcode_DESC_lbl.FontUnderline = True
1710              .revcode_DESC_lbl.Caption = .txtRevcode_desc
1720            End If
1730          Case Else
1740            .revcode_DESC_lbl.Left = 0&
1750            .revcode_DESC_lbl.FontUnderline = True
1760            .revcode_DESC_lbl.Caption = .txtRevcode_desc
1770          End Select

1780        End Select
1790      End With
1800    End If

EXITP:
1810    Exit Sub

ERRH:
1820    Select Case ERR.Number
        Case Else
1830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1840    End Select
1850    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1910    With Me
1920  On Error Resume Next
1930      strTmp01 = .totdesc
1940      If ERR.Number = 2427 Then  ' ** You entered an expression that has no value.
1950  On Error GoTo ERRH
1960        Report_NoData Cancel  ' ** Though this may be triggered anyway.
1970      Else
1980  On Error GoTo ERRH
1990      End If
2000    End With

EXITP:
2010    Exit Sub

ERRH:
2020    Select Case ERR.Number
        Case Else
2030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2040    End Select
2050    Resume EXITP

End Sub

Private Sub GroupFooter4_Format(Cancel As Integer, FormatCount As Integer)
' ** revcode_SORTORDER Footer.

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter4_Format"

2110    With Me
2120      If blnNoData Then
2130        .txtRevcodeTotal_lbl.Visible = False
2140      Else
2150        Cancel = 0
2160        If .txtRevcode_desc = "Unspecified Income" Or (.txtjournaltype = "Received" And .revcode_ID <= 2&) Then
2170          .txtRevcodeTotal_lbl.Caption = vbNullString
2180          Cancel = -1
2190        Else
2200          .txtRevcodeTotal_lbl.Caption = "Total " & .txtRevcode_desc
2210        End If
2220      End If
2230    End With

EXITP:
2240    Exit Sub

ERRH:
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)
' ** journaltype Footer.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter2_Format"

2310    With Me
2320      If blnNoData Then
2330        .txtJournalTypeTotal_lbl.Visible = False
2340      Else
2350        Cancel = 0
2360        Select Case [txtjournaltype]
            Case "Dividend"
2370          If .SpecialSort2 = 3 Then
2380            Cancel = -1
2390          Else
2400            .txtJournalTypeTotal_lbl.Caption = "Total Dividends"
2410          End If
2420        Case "Interest"
2430          If .SpecialSort2 = 3 Then
2440            Cancel = -1
2450          Else
2460            .txtJournalTypeTotal_lbl.Caption = "Total Interest"
2470          End If
2480        Case "Misc."
2490          .txtJournalTypeTotal_lbl.Caption = "Total Miscellaneous"
2500        Case "Received"
2510          If .revcode_ID <= 2& Then
2520            .txtJournalTypeTotal_lbl.Caption = "Total Other Receipts"
2530          Else
2540            .txtJournalTypeTotal_lbl.Caption = "Total " & .txtjournaltype
2550          End If
2560        Case Else
2570          .txtJournalTypeTotal_lbl.Caption = "Total " & .txtjournaltype
2580        End Select
2590      End If
2600    End With

EXITP:
2610    Exit Sub

ERRH:
2620    Select Case ERR.Number
        Case Else
2630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2640    End Select
2650    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

2710    blnNoData = True

2720    With Me
2730      .NoData_lbl.BackColor = CLR_WHT
2740      .NoData_lbl.Visible = True  ' ** Turn on the message.
2750      .transdate.Visible = False
2760      .TotalShareFace.Visible = False
2770      .totdesc.Visible = False
2780      .Gain.Visible = False
2790      .txtTotalAmount_line.Visible = False
2800      .txtTotalAmount.Visible = False
2810      .txtJournalTypeTotal_lbl.Visible = False
2820      .txtJournalTypeTotal.Visible = False
2830      .txtGrandTotal_lbl.Visible = False
2840      .txtGrandTotal_line.Visible = False
2850      .txtGrandTotal.Visible = False
2860      .txtRevcodeTotal_lbl.Visible = False
2870      .GroupHeader2.Visible = False
2880      .GroupHeader4.Visible = False
2890      .txtRevcodeTotal.Visible = False
2900      .txtRevcodeTotal_line.Visible = False
2910      .txtJournalTypeTotal_line.Visible = False
2920      If gblnLegalName = True Then
            ' ** Column 4: legalname
2930        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
2940      Else
            ' ** Column 3: shorname
2950        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
2960      End If
2970      .AcctName_Legal.Visible = False
2980      .AcctName_Short.Visible = False
2990      .AcctName_lbl.Visible = True
3000    End With

EXITP:
3010    Exit Sub

ERRH:
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Sub

Private Sub Report_Close()

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

3110    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
3120    Set frm = Nothing
3130    Exit Sub

ERRH:
3140    Select Case ERR.Number
        Case Else
3150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3160    End Select
3170    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

3210    varRetVal = Null
3220    With Me
3230      If frm.chkIncludeFootnote = True Then
3240        varRetVal = frm.CourtReports_Footnote
3250      End If
3260    End With

EXITP:
3270    ThisFootnote = varRetVal
3280    Exit Function

ERRH:
3290    varRetVal = RET_ERR
3300    Select Case ERR.Number
        Case Else
3310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3320    End Select
3330    Resume EXITP

End Function
