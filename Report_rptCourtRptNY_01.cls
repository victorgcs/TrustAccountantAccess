VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_01"

'VGC 01/18/2017: CHANGES!

Private lngTpp As Long

Private Const strCallingForm As String = "frmRpt_CourtReports_NY"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim Sec As Access.Section, ctl As Access.Control
        Dim strFilter As String
        Dim lngLegalNameMove_Height As Long
        Dim intLen As Integer
        Dim varFootnote As Variant
        Dim blnFound As Boolean
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Procedure: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
140         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

150         Select Case Forms(strCallingForm).chkPageOf
            Case True
160           .PageNumOf.Visible = True
170           .PageNum.Visible = False
180         Case False
190           .PageNum.Visible = True
200           .PageNumOf.Visible = False
210           .PageNumOf.ControlSource = vbNullString
220         End Select

230         If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
              ' ** Report is in Preview mode.
240           Forms(strCallingForm).Visible = False
250         End If

            ' ** Because the PageFooter does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
260         If Forms(strCallingForm).chkIncludeFootnote = True Then
270           varFootnote = Forms(strCallingForm).CourtReports_Footnote
280           If IsNull(varFootnote) = False Then
290             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
300               intLen = Len(Trim(varFootnote))
310               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
320                 .Section("PageFooter").Height = (32& * lngTpp) '480&
330               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
340                 .Section("PageFooter").Height = (46& * lngTpp) '690&
350               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
360                 .Section("PageFooter").Height = (60& * lngTpp) '900&
370               End If
380             End If
390           End If
400         End If

410         varTmp00 = DSum("[TotalCost]", "tmpAssetList2")
420         Select Case IsNull(varTmp00)
            Case True
430           .GroupHeader2.Visible = False
440           .Detail.Visible = False
450           .GroupFooter2.Visible = False
460         Case False
470           If varTmp00 = 0 Then
480             .GroupHeader2.Visible = False
490             .Detail.Visible = False
500             .GroupFooter2.Visible = False
510           End If
520         End Select

530         lngLegalNameMove_Height = (.AcctName_Short.Height - (8& * lngTpp))

            ' ** Set the account name heading.
540         If gblnLegalName = True Then

550           Set Sec = .Section("PageHeaderSection")
560           Sec.Height = (Sec.Height + lngLegalNameMove_Height)
570           .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
580           .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
590           .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
600           .AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height

              ' ** Check space before actually moving.
610           lngTmp01 = lngLegalNameMove_Height
620           lngTmp02 = .AcctName_Short.Height
630           lngTmp03 = 8&
640           blnFound = True
650           Do Until blnFound = False
660             blnFound = False
670             For Each ctl In Sec.Controls
680               With ctl
690                 If .Visible = True Then
700                   Select Case .ControlType
                      Case acLabel
710                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Already moved.
720                     Case Else
730                       If ((.Top + lngTmp01) + .Height) > Sec.Height Then
740                         blnFound = True
750                         lngTmp03 = lngTmp03 - 1&
760                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
770                       End If
780                     End Select
790                   Case acLine
800                     If InStr(.Name, "_lbl_line") > 0 Then
810                       If .Top + lngTmp01 > Sec.Height Then  ' ** Label lines have Zero height.
820                         blnFound = True
830                         lngTmp03 = lngTmp03 - 1&
840                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
850                       End If
860                     End If
870                   End Select
880                 End If
890               End With
900             Next
910             If lngTmp03 = 0& Then
920               Exit Do
930             End If
940           Loop  ' ** blnFound.

950           lngLegalNameMove_Height = lngTmp01

960           For Each ctl In Sec.Controls
970             With ctl
980               If .Visible = True Then
990                 Select Case .ControlType
                    Case acLabel
1000                  Select Case .Name
                      Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                        ' ** Ignore.
1010                  Case Else
1020                    .Top = (.Top + lngLegalNameMove_Height)
1030                  End Select
1040                Case acLine
1050                  If InStr(.Name, "_lbl_line") > 0 Then
1060                    .Top = (.Top + lngLegalNameMove_Height)
1070                  End If
1080                End Select
1090              End If
1100            End With
1110          Next
1120          .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
1130          .AcctName_Legal.Visible = True
1140          .AcctName_Short.Visible = False

1150        End If

            ' ** Set the date range heading.
1160        .AsOf_lbl.Caption = "As of " & Format(gdatStartDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
1170        .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Leave off the Schedule.
            '.AccountHead_lbl.Caption = vbNullString

            ' ** Set the date and account filter.
1180        strFilter = "accountno = '" & gstrAccountNo & "'"
1190        .Filter = strFilter
1200        .FilterOn = True

1210        .Caption = "CourtReport_NY_Principal_Received_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

1220      Else
1230        Cancel = -1
1240      End If
1250    End With

1260    DoCmd.Hourglass False

1270    If Cancel = -1 Then
1280      Beep
1290      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1300    End If

EXITP:
1310    Set ctl = Nothing
1320    Set Sec = Nothing
1330    Exit Sub

ERRH:
1340    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1350      Forms(strCallingForm).Visible = True
1360      DoCmd.SelectObject acForm, strCallingForm, False
1370      DoCmd.Restore
1380    End If
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

        Dim lngTmp01 As Long

1510    With Me
1520      If .TotalCost_Sum1 = 0 Then
1530        lngTmp01 = 0&
1540        If IsNull(.ICash) = False Then
1550          If .ICash <> 0 Then lngTmp01 = lngTmp01 + 1&
1560        End If
1570        If IsNull(.PCash) = False Then
1580          If .PCash <> 0 Then lngTmp01 = lngTmp01 + 1&
1590        End If
1600        If gcurCrtRpt_NY_InputNew <> 0@ Then lngTmp01 = lngTmp01 + 1&
1610        If lngTmp01 = 0& Then
1620          .TotalCost_Sum2.Visible = False
1630          .TotalCost_Sum2_line.Visible = False
1640          .txtIcash.Visible = False
1650          .txtPcash.Visible = False
1660          .TotalCash.Visible = False
1670          .TotalCash_line.Visible = False
1680          .tInvestedIncome.Visible = False
1690          .txtLessIncomeCash.Visible = False
1700          .grandTotalCost_line.Top = (.TotalCost_Sum2_line.Top + (8& * lngTpp))
1710          .grandTotalCost_line.Visible = False
1720          .grandTotalCost.Top = (.TotalCost_Sum2.Top + (8& * lngTpp))
1730          .grandTotalCost_lbl.Top = (.TotalCost_Sum2_lbl.Top + (8& * lngTpp))
1740          .txtIcash.Top = 0&
1750          .txtIcash_lbl.Top = 0&
1760          .txtPcash.Top = 0&
1770          .txtPcash_lbl.Top = 0&
1780          .TotalCash.Top = 0&
1790          .TotalCash_lbl.Top = 0&
1800          .TotalCash_line.Top = 0&
1810          .tInvestedIncome.Top = 0&
1820          .tInvestedIncome_lbl.Top = 0&
1830          .txtLessIncomeCash.Top = 0&
1840          .txtLessIncomeCash_lbl.Top = 0&
1850          .txtCash_lbl.Top = 0&
1860          .txtCash2.Top = 0&
1870          .txtCash.Top = 0&
1880          .group_footer1_height_line.Top = 0&
1890          .GroupFooter1.Height = ((.grandTotalCost.Top + .grandTotalCost.Height) + (8& * lngTpp))
1900        End If
1910      End If
1920    End With

EXITP:
1930    Exit Sub

ERRH:
1940    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1950      Forms(strCallingForm).Visible = True
1960      DoCmd.SelectObject acForm, strCallingForm, False
1970      DoCmd.Restore
1980    End If
1990    Select Case ERR.Number
        Case Else
2000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2010    End Select
2020    Resume EXITP

End Sub

Private Sub ReportFooter_Format(Cancel As Integer, FormatCount As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ReportFooter_Format"

        ' ** qryCourtReport_NY_01_44 (qryCourtReport_07 (qryCourtReport_06 (Ledger, linked to Account,
        ' ** qryCourtReport_05 (Balance, grouped by accountno, with max Balance Date), with add'l fields),
        ' ** grouped and summed); excluding Misc. [icash]+[pcash]=0), grouped and summed.
2110    Cancel = IsNull(DSum("TotAmount", "qryCourtReport_NY_01_45"))

EXITP:
2120    Exit Sub

ERRH:
2130    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2140      Forms(strCallingForm).Visible = True
2150      DoCmd.SelectObject acForm, strCallingForm, False
2160      DoCmd.Restore
2170    End If
2180    Select Case ERR.Number
        Case Else
2190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2200    End Select
2210    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

        Dim intPos01 As Integer
        Dim strTmp01 As String

2310    With Me
2320      .NoData_lbl.BackColor = CLR_WHT
2330      .NoData_lbl.Visible = True
2340      .GroupHeader2.Visible = False
2350      .GroupFooter1.Visible = False
2360      .GroupFooter2.Visible = False
2370      Select Case gblnLegalName
          Case True
            ' ** Column 4: legalname
2380        strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
2390      Case False
            ' ** Column 3: shorname
2400        strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
2410      End Select
2420      intPos01 = InStr(strTmp01, "&")
2430      Do While intPos01 > 0
2440        strTmp01 = Left(strTmp01, intPos01) & Mid(strTmp01, intPos01)  ' ** Double the ampersand.
2450        intPos01 = InStr((intPos01 + 2), strTmp01, "&")
2460      Loop
2470      .AcctName_lbl.Caption = strTmp01
2480      .AcctName_Legal.Visible = False
2490      .AcctName_Short.Visible = False
2500      .AcctName_lbl.Visible = True
2510      .tTotal.Visible = False
2520      .tTotal_lbl.Visible = False
2530    End With

EXITP:
2540    Exit Sub

ERRH:
2550    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2560      Forms(strCallingForm).Visible = True
2570      DoCmd.SelectObject acForm, strCallingForm, False
2580      DoCmd.Restore
2590    End If
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Private Sub Report_Close()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2710    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

2720    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2730      Forms(strCallingForm).Visible = True
2740      If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
2750        DoCmd.SelectObject acForm, strCallingForm, True
2760        DoCmd.Restore
2770        gblnSetFocus = True
2780        DoCmd.SelectObject acForm, strCallingForm, False
2790      Else
2800        DoCmd.SelectObject acForm, strCallingForm, False
2810        DoCmd.Restore
2820      End If
2830      gblnMessage = False
2840    End If

EXITP:
2850    Exit Sub

ERRH:
2860    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2870      Forms(strCallingForm).Visible = True
2880      DoCmd.SelectObject acForm, strCallingForm, False
2890      DoCmd.Restore
2900    End If
2910    Select Case ERR.Number
        Case Else
2920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2930    End Select
2940    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooter.

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

3010    varRetVal = Null
3020    With Me
3030      If Forms(strCallingForm).chkIncludeFootnote = True Then
3040        varRetVal = Forms(strCallingForm).CourtReports_Footnote
3050      End If
3060    End With

EXITP:
3070    ThisFootnote = varRetVal
3080    Exit Function

ERRH:
3090    varRetVal = RET_ERR
3100    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3110      Forms(strCallingForm).Visible = True
3120      DoCmd.SelectObject acForm, strCallingForm, False
3130      DoCmd.Restore
3140    End If
3150    Select Case ERR.Number
        Case Else
3160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3170    End Select
3180    Resume EXITP

End Function
