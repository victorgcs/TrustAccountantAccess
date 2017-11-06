VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_09A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_09A"

'VGC 01/18/2017: CHANGES!

Private lngTpp As Long
Private blnNoData As Boolean
Private lngLegalNameMove_Height As Long

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
        Dim lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long

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

410         lngLegalNameMove_Height = (.AcctName_Short.Height - (8& * lngTpp))

            ' ** Set the account name heading.
420         If gblnLegalName = True Then

430           Set Sec = .Section("PageHeaderSection")
440           Sec.Height = (Sec.Height + lngLegalNameMove_Height)
450           .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
460           .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
470           .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
480           .AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height
490           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height

              ' ** Check space before actually moving.
500           lngTmp01 = lngLegalNameMove_Height
510           lngTmp02 = .AcctName_Short.Height
520           lngTmp03 = 8&
530           blnFound = True
540           Do Until blnFound = False
550             blnFound = False
560             For Each ctl In Sec.Controls
570               With ctl
580                 If .Visible = True Then
590                   Select Case .ControlType
                      Case acLabel
600                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Already moved.
610                     Case Else
620                       If ((.Top + lngTmp01) + .Height) > Sec.Height Then
630                         blnFound = True
640                         lngTmp03 = lngTmp03 - 1&
650                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
660                       End If
670                     End Select
680                   Case acLine
690                     If InStr(.Name, "_lbl_line") > 0 Then
700                       If .Top + lngTmp01 > Sec.Height Then  ' ** Label lines have Zero height.
710                         blnFound = True
720                         lngTmp03 = lngTmp03 - 1&
730                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
740                       End If
750                     End If
760                   End Select
770                 End If
780               End With
790             Next
800             If lngTmp03 = 0& Then
810               Exit Do
820             End If
830           Loop  ' ** blnFound.

840           lngLegalNameMove_Height = lngTmp01

850           For Each ctl In Sec.Controls
860             With ctl
870               If .Visible = True Then
880                 Select Case .ControlType
                    Case acLabel
890                   Select Case .Name
                      Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                        ' ** Ignore.
900                   Case Else
910                     .Top = (.Top + lngLegalNameMove_Height)
920                   End Select
930                 Case acLine
940                   If InStr(.Name, "_lbl_line") > 0 Then
950                     .Top = (.Top + lngLegalNameMove_Height)
960                   End If
970                 End Select
980               End If
990             End With
1000          Next

1010          .AcctName_Legal.Visible = True
1020          .AcctName_Short.Visible = False

1030        End If

            ' ** Set the date range heading.
1040        .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
1050        .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
1060        strFilter = "transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & _
              Format(gdatEndDate + 1, "mm/dd/yyyy") & "# And "
1070        strFilter = strFilter & "accountno = '" & gstrAccountNo & "' And "
1080        strFilter = strFilter & "("
1090        strFilter = strFilter & "(journaltype <> 'Purchase' And journaltype <> 'Sold' And icash > 0) Or "
1100        strFilter = strFilter & "(journaltype = 'Purchase' And icash < 0 And pcash = -cost)"
1110        strFilter = strFilter & ")"
1120        .Filter = strFilter
1130        .FilterOn = True

            ' ** Assume there is data to be printed.
1140        blnNoData = False

1150        .Caption = "CourtReport_NY_All_Income_Collected_Grouped_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

1160      Else
1170        Cancel = -1
1180      End If
1190    End With

1200    DoCmd.Hourglass False

1210    If Cancel = -1 Then
1220      Beep
1230      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1240    End If

EXITP:
1250    Set ctl = Nothing
1260    Set Sec = Nothing
1270    Exit Sub

ERRH:
1280    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1290      Forms(strCallingForm).Visible = True
1300      DoCmd.SelectObject acForm, strCallingForm, False
1310      DoCmd.Restore
1320    End If
1330    Select Case ERR.Number
        Case Else
1340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1350    End Select
1360    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

1410    With Me
1420      If blnNoData Then
1430        .AmountJTypeSum_lbl.Visible = False
1440      Else
1450        Select Case .journaltype
            Case "Received"
1460          .AmountJTypeSum_lbl.Caption = "Other Income"
1470        Case "Dividend"
1480          .AmountJTypeSum_lbl.Caption = "Dividends"
1490        Case "Misc."
1500          .AmountJTypeSum_lbl.Caption = "Miscellaneous"
1510        Case Else
1520          .AmountJTypeSum_lbl.Caption = .journaltype
1530        End Select
1540      End If
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1580      Forms(strCallingForm).Visible = True
1590      DoCmd.SelectObject acForm, strCallingForm, False
1600      DoCmd.Restore
1610    End If
1620    Select Case ERR.Number
        Case Else
1630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1640    End Select
1650    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

1710    With Me
1720  On Error Resume Next  ' ** Allow processing of errors inline.
1730      If .totdesc = vbNullString Then  ' ** If there is no data in the report force an error.
1740        If ERR.Number = 2427 Then  ' ** This means the report field is empty.
1750          .transdate.Visible = False  ' ** Make everything invisible that might show an error.
1760          .TotalShareFace.Visible = False
1770          .totdesc.Visible = False
1780          .amount.Visible = False
1790          .AmountSum.Visible = False
1800          .AmountSum_line.Visible = False
1810          .AmountJTypeSum.Visible = False
1820          .AmountJTypeSum_lbl.Visible = False
1830          .AmountSumTotal.Visible = False
1840          .AmountSumTotal_line.Visible = False
1850          .TotalRevcode.Visible = False
1860          .NoData_lbl.Visible = True  ' ** Turn on the message.
1870          .NoData_lbl.BackColor = CLR_WHT
1880        End If
1890      End If
1900    End With

EXITP:
1910    Exit Sub

ERRH:
1920    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1930      Forms(strCallingForm).Visible = True
1940      DoCmd.SelectObject acForm, strCallingForm, False
1950      DoCmd.Restore
1960    End If
1970    Select Case ERR.Number
        Case Else
1980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1990    End Select
2000    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

2110    With Me
2120      If blnNoData Then
2130        .AmountJTypeSum_lbl.Visible = False
2140      Else
2150        Select Case .journaltype
            Case "Received"
2160          .AmountJTypeSum_lbl.Caption = "Total Other Income"
2170        Case "Dividend"
2180          .AmountJTypeSum_lbl.Caption = "Total Dividends"
2190        Case "Misc."
2200          .AmountJTypeSum_lbl.Caption = "Total Miscellaneous"
2210        Case Else
2220          .AmountJTypeSum_lbl.Caption = "Total " & .journaltype
2230        End Select
2240      End If
2250    End With

EXITP:
2260    Exit Sub

ERRH:
2270    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2280      Forms(strCallingForm).Visible = True
2290      DoCmd.SelectObject acForm, strCallingForm, False
2300      DoCmd.Restore
2310    End If
2320    Select Case ERR.Number
        Case Else
2330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2340    End Select
2350    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

2410    blnNoData = True

2420    With Me
2430      .GroupHeader1.Visible = False
2440      Select Case gblnLegalName
          Case True
            ' ** Column 4: legalname
2450        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
2460      Case False
            ' ** Column 3: shorname
2470        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
2480      End Select
2490      .AcctName_Legal.Visible = False
2500      .AcctName_Short.Visible = False
2510      .AcctName_lbl.Visible = True
          '.tTotalRevcode_DESClbl.Visible = False
2520      .totRevcode.Visible = False
2530      .AmountJTypeSum.Visible = False
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2570      Forms(strCallingForm).Visible = True
2580      DoCmd.SelectObject acForm, strCallingForm, False
2590      DoCmd.Restore
2600    End If
2610    Select Case ERR.Number
        Case Else
2620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2630    End Select
2640    Resume EXITP

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
