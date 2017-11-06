VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_09"

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

1150        .Caption = "CourtReport_NY_All_Income_Collected_" & gstrAccountNo & "_" & _
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
1850          .NoData_lbl.Visible = True  ' ** Turn on the message.
1860          .NoData_lbl.BackColor = CLR_WHT
1870        End If
1880      End If
1890    End With

EXITP:
1900    Exit Sub

ERRH:
1910    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1920      Forms(strCallingForm).Visible = True
1930      DoCmd.SelectObject acForm, strCallingForm, False
1940      DoCmd.Restore
1950    End If
1960    Select Case ERR.Number
        Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

2010    With Me
2020      If blnNoData Then
2030        .AmountJTypeSum_lbl.Visible = False
2040      Else
2050        Select Case .journaltype
            Case "Received"
2060          .AmountJTypeSum_lbl.Caption = "Total Other Income"
2070        Case "Dividend"
2080          .AmountJTypeSum_lbl.Caption = "Total Dividends"
2090        Case "Misc."
2100          .AmountJTypeSum_lbl.Caption = "Total Miscellaneous"
2110        Case Else
2120          .AmountJTypeSum_lbl.Caption = "Total " & .journaltype
2130        End Select
2140      End If
2150    End With

EXITP:
2160    Exit Sub

ERRH:
2170    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2180      Forms(strCallingForm).Visible = True
2190      DoCmd.SelectObject acForm, strCallingForm, False
2200      DoCmd.Restore
2210    End If
2220    Select Case ERR.Number
        Case Else
2230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2240    End Select
2250    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

2310    blnNoData = True

2320    With Me
2330      .GroupHeader1.Visible = False
2340      Select Case gblnLegalName
          Case True
            ' ** Column 4: legalname
2350        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
2360      Case False
            ' ** Column 3: shorname
2370        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
2380      End Select
2390      .AcctName_Legal.Visible = False
2400      .AcctName_Short.Visible = False
2410      .AcctName_lbl.Visible = True
2420    End With

EXITP:
2430    Exit Sub

ERRH:
2440    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2450      Forms(strCallingForm).Visible = True
2460      DoCmd.SelectObject acForm, strCallingForm, False
2470      DoCmd.Restore
2480    End If
2490    Select Case ERR.Number
        Case Else
2500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2510    End Select
2520    Resume EXITP

End Sub

Private Sub Report_Close()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2610    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

2620    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2630      Forms(strCallingForm).Visible = True
2640      If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
2650        DoCmd.SelectObject acForm, strCallingForm, True
2660        DoCmd.Restore
2670        gblnSetFocus = True
2680        DoCmd.SelectObject acForm, strCallingForm, False
2690      Else
2700        DoCmd.SelectObject acForm, strCallingForm, False
2710        DoCmd.Restore
2720      End If
2730      gblnMessage = False
2740    End If

EXITP:
2750    Exit Sub

ERRH:
2760    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2770      Forms(strCallingForm).Visible = True
2780      DoCmd.SelectObject acForm, strCallingForm, False
2790      DoCmd.Restore
2800    End If
2810    Select Case ERR.Number
        Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooter.

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2910    varRetVal = Null
2920    With Me
2930      If Forms(strCallingForm).chkIncludeFootnote = True Then
2940        varRetVal = Forms(strCallingForm).CourtReports_Footnote
2950      End If
2960    End With

EXITP:
2970    ThisFootnote = varRetVal
2980    Exit Function

ERRH:
2990    varRetVal = RET_ERR
3000    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3010      Forms(strCallingForm).Visible = True
3020      DoCmd.SelectObject acForm, strCallingForm, False
3030      DoCmd.Restore
3040    End If
3050    Select Case ERR.Number
        Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Function
