VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_05"

'VGC 01/18/2017: CHANGES!

Private lngTpp As Long
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

260         If glngTaxCode_Distribution = 0& Then
270           glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
280         End If

            ' ** Because the PageFooter does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
290         If Forms(strCallingForm).chkIncludeFootnote = True Then
300           varFootnote = Forms(strCallingForm).CourtReports_Footnote
310           If IsNull(varFootnote) = False Then
320             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
330               intLen = Len(Trim(varFootnote))
340               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
350                 .Section("PageFooter").Height = (32& * lngTpp) '480&
360               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
370                 .Section("PageFooter").Height = (46& * lngTpp) '690&
380               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
390                 .Section("PageFooter").Height = (60& * lngTpp) '900&
400               End If
410             End If
420           End If
430         End If

440         lngLegalNameMove_Height = (.AcctName_Short.Height - (8& * lngTpp))

            ' ** Set the account name heading.
450         If gblnLegalName = True Then

460           Set Sec = .Section("PageHeaderSection")
470           Sec.Height = (Sec.Height + lngLegalNameMove_Height)
480           .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
490           .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
500           .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
510           .AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height
520           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height

              ' ** Check space before actually moving.
530           lngTmp01 = lngLegalNameMove_Height
540           lngTmp02 = .AcctName_Short.Height
550           lngTmp03 = 8&
560           blnFound = True
570           Do Until blnFound = False
580             blnFound = False
590             For Each ctl In Sec.Controls
600               With ctl
610                 If .Visible = True Then
620                   Select Case .ControlType
                      Case acLabel
630                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Already moved.
640                     Case Else
650                       If ((.Top + lngTmp01) + .Height) > Sec.Height Then
660                         blnFound = True
670                         lngTmp03 = lngTmp03 - 1&
680                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
690                       End If
700                     End Select
710                   Case acLine
720                     If InStr(.Name, "_lbl_line") > 0 Then
730                       If .Top + lngTmp01 > Sec.Height Then  ' ** Label lines have Zero height.
740                         blnFound = True
750                         lngTmp03 = lngTmp03 - 1&
760                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
770                       End If
780                     End If
790                   End Select
800                 End If
810               End With
820             Next
830             If lngTmp03 = 0& Then
840               Exit Do
850             End If
860           Loop  ' ** blnFound.

870           lngLegalNameMove_Height = lngTmp01

880           For Each ctl In Sec.Controls
890             With ctl
900               If .Visible = True Then
910                 Select Case .ControlType
                    Case acLabel
920                   Select Case .Name
                      Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                        ' ** Ignore.
930                   Case Else
940                     .Top = (.Top + lngLegalNameMove_Height)
950                   End Select
960                 Case acLine
970                   If InStr(.Name, "_lbl_line") > 0 Then
980                     .Top = (.Top + lngLegalNameMove_Height)
990                   End If
1000                End Select
1010              End If
1020            End With
1030          Next

1040          .AcctName_Legal.Visible = True
1050          .AcctName_Short.Visible = False

1060        End If

            ' ** Set the date range heading.
1070        .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
1080        .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
1090        strFilter = "transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & _
              Format(gdatEndDate, "mm/dd/yyyy") & "# And "
1100        strFilter = strFilter & "accountno = '" & gstrAccountNo & "' And "
1110        strFilter = strFilter & "((journaltype = 'Paid' And pcash <> 0 And taxcode = " & CStr(glngTaxCode_Distribution) & ") Or "  '= "Distribution"
1120        strFilter = strFilter & "(journaltype = 'Withdrawn' And taxcode = " & CStr(glngTaxCode_Distribution) & "))"  '= "Distribution"
1130        .Filter = strFilter    '####  TAXCODE  ####
1140        .FilterOn = True

1150        .OrderBy = "transdate"
1160        .OrderByOn = True

1170        .Caption = "CourtReport_NY_Distributions_of_Principal_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

1180      Else
1190        Cancel = -1
1200      End If
1210    End With

1220    DoCmd.Hourglass False

1230    If Cancel = -1 Then
1240      Beep
1250      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1260    End If

EXITP:
1270    Set ctl = Nothing
1280    Set Sec = Nothing
1290    Exit Sub

ERRH:
1300    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1310      Forms(strCallingForm).Visible = True
1320      DoCmd.SelectObject acForm, strCallingForm, False
1330      DoCmd.Restore
1340    End If
1350    Select Case ERR.Number
        Case Else
1360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1370    End Select
1380    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

1410    With Me
1420  On Error Resume Next  ' ** Allow processing of errors inline.
1430      If .totdesc = vbNullString Then  ' ** If there is no data in the report force an error.
1440        If ERR.Number = 2427 Then  ' ** This means the report field is empty.
1450          .totdesc.Visible = False  ' ** Make everything invisible that might show an error.
1460          .amount.Visible = False
1470          .AmountTotalSumCost_lbl.Visible = False
1480          .AmountTotalSumCost.Visible = False
1490          .AmountTotalSumCost_line.Visible = False
1500          .NoData_lbl.Visible = True  ' ** Turn on the message.
1510          .NoData_lbl.BackColor = CLR_WHT
1520        End If
1530      End If
1540    End With

EXITP:
1550    Exit Sub

ERRH:
1560    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1570      Forms(strCallingForm).Visible = True
1580      DoCmd.SelectObject acForm, strCallingForm, False
1590      DoCmd.Restore
1600    End If
1610    Select Case ERR.Number
        Case Else
1620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1630    End Select
1640    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1710    With Me
1720      Select Case gblnLegalName
          Case True
            ' ** Column 4: legalname
1730        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1740      Case False
            ' ** Column 3: shorname
1750        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1760      End Select
1770      .AcctName_Legal.Visible = False
1780      .AcctName_Short.Visible = False
1790      .AcctName_lbl.Visible = True
1800    End With

EXITP:
1810    Exit Sub

ERRH:
1820    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1830      Forms(strCallingForm).Visible = True
1840      DoCmd.SelectObject acForm, strCallingForm, False
1850      DoCmd.Restore
1860    End If
1870    Select Case ERR.Number
        Case Else
1880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1890    End Select
1900    Resume EXITP

End Sub

Private Sub Report_Close()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2010    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

2020    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2030      Forms(strCallingForm).Visible = True
2040      If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
2050        DoCmd.SelectObject acForm, strCallingForm, True
2060        DoCmd.Restore
2070        gblnSetFocus = True
2080        DoCmd.SelectObject acForm, strCallingForm, False
2090      Else
2100        DoCmd.SelectObject acForm, strCallingForm, False
2110        DoCmd.Restore
2120      End If
2130      gblnMessage = False
2140    End If

EXITP:
2150    Exit Sub

ERRH:
2160    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2170      Forms(strCallingForm).Visible = True
2180      DoCmd.SelectObject acForm, strCallingForm, False
2190      DoCmd.Restore
2200    End If
2210    Select Case ERR.Number
        Case Else
2220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2230    End Select
2240    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooter.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2310    varRetVal = Null
2320    With Me
2330      If Forms(strCallingForm).chkIncludeFootnote = True Then
2340        varRetVal = Forms(strCallingForm).CourtReports_Footnote
2350      End If
2360    End With

EXITP:
2370    ThisFootnote = varRetVal
2380    Exit Function

ERRH:
2390    varRetVal = RET_ERR
2400    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2410      Forms(strCallingForm).Visible = True
2420      DoCmd.SelectObject acForm, strCallingForm, False
2430      DoCmd.Restore
2440    End If
2450    Select Case ERR.Number
        Case Else
2460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2470    End Select
2480    Resume EXITP

End Function
