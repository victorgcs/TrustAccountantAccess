VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_02"

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
1060        strFilter = "([transdate] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And " & _
              "[transdate] < #" & Format(gdatEndDate + 1, "mm/dd/yyyy") & "#) And "
1070        strFilter = strFilter & "([accountno] = '" & gstrAccountNo & "') And "
1080        strFilter = strFilter & "([journaltype] = 'Sold' AND (([pcash] + [cost]) > 0))"
1090        .Filter = strFilter
1100        .FilterOn = True

1110        .Caption = "CourtReport_NY_Increases_on_Sales_Liquidation_Or_Distribution_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

1120      Else
1130        Cancel = -1
1140      End If
1150    End With

1160    DoCmd.Hourglass False

1170    If Cancel = -1 Then
1180      Beep
1190      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1200    End If

EXITP:
1210    Set ctl = Nothing
1220    Set Sec = Nothing
1230    Exit Sub

ERRH:
1240    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1250      Forms(strCallingForm).Visible = True
1260      DoCmd.SelectObject acForm, strCallingForm, False
1270      DoCmd.Restore
1280    End If
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

        Dim strTmp01 As String

1410    With Me
1420  On Error Resume Next
1430      strTmp01 = .totdesc
1440      If ERR.Number = 2427 Then
1450  On Error GoTo ERRH
1460        .transdate.Visible = False
1470        .TotalShareFace.Visible = False
1480        .totdesc.Visible = False
1490        .netProceeds.Visible = False
1500        .ourCost.Visible = False
1510        .Gain.Visible = False
1520        .txtGrandTotal.Visible = False
1530        .txtGrandTotal_line.Visible = False
1540        .txtGrandTotal_lbl.Visible = False
1550        .NoData_lbl.Visible = True  ' ** Turn on the message.
1560        .NoData_lbl.BackColor = CLR_WHT
1570        .TotalOurCost.Visible = False
1580        .TotalOurCost_line.Visible = False
1590        .TotalNetProceeds.Visible = False
1600        .TotalNetProceeds_line.Visible = False
1610      Else
1620  On Error GoTo ERRH
1630      End If
1640    End With

EXITP:
1650    Exit Sub

ERRH:
1660    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1670      Forms(strCallingForm).Visible = True
1680      DoCmd.SelectObject acForm, strCallingForm, False
1690      DoCmd.Restore
1700    End If
1710    Select Case ERR.Number
        Case Else
1720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1730    End Select
1740    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1810    With Me
1820      Select Case gblnLegalName
          Case True
            ' ** Column 4: legalname
1830        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1840      Case False
            ' ** Column 3: shorname
1850        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1860      End Select
1870      .AcctName_Legal.Visible = False
1880      .AcctName_Short.Visible = False
1890      .AcctName_lbl.Visible = True
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

Private Sub Report_Close()

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2110    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

2120    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2130      Forms(strCallingForm).Visible = True
2140      If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
2150        DoCmd.SelectObject acForm, strCallingForm, True
2160        DoCmd.Restore
2170        gblnSetFocus = True
2180        DoCmd.SelectObject acForm, strCallingForm, False
2190      Else
2200        DoCmd.SelectObject acForm, strCallingForm, False
2210        DoCmd.Restore
2220      End If
2230      gblnMessage = False
2240    End If

EXITP:
2250    Exit Sub

ERRH:
2260    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2270      Forms(strCallingForm).Visible = True
2280      DoCmd.SelectObject acForm, strCallingForm, False
2290      DoCmd.Restore
2300    End If
2310    Select Case ERR.Number
        Case Else
2320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2330    End Select
2340    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooter.

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2410    varRetVal = Null
2420    With Me
2430      If Forms(strCallingForm).chkIncludeFootnote = True Then
2440        varRetVal = Forms(strCallingForm).CourtReports_Footnote
2450      End If
2460    End With

EXITP:
2470    ThisFootnote = varRetVal
2480    Exit Function

ERRH:
2490    varRetVal = RET_ERR
2500    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2510      Forms(strCallingForm).Visible = True
2520      DoCmd.SelectObject acForm, strCallingForm, False
2530      DoCmd.Restore
2540    End If
2550    Select Case ERR.Number
        Case Else
2560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2570    End Select
2580    Resume EXITP

End Function
