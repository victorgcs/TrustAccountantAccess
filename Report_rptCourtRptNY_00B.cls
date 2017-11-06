VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_00B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_00B"

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
              '.AccountHead_lbl.Top = .AccountHead_lbl.Top + lngLegalNameMove_Height
480           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height

              ' ** Check space before actually moving.
490           lngTmp01 = lngLegalNameMove_Height
500           lngTmp02 = .AcctName_Short.Height
510           lngTmp03 = 8&
520           blnFound = True
530           Do Until blnFound = False
540             blnFound = False
550             For Each ctl In Sec.Controls
560               With ctl
570                 If .Visible = True Then
580                   Select Case .ControlType
                      Case acLabel
590                     Select Case .Name
                        Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                          ' ** Already moved.
600                     Case Else
610                       If ((.Top + lngTmp01) + .Height) > Sec.Height Then
620                         blnFound = True
630                         lngTmp03 = lngTmp03 - 1&
640                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
650                       End If
660                     End Select
670                   Case acLine
680                     If InStr(.Name, "_lbl_line") > 0 Then
690                       If .Top + lngTmp01 > Sec.Height Then  ' ** Label lines have Zero height.
700                         blnFound = True
710                         lngTmp03 = lngTmp03 - 1&
720                         lngTmp01 = (lngTmp02 - (lngTmp03 * lngTpp))  ' ** 1 pixel less movement.
730                       End If
740                     End If
750                   End Select
760                 End If
770               End With
780             Next
790             If lngTmp03 = 0& Then
800               Exit Do
810             End If
820           Loop  ' ** blnFound.

830           lngLegalNameMove_Height = lngTmp01

840           For Each ctl In Sec.Controls
850             With ctl
860               If .Visible = True Then
870                 Select Case .ControlType
                    Case acLabel
880                   Select Case .Name
                      Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl", "AcctName_lbl"
                        ' ** Ignore.
890                   Case Else
900                     .Top = (.Top + lngLegalNameMove_Height)
910                   End Select
920                 Case acLine
930                   If InStr(.Name, "_lbl_line") > 0 Then
940                     .Top = (.Top + lngLegalNameMove_Height)
950                   End If
960                 End Select
970               End If
980             End With
990           Next

1000          .AcctName_Legal.Visible = True
1010          .AcctName_Short.Visible = False

1020        End If

            ' ** Set the date range heading.
1030        .AsOf_lbl.Caption = "As of " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
1040        .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
1050        strFilter = "accountno = '" & gstrAccountNo & "'"
1060        .Filter = strFilter
1070        .FilterOn = True

1080        .Caption = "CourtReport_NY_Property_on_Hand_at_Ending_of_Account_Period_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

1090      Else
1100        Cancel = -1
1110      End If
1120    End With

1130    DoCmd.Hourglass False

1140    If Cancel = -1 Then
1150      Beep
1160      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1170    End If

EXITP:
1180    Set ctl = Nothing
1190    Set Sec = Nothing
1200    Exit Sub

ERRH:
1210    DoCmd.Hourglass False
1220    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1230      Forms(strCallingForm).Visible = True
1240      DoCmd.SelectObject acForm, strCallingForm, False
1250      DoCmd.Restore
1260    End If
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

        Dim intPos01 As Integer
        Dim strTmp01 As String

1410    With Me
1420      .NoData_lbl.BackColor = CLR_WHT
1430      .NoData_lbl.Visible = True
1440      .GroupHeader2.Visible = False
1450      .GroupFooter1.Visible = False
1460      .GroupFooter2.Visible = False
1470      Select Case gblnLegalName
          Case True
            ' ** Column 4: legalname
1480        strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1490      Case False
            ' ** Column 3: shorname
1500        strTmp01 = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1510      End Select
1520      intPos01 = InStr(strTmp01, "&")
1530      Do While intPos01 > 0
1540        strTmp01 = Left(strTmp01, intPos01) & Mid(strTmp01, intPos01)  ' ** Double the ampersand.
1550        intPos01 = InStr((intPos01 + 2), strTmp01, "&")
1560      Loop
1570      .AcctName_lbl.Caption = strTmp01
1580      .AcctName_Legal.Visible = False
1590      .AcctName_Short.Visible = False
1600      .AcctName_lbl.Visible = True
1610    End With

EXITP:
1620    Exit Sub

ERRH:
1630    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1640      Forms(strCallingForm).Visible = True
1650      DoCmd.SelectObject acForm, strCallingForm, False
1660      DoCmd.Restore
1670    End If
1680    Select Case ERR.Number
        Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub Report_Close()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1810    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1820    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1830      Forms(strCallingForm).Visible = True
1840      If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
1850        DoCmd.SelectObject acForm, strCallingForm, True
1860        DoCmd.Restore
1870        gblnSetFocus = True
1880        DoCmd.SelectObject acForm, strCallingForm, False
1890      Else
1900        DoCmd.SelectObject acForm, strCallingForm, False
1910        DoCmd.Restore
1920      End If
1930      gblnMessage = False
1940    End If

EXITP:
1950    Exit Sub

ERRH:
1960    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1970      Forms(strCallingForm).Visible = True
1980      DoCmd.SelectObject acForm, strCallingForm, False
1990      DoCmd.Restore
2000    End If
2010    Select Case ERR.Number
        Case Else
2020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2030    End Select
2040    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooter.

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2110    varRetVal = Null
2120    With Me
2130      If Forms(strCallingForm).chkIncludeFootnote = True Then
2140        varRetVal = Forms(strCallingForm).CourtReports_Footnote
2150      End If
2160    End With

EXITP:
2170    ThisFootnote = varRetVal
2180    Exit Function

ERRH:
2190    varRetVal = RET_ERR
2200    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2210      Forms(strCallingForm).Visible = True
2220      DoCmd.SelectObject acForm, strCallingForm, False
2230      DoCmd.Restore
2240    End If
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Function
