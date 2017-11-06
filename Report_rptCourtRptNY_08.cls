VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_08"

'VGC 01/18/2017: CHANGES!

' ** Statement of Income Received
' ** Schedule AA-1

' ** Income on Hand  : tAmount         : =Nz(FormRef('IncomeCash'),0)
' ** Invested Income : tInvestedIncome : =Nz(FormRef('NewInput'),0)
' ** Total           : tsumAmount      : =(Nz(FormRef('IncomeCash'),0)+Nz(FormRef('NewInput'),0))

' ** FormRef('IncomeCash') = gcurCrtRpt_NY_ICash = Nz(DLookup("[icash]", "qryCourtReport_NY_00_B_01"), 0)
' ** FormRef('NewInput') = gcurCrtRpt_NY_InputNew =
' **   ENTERED VIA frmRpt_CourtReports_NY_Input_InvestedIncome AND gblnCrtRpt_NY_InvIncChange = True
' **   THEN Rpt_CourtReports_NY_Input GETS gcurCrtRpt_NY_InputNew
' **   OTHERWISE
' **   ENTERED VIA Rpt_CourtReports_NY_Input AND gblnCrtRpt_NY_InvIncChange = True

Private lngTpp As Long
Private lngLegalNameMove_Height As Long

Private Const strCallingForm As String = "frmRpt_CourtReports_NY"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim Sec As Access.Section, ctl As Access.Control
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

1060        .Caption = "Statement_of_Income_Received_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

1070      Else
1080        Cancel = -1
1090      End If
1100    End With

1110    DoCmd.Hourglass False

1120    If Cancel = -1 Then
1130      Beep
1140      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1150    End If

EXITP:
1160    Set ctl = Nothing
1170    Set Sec = Nothing
1180    Exit Sub

ERRH:
1190    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1200      Forms(strCallingForm).Visible = True
1210      DoCmd.SelectObject acForm, strCallingForm, False
1220      DoCmd.Restore
1230    End If
1240    Select Case ERR.Number
        Case Else
1250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1260    End Select
1270    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

1310    With Me
1320      Select Case gblnLegalName
          Case True
            ' ** Column 4: legalname
1330        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(4)
1340      Case False
            ' ** Column 3: shorname
1350        .AcctName_lbl.Caption = Forms("frmRpt_CourtReports_" & Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)).cmbAccounts.Column(3)
1360      End Select
1370      .AcctName_Legal.Visible = False
1380      .AcctName_Short.Visible = False
1390      .AcctName_lbl.Visible = True
1400    End With

EXITP:
1410    Exit Sub

ERRH:
1420    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1430      Forms(strCallingForm).Visible = True
1440      DoCmd.SelectObject acForm, strCallingForm, False
1450      DoCmd.Restore
1460    End If
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub Report_Close()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1610    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

1620    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1630      Forms(strCallingForm).Visible = True
1640      If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
1650        DoCmd.SelectObject acForm, strCallingForm, True
1660        DoCmd.Restore
1670        gblnSetFocus = True
1680        DoCmd.SelectObject acForm, strCallingForm, False
1690      Else
1700        DoCmd.SelectObject acForm, strCallingForm, False
1710        DoCmd.Restore
1720      End If
1730      gblnMessage = False
1740    End If

EXITP:
1750    Exit Sub

ERRH:
1760    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1770      Forms(strCallingForm).Visible = True
1780      DoCmd.SelectObject acForm, strCallingForm, False
1790      DoCmd.Restore
1800    End If
1810    Select Case ERR.Number
        Case Else
1820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1830    End Select
1840    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooter.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

1910    varRetVal = Null
1920    With Me
1930      If Forms(strCallingForm).chkIncludeFootnote = True Then
1940        varRetVal = Forms(strCallingForm).CourtReports_Footnote
1950      End If
1960    End With

EXITP:
1970    ThisFootnote = varRetVal
1980    Exit Function

ERRH:
1990    varRetVal = RET_ERR
2000    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2010      Forms(strCallingForm).Visible = True
2020      DoCmd.SelectObject acForm, strCallingForm, False
2030      DoCmd.Restore
2040    End If
2050    Select Case ERR.Number
        Case Else
2060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2070    End Select
2080    Resume EXITP

End Function
