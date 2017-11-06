VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNY_00A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNY_00A"

'VGC 01/18/2017: CHANGES!

Private lngTpp As Long
Private intPageCount As Integer
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

110     gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
120     gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
130     gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString

140     With Me
150       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Procedure: modFileUtilities.

160         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
170         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

180         Select Case Forms(strCallingForm).chkPageOf
            Case True
190           .PageNumOf.Visible = True
200           .PageNum.Visible = False
210         Case False
220           .PageNum.Visible = True
230           .PageNumOf.Visible = False
240           .PageNumOf.ControlSource = vbNullString
250         End Select

260         If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
              ' ** Report is in Preview mode.
270           Forms(strCallingForm).Visible = False
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
                    ' **  2 lines : 480&
350                 .Section("PageFooter").Height = (32& * lngTpp) '480&
360               ElseIf intLen > 140 And intLen < 211 Then
                    ' **  3 lines : 690&
370                 .Section("PageFooter").Height = (46& * lngTpp) '690&
380               ElseIf intLen > 210 Then
                    ' **  4 lines : 900&
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

1090        .Caption = "CourtReport_NY_Summary_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

            ' ** Set the ordinal and version heading.
1100        .AccountHead_lbl.Caption = Forms(strCallingForm).Ordinal & " And " & _
              Forms(strCallingForm).Version & " Account"

1110      Else
1120        Cancel = -1
1130      End If
1140    End With

1150    DoCmd.Hourglass False

1160    If Cancel = -1 Then
1170      Beep
1180      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
1190    End If

EXITP:
1200    Set ctl = Nothing
1210    Set Sec = Nothing
1220    Exit Sub

ERRH:
1230    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1240      Forms(strCallingForm).Visible = True
1250      DoCmd.SelectObject acForm, strCallingForm, False
1260      DoCmd.Restore
1270    End If
1280    Select Case ERR.Number
        Case Else
1290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1300    End Select
1310    Resume EXITP

End Sub

Private Sub GroupHeader2_Print(Cancel As Integer, PrintCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader2_Print"

        ' ** Reset intPageCount for new section to print dollarsign only on the first line of the detail section.
1410    intPageCount = 0

EXITP:
1420    Exit Sub

ERRH:
1430    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1440      Forms(strCallingForm).Visible = True
1450      DoCmd.SelectObject acForm, strCallingForm, False
1460      DoCmd.Restore
1470    End If
1480    Select Case ERR.Number
        Case Else
1490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1500    End Select
1510    Resume EXITP

End Sub

Private Sub GroupHeader4_Format(Cancel As Integer, FormatCount As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader4_Format"

1610    With Me
1620      Select Case .ReportNumber2
          Case Is = 40
1630        If Nz(DSum("tamount", "qryCourtReport_NY_00A", "reportnumber=40"), 0) <> 0 Then
1640          Cancel = False
1650        Else
1660          Cancel = True
1670        End If
1680      Case Is = 100
1690        If Nz(DSum("tamount", "qryCourtReport_NY_00A", "reportnumber=100"), 0) <> 0 Then
1700          Cancel = False
1710        Else
1720          Cancel = True
1730        End If
1740      Case Else
1750        Cancel = True
1760      End Select
1770    End With

EXITP:
1780    Exit Sub

ERRH:
1790    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1800      Forms(strCallingForm).Visible = True
1810      DoCmd.SelectObject acForm, strCallingForm, False
1820      DoCmd.Restore
1830    End If
1840    Select Case ERR.Number
        Case Else
1850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1860    End Select
1870    Resume EXITP

End Sub

Private Sub GroupHeader4_Print(Cancel As Integer, PrintCount As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader4_Print"

        ' ** Reset intPageCount for new section to print dollarsign only on the first line of the detail section.
1910    intPageCount = 0

EXITP:
1920    Exit Sub

ERRH:
1930    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
1940      Forms(strCallingForm).Visible = True
1950      DoCmd.SelectObject acForm, strCallingForm, False
1960      DoCmd.Restore
1970    End If
1980    Select Case ERR.Number
        Case Else
1990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000    End Select
2010    Resume EXITP

End Sub

Private Sub GroupHeader5_Format(Cancel As Integer, FormatCount As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader5_Format"

2110    With Me
2120      Select Case .ReportNumber
          Case Is = 40
2130        Cancel = False
2140      Case Is = 100
2150        Cancel = False
2160      Case Else
2170        Cancel = True
2180      End Select
2190    End With

EXITP:
2200    Exit Sub

ERRH:
2210    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2220      Forms(strCallingForm).Visible = True
2230      DoCmd.SelectObject acForm, strCallingForm, False
2240      DoCmd.Restore
2250    End If
2260    Select Case ERR.Number
        Case Else
2270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2280    End Select
2290    Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

2310    With Me
2320      If Nz(.revcode_DESC, vbNullString) = vbNullString Then
2330        Cancel = True
2340      Else
2350        Select Case .ReportNumber2
            Case Is = 40
2360          Cancel = False
2370        Case Is = 100
2380          Cancel = False
2390        Case Else
2400          Cancel = True
2410        End Select
2420      End If
2430    End With

EXITP:
2440    Exit Sub

ERRH:
2450    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2460      Forms(strCallingForm).Visible = True
2470      DoCmd.SelectObject acForm, strCallingForm, False
2480      DoCmd.Restore
2490    End If
2500    Select Case ERR.Number
        Case Else
2510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2520    End Select
2530    Resume EXITP

End Sub

Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Print"

2610    With Me
          ' ** Only show total line at last record of group.
2620      If .txtGroupCount = .txtDetailCount Then
2630        .Totalrevcode_DESC_line.Visible = True
2640      Else
2650        .Totalrevcode_DESC_line.Visible = False
2660      End If
          ' ** Print dollarsign only on the first line of the detail section.
2670      If PrintCount = 1 Then
2680        intPageCount = intPageCount + 1
2690        .tDollarsign4.Visible = (intPageCount = 1)
2700      End If
2710    End With

EXITP:
2720    Exit Sub

ERRH:
2730    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
2740      Forms(strCallingForm).Visible = True
2750      DoCmd.SelectObject acForm, strCallingForm, False
2760      DoCmd.Restore
2770    End If
2780    Select Case ERR.Number
        Case Else
2790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2800    End Select
2810    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

2910    With Me
2920      Select Case .tReportDivisionTitle2
          Case "Principal"
2930        .tScheduleF = "Schedule F"
2940      Case "Income"
2950        .tScheduleF = "Schedule F-1"
2960      End Select
2970    End With

EXITP:
2980    Exit Sub

ERRH:
2990    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3000      Forms(strCallingForm).Visible = True
3010      DoCmd.SelectObject acForm, strCallingForm, False
3020      DoCmd.Restore
3030    End If
3040    Select Case ERR.Number
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub

Private Sub GroupFooter4_Format(Cancel As Integer, FormatCount As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter4_Format"

3110    With Me
3120      If .ReportNumber2 = 40 Or .ReportNumber2 = 100 Then
3130        .tSchedule.Visible = False
3140        .tCategory.Left = 1700
3150        .tDollarsign1.Visible = True
3160      Else
3170        .tSchedule.Visible = True
3180        .tCategory.Left = 0
3190      End If
3200    End With

EXITP:
3210    Exit Sub

ERRH:
3220    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3230      Forms(strCallingForm).Visible = True
3240      DoCmd.SelectObject acForm, strCallingForm, False
3250      DoCmd.Restore
3260    End If
3270    Select Case ERR.Number
        Case Else
3280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3290    End Select
3300    Resume EXITP

End Sub

Private Sub GroupFooter4_Print(Cancel As Integer, PrintCount As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter4_Print"

3410    With Me
          ' ** Print dollarsign only on the first line of the detail section
3420      If PrintCount = 1 Then
3430        intPageCount = intPageCount + 1
3440        .tDollarsign1.Visible = (intPageCount = 1)
3450      End If
3460      If .ReportNumber2 = 40 Or .ReportNumber2 = 100 Then
3470        .tDollarsign1.Visible = True
3480      End If
3490    End With

EXITP:
3500    Exit Sub

ERRH:
3510    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3520      Forms(strCallingForm).Visible = True
3530      DoCmd.SelectObject acForm, strCallingForm, False
3540      DoCmd.Restore
3550    End If
3560    Select Case ERR.Number
        Case Else
3570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3580    End Select
3590    Resume EXITP

End Sub

Private Sub Report_Close()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

3610    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

3620    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3630      Forms(strCallingForm).Visible = True
3640      If Parse_Ext(CurrentAppName) = gstrExt_AppDev Then  ' ** Module Functions: modFileUtilities.
3650        DoCmd.SelectObject acForm, strCallingForm, True
3660        DoCmd.Restore
3670        gblnSetFocus = True
3680        DoCmd.SelectObject acForm, strCallingForm, False
3690      Else
3700        DoCmd.SelectObject acForm, strCallingForm, False
3710        DoCmd.Restore
3720      End If
3730      gblnMessage = False
3740    End If

EXITP:
3750    Exit Sub

ERRH:
3760    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3770      Forms(strCallingForm).Visible = True
3780      DoCmd.SelectObject acForm, strCallingForm, False
3790      DoCmd.Restore
3800    End If
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub Report_NoData(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_NoData"

3910    With Me
3920      .NoData_lbl.Visible = True
3930    End With

EXITP:
3940    Exit Sub

ERRH:
3950    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
3960      Forms(strCallingForm).Visible = True
3970      DoCmd.SelectObject acForm, strCallingForm, False
3980      DoCmd.Restore
3990    End If
4000    Select Case ERR.Number
        Case Else
4010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4020    End Select
4030    Resume EXITP

End Sub

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooter.

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

4110    varRetVal = Null
4120    With Me
4130      If Forms(strCallingForm).chkIncludeFootnote = True Then
4140        varRetVal = Forms(strCallingForm).CourtReports_Footnote
4150      End If
4160    End With

EXITP:
4170    ThisFootnote = varRetVal
4180    Exit Function

ERRH:
4190    varRetVal = RET_ERR
4200    If gblnMessage = True And Forms(strCallingForm).PopUp = True Then
4210      Forms(strCallingForm).Visible = True
4220      DoCmd.SelectObject acForm, strCallingForm, False
4230      DoCmd.Restore
4240    End If
4250    Select Case ERR.Number
        Case Else
4260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4270    End Select
4280    Resume EXITP

End Function
