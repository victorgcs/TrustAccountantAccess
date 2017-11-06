VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_00A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_00A"

'VGC 03/04/2013: CHANGES!

Private dblRemoveFromTotal As Double
Private dblTmp01 As Double

Private Const strCallingForm As String = "frmRpt_CourtReports_NS"
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim ctl As Access.Control
        Dim strFilter As String
        Dim lngLegalNameMove_Height As Long
        Dim intLen As Integer
        Dim varFootnote As Variant

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Procedure: modFileUtilities.

130         dblRemoveFromTotal = 0#
140         dblTmp01 = 0#
150         gdblCrtRpt_CostTot = 0#

160         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.

170         Select Case Forms(strCallingForm).chkPageOf
            Case True
180           .PageNumOf.Visible = True
190           .PageNum.Visible = False
200         Case False
210           .PageNum.Visible = True
220           .PageNumOf.Visible = False
230           .PageNumOf.ControlSource = vbNullString
240         End Select

            ' ** Because the PageFooterSection does not have the CanGrow property,
            ' ** and because the footnote can be up to 255 characters long,
            ' ** the section must be adjusted manually.
250         If Forms(strCallingForm).chkIncludeFootnote = True Then
260           varFootnote = Forms(strCallingForm).CourtReports_Footnote
270           If IsNull(varFootnote) = False Then
280             If Trim(varFootnote) <> vbNullString Then
                  ' ** These heights are not adjusted for Twips-Per-Pixel.
290               intLen = Len(Trim(varFootnote))
300               If intLen > 70 And intLen < 141 Then
                    ' ** 2 lines : 480&
310                 .Section("PageFooterSection").Height = 480&
320               ElseIf intLen > 140 And intLen < 211 Then
                    ' ** 3 lines : 690&
330                 .Section("PageFooterSection").Height = 690&
340               ElseIf intLen > 210 Then
                    ' ** 4 lines : 900&
350                 .Section("PageFooterSection").Height = 900&
360               End If
370             End If
380           End If
390         End If

400         lngLegalNameMove_Height = (.AcctName_Short.Height - 120&)

            ' ** Set the account name heading.
410         If gblnLegalName = True Then
420           .Section(acPageHeader).Height = (.Section(acPageHeader).Height + lngLegalNameMove_Height)
430           .Header_hline03.Top = .Header_hline03.Top + lngLegalNameMove_Height
440           .Header_hline04.Top = .Header_hline04.Top + lngLegalNameMove_Height
450           .AsOf_lbl.Top = .AsOf_lbl.Top + lngLegalNameMove_Height
460           .Header_lbl.Top = .Header_lbl.Top + lngLegalNameMove_Height
470           For Each ctl In .Section(acPageHeader).Controls
480             With ctl
490               Select Case .ControlType
                  Case acLabel
500                 Select Case .Name
                    Case "Header_lbl", "AccountHead_lbl", "AsOf_lbl"
                      ' ** Ignore.
510                 Case Else
520                   .Top = .Top + lngLegalNameMove_Height
530                 End Select
540               Case acLine
550                 If .Width <> Me.Width Then
560                   .Top = .Top + lngLegalNameMove_Height
570                 End If
580               End Select
590             End With
600           Next
610           .AcctName_Legal.Visible = True
620           .AcctName_Short.Visible = False
630         End If

            ' ** Set the date range heading.
640         .AsOf_lbl.Caption = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")

            ' ** Set the bottom-left report version footer.
650         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
660         strFilter = "[date] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And [date] < #" & _
              Format(gdatEndDate + 1, "mm/dd/yyyy") & "# And "
670         strFilter = strFilter & "[accountno] = '" & gstrAccountNo & "'"
680         .Filter = strFilter
690         .FilterOn = True

700         dblRemoveFromTotal = 0#

            ' ** Turn on the group footer for the reveune / expense code lines.
710         If gblnUseReveuneExpenseCodes = True Then
720           .GroupFooter5.Visible = True
730         Else
740           .GroupFooter5.Visible = False
750         End If

760         .Caption = "CourtReport_NS_Summary_Grouped_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

770       Else
780         Cancel = -1
790       End If
800     End With

810     If Cancel = -1 Then
820       Beep
830       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
840     End If

EXITP:
850     Set ctl = Nothing
860     Exit Sub

ERRH:
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub GroupHeader3_Format(Cancel As Integer, FormatCount As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader3_Format"

1010    With Me
1020      If gblnUseReveuneExpenseCodes = True Then
1030        Select Case .ReportNumber3h
            Case 30
1040          If .ReportDivision1 = 20 And .ReportGroup2 = 20 Then
1050            .GroupFooter3.Visible = True
1060            .RptCategoryHeaderSum.Visible = False
1070          End If
1080        Case 80
1090          If .ReportDivision1 = 60 And .ReportGroup2 = 50 Then
1100            .GroupFooter3.Visible = True
1110            .RptCategoryHeaderSum.Visible = False
1120          End If
1130        Case Else
1140          .GroupFooter3.Visible = False
1150          .RptCategoryHeaderSum.Visible = True
1160        End Select
1170      End If
1180    End With

EXITP:
1190    Exit Sub

ERRH:
1200    Select Case ERR.Number
        Case Else
1210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1220    End Select
1230    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

1310    With Me
1320      If .ReportDivision1 = 40 Then
1330        Cancel = True
1340      End If
1350    End With

EXITP:
1360    Exit Sub

ERRH:
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter2_Format"

1510    With Me
1520      If .ReportGroup2 = 30 Then
1530        Cancel = True
1540      End If
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter3_Format"

        Static intIter50 As Integer, intIter60 As Integer

1710    With Me
1720      If .ReportNumber3f = 50 Then
            ' ** Investments Made.
1730        If intIter50 = 0 Then
1740          dblRemoveFromTotal = dblRemoveFromTotal + .RptCategoryFooterSum
1750        End If
1760        intIter50 = intIter50 + 1
1770      ElseIf .ReportNumber3f = 60 Then
            ' ** Changes in Investment Holdings.
1780        If intIter60 = 0 Then
1790          dblRemoveFromTotal = dblRemoveFromTotal + .RptCategoryFooterSum
1800        End If
1810        intIter60 = intIter60 + 1
1820      End If
1830    End With

EXITP:
1840    Exit Sub

ERRH:
1850    Select Case ERR.Number
        Case Else
1860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1870    End Select
1880    Resume EXITP

End Sub

Private Sub GroupHeader5_Format(Cancel As Integer, FormatCount As Integer)
' ** Select only details for the disbursements areas.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader5_Format"

1910    With Me
1920      Select Case .ReportNumber3h
          Case 30
1930        If .ReportDivision1 = 20 And .ReportGroup2 = 20 Then
1940          Cancel = False
1950          If .revcode_DESC = "Dummy entry" Then
1960            Cancel = True  ' ** Do not print dummy entries.
1970          End If
1980        Else
1990          Cancel = True
2000        End If
2010      Case 80
2020        If .ReportDivision1 = 60 And .ReportGroup2 = 50 Then
2030          Cancel = False
2040          If .revcode_DESC = "Dummy entry" Then
2050            Cancel = True  ' ** Do not print dummy entries.
2060          End If
2070        Else
2080          Cancel = True
2090        End If
2100      Case Else
2110        Cancel = True
2120      End Select
2130    End With

EXITP:
2140    Exit Sub

ERRH:
2150    Select Case ERR.Number
        Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

End Sub

Private Sub Report_Close()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

2210    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Function CalcTotal(dblReportTotal As Double) As Double
' ** Called from control AmountSumTotal in ReportFooter.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "CalcTotal"

        Dim dblRetVal As Double

2310    If gdblCrtRpt_CostTot = 0# Then
2320      gdblCrtRpt_CostTot = dblRemoveFromTotal
2330    End If
2340    If dblRemoveFromTotal = 0# And gdblCrtRpt_CostTot <> 0# Then
2350      dblRemoveFromTotal = gdblCrtRpt_CostTot
2360    End If
2370    dblTmp01 = dblRemoveFromTotal

2380    dblRetVal = dblReportTotal - dblRemoveFromTotal
2390    dblRemoveFromTotal = 0#  ' ** Need to reset to zero in case of multiple prints.

EXITP:
2400    CalcTotal = dblRetVal
2410    Exit Function

ERRH:
2420    dblRetVal = 0#
2430    Select Case ERR.Number
        Case Else
2440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2450    End Select
2460    Resume EXITP

End Function

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "ThisFootnote"

        Dim varRetVal As Variant

2510    varRetVal = Null
2520    With Me
2530      If Forms(strCallingForm).chkIncludeFootnote = True Then
2540        varRetVal = Forms(strCallingForm).CourtReports_Footnote
2550      End If
2560    End With

EXITP:
2570    ThisFootnote = varRetVal
2580    Exit Function

ERRH:
2590    varRetVal = RET_ERR
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Function
