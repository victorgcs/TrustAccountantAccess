VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptCourtRptNS_00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptCourtRptNS_00"

'VGC 12/07/2014: CHANGES!

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

            ' ** Set the ordinal and version heading.
650         .AccountHead_lbl.Caption = Forms(strCallingForm).Ordinal & " And " & _
              Forms(strCallingForm).Version & " Account"

            ' ** Set the bottom-left report version footer.
660         .RptVer_lbl.Caption = Mid(THIS_NAME, (Len("rptCourtRpt") + 1), 2)

            ' ** Set the date and account filter.
670         strFilter = "[date] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And [date] < #" & _
              Format(gdatEndDate + 1, "mm/dd/yyyy") & "# And "
680         strFilter = strFilter & "[accountno] = '" & gstrAccountNo & "'"
690         .Filter = strFilter
700         .FilterOn = True

            ' ** Turn on the group footer for the reveune / expense code lines.
710         If gblnUseReveuneExpenseCodes = True Then
720           .GroupFooter5.Visible = True
730         Else
740           .GroupFooter5.Visible = False
750         End If

760         .Caption = "CourtReport_NS_Summary_" & gstrAccountNo & "_" & _
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
        'Debug.Print "'" & ERR.Number
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

1010    With Me
1020      If .ReportDivision1 = 40 Then
1030        Cancel = True
1040      End If
1050    End With

EXITP:
1060    Exit Sub

ERRH:
1070    Select Case ERR.Number
        Case Else
1080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1090    End Select
1100    Resume EXITP

End Sub

Private Sub GroupFooter2_Format(Cancel As Integer, FormatCount As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter2_Format"

1210    With Me
1220      If .ReportGroup2 = 30 Then
1230        Cancel = True
1240      End If
1250    End With

EXITP:
1260    Exit Sub

ERRH:
1270    Select Case ERR.Number
        Case Else
1280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1290    End Select
1300    Resume EXITP

End Sub

Private Sub GroupFooter3_Format(Cancel As Integer, FormatCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter3_Format"

        Static intIter50 As Integer, intIter60 As Integer

1410    With Me
1420      If .ReportNumber3 = 50 Then
            ' ** Investments Made.
1430        If intIter50 = 0 Then
1440          dblRemoveFromTotal = dblRemoveFromTotal + .Amount_RptNum_Sum
1450        End If
1460        intIter50 = intIter50 + 1
1470      ElseIf .ReportNumber3 = 60 Then
            ' ** Changes in Investment Holdings.
1480        If intIter60 = 0 Then
1490          dblRemoveFromTotal = dblRemoveFromTotal + .Amount_RptNum_Sum
1500        End If
1510        intIter60 = intIter60 + 1
1520      End If
1530    End With

EXITP:
1540    Exit Sub

ERRH:
1550    Select Case ERR.Number
        Case Else
1560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1570    End Select
1580    Resume EXITP

End Sub

Private Sub GroupFooter5_Format(Cancel As Integer, FormatCount As Integer)
' ** Select only details for the disbursements areas

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter5_Format"

        'Debug.Print "'5 " & THIS_PROC & "()"
1610    With Me
1620      Select Case .ReportNumber5
          Case 30
1630        If .ReportDivision5 = 20 And .ReportGroup5 = 20 Then
1640          Cancel = False
1650          If .revcode_DESC = "Dummy entry" Then
1660            Cancel = True  ' ** Do not print dummy entries.
1670          End If
1680        Else
1690          Cancel = True
1700        End If
1710      Case 80
1720        If .ReportDivision5 = 60 And .ReportGroup5 = 50 Then
1730          Cancel = False
1740          If .revcode_DESC = "Dummy entry" Then
1750            Cancel = True  ' ** Do not print dummy entries.
1760          End If
1770        Else
1780          Cancel = True
1790        End If
1800      Case Else
1810        Cancel = True
1820      End Select
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

Private Sub Report_Close()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1910    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Function CalcTotal(dblReportTotal As Double) As Double

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "CalcTotal"

        Dim dblRetVal As Double

2010    If gdblCrtRpt_CostTot = 0# Then
2020      gdblCrtRpt_CostTot = dblRemoveFromTotal
2030    End If
2040    If dblRemoveFromTotal = 0# And gdblCrtRpt_CostTot <> 0# Then
2050      dblRemoveFromTotal = gdblCrtRpt_CostTot
2060    End If
2070    dblTmp01 = dblRemoveFromTotal

2080    dblRetVal = dblReportTotal - dblRemoveFromTotal
2090    dblRemoveFromTotal = 0#  ' ** Need to reset to zero in case of multiple prints.

EXITP:
2100    CalcTotal = dblRetVal
2110    Exit Function

ERRH:
2120    dblRetVal = 0#
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Function

Private Function RemoveAmount() As Double

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "RemoveAmount"

        Dim dblRetVal As Double

2210    dblRetVal = dblTmp01

EXITP:
2220    RemoveAmount = dblRetVal
2230    Exit Function

ERRH:
2240    dblRetVal = 0#
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Function

Public Function ThisFootnote() As Variant
' ** This is called by the CourtReports_Footnote control in the PageFooterSection.

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
2400    Select Case ERR.Number
        Case Else
2410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2420    End Select
2430    Resume EXITP

End Function
