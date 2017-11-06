VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptAccountSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptAccountSummary"

'VGC 07/07/2016: CHANGES!

Private Const strCallingForm As String = "frmStatementParameters"

'Beginning Account Value:
'=IIf(IsNull([PreviousTotalMarketValue]),0,[PreviousTotalMarketValue])
'[TotalMarketValue] from Balance table, as of the last statement balance.

'Realized Market Adjustment for Current & Prior Periods:
'=IIf(IsNull([RMA]),0,[RMA])
'RMA = Sum((IIf([ledger].[pcash]<0,0,[ledger].[pcash])*-1)+([ledger].[cost]*-1))
'  WHERE (((ledger.journaltype)='Sold') AND ((ledger.icash)>=0) AND
'  ((ledger.transdate)>=CDate(Format([qryAccountSummary_12x].[MaxOfbalance date],'mm/dd/yyyy'))+1 And
'  (ledger.transdate)<=#9/30/2008#) AND ((ledger.pcash)>=0))
'  'Sold' transactions between last Statement and this Statement date, with Income/Principal Cash > 0.

'Change In Market Value:
'=IIf(IsNull([CurrentTotalMarketValue]),0,([CurrentTotalMarketValue]) -         $2,279.52
'    (IIf(IsNull([PreviousTotalMarketValue]),0,[PreviousTotalMarketValue]) +    $2,479.52
'    IIf(IsNull([SumPositiveIcash]),0,[SumPositiveIcash]) +                         $0.00
'    IIf(IsNull([SumPositivePcash]),0,[SumPositivePcash]) +                       $200.00
'    IIf(IsNull([SumNegativeIcash]),0,[SumNegativeIcash]) +                         $0.00
'    IIf(IsNull([SumNegativePcash]),0,[SumNegativePcash]) +                         $0.00
'    IIf(IsNull([SumPositiveCost]),0,[SumPositiveCost]) +                        ($400.00)
'    IIf(IsNull([SumNegativeCost]),0,[SumNegativeCost])) +                          $0.00
'REMOVE THIS NEXT LINE, AND SET RMA LINE ON REPORT TO Visible = False: 02/21/09, per Rich.
'    (IIf(IsNull([RMA]),0,[RMA])*-1) )                                            $200.00 * -1
'  SumPositiveIcash = Sum(IIf([ledger].[icash]<0,0,[ledger].[icash]))
'  SumPositivePcash = Sum(IIf([ledger].[pcash]<0,0,[ledger].[pcash]))
'  SumPositiveCost = Sum([ledger].[cost])
'    Numerous criteria for JournalType and Cash values, between last Statement and this Statement date.
'  SumNegativeIcash = Sum(IIf([ledger].[icash]>0,0,[ledger].[icash]))
'  SumNegativePcash = Sum(IIf([ledger].[pcash]>0,0,[ledger].[pcash]))
'  SumNegativeCost = Sum([ledger].[cost])
'    Numerous criteria for JournalType and Cash values, between last Statement and this Statement date.
'  TotalMarketValue (Balance table, last Statement) - Total positive/negative Income, Principal, Cost, and RMA,
'    between last Statement and this Statement date.

' ** VGC 11/25/2009: «='90',-1,1)» to «='90',1,1)».
'Ending Account Value:
'=IIf(IsNull([CurrentTotalMarketValue]),0,[CurrentTotalMarketValue])
'CurrentTotalMarketValue = [qryAccountSummary_11x].[TotalMarketValue]+[qryAccountSummary_11x].[IcashAndPcash]
'  TotalMarketValue = Sum([TotalShareface]*[MarketValueCurrentX])
'    TotalShareface = Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface]))*IIf([assettype].[assettype]='90',1,1)
'    MarketValueCurrentX = IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent])
'  IcashAndPcash = IIf(IsNull([account].[icash]),0,[account].[icash])+IIf(IsNull([account].[pcash]),0,[account].[pcash])

Private datStartDate As Date, datEndDate As Date
Private lngPageNum As Long
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim blnSkip As Boolean
        Dim varTmp00 As Variant
        Dim lngZ As Long

110     With Me
120       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.

130         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
140         blnSkip = False

150         Select Case Forms(strCallingForm).chkPageOf
            Case True
160           .PageNumOf.Visible = True
170           .PageNum.Visible = False
180         Case False
190           .PageNum.Visible = True
200           .PageNumOf.Visible = False
210           .PageNumOf.ControlSource = vbNullString
220         End Select

230         .MarketAsOf_lbl.Caption = vbNullString
240         .Name_lbl.Caption = vbNullString
250         .DateBeg_lbl.Caption = vbNullString
260         .DateEnd_lbl.Caption = vbNullString
900         .TotalIncreases.Visible = False
910         .TotalDecreases.Visible = False

270         lngPageNum = 0&
280         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

290         datStartDate = 0: datEndDate = 0
300         varTmp00 = Forms(strCallingForm).cmbAccounts
310         If IsNull(varTmp00) = False Then
320           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
330           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
340           If IsNull(Forms(strCallingForm).DateStart) = False Then
350             If IsDate(Forms(strCallingForm).DateStart) = True Then
360               datStartDate = CDate(Forms(strCallingForm).DateStart)
370             End If
380           End If
390           If IsNull(Forms(strCallingForm).DateEnd) = False Then
400             If IsDate(Forms(strCallingForm).DateEnd) = True Then
410               datEndDate = CDate(Forms(strCallingForm).DateEnd)
420               .Caption = "Account_Summary_" & varTmp00 & "_" & Format(CDate(datEndDate), "yymmdd")
430             End If
440           End If
450         End If

460         .AsOf_lbl.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
            '.AsOf_lbl2.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
470         .DateBeg_lbl.Caption = Format(datStartDate, "mm/dd/yyyy")
480         .DateEnd_lbl.Caption = Format(datEndDate, "mm/dd/yyyy")

500         If glngPrintRpts > 0& Then
510           For lngZ = 0& To (glngPrintRpts - 1&)
520             If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
530               If garr_varPrintRpt(PR_SUMRY, lngZ) = True Then
540                 blnSkip = True
550               Else
560                 garr_varPrintRpt(PR_SUMRY, lngZ) = CBool(True)
570               End If
580               Exit For
590             End If
600           Next
610         End If

620         If gblnCrtRpt_Zero = True Then
630           .rptAccountSummary_Sub_Increases_lbl.Visible = True
640           .NoData_Increases_lbl.Visible = True
650         End If

660         If gblnCrtRpt_ZeroDialog = True Then
670           .rptAccountSummary_Sub_Decreases_lbl.Visible = True
680           .NoData_Decreases_lbl.Visible = True
690         End If

700       Else
710         Cancel = -1
720       End If
730     End With

740     If Cancel = -1 Then
750       Beep
760       MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
770     ElseIf blnSkip = True Then
780       Cancel = -1
790     End If

EXITP:
800     Exit Sub

ERRH:
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub PageHeaderSection_Format(Cancel As Integer, FormatCount As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "PageHeaderSection_Format"

910     With Me
970       .MarketAsOf_lbl.Caption = "Market as of " & Format(gdatMarketDate, "mm/dd/yyyy")
990     End With

EXITP:
1000    Exit Sub

ERRH:
1010    Select Case ERR.Number
        Case Else
1020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1030    End Select
1040    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)
' ** accountno Header.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim varTmp00 As Variant

1110    With Me

1120      lngPageNum = 0&

1130      Select Case Forms(strCallingForm).chkLegalName
          Case True
1140        If IsNull(.legalname) = True Then
1150          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1160        Else
1170          varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
1180        End If
1190        If IsNull(varTmp00) = True Then
1200          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1210        End If
1220        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1230        .Name_lbl.Caption = varTmp00
1240      Case False
1250        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
1260        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
1270        .Name_lbl.Caption = varTmp00
1280      End Select  ' ** chkLegalName.

1290    End With

EXITP:
1300    Exit Sub

ERRH:
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Sub

Private Sub PageFooterSection_Format(Cancel As Integer, FormatCount As Integer)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "PageFooterSection_Format"

1410    lngPageNum = lngPageNum + 1&

EXITP:
1420    Exit Sub

ERRH:
1430    Select Case ERR.Number
        Case Else
1440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1450    End Select
1460    Resume EXITP

End Sub

Private Sub Report_Close()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

1510    TAReports_Access2007 False  ' ** Module Function: modReportFunctions.

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub
