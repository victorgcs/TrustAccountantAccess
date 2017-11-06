VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Statement_ForEx_SortDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Statement_ForEx_SortDate"

'VGC 11/27/2016: CHANGES!

'gstrFormQuerySpec = "frmStatementParameters"

Private Const strCallingForm As String = "frmStatementParameters"

' ** Array: arr_varFld().
Private lngFlds As Long, arr_varFld() As Variant
Private Const F_ELEMS As Integer = 9  ' ** Array's first-element UBound().
Private Const F_ACTNO As Integer = 0
Private Const F_DATE  As Integer = 1
Private Const F_FOREX As Integer = 2
Private Const F_ICSHB As Integer = 3
Private Const F_PCSHB As Integer = 4
Private Const F_COSTB As Integer = 5
Private Const F_ICSHE As Integer = 6
Private Const F_PCSHE As Integer = 7
Private Const F_COSTE As Integer = 8
Private Const F_RECS  As Integer = 9

Private strAccountNo As String
Private datStartDate As Date, datEndDate As Date
Private blnStatements As Boolean, blnSaveData As Boolean
Private strThisProc As String
Private varTmp00 As Variant
' **

Private Sub Report_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Open"

        Dim frm As Access.Form
        Dim blnSkip As Boolean
        Dim lngZ As Long

110     strThisProc = THIS_PROC

120     With Me
130       If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Functions: modFileUtilities.

140         TAReports_Access2007 True  ' ** Module Function: modReportFunctions.
150         blnSkip = False: blnSaveData = False

160         Set frm = Forms(strCallingForm)

170         If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
180           If frm.PopUp = True Then
190             frm.Visible = False
200           End If
210         End If

220         Select Case frm.opgAccountNumber
            Case frm.opgAccountNumber_optSpecified.OptionValue
230           Select Case frm.chkTransactions
              Case True
240             If .RecordSource <> "qryStatementParameters_Trans_02_08" Then
250               .RecordSource = "qryStatementParameters_Trans_02_08"
260             End If
270           Case False
                ' ** Let's try this.
280             If .RecordSource <> "qryStatementParameters_Trans_02_04" Then
290               .RecordSource = "qryStatementParameters_Trans_02_04"
300             End If
310           End Select
320         Case frm.opgAccountNumber_optAll.OptionValue
330           Select Case frm.chkTransactions
              Case True
340             If .RecordSource <> "qryStatementParameters_Trans_02_06" Then
350               .RecordSource = "qryStatementParameters_Trans_02_06"
360             End If
370           Case False
                ' ** Let's try this other one.
380             If .RecordSource <> "qryStatementParameters_Trans_02_03" Then
390               .RecordSource = "qryStatementParameters_Trans_02_03"
400             End If
410           End Select
420         End Select

430         Select Case frm.chkPageOf
            Case True
440           .PageNumOf.Visible = True
450           .PageNum.Visible = False
460         Case False
470           .PageNum.Visible = True
480           .PageNumOf.Visible = False
490           .PageNumOf.ControlSource = vbNullString
500         End Select

510         .Name_lbl.Caption = vbNullString

520         Select Case frm.opgAccountNumber
            Case frm.opgAccountNumber_optSpecified.OptionValue
530           .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
540         Case frm.opgAccountNumber_optAll.OptionValue
550           .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
560         End Select

570         Select Case frm.chkTransactions
            Case True
580           blnStatements = False
590           datStartDate = CDate(frm.TransDateStart)
600           datEndDate = CDate(frm.TransDateEnd)
610           .AsOf_lbl.Caption = "From " & Format(datStartDate, "mm/dd/yyyy") & " To " & Format(datEndDate, "mm/dd/yyyy")
620           .AsOf_lbl2.Caption = "From " & Format(datStartDate, "mm/dd/yyyy") & " To " & Format(datEndDate, "mm/dd/yyyy")
630           Select Case frm.chkArchiveOnly_Trans
              Case True
640             .AsOf.Visible = True
650             .AsOf2.Visible = True
660           Case False
670             .AsOf.Visible = False
680             .AsOf2.Visible = False
690           End Select  ' ** chkArchiveOnly_Trans.
700         Case False
710           blnStatements = True
720           datStartDate = 0
730           datEndDate = CDate(frm.DateEnd)
740           .AsOf_lbl.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
750           .AsOf_lbl2.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
760           .AsOf.Visible = False
770           .AsOf2.Visible = False
780         End Select  ' ** chkTransactions.

790         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
800         .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

810         varTmp00 = frm.cmbAccounts
820         If IsNull(varTmp00) = False Then
830           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
840           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
850           If CLng(datStartDate) = 0 Then
860             .Caption = "Transaction_Statement_" & varTmp00 & "_" & Format(datEndDate, "yymmdd")
870           Else
880             .Caption = "Transaction_Statement_" & varTmp00 & "_" & Format(datStartDate, "yymmdd") & "_To_" & Format(datEndDate, "yymmdd")
890           End If
900           strAccountNo = varTmp00
910         End If

920         If frm.chkDevSaveData.Visible = True And frm.chkDevSaveData = True Then
930           blnSaveData = True
940           lngFlds = 1&
950           ReDim arr_varFld(F_ELEMS, 0)
              ' *************************************************
              ' ** Array: arr_varFld()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     accountno       F_ACTNO
              ' **     2       1     st_period       F_DATE
              ' **     3       2     st_forex        F_FOREX
              ' **     4       3     st_icash_beg    F_ICSHB
              ' **     5       4     st_pcash_beg    F_PCSHB
              ' **     6       5     st_cost_beg     F_COSTB
              ' **     7       6     st_icash_end    F_ICSHE
              ' **     8       7     st_pcash_end    F_PCSHE
              ' **     9       8     st_cost_end     F_COSTE
              ' **    10       9     st_recs         F_RECS
              ' **
              ' *************************************************
960           Select Case frm.opgAccountNumber
              Case frm.opgAccountNumber_optSpecified.OptionValue
970             arr_varFld(F_ACTNO, 0) = strAccountNo
980             arr_varFld(F_DATE, 0) = datEndDate
990             arr_varFld(F_FOREX, 0) = CBool(True)
1000            arr_varFld(F_ICSHB, 0) = CDbl(0)
1010            arr_varFld(F_PCSHB, 0) = CDbl(0)
1020            arr_varFld(F_COSTB, 0) = CDbl(0)
1030            arr_varFld(F_ICSHE, 0) = CDbl(0)
1040            arr_varFld(F_PCSHE, 0) = CDbl(0)
1050            arr_varFld(F_COSTE, 0) = CDbl(0)
1060            arr_varFld(F_RECS, 0) = CLng(0)
1070          Case frm.opgAccountNumber_optAll.OptionValue
1080            arr_varFld(F_ACTNO, 0) = Null
1090            arr_varFld(F_DATE, 0) = datEndDate
1100            arr_varFld(F_FOREX, 0) = CBool(True)
1110            arr_varFld(F_ICSHB, 0) = CDbl(0)
1120            arr_varFld(F_PCSHB, 0) = CDbl(0)
1130            arr_varFld(F_COSTB, 0) = CDbl(0)
1140            arr_varFld(F_ICSHE, 0) = CDbl(0)
1150            arr_varFld(F_PCSHE, 0) = CDbl(0)
1160            arr_varFld(F_COSTE, 0) = CDbl(0)
1170            arr_varFld(F_RECS, 0) = CLng(0)
1180          End Select
1190        End If

            ' ** This frist group of statements doesn't deal with the Detail fields.
1200        Select Case frm.chkTransactions
            Case True
              ' ** Transactions only, don't show the Balance data; assumes separate icash, pcash.
1210          .BeginningBalance_lbl.Visible = False
1220          .PreviousIcash.Visible = False
1230          .PreviousPcash.Visible = False
1240          .PreviousCash.Visible = False
1250          .PreviousCost.Visible = False
1260          .TotalICash_lbl.Visible = False
1270          .TotalICash.Visible = False
1280          .TotalPCash.Visible = False
1290          .TotalCash.Visible = False
1300          .TotalCost.Visible = False
1310          .TotalICash_line.Visible = False
1320          .TotalPCash_line.Visible = False
1330          .TotalCash_line.Visible = False
1340          .TotalCost_line.Visible = False
1350          .EndingBalance_lbl.Visible = False
1360          .EndingICash.Visible = False
1370          .EndingPCash.Visible = False
1380          .EndingCash.Visible = False
1390          .EndingCost.Visible = False
1400          .EndingICash_line.Visible = False
1410          .EndingPCash_line.Visible = False
1420          .EndingCash_line.Visible = False
1430          .EndingCost_line.Visible = False
1440          .Section("GroupHeader2").Visible = False
1450          .Section("GroupFooter1").Visible = False
1460        Case False
              ' ** In combination with Statements, do show the Balance data; assumes separate icash, pcash.
1470          .BeginningBalance_lbl.Visible = True
1480          .PreviousIcash.Visible = True
1490          .PreviousPcash.Visible = True
1500          .PreviousCash.Visible = False     'CASH
1510          .PreviousCost.Visible = True
1520          .TotalICash_lbl.Visible = True
1530          .TotalICash.Visible = True
1540          .TotalPCash.Visible = True
1550          .TotalCash.Visible = False        'CASH
1560          .TotalCost.Visible = True
1570          .TotalICash_line.Visible = True
1580          .TotalPCash_line.Visible = True
1590          .TotalCash_line.Visible = False   'CASH
1600          .TotalCost_line.Visible = True
1610          .EndingBalance_lbl.Visible = True
1620          .EndingICash.Visible = True
1630          .EndingPCash.Visible = True
1640          .EndingCash.Visible = False       'CASH
1650          .EndingCost.Visible = True
1660          .EndingICash_line.Visible = True
1670          .EndingPCash_line.Visible = True
1680          .EndingCash_line.Visible = False  'CASH
1690          .EndingCost_line.Visible = True
1700        End Select  ' ** chkTransactions.

            ' ** This group does deal with the Detail fields.
1710        Select Case frm.chkCombineCash
            Case True
1720          Select Case frm.chkTransactions
              Case True
                ' ** Transactions only, don't show the Balance data.
1730            .BeginningBalance_lbl.Visible = False
1740            .PreviousIcash.Visible = False
1750            .PreviousPcash.Visible = False
1760            .PreviousCash.Visible = False
1770            .PreviousCost.Visible = False
1780            .TotalICash_lbl.Visible = False
1790            .TotalICash.Visible = False
1800            .TotalPCash.Visible = False
1810            .TotalCash.Visible = False
1820            .TotalCost.Visible = False
1830            .TotalICash_line.Visible = False
1840            .TotalPCash_line.Visible = False
1850            .TotalCash_line.Visible = False
1860            .TotalCost_line.Visible = False
1870            .EndingBalance_lbl.Visible = False
1880            .EndingICash.Visible = False
1890            .EndingPCash.Visible = False
1900            .EndingCash.Visible = False
1910            .EndingCost.Visible = False
1920            .EndingICash_line.Visible = False
1930            .EndingPCash_line.Visible = False
1940            .EndingCash_line.Visible = False
1950            .EndingCost_line.Visible = False
1960          Case False
                ' ** In combination with Statements, do show the Balance data.
1970            .BeginningBalance_lbl.Visible = True
1980            .PreviousIcash.Visible = False
1990            .PreviousPcash.Visible = False
2000            .PreviousCash.Visible = True
2010            .PreviousCost.Visible = True
2020            .TotalICash_lbl.Visible = True
2030            .TotalICash.Visible = False
2040            .TotalPCash.Visible = False
2050            .TotalCash.Visible = True
2060            .TotalCost.Visible = True
2070            .TotalICash_line.Visible = False
2080            .TotalPCash_line.Visible = False
2090            .TotalCash_line.Visible = True
2100            .TotalCost_line.Visible = True
2110            .EndingBalance_lbl.Visible = True
2120            .EndingICash.Visible = False
2130            .EndingPCash.Visible = False
2140            .EndingCash.Visible = True
2150            .EndingCost.Visible = True
2160            .EndingICash_line.Visible = False
2170            .EndingPCash_line.Visible = False
2180            .EndingCash_line.Visible = True
2190            .EndingCost_line.Visible = True
2200          End Select  ' ** chkTransactions.
2210          .ICash.Visible = False
2220          .icash_lbl.Visible = False
2230          .icash_lbl_line.Visible = False
2240          .icash_usd.Visible = False
2250          .icash_usd_lbl.Visible = False
2260          .icash_usd_lbl_line.Visible = False
2270          .PCash.Visible = False
2280          .pcash_lbl.Visible = False
2290          .pcash_lbl_line.Visible = False
2300          .pcash_usd.Visible = False
2310          .pcash_usd_lbl.Visible = False
2320          .pcash_usd_lbl_line.Visible = False
2330          .cash.Visible = True
2340          .cash_lbl.Visible = True
2350          .cash_lbl_line.Visible = True
2360          .cash_usd.Visible = True
2370          .cash_usd_lbl.Visible = True
2380          .cash_usd_lbl_line.Visible = True
2390          .Cost.Visible = True
2400          .cost_lbl.Visible = True
2410          .cost_lbl_line.Visible = True
2420          .cost_usd.Visible = True
2430          .cost_usd_lbl.Visible = True
2440          .cost_usd_lbl_line.Visible = True
2450        Case False
2460          .ICash.Visible = True
2470          .icash_lbl.Visible = True
2480          .icash_lbl_line.Visible = True
2490          .icash_usd.Visible = True
2500          .icash_usd_lbl.Visible = True
2510          .icash_usd_lbl_line.Visible = True
2520          .PCash.Visible = True
2530          .pcash_lbl.Visible = True
2540          .pcash_lbl_line.Visible = True
2550          .pcash_usd.Visible = True
2560          .pcash_usd_lbl.Visible = True
2570          .pcash_usd_lbl_line.Visible = True
2580          .cash.Visible = False
2590          .cash_lbl.Visible = False
2600          .cash_lbl_line.Visible = False
2610          .cash_usd.Visible = False
2620          .cash_usd_lbl.Visible = False
2630          .cash_usd_lbl_line.Visible = False
2640          .Cost.Visible = True
2650          .cost_lbl.Visible = True
2660          .cost_lbl_line.Visible = True
2670          .cost_usd.Visible = True
2680          .cost_usd_lbl.Visible = True
2690          .cost_usd_lbl_line.Visible = True
2700        End Select  ' ** chkCombineCash.

2710        If glngPrintRpts > 0& Then
2720          For lngZ = 0& To (glngPrintRpts - 1&)
2730            If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
2740              If garr_varPrintRpt(PR_TRANS, lngZ) = True Then
2750                blnSkip = True
2760              Else
2770                garr_varPrintRpt(PR_TRANS, lngZ) = CBool(True)
2780              End If
2790              Exit For
2800            End If
2810          Next
2820        End If

2830      Else
2840        Cancel = -1
2850      End If
2860    End With

2870    If Cancel = -1 Then
2880      Beep
2890      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
2900    ElseIf blnSkip = True Then
2910      Cancel = -1
2920    End If

EXITP:
2930    Set frm = Nothing
2940    Exit Sub

ERRH:
2950    If Forms(strCallingForm).Visible = False Then
2960      Forms(strCallingForm).Visible = True
2970    End If
2980    Select Case ERR.Number
        Case Else
2990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3000    End Select
3010    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim blnFound As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long

3110    strThisProc = THIS_PROC

3120    With Me

3130      Select Case Forms(strCallingForm).chkLegalName
          Case True
3140        varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
3150        If IsNull(varTmp00) = True Then
3160          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
3170        End If
3180        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
3190        .Name_lbl.Caption = varTmp00
3200      Case False
3210  On Error Resume Next
3220        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
3230        If ERR.Number <> 0 Then
3240  On Error GoTo ERRH
3250          varTmp00 = Null
3260        Else
3270  On Error GoTo ERRH
3280        End If
3290        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
3300        .Name_lbl.Caption = varTmp00
3310      End Select  ' ** chkLegalName.

3320      If blnSaveData = True And blnStatements = True Then
3330        Select Case IsNull(arr_varFld(F_ACTNO, 0))
            Case True
              ' ** First accountno.
3340          arr_varFld(F_ACTNO, 0) = .accountno
3350        Case False
3360          blnFound = False
3370          For lngX = 0& To (lngFlds - 1&)
3380            If arr_varFld(F_ACTNO, lngX) = .accountno Then
3390              blnFound = True
3400              Exit For
3410            End If
3420          Next
3430          If blnFound = False Then
3440            lngFlds = lngFlds + 1&
3450            lngE = lngFlds - 1&
3460            ReDim Preserve arr_varFld(F_ELEMS, lngE)
3470            arr_varFld(F_ACTNO, lngE) = .accountno
3480            arr_varFld(F_DATE, lngE) = datEndDate
3490            arr_varFld(F_FOREX, lngE) = CBool(True)
3500            arr_varFld(F_ICSHB, lngE) = CDbl(0)
3510            arr_varFld(F_PCSHB, lngE) = CDbl(0)
3520            arr_varFld(F_COSTB, lngE) = CDbl(0)
3530            arr_varFld(F_ICSHE, lngE) = CDbl(0)
3540            arr_varFld(F_PCSHE, lngE) = CDbl(0)
3550            arr_varFld(F_COSTE, lngE) = CDbl(0)
3560            arr_varFld(F_RECS, lngE) = CLng(0)
3570          End If
3580        End Select
3590      End If

3600    End With

EXITP:
3610    Exit Sub

ERRH:
3620    If Forms(strCallingForm).Visible = False Then
3630      Forms(strCallingForm).Visible = True
3640    End If
3650    Select Case ERR.Number
        Case Else
3660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3670    End Select
3680    Resume EXITP

End Sub

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader2_Format"

        Dim lngX As Long

3710    With Me
3720      If blnSaveData = True Then
3730        For lngX = 0& To (lngFlds - 1&)
3740          If arr_varFld(F_ACTNO, lngX) = .accountno Then
3750            arr_varFld(F_ICSHB, lngX) = .PreviousIcash
3760            arr_varFld(F_PCSHB, lngX) = .PreviousPcash
3770            arr_varFld(F_COSTB, lngX) = .PreviousCost
3780            Exit For
3790          End If
3800        Next
3810      End If
3820    End With

EXITP:
3830    Exit Sub

ERRH:
3840    If Forms(strCallingForm).Visible = False Then
3850      Forms(strCallingForm).Visible = True
3860    End If
3870    Select Case ERR.Number
        Case Else
3880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3890    End Select
3900    Resume EXITP

End Sub

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Format"

        Dim strFormat As String

4010    With Me
4020      Select Case .currsym_symbol
          Case "CHF"
4030        strFormat = "\C\H\F#,##0.00;(\C\H\F#,##0.00);\C\H\F0.00;\C\H\F0.00"
4040      Case Else
4050        strFormat = .currsym_symbol & "#,##0.00;(" & .currsym_symbol & "#,##0.00);" & .currsym_symbol & "0.00;" & .currsym_symbol & "0.00"
4060      End Select
4070      .ICash.Format = strFormat
4080      .PCash.Format = strFormat
4090      .cash.Format = strFormat
4100      .Cost.Format = strFormat
4110    End With

EXITP:
4120    Exit Sub

ERRH:
4130    If Forms(strCallingForm).Visible = False Then
4140      Forms(strCallingForm).Visible = True
4150    End If
4160    Select Case ERR.Number
        Case Else
4170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4180    End Select
4190    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

        Dim lngX As Long

4210    With Me
4220      If blnSaveData = True Then
4230        For lngX = 0& To (lngFlds - 1&)
4240          If arr_varFld(F_ACTNO, lngX) = .accountno Then
4250            arr_varFld(F_ICSHE, lngX) = .EndingICash
4260            arr_varFld(F_PCSHE, lngX) = .EndingPCash
4270            arr_varFld(F_COSTE, lngX) = .EndingCost
4280            Exit For
4290          End If
4300        Next
4310      End If
4320    End With

EXITP:
4330    Exit Sub

ERRH:
4340    If Forms(strCallingForm).Visible = False Then
4350      Forms(strCallingForm).Visible = True
4360    End If
4370    Select Case ERR.Number
        Case Else
4380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4390    End Select
4400    Resume EXITP

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Error"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim dblErrNum As Double, varErrLine As Variant
        Dim strErrDesc As String

4510    If DataErr <> 0 Then
4520      dblErrNum = DataErr
4530      varErrLine = Nz(Erl, 0)
4540      Select Case dblErrNum
          Case Else
4550        If IsNull(ERR.description) = False Then
4560          If Trim(ERR.description) <> vbNullString Then
4570            strErrDesc = ERR.description
4580          Else
4590            strErrDesc = zErrorDescription(DataErr)  ' ** Module Function: modErrorHandler.
4600          End If
4610        Else
4620          strErrDesc = zErrorDescription(DataErr)  ' ** Module Function: modErrorHandler.
4630        End If
4640        Set dbs = CurrentDb
4650        With dbs
4660          Set rst = .OpenRecordset("tblErrorLog", dbOpenDynaset, dbConsistent)
4670          zErrorWriteRecord dblErrNum, strErrDesc, THIS_NAME, strThisProc, CLng(varErrLine), rst  ' ** Module Function: modErrorHandler.
4680          rst.Close
4690          .Close
4700        End With
4710        Response = acDataErrContinue
4720      End Select
4730    End If

EXITP:
4740    Set rst = Nothing
4750    Set dbs = Nothing
4760    Exit Sub

ERRH:
4770    If Forms(strCallingForm).Visible = False Then
4780      Forms(strCallingForm).Visible = True
4790    End If
4800    Select Case ERR.Number
        Case Else
4810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4820    End Select
4830    Resume EXITP

End Sub

Private Sub Report_Close()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnAdd As Boolean, blnAddAll As Boolean
        Dim lngX As Long

4910    With Me

4920      strThisProc = THIS_PROC

4930      If blnSaveData = True Then
4940        blnAdd = False: blnAddAll = False
4950        Set dbs = CurrentDb
4960        With dbs
4970          Set rst = .OpenRecordset("zz_tbl_Statement_Transaction", dbOpenDynaset, dbConsistent)
4980          With rst
4990            If .BOF = True And .EOF = True Then
5000              blnAddAll = True
5010            End If
5020            For lngX = 0& To (lngFlds - 1&)
5030              If IsNull(arr_varFld(F_ACTNO, lngX)) = False Then
5040                If Trim(arr_varFld(F_ACTNO, lngX)) <> vbNullString Then
5050                  blnAdd = False
5060                  Select Case blnAddAll
                      Case True
5070                    blnAdd = True
5080                  Case False
5090                    .FindFirst "[accountno] = '" & arr_varFld(F_ACTNO, lngX) & "' And " & _
                          "[st_period] = #" & Format(arr_varFld(F_DATE, lngX), "mm/dd/yyyy") & "# And " & _
                          "[st_forex] = True"
5100                    If .NoMatch = True Then
5110                      blnAdd = True
5120                    End If
5130                  End Select
5140                  Select Case blnAdd
                      Case True
5150                    .AddNew
5160                    ![accountno] = arr_varFld(F_ACTNO, lngX)
5170                    ![st_period] = arr_varFld(F_DATE, lngX)
5180                    ![st_forex] = arr_varFld(F_FOREX, lngX)
5190                  Case False
5200                    .Edit
5210                  End Select
5220                  ![st_icash_beg] = arr_varFld(F_ICSHB, lngX)
5230                  ![st_pcash_beg] = arr_varFld(F_PCSHB, lngX)
5240                  ![st_cost_beg] = arr_varFld(F_COSTB, lngX)
5250                  ![st_icash_end] = arr_varFld(F_ICSHE, lngX)
5260                  ![st_pcash_end] = arr_varFld(F_PCSHE, lngX)
5270                  ![st_cost_end] = arr_varFld(F_COSTE, lngX)
5280                  ![st_recs] = arr_varFld(F_RECS, lngX)
5290                  ![st_datemodified] = Now()
5300                  .Update
5310                End If
5320              End If
5330            Next
5340            .Close
5350          End With
5360          .Close
5370        End With
5380      End If

5390      If strCallingForm <> vbNullString Then
5400        gblnReportClose = True
5410        If Forms(strCallingForm).Visible = False Then
5420          DoCmd.Restore
5430          Forms(strCallingForm).Visible = True
5440        End If
5450        Forms(strCallingForm).TimerInterval = 250&
5460      End If

5470    End With

EXITP:
5480    Set rst = Nothing
5490    Set dbs = Nothing
5500    Exit Sub

ERRH:
5510    If Forms(strCallingForm).Visible = False Then
5520      Forms(strCallingForm).Visible = True
5530    End If
5540    Select Case ERR.Number
        Case Else
5550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5560    End Select
5570    Resume EXITP

End Sub
