VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptTransaction_Statement_SortType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "rptTransaction_Statement_SortType"

'VGC 11/27/2016: CHANGES!

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

220         Select Case frm.chkPageOf
            Case True
230           .PageNumOf.Visible = True
240           .PageNum.Visible = False
250         Case False
260           .PageNum.Visible = True
270           .PageNumOf.Visible = False
280           .PageNumOf.ControlSource = vbNullString
290         End Select

300         .Name_lbl.Caption = vbNullString

310         Select Case frm.opgAccountNumber
            Case frm.opgAccountNumber_optSpecified.OptionValue
320           .Section("GroupFooter1").ForceNewPage = acForceNewPageNone
330         Case frm.opgAccountNumber_optAll.OptionValue
340           .Section("GroupFooter1").ForceNewPage = acForceNewPageAfter
350         End Select

360         Select Case frm.chkTransactions
            Case True
370           blnStatements = False
380           datStartDate = CDate(frm.TransDateStart)
390           datEndDate = CDate(frm.TransDateEnd)
400           .AsOf_lbl.Caption = "From " & Format(datStartDate, "mm/dd/yyyy") & " To " & Format(datEndDate, "mm/dd/yyyy")
410           .AsOf_lbl2.Caption = "From " & Format(datStartDate, "mm/dd/yyyy") & " To " & Format(datEndDate, "mm/dd/yyyy")
420           Select Case frm.chkArchiveOnly_Trans
              Case True
430             .AsOf.Visible = True
440             .AsOf2.Visible = True
450           Case False
460             .AsOf.Visible = False
470             .AsOf2.Visible = False
480           End Select  ' ** chkArchiveOnly_Trans.
490         Case False
500           blnStatements = True
510           datStartDate = 0
520           datEndDate = CDate(frm.DateEnd)
530           .AsOf_lbl.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
540           .AsOf_lbl2.Caption = "For Period Ending " & Format(datEndDate, "mmmm dd, yyyy")
550           .AsOf.Visible = False
560           .AsOf2.Visible = False
570         End Select  ' ** chkTransactions

580         .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.
590         .CoInfoBlock2.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

600         varTmp00 = frm.cmbAccounts
610         If IsNull(varTmp00) = False Then
620           varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
630           varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
640           If CLng(datStartDate) = 0 Then
650             .Caption = "Transaction_Statement_" & varTmp00 & "_" & Format(datEndDate, "yymmdd")
660           Else
670             .Caption = "Transaction_Statement_" & varTmp00 & "_" & Format(datStartDate, "yymmdd") & "_To_" & Format(datEndDate, "yymmdd")
680           End If
690           strAccountNo = varTmp00
700         End If

710         If frm.chkDevSaveData.Visible = True And frm.chkDevSaveData = True Then
720           blnSaveData = True
730           lngFlds = 1&
740           ReDim arr_varFld(F_ELEMS, 0)
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
750           Select Case frm.opgAccountNumber
              Case frm.opgAccountNumber_optSpecified.OptionValue
760             arr_varFld(F_ACTNO, 0) = strAccountNo
770             arr_varFld(F_DATE, 0) = datEndDate
780             arr_varFld(F_FOREX, 0) = CBool(False)
790             arr_varFld(F_ICSHB, 0) = CDbl(0)
800             arr_varFld(F_PCSHB, 0) = CDbl(0)
810             arr_varFld(F_COSTB, 0) = CDbl(0)
820             arr_varFld(F_ICSHE, 0) = CDbl(0)
830             arr_varFld(F_PCSHE, 0) = CDbl(0)
840             arr_varFld(F_COSTE, 0) = CDbl(0)
850             arr_varFld(F_RECS, 0) = CLng(0)
860           Case frm.opgAccountNumber_optAll.OptionValue
870             arr_varFld(F_ACTNO, 0) = Null
880             arr_varFld(F_DATE, 0) = datEndDate
890             arr_varFld(F_FOREX, 0) = CBool(False)
900             arr_varFld(F_ICSHB, 0) = CDbl(0)
910             arr_varFld(F_PCSHB, 0) = CDbl(0)
920             arr_varFld(F_COSTB, 0) = CDbl(0)
930             arr_varFld(F_ICSHE, 0) = CDbl(0)
940             arr_varFld(F_PCSHE, 0) = CDbl(0)
950             arr_varFld(F_COSTE, 0) = CDbl(0)
960             arr_varFld(F_RECS, 0) = CLng(0)
970           End Select
980         End If

            ' ** This frist group of statements don't deal with the Detail fields; assumes separate icash, pcash.
990         Select Case frm.chkTransactions
            Case True
              ' ** Transactions only, don't show the Balance data.
1000          .BeginningBalance_lbl.Visible = False
1010          .PreviousIcash.Visible = False
1020          .PreviousPcash.Visible = False
1030          .PreviousCash.Visible = False
1040          .PreviousCost.Visible = False
1050          .TotalICash_lbl.Visible = False
1060          .TotalICash.Visible = False
1070          .TotalPCash.Visible = False
1080          .TotalCash.Visible = False
1090          .TotalCost.Visible = False
1100          .TotalICash_line.Visible = False
1110          .TotalPCash_line.Visible = False
1120          .TotalCash_line.Visible = False
1130          .TotalCost_line.Visible = False
1140          .EndingBalance_lbl.Visible = False
1150          .EndingICash.Visible = False
1160          .EndingPCash.Visible = False
1170          .EndingCash.Visible = False
1180          .EndingCost.Visible = False
1190          .EndingICash_line.Visible = False
1200          .EndingPCash_line.Visible = False
1210          .EndingCash_line.Visible = False
1220          .EndingCost_line.Visible = False
1230        Case False
              ' ** In combination with Statements, do show the Balance data; assumes separate icash, pcash.
1240          .BeginningBalance_lbl.Visible = True
1250          .PreviousIcash.Visible = True
1260          .PreviousPcash.Visible = True
1270          .PreviousCash.Visible = False
1280          .PreviousCost.Visible = True
1290          .TotalICash_lbl.Visible = True
1300          .TotalICash.Visible = True
1310          .TotalPCash.Visible = True
1320          .TotalCash.Visible = False
1330          .TotalCost.Visible = True
1340          .TotalICash_line.Visible = True
1350          .TotalPCash_line.Visible = True
1360          .TotalCash_line.Visible = False
1370          .TotalCost_line.Visible = True
1380          .EndingBalance_lbl.Visible = True
1390          .EndingICash.Visible = True
1400          .EndingPCash.Visible = True
1410          .EndingCash.Visible = False       'CASH
1420          .EndingCost.Visible = True
1430          .EndingICash_line.Visible = True
1440          .EndingPCash_line.Visible = True
1450          .EndingCash_line.Visible = False  'CASH
1460          .EndingCost_line.Visible = True
1470        End Select  ' ** chkTransactions.

            ' ** This group does deal with the Detail fields.
1480        Select Case frm.chkCombineCash
            Case True
1490          Select Case frm.chkTransactions
              Case True
1500            .BeginningBalance_lbl.Visible = False
1510            .PreviousIcash.Visible = False
1520            .PreviousPcash.Visible = False
1530            .PreviousCash.Visible = False
1540            .PreviousCost.Visible = False
1550            .TotalICash_lbl.Visible = False
1560            .TotalICash.Visible = False
1570            .TotalPCash.Visible = False
1580            .TotalCash.Visible = False
1590            .TotalCost.Visible = False
1600            .TotalICash_line.Visible = False
1610            .TotalPCash_line.Visible = False
1620            .TotalCash_line.Visible = False
1630            .TotalCost_line.Visible = False
1640            .EndingBalance_lbl.Visible = False
1650            .EndingICash.Visible = False
1660            .EndingPCash.Visible = False
1670            .EndingCash.Visible = False
1680            .EndingCost.Visible = False
1690            .EndingICash_line.Visible = False
1700            .EndingPCash_line.Visible = False
1710            .EndingCash_line.Visible = False
1720            .EndingCost_line.Visible = False
1730          Case False
1740            .BeginningBalance_lbl.Visible = True
1750            .PreviousIcash.Visible = False
1760            .PreviousPcash.Visible = False
1770            .PreviousCash.Visible = True
1780            .PreviousCost.Visible = True
1790            .TotalICash_lbl.Visible = True
1800            .TotalICash.Visible = False
1810            .TotalPCash.Visible = False
1820            .TotalCash.Visible = True
1830            .TotalCost.Visible = True
1840            .TotalICash_line.Visible = False
1850            .TotalPCash_line.Visible = False
1860            .TotalCash_line.Visible = True
1870            .TotalCost_line.Visible = True
1880            .EndingBalance_lbl.Visible = True
1890            .EndingICash.Visible = False
1900            .EndingPCash.Visible = False
1910            .EndingCash.Visible = True
1920            .EndingCost.Visible = True
1930            .EndingICash_line.Visible = False
1940            .EndingPCash_line.Visible = False
1950            .EndingCash_line.Visible = True
1960            .EndingCost_line.Visible = True
1970          End Select  ' ** chkTransactions.
1980          .ICash.Visible = False
1990          .icash_lbl.Visible = False
2000          .icash_lbl_line.Visible = False
2010          .PCash.Visible = False
2020          .pcash_lbl.Visible = False
2030          .pcash_lbl_line.Visible = False
2040          .cash.Visible = True
2050          .cash_lbl.Visible = True
2060          .cash_lbl_line.Visible = True
2070          .Cost.Visible = True
2080          .cost_lbl.Visible = True
2090          .cost_lbl_line.Visible = True
2100        Case False
2110          .ICash.Visible = True
2120          .icash_lbl.Visible = True
2130          .icash_lbl_line.Visible = True
2140          .PCash.Visible = True
2150          .pcash_lbl.Visible = True
2160          .pcash_lbl_line.Visible = True
2170          .cash.Visible = False
2180          .cash_lbl.Visible = False
2190          .cash_lbl_line.Visible = False
2200          .Cost.Visible = True
2210          .cost_lbl.Visible = True
2220          .cost_lbl_line.Visible = True
2230        End Select  ' ** chkCombineCash.

2240        If glngPrintRpts > 0& Then
2250          For lngZ = 0& To (glngPrintRpts - 1&)
2260            If garr_varPrintRpt(PR_ACTNO, lngZ) = gstrAccountNo Then
2270              If garr_varPrintRpt(PR_TRANS, lngZ) = True Then
2280                blnSkip = True
2290              Else
2300                garr_varPrintRpt(PR_TRANS, lngZ) = CBool(True)
2310              End If
2320              Exit For
2330            End If
2340          Next
2350        End If

2360      Else
2370        Cancel = -1
2380      End If
2390    End With

2400    If Cancel = -1 Then
2410      Beep
2420      MsgBox "This report is not to be opened independently.", vbCritical + vbOKOnly, "Report Inaccessible"
2430    ElseIf blnSkip = True Then
2440      Cancel = -1
2450    End If

EXITP:
2460    Set frm = Nothing
2470    Exit Sub

ERRH:
2480    If Forms(strCallingForm).Visible = False Then
2490      Forms(strCallingForm).Visible = True
2500    End If
2510    Select Case ERR.Number
        Case Else
2520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2530    End Select
2540    Resume EXITP

End Sub

Private Sub GroupHeader1_Format(Cancel As Integer, FormatCount As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader1_Format"

        Dim blnFound As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long

2610    strThisProc = THIS_PROC

2620    With Me

2630      Select Case Forms(strCallingForm).chkLegalName
          Case True
2640        varTmp00 = FixAmps(Nz(.legalname, vbNullString))  ' ** Module Functions: modStringFuncs.
2650        If IsNull(varTmp00) = True Then
2660          varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
2670        End If
2680        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
2690        .Name_lbl.Caption = varTmp00
2700      Case False
2710        varTmp00 = FixAmps(Nz(.shortname, vbNullString))  ' ** Module Functions: modStringFuncs.
2720        If IsNull(varTmp00) = True Then varTmp00 = "{no name}"
2730        .Name_lbl.Caption = varTmp00
2740      End Select  ' ** chkLegalName.

2750      If blnSaveData = True And blnStatements = True Then
2760        Select Case IsNull(arr_varFld(F_ACTNO, 0))
            Case True
              ' ** First accountno.
2770          arr_varFld(F_ACTNO, 0) = .accountno
2780        Case False
2790          blnFound = False
2800          For lngX = 0& To (lngFlds - 1&)
2810            If arr_varFld(F_ACTNO, lngX) = .accountno Then
2820              blnFound = True
2830              Exit For
2840            End If
2850          Next
2860          If blnFound = False Then
2870            lngFlds = lngFlds + 1&
2880            lngE = lngFlds - 1&
2890            ReDim Preserve arr_varFld(F_ELEMS, lngE)
2900            arr_varFld(F_ACTNO, lngE) = .accountno
2910            arr_varFld(F_DATE, lngE) = datEndDate
2920            arr_varFld(F_FOREX, lngE) = CBool(False)
2930            arr_varFld(F_ICSHB, lngE) = CDbl(0)
2940            arr_varFld(F_PCSHB, lngE) = CDbl(0)
2950            arr_varFld(F_COSTB, lngE) = CDbl(0)
2960            arr_varFld(F_ICSHE, lngE) = CDbl(0)
2970            arr_varFld(F_PCSHE, lngE) = CDbl(0)
2980            arr_varFld(F_COSTE, lngE) = CDbl(0)
2990            arr_varFld(F_RECS, lngE) = CLng(0)
3000          End If
3010        End Select
3020      End If

3030    End With

EXITP:
3040    Exit Sub

ERRH:
3050    If Forms(strCallingForm).Visible = False Then
3060      Forms(strCallingForm).Visible = True
3070    End If
3080    Select Case ERR.Number
        Case Else
3090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3100    End Select
3110    Resume EXITP

End Sub

Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupHeader2_Format"

        Dim lngX As Long

3210    With Me
3220      If blnSaveData = True Then
3230        For lngX = 0& To (lngFlds - 1&)
3240          If arr_varFld(F_ACTNO, lngX) = .accountno Then
3250            arr_varFld(F_ICSHB, lngX) = .PreviousIcash
3260            arr_varFld(F_PCSHB, lngX) = .PreviousPcash
3270            arr_varFld(F_COSTB, lngX) = .PreviousCost
3280            Exit For
3290          End If
3300        Next
3310      End If
3320    End With

EXITP:
3330    Exit Sub

ERRH:
3340    If Forms(strCallingForm).Visible = False Then
3350      Forms(strCallingForm).Visible = True
3360    End If
3370    Select Case ERR.Number
        Case Else
3380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3390    End Select
3400    Resume EXITP

End Sub

Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "GroupFooter1_Format"

        Dim lngX As Long

3510    With Me
3520      If blnSaveData = True Then
3530        For lngX = 0& To (lngFlds - 1&)
3540          If arr_varFld(F_ACTNO, lngX) = .accountno Then
3550            arr_varFld(F_ICSHE, lngX) = .EndingICash
3560            arr_varFld(F_PCSHE, lngX) = .EndingPCash
3570            arr_varFld(F_COSTE, lngX) = .EndingCost
3580            Exit For
3590          End If
3600        Next
3610      End If
3620    End With

EXITP:
3630    Exit Sub

ERRH:
3640    If Forms(strCallingForm).Visible = False Then
3650      Forms(strCallingForm).Visible = True
3660    End If
3670    Select Case ERR.Number
        Case Else
3680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3690    End Select
3700    Resume EXITP

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Error"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim dblErrNum As Double, varErrLine As Variant
        Dim strErrDesc As String

3810    If DataErr <> 0 Then
3820      dblErrNum = DataErr
3830      varErrLine = Nz(Erl, 0)
3840      Select Case dblErrNum
          Case Else
3850        If IsNull(ERR.description) = False Then
3860          If Trim(ERR.description) <> vbNullString Then
3870            strErrDesc = ERR.description
3880          Else
3890            strErrDesc = zErrorDescription(DataErr)  ' ** Module Function: modErrorHandler.
3900          End If
3910        Else
3920          strErrDesc = zErrorDescription(DataErr)  ' ** Module Function: modErrorHandler.
3930        End If
3940        Set dbs = CurrentDb
3950        With dbs
3960          Set rst = .OpenRecordset("tblErrorLog", dbOpenDynaset, dbConsistent)
3970          zErrorWriteRecord dblErrNum, strErrDesc, THIS_NAME, strThisProc, CLng(varErrLine), rst  ' ** Module Function: modErrorHandler.
3980          rst.Close
3990          .Close
4000        End With
4010        Response = acDataErrContinue
4020      End Select
4030    End If

EXITP:
4040    Set rst = Nothing
4050    Set dbs = Nothing
4060    Exit Sub

ERRH:
4070    If Forms(strCallingForm).Visible = False Then
4080      Forms(strCallingForm).Visible = True
4090    End If
4100    Select Case ERR.Number
        Case Else
4110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4120    End Select
4130    Resume EXITP

End Sub

Private Sub Report_Close()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Report_Close"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnAdd As Boolean, blnAddAll As Boolean
        Dim lngX As Long

4210    With Me

4220      strThisProc = THIS_PROC

4230      If blnSaveData = True Then
4240        blnAdd = False: blnAddAll = False
4250        Set dbs = CurrentDb
4260        With dbs
4270          Set rst = .OpenRecordset("zz_tbl_Statement_Transaction", dbOpenDynaset, dbConsistent)
4280          With rst
4290            If .BOF = True And .EOF = True Then
4300              blnAddAll = True
4310            End If
4320            For lngX = 0& To (lngFlds - 1&)
4330              If IsNull(arr_varFld(F_ACTNO, lngX)) = False Then
4340                If Trim(arr_varFld(F_ACTNO, lngX)) <> vbNullString Then
4350                  blnAdd = False
4360                  Select Case blnAddAll
                      Case True
4370                    blnAdd = True
4380                  Case False
4390                    .FindFirst "[accountno] = '" & arr_varFld(F_ACTNO, lngX) & "' And " & _
                          "[st_period] = #" & Format(arr_varFld(F_DATE, lngX), "mm/dd/yyyy") & "# And " & _
                          "[st_forex] = False"
4400                    If .NoMatch = True Then
4410                      blnAdd = True
4420                    End If
4430                  End Select
4440                  Select Case blnAdd
                      Case True
4450                    .AddNew
4460                    ![accountno] = arr_varFld(F_ACTNO, lngX)
4470                    ![st_period] = arr_varFld(F_DATE, lngX)
4480                    ![st_forex] = arr_varFld(F_FOREX, lngX)
4490                  Case False
4500                    .Edit
4510                  End Select
4520                  ![st_icash_beg] = arr_varFld(F_ICSHB, lngX)
4530                  ![st_pcash_beg] = arr_varFld(F_PCSHB, lngX)
4540                  ![st_cost_beg] = arr_varFld(F_COSTB, lngX)
4550                  ![st_icash_end] = arr_varFld(F_ICSHE, lngX)
4560                  ![st_pcash_end] = arr_varFld(F_PCSHE, lngX)
4570                  ![st_cost_end] = arr_varFld(F_COSTE, lngX)
4580                  ![st_recs] = arr_varFld(F_RECS, lngX)
4590                  ![st_datemodified] = Now()
4600                  .Update
4610                End If
4620              End If
4630            Next
4640            .Close
4650          End With
4660          .Close
4670        End With
4680      End If

4690      If strCallingForm <> vbNullString Then
4700        gblnReportClose = True
4710        If Forms(strCallingForm).Visible = False Then
4720          DoCmd.Restore
4730          Forms(strCallingForm).Visible = True
4740        End If
4750        Forms(strCallingForm).TimerInterval = 250&
4760      End If

4770    End With

EXITP:
4780    Set rst = Nothing
4790    Set dbs = Nothing
4800    Exit Sub

ERRH:
4810    If Forms(strCallingForm).Visible = False Then
4820      Forms(strCallingForm).Visible = True
4830    End If
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub
