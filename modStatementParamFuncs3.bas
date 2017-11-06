Attribute VB_Name = "modStatementParamFuncs3"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modStatementParamFuncs3"

'VGC 09/05/2017: CHANGES!

' ** cmbAccounts combo box constants:
Private Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
'Private Const CBX_A_DESC   As Integer = 1  ' ** Desc
'Private Const CBX_A_PREDAT As Integer = 2  ' ** predate
'Private Const CBX_A_SHORT  As Integer = 3  ' ** shortname
'Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname
'Private Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
'Private Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
'Private Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
'Private Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

' ** cmbMonth combo box constants:
'Private Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
Private Const CBX_MON_NAME  As Integer = 1  ' ** month_name
'Private Const CBX_MON_SHORT As Integer = 2  ' ** month_short
' **

Public Function Test_AList_SP(strAccountNo As String) As Boolean
' ** SetQrys_AList_SP() started throwing a 'Bad DLL calling convention'
' ** error, and I have no idea why.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Test_AList_SP"

        Dim frm As Access.Form
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

110     blnRetVal = True

120     Set frm = Forms("frmStatementParameters")

130     varTmp00 = SetQrys_AList_SP(strAccountNo, frm)  ' ** Function: Below.
        ' ** Return codes:
        ' **    0  Success.
        ' **    1  Success, with Archive.
        ' **    2  Success, Archive only.
        ' **   -2  No data.
        ' **   -4  Date criteria not met.
        ' **   -9  Error.

140     If varTmp00 < 0 Then
150       blnRetVal = False
160     End If

EXITP:
170     Set frm = Nothing
180     Test_AList_SP = blnRetVal
190     Exit Function

ERRH:
200     blnRetVal = False
210     Select Case ERR.Number
        Case Else
220       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
230     End Select
240     Resume EXITP

End Function

Public Function SetQrys_AList_SP(strAccountNo As String, frm As Access.Form) As Integer
' ** This will now be for the Asset List exclusively!
' ** Return codes:
' **    0  Success.
' **    1  Success, with Archive.
' **    2  Success, Archive only.
' **   -2  No data.
' **   -4  Date criteria not met.
' **   -9  Error.

        'SetDateSpecificSQL(.cmbAccounts, "Statements", THIS_NAME)
        ' blnBuildAssetListInfo, for .opgAccountNumber_optSpecified
        'SetDateSpecificSQL("All", "Statements", THIS_NAME)
        ' blnBuildAssetListInfo, for .opgAccountNumber_optAll

300   On Error GoTo ERRH

        Const THIS_PROC As String = "SetQrys_AList_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strStartDate As String, strEndDate As String
        Dim blnNoAccountsUpdated As Boolean
        Dim lngRecs As Long
        Dim varTmp00 As Variant
        Dim intRetVal As Integer, intRetVal_ChkFirstBal As Integer
        Dim lngX As Long

310     intRetVal = 0
320     blnNoAccountsUpdated = True

        'I DON'T REALLY KNOW WHAT THIS IS DOING!
        'I DON'T THINK IT'S SETTING ANY DATES FOR THE ASSET LIST!
330     strEndDate = frm.DateEnd
        'strStartDate = frm.DateStart
        'I THINK THEY SHOULD ALL USE THIS DLOOKUP()!!!!
        'If ERR <> 0 Then
340     If strAccountNo <> "All" Then
350       varTmp00 = DLookup("[BalDate_max]", "qryStatementParameters_AssetList_03", "[accountno] = '" & strAccountNo & "'")
360       Select Case IsNull(varTmp00)
          Case True
            ' ** Get date from Statement Date table?
370         varTmp00 = DLookup("[Statement_Date]", "Statement Date")
380         Select Case IsNull(varTmp00)
            Case True
390           frm.DateStart = #1/1/1900#
400           strStartDate = "01/01/1900"
410         Case False
420           frm.DateStart = CDate(varTmp00)
430           strStartDate = Format(varTmp00, "mm/dd/yyyy")
440         End Select
450       Case False
460         frm.DateStart = CDate(varTmp00)
470         strStartDate = Format(varTmp00, "mm/dd/yyyy")
480       End Select
490       DoEvents
500     End If
        'If IsNull(DLookup("[BalDate_max]", "qryStatementParameters_AssetList_03", "[accountno] = '" & strAccountNo & "'")) = True Then
        '  strStartDate = "1/1/1900"
        'Else
        '  strStartDate = DLookup("[BalDate_max]", "qryStatementParameters_AssetList_03", "[accountno] = '" & strAccountNo & "'")
        'End If

510     Set dbs = CurrentDb

520     If strAccountNo = "All" Then
          ' ** Loop through each account and set the right Balance information.
          ' ** The period ending date might be valid for some of the accounts,
          ' ** but invalid for others. So here we only change those that are
          ' ** invalid so that they have a valid balance data.
          ' ** However, the accounts that have no transactions prior to the
          ' ** chosen period ending date will be passed by.

          ' ** Account, just accountno.
530       Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_02")
540       Set rst = qdf.OpenRecordset
550       With rst
560         .MoveLast
570         lngRecs = .RecordCount
580         .MoveFirst
590         For lngX = 1& To lngRecs
600           intRetVal_ChkFirstBal = ChkFirstBal(dbs, ![accountno], strEndDate) ' ** Module Function: modStatementParamFuncs2.
              ' ** Return codes:
              ' **    0  Success.
              ' **    1  Success, with Archive.
              ' **    2  Success, Archive only.
              ' **   -2  No data.
              ' **   -4  Date criteria not met.
              ' **   -9  Error.
610           If intRetVal_ChkFirstBal >= 0 Then
                ' ** If it gets at least one good hit, the variable's turned off.
620             blnNoAccountsUpdated = False
630           End If
640           If lngX < lngRecs Then .MoveNext
650         Next
660         .Close
670       End With

          ' ** For all accounts, the return value only represents 1 of many, and the Boolean (next line), represents the entire function.
680       If blnNoAccountsUpdated = True Then
690         intRetVal = -4
            ' ** None of the accounts were updated due to the fact that
            ' ** none of the accounts had transactions that were prior
            ' ** to the selected Period Ending date.
700       End If

710     Else
          ' ** Specific account was chosen.

720       intRetVal_ChkFirstBal = ChkFirstBal(dbs, strAccountNo, strEndDate) ' ** Module Function: modStatementParamFuncs2.
          ' ** Return codes:
          ' **    0  Success.
          ' **    1  Success, with Archive.
          ' **    2  Success, Archive only.
          ' **   -2  No data.
          ' **   -4  Date criteria not met.
          ' **       The account that was selected was not updated due to the fact that it
          ' **       had no transactions that were prior to the selected Period Ending date.
          ' **   -9  Error.
          ' ** For 1 account, this return value represents the entire function.
730       intRetVal = intRetVal_ChkFirstBal

740     End If

750     dbs.Close

        ' ** OLD QUERIES:

        'strSQL = "SELECT Balance.accountno AS accountno, Max(Balance.[balance date]) AS [MaxOfbalance date] " & _
        '  "FROM Balance " & _
        '  "WHERE (((Balance.[balance date]) < #" & strEndDate & "#)) " & _
        '  "GROUP BY Balance.accountno;"
        'dbs.QueryDefs(strQry_BalDate).SQL = strSQL

        'dbs.QueryDefs("qryAssetList").SQL = strSQL
        ' ** qryCurrentTotalMarketValue comes after qryAssetList has been written.
        ' ** qryTransRangeTotals uses qryMaxBalDates!

        ' ** NEW QUERIES:

        ' ** Balance, grouped, with BalDate_max, by specified FormRef('EndDate').
        'qryStatementParameters_AssetList_03

        ' ** Account, linked to ActiveAssets, with add'l fields; all accounts.
        'qryStatementParameters_AssetList_06a

        ' ** Account, linked to ActiveAssets, with add'l fields; specified FormRef('accountno').
        'qryStatementParameters_AssetList_06b

EXITP:
760     Set rst = Nothing
770     Set qdf = Nothing
780     Set dbs = Nothing
790     SetQrys_AList_SP = intRetVal
800     Exit Function

ERRH:
810     intRetVal = -9
820     Select Case ERR.Number
        Case Else
830       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
840     End Select
850     Resume EXITP

End Function

Public Function FillAListTmp_SP(dbs As DAO.Database, rst1 As DAO.Recordset, strTableName As String) As Boolean

900   On Error GoTo ERRH

        Const THIS_PROC As String = "FillAListTmp_SP"

        Dim qdf As DAO.QueryDef, rst2 As DAO.Recordset, fld As DAO.Field
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

910     blnRetVal = True

920     With dbs

930       With rst1
940         If .BOF = True And .EOF = True Then
950           blnRetVal = False
960         Else
970           .MoveLast
980           lngRecs = .RecordCount
990           .MoveFirst

1000          Select Case strTableName
              Case "tmpAssetList2"
                ' ** Empty tmpAssetList2.
1010            Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_09c")
1020            qdf.Execute
1030            Set qdf = Nothing
1040          Case "tmpAssetList5"
                ' ** Empty tmpAssetList5.
1050            Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_70_52")
1060            qdf.Execute
1070            Set qdf = Nothing
1080          Case "tmpAccountInfo"
                ' ** Empty tmpAccountInfo.
1090            Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_09d")
1100            qdf.Execute
1110            Set qdf = Nothing
1120          Case "tmpAccountInfo2"
                ' ** Empty tmpAccountInfo2.
1130            Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_70_53")
1140            qdf.Execute
1150            Set qdf = Nothing
1160          End Select

1170          Set rst2 = dbs.OpenRecordset(strTableName, dbOpenDynaset, dbAppendOnly)
1180          For lngX = 1& To lngRecs
1190            rst2.AddNew
1200            For Each fld In rst2.Fields  ' ** Don't get confused about which is which, I'm just using the field Name.
1210              rst2.Fields(fld.Name) = .Fields(fld.Name)
1220            Next
1230            rst2.Update
1240            If lngX < lngRecs Then .MoveNext
1250          Next

1260        End If  ' ** BOF, EOF.
1270      End With  ' ** rst1.
1280    End With  ' ** dbs.

        ' ** assetno
        ' ** MasterAssetDescription
        ' ** due
        ' ** rate
        ' ** TotalCost
        ' ** TotalShareface
        ' ** accountno
        ' ** shortname
        ' ** legalname
        ' ** assettype
        ' ** assettype_description
        ' ** totdesc
        ' ** icash
        ' ** pcash
        ' ** currentDate
        ' ** CompanyName
        ' ** CompanyAddress1
        ' ** CompanyAddress2
        ' ** CompanyCity
        ' ** CompanyState
        ' ** CompanyZip
        ' ** CompanyPhone
        ' ** MarketValueX
        ' ** MarketValueCurrentX
        ' ** YieldX

EXITP:
1290    Set fld = Nothing
1300    Set rst2 = Nothing
1310    Set qdf = Nothing
1320    FillAListTmp_SP = blnRetVal
1330    Exit Function

ERRH:
1340    blnRetVal = False
1350    Select Case ERR.Number
        Case Else
1360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1370    End Select
1380    Resume EXITP

End Function

Public Sub DetailMouse_SP(blnAnnualStatement_Focus As Boolean, blnBalanceTable_Focus As Boolean, blnCalendar1_Focus As Boolean, blnCalendar2_Focus As Boolean, blnCalendar3_Focus As Boolean, frm As Access.Form)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "DetailMouse_SP"

1410    With frm
1420      If .cmdAnnualStatement_raised_focus_dots_img.Visible = True Or .cmdAnnualStatement_raised_focus_img.Visible = True Then
1430        Select Case blnAnnualStatement_Focus
            Case True
1440          .cmdAnnualStatement_raised_semifocus_dots_img.Visible = True
1450          .cmdAnnualStatement_raised_img.Visible = False
1460        Case False
1470          .cmdAnnualStatement_raised_img.Visible = True
1480          .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
1490        End Select
1500        .cmdAnnualStatement_raised_focus_img.Visible = False
1510        .cmdAnnualStatement_raised_focus_dots_img.Visible = False
1520        .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
1530        .cmdAnnualStatement_raised_img_dis.Visible = False
1540      End If
1550      If .cmdBalanceTable_raised_focus_dots_img.Visible = True Or .cmdBalanceTable_raised_focus_img.Visible = True Then
1560        Select Case blnBalanceTable_Focus
            Case True
1570          .cmdBalanceTable_raised_semifocus_dots_img.Visible = True
1580          .cmdBalanceTable_raised_img.Visible = False
1590        Case False
1600          .cmdBalanceTable_raised_img.Visible = True
1610          .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
1620        End Select
1630        .cmdBalanceTable_raised_focus_img.Visible = False
1640        .cmdBalanceTable_raised_focus_dots_img.Visible = False
1650        .cmdBalanceTable_sunken_focus_dots_img.Visible = False
1660        .cmdBalanceTable_raised_img_dis.Visible = False
1670      End If
1680      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
1690        Select Case blnCalendar1_Focus
            Case True
1700          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
1710          .cmdCalendar1_raised_img.Visible = False
1720        Case False
1730          .cmdCalendar1_raised_img.Visible = True
1740          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
1750        End Select
1760        .cmdCalendar1_raised_focus_img.Visible = False
1770        .cmdCalendar1_raised_focus_dots_img.Visible = False
1780        .cmdCalendar1_sunken_focus_dots_img.Visible = False
1790        .cmdCalendar1_raised_img_dis.Visible = False
1800      End If
1810      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
1820        Select Case blnCalendar2_Focus
            Case True
1830          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
1840          .cmdCalendar2_raised_img.Visible = False
1850        Case False
1860          .cmdCalendar2_raised_img.Visible = True
1870          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
1880        End Select
1890        .cmdCalendar2_raised_focus_img.Visible = False
1900        .cmdCalendar2_raised_focus_dots_img.Visible = False
1910        .cmdCalendar2_sunken_focus_dots_img.Visible = False
1920        .cmdCalendar2_raised_img_dis.Visible = False
1930      End If
1940      If .cmdCalendar3_raised_focus_dots_img.Visible = True Or .cmdCalendar3_raised_focus_img.Visible = True Then
1950        Select Case blnCalendar3_Focus
            Case True
1960          .cmdCalendar3_raised_semifocus_dots_img.Visible = True
1970          .cmdCalendar3_raised_img.Visible = False
1980        Case False
1990          .cmdCalendar3_raised_img.Visible = True
2000          .cmdCalendar3_raised_semifocus_dots_img.Visible = False
2010        End Select
2020        .cmdCalendar3_raised_focus_img.Visible = False
2030        .cmdCalendar3_raised_focus_dots_img.Visible = False
2040        .cmdCalendar3_sunken_focus_dots_img.Visible = False
2050        .cmdCalendar3_raised_img_dis.Visible = False
2060      End If
2070    End With

EXITP:
2080    Exit Sub

ERRH:
2090    Select Case ERR.Number
        Case Else
2100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2110    End Select
2120    Resume EXITP

End Sub

Public Sub AnnualBalance_Handler_SP(strProc As String, blnAnnualStatement_Focus As Boolean, blnAnnualStatement_MouseDown As Boolean, blnBalanceTable_Focus As Boolean, blnBalanceTable_MouseDown As Boolean, frm As Access.Form)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "AnnualBalance_Handler_SP"

        Dim strCalled As String, strAction As String
        Dim intPos01 As Integer

2210    With frm

2220      intPos01 = InStr(strProc, "_")
2230      strCalled = Left(strProc, (intPos01 - 1))
2240      strAction = Mid(strProc, (intPos01 + 1))

2250      Select Case strCalled
          Case "cmdAnnualStatement"

2260        Select Case strAction
            Case "GotFocus"
2270          blnAnnualStatement_Focus = True
2280          .cmdAnnualStatement_raised_semifocus_dots_img.Visible = True
2290          .cmdAnnualStatement_raised_img.Visible = False
2300          .cmdAnnualStatement_raised_focus_img.Visible = False
2310          .cmdAnnualStatement_raised_focus_dots_img.Visible = False
2320          .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
2330          .cmdAnnualStatement_raised_img_dis.Visible = False
2340        Case "MouseDown"
2350          blnAnnualStatement_MouseDown = True
2360          .cmdAnnualStatement_sunken_focus_dots_img.Visible = True
2370          .cmdAnnualStatement_raised_img.Visible = False
2380          .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
2390          .cmdAnnualStatement_raised_focus_img.Visible = False
2400          .cmdAnnualStatement_raised_focus_dots_img.Visible = False
2410          .cmdAnnualStatement_raised_img_dis.Visible = False
2420        Case "MouseMove"
2430          If blnAnnualStatement_MouseDown = False Then
2440            Select Case blnAnnualStatement_Focus
                Case True
2450              .cmdAnnualStatement_raised_focus_dots_img.Visible = True
2460              .cmdAnnualStatement_raised_focus_img.Visible = False
2470            Case False
2480              .cmdAnnualStatement_raised_focus_img.Visible = True
2490              .cmdAnnualStatement_raised_focus_dots_img.Visible = False
2500            End Select
2510            .cmdAnnualStatement_raised_img.Visible = False
2520            .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
2530            .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
2540            .cmdAnnualStatement_raised_img_dis.Visible = False
2550          End If
2560          If .cmdBalanceTable_raised_focus_dots_img.Visible = True Or .cmdBalanceTable_raised_focus_img.Visible = True Then
2570            Select Case blnBalanceTable_Focus
                Case True
2580              .cmdBalanceTable_raised_semifocus_dots_img.Visible = True
2590              .cmdBalanceTable_raised_img.Visible = False
2600            Case False
2610              .cmdBalanceTable_raised_img.Visible = True
2620              .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
2630            End Select
2640            .cmdBalanceTable_raised_focus_img.Visible = False
2650            .cmdBalanceTable_raised_focus_dots_img.Visible = False
2660            .cmdBalanceTable_sunken_focus_dots_img.Visible = False
2670            .cmdBalanceTable_raised_img_dis.Visible = False
2680          End If
2690        Case "MouseUp"
2700          .cmdAnnualStatement_raised_focus_dots_img.Visible = True
2710          .cmdAnnualStatement_raised_img.Visible = False
2720          .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
2730          .cmdAnnualStatement_raised_focus_img.Visible = False
2740          .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
2750          .cmdAnnualStatement_raised_img_dis.Visible = False
2760          blnAnnualStatement_MouseDown = False
2770        Case "LostFocus"
2780          .cmdAnnualStatement_raised_img.Visible = True
2790          .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
2800          .cmdAnnualStatement_raised_focus_img.Visible = False
2810          .cmdAnnualStatement_raised_focus_dots_img.Visible = False
2820          .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
2830          .cmdAnnualStatement_raised_img_dis.Visible = False
2840          blnAnnualStatement_Focus = False
2850        End Select

2860      Case "cmdBalanceTable"

2870        Select Case strAction
            Case "GotFocus"
2880          blnBalanceTable_Focus = True
2890          .cmdBalanceTable_raised_semifocus_dots_img.Visible = True
2900          .cmdBalanceTable_raised_img.Visible = False
2910          .cmdBalanceTable_raised_focus_img.Visible = False
2920          .cmdBalanceTable_raised_focus_dots_img.Visible = False
2930          .cmdBalanceTable_sunken_focus_dots_img.Visible = False
2940          .cmdBalanceTable_raised_img_dis.Visible = False
2950        Case "MouseDown"
2960          blnBalanceTable_MouseDown = True
2970          .cmdBalanceTable_sunken_focus_dots_img.Visible = True
2980          .cmdBalanceTable_raised_img.Visible = False
2990          .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
3000          .cmdBalanceTable_raised_focus_img.Visible = False
3010          .cmdBalanceTable_raised_focus_dots_img.Visible = False
3020          .cmdBalanceTable_raised_img_dis.Visible = False
3030        Case "MouseMove"
3040          If blnBalanceTable_MouseDown = False Then
3050            Select Case blnBalanceTable_Focus
                Case True
3060              .cmdBalanceTable_raised_focus_dots_img.Visible = True
3070              .cmdBalanceTable_raised_focus_img.Visible = False
3080            Case False
3090              .cmdBalanceTable_raised_focus_img.Visible = True
3100              .cmdBalanceTable_raised_focus_dots_img.Visible = False
3110            End Select
3120            .cmdBalanceTable_raised_img.Visible = False
3130            .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
3140            .cmdBalanceTable_sunken_focus_dots_img.Visible = False
3150            .cmdBalanceTable_raised_img_dis.Visible = False
3160          End If
3170          If .cmdAnnualStatement_raised_focus_dots_img.Visible = True Or .cmdAnnualStatement_raised_focus_img.Visible = True Then
3180            Select Case blnAnnualStatement_Focus
                Case True
3190              .cmdAnnualStatement_raised_semifocus_dots_img.Visible = True
3200              .cmdAnnualStatement_raised_img.Visible = False
3210            Case False
3220              .cmdAnnualStatement_raised_img.Visible = True
3230              .cmdAnnualStatement_raised_semifocus_dots_img.Visible = False
3240            End Select
3250            .cmdAnnualStatement_raised_focus_img.Visible = False
3260            .cmdAnnualStatement_raised_focus_dots_img.Visible = False
3270            .cmdAnnualStatement_sunken_focus_dots_img.Visible = False
3280            .cmdAnnualStatement_raised_img_dis.Visible = False
3290          End If
3300        Case "MouseUp"
3310          .cmdBalanceTable_raised_focus_dots_img.Visible = True
3320          .cmdBalanceTable_raised_img.Visible = False
3330          .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
3340          .cmdBalanceTable_raised_focus_img.Visible = False
3350          .cmdBalanceTable_sunken_focus_dots_img.Visible = False
3360          .cmdBalanceTable_raised_img_dis.Visible = False
3370          blnBalanceTable_MouseDown = False
3380        Case "LostFocus"
3390          .cmdBalanceTable_raised_img.Visible = True
3400          .cmdBalanceTable_raised_semifocus_dots_img.Visible = False
3410          .cmdBalanceTable_raised_focus_img.Visible = False
3420          .cmdBalanceTable_raised_focus_dots_img.Visible = False
3430          .cmdBalanceTable_sunken_focus_dots_img.Visible = False
3440          .cmdBalanceTable_raised_img_dis.Visible = False
3450          blnBalanceTable_Focus = False
3460        End Select

3470      End Select
3480    End With

EXITP:
3490    Exit Sub

ERRH:
3500    Select Case ERR.Number
        Case Else
3510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3520    End Select
3530    Resume EXITP

End Sub

Public Sub TransAssetState_Handler_SP(strProc As String, frm As Access.Form)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "TransAssetState_Handler_SP"

        Dim strCalled As String, strAction As String
        Dim intPos01 As Integer

3610    With frm

3620      intPos01 = InStr(strProc, "_")
3630      strCalled = Left(strProc, (intPos01 - 1))
3640      strAction = Mid(strProc, (intPos01 + 1))

3650      Select Case strCalled
          Case "chkTransactions"

3660        Select Case strAction
            Case "GotFocus"
3670          .chkTransactions_box.BackColor = MY_CLR_VLTBGE
3680          .TransDateStart_lbl_box.BorderColor = MY_CLR_VLTBGE  ' ** Until they get turned off.
3690          .TransDateEnd_lbl_box.BorderColor = MY_CLR_VLTBGE
3700          .chkArchive_Trans_box.BackColor = MY_CLR_VLTBGE
3710          .chkAssetList_box.BackColor = MY_CLR_LTBGE
3720          .chkIncludeArchive_Asset_box.BackColor = MY_CLR_LTBGE
3730          .chkStatements_box.BackColor = MY_CLR_LTBGE
3740          .chkStatements_box2.BackColor = MY_CLR_LTBGE
3750          .cmdAnnualStatement_box.BackColor = MY_CLR_LTBGE
3760          If IsNull(.cmbMonth) = True And IsNull(.StatementsYear) = False Then
3770            .StatementsYear = Null
3780          End If
3790        End Select

3800      Case "chkAssetList"

3810        Select Case strAction
            Case "GotFocus"
3820          .chkAssetList_box.BackColor = MY_CLR_VLTBGE
3830          .AssetListDate_lbl_box.BorderColor = MY_CLR_VLTBGE  ' ** Until it gets turned off.
3840          .chkIncludeArchive_Asset_box.BackColor = MY_CLR_VLTBGE
3850          .chkTransactions_box.BackColor = MY_CLR_LTBGE
3860          .chkArchive_Trans_box.BackColor = MY_CLR_LTBGE
3870          .chkStatements_box.BackColor = MY_CLR_LTBGE
3880          .chkStatements_box2.BackColor = MY_CLR_LTBGE
3890          .cmdAnnualStatement_box.BackColor = MY_CLR_LTBGE
3900          If IsNull(.cmbMonth) = True And IsNull(.StatementsYear) = False Then
3910            .StatementsYear = Null
3920          End If
3930        End Select

3940      Case "chkStatements"

3950        Select Case strAction
            Case "GotFocus"
3960          .chkStatements_box.BackColor = MY_CLR_VLTBGE
3970          .cmbMonth_lbl_box.BorderColor = MY_CLR_VLTBGE  ' ** Until it gets turned off.
3980          .chkStatements_box2.BackColor = MY_CLR_VLTBGE
3990          .cmdAnnualStatement_box.BackColor = MY_CLR_VLTBGE
4000          .chkAssetList_box.BackColor = MY_CLR_LTBGE
4010          .chkIncludeArchive_Asset_box.BackColor = MY_CLR_LTBGE
4020          .chkTransactions_box.BackColor = MY_CLR_LTBGE
4030          .chkArchive_Trans_box.BackColor = MY_CLR_LTBGE
4040        End Select

4050      End Select

4060    End With

EXITP:
4070    Exit Sub

ERRH:
4080    Select Case ERR.Number
        Case Else
4090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4100    End Select
4110    Resume EXITP

End Sub

Public Sub Remember_Handler_SP(strProc As String, frm As Access.Form)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Remember_Handler_SP"

        Dim strCalled As String, strAction As String
        Dim intPos01 As Integer, lngCnt As Long

4210    With frm

4220      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
4230      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
4240      strCalled = Left(strProc, (intPos01 - 1))
4250      strAction = Mid(strProc, (intPos01 + 1))

4260      Select Case strCalled
          Case "chkRememberDates_Trans"
4270        Select Case strAction
            Case "AfterUpdate"
4280          Select Case .chkRememberDates_Trans
              Case True
4290            .chkRememberDates_Trans_lbl.FontBold = True
4300            .chkRememberDates_Trans_lbl2_dim.FontBold = True
4310            .chkRememberDates_Trans_lbl2_dim_hi.FontBold = True
4320          Case False
4330            .chkRememberDates_Trans_lbl.FontBold = False
4340            .chkRememberDates_Trans_lbl2_dim.FontBold = False
4350            .chkRememberDates_Trans_lbl2_dim_hi.FontBold = False
4360          End Select
4370        End Select
4380      Case "chkRememberDates_Asset"
4390        Select Case strAction
            Case "AfterUpdate"
4400          Select Case .chkRememberDates_Asset
              Case True
4410            .chkRememberDates_Asset_lbl.FontBold = True
4420            .chkRememberDates_Asset_lbl2_dim.FontBold = True
4430            .chkRememberDates_Asset_lbl2_dim_hi.FontBold = True
4440          Case False
4450            .chkRememberDates_Asset_lbl.FontBold = False
4460            .chkRememberDates_Asset_lbl2_dim.FontBold = False
4470            .chkRememberDates_Asset_lbl2_dim_hi.FontBold = False
4480          End Select
4490        End Select
4500      Case "chkRememberMe"
4510        Select Case strAction
            Case "AfterUpdate"
4520          Select Case .chkRememberMe
              Case True
4530            .chkRememberMe_lbl.FontBold = True
4540            .chkRememberMe_lbl2_dim.FontBold = True
4550            .chkRememberMe_lbl2_dim_hi.FontBold = True
4560          Case False
4570            .chkRememberMe_lbl.FontBold = False
4580            .chkRememberMe_lbl2_dim.FontBold = False
4590            .chkRememberMe_lbl2_dim_hi.FontBold = False
4600          End Select
4610        End Select
4620      End Select

4630    End With

EXITP:
4640    Exit Sub

ERRH:
4650    Select Case ERR.Number
        Case Else
4660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4670    End Select
4680    Resume EXITP

End Sub

Public Sub Archive_Handler_SP(strProc As String, blnLoop As Boolean, frm As Access.Form)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Archive_Handler_SP"

        Dim strCalled As String, strAction As String
        Dim intPos01 As Integer, lngCnt As Long

4710    With frm

4720      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
4730      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
4740      strCalled = Left(strProc, (intPos01 - 1))
4750      strAction = Mid(strProc, (intPos01 + 1))

4760      Select Case strCalled
          Case "chkIncludeArchive_Trans"
4770        Select Case strAction
            Case "AfterUpdate"
4780          Select Case .chkIncludeArchive_Trans
              Case True
4790            .chkIncludeArchive_Trans_lbl.FontBold = True
4800            If .chkArchiveOnly_Trans = True Then
4810              .chkArchiveOnly_Trans = False
4820              Select Case blnLoop
                  Case True
4830                blnLoop = False
4840              Case False
                    ' ** 1st.
4850                blnLoop = True
4860                .chkArchiveOnly_Trans_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
                    ' ** Calls chkIncludeArchive_Trans_AfterUpdate()  '** Recursive loop!
4870              End Select
4880            End If
4890          Case False
4900            .chkIncludeArchive_Trans_lbl.FontBold = False
4910          End Select
4920        End Select
4930      Case "chkArchiveOnly_Trans"
4940        Select Case strAction
            Case "AfterUpdate"
4950          Select Case .chkArchiveOnly_Trans
              Case True
4960            .chkArchiveOnly_Trans_lbl.FontBold = True
4970            If .chkIncludeArchive_Trans = True Then
4980              .chkIncludeArchive_Trans = False
4990              Select Case blnLoop
                  Case True
                    ' ** 2nd.
5000                blnLoop = False
5010              Case False
5020                blnLoop = True
5030                .chkIncludeArchive_Trans_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
                    'Calls chkArchiveOnly_Trans_AfterUpdate()  '** Recursive loop!
5040              End Select
5050            End If
                ' ** Force Specified account.
5060            If .opgAccountNumber <> .opgAccountNumber_optSpecified.OptionValue Then
5070              .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
5080              .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
5090            End If
5100            .opgAccountNumber_optAll.Enabled = False
5110          Case False
5120            .chkArchiveOnly_Trans_lbl.FontBold = False
5130            .opgAccountNumber_optAll.Enabled = True
5140            .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
5150          End Select
5160        End Select
5170      Case "chkIncludeArchive_Asset"
5180        Select Case strAction
            Case "AfterUpdate"
5190          Select Case .chkIncludeArchive_Asset
              Case True
5200            .chkIncludeArchive_Asset_lbl.FontBold = True
5210            .chkIncludeArchive_Asset_lbl2.FontBold = True
5220            .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = True
5230          Case False
5240            .chkIncludeArchive_Asset_lbl.FontBold = False
5250            .chkIncludeArchive_Asset_lbl2.FontBold = False
5260            .chkIncludeArchive_Asset_lbl2_dim_hi.FontBold = False
5270          End Select
5280          Select Case .chkIncludeArchive_Asset.Enabled
              Case True
5290            .chkIncludeArchive_Asset_lbl2.ForeColor = .chkIncludeArchive_Asset_lbl.ForeColor
5300            .chkIncludeArchive_Asset_lbl2_dim_hi.Visible = False
5310          Case False
5320            .chkIncludeArchive_Asset_lbl2.ForeColor = WIN_CLR_DISF
5330            .chkIncludeArchive_Asset_lbl2_dim_hi.Visible = True
5340          End Select
5350        End Select
5360      End Select

5370    End With

EXITP:
5380    Exit Sub

ERRH:
5390    Select Case ERR.Number
        Case Else
5400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5410    End Select
5420    Resume EXITP

End Sub

Public Sub AssetListDateLost_SP(frm As Access.Form)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "AssetListDateLost_SP"

        Dim strMsg As String

5510    With frm
5520      If IsNull(.AssetListDate) = True Then
            ' ** Populate it with today's date.
5530        .AssetListDate = Date
5540        .DateEnd = .AssetListDate
5550      Else
5560        If .AssetListDate = vbNullString Then
              ' ** Populate it with today's date.
5570          .AssetListDate = Date
5580          .DateEnd = .AssetListDate
5590        Else
5600          .DateEnd = .AssetListDate
5610          If .DateEnd <> Date Then
5620            If (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
                  ' ** Skip message.
5630            Else
5640              Select Case .HasForeign
                  Case True
5650                strMsg = "You have changed the date of this Asset List report." & vbCrLf & vbCrLf & _
                      "Shares/Face and Cost for the account selected will be as of this date." & vbCrLf & vbCrLf
5660                Select Case .chkForeignExchange
                    Case True
5670                  strMsg = strMsg & "Market Value will reflect the most appropriate data found in" & vbCrLf & _
                        "Asset Pricing History, and currency conversion rates will" & vbCrLf & _
                        "reflect the most appropriate data found in Currency History."
5680                Case False
5690                  strMsg = strMsg & "However, the Market Value will still be as of a Current Date."
5700                End Select
5710              Case False
5720                strMsg = "You have changed the date of this Asset List report." & vbCrLf & vbCrLf & _
                      "Shares/Face and Cost for the account selected will be as of this date." & vbCrLf & vbCrLf & _
                      "However, the Market Value will still be as of a Current Date."
5730              End Select
5740              MsgBox strMsg, vbInformation + vbOKOnly, (Left(("Date Change" & Space(55)), 55) & "D01")
5750            End If
5760          End If
5770        End If

5780      End If
5790    End With

EXITP:
5800    Exit Sub

ERRH:
5810    Select Case ERR.Number
        Case Else
5820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5830    End Select
5840    Resume EXITP

End Sub

Public Sub TransDateStmtYearExit_SP(intMode As Integer, frm As Access.Form)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStmtYearExit_SP"

        Dim msgResponse As VbMsgBoxResult

5910    With frm
5920      Select Case intMode
          Case 1
5930        If .chkTransactions = True And IsNull(.TransDateStart) = False Then
5940          If IsDate(.TransDateEnd) = True And IsDate(.TransDateStart) = True Then
5950            If CDate(.TransDateStart) > CDate(.TransDateEnd) Then
5960              msgResponse = MsgBox("The start date must be less than or equal to the end date." & vbCrLf & _
                    "Clear the end date?", vbQuestion + vbOKCancel, (Left(("Invalid Date" & Space(55)), 55) & "A01"))
5970              If msgResponse = vbOK Then
5980                .TransDateEnd = Null
5990                .TransDateEnd.SetFocus
6000              Else
6010                DoCmd.CancelEvent
6020              End If
6030            End If
6040          End If
6050        End If
6060      Case 2
6070        If .chkTransactions = True And IsNull(.TransDateStart) = False Then
6080          If IsNull(.TransDateEnd) Then
                ' ** Populate it with today's date.
6090            .TransDateEnd = Date
6100          Else
6110            If .TransDateEnd = vbNullString Then
                  ' ** Populate it with today's date.
6120              .TransDateEnd = Date
6130            End If
6140          End If
6150          If IsDate(.TransDateStart) = False Then
6160            MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, _
                  (Left(("Invalid Date" & Space(55)), 55) & "C01")
6170            .TransDateStart.SetFocus
6180          Else
6190            If CDate(.TransDateStart) > CDate(.TransDateEnd) Then
6200              MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, _
                    (Left(("Invalid Date" & Space(55)), 55) & "C02")
6210              .TransDateStart.SetFocus
6220            End If
6230          End If
6240        End If
6250      Case 3
6260        Select Case .cmbMonth.Column(CBX_MON_NAME)
            Case "January"
6270          .DateEnd = "01/31/" & .StatementsYear
6280        Case "February"
6290          .DateEnd = Format(CDate("03/01/" & .StatementsYear) - 1, "mm/dd/yyyy")
6300        Case "March"
6310          .DateEnd = "03/31/" & .StatementsYear
6320        Case "April"
6330          .DateEnd = "04/30/" & .StatementsYear
6340        Case "May"
6350          .DateEnd = "05/31/" & .StatementsYear
6360        Case "June"
6370          .DateEnd = "06/30/" & .StatementsYear
6380        Case "July"
6390          .DateEnd = "07/31/" & .StatementsYear
6400        Case "August"
6410          .DateEnd = "08/31/" & .StatementsYear
6420        Case "September"
6430          .DateEnd = "09/30/" & .StatementsYear
6440        Case "October"
6450          .DateEnd = "10/31/" & .StatementsYear
6460        Case "November"
6470          .DateEnd = "11/30/" & .StatementsYear
6480        Case "December"
6490          .DateEnd = "12/31/" & .StatementsYear
6500        End Select
6510        .AssetListDate = .DateEnd
6520      End Select
6530    End With

EXITP:
6540    Exit Sub

ERRH:
6550    Select Case ERR.Number
        Case Else
6560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6570    End Select
6580    Resume EXITP

End Sub

Public Sub AcctSourceAfter_SP(frm As Access.Form)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "AcctSourceAfter_SP"

        Dim strAccountNo As String

6610    With frm
6620      strAccountNo = vbNullString
6630      If IsNull(.cmbAccounts) = False Then
6640        If Len(.cmbAccounts.Column(CBX_A_ACTNO)) > 0 Then
6650          strAccountNo = .cmbAccounts.Column(CBX_A_ACTNO)
6660        End If
6670      End If
6680      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
6690        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
6700        .opgAccountSource_optNumber_lbl.FontBold = True
6710        .opgAccountSource_optNumber_lbl2.FontBold = True
6720        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
6730        .opgAccountSource_optName_lbl.FontBold = False
6740        .opgAccountSource_optName_lbl2.FontBold = False
6750        .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
6760      Case .opgAccountSource_optName.OptionValue
6770        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
6780        .opgAccountSource_optNumber_lbl.FontBold = False
6790        .opgAccountSource_optNumber_lbl2.FontBold = False
6800        .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
6810        .opgAccountSource_optName_lbl.FontBold = True
6820        .opgAccountSource_optName_lbl2.FontBold = True
6830        .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
6840      End Select
6850      DoEvents
6860      If strAccountNo <> vbNullString Then
6870        .cmbAccounts = strAccountNo
6880      End If
6890    End With

EXITP:
6900    Exit Sub

ERRH:
6910    Select Case ERR.Number
        Case Else
6920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6930    End Select
6940    Resume EXITP

End Sub

Public Sub Calendar_KeyDown_SP(KeyCode As Integer, Shift As Integer, strProc As String, frm As Access.Form)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_KeyDown_SP"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strCalled As String
        Dim intPos01 As Integer
        Dim intRetVal As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

7050    intPos01 = InStr(strProc, "_")
7060    strCalled = Left(strProc, (intPos01 - 1))

        ' ** Plain keys.
7070    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7080      Select Case intRetVal
          Case vbKeyTab
7090        With frm
7100          intRetVal = 0
7110          Select Case strCalled
              Case "cmdCalendar1"
7120            .TransDateEnd.SetFocus
7130          Case "cmdCalendar2"
7140            .chkRememberDates_Trans.SetFocus
7150          Case "cmdCalendar3"
7160            .chkRememberDates_Asset.SetFocus
7170          End Select
7180        End With
7190      End Select
7200    End If

        ' ** Shift keys.
7210    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7220      Select Case intRetVal
          Case vbKeyTab
7230        With frm
7240          intRetVal = 0
7250          Select Case strCalled
              Case "cmdCalendar1"
7260            .TransDateStart.SetFocus
7270          Case "cmdCalendar2"
7280            .TransDateEnd.SetFocus
7290          Case "cmdCalendar3"
7300            .AssetListDate.SetFocus
7310          End Select
7320        End With
7330      End Select
7340    End If

EXITP:
7350    KeyCode = intRetVal
7360    Exit Sub

ERRH:
7370    intRetVal = 0
7380    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7390    Case Else
7400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7410    End Select
7420    Resume EXITP

End Sub

Public Sub PrintStmt_KeyDown_SP(KeyCode As Integer, Shift As Integer, strProc As String, frm As Access.Form)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintStmt_KeyDown_SP"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strCalled As String
        Dim intPos01 As Integer, lngCnt As Long
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

7550    lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
7560    intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
7570    strCalled = Left(strProc, (intPos01 - 1))

        ' ** Plain keys.
7580    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7590      Select Case intRetVal
          Case vbKeyTab
7600        With frm
7610          intRetVal = 0
7620          Select Case strCalled
              Case "cmdPrintStatement_All"
7630            If .cmdPrintStatement_Single.Enabled = True Then
7640              .cmdPrintStatement_Single.SetFocus
7650            ElseIf .cmdPrintStatement_Summary.Enabled = True And .cmdPrintStatement_Summary.Visible = True Then
7660              .cmdPrintStatement_Summary.SetFocus
7670            Else
7680              .cmdClose.SetFocus
7690            End If
7700          Case "cmdPrintStatement_Single"
7710            If .cmdPrintStatement_Summary.Enabled = True And .cmdPrintStatement_Summary.Visible = True Then
7720              .cmdPrintStatement_Summary.SetFocus
7730            Else
7740              .cmdClose.SetFocus
7750            End If
7760          Case "cmdPrintStatement_Summary"
7770            .cmdClose.SetFocus
7780          End Select
7790        End With
7800      Case vbKeyLeft
7810        With frm
7820          intRetVal = 0
7830          Select Case strCalled
              Case "cmdPrintStatement_All"
7840            If .cmdAssetListExcel.Enabled = True Then
7850              .cmdAssetListExcel.SetFocus
7860            ElseIf .cmdAssetListWord.Enabled = True Then
7870              .cmdAssetListWord.SetFocus
7880            ElseIf .cmdTransactionsExcel.Enabled = True Then
7890              .cmdTransactionsExcel.SetFocus
7900            ElseIf .cmdTransactionsWord.Enabled = True Then
7910              .cmdTransactionsWord.SetFocus
7920            End If
7930          Case "cmdPrintStatement_Single"
7940            If .cmdPrintStatement_All.Enabled = True Then
7950              .cmdPrintStatement_All.SetFocus
7960            ElseIf .cmdAssetListExcel.Enabled = True Then
7970              .cmdAssetListExcel.SetFocus
7980            ElseIf .cmdAssetListWord.Enabled = True Then
7990              .cmdAssetListWord.SetFocus
8000            ElseIf .cmdTransactionsExcel.Enabled = True Then
8010              .cmdTransactionsExcel.SetFocus
8020            ElseIf .cmdTransactionsWord.Enabled = True Then
8030              .cmdTransactionsWord.SetFocus
8040            End If
8050          Case "cmdPrintStatement_Summary"
8060            If .cmdPrintStatement_Single.Enabled = True Then
8070              .cmdPrintStatement_Single.SetFocus
8080            ElseIf .cmdPrintStatement_All.Enabled = True Then
8090              .cmdPrintStatement_All.SetFocus
8100            ElseIf .cmdAssetListExcel.Enabled = True Then
8110              .cmdAssetListExcel.SetFocus
8120            ElseIf .cmdAssetListWord.Enabled = True Then
8130              .cmdAssetListWord.SetFocus
8140            ElseIf .cmdTransactionsExcel.Enabled = True Then
8150              .cmdTransactionsExcel.SetFocus
8160            ElseIf .cmdTransactionsWord.Enabled = True Then
8170              .cmdTransactionsWord.SetFocus
8180            End If
8190          End Select
8200        End With
8210      Case vbKeyRight
8220        With frm
8230          intRetVal = 0
8240          Select Case strCalled
              Case "cmdPrintStatement_All"
8250            If .cmdPrintStatement_Single.Enabled = True Then
8260              .cmdPrintStatement_Single.SetFocus
8270            ElseIf .cmdPrintStatement_Summary.Enabled = True Then
8280              .cmdPrintStatement_Summary.SetFocus
8290            ElseIf .cmdTransactionsPreview.Enabled = True Then
8300              .cmdTransactionsPreview.SetFocus
8310            Else
8320              .cmdAssetListPreview.SetFocus
8330            End If
8340          Case "cmdPrintStatement_Single"
8350            If .cmdPrintStatement_Summary.Enabled = True Then
8360              .cmdPrintStatement_Summary.SetFocus
8370            ElseIf .cmdTransactionsPreview.Enabled = True Then
8380              .cmdTransactionsPreview.SetFocus
8390            Else
8400              .cmdAssetListPreview.SetFocus
8410            End If
8420          Case "cmdPrintStatement_Summary"
8430            If .cmdTransactionsPreview.Enabled = True Then
8440              .cmdTransactionsPreview.SetFocus
8450            Else
8460              .cmdAssetListPreview.SetFocus
8470            End If
8480          End Select
8490        End With
8500      Case vbKeyUp
8510        With frm
8520          intRetVal = 0
8530          Select Case strCalled
              Case "cmdPrintStatement_All"
8540            If .cmdTransactionsWord.Enabled = True Then
8550              .cmdTransactionsWord.SetFocus
8560            ElseIf .cmdPrintStatement_Summary.Enabled = True Then
8570              .cmdPrintStatement_Summary.SetFocus
8580            ElseIf .cmdAssetListExcel.Enabled = True Then
8590              .cmdAssetListExcel.SetFocus
8600            ElseIf .cmdAssetListPrint.Enabled = True Then
8610              .cmdAssetListPrint.SetFocus
8620            ElseIf .cmdTransactionsExcel.Enabled = True Then
8630              .cmdTransactionsExcel.SetFocus
8640            Else
8650              .cmdTransactionsPrint.SetFocus
8660            End If
8670          Case "cmdPrintStatement_Single"
8680            If .cmdAssetListWord.Enabled = True Then
8690              .cmdAssetListWord.SetFocus
8700            ElseIf .cmdTransactionsExcel.Enabled = True Then
8710              .cmdTransactionsExcel.SetFocus
8720            Else
8730              .cmdTransactionsPrint.SetFocus
8740            End If
8750          Case "cmdPrintStatement_Summary"
8760            If .cmdAssetListExcel.Enabled = True Then
8770              .cmdAssetListExcel.SetFocus
8780            ElseIf .cmdAssetListPrint.Enabled = True Then
8790              .cmdAssetListPrint.SetFocus
8800            ElseIf .cmdPrintStatement_Single.Enabled = True Then
8810              .cmdPrintStatement_Single.SetFocus
8820            ElseIf .cmdTransactionsExcel.Enabled = True Then
8830              .cmdTransactionsExcel.SetFocus
8840            Else
8850              .cmdTransactionsPrint.SetFocus
8860            End If
8870          End Select
8880        End With
8890      Case vbKeyDown
8900        With frm
8910          intRetVal = 0
8920          Select Case strCalled
              Case "cmdPrintStatement_All"
8930            If .cmdTransactionsPrint.Enabled = True Then
8940              .cmdTransactionsPrint.SetFocus
8950            Else
8960              .cmdAssetListPreview.SetFocus
8970            End If
8980          Case "cmdPrintStatement_Single"
8990            If .cmdAssetListPrint.Enabled = True Then
9000              .cmdAssetListPrint.SetFocus
9010            ElseIf .cmdPrintStatement_Summary.Enabled = True Then
9020              .cmdPrintStatement_Summary.SetFocus
9030            Else
9040              .cmdTransactionsPreview.SetFocus
9050            End If
9060          Case "cmdPrintStatement_Summary"
9070            If .cmdTransactionsPreview.Enabled = True Then
9080              .cmdTransactionsPreview.SetFocus
9090            ElseIf .cmdPrintStatement_All.Enabled = True Then
9100              .cmdPrintStatement_All.SetFocus
9110            Else
9120              .cmdAssetListPreview.SetFocus
9130            End If
9140          End Select
9150        End With
9160      End Select
9170    End If

        ' ** Shift keys.
9180    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9190      Select Case intRetVal
          Case vbKeyTab
9200        With frm
9210          intRetVal = 0
9220          Select Case strCalled
              Case "cmdPrintStatement_All"
9230            If .cmdAssetListExcel.Enabled = True Then
9240              .cmdAssetListExcel.SetFocus
9250            ElseIf .cmdAssetListWord.Enabled = True Then
9260              .cmdAssetListWord.SetFocus
9270            ElseIf .cmdTransactionsExcel.Enabled = True Then
9280              .cmdTransactionsExcel.SetFocus
9290            ElseIf .cmdTransactionsWord.Enabled = True Then
9300              .cmdTransactionsWord.SetFocus
9310            ElseIf .cmbAccounts.Enabled = True Then
9320              .cmbAccounts.SetFocus
9330            Else
9340              .opgAccountNumber.SetFocus
9350            End If
9360          Case "cmdPrintStatement_Single"
9370            If .cmdPrintStatement_All.Enabled = True Then
9380              .cmdPrintStatement_All.SetFocus
9390            ElseIf .cmdAssetListExcel.Enabled = True Then
9400              .cmdAssetListExcel.SetFocus
9410            ElseIf .cmdAssetListWord.Enabled = True Then
9420              .cmdAssetListWord.SetFocus
9430            ElseIf .cmdTransactionsExcel.Enabled = True Then
9440              .cmdTransactionsExcel.SetFocus
9450            ElseIf .cmdTransactionsWord.Enabled = True Then
9460              .cmdTransactionsWord.SetFocus
9470            ElseIf .cmbAccounts.Enabled = True Then
9480              .cmbAccounts.SetFocus
9490            Else
9500              .opgAccountNumber.SetFocus
9510            End If
9520          Case "cmdPrintStatement_Summary"
9530            If .cmdPrintStatement_Single.Enabled = True Then
9540              .cmdPrintStatement_Single.SetFocus
9550            ElseIf .cmdPrintStatement_All.Enabled = True Then
9560              .cmdPrintStatement_All.SetFocus
9570            ElseIf .cmdAssetListExcel.Enabled = True Then
9580              .cmdAssetListExcel.SetFocus
9590            ElseIf .cmdAssetListWord.Enabled = True Then
9600              .cmdAssetListWord.SetFocus
9610            ElseIf .cmdTransactionsExcel.Enabled = True Then
9620              .cmdTransactionsExcel.SetFocus
9630            ElseIf .cmdTransactionsWord.Enabled = True Then
9640              .cmdTransactionsWord.SetFocus
9650            ElseIf .cmbAccounts.Enabled = True Then
9660              .cmbAccounts.SetFocus
9670            Else
9680              .opgAccountNumber.SetFocus
9690            End If
9700          End Select
9710        End With
9720      End Select
9730    End If

EXITP:
9740    KeyCode = intRetVal
9750    Exit Sub

ERRH:
9760    intRetVal = 0
9770    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9780    Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
9790    Case Else
9800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9810    End Select
9820    Resume EXITP

End Sub
