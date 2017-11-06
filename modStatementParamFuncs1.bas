Attribute VB_Name = "modStatementParamFuncs1"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modStatementParamFuncs1"

'VGC 09/06/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  39180
' #########################

' ** Array: arr_varAcctArch().
Private lngAcctArchs As Long, arr_varAcctArch As Variant
Private Const AR_ACTNO As Integer = 0
'Private Const AR_TDATE As Integer = 1
'Private Const AR_CNT   As Integer = 2

' ** Array: arr_varAcctFor().
Private lngAcctFors As Long, arr_varAcctFor As Variant
Private Const F_ACTNO As Integer = 0
Private Const F_JCNT  As Integer = 1
Private Const F_ACNT  As Integer = 2
Private Const F_SUPP  As Integer = 3

' ** Array: arr_varStmt().
Private lngStmts As Long, lngStmtCnt As Long, arr_varStmt() As Variant
Private Const S_ELEMS1 As Integer = 12  ' ** Array's first-element UBound().
Private Const S_ELEMS2 As Integer = 4   ' ** Array's second-element UBound().
Private Const S_MID   As Integer = 0  'month_id
Private Const S_MSHT  As Integer = 1  'month_short
Private Const S_CNT   As Integer = 2  'cnt_smt
Private Const S_ACTNO As Integer = 3  'accountno
Private Const S_SNAM  As Integer = 4  'shortname

' ** cmbAccounts combo box constants:
Private Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
'Private Const CBX_A_DESC   As Integer = 1  ' ** Desc
Private Const CBX_A_PREDAT As Integer = 2  ' ** predate
Private Const CBX_A_SHORT  As Integer = 3  ' ** shortname
'Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname
Private Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
Private Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
'Private Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
Private Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

' ** cmbMonth combo box constants:
Private Const CBX_MON_ID    As Integer = 0  ' ** month_id (same as month number)
Private Const CBX_MON_NAME  As Integer = 1  ' ** month_name
Private Const CBX_MON_SHORT As Integer = 2  ' ** month_short

Private blnIncludeCurrency As Boolean
' **

Public Function BuildTransactionInfo_SP(frm As Access.Form, strFileName As String, strReportName As String, blnAllStatements As Boolean, blnContinue As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean, blnFromStmts As Boolean, Optional varOutput As Variant) As Boolean
'cmdTransactionsPreview_Click: {missing}.
'cmdTransactionsPrint_Click: {missing}.
'cmdTransactionsWord_Click: "Word".
'cmdTransactionsExcel_Click: "Excel".

100   On Error GoTo ERRH

        Const THIS_PROC As String = "BuildTransactionInfo_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strSQL As String, strDocName As String
        Dim blnNoAccount As Boolean
        Dim intRetVal_SetDateSpecificSQL As Integer
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, blnTmp04 As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

110     blnContinue = True
120     blnRetVal = True  ' ** Unless proven otherwise.
130     blnNoAccount = False

140     With frm

150       If lngAcctFors = 0& Or IsEmpty(arr_varAcctFor) = True Then
160         ForExArr_Load  ' ** Function: Below.
170       End If

180       Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
190         gstrAccountNo = .cmbAccounts
200       Case .opgAccountNumber_optAll.OptionValue
210         gstrAccountNo = "All"
220       End Select

230       Select Case .chkTransactions
          Case True
            ' ** Transactions selected.
240         If IsNull(.TransDateStart) Then
250           blnRetVal = False
260           DoCmd.Hourglass False
270           MsgBox "Must enter the start date to continue.", vbInformation + vbOKOnly, (Left(("Entry Required" & Space(55)), 55) & "Y01")
280           .TransDateStart.SetFocus
290           blnContinue = False
300         Else
310           If IsNull(.TransDateEnd) Then
320             blnRetVal = False
330             DoCmd.Hourglass False
340             MsgBox "Must enter the end date to continue.", vbInformation + vbOKOnly, (Left(("Entry Required" & Space(55)), 55) & "Y02")
350             .TransDateStart.SetFocus
360             blnContinue = False
370           End If
380         End If
390       Case False
            ' ** Transactions not selected, so it's for Statements.
400         If IsNull(.DateEnd) = True And IsNull(.AssetListDate) = True Then
410           blnRetVal = False
420           DoCmd.Hourglass False
430           MsgBox "Must enter Period Ending date to continue.", vbInformation + vbOKOnly, (Left(("Entry Required" & Space(55)), 55) & "Y03")
440           .AssetListDate.SetFocus
450           blnContinue = False
460         ElseIf IsNull(.DateEnd) = True Then
470           .DateEnd = .AssetListDate
480         End If
490       End Select  ' ** chkTransactions.

500       If blnRetVal = True Then
            ' ** A DateStart and DateEnd are needed for SetDateSpecificSQL().
            ' ** DateEnd is already assured by the time it gets here.
510         If .chkStatements = True Then
520           glngMonthID = .cmbMonth.Column(CBX_MON_ID)
              ' ** DateStart defaults to the last statement date.
              ' ** It can be different dates depending on the last balance date,
              ' ** which I don't believe has been calculated yet.
530           strTmp01 = Right("00" & CStr(glngMonthID), 2)
              'NOT ALWAYS THE 31ST!
540           strTmp01 = strTmp01 & "/01/" & CStr(.StatementsYear)  ' ** Should be same as DateEnd.
550           gdatEndDate = CDate(DateAdd("y", -1, CDate(DateAdd("m", 1, CDate(strTmp01)))))
              ' ** Balance, by GlobalVarGet('gstrAccountno'), GlobalVarGet('gdatEndDate').
560           varTmp00 = DLookup("[balance_date]", "qryStatementParameters_32")  ' ** Includes underscore.
570           Select Case IsNull(varTmp00)
              Case True
580             Select Case IsNull(.DateStart)
                Case True
590               gdatStartDate = CDate(DateAdd("y", 1, CDate(DateAdd("yyyy", -1, gdatEndDate))))
600             Case False
610               gdatStartDate = .DateStart
620             End Select
630           Case False
640             gdatStartDate = varTmp00
650           End Select
660           If gdatStartDate >= gdatEndDate Then
                ' ** Now default to 1 year prior to end date, plus 1 day.
670             gdatStartDate = CDate(DateAdd("y", 1, CDate(DateAdd("yyyy", -1, gdatEndDate))))
680           End If
690           .DateStart = gdatStartDate
700         End If
710       End If  ' ** blnRetVal.
720       DoEvents

730       If blnRetVal = True Then
740         If .cmbAccounts.Enabled = True Then
750           If IsNull(.cmbAccounts) = True Then
760             blnNoAccount = True
770           Else
780             If .cmbAccounts = vbNullString Then
790               blnNoAccount = True
800             End If
810           End If
820           If blnNoAccount = True Then
830             blnRetVal = False
840             DoCmd.Hourglass False
850             MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "Y04")
860             blnContinue = False
870           End If
880         End If
890       End If  ' ** blnRetVal.

900       If blnRetVal = True Then

910         If .chkArchiveOnly_Trans = True And .chkTransactions = True Then  ' ** Archived transactions only.
              ' ** This is ONLY LedgerArchive.

920           Set dbs = CurrentDb
              ' ** LedgerArchive, by specified FormRef('accountno'), FormRef('StartDateTrans'), FormRef('EndDateTrans').  #curr_id
930           Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_14")  ' ** #curr_id
940           strSQL = qdf.SQL
950           dbs.Close
960           Set qdf = Nothing

970         Else  ' ** Normal.
              ' ** This is ONLY Legder.
              ' ** No! If it's statements, LedgerArchive is automatically included.

              ' ** This code will update the qryMaxBalDates query to give us the
              ' ** balance numbers from the previous statement.
980           Select Case .opgAccountNumber
              Case .opgAccountNumber_optSpecified.OptionValue
                ' ** Specified Account.
990             Select Case .chkTransactions
                Case True
                  ' ** Checked.
1000              intRetVal_SetDateSpecificSQL = SetDateSpecificSQL(gstrAccountNo, "StatementTransactions", frm.Name)
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -4  Date criteria not met.
                  ' **   -9  Error.
1010              If intRetVal_SetDateSpecificSQL <> 0 Then
1020                blnRetVal = False
1030                DoCmd.Hourglass False
1040                MsgBox "There are no transactions for this account for the period selected.", _
                      vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "Y05")
1050                blnContinue = False
1060              End If
1070            Case False
1080              intRetVal_SetDateSpecificSQL = SetDateSpecificSQL(gstrAccountNo, "Statements", frm.Name)
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -4  Date criteria not met.
                  ' **   -9  Error.
1090              If intRetVal_SetDateSpecificSQL <> 0 And blnAllStatements = False Then
1100                blnRetVal = False
1110                DoCmd.Hourglass False
1120                MsgBox "There are no transactions for this account for the period selected.", _
                      vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "Y06")
1130                blnContinue = False
1140              ElseIf blnAllStatements = True Then
                    'Debug.Print "'Account " & .cmbAccounts & " has no transactions."
1150              End If
1160            End Select  ' ** chkTransactions.
1170          Case .opgAccountNumber_optAll.OptionValue

                ' ** All accounts.
1180            Select Case .chkTransactions
                Case True
                  ' ** Checked.
1190              intRetVal_SetDateSpecificSQL = SetDateSpecificSQL("All", "StatementTransactions", frm.Name)
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -4  Date criteria not met.
                  ' **   -9  Error.
1200              If intRetVal_SetDateSpecificSQL <> 0 Then
                    'Debug.Print "'intRetVal_SetDateSpecificSQL = " & CStr(intRetVal_SetDateSpecificSQL)
1210                blnRetVal = False
1220                DoCmd.Hourglass False
1230                MsgBox "No Transactions.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "Y07")
1240                blnContinue = False
1250              End If
1260            Case False
1270              intRetVal_SetDateSpecificSQL = SetDateSpecificSQL("All", "Statements", frm.Name)
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -4  Date criteria not met.
                  ' **   -9  Error.
1280              If intRetVal_SetDateSpecificSQL <> 0 Then
1290                blnRetVal = False
1300                DoCmd.Hourglass False
1310                MsgBox "No Transactions.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "Y08")
1320                blnContinue = False
1330              End If
1340            End Select  ' ** chkTransactions.
1350          End Select  ' ** opgAccountNumber.

1360          If blnRetVal = True Then
1370            Select Case .chkTransactions
                Case True
                  'WHERE'S ARCHIVE?!
                  ' ** Transactions has been selected, so only transactions can be viewed.
                  ' ** We must query on the dates that are in the Transaction block.
1380              Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
                    ' ** Specified account.
1390                Set dbs = CurrentDb
1400                Select Case .chkArchiveOnly_Trans
                    Case True
                      ' ** LedgerArchive, by specified FormRef('accountno'), FormRef('StartDateTrans'), FormRef('EndDateTrans').
1410                  Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_10_01")
1420                Case False
1430                  Select Case .chkIncludeArchive_Trans
                      Case True
                        ' ** Union of qryStatementParameters_Trans_10 (Ledger, by specified FormRef('accountno'),
                        ' ** FormRef('StartDateTrans'), FormRef('EndDateTrans')), qryStatementParameters_Trans_10_01
                        ' ** (LedgerArchive, by specified FormRef('accountno'), FormRef('StartDateTrans'),
                        ' ** FormRef('EndDateTrans')).
1440                    Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_10_02")
1450                  Case False
                        ' ** Ledger, by specified FormRef('accountno'), FormRef('StartDateTrans'), FormRef('EndDateTrans').
1460                    Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_10")
1470                  End Select
1480                End Select
1490                strSQL = qdf.SQL
1500                Set qdf = Nothing
1510                dbs.Close
1520                Set dbs = Nothing
1530                strFileName = .cmbAccounts.Column(CBX_A_ACTNO) & " Transaction Rpts " & .TransDateStart & " through " & .TransDateEnd
1540              Case .opgAccountNumber_optAll.OptionValue
                    ' ** All accounts.
1550                Set dbs = CurrentDb
1560                Select Case .chkArchiveOnly_Trans
                    Case True
                      ' ** LedgerArchive, by specified FormRef('StartDateTrans','EndDateTrans').
1570                  Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_11_01")
1580                Case False
1590                  Select Case .chkIncludeArchive_Trans
                      Case True
                        ' ** Union of qryStatementParameters_Trans_11 (Ledger, by specified FormRef('StartDateTrans'),
                        ' ** FormRef('EndDateTrans')), qryStatementParameters_Trans_11_01 (LedgerArchive, by specified
                        ' ** FormRef('StartDateTrans','EndDateTrans')).
1600                    Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_11_02")
1610                  Case False
                        ' ** Ledger, by specified FormRef('StartDateTrans'), FormRef('EndDateTrans').
1620                    Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_11")
1630                  End Select
1640                End Select
1650                strSQL = qdf.SQL
1660                Set qdf = Nothing
1670                dbs.Close
1680                Set dbs = Nothing
1690                strFileName = "All Transaction Rpts " & .TransDateStart & " through " & .TransDateEnd
1700              End Select  ' ** opgAccountNumber.
1710            Case False
                  ' ** Transactions not selected, so it's for Statements.
1720              Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
                    ' ** Specified account.
1730                Set dbs = CurrentDb

                    ' ** Ledger, linked to qryStatementParameters_Trans_09a (Balance table, by specified
                    ' ** FormRef('EndDate')), with add'l fields, by specified FormRef('accountno'), FormRef('EndDate').
                    'Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_12")
                    'strSQL = qdf.SQL
                    'THIS strSQL IS REPLACED BELOW!
                    ' ************************************************************************************************************
                    ' ************************************************************************************************************
                    ' ** This all new (11/28/2015) in order to fix that 'Unknown Jet error' Message.
                    ' ************************************************************************************************************
                    ' ************************************************************************************************************

                    ' ** qryStatementParameters_Trans_09a -> qryStatementParameters_Trans_09_01
1740                strTmp01 = "SELECT Balance.accountno AS accountno, Max(Balance.[balance date]) AS [MaxOfbalance date]" & vbCrLf & _
                      "FROM Balance" & vbCrLf & _
                      "WHERE (((Balance.[balance date])<FormRef('EndDate')))" & vbCrLf & _
                      "GROUP BY Balance.accountno;"
1750                strTmp02 = "FormRef('EndDate')"
1760                strTmp03 = "#" & Format(CDate(.DateEnd), "mm/dd/yyyy") & "#"
1770                strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
                    ' ** Balance table, by specified FormRef('EndDate').
1780                Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_09_01")
1790                qdf.SQL = strTmp01
1800                Set qdf = Nothing

                    ' ** This one doesn't include LedgerArchive.
                    ' ** qryStatementParameters_Trans_12 -> qryStatementParameters_Trans_12_01
                    'strTmp01 = "SELECT DISTINCTROW ledger.journalno, account.shortname, account.legalname, ledger.accountno, " & _
                    '  "ledger.transdate, ledger.journaltype, ledger.assetdate, ledger.shareface, masterasset.due, masterasset.rate, " & _
                    '  "ledger.pershare, ledger.icash, ledger.pcash, ledger.cost, ledger.posted, masterasset.description, " & _
                    '  "ledger.description AS jcomment, masterasset.assetno, masterasset.yield, Balance.icash AS PreviousIcash, " & _
                    '  "Balance.pcash AS PreviousPcash, Balance.cost AS PreviousCost, journaltype.sortOrder, ledger.RecurringItem, " & _
                    '  "CoInfoGet('gstrCo_Name') AS CompanyName, CoInfoGet('gstrCo_Address1') AS CompanyAddress1, CoInfoGet('gstrCo_Address2') " & _
                    '  "AS CompanyAddress2, CoInfoGet('gstrCo_City') AS CompanyCity, CoInfoGet('gstrCo_State') AS CompanyState, " & _
                    '  "CoInfoGet('gstrCo_Zip') AS CompanyZip, CoInfoGet('gstrCo_Phone') AS CompanyPhone, CoInfoGet('gstrCo_Country') AS " & _
                    '  "CompanyCountry, CoInfoGet('gstrCo_PostalCode') AS CompanyPostalCode, ledger.curr_id, " & _
                    '  "IIf(IsNull([ledger].[PurchaseDate])=True,Null," & _
                    '  "CDate(Format([ledger].[PurchaseDate]," & Chr(34) & "mm/dd/yyyy" & Chr(34) & "))) AS PurchaseDate" & vbCrLf
                    'strTmp01 = strTmp01 & "FROM ((account LEFT JOIN qryStatementParameters_Trans_09_01 ON account.accountno = " & _
                    '  "qryStatementParameters_Trans_09_01.accountno) LEFT JOIN Balance ON (qryStatementParameters_Trans_09_01." & _
                    '  "[MaxOfbalance date] = Balance.[balance date]) AND (qryStatementParameters_Trans_09_01.accountno = Balance." & _
                    '  "accountno)) INNER JOIN ((ledger LEFT JOIN masterasset ON ledger.assetno = masterasset.assetno) LEFT JOIN " & _
                    '  "journaltype ON ledger.journaltype = journaltype.journaltype) ON account.accountno = ledger.accountno" & vbCrLf
                    'strTmp01 = strTmp01 & "WHERE (((ledger.accountno)=FormRef('accountno')) AND " & _
                    '  "((ledger.transdate)>=DateAdd('d',1,[qryStatementParameters_Trans_09_01].[MaxOfbalance date]) And " & _
                    '  "(ledger.transdate)<=FormRef('EndDate')) AND ((ledger.ledger_HIDDEN)=False));"
                    'strTmp02 = "FormRef('accountno')"
                    'strTmp03 = "'" & .cmbAccounts & "'"
                    'strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
                    'strTmp02 = "FormRef('EndDate')"
                    'strTmp03 = "#" & Format(CDate(.DateEnd), "mm/dd/yyyy") & "#"
                    'strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.

                    ' ** This one includes LedgerArchive.
                    'strTmp01 = "SELECT DISTINCTROW qryStatementParameters_Trans_12_04.journalno, account.shortname, " & _
                    '  "account.legalname, qryStatementParameters_Trans_12_04.accountno, qryStatementParameters_Trans_12_04.transdate, " & _
                    '  "qryStatementParameters_Trans_12_04.journaltype, qryStatementParameters_Trans_12_04.assetdate, " & _
                    '  "qryStatementParameters_Trans_12_04.shareface, masterasset.due, masterasset.rate, " & _
                    '  "qryStatementParameters_Trans_12_04.pershare, qryStatementParameters_Trans_12_04.icash, " & _
                    '  "qryStatementParameters_Trans_12_04.pcash, qryStatementParameters_Trans_12_04.cost, " & _
                    '  "qryStatementParameters_Trans_12_04.posted, masterasset.description, qryStatementParameters_Trans_12_04.jcomment, " & _
                    '  "masterasset.assetno, masterasset.yield, Balance.icash AS PreviousIcash, Balance.pcash AS PreviousPcash, " & _
                    '  "Balance.cost AS PreviousCost, journaltype.sortOrder, qryStatementParameters_Trans_12_04.RecurringItem, " & _
                    '  "CoInfoGet('gstrCo_Name') AS CompanyName, CoInfoGet('gstrCo_Address1') AS CompanyAddress1, CoInfoGet('gstrCo_Address2') " & _
                    '  "AS CompanyAddress2, CoInfoGet('gstrCo_City') AS CompanyCity, CoInfoGet('gstrCo_State') AS CompanyState, " & _
                    '  "CoInfoGet('gstrCo_Zip') AS CompanyZip, CoInfoGet('gstrCo_Phone') AS CompanyPhone, CoInfoGet('gstrCo_Country') AS " & _
                    '  "CompanyCountry, CoInfoGet('gstrCo_PostalCode') AS CompanyPostalCode, qryStatementParameters_Trans_12_04.curr_id, " & _
                    '  "IIf(IsNull([qryStatementParameters_Trans_12_04].[PurchaseDate])=True,Null," & _
                    '  "CDate(Format([qryStatementParameters_Trans_12_04].[PurchaseDate],'mm/dd/yyyy'))) AS PurchaseDate" & vbCrLf
                    'strTmp01 = strTmp01 & "FROM ((account LEFT JOIN qryStatementParameters_Trans_09_01 ON " & _
                    '  "account.accountno = qryStatementParameters_Trans_09_01.accountno) LEFT JOIN Balance ON " & _
                    '  "(qryStatementParameters_Trans_09_01.[MaxOfbalance date] = Balance.[balance date]) AND " & _
                    '  "(qryStatementParameters_Trans_09_01.accountno = Balance.accountno)) INNER JOIN " & _
                    '  "((qryStatementParameters_Trans_12_04 LEFT JOIN masterasset ON " & _
                    '  "qryStatementParameters_Trans_12_04.assetno = masterasset.assetno) LEFT JOIN journaltype ON " & _
                    '  "qryStatementParameters_Trans_12_04.journaltype = journaltype.journaltype) ON " & _
                    '  "account.accountno = qryStatementParameters_Trans_12_04.accountno" & vbCrLf
                    'strTmp01 = strTmp01 & "WHERE (((qryStatementParameters_Trans_12_04.accountno)='" & gstrAccountNo & "') AND " & _
                    '  "((qryStatementParameters_Trans_12_04.transdate)>=" & _
                    '  "DateAdd('d',1,[qryStatementParameters_Trans_09_01].[MaxOfbalance date]) And " & _
                    '  "(qryStatementParameters_Trans_12_04.transdate)<=#" & Format(gdatEndDate, "mm/dd/yyyy") & "#) AND " & _
                    '  "((qryStatementParameters_Trans_12_04.ledger_HIDDEN)=False));"

                    ' ** qryStatementParameters_Trans_12_04 (Union of qryStatementParameters_Trans_12_02
                    ' ** (Ledger, just needed fields, by specified GlobalVarGet("gstrAccountNo","gdatStartDate",
                    ' ** "gdatEndDate")), qryStatementParameters_Trans_12_03 (LedgerArchive, just needed fields,
                    ' ** by specified GlobalVarGet("gstrAccountNo","gdatStartDate","gdatEndDate"))), linked to
                    ' ** qryStatementParameters_Trans_09_01 (Balance table, by specified FormRef('EndDate')),
                    ' ** with add'l fields, by specified FormRef('accountno'), FormRef('EndDate').
1810                Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_12_05")
1820                strTmp01 = qdf.SQL
1830                Set qdf = Nothing
1840                Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_12_01")
1850                qdf.SQL = strTmp01
1860                Set qdf = Nothing
1870                strSQL = strTmp01

                    ' ************************************************************************************************************
                    ' ************************************************************************************************************
1880                dbs.Close
1890                Set dbs = Nothing
1900                strFileName = .cmbAccounts.Column(CBX_A_ACTNO) & " Transaction Rpts " & .TransDateStart & " through " & .TransDateEnd
1910              Case .opgAccountNumber_optAll.OptionValue
                    ' ** All accounts.
1920                Select Case .chkStatements
                    Case True
                      ' ** Only those scheduled.
1930                  Set dbs = CurrentDb

                      ' ** Ledger, linked to Account, qryStatementParameters_Trans_15b (qryStatementParameters_Trans_15a
                      ' ** (Account, with MonthNum = FormRef('MonthNum')), just those matching MonthNum),
                      ' ** qryStatementParameters_Trans_15c (Balance, grouped by accountno, with Max(balance date),
                      ' ** by specified FormRef('MaxBalDate')), by specified FormRef('EndDate'), FormRef('MonthNum').  #curr_id
                      'Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_15")
                      'strSQL = qdf.SQL
                      'Set qdf = Nothing
                      'THIS strSQL IS REPLACED BELOW!
                      ' ************************************************************************************************************
                      ' ************************************************************************************************************
                      ' ** This all new (11/28/2015) in order to fix that 'Unknown Jet error' Message.
                      ' ************************************************************************************************************
                      ' ************************************************************************************************************

                      ' ** This one now uses GlobalVarGet("glngMonthID").
                      ' ** qryStatementParameters_Trans_15a -> qryStatementParameters_Trans_15_01
                      'strTmp01 = "SELECT account.accountno, account.smtjan, account.smtfeb, account.smtmar, account.smtapr, account.smtmay, " & _
                      '  "account.smtjun, account.smtjul, account.smtaug, account.smtsep, account.smtoct, account.smtnov, account.smtdec, " & _
                      '  "FormRef('MonthNum') AS MonthNum" & vbCrLf & "FROM account;"
                      'strTmp02 = "FormRef('MonthNum')"
                      'strTmp03 = CStr(.cmbMonth.Column(CBX_MON_ID))
                      'strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
                      'Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_15_01")
                      'qdf.SQL = strTmp01
                      'Set qdf = Nothing

                      ' ** qryStatementParameters_Trans_15b -> qryStatementParameters_Trans_15_02
                      ' ** NO CHANGE!
                      'strTmp01 = "SELECT qryStatementParameters_Trans_15_01.accountno, qryStatementParameters_Trans_15_01.MonthNum" & vbCrLf & _
                      '  "FROM qryStatementParameters_Trans_15_01" & vbCrLf & _
                      '  "WHERE (((IIf(([MonthNum]=1 And [smtjan]=True) Or ([MonthNum]=2 And [smtfeb]=True) Or ([MonthNum]=3 And [smtmar]=True) Or " & _
                      '  "([MonthNum]=4 And [smtapr]=True) Or ([MonthNum]=5 And [smtmay]=True) Or ([MonthNum]=6 And [smtjun]=True) Or " & _
                      '  "([MonthNum]=7 And [smtjul]=True) Or ([MonthNum]=8 And [smtaug]=True) Or ([MonthNum]=9 And [smtsep]=True) Or " & _
                      '  "([MonthNum]=10 And [smtoct]=True) Or ([MonthNum]=11 And [smtnov]=True) Or ([MonthNum]=12 And [smtdec]=True),-1,0))=-1));"
                      'Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_15_02")
                      'qdf.SQL = strTmp01
                      'Set qdf = Nothing

                      ' ** qryStatementParameters_Trans_15c -> qryStatementParameters_Trans_15_03
1940                  strTmp01 = "SELECT Balance.accountno AS accountno, Max(Balance.[balance date]) AS MaxOfBalance_Date" & vbCrLf & _
                        "FROM Balance" & vbCrLf & _
                        "WHERE (((Balance.[balance date])<FormRef('MaxBalDate')))" & vbCrLf & _
                        "GROUP BY Balance.accountno;"
1950                  strTmp02 = "FormRef('MaxBalDate')"
1960                  strTmp03 = "#" & Format(CDate(DateAdd("y", -1, (DateAdd("m", 1, _
                        DateSerial(.StatementsYear, glngMonthID, 1))))), "mm/dd/yyyy") & "#"
1970                  strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
1980                  Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_15_03")
1990                  qdf.SQL = strTmp01
2000                  Set qdf = Nothing

                      ' ** This one doesn't include LedgerArchive.
                      ' ** qryStatementParameters_Trans_15 -> qryStatementParameters_Trans_15_05
                      'strTmp01 = "SELECT DISTINCTROW ledger.journalno, account.shortname, account.legalname, ledger.accountno, " & _
                      '  "ledger.transdate, ledger.journaltype, ledger.assetdate, ledger.shareface, masterasset.due, masterasset.rate, " & _
                      '  "ledger.pershare, ledger.icash, ledger.pcash, ledger.cost, ledger.posted, masterasset.description, " & _
                      '  "ledger.description AS jcomment, masterasset.assetno, masterasset.yield, Balance.icash AS PreviousIcash, " & _
                      '  "Balance.pcash AS PreviousPcash, Balance.cost AS PreviousCost, journaltype.sortOrder, ledger.RecurringItem, " & _
                      '  "CoInfoGet('gstrCo_Name') AS CompanyName, CoInfoGet('gstrCo_Address1') AS CompanyAddress1, CoInfoGet('gstrCo_Address2') " & _
                      '  "AS CompanyAddress2, CoInfoGet('gstrCo_City') AS CompanyCity, CoInfoGet('gstrCo_State') AS CompanyState, " & _
                      '  "CoInfoGet('gstrCo_Zip') AS CompanyZip, CoInfoGet('gstrCo_Phone') AS CompanyPhone, CoInfoGet('gstrCo_Country') AS " & _
                      '  "CompanyCountry, CoInfoGet('gstrCo_PostalCode') AS CompanyPostalCode, ledger.curr_id, " & _
                      '  "IIf(IsNull([ledger].[PurchaseDate])=True,Null," & _
                      '  "CDate(Format([ledger].[PurchaseDate]," & Chr(34) & "mm/dd/yyyy" & Chr(34) & "))) AS PurchaseDate" & vbCrLf
                      'strTmp01 = strTmp01 & "FROM (((account LEFT JOIN qryStatementParameters_Trans_15_03 ON account.accountno = " & _
                      '  "qryStatementParameters_Trans_15_03.accountno) LEFT JOIN Balance ON (qryStatementParameters_Trans_15_03." & _
                      '  "MaxOfBalance_Date = Balance.[balance date]) AND (qryStatementParameters_Trans_15_03.accountno = Balance." & _
                      '  "accountno)) INNER JOIN qryStatementParameters_Trans_15_02 ON account.accountno = qryStatementParameters_Trans_15_02." & _
                      '  "accountno) INNER JOIN ((ledger LEFT JOIN masterasset ON ledger.assetno = masterasset.assetno) LEFT JOIN journaltype " & _
                      '  "ON ledger.journaltype = journaltype.journaltype) ON account.accountno = ledger.accountno" & vbCrLf
                      'strTmp01 = strTmp01 & "WHERE (((ledger.transdate)>=DateAdd('y',1,[qryStatementParameters_Trans_15_03].[MaxOfBalance_Date]) " & _
                      '  "And (ledger.transdate)<=FormRef('EndDate')) AND ((ledger.ledger_HIDDEN)=False));"
                      'strTmp02 = "FormRef('EndDate')"
                      'strTmp03 = "#" & Format(CDate(.DateEnd), "mm/dd/yyyy") & "#"
                      'strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.

                      ' ** This one does include LedgerArchive.
                      ' ** Union of qryStatementParameters_Trans_15_05 (Ledger, linked to Account,
                      ' ** qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account,
                      ' ** by specified GlobalVarGet("glngMonthID")), just those matching MonthNum),
                      ' ** qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with Max(balance date),
                      ' ** by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                      ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account,
                      ' ** qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account, by
                      ' ** specified GlobalVarGet("glngMonthID")), just those matching MonthNum),
                      ' ** qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with Max(balance date),
                      ' ** by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")).
2010                  Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_15_07")
2020                  strTmp01 = qdf.SQL
2030                  Set qdf = Nothing
2040                  strSQL = strTmp01

                      ' ************************************************************************************************************
                      ' ************************************************************************************************************
                      ' ** qryStatementParameters_Trans_15c (Balance, grouped by accountno, with Max(balance date),
                      ' ** by specified FormRef('MaxBalDate')), linked to qryStatementParameters_Trans_15b
                      ' ** (qryStatementParameters_Trans_15a (Account, with MonthNum = FormRef('MonthNum')),
                      ' ** just those matching MonthNum), with FromDate.
                      'varTmp00 = DLookup("[FromDate]", "qryStatementParameters_Trans_15d")
                      ' ************************************************************************************************************
                      ' ** New!
                      ' ************************************************************************************************************

                      ' ** qryStatementParameters_Trans_15d -> qryStatementParameters_Trans_15_04
                      ' ** NO CHANGE!
                      'strTmp01 = "SELECT qryStatementParameters_Trans_15_03.accountno, qryStatementParameters_Trans_15_03.MaxOfBalance_Date, " & _
                      '  "DateAdd(" & Chr(34) & "y" & Chr(34) & ",1,[MaxOfBalance_Date]) AS FromDate" & vbCrLf & _
                      '  "FROM qryStatementParameters_Trans_15_03 INNER JOIN qryStatementParameters_Trans_15_02 ON " & _
                      '  "qryStatementParameters_Trans_15_03.accountno = qryStatementParameters_Trans_15_02.accountno;"
                      'Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_15_04")
                      'qdf.SQL = strTmp01
                      'Set qdf = Nothing
2050                  varTmp00 = DLookup("[FromDate]", "qryStatementParameters_Trans_15_04")

                      ' ************************************************************************************************************
                      ' ************************************************************************************************************
2060                  dbs.Close
2070                  Set dbs = Nothing
2080                  strFileName = "Scheduled Transaction Rpts " & Format(CDate(varTmp00), "mm/dd/yyyy") & " through " & _
                        Format(FormRef("MaxBalDate"), "mm/dd/yyyy")
2090                Case False
                      ' ** Statements not selected, so it's for Transactions.
2100                  Set dbs = CurrentDb

                      ' ** Ledger, linked to qryStatementParameters_Trans_09a (Balance table, by specified
                      ' ** FormRef('EndDate')), with add'l fields, by specified FormRef('EndDate').  #curr_id
                      'Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_13")  ' ** #curr_id
                      'strSQL = qdf.SQL
                      'THIS strSQL IS REPLACED BELOW!
                      ' ************************************************************************************************************
                      ' ************************************************************************************************************
                      ' ** This all new (11/28/2015) in order to fix that 'Unknown Jet error' Message.
                      ' ************************************************************************************************************
                      ' ************************************************************************************************************

                      ' ** qryStatementParameters_Trans_09a -> qryStatementParameters_Trans_09_01
2110                  strTmp01 = "SELECT Balance.accountno AS accountno, Max(Balance.[balance date]) AS [MaxOfbalance date]" & vbCrLf & _
                        "FROM Balance" & vbCrLf & _
                        "WHERE (((Balance.[balance date])<FormRef('EndDate')))" & vbCrLf & _
                        "GROUP BY Balance.accountno;"
2120                  strTmp02 = "FormRef('EndDate')"
2130                  strTmp03 = "#" & Format(CDate(.DateEnd), "mm/dd/yyyy") & "#"
2140                  strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
2150                  Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_09_01")
2160                  qdf.SQL = strTmp01
2170                  Set qdf = Nothing

                      ' ** This one doesn't include LedgerArchive.
                      ' ** qryStatementParameters_Trans_13 -> qryStatementParameters_Trans_13_01
                      'strTmp01 = "SELECT DISTINCTROW ledger.journalno, account.shortname, account.legalname, ledger.accountno, " & _
                      '  "ledger.transdate, ledger.journaltype, ledger.assetdate, ledger.shareface, masterasset.due, masterasset.rate, " & _
                      '  "ledger.pershare, ledger.icash, ledger.pcash, ledger.cost, ledger.posted, masterasset.description, " & _
                      '  "ledger.description AS jcomment, masterasset.assetno, masterasset.yield, Balance.icash AS PreviousIcash, " & _
                      '  "Balance.pcash AS PreviousPcash, Balance.cost AS PreviousCost, journaltype.sortOrder, ledger.RecurringItem, " & _
                      '  "CoInfoGet('gstrCo_Name') AS CompanyName, CoInfoGet('gstrCo_Address1') AS CompanyAddress1, CoInfoGet('gstrCo_Address2') " & _
                      '  "AS CompanyAddress2, CoInfoGet('gstrCo_City') AS CompanyCity, CoInfoGet('gstrCo_State') AS CompanyState, " & _
                      '  "CoInfoGet('gstrCo_Zip') AS CompanyZip, CoInfoGet('gstrCo_Phone') AS CompanyPhone, CoInfoGet('gstrCo_Country') AS " & _
                      '  "CompanyCountry, CoInfoGet('gstrCo_PostalCode') AS CompanyPostalCode, ledger.curr_id, " & _
                      '  "IIf(IsNull([ledger].[PurchaseDate])=True,Null," & _
                      '  "CDate(Format([ledger].[PurchaseDate]," & Chr(34) & "mm/dd/yyyy" & Chr(34) & "))) AS PurchaseDate" & vbCrLf
                      'strTmp01 = strTmp01 & "FROM ((account LEFT JOIN qryStatementParameters_Trans_09a ON account.accountno = " & _
                      '  "qryStatementParameters_Trans_09a.accountno) LEFT JOIN Balance ON (qryStatementParameters_Trans_09a.accountno = " & _
                      '  "Balance.accountno) AND (qryStatementParameters_Trans_09a.[MaxOfbalance date] = Balance.[balance date])) " & _
                      '  "INNER JOIN ((ledger LEFT JOIN masterasset ON ledger.assetno = masterasset.assetno) LEFT JOIN journaltype ON " & _
                      '  "ledger.journaltype = journaltype.journaltype) ON account.accountno = ledger.accountno" & vbCrLf
                      'strTmp01 = strTmp01 & "WHERE (((ledger.transdate)>=DateAdd('d',1,[qryStatementParameters_Trans_09a]." & _
                      '  "[MaxOfbalance date]) And (ledger.transdate)<=FormRef('EndDate')) AND ((ledger.ledger_HIDDEN)=False));"
                      'strTmp02 = "FormRef('EndDate')"
                      'strTmp03 = "#" & Format(CDate(.DateEnd), "mm/dd/yyyy") & "#"
                      'strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.

2180                  Select Case .chkArchiveOnly_Trans
                      Case True
                        ' ** LedgerArchive, linked to qryStatementParameters_Trans_09_02 (Balance table, by specified
                        ' ** FormRef('EndDateTrans')), with add'l fields, by specified GlobalVarGet("gdatEndDate").
2190                    Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_13_02")
2200                  Case False
2210                    Select Case .chkIncludeArchive_Trans
                        Case True
                          ' ** This one does include LedgerArchive.
                          ' ** Union of qryStatementParameters_Trans_13_01 (Ledger, linked to
                          ' ** qryStatementParameters_Trans_09_02 (Balance table, by specified
                          ' ** FormRef('EndDateTrans')), with add'l fields, by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_13_02 (LedgerArchive, linked to
                          ' ** qryStatementParameters_Trans_09_02 (Balance table, by specified
                          ' ** FormRef('EndDateTrans')), with add'l fields, by specified GlobalVarGet("gdatEndDate")).
2220                      Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_13_03")
2230                    Case False
                          ' ** Ledger, linked to qryStatementParameters_Trans_09_02 (Balance table, by specified
                          ' ** FormRef('EndDateTrans')), with add'l fields, by specified GlobalVarGet("gdatEndDate").
2240                      Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_13_01")
2250                    End Select
2260                  End Select
2270                  strTmp01 = qdf.SQL
2280                  Set qdf = Nothing
2290                  strSQL = strTmp01

                      ' ************************************************************************************************************
                      ' ************************************************************************************************************
2300                  dbs.Close
2310                  Set qdf = Nothing
2320                  Set dbs = Nothing
2330                  strFileName = "All Transaction Rpts " & .TransDateStart & " through " & .TransDateEnd
2340                End Select  ' ** chkStatements.
2350              End Select  ' ** opgAccountNumber.
2360            End Select  ' ** chkTransactions.
2370          End If  ' ** blnRetVal.

2380        End If  ' ** chkArchiveOnly_Trans.

2390      End If  ' ** blnRetVal.

2400      If blnRetVal = True Then

2410        Set dbs = CurrentDb
2420        Set rst = dbs.OpenRecordset(strSQL)
2430        If rst.BOF = True And rst.EOF = True Then
2440          blnRetVal = False
2450          If .chkStatements = False Or (.chkStatements = True And blnFromStmts = False) Then  ' ** Only show this message if NOT doing statements.
2460            DoCmd.Hourglass False
2470            MsgBox "There is no data for this report.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "Y09")
2480          End If
2490          rst.Close
2500          dbs.Close
2510        Else
2520          rst.MoveFirst
2530          rst.Close

2540          If .chkTransactions = True Then
2550            dbs.QueryDefs("qryStatementParameters_Trans_01_02").SQL = strSQL
2560            dbs.QueryDefs("qryStatementParameters_Trans_02").SQL = dbs.QueryDefs("qryStatementParameters_Trans_02_02").SQL
                ' ************************************************************************************************************
                ' ************************************************************************************************************
                ' ** This all new (11/28/2015) in order to fix that 'Unknown Jet error' Message.
                ' ************************************************************************************************************
                ' ************************************************************************************************************

                ' ** qryStatementParameters_Trans_09b -> qryStatementParameters_Trans_09_02
2570            strTmp01 = "SELECT Balance.accountno AS accountno, Max(Balance.[balance date]) AS [MaxOfbalance date]" & vbCrLf & _
                  "FROM Balance" & vbCrLf & _
                  "WHERE (((Balance.[balance date])<FormRef('EndDateTrans')))" & vbCrLf & _
                  "GROUP BY Balance.accountno;"
2580            strTmp02 = "FormRef('EndDateTrans')"  'IS THIS EVEN POPULATED FOR STATEMENTS?!
2590            Select Case IsNull(.TransDateEnd)
                Case True
2600              strTmp03 = "#" & Format(CDate(.DateEnd), "mm/dd/yyyy") & "#"
2610            Case False
2620              strTmp03 = "#" & Format(CDate(.TransDateEnd), "mm/dd/yyyy") & "#"
2630            End Select
2640            strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
2650            Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_09_02")
2660            qdf.SQL = strTmp01
2670            Set qdf = Nothing

                ' ** qryStatementParameters_Trans_01_02 -> qryStatementParameters_Trans_01_06
2680            Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_01_06")
2690            qdf.SQL = strSQL
2700            Set qdf = Nothing

                'qryStatementParameters_Trans_02_02 -> qryStatementParameters_Trans_02_07
                'qryStatementParameters_Trans_01_02 has gotten strSQL.

2710            dbs.QueryDefs("qryStatementParameters_Trans_02").SQL = dbs.QueryDefs("qryStatementParameters_Trans_02_07").SQL

                'qryStatementParameters_Trans_02 now has qryStatementParameters_Trans_02_07.
                'qryStatementParameters_Trans_01_04_01 -> qryStatementParameters_Trans_01_04_08
                'qryStatementParameters_Trans_01_04_02_01 -> qryStatementParameters_Trans_01_04_09
                'qryStatementParameters_Trans_01_04_02 -> qryStatementParameters_Trans_01_04_10
                'qryStatementParameters_Trans_01_04_03 -> qryStatementParameters_Trans_01_04_11
                'qryStatementParameters_Trans_01_04_04_01 -> qryStatementParameters_Trans_01_04_12
                'qryStatementParameters_Trans_01_04_04 -> qryStatementParameters_Trans_01_04_13
                'qryStatementParameters_Trans_01_04_05 -> qryStatementParameters_Trans_01_04_14
                'qryStatementParameters_Trans_01_04_06 -> qryStatementParameters_Trans_01_04_15
                'qryStatementParameters_Trans_01_04_07 -> qryStatementParameters_Trans_01_04_16
                'qryStatementParameters_Trans_02_04 -> qryStatementParameters_Trans_02_08

                ' ************************************************************************************************************
                ' ************************************************************************************************************
2720          ElseIf .chkStatements = True Then
2730            dbs.QueryDefs("qryStatementParameters_Trans_01_01").SQL = strSQL
2740            dbs.QueryDefs("qryStatementParameters_Trans_02").SQL = dbs.QueryDefs("qryStatementParameters_Trans_02_01").SQL
                ' ************************************************************************************************************
                ' ************************************************************************************************************
                ' ** This all new (11/28/2015) in order to fix that 'Unknown Jet error' Message.
                ' ************************************************************************************************************
                ' ************************************************************************************************************

                ' ** qryStatementParameters_Trans_09a -> qryStatementParameters_Trans_09_01
2750            strTmp01 = "SELECT Balance.accountno AS accountno, Max(Balance.[balance date]) AS [MaxOfbalance date]" & vbCrLf & _
                  "FROM Balance" & vbCrLf & _
                  "WHERE (((Balance.[balance date])<FormRef('EndDate')))" & vbCrLf & _
                  "GROUP BY Balance.accountno;"
2760            strTmp02 = "FormRef('EndDate')"
2770            strTmp03 = "#" & Format(CDate(.DateEnd), "mm/dd/yyyy") & "#"
2780            strTmp01 = StringReplace(strTmp01, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
2790            Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_09_01")
2800            qdf.SQL = strTmp01
2810            Set qdf = Nothing

                ' ** qryStatementParameters_Trans_01_01 -> qryStatementParameters_Trans_01_05
2820            Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_01_05")
2830            qdf.SQL = strSQL
2840            Set qdf = Nothing

                'qryStatementParameters_Trans_02_01 -> qryStatementParameters_Trans_02_05
                'qryStatementParameters_Trans_01_01 has gotten strSQL.

2850            dbs.QueryDefs("qryStatementParameters_Trans_02").SQL = dbs.QueryDefs("qryStatementParameters_Trans_02_05").SQL

                'qryStatementParameters_Trans_02 now has qryStatementParameters_Trans_02_05.
                'qryStatementParameters_Trans_01_03_01 -> qryStatementParameters_Trans_01_03_08
                'qryStatementParameters_Trans_01_03_02_01 -> qryStatementParameters_Trans_01_03_09
                'qryStatementParameters_Trans_01_03_02 -> qryStatementParameters_Trans_01_03_10
                'qryStatementParameters_Trans_01_03_03 -> qryStatementParameters_Trans_01_03_11
                'qryStatementParameters_Trans_01_03_04_01 -> qryStatementParameters_Trans_01_03_12
                'qryStatementParameters_Trans_01_03_04 -> qryStatementParameters_Trans_01_03_13
                'qryStatementParameters_Trans_01_03_05 -> qryStatementParameters_Trans_01_03_14
                'qryStatementParameters_Trans_01_03_06 -> qryStatementParameters_Trans_01_03_15
                'qryStatementParameters_Trans_01_03_07 -> qryStatementParameters_Trans_01_03_16
                'qryStatementParameters_Trans_02_03 -> qryStatementParameters_Trans_02_06

                ' ************************************************************************************************************
                ' ************************************************************************************************************
2860          End If
2870          dbs.Close

2880          blnIncludeCurrency = False
2890          If blnHasForEx = True Then
2900            blnTmp04 = True
2910            If .chkStatements = True Then
                  ' ** This only applies when dealing with scheduled accounts.
2920              If IsNull(.HasForeign_Sched) = True Then
                    ' ** Shouldn't be; it has a DefaultValue!
2930              Else
2940                If .HasForeign_Sched = vbNullString Then
                      ' ** Shouldn't be; it has a DefaultValue!
2950                Else
2960                  If .HasForeign_Sched = "NOT CHECKED" Then
                        ' ** Proceed below.
2970                  ElseIf .HasForeign_Sched = "SOME" Then
                        ' ** Proceed below.
2980                  ElseIf .HasForeign_Sched = "NONE" Then
                        ' ** No scheduled accounts have foreign currency.
2990                    blnTmp04 = False
3000                  End If
3010                End If
3020              End If
3030            End If  ' ** chkStatements.
3040            If blnTmp04 = True Then
3050              Select Case .opgAccountNumber
                  Case .opgAccountNumber_optAll.OptionValue
3060                blnIncludeCurrency = True
3070                .chkIncludeCurrency = True
3080                .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
3090              Case .opgAccountNumber_optSpecified.OptionValue
3100                blnHasForExThis = False
3110                gblnHasForExThis = False
3120                blnHasForExThis = HasForEx_SP(gstrAccountNo)  ' ** Module Function: modStatementParamFuncs1.
3130                gblnSwitchTo = False
3140                For lngX = 0& To (lngAcctFors - 1&)
3150                  If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
3160                    If arr_varAcctFor(F_JCNT, lngX) > 0 Then
3170                      gblnHasForExThis = True
3180                      blnHasForExThis = True
3190                    End If
3200                    gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)  ' ** True means supress foreign exchange columns.
3210                    If gblnHasForExThis = True And gblnSwitchTo = True Then
                          ' ** Don't suppress if now they do have foreign exchange transactions.
3220                      gblnSwitchTo = False
3230                    End If
3240                    Exit For
3250                  End If
3260                Next
3270                Select Case gblnSwitchTo
                    Case True
3280                  blnIncludeCurrency = False
3290                  .chkIncludeCurrency = False
3300                  .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
3310                Case False
3320                  Select Case .chkIncludeCurrency
                      Case True
3330                    blnIncludeCurrency = True
3340                  Case False
3350                    Select Case gblnHasForExThis
                        Case True
3360                      blnIncludeCurrency = True
3370                      .chkIncludeCurrency = True
3380                      .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
3390                    Case False
                          ' ** Ask if they want to suppress.
3400                      strDocName = "frmStatementParameters_ForEx"
3410                      gblnSetFocus = True
3420                      gblnMessage = True  ' ** False return means cancel.
3430                      gblnSwitchTo = True  ' ** False return means show ForEx, don't supress.
3440                      DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & gstrAccountNo
3450                      DoCmd.Hourglass True
3460                      DoEvents
3470                      Select Case gblnMessage
                          Case True
3480                        Select Case gblnSwitchTo
                            Case True
                              ' ** Leave chkIncludeCurrency = False.
3490                          blnIncludeCurrency = False
3500                        Case False
3510                          blnIncludeCurrency = True
3520                          .chkIncludeCurrency = True
3530                          .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
3540                        End Select
3550                      Case False
3560                        blnRetVal = False
3570                        DoCmd.Hourglass False
3580                      End Select
3590                    End Select
3600                  End Select
3610                End Select
3620                DoEvents
3630              End Select
3640            End If  ' ** blnTmp04.
3650          End If  ' ** blnHasForEx.

3660          If blnRetVal = True Then

3670            Select Case IsMissing(varOutput)
                Case True
3680              Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
3690                Select Case .opgOrderBy
                    Case .opgOrderBy_optDate.OptionValue
                      ' ** Order by date.
3700                  Select Case blnHasForEx
                      Case True
                        ' ** Foreign currencies are present somewhere.
3710                    Select Case blnIncludeCurrency
                        Case True
                          ' ** This acccount has foreign currencies.
3720                      strReportName = "rptTransaction_Statement_ForEx_SortDate"
3730                    Case False
                          ' ** User wants the standard layout.
3740                      strReportName = "rptTransaction_Statement_SortDate"
3750                    End Select
3760                  Case False
                        ' ** There are no foreign currencies whatsoever.
3770                    strReportName = "rptTransaction_Statement_SortDate"
3780                  End Select
3790                  DoCmd.OpenReport strReportName, acViewPreview
3800                Case .opgOrderBy_optType.OptionValue
                      ' ** Order by type.
3810                  Select Case blnHasForEx  'blnHasForeign
                      Case True
                        ' ** Foreign currencies are present somewhere.
3820                    Select Case blnIncludeCurrency
                        Case True
                          ' ** This acccount has foreign currencies.
3830                      strReportName = "rptTransaction_Statement_ForEx_SortType"
3840                    Case False
                          ' ** User wants the standard layout.
3850                      strReportName = "rptTransaction_Statement_SortType"
3860                    End Select
3870                  Case False
                        ' ** There are no foreign currencies whatsoever.
3880                    strReportName = "rptTransaction_Statement_SortType"
3890                  End Select
3900                  DoCmd.OpenReport strReportName, acViewPreview
3910                End Select  ' ** opgOrderBy.
3920              Case .opgAccountNumber_optAll.OptionValue
3930                Select Case .opgOrderBy
                    Case .opgOrderBy_optDate.OptionValue
                      ' ** Order by date.
3940                  Select Case blnIncludeCurrency
                      Case True
3950                    strReportName = "rptTransaction_Statement_ForEx_SortDate"
3960                  Case False
3970                    strReportName = "rptTransaction_Statement_SortDate"
3980                  End Select
3990                  DoCmd.OpenReport strReportName, acViewPreview
4000                Case .opgOrderBy_optType.OptionValue
                      ' ** Order by type.
4010                  Select Case blnIncludeCurrency
                      Case True
4020                    strReportName = "rptTransaction_Statement_ForEx_SortType"
4030                  Case False
4040                    strReportName = "rptTransaction_Statement_SortType"
4050                  End Select
4060                  DoCmd.OpenReport strReportName, acViewPreview
4070                End Select
4080              End Select

4090            Case False
                  ' ** Word/Excel.
4100            End Select  ' ** varOutput.

4110          End If  ' ** blnRetVal

4120        End If  ' ** EOF/BOF.

4130      End If  ' ** blnRetVal.

4140    End With

EXITP:
4150    Set rst = Nothing
4160    Set qdf = Nothing
4170    Set dbs = Nothing
4180    BuildTransactionInfo_SP = blnRetVal
4190    Exit Function

ERRH:
4200    blnRetVal = False
4210    DoCmd.Hourglass False
4220    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
4230      blnContinue = False
4240      If Reports.Count > 0 Then
4250        DoCmd.Close acReport, Reports(0).Name  ' ** Close report in preview.
4260      End If
4270    Case Else
4280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290    End Select
4300    Resume EXITP

End Function

Public Function BuildAssetListInfo_SP(frm As Access.Form, blnContinue As Boolean, datAssetListDate As Date, blnPrintAnnualStatement As Boolean, blnAllStatements As Boolean, blnNoDataAll As Boolean, blnRollbackNeeded As Boolean) As Boolean
' ** Called by:
' **  cmdAssetListPreview_Click()
' **  cmdAssetListPrint_Click
' **  cmdAssetListWord_Click
' **  cmdAssetListExcel_Click()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "BuildAssetListInfo_SP"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, qdf3 As DAO.QueryDef, qdf4 As DAO.QueryDef
        Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim rstWork As DAO.Recordset, rstTmpAssetList As DAO.Recordset, rstTmpAccountInfo As DAO.Recordset
        Dim rstAccount As DAO.Recordset, rstMasterAsset As DAO.Recordset
        Dim strSQL As String, strWorkSQL As String
        Dim strRelatedAccounts As String, strRelatedAccountsIN As String
        Dim strAccountNo As String, strWorkQry As String
        Dim blnNoAccount As Boolean, blnNoMAsset As Boolean
        Dim blnRollbackFailed As Boolean, blnNoData As Boolean
        Dim blnPriceHistory As Boolean, blnSkip As Boolean
        Dim dblCurr_Rate2 As Double, strCurr_Symbol As String
        Dim lngRecs As Long
        Dim intPos01 As Integer
        Dim lngTmpAccts As Long, arr_varTmpAcct As Variant
        Dim arr_varTmp00 As Variant, varTmp01 As Variant, dblTmp02 As Double, dblTmp03 As Double, dblTmp04 As Double
        Dim lngX As Long, lngY As Long
        Dim intRetVal_SetQrys_AList As Integer
        Dim blnRetVal As Boolean

        ' ** Array: arr_varTmp00().
        Const RV_ERR As Integer = 0
        Const RV_REL As Integer = 1
        Const RV_IN  As Integer = 2

4210    blnRetVal = True   ' ** Unless proven otherwise.
4220    blnContinue = True  ' ** Default.
4230    strRelatedAccountsIN = vbNullString   ' ** Default.
4240    blnNoAccount = False: blnPriceHistory = False
4250    gblnMessage = False

4260    With frm

4270      DoCmd.Hourglass True
4280      DoEvents

4290      If .cmbAccounts.Enabled = True Then
4300        If IsNull(.cmbAccounts) = True Then
4310          blnNoAccount = True
4320        Else
4330          If .cmbAccounts = vbNullString Then
4340            blnNoAccount = True
4350          Else
4360            strAccountNo = .cmbAccounts
4370          End If
4380        End If
4390        If blnNoAccount = True Then
4400          blnRetVal = False
4410          DoCmd.Hourglass False
4420          blnContinue = False
4430          MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "W01")
4440        End If
4450      End If

4460      If blnRetVal = True Then
4470        If IsNull(.DateEnd) = True And IsNull(.AssetListDate) = True Then
4480          blnRetVal = False
4490          DoCmd.Hourglass False
4500          blnContinue = False
4510          MsgBox "Must enter Period Ending date to continue.", vbInformation + vbOKOnly, (Left(("Entry Required" & Space(55)), 55) & "W02")
4520          .AssetListDate.SetFocus
4530        ElseIf IsNull(.DateEnd) = True Then
4540          .DateEnd = .AssetListDate
4550        End If
4560      End If

4570      If blnRetVal = True Then

4580        datAssetListDate = .AssetListDate
4590        blnIncludeCurrency = .chkIncludeCurrency
4600        .UsePriceHistory = False
4610        gdatEndDate = datAssetListDate
4620        .DateEnd = datAssetListDate
            ' ** qryStatementParameters_AssetList_74_45_01 (Balance, grouped by accountno, with Max(balance date),
            ' ** by specified GlobalVarGet("gdatEndDate")), grouped by balance_date, with cnt.
4630        varTmp01 = DLookup("[balance_date]", "qryStatementParameters_AssetList_74_45_02")
4640        If IsNull(varTmp01) = True Then
4650          gdatStartDate = #1/1/1900#
4660        Else
4670          gdatStartDate = varTmp01
4680        End If
4690        gdatStartDate = DateAdd("y", 1, gdatStartDate)  ' ** One day after last balance date.
4700        .DateStart = gdatStartDate
            'gdatStartDate
            'gdatEndDate
            'gdatMarketDate

4710        If .chkForeignExchange = True And blnIncludeCurrency = True Then
4720          .currentDate = Null
4730          blnPriceHistory = PricingHistory(datAssetListDate)  ' ** Module Function: modStatementParamFuncs2.
              ' ** blnPriceHistory indicates whether current pricing or pricing history should be used.
              ' ** IT HAS NOTHING TO DO WITH FOREIGN EXCHANGE!
4740          .UsePriceHistory = blnPriceHistory
4750        End If
4760        DoEvents

            ' ** This code will update the qryStatementParameters_AssetList_03 {qryMaxBalDates}
            ' ** query to give us the balance numbers from the previous statement.
4770        Select Case .opgAccountNumber.Value
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specified account.
4780          Select Case blnPrintAnnualStatement
              Case True
4790            blnRetVal = Test_AList_SP(gstrAccountNo)  ' ** Module Function: modStatementParamFuncs3.
                'intRetVal_SetQrys_AList = SetQrys_AList_SP(gstrAccountNo, Me)  ' ** Module Function: modStatementParamFuncs3.
4800            If blnRetVal = False Then
4810              blnRetVal = True
4820              intRetVal_SetQrys_AList = -9
4830            End If
4840          Case False
                ' ** SetQrys_AList() started throwing a 'Bad DLL calling convention'
                ' ** error, and I have no idea why.
4850            blnRetVal = Test_AList_SP(gstrAccountNo)  ' ** Module Function: modStatementParamFuncs3.
                'intRetVal_SetQrys_AList = SetQrys_AList_SP(strAccountNo, Me)  ' ** Module Function: modStatementParamFuncs3.
4860            If blnRetVal = False Then
4870              blnRetVal = True  ' ** We'll let the regular code change this.
4880              intRetVal_SetQrys_AList = -9
4890            End If
4900          End Select
              ' ** Return codes:
              ' **    0  Success.
              ' **    1  Success, with Archive.
              ' **    2  Success, Archive only.
              ' **   -2  No data.
              ' **   -4  Date criteria not met.
              ' **   -9  Error.
4910          If intRetVal_SetQrys_AList < 0 Then
4920            If .chkAssetList = False And blnAllStatements = False Then
4930              blnRetVal = False
4940              DoCmd.Hourglass False
4950              MsgBox "No Transactions.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "W03")
4960            ElseIf blnAllStatements = True Then
                  ' ** Remember, for statements, this is a list at beginning of period, not end.
4970            End If
4980          End If
4990        Case .opgAccountNumber_optAll.OptionValue
              ' ** All accounts.
5000          blnRetVal = Test_AList_SP(gstrAccountNo)  ' ** Module Function: modStatementParamFuncs3.
              'intRetVal_SetQrys_AList = SetQrys_AList_SP("All", Me)  ' ** Module Function: modStatementParamFuncs3.
5010          If blnRetVal = False Then
5020            blnRetVal = True
5030            intRetVal_SetQrys_AList = -9
5040          End If
              ' ** Return codes:
              ' **    0  Success.
              ' **    1  Success, with Archive.
              ' **    2  Success, Archive only.
              ' **   -2  No data.
              ' **   -4  Date criteria not met.
              ' **   -9  Error.
5050          If intRetVal_SetQrys_AList < 0 Then
5060            blnContinue = False
5070            Select Case blnAllStatements
                Case True
5080              blnNoData = True
5090              blnNoDataAll = True
5100            Case False
5110              DoCmd.Hourglass False
                  ' ** Since this is all under opgAccountNumber_optAll, we shouldn't be here!
5120            End Select
5130            blnRetVal = False
5140          End If
5150        End Select

5160      End If  ' ** blnRetVal.
5170      DoEvents

5180      If blnRetVal = True Then

5190        DoCmd.Hourglass True
5200        DoEvents

5210        Set dbs = CurrentDb

            ' ** Empty tmpAccountInfo.
5220        Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_09d")
5230        qdf1.Execute
5240        Set qdf1 = Nothing
            ' ** Empty tmpAccountInfo2.
5250        Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_53")
5260        qdf1.Execute
5270        Set qdf1 = Nothing
5280        DoEvents

5290        Select Case .opgAccountNumber.Value
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specified account.

5300          Select Case .chkRelatedAccounts
              Case True  ' ** Related Fields has been checked.
5310            arr_varTmp00 = SetRelatedAccts(frm)  ' ** Module Function: modStatementParamFuncs2.
5320            If arr_varTmp00(RV_ERR, 0) = vbNullString Then
5330              strRelatedAccounts = arr_varTmp00(RV_REL, 0)
5340              strRelatedAccountsIN = arr_varTmp00(RV_IN, 0)
                  ' ** Pricing History is only needed for foreign currencies, because it requires converting
                  ' ** older assets and transactions. Non-Foreign currency reports don't have those calculations.
5350              If .chkForeignExchange = True And blnIncludeCurrency = True Then
5360                Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** Append qryStatementParameters_AssetList_73_21 (xx) to tmpRelatedAccount_03.
5370                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_73_22")
5380                  With qdf1.Parameters
5390                    ![curdat] = datAssetListDate
5400                  End With
5410                Case False
                      ' ** Append qryStatementParameters_AssetList_73_01 (tmpRelatedAccount_01,
                      ' ** linked to Account, ActiveAssets, MasterAsset, AssetType, with
                      ' ** TotalCost1_usd, TotalMarket1_usd) to tmpRelatedAccount_03.
5420                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_73_02")
5430                End Select
5440              Else
                    ' ** Append qryStatementParameters_AssetList_23 (tmpRelatedAccount_01, linked to
                    ' ** Account, ActiveAssets, MasterAsset, AssetType.) to tmpRelatedAccount_02.
5450                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_24")
5460              End If
5470              qdf1.Execute
5480              Set qdf1 = Nothing
5490              If .chkForeignExchange = True And blnIncludeCurrency = True Then
5500                Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** tmpRelatedAccount_03, with qryStatementParameters_AssetList_70_31 (xx),
                      ' ** by specified [ractnos], with TotalCost_usd, TotalMarket_usd; Cartesian.
5510                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_32")
                      'With qdf1.Parameters
                      '  ![curdat] = datAssetListDate
                      'End With
5520                Case False
                      ' ** tmpRelatedAccount_03, with qryStatementParameters_AssetList_70_11
                      ' ** (qryStatementParameters_AssetList_70_10 (tmpRelatedAccount_03, with Foreign
                      ' ** Exchange, linked to Account, grouped and summed, by accountno), grouped and
                      ' ** summed), by specified [ractnos], with TotalCost_usd, TotalMarket_usd; Cartesian.
5530                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_12")
5540                End Select
5550              Else
                    ' ** tmpRelatedAccount_02, with qryStatementParameters_AssetList_26b (xx), by specified [ractnos]; Cartesian.
5560                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_27")
5570              End If
5580              strSQL = qdf1.SQL
5590              intPos01 = InStr(strSQL, "[ractnos]")
5600              strSQL = Left(strSQL, (intPos01 - 1)) & "'" & strRelatedAccounts & "'" & Mid(strSQL, (intPos01 + Len("[ractnos]")))
5610              intPos01 = InStr(strSQL, "[ractnos]")
5620              strSQL = Left(strSQL, (intPos01 - 1)) & "'" & strRelatedAccounts & "'" & Mid(strSQL, (intPos01 + Len("[ractnos]")))
5630              If .chkForeignExchange = True And blnIncludeCurrency = True Then
5640                Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** tmpRelatedAccount_03, with qryStatementParameters_AssetList_70_31 (xx),
                      ' ** from code, with [ractnos] replaced with actual accountno's, with
                      ' ** TotalCost_usd, TotalMarket_usd; Cartesian.
5650                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_33")
                      'With qdf1.Parameters
                      '  ![curdat] = datAssetListDate
                      'End With
5660                Case False
                      ' ** qryStatementParameters_AssetList_70_12 (tmpRelatedAccount_03, with
                      ' ** qryStatementParameters_AssetList_70_11 (qryStatementParameters_AssetList_70_10
                      ' ** (tmpRelatedAccount_03, with Foreign Exchange, linked to Account, grouped and summed,
                      ' ** by accountno), grouped and summed), by specified [ractnos], with TotalCost_usd,
                      ' ** TotalMarket_usd; Cartesian), with qryStatementParameters_AssetList_70_11
                      ' ** (qryStatementParameters_AssetList_70_10 (tmpRelatedAccount_03, with Foreign Exchange,
                      ' ** linked to Account, grouped and summed, by accountno), grouped and summed), from code, with
                      ' ** [ractnos] replaced with actual accountno's, with TotalCost_usd, TotalMarket_usd; Cartesian.
5670                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_13")
5680                End Select
5690                qdf1.SQL = strSQL
5700                Set qdf1 = Nothing
5710                Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** tmpRelatedAccount_03, with qryStatementParameters_AssetList_70_31 (xx),
                      ' ** from code, with [ractnos] replaced with actual accountno's, with
                      ' ** TotalCost_usd, TotalMarket_usd; Cartesian.
5720                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_33")
                      'With qdf1.Parameters
                      '  ![curdat] = datAssetListDate
                      'End With
5730                Case False
                      ' ** qryStatementParameters_AssetList_70_12 (tmpRelatedAccount_03, with
                      ' ** qryStatementParameters_AssetList_70_11 (qryStatementParameters_AssetList_70_10
                      ' ** (tmpRelatedAccount_03, with Foreign Exchange, linked to Account, grouped and summed,
                      ' ** by accountno), grouped and summed), by specified [ractnos], with TotalCost_usd,
                      ' ** TotalMarket_usd; Cartesian), with qryStatementParameters_AssetList_70_11
                      ' ** (qryStatementParameters_AssetList_70_10 (tmpRelatedAccount_03, with Foreign Exchange,
                      ' ** linked to Account, grouped and summed, by accountno), grouped and summed), from code, with
                      ' ** [ractnos] replaced with actual accountno's, with TotalCost_usd, TotalMarket_usd; Cartesian.
5740                  Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_13")
5750                End Select
5760              Else
                    ' ** qryStatementParameters_AssetList_27 (xx), from code, with [ractnos] replaced with actual accountno's.
5770                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_28")
5780                qdf1.SQL = strSQL
5790                Set qdf1 = Nothing
                    ' ** qryStatementParameters_AssetList_27 (xx), from code, with [ractnos] replaced with actual accountno's.
5800                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_28")
5810              End If
5820            End If  ' ** arr_varTmp00.
5830          Case False  ' ** Without Related Accounts.
5840            If .chkForeignExchange = True And blnIncludeCurrency = True Then
5850              Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' **
5860                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_23")
5870              Case False
                    ' ** Account, linked to ActiveAssets, grouped, with add'l fields, specified
                    ' ** by FormRef('accountno'), with TotalCost_usd, TotalMarket_usd.
5880                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_03")
5890              End Select
5900            Else
                  ' ** Account, linked to ActiveAssets, grouped, with add'l fields; specified FormRef('accountno').
5910              Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_07b")
5920            End If
5930          End Select  ' ** chkRelatedAccounts.
5940        Case .opgAccountNumber_optAll.OptionValue
5950          Select Case .chkStatements
              Case True  ' ** Qualified.
5960            If .chkForeignExchange = True And blnIncludeCurrency = True Then
5970              Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' **
5980                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_22")
5990              Case False
                    ' ** Account, linked to ActiveAssets, grouped, with add'l fields;
                    ' ** all accounts, qualified, with TotalCost_usd, TotalMarket_usd.
6000                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_02")
6010              End Select
6020            Else
                  ' ** Account, linked to ActiveAssets, grouped, with add'l fields; all accounts, qualified.
6030              Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_07aq")
6040            End If
6050          Case False  ' ** Unqualified.
6060            If .chkForeignExchange = True And blnIncludeCurrency = True Then
6070              Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' **
6080                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_21")
6090              Case False
                    ' ** Account, linked to ActiveAssets, grouped, with add'l fields;
                    ' ** all accounts, with TotalCost_usd, TotalMarket_usd.
6100                Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_01")
6110              End Select
6120            Else
                  ' ** Account, linked to ActiveAssets, grouped, with add'l fields; all accounts.
6130              Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_07a")
6140            End If
6150          End Select
6160        End Select

6170      End If  ' ** blnRetVal.
6180      DoEvents

6190      If blnRetVal = True Then

            'VGC: QDF1 HAS ACCOUNT LINKED TO ACTIVE ASSETS!

6200        Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specified account.
6210          Select Case .chkRelatedAccounts
              Case True
6220            If Len(strRelatedAccounts) = 0 Then
6230              blnContinue = False
6240            Else
                  ' ** Ledger, just needed fields, with shareface adjusted +/-; specified [datasof], [ractnos].
6250              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08c")  '#curr_id
6260              strSQL = qdf2.SQL
6270              intPos01 = InStr(strSQL, "[ractnos]")
6280              strSQL = Left(strSQL, (intPos01 - 1)) & strRelatedAccountsIN & Mid(strSQL, (intPos01 + Len("[ractnos]")))
                  ' ** qryStatementParameters_AssetList_08c, from code, with [ractnos] replaced with actual accountno's.
6290              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08d")  '#curr_id
6300              qdf2.SQL = strSQL
6310              Set qdf2 = Nothing
6320              Select Case .chkIncludeArchive_Asset
                  Case True
                    ' ** LedgerArchive, just needed fields, with shareface adjusted +/-; specified [datasof], [ractnos].
6330                Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08i")  '#curr_id
6340                strSQL = qdf2.SQL
6350                intPos01 = InStr(strSQL, "[ractnos]")
6360                strSQL = Left(strSQL, (intPos01 - 1)) & strRelatedAccountsIN & Mid(strSQL, (intPos01 + Len("[ractnos]")))
                    ' ** qryStatementParameters_AssetList_08i, from code, with [ractnos] replaced with actual accountno's.
6370                Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08j")  '#curr_id
6380                qdf2.SQL = strSQL
6390                Set qdf2 = Nothing
                    ' ** Union of qryStatementParameters_AssetList_08d, qryStatementParameters_AssetList_08j.
6400                Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08k")  '#curr_id
6410                With qdf2.Parameters
6420                  ![datasof] = datAssetListDate
6430                End With
6440              Case False
6450                Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08d")  '#curr_id
6460                With qdf2.Parameters
6470                  ![datasof] = datAssetListDate
6480                End With
6490              End Select
6500            End If
6510          Case False
6520            Select Case .chkIncludeArchive_Asset
                Case True
                  ' ** Union of qryStatementParameters_AssetList_08b (Ledger, just needed fields, with shareface adjusted +/-;
                  ' ** specified [datasof], [actno]), qryStatementParameters_AssetList_08g (LedgerArchive, just needed fields,
                  ' ** with shareface adjusted +/-; specified [datasof], [actno]).
6530              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08h")  '#curr_id
6540              With qdf2.Parameters
6550                ![datasof] = datAssetListDate
6560                If blnPrintAnnualStatement = True Then
6570                  ![actno] = gstrAccountNo
6580                Else
6590                  ![actno] = strAccountNo
6600                End If
6610              End With
6620            Case False
                  ' ** Ledger, just needed fields, with shareface adjusted +/-; specified [datasof], [actno].
6630              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08b")  '#curr_id
6640              With qdf2.Parameters
6650                ![datasof] = datAssetListDate
6660                If blnPrintAnnualStatement = True Then
6670                  ![actno] = gstrAccountNo
6680                Else
6690                  ![actno] = strAccountNo
6700                End If
6710              End With
6720            End Select
6730          End Select
6740        Case .opgAccountNumber_optAll.OptionValue
6750          Select Case .chkIncludeArchive_Asset
              Case True
6760            Select Case .chkStatements
                Case True  ' ** Qualified.
                  ' ** Union of qryStatementParameters_AssetList_08aq (Ledger, just needed fields, with shareface adjusted +/-;
                  ' ** specified [datasof], all accounts, qualified), qryStatementParameters_AssetList_08eq (LedgerArchive, just needed fields,
                  ' ** with shareface adjusted +/-; specified [datasof], all accounts, qualified), qualified.
6770              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08fq")  '#curr_id
6780            Case False  ' ** Unqualified.
                  ' ** Union of qryStatementParameters_AssetList_08a (Ledger, just needed fields, with shareface adjusted +/-;
                  ' ** specified [datasof], all accounts), qryStatementParameters_AssetList_08e (LedgerArchive, just needed fields,
                  ' ** with shareface adjusted +/-; specified [datasof], all accounts).
6790              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08f")  '#curr_id
6800            End Select
6810            With qdf2.Parameters
6820              ![datasof] = datAssetListDate
6830            End With
6840          Case False
6850            Select Case .chkStatements
                Case True  ' ** Qualified.
                  ' ** Ledger, just needed fields, with shareface adjusted +/-; specified [datasof], all accounts, qualified.
6860              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08aq")  '#curr_id
6870            Case False  ' ** Unqualified.
                  ' ** Ledger, just needed fields, with shareface adjusted +/-; specified [datasof], all accounts.
6880              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_08a")  '#curr_id
6890            End Select
6900            With qdf2.Parameters
6910              ![datasof] = datAssetListDate
6920            End With
6930          End Select
6940        End Select

6950      End If  ' ** blnRetVal.
6960      DoEvents

6970      If blnRetVal = True Then

            'VGC: QDF2 HAS LEDGER WHERE ASSETNO > 0.

            'THIS NEXT PIECE JUST ESTABLISHES WHETHER A ROLLBACK IS NEEDED,
            'BUT USING blnContinue SEEMS TO PREVENT OTHER THINGS FROM
            'HAPPENING EVEN IF NO ROLLBACK IS NEEDED!!!
6980        blnRollbackNeeded = False: blnRollbackFailed = False

            ' ** qdf1 is Account info, qdf2 is Ledger info.
6990        If blnContinue = True Then
              ' ** strWorkSQL is just Ledger (or Ledger/LedgerArchive), by accountno, transdate.
7000          Set rstWork = qdf2.OpenRecordset
7010          If rstWork.BOF = True And rstWork.EOF = True Then
7020            blnRollbackNeeded = False
7030            rstWork.Close
7040            Set rstWork = Nothing
7050            Set rst1 = qdf1.OpenRecordset
7060          Else
7070            rstWork.MoveFirst
                ' ** VGC 12/07/2014: Even though later transactions may not be asset-
                ' ** related, they may still affect cash totals; so always rollback.
7080            blnRollbackNeeded = True
7090            strWorkQry = qdf2.Name
7100          End If
7110        End If
7120        DoEvents

7130        If blnRollbackNeeded = True Then  ' ** Yes, a rollback is necessary.

              'qryStatementParameters_AssetList_08a
              'qryStatementParameters_AssetList_08aq
              'qryStatementParameters_AssetList_08b
              'qryStatementParameters_AssetList_08c
              'qryStatementParameters_AssetList_08d
              'qryStatementParameters_AssetList_08e
              'qryStatementParameters_AssetList_08eq
              'qryStatementParameters_AssetList_08f
              'qryStatementParameters_AssetList_08fq
              'qryStatementParameters_AssetList_08h
              'qryStatementParameters_AssetList_08i
              'qryStatementParameters_AssetList_08j
              'qryStatementParameters_AssetList_08k

              ' ** This has current totals, without any rollback.
7140          Set rst1 = qdf1.OpenRecordset  ' ** This is what's used for appending to the tmpAssetList2 table and rstTmpAssetList.
7150          If .chkForeignExchange = True And blnIncludeCurrency = True Then
7160            blnRetVal = FillAListTmp_SP(dbs, rst1, "tmpAssetList5")  ' ** Module Function: modStatementParamFuncs3.
7170          Else
7180            blnRetVal = FillAListTmp_SP(dbs, rst1, "tmpAssetList2")  ' ** Module Function: modStatementParamFuncs3.
7190          End If
7200          rst1.Close
7210          Set rst1 = Nothing
7220          DoEvents

7230          If blnRetVal = True Then

7240            If .chkForeignExchange = True And blnIncludeCurrency = True Then
7250              Set rstTmpAssetList = dbs.OpenRecordset("tmpAssetList5", dbOpenDynaset, dbConsistent)
7260            Else
7270              Set rstTmpAssetList = dbs.OpenRecordset("tmpAssetList2", dbOpenDynaset, dbConsistent)
7280            End If

7290            If (.opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue) And (.chkRelatedAccounts = True) Then
                  ' ** Related.
                  ' ** qryStatementParameters_AssetList_28, just needed fields, Top 1; related accounts.
7300              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_10b")
7310              Set rstAccount = qdf2.OpenRecordset
7320            Else
7330              Select Case .chkStatements
                  Case True  ' ** Qualified.
                    ' ** Account, just needed fields; all accounts, qualified.
7340                Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_10aq")
7350              Case False  ' ** Unqualified.
                    ' ** Account, just needed fields; all accounts.
7360                Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_10a")
7370              End Select
7380              Set rstAccount = qdf2.OpenRecordset
7390            End If
7400            Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' **
                  'Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_79_05")
7410              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_79_09")
7420              With qdf2.Parameters
7430                ![curdat] = datAssetListDate
7440              End With
7450            Case False
                  ' ** MasterAsset, linked to AssetType; all assets.
7460              Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_01")
7470            End Select
7480            Set rstMasterAsset = qdf2.OpenRecordset()
7490            DoEvents

7500            If (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
                  ' ** If developer, only leave final Asset List open.
7510              If IsLoaded("rptAssetList", acReport) = True Then  ' ** Module Function: modFileUtilities.
7520                DoCmd.Close acReport, "rptAssetList"
7530              End If
7540            End If

7550          Else
                ' ** ERROR.
7560            blnRetVal = False
7570            blnContinue = False
7580            blnRollbackFailed = True
7590            rstWork.Close
7600            dbs.Close
7610            Set dbs = Nothing
7620            Set rstWork = Nothing
7630            DoCmd.Hourglass False
7640            MsgBox "Unable to create temporary table for reporting.", vbCritical + vbOKOnly, _
                  (Left(("Error Creating Temporary Table" & Space(55)), 55) & "W05")
7650          End If

7660        End If  ' ** blnRollbackNeeded.

7670      End If  ' ** blnRetVal.
7680      DoEvents

7690      If blnRetVal = True And blnRollbackNeeded = True Then

            'VGC: rstACCOUNT HAS ACCOUNT TABLE!

            ' ** Queries behind qdf1 are:
7700        Select Case qdf1.Name
            Case "qryStatementParameters_AssetList_28"
              ' **   Specified, Related:
              ' ** qryStatementParameters_AssetList_28, grouped and summed, by accountno.
7710          Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_28s")
7720        Case "qryStatementParameters_AssetList_70_13", "qryStatementParameters_AssetList_70_33"  ' ** From .._28.
              ' **   Specified, Related, Foreign Exchange:
7730          Select Case blnPriceHistory
              Case True
                ' ** PRICING HISTORY!
                ' **
7740            Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_35")
                'With qdf1.Parameters
                '  ![curdat] = datAssetListDate
                'End With
7750          Case False
                ' ** qryStatementParameters_AssetList_70_13 (qryStatementParameters_AssetList_70_12
                ' ** (tmpRelatedAccount_03, with qryStatementParameters_AssetList_70_11
                ' ** (qryStatementParameters_AssetList_70_10 (tmpRelatedAccount_03, with Foreign Exchange,
                ' ** linked to Account, grouped and summed, by accountno), grouped and summed), by
                ' ** specified [ractnos], with TotalCost_usd, TotalMarket_usd; Cartesian), with
                ' ** qryStatementParameters_AssetList_70_11 (qryStatementParameters_AssetList_70_10
                ' ** (tmpRelatedAccount_03, with Foreign Exchange, linked to Account, grouped and summed,
                ' ** by accountno), grouped and summed), from code, with [ractnos] replaced with actual
                ' ** accountno's, with TotalCost_usd, TotalMarket_usd; Cartesian), grouped and summed, by accountno.
7760            Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_70_15")
7770          End Select
7780        Case "qryStatementParameters_AssetList_07b"
              ' **   Specified:
              ' ** qryStatementParameters_AssetList_07b, grouped and summed, by accountno.
7790          Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_07t")
7800        Case "qryStatementParameters_AssetList_74_03", "qryStatementParameters_AssetList_74_23"  ' ** From .._07b.
              ' **   Specified, Foreign Exchange:
7810          Select Case blnPriceHistory
              Case True
                ' ** PRICING HISTORY!
                ' **
7820            Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_27")
7830          Case False
                ' ** qryStatementParameters_AssetList_74_03 (Account, linked to ActiveAssets,
                ' ** grouped, with add'l fields, specified by FormRef('accountno'), with
                ' ** TotalCost_usd, TotalMarket_usd), grouped and summed, by accountno.
7840            Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_07")
7850          End Select
7860        Case "qryStatementParameters_AssetList_07a", "qryStatementParameters_AssetList_07aq"
              ' **   All:
7870          Select Case .chkStatements
              Case True  ' ** Qualified.
                ' ** qryStatementParameters_AssetList_07a, grouped and summed, by accountno, qalified.
7880            Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_07sq")
7890          Case False  ' ** Unqualified.
                ' ** qryStatementParameters_AssetList_07a, grouped and summed, by accountno.
7900            Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_07s")
7910          End Select
7920        Case "qryStatementParameters_AssetList_74_01", "qryStatementParameters_AssetList_74_02", _
                "qryStatementParameters_AssetList_74_21", "qryStatementParameters_AssetList_74_22"  ' ** From .._07s, .._07sq.
              ' **   All, Foreign Exchange:
7930          Select Case .chkStatements
              Case True
7940            Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' **
7950              Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_26")
7960            Case False
                  ' ** qryStatementParameters_AssetList_74_02 (Account, linked to ActiveAssets,
                  ' ** grouped, with add'l fields; all accounts, qualified, with TotalCost_usd,
                  ' ** TotalMarket_usd), grouped and summed, by accountno, qualified.
7970              Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_06")
7980            End Select
7990          Case False
8000            Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' **
8010              Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_25")
8020            Case False
                  ' ** qryStatementParameters_AssetList_74_01 (Account, linked to ActiveAssets,
                  ' ** grouped, with add'l fields; all accounts, with TotalCost_usd,
                  ' ** TotalMarket_usd), grouped and summed, by accountno.
8030              Set qdf1 = dbs.QueryDefs("qryStatementParameters_AssetList_74_05")
8040            End Select
8050          End Select
8060        End Select
8070        DoEvents

            ' ** Create table to contain temporary account master records to track "global" account info.
8080        Set rst1 = qdf1.OpenRecordset
8090        If .chkForeignExchange = True And blnIncludeCurrency = True Then
8100          blnRetVal = FillAListTmp_SP(dbs, rst1, "tmpAccountInfo2")  ' ** Module Function: modStatementParamFuncs3.
8110        Else
8120          blnRetVal = FillAListTmp_SP(dbs, rst1, "tmpAccountInfo")  ' ** Module Function: modStatementParamFuncs3.
8130        End If
8140        rst1.Close
8150        Set rst1 = Nothing
8160        Set qdf1 = Nothing
8170        DoEvents

8180        If blnRetVal = True Then

              'VGC: tmpACCOUNTINFO NOW HAS ACTIVE ASSETS, WITH ICASH/PCASH FROM ACCOUNT!

              ' ** Now, make records here distinct for each account.
8190          If .chkForeignExchange = True And blnIncludeCurrency = True Then
                ' ** Update tmpAccountInfo2, set fields to Null.
8200            Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_75_01")
8210            qdf2.Execute
8220            Set qdf2 = Nothing
                ' ** tmpAccountInfo2, all fields, DISTINCTROW.
8230            Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_75_02")
8240          Else
                ' ** Update tmpAccountInfo, set fields to Null.
8250            Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_11")
8260            qdf2.Execute
8270            Set qdf2 = Nothing
                ' ** tmpAccountInfo, all fields, DISTINCTROW.
8280            Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_12")
8290          End If
8300          Set rstTmpAccountInfo = qdf2.OpenRecordset
8310          With rstTmpAccountInfo
8320            If .BOF = True And .EOF = True Then
                  ' ** Unlikely.
8330              lngTmpAccts = 0&
8340            Else
8350              .MoveLast
8360              lngTmpAccts = .RecordCount
8370              .MoveFirst
8380              arr_varTmpAcct = .GetRows(lngTmpAccts)
8390            End If
8400            .Close
8410          End With
8420          Set rstTmpAccountInfo = Nothing
8430          Set qdf2 = Nothing
8440          DoEvents

              'VGC: ARRAY HAS ACTIVE ASSETS WITH ICASH/PCASH FROM ACCOUNT TABLE!

              ' ** Empty tmpAccountInfo.
8450          Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_09d")
8460          qdf2.Execute
8470          Set qdf2 = Nothing
              ' ** Empty tmpAccountInfo2.
8480          Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_70_53")
8490          qdf2.Execute
8500          Set qdf2 = Nothing
8510          DoEvents

8520          If .chkForeignExchange = True And blnIncludeCurrency = True Then
                ' ** tmpAccountInfo2, all fields, DISTINCTROW.
8530            Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_75_02")
8540          Else
                ' ** tmpAccountInfo, all fields, DISTINCTROW.
8550            Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_12")  ' ** Though it's empty here.
8560          End If
8570          Set rstTmpAccountInfo = qdf2.OpenRecordset
8580          With rstTmpAccountInfo
8590            For lngX = 0& To (lngTmpAccts - 1&)
8600              .AddNew
8610              For lngY = 0& To UBound(arr_varTmpAcct, 1)
8620                .Fields(lngY) = arr_varTmpAcct(lngY, lngX)
8630              Next
8640              .Update
8650            Next
8660          End With
8670          DoEvents

              'VGC: rstTMPACCOUNTINFO NOW HAS ARRAY TOTALS FROM ACTIVE ASSETS, WITH ICASH/PCASH FROM ACCOUNT TABLE!

8680          If lngTmpAccts = 0& Then 'CopyToTempTable(rstTmpAccountInfo, "tmpAccountInfo") = False Then  ' ** Module Function: modFileUtilities.
                ' ** ERROR.
8690            blnRetVal = False
8700            DoCmd.Hourglass False
8710            blnContinue = False
8720            blnRollbackFailed = True
8730            rstWork.Close
8740            rstAccount.Close
8750            rstMasterAsset.Close
8760            rstTmpAssetList.Close
8770            rstTmpAccountInfo.Close
8780            dbs.Close
8790            Set dbs = Nothing
8800            Set rstWork = Nothing
8810            Set rstAccount = Nothing
8820            Set rstMasterAsset = Nothing
8830            Set rstTmpAssetList = Nothing
8840            Set rstTmpAccountInfo = Nothing
8850            MsgBox "Unable to copy data to temporary table.", vbCritical + vbOKOnly, _
                  (Left(("Error Creating Temporary Table" & Space(55)), 55) & "W06")
8860          Else
8870            rstTmpAccountInfo.Close  ' ** Close old snapshot, then open as dynaset.
8880            Set rstTmpAccountInfo = Nothing
8890            If .chkForeignExchange = True And blnIncludeCurrency = True Then
8900              Set rstTmpAccountInfo = dbs.OpenRecordset("tmpAccountInfo2", dbOpenDynaset, dbConsistent)
8910            Else
8920              Set rstTmpAccountInfo = dbs.OpenRecordset("tmpAccountInfo", dbOpenDynaset, dbConsistent)
8930            End If
8940          End If
8950        Else
              ' ** ERROR.
8960          blnRetVal = False
8970          DoCmd.Hourglass False
8980          blnContinue = False
8990          blnRollbackFailed = True
9000          rstWork.Close
9010          rstAccount.Close
9020          rstMasterAsset.Close
9030          rstTmpAssetList.Close
9040          dbs.Close
9050          Set dbs = Nothing
9060          Set rstWork = Nothing
9070          Set rstAccount = Nothing
9080          Set rstMasterAsset = Nothing
9090          Set rstTmpAssetList = Nothing
9100          MsgBox "Unable to create temporary account table for reporting.", vbCritical + vbOKOnly, _
                (Left(("Error Creating Temporary Table" & Space(55)), 55) & "W07")
9110        End If

9120      End If  ' ** blnRetVal, blnRollbackNeeded.
9130      DoEvents

9140      If blnRetVal = True And blnRollbackNeeded = True Then

9150        rstTmpAccountInfo.MoveLast
9160        rstTmpAccountInfo.MoveFirst

            ' ** Get total number of records in recordset of subsequent ledger entries.
9170        rstWork.MoveLast
9180        lngRecs = rstWork.RecordCount
9190        rstWork.MoveFirst

            ' ** Still without rollbacks; here's where the rollbacks start.
9200        For lngX = 1& To lngRecs  ' ** Move through each record, tracking changes.

              ' ** THE WORK RST HAS ALL ASSET-RELATED RECORDS, INCLUDING DIVIDEND AND INTEREST!
              ' ** I CAN'T BELIEVE IT'S ALWAYS INCLUDED THOSE IN THE TOTALS! 12/19/2015
9210          Select Case rstWork![journaltype]
              Case "Deposit", "Purchase", "Withdrawn", "Sold", "Liability", "Cost Adj."

9220            dblCurr_Rate2 = 0#: strCurr_Symbol = vbNullString

9230            If (.opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue) And (.chkRelatedAccounts = True) Then
                  ' ** Related - there can be only one record.
9240              rstTmpAssetList.FindFirst " assetno = " & CStr(rstWork![assetno])
9250            Else
9260              rstTmpAssetList.FindFirst " accountno = '" & Trim(rstWork![accountno]) & "' AND assetno = " & CStr(rstWork![assetno])
9270            End If
9280            Select Case rstTmpAssetList.NoMatch
                Case True
                  ' ** We need a new record because this asset was not in the reporting query after subsequent changes.
9290              rstMasterAsset.FindFirst "assetno = " & CStr(rstWork![assetno])
9300              blnNoMAsset = rstMasterAsset.NoMatch
9310              If blnNoMAsset Then
9320                Select Case Trim(rstWork![journaltype])
                    Case "Misc.", "Paid", "Received"
                      ' ** OK to continue.
9330                Case Else
                      ' ** ERROR.
9340                  blnRetVal = False
9350                  DoCmd.Hourglass False
9360                  blnContinue = False
9370                  blnRollbackFailed = True
9380                  rstWork.Close
9390                  rstAccount.Close
9400                  rstMasterAsset.Close
9410                  rstTmpAssetList.Close
9420                  rstTmpAccountInfo.Close
9430                  dbs.Close
9440                  Set dbs = Nothing
9450                  Set rstWork = Nothing
9460                  Set rstAccount = Nothing
9470                  Set rstMasterAsset = Nothing
9480                  Set rstTmpAssetList = Nothing
9490                  Set rstTmpAccountInfo = Nothing
9500                  MsgBox "Missing Master Asset record.", vbCritical + vbOKOnly, (Left(("Asset Record Missing" & Space(55)), 55) & "W08")
9510                End Select
9520              Else
                    ' ** rstTmpAssetList is tmpAssetList2/tmpAssetList4.
                    ' ** HOW DOES MARKETVALUECURRENT GET INTO THIS LIST?!!!!  ##########################################
                    ' ** WITH A CARTESIAN QUERY CONTAINING JUST THE MARKETVALUE DATA!
9530                rstTmpAssetList.AddNew
9540                rstTmpAssetList![assetno] = rstMasterAsset![assetno]
9550                rstTmpAssetList![MasterAssetDescription] = rstMasterAsset![description]
9560                rstTmpAssetList![due] = rstMasterAsset![due]
9570                rstTmpAssetList![rate] = rstMasterAsset![rate]
9580                rstTmpAssetList![TotalCost] = 0  'rstTmpAssetList![TotalCost]            UPDATED BELOW!
9590                If .chkForeignExchange = True And blnIncludeCurrency = True Then
9600                  rstTmpAssetList![TotalCost_usd] = 0
9610                  rstTmpAssetList![TotalMarket_usd] = 0
9620                End If
9630                rstTmpAssetList![TotalShareFace] = 0  'rstTmpAssetList![TotalShareFace]  UPDATED BELOW!
9640                rstTmpAssetList![accountno] = IIf(Len(strRelatedAccounts) = 0, rstWork![accountno], strRelatedAccounts)
                    ' ** rstTmpAssetList![shortname]                                          UPDATED BELOW!
                    ' ** rstTmpAssetList![legalname]                                          UPDATED BELOW!
9650                rstTmpAssetList![assettype] = rstMasterAsset![assettype]
9660                rstTmpAssetList![assettype_description] = rstMasterAsset![assettype_description]
9670                rstTmpAssetList![totdesc] = CStr(rstMasterAsset![description]) & _
                      IIf(rstMasterAsset![rate] > 0, " " & Format(rstMasterAsset![rate], "0.000%"), vbNullString) & _
                      IIf(Not IsNull(rstMasterAsset![due]), "  Due " & Format(rstMasterAsset![due], "mm/dd/yyyy"), vbNullString)
9680                rstTmpAssetList![ICash] = 0  'rstTmpAssetList![ICash]                    UPDATED BELOW!
9690                rstTmpAssetList![PCash] = 0  'rstTmpAssetList![PCash]                    UPDATED BELOW!
9700                rstTmpAssetList![currentDate] = rstMasterAsset![currentDate]
9710                If (.chkForeignExchange = False Or (.chkForeignExchange = True And blnIncludeCurrency = False)) Then
9720                  rstTmpAssetList![CompanyName] = CoInfoGet("gstrCo_Name")  ' ** Module Function: modQueryFunctions2.
9730                  rstTmpAssetList![CompanyAddress1] = CoInfoGet("gstrCo_Address1")  ' ** Module Function: modQueryFunctions2.
9740                  rstTmpAssetList![CompanyAddress2] = CoInfoGet("gstrCo_Address2")  ' ** Module Function: modQueryFunctions2.
9750                  rstTmpAssetList![CompanyCity] = CoInfoGet("gstrCo_City")  ' ** Module Function: modQueryFunctions2.
9760                  rstTmpAssetList![CompanyState] = CoInfoGet("gstrCo_State")  ' ** Module Function: modQueryFunctions2.
9770                  rstTmpAssetList![CompanyZip] = CoInfoGet("gstrCo_Zip")  ' ** Module Function: modQueryFunctions2.
9780                  rstTmpAssetList![CompanyCountry] = CoInfoGet("gstrCo_Country")  ' ** Module Function: modQueryFunctions2.
9790                  rstTmpAssetList![CompanyPostalCode] = CoInfoGet("gstrCo_PostalCode")  ' ** Module Function: modQueryFunctions2.
9800                  rstTmpAssetList![CompanyPhone] = CoInfoGet("gstrCo_Phone")  ' ** Module Function: modQueryFunctions2.
9810                End If
9820                rstTmpAssetList![MarketValueX] = ZeroIfNull(rstMasterAsset![marketvalue])  ' ** Module Function: modStringFuncs.
9830                rstTmpAssetList![MarketValueCurrentX] = rstMasterAsset![marketvaluecurrent]
9840                rstTmpAssetList![YieldX] = rstMasterAsset![yield]
9850                DoEvents
9860                If .chkForeignExchange = True And blnIncludeCurrency = True Then
9870                  Select Case blnPriceHistory
                      Case True
                        ' ** CURRENCY HISTORY!
                        ' ** qryStatementParameters_AssetList_77_05 (Union of qryStatementParameters_AssetList_77_03
                        ' ** (qryStatementParameters_AssetList_77_02 (tblCurrency_History, all rates <= asset list date, by
                        ' ** specified [curdat]), grouped by curr_id, with Max(curr_date)), qryStatementParameters_AssetList_77_04
                        ' ** (tblCurrency, not in qryStatementParameters_AssetList_77_03 (qryStatementParameters_AssetList_77_02
                        ' ** (tblCurrency_History, all rates <= asset list date, by specified [curdat]), grouped by curr_id, with
                        ' ** Max(curr_date)))), linked back to tblCurrency_History, tblCurrency, tblCurrency_Symbol, by specified [currid].
9880                    Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_77_06")
9890                    With qdf3.Parameters
9900                      ![curdat] = datAssetListDate
9910                    End With
9920                  Case False
                        ' ** tblCurrency, linked to tblCurrency_Symbol, by specified [currid].
9930                    Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_77_01")
9940                  End Select
9950                  With qdf3.Parameters
9960                    ![currid] = rstMasterAsset![curr_id]
9970                  End With
9980                  Set rst1 = qdf3.OpenRecordset
9990                  With rst1
10000                   If .BOF = True And .EOF = True Then
                          ' ** Shouldn't happen.
10010                     dblCurr_Rate2 = 1#
10020                     strCurr_Symbol = "$"
10030                   Else
10040                     If rstMasterAsset![curr_id] = 150& Then  ' ** USD.
10050                       dblCurr_Rate2 = 1#
10060                       strCurr_Symbol = "$"
10070                     Else
10080                       dblCurr_Rate2 = ![curr_rate2]  ' ** This is the one that converts theirs to ours.
10090                       strCurr_Symbol = ![currsym_symbol]
10100                     End If
10110                   End If
10120                   .Close
10130                 End With
10140                 Set rst1 = Nothing
10150                 Set qdf3 = Nothing
10160                 rstTmpAssetList![curr_id] = rstMasterAsset![curr_id]
10170                 rstTmpAssetList![MarketValueCurrentX_usd] = (rstMasterAsset![marketvaluecurrent] * dblCurr_Rate2)
10180               End If
                    ' ** LEAVE in editing mode.
10190             End If
10200           Case False
10210             blnNoMAsset = False
10220             rstTmpAssetList.Edit  ' ** Edit existing temp record.
10230           End Select  ' ** NoMatch.
10240           DoEvents

10250           If blnRetVal = True Then

10260             If (.opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue) And _
                      (.chkRelatedAccounts = True) Then
                    ' ** Related - there can be only one record.
10270               rstTmpAccountInfo.FindFirst "accountno = '" & strRelatedAccounts & "'"
10280             Else
10290               rstTmpAccountInfo.FindFirst "accountno = '" & Trim(rstWork![accountno]) & "'"
10300             End If

10310             If rstTmpAccountInfo.NoMatch Then
                    ' ** We need a new account record.

10320               If (.opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue) And _
                        (.chkRelatedAccounts = True) Then
                      ' ** Related - there can be only one record.
10330                 rstTmpAccountInfo.AddNew
                      ' ** Copy info from Account table.
10340                 rstTmpAccountInfo![accountno] = strRelatedAccounts
10350                 rstTmpAccountInfo![shortname] = "Related Accounts"
10360                 rstTmpAccountInfo![legalname] = "Related Accounts"
10370                 rstTmpAccountInfo![ICash] = 0
10380                 rstTmpAccountInfo![PCash] = 0
                      ' ** LEAVE in editing mode.
10390               Else
10400                 rstAccount.FindFirst "accountno = '" & Trim(rstWork![accountno]) & "'"
10410                 If rstAccount.NoMatch Then
                        ' ** ERROR.
10420                   blnRetVal = False
10430                   DoCmd.Hourglass False
10440                   blnContinue = False
10450                   blnRollbackFailed = True
10460                   MsgBox "Missing Account master record.", vbCritical + vbOKOnly, (Left(("Asset Record Missing" & Space(55)), 55) & "W09")
10470                 Else
10480                   rstTmpAccountInfo.AddNew
                        ' ** Copy info from Account table.
10490                   rstTmpAccountInfo![accountno] = rstAccount![accountno]
10500                   rstTmpAccountInfo![shortname] = rstAccount![shortname]
10510                   rstTmpAccountInfo![legalname] = rstAccount![legalname]
10520                   rstTmpAccountInfo![ICash] = rstAccount![ICash]
10530                   rstTmpAccountInfo![PCash] = rstAccount![PCash]
                        ' ** LEAVE in editing mode.
10540                 End If
10550               End If
10560             Else
10570               rstTmpAccountInfo.Edit  ' ** Edit existing temp record.
10580             End If

                  'VGC: FOR NEW ACCOUNT RECORD, ICASH/PCASH COME FROM ACCOUNT TABLE!

10590           End If  ' ** blnRetVal.
10600           DoEvents

10610           If blnRetVal = True Then

10620             If (Not blnNoMAsset) And (Trim(rstWork![journaltype]) <> "Received") Then
                    ' ** Add/Subtract info from ASSET record. NOTE: these are SUBTRACTIONS because
                    ' ** the query returns the changes since the date.
10630               rstTmpAssetList![TotalShareFace] = _
                      rstTmpAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, _
                      rstWork![shareface] * IIf(rstTmpAssetList![assettype] = 90, 1, 1)))
10640               If .chkForeignExchange = True And blnIncludeCurrency = True Then
                      '###################################################################
                      '## Foreign Currency conversion calculations.
                      '###################################################################
10650                 If rstWork![curr_id] = 150& Then  ' ** 150 = USD.
10660                   rstTmpAssetList![TotalCost] = (rstTmpAssetList![TotalCost] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost]))
10670                   rstTmpAssetList![TotalCost_usd] = (rstTmpAssetList![TotalCost_usd] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost]))
                        ' ** MarketValueX should generally always be Null,
                        ' ** with only MarketValueCurrentX having the current unit value.
10680                   rstTmpAssetList![TotalMarket_usd] = (rstTmpAssetList![TotalMarket_usd] - _
                          (IIf(IsNull(rstWork![shareface]), 0, rstWork![shareface]) * rstTmpAssetList![MarketValueCurrentX]))
10690                 Else
                        ' ** TotalCost should be for this asset only, so it will be in the local currency.
10700                   rstTmpAssetList![TotalCost] = (rstTmpAssetList![TotalCost] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost]))
10710                   Select Case blnPriceHistory
                        Case True
                          ' ** CURRENCY HISTORY!
                          ' **
10720                     Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_77_06")
10730                     With qdf3.Parameters
10740                       ![curdat] = datAssetListDate
10750                     End With
10760                   Case False
                          ' ** tblCurrency, linked to tblCurrency_Symbol, by specified [currid].
10770                     Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_77_01")
10780                   End Select
10790                   With qdf3.Parameters
10800                     ![currid] = rstWork![curr_id]
10810                   End With
10820                   Set rst1 = qdf3.OpenRecordset
10830                   With rst1
10840                     If .BOF = True And .EOF = True Then
                            ' ** Shouldn't happen.
10850                       dblCurr_Rate2 = 1#
10860                       strCurr_Symbol = "$"
10870                     Else
10880                       dblCurr_Rate2 = ![curr_rate2]  ' ** This is the one that converts theirs to ours.
10890                       strCurr_Symbol = ![currsym_symbol]
10900                     End If
10910                     .Close
10920                   End With
10930                   Set rst1 = Nothing
10940                   Set qdf3 = Nothing
10950                   dblTmp02 = Round((IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost]) * dblCurr_Rate2), 2)
10960                   rstTmpAssetList![TotalCost_usd] = (rstTmpAssetList![TotalCost_usd] - dblTmp02)
                        ' ** MarketValueCurrentX is the unit price in the local currency.
10970                   dblTmp02 = (IIf(IsNull(rstWork![shareface]), 0, rstWork![shareface]) * rstTmpAssetList![MarketValueCurrentX])
10980                   dblTmp02 = Round((dblTmp02 * dblCurr_Rate2), 2)
10990                   rstTmpAssetList![TotalMarket_usd] = (rstTmpAssetList![TotalMarket_usd] - dblTmp02)
11000                 End If  ' ** curr_id.
                      '###################################################################
                      '###################################################################
11010               Else
11020                 rstTmpAssetList![TotalCost] = rstTmpAssetList![TotalCost] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost])
11030               End If
                    ' ** Save ASSET temp table record.
11040               rstTmpAssetList.Update
11050             End If
11060             DoEvents

                  'curr_id: 8  dblCurr_Rate2: 0.777
                  '  TotalCost_usd 1: 377.97  dblTmp02: 358.97  TotalCost_usd 2: 19
                  'curr_id: 27  dblCurr_Rate2: 0.7897
                  '  TotalCost_usd 1: 2195.37  dblTmp02: 2195.37  TotalCost_usd 2: 0
                  'curr_id: 27  dblCurr_Rate2: 0.7897
                  '  TotalCost_usd 1: 4295.97  dblTmp02: 4295.97  TotalCost_usd 2: 0
                  'curr_id: 52  dblCurr_Rate2: 1.5062
                  '  TotalCost_usd 1: 4364.2  dblTmp02: 4217.36  TotalCost_usd 2: 146.84
                  'curr_id: 109  dblCurr_Rate2: 0.1282
                  '  TotalCost_usd 1: 1000  dblTmp02: 1000  TotalCost_usd 2: 0

                  ' ** Add/Subtract info from ACCOUNT record.  NOTE: these are SUBTRACTIONS because
                  ' ** the query returns the changes since the date.
11070             If .chkForeignExchange = True And blnIncludeCurrency = True Then
                    '###################################################################
                    '## Foreign Currency conversion calculations.
                    '###################################################################
                    ' ** Individual transactions are always in the local currency,
                    ' ** and the Account totals will, at this time, always be in USD.
11080               If rstWork![curr_id] = 150& Then  ' ** 150 = USD.
11090                 rstTmpAccountInfo![ICash] = rstTmpAccountInfo![ICash] - rstWork![ICash]
11100                 rstTmpAccountInfo![PCash] = rstTmpAccountInfo![PCash] - rstWork![PCash]
11110               Else
11120                 If dblCurr_Rate2 = 0# Then
11130                   Select Case blnPriceHistory
                        Case True
                          ' ** CURRENCY HISTORY!
                          ' ** qryStatementParameters_AssetList_77_05 (Union of qryStatementParameters_AssetList_77_03
                          ' ** (qryStatementParameters_AssetList_77_02 (tblCurrency_History, all rates <= asset
                          ' ** list date, by specified [curdat]), grouped by curr_id, with Max(curr_date)),
                          ' ** qryStatementParameters_AssetList_77_04 (tblCurrency, not in
                          ' ** qryStatementParameters_AssetList_77_03 (qryStatementParameters_AssetList_77_02
                          ' ** (tblCurrency_History, all rates <= asset list date, by specified [curdat]), grouped
                          ' ** by curr_id, with Max(curr_date)))), linked back to tblCurrency_History, tblCurrency.
11140                     Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_77_06")
11150                     With qdf3.Parameters
11160                       ![curdat] = datAssetListDate
11170                     End With
11180                   Case False
                          ' ** tblCurrency, linked to tblCurrency_Symbol, by specified [currid].
11190                     Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_77_01")
11200                   End Select
11210                   With qdf3.Parameters
11220                     ![currid] = rstWork![curr_id]
11230                   End With
11240                   Set rst1 = qdf3.OpenRecordset
11250                   With rst1
11260                     If .BOF = True And .EOF = True Then
                            ' ** Shouldn't happen.
11270                       dblCurr_Rate2 = 1#
11280                       strCurr_Symbol = "$"
11290                     Else
11300                       dblCurr_Rate2 = ![curr_rate2]  ' ** This is the one that converts theirs to ours.
11310                       strCurr_Symbol = ![currsym_symbol]
11320                     End If
11330                     .Close
11340                   End With
11350                   Set rst1 = Nothing
11360                   Set qdf3 = Nothing
11370                 End If
11380                 dblTmp02 = Round((rstWork![ICash] * dblCurr_Rate2), 2)
11390                 rstTmpAccountInfo![ICash] = (rstTmpAccountInfo![ICash] - dblTmp02)
11400                 dblTmp02 = Round((rstWork![PCash] * dblCurr_Rate2), 2)
11410                 rstTmpAccountInfo![PCash] = (rstTmpAccountInfo![PCash] - dblTmp02)
11420               End If  ' ** curr_id.
                    '###################################################################
                    '###################################################################
11430             Else
                    ' ** rstWork is strWorkSQL, which is just Ledger, by accountno, transdate.
11440               rstTmpAccountInfo![ICash] = rstTmpAccountInfo![ICash] - rstWork![ICash]
11450               rstTmpAccountInfo![PCash] = rstTmpAccountInfo![PCash] - rstWork![PCash]
11460             End If

                  ' ** Save ACCOUNT temp table record.
11470             rstTmpAccountInfo.Update

11480           End If  ' ** blnRetVal.

11490         Case Else
                ' ** These shouldn't be in the totals!
11500         End Select  ' ** journaltype.
11510         DoEvents

11520         If blnRetVal = False Then
11530           Exit For
11540         ElseIf lngX < lngRecs Then
11550           rstWork.MoveNext
11560         End If

11570       Next  ' ** Move through each record, tracking changes.

11580       If blnRetVal = True Then
              ' ** Rollbacks can't include the Market Value, because each transaction
              ' ** may have had a different MarketValueCurrent, and the final total
              ' ** will end up really screwy!
              ' ** So, it should only be applied after all the rollbacks
              ' ** using the marketvaluecurrent as of the ending date.
11590         If .chkForeignExchange = True And blnIncludeCurrency = True Then
11600           With rstTmpAssetList
11610             .MoveLast
11620             lngRecs = .RecordCount
11630             .MoveFirst
11640             For lngX = 1& To lngRecs
11650               If Round(![TotalShareFace], 4) > 0 Then
11660                 .Edit
11670                 ![TotalMarket_usd] = CCur(Round((![TotalShareFace] * ![MarketValueCurrentX_usd]), 2))
11680                 .Update
11690               End If
11700               If lngX < lngRecs Then .MoveNext
11710             Next
11720           End With
11730         Else
                ' ** qryStatementParameters_35_03 (qryStatementParameters_35_02 (qryStatementParameters_35_01
                ' ** (tmpAssetList2, with currentDate_new, by specified GlobalVarGet("gdatEndDate")), linked to
                ' ** tblPricing_MasterAsset_History, with currentDate_new), grouped, with Max(currentDate_new)),
                ' ** linked back to tblPricing_MasterAsset_History, with marketvaluecurrentX_new, currentDate_new.
11740           Set qdf3 = dbs.QueryDefs("qryStatementParameters_35_04")
11750           Set rst1 = qdf3.OpenRecordset
11760           rst1.MoveFirst
11770           With rstTmpAssetList
11780             .MoveLast
11790             lngRecs = .RecordCount
11800             .MoveFirst
11810             For lngX = 1& To lngRecs
11820               If Round(![TotalShareFace], 4) > 0 Then
11830                 rst1.FindFirst "[assetno] = " & CStr(![assetno])
11840                 If rst1.NoMatch = False Then
11850                   .Edit
11860                   ![MarketValueCurrentX] = rst1![marketvaluecurrentX_new]
11870                   ![currentDate] = rst1![currentDate_new]
11880                   .Update
11890                 End If
11900               End If
11910               If lngX < lngRecs Then .MoveNext
11920             Next
11930           End With
11940           rst1.Close
11950           Set rst1 = Nothing
11960           Set qdf3 = Nothing
11970         End If
11980         DoEvents
11990         dblTmp02 = 0#: dblTmp03 = 0#: dblTmp04 = 0#
12000         Select Case strWorkQry
              Case "qryStatementParameters_AssetList_08a", "qryStatementParameters_AssetList_08aq", "qryStatementParameters_AssetList_08b", _
                  "qryStatementParameters_AssetList_08c", "qryStatementParameters_AssetList_08d", "qryStatementParameters_AssetList_08i", _
                  "qryStatementParameters_AssetList_08j"
12010           Set qdf3 = dbs.QueryDefs(strWorkQry)
12020           strWorkSQL = qdf3.SQL
12030           Set qdf3 = Nothing
12040           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12050           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12060           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08m")
12070           qdf3.SQL = strWorkSQL
12080           Set qdf3 = Nothing
12090           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08m")
12100         Case "qryStatementParameters_AssetList_08f"
12110           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08a")
12120           strWorkSQL = qdf3.SQL
12130           Set qdf3 = Nothing
12140           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12150           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12160           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08m")
12170           qdf3.SQL = strWorkSQL
12180           Set qdf3 = Nothing
12190           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08e")
12200           strWorkSQL = qdf3.SQL
12210           Set qdf3 = Nothing
12220           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12230           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12240           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08n")
12250           qdf3.SQL = strWorkSQL
12260           Set qdf3 = Nothing
12270           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08o")
12280         Case "qryStatementParameters_AssetList_08fq"
12290           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08aq")
12300           strWorkSQL = qdf3.SQL
12310           Set qdf3 = Nothing
12320           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12330           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12340           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08m")
12350           qdf3.SQL = strWorkSQL
12360           Set qdf3 = Nothing
12370           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08eq")
12380           strWorkSQL = qdf3.SQL
12390           Set qdf3 = Nothing
12400           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12410           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12420           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08n")
12430           qdf3.SQL = strWorkSQL
12440           Set qdf3 = Nothing
12450           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08o")
12460         Case "qryStatementParameters_AssetList_08h"
12470           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08b")
12480           strWorkSQL = qdf3.SQL
12490           Set qdf3 = Nothing
12500           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12510           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12520           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08m")
12530           qdf3.SQL = strWorkSQL
12540           Set qdf3 = Nothing
12550           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08g")
12560           strWorkSQL = qdf3.SQL
12570           Set qdf3 = Nothing
12580           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12590           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12600           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08n")
12610           qdf3.SQL = strWorkSQL
12620           Set qdf3 = Nothing
12630           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08o")
12640         Case "qryStatementParameters_AssetList_08k"
12650           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08d")
12660           strWorkSQL = qdf3.SQL
12670           Set qdf3 = Nothing
12680           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12690           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12700           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08m")
12710           qdf3.SQL = strWorkSQL
12720           Set qdf3 = Nothing
12730           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08j")
12740           strWorkSQL = qdf3.SQL
12750           Set qdf3 = Nothing
12760           strWorkSQL = StringReplace(strWorkSQL, "((ledger.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12770           strWorkSQL = StringReplace(strWorkSQL, "((LedgerArchive.assetno)>0) AND ", vbNullString)  ' ** Module Function: modStringFuncs.
12780           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08n")
12790           qdf3.SQL = strWorkSQL
12800           Set qdf3 = Nothing
12810           Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_08o")
12820         End Select  ' ** strWorkQry.
12830         lngY = qdf3.Parameters.Count
12840         If lngY > 0& Then
12850           For lngX = 0& To (lngY - 1&)
12860             If qdf3.Parameters(lngX).Name = "[datasof]" Then
12870               qdf3.Parameters(lngX) = datAssetListDate
12880             ElseIf qdf3.Parameters(lngX).Name = "[actno]" Then
12890               qdf3.Parameters(lngX) = strAccountNo
12900             End If
12910           Next
12920         End If
12930         DoEvents
12940         Set rst1 = qdf3.OpenRecordset
12950         Select Case qdf3.Name
              Case "qryStatementParameters_AssetList_08m"
                ' ** qryStatementParameters_AssetList_08q (qryStatementParameters_AssetList_08p
                ' ** (qryStatementParameters_AssetList_08m (xx), linked to tblCurrency_History,
                ' ** with curr_date), grouped, with Max(curr_date)), linked back to
                ' ** tblCurrency_History, with icash_usd, pcash_usd, cost_usd.
12960           Set qdf4 = dbs.QueryDefs("qryStatementParameters_AssetList_08r")
12970         Case "qryStatementParameters_AssetList_08o"
                ' ** qryStatementParameters_AssetList_08t (qryStatementParameters_AssetList_08s
                ' ** (qryStatementParameters_AssetList_08o (xx), linked to tblCurrency_History,
                ' ** with curr_date), grouped, with Max(curr_date)), linked back to
                ' ** tblCurrency_History, with icash_usd, pcash_usd, cost_usd.
12980           Set qdf4 = dbs.QueryDefs("qryStatementParameters_AssetList_08u")
12990         End Select
13000         lngY = qdf4.Parameters.Count
13010         If lngY > 0& Then
13020           For lngX = 0& To (lngY - 1&)
13030             If qdf4.Parameters(lngX).Name = "[datasof]" Then
13040               qdf4.Parameters(lngX) = datAssetListDate
13050             ElseIf qdf4.Parameters(lngX).Name = "[actno]" Then
13060               qdf4.Parameters(lngX) = strAccountNo
13070             End If
13080           Next
13090         End If
13100         Set rst2 = qdf4.OpenRecordset
13110         With rst1
13120           .MoveLast
13130           lngRecs = .RecordCount
13140           .MoveFirst
13150           For lngX = 1& To lngRecs
13160             If ![curr_id] = 150& Then
13170               dblTmp02 = dblTmp02 + ![ICash]
13180               dblTmp03 = dblTmp03 + ![PCash]
13190               dblTmp04 = dblTmp04 + ![Cost]
13200             Else
13210               rst2.FindFirst "[journalno] = " & CStr(![journalno])
13220               If rst2.NoMatch = False Then
13230                 dblTmp02 = dblTmp02 + rst2![icash_usd]
13240                 dblTmp03 = dblTmp03 + rst2![pcash_usd]
13250                 dblTmp04 = dblTmp04 + rst2![cost_usd]
13260               End If
13270             End If
13280             If lngX < lngRecs Then .MoveNext
13290           Next
13300           .Close
13310         End With
13320         rst2.Close
13330         Set rst1 = Nothing
13340         Set rst2 = Nothing
13350         Set qdf3 = Nothing
13360         Set qdf4 = Nothing
13370         DoEvents
13380         Set rst1 = dbs.OpenRecordset("account", dbOpenDynaset, dbReadOnly)
13390         With rst1
13400           .FindFirst "[accountno] = '" & rstWork![accountno] & "'"
13410           If .NoMatch = False Then
13420             dblTmp02 = ![ICash] - dblTmp02
13430             dblTmp03 = ![PCash] - dblTmp03
13440             dblTmp04 = ![Cost] - dblTmp04
13450           End If
13460           .Close
13470         End With
13480         Set rst1 = Nothing
              ' ** These numbers will later revert in the ForEx version!
13490         With rstTmpAccountInfo
13500           .MoveLast
13510           lngRecs = .RecordCount
13520           .MoveFirst
13530           For lngX = 1& To lngRecs
13540             If ![accountno] = strAccountNo Then
13550               .Edit
13560               ![ICash] = dblTmp02
13570               ![PCash] = dblTmp03
13580               .Update
13590             End If
13600             If lngX < lngRecs Then .MoveNext
13610           Next
13620         End With
              ' ** These numbers will later revert in the ForEx version!
13630         With rstTmpAssetList
13640           .MoveLast
13650           lngRecs = .RecordCount
13660           .MoveFirst
13670           For lngX = 1& To lngRecs
13680             .Edit
13690             ![ICash] = dblTmp02
13700             ![PCash] = dblTmp03
13710             .Update
13720             If lngX < lngRecs Then .MoveNext
13730           Next
13740         End With
13750       End If  ' ** blnRetVal.
13760       DoEvents

13770       If blnRetVal = False Then
13780         If IsNothing(rstWork) = False Then  ' ** Module Function: modUtilities.
13790           rstWork.Close
13800         End If
13810         If IsNothing(rstAccount) = False Then  ' ** Module Function: modUtilities.
13820           rstAccount.Close
13830         End If
13840         If IsNothing(rstMasterAsset) = False Then  ' ** Module Function: modUtilities.
13850           rstMasterAsset.Close
13860         End If
13870         If IsNothing(rstTmpAssetList) = False Then  ' ** Module Function: modUtilities.
13880           rstTmpAssetList.Close
13890         End If
13900         If IsNothing(rstTmpAccountInfo) = False Then  ' ** Module Function: modUtilities.
13910           rstTmpAccountInfo.Close
13920         End If
13930         If IsNothing(dbs) = False Then  ' ** Module Function: modUtilities.
13940           dbs.Close
13950         End If
13960         Set rstWork = Nothing
13970         Set rstAccount = Nothing
13980         Set rstMasterAsset = Nothing
13990         Set rstTmpAssetList = Nothing
14000         Set rstTmpAccountInfo = Nothing
14010       End If

14020     End If  ' ** blnRetVal, blnRollbackNeeded.
14030     DoEvents

14040     If blnRetVal = True And blnRollbackNeeded = True Then

            ' ** Remove any asset records which now have a ZERO totalshareface AND totalcost.
14050       If .chkForeignExchange = True And blnIncludeCurrency = True Then
              ' ** Delete tmpAssetList5, for Zero totalshareface, totalcost.
14060         Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_75_03")
14070       Else
              ' ** Delete tmpAssetList2, for Zero totalshareface, totalcost.
14080         Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_13")
14090       End If
14100       qdf2.Execute
14110       Set qdf2 = Nothing
14120       DoEvents

            'THIS DIDN'T WORK!  WHY?
14130       blnSkip = True
14140       If blnSkip = False Then
              ' ** Update tmpAccountInfo with the cash as of the asset list date.
14150         If .chkForeignExchange = True And blnIncludeCurrency = True Then
                ' ** Update qryStatementParameters_AssetList_75_05_09 (tmpAccountInfo2, with DLookups() to
                ' ** qryStatementParameters_AssetList_75_05_08 (qryStatementParameters_AssetList_75_05_04
                ' ** (qryStatementParameters_AssetList_75_05_03 (Union of qryStatementParameters_AssetList_75_05_01
                ' ** (Ledger, by GlobalVarGet("gdatEndDate")), qryStatementParameters_AssetList_75_05_02
                ' ** (LedgerArchive, by GlobalVarGet("gdatEndDate"))), grouped and summed, by accountno),
                ' ** linked to tmpAccountInfo2, with icash_new, pcash_new)).
14160           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_75_05_10")
14170         Else
                ' ** Update qryStatementParameters_AssetList_75_05_06 (tmpAccountInfo, with DLookups() to
                ' ** qryStatementParameters_AssetList_75_05_05 (qryStatementParameters_AssetList_75_05_04
                ' ** (qryStatementParameters_AssetList_75_05_03 (Union of qryStatementParameters_AssetList_75_05_01
                ' ** (Ledger, by GlobalVarGet("gdatEndDate")), qryStatementParameters_AssetList_75_05_02
                ' ** (LedgerArchive, by GlobalVarGet("gdatEndDate"))), grouped and summed, by accountno),
                ' ** linked to tmpAccountInfo, with icash_new, pcash_new)).
14180           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_75_05_07")
14190         End If
14200         qdf2.Execute
14210         Set qdf2 = Nothing
14220         DoEvents
14230       End If  ' ** blnSkip.

            ' ** Now, update tmpAssetList2 to reflect the needed account information
            ' ** which is global to the report.
14240       rstTmpAccountInfo.MoveLast
14250       lngRecs = rstTmpAccountInfo.RecordCount
14260       rstTmpAccountInfo.MoveFirst

            ' ** qryStatementParameters_AssetList_75_05_04 (qryStatementParameters_AssetList_75_05_03
            ' ** (Union of qryStatementParameters_AssetList_75_05_01 (Ledger, by GlobalVarGet("gdatEndDate")),
            ' ** qryStatementParameters_AssetList_75_05_02 (LedgerArchive, by GlobalVarGet("gdatEndDate"))),
            ' ** grouped and summed, by accountno), linked to Account.
14270       Set qdf3 = dbs.QueryDefs("qryStatementParameters_AssetList_75_05_11")
14280       Set rst1 = qdf3.OpenRecordset
14290       rst1.MoveLast
14300       lngRecs = rst1.RecordCount
14310       rst1.MoveFirst

14320       For lngX = 1& To lngRecs
14330         If .chkForeignExchange = True And blnIncludeCurrency = True Then
                ' ** Update tmpAssetList5, by specified [actno], [snam], [lnam], [icsh], [pcsh].
14340           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_75_04")
14350         Else
                ' ** Update tmpAssetList2, by specified [actno], [snam], [lnam], [icsh], [pcsh].
14360           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_14")
14370         End If
14380         With qdf2.Parameters
14390           ![actno] = rst1![accountno] 'rstTmpAccountInfo![accountno]
14400           ![snam] = rst1![shortname] 'rstTmpAccountInfo![shortname]
14410           ![lnam] = rst1![legalname] 'rstTmpAccountInfo![legalname]
14420           ![icsh] = rst1![ICash] 'rstTmpAccountInfo![ICash]
14430           ![pcsh] = rst1![PCash] 'rstTmpAccountInfo![PCash]
14440         End With
14450         qdf2.Execute
14460         If lngX < lngRecs Then rst1.MoveNext 'rstTmpAccountInfo.MoveNext
14470       Next
14480       Set qdf2 = Nothing
14490       rst1.Close
14500       Set rst1 = Nothing
14510       Set qdf3 = Nothing
14520       DoEvents

            ' ** Close temp & working recordsets.
14530       rstWork.Close
14540       rstAccount.Close
14550       rstMasterAsset.Close
14560       rstTmpAssetList.Close
14570       rstTmpAccountInfo.Close
14580       Set rstWork = Nothing
14590       Set rstAccount = Nothing
14600       Set rstMasterAsset = Nothing
14610       Set rstTmpAssetList = Nothing
14620       Set rstTmpAccountInfo = Nothing

            ' ** Finally, base report on this instead of on qryAssetList.
14630       If .chkForeignExchange = True And blnIncludeCurrency = True Then
14640         Select Case blnPriceHistory
              Case True
                ' ** PRICING HISTORY!
                ' **
14650           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_70_28")
                'With qdf2.Parameters
                '  ![curdat] = datAssetListDate
                'End With
14660         Case False
                ' ** qryStatementParameters_AssetList_70_07 (tmpAssetList5, all fields, with
                ' ** rollback, with TotalCost_usd, TotalMarket_usd), with Liability adjustment.
14670           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_70_08")
14680         End Select
14690       Else
              ' ** tmpAssetList2, all fields, with rollback.
14700         Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_15")
14710       End If
14720       Set rst1 = qdf2.OpenRecordset
14730       gblnMessage = True

14740     End If  ' ** blnRetVal, blnRollbackNeeded.
14750     DoEvents

14760     If blnRetVal = True And blnContinue = True Then
14770       If IsNothing(rst1) = True Then  ' ** Module Function: modUtilities.
              ' ** No rollbacks and no assets.
14780         If .chkForeignExchange = True And blnIncludeCurrency = True Then
                ' ** Append qryStatementParameters_AssetList_76_01 (Account, as tmpAssetList5
                ' ** record, for no assets, by specified FormRef('accountno')) to tmpAssetList5.
14790           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_76_02")
14800           qdf2.Execute
14810           Set qdf2 = Nothing
                ' ** Finally, base report on this instead of on qryAssetList.
14820           Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' **
14830             Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_70_28")
14840           Case False
                  ' ** qryStatementParameters_AssetList_70_07 (tmpAssetList5, all fields, with
                  ' ** rollback, with TotalCost_usd, TotalMarket_usd), with Liability adjustment.
14850             Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_70_08")
14860           End Select
14870           Set rst1 = qdf2.OpenRecordset
14880         Else
                'THIS DOESN'T GET HIT WHEN IT NEEDS TO!
                ' ** Append qryStatementParameters_AssetList_15c (Account,
                ' ** as tmpAssetList2 record, for no assets) to tmpAssetList2.
14890           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_15d")
14900           qdf2.Execute
14910           Set qdf2 = Nothing
                ' ** Finally, base report on this instead of on qryAssetList.
                ' ** tmpAssetList2, all fields, with rollback.
14920           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_15")
14930           Set rst1 = qdf2.OpenRecordset
14940         End If
              'WHY IS THIS SET TRUE IF NO ROLLBACKS ARE NEEDED?
14950         gblnMessage = True  ' ** Indicates rollbacks were needed.
14960       End If
14970     End If
14980     DoEvents

          ' ** Check for accounts with only cash transactions, and no assets.
14990     If blnRetVal = True And blnContinue = True Then
15000       blnNoData = False
15010       .chkNoAssets = False
15020       .chkNoAssets_lbl.FontBold = False
15030       .chkNoAssets_All = False
15040       .chkNoAssets_All_lbl.FontBold = False
15050       If rst1.BOF = True And rst1.EOF = True Then
15060         .chkNoAssets = True
15070         .chkNoAssets_lbl.FontBold = True
15080         rst1.Close
15090         Set rst1 = Nothing
15100         If .chkForeignExchange = True And blnIncludeCurrency = True Then
15110           Select Case blnPriceHistory
                Case True
                  ' ** qryStatementParameters_AssetList_74_37 (Union of qryStatementParameters_AssetList_74_33
                  ' ** (qryStatementParameters_AssetList_74_32 (qryStatementParameters_AssetList_74_31
                  ' ** (qryStatementParameters_AssetList_74_30 (Union of qryStatementParameters_AssetList_74_28
                  ' ** (Ledger, by specified GlobalVarGet("gstrAccountNo","gdatEndDate")),
                  ' ** qryStatementParameters_AssetList_74_29 (LedgerArchive, by specified GlobalVarGet("gstrAccountNo",
                  ' ** "gdatEndDate"))), linked to tblCurrency_History, just curr_date <= transdate), grouped, with
                  ' ** Max(curr_date)), linked back to tblCurrency_History, with icash_usd, pcash_usd, cost_usd),
                  ' ** qryStatementParameters_AssetList_74_36 (qryStatementParameters_AssetList_74_35
                  ' ** (qryStatementParameters_AssetList_74_34 (qryStatementParameters_AssetList_74_30 (Union of
                  ' ** qryStatementParameters_AssetList_74_28 (Ledger, by specified GlobalVarGet("gstrAccountNo",
                  ' ** "gdatEndDate")), qryStatementParameters_AssetList_74_29 (LedgerArchive, by specified
                  ' ** GlobalVarGet("gstrAccountNo","gdatEndDate"))), not in qryStatementParameters_AssetList_74_33
                  ' ** (qryStatementParameters_AssetList_74_32 (qryStatementParameters_AssetList_74_31
                  ' ** (qryStatementParameters_AssetList_74_30 (Union of qryStatementParameters_AssetList_74_28
                  ' ** (Ledger, by specified GlobalVarGet("gstrAccountNo","gdatEndDate")),
                  ' ** qryStatementParameters_AssetList_74_29 (LedgerArchive, by specified GlobalVarGet("gstrAccountNo",
                  ' ** "gdatEndDate"))), linked to tblCurrency_History, just curr_date <= transdate), grouped, with
                  ' ** Max(curr_date)), linked back to tblCurrency_History, with icash_usd, pcash_usd, cost_usd),
                  ' ** linked to tblCurrency_History, just curr_date > transdate), grouped, with Min(curr_date)),
                  ' ** linked back to tblCurrency_History, with icash_usd, pcash_usd, cost_usd)),
                  ' ** grouped and summed, by accountno.
15120             Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_74_40")
15130           Case False
                  ' ** qryStatementParameters_AssetList_74_30 (Union of qryStatementParameters_AssetList_74_28
                  ' ** (Ledger, by specified GlobalVarGet("gstrAccountNo","gdatEndDate")),
                  ' ** qryStatementParameters_AssetList_74_29 (LedgerArchive, by specified
                  ' ** GlobalVarGet("gstrAccountNo","gdatEndDate"))), grouped and summed, by accountno.
15140             Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_74_41")
15150           End Select
15160         Else
                ' ** qryStatementParameters_AssetList_74_30 (Union of qryStatementParameters_AssetList_74_28
                ' ** (Ledger, by specified GlobalVarGet("gstrAccountNo","gdatEndDate")),
                ' ** qryStatementParameters_AssetList_74_29 (LedgerArchive, by specified
                ' ** GlobalVarGet("gstrAccountNo","gdatEndDate"))), grouped and summed, by accountno.
15170           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_74_41")
15180         End If
15190         Set rst1 = qdf2.OpenRecordset
15200         If rst1.BOF = True And rst1.EOF = True Then
15210           blnNoData = True
15220         Else
15230           rst1.MoveFirst
15240           If Nz(rst1![ICash], 0) = 0 And Nz(rst1![PCash], 0) = 0 Then
15250             blnNoData = True
15260           End If
15270         End If
15280       Else
15290         If .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then
                ' ** qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a
                ' ** (Account, with MonthNum = FormRef('MonthNum')), just those matching MonthNum),
                ' ** not in qryStatementParameters_AssetList_74_50 (tmpAssetList5, grouped by accountno),
                ' ** accounts with no asset trans or no trans, period.
15300           Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_74_51")
15310           Set rst1 = qdf2.OpenRecordset
15320           If rst1.BOF = True And rst1.EOF = True Then
                  ' ** All pertinent accounts have asset transactions.
15330           Else
                  ' ** Some accounts have either no transactions at
                  ' ** all in the period, or no asset transactions.
15340             rst1.MoveFirst
15350             .chkNoAssets_All = True
15360             .chkNoAssets_All_lbl.FontBold = True
15370           End If
15380           rst1.Close
15390           Set rst1 = Nothing
15400           Set qdf2 = Nothing
15410         End If
15420       End If
15430       DoEvents

15440       Select Case blnNoData
            Case True
15450         Select Case blnAllStatements
              Case True
15460           blnNoDataAll = True
15470         Case False
15480           DoCmd.Hourglass False
15490           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "W10")
15500         End Select
15510         If IsNothing(rst1) = False Then  ' ** Module Function: modUtilities.
15520           rst1.Close
15530         End If
15540         dbs.Close
15550         Set rst1 = Nothing
15560         Set dbs = Nothing
15570         blnRetVal = False
15580       Case False
15590         If IsNothing(rst1) = False Then  ' ** Module Function: modUtilities.
15600           rst1.Close
15610         End If
              ' ** qryStatementParameters_AssetList_01, grouped, with Max(currentDate).
15620         Set qdf2 = dbs.QueryDefs("qryStatementParameters_AssetList_16")
15630         Set rst1 = qdf2.OpenRecordset
15640         With rst1
15650           If .BOF = True And .EOF = True Then
15660             gdatMarketDate = 0
15670           Else
15680             .MoveFirst
15690             gdatMarketDate = ![currentDate]  ' ** Used by report.
15700           End If
15710           .Close
15720         End With
15730         Set rst1 = Nothing
15740         dbs.Close
15750         Set dbs = Nothing
15760       End Select  ' ** blnNoData.
15770     End If  ' ** blnRetVal, blnContinue.
15780     DoEvents

          ' ** Balance, grouped, with BalDate_max, by specified FormRef('EndDate').
          'qryStatementParameters_AssetList_03

          ' ** If no rollbacks were needed, these are the queries:
          ' ** Account, linked to ActiveAssets, with add'l fields; all accounts.
          'qryStatementParameters_AssetList_06a
          ' ** Account, linked to ActiveAssets, with add'l fields; specified FormRef('accountno').
          'qryStatementParameters_AssetList_06b
          ' ** qryStatementParameters_AssetList_27, from code, with [ractnos] replaced with actual accountno's.
          'qryStatementParameters_AssetList_28  'I THINK!

          ' ** If rollbacks were needed, this is the query:
          ' ** tmpAssetList2, all fields, with rollback.
          'qryStatementParameters_AssetList_15

15790   End With  ' ** frm.

15800   DoCmd.Hourglass False

EXITP:
15810   Set rst1 = Nothing
15820   Set rst2 = Nothing
15830   Set qdf1 = Nothing
15840   Set qdf2 = Nothing
15850   Set qdf3 = Nothing
15860   Set qdf4 = Nothing
15870   Set dbs = Nothing
15880   BuildAssetListInfo_SP = blnRetVal
15890   Exit Function

ERRH:
100     DoCmd.Hourglass False
110     Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
120       blnContinue = False
130     Case Else
140       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
150     End Select
160     Resume EXITP

End Function

Public Function Statements_Print(frm As Access.Form, blnPrintStatements As Boolean, blnAllStatements As Boolean, blnSingleStatement As Boolean, blnRunPriorStatement As Boolean, blnAcctsSchedRpt As Boolean, datFirstDate As Date, blnContinue As Boolean, blnFromStmts As Boolean, blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, blnWasGTR As Boolean, Optional varAnnual As Variant) As Boolean
' ** Called by:
' **   cmdAnnualStatement_Click(), Above.
' **   cmdPrintStatement_All_Click(), Above.
' **   cmdPrintStatement_Single_Click(), Above.

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "Statements_Print"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim datStartDate As Date, datEndDate As Date
        Dim lngStatements As Long, lngStmtCnt As Long
        Dim strSQL As String
        Dim blnNoAccount As Boolean, blnAnnual As Boolean, blnPrintRpt As Boolean, blnPrintAll As Boolean, blnFound As Boolean
        Dim strMsg1 As String
        Dim strAccountNo As String
        Dim blnSwitchedOptGrp As Boolean, blnAccountsDisabled As Boolean
        Dim lngRecs As Long
        Dim varTmp00 As Variant
        Dim lngX As Long, intY As Integer, lngZ As Long
        Dim blnRetVal As Boolean

15910   blnRetVal = True
15920   blnContinue = True  ' ** Unless user cancels.
15930   lngStatements = 0&
15940   blnNoAccount = False: blnAccountsDisabled = False: blnPrintAll = False: blnFromStmts = False

15950   If IsMissing(varAnnual) = True Then
15960     blnAnnual = False
15970   Else
15980     blnAnnual = CBool(varAnnual)
15990   End If

16000   With frm

16010     If blnAnnual = False Then
16020       If IsNull(.cmbMonth) Then
16030         blnRetVal = False
16040         MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                (Left(("Entry Required" & Space(55)), 55) & "V01")
16050         .cmbMonth.SetFocus
16060       Else
16070         If .cmbMonth = vbNullString Then
16080           blnRetVal = False
16090           MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                  (Left(("Entry Required" & Space(55)), 55) & "V02")
16100           .cmbMonth.SetFocus
16110         Else
16120           glngMonthID = .cmbMonth.Column(CBX_MON_ID)
                'lngTmp01 = .cmbMonth  ' ** Just to make sure we get it!  RETURNS MONTH NAME!
16130           If .cmbAccounts.Enabled = True Then
16140             Select Case IsNull(.cmbAccounts)
                  Case True
16150               blnNoAccount = True
16160             Case False
16170               If .cmbAccounts = vbNullString Then
16180                 blnNoAccount = True
16190               End If
16200             End Select
16210             If blnNoAccount = True Then
16220               blnRetVal = False
16230               MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, _
                      (Left(("Nothing To Do" & Space(40)), 40) & "V03")
16240             End If
16250           End If
16260         End If
16270       End If
16280     End If  ' ** blnAnnual.

16290     Set dbs = CurrentDb

16300     If blnRetVal = True Then

            ' ** Disable this while it's running.
16310       If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
16320         .cmbAccounts.ForeColor = WIN_CLR_DISF
16330         .cmbAccounts.BackColor = WIN_CLR_DISB
16340         .cmbAccounts.BorderColor = WIN_CLR_DISR
16350         .cmbAccounts_lbl.ForeColor = WIN_CLR_DISF
16360         .cmbAccounts_lbl.BackStyle = acBackStyleTransparent
16370         .opgAccountSource.Enabled = False
16380         .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
16390         .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
16400         .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
16410         .opgAccountSource_optName_lbl2_dim_hi.Visible = True
16420         .chkRememberMe.Enabled = False
16430         .chkRememberMe_lbl2_dim.ForeColor = WIN_CLR_DISF
16440         .chkRememberMe_lbl2_dim_hi.Visible = True
16450         blnAccountsDisabled = True
16460       End If

16470       Select Case .cmbMonth.Column(CBX_MON_NAME)
            Case "January"
16480         .DateEnd = "01/31/" & .StatementsYear
16490       Case "February"
16500         .DateEnd = Format(CDate("03/01/" & .StatementsYear) - 1, "mm/dd/yyyy")  ' ** March 1st minus 1 day.
16510       Case "March"
16520         .DateEnd = "03/31/" & .StatementsYear
16530       Case "April"
16540         .DateEnd = "04/30/" & .StatementsYear
16550       Case "May"
16560         .DateEnd = "05/31/" & .StatementsYear
16570       Case "June"
16580         .DateEnd = "06/30/" & .StatementsYear
16590       Case "July"
16600         .DateEnd = "07/31/" & .StatementsYear
16610       Case "August"
16620         .DateEnd = "08/31/" & .StatementsYear
16630       Case "September"
16640         .DateEnd = "09/30/" & .StatementsYear
16650       Case "October"
16660         .DateEnd = "10/31/" & .StatementsYear
16670       Case "November"
16680         .DateEnd = "11/30/" & .StatementsYear
16690       Case "December"
16700         .DateEnd = "12/31/" & .StatementsYear
16710       End Select

16720       Select Case blnAnnual
            Case True
16730         .DateStart = "12/31/" & CStr(Val(.StatementsYear) - 1)
16740         datStartDate = .DateStart
16750       Case False
              ' ** .DateStart filled below.
16760       End Select
16770       datEndDate = CDate(.DateEnd)

16780       strMsg1 = vbNullString

            ' ** Note: The Annual Statement is sent as separate, single commands.
16790       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Single statement.
16800         Select Case blnAnnual
              Case True
16810           strAccountNo = gstrAccountNo
16820         Case False
16830           strAccountNo = .cmbAccounts
16840         End Select
16850         lngStatements = 1&
16860         Set rst1 = dbs.OpenRecordset("Statement Date", dbOpenDynaset, dbConsistent)
16870         If rst1.BOF = True And rst1.EOF = True Then
16880           rst1.AddNew
16890           rst1![Statement_Date] = #1/1/1900#
16900           rst1.Update
16910         End If
16920         rst1.MoveFirst  ' ** Must be exactly one record.
16930         Select Case blnRunPriorStatement  ' ** Set elsewhere.
              Case True
16940           If rst1![Statement_Date] = datEndDate Then
                  ' ** That's what we want to see.
16950             Set rst2 = CurrentDb.OpenRecordset("SELECT balance.* FROM balance WHERE accountno = '" & _
                    strAccountNo & "' ORDER BY [balance date] DESC;", dbOpenSnapshot)
16960             If rst2.BOF = True And rst2.EOF = True Then
                    ' ** This shouldn't ever happen, but what do we do if it does?
16970               rst2.Close
16980               Set rst2 = CurrentDb.OpenRecordset("Balance", dbOpenDynaset, dbAppendOnly)
16990               varTmp00 = DLookup("[predate]", "account", "[accountno] = '" & strAccountNo & "'")
17000               Select Case IsNull(varTmp00)
                    Case True
17010                 varTmp00 = DMin("[transdate]", "ledger", "[accountno] = '" & strAccountNo & "'")
17020                 Select Case IsNull(varTmp00)
                      Case True
                        ' ** No transactions!
17030                   blnRetVal = False
17040                   strMsg1 = "There are no transactions for this account"
17050                 Case False
17060                   With rst2
17070                     .AddNew
17080                     ![accountno] = strAccountNo
17090                     ![balance date] = (varTmp00 - 1)
17100                     ![ICash] = 0@
17110                     ![PCash] = 0@
17120                     ![Cost] = 0@
17130                     ![TotalMarketValue] = 0@
17140                     ![AccountValue] = 0@
17150                     .Update
17160                     .Bookmark = .LastModified
17170                   End With
17180                 End Select
17190               Case False
17200                 With rst2
17210                   .AddNew
17220                   ![accountno] = strAccountNo
17230                   ![balance date] = varTmp00
17240                   ![ICash] = 0@
17250                   ![PCash] = 0@
17260                   ![Cost] = 0@
17270                   ![TotalMarketValue] = 0@
17280                   ![AccountValue] = 0@
17290                   .Update
17300                   .Bookmark = .LastModified
17310                 End With
17320               End Select
17330             Else
17340               rst2.MoveFirst
17350             End If
17360             If blnRetVal = True Then
17370               If rst2![balance date] > datEndDate Then
17380                 strMsg1 = "This account already has a more recent statement."
17390               Else
17400                 If rst2![balance date] < datEndDate Then
                        ' ** This is what we want to see.
17410                 ElseIf rst2![balance date] = datEndDate Then
17420                   strMsg1 = "A statement has already been printed for the period ending " & _
                          Format(datEndDate, "mm/dd/yyyy") & " for account " & strAccountNo & "."
17430                 Else
                        ' ** There are no more options!
17440                   blnRetVal = False
17450                 End If
17460               End If
17470             End If  ' ** blnRetVal.
17480             rst2.Close
17490           Else
17500             If rst1![Statement_Date] > datEndDate Then
17510               strMsg1 = "The specified month is too old."
17520             ElseIf rst1![Statement_Date] < datEndDate Then
17530               strMsg1 = "A single future statement cannot be run."
17540             Else
                    ' ** There are no more options!
17550               blnRetVal = False
17560             End If
17570           End If
17580         Case False
17590           If rst1![Statement_Date] < datEndDate Then
17600             strMsg1 = "No statements have been printed for the period " & Format(datEndDate, "mm/dd/yyyy")
17610           Else
17620             Set rst2 = CurrentDb.OpenRecordset("SELECT balance.* FROM balance WHERE accountno = '" & _
                    strAccountNo & "' ORDER BY [balance date] DESC;", dbOpenSnapshot)
17630             With rst2
17640               If .BOF = True And .EOF = True Then
                      ' ** No balance records.
17650                 strMsg1 = "No statement has been printed for the period ending " & _
                        Format(datEndDate, "mm/dd/yyyy") & " for account " & strAccountNo & "."
17660               Else
17670                 .MoveFirst
17680                 If ![balance date] > datEndDate Then
17690                   strMsg1 = "This account already has a more recent statement."
17700                 Else
17710                   If ![balance date] < datEndDate Then
17720                     strMsg1 = "No statement has been printed for the period ending " & _
                            Format(datEndDate, "mm/dd/yyyy") & " for account " & strAccountNo & "."
17730                   End If
17740                 End If
17750               End If
17760               .Close
17770             End With  ' ** rst2.
17780             Set rst2 = Nothing
17790           End If
17800         End Select
17810         rst1.Close
17820         Set rst1 = Nothing

17830         If strMsg1 <> vbNullString Then
                ' ** Had some kind of problem.
17840           blnRetVal = False
17850           MsgBox strMsg1, vbExclamation + vbOKOnly, (Left(("Unable To Complete Operation" & Space(55)), 55) & "V04")
17860           strMsg1 = vbNullString
17870         Else
17880           Select Case blnAnnual
                Case True
17890             strMsg1 = "Do you want to print an Annual Statement for the period" & vbCrLf & _
                    Format(datStartDate, "mm/dd/yyyy") & " through " & Format(datEndDate, "mm/dd/yyyy")
17900             strMsg1 = strMsg1 & " for account " & strAccountNo & "?"  ' ** No vbCrLf.
17910           Case False
17920             strMsg1 = "Do you want to " & IIf(blnRunPriorStatement = True, vbNullString, "re") & _
                    "print the statement for the period ending " & Format(datEndDate, "mm/dd/yyyy")
17930             strMsg1 = strMsg1 & vbCrLf & "for account " & strAccountNo & "?"
17940           End Select
17950         End If

17960       Case .opgAccountNumber_optAll.OptionValue
              ' ** All for period.
17970         blnPrintAll = True
17980         strMsg1 = "Do you want to print statements for the period ending " & _
                Format(datEndDate, "mm/dd/yyyy") & "?"

17990       End Select

18000     End If  ' ** blnRetVal.

18010     If blnRetVal = True Then
18020       If SubsequentTransactionCheck(datEndDate) = False Then  ' ** Module Function: modStatementParamFuncs1.
18030         blnRetVal = False
18040       Else
18050         If .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then  ' ** All statements.
18060           Set rst1 = dbs.OpenRecordset("Statement Date", dbOpenDynaset, dbReadOnly)
18070           rst1.MoveFirst
18080           If rst1![Statement_Date] > datEndDate Then
18090             blnRetVal = False
18100             MsgBox "Period ending must be later or the same as " & Format(rst1![Statement_Date], "mm/dd/yyyy"), _
                    vbInformation + vbOKOnly, (Left(("Invalid Entry" & Space(40)), 40) & "V05")
18110           End If
18120           rst1.Close
18130         End If
18140       End If
18150     End If  ' ** blnRetVal.
          ' ** rst1 is closed.

18160     If blnRetVal = True Then
18170       Select Case blnAnnual
            Case True
18180         lngStatements = 1&
18190         strMsg1 = strMsg1 & " 1 Statement will be printed."
18200       Case False
18210         If blnSingleStatement = False Then
                ' ** Check this again.
18220           If glngMonthID = 0& Then
18230             glngMonthID = 12& 'lngTmp01
18240           End If
                ' ** Account, now with DateClosed = Null, by specified GlobalVarGet("glngMonthID").
18250           Set qdf2 = dbs.QueryDefs("qryStatementParameters_20")
18260           Set rst2 = qdf2.OpenRecordset
18270           If rst2.BOF = True And rst2.EOF = True Then
                  ' ** None scheduled!
18280             blnRetVal = False
18290             Select Case gblnGoToReport
                  Case True
18300               .TimerInterval = 0&
18310               .GoToReport_arw_sp_printall_img.Visible = False
18320               .cmdPrintStatement_Single.Visible = True
18330               blnGoingToReport2 = False
18340               blnGoingToReport = False
18350               gblnGoToReport = False
18360               blnGTR_Emblem = False
18370               .GTREmblem_Off  ' ** Procedure: Below.
18380               Beep
18390               DoCmd.Hourglass False
18400               MsgBox "Trust Accountant is unable to show the requested report." & vbCrLf & vbCrLf & _
                      "There is insufficient data to demonstrate.", vbInformation + vbOKOnly, "Report Location Unavailable"
18410             Case False
18420               MsgBox "There are no statements scheduled!", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "V06")
18430             End Select
18440             rst2.Close
18450           Else
18460             rst2.MoveLast
18470             lngStatements = rst2.RecordCount
18480             strMsg1 = strMsg1 & vbCrLf & CStr(lngStatements) & " Statement" & _
                    IIf(lngStatements > 1&, "s", vbNullString) & " will be printed."
18490             rst2.Close
18500           End If
18510           Set rst2 = Nothing
18520           Set qdf2 = Nothing
18530         End If  ' ** blnSingleStatement.
18540       End Select  ' ** blnAnnual.
18550     End If  ' ** blnRetVal.
          ' ** rst1 still closed.

18560     If blnRetVal = True Then
18570       Select Case blnAnnual
            Case True
18580         strMsg1 = "Do you want to print statements for the period ending " & _
                Format(datEndDate, "mm/dd/yyyy")
18590         strMsg1 = strMsg1 & vbCrLf & "for account " & strAccountNo & "?"
18600       Case False
18610         If (year(datEndDate) > year(Date)) Then
18620           blnRetVal = False
18630           MsgBox "Period ending must be before or the same as the current month", vbExclamation + vbOKOnly, _
                  (Left(("Invalid Entry" & Space(40)), 40) & "V07")
18640         Else
18650           If (year(datEndDate) = year(Date)) And (month(datEndDate) > month(Date)) Then
18660             blnRetVal = False
18670             MsgBox "Period ending must be before or the same as the current month", vbExclamation + vbOKOnly, _
                    (Left(("Invalid Entry" & Space(40)), 40) & "V08")
18680           End If
18690         End If
18700       End Select  ' ** blnAnnual.
18710     End If  ' ** blnRetVal.
          ' ** rst1 still closed.

18720     If blnRetVal = True Then
18730       If blnAnnual = True And .PrintAnnual_cnt > 1& Then
              ' ** Skip this extra window.
18740         blnPrintStatements = True
18750       Else
18760         blnPrintStatements = False: blnAcctsSchedRpt = False
18770         If gblnGoToReport = False Then
18780           DoCmd.Hourglass False
18790           Beep
18800         End If
              ' ** Confirm once more to assure user sees the date and is really, really ready to print statements.
18810         DoCmd.OpenForm "frmStatementParameters_Print", acNormal, , , acFormPropertySettings, acDialog, frm.Name & "~" & strMsg1
              ' ** This separate form sets blnPrintStatements via SetPrintStatements(), below.
              ' ** Because it's opened in acDialog mode, processing stops till the window is closed.
18820         DoCmd.Hourglass True
18830         DoEvents
18840         If blnPrintStatements = False Then
18850           blnRetVal = False
18860           DoCmd.Hourglass False
18870           .PrintAnnual_chk = False
18880         End If
18890       End If
18900     End If  ' ** blnRetVal.
          ' ** rst1 still closed.

          ' ** If they're not continuing, undo what was done above.
18910     If blnRetVal = False And blnAccountsDisabled = True Then
18920       .cmbAccounts.ForeColor = CLR_BLK
18930       .cmbAccounts.BackColor = CLR_WHT
18940       .cmbAccounts.BorderColor = CLR_LTBLU2
18950       .cmbAccounts_lbl.ForeColor = CLR_WHT
18960       .cmbAccounts_lbl.BackStyle = acBackStyleNormal
18970       .opgAccountSource.Enabled = True
18980       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
18990       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
19000       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
19010       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
19020       .chkRememberMe.Enabled = True
19030       .chkRememberMe_lbl2_dim.ForeColor = CLR_DKGRY2
19040       .chkRememberMe_lbl2_dim_hi.Visible = False
19050     End If

19060     If blnPrintStatements = True And blnRetVal = True Then

19070       gstrAccountNo = strAccountNo
19080       gdatStartDate = datStartDate
19090       gdatEndDate = datEndDate
19100       blnIncludeCurrency = .chkIncludeCurrency

19110       Select Case blnAnnual
            Case True
              ' ** qryStatementAnnual_04.
19120         strSQL = "SELECT ledger.accountno, Sum(IIf([ledger].[icash]<0,0,[ledger].[icash])) AS SumPositiveIcash, " & _
                "Sum(IIf([ledger].[pcash]<0,0,[ledger].[pcash])) AS SumPositivePcash, Sum(ledger.cost) AS SumPositiveCost, " & _
                "Sum((IIf([ledger].[pcash]<0,0,[ledger].[pcash])*-1)+([ledger].[cost]*-1)) AS RMA "
19130         strSQL = strSQL & "FROM ledger "
19140         strSQL = strSQL & "WHERE (((ledger.accountno) = '" & strAccountNo & "') AND " & _
                "((ledger.journaltype) = 'Sold') AND ((ledger.icash) >= 0) AND " & _
                "((ledger.transdate) >= #" & Format(datStartDate, "mm/dd/yyyy") & "# AND (ledger.transdate) <= " & _
                "#" & Format(datEndDate, "mm/dd/yyyy") & "#) AND " & _
                "((ledger.pcash) >= 0)) "
19150         strSQL = strSQL & "GROUP BY ledger.accountno;"
19160       Case False
              ' ** Regular statements, not annual.
19170         Select Case blnSingleStatement
              Case True  'cmdPrintStatement_Single

                ' ** Recordset Source; frmStatementParameters.PrintStatements(), by specified accountno.
                ' ** qryStatementParameters_02b.
                'strSQL = "SELECT ledger.accountno, Sum(IIf([ledger].[icash]<0,0,[ledger].[icash])) AS SumPositiveIcash, " & _
                '  "Sum(IIf([ledger].[pcash]<0,0,[ledger].[pcash])) AS SumPositivePcash, Sum(ledger.cost) AS SumPositiveCost, " & _
                '  "Sum((IIf([ledger].[pcash]<0,0,[ledger].[pcash])*-1)+([ledger].[cost]*-1)) AS RMA "
                'strSQL = strSQL & "FROM ledger INNER JOIN qryMaxBalDates ON ledger.accountno = qryMaxBalDates.accountno "
                'strSQL = strSQL & "WHERE (((ledger.accountno)='" & strAccountNo & "') AND " & _
                '  "((ledger.journaltype)='Sold') AND ((ledger.icash)>=0) AND " & _
                '  "((ledger.transdate)>=CDate(Format([qryMaxBalDates].[MaxOfbalance date],'mm/dd/yyyy'))+1 And " & _
                '  "(ledger.transdate)<=#" & Format(datEndDate, "mm/dd/yyyy") & "#) AND ((ledger.pcash)>=0)) "
                'strSQL = strSQL & "GROUP BY ledger.accountno;"

19180           Select Case blnIncludeCurrency
                Case True
                  ' ** qryStatementParameters_02d_06 (qryStatementParameters_02d_05 (qryStatementParameters_02d_04
                  ' ** (qryStatementParameters_02d_03 (qryStatementParameters_02d_01 (Ledger, linked to qryMaxBalDates,
                  ' ** by specified GlobalVarGet("gstrAccountNo","gdatEndDate"); with ForEx, for specified accountno),
                  ' ** qryStatementParameters_02d_02 (LedgerArchive, linked to qryMaxBalDates, by specified
                  ' ** GlobalVarGet("gstrAccountNo","gdatEndDate"); with ForEx, for specified accountno) with ForEx,
                  ' ** for specified accountno), linked to tblCurrency_History; with ForEx, for specified accountno),
                  ' ** grouped, with Max(currentDate); with ForEx, for specified accountno), linked back to
                  ' ** tblCurrency_History, with .._usd fields; with ForEx, for specified accountno), linked to
                  ' ** qryMaxBalDates, grouped and summed; with ForEx, for specified accountno.
19190             Set qdf1 = dbs.QueryDefs("qryStatementParameters_02d")
19200             strSQL = qdf1.SQL
19210           Case False
                  ' ** qryStatementParameters_02a_03 (Union of qryStatementParameters_02b_01 (Ledger, linked to
                  ' ** qryMaxBalDates, by specified GlobalVarGet('gstrAccountNo','gdatEndDate'); for specified accountno),
                  ' ** qryStatementParameters_02b_02 (LedgerArchive, linked to qryMaxBalDates, by specified
                  ' ** GlobalVarGet('gstrAccountNo','gdatEndDate'); for specified accountno)), grouped and summed.
19220             Set qdf1 = dbs.QueryDefs("qryStatementParameters_02b")
19230             strSQL = qdf1.SQL
19240           End Select

19250         Case False

                ' ** Recordset Source; frmStatementParameters.PrintStatements(), for all.
                ' ** qryStatementParameters_02a.
                'strSQL = "SELECT ledger.accountno, Sum(IIf([ledger].[icash]<0,0,[ledger].[icash])) AS SumPositiveIcash, " & _
                '  "Sum(IIf([ledger].[pcash]<0,0,[ledger].[pcash])) AS SumPositivePcash, Sum(ledger.cost) AS SumPositiveCost, " & _
                '  "Sum((IIf([ledger].[pcash]<0,0,[ledger].[pcash])*-1)+([ledger].[cost]*-1)) AS RMA "
                'strSQL = strSQL & "FROM ledger LEFT JOIN qryMaxBalDates ON ledger.accountno = qryMaxBalDates.accountno "
                'strSQL = strSQL & "WHERE (((ledger.journaltype)='Sold') AND ((ledger.icash)>=0) AND " & _
                '  "((ledger.transdate)>=CDate(Format([qryMaxBalDates].[MaxOfbalance date],'mm/dd/yyyy'))+1 And " & _
                '  "(ledger.transdate)<=#" & Format(datEndDate, "mm/dd/yyyy") & "#) AND ((ledger.pcash)>=0)) "
                'strSQL = strSQL & "GROUP BY ledger.accountno;"

19260           Select Case blnIncludeCurrency
                Case True
                  ' ** qryStatementParameters_02c_06 (qryStatementParameters_02c_05 (qryStatementParameters_02c_04
                  ' ** (qryStatementParameters_02c_03 (Union of qryStatementParameters_02c_01 (Ledger, linked to
                  ' ** qryMaxBalDates, by specified GlobalVarGet("gdatEndDate"); with ForEx, for all),
                  ' ** qryStatementParameters_02c_02 (LedgerArchive, linked to qryMaxBalDates, by specified
                  ' ** GlobalVarGet("gdatEndDate"); with ForEx, for all); with ForEx, for all), linked to
                  ' ** tblCurrency_History; with ForEx, for all), grouped, with Max(currentDate); with ForEx,
                  ' ** for all), linked back to tblCurrency_History, with .._usd fields; with ForEx, for all),
                  ' ** linked to qryMaxBalDates, grouped and summed; with ForEx, for all.
19270             Set qdf1 = dbs.QueryDefs("qryStatementParameters_02c")
19280             strSQL = qdf1.SQL
19290           Case False
                  ' ** qryStatementParameters_02a_03 (Union of qryStatementParameters_02a_01 (Ledger,
                  ' ** linked to qryMaxBalDates, by specified GlobalVarGet("gdatEndDate"); for all),
                  ' ** qryStatementParameters_02a_02 (LedgerArchive, linked to qryMaxBalDates, by
                  ' ** specified GlobalVarGet("gdatEndDate"); for all)), grouped and summed.
19300             Set qdf1 = dbs.QueryDefs("qryStatementParameters_02a")
19310             strSQL = qdf1.SQL
19320           End Select

19330         End Select  ' ** blnSingleStatement.
19340       End Select  ' ** blnAnnual.

19350       dbs.QueryDefs("qrySumIncreasesRMA").SQL = strSQL

19360       If blnAnnual = True Then
19370         strAccountNo = gstrAccountNo
19380       ElseIf blnSingleStatement = True Then
19390         strAccountNo = .cmbAccounts
19400       End If

            ' ** Somewhere, we got an error that Statement Date couldn't be updated.
            ' ** I'm juggling a little bit here to see about preventing that.
19410       Set rst1 = Nothing
19420       Set rst2 = Nothing
19430       dbs.Close
19440       Set dbs = Nothing
19450       Set dbs = CurrentDb
19460       DoEvents

19470       If blnAnnual = False And blnRunPriorStatement = False Then  ' ** Should already be datEndDate.
              ' ** Now, finally, use rst1!
              ' ** Now update Statement_Date in the Statement Date table.
19480         Set rst1 = dbs.OpenRecordset("Statement Date", dbOpenDynaset, dbConsistent)
19490         With rst1
19500           If ![Statement_Date] < datEndDate Then
19510             .Edit
19520             ![Statement_Date] = datEndDate
19530             .Update
19540           End If
19550           .Close
19560         End With  ' ** rst1.
19570         Set rst1 = Nothing
19580       End If  ' ** blnAnnual.

19590       Select Case blnAnnual
            Case True
              ' ** Account, with CoInfo, by specified [actno].
19600         Set qdf2 = dbs.QueryDefs("qryStatementParameters_22")
19610         With qdf2.Parameters
19620           ![actno] = gstrAccountNo
19630         End With
19640         Set rst2 = qdf2.OpenRecordset  'dbs.OpenRecordset(strSQL)
19650       Case False
              ' ** The qryStatementParameters_20 query is set in cmbMonth_AfterUpdate(), Above.
19660         Select Case blnSingleStatement
              Case True
                ' ** qryStatementParameters_20 (Account, now with DateClosed = Null), by specified [actno].
19670           Set qdf2 = dbs.QueryDefs("qryStatementParameters_21")
19680           With qdf2.Parameters
19690             ![actno] = strAccountNo
19700           End With
19710           Set rst2 = qdf2.OpenRecordset  'dbs.OpenRecordset(strSQL)
19720         Case False
                ' ** Account, now with DateClosed = Null.
19730           Set qdf2 = dbs.QueryDefs("qryStatementParameters_20")
19740           Set rst2 = qdf2.OpenRecordset  'dbs.OpenRecordset("qryQualifyingAccountsForStatement")
19750         End Select
              ' ** Balance table, by specified FormRef('EndDate').
              'Set qdf1 = dbs.QueryDefs("qryAccountSummary_02")
              ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), linked to
              ' ** qryAccountSummary_02a (Account, with JanX - DecX, by specified FormRef('Month')).
19760         Set qdf1 = dbs.QueryDefs("qryAccountSummary_02b")
19770         Set rst1 = qdf1.OpenRecordset
19780         If rst1.BOF = True And rst1.EOF = True Then
                ' ** Shouldn't have gotten this far!
19790           .DateStart = 0
19800           datStartDate = 0
19810         Else
19820           rst1.MoveFirst
                ' ** This date should update with each account, within the loop!
19830           .DateStart = rst1![MaxOfbalance Date]  ' ** [balance date] < FormRef('EndDate')
19840           datStartDate = rst1![MaxOfbalance Date]
19850         End If
              'THIS IS THE FIRST ONE IN THE LIST, BUT DOESN'T REFLECT THE INIDIVIDUAL ACCT!
19860         rst1.Close
19870         Set rst1 = Nothing
19880         Set qdf1 = Nothing
19890       End Select  ' ** blnAnnual.

19900       If rst2.BOF = True And rst2.EOF = True Then
19910         lngRecs = 0&
19920       Else
19930         rst2.MoveLast
19940         lngRecs = rst2.RecordCount
19950         rst2.MoveFirst
19960       End If

            ' ################################
            ' ## HERE'S THE STATEMENT LOOP!
            ' ################################
            'If .chkDevMsg = True Then
            '  MsgBox "Statement Loop Begins"
            'End If
19970       lngStmtCnt = 0&
19980       For lngX = 1& To lngRecs
19990         Select Case blnAnnual
              Case True
20000           datFirstDate = #1/1/1990#
20010         Case False
20020           .cmbAccounts = rst2![accountno]  ' ** THIS IS WHY cmbAccounts MUST BE VISIBLE AND ENABLED!
                ' ** Predate added to qryAccountNoDropDown_01 and qryAccountNoDropDown_02.
20030           If IsNull(.cmbAccounts.Column(CBX_A_PREDAT)) = True Then  ' ** 3rd column.
20040             datFirstDate = #1/1/1990#
20050           Else
20060             If CDate(.cmbAccounts.Column(CBX_A_BALDAT)) < CDate(.cmbAccounts.Column(CBX_A_PREDAT)) Then
20070               datFirstDate = CDate(.cmbAccounts.Column(CBX_A_BALDAT))
20080             Else
20090               datFirstDate = CDate(.cmbAccounts.Column(CBX_A_PREDAT))
20100             End If
20110           End If
20120           strAccountNo = .cmbAccounts
                ' ** Update DateStart for each account!
20130           varTmp00 = DLookup("[MaxOfbalance date]", "qryAccountSummary_02b", "[accountno] = '" & strAccountNo & "'")
20140           Select Case IsNull(varTmp00)
                Case True
20150             .DateStart = 0
20160             datStartDate = 0
20170           Case False
20180             .DateStart = CDate(varTmp00)
20190             datStartDate = CDate(varTmp00)
20200           End Select
20210           DoEvents
20220         End Select  ' ** blnAnnual.
20230         If datEndDate < datFirstDate Then
                ' ** Skip Accounts opened after the statement date, unless they've got transactions within the period.
20240         Else
20250           If gblnDev_Debug = True Then
20260             MsgBox "We would be printing statement for account " & strAccountNo & "!!", _
                    vbExclamation + vbOKOnly, (Left(("Print Successful" & Space(55)), 55) & "V09")
20270           Else
20280             For intY = 1 To IIf(IsNull(rst2![numCopies]), 1, IIf(rst2![numCopies] = 0, 1, rst2![numCopies]))
20290               If blnContinue = True Then
                      ' ** 2nd Annual Statement branching.
20300                 blnSwitchedOptGrp = False
20310                 If .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then
20320                   .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
20330                   .opgAccountNumber_optSpecified_lbl_box.Visible = True
20340                   .opgAccountNumber_optAll_lbl_box.Visible = False
20350                   blnSwitchedOptGrp = True
20360                 End If
20370                 blnPrintRpt = True
20380                 If blnAnnual = True And blnPrintAll = True Then
20390                   blnFound = False
20400                   For lngZ = 0& To (glngPrintRpts - 1&)
20410                     If garr_varPrintRpt(PR_ACTNO, lngZ) = rst2![accountno] Then
20420                       blnFound = True
20430                       Exit For
20440                     End If
20450                   Next
20460                   For lngZ = 0& To (glngPrintRpts - 1&)
20470                     If garr_varPrintRpt(PR_ACTNO, lngZ) = rst2![accountno] And garr_varPrintRpt(PR_ALIST, lngZ) = True Then
20480                       blnPrintRpt = False
20490                       Exit For
20500                     End If
20510                   Next
20520                 End If
                      ' ****************************************************
                      ' ****************************************************
                      ' ** Print Asset List.
20530                 If blnPrintRpt = True Then
20540                   blnFromStmts = True
20550                   .cmdAssetListPrint_Click  ' ** Procedure: Above.
20560                   blnFromStmts = False
20570                 End If
20580                 DoEvents
                      ' ****************************************************
                      ' ****************************************************
                      ' ** 1. cmdPrintStatement_All_Click()  Above
                      ' ** 1. Here.                          Here
                      ' ** 2. cmdAssetListPrint_Click()      Above
                      ' ** 3. CommonAssetListCode()          Below
                      ' ** 4. BuildAssetListInfo_SP()           Below
                      ' ** 5. MakeTempTable()                modFileUtilities
                      ' ** 6. TableDelete()                  modFileUtilities
                      ' ** 7. Back to here.                  Here
20590                 If blnSwitchedOptGrp = True Then
20600                   blnSwitchedOptGrp = False
20610                   .opgAccountNumber = .opgAccountNumber_optAll.OptionValue
20620                   .opgAccountNumber_optSpecified_lbl_box.Visible = False
20630                   .opgAccountNumber_optAll_lbl_box.Visible = True
20640                 End If
20650                 blnRetVal = blnContinue
20660               End If
20670               DoEvents
20680               If blnContinue = True Then
20690                 If .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then
20700                   .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
20710                   .opgAccountNumber_optSpecified_lbl_box.Visible = True
20720                   .opgAccountNumber_optAll_lbl_box.Visible = False
20730                   blnSwitchedOptGrp = True
20740                 End If
20750                 blnPrintRpt = True
20760                 If blnAnnual = True And blnPrintAll = True Then
20770                   For lngZ = 0& To (glngPrintRpts - 1&)
20780                     If garr_varPrintRpt(PR_ACTNO, lngZ) = rst2![accountno] And garr_varPrintRpt(PR_TRANS, lngZ) = True Then
20790                       blnPrintRpt = False
20800                       Exit For
20810                     End If
20820                   Next
20830                 End If
                      ' ****************************************************
                      ' ****************************************************
                      ' ** Print Transactions.
20840                 If blnPrintRpt = True Then
20850                   blnFromStmts = True
20860                   .cmdTransactionsPrint_Click  ' ** Procedure: Above.
20870                   blnFromStmts = False
20880                 End If
20890                 DoEvents
                      ' ****************************************************
                      ' ****************************************************
20900                 If blnSwitchedOptGrp = True Then
20910                   blnSwitchedOptGrp = False
20920                   .opgAccountNumber = .opgAccountNumber_optAll.OptionValue
20930                   .opgAccountNumber_optSpecified_lbl_box.Visible = False
20940                   .opgAccountNumber_optAll_lbl_box.Visible = True
20950                 End If
20960                 blnRetVal = blnContinue
20970               End If
20980               DoEvents
20990               If blnContinue = True Then
21000                 blnPrintRpt = True
21010                 If blnAnnual = True And blnPrintAll = True Then
21020                   For lngZ = 0& To (glngPrintRpts - 1&)
21030                     If garr_varPrintRpt(PR_ACTNO, lngZ) = rst2![accountno] And garr_varPrintRpt(PR_SUMRY, lngZ) = True Then
21040                       blnPrintRpt = False
21050                       Exit For
21060                     End If
21070                   Next
21080                 End If
                      ' ****************************************************
                      ' ****************************************************
                      ' ** Print Summary.
21090                 If blnPrintRpt = True Then
21100                   blnFromStmts = True
21110                   .cmdSummaryPrint  ' ** Procedure: Above.
21120                   blnFromStmts = False
21130                 End If
21140                 DoEvents
                      ' ****************************************************
                      ' ****************************************************
21150                 blnRetVal = blnContinue
21160               End If
21170               DoEvents
21180             Next  ' ** numCopies: lngY.
21190             If blnContinue = False Then
21200               rst2.MoveLast  ' ** For skip to EOF on user cancel.
21210               lngX = lngRecs
21220             End If
21230           End If
21240         End If
21250         Select Case blnAnnual
              Case True
21260           Exit For  ' ** It should exit anyway.
21270         Case False
21280           If lngX < lngRecs Then rst2.MoveNext
21290           If (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
21300             lngStmtCnt = lngStmtCnt + 1&
21310           End If
21320         End Select  ' ** blnAnnual.

21330         DoEvents
21340         .cmdDevCloseReports_Click  ' ** Procedure: Below.

21350         DoEvents
21360         blnFound = EXE_IsRunning("Acrobat.exe")  ' ** Module Function: modProcessFuncs.
21370         If blnFound = True Then
                ' ** Let's try to close these so they don't pile up and trigger the Windows Max!
21380           Do While blnFound = True
21390             blnFound = EXE_Terminate("Acrobat.exe")  ' ** Module Function: modProcessFuncs.
21400             DoEvents
21410             blnFound = EXE_IsRunning("Acrobat.exe")  ' ** Module Function: modProcessFuncs.
21420           Loop
21430         End If
21440         DoEvents

21450       Next  ' ** lngX.

21460       If blnContinue = True And blnRetVal = True Then

21470         If blnAnnual = False Then
21480           If gblnDev_Debug = False Then  ' ** Don't update values if in Debug mode.
21490             blnFromStmts = True
21500             UpdateBalanceTable1 frm, blnFromStmts, blnIncludeCurrency   ' ** Module Procedure: modStatementParamFuncs2.
21510             blnFromStmts = False
21520           End If
21530         End If  ' ** blnAnnual.

21540         rst2.Close
21550         Set rst2 = Nothing
21560         Set qdf2 = Nothing

              ' ** Empty tmpUpdatedValues.
21570         Set qdf1 = dbs.QueryDefs("qryStatementParameters_17")
21580         qdf1.Execute
21590         Set qdf1 = Nothing

21600         Select Case blnAnnual
              Case True
21610           If .PrintAnnual_cnt > 1& Then
                  ' ** Skip this.
21620           Else
21630             .cmdAnnualStatement.SetFocus
21640             strMsg1 = vbNullString
21650             DoCmd.Hourglass False
21660             MsgBox "Annual Statement done for period ending " & Format(datEndDate, "mm/dd/yyyy") & "." & vbCrLf & vbCrLf & _
                    "1 Statement processed.", vbInformation + vbOKOnly, (Left(("Statement Finished" & Space(55)), 55) & "V10")
21670           End If
21680         Case False

21690           If .cmdPrintStatement_All.Enabled = True Then
21700             .cmdPrintStatement_All.SetFocus
21710           ElseIf .cmdPrintStatement_Single.Enabled = True Then
21720             .cmdPrintStatement_Single.SetFocus
21730           Else
21740             .chkStatements.SetFocus
21750           End If

21760           Select Case blnAllStatements
                Case True
21770             If lngStatements > 1& Then strMsg1 = "s" Else strMsg1 = vbNullString
21780             .cmbAccounts.Locked = False
21790             .cmbAccounts.Enabled = False
21800             .cmbAccounts.BorderColor = WIN_CLR_DISR
21810             .cmbAccounts.BackStyle = acBackStyleTransparent
21820             .cmbAccounts_lbl_box.Visible = True
21830             .cmbAccounts_lbl.ForeColor = CLR_WHT
21840             DoEvents
21850             .cmbAccounts_lbl.BackStyle = acBackStyleTransparent
21860             .cmbAccounts = Null
21870             .cmbAccounts.ForeColor = CLR_BLK
21880             .cmbAccounts.BackColor = CLR_WHT
21890             DoEvents
21900           Case False
21910             strMsg1 = vbNullString
21920           End Select

21930           .chkStatements_lbl3.Caption = Format(DLookup("[Statement_Date]", "Statement Date"), "mm/dd/yyyy")
                ' ** qryStatementParameters_08c gives Max Balance Date for specified accountno.

21940           DoCmd.Hourglass False
21950           DoEvents
21960           Beep
21970           MsgBox "Statement" & strMsg1 & " done for period ending " & Format(datEndDate, "mm/dd/yyyy") & "." & vbCrLf & vbCrLf & _
                  CStr(lngStatements) & " Statement" & strMsg1 & " processed.", _
                  vbInformation + vbOKOnly, (Left(("Statement" & strMsg1 & " Finished" & Space(55)), 55) & "V11")

21980         End Select  ' ** blnAnnual.

21990       End If  ' ** blnContinue, blnRetVal.

22000     Else
22010       Select Case blnWasGTR
            Case True
22020         .cmdPrintStatement_All.SetFocus
22030       Case False
22040         .StatementsYear.SetFocus
22050       End Select
22060     End If  ' ** blnPrintStatements, blnRetVal.

22070     dbs.Close

22080   End With  ' ** frm.

EXITP:
22090   Set rst1 = Nothing
22100   Set rst2 = Nothing
22110   Set qdf1 = Nothing
22120   Set qdf2 = Nothing
22130   Set dbs = Nothing
22140   Statements_Print = blnRetVal
22150   Exit Function

ERRH:
4200    blnRetVal = False
4210    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** I canceled something, closed the window, and THEN the errors popped up!
4220    Case Else
4230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4240    End Select
4250    Resume EXITP

End Function

Public Sub AccountSummary_Print(frm As Access.Form, blnContinue As Boolean, blnFromStmts As Boolean, strReportName As String, blnHasForEx As Boolean, blnHasForExThis As Boolean)
' ** Called by:
' **   Statements_Print(), Below.
' **   cmdPrintStatement_Summary_Click(), Above.
' ** (Not really a command button, but I was organizing...)

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "AccountSummary_Print"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, rpt1 As Access.Report, rpt2 As Access.Report
        Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, datTmp03 As Date
        Dim blnRetVal As Boolean

22210   With frm

22220     DoCmd.Hourglass True
22230     DoEvents

22240     blnRetVal = True

22250     Set dbs = CurrentDb

          ' ** Grab current masterasset date (the same for all records, so just get one).
          ' ** MasterAsset, grouped by currentDate, with MaxDate.
22260     Set qdf = dbs.QueryDefs("qryAccountSummary_01")
22270     Set rst = qdf.OpenRecordset
22280     If IsNull(rst![MaxDate]) Then
22290       blnRetVal = False
22300       DoCmd.Hourglass False
22310       MsgBox "All assets must be priced in order to run statements.", vbInformation + vbOKOnly, _
              (Left(("Missing An Asset Current Date" & Space(55)), 55) & "U01")
22320       rst.Close
22330       dbs.Close
22340       Set rst = Nothing
22350       Set qdf = Nothing
22360       Set dbs = Nothing
22370     Else
22380       rst.MoveFirst
22390       If IsNull(rst![MaxDate]) Then
22400         blnRetVal = False
22410         DoCmd.Hourglass False
22420         MsgBox "All assets must be priced in order to run statements.", vbInformation + vbOKOnly, _
                (Left(("Missing An Asset Current Date" & Space(55)), 55) & "U02")
22430         rst.Close
22440         dbs.Close
22450         Set rst = Nothing
22460         Set qdf = Nothing
22470         Set dbs = Nothing
22480       Else
22490         gdatMarketDate = rst![MaxDate]
22500         rst.Close
22510         Set rst = Nothing
22520         Set qdf = Nothing
22530         dbs.Close
22540         Set dbs = Nothing
22550       End If
22560     End If

          'SHOULD WE CHECK FIRST, AND ONLY DO THIS IF THEY'RE EMPTY?

22570     If blnRetVal = True Then

            ' ** We need to make sure .DateStart and .DateEnd are populated!
22580       lngTmp01 = .cmbMonth.Column(CBX_MON_ID)
22590       lngTmp02 = .StatementsYear
22600       gdatStartDate = DateSerial(lngTmp02, lngTmp01, 1)  ' ** Default to 1 month only.
22610       gdatEndDate = DateAdd("y", -1, DateAdd("m", 1, gdatStartDate))  ' ** 1st of month, plus 1 month, minus 1 day.

22620       .DateEnd = gdatEndDate

            ' ** Balance, grouped by accountno, with Max(balance date), by specified FormRef('accountno','EndDate').
22630       varTmp00 = DLookup("[balance date]", "qryStatementParameters_Summary_01")
22640       If IsNull(varTmp00) = False Then
22650         gdatStartDate = DateAdd("y", 1, CDate(varTmp00))
22660       End If

22670       .DateStart = gdatStartDate

22680       Set dbs = CurrentDb
22690       With dbs
22700         Set rst = .OpenRecordset("tblPricing_MasterAsset_History", dbOpenDynaset, dbReadOnly)
22710         If rst.BOF = True And rst.EOF = True Then
                ' ** No pricing history records exist!
22720           rst.Close
22730           Set rst = Nothing
                ' ** Check that they all have the same currentDate.
22740           datTmp03 = 0
                ' ** MasterAsset, grouped by currentDate, with cnt_ast.
22750           Set qdf = .QueryDefs("qryAccountSummary_16_02")
22760           Set rst = qdf.OpenRecordset
22770           With rst
                  ' ** This shouldn't ever be empty!
22780             .MoveLast
22790             If .RecordCount > 1 Then
                    ' ** There shouldn't be more than one currentDate!
22800               .MoveFirst
22810               datTmp03 = ![currentDate]  ' ** Take the most prevalent one.
22820             End If
22830             .Close
22840           End With
22850           Set rst = Nothing
22860           If datTmp03 <> 0 Then
                  ' ** Update qryAccountSummary_16_03 (MasterAsset, with currentDate_new, by specified [curdat]).
22870             Set qdf = .QueryDefs("qryAccountSummary_16_04")
22880             With qdf.Parameters
22890               ![curdat] = datTmp03
22900             End With
22910             qdf.Execute
22920             Set qdf = Nothing
22930           End If
22940           DoEvents
                ' ** Now copy the current pricing to pricing history.
                ' ** Append qryAccountSummary_16_01 (MasterAsset, linked to tblCurrency, as new
                ' ** tblPricing_MasterAsset_History records) to tblPricing_MasterAsset_History.
22950           Set qdf = .QueryDefs("qryAccountSummary_16_05")
22960           qdf.Execute
22970           Set qdf = Nothing
22980           DoEvents
22990         Else
                ' ** Pricing history records exist.
23000           rst.Close
23010         End If
23020         Set rst = Nothing
              ' ** tblPricing_MasterAsset_History, grouped, with Max(currentDate), by specified [datend].
23030         Set qdf = .QueryDefs("qryAccountSummary_15")
23040         With qdf.Parameters
23050           ![datEnd] = gdatEndDate
23060         End With
23070         Set rst = qdf.OpenRecordset
23080         With rst
23090           .MoveFirst
23100           gdatMarketDate = ![currentDate]
23110           .Close
23120         End With
23130         Set rst = Nothing
23140         Set qdf = Nothing
23150         .Close
23160       End With
23170       Set dbs = Nothing

23180       .currentDate = gdatMarketDate

23190       If (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
23200         If IsLoaded("rptAccountSummary", acReport) = True Then  ' ** Module Function: modFileUtilities.
23210           DoCmd.Close acReport, "rptAccountSummary"
23220         End If
23230       End If

23240       gstrReportQuerySpec = "rptAccountSummary"

            ' ** These are for all accounts, whether or not they're scheduled for this month.
            ' ** Certainly, if they're all Zero, that includes scheduled, but,
            ' ** the scheduled could be Zero, with others not!
23250       gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False  ' ** Borrowing these variables from Court Reports.

23260       Select Case blnFromStmts
            Case True
23270         Select Case blnIncludeCurrency
              Case True
23280           strReportName = "rptAccountSummary_ForEx"
                'varTmp00 = DCount("*", "qryAccountSummary_12_02_03")  ' ** Ledger Only, ForEx.
                ' ** qryAccountSummary_12_04_01 (qryAccountSummary_11_05 (Union of qryAccountSummary_11_01
                ' ** (Ledger, linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate'); Ledger Only, ForEx),
                ' ** qryAccountSummary_11_03 (LedgerArchive, linked to qryAccountSummary_02 (Balance, by
                ' ** specified FormRef('EndDate')), tblCurrency, with .._usd fields, by specified
                ' ** FormRef('EndDate'); LedgerArchive Only, ForEx); Ledger, LedgerArchive, ForEx), linked
                ' ** to JournalType; Ledger, LedgerArchive), linked to qryAccountSummary_12_04_02 (Account,
                ' ** with JanX - DecX, by specified FormRef('Month')); Ledger, LedgerArchive.
23290           varTmp00 = DCount("*", "qryAccountSummary_12_04_03")  ' ** Ledger, LedgerArchive, ForEx.
23300         Case False
23310           strReportName = "rptAccountSummary"
                'varTmp00 = DCount("*", "qryAccountSummary_12_01_03")  ' ** Ledger Only.
                ' ** qryAccountSummary_12_03_01 (qryAccountSummary_11_04 (Union of qryAccountSummary_11
                ' ** (Ledger, linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                ' ** by specified FormRef('EndDate'); Ledger Only), qryAccountSummary_11_02 (LedgerArchive,
                ' ** linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                ' ** FormRef('EndDate'); LedgerArchive Only); Ledger, LedgerArchive), linked to JournalType;
                ' ** Ledger, LedgerArchive), linked to qryAccountSummary_12_03_02 (Account, with JanX - DecX,
                ' ** by specified FormRef('Month')); Ledger, LedgerArchive.
23320           varTmp00 = DCount("*", "qryAccountSummary_12_03_03")  ' ** Ledger, LedgerArchive.
23330         End Select
23340       Case False
23350         Select Case blnHasForEx
              Case True
23360           blnHasForExThis = HasForEx_SP(.cmbAccounts)  ' ** Module Function: modStatementParamFuncs1.
23370           If .chkIncludeCurrency <> blnHasForExThis Then
23380             .chkIncludeCurrency = blnHasForExThis
23390             .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
23400             DoEvents
23410           End If
23420           Select Case blnHasForExThis
                Case True
23430             strReportName = "rptAccountSummary_ForEx"
                  ' ** qryAccountSummary_12_02_01 (qryAccountSummary_11_01 (Ledger, linked
                  ' ** to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                  ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate')),
                  ' ** linked to JournalType), linked to qryAccountSummary_12_02_02
                  ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                  'varTmp00 = DCount("*", "qryAccountSummary_12_02_03")  ' ** Ledger Only, ForEx.
                  ' ** qryAccountSummary_12_04_01 (qryAccountSummary_11_05 (Union of qryAccountSummary_11_01
                  ' ** (Ledger, linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                  ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate'); Ledger Only, ForEx),
                  ' ** qryAccountSummary_11_03 (LedgerArchive, linked to qryAccountSummary_02 (Balance, by
                  ' ** specified FormRef('EndDate')), tblCurrency, with .._usd fields, by specified
                  ' ** FormRef('EndDate'); LedgerArchive Only, ForEx); Ledger, LedgerArchive, ForEx), linked
                  ' ** to JournalType; Ledger, LedgerArchive, ForEx), linked to qryAccountSummary_12_04_02
                  ' ** (Account, with JanX - DecX, by specified FormRef('Month')); Ledger, LedgerArchive, ForEx.
23440             varTmp00 = DCount("*", "qryAccountSummary_12_04_03")  ' ** Ledger, LedgerArchive, ForEx.
23450           Case False
23460             Select Case .chkIncludeCurrency
                  Case True
23470               strReportName = "rptAccountSummary_ForEx"
                    ' ** qryAccountSummary_12_02_01 (qryAccountSummary_11_01 (Ledger, linked
                    ' ** to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                    ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate')),
                    ' ** linked to JournalType), linked to qryAccountSummary_12_02_02
                    ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                    'varTmp00 = DCount("*", "qryAccountSummary_12_02_03")  ' ** Ledger Only, ForEx.
                    ' ** qryAccountSummary_12_04_01 (qryAccountSummary_11_05 (Union of qryAccountSummary_11_01
                    ' ** (Ledger, linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                    ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate'); Ledger Only, ForEx),
                    ' ** qryAccountSummary_11_03 (LedgerArchive, linked to qryAccountSummary_02 (Balance, by
                    ' ** specified FormRef('EndDate')), tblCurrency, with .._usd fields, by specified
                    ' ** FormRef('EndDate'); LedgerArchive Only, ForEx); Ledger, LedgerArchive, ForEx), linked
                    ' ** to JournalType; Ledger, LedgerArchive, ForEx), linked to qryAccountSummary_12_04_02
                    ' ** (Account, with JanX - DecX, by specified FormRef('Month')); Ledger, LedgerArchive, ForEx.
23480               varTmp00 = DCount("*", "qryAccountSummary_12_04_03")  ' ** Ledger, LedgerArchive, ForEx.
23490             Case False
23500               strReportName = "rptAccountSummary"
                    ' ** qryAccountSummary_12_01_01 (qryAccountSummary_11 (Ledger, linked to
                    ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                    ' ** FormRef('EndDate')), linked to JournalType), linked to qryAccountSummary_12_01_02
                    ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                    'varTmp00 = DCount("*", "qryAccountSummary_12_01_03")  ' ** Ledger Only.
                    ' ** qryAccountSummary_12_03_01 (qryAccountSummary_11_04 (Union of qryAccountSummary_11
                    ' ** (Ledger, linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                    ' ** by specified FormRef('EndDate'); Ledger Only), qryAccountSummary_11_02 (LedgerArchive,
                    ' ** linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                    ' ** FormRef('EndDate'); LedgerArchive Only); Ledger, LedgerArchive), linked to JournalType;
                    ' ** Ledger, LedgerArchive), linked to qryAccountSummary_12_03_02 (Account, with JanX - DecX,
                    ' ** by specified FormRef('Month')); Ledger, LedgerArchive.
23510               varTmp00 = DCount("*", "qryAccountSummary_12_03_03")  ' ** Ledger, LedgerArchive.
23520             End Select
23530           End Select
23540         Case False
23550           strReportName = "rptAccountSummary"
                ' ** qryAccountSummary_12_01_01 (qryAccountSummary_11 (Ledger, linked to
                ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                ' ** FormRef('EndDate')), linked to JournalType), linked to qryAccountSummary_12_01_02
                ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                'varTmp00 = DCount("*", "qryAccountSummary_12_01_03")  ' ** Ledger Only.
                ' ** qryAccountSummary_12_03_01 (qryAccountSummary_11_04 (Union of qryAccountSummary_11
                ' ** (Ledger, linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                ' ** by specified FormRef('EndDate'); Ledger Only), qryAccountSummary_11_02 (LedgerArchive,
                ' ** linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                ' ** FormRef('EndDate'); LedgerArchive Only); Ledger, LedgerArchive), linked to JournalType;
                ' ** Ledger, LedgerArchive), linked to qryAccountSummary_12_03_02 (Account, with JanX - DecX,
                ' ** by specified FormRef('Month')); Ledger, LedgerArchive.
23560           varTmp00 = DCount("*", "qryAccountSummary_12_03_03")  ' ** Ledger, LedgerArchive.
23570         End Select
23580       End Select  ' ** blnFromStmts.

23590       Select Case IsNull(varTmp00)
            Case True
23600         gblnCrtRpt_Zero = True
23610       Case False
23620         If varTmp00 = 0 Then gblnCrtRpt_Zero = True
23630       End Select

23640       Select Case blnFromStmts
            Case True
23650         Select Case blnIncludeCurrency
              Case True
                'varTmp00 = DCount("*", "qryAccountSummary_14_02_03")  ' ** Ledger Only, ForEx.
                ' ** qryAccountSummary_14_04_01 (qryAccountSummary_13_05 (Union of qryAccountSummary_13_01 (Ledger,
                ' ** linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), tblCurrency, with
                ' ** .._usd fields, by specified FormRef('EndDate'); Ledger Only, ForEx), qryAccountSummary_13_03
                ' ** (LedgerArchive, linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate'); LedgerArchive Only, ForEx);
                ' ** Ledger, LedgerArchive, ForEx), linked to JournalType; Ledger, LedgerArchive, ForEx), linked
                ' ** to qryAccountSummary_14_04_02 (Account, with JanX - DecX, by specified FormRef('Month'));
                ' ** Ledger, LedgerArchive, ForEx.
23660           varTmp00 = DCount("*", "qryAccountSummary_14_04_03")  ' ** Ledger, LedgerArchive, ForEx.
23670         Case False
                'varTmp00 = DCount("*", "qryAccountSummary_14_01_03")  ' ** Ledger Only.
                ' ** qryAccountSummary_14_03_01 (qryAccountSummary_13_04 (Union of qryAccountSummary_13 (Ledger,
                ' ** linked to qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                ' ** FormRef('EndDate'); Ledger Only), qryAccountSummary_13_02 (LedgerArchive, linked to
                ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                ' ** FormRef('EndDate'); LedgerArchive Only); Ledger, LedgerArchive), linked to JournalType;
                ' ** Ledger, LedgerArchive), linked to qryAccountSummary_14_03_02 (Account, with JanX - DecX,
                ' ** by specified FormRef('Month')); Ledger, LedgerArchive.
23680           varTmp00 = DCount("*", "qryAccountSummary_14_03_03")  ' ** Ledger, LedgerArchive.
23690         End Select
23700       Case False
23710         Select Case blnHasForEx
              Case True
23720           Select Case blnHasForExThis
                Case True
                  ' ** qryAccountSummary_14_02_01 (qryAccountSummary_13_01 (Ledger, linked to
                  ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                  ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate')),
                  ' ** linked to JournalType), linked to qryAccountSummary_14_02_02
                  ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                  'varTmp00 = DCount("*", "qryAccountSummary_14_02_03")  ' ** Ledger Only, ForEx.
                  ' ** qryAccountSummary_14_04_01 (xx), linked to qryAccountSummary_14_04_02 (xx); Ledger, LedgerArchive, ForEx.
23730             varTmp00 = DCount("*", "qryAccountSummary_14_04_03")  ' ** Ledger, LedgerArchive, ForEx.
23740           Case False
23750             Select Case .chkIncludeCurrency
                  Case True
                    ' ** qryAccountSummary_14_02_01 (qryAccountSummary_13_01 (Ledger, linked to
                    ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')),
                    ' ** tblCurrency, with .._usd fields, by specified FormRef('EndDate')),
                    ' ** linked to JournalType), linked to qryAccountSummary_14_02_02
                    ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                    'varTmp00 = DCount("*", "qryAccountSummary_14_02_03")  ' ** Ledger Only, ForEx.
                    ' ** qryAccountSummary_14_04_01 (xx), linked to qryAccountSummary_14_04_02 (xx); Ledger, LedgerArchive, ForEx.
23760               varTmp00 = DCount("*", "qryAccountSummary_14_04_03")  ' ** Ledger, LedgerArchive, ForEx.
23770             Case False
                    ' ** qryAccountSummary_14_01_01 (qryAccountSummary_13 (Ledger, linked to
                    ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                    ' ** FormRef('EndDate')), linked to JournalType), linked to qryAccountSummary_14_01_02
                    ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                    'varTmp00 = DCount("*", "qryAccountSummary_14_01_03")  ' ** Ledger Only.
                    ' ** qryAccountSummary_14_03_01 (xx), linked to qryAccountSummary_14_03_02 (xx); Ledger, LedgerArchive.
23780               varTmp00 = DCount("*", "qryAccountSummary_14_03_03")  ' ** Ledger, LedgerArchive.
23790             End Select
23800           End Select
23810         Case False
                ' ** qryAccountSummary_14_01_01 (qryAccountSummary_13 (Ledger, linked to
                ' ** qryAccountSummary_02 (Balance, by specified FormRef('EndDate')), by specified
                ' ** FormRef('EndDate')), linked to JournalType), linked to qryAccountSummary_14_01_02
                ' ** (Account, with JanX - DecX, by specified FormRef('Month')).
                'varTmp00 = DCount("*", "qryAccountSummary_14_01_03")  ' ** Ledger Only.
                ' ** qryAccountSummary_14_03_01 (xx), linked to qryAccountSummary_14_03_02 (xx); Ledger, LedgerArchive.
23820           varTmp00 = DCount("*", "qryAccountSummary_14_03_03")  ' ** Ledger, LedgerArchive.
23830         End Select
23840       End Select  ' ** blnFromStmts.

23850       Select Case IsNull(varTmp00)
            Case True
23860         gblnCrtRpt_ZeroDialog = True
23870       Case False
23880         If varTmp00 = 0 Then gblnCrtRpt_ZeroDialog = True
23890       End Select

23900       If strReportName = vbNullString Then strReportName = "rptAccountSummary"  ' ** Redundant

23910       DoCmd.OpenReport strReportName, acViewPreview  ' ** This also gets the caption changed!
23920       If .chkCombineCash.Value = True Then
23930         Select Case strReportName
              Case "rptAccountSummary"
23940           Set rpt1 = Reports(strReportName)
23950           With rpt1
23960             Set rpt2 = .rptAccountSummary_Sub_Increases.Report
23970             With rpt2
23980               .CurrentPositiveIcash_lbl.Visible = False
23990               .CurrentPositivePCash_lbl.Visible = False
24000               .CurrentPositiveCash_lbl.Visible = True
24010               .CurrentPositiveICash_lbl_line.Visible = False
24020               .CurrentPositivePCash_lbl_line.Visible = False
24030               .CurrentPositiveCash_lbl_line.Visible = True
24040               .CurrentPositiveICash.Visible = False
24050               .CurrentPositivePCash.Visible = False
24060               .CurrentPositiveCash.Visible = True
24070               .TotalPositiveICash.Visible = False
24080               .TotalPositivePCash.Visible = False
24090               .TotalPositiveCash.Visible = True
24100               .TotalPositiveICash_line.Visible = False
24110               .TotalPositivePCash_line.Visible = False
24120               .TotalPositiveCash_line.Visible = True
                    ' ** We could do, If gblnCrtRpt_Zero = True Then turn on/off things.
24130             End With
24140             Set rpt2 = Nothing
24150             DoEvents
24160             Set rpt2 = .rptAccountSummary_Sub_Decreases.Report
24170             With rpt2
24180               .CurrentNegativeICash_lbl.Visible = False
24190               .CurrentNegativePCash_lbl.Visible = False
24200               .CurrentNegativeCash_lbl.Visible = True
24210               .CurrentNegativeICash_lbl_line.Visible = False
24220               .CurrentNegativePCash_lbl_line.Visible = False
24230               .CurrentNegativeCash_lbl_line.Visible = True
24240               .CurrentNegativeIcash.Visible = False
24250               .CurrentNegativePcash.Visible = False
24260               .CurrentNegativeCash.Visible = True
24270               .TotalNegativeICash.Visible = False
24280               .TotalNegativePCash.Visible = False
24290               .TotalNegativeCash.Visible = True
24300               .TotalNegativeICash_line.Visible = False
24310               .TotalNegativePCash_line.Visible = False
24320               .TotalNegativeCash_line.Visible = True
                    ' ** We could do, If gblnCrtRpt_ZeroDialog = True Then turn on/off things.
24330             End With
24340             Set rpt2 = Nothing
24350             DoEvents
24360           End With
24370           Set rpt1 = Nothing
24380         Case "rptAccountSummary_ForEx"
24390           Set rpt1 = Reports(strReportName)
24400           With rpt1
24410             Set rpt2 = .rptAccountSummary_ForEx_Sub_Increases.Report
24420             With rpt2
24430               .CurrentPositiveICash_usd_lbl.Visible = False
24440               .CurrentPositivePCash_usd_lbl.Visible = False
24450               .CurrentPositiveCash_usd_lbl.Visible = True
24460               .CurrentPositiveICash_usd_lbl_line.Visible = False
24470               .CurrentPositivePCash_usd_lbl_line.Visible = False
24480               .CurrentPositiveCash_usd_lbl_line.Visible = True
24490               .CurrentPositiveICash_usd.Visible = False
24500               .CurrentPositivePCash_usd.Visible = False
24510               .CurrentPositiveCash_usd.Visible = True
24520               .TotalPositiveICash_usd.Visible = False
24530               .TotalPositivePCash_usd.Visible = False
24540               .TotalPositiveCash_usd.Visible = True
24550               .TotalPositiveICash_usd_line.Visible = False
24560               .TotalPositivePCash_usd_line.Visible = False
24570               .TotalPositiveCash_usd_line.Visible = True
24580             End With
24590             Set rpt2 = Nothing
24600             DoEvents
24610             Set rpt2 = .rptAccountSummary_ForEx_Sub_Decreases.Report
24620             With rpt2
24630               .CurrentNegativeIcash_usd_lbl.Visible = False
24640               .CurrentNegativePcash_usd_lbl.Visible = False
24650               .CurrentNegativeCash_usd_lbl.Visible = True
24660               .CurrentNegativeIcash_usd_lbl_line.Visible = False
24670               .CurrentNegativePcash_usd_lbl_line.Visible = False
24680               .CurrentNegativeCash_usd_lbl_line.Visible = True
24690               .CurrentNegativeIcash_usd.Visible = False
24700               .CurrentNegativePcash_usd.Visible = False
24710               .CurrentNegativeCash_usd.Visible = True
24720               .TotalNegativeICash_usd.Visible = False
24730               .TotalNegativePCash_usd.Visible = False
24740               .TotalNegativeCash_usd.Visible = True
24750               .TotalNegativeICash_usd_line.Visible = False
24760               .TotalNegativePCash_usd_line.Visible = False
24770               .TotalNegativeCash_usd_line.Visible = True
24780             End With
24790             Set rpt2 = Nothing
24800             DoEvents
24810           End With
24820           Set rpt1 = Nothing
24830         End Select
24840       End If

24850       If gblnDev_Debug = True Or (CurrentUser = "Superuser") Then ' And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
              'DoCmd.Maximize
              'DoCmd.RunCommand acCmdFitToWindow
24860         DoCmd.OpenReport strReportName, acViewNormal
24870         DoEvents
24880         If IsLoaded(strReportName, acReport) = True Then  ' ** Module Function: modFileUtilities.
24890           DoCmd.Close acReport, strReportName
24900         End If
24910       Else
              '##GTR_Ref: rptAccountSummary
              '##GTR_Ref: rptAccountSummary_ForEx
24920         DoCmd.OpenReport strReportName, acViewNormal
24930         DoEvents
24940         If IsLoaded(strReportName, acReport) = True Then  ' ** Module Function: modFileUtilities.
24950           DoCmd.Close acReport, strReportName
24960         End If
24970       End If

            '=IIf(
            '     IsNull([CurrentTotalMarketValue]),0,
            '     (
            '      ([CurrentTotalMarketValue])-
            '      (
            '       IIf(
            '           IsNull([PreviousTotalMarketValue]),0,
            '           [PreviousTotalMarketValue]
            '          )+
            '       IIf(
            '           IsNull([SumPositiveIcash]),0,
            '           [SumPositiveIcash]
            '          )+
            '       IIf(
            '           IsNull([SumPositivePcash]),0,
            '           [SumPositivePcash]
            '          )+
            '       IIf(
            '           IsNull([SumNegativeIcash]),0,
            '           [SumNegativeIcash]
            '          )+
            '       IIf(
            '           IsNull([SumNegativePcash]),0,
            '           [SumNegativePcash]
            '          )+
            '       IIf(
            '           IsNull([SumPositiveCost]),0,
            '           [SumPositiveCost]
            '          )+
            '       IIf(
            '           IsNull([SumNegativeCost]),0,
            '           [SumNegativeCost]
            '          )
            '      )
            '     )
            '    )

24980     End If  ' ** blnRetVal.

24990     DoCmd.Hourglass False

25000   End With

EXITP:
25010   Set rpt1 = Nothing
25020   Set rpt2 = Nothing
25030   Set rst = Nothing
25040   Set qdf = Nothing
25050   Set dbs = Nothing
25060   Exit Sub

ERRH:
4200    DoCmd.Hourglass False
4210    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
4220      blnContinue = False
4230      If Reports.Count > 0 Then
4240        DoCmd.Close acReport, Reports(0).Name  ' ** Close report in preview.
4250      End If
4260    Case Else
4270      Select Case ERR.Number
          Case Else
4280        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4290      End Select
4300    End Select
4310    Resume EXITP

End Sub

Public Sub AnnualStatement_Print(frm As Access.Form, blnContinue As Boolean, blnFromStmts As Boolean, blnPrintAnnualStatement As Boolean, blnPrintStatements As Boolean, blnAllStatements As Boolean, blnSingleStatement As Boolean, blnRunPriorStatement As Boolean, blnAcctsSchedRpt As Boolean, datFirstDate As Date, blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnGTR_Emblem As Boolean, blnWasGTR As Boolean)
' ** December or 4th Quarter statement must have been run.
' ** Called by:
' **   frmStatementParameters:
' **     cmdAnnualStatement_Click()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "AnnualStatement_Print"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim datLastYearEnd As Date, datPrevYearEnd As Date, blnLast As Boolean, blnPrev As Boolean, blnNoTrans As Boolean
        Dim datFirstYearEnd As Date
        Dim blnFirstIsZero As Boolean, blnLastIsZero As Boolean, blnPrevIsZero As Boolean
        Dim strAccountNo As String, strShortName As String, lngFirstElem As Long
        Dim lngBals As Long, arr_varBal As Variant
        Dim lngAccts As Long, arr_varAcct() As Variant
        Dim lngNoBals As Long, lngRuns As Long, lngAcctsScheduled As Long, lngAcctsCnt As Long
        Dim strMsg1 As String, strMsg2 As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue2 As Boolean, blnMsgShown As Boolean, blnPrintAll As Boolean, blnFound As Boolean, blnOpgChanged As Boolean
        Dim varTmp00 As Variant, lngTmp01 As Long
        Dim lngX As Long, lngY As Long, lngE As Long

        ' ** Array: arr_varBal().
        Const B_ACTNO As Integer = 0
        Const B_DATE  As Integer = 1
        Const B_ICASH As Integer = 2
        Const B_PCASH As Integer = 3
        Const B_COST  As Integer = 4
        'Const B_BAL   As Integer = 7

        ' ** Array: arr_varAcct().
        Const A_ELEMS As Integer = 14  ' ** Array's first-element UBound().
        Const A_ACTNO  As Integer = 0
        Const A_BAL    As Integer = 1
        Const A_BELEM1 As Integer = 2
        Const A_BCNT   As Integer = 3
        Const A_FDAT   As Integer = 4
        Const A_FDATZ  As Integer = 5
        Const A_FELEM  As Integer = 6
        Const A_PDAT   As Integer = 7
        Const A_PDATZ  As Integer = 8
        Const A_PELEM  As Integer = 9
        Const A_LDAT   As Integer = 10
        Const A_LDATZ  As Integer = 11
        Const A_LELEM  As Integer = 12
        Const A_ALLZ   As Integer = 13
        Const A_RUN    As Integer = 14

25110   blnContinue2 = True
25120   blnMsgShown = False
25130   lngAcctsScheduled = 0&
25140   DoCmd.Hourglass True
25150   DoEvents

25160   With frm

25170     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
25180       blnPrintAll = False
25190     Case .opgAccountNumber_optAll.OptionValue
25200       blnPrintAll = True
25210       lngAccts = 0&
25220       ReDim arr_varAcct(A_ELEMS, 0)
25230       glngPrintRpts = 0&
25240       ReDim garr_varPrintRpt(PR_ELEMS, 0)
25250     End Select

          ' ** Account, grouped by smtdec = true, with cnt.
25260     varTmp00 = DLookup("[cnt]", "qryStatementAnnual_06")
25270     If IsNull(varTmp00) = True Then
25280       blnContinue2 = False
25290       Beep
25300       MsgBox "There are no accounts scheduled for December statements." & vbCrLf & vbCrLf & _
              "An Annual Statement can only be run for Accounts" & vbCrLf & "that have a December statement balance.", _
              vbInformation + vbOKOnly, "No December Statements"
25310     Else
25320       If varTmp00 = 0 Then
25330         blnContinue2 = False
25340         Beep
25350         MsgBox "There are no accounts scheduled for December statements." & vbCrLf & vbCrLf & _
                "An Annual Statement can only be run for Accounts" & vbCrLf & "that have a December statement balance.", _
                vbInformation + vbOKOnly, "No December Statements"
25360       Else
25370         lngAcctsScheduled = varTmp00
25380         varTmp00 = Empty
              ' ** Account, all accounts, with cnt.
25390         varTmp00 = DLookup("[cnt]", "qryStatementAnnual_07")
25400         lngAcctsCnt = varTmp00
25410         varTmp00 = Empty

25420         If blnPrintAll = False Then
25430           If IsNull(.cmbAccounts) = True Then
25440             blnContinue2 = False
25450             DoCmd.Hourglass False
25460             MsgBox "You must select an account to continue," & vbCrLf & _
                    "or choose All.", vbInformation + vbOKOnly, _
                    (Left("Entry Required" & Space(55), 55) & "F01")
25470             If .opgAccountNumber <> .opgAccountNumber_optSpecified.OptionValue Then
25480               .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
25490               .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
25500             End If
25510             If IsNull(.cmbMonth) = True Then
25520               .cmbMonth = "December"
25530             Else
25540               If .cmbMonth = vbNullString Then
25550                 .cmbMonth = "December"
25560               End If
25570             End If
25580             .StatementsYear = (year(Date) - 1)
25590             DoEvents
25600             .cmbAccounts.SetFocus
25610           Else
25620             strAccountNo = .cmbAccounts.Column(CBX_A_ACTNO)
25630             strShortName = .cmbAccounts.Column(CBX_A_SHORT)
25640             varTmp00 = DLookup("[smtdec]", "account", "[accountno] = '" & strAccountNo & "'")
25650             If CBool(varTmp00) = False Then
25660               blnContinue2 = False
25670               Beep
25680               MsgBox "The selected account does not have a statement scheduled for December.", _
                      vbInformation + vbOKOnly, "No December Statement Scheduled"
25690             End If
25700           End If  ' ** cmbAccounts.
25710         End If  ' ** blnPrintAll.

25720         If blnContinue2 = True Then

25730           datLastYearEnd = DateSerial((year(Date) - 1), 12, 31)
25740           datPrevYearEnd = DateSerial((year(Date) - 2), 12, 31)

                ' ** I think my covering all possibilities got carried away!
25750           If IsNull(.cmbMonth) = True And IsNull(.StatementsYear) = True Then
25760             .cmbMonth = Format(datLastYearEnd, "mmmm")
25770             .StatementsYear = year(datLastYearEnd)
25780             DoCmd.Hourglass False
25790             DoEvents
25800             msgResponse = MsgBox("Print Annual Statement as of " & Format(datLastYearEnd, "mmmm d, yyyy") & "?", _
                    vbQuestion + vbOKCancel, (Left("Annual Statement For Last Year" & Space(55), 55) & "F02"))
25810             blnMsgShown = True
25820             If msgResponse <> vbOK Then
25830               blnContinue2 = False
25840             Else
                    ' ** OK path...
25850               DoCmd.Hourglass True
25860               DoEvents
25870             End If
25880           Else
25890             If .cmbMonth = vbNullString Then
25900               .cmbMonth = Format(datLastYearEnd, "mmmm")
25910               DoEvents
25920               If IsNull(.StatementsYear) = True Then
25930                 .StatementsYear = year(datLastYearEnd)
25940                 DoEvents
25950               Else
25960                 If (CLng(.StatementsYear) > year(Date)) Or _
                          ((CLng(.StatementsYear) = year(Date) And Format(Date, "mm/dd") <> "12/31")) Then
25970                   blnContinue2 = False
25980                   DoCmd.Hourglass False
25990                   MsgBox "The Annual Statement can only be for a prior December 31st.", vbInformation + vbOKOnly, _
                          (Left("Invalid Entry" & Space(40), 40) & "F03")
26000                   .StatementsYear.SetFocus
26010                 Else
26020                   If CLng(.StatementsYear) < 1950& Then
26030                     blnContinue2 = False
26040                     DoCmd.Hourglass False
26050                     MsgBox "Please enter a valid year for the Annual Statement", vbInformation + vbOKOnly, _
                            (Left("Invalid Entry" & Space(40), 40) & "F04")
26060                     .StatementsYear.SetFocus
26070                   ElseIf CLng(.StatementsYear) <= (year(Date) - 5) Then
26080                     DoCmd.Hourglass False
26090                     msgResponse = MsgBox("Are you sure you want to print an Annual Statement for data this old?" & vbCrLf & vbCrLf & _
                            "December 31, " & CStr(.StatementsYear), vbQuestion + vbYesNo, _
                            (Left("Old Annual Statement" & Space(55), 55) & "F05"))
26100                     blnMsgShown = True
26110                     If msgResponse <> vbYes Then
26120                       blnContinue2 = False
26130                     Else
                            ' ** OK path...
26140                       DoCmd.Hourglass True
26150                       DoEvents
26160                     End If
26170                   Else
26180                     DoCmd.Hourglass False
26190                     msgResponse = MsgBox("Print Annual Statement as of December 31, " & CStr(.StatementsYear) & "?", _
                            vbQuestion + vbOKCancel, _
                            (Left(("Annual Statement For The Year " & CStr(.StatementsYear)) & Space(55), 55) & "F06"))
26200                     blnMsgShown = True
26210                     If msgResponse <> vbOK Then
26220                       blnContinue2 = False
26230                     Else
                            ' ** OK path...
26240                       DoCmd.Hourglass True
26250                       DoEvents
26260                     End If
26270                   End If
26280                 End If
26290               End If
26300             Else
26310               If .cmbMonth.Column(CBX_MON_NAME) <> "December" Then
26320                 blnContinue2 = False
26330                 DoCmd.Hourglass False
26340                 MsgBox "Annual Statement must be as of December.", vbInformation + vbOKOnly, _
                        (Left("Invalid Entry" & Space(40), 40) & "F07")
26350                 .cmbMonth.SetFocus
26360               Else
26370                 If IsNull(.StatementsYear) = True Then
26380                   .StatementsYear = year(datLastYearEnd)
26390                   DoEvents
                        ' ** OK path...
26400                 Else
26410                   If (CLng(.StatementsYear) > year(Date)) Or _
                            ((CLng(.StatementsYear) = year(Date) And Format(Date, "mm/dd") <> "12/31")) Then
26420                     blnContinue2 = False
26430                     DoCmd.Hourglass False
26440                     MsgBox "The Annual Statement can only be for a prior December 31st.", vbInformation + vbOKOnly, _
                            (Left("Invalid Entry" & Space(40), 40) & "F08")
26450                     .StatementsYear.SetFocus
26460                   Else
26470                     If CLng(.StatementsYear) < 1950& Then
26480                       blnContinue2 = False
26490                       DoCmd.Hourglass False
26500                       MsgBox "Please enter a valid year for the Annual Statement", vbInformation + vbOKOnly, _
                              (Left("Invalid Entry" & Space(40), 40) & "F09")
26510                       .StatementsYear.SetFocus
26520                     ElseIf CLng(.StatementsYear) <= (year(Date) - 5) Then
26530                       DoCmd.Hourglass False
26540                       msgResponse = MsgBox("Are you sure you want to print an Annual Statement for data this old?" & vbCrLf & vbCrLf & _
                              "December 31, " & CStr(.StatementsYear), vbQuestion + vbYesNo, _
                              (Left("Old Annual Statement" & Space(55), 55) & "F10"))
26550                       blnMsgShown = True
26560                       If msgResponse <> vbYes Then
26570                         blnContinue2 = False
26580                       Else
                              ' ** OK path...
26590                         DoCmd.Hourglass True
26600                         DoEvents
26610                       End If
26620                     Else
26630                       DoCmd.Hourglass False
26640                       msgResponse = MsgBox("Print Annual Statement as of December 31, " & CStr(.StatementsYear) & "?", _
                              vbQuestion + vbOKCancel, _
                              (Left(("Annual Statement For The Year " & CStr(.StatementsYear)) & Space(55), 55) & "F11"))
26650                       blnMsgShown = True
26660                       If msgResponse <> vbOK Then
26670                         blnContinue2 = False
26680                       Else
                              ' ** OK path...
26690                         DoCmd.Hourglass True
26700                         DoEvents
26710                       End If
26720                     End If
26730                   End If
26740                 End If
26750               End If
26760             End If
26770           End If

26780         End If  ' ** blnContinue2.

26790         If blnContinue2 = True Then

                ' ** Fields in cmbAccounts:
                ' **   Desc
                ' **   accountno
                ' **   predate
                ' **   shortname
                ' **   legalname
                ' **   BalanceDate

26800           datLastYearEnd = DateSerial(CLng(.StatementsYear), CLng(.cmbMonth.Column(CBX_MON_ID)), 31)
26810           datPrevYearEnd = DateSerial((CLng(.StatementsYear) - 1), CLng(.cmbMonth.Column(CBX_MON_ID)), 31)

26820           Set dbs = CurrentDb
26830           With dbs
26840             Select Case blnPrintAll
                  Case True
                    ' ** Account, linked to Balance, for smtdec = True, with HasBal.
26850               Set qdf = .QueryDefs("qryStatementAnnual_10")
26860             Case False
                    ' ** Balance, by specified [actno].
26870               Set qdf = .QueryDefs("qryStatementAnnual_01")
26880               With qdf.Parameters
26890                 ![actno] = strAccountNo
26900               End With
26910             End Select
26920             Set rst = qdf.OpenRecordset
26930             With rst
26940               If .BOF = True And .EOF = True Then
                      ' ** Can't really happen for All.
26950                 blnContinue2 = False
26960                 DoCmd.Hourglass False
26970                 lngBals = 0&
26980                 MsgBox "No previous balance was found for this account." & vbCrLf & vbCrLf & _
                        strAccountNo & "  " & strShortName, vbInformation + vbOKOnly, _
                        (Left("Nothing To Do" & Space(40), 55) & "F12")
26990               Else
27000                 .MoveLast
27010                 lngBals = .RecordCount
27020                 .MoveFirst
27030                 arr_varBal = .GetRows(lngBals)
                      ' ****************************************************
                      ' ** Array: arr_varBal()
                      ' **
                      ' **  Field  Element  Name                Constant
                      ' **  =====  =======  ==================  ==========
                      ' **    1       0     accountno           B_ACTNO
                      ' **    2       1     balance date        B_DATE
                      ' **    3       2     icash               B_ICASH
                      ' **    4       3     pcash               B_PCASH
                      ' **    5       4     cost                B_COST
                      ' **    6       5     TotalMarketValue
                      ' **    7       6     AccountValue
                      ' **    8       7     HasBal              B_BAL
                      ' **
                      ' ****************************************************
27040               End If
27050               .Close
27060             End With
27070             .Close
27080           End With

27090         End If  ' ** blnContinue2.

27100       End If
27110     End If
27120   End With  ' ** frm.

27130   If blnContinue2 = True Then
27140     If blnPrintAll = True Then

            ' ** First, collect the accounts.
27150       For lngX = 0& To (lngBals - 1&)
27160         strAccountNo = arr_varBal(B_ACTNO, lngX)
27170         blnFound = False
27180         For lngY = 0& To (lngAccts - 1&)
27190           If arr_varAcct(A_ACTNO, lngY) = strAccountNo Then
27200             blnFound = True
27210             Exit For
27220           End If
27230         Next
27240         If blnFound = False Then
27250           lngAccts = lngAccts + 1&
27260           lngE = lngAccts - 1&
27270           ReDim Preserve arr_varAcct(A_ELEMS, lngE)
27280           arr_varAcct(A_ACTNO, lngE) = strAccountNo
27290           arr_varAcct(A_BAL, lngE) = CBool(False)
27300           arr_varAcct(A_BELEM1, lngE) = lngX         ' ** First element number in arr_varBal().
27310           arr_varAcct(A_BCNT, lngE) = CLng(0)        ' ** Number of balance records.
27320           arr_varAcct(A_FDAT, lngE) = Null           ' ** First balance date.
27330           arr_varAcct(A_FDATZ, lngE) = CBool(False)  ' ** First balance is all zeroes.
27340           arr_varAcct(A_FELEM, lngE) = CLng(-1)      ' ** First balance element number. (Most likely same as A_BELEM1.)
27350           arr_varAcct(A_PDAT, lngE) = Null           ' ** Previous balance date.
27360           arr_varAcct(A_PDATZ, lngE) = CBool(False)  ' ** Previous balance is all zeroes.
27370           arr_varAcct(A_PELEM, lngE) = CLng(-1)      ' ** Previous balance element number.
27380           arr_varAcct(A_LDAT, lngE) = Null           ' ** Last balance date.
27390           arr_varAcct(A_LDATZ, lngE) = CBool(False)  ' ** Last balance is all zeroes.
27400           arr_varAcct(A_LELEM, lngE) = CLng(-1)      ' ** Last balance element number.
27410           arr_varAcct(A_ALLZ, lngE) = CBool(True)    ' ** All balances are zero.
27420           arr_varAcct(A_RUN, lngE) = CBool(False)    ' ** Run the statement.
27430         End If
27440       Next

27450       lngNoBals = 0&

            ' ** Now see which ones have balance records.
27460       For lngX = 0& To (lngAccts - 1&)
27470         If IsNull(arr_varBal(B_DATE, arr_varAcct(A_BELEM1, lngX))) = True Then
27480           lngNoBals = lngNoBals + 1&
27490           arr_varAcct(A_RUN, lngX) = CBool(False)
27500         Else
27510           strAccountNo = arr_varAcct(A_ACTNO, lngX)
27520           For lngY = arr_varAcct(A_BELEM1, lngX) To (lngBals - 1&)
27530             If arr_varBal(B_ACTNO, lngY) = strAccountNo Then
27540               arr_varAcct(A_BAL, lngX) = CBool(True)
27550               arr_varAcct(A_RUN, lngX) = CBool(True)
27560               arr_varAcct(A_BCNT, lngX) = arr_varAcct(A_BCNT, lngX) + 1&
                    ' ** Update first balance date.
27570               If IsNull(arr_varAcct(A_FDAT, lngX)) = True Then
27580                 arr_varAcct(A_FDAT, lngX) = arr_varBal(B_DATE, lngY)
27590                 arr_varAcct(A_FELEM, lngX) = lngY
27600               Else
27610                 If arr_varBal(B_DATE, lngY) < arr_varAcct(A_FDAT, lngX) Then
27620                   arr_varAcct(A_FDAT, lngX) = arr_varBal(B_DATE, lngY)
27630                   arr_varAcct(A_FELEM, lngX) = lngY
27640                 End If
27650               End If
                    ' ** Update last balance date.
27660               If arr_varBal(B_DATE, lngY) = datLastYearEnd Then
27670                 arr_varAcct(A_LDAT, lngX) = arr_varBal(B_DATE, lngY)
27680                 arr_varAcct(A_LELEM, lngX) = lngY
27690                 If arr_varBal(B_ICASH, lngY) <> 0 Or arr_varBal(B_PCASH, lngY) <> 0 Or arr_varBal(B_COST, lngY) <> 0 Then
27700                   arr_varAcct(A_LDATZ, lngX) = CBool(True)
27710                 End If
27720               End If
                    ' ** Update previous balance date.
27730               If arr_varBal(B_DATE, lngY) = datPrevYearEnd Then
27740                 arr_varAcct(A_PDAT, lngX) = arr_varBal(B_DATE, lngY)
27750                 arr_varAcct(A_PELEM, lngX) = lngY
27760                 If arr_varBal(B_ICASH, lngY) <> 0 Or arr_varBal(B_PCASH, lngY) <> 0 Or arr_varBal(B_COST, lngY) <> 0 Then
27770                   arr_varAcct(A_PDATZ, lngX) = CBool(True)
27780                 End If
27790               End If
27800               If arr_varBal(B_ICASH, lngY) <> 0 Or arr_varBal(B_PCASH, lngY) <> 0 Or arr_varBal(B_COST, lngY) <> 0 Then
27810                 arr_varAcct(A_ALLZ, lngX) = CBool(False)  ' ** Just one hit is enough.
27820               End If
27830             Else
27840               Exit For
27850             End If  ' ** strAcctountNo.
27860           Next  ' ** lngBals: lngY.
27870         End If  ' ** lngNoBals.
27880       Next  ' ** lngAccts: lngX.

            ' ** Check the first balance zero status.
27890       For lngX = 0& To (lngAccts - 1&)
27900         If arr_varAcct(A_ALLZ, lngX) = True Then
27910           arr_varAcct(A_FDATZ, lngX) = CBool(True)
27920           arr_varAcct(A_PDATZ, lngX) = CBool(True)
27930           arr_varAcct(A_LDATZ, lngX) = CBool(True)
27940         Else
27950           lngTmp01 = arr_varAcct(A_FELEM, lngX)
27960           If arr_varBal(B_ICASH, lngTmp01) = 0 And arr_varBal(B_PCASH, lngTmp01) = 0 And arr_varBal(B_COST, lngTmp01) = 0 Then
27970             arr_varAcct(A_FDATZ, lngX) = CBool(True)
27980           End If
27990         End If
28000       Next

28010     End If  ' ** blnPrintAll.
28020   End If  ' ** blnContinue2.

28030   If blnContinue2 = True Then

28040     Select Case blnPrintAll
          Case True

            ' ** The 2 arrays already have this info:
            ' **   blnLastIsZero
            ' **   blnPrevIsZero
            ' **   blnFirstIsZero
            ' **   datFirstYearEnd
            ' **   lngFirstElem
            ' **   blnPrev
            ' **   blnLast

            ' ** Check the criteria.
28050       For lngX = 0& To (lngAccts - 1&)
28060         blnLast = False: blnPrev = False: blnNoTrans = False
28070         If arr_varAcct(A_BAL, lngX) = True And arr_varAcct(A_RUN, lngX) = True Then
28080           strAccountNo = arr_varAcct(A_ACTNO, lngX)
28090           If IsNull(arr_varAcct(A_PDAT, lngX)) = False Then blnPrev = True
28100           If IsNull(arr_varAcct(A_LDAT, lngX)) = False Then blnLast = True
28110           blnPrevIsZero = arr_varAcct(A_PDATZ, lngX)
28120           datFirstYearEnd = Nz(arr_varAcct(A_FDAT, lngX), 0)
28130           If blnLast = False And blnPrev = False Then
28140             arr_varAcct(A_RUN, lngX) = CBool(False)
28150           ElseIf blnLast = False And blnPrev = True Then
28160             If blnPrevIsZero = True Then  ' ** Let the 'Else' go through.
28170               arr_varAcct(A_RUN, lngX) = CBool(False)
28180             End If
28190           ElseIf blnLast = True And blnPrev = False Then
28200             If year(datLastYearEnd) = year(datFirstYearEnd) Then  ' ** Even if they're the same date.
28210               blnContinue2 = AnnualStatement_PrevTrans(strAccountNo, datLastYearEnd, blnPrintAll)  ' ** Module Function: modStatementParamFuncs1.
28220               blnNoTrans = Not (blnContinue2)  ' ** They're opposite.
28230               blnContinue2 = True
28240               If blnNoTrans = True Then
28250                 arr_varAcct(A_RUN, lngX) = CBool(False)
28260               End If
28270             Else
28280               arr_varAcct(A_RUN, lngX) = CBool(False)
28290             End If
28300           End If  ' ** blnLast, blnPrev.
28310         End If  ' ** A_RUN.
28320       Next  ' ** lngX

28330       lngRuns = 0&
28340       For lngX = 0& To (lngAccts - 1&)
28350         If arr_varAcct(A_RUN, lngX) = True Then
28360           lngRuns = lngRuns + 1&
28370         End If
28380       Next

            ' ** Let the user know the counts.
28390       If lngRuns < lngAccts Then
28400         If lngRuns = 0& Then
28410           If lngAccts = 1& Then
28420             strMsg1 = "The Annual Statement cannot be run."
28430             If lngAcctsCnt > 1& Then
28440               strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                      "only 1 is scheduled for a December statement."
28450             End If
28460           Else
28470             strMsg1 = "Annual Statements cannot be run."
28480             If lngAcctsCnt > lngAcctsScheduled Then
28490               strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                      CStr(lngAcctsScheduled) & " are scheduled for December statements."
28500             End If
28510           End If
28520         ElseIf lngRuns = 1& Then
28530           If lngAccts = 2& Then
28540             strMsg1 = "An Annual Statement for one of the two accounts will be run."
28550             If lngAcctsCnt > 2& Then
28560               If lngAcctsScheduled = 1& Then
28570                 strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                        "only 1 is scheduled for a December statement."
28580               Else
28590                 strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                        CStr(lngAcctsScheduled) & " are scheduled for December statements."
28600               End If
28610             End If
28620           Else
28630             strMsg1 = "An Annual Statement for one of the " & CStr(lngAccts) & " will be run."
28640             If lngAcctsScheduled = 1& Then
28650               strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                      "only 1 is scheduled for a December statement."
28660             Else
28670               strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                      CStr(lngAcctsScheduled) & " are scheduled for December statements."
28680             End If
28690           End If
28700         ElseIf lngRuns = 2& Then
28710           strMsg1 = "Annual Statements for two of the " & CStr(lngAccts) & " will be run."
28720           strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                  CStr(lngAcctsScheduled) & " are scheduled for December statements."
28730         Else
28740           strMsg1 = CStr(lngRuns) & " Annual Statements can be run, out of " & CStr(lngAccts) & " accounts."
28750           strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                  CStr(lngAcctsScheduled) & " are scheduled for December statements."
28760         End If
28770         If lngNoBals > 0& Then
28780           If (lngAccts - lngRuns) = lngNoBals Then
28790             strMsg2 = "The remaining " & CStr(lngNoBals) & " account" & _
                    IIf(lngNoBals = 1, " has", "s have") & " no Statement Balance records."
28800           Else
28810             strMsg2 = "Of the remaining " & CStr(lngAccts - lngRuns) & " accounts, " & _
                    IIf(lngNoBals = 1, "1 has", CStr(lngNoBals) & " have") & " no balance records, and " & _
                    IIf(((lngAccts - lngRuns) - lngNoBals) = 1, "1 has", _
                    CStr((lngAccts - lngRuns) - lngNoBals) & " have") & " insufficient data."
28820           End If
28830         Else
28840           If lngRuns = 0& Then
28850             If lngAccts = 1& Then
28860               strMsg2 = "The account has insufficient data."
28870             Else
28880               strMsg2 = "The " & CStr(lngAccts) & " accounts have insufficient data."
28890             End If
28900           Else
28910             strMsg2 = "The remaining " & CStr(lngAccts - lngRuns) & " account" & _
                    IIf((lngAccts - lngRuns) = 1, " has", "s have") & " insufficient data."
28920           End If
28930         End If
28940       Else
28950         If lngAccts = 1& Then
28960           strMsg1 = "An Annual Statement for the account will be run."
28970           If lngAcctsCnt > 1& Then
28980             If lngAcctsScheduled = 1& Then
28990               strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                      "only 1 is scheduled for a December statement."
29000             Else
29010               strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                      CStr(lngAcctsScheduled) & " are scheduled for December statements."
29020             End If
29030           End If
29040         ElseIf lngAccts = 2& Then
29050           strMsg1 = "Annual Statements for both accounts will be run."
29060           If lngAcctsCnt > 2& Then
29070             strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                    CStr(lngAcctsScheduled) & " are scheduled for December statements."
29080           End If
29090         Else
29100           strMsg1 = "Annual Statements for all " & CStr(lngAccts) & " accounts will be run."
29110           If lngAcctsCnt > lngAccts Then
29120             strMsg1 = strMsg1 & vbCrLf & "Of " & CStr(lngAcctsCnt) & " accounts, " & _
                    CStr(lngAcctsScheduled) & " are scheduled for December statements."
29130           End If
29140         End If
29150         strMsg2 = vbNullString
29160       End If

            ' **  Don't I just carry on...
29170       If lngRuns > 0& Then
29180         If strMsg2 = vbNullString Then
29190           DoCmd.Hourglass False
29200           msgResponse = MsgBox(strMsg1 & vbCrLf & vbCrLf & "Click OK to proceed.", _
                  vbExclamation + vbOKCancel, Left(("Run Annual Statements" & Space(55)), 55) & "F13")
29210         Else
29220           DoCmd.Hourglass False
29230           msgResponse = MsgBox(strMsg1 & vbCrLf & vbCrLf & strMsg2 & vbCrLf & vbCrLf & "Click OK to proceed.", _
                  vbExclamation + vbOKCancel, Left(("Run Annual Statements" & Space(55)), 55) & "F14")
29240         End If
29250         If msgResponse <> vbOK Then
29260           blnContinue2 = False
29270         Else
                ' ** OK path...
29280           DoCmd.Hourglass True
29290           DoEvents
29300         End If
29310       Else
29320         blnContinue2 = False
29330         DoCmd.Hourglass False
29340         msgResponse = MsgBox(strMsg1 & vbCrLf & vbCrLf & strMsg2, _
                vbInformation + vbOKOnly, Left(("Run Annual Statements" & Space(55)), 55) & "F15")
29350       End If  ' ** lngRuns.

29360     Case False

29370       blnLast = False: blnPrev = False
29380       blnLastIsZero = False: blnPrevIsZero = False: blnFirstIsZero = False
29390       lngFirstElem = -1&
29400       datFirstYearEnd = Date
29410       For lngX = 0& To (lngBals - 1&)
29420         If arr_varBal(B_DATE, lngX) = datLastYearEnd Then
29430           blnLast = True
29440           If arr_varBal(B_ICASH, lngX) = 0@ And arr_varBal(B_PCASH, lngX) = 0@ And arr_varBal(B_COST, lngX) = 0@ Then
29450             blnLastIsZero = True
29460           End If
29470         ElseIf arr_varBal(B_DATE, lngX) = datPrevYearEnd Then
29480           blnPrev = True
29490           If arr_varBal(B_ICASH, lngX) = 0@ And arr_varBal(B_PCASH, lngX) = 0@ And arr_varBal(B_COST, lngX) = 0@ Then
29500             blnPrevIsZero = True
29510           End If
29520         End If
29530         If arr_varBal(B_DATE, lngX) < datFirstYearEnd Then
                ' ** This could be any day of the year.
29540           datFirstYearEnd = arr_varBal(B_DATE, lngX)
29550           lngFirstElem = lngX
29560         End If
29570       Next

29580       If arr_varBal(B_ICASH, lngFirstElem) = 0@ And arr_varBal(B_PCASH, lngFirstElem) = 0@ And arr_varBal(B_COST, lngFirstElem) = 0@ Then
29590         blnFirstIsZero = True
29600       End If

            ' ** THIS IS WAY OUT-OF-HAND!!!
29610       If blnLast = False And blnPrev = False Then
29620         blnContinue2 = False
29630         DoCmd.Hourglass False
29640         MsgBox "No balance found for the specified year," & vbCrLf & _
                "nor the year prior to it." & vbCrLf & vbCrLf & _
                "Annual Statement cannot be run.", vbInformation + vbOKOnly, _
                (Left("Insufficient Data" & Space(55), 55) & "F16")
29650       ElseIf blnLast = False And blnPrev = True Then
29660         If blnPrevIsZero = True Then
                ' ** Only Zeroes, so we can't do it.
29670           blnContinue2 = False
29680           DoCmd.Hourglass False
29690           If blnMsgShown = False Then
29700             MsgBox "No balance found for last December 31st," & vbCrLf & _
                    "and the previous year's balance was zero." & vbCrLf & vbCrLf & _
                    "Annual Statement cannot be run.", vbInformation + vbOKOnly, _
                    (Left("Insufficient Data" & Space(55), 55) & "F17")
29710           Else
29720             MsgBox "No balance found for the specified year," & vbCrLf & _
                    "and the balance was zero for the year prior to that." & vbCrLf & vbCrLf & _
                    "Annual Statement cannot be run.", vbInformation + vbOKOnly, _
                    (Left("Insufficient Data" & Space(55), 55) & "F18")
29730           End If
29740         Else
29750           If blnMsgShown = False Then
29760             DoCmd.Hourglass False
29770             msgResponse = MsgBox("The latest year-end balance is over a year old." & vbCrLf & _
                    "Do you want to run a statement for the year ending " & CStr(year(datPrevYearEnd)) & "?", _
                    vbInformation + vbYesNo, _
                    (Left("Old Year-End Balance" & Space(55), 55) & "F19"))
29780             blnMsgShown = True
29790           Else
29800             msgResponse = vbYes
29810           End If
29820           If msgResponse <> vbYes Then
29830             blnContinue2 = False
29840           Else
                  ' ** OK path...
29850             DoCmd.Hourglass True
29860             DoEvents
29870           End If
29880         End If
29890       ElseIf blnLast = True And blnPrev = False Then
              ' ** If last year was their first year, it'll still work.
29900         If year(datLastYearEnd) = year(datFirstYearEnd) Then  ' ** Even if they're the same date.
                ' ** See if there are transactions prior to the date.

29910           blnContinue2 = AnnualStatement_PrevTrans(strAccountNo, datLastYearEnd, blnPrintAll)  ' ** Module Function: modStatementParamFuncs1.

29920         Else  ' ** year(datLastYearEnd) <> year(datFirstYearEnd).
                ' ** Prior data covers multiple years.
29930           blnContinue2 = False
29940           DoCmd.Hourglass False
29950           MsgBox "No previous year-end balance exists." & vbCrLf & vbCrLf & _
                  "Annual Statement cannot be run.", vbInformation + vbOKOnly, _
                  (Left("Start-of-Year Balance Not Found" & Space(55), 55) & "F20")
29960         End If
29970       Else  ' ** blnLast = True And blnPrev = True
              ' ** Balances exist, but are they valid?  I DON'T CARE!
29980       End If  ' ** blnLast, blnPrev.

29990     End Select  ' ** blnPrintAll.

30000   End If  ' ** blnContinue2.

30010   If blnContinue2 = True Then
30020     With frm

30030       .PrintAnnual_chk = True
30040       Select Case blnPrintAll
            Case True
30050         .PrintAnnual_cnt = lngRuns
30060       Case False
30070         .PrintAnnual_cnt = 1&
30080       End Select
30090       DoEvents

30100       Beep
30110       MsgBox "Note: The Annual Statement does not update the year-end balance." & vbCrLf & vbCrLf & _
              "Values are only reported. " & _
              "You must still run periodic statements in order to properly generate year-end balances.", _
              vbInformation + vbOKOnly, "Annual Statement Report"

30120       blnPrintAnnualStatement = True
30130       Select Case blnPrintAll
            Case True
30140         .cmbAccounts.ForeColor = WIN_CLR_DISF
30150         .cmbAccounts.BackColor = WIN_CLR_DISB
30160         .cmbAccounts_lbl.ForeColor = WIN_CLR_DISF
30170         .cmbAccounts.Locked = True
30180         .cmbAccounts.Enabled = True
30190         .cmbAccounts.BorderColor = CLR_LTBLU2
30200         .cmbAccounts.BackStyle = acBackStyleNormal
30210         DoEvents
30220         .cmdAnnualStatement.SetFocus
30230         blnOpgChanged = False
30240         For lngX = 0& To (lngAccts - 1&)
30250           If arr_varAcct(A_RUN, lngX) = True Then
30260             gstrAccountNo = arr_varAcct(A_ACTNO, lngX)
                  ' ** This is in an effort to suppress errant reports, and assure that just the right reports get printed.
30270             glngPrintRpts = glngPrintRpts + 1&
30280             lngE = glngPrintRpts - 1&
30290             ReDim Preserve garr_varPrintRpt(PR_ELEMS, lngE)
30300             garr_varPrintRpt(PR_ACTNO, lngE) = gstrAccountNo
30310             garr_varPrintRpt(PR_ALIST, lngE) = CBool(False)
30320             garr_varPrintRpt(PR_TRANS, lngE) = CBool(False)
30330             garr_varPrintRpt(PR_SUMRY, lngE) = CBool(False)
30340             If .opgAccountNumber = .opgAccountNumber_optAll.OptionValue Then
30350               .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue
30360               .opgAccountNumber_optSpecified_lbl_box.Visible = True
30370               .opgAccountNumber_optAll_lbl_box.Visible = False
30380               blnOpgChanged = True
30390             End If
30400             .cmbAccounts = arr_varAcct(A_ACTNO, lngX)
30410             DoEvents
                  ' ** 1st Annual Statement branching.
                  'blnContinue2 = PrintStatements(True)  ' ** Procedure: Below.
30420             blnContinue2 = Statements_Print(frm, blnPrintStatements, blnAllStatements, blnSingleStatement, _
                    blnRunPriorStatement, blnAcctsSchedRpt, datFirstDate, blnContinue, blnFromStmts, _
                    blnGoingToReport, blnGoingToReport2, blnGTR_Emblem, blnWasGTR, True)  ' ** Module Function: modStatementParamFuncs1.
30430             If blnContinue2 = False Then
30440               Exit For
30450             End If
30460             .cmdDevCloseReports_Click  ' ** Form Procedure: frmStatementParameters.
30470             DoEvents
30480           End If
30490         Next
30500         If blnContinue2 = True Then
30510           .cmdAnnualStatement.SetFocus
30520           .cmbAccounts = Null
30530           .cmbAccounts.Locked = False
30540           .cmbAccounts.ForeColor = CLR_BLK
30550           .cmbAccounts.BackColor = CLR_WHT
30560           DoEvents
30570           .opgAccountNumber = .opgAccountNumber_optAll.OptionValue  '#Covered.
30580           .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
30590           DoEvents
30600           .cmdAnnualStatement.SetFocus
30610           If lngRuns > 1& Then
30620             strMsg1 = "s"
30630             If Reports.Count > 0& Then  ' ** If any reports are on the screen, close 'em.
30640               Do While Reports.Count > 0&
30650                 DoCmd.Close acReport, Reports(0).Name
30660                 DoEvents
30670               Loop
30680             End If
30690           Else
30700             strMsg1 = vbNullString
30710           End If
30720           If .PrintAnnual_chk = True Then
30730             Beep
30740             DoCmd.Hourglass False
30750             MsgBox "Annual Statement" & strMsg1 & " done for period ending " & _
                    .DateEnd & "." & vbCrLf & vbCrLf & _
                    CStr(lngRuns) & " Statement" & strMsg1 & " processed.", _
                    vbInformation + vbOKOnly, (Left(("Statement" & strMsg1 & " Finished" & Space(55)), 55) & "F21")
30760           End If
30770         End If  ' ** blnContinue2.
30780       Case False
30790         gstrAccountNo = .cmbAccounts
              'blnContinue2 = PrintStatements(True)  ' ** Function: Below.
30800         blnContinue2 = Statements_Print(frm, blnPrintStatements, blnAllStatements, blnSingleStatement, _
                blnRunPriorStatement, blnAcctsSchedRpt, datFirstDate, blnContinue, blnFromStmts, _
                blnGoingToReport, blnGoingToReport2, blnGTR_Emblem, blnWasGTR, True)  ' ** Module Function: modStatementParamFuncs1.
30810       End Select

30820       .cmdDevCloseReports_Click  ' ** Form Procedure: frmStatementParameters.
30830       .PrintAnnual_chk = False

30840     End With
30850   End If  ' ** blnContinue2.

30860   blnPrintAnnualStatement = False
30870   gstrAccountNo = vbNullString

30880   DoCmd.Hourglass False

EXITP:
30890   Set rst = Nothing
30900   Set qdf = Nothing
30910   Set dbs = Nothing
30920   Exit Sub

ERRH:
4200    DoCmd.Hourglass False
4210    frm.PrintAnnual_chk = False
4220    frm.cmbAccounts.Locked = False
4230    Select Case ERR.Number
        Case Else
4240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250    End Select
4260    Resume EXITP

End Sub

Public Sub Transactions_Excel(frm As Access.Form, blnContinue As Boolean, blnFromStmts As Boolean, strFirstDateMsg As String, strFileName As String, strReportName As String, blnAllStatements As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean)

31000 On Error GoTo ERRH

        Const THIS_PROC As String = "Transactions_Excel"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim blnRetVal As Boolean

        '#####################
        'NOT CURRENCY READY!
        '#####################

31010   With frm
31020     blnRetVal = True
31030     If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
31040       MsgBox "You must select an account to continue," & vbCrLf & _
              "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "L01")
31050     Else

31060       If .chkStatements = True Then
31070         If IsNull(.cmbMonth) Then
31080           blnRetVal = False
31090           MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                  (Left(("Entry Required" & Space(55)), 55) & "L02")
31100           .cmbMonth.SetFocus
31110         Else
31120           If .cmbMonth = vbNullString Then
31130             blnRetVal = False
31140             MsgBox "You must select a report month to continue.", vbInformation + vbOKOnly, _
                    (Left(("Entry Required" & Space(55)), 55) & "L03")
31150             .cmbMonth.SetFocus
31160           Else
31170             If IsNull(.StatementsYear) = True Then
31180               blnRetVal = False
31190               MsgBox "You must enter a report year to continue.", vbInformation + vbOKOnly, _
                      (Left(("Entry Required" & Space(55)), 55) & "L04")
31200               .StatementsYear.SetFocus
31210             End If
31220           End If
31230         End If
31240       Else
31250         If FirstDate_SP(frm) = False Then  ' ** Module Function: modStatementParamFuncs2.
31260           blnRetVal = False
31270           MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "L05")
31280         End If
31290       End If

31300       If blnRetVal = True Then

31310         DoCmd.Hourglass True
31320         DoEvents

31330         strRptCap = vbNullString: strRptPathFile = vbNullString

              'If .chkStatements = True, then just those scheduled!

              ' ** Execute the common code.
31340         blnRetVal = BuildTransactionInfo_SP(frm, strFileName, strReportName, blnAllStatements, blnContinue, blnHasForEx, blnHasForExThis, blnFromStmts, "Excel")  ' ** Module Function: modStatementParamFuncs1.

31350         If blnRetVal = True Then
31360           If blnContinue = True Then

31370             If IsNull(.UserReportPath) = True Then
31380               strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
31390             Else
31400               strRptPath = .UserReportPath
31410             End If
31420             Select Case .opgAccountNumber
                  Case .opgAccountNumber_optSpecified.OptionValue
31430               strRptCap = "rptTransaction_Statement_" & .cmbAccounts.Column(CBX_A_ACTNO) & "_"
31440             Case .opgAccountNumber_optAll.OptionValue
31450               strRptCap = "rptTransaction_Statement_All_"
31460             End Select
31470             strRptCap = StringReplace(strRptCap, "/", "_")  ' ** Module Function: modStringFuncs.
31480             If .chkTransactions = True Then
31490               strRptCap = strRptCap & "_" & Format(.TransDateStart, "yymmdd") & "-" & Format(.TransDateEnd, "yymmdd")
31500             ElseIf .chkStatements = True Then
31510               Select Case .opgOrderBy
                    Case .opgOrderBy_optDate.OptionValue
31520                 strRptCap = strRptCap & "Date_"
31530               Case .opgOrderBy_optType.OptionValue
31540                 strRptCap = strRptCap & "Type_"
31550               End Select
31560               strRptCap = strRptCap & Format(.DateEnd, "yymmdd")
31570             End If

31580             If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
31590               Select Case .opgOrderBy
                    Case .opgOrderBy_optDate.OptionValue
31600                 strQry = "rptTransaction_Statement_SortDate"
31610               Case .opgOrderBy_optType.OptionValue
31620                 strQry = "rptTransaction_Statement_SortType"
31630               End Select
31640               DoCmd.OpenReport strQry, acViewPreview
31650             Else

31660               strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

31670               If strRptPathFile <> vbNullString Then

31680                 strQry = vbNullString
31690                 Select Case .opgOrderBy
                      Case .opgOrderBy_optDate.OptionValue
31700                   If .chkTransactions = True Then
                          ' ** qryStatementParameters_Trans_01_02 (Ledger, linked to Account, qryStatementParameters_Trans_09a
                          ' ** (Balance table, by specified FormRef('EndDate')), with add'l fields, by specified
                          ' ** FormRef('EndDate')), For Export, by Date, Transactions only.
31710                     strQry = "qryStatementParameters_Trans_03a"
31720                   ElseIf .chkStatements = True Then
31730                     Set dbs = CurrentDb
                          ' ** Empty tmpEdit11.
31740                     Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_06l_01")
31750                     qdf.Execute
31760                     Set qdf = Nothing
31770                     DoEvents
                          ' ** I CAN'T UNDERSTAND WHY IT SAID THERE WERE TOO MANY DATABASES OPEN!
                          ' ** Append qryStatementParameters_Trans_06e_01 (Union of qryStatementParameters_Trans_06a
                          ' ** (qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01 (Union of
                          ' ** qryStatementParameters_Trans_15_05 (Ledger, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just
                          ' ** those matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate"))),
                          ' ** by Date, Statements), linked to qryStatementParameters_Trans_05c (qryStatementParameters_Trans_05b
                          ' ** (qryStatementParameters_Trans_05a (qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01
                          ' ** (Union of qryStatementParameters_Trans_15_05 (Ledger, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate"))), by Date,
                          ' ** Statements), grouped by accountno, with Min(transdate)), linked back to qryStatementParameters_Trans_04a
                          ' ** (qryStatementParameters_Trans_01_01 (Union of qryStatementParameters_Trans_15_05 (Ledger, linked to
                          ' ** Account, qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account, by specified
                          ' ** GlobalVarGet("glngMonthID")), just those matching MonthNum), qryStatementParameters_Trans_15_03 (Balance,
                          ' ** grouped by accountno, with Max(balance date), by specified FormRef('MaxBalDate')), by specified
                          ' ** GlobalVarGet("gdatEndDate")), qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account,
                          ' ** qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account, by specified
                          ' ** GlobalVarGet("glngMonthID")), just those matching MonthNum), qryStatementParameters_Trans_15_03 (Balance,
                          ' ** grouped by accountno, with Max(balance date), by specified FormRef('MaxBalDate')), by specified
                          ' ** GlobalVarGet("gdatEndDate"))), by Date, Statements), grouped by accountno, transdate, with Min(sortOrder)),
                          ' ** linked back to qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01 (Union of
                          ' ** qryStatementParameters_Trans_15_05 (Ledger, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate"))),
                          ' ** by Date, Statements), grouped by accountno, transdate, sortOrder, with Min(journalno), by Date), with accountnox,
                          ' ** shortnamex, journaltypex, ratex), qryStatementParameters_Trans_06e (qryStatementParameters_Trans_06d
                          ' ** (Union of qryStatementParameters_Trans_06b (qryStatementParameters_Trans_05c (qryStatementParameters_Trans_05b
                          ' ** (qryStatementParameters_Trans_05a (qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01
                          ' ** (Union of qryStatementParameters_Trans_15_05 (Ledger, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those linked
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with Max(balance date),
                          ' ** by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate"))), by Date, Statements), grouped
                          ' ** by accountno, with Min(transdate)), back to qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01
                          ' ** (Union of qryStatementParameters_Trans_15_05 (Ledger, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate"))),
                          ' ** by Date, Statements), grouped by accountno, transdate, with Min(sortOrder)), linked back to
                          ' ** qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01 (Union of qryStatementParameters_Trans_15_05
                          ' ** (Ledger, linked to Account, qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account,
                          ' ** by specified GlobalVarGet("glngMonthID")), just those matching MonthNum), qryStatementParameters_Trans_15_03
                          ' ** (Balance, grouped by accountno, with Max(balance date), by specified FormRef('MaxBalDate')), by specified
                          ' ** GlobalVarGet("gdatEndDate")), qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account,
                          ' ** qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account, by specified
                          ' ** GlobalVarGet("glngMonthID")), just those matching MonthNum), qryStatementParameters_Trans_15_03
                          ' ** (Balance, grouped by accountno, with Max(balance date), by specified FormRef('MaxBalDate')), by
                          ' ** specified GlobalVarGet("gdatEndDate"))), by Date, Statements), grouped by accountno, transdate, sortOrder,
                          ' ** with Min(journalno), by Date), with add'l fields, Beginning Balance), qryStatementParameters_Trans_06c
                          ' ** (qryStatementParameters_Trans_06b (qryStatementParameters_Trans_05c (qryStatementParameters_Trans_05b
                          ' ** (qryStatementParameters_Trans_05a (qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01
                          ' ** (Union of qryStatementParameters_Trans_15_05 (Ledger, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate"))), by Date,
                          ' ** Statements), grouped by accountno, with Min(transdate)), linked back to qryStatementParameters_Trans_04a
                          ' ** (qryStatementParameters_Trans_01_01 (Union of qryStatementParameters_Trans_15_05 (Ledger, linked to Account,
                          ' ** qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account, by specified
                          ' ** GlobalVarGet("glngMonthID")), just those matching MonthNum), qryStatementParameters_Trans_15_03 (Balance,
                          ' ** grouped by accountno, with Max(balance date), by specified FormRef('MaxBalDate')), by specified
                          ' ** GlobalVarGet("gdatEndDate")), qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account,
                          ' ** qryStatementParameters_Trans_15_02 (qryStatementParameters_Trans_15_01 (Account, by specified
                          ' ** GlobalVarGet("glngMonthID")), just those matching MonthNum), qryStatementParameters_Trans_15_03
                          ' ** (Balance, grouped by accountno, with Max(balance date), by specified FormRef('MaxBalDate')), by specified
                          ' ** GlobalVarGet("gdatEndDate"))), by Date, Statements), grouped by accountno, transdate, with Min(sortOrder)),
                          ' ** linked back to qryStatementParameters_Trans_04a (qryStatementParameters_Trans_01_01 (Union of
                          ' ** qryStatementParameters_Trans_15_05 (Ledger, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID")), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate")),
                          ' ** qryStatementParameters_Trans_15_06 (LedgerArchive, linked to Account, qryStatementParameters_Trans_15_02
                          ' ** (qryStatementParameters_Trans_15_01 (Account, by specified GlobalVarGet("glngMonthID"), just those
                          ' ** matching MonthNum), qryStatementParameters_Trans_15_03 (Balance, grouped by accountno, with
                          ' ** Max(balance date), by specified FormRef('MaxBalDate')), by specified GlobalVarGet("gdatEndDate"))),
                          ' ** by Date, Statements), grouped by accountno, transdate, sortOrder, with Min(journalno), by Date),
                          ' ** with add'l fields, Beginning Balance), blank line after)), sorted)) to tmpEdit11.
31780                     Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_06e_02")
31790                     qdf.Execute
31800                     Set qdf = Nothing
31810                     DoEvents
                          ' ** Append qryStatementParameters_Trans_06k_05 (xx) to tmpEdit11.
31820                     Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_06k_06")
31830                     qdf.Execute
31840                     Set qdf = Nothing
31850                     dbs.Close
31860                     Set dbs = Nothing
31870                     DoEvents
                          ' ** qryStatementParameters_Trans_07d (Union of qryStatementParameters_Trans_07a
                          ' ** (qryStatementParameters_Trans_06l (tmpEdit11), with some new field names),
                          ' ** qryStatementParameters_Trans_07b (Report title), qryStatementParameters_Trans_07c
                          ' ** (Report period), by Date), For Export, by Date, Statements.
31880                     strQry = "qryStatementParameters_Trans_08a"
                          ' ** qryStatementParameters_Trans_08c seems unneeded.
31890                   End If
31900                 Case .opgOrderBy_optType.OptionValue
31910                   If .chkTransactions = True Then
                          ' ** qryStatementParameters_Trans_01_02 (Ledger, linked to Account,
                          ' ** qryStatementParameters_Trans_09a (xx), with add'l fields, by
                          ' ** specified FormRef('EndDate')), For Export, by Type, Transactions only.
31920                     strQry = "qryStatementParameters_Trans_03b"
31930                   ElseIf .chkStatements = True Then
31940                     Set dbs = CurrentDb
                          ' ** Empty tmpEdit12.
31950                     Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_06x_01")
31960                     qdf.Execute
31970                     Set qdf = Nothing
31980                     DoEvents
                          ' ** Append qryStatementParameters_Trans_06q_01 (xx) to tmpEdit12.
31990                     Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_06q_02")
32000                     qdf.Execute
32010                     Set qdf = Nothing
32020                     DoEvents
                          ' ** Append qryStatementParameters_Trans_06w_05 (xx) to tmpEdit12.
32030                     Set qdf = dbs.QueryDefs("qryStatementParameters_Trans_06w_06")
32040                     qdf.Execute
32050                     Set qdf = Nothing
32060                     DoEvents
32070                     dbs.Close
32080                     Set dbs = Nothing
                          ' ** qryStatementParameters_Trans_07h (Union of qryStatementParameters_Trans_07e
                          ' ** (qryStatementParameters_Trans_06x (tmpEdit12), with some new field names),
                          ' ** qryStatementParameters_Trans_07f (Report title), qryStatementParameters_Trans_07g
                          ' ** (Report period), by Type), For Export, by Type, Statements.
32090                     strQry = "qryStatementParameters_Trans_08b"
                          ' ** qryStatementParameters_Trans_08d seems unneeded.
32100                   End If
32110                 End Select

32120                 If .chkTransactions = True Then

32130                   DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
32140                   .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

32150                 ElseIf .chkStatements = True Then

32160                   If FileExists(CurrentAppPath & LNK_SEP & "TransStmt_xxx.xls") = True Then  ' ** Module Functions: modFileUtilities.
32170                     Kill (CurrentAppPath & LNK_SEP & "TransStmt_xxx.xls")
32180                   End If

32190                   If CurDir <> CurrentAppPath Then  ' ** Module Function: modFileUtilities.
                          ' ** Since I'm not specifying a path in the macro, I want to make sure it's here.
32200                     ChDir CurrentAppPath  ' ** Module Function: modFileUtilities.
32210                   End If

32220                   Select Case .opgOrderBy
                        Case .opgOrderBy_optDate.OptionValue
                          ' ** qryStatementParameters_Trans_08a.
32230                     DoCmd.RunMacro "mcrExcelExport_TransStmt_All_Date_01"
32240                   Case .opgOrderBy_optType.OptionValue
                          ' ** qryStatementParameters_Trans_08b
32250                     DoCmd.RunMacro "mcrExcelExport_TransStmt_All_Type_01"
32260                   End Select

                        ' ** The macro specifies qryStatementParameters_Trans_08x, but cannot be given a dynamic file name.
                        ' ** So, it's exported to 'TransStmt_xxx.xls', which is then renamed.
32270                   If FileExists(CurrentAppPath & LNK_SEP & "TransStmt_xxx.xls") = True Then  ' ** Module Functions: modFileUtilities.
32280                     If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
32290                       Kill strRptPathFile
32300                     End If
32310                     Name (CurrentAppPath & LNK_SEP & "TransStmt_xxx.xls") As (strRptPathFile)  ' ** Module Function: modFileUtilities.
32320                     Excel_Trans strRptPathFile  ' ** Module Function: modExcelFuncs.
32330                     DoEvents
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
32340                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
32350                   End If

32360                 End If
32370               End If

32380             End If  ' ** gblnDev_Debug.

32390           End If  ' ** blnContinue.
32400         End If  ' ** blnRetVal.

32410       End If  ' ** blnRetVal
32420     End If
32430   End With

32440   DoCmd.Hourglass False

EXITP:
32450   Set qdf = Nothing
32460   Set dbs = Nothing
32470   DoCmd.SetWarnings True
32480   Exit Sub

ERRH:
4200    DoCmd.Hourglass False
4210    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
4220      blnContinue = False
4230    Case Else
4240      Select Case ERR.Number
          Case Else
4250        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4260      End Select
4270    End Select
4280    Resume EXITP

End Sub

Public Sub AssetList_Excel_SP(frm As Access.Form, blnContinue As Boolean, strFirstDateMsg As String, datAssetListDate As Date, blnPrintAnnualStatement As Boolean, blnAllStatements As Boolean, blnNoDataAll As Boolean, blnRollbackNeeded As Boolean, blnHasForExClick As Boolean, blnHasForEx As Boolean, blnHasForExThis As Boolean)

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetList_Excel_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQry1 As String, strQry2 As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim intOptGrpAcctNum As Integer
        Dim blnPriceHistory As Boolean
        Dim strDocName As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

32510   With frm
32520     If blnHasForExClick = False Then

32530       blnContinue = True
32540       blnRetVal = True

32550       If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue And IsNull(.cmbAccounts) = True Then
32560         blnContinue = False
32570         MsgBox "You must select an account to continue," & vbCrLf & _
                "or choose All for Account.", vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "P01")
32580       Else

32590         If .chkAssetList = True Then
32600           If FirstDate_SP(frm) = False Then  ' ** Module Function: modStatementParamFuncs2.
32610             blnRetVal = False
32620             MsgBox strFirstDateMsg, vbInformation + vbOKOnly, (Left(("Nothing To Do" & Space(40)), 40) & "P02")
32630           End If
32640         End If

32650         If blnRetVal = True Then

32660           intOptGrpAcctNum = .opgAccountNumber
32670           gblnCombineAssets = .chkCombineCash.Value
32680           strRptCap = vbNullString: strRptPathFile = vbNullString

32690           Select Case .opgAccountNumber
                Case .opgAccountNumber_optSpecified.OptionValue
32700             gstrAccountNo = .cmbAccounts
32710           Case .opgAccountNumber_optAll.OptionValue
32720             gstrAccountNo = "All"
32730           End Select
32740           blnIncludeCurrency = .chkIncludeCurrency
32750           datAssetListDate = .AssetListDate

32760           If .chkForeignExchange = True Then
32770             .currentDate = Null
32780             blnPriceHistory = PricingHistory(datAssetListDate)  ' ** Module Function: modStatementParamFuncs2.
                  ' ** blnPriceHistory indicates whether current pricing or pricing history should be used.
                  ' ** It ONLY applies to foreign exchange, since regular reports don't require that info.
32790             .UsePriceHistory = blnPriceHistory
32800           End If

32810           If .opgAccountNumber = .opgAccountNumber_optSpecified.OptionValue Then
32820             If blnHasForEx = True And .UsePriceHistory = True Then
                    ' ** If all the rest of this code is using the foreign currency tables and queries,
                    ' ** but the user unchecked the box because this particular account has no foreign currency,
                    ' ** the report will end up looking in the wrong tables for the data!
32830               Select Case blnIncludeCurrency
                    Case False
32840                 gblnHasForExThis = False
32850                 blnHasForExThis = False
32860                 gblnSwitchTo = False
32870                 For lngX = 0& To (lngAcctFors - 1&)
32880                   If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
32890                     If arr_varAcctFor(F_ACNT, lngX) > 0 Then
32900                       gblnHasForExThis = True
32910                       blnHasForExThis = True
32920                     End If
32930                     gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
32940                     Exit For
32950                   End If
32960                 Next
32970                 Select Case gblnHasForExThis
                      Case True
32980                   If gblnSwitchTo = True Then
                          ' ** Turn it off since they now do have foreign currencies.
32990                     Set dbs = CurrentDb
33000                     With dbs
                            ' ** Update tblCurrency_Account for curracct_supress = False, by specified [actno].
33010                       Set qdf = .QueryDefs("qryCurrency_17_02")
33020                       With qdf.Parameters
33030                         ![actno] = gstrAccountNo
33040                       End With
33050                       qdf.Execute
33060                       Set qdf = Nothing
33070                       .Close
33080                     End With
33090                     Set dbs = Nothing
33100                   End If
33110                 Case False
33120                   If gblnSwitchTo = True Then
                          ' ** If they've specified to suppress, then turn chkIncludeCurrency off.
33130                     blnIncludeCurrency = False
33140                     .chkIncludeCurrency = False
33150                     .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
33160                   End If
33170                 End Select
33180               Case False
33190                 gblnHasForExThis = False
33200                 blnHasForExThis = False
33210                 gblnSwitchTo = False
33220                 For lngX = 0& To (lngAcctFors - 1&)
33230                   If arr_varAcctFor(F_ACTNO, lngX) = gstrAccountNo Then
33240                     If arr_varAcctFor(F_ACNT, lngX) > 0 Then
33250                       gblnHasForExThis = True
33260                       blnHasForExThis = True
33270                     End If
33280                     gblnSwitchTo = arr_varAcctFor(F_SUPP, lngX)
33290                     Exit For
33300                   End If
33310                 Next
33320                 Select Case gblnHasForExThis
                      Case True
                        ' ** This account does have foreign currencies, and
                        ' ** the user shouldn't have been able to turn it off.
33330                   blnIncludeCurrency = True
33340                   .chkIncludeCurrency = True
33350                   .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
33360                 Case False
33370                   Select Case gblnSwitchTo
                        Case True
33380                     gblnMessage = True
33390                   Case False
33400                     strDocName = "frmStatementParameters_ForEx"
33410                     gblnSetFocus = True
33420                     gblnMessage = True  ' ** False return means cancel.
33430                     gblnSwitchTo = True  ' ** False return means show ForEx, don't supress.
33440                     DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name & "~" & gstrAccountNo
33450                   End Select
33460                   Select Case gblnMessage
                        Case True
33470                     Select Case gblnSwitchTo
                          Case True
                            ' ** Let blnIncludeCurrency remain False.
33480                       .UsePriceHistory = False
33490                     Case False
                            ' ** Turn it back on then.
33500                       blnIncludeCurrency = True
33510                       .chkIncludeCurrency = True
33520                       .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
33530                     End Select
33540                     gblnMessage = False
33550                     gblnSwitchTo = False
33560                   Case False
                          ' ** Cancel this button.
33570                     blnRetVal = False
33580                     DoCmd.Hourglass False
33590                   End Select
33600                 End Select
33610               End Select
33620             End If
33630           End If

33640           If blnRetVal = True Then

                  ' ** Execute the common code.
33650             If BuildAssetListInfo_SP(frm, blnContinue, datAssetListDate, blnPrintAnnualStatement, blnAllStatements, blnNoDataAll, blnRollbackNeeded) = True Then  ' ** Module Function: modStatementParamFuncs1.
33660               If blnContinue = True Then

33670                 If IsNull(.UserReportPath) = True Then
33680                   strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
33690                 Else
33700                   strRptPath = .UserReportPath
33710                 End If

33720                 Select Case .opgAccountNumber
                      Case .opgAccountNumber_optSpecified.OptionValue
33730                   strRptCap = "rptAssetList_" & .cmbAccounts.Column(CBX_A_ACTNO) & "_"
33740                 Case .opgAccountNumber_optAll.OptionValue
33750                   strRptCap = "rptAssetList_All_"
33760                 End Select
33770                 strRptCap = StringReplace(strRptCap, "/", "_")  ' ** Module Function: modStringFuncs.
33780                 If .chkAssetList = True Then
33790                   strRptCap = strRptCap & "_" & Format(.AssetListDate, "yymmdd")
33800                 ElseIf .chkStatements = True Then
33810                   strRptCap = strRptCap & "_" & Format(.DateEnd, "yymmdd")
33820                 End If

33830                 If gblnDev_Debug = True Or (CurrentUser = "Superuser" And .chkAsDev = True) Then  ' ** Internal Access Function: Trust Accountant login.
33840                   strQry1 = "rptAssetList"
33850                   DoCmd.OpenReport strQry1, acViewPreview
33860                 Else

33870                   strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

33880                   If strRptPathFile <> vbNullString Then

                          ' **  All Accounts     : qryStatementParameters_AssetList_06a
                          ' **  One Account      : qryStatementParameters_AssetList_06b
                          ' **  Related Accounts : qryStatementParameters_AssetList_28
                          ' **  Any Rollbacks    : qryStatementParameters_AssetList_15

33890                     strQry1 = vbNullString: strQry2 = vbNullString
33900                     Select Case gblnMessage
                          Case True
                            ' ** Rollbacks were needed.
33910                       If .chkForeignExchange = True And blnIncludeCurrency = True Then
                              ' ** Append qryStatementParameters_AssetList_71_04_02 (xx) to AssetList2; any rollbacks.
33920                         strQry1 = "qryStatementParameters_AssetList_71_04_03"
                              'PRICING HISTORY!
                              'qryStatementParameters_AssetList_71_04_23
                              '![curdat] = datAssetListDate
                              ' ** Append qryStatementParameters_AssetList_72_04_12 (xx) to tmpAssetList4; any rollbacks.
33930                         strQry2 = "qryStatementParameters_AssetList_72_04_13"
                              'PRICING HISTORY!
                              'qryStatementParameters_AssetList_72_04_33
                              'qryStatementParameters_AssetList_72_04_33a
                              'qryStatementParameters_AssetList_72_04_33b
                              '![curdat] = datAssetListDate
33940                       Else
                              ' ** Append qryStatementParameters_AssetList_41d (qryStatementParameters_AssetList_40d
                              ' ** (qryStatementParameters_AssetList_15 (tmpAssetList2, all fields, with rollback), with add'l
                              ' ** fields; rollbacks, any), with add'l fields; all accounts) to AssetList table; rollbacks, any.
33950                         strQry1 = "qryStatementParameters_AssetList_42d"
                              ' ** Append qryStatementParameters_AssetList_56d (...) to tmpAssetList1; any rollbacks.
33960                         strQry2 = "qryStatementParameters_AssetList_57d"
33970                       End If
33980                     Case False
                            ' ** No Rollbacks needed.
33990                       Select Case .chkRelatedAccounts
                            Case True
                              ' ** With Related Accounts.
34000                         If .chkForeignExchange = True And blnIncludeCurrency = True Then
                                ' ** Append qryStatementParameters_AssetList_71_03_02 (xx) to AssetList2 table; related accounts.
34010                           strQry1 = "qryStatementParameters_AssetList_71_03_03"
                                'PRICING HISTORY!
                                'qryStatementParameters_AssetList_71_03_23
                                '![curdat] = datAssetListDate
                                ' ** Append qryStatementParameters_AssetList_72_03_12 (xx) to tmpAssetList4; related accounts.
34020                           strQry2 = "qryStatementParameters_AssetList_72_03_13"
                                'PRICING HISTORY!
                                'qryStatementParameters_AssetList_72_03_33
                                'qryStatementParameters_AssetList_72_03_33a
                                'qryStatementParameters_AssetList_72_03_33b
                                '![curdat] = datAssetListDate
34030                         Else
                                ' ** Append qryStatementParameters_AssetList_41c (qryStatementParameters_AssetList_40c
                                ' ** (qryStatementParameters_AssetList_28 (qryStatementParameters_AssetList_27 (tmpRelatedAccount_02,
                                ' ** with qryStatementParameters_AssetList_26b (qryStatementParameters_AssetList_26a (tmpRelatedAccount_02,
                                ' ** linked to Account, grouped and summed, by accountno), grouped and summed), by specified [ractnos];
                                ' ** Cartesian), from code, with [ractnos] replaced with actual accountno's), with add'l fields;
                                ' ** related accounts), with add'l fields; all accounts) to AssetList table; related accounts.
34040                           strQry1 = "qryStatementParameters_AssetList_42c"
                                ' ** Append qryStatementParameters_AssetList_56c (...) to tmpAssetList1; related accounts.
34050                           strQry2 = "qryStatementParameters_AssetList_57c"
34060                         End If
34070                       Case False
                              ' ** Without Related Accounts.
34080                         Select Case .opgAccountNumber
                              Case .opgAccountNumber_optSpecified.OptionValue
                                ' ** One Account.
34090                           If .chkForeignExchange = True And blnIncludeCurrency = True Then
                                  ' ** Append qryStatementParameters_AssetList_71_02_02 (xx) to AssetList2 table; one account.
34100                             strQry1 = "qryStatementParameters_AssetList_71_02_03"
                                  'PRICING HISTORY!
                                  'qryStatementParameters_AssetList_71_02_23
                                  '![curdat] = datAssetListDate
                                  ' ** Append qryStatementParameters_AssetList_72_02_12 (xx) to tmpAssetList4; one account.
34110                             strQry2 = "qryStatementParameters_AssetList_72_02_13"
                                  'strQry2 = "qryStatementParameters_AssetList_72_02_13a"
                                  'strQry2 = "qryStatementParameters_AssetList_72_02_13b"
                                  'PRICING HISTORY!
                                  ' ** Query too complex, so broken down into pieces.
                                  'Set dbs = CurrentDb
                                  'Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_72_02_25a")
                                  'qdf.Execute
                                  'Set qdf = Nothing
                                  'Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_72_02_26a")
                                  'qdf.Execute
                                  'Set qdf = Nothing
                                  'Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_72_02_30a")
                                  'Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_72_02_30b")
                                  'qdf.Execute
                                  'Set qdf = Nothing
                                  'Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_72_02_31a")
                                  'qdf.Execute
                                  'Set qdf = Nothing
                                  'Set dbs = Nothing
                                  ' 'qryStatementParameters_AssetList_72_02_33
                                  ' '![curdat] = datAssetListDate
34120                           Else
                                  ' ** Append qryStatementParameters_AssetList_41b (qryStatementParameters_AssetList_40b
                                  ' ** (qryStatementParameters_AssetList_06b (Account, linked to ActiveAssets, grouped,
                                  ' ** with add'l fields; specified FormRef('accountno')), with add'l fields; one account),
                                  ' ** with add'l fields; all accounts) to AssetList table; one account.
34130                             strQry1 = "qryStatementParameters_AssetList_42b"
                                  ' ** Append qryStatementParameters_AssetList_56b (...) to tmpAssetList1; one account.
34140                             strQry2 = "qryStatementParameters_AssetList_57b"
34150                           End If
34160                         Case .opgAccountNumber_optAll.OptionValue
                                ' ** All Accounts.
34170                           If .chkForeignExchange = True And blnIncludeCurrency = True Then
                                  ' ** Append qryStatementParameters_AssetList_71_01_02 (xx) to AssetList2 table; all accounts
34180                             strQry1 = "qryStatementParameters_AssetList_71_01_03"
                                  'PRICING HISTORY!
                                  'qryStatementParameters_AssetList_71_01_23
                                  '![curdat] = datAssetListDate
                                  ' ** Append qryStatementParameters_AssetList_72_01_12 (xx) to tmpAssetList4; all accounts.
34190                             strQry2 = "qryStatementParameters_AssetList_72_01_13"
                                  'PRICING HISTORY!
                                  'qryStatementParameters_AssetList_72_01_33
                                  'qryStatementParameters_AssetList_72_01_33a
                                  'qryStatementParameters_AssetList_72_01_33b
                                  '![curdat] = datAssetListDate
34200                           Else
                                  ' ** Append qryStatementParameters_AssetList_41a (qryStatementParameters_AssetList_40a
                                  ' ** (qryStatementParameters_AssetList_06a (Account, linked to ActiveAssets, grouped,
                                  ' ** with add'l fields; all accounts), with add'l fields; all accounts), with add'l
                                  ' ** fields; all accounts) to AssetList table; all accounts.
34210                             strQry1 = "qryStatementParameters_AssetList_42a"
                                  ' ** Append qryStatementParameters_AssetList_56a (...) to tmpAssetList1; all accounts.
34220                             strQry2 = "qryStatementParameters_AssetList_57a"
34230                           End If
34240                         End Select
34250                       End Select
34260                     End Select

34270                     Set dbs = CurrentDb
34280                     With dbs
                            ' ** Empty AssetList.
34290                       Set qdf = .QueryDefs("qryStatementParameters_AssetList_09a")
34300                       qdf.Execute
34310                       Set qdf = Nothing
                            ' ** Empty AssetList2.
34320                       Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_50")
34330                       qdf.Execute
34340                       Set qdf = Nothing
                            ' ** Empty tmpAssetList1.
34350                       Set qdf = .QueryDefs("qryStatementParameters_AssetList_09b")
34360                       qdf.Execute
34370                       Set qdf = Nothing
                            ' ** Empty tmpAssetList4.
34380                       Set qdf = .QueryDefs("qryStatementParameters_AssetList_70_51")
34390                       qdf.Execute
34400                       Set qdf = Nothing
                            ' ** Append qryStatementParameters_AssetList_nn to AssetList/AssetList2.
34410                       Set qdf = .QueryDefs(strQry1)
34420                       qdf.Execute
34430                       Set qdf = Nothing
                            ' ** Append qryStatementParameters_AssetList_nn to tmpAssetList1/tmpAssetList4.
34440                       Set qdf = .QueryDefs(strQry2)
34450                       qdf.Execute
34460                       Set qdf = Nothing
34470                       .Close
34480                     End With

34490                     Select Case .chkStatements
                          Case True
34500                       If .chkForeignExchange = True And blnIncludeCurrency = True Then
                              ' ** Asset List: qryStatementParameters_AssetList_78_02_04 (Union of
                              ' ** qryStatementParameters_AssetList_78_02_01 (qryStatementParameters_AssetList_78_01_03
                              ' ** (qryStatementParameters_AssetList_78_01_02 (tmpAssetList4, with
                              ' ** qryStatementParameters_AssetList_78_01_01 (tmpAssetList4, grouped by accountno, with
                              ' ** Min([Sortx]) as Sortz), with Sortz), with field name changes), linked to
                              ' ** qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a (Account,
                              ' ** with MonthNum = FormRef('MonthNum')), just those matching MonthNum), just scheduled
                              ' ** accounts), qryStatementParameters_AssetList_78_02_02 (Report title),
                              ' ** qryStatementParameters_AssetList_78_02_03 (Report period)), For Export; From 64a.
34510                         strQry1 = "qryStatementParameters_AssetList_78_02_05"
34520                       Else
                              ' ** Asset List: qryStatementParameters_AssetList_63a (Union of
                              ' ** qryStatementParameters_AssetList_60a (qryStatementParameters_AssetList_60
                              ' ** (qryStatementParameters_AssetList_59 (tmpAssetList1, with
                              ' ** qryStatementParameters_AssetList_58 (tmpAssetList1, grouped by accountno,
                              ' ** with Min([Sortx]) as Sortz), with Sortz), with field name changes), linked
                              ' ** to qryStatementParameters_AssetList_19b (qryStatementParameters_AssetList_19a
                              ' ** (Account, with MonthNum = FormRef('MonthNum')), just those matching MonthNum),
                              ' ** just scheduled accounts), qryStatementParameters_AssetList_61 (Report title),
                              ' ** qryStatementParameters_AssetList_62 (Report period)), For Export.
34530                         strQry1 = "qryStatementParameters_AssetList_64a"
34540                       End If
34550                     Case False
34560                       If .chkForeignExchange = True And blnIncludeCurrency = True Then
                              ' ** Asset List: qryStatementParameters_AssetList_78_01_06 (Union of
                              ' ** qryStatementParameters_AssetList_78_01_03 (qryStatementParameters_AssetList_78_01_02
                              ' ** (tmpAssetList4, with qryStatementParameters_AssetList_78_01_01 (tmpAssetList4,
                              ' ** grouped by accountno, with Min([Sortx]) as Sortz), with Sortz), with field name
                              ' ** changes), qryStatementParameters_AssetList_78_01_04 (Report title),
                              ' ** qryStatementParameters_AssetList_78_01_05 (Report period)), For Export; From .._64.
34570                         strQry1 = "qryStatementParameters_AssetList_78_01_07"
34580                       Else
                              ' ** Asset List: qryStatementParameters_AssetList_63 (Union of
                              ' ** qryStatementParameters_AssetList_60 (qryStatementParameters_AssetList_59
                              ' ** (tmpAssetList1, with qryStatementParameters_AssetList_58 (tmpAssetList1,
                              ' ** grouped by accountno, with Min([Sortx]) as Sortz), with Sortz), with field
                              ' ** name changes), qryStatementParameters_AssetList_61 (Report title),
                              ' ** qryStatementParameters_AssetList_62 (Report period)), For Export.
34590                         strQry1 = "qryStatementParameters_AssetList_64"
34600                       End If
34610                     End Select

34620                     DoCmd.OutputTo acOutputQuery, strQry1, acFormatXLS, strRptPathFile, True
34630                     .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.

34640                   End If

34650                   If IsLoaded("rptAssetList", acReport) = True Then  ' ** Module Function: modFileUtilities.
34660                     DoCmd.Close acReport, "rptAssetList"
34670                   End If

34680                 End If  ' ** gblnDev_Debug.

34690                 If intOptGrpAcctNum <> .opgAccountNumber Then
34700                   .opgAccountNumber = intOptGrpAcctNum  '#Covered.
34710                   .opgAccountNumber_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
34720                 End If

34730               End If  ' ** blnContinue.
34740             End If  ' ** BuildAssetListInfo_SP().

34750           End If  ' ** blnRetVal.

34760         End If
34770       End If  ' ** cmbAccounts.

34780     Else
34790       blnHasForExClick = False
34800       DoCmd.Hourglass False
34810     End If
34820   End With

EXITP:
34830   Set qdf = Nothing
34840   Set dbs = Nothing
34850   DoCmd.SetWarnings True
34860   Exit Sub

ERRH:
4200    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** User Canceled.
4210      blnContinue = False
4220    Case Else
4230      Select Case ERR.Number
          Case Else
4240        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4250      End Select
4260    End Select
4270    Resume EXITP

End Sub

Public Function SubsequentTransactionCheck(datDate As Date) As Boolean
' ** Check to see if there are transactions after the date selected.
' ** True  = No transactions after, or user clicked to proceed anyway.
' ** False = Otherwise.

34900 On Error GoTo ERRH

        Const THIS_PROC As String = "SubsequentTransactionCheck"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim intCount As Integer
        Dim blnRetVal As Boolean

34910   intCount = 0
34920   blnRetVal = True    ' ** Unless proven otherwise.

34930   Set dbs = CurrentDb
34940   With dbs
          ' ** Ledger, linked to qryStatementParameters_20 (Account, now with
          ' ** DateClosed = Null, by specified FormRef('MonthNum')), by specified [datend].
34950     Set qdf = .QueryDefs("qryStatementParameters_12")
34960     With qdf.Parameters
34970       ![datEnd] = datDate
34980     End With
34990     Set rst = qdf.OpenRecordset
35000     With rst
35010       .MoveFirst
35020       intCount = ![TranCount]
35030       .Close
35040     End With
35050     .Close
35060   End With

35070   If intCount > 0 Then
35080     If MsgBox("There are transactions posted after " & CStr(datDate) & "." & vbCrLf & vbCrLf & _
              "Do you want to continue anyway?", vbQuestion + vbYesNo + vbDefaultButton2, _
              (Left(("Confirm After-Statement Posting" & Space(55)), 55) & "X01")) = vbNo Then
35090       blnRetVal = False
35100     End If
35110   End If

EXITP:
35120   Set rst = Nothing
35130   Set qdf = Nothing
35140   Set dbs = Nothing
35150   SubsequentTransactionCheck = blnRetVal
35160   Exit Function

ERRH:
4200    Select Case ERR.Number
        Case Else
4210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4220    End Select
4230    Resume EXITP

End Function

Public Function AcctSched_Load() As Variant
' ** Called by:
' **   modStatementParamFuncs1:
' **     ForEx_ChkScheduled()
' **  modStatementParamFuncs2:
' **     CmbAccts_After_SP()
' **     Btn_Enable_SP()

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctSched_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngLastMonthID As Long
        Dim lngX As Long, lngE As Long, lngF As Long
        Dim arr_varRetVal As Variant

35210   lngStmts = 12&
35220   ReDim arr_varStmt(S_ELEMS1, S_ELEMS2, 0)

35230   Set dbs = CurrentDb
35240   With dbs
          ' ** qryStatementParameters_34_03 (xx), linked to qryStatementParameters_34_04 (xx), with cnt_smt.
35250     Set qdf = .QueryDefs("qryStatementParameters_34_05")
35260     Set rst = qdf.OpenRecordset
35270     With rst
35280       If .BOF = True And .EOF = True Then
35290         For lngX = 1& To 12&
35300           arr_varStmt(lngX, 0, 0) = lngX
35310           arr_varStmt(lngX, 1, 0) = Choose(lngX, "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
35320           arr_varStmt(lngX, 2, 0) = 0&
35330           arr_varStmt(lngX, 3, 0) = Null
35340           arr_varStmt(lngX, 4, 0) = Null
35350         Next  ' ** lngX.
35360       Else
35370         .MoveLast
35380         lngRecs = .RecordCount
35390         .MoveFirst
35400         lngLastMonthID = 0&
35410         lngF = 0&  ' ** Max accounts.
35420         For lngX = 1& To lngRecs
35430           If ![month_id] <> lngLastMonthID Then
35440             lngLastMonthID = ![month_id]
35450             lngE = 0&
35460           Else
35470             lngE = lngE + 1&
35480           End If
                ' ** This isn't quite the way I wanted, but I'm not sure how to do it that way.
                ' ** The idea was that each month should have only as many account elements as the number of accounts,
                ' ** and no more, with the 0 element being empty for months with no accounts.
                ' ** This gives every month the number of account elements as the greatest number for any month.
                'WHEN WE START A NEW MONTH, THE REDIM WAS WIPING OUT ALL THE REST!!!!!!!!!!!!!!
35490           If lngE > lngF Then
35500             ReDim Preserve arr_varStmt(S_ELEMS1, S_ELEMS2, lngE)
35510             lngF = lngE
35520           End If
                ' ************************************************
                ' ** Array: arr_varStmt()
                ' **
                ' **   Field  Element  Name           Constant
                ' **   =====  =======  =============  ==========
                ' **     1       0     month_id       S_MID
                ' **     2       1     month_short    S_MSHT
                ' **     3       2     cnt_smt        S_CNT
                ' **     4       3     accountno      S_ACTNO
                ' **     5       4     shortname      S_SNAM
                ' **
                ' ************************************************
35530           arr_varStmt(lngLastMonthID, S_MID, lngE) = ![month_id]
35540           arr_varStmt(lngLastMonthID, S_MSHT, lngE) = ![month_short]
35550           arr_varStmt(lngLastMonthID, S_CNT, lngE) = ![cnt_smt]
35560           arr_varStmt(lngLastMonthID, S_ACTNO, lngE) = ![accountno]
35570           arr_varStmt(lngLastMonthID, S_SNAM, lngE) = ![shortname]
35580           If lngX < lngRecs Then .MoveNext
35590         Next
35600       End If
35610       .Close
35620     End With
35630     .Close
35640   End With
35650   arr_varRetVal = arr_varStmt

EXITP:
35660   Set rst = Nothing
35670   Set qdf = Nothing
35680   Set dbs = Nothing
35690   AcctSched_Load = arr_varRetVal
35700   Exit Function

ERRH:
4200    arr_varRetVal(0, 0) = RET_ERR
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Function

Public Sub ForEx_ChkScheduled(frm As Access.Form)
' ** Called by:
' **   frmStatementParameters:
' **     cmbMonth_AfterUpdate()

35800 On Error GoTo ERRH

        Const THIS_PROC As String = "ForEx_ChkScheduled"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngMonthID As Long, lngHits As Long
        Dim strAccountNo As String
        Dim blnFound1 As Boolean, blnFound2 As Boolean, blnTryArchive As Boolean
        Dim lngX As Long, lngY As Long

35810   With frm

35820     .ForEx_ChkScheduled_lbl.Visible = False

35830     If lngAcctFors = 0& Or IsEmpty(arr_varAcctFor) = True Then
35840       ForExArr_Load  ' ** Function: Above.
            'Runs Qyrs, loads arrays.
35850     End If

35860     If lngStmts = 0& Or IsEmpty(arr_varStmt) = True Then
35870       AcctSched_Load  ' ** Function: Above.
            'Runs Qyrs, loads arrays.
35880     End If

35890     lngMonthID = .cmbMonth.Column(CBX_MON_ID)
35900     lngStmtCnt = arr_varStmt(lngMonthID, S_CNT, 0)
35910     lngHits = 0&
35920     If lngStmtCnt = 0& Then
            ' ** If no accounts are scheduled for this month,
            ' ** I think it's handled elsewhere.
            ' ** Anyway, turn off currency.
35930       .chkIncludeCurrency = False
35940       .chkIncludeCurrency.Locked = False
35950       .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.

35960     Else
            ' ** See if any of these have foreign currency.
35970       For lngX = 0& To (lngStmtCnt - 1&)
35980         strAccountNo = arr_varStmt(lngMonthID, S_ACTNO, lngX)
35990         blnFound1 = False: blnFound2 = False: blnTryArchive = False
36000         For lngY = 0& To (lngAcctFors - 1&)
36010           If arr_varAcctFor(F_ACTNO, lngY) = strAccountNo Then
36020             If arr_varAcctFor(F_JCNT, lngY) > 0 Then
36030               blnFound1 = True
36040             End If
36050             If arr_varAcctFor(F_ACNT, lngY) > 0 Then
36060               blnFound2 = True
36070             End If
36080             Exit For
36090           End If
36100         Next  ' ** lngY.
36110         If blnFound2 = True Then
                ' ** A scheduled account currently holds at least one foreign currency asset.
36120           .chkIncludeCurrency = True
36130           .chkIncludeCurrency.Locked = True
36140           .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
36150           .ForEx_ChkScheduled_lbl.Caption = "At least one scheduled account currently holds a foreign currency asset."
36160           .ForEx_ChkScheduled_lbl.Visible = True
36170           lngHits = lngHits + 1&
36180           Exit For
36190         ElseIf blnFound1 = True Then
                ' ** A scheduled account does not currently hold a foregin currency asset,
                ' ** but does have previous foreign currency transactions.
                ' ** See if any are in the scheduled period.
36200           gdatStartDate = .DateStart
36210           gdatEndDate = .DateEnd
36220           Set dbs = CurrentDb
                ' ** qryStatementParameters_37_01 (Ledger, just curr_id <> 150, by
                ' ** specified [actno], [datstart], [datend]), grouped by src, with cnt_jno.
36230           Set qdf = dbs.QueryDefs("qryStatementParameters_37_02")
36240           With qdf.Parameters
36250             ![actno] = strAccountNo
36260             ![datStart] = gdatStartDate
36270             ![datEnd] = gdatEndDate
36280           End With
36290           Set rst = qdf.OpenRecordset
36300           If rst.BOF = True And rst.EOF = True Then
                  ' ** None here.
36310             If lngAcctArchs = 0& Then
                    ' ** We're done here.
36320             Else
36330               blnTryArchive = True
36340             End If
36350           Else
36360             rst.MoveFirst
36370             If IsNull(rst![cnt_jno]) = True Then
                    ' ** None here.
36380               blnTryArchive = True
36390             Else
36400               If rst![cnt_jno] = 0 Then
                      ' ** None here.
36410                 blnTryArchive = True
36420               Else
                      ' ** This account has scheduled-period foreign currency transactions.
36430                 .chkIncludeCurrency = True
36440                 .chkIncludeCurrency.Locked = True
36450                 .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
36460                 .ForEx_ChkScheduled_lbl.Caption = "At least one scheduled account has schedule-period foreign currency transactions."
36470                 .ForEx_ChkScheduled_lbl.Visible = True
36480                 lngHits = lngHits + 1&
36490                 Exit For
36500               End If
36510             End If
36520           End If
36530           rst.Close
36540           Set rst = Nothing
36550           Set qdf = Nothing
36560         Else
                ' ** This account doesn't hold a foreign currency asset,
                ' ** nor does it have any foreign currency transactions.
36570         End If
36580         If blnTryArchive = True Then
36590           For lngY = 0& To (lngAcctArchs - 1&)
36600             If arr_varAcctArch(AR_ACTNO, lngY) = strAccountNo Then
                    ' ** qryStatementParameters_37_03 (LedgerArchive, just curr_id <> 150, by
                    ' ** specified [actno], [datstart], [datend]), grouped by src, with cnt_jno.
36610               Set qdf = dbs.QueryDefs("qryStatementParameters_37_04")
36620               With qdf.Parameters
36630                 ![actno] = strAccountNo
36640                 ![datStart] = gdatStartDate
36650                 ![datEnd] = gdatEndDate
36660               End With
36670               Set rst = qdf.OpenRecordset
36680               If rst.BOF = True And rst.EOF = True Then
                      ' ** None here either.
36690               Else
36700                 rst.MoveFirst
36710                 If IsNull(![cnt_jno]) = True Then
                        ' ** None here either.
36720                 Else
36730                   If ![cnt_jno] = 0 Then
                          ' ** None here either.
36740                   Else
                          ' ** This account has scheduled-period foreign currency transactions.
36750                     .chkIncludeCurrency = True
36760                     .chkIncludeCurrency.Locked = True
36770                     .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
36780                     .ForEx_ChkScheduled_lbl.Caption = "At least one scheduled account has schedule-period foreign currency transactions."
36790                     .ForEx_ChkScheduled_lbl.Visible = True
36800                     lngHits = lngHits + 1&
36810                   End If
36820                 End If
36830               End If
36840               rst.Close
36850               Set rst = Nothing
36860               Set qdf = Nothing
36870               Exit For
36880             End If
36890           Next  ' ** lngY.
36900         End If
36910       Next  ' ** lngX.
36920     End If  ' ** lngStmtCnt.
36930     If lngHits = 0& Then
            ' ** No scheduled account holds a foreign currency asset,
            ' ** nor any scheduled-period foreign currency transactions.
36940       .chkIncludeCurrency = False
36950       .chkIncludeCurrency.Locked = False
36960       .chkIncludeCurrency_AfterUpdate  ' ** Form Procedure: frmStatementParameters.
36970       .ForEx_ChkScheduled_lbl.Caption = "No scheduled account holds a foreign currency asset, " & _
              "nor any schedule-period foreign currency transactions."
36980       .ForEx_ChkScheduled_lbl.Visible = True
36990     End If

37000   End With

EXITP:
37010   Set rst = Nothing
37020   Set qdf = Nothing
37030   Set dbs = Nothing
37040   Exit Sub

ERRH:
4200    Select Case ERR.Number
        Case Else
4210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4220    End Select
4230    Resume EXITP

End Sub

Public Function ForExArr_Get() As Variant

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "ForExArr_Get"

        Dim arr_varRetVal As Variant, blnRetVal As Boolean

37110   blnRetVal = True
37120   If lngAcctFors = 0& Or IsEmpty(arr_varAcctFor) = True Then
37130     blnRetVal = ForExArr_Load  ' ** Function: Above.
37140     DoEvents
37150   End If
37160   arr_varRetVal = arr_varAcctFor

EXITP:
37170   ForExArr_Get = arr_varRetVal
37180   Exit Function

ERRH:
4200    arr_varRetVal(0, 0) = RET_ERR
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Function

Public Function ForExArr_Load() As Boolean

37200 On Error GoTo ERRH

        Const THIS_PROC As String = "ForExArr_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

37210   blnRetVal = False

37220   If lngAcctFors = 0& Or IsEmpty(arr_varAcctFor) = True Then
37230     Set dbs = CurrentDb
37240     With dbs
            ' ** tblCurrency_Account, all records.
37250       Set qdf = .QueryDefs("qryStatementParameters_36")
37260       Set rst = qdf.OpenRecordset
37270       With rst
37280         .MoveLast
37290         lngAcctFors = .RecordCount
37300         .MoveFirst
37310         arr_varAcctFor = .GetRows(lngAcctFors)
              ' ******************************************************
              ' ** Array: arr_varAcctFor()
              ' **
              ' **   Field  Element  Name                 Constant
              ' **   =====  =======  ===================  ==========
              ' **     1       0     accountno            F_ACTNO
              ' **     2       1     curracct_jno         F_JCNT
              ' **     3       2     curracct_aa          F_ACNT
              ' **     4       3     curracct_suppress    F_SUPP
              ' **
              ' ******************************************************
37320         blnRetVal = True
37330         .Close
37340       End With
37350       .Close
37360     End With
37370   End If

EXITP:
37380   Set rst = Nothing
37390   Set qdf = Nothing
37400   Set dbs = Nothing
37410   ForExArr_Load = blnRetVal
37420   Exit Function

ERRH:
4200    blnRetVal = False
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Function

Public Sub ForExRptSub_Load(frm As Access.Form, blnRollbackNeeded As Boolean)

37500 On Error GoTo ERRH

        Const THIS_PROC As String = "ForExRptSub_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQryName As String, strSQL As String
        Dim blnPriceHistory As Boolean, blnNoAssets As Boolean

37510   With frm

37520     blnPriceHistory = .UsePriceHistory
37530     blnNoAssets = frm.chkNoAssets

37540     Select Case blnRollbackNeeded
          Case True
            ' ** Rollbacks were needed.
37550       Select Case .chkStatements
            Case True
37560         Select Case blnNoAssets
              Case True
                'SINGLES ONLY!
                ' ** qryStatementParameters_AssetList_81_36_01 (xx), rounded, with TotalCost_str, TotalMarket_str.
37570           strQryName = "qryStatementParameters_AssetList_81_36_02"
37580         Case False
37590           Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' ** qryStatementParameters_AssetList_81_29_01 (xx), rounded, with TotalCost_str, TotalMarket_str.
37600             strQryName = "qryStatementParameters_AssetList_81_29_02"
37610           Case False
                  ' ** qryStatementParameters_AssetList_81_09_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37620             strQryName = "qryStatementParameters_AssetList_81_09_02"
37630           End Select  ' ** blnPriceHistory.
37640         End Select

37650       Case False
37660         Select Case blnNoAssets
              Case True
                'SINGLES ONLY!
                ' ** qryStatementParameters_AssetList_81_36_01 (xx), rounded, with TotalCost_str, TotalMarket_str.
37670           strQryName = "qryStatementParameters_AssetList_81_36_02"
37680         Case False
37690           Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' ** qryStatementParameters_AssetList_81_28_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37700             strQryName = "qryStatementParameters_AssetList_81_28_02"
37710           Case False
                  ' ** qryStatementParameters_AssetList_81_08_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37720             strQryName = "qryStatementParameters_AssetList_81_08_02"
37730           End Select  ' ** blnPriceHistory.
37740         End Select

37750       End Select  ' ** chkStatements.
37760     Case False
            ' ** No Rollbacks needed.
37770       Select Case frm.chkRelatedAccounts
            Case True
              'NO ASSETS!
              ' ** With Related Accounts.
37780         Select Case .chkStatements
              Case True
37790           Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' ** qryStatementParameters_AssetList_81_34_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37800             strQryName = "qryStatementParameters_AssetList_81_34_02"
37810           Case False
                  ' ** qryStatementParameters_AssetList_81_14_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37820             strQryName = "qryStatementParameters_AssetList_81_14_02"
37830           End Select  ' ** blnPriceHistory.
37840         Case False
37850           Select Case blnPriceHistory
                Case True
                  ' ** PRICING HISTORY!
                  ' ** qryStatementParameters_AssetList_81_33_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37860             strQryName = "qryStatementParameters_AssetList_81_33_02"
37870           Case False
                  ' ** qryStatementParameters_AssetList_81_13_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37880             strQryName = "qryStatementParameters_AssetList_81_13_02"
37890           End Select  ' ** blnPriceHistory.
37900         End Select  ' ** chkStatements.
37910       Case False
              ' ** Without Related Accounts.
37920         Select Case frm.opgAccountNumber
              Case frm.opgAccountNumber_optSpecified.OptionValue
                ' ** One Account.
                'NO ASSETS!
37930           Select Case blnNoAssets
                Case True
37940             Select Case .chkStatements
                  Case True
37950               Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** qryStatementParameters_AssetList_81_26_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37960                 strQryName = "qryStatementParameters_AssetList_81_26_02"
37970               Case False
                      ' ** qryStatementParameters_AssetList_81_06_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
37980                 strQryName = "qryStatementParameters_AssetList_81_06_02"
37990               End Select  ' ** blnPriceHistory.
38000             Case False
38010               Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** qryStatementParameters_AssetList_81_23_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38020                 strQryName = "qryStatementParameters_AssetList_81_23_02"
38030               Case False
                      ' ** qryStatementParameters_AssetList_81_03_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38040                 strQryName = "qryStatementParameters_AssetList_81_03_02"
38050               End Select  ' ** blnPriceHistory.
38060             End Select  ' ** chkStatements.
38070           Case False
38080             Select Case .chkStatements
                  Case True
38090               Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** qryStatementParameters_AssetList_81_25_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38100                 strQryName = "qryStatementParameters_AssetList_81_25_02"
38110               Case False
                      ' ** qryStatementParameters_AssetList_81_05_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38120                 strQryName = "qryStatementParameters_AssetList_81_05_02"
38130               End Select  ' ** blnPriceHistory.
38140             Case False
38150               Select Case blnPriceHistory
                    Case True
                      ' ** PRICING HISTORY!
                      ' ** qryStatementParameters_AssetList_81_22_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38160                 strQryName = "qryStatementParameters_AssetList_81_22_02"
38170               Case False
                      ' ** qryStatementParameters_AssetList_81_02_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38180                 strQryName = "qryStatementParameters_AssetList_81_02_02"
38190               End Select  ' ** blnPriceHistory.
38200             End Select  ' ** chkStatements.
38210           End Select  ' ** chkNoAssets.
38220         Case frm.opgAccountNumber_optAll.OptionValue
                ' ** All Accounts.
                'NO ASSETS!
38230           Select Case .chkStatements
                Case True
38240             Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' ** qryStatementParameters_AssetList_81_24_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38250               strQryName = "qryStatementParameters_AssetList_81_24_02"
38260             Case False
                    ' ** qryStatementParameters_AssetList_81_04_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38270               strQryName = "qryStatementParameters_AssetList_81_04_02"
38280             End Select  ' ** blnPriceHistory.
38290           Case False
38300             Select Case blnPriceHistory
                  Case True
                    ' ** PRICING HISTORY!
                    ' ** qryStatementParameters_AssetList_81_21_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38310               strQryName = "qryStatementParameters_AssetList_81_21_02"
38320             Case False
                    ' ** qryStatementParameters_AssetList_81_01_01 (xx). rounded, with TotalCost_str, TotalMarket_str.
38330               strQryName = "qryStatementParameters_AssetList_81_01_02"
38340             End Select  ' ** blnPriceHistory.
38350           End Select  ' ** chkStatements.
38360         End Select  ' ** opgAccountNumber.
38370       End Select  ' ** chkRelatedAccounts.
38380     End Select  ' ** gblnMessage.

38390     Set dbs = CurrentDb
38400     With dbs
38410       Set qdf = .QueryDefs(strQryName)
38420       strSQL = qdf.SQL
38430       Set qdf = Nothing
38440       Set qdf = .QueryDefs("qryStatementParameters_AssetList_82")
38450       qdf.SQL = strSQL
38460       Set qdf = Nothing
38470       .Close
38480     End With

38490   End With

EXITP:
38500   Set qdf = Nothing
38510   Set dbs = Nothing
38520   Exit Sub

ERRH:
4200    Select Case ERR.Number
        Case Else
4210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4220    End Select
4230    Resume EXITP

End Sub

Public Function HasForEx_SP(strAccountNo As String) As Boolean

38600 On Error GoTo ERRH

        Const THIS_PROC As String = "HasForEx_SP"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

38610   blnRetVal = False

38620   If Trim(strAccountNo) <> vbNullString Then
38630     If strAccountNo = "All" Then
38640       Set dbs = CurrentDb
38650       With dbs
              ' ** ActiveAssets, grouped by curr_id, just curr_id <> 150, with cnt.
38660         Set qdf = .QueryDefs("qryStatementParameters_AssetList_80_01")
38670         Set rst = qdf.OpenRecordset
38680         With rst
38690           If .BOF = True And .EOF = True Then
                  ' ** No foreign currencies whatsoever.
38700           Else
38710             .MoveFirst
38720             blnRetVal = True
38730           End If
38740           .Close
38750         End With
38760         Set rst = Nothing
38770         Set qdf = Nothing
38780         .Close
38790       End With
38800       Set dbs = Nothing
38810     Else
            ' ** ActiveAssets, grouped by curr_id, just curr_id <> 150, with cnt, by specified [actno].
38820       Set dbs = CurrentDb
38830       With dbs
38840         Set qdf = .QueryDefs("qryStatementParameters_AssetList_80_02")
38850         With qdf.Parameters
38860           ![actno] = strAccountNo
38870         End With
38880         Set rst = qdf.OpenRecordset
38890         With rst
38900           If .BOF = True And .EOF = True Then
                  ' ** This account has no foreign currencies.
38910           Else
38920             .MoveFirst
38930             blnRetVal = True
38940           End If
38950           .Close
38960         End With
38970         Set rst = Nothing
38980         Set qdf = Nothing
38990         .Close
39000       End With
39010       Set dbs = Nothing
39020     End If
39030   End If

EXITP:
39040   Set rst = Nothing
39050   Set qdf = Nothing
39060   Set dbs = Nothing
39070   HasForEx_SP = blnRetVal
39080   Exit Function

ERRH:
4200    DoCmd.Hourglass False
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Function

Public Sub ErrSave(dblErrNum As Double, strErrDesc As String, strFunctionName As String, lngLineNum As Long)
' ** Report ALL errors on form.

39100 On Error GoTo ERRH

        Const THIS_PROC As String = "ErrSave"

        Dim dbs As DAO.Database, rst As DAO.Recordset

39110   Set dbs = CurrentDb
39120   Set rst = dbs.OpenRecordset("tblErrorLog", dbOpenDynaset, dbConsistent)

39130   zErrorWriteRecord dblErrNum, strErrDesc, THIS_NAME, strFunctionName, lngLineNum, rst  ' ** Module Function: modErrorHandler.

39140   rst.Close
39150   dbs.Close

EXITP:
39160   Set rst = Nothing
39170   Set dbs = Nothing
39180   Exit Sub

ERRH:
4200    Beep
4210    MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
          "Module: " & THIS_NAME & vbCrLf & "Sub/Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & CStr(Erl), _
          vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
4220    Resume EXITP

End Sub
