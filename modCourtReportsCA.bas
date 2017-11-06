Attribute VB_Name = "modCourtReportsCA"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCourtReportsCA"

'VGC 09/13/2017: CHANGES!

' ** cmbAccounts combo box constants:
Private Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
'Private Const CBX_A_DESC   As Integer = 1  ' ** Desc
'Private Const CBX_A_PREDAT As Integer = 2  ' ** predate
Private Const CBX_A_SHORT  As Integer = 3  ' ** shortname
'Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname  {Constant is used in reports!}
'Private Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
'Private Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
'Private Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
'Private Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

' ** Array: arr_varCARpt().
Private lngCARpts As Long, arr_varCARpt As Variant
'Private Const CR_ID     As Integer = 0
Private Const CR_NUM    As Integer = 1
Private Const CR_CAT    As Integer = 2
Private Const CR_CON    As Integer = 3
Private Const CR_DIV    As Integer = 4
Private Const CR_DIVTXT As Integer = 5
Private Const CR_DIVTTL As Integer = 6
Private Const CR_GRP    As Integer = 7
Private Const CR_GRPTXT As Integer = 8
Private Const CR_DATE   As Integer = 9
Private Const CR_SCHED  As Integer = 10

' ** Array: arr_varFile().
Private lngFiles As Long, arr_varFile As Variant
Private Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const F_RNAM As Integer = 0
Private Const F_FILE As Integer = 1
Private Const F_PATH As Integer = 2

' ** Array: arr_varCap().
Private lngCaps As Long, arr_varCap As Variant
'Private Const C_RID  As Integer = 0
Private Const C_RNAM As Integer = 1
'Private Const C_CAP  As Integer = 2
Private Const C_CAPN As Integer = 3

' ** CourtReport family constants.
Public CRPT_DIV_CHARGES As Integer
Public CRPT_DIV_CREDITS As Integer
Public CRPT_DIV_ADDL   As Integer
Public CRPT_ON_HAND_BEGL As Integer
Public CRPT_CASH_BEG As Integer
Public CRPT_NON_CASH_BEG As Integer
Public CRPT_ON_HAND_BEG As Integer
Public CRPT_ADDL_PROP As Integer
Public CRPT_RECEIPTS As Integer
Public CRPT_GAINS As Integer
Public CRPT_OTH_CHG As Integer
Public CRPT_NET_INCOME As Integer
Public CRPT_DISBURSEMENTS As Integer
Public CRPT_LOSSES As Integer
Public CRPT_NET_LOSS As Integer
Public CRPT_DISTRIBUTIONS As Integer
Public CRPT_OTH_CRED As Integer
Public CRPT_ON_HAND_ENDL As Integer
Public CRPT_CASH_END As Integer
Public CRPT_NON_CASH_END As Integer
Public CRPT_ON_HAND_END As Integer
Public CRPT_INVEST_INFO As Integer
Public CRPT_CHANGES As Integer

Private blnExcel As Boolean, blnAllCancel As Boolean
Private strCaseNum As String, strThisProc As String
' **

Public Function CANum(strConst As String) As Integer

100   On Error GoTo ERRH

        Const THIS_PROC As String = "CANum"

        Dim lngX As Long
        Dim intRetVal As Integer

110     For lngX = 0& To (lngCARpts - 1&)
120       If arr_varCARpt(CR_CON, lngX) = strConst Then
130         intRetVal = arr_varCARpt(CR_NUM, lngX)
140         Exit For
150       End If
160     Next

EXITP:
170     CANum = intRetVal
180     Exit Function

ERRH:
190     intRetVal = 0
200     Select Case ERR.Number
        Case Else
210       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
220     End Select
230     Resume EXITP

End Function

Public Function CACourtReportLoad() As Boolean
'modCourtReportsCA.CABuildCourtReportData()
'frmRpt_CourtReports_CA.PreviewOrPrint()
'frmRpt_CourtReports_CA.SendToFile()

300   On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportLoad"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngX As Long
        Dim blnRetVal As Boolean

310     blnRetVal = True

320     Set dbs = CurrentDb
330     Set qdf = dbs.QueryDefs("qryCourtReport_CA_20")
340     Set rst = qdf.OpenRecordset
350     With rst
360       .MoveLast
370       lngCARpts = .RecordCount
380       .MoveFirst
390       arr_varCARpt = .GetRows(lngCARpts)
          ' *******************************************************
          ' ** Array: arr_varCARpt()
          ' **
          ' **   Field  Element  Name                 Constant
          ' **   =====  =======  ===================  ===========
          ' **     1       0     cr_id                CR_ID
          ' **     2       1     cr_number            CR_NUM
          ' **     3       2     cr_category          CR_CAT
          ' **     4       3     cr_constant          CR_CON
          ' **     5       4     cr_division          CR_DIV
          ' **     6       5     cr_division_text     CR_DIVTXT
          ' **     7       6     cr_division_title    CR_DIVTTL
          ' **     8       7     cr_group             CR_GRP
          ' **     9       8     cr_group_text        CR_GRPTXT
          ' **    10       9     cr_date_source       CR_DATE
          ' **    11      10     cr_schedule          CR_SCHED
          ' **    12      11     cr_user
          ' **    13      12     cr_datecreated
          ' **    14      13     cr_datemodified
          ' **
          ' *******************************************************
400       .Close
410     End With
420     dbs.Close

430     For lngX = 0& To (lngCARpts - 1&)
440       If IsEmpty(arr_varCARpt(CR_GRPTXT, lngX)) = True Then
450         arr_varCARpt(CR_GRPTXT, lngX) = vbNullString
460       ElseIf IsNull(arr_varCARpt(CR_GRPTXT, lngX)) = True Then
470         arr_varCARpt(CR_GRPTXT, lngX) = vbNullString
480       End If
490     Next

500     For lngX = 0& To (lngCARpts - 1&)

510       If arr_varCARpt(CR_DIVTTL, lngX) = "CHARGES" And CRPT_DIV_CHARGES = 0 Then
520         CRPT_DIV_CHARGES = arr_varCARpt(CR_DIV, lngX)  ' ** 20
530       ElseIf arr_varCARpt(CR_DIVTTL, lngX) = "CREDITS" And CRPT_DIV_CREDITS = 0 Then
540         CRPT_DIV_CREDITS = arr_varCARpt(CR_DIV, lngX)  ' ** 40
550       ElseIf arr_varCARpt(CR_DIVTTL, lngX) = "ADDITIONAL INFORMATION" And CRPT_DIV_ADDL = 0 Then
560         CRPT_DIV_ADDL = arr_varCARpt(CR_DIV, lngX)     ' ** 60
570       End If

580       If arr_varCARpt(CR_CON, lngX) = "CRPT_ON_HAND_BEGL" Then       '      2
590         CRPT_ON_HAND_BEGL = arr_varCARpt(CR_NUM, lngX)
600       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_Cash_BEG" Then       '      3
610         CRPT_CASH_BEG = arr_varCARpt(CR_NUM, lngX)
620       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_NON_Cash_BEG" Then   '      4
630         CRPT_NON_CASH_BEG = arr_varCARpt(CR_NUM, lngX)
640       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_ON_HAND_BEG" Then    '5     5
650         CRPT_ON_HAND_BEG = arr_varCARpt(CR_NUM, lngX)
660       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_ADDL_PROP" Then      '10   10
670         CRPT_ADDL_PROP = arr_varCARpt(CR_NUM, lngX)
680       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_RECEIPTS" Then       '20   20
690         CRPT_RECEIPTS = arr_varCARpt(CR_NUM, lngX)
700       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_GAINS" Then          '30   30
710         CRPT_GAINS = arr_varCARpt(CR_NUM, lngX)
720       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_OTH_CHG" Then        '     40
730         CRPT_OTH_CHG = arr_varCARpt(CR_NUM, lngX)
740       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_NET_INCOME" Then     '40   50
750         CRPT_NET_INCOME = arr_varCARpt(CR_NUM, lngX)
760       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_DISBURSEMENTS" Then  '50   60
770         CRPT_DISBURSEMENTS = arr_varCARpt(CR_NUM, lngX)
780       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_LOSSES" Then         '60   70
790         CRPT_LOSSES = arr_varCARpt(CR_NUM, lngX)
800       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_DISTRIBUTIONS" Then  '80   80
810         CRPT_DISTRIBUTIONS = arr_varCARpt(CR_NUM, lngX)
820       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_OTH_CRED" Then       '     90
830         CRPT_OTH_CRED = arr_varCARpt(CR_NUM, lngX)
840       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_NET_LOSS" Then       '70  100
850         CRPT_NET_LOSS = arr_varCARpt(CR_NUM, lngX)
860       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_ON_HAND_ENDL" Then   '    107
870         CRPT_ON_HAND_ENDL = arr_varCARpt(CR_NUM, lngX)
880       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_Cash_END" Then       '    108
890         CRPT_CASH_END = arr_varCARpt(CR_NUM, lngX)
900       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_NON_Cash_END" Then   '    109
910         CRPT_NON_CASH_END = arr_varCARpt(CR_NUM, lngX)
920       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_ON_HAND_END" Then    '90  110
930         CRPT_ON_HAND_END = arr_varCARpt(CR_NUM, lngX)
940       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_INVEST_INFO" Then    '100 120
950         CRPT_INVEST_INFO = arr_varCARpt(CR_NUM, lngX)
960       ElseIf arr_varCARpt(CR_CON, lngX) = "CRPT_CHANGES" Then        '110 130
970         CRPT_CHANGES = arr_varCARpt(CR_NUM, lngX)
980       End If

990     Next

EXITP:
1000    Set rst = Nothing
1010    Set qdf = Nothing
1020    Set dbs = Nothing
1030    CACourtReportLoad = blnRetVal
1040    Exit Function

ERRH:
1050    blnRetVal = False
1060    Select Case ERR.Number
        Case Else
1070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1080    End Select
1090    Resume EXITP

End Function

Public Function CABuildCourtReportData(ByVal strReportNumber As String, strProc As String, Optional varIsArchive As Variant, Optional varIncludeCheckNum As Variant) As Integer
' ** Called by:
' **   frmRpt_CourtReports_CA.PreviewOrPrint()
' **   frmRpt_CourtReports_CA.SendToFile()
' ** Return Codes:
' **   0  Success.
' **  -1  Canceled.
' **  -9  Error.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "CABuildCourtReportData"

        'Dim rsxDataIn As ADODB.Recordset, rsxDataOut As ADODB.Recordset  ' ** Early binding.
        Dim rsxDataIn As Object, rsxDataOut As Object                     ' ** Late binding.
        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intReportNumber As Integer
        Dim blnIsArchive As Boolean, blnIncludeCheckNum As Boolean
        Dim strTmp01 As String, datTmp02 As Date, lngTmp03 As Long
        Dim lngX As Long
        Dim intRetVal As Integer

1110    intRetVal = 0

1120    CACourtReportLoad  ' ** Function: Above.

1130    strCaseNum = CACourtReportCaseNum  ' ** Function: Below.

1140    Select Case IsMissing(varIsArchive)
        Case True
1150      blnIsArchive = False
1160    Case False
1170      blnIsArchive = CBool(varIsArchive)
1180    End Select

1190    Select Case IsMissing(varIncludeCheckNum)
        Case True
1200      blnIncludeCheckNum = False
1210    Case False
1220      blnIncludeCheckNum = CBool(varIncludeCheckNum)
1230    End Select

        ' ** Delete the data from the tmpCourtReport table.
1240    Set dbs = CurrentDb
1250    With dbs
          ' ** Empty tmpCourtReportData.
1260      Set qdf = .QueryDefs("qryCourtReport_02")
1270      qdf.Execute
1280      .Close
1290    End With
1300    Set qdf = Nothing
1310    Set dbs = Nothing

        'Set rsxDataOut = New ADODB.Recordset             ' ** Early binding.
1320    Set rsxDataOut = CreateObject("ADODB.Recordset")  ' ** Late binding.
1330  On Error Resume Next
1340    rsxDataOut.Open "tmpCourtReportData", CurrentProject.Connection, , adLockOptimistic, adCmdTable
1350    If ERR.Number <> 0 Then
1360      Select Case ERR.Number
          Case -2147217838  ' ** Data source object is already initialized.
1370  On Error GoTo ERRH
            ' ** For now, just let it go, since I think that means it's already available.
1380      Case Else
1390        intRetVal = -9
1400        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410  On Error GoTo ERRH
1420      End Select
1430    Else
1440  On Error GoTo ERRH
1450    End If

1460    If intRetVal = 0 Then

          ' ** Build dummy records with zero in the amount to insure that all report sections are displayed.
          ' ** One for each 10's section in tblCourtReport. Report 5 gets added later in CAGetCourtReportData(), below.
1470      For lngX = 0& To (lngCARpts - 1&)
1480        With rsxDataOut
1490          .AddNew
1500          .Fields("ReportNumber") = arr_varCARpt(CR_NUM, lngX)
1510          .Fields("ReportCategory") = CACourtReportCategory(CInt(arr_varCARpt(CR_NUM, lngX)))  ' ** Function: Below.
1520          .Fields("ReportGroup") = CACourtReportGroup(CInt(arr_varCARpt(CR_NUM, lngX)))  ' ** Function: Below.
1530          .Fields("ReportDivision") = CACourtReportDivision(CInt(arr_varCARpt(CR_NUM, lngX)))  ' ** Function: Below.
1540          .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CInt(arr_varCARpt(CR_NUM, lngX)))  ' ** Function: Below.
1550          .Fields("ReportDivisionText") = CACourtReportDivisionText(CInt(arr_varCARpt(CR_NUM, lngX)))  ' ** Function: Below.
1560          .Fields("ReportGroupText") = CACourtReportGroupText(CInt(arr_varCARpt(CR_NUM, lngX)))  ' ** Function: Below.
1570          .Fields("accountno") = gstrAccountNo
1580          .Fields("date") = gdatEndDate
1590          .Fields("journaltype") = "Miscellaneous"
1600          .Fields("shareface") = 0
1610          .Fields("Description") = "Dummy"
1620          .Fields("Amount") = 0
1630          .Fields("revcode_ID") = 0
1640          .Fields("revcode_DESC") = "Dummy entry"
1650          .Fields("revcode_TYPE") = 1
1660          .Fields("revcode_SORTORDER") = 0
1670          .Fields("ReportSchedule") = CACourtReportSchedule(CInt(arr_varCARpt(CR_NUM, lngX)))  ' ** Function: Below.
1680          .Fields("CaseNum") = strCaseNum
1690          .Update
1700        End With
1710      Next

          ' ** When called from the TAReports CommandBar, these variables should already be filled.
          ' **   gstrCrtRpt_Ordinal
          ' **   gstrCrtRpt_Version
          ' **   gstrCrtRpt_CashAssets_Beg
          ' **   gstrCrtRpt_NetIncome
          ' **   gstrCrtRpt_NetLoss
          ' **   gstrCrtRpt_CashAssets_End

          ' ** Summary.
1720      If strReportNumber = "0" Or strReportNumber = "0A" Then
            ' ** Get entered data.
1730        intRetVal = CAGetCourtReportData(THIS_PROC & "^" & strProc)  ' ** Function: Below.
1740      End If

1750    Else
          ' ** rsxDataOut failed to open.
1760    End If  ' ** intRetVal.

1770    If intRetVal = 0 Then

          'Set rsxDataIn = New ADODB.Recordset             ' ** Early binding.
1780      Set rsxDataIn = CreateObject("ADODB.Recordset")  ' ** Late binding.
1790      Select Case blnIsArchive
          Case True
            ' **** ARCHIVE ****
1800        rsxDataIn.Open "qryCourtReport_CA_21_archive", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
            ' *****************
1810      Case False
1820        rsxDataIn.Open "qryCourtReport_CA_21", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
1830      End Select

          ' ** Loop through data, processing records for requested account.
1840      Do While rsxDataIn.EOF = False
1850        If Trim(rsxDataIn.Fields("accountno")) = gstrAccountNo Then
              ' ** I have no explanation for why this Trim() is necessary! VGC 02/27/2013.
1860          intReportNumber = rsxDataIn.Fields("Reportnumber")  ' ** Do this because it recalcs each time it's referenced.
              ' ** Find the right date to use.
1870          datTmp02 = CACourtReportDate(intReportNumber, rsxDataIn.Fields("transdate"), rsxDataIn.Fields("assetdate"))
              ' ** If the date for the transaction is within range, build a report record.
              'USES assetdate!    ####
1880          strTmp01 = CStr(CDbl(datTmp02))
1890          If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
1900          lngTmp03 = CLng(strTmp01)
1910          If lngTmp03 >= glngStartDateLong And lngTmp03 < (glngEndDateLong + 1@) Then
                ' ** If the journal type is MISC, create 2 transactions.
1920            If rsxDataIn.Fields("Reportnumber") = 0 And rsxDataIn.Fields("journaltype") = "Misc." Then
                  ' ** It'll only come back '0' if CACourtReportDataID(), below, doesn't find a good match.
1930              With rsxDataOut
1940                .AddNew
1950                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
1960                .Fields("date") = datTmp02
1970                .Fields("journaltype") = "Miscellaneous"
1980                .Fields("shareface") = rsxDataIn.Fields("shareface")
1990                .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                      rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                      rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
2000                .Fields("Amount") = rsxDataIn.Fields("icash")
2010                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2020                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2030                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2040                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2050                .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
2060                .Fields("CaseNum") = strCaseNum
2070                If rsxDataIn.Fields("icash") > 0 Then
                      ' ** Does this '7' have any relation to Report 70 (Net Loss), or is it just for sort order?
2080                  .Fields("ReportNumber") = 7
2090                  .Fields("ReportCategory") = CACourtReportCategory(7)  ' ** This'll come back 'Unknown'!
2100                  .Fields("ReportGroup") = CACourtReportGroup(7)  ' ** This'll come back 0!
2110                  .Fields("ReportDivision") = CACourtReportDivision(7)  ' ** This'll come back 0!
2120                  .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(7)  ' ** This'll come back 'Unknown'!
2130                  .Fields("ReportDivisionText") = CACourtReportDivisionText(7)  ' ** This'll come back 'Unknown'!
2140                  .Fields("ReportGroupText") = CACourtReportGroupText(7)  ' ** This'll come back 'Unknown'!
2150                Else
                      ' ** Does this '8' have any relation to Report 80 (Distributions), or is it just for sort order?
2160                  .Fields("ReportNumber") = 8
2170                  .Fields("ReportCategory") = CACourtReportCategory(8)  ' ** This'll come back 'Unknown'!
2180                  .Fields("ReportGroup") = CACourtReportGroup(8)  ' ** This'll come back 0!
2190                  .Fields("ReportDivision") = CACourtReportDivision(8)  ' ** This'll come back 0!
2200                  .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(8)  ' ** This'll come back 'Unknown'!
2210                  .Fields("ReportDivisionText") = CACourtReportDivisionText(8)  ' ** This'll come back 'Unknown'!
2220                  .Fields("ReportGroupText") = CACourtReportGroupText(8)  ' ** This'll come back 'Unknown'!
2230                End If
2240                .Update
2250                .AddNew
2260                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2270                .Fields("date") = datTmp02
2280                .Fields("journaltype") = rsxDataIn.Fields("journaltype")
2290                .Fields("shareface") = rsxDataIn.Fields("shareface")
2300                .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                      rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                      rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
2310                .Fields("Amount") = rsxDataIn.Fields("icash")
2320                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2330                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2340                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2350                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2360                .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
2370                .Fields("CaseNum") = strCaseNum
2380                If rsxDataIn.Fields("pcash") > 0 Then
                      ' ** Does this '1' have any relation to Report 10 (Additional Property Received), or is it just for sort order?
2390                  .Fields("ReportNumber") = 1
2400                  .Fields("ReportCategory") = CACourtReportCategory(1)  ' ** This'll come back 'Unknown'!
2410                  .Fields("ReportGroup") = CACourtReportGroup(1)  ' ** This'll come back 0!
2420                  .Fields("ReportDivision") = CACourtReportDivision(1)  ' ** This'll come back 0!
2430                  .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(1)  ' ** This'll come back 'Unknown'!
2440                  .Fields("ReportDivisionText") = CACourtReportDivisionText(1)  ' ** This'll come back 'Unknown'!
2450                  .Fields("ReportGroupText") = CACourtReportGroupText(1)  ' ** This'll come back 'Unknown'!
2460                Else
                      ' ** Does this '3' have any relation to Report 30 (Gains on Sales), or is it just for sort order?
2470                  .Fields("ReportNumber") = 3
2480                  .Fields("ReportCategory") = CACourtReportCategory(3)  ' ** This'll come back 'Unknown'!
2490                  .Fields("ReportGroup") = CACourtReportGroup(3)  ' ** This'll come back 0!
2500                  .Fields("ReportDivision") = CACourtReportDivision(3)  ' ** This'll come back 0!
2510                  .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(3)  ' ** This'll come back 'Unknown'!
2520                  .Fields("ReportDivisionText") = CACourtReportDivisionText(3)  ' ** This'll come back 'Unknown'!
2530                  .Fields("ReportGroupText") = CACourtReportGroupText(3)  ' ** This'll come back 'Unknown'!
2540                End If
2550                .Update
2560              End With  ' ** rsxDataOut.
2570            Else
                  ' ** Handle journal type special cases here.
2580              With rsxDataOut
2590                Select Case rsxDataIn.Fields("journaltype")

                    Case "Purchase"
                      'PURCHASE: CHECK FOR EXPENSES CODED 'Unspecified Income'
                      'ALL 'Purchase' ARE EXPENSES!
                      'Purchase, Deposit, Liability -icash: Expense
                      ' ** Always add the purchase record.
                      ' ** intReportNumber is already a 10's number.
2600                  .AddNew
2610                  .Fields("ReportNumber") = intReportNumber
2620                  .Fields("ReportCategory") = CACourtReportCategory(intReportNumber)
2630                  .Fields("ReportGroup") = CACourtReportGroup(intReportNumber)
2640                  .Fields("ReportDivision") = CACourtReportDivision(intReportNumber)
2650                  .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(intReportNumber)
2660                  .Fields("ReportDivisionText") = CACourtReportDivisionText(intReportNumber)
2670                  .Fields("ReportGroupText") = CACourtReportGroupText(intReportNumber)
2680                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2690                  .Fields("date") = datTmp02
2700                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
2710                  .Fields("shareface") = rsxDataIn.Fields("shareface")
2720                  .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                        rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                        rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
                      ' ** 07/17/2008: Added this special case Purchase (correctly, I hope!), per Rich.
2730                  If rsxDataIn.Fields("icash") <> 0 And rsxDataIn.Fields("pcash") <> 0 Then
2740                    If intReportNumber = CRPT_INVEST_INFO Then
2750                      .Fields("Amount") = _
                            ((IIf(rsxDataIn.Fields("icash") = 0, rsxDataIn.Fields("pcash"), _
                            IIf(rsxDataIn.Fields("pcash") = 0, rsxDataIn.Fields("icash"), _
                            rsxDataIn.Fields("pcash")))) * -1)
2760                    Else
2770                      .Fields("Amount") = CACourtReportDollars(intReportNumber, 0@, rsxDataIn.Fields("icash"), _
                            rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
2780                    End If
2790                  Else
2800                    .Fields("Amount") = CACourtReportDollars(intReportNumber, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
2810                  End If
2820                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2830                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2840                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2850                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2860                  .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
2870                  .Fields("CaseNum") = strCaseNum
2880                  .Update
                      ' ** Force an Information For Investments Made record from other Purchases.
2890                  If intReportNumber <> CRPT_INVEST_INFO Then
                        'If rsxDataIn.Fields("icash") < 0 And rsxDataIn.Fields("pcash") < 0 Then
                        ' ** 01/24/2009: Adding another special case, though not yet checked with Rich.
                        ' ** TOTALLY CONFUSED ABOUT ReportNumber!!!
                        ' **   strReportNumber: a single digit, 0-11, indicating each of the 12 reports, 00-11.
                        ' **   intReportNumber: found in tblCourtReport, these are internal numbers set in 20 constants,
                        ' **                    like the one above, numbered 2-130. See CACourtReportLoad(), above.
                        ' ** This intReportNumber is of the 2nd type, in this case: CRPT_DISBURSEMENTS, 60.
                        '.Fields("Amount") = _
                        '  (IIf(.Fields("icash") = 0, .Fields("pcash"), IIf(.Fields("pcash") = 0, .Fields("icash"), .Fields("pcash")))) * -1
                        ' ** 01/24/2009: New formula for Amount: =(IIf([icash]=0,[pcash],IIf([pcash]=0,[icash],[pcash])))*-1
                        ' ** If the original (icash + pcash) was intended to simplify choosing whichever column has the
                        ' ** value (since most Purchases only have a value in one or the other), no provision was given
                        ' ** to the case when both have a value. This new formula, then, chooses pcash over both.
                        ' ** 2 examples: 10/18/2007, ($17.18) ($9,845.28); 01/22/2009, ($100.00) ($1,000.00)
                        ' ** Do taxcode's matter?            4 (Non-Taxable)                    2 (Fed. Tax Only)
                        ' ** Both have a revcode_ID of 1, Unspecified Income  ?
                        ' ** See also qryCourtReport_CA_00_A2_01.
2900                    .AddNew
2910                    .Fields("ReportNumber") = 120
2920                    .Fields("ReportCategory") = CACourtReportCategory(120)
2930                    .Fields("ReportGroup") = CACourtReportGroup(120)
2940                    .Fields("ReportDivision") = CACourtReportDivision(120)
2950                    .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(120)
2960                    .Fields("ReportDivisionText") = CACourtReportDivisionText(120)
2970                    .Fields("ReportGroupText") = CACourtReportGroupText(120)
2980                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2990                    .Fields("date") = datTmp02
3000                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3010                    .Fields("shareface") = rsxDataIn.Fields("shareface")
3020                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
3030                    .Fields("Amount") = _
                          ((IIf(rsxDataIn.Fields("icash") = 0, rsxDataIn.Fields("pcash"), _
                          IIf(rsxDataIn.Fields("pcash") = 0, rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("pcash")))) * -1)
3040                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3050                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3060                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3070                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3080                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
3090                    .Fields("CaseNum") = strCaseNum
3100                    .Update
3110                  End If
                      ' ** Force a Change in Investment Holdings record from other Purchases.
3120                  If intReportNumber <> CRPT_CHANGES Then
3130                    .AddNew
3140                    .Fields("ReportNumber") = 130
3150                    .Fields("ReportCategory") = CACourtReportCategory(130)
3160                    .Fields("ReportGroup") = CACourtReportGroup(130)
3170                    .Fields("ReportDivision") = CACourtReportDivision(130)
3180                    .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(130)
3190                    .Fields("ReportDivisionText") = CACourtReportDivisionText(130)
3200                    .Fields("ReportGroupText") = CACourtReportGroupText(130)
3210                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3220                    .Fields("date") = datTmp02
3230                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3240                    .Fields("shareface") = rsxDataIn.Fields("shareface")
3250                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
3260                    .Fields("Amount") = _
                          ((IIf(rsxDataIn.Fields("icash") = 0, rsxDataIn.Fields("pcash"), _
                          IIf(rsxDataIn.Fields("pcash") = 0, rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("pcash")))) * -1)
3270                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3280                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3290                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3300                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3310                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
3320                    .Fields("CaseNum") = strCaseNum
3330                    .Update
3340                  End If

3350                Case "Sold"
3360                  If rsxDataIn.Fields("icash") > 0 And (rsxDataIn.Fields("icash") <> -rsxDataIn.Fields("Cost")) Then
                        ' ** Add Sold into Interest.
                        ' ** intReportNumber is already a 10's number.
3370                    .AddNew
3380                    .Fields("ReportNumber") = intReportNumber
3390                    .Fields("ReportCategory") = CACourtReportCategory(intReportNumber)
3400                    .Fields("ReportGroup") = CACourtReportGroup(intReportNumber)
3410                    .Fields("ReportDivision") = CACourtReportDivision(intReportNumber)
3420                    .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(intReportNumber)
3430                    .Fields("ReportDivisionText") = CACourtReportDivisionText(intReportNumber)
3440                    .Fields("ReportGroupText") = CACourtReportGroupText(intReportNumber)
3450                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3460                    .Fields("date") = datTmp02
3470                    .Fields("journaltype") = "Interest"
3480                    .Fields("shareface") = rsxDataIn.Fields("shareface")
3490                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
3500                    .Fields("Amount") = CACourtReportDollars(intReportNumber, 0, rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
3510                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3520                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3530                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3540                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3550                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
3560                    .Fields("CaseNum") = strCaseNum
3570                    .Update
3580                  End If
3590                  If rsxDataIn.Fields("Pcash") = 0 And (rsxDataIn.Fields("icash") = -rsxDataIn.Fields("Cost")) Then
3600                    .AddNew  ' ** Force a Change in Investment Holdings record.
3610                    .Fields("ReportNumber") = 130
3620                    .Fields("ReportCategory") = CACourtReportCategory(130)
3630                    .Fields("ReportGroup") = CACourtReportGroup(130)
3640                    .Fields("ReportDivision") = CACourtReportDivision(130)
3650                    .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(130)
3660                    .Fields("ReportDivisionText") = CACourtReportDivisionText(130)
3670                    .Fields("ReportGroupText") = CACourtReportGroupText(130)
3680                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3690                    .Fields("date") = datTmp02
3700                    .Fields("journaltype") = "Interest"
3710                    .Fields("shareface") = rsxDataIn.Fields("shareface")
3720                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
3730                    .Fields("Amount") = rsxDataIn.Fields("cost")
3740                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3750                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3760                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3770                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3780                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
3790                    .Fields("CaseNum") = strCaseNum
3800                    .Update
3810                  Else
                        ' ** Always add the sold record.
3820                    .AddNew
3830                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3840                    .Fields("date") = datTmp02
3850                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3860                    .Fields("shareface") = rsxDataIn.Fields("shareface")
3870                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
3880                    .Fields("Amount") = CACourtReportDollars(CRPT_GAINS, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
3890                    Select Case .Fields("Amount")
                        Case Is > 0
3900                      .Fields("ReportNumber") = CRPT_GAINS
3910                      .Fields("ReportCategory") = CACourtReportCategory(CRPT_GAINS)
3920                      .Fields("ReportGroup") = CACourtReportGroup(CRPT_GAINS)
3930                      .Fields("ReportDivision") = CACourtReportDivision(CRPT_GAINS)
3940                      .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_GAINS)
3950                      .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_GAINS)
3960                      .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_GAINS)
3970                      .Fields("Amount") = CACourtReportDollars(CRPT_GAINS, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                            rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
3980                    Case Is = 0
3990                      .Fields("ReportNumber") = CRPT_CHANGES
4000                      .Fields("ReportCategory") = CACourtReportCategory(CRPT_CHANGES)
4010                      .Fields("ReportGroup") = CACourtReportGroup(CRPT_CHANGES)
4020                      .Fields("ReportDivision") = CACourtReportDivision(CRPT_CHANGES)
4030                      .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_CHANGES)
4040                      .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_CHANGES)
4050                      .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_CHANGES)
                          ' ** Reset the date for this report.
4060                      .Fields("date") = CACourtReportDate(CRPT_CHANGES, rsxDataIn.Fields("transdate"), rsxDataIn.Fields("assetdate"))
                          ' ** Reset the amount for the report.
4070                      .Fields("Amount") = CACourtReportDollars(CRPT_CHANGES, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                            rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
4080                    Case Is < 0
4090                      .Fields("ReportNumber") = CRPT_LOSSES
4100                      .Fields("ReportCategory") = CACourtReportCategory(CRPT_LOSSES)
4110                      .Fields("ReportGroup") = CACourtReportGroup(CRPT_LOSSES)
4120                      .Fields("ReportDivision") = CACourtReportDivision(60)
4130                      .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_LOSSES)
4140                      .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_LOSSES)
4150                      .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_LOSSES)
4160                      .Fields("Amount") = CACourtReportDollars(CRPT_LOSSES, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                            rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
4170                    End Select
4180                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4190                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4200                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4210                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4220                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
4230                    .Fields("CaseNum") = strCaseNum
4240                    .Update
                        'End If
4250                  End If

4260                Case "Liability"
                      'Liability Criteria:
                      'Liability -icash/pcash/+cost: Expense
                      'Liability +icash/pcash/-cost: Income
4270                  If ((rsxDataIn.Fields("cost") < 0) And (rsxDataIn.Fields("cost") = (-rsxDataIn.Fields("pcash")))) Then
                        ' ** Information for Investments Made - Report 5; Case 50 = "Investments Made"  'VGC 02/20/2013: PER RICH!
                        ' ** gdblCrtRpt_CA_InvestInfo: qryCourtReport_CA_00_A2_02
                        'Debug.Print "'HERE! 2  " & CRPT_INVEST_INFO
4280                    .AddNew  ' ** Force an Investments Made record.
4290                    .Fields("ReportNumber") = CRPT_INVEST_INFO
4300                    .Fields("ReportCategory") = "Investments Made During Period of Account"
4310                    .Fields("ReportGroup") = 60
4320                    .Fields("ReportDivision") = 60
4330                    .Fields("ReportDivisionTitle") = "ADDITIONAL INFORMATION"
4340                    .Fields("ReportDivisionText") = "xxxxx"
4350                    .Fields("ReportGroupText") = vbNullString
4360                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4370                    .Fields("date") = rsxDataIn.Fields("transdate")
4380                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4390                    .Fields("shareface") = rsxDataIn.Fields("shareface")
4400                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
4410                    .Fields("Amount") = rsxDataIn.Fields("cost")
4420                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4430                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4440                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4450                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4460                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
4470                    .Fields("CaseNum") = strCaseNum
4480                    .Update
4490                  End If
4500                  If ((rsxDataIn.Fields("cost") > 0) And (rsxDataIn.Fields("cost") = (-rsxDataIn.Fields("pcash")))) Then
                        ' ** Changes in Investment Holdings - Report 6; Case 60 = "Changes in Investment Holdings"  'VGC 02/20/2013: PER RICH!
                        ' ** gdblCrtRpt_CA_InvestChange: qryCourtReport_CA_00_A3_02
                        'Debug.Print "'HERE! 1  " & CRPT_CHANGES
4510                    .AddNew  ' ** Force a Changes in Investment Holdings record.
4520                    .Fields("ReportNumber") = CRPT_CHANGES
4530                    .Fields("ReportCategory") = "Changes in Investment Holdings During Period of Account"
4540                    .Fields("ReportGroup") = 60
4550                    .Fields("ReportDivision") = 60
4560                    .Fields("ReportDivisionTitle") = "ADDITIONAL INFORMATION"
4570                    .Fields("ReportDivisionText") = "xxxxx"
4580                    .Fields("ReportGroupText") = vbNullString
4590                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4600                    .Fields("date") = datTmp02
4610                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4620                    .Fields("shareface") = rsxDataIn.Fields("shareface")
4630                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
4640                    .Fields("Amount") = rsxDataIn.Fields("cost")
4650                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4660                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4670                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4680                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4690                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
4700                    .Fields("CaseNum") = strCaseNum
4710                    .Update
4720                  End If
4730                  If rsxDataIn.Fields("icash") < 0 Then
                        ' ** Add Icash as a disbursement.
4740                    .AddNew
4750                    .Fields("ReportNumber") = CRPT_DISBURSEMENTS
4760                    .Fields("ReportCategory") = CACourtReportCategory(CRPT_DISBURSEMENTS)
4770                    .Fields("ReportGroup") = CACourtReportGroup(CRPT_DISBURSEMENTS)
4780                    .Fields("ReportDivision") = CACourtReportDivision(CRPT_DISBURSEMENTS)
4790                    .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_DISBURSEMENTS)
4800                    .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_DISBURSEMENTS)
4810                    .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_DISBURSEMENTS)
4820                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4830                    .Fields("date") = datTmp02
4840                    .Fields("journaltype") = "Liability"
4850                    .Fields("shareface") = rsxDataIn.Fields("shareface")
4860                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
4870                    .Fields("Amount") = rsxDataIn.Fields("icash") * -1
4880                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4890                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4900                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4910                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4920                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
4930                    .Fields("CaseNum") = strCaseNum
4940                    .Update
4950                  Else
                        ' ** Always add the liability record.
                        ' ** intReportNumber is already a 10's number.
4960                    .AddNew
4970                    .Fields("ReportNumber") = intReportNumber
4980                    .Fields("ReportCategory") = CACourtReportCategory(intReportNumber)
4990                    .Fields("ReportGroup") = CACourtReportGroup(intReportNumber)
5000                    .Fields("ReportDivision") = CACourtReportDivision(intReportNumber)
5010                    .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(intReportNumber)
5020                    .Fields("ReportDivisionText") = CACourtReportDivisionText(intReportNumber)
5030                    .Fields("ReportGroupText") = CACourtReportGroupText(intReportNumber)
5040                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
5050                    .Fields("date") = datTmp02
5060                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
5070                    .Fields("shareface") = rsxDataIn.Fields("shareface")
5080                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
5090                    .Fields("Amount") = CACourtReportDollars(intReportNumber, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
5100                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
5110                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
5120                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5130                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5140                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
5150                    .Fields("CaseNum") = strCaseNum
5160                    .Update
5170                  End If

5180                Case "Cost Adj."
5190                  .AddNew  ' ** Force an Additional Property Received record.
5200                  .Fields("ReportNumber") = 10
5210                  .Fields("ReportCategory") = CACourtReportCategory(10)
5220                  .Fields("ReportGroup") = CACourtReportGroup(10)
5230                  .Fields("ReportDivision") = CACourtReportDivision(10)
5240                  .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(10)
5250                  .Fields("ReportDivisionText") = CACourtReportDivisionText(10)
5260                  .Fields("ReportGroupText") = CACourtReportGroupText(10)
5270                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
5280                  .Fields("date") = datTmp02
5290                  .Fields("journaltype") = "Interest"
5300                  .Fields("shareface") = rsxDataIn.Fields("shareface")
5310                  .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                        rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                        rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
5320                  .Fields("Amount") = rsxDataIn.Fields("cost")
5330                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
5340                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
5350                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5360                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5370                  .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
5380                  .Fields("CaseNum") = strCaseNum
5390                  .Update

5400                Case Else  ' ** All other journaltypes go through here.
                      'WITHDRAWN: CHECK FOR EXPENSES CODED 'Unspecified Income'
                      'MISC.: CHECK FOR EXPENSES CODED 'Unspecified Income'
                      'Purchase, Deposit, Liability -icash: Expense
                      'Withdrawn?
                      ' ** intReportNumber is already a 10's number.
5410                  .AddNew
5420                  .Fields("ReportNumber") = intReportNumber
5430                  .Fields("ReportCategory") = CACourtReportCategory(intReportNumber)
5440                  .Fields("ReportGroup") = CACourtReportGroup(intReportNumber)
5450                  .Fields("ReportDivision") = CACourtReportDivision(intReportNumber)
5460                  .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(intReportNumber)
5470                  .Fields("ReportDivisionText") = CACourtReportDivisionText(intReportNumber)
5480                  .Fields("ReportGroupText") = CACourtReportGroupText(intReportNumber)
5490                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
5500                  .Fields("date") = datTmp02
5510                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
5520                  .Fields("shareface") = rsxDataIn.Fields("shareface")
5530                  Select Case blnIncludeCheckNum
                      Case True
5540                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"), rsxDataIn.Fields("CheckNum"))  ' ** Module Function: modReportFunctions.
5550                  Case False
5560                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                          rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                          rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
5570                  End Select
5580                  .Fields("Amount") = CACourtReportDollars(intReportNumber, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                        rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
5590                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
5600                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
5610                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5620                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5630                  .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
5640                  .Fields("CaseNum") = strCaseNum
5650                  .Update

5660                End Select  ' ** journaltype.
5670              End With  ' ** rsxDataOut.
5680            End If  ' ** journaltype.

                ' ** Other Charges.
5690            If rsxDataIn.Fields("revcode_ID") = REVID_OCHG Then
5700              With rsxDataOut
5710                .AddNew
5720                .Fields("ReportNumber") = 40
5730                .Fields("ReportCategory") = CACourtReportCategory(40)
5740                .Fields("ReportGroup") = CACourtReportGroup(40)
5750                .Fields("ReportDivision") = CACourtReportDivision(40)
5760                .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(40)
5770                .Fields("ReportDivisionText") = CACourtReportDivisionText(40)
5780                .Fields("ReportGroupText") = CACourtReportGroupText(40)
5790                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
5800                .Fields("date") = datTmp02
5810                .Fields("journaltype") = rsxDataIn.Fields("journaltype")
5820                .Fields("shareface") = rsxDataIn.Fields("shareface")
5830                .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                      rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                      rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
5840                .Fields("Amount") = CACourtReportDollars(40, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                      rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
5850                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
5860                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
5870                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5880                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5890                .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
5900                .Fields("CaseNum") = strCaseNum
5910                .Update
5920              End With  ' ** rsxDataOut.
5930            End If  ' ** lngOtherChargesID.

                ' ** Other Credits.
5940            If rsxDataIn.Fields("revcode_ID") = REVID_OCRED Then
5950              With rsxDataOut
5960                .AddNew
5970                .Fields("ReportNumber") = 90
5980                .Fields("ReportCategory") = CACourtReportCategory(90)
5990                .Fields("ReportGroup") = CACourtReportGroup(90)
6000                .Fields("ReportDivision") = CACourtReportDivision(90)
6010                .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(90)
6020                .Fields("ReportDivisionText") = CACourtReportDivisionText(90)
6030                .Fields("ReportGroupText") = CACourtReportGroupText(90)
6040                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
6050                .Fields("date") = datTmp02
6060                .Fields("journaltype") = rsxDataIn.Fields("journaltype")
6070                .Fields("shareface") = rsxDataIn.Fields("shareface")
6080                .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), _
                      rsxDataIn.Fields("Description"), rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), _
                      rsxDataIn.Fields("jComment"))  ' ** Module Function: modReportFunctions.
6090                .Fields("Amount") = CACourtReportDollars(90, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                      rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
6100                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
6110                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
6120                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
6130                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
6140                .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
6150                .Fields("CaseNum") = strCaseNum
6160                .Update
6170              End With  ' ** rsxDataOut.
6180            End If  ' ** lngOtherCreditsID.

6190          End If  ' ** glngStartDateLong, glngEndDateLong.
6200        End If  ' ** gstrAccountNo.
6210        rsxDataIn.MoveNext
6220      Loop
6230      rsxDataIn.Close

          ' ** Get asset information.
6240      rsxDataIn.Open "qryAssetList-CA", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
6250      intReportNumber = CRPT_ON_HAND_END

          ' ** Loop through data, processing records for requested account.
6260      Do While rsxDataIn.EOF = False
6270        If Trim(rsxDataIn.Fields("accountno")) = gstrAccountNo Then
6280          With rsxDataOut
                ' ** intReportNumber is already a 10's number.
6290            .AddNew
6300            .Fields("ReportNumber") = intReportNumber
6310            .Fields("ReportCategory") = CACourtReportCategory(intReportNumber)
6320            .Fields("ReportGroup") = CACourtReportGroup(intReportNumber)
6330            .Fields("ReportDivision") = CACourtReportDivision(intReportNumber)
6340            .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(intReportNumber)
6350            .Fields("ReportDivisionText") = CACourtReportDivisionText(intReportNumber)
6360            .Fields("ReportGroupText") = CACourtReportGroupText(intReportNumber)
6370            .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
6380            .Fields("date") = gdatStartDate  ' ** Use this date to get by the filter on the report.
6390            .Fields("journaltype") = "Asset"
6400            .Fields("Description") = rsxDataIn.Fields("totdesc")
6410            .Fields("Amount") = rsxDataIn.Fields("TotalCost")
6420            .Fields("revcode_ID") = 0
6430            .Fields("revcode_DESC") = "Dummy entry"
6440            .Fields("revcode_TYPE") = 1
6450            .Fields("revcode_SORTORDER") = 0
6460            .Fields("ReportSchedule") = CACourtReportSchedule(intReportNumber)
6470            .Fields("CaseNum") = strCaseNum
6480            .Update
6490          End With
6500        End If
6510        rsxDataIn.MoveNext
6520      Loop

6530      rsxDataIn.Close
6540      rsxDataIn.Open "account", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
6550      intReportNumber = CRPT_ON_HAND_END

          ' ** Loop through data, processing records for requested account.
6560      Do While rsxDataIn.EOF = False
6570        If Trim(rsxDataIn.Fields("accountno")) = gstrAccountNo Then
6580          With rsxDataOut
                ' ** intReportNumber is already a 10's number.
6590            .AddNew
6600            .Fields("ReportNumber") = intReportNumber
6610            .Fields("ReportCategory") = CACourtReportCategory(intReportNumber)
6620            .Fields("ReportGroup") = CACourtReportGroup(intReportNumber)
6630            .Fields("ReportDivision") = CACourtReportDivision(intReportNumber)
6640            .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(intReportNumber)
6650            .Fields("ReportDivisionText") = CACourtReportDivisionText(intReportNumber)
6660            .Fields("ReportGroupText") = CACourtReportGroupText(intReportNumber)
6670            .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
6680            .Fields("date") = gdatStartDate  ' ** Use this date to get by the filter on the report.
6690            .Fields("journaltype") = "Asset"
6700            .Fields("Description") = "Account info"
6710            .Fields("Amount") = rsxDataIn.Fields("pcash") + IIf(IsNull(rsxDataIn.Fields("icash")), 0, rsxDataIn.Fields("icash"))
6720            .Fields("revcode_ID") = 0
6730            .Fields("revcode_DESC") = "Dummy entry"
6740            .Fields("revcode_TYPE") = 1
6750            .Fields("revcode_SORTORDER") = 0
6760            .Fields("ReportSchedule") = CACourtReportSchedule(intReportNumber)
6770            .Fields("CaseNum") = strCaseNum
6780            .Update
6790          End With
6800        End If
6810        rsxDataIn.MoveNext
6820      Loop

6830      rsxDataOut.Close

6840    End If

EXITP:
6850    Set rsxDataIn = Nothing
6860    Set rsxDataOut = Nothing
6870    Set qdf = Nothing
6880    Set dbs = Nothing
6890    CABuildCourtReportData = intRetVal
6900    Exit Function

ERRH:
6910    intRetVal = -9
6920    Select Case ERR.Number
        Case Else
6930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6940    End Select
6950    Resume EXITP

End Function

Public Function CACourtReportCategory(intCourtReport As Integer) As String
' ** HOW COME SO MANY OF THE LINES CALLING THIS USE THE REPORT NUMBER ALONE WITHOUT MULTIPLYING BY 10?
' ** THEY'LL ALL COME BACK 'Unknown'?
'modCourtReportsCA.CAGetCourtReportData()
'modCourtReportsCA.CABuildCourtReportData()

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportCategory"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

7010    strRetVal = vbNullString

7020    blnFound = False
7030    For lngX = 0& To (lngCARpts - 1&)
7040      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
7050        strRetVal = arr_varCARpt(CR_CAT, lngX)
7060        blnFound = True
7070        Exit For
7080      End If
7090    Next
7100    If blnFound = False Then strRetVal = "Unknown"

        'Select Case intCourtReport
        'Case 5
        '  strRetVal = "Property on Hand at Beginning of Account Period"
        'Case 10
        '  strRetVal = "Additional Property Received During Period of Account"
        'Case 20
        '  strRetVal = "Receipts During Period of Account"
        'Case 30
        '  strRetVal = "Gains on Sale During Period of Account"
        'Case 40
        '  strRetVal = "Net Income from Trade or Business During Period of Account"
        'Case 50
        '  strRetVal = "Disbursements During Period of Account"
        'Case 60
        '  strRetVal = "Losses on Sale During Period of Account"
        'Case 70
        '  strRetVal = "Net Loss from Trade or Business During Period of Account"
        'Case 80
        '  strRetVal = "Distributions During Period of Account"
        'Case 90
        '  strRetVal = "Property on Hand at Close of Account Period"
        'Case 100
        '  strRetVal = "Investments Made During Period of Account"
        'Case 110
        '  strRetVal = "Changes in Investment Holdings During Period of Account"
        'Case Else
        '  strRetVal = "Unknown"
        'End Select

EXITP:
7110    CACourtReportCategory = strRetVal
7120    Exit Function

ERRH:
7130    strRetVal = RET_ERR
7140    Select Case ERR.Number
        Case Else
7150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7160    End Select
7170    Resume EXITP

End Function

Public Function CACourtReportDataID(strComment As String, curPCash As Currency, curICash As Currency, curCost As Currency, curGainLoss As Currency, strTaxCode As String, strJournalType As String) As Integer
' ** The code returned is the Court Report that the data belongs to.
' ** Used in queries:
'QRY: 'qryCourtReport - Summary-1-CA' CACourtReportDataID([jcomment  NOT USED ANYWHERE!!!
'QRY: 'qryCourtReport - Summary-1-NY' CACourtReportDataID([jcomment  NOT USED
'QRY: 'qryCourtReport_CA_21' CACourtReportDataID([jcomment           SUPPLIES DATA FOR CABuildCourtReportData()!
'QRY: 'qryCourtReport_CA_21x' CACourtReportDataID([jcomment

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportDataID"

        Dim lngTaxcode As Long
        Dim intRetVal As Integer

7210    intRetVal = 0  ' ** Force to 0 just in case.

7220    If glngTaxCode_Distribution = 0& Then
7230      glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
7240    End If
7250    If Trim(strTaxCode) <> vbNullString Then
7260      lngTaxcode = Val(strTaxCode)
7270    Else
7280      lngTaxcode = 0&
7290    End If

        ' ** Additional Property Received.  'VGC 08/20/2010: CHANGES!
        ' ** All Liabilities, both sides, included here, per Rich.  'VGC 08/26/2010: CHANGES!
7300    If (strJournalType = "Deposit" And Not (strComment Like "*stock split*")) Or _
            (strJournalType = "Cost Adj." And curCost > 0) Or _
            (strJournalType = "Liability") Then
7310      intRetVal = CRPT_ADDL_PROP
7320    Else
          'IT APPEARS THAT ANYTHING CAUGHT ABOVE WILL NEVER, EVER GET DOWN BELOW!!!!
          'HOW CAN WE DO THAT???????????????????????????????
          '### PURCHASE: ICASH>0 OR ICASH/PCASH<0,                               NEVER GET TO INVESTINFO OR CHANGES!!
          'CABuildCourtReportData() DOES INVESTINFO, BUT DIDN'T DO CHANGES, NOW IT DOES!

          '### SOLD: ICASH>0, GAINLOSS>0 OR GAINLOSS<0 OR GAINLOSS=0, IF ICASH>0 NEVER GETS TO GAINS OR LOSSES OR CHANGES!!!
          'CABuildCourtReportData() HANDLES 1ST CRITERIA SEPARATELY, SO ALWAYS GOES ON TO NEXT 3!

          'DEPOSIT: NON-STOCKSPLIT OR STOCKSPLIT           - OK, MUTUALLY EXCLUSIVE
          'WITHDRAWN: <>11 OR =11                          - OK, MUTUALLY EXCLUSIVE
          'COST ADJ: COST>0 OR COST<0                      - OK, MUTUALLY EXCLUSIVE
          'PAID: PCASH/ICASH<>0/<>11 OR PCASH/ICASH<>0/=11 - OK, MUTUALLY EXCLUSIVE
          'MISC: PCASH+ICASH>0 OR PCASH+ICASH<0            - OK, MUTUALLY EXCLUSIVE
          'DIVIDEND: ICASH>0       - OK, ONLY ONE
          'INTEREST: ICASH>0       - OK, ONLY ONE
          'RECEIVED: ICASH/PCASH>0 - OK, ONLY ONE

          ' ** Receipts.
7330      If (strJournalType = "Received" And (curPCash > 0 Or curICash > 0)) Or (strJournalType = "Misc." And (curPCash + curICash > 0)) _
              Or (strJournalType = "Dividend" And curICash > 0) Or (strJournalType = "Purchase" And curICash > 0) _
              Or (strJournalType = "Sold" And curICash > 0) Or (strJournalType = "Interest" And curICash > 0) Then
            '### PURCHASE: ICASH>0 OR ICASH/PCASH<0, NEVER GET TO INVESTINFO OR CHANGES!!
            '### SOLD: ICASH>0 , GAINLOSS>0 OR GAINLOSS<0 OR GAINLOSS=0, IF ICASH>0 NEVER GETS BEYOND!!!
7340        intRetVal = CRPT_RECEIPTS
7350      Else
            ' ** Gains on Sales.
            'VGC 08/26/2010: Will try adding 'Cost Adj.' like LOSSES to see if it clears anything up.
7360        If (strJournalType = "Sold" And curGainLoss > 0) Or (strJournalType = "Cost Adj." And curCost > 0) Then
              'If (strJournalType = "Sold" And curGainLoss > 0) Then
              ' ** GainLoss: ([ledger.pcash]-([ledger.cost]*-1))
7370          intRetVal = CRPT_GAINS
7380        Else
              ' ** Net Income is entered by hand (intRetVal = CRPT_NET_INCOME).
              ' ** Disbursements.
              ' ** 07/17/08: Added negative Purchases, per Rich.
              ' **           Moved negative Cost Adjustments to Losses, per Rich.
7390          If ((strJournalType = "Paid" And curPCash <> 0 And lngTaxcode <> glngTaxCode_Distribution) _
                  Or (strJournalType = "Paid" And curICash <> 0 And lngTaxcode <> glngTaxCode_Distribution) _
                  Or (strJournalType = "Misc." And (curPCash + curICash < 0)) _
                  Or (strJournalType = "Withdrawn" And lngTaxcode <> glngTaxCode_Distribution) _
                  Or (strJournalType = "Liability" And curICash < 0) _
                  Or (strJournalType = "Purchase" And curICash < 0 And curPCash < 0)) Then  '<> "Distribution"
                '### PURCHASE: ICASH>0 OR ICASH/PCASH<0, NEVER GET TO INVESTINFO OR CHANGES!!
7400            intRetVal = CRPT_DISBURSEMENTS    '####  TAXCODE  ####
7410          Else
                ' ** Losses on Sales.
                ' ** 07/17/08: Added negative Cost Adjustments, per Rich.
7420            If (strJournalType = "Sold" And curGainLoss < 0) Or (strJournalType = "Cost Adj." And curCost < 0) Then
                  ' ** GainLoss: ([ledger.pcash]-([ledger.cost]*-1))
7430              intRetVal = CRPT_LOSSES
7440            Else
                  ' ** Net Loss is entered by hand (intRetVal = CRPT_NET_LOSS).
                  ' ** Distributions.
7450              If ((strJournalType = "Paid" And curPCash <> 0 And lngTaxcode = glngTaxCode_Distribution) _
                      Or (strJournalType = "Paid" And curICash <> 0 And lngTaxcode = glngTaxCode_Distribution) _
                      Or (strJournalType = "Withdrawn" And lngTaxcode = glngTaxCode_Distribution)) Then  '= "Distribution"
7460                intRetVal = CRPT_DISTRIBUTIONS    '####  TAXCODE  ####
7470              Else
                    ' ** Property on Hand added separately  (intRetVal = CRPT_ON_HAND_END).
                    ' ** Information for Investments Made.
7480                If (strJournalType = "Purchase") Then
7490                  intRetVal = CRPT_INVEST_INFO
7500                Else
                      ' ** Changes in Investment Holdings.
7510                  If (strJournalType = "Sold" And curGainLoss = 0) _
                          Or (strJournalType = "Deposit" And strComment Like "*stock split*") _
                          Or (strJournalType = "Purchase" And curICash <= 0) _
                          Or (strJournalType = "Liability") Then
7520                    intRetVal = CRPT_CHANGES
7530                  End If
7540                End If
7550              End If
7560            End If
7570          End If
7580        End If
7590      End If
7600    End If

EXITP:
7610    CACourtReportDataID = intRetVal
7620    Exit Function

ERRH:
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Function

Public Function CACourtReportDate(intCourtReport As Integer, datTransDate As Date, datAssetDate As Date) As Date
'modCourtReportsCA.CABuildCourtReportData()

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportDate"

        Dim lngX As Long, blnFound As Boolean
        Dim datRetVal As Date

7710    blnFound = False
7720    For lngX = 0& To (lngCARpts - 1&)
7730      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
7740        Select Case arr_varCARpt(CR_DATE, lngX)
            Case "TransDate"
7750          datRetVal = datTransDate
7760        Case "AssetDate"
7770          datRetVal = datAssetDate
7780        End Select
7790        blnFound = True
7800        Exit For
7810      End If
7820    Next
7830    If blnFound = False Then datRetVal = #1/1/1900#

EXITP:
7840    CACourtReportDate = datRetVal
7850    Exit Function

ERRH:
7860    datRetVal = #1/1/1900#
7870    Select Case ERR.Number
        Case Else
7880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7890    End Select
7900    Resume EXITP

End Function

Public Function CACourtReportDivision(intCourtReport As Integer) As Integer
'modCourtReportsCA.CAGetCourtReportData()
'modCourtReportsCA.CABuildCourtReportData()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportDivision"

        Dim lngX As Long, blnFound As Boolean
        Dim intRetVal As Integer

8010    blnFound = False
8020    For lngX = 0& To (lngCARpts - 1&)
8030      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
8040        intRetVal = arr_varCARpt(CR_DIV, lngX)
8050        blnFound = True
8060        Exit For
8070      End If
8080    Next
8090    If blnFound = False Then intRetVal = 0

EXITP:
8100    CACourtReportDivision = intRetVal
8110    Exit Function

ERRH:
8120    intRetVal = 0
8130    Select Case ERR.Number
        Case Else
8140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8150    End Select
8160    Resume EXITP

End Function

Public Function CACourtReportDivisionText(intCourtReport As Integer) As String
'modCourtReportsCA.CAGetCourtReportData()
'modCourtReportsCA.CABuildCourtReportData()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportDivisionText"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

8210    blnFound = False
8220    For lngX = 0& To (lngCARpts - 1&)
8230      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
8240        strRetVal = arr_varCARpt(CR_DIVTXT, lngX)
8250        blnFound = True
8260        Exit For
8270      End If
8280    Next
8290    If blnFound = False Then strRetVal = "Unknown"

EXITP:
8300    CACourtReportDivisionText = strRetVal
8310    Exit Function

ERRH:
8320    strRetVal = RET_ERR
8330    Select Case ERR.Number
        Case Else
8340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8350    End Select
8360    Resume EXITP

End Function

Public Function CACourtReportDivisionTitle(intCourtReport As Integer) As String
'modCourtReportsCA.CAGetCourtReportData()
'modCourtReportsCA.CABuildCourtReportData()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportDivisionTitle"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

8410    blnFound = False
8420    For lngX = 0& To (lngCARpts - 1&)
8430      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
8440        strRetVal = arr_varCARpt(CR_DIVTTL, lngX)
8450        blnFound = True
8460        Exit For
8470      End If
8480    Next
8490    If blnFound = False Then strRetVal = "Unknown"

EXITP:
8500    CACourtReportDivisionTitle = strRetVal
8510    Exit Function

ERRH:
8520    strRetVal = RET_ERR
8530    Select Case ERR.Number
        Case Else
8540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8550    End Select
8560    Resume EXITP

End Function

Public Function CACourtReportGroup(intCourtReport As Integer) As Integer
'modCourtReportsCA.CAGetCourtReportData()
'modCourtReportsCA.CABuildCourtReportData()

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportGroup"

        Dim lngX As Long, blnFound As Boolean
        Dim intRetVal As Integer

8610    blnFound = False
8620    For lngX = 0& To (lngCARpts - 1&)
8630      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
8640        intRetVal = arr_varCARpt(CR_GRP, lngX)
8650        blnFound = True
8660        Exit For
8670      End If
8680    Next
8690    If blnFound = False Then intRetVal = 0

EXITP:
8700    CACourtReportGroup = intRetVal
8710    Exit Function

ERRH:
8720    intRetVal = 0
8730    Select Case ERR.Number
        Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Function

Public Function CACourtReportGroupText(intCourtReport As Integer) As String
'modCourtReportsCA.CAGetCourtReportData()
'modCourtReportsCA.CABuildCourtReportData()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportGroupText"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

8810    blnFound = False
8820    For lngX = 0& To (lngCARpts - 1&)
8830      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
8840        strRetVal = arr_varCARpt(CR_GRPTXT, lngX)
8850        blnFound = True
8860        Exit For
8870      End If
8880    Next
8890    If blnFound = False Then strRetVal = "Unknown"

EXITP:
8900    CACourtReportGroupText = strRetVal
8910    Exit Function

ERRH:
8920    strRetVal = RET_ERR
8930    Select Case ERR.Number
        Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Function

Public Function CACourtReportSchedule(intCourtReport As Integer) As String
'modCourtReportsCA.CAGetCourtReportData()
'modCourtReportsCA.CABuildCourtReportData()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportSchedule"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

9010    blnFound = False
9020    For lngX = 0& To (lngCARpts - 1&)
9030      If arr_varCARpt(CR_NUM, lngX) = intCourtReport Then
9040        strRetVal = arr_varCARpt(CR_SCHED, lngX)
9050        blnFound = True
9060        Exit For
9070      End If
9080    Next
9090    If blnFound = False Then strRetVal = vbNullString

EXITP:
9100    CACourtReportSchedule = strRetVal
9110    Exit Function

ERRH:
9120    strRetVal = RET_ERR
9130    Select Case ERR.Number
        Case Else
9140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9150    End Select
9160    Resume EXITP

End Function

Private Function CACourtReportCaseNum() As String
'modCourtReportsCA.CABuildCourtReportData()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportCaseNum"

        Dim strRetVal As String

9210    strRetVal = vbNullString
9220  On Error Resume Next

9230    strRetVal = DLookup("[CaseNum]", "account", "[accountno] = '" & gstrAccountNo & "'")
9240  On Error GoTo ERRH

EXITP:
9250    CACourtReportCaseNum = strRetVal
9260    Exit Function

ERRH:
9270    strRetVal = RET_ERR
9280    Select Case ERR.Number
        Case Else
9290      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9300    End Select
9310    Resume EXITP

End Function

Public Function CACourtReportDollars(intCourtReport As Integer, curPCash As Currency, curICash As Currency, curCost As Currency, strJournalType As String) As Double
'modCourtReportsCA.CABuildCourtReportData()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "CACourtReportDollars"

        Dim dblRetVal As Double

9410    Select Case intCourtReport
        Case CRPT_ADDL_PROP      ' ** WAS: 10   NOW: 10
9420      dblRetVal = curCost + curPCash
9430    Case CRPT_RECEIPTS       ' ** WAS: 20   NOW: 20
9440      If (strJournalType = "Sold" And curICash > 0) Or strJournalType = "Purchase" And curICash > 0 And curPCash + curCost = 0 Then
9450        dblRetVal = curICash
9460      Else
9470        dblRetVal = curICash + curPCash
9480      End If
9490    Case CRPT_GAINS          ' ** WAS: 30   NOW: 30
9500      dblRetVal = curPCash + curCost
9510    Case CRPT_OTH_CHG        ' ** WAS:      NOW: 40
          ' ** Unknown.  'VGC 05/20/2011: JUST A GUESS!
9520      If (strJournalType = "Sold" And curICash > 0) Or strJournalType = "Purchase" And curICash > 0 And curPCash + curCost = 0 Then
9530        dblRetVal = curICash
9540      Else
9550        dblRetVal = curICash + curPCash
9560      End If
9570    Case CRPT_NET_INCOME     ' ** WAS: 40   NOW: 50
          ' ** Nothing.
9580    Case CRPT_DISBURSEMENTS  ' ** WAS: 50   NOW: 60
9590      If strJournalType = "Withdrawn" Then
9600        dblRetVal = curCost * -1
9610      Else
9620        dblRetVal = (curICash + curPCash) * -1
9630      End If
9640    Case CRPT_LOSSES         ' ** WAS: 60   NOW: 70
9650      dblRetVal = (curPCash + curCost) * -1
9660    Case CRPT_DISTRIBUTIONS  ' ** WAS: 80   NOW: 80
9670      If strJournalType = "Withdrawn" Then
9680        dblRetVal = curCost * -1
9690      Else
9700        dblRetVal = (curICash + curPCash) * -1
9710      End If
9720    Case CRPT_OTH_CRED       ' ** WAS:      NOW: 90
          ' ** Unknown.  'VGC 05/20/2011: JUST A GUESS!
9730      If strJournalType = "Withdrawn" Then
9740        dblRetVal = curCost * -1
9750      Else
9760        dblRetVal = (curICash + curPCash) * -1
9770      End If
9780    Case CRPT_NET_LOSS       ' ** WAS: 70   NOW: 100
          ' ** Nothing.
9790    Case CRPT_ON_HAND_END    ' ** WAS: 90   NOW: 110
          ' ** Nothing.
9800    Case CRPT_INVEST_INFO    ' ** WAS: 100  NOW: 120
9810      dblRetVal = (curICash + curPCash) * -1
9820    Case CRPT_CHANGES        ' ** WAS: 110  NOW: 130
9830      dblRetVal = curCost
9840    Case Else
9850      dblRetVal = -999999
9860    End Select

EXITP:
9870    CACourtReportDollars = dblRetVal
9880    Exit Function

ERRH:
9890    dblRetVal = 0#
9900    Select Case ERR.Number
        Case Else
9910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9920    End Select
9930    Resume EXITP

End Function

Public Function CAGetCourtReportData(Optional varType As Variant) As Integer
' ** Called by:
' **   CABuildCourtReportData(), above
' ** Return Codes:
' **   0  Success.
' **  -1  Canceled.
' **  -9  Error.

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "CAGetCourtReportData"

        'Dim rsxDataOut As ADODB.Recordset  ' ** Early binding.
        Dim rsxDataOut As Object            ' ** Late binding.
        Dim dblNetIncome As Double, dblNetLoss As Double
        Dim dblCashAssets_Beg As Double, dblCashAssets_End As Double
        Dim blnShowForm As Boolean, strSource As String, strProc As String
        Dim intRetVal As Integer

10010   intRetVal = 0

        ' ** When called from the TAReports CommandBar, these variables should already be filled.
        ' **   gstrCrtRpt_Ordinal
        ' **   gstrCrtRpt_Version
        ' **   gstrCrtRpt_CashAssets_Beg
        ' **   gstrCrtRpt_NetIncome
        ' **   gstrCrtRpt_NetLoss
        ' **   gstrCrtRpt_CashAssets_End

        ' ** Get ordinal and version info.
10020   If gblnCrtRpt_Zero = False Then
          ' ** If report is open in preview, don't call this form.

10030     If IsMissing(varType) = True Then
10040       blnShowForm = True
10050     Else
            ' ** 'CABuildCourtReportData' ^ strProc
            ' ** 'cmdPrintAll_Click' ^ 'cmdPrintAll'
10060       strSource = Left(varType, (InStr(varType, "^") - 1))
10070       strProc = Mid(varType, (InStr(varType, "^") + 1))
10080       If strSource = "cmdPrintAll_Click" Or strSource = "cmdWordAll_Click" Or strSource = "cmdExcelAll_Click" Then
10090         blnShowForm = True
10100       Else
10110         Select Case gblnPrintAll
              Case True
10120           blnShowForm = False
10130         Case False
10140           blnShowForm = True
10150         End Select
10160       End If
10170     End If

10180     If blnShowForm = True Then
10190       DoCmd.Hourglass False
            ' ** Leave these as they are.
            ' **   gstrCrtRpt_Ordinal
            ' **   gstrCrtRpt_Version
10200       gblnMessage = True  ' ** If this returns False, the dialog was canceled.
10210       DoCmd.OpenForm "frmRpt_CourtReports_CA_Input", , , , , acDialog, "frmRpt_CourtReports_CA"

10220       If gblnMessage = False Then
10230         intRetVal = -1  ' ** Canceled.
10240       Else
10250         Forms("frmRpt_CourtReports_CA").Ordinal = gstrCrtRpt_Ordinal
10260         Forms("frmRpt_CourtReports_CA").Version = gstrCrtRpt_Version
10270       End If
            'DoCmd.OpenForm "frmRpt_CourtReports_CA_Input", , , , , acDialog
            'If IsNull(Forms("frmRpt_CourtReports_CA").Ordinal) = True Or IsNull(Forms("frmRpt_CourtReports_CA").Version) = True Then
            '  intRetVal = -1  ' ** Canceled.
            'Else
            '  If Forms("frmRpt_CourtReports_CA").Ordinal = vbNullString Or Forms("frmRpt_CourtReports_CA").Version = vbNullString Then
            '    intRetVal = -1  ' ** Canceled.
            '  End If
            'End If
10280     End If

          'If intRetVal <> -1 Then
          '  If blnShowForm = True Then
          'gstrCrtRpt_Ordinal = Forms("frmRpt_CourtReports_CA").Ordinal
          'gstrCrtRpt_Version = Forms("frmRpt_CourtReports_CA").Version
          'gstrCrtRpt_CashAssets_Beg = InputBox("Please enter the Cash Assets at Beginning of Period.", _
          '  "Beginning Cash Assets", IIf(gstrCrtRpt_CashAssets_Beg = vbNullString, 0, Val(gstrCrtRpt_CashAssets_Beg)))
          'If gstrCrtRpt_CashAssets_Beg = vbNullString Then
          '  intRetVal = -1
          'ElseIf IsNumeric(gstrCrtRpt_CashAssets_Beg) = False And InStr(gstrCrtRpt_CashAssets_Beg, "$") = 0 And _
          '    InStr(gstrCrtRpt_CashAssets_Beg, "(") = 0 And InStr(gstrCrtRpt_CashAssets_Beg, ",") = 0 Then
          '  MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
          '  intRetVal = -1
          'Else
          '  ' ** Remove '$', '()', ','.
          '  gstrCrtRpt_CashAssets_Beg = CleanInputBox(Trim(gstrCrtRpt_CashAssets_Beg))  ' ** Module Function: modReportFunctions.
          '  gstrCrtRpt_NetIncome = InputBox("Please enter the Net Income From a Trade or Business amount.", _
          '    "Net Income From a Trade or Business", IIf(gstrCrtRpt_NetIncome = vbNullString, 0, Val(gstrCrtRpt_NetIncome)))
          '  If gstrCrtRpt_NetIncome = vbNullString Then
          '    intRetVal = -1
          '  ElseIf IsNumeric(gstrCrtRpt_NetIncome) = False And InStr(gstrCrtRpt_NetIncome, "$") = 0 And _
          '      InStr(gstrCrtRpt_NetIncome, "(") = 0 And InStr(gstrCrtRpt_NetIncome, ",") = 0 Then
          '    MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
          '    intRetVal = -1
          '  Else
          '    ' ** Remove '$', '()', ','.
          '    gstrCrtRpt_NetIncome = CleanInputBox(Trim(gstrCrtRpt_NetIncome))  ' ** Module Function: modReportFunctions.
          '    gstrCrtRpt_NetLoss = InputBox("Please enter the Net Loss From a Trade or Business amount.", _
          '      "Net Loss From a Trade or Business", IIf(gstrCrtRpt_NetLoss = vbNullString, 0, Val(gstrCrtRpt_NetLoss)))
          '    If gstrCrtRpt_NetLoss = vbNullString Then
          '      intRetVal = -1
          '    ElseIf IsNumeric(gstrCrtRpt_NetLoss) = False And InStr(gstrCrtRpt_NetLoss, "$") = 0 And _
          '        InStr(gstrCrtRpt_NetLoss, "(") = 0 And InStr(gstrCrtRpt_NetLoss, ",") = 0 Then
          '      MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
          '      intRetVal = -1
          '    Else
          '      ' ** Remove '$', '()', ','.
          '      gstrCrtRpt_NetLoss = CleanInputBox(Trim(gstrCrtRpt_NetLoss))  ' ** Module Function: modReportFunctions.
          '      gstrCrtRpt_CashAssets_End = InputBox("Please enter the Cash Assets at End of Period.", _
          '        "Ending Cash Assets", IIf(gstrCrtRpt_CashAssets_End = vbNullString, 0, Val(gstrCrtRpt_CashAssets_End)))
          '      If gstrCrtRpt_CashAssets_End = vbNullString Then
          '        intRetVal = -1
          '      ElseIf IsNumeric(gstrCrtRpt_CashAssets_End) = False And InStr(gstrCrtRpt_CashAssets_End, "$") = 0 And _
          '          InStr(gstrCrtRpt_CashAssets_End, "(") = 0 And InStr(gstrCrtRpt_CashAssets_End, ",") = 0 Then
          '        MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
          '        intRetVal = -1
          '      Else
          '        ' ** Remove '$', '()', ','.
          '        gstrCrtRpt_CashAssets_End = CleanInputBox(Trim(gstrCrtRpt_CashAssets_End))  ' ** Module Function: modReportFunctions.
          '      End If
          '    End If
          '  End If
          'End If
          '  Else
          'If gstrCrtRpt_Ordinal = vbNullString Or gstrCrtRpt_Version = vbNullString Or _
          '    gstrCrtRpt_CashAssets_Beg = vbNullString Or IsNull(Forms("frmRpt_CourtReports_CA").CashAssets_Beg) = True Then
          '  Beep
          '  MsgBox "Something's missing!"
          '  Stop
          'End If
          '  End If
          'End If

10290   End If

10300   If intRetVal = 0 Then

10310     DoCmd.Hourglass True

10320     If Trim(gstrCrtRpt_CashAssets_Beg) = vbNullString Then gstrCrtRpt_CashAssets_Beg = "0"
10330     If IsNumeric(Trim(gstrCrtRpt_CashAssets_Beg)) = False Then gstrCrtRpt_CashAssets_Beg = "0"
10340     If Trim(gstrCrtRpt_CashAssets_End) = vbNullString Then gstrCrtRpt_CashAssets_End = "0"
10350     If IsNumeric(Trim(gstrCrtRpt_CashAssets_End)) = False Then gstrCrtRpt_CashAssets_End = "0"

          ' ** Get Beginning amounts.
10360     dblCashAssets_Beg = CDbl(gstrCrtRpt_CashAssets_Beg)
10370     gdblCrtRpt_CA_COHBeg = dblCashAssets_Beg
10380     Forms("frmRpt_CourtReports_CA").CashAssets_Beg = dblCashAssets_Beg
10390     dblCashAssets_End = CDbl(gstrCrtRpt_CashAssets_End)
10400     Forms("frmRpt_CourtReports_CA").CashAssets_End = dblCashAssets_End
10410     gdblCrtRpt_CA_COHEnd = dblCashAssets_End
10420     dblNetIncome = CDbl(gstrCrtRpt_NetIncome)
10430     dblNetLoss = CDbl(gstrCrtRpt_NetLoss)

          'Set rsxDataOut = New ADODB.Recordset             ' ** Early binding.
10440     Set rsxDataOut = CreateObject("ADODB.Recordset")  ' ** Late binding.
10450     rsxDataOut.Open "tmpCourtReportData", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTable

10460     With rsxDataOut
            ' ** Get the beginning Cash Assets amount for the report.
10470       .AddNew
10480       .Fields("ReportNumber") = CRPT_CASH_BEG
10490       .Fields("ReportCategory") = CACourtReportCategory(CRPT_CASH_BEG)
10500       .Fields("ReportGroup") = CACourtReportGroup(CRPT_CASH_BEG)
10510       .Fields("ReportDivision") = CACourtReportDivision(CRPT_CASH_BEG)
10520       .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_CASH_BEG)
10530       .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_CASH_BEG)
10540       .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_CASH_BEG)
10550       .Fields("accountno") = gstrAccountNo
10560       .Fields("date") = gdatStartDate
10570       .Fields("journaltype") = "Entered"
10580       .Fields("shareface") = 0
10590       .Fields("Description") = "Beginning Cash Assets"
10600       .Fields("Amount") = dblCashAssets_Beg
10610       .Fields("revcode_ID") = 0
10620       .Fields("revcode_DESC") = "Dummy entry"
10630       .Fields("revcode_TYPE") = 1
10640       .Fields("revcode_SORTORDER") = 0
10650       .Fields("ReportSchedule") = CACourtReportSchedule(CRPT_CASH_BEG)
10660       .Fields("CaseNum") = strCaseNum
10670       .Update
            ' ** Get the Net Income Amount for the report.
10680       .AddNew
10690       .Fields("ReportNumber") = CRPT_NET_INCOME
10700       .Fields("ReportCategory") = CACourtReportCategory(CRPT_NET_INCOME)
10710       .Fields("ReportGroup") = CACourtReportGroup(CRPT_NET_INCOME)
10720       .Fields("ReportDivision") = CACourtReportDivision(CRPT_NET_INCOME)
10730       .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_NET_INCOME)
10740       .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_NET_INCOME)
10750       .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_NET_INCOME)
10760       .Fields("accountno") = gstrAccountNo
10770       .Fields("date") = gdatStartDate
10780       .Fields("journaltype") = "Entered"
10790       .Fields("shareface") = 0
10800       .Fields("Description") = "Net Income"
10810       .Fields("Amount") = dblNetIncome
10820       .Fields("revcode_ID") = 0
10830       .Fields("revcode_DESC") = "Dummy entry"
10840       .Fields("revcode_TYPE") = 1
10850       .Fields("revcode_SORTORDER") = 0
10860       .Fields("ReportSchedule") = CACourtReportSchedule(CRPT_NET_INCOME)
10870       .Fields("CaseNum") = strCaseNum
10880       .Update
            ' ** Get the Net Loss Amount for the report.
10890       .AddNew
10900       .Fields("ReportNumber") = CRPT_NET_LOSS
10910       .Fields("ReportCategory") = CACourtReportCategory(CRPT_NET_LOSS)
10920       .Fields("ReportGroup") = CACourtReportGroup(CRPT_NET_LOSS)
10930       .Fields("ReportDivision") = CACourtReportDivision(CRPT_NET_LOSS)
10940       .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_NET_LOSS)
10950       .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_NET_LOSS)
10960       .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_NET_LOSS)
10970       .Fields("accountno") = gstrAccountNo
10980       .Fields("date") = gdatStartDate
10990       .Fields("journaltype") = "Entered"
11000       .Fields("shareface") = 0
11010       .Fields("Description") = "Net Loss"
11020       .Fields("Amount") = dblNetLoss
11030       .Fields("revcode_ID") = 0
11040       .Fields("revcode_DESC") = "Dummy entry"
11050       .Fields("revcode_TYPE") = 1
11060       .Fields("revcode_SORTORDER") = 0
11070       .Fields("ReportSchedule") = CACourtReportSchedule(CRPT_NET_LOSS)
11080       .Fields("CaseNum") = strCaseNum
11090       .Update
            ' ** Get the ending Cash Assets amount for the report.
11100       .AddNew
11110       .Fields("ReportNumber") = CRPT_CASH_END
11120       .Fields("ReportCategory") = CACourtReportCategory(CRPT_CASH_END)
11130       .Fields("ReportGroup") = CACourtReportGroup(CRPT_CASH_END)
11140       .Fields("ReportDivision") = CACourtReportDivision(CRPT_CASH_END)
11150       .Fields("ReportDivisionTitle") = CACourtReportDivisionTitle(CRPT_CASH_END)
11160       .Fields("ReportDivisionText") = CACourtReportDivisionText(CRPT_CASH_END)
11170       .Fields("ReportGroupText") = CACourtReportGroupText(CRPT_CASH_END)
11180       .Fields("accountno") = gstrAccountNo
11190       .Fields("date") = gdatStartDate
11200       .Fields("journaltype") = "Entered"
11210       .Fields("shareface") = 0
11220       .Fields("Description") = "Ending Cash Assets"
11230       .Fields("Amount") = dblCashAssets_End
11240       .Fields("revcode_ID") = 0
11250       .Fields("revcode_DESC") = "Dummy entry"
11260       .Fields("revcode_TYPE") = 1
11270       .Fields("revcode_SORTORDER") = 0
11280       .Fields("ReportSchedule") = CACourtReportSchedule(CRPT_CASH_END)
11290       .Fields("CaseNum") = strCaseNum
11300       .Update
11310     End With

11320   End If

EXITP:
11330   Set rsxDataOut = Nothing
11340   CAGetCourtReportData = intRetVal
11350   Exit Function

ERRH:
11360   intRetVal = -9  ' ** Error.
11370   Select Case ERR.Number
        Case 13 ' ** Type mismatch.
11380     MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
11390   Case Else
11400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11410   End Select
11420   Resume EXITP

End Function

Public Sub WordAll_CA(frm As Access.Form)

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "WordAll_CA"

        Dim strOrd As String, strVer As String
        Dim strCashBeg As String, dblCashBeg As Double, strCashEnd As String, dblCashEnd As Double
        Dim strDocName As String
        Dim blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal As Integer
        Dim strTmp01 As String
        Dim lngX As Long

        ' ** Access: 14272506  Very Light Red
        ' ** Access: 12295153  Medium Red
        ' ** Word:   16770233  Very Light Blue
        ' ** Word:   16434048  Medium Blue
        ' ** Excel:  14677736  Very Light Green
        ' ** Excel:  5952646   Medium Green

11510   With frm
11520     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_CA.

11530       DoCmd.Hourglass True
11540       DoEvents

11550       .cmdWordAll_box01.Visible = True
11560       .cmdWordAll_box02.Visible = True
11570       If .chkAssetList = True Then
11580         .cmdWordAll_box03.Visible = True
11590       End If
11600       .cmdWordAll_box04.Visible = True
11610       DoEvents

11620       blnExcel = False
11630       blnAutoStart = .chkOpenWord
11640       Beep
11650       DoCmd.Hourglass False
11660       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Word" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

11670       If msgResponse = vbOK Then

11680         DoCmd.Hourglass True
11690         DoEvents

11700         blnAllCancel = False
11710         .AllCancelSet1_CA blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
11720         gblnPrintAll = True
11730         blnAutoStart = False  ' ** They'll open only after all have been exported.
11740         strThisProc = "cmdWordAll_Click"

              ' ** Get the Summary inputs first.
11750         intRetVal = CAGetCourtReportData(strThisProc & "^" & "cmdWordAll")  ' ** Function: Above.
11760         If intRetVal <> 0 Then
11770           blnAllCancel = True
11780           .AllCancelSet1_CA blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
11790         Else
                ' ** Save these for later.
11800           strOrd = gstrCrtRpt_Ordinal
11810           strVer = gstrCrtRpt_Version
11820           If Trim$(gstrCrtRpt_CashAssets_Beg) = vbNullString Then gstrCrtRpt_CashAssets_Beg = "0"
11830           If IsNumeric(Trim$(gstrCrtRpt_CashAssets_Beg)) = False Then gstrCrtRpt_CashAssets_Beg = "0"
11840           If Trim$(gstrCrtRpt_CashAssets_End) = vbNullString Then gstrCrtRpt_CashAssets_End = "0"
11850           If IsNumeric(Trim$(gstrCrtRpt_CashAssets_End)) = False Then gstrCrtRpt_CashAssets_End = "0"
11860           strCashBeg = gstrCrtRpt_CashAssets_Beg
11870           dblCashBeg = Nz(.CashAssets_Beg, 0)
11880           strCashEnd = gstrCrtRpt_CashAssets_End
11890           dblCashEnd = Nz(.CashAssets_End, 0)
11900         End If

11910         If blnAllCancel = False Then
                ' ** Summary of Account.
11920           .cmdWord00.SetFocus
11930           .cmdWord00_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
11940           DoEvents
11950         End If
11960         If blnAllCancel = False Then
                ' ** Additional Property Received.
11970           .cmdWord01.SetFocus
11980           .cmdWord01_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
11990           DoEvents
12000         End If
12010         If blnAllCancel = False Then
                ' ** Receipts.
12020           .cmdWord02.SetFocus
12030           .cmdWord02_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12040           DoEvents
12050         End If
12060         If blnAllCancel = False Then
                ' ** Gains on Sales.
12070           .cmdWord03.SetFocus
12080           .cmdWord03_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12090           DoEvents
12100         End If
12110         If blnAllCancel = False Then
                ' ** Other Charges.
12120           .cmdWord10.SetFocus
12130           .cmdWord10_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12140           DoEvents
12150         End If
12160         If blnAllCancel = False Then
                ' ** Disbursements.
12170           .cmdWord04.SetFocus
12180           .cmdWord04_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12190           DoEvents
12200         End If
12210         If blnAllCancel = False Then
                ' ** Losses on Sales.
12220           .cmdWord05.SetFocus
12230           .cmdWord05_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12240           DoEvents
12250         End If
12260         If blnAllCancel = False Then
                ' ** Distributions.
12270           .cmdWord06.SetFocus
12280           .cmdWord06_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12290           DoEvents
12300         End If
12310         If blnAllCancel = False Then
                ' ** Other Credits.
12320           .cmdWord11.SetFocus
12330           .cmdWord11_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12340           DoEvents
12350         End If
12360         If blnAllCancel = False Then
                ' ** Property on Hand at Close of Accounting Period.
12370           .cmdWord07.SetFocus
12380           .cmdWord07_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12390           DoEvents
12400         End If
12410         If blnAllCancel = False Then
                ' ** Information for Investments Made.
12420           .cmdWord08.SetFocus
12430           .cmdWord08_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12440           DoEvents
12450         End If
12460         If blnAllCancel = False Then
                ' ** Change in Investment Holdings.
12470           .cmdWord09.SetFocus
12480           .cmdWord09_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
12490           DoEvents
12500         End If

12510         DoCmd.Hourglass True
12520         DoEvents

12530         .cmdWordAll.SetFocus

12540         gblnPrintAll = False
12550         Beep

12560         If lngFiles > 0& Then

12570           DoCmd.Hourglass False

12580           strTmp01 = CStr(lngFiles) & " documents were created."
12590           If .chkOpenWord = True Then
12600             strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
12610             msgResponse = MsgBox(strTmp01, vbInformation + vbOKCancel, "Reports Exported")
12620           Else
12630             msgResponse = MsgBox(strTmp01, vbInformation + vbOKOnly, "Reports Exported")
12640           End If

12650           .cmdWordAll_box01.Visible = False
12660           .cmdWordAll_box02.Visible = False
12670           .cmdWordAll_box03.Visible = False
12680           .cmdWordAll_box04.Visible = False

12690           If .chkOpenWord = True And msgResponse = vbOK Then
12700             DoCmd.Hourglass True
12710             DoEvents
12720             For lngX = 0& To (lngFiles - 1&)
12730               strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
12740               OpenExe strDocName  ' ** Module Function: modShellFuncs.
12750               DoEvents
12760               If lngX < (lngFiles - 1&) Then
12770                 ForcePause 2  ' ** Module Function: modCodeUtilities.
12780               End If
12790             Next
12800             Beep
12810           End If

12820         Else
12830           DoCmd.Hourglass False
12840           MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
12850           .cmdWordAll_box01.Visible = False
12860           .cmdWordAll_box02.Visible = False
12870           .cmdWordAll_box03.Visible = False
12880           .cmdWordAll_box04.Visible = False
12890         End If  ' ** lngFiles.

12900       Else
12910         .cmdWordAll_box01.Visible = False
12920         .cmdWordAll_box02.Visible = False
12930         .cmdWordAll_box03.Visible = False
12940         .cmdWordAll_box04.Visible = False
12950       End If  ' ** msgResponse.

12960       DoCmd.Hourglass False
12970     End If  ' ** Validate.
12980   End With

EXITP:
12990   Exit Sub

ERRH:
13000   gblnPrintAll = False
13010   DoCmd.Hourglass False
13020   Select Case ERR.Number
        Case Else
13030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13040   End Select
13050   Resume EXITP

End Sub

Public Sub ExcelAll_CA(frm As Access.Form)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "ExcelAll_CA"

        Dim strOrd As String, strVer As String
        Dim strCashBeg As String, dblCashBeg As Double, strCashEnd As String, dblCashEnd As Double
        Dim strDocName As String
        Dim blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal As Integer
        Dim strTmp01 As String
        Dim lngX As Long

        ' ** Access: 14272506  Very Light Red
        ' ** Access: 12295153  Medium Red
        ' ** Word:   16770233  Very Light Blue
        ' ** Word:   16434048  Medium Blue
        ' ** Excel:  14677736  Very Light Green
        ' ** Excel:  5952646   Medium Green

13110   With frm
13120     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_CA.

13130       DoCmd.Hourglass True
13140       DoEvents

13150       .cmdExcelAll_box01.Visible = True
13160       .cmdExcelAll_box02.Visible = True
13170       If .chkAssetList = True Then
13180         .cmdExcelAll_box03.Visible = True
13190       End If
13200       .cmdExcelAll_box04.Visible = True
13210       DoEvents

13220       blnExcel = True
13230       blnAutoStart = .chkOpenExcel
13240       Beep
13250       DoCmd.Hourglass False
13260       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Excel" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

13270       If msgResponse = vbOK Then

13280         DoCmd.Hourglass True
13290         DoEvents

13300         blnAllCancel = False
13310         .AllCancelSet1_CA blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
13320         gblnPrintAll = True
13330         blnAutoStart = False  ' ** They'll open only after all have been exported.
13340         strThisProc = "cmdExcelAll_Click"

13350         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13360           DoCmd.Hourglass False
13370           msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                  "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                  "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                  "You may close Excel before proceding, then click Retry." & vbCrLf & _
                  "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
                ' ** ... Otherwise Trust Accountant will do it for you.
13380           If msgResponse <> vbRetry Then
13390             blnAllCancel = True
13400             .AllCancelSet1_CA blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
13410           End If
13420         End If

13430         If blnAllCancel = False Then

13440           DoCmd.Hourglass True
13450           DoEvents

13460           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13470             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
13480           End If
13490           DoEvents

                ' ** Get the Summary inputs first.
13500           intRetVal = CAGetCourtReportData(strThisProc & "^" & "cmdExcelAll")  ' ** Function: Above.
13510           If intRetVal <> 0 Then
13520             blnAllCancel = True
13530             .AllCancelSet1_CA blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
13540           Else
                  ' ** Save these for later.
13550             strOrd = gstrCrtRpt_Ordinal
13560             strVer = gstrCrtRpt_Version
13570             If Trim$(gstrCrtRpt_CashAssets_Beg) = vbNullString Then gstrCrtRpt_CashAssets_Beg = "0"
13580             If IsNumeric(Trim$(gstrCrtRpt_CashAssets_Beg)) = False Then gstrCrtRpt_CashAssets_Beg = "0"
13590             If Trim$(gstrCrtRpt_CashAssets_End) = vbNullString Then gstrCrtRpt_CashAssets_End = "0"
13600             If IsNumeric(Trim$(gstrCrtRpt_CashAssets_End)) = False Then gstrCrtRpt_CashAssets_End = "0"
13610             strCashBeg = gstrCrtRpt_CashAssets_Beg
13620             dblCashBeg = Nz(.CashAssets_Beg, 0)
13630             strCashEnd = gstrCrtRpt_CashAssets_End
13640             dblCashEnd = Nz(.CashAssets_End, 0)
13650           End If

13660           If blnAllCancel = False Then
                  ' ** Summary of Account.
13670             .cmdExcel00.SetFocus
13680             .cmdExcel00_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
13690             DoEvents
13700           End If
13710           If blnAllCancel = False Then
                  ' ** Additional Property Received.
13720             .cmdExcel01.SetFocus
13730             .cmdExcel01_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
13740             DoEvents
13750           End If
13760           If blnAllCancel = False Then
                  ' ** Receipts.
13770             .cmdExcel02.SetFocus
13780             .cmdExcel02_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
13790             DoEvents
13800           End If
13810           If blnAllCancel = False Then
                  ' ** Gains on Sales.
13820             .cmdExcel03.SetFocus
13830             .cmdExcel03_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
13840             DoEvents
13850           End If
13860           If blnAllCancel = False Then
                  ' ** Other Charges.
13870             .cmdExcel10.SetFocus
13880             .cmdExcel10_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
13890             DoEvents
13900           End If
13910           If blnAllCancel = False Then
                  ' ** Disbursements.
13920             .cmdExcel04.SetFocus
13930             .cmdExcel04_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
13940             DoEvents
13950           End If
13960           If blnAllCancel = False Then
                  ' ** Losses on Sales.
13970             .cmdExcel05.SetFocus
13980             .cmdExcel05_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
13990             DoEvents
14000           End If
14010           If blnAllCancel = False Then
                  ' ** Distributions.
14020             .cmdExcel06.SetFocus
14030             .cmdExcel06_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
14040             DoEvents
14050           End If
14060           If blnAllCancel = False Then
                  ' ** Other Credits.
14070             .cmdExcel11.SetFocus
14080             .cmdExcel11_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
14090             DoEvents
14100           End If
14110           If blnAllCancel = False Then
                  ' ** Property on Hand at Close of Accounting Period.
14120             .cmdExcel07.SetFocus
14130             .cmdExcel07_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
14140             DoEvents
14150           End If
14160           If blnAllCancel = False Then
                  ' ** Information for Investments Made.
14170             .cmdExcel08.SetFocus
14180             .cmdExcel08_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
14190             DoEvents
14200           End If
14210           If blnAllCancel = False Then
                  ' ** Change in Investment Holdings.
14220             .cmdExcel09.SetFocus
14230             .cmdExcel09_Click  ' ** Form Procedure: frmRpt_CourtReports_CA.
14240             DoEvents
14250           End If

14260           DoCmd.Hourglass True
14270           DoEvents

14280           .cmdExcelAll.SetFocus

14290           gblnPrintAll = False
14300           Beep

14310           If lngFiles > 0& Then

14320             DoCmd.Hourglass False

14330             strTmp01 = CStr(lngFiles) & " documents were created."
14340             If .chkOpenExcel = True Then
14350               strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
14360             End If

14370             MsgBox strTmp01, vbInformation + vbOKOnly, "Reports Exported"

14380             .cmdExcelAll_box01.Visible = False
14390             .cmdExcelAll_box02.Visible = False
14400             .cmdExcelAll_box03.Visible = False
14410             .cmdExcelAll_box04.Visible = False

14420             If .chkOpenExcel = True Then
14430               DoCmd.Hourglass True
14440               DoEvents
14450               For lngX = 0& To (lngFiles - 1&)
14460                 strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
14470                 OpenExe strDocName  ' ** Module Function: modShellFuncs.
14480                 DoEvents
14490                 If lngX < (lngFiles - 1&) Then
14500                   ForcePause 2  ' ** Module Function: modCodeUtilities.
14510                 End If
14520               Next
14530             End If

14540           Else
14550             DoCmd.Hourglass False
14560             MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
14570             .cmdExcelAll_box01.Visible = False
14580             .cmdExcelAll_box02.Visible = False
14590             .cmdExcelAll_box03.Visible = False
14600             .cmdExcelAll_box04.Visible = False
14610           End If

14620         End If  ' ** blnAllCancel.

14630       Else
14640         .cmdExcelAll_box01.Visible = False
14650         .cmdExcelAll_box02.Visible = False
14660         .cmdExcelAll_box03.Visible = False
14670         .cmdExcelAll_box04.Visible = False
14680       End If  ' ** msgResponse.

14690       DoCmd.Hourglass False
14700     End If  ' ** Validate.
14710   End With

EXITP:
14720   Exit Sub

ERRH:
14730   gblnPrintAll = False
14740   DoCmd.Hourglass False
14750   Select Case ERR.Number
        Case Else
14760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14770   End Select
14780   Resume EXITP

End Sub

Public Sub FileArraySet_CA(arr_varTmp00 As Variant)

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "FileArraySet_CA"

14810   arr_varFile = arr_varTmp00
14820   lngFiles = UBound(arr_varFile, 2) + 1&

EXITP:
14830   Exit Sub

ERRH:
14840   Select Case ERR.Number
        Case Else
14850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14860   End Select
14870   Resume EXITP

End Sub

Public Sub AllCancelSet2_CA(blnCancel As Boolean)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet2_CA"

14910   blnAllCancel = blnCancel

EXITP:
14920   Exit Sub

ERRH:
14930   Select Case ERR.Number
        Case Else
14940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14950   End Select
14960   Resume EXITP

End Sub

Public Sub SetArchiveOption_CA(blnIsOpen As Boolean, frm As Access.Form)

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "SetArchiveOption_CA"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccountNo As String
        Dim blnArchive As Boolean

15010   With frm
15020     Select Case blnIsOpen
          Case True
15030       blnArchive = False
15040       Set dbs = CurrentDb
15050       With dbs
              ' ** LedgerArchive, grouped, with cnt.
15060         Set qdf = .QueryDefs("qryStatementParameters_30")
15070         Set rst = qdf.OpenRecordset
15080         With rst
15090           If .BOF = True And .EOF = True Then
                  ' ** No Archive.
15100           Else
15110             .MoveFirst
15120             If ![cnt] = 0 Then
                    ' ** No archive.
15130             Else
15140               blnArchive = True
15150             End If
15160           End If
15170           .Close
15180         End With
15190         Set rst = Nothing
15200         Set qdf = Nothing
15210         .Close
15220       End With
15230       DoEvents
15240       Set dbs = Nothing
15250       Select Case blnArchive
            Case True
15260         .HasArchive = True
15270         .chkIncludeArchive_lbl.Visible = True
15280         .chkIncludeArchive_lbl2.Visible = False
15290         .chkIncludeArchive_lbl2_dim_hi.Visible = False
15300         .chkIncludeArchive.Enabled = True
15310       Case False
15320         .HasArchive = False
15330         .chkIncludeArchive_lbl.Visible = False
15340         .chkIncludeArchive_lbl2.Caption = "There are no Archived Transactions"
15350         .chkIncludeArchive_lbl2.FontBold = False
15360         .chkIncludeArchive_lbl2.Visible = True
15370         .chkIncludeArchive_lbl2_dim_hi.Caption = "There are no Archived Transactions"
15380         .chkIncludeArchive_lbl2_dim_hi.FontBold = False
15390         .chkIncludeArchive_lbl2_dim_hi.Visible = True
15400         .chkIncludeArchive.Enabled = False
15410       End Select
15420     Case False
15430       If .HasArchive = True Then
15440         If IsNull(.cmbAccounts) = False Then
15450           DoCmd.Hourglass True
15460           DoEvents
15470           strAccountNo = .cmbAccounts
15480           blnArchive = False
15490           DoEvents
15500           Set dbs = CurrentDb
15510           With dbs
                  ' ** LedgerArchive, by specified [actno].
15520             Set qdf = .QueryDefs("qryStatementParameters_31")
15530             With qdf.Parameters
15540               ![actno] = strAccountNo
15550             End With
15560             Set rst = qdf.OpenRecordset
15570             With rst
15580               If .BOF = True And .EOF = True Then
                      ' ** No archive for this account.
15590               Else
15600                 .MoveFirst
15610                 If ![cnt] = 0 Then
                        ' ** No archive for this account.
15620                 Else
15630                   blnArchive = True
15640                 End If
15650               End If
15660               .Close
15670             End With
15680             Set rst = Nothing
15690             Set qdf = Nothing
15700             .Close
15710           End With
15720           Set dbs = Nothing
15730           DoEvents
15740           Select Case blnArchive
                Case True
15750             .chkIncludeArchive_lbl.Visible = True
15760             .chkIncludeArchive_lbl2.Visible = False
15770             .chkIncludeArchive_lbl2_dim_hi.Visible = False
15780             .chkIncludeArchive.Enabled = True
15790           Case False
15800             .chkIncludeArchive_lbl.Visible = False
15810             Select Case .chkIncludeArchive
                  Case True
15820               .chkIncludeArchive_lbl2.FontBold = True
15830               .chkIncludeArchive_lbl2_dim_hi.FontBold = True
15840             Case False
15850               .chkIncludeArchive_lbl2.FontBold = False
15860               .chkIncludeArchive_lbl2_dim_hi.FontBold = False
15870             End Select
15880             .chkIncludeArchive_lbl2.Caption = "No Archived Transactions For This Account"
15890             .chkIncludeArchive_lbl2.Visible = True
15900             .chkIncludeArchive_lbl2_dim_hi.Caption = "No Archived Transactions For This Account"
15910             .chkIncludeArchive_lbl2_dim_hi.Visible = True
15920             .chkIncludeArchive.Enabled = False
15930           End Select
15940           DoCmd.Hourglass False
15950         End If
15960       End If
15970     End Select
15980   End With

EXITP:
15990   Set rst = Nothing
16000   Set qdf = Nothing
16010   Set dbs = Nothing
16020   Exit Sub

ERRH:
16030   DoCmd.Hourglass False
16040   Select Case ERR.Number
        Case Else
16050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16060   End Select
16070   Resume EXITP

End Sub

Public Sub SetUserReportPath_CA(frm As Access.Form)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "SetUserReportPath_CA"

        Dim blnEnable As Boolean

16110   With frm
16120     blnEnable = True
16130     Select Case IsNull(.UserReportPath)
          Case True
16140       blnEnable = False
16150     Case False
16160       If Trim(.UserReportPath) = vbNullString Then
16170         blnEnable = False
16180       End If
16190     End Select
16200     Select Case blnEnable
          Case True
16210       .UserReportPath.BorderColor = CLR_LTBLU2
16220       .UserReportPath.BackStyle = acBackStyleNormal
16230       .UserReportPath.Enabled = True  ' ** It remains locked.
16240       .UserReportPath_chk.Enabled = True
16250       .UserReportPath_chk.Locked = False
16260       .UserReportPath_chk_lbl1.Visible = True
16270       .UserReportPath_chk_lbl1_dim.Visible = False
16280       .UserReportPath_chk_lbl1_dim_hi.Visible = False
16290       .UserReportPath_chk_lbl2.Visible = True
16300       .UserReportPath_chk_lbl2_dim.Visible = False
16310       .UserReportPath_chk_lbl2_dim_hi.Visible = False
16320     Case False
16330       .UserReportPath = vbNullString
16340       .UserReportPath.BorderColor = WIN_CLR_DISR
16350       .UserReportPath.BackStyle = acBackStyleTransparent
16360       .UserReportPath.Enabled = False
16370       .UserReportPath_chk.Enabled = False
16380       .UserReportPath_chk.Locked = False
16390       .UserReportPath_chk_lbl1.Visible = False
16400       .UserReportPath_chk_lbl1_dim.Visible = True
16410       .UserReportPath_chk_lbl1_dim_hi.Visible = True
16420       .UserReportPath_chk_lbl2.Visible = False
16430       .UserReportPath_chk_lbl2_dim.Visible = True
16440       .UserReportPath_chk_lbl2_dim_hi.Visible = True
16450     End Select
16460   End With

EXITP:
16470   Exit Sub

ERRH:
16480   DoCmd.Hourglass False
16490   Select Case ERR.Number
        Case Else
16500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16510   End Select
16520   Resume EXITP

End Sub

Public Sub Preview00_CA(frm As Access.Form)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "Preview00_CA"

16610   With frm

16620     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_CA.

16630       DoCmd.Hourglass True
16640       DoEvents

16650       TAReports_SetZero True  ' ** Module Function: modReportFunctions.

16660       gdblCrtRpt_CA_POHBeg = 0#
16670       gdblCrtRpt_CA_POHEnd = 0#
16680       gdblCrtRpt_CA_COHBeg = 0#
16690       gdblCrtRpt_CA_COHEnd = 0#
16700       gdblCrtRpt_CA_InvestInfo = 0#
16710       gdblCrtRpt_CA_InvestChange = 0#
16720       .PropOnHand_Beg = Null
16730       .CashAssets_Beg = Null
16740       .PropOnHand_End = Null
16750       .CashAssets_End = Null
16760       .Ordinal = Null
16770       .Version = Null

            'INDIVIDUAL MEANS SEPARATED GAINS AND LOSSES, NOT INDIVIDUAL REPORTS!
            'ERGO, COMBINED IS COMBINED GAINS AND LOSSES, NOT THE COMBINED SUMMARY REPORT!
            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
16780       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
16790         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
16800           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
16810             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
16820             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
16830             CurrentDb.QueryDefs.Refresh
16840             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
16850             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
16860             CurrentDb.QueryDefs.Refresh
16870             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
16880             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
16890             CurrentDb.QueryDefs.Refresh
16900             CurrentDb.QueryDefs.Refresh
16910           End If
16920         End If
16930       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR THE 2 INDIVIDUAL REPORTS, rptCourtRptCA_03 AND rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
16940       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
16950         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Or _
                  InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
16960           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
16970             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
16980             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
16990             CurrentDb.QueryDefs.Refresh
17000             CurrentDb.QueryDefs.Refresh
17010           End If
17020         End If
17030       End If

17040       If gblnUseReveuneExpenseCodes = True Then
17050         .PreviewOrPrint "0A", THIS_PROC, acViewPreview  ' ** Form Procedure: frmRpt_CourtReports_CA.
17060       Else
17070         .PreviewOrPrint "0", THIS_PROC, acViewPreview  ' ** Form Procedure: frmRpt_CourtReports_CA.
17080       End If

17090     End If  ' ** Validate.

17100   End With

EXITP:
17110   Exit Sub

ERRH:
17120   DoCmd.Hourglass False
17130   Select Case ERR.Number
        Case Else
17140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17150   End Select
17160   Resume EXITP

End Sub

Public Sub Print00_CA(frm As Access.Form)
' ** Public, because it's called from the TAReports CommandBar as well.

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "Print00_CA"

        ' ** When called from the TAReports CommandBar, these variables should already be filled.
        ' **   gstrCrtRpt_Ordinal
        ' **   gstrCrtRpt_Version
        ' **   gstrCrtRpt_CashAssets_Beg
        ' **   gstrCrtRpt_NetIncome
        ' **   gstrCrtRpt_NetLoss
        ' **   gstrCrtRpt_CashAssets_End

17210   With frm

17220     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_CA.

17230       DoCmd.Hourglass True
17240       DoEvents

17250       If gblnCrtRpt_Zero = False Then
17260         TAReports_SetZero False  ' ** Module Function: modReportFunctions.
17270       End If

17280       gdblCrtRpt_CA_POHBeg = 0#
17290       gdblCrtRpt_CA_POHEnd = 0#
17300       gdblCrtRpt_CA_COHBeg = 0#
17310       gdblCrtRpt_CA_COHEnd = 0#
17320       gdblCrtRpt_CA_InvestInfo = 0#
17330       gdblCrtRpt_CA_InvestChange = 0#
17340       .PropOnHand_Beg = Null
17350       .CashAssets_Beg = Null
17360       .PropOnHand_End = Null
17370       .CashAssets_End = Null
17380       .Ordinal = Null
17390       .Version = Null

            'INDIVIDUAL MEANS SEPARATED GAINS AND LOSSES, NOT INDIVIDUAL REPORTS!
            'ERGO, COMBINED IS COMBINED GAINS AND LOSSES, NOT THE COMBINED SUMMARY REPORT!
            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
17400       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
17410         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
17420           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
17430             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
17440             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
17450             CurrentDb.QueryDefs.Refresh
17460             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
17470             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
17480             CurrentDb.QueryDefs.Refresh
17490             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
17500             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
17510             CurrentDb.QueryDefs.Refresh
17520             CurrentDb.QueryDefs.Refresh
17530           End If
17540         End If
17550       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR THE 2 INDIVIDUAL REPORTS, rptCourtRptCA_03 AND rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
17560       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
17570         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Or _
                  InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
17580           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
17590             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
17600             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
17610             CurrentDb.QueryDefs.Refresh
17620             CurrentDb.QueryDefs.Refresh
17630           End If
17640         End If
17650       End If

17660       If gblnUseReveuneExpenseCodes = True Then
              '##GTR_Ref: rptCourtRptCA_00A
17670         .PreviewOrPrint "0A", THIS_PROC, acViewNormal  ' ** Form Procedure: frmRpt_CourtReports_CA.
17680       Else
              '##GTR_Ref: rptCourtRptCA_00
17690         .PreviewOrPrint "0", THIS_PROC, acViewNormal  ' ** Form Procedure: frmRpt_CourtReports_CA.
17700       End If

17710     End If  ' ** Validate

17720   End With

EXITP:
17730   Exit Sub

ERRH:
17740   DoCmd.Hourglass False
17750   Select Case ERR.Number
        Case Else
17760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17770   End Select
17780   Resume EXITP

End Sub

Public Sub Word00_CA(blnRebuildTable As Boolean, frm As Access.Form)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "Word00_CA"

17810   With frm

17820     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_CA.

17830       DoCmd.Hourglass True
17840       DoEvents

17850       gdblCrtRpt_CA_POHBeg = 0#
17860       gdblCrtRpt_CA_POHEnd = 0#
17870       gdblCrtRpt_CA_COHBeg = 0#
17880       gdblCrtRpt_CA_COHEnd = 0#
17890       gdblCrtRpt_CA_InvestInfo = 0#
17900       gdblCrtRpt_CA_InvestChange = 0#
17910       .PropOnHand_Beg = Null
17920       .CashAssets_Beg = Null
17930       .PropOnHand_End = Null
17940       .CashAssets_End = Null
17950       .Ordinal = Null
17960       .Version = Null

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
17970       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
17980         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
17990           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
18000             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
18010             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
18020             CurrentDb.QueryDefs.Refresh
18030             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
18040             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
18050             CurrentDb.QueryDefs.Refresh
18060             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
18070             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
18080             CurrentDb.QueryDefs.Refresh
18090             CurrentDb.QueryDefs.Refresh
18100           End If
18110         End If
18120       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR THE 2 INDIVIDUAL REPORTS, rptCourtRptCA_03 AND rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
18130       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
18140         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Or _
                  InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
18150           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
18160             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
18170             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
18180             CurrentDb.QueryDefs.Refresh
18190             CurrentDb.QueryDefs.Refresh
18200           End If
18210         End If
18220       End If

18230       If gblnUseReveuneExpenseCodes = True Then
18240         SendToFile_CA "0A", THIS_PROC, blnRebuildTable, frm, False  ' ** Procedure: Below.
18250       Else
18260         SendToFile_CA "0", THIS_PROC, blnRebuildTable, frm, False  ' ** Procedure: Below.
18270       End If

18280     End If  ' ** Validate

18290   End With

EXITP:
18300   Exit Sub

ERRH:
18310   DoCmd.Hourglass False
18320   Select Case ERR.Number
        Case Else
18330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18340   End Select
18350   Resume EXITP

End Sub

Public Sub Excel00_CA(blnRebuildTable As Boolean, frm As Access.Form)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "Excel00_CA"

18410   With frm

18420     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_CA.

18430       DoCmd.Hourglass True
18440       DoEvents

18450       gdblCrtRpt_CA_POHBeg = 0#
18460       gdblCrtRpt_CA_POHEnd = 0#
18470       gdblCrtRpt_CA_COHBeg = 0#
18480       gdblCrtRpt_CA_COHEnd = 0#
18490       gdblCrtRpt_CA_InvestInfo = 0#
18500       gdblCrtRpt_CA_InvestChange = 0#
18510       .PropOnHand_Beg = Null
18520       .CashAssets_Beg = Null
18530       .PropOnHand_End = Null
18540       .CashAssets_End = Null
18550       .Ordinal = Null
18560       .Version = Null

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
18570       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
18580         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
18590           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
18600             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
18610             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
18620             CurrentDb.QueryDefs.Refresh
18630             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
18640             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
18650             CurrentDb.QueryDefs.Refresh
18660             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
18670             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
18680             CurrentDb.QueryDefs.Refresh
18690             CurrentDb.QueryDefs.Refresh
18700           End If
18710         End If
18720       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR THE 2 INDIVIDUAL REPORTS, rptCourtRptCA_03 AND rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
18730       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
18740         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Or _
                  InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
18750           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
18760             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
18770             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
18780             CurrentDb.QueryDefs.Refresh
18790             CurrentDb.QueryDefs.Refresh
18800           End If
18810         End If
18820       End If

            ' ###############################################################################################################
            ' ## qryCourtReport_CA_00a IS THE BASIS FOR THE QUERY THAT GETS EXPORTED TO EXCEL, qryCourtReport_CA_00i!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00a_bak1  ' ** tmpCourtReportData, grouped and summed.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00a_bak2  ' ** .._CA_00a_08.                              ####  Combined Gains and Losses.    ####
18830       If QueryExists("qryCourtReport_CA_00a") = True Then  ' ** Module Function: modFileUtilities.
18840         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00a").SQL, "_00a_08") > 0 Then
18850           If QueryExists("qryCourtReport_CA_00a_bak1") = True Then  ' ** Module Function: modFileUtilities.
18860             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00a"
18870             DoCmd.CopyObject , "qryCourtReport_CA_00a", acQuery, "qryCourtReport_CA_00a_bak1"
18880             CurrentDb.QueryDefs.Refresh
18890             CurrentDb.QueryDefs.Refresh
18900           End If
18910         End If
18920       End If

18930       If gblnUseReveuneExpenseCodes = True Then
18940         .SendToFile_CA "0A", THIS_PROC, blnRebuildTable, frm, True  ' ** Procedure: Below.
18950       Else
18960         .SendToFile_CA "0", THIS_PROC, blnRebuildTable, frm, True  ' ** Procedure: Below.
18970       End If

18980     End If  ' ** Validate

18990   End With

EXITP:
19000   Exit Sub

ERRH:
19010   DoCmd.Hourglass False
19020   Select Case ERR.Number
        Case Else
19030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19040   End Select
19050   Resume EXITP

End Sub

Public Sub SendToFile_CA(strReportNumber As String, strProc As String, blnRebuildTable As Boolean, frm As Access.Form, Optional varExcel As Variant)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "SendToFile_CA"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strQry As String, strMacro As String
        Dim strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim blnExcel As Boolean, blnNoTmp2 As Boolean, blnNoData As Boolean
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean
        Dim intRetVal_BuildCourtReportData As Integer, intRetVal_BuildAssetListInfo As Integer

        Static intSendIter As Integer

19110   blnContinue = True
19120   blnUseSavedPath = False

19130   With frm

19140     DoCmd.Hourglass True
19150     DoEvents

19160     If IsMissing(varExcel) = True Then
19170       blnExcel = False
19180     Else
19190       blnExcel = varExcel
19200     End If

19210     If blnExcel = True Then
19220       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
19230         DoCmd.Hourglass False
19240         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
19250         If msgResponse <> vbRetry Then
19260           blnContinue = False
19270           blnAllCancel = True
19280           AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
19290         End If
19300       End If
19310     End If

19320     If blnContinue = True Then

19330       DoCmd.Hourglass True
19340       DoEvents

19350       intRetVal_BuildCourtReportData = 0: intRetVal_BuildAssetListInfo = 0
19360       blnNoTmp2 = False
19370       intSendIter = intSendIter + 1

19380       ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

            ' ** Set global variables for report headers.
19390       gdatStartDate = .DateStart.Value
19400       strTmp01 = CStr(CDbl(gdatStartDate))
19410       If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
19420       glngStartDateLong = CLng(strTmp01)
19430       gdatEndDate = .DateEnd.Value
19440       strTmp01 = CStr(CDbl(gdatEndDate))
19450       If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
19460       glngEndDateLong = CLng(strTmp01)
19470       gstrAccountNo = .cmbAccounts.Column(CBX_A_ACTNO)
19480       gstrAccountName = .cmbAccounts.Column(CBX_A_SHORT)

19490       Set dbs = CurrentDb
19500       With dbs
              ' ** tblReport, captions of Court Reports, by specified [CrtTyp].
19510         Set qdf = .QueryDefs("qryCourtReport_15")
19520         With qdf.Parameters
19530           ![CrtTyp] = "CA"
19540         End With
19550         Set rst = qdf.OpenRecordset
19560         With rst
19570           .MoveLast
19580           lngCaps = .RecordCount
19590           .MoveFirst
19600           arr_varCap = .GetRows(lngCaps)
                ' ****************************************************
                ' ** Array: arr_varCap()
                ' **
                ' **   Field  Element  Name               Constant
                ' **   =====  =======  =================  ==========
                ' **     1       0     rpt_id             C_RID
                ' **     2       1     rpt_name           C_RNAM
                ' **     3       2     rpt_caption        C_CAP
                ' **     4       3     rpt_caption_new    C_CAPN
                ' **
                ' ****************************************************
19610           .Close
19620         End With
19630         .Close
19640       End With

19650       If IsNull(.UserReportPath) = False Then
19660         If .UserReportPath <> vbNullString Then
19670           If .UserReportPath_chk = True Then
19680             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
19690               blnUseSavedPath = True
19700             End If
19710           End If
19720         End If
19730       End If

19740       If strReportNumber <> "7" Then
              ' ** Build a new summary report table for everything except
              ' ** report 7, the property on hand report.
              ' **** ARCHIVE ****
19750         intRetVal_BuildCourtReportData = CABuildCourtReportData(strReportNumber, strProc, .chkIncludeArchive, .chkIncludeCheckNum)  ' ** Module Function: modCourtReportsCA.
              ' *****************

19760         If intRetVal_BuildCourtReportData = 0 Then

19770           blnRebuildTable = False
19780           strTmp02 = "Ending"
19790           intRetVal_BuildAssetListInfo = .BuildAssetListInfo(.DateStart, .DateEnd, strTmp02, strProc)  ' ** Form Function: frmRpt_CourtReports_CA.

19800           If intRetVal_BuildAssetListInfo = 0 Then
                  ' ** Property On Hand = Sum([TotalCost])+IIf(IsNull([icash]),0,[icash])+[pcash].

                  ' ** Property on Hand at End of Account Period.
19810             Set dbs = CurrentDb
19820             Set qdf = dbs.QueryDefs("qryCourtReport_CA_07g")  ' ** tmpAssetList2, summed.
19830             Set rst = qdf.OpenRecordset
19840             If rst.BOF = True And rst.EOF = True Then
                    ' ** In the absense of newer data to roll back, take the total cost out of tmpCourtReportData.
19850               rst.Close
                    ' ** tmpCourtReportData, summed with PropertyOnHand.
19860               Set qdf = dbs.QueryDefs("qryCourtReport_CA_07e")
19870               With qdf.Parameters
19880                 ![actno] = frm.cmbAccounts
19890               End With
19900               Set rst = qdf.OpenRecordset
19910               rst.MoveFirst
19920               gdblCrtRpt_CA_POHEnd = Nz(rst![PropertyOnHand], 0)
19930               .PropOnHand_End = Nz(rst![PropertyOnHand], 0)
19940               rst.Close
19950             Else
19960               rst.MoveFirst
19970               gdblCrtRpt_CA_POHEnd = Nz(rst![PropertyOnHand], 0)
19980               .PropOnHand_End = Nz(rst![PropertyOnHand], 0)
19990               rst.Close
20000             End If

                  ' ** For Property On Hand at Beginning of Account Period, we need 1 day prior to DateStart.
20010             strTmp02 = "Beginning"
20020             intRetVal_BuildAssetListInfo = .BuildAssetListInfo("01/01/1900", (.DateStart - 1), strTmp02, strProc)  ' ** Form Function: frmRpt_CourtReports_CA.
20030             If intRetVal_BuildAssetListInfo = 0 Then
20040               Set qdf = dbs.QueryDefs("qryCourtReport_CA_07g")  ' ** tmpAssetList2, summed.
20050               Set rst = qdf.OpenRecordset
20060               If rst.BOF = True And rst.EOF = True Then
                      ' ** In the absense of newer data to roll back, take the total cost out of tmpCourtReportData.
20070                 rst.Close
20080                 blnNoTmp2 = True
                      ' ** tmpCourtReportData, summed with PropertyOnHand.
20090                 Set qdf = dbs.QueryDefs("qryCourtReport_CA_07e")
20100                 With qdf.Parameters
20110                   ![actno] = frm.cmbAccounts
20120                 End With
20130                 Set rst = qdf.OpenRecordset
20140                 rst.MoveFirst
20150                 gdblCrtRpt_CA_POHBeg = Nz(rst![PropertyOnHand], 0)
20160                 .PropOnHand_Beg = Nz(rst![PropertyOnHand], 0)
20170                 rst.Close
20180               Else
20190                 rst.MoveFirst
20200                 gdblCrtRpt_CA_POHBeg = Nz(rst![PropertyOnHand], 0)
20210                 .PropOnHand_Beg = Nz(rst![PropertyOnHand], 0)
20220                 rst.Close
20230               End If

20240               If blnContinue = True Then

                      ' ** Information for Investments Made total, report page 2.
20250                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
20260                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A2_02_archive")
                        ' *****************
20270                 Case False
20280                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A2_02")
20290                 End Select
20300                 Set rst = qdf.OpenRecordset
20310                 If rst.BOF = True And rst.EOF = True Then
                        ' ** Nothing to total.
20320                   gdblCrtRpt_CA_InvestInfo = 0@
20330                   .InvestInfo = 0&
20340                 Else
20350                   rst.MoveFirst
20360                   gdblCrtRpt_CA_InvestInfo = Nz(rst![Amount_Tot], 0)
20370                   .InvestInfo = Nz(rst![Amount_Tot], 0)
20380                 End If
20390                 rst.Close

                      ' ** Change in Investment Holdings total, report page 2.
20400                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
20410                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A3_02_archive")
                        ' *****************
20420                 Case False
20430                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A3_02")
20440                 End Select
20450                 Set rst = qdf.OpenRecordset
20460                 If rst.BOF = True And rst.EOF = True Then
                        ' ** Nothing to total.
20470                   gdblCrtRpt_CA_InvestChange = 0@
20480                   .InvestChange = 0&
20490                 Else
20500                   rst.MoveFirst
20510                   gdblCrtRpt_CA_InvestChange = Nz(rst![cost_tot], 0)
20520                   .InvestChange = Nz(rst![cost_tot], 0)
20530                 End If
20540                 rst.Close

20550               End If

20560             Else
                    ' ** No beginning data: strTmp02 = "Beginning.
20570               blnContinue = False
                    ' ** Return codes:
                    ' **    0  Success.
                    ' **   -3  Missing entry, e.g., date.
                    ' **   -2  No data.
                    ' **   -9  Error.
20580             End If
20590             dbs.Close
20600           Else
                  ' ** No ending data: strTmp02 = "Ending".
20610             blnContinue = False
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -3  Missing entry, e.g., date.
                  ' **   -2  No data.
                  ' **   -9  Error.
20620           End If

20630           If strTmp02 = "Beginning" And blnContinue = False And intRetVal_BuildAssetListInfo = -2 Then
                  ' ** There was an ending balance, but no beginning balance.
                  ' ** Since that is a reasonable combination, continue.
20640             blnContinue = True
20650             intRetVal_BuildAssetListInfo = 0
20660           End If

20670           If blnExcel = True Then
20680             If strReportNumber <> "0B" Then
20690               If blnNoTmp2 = True Then
                      ' ** No data to roll back.
20700               End If
20710               strQry = vbNullString: blnNoData = False
20720               Select Case strReportNumber
                    Case "0"  ' ** Summary of Account.
20730                 strQry = "qryCourtReport_CA_00i"
20740               Case "0A"  ' ** Summary of Account - Grouped.
20750                 strQry = "qryCourtReport_CA_00_Al"
20760               Case "1"  ' ** Additional Property Received.
20770                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
20780                   strQry = "qryCourtReport_CA_01h_archive"
                        ' *****************
20790                 Case False
20800                   strQry = "qryCourtReport_CA_01h"
20810                 End Select
20820                 varTmp00 = DCount("*", strQry)
20830                 If IsNull(varTmp00) = True Then
20840                   blnNoData = True
20850                   strQry = "qryCourtReport_CA_01k"  ' ** For Export - No Data.
20860                 Else
20870                   If varTmp00 = 0 Then
20880                     blnNoData = True
20890                     strQry = "qryCourtReport_CA_01k"  ' ** For Export - No Data.
20900                   End If
20910                 End If
20920               Case "2"  ' ** Receipts.
20930                 strQry = "qryCourtReport_CA_02l"  ' ** Uses tmpCourtReportData.
20940                 varTmp00 = DCount("*", strQry)
20950                 If IsNull(varTmp00) = True Then
20960                   blnNoData = True
20970                   strQry = "qryCourtReport_CA_02o"  ' ** For Export - No Data.
20980                 Else
20990                   If varTmp00 = 0 Then
21000                     blnNoData = True
21010                     strQry = "qryCourtReport_CA_02o"  ' ** For Export - No Data.
21020                   End If
21030                 End If
21040               Case "2A"  ' ** Receipts - Grouped.
21050                 strQry = "qryCourtReport_CA_02_Aq"  ' ** Uses tmpCourtReportData.
21060                 varTmp00 = DCount("*", strQry)
21070                 If IsNull(varTmp00) = True Then
21080                   blnNoData = True
21090                   strQry = "qryCourtReport_CA_02_At"  ' ** For Export - No Data.
21100                 Else
21110                   If varTmp00 = 0 Then
21120                     blnNoData = True
21130                     strQry = "qryCourtReport_CA_02_At"  ' ** For Export - No Data.
21140                   End If
21150                 End If
21160               Case "3"  ' ** Gains on Sale or Other Dispositions.
21170                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
21180                   strQry = "qryCourtReport_CA_03i_archive"
                        ' *****************
21190                 Case False
21200                   strQry = "qryCourtReport_CA_03i"
21210                 End Select
21220                 varTmp00 = DCount("*", strQry)
21230                 If IsNull(varTmp00) = True Then
21240                   blnNoData = True
21250                   strQry = "qryCourtReport_CA_03l"  ' ** For Export - No Data.
21260                 Else
21270                   If varTmp00 = 0 Then
21280                     blnNoData = True
21290                     strQry = "qryCourtReport_CA_03l"  ' ** For Export - No Data.
21300                   End If
21310                 End If
21320               Case "4"  ' ** Disbursements.
21330                 strQry = "qryCourtReport_CA_04h"  ' ** Uses tmpCourtReportData.
21340                 varTmp00 = DCount("*", strQry)
21350                 If IsNull(varTmp00) = True Then
21360                   blnNoData = True
21370                   strQry = "qryCourtReport_CA_04l"  ' ** For Export - No Data.
21380                 Else
21390                   If varTmp00 = 0 Then
21400                     blnNoData = True
21410                     strQry = "qryCourtReport_CA_04l"  ' ** For Export - No Data.
21420                   End If
21430                 End If
21440               Case "4A"  ' ** Disbursements - Grouped.
21450                 strQry = "qryCourtReport_CA_04_Ai"  ' ** Uses tmpCourtReportData.
21460                 varTmp00 = DCount("*", strQry)
21470                 If IsNull(varTmp00) = True Then
21480                   blnNoData = True
21490                   strQry = "qryCourtReport_CA_04_Al"  ' ** For Export - No Data.
21500                 Else
21510                   If varTmp00 = 0 Then
21520                     blnNoData = True
21530                     strQry = "qryCourtReport_CA_04_Al"  ' ** For Export - No Data.
21540                   End If
21550                 End If
21560               Case "5"  ' ** Losses on Sale or Other Dispositions.
21570                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
21580                   strQry = "qryCourtReport_CA_05i_archive"
                        ' *****************
21590                 Case False
21600                   strQry = "qryCourtReport_CA_05i"
21610                 End Select
21620                 varTmp00 = DCount("*", strQry)
21630                 If IsNull(varTmp00) = True Then
21640                   blnNoData = True
21650                   strQry = "qryCourtReport_CA_05l"  ' ** For Export - No Data.
21660                 Else
21670                   If varTmp00 = 0 Then
21680                     blnNoData = True
21690                     strQry = "qryCourtReport_CA_05l"  ' ** For Export - No Data.
21700                   End If
21710                 End If
21720               Case "6"  ' ** Distributions.
21730                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
21740                   strQry = "qryCourtReport_CA_06i_archive"
                        ' *****************
21750                 Case False
21760                   strQry = "qryCourtReport_CA_06i"
21770                 End Select
21780                 varTmp00 = DCount("*", strQry)
21790                 If IsNull(varTmp00) = True Then
21800                   blnNoData = True
21810                   strQry = "qryCourtReport_CA_06l"  ' ** For Export - No Data.
21820                 Else
21830                   If varTmp00 = 0 Then
21840                     blnNoData = True
21850                     strQry = "qryCourtReport_CA_06l"  ' ** For Export - No Data.
21860                   End If
21870                 End If
21880               Case "8"  ' ** Information For Investments Made.
21890                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
21900                   strQry = "qryCourtReport_CA_08i_archive"
                        ' *****************
21910                 Case False
21920                   strQry = "qryCourtReport_CA_08i"
21930                 End Select
21940                 varTmp00 = DCount("*", strQry)
21950                 If IsNull(varTmp00) = True Then
21960                   blnNoData = True
21970                   strQry = "qryCourtReport_CA_08l"  ' ** For Export - No Data.
21980                 Else
21990                   If varTmp00 = 0 Then
22000                     blnNoData = True
22010                     strQry = "qryCourtReport_CA_08l"  ' ** For Export - No Data.
22020                   End If
22030                 End If
22040               Case "9"  ' ** Change in Investment Holdings.
22050                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
22060                   strQry = "qryCourtReport_CA_09i_archive"
                        ' *****************
22070                 Case False
22080                   strQry = "qryCourtReport_CA_09i"
22090                 End Select
22100                 varTmp00 = DCount("*", strQry)
22110                 If IsNull(varTmp00) = True Then
22120                   blnNoData = True
22130                   strQry = "qryCourtReport_CA_09l"  ' ** For Export - No Data.
22140                 Else
22150                   If varTmp00 = 0 Then
22160                     blnNoData = True
22170                     strQry = "qryCourtReport_CA_09l"  ' ** For Export - No Data.
22180                   End If
22190                 End If
22200               Case "10"  ' ** Other Charges.
22210                 strQry = "qryCourtReport_CA_10h"  ' ** Uses tmpCourtReportData.
22220                 varTmp00 = DCount("*", strQry)
22230                 If IsNull(varTmp00) = True Then
22240                   blnNoData = True
22250                   strQry = "qryCourtReport_CA_10k"  ' ** For Export - No Data.
22260                 Else
22270                   If varTmp00 <= 2 Then  ' ** Title, Period.
22280                     blnNoData = True
22290                     strQry = "qryCourtReport_CA_10k"  ' ** For Export - No Data.
22300                   End If
22310                 End If
22320               Case "11"  ' ** Other Credits.
22330                 strQry = "qryCourtReport_CA_11h"  ' ** Uses tmpCourtReportData.
22340                 varTmp00 = DCount("*", strQry)
22350                 If IsNull(varTmp00) = True Then
22360                   blnNoData = True
22370                   strQry = "qryCourtReport_CA_11k"  ' ** For Export - No Data.
22380                 Else
22390                   If varTmp00 <= 2 Then  ' ** Title, Period.
22400                     blnNoData = True
22410                     strQry = "qryCourtReport_CA_11k"  ' ** For Export - No Data.
22420                   End If
22430                 End If
22440               End Select

22450               strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
22460               strRptPath = .UserReportPath

22470               If Len(strReportNumber) = 1 Then
22480                 strRptName = ("rptCourtRptCA_0" & strReportNumber)
22490               ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
22500                 strRptName = ("rptCourtRptCA_" & strReportNumber)
22510               Else
22520                 strRptName = ("rptCourtRptCA_0" & strReportNumber)
22530               End If

22540               strMacro = "mcrExcelExport_CR_CA" & Mid(strRptName, InStr(strRptName, "_"))
22550               If blnNoData = True Then
22560                 strMacro = strMacro & "_nd"
22570               End If

22580               For lngX = 0& To (lngCaps - 1&)
22590                 If arr_varCap(C_RNAM, lngX) = strRptName Then
22600                   strRptCap = arr_varCap(C_CAPN, lngX)
22610                   Exit For
22620                 End If
22630               Next

22640               Select Case blnUseSavedPath
                    Case True
22650                 strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
22660               Case False
22670                 DoCmd.Hourglass False
22680                 strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
22690               End Select

22700               If strRptPathFile <> vbNullString Then
22710                 DoCmd.Hourglass True
22720                 DoEvents
22730                 Select Case blnExcel
                      Case True
22740                   blnAutoStart = .chkOpenExcel
22750                 Case False
22760                   blnAutoStart = .chkOpenWord
22770                 End Select
22780                 If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
22790                 If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
22800                   Kill strRptPathFile
22810                 End If
22820                 If strQry <> vbNullString Then
22830                   CACourtReportLoad  ' ** Module Function: modCourReportsCA.
                        ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                        ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
22840                   DoCmd.RunMacro strMacro
                        ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                        ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
22850                   If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_CA_xxx.xls") = True Or _
                            FileExists(strRptPath & LNK_SEP & "CourtReport_CA_xxx.xls") = True Then
22860                     If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_CA_xxx.xls") = True Then
22870                       Name (CurrentAppPath & LNK_SEP & "CourtReport_CA_xxx.xls") As (strRptPathFile)
                            ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
22880                     Else
22890                       Name (strRptPath & LNK_SEP & "CourtReport_CA_xxx.xls") As (strRptPathFile)
                            ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
22900                     End If
22910                     DoEvents
22920                     If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
22930                       DoEvents
22940                       Select Case gblnPrintAll
                            Case True
22950                         lngFiles = lngFiles + 1&
22960                         lngE = lngFiles - 1&
22970                         ReDim Preserve arr_varFile(F_ELEMS, lngE)
22980                         arr_varFile(F_RNAM, lngE) = strRptName
22990                         arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
23000                         arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23010                         FileArraySet_CA arr_varFile  ' ** Module Procedure: modCourtReportsCA.
23020                       Case False
23030                         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
23040                           EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
23050                         End If
23060                         DoEvents
23070                         If blnAutoStart = True Then
23080                           OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
23090                         End If
23100                       End Select
                            'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                            '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                            'End If
                            'DoEvents
                            'If blnAutoStart = True Then
                            '  OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
                            'End If
23110                     End If
23120                   End If
23130                 Else
23140                   DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
23150                 End If
23160                 strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23170                 If strRptPath <> .UserReportPath Then
23180                   .UserReportPath = strRptPath
23190                   SetUserReportPath_CA frm  ' ** Module Procedure: modCourtReportsCA.
23200                 End If
23210               Else
23220                 blnContinue = False
23230               End If

23240             End If  ' ** 0B.

23250           Else
                  ' ** Not Excel.

23260             If strReportNumber <> "0B" Then

23270               strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
23280               strRptPath = .UserReportPath

23290               If Len(strReportNumber) = 1 Then
23300                 strRptName = ("rptCourtRptCA_0" & strReportNumber)
23310               ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
23320                 strRptName = ("rptCourtRptCA_" & strReportNumber)
23330               Else
23340                 strRptName = ("rptCourtRptCA_0" & strReportNumber)
23350               End If

23360               For lngX = 0& To (lngCaps - 1&)
23370                 If arr_varCap(C_RNAM, lngX) = strRptName Then
23380                   strRptCap = arr_varCap(C_CAPN, lngX)
23390                   Exit For
23400                 End If
23410               Next

23420               Select Case blnUseSavedPath
                    Case True
23430                 strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
23440               Case False
23450                 DoCmd.Hourglass False
23460                 strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
23470               End Select

23480               If strRptPathFile <> vbNullString Then
23490                 DoCmd.Hourglass True
23500                 DoEvents
23510                 Select Case blnExcel
                      Case True
23520                   blnAutoStart = .chkOpenExcel
23530                 Case False
23540                   blnAutoStart = .chkOpenWord
23550                 End Select
23560                 If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
23570                 If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
23580                   Kill strRptPathFile
23590                 End If
23600                 Select Case gblnPrintAll
                      Case True
23610                   lngFiles = lngFiles + 1&
23620                   lngE = lngFiles - 1&
23630                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
23640                   arr_varFile(F_RNAM, lngE) = strRptName
23650                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
23660                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23670                   FileArraySet_CA arr_varFile  ' ** Module Procedure: modCourtReportsCA.
23680                   DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
23690                 Case False
23700                   DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
23710                 End Select
                      'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
23720                 strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23730                 If strRptPath <> .UserReportPath Then
23740                   .UserReportPath = strRptPath
23750                   SetUserReportPath_CA frm  ' ** Module Procedure: modCourtReportsCA.
23760                 End If
23770               Else
23780                 blnContinue = False
23790               End If

23800             End If  ' ** 0B.

23810           End If

23820           If ((Left(strReportNumber, 1) = "0" And .chkAssetList = True) Or (strReportNumber = "0B")) Then

23830             .chkAssetList_Start = True
23840             intRetVal_BuildAssetListInfo = .BuildAssetListInfo(#1/1/1900#, (CDate(.DateStart) - 1), "Beginning", strProc)  ' ** Form Function: frmRpt_CourtReports_CA.

23850             If intRetVal_BuildAssetListInfo = 0 Then
23860               If blnExcel = True Then

                      ' ** Prop On Hand: qryCourtReport_CA_07x (xx), For export.
23870                 strQry = "qryCourtReport_CA_07y"

23880                 strRptCap = vbNullString: strRptPathFile = vbNullString
23890                 strRptPath = .UserReportPath
23900                 strRptName = "rptCourtRptCA_00B"

23910                 strMacro = "mcrExcelExport_CR_CA" & Mid(strRptName, InStr(strRptName, "_"))

23920                 For lngX = 0& To (lngCaps - 1&)
23930                   If arr_varCap(C_RNAM, lngX) = strRptName Then
23940                     strRptCap = arr_varCap(C_CAPN, lngX)
23950                     Exit For
23960                   End If
23970                 Next

23980                 Select Case blnUseSavedPath
                      Case True
23990                   strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
24000                 Case False
24010                   DoCmd.Hourglass False
24020                   strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
24030                 End Select

24040                 If strRptPathFile <> vbNullString Then
24050                   DoCmd.Hourglass True
24060                   DoEvents
24070                   If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
24080                     Kill strRptPathFile
24090                   End If
                        ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                        ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
24100                   DoCmd.RunMacro strMacro
                        ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                        ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
24110                   If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_CA_xxxx.xls") = True Or _
                            FileExists(strRptPath & LNK_SEP & "CourtReport_CA_xxxx.xls") = True Then  ' ** 4 X's!
24120                     If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_CA_xxxx.xls") = True Then
24130                       Name (CurrentAppPath & LNK_SEP & "CourtReport_CA_xxxx.xls") As (strRptPathFile)
                            ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
24140                     Else
24150                       Name (strRptPath & LNK_SEP & "CourtReport_CA_xxxx.xls") As (strRptPathFile)
                            ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
24160                     End If
24170                     DoEvents
24180                     If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
24190                       DoEvents
24200                       Select Case blnExcel
                            Case True
24210                         blnAutoStart = .chkOpenExcel
24220                       Case False
24230                         blnAutoStart = .chkOpenWord
24240                       End Select
24250                       If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
24260                       Select Case gblnPrintAll
                            Case True
24270                         lngFiles = lngFiles + 1&
24280                         lngE = lngFiles - 1&
24290                         ReDim Preserve arr_varFile(F_ELEMS, lngE)
24300                         arr_varFile(F_RNAM, lngE) = strRptName
24310                         arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
24320                         arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24330                         FileArraySet_CA arr_varFile  ' ** Module Procedure: modCourtReportsCA.
24340                       Case False
24350                         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
24360                           EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
24370                         End If
24380                         DoEvents
24390                         If blnAutoStart = True Then
24400                           OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
24410                         End If
24420                       End Select
                            'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                            '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                            'End If
                            'DoEvents
                            'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
24430                     End If
24440                   End If
24450                   strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24460                   If strRptPath <> .UserReportPath Then
24470                     .UserReportPath = strRptPath
24480                     SetUserReportPath_CA frm  ' ** Module Procedure: modCourtReportsCA.
24490                   End If
24500                 Else
24510                   blnContinue = False
24520                 End If

24530               Else

24540                 strRptCap = vbNullString: strRptPathFile = vbNullString
24550                 strRptPath = .UserReportPath
24560                 strRptName = "rptCourtRptCA_00B"

24570                 For lngX = 0& To (lngCaps - 1&)
24580                   If arr_varCap(C_RNAM, lngX) = strRptName Then
24590                     strRptCap = arr_varCap(C_CAPN, lngX)
24600                     Exit For
24610                   End If
24620                 Next

24630                 Select Case blnUseSavedPath
                      Case True
24640                   strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
24650                 Case False
24660                   DoCmd.Hourglass False
24670                   strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
24680                 End Select

24690                 If strRptPathFile <> vbNullString Then
24700                   DoCmd.Hourglass True
24710                   DoEvents
24720                   If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
24730                     Kill strRptPathFile
24740                   End If
24750                   DoEvents
24760                   Select Case blnExcel
                        Case True
24770                     blnAutoStart = .chkOpenExcel
24780                   Case False
24790                     blnAutoStart = .chkOpenWord
24800                   End Select
24810                   If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
24820                   Select Case gblnPrintAll
                        Case True
24830                     lngFiles = lngFiles + 1&
24840                     lngE = lngFiles - 1&
24850                     ReDim Preserve arr_varFile(F_ELEMS, lngE)
24860                     arr_varFile(F_RNAM, lngE) = strRptName
24870                     arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
24880                     arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24890                     FileArraySet_CA arr_varFile  ' ** Module Procedure: modCourtReportsCA.
24900                     DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
24910                   Case False
24920                     DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
24930                   End Select
                        'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
24940                   strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24950                   If strRptPath <> .UserReportPath Then
24960                     .UserReportPath = strRptPath
24970                     SetUserReportPath_CA frm  ' ** Module Procedure: modCourtReportsCA.
24980                   End If
24990                 Else
25000                   blnContinue = False
25010                 End If

25020               End If
25030             Else
25040               blnContinue = False
                    ' ** Return codes:
                    ' **    0  Success.
                    ' **   -3  Missing entry, e.g., date.
                    ' **   -2  No data.
                    ' **   -9  Error.
25050             End If

25060             .chkAssetList_Start = False

25070           End If
25080         Else
25090           blnAllCancel = True
25100           AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
25110         End If
25120       Else
              ' ** Property On Hand, 7.

25130         CACourtReportLoad  ' ** Module Function: modCourReportsCA.
25140         intRetVal_BuildAssetListInfo = .BuildAssetListInfo(.DateStart, .DateEnd, "Ending", strProc)  ' ** Form Function: frmRpt_CourtReports_CA.

25150         If intRetVal_BuildAssetListInfo = 0 Then
25160           If blnExcel = True Then

25170             strQry = "qryCourtReport_CA_07y"  ' ** Uses tblCourtReportData.

25180             strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
25190             strRptPath = .UserReportPath

25200             If Len(strReportNumber) = 1 Then
25210               strRptName = ("rptCourtRptCA_0" & strReportNumber)
25220             ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
25230               strRptName = ("rptCourtRptCA_" & strReportNumber)
25240             Else
25250               strRptName = ("rptCourtRptCA_0" & strReportNumber)
25260             End If

25270             strMacro = "mcrExcelExport_CR_CA" & Mid(strRptName, InStr(strRptName, "_"))

25280             For lngX = 0& To (lngCaps - 1&)
25290               If arr_varCap(C_RNAM, lngX) = strRptName Then
25300                 strRptCap = arr_varCap(C_CAPN, lngX)
25310                 Exit For
25320               End If
25330             Next

25340             Select Case blnUseSavedPath
                  Case True
25350               strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
25360             Case False
25370               DoCmd.Hourglass False
25380               strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
25390             End Select

25400             If strRptPathFile <> vbNullString Then
25410               DoCmd.Hourglass True
25420               DoEvents
25430               If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
25440                 Kill strRptPathFile
25450               End If
                    ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                    ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
25460               DoCmd.RunMacro strMacro
                    ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                    ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
25470               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_CA_xxxx.xls") = True Or _
                        FileExists(strRptPath & LNK_SEP & "CourtReport_CA_xxxx.xls") = True Then  ' ** 4 X's!
25480                 If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_CA_xxxx.xls") = True Then
25490                   Name (CurrentAppPath & LNK_SEP & "CourtReport_CA_xxxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
25500                 Else
25510                   Name (strRptPath & LNK_SEP & "CourtReport_CA_xxxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
25520                 End If
25530                 DoEvents
25540                 If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
25550                   DoEvents
25560                   Select Case blnExcel
                        Case True
25570                     blnAutoStart = .chkOpenExcel
25580                   Case False
25590                     blnAutoStart = .chkOpenWord
25600                   End Select
25610                   If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
25620                   Select Case gblnPrintAll
                        Case True
25630                     lngFiles = lngFiles + 1&
25640                     lngE = lngFiles - 1&
25650                     ReDim Preserve arr_varFile(F_ELEMS, lngE)
25660                     arr_varFile(F_RNAM, lngE) = strRptName
25670                     arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
25680                     arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
25690                     FileArraySet_CA arr_varFile  ' ** Module Procedure: modCourtReportsCA.
25700                   Case False
25710                     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
25720                       EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
25730                     End If
25740                     DoEvents
25750                     If .chkOpenExcel = True Then
25760                       OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
25770                     End If
25780                   End Select
                        'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                        '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                        'End If
                        'DoEvents
                        'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
25790                 End If
25800               End If
25810               strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
25820               If strRptPath <> .UserReportPath Then
25830                 .UserReportPath = strRptPath
25840                 SetUserReportPath_CA frm  ' ** Module Procedure: modCourtReportsCA.
25850               End If
25860             Else
25870               blnContinue = False
25880             End If

25890           Else

25900             strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
25910             strRptPath = .UserReportPath

25920             If Len(strReportNumber) = 1 Then
25930               strRptName = ("rptCourtRptCA_0" & strReportNumber)
25940             ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
25950               strRptName = ("rptCourtRptCA_" & strReportNumber)
25960             Else
25970               strRptName = ("rptCourtRptCA_0" & strReportNumber)
25980             End If

25990             For lngX = 0& To (lngCaps - 1&)
26000               If arr_varCap(C_RNAM, lngX) = strRptName Then
26010                 strRptCap = arr_varCap(C_CAPN, lngX)
26020                 Exit For
26030               End If
26040             Next

26050             Select Case blnUseSavedPath
                  Case True
26060               strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
26070             Case False
26080               DoCmd.Hourglass False
26090               strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
26100             End Select

26110             If strRptPathFile <> vbNullString Then
26120               DoCmd.Hourglass True
26130               DoEvents
26140               Select Case blnExcel
                    Case True
26150                 blnAutoStart = .chkOpenExcel
26160               Case False
26170                 blnAutoStart = .chkOpenWord
26180               End Select
26190               If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
26200               If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
26210                 Kill strRptPathFile
26220               End If
26230               DoEvents
26240               Select Case gblnPrintAll
                    Case True
26250                 lngFiles = lngFiles + 1&
26260                 lngE = lngFiles - 1&
26270                 ReDim Preserve arr_varFile(F_ELEMS, lngE)
26280                 arr_varFile(F_RNAM, lngE) = strRptName
26290                 arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
26300                 arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
26310                 FileArraySet_CA arr_varFile  ' ** Module Procedure: modCourtReportsCA.
26320                 DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
26330               Case False
26340                 DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
26350               End Select
                    'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
26360               strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
26370               If strRptPath <> .UserReportPath Then
26380                 .UserReportPath = strRptPath
26390                 SetUserReportPath_CA frm  ' ** Module Procedure: modCourtReportsCA.
26400               End If
26410             Else
26420               blnContinue = False
26430             End If

26440           End If
26450         Else
26460           blnContinue = False
                ' ** Return codes:
                ' **    0  Success.
                ' **   -3  Missing entry, e.g., date.
                ' **   -2  No data.
                ' **   -9  Error.
26470         End If
26480       End If

26490       If blnContinue = False And intRetVal_BuildAssetListInfo = -2 Then
26500         DoCmd.Hourglass False
26510         MsgBox "There is no data for the report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
26520       End If

26530     End If  ' ** blnContinue.

26540   End With

26550   DoCmd.Hourglass False

EXITP:
        ' ** Check some Public variables in case there's been an error.
26560   blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
26570   blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
26580   Set rst = Nothing
26590   Set qdf = Nothing
26600   Set dbs = Nothing
26610   Exit Sub

ERRH:
26620   DoCmd.Hourglass False
26630   Select Case ERR.Number
        Case 70  ' ** Permission denied.
26640     Beep
26650     MsgBox "Trust Accountant is unable to save the file." & vbCrLf & vbCrLf & _
            "If the program to which you're exporting is open," & vbCrLf & _
            "please close it and try again.", vbInformation + vbOKOnly, "Save Failed"
26660   Case 2501  ' ** The '|' action was Canceled.
          '  ' ** Do nothing.
26670   Case 2302  ' ** Access can't save the output data to the file you've selected.
26680     Beep
26690     MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
26700   Case Else
26710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26720   End Select
26730   Resume EXITP

End Sub

Public Sub Detail_Mouse_CA(blnCalendar1_Focus As Boolean, blnCalendar2_Focus As Boolean, blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean, frm As Access.Form)

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Mouse_CA"

26810   With frm
26820     If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
26830       Select Case blnCalendar1_Focus
            Case True
26840         .cmdCalendar1_raised_semifocus_dots_img.Visible = True
26850         .cmdCalendar1_raised_img.Visible = False
26860       Case False
26870         .cmdCalendar1_raised_img.Visible = True
26880         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
26890       End Select
26900       .cmdCalendar1_raised_focus_dots_img.Visible = False
26910       .cmdCalendar1_raised_focus_img.Visible = False
26920       .cmdCalendar1_sunken_focus_dots_img.Visible = False
26930       .cmdCalendar1_raised_img_dis.Visible = False
26940     End If
26950     If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
26960       Select Case blnCalendar2_Focus
            Case True
26970         .cmdCalendar2_raised_semifocus_dots_img.Visible = True
26980         .cmdCalendar2_raised_img.Visible = False
26990       Case False
27000         .cmdCalendar2_raised_img.Visible = True
27010         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
27020       End Select
27030       .cmdCalendar2_raised_focus_dots_img.Visible = False
27040       .cmdCalendar2_raised_focus_img.Visible = False
27050       .cmdCalendar2_sunken_focus_dots_img.Visible = False
27060       .cmdCalendar2_raised_img_dis.Visible = False
27070     End If
27080     If blnPrintAll_Focus = False And (.cmdPrintAll_box01.Visible = True Or .cmdPrintAll_box02.Visible = True Or _
              .cmdPrintAll_box03.Visible = True Or .cmdPrintAll_box01.Visible = True) Then
27090       .cmdPrintAll_box01.Visible = False
27100       .cmdPrintAll_box02.Visible = False
27110       .cmdPrintAll_box03.Visible = False
27120       .cmdPrintAll_box04.Visible = False
27130     End If
27140     If blnWordAll_Focus = False And (.cmdWordAll_box01.Visible = True Or .cmdWordAll_box02.Visible = True Or _
              .cmdWordAll_box03.Visible = True Or .cmdWordAll_box01.Visible = True) Then
27150       .cmdWordAll_box01.Visible = False
27160       .cmdWordAll_box02.Visible = False
27170       .cmdWordAll_box03.Visible = False
27180       .cmdWordAll_box04.Visible = False
27190     End If
27200     If blnExcelAll_Focus = False And (.cmdExcelAll_box01.Visible = True Or .cmdExcelAll_box02.Visible = True Or _
              .cmdExcelAll_box03.Visible = True Or .cmdExcelAll_box01.Visible = True) Then
27210       .cmdExcelAll_box01.Visible = False
27220       .cmdExcelAll_box02.Visible = False
27230       .cmdExcelAll_box03.Visible = False
27240       .cmdExcelAll_box04.Visible = False
27250     End If
27260   End With

EXITP:
27270   Exit Sub

ERRH:
27280   Select Case ERR.Number
        Case Else
27290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27300   End Select
27310   Resume EXITP

End Sub

Public Sub Calendar_Handler_CA(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_CA"

        Dim strEvent As String, strCtlName As String
        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer, intNum As Integer
        Dim blnRetVal As Boolean

27410   With frm

27420     strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
27430     strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.
27440     intNum = Val(Right(strCtlName, 1))

27450     Select Case strEvent
          Case "Click"
27460       Select Case intNum
            Case 1
27470         datStartDate = Date
27480         datEndDate = 0
27490         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
27500         If blnRetVal = True Then
27510           .DateStart = datStartDate
27520         Else
27530           .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
27540         End If
27550         .DateStart.SetFocus
27560       Case 2
27570         datStartDate = Date
27580         datEndDate = 0
27590         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
27600         If blnRetVal = True Then
27610           .DateEnd = datStartDate
27620         Else
27630           .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
27640         End If
27650         .DateEnd.SetFocus
27660         Cancel = 0
27670         .DateEnd_Exit Cancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
27680         If Cancel = 0 Then
27690           .cmbAccounts.SetFocus
27700         End If
27710       End Select
27720     Case "GotFocus"
27730       Select Case intNum
            Case 1
27740         blnCalendar1_Focus = True
27750         .cmdCalendar1_raised_semifocus_dots_img.Visible = True
27760         .cmdCalendar1_raised_img.Visible = False
27770         .cmdCalendar1_raised_focus_img.Visible = False
27780         .cmdCalendar1_raised_focus_dots_img.Visible = False
27790         .cmdCalendar1_sunken_focus_dots_img.Visible = False
27800         .cmdCalendar1_raised_img_dis.Visible = False
27810       Case 2
27820         blnCalendar2_Focus = True
27830         .cmdCalendar2_raised_semifocus_dots_img.Visible = True
27840         .cmdCalendar2_raised_img.Visible = False
27850         .cmdCalendar2_raised_focus_img.Visible = False
27860         .cmdCalendar2_raised_focus_dots_img.Visible = False
27870         .cmdCalendar2_sunken_focus_dots_img.Visible = False
27880         .cmdCalendar2_raised_img_dis.Visible = False
27890       End Select
27900     Case "MouseDown"
27910       Select Case intNum
            Case 1
27920         blnCalendar1_MouseDown = True
27930         .cmdCalendar1_sunken_focus_dots_img.Visible = True
27940         .cmdCalendar1_raised_img.Visible = False
27950         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
27960         .cmdCalendar1_raised_focus_img.Visible = False
27970         .cmdCalendar1_raised_focus_dots_img.Visible = False
27980         .cmdCalendar1_raised_img_dis.Visible = False
27990       Case 2
28000         blnCalendar2_MouseDown = True
28010         .cmdCalendar2_sunken_focus_dots_img.Visible = True
28020         .cmdCalendar2_raised_img.Visible = False
28030         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
28040         .cmdCalendar2_raised_focus_img.Visible = False
28050         .cmdCalendar2_raised_focus_dots_img.Visible = False
28060         .cmdCalendar2_raised_img_dis.Visible = False
28070       End Select
28080     Case "MouseMove"
28090       Select Case intNum
            Case 1
28100         If blnCalendar1_MouseDown = False Then
28110           Select Case blnCalendar1_Focus
                Case True
28120             .cmdCalendar1_raised_focus_dots_img.Visible = True
28130             .cmdCalendar1_raised_focus_img.Visible = False
28140           Case False
28150             .cmdCalendar1_raised_focus_img.Visible = True
28160             .cmdCalendar1_raised_focus_dots_img.Visible = False
28170           End Select
28180           .cmdCalendar1_raised_img.Visible = False
28190           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
28200           .cmdCalendar1_sunken_focus_dots_img.Visible = False
28210           .cmdCalendar1_raised_img_dis.Visible = False
28220         End If
28230       Case 2
28240         If blnCalendar2_MouseDown = False Then
28250           Select Case blnCalendar2_Focus
                Case True
28260             .cmdCalendar2_raised_focus_dots_img.Visible = True
28270             .cmdCalendar2_raised_focus_img.Visible = False
28280           Case False
28290             .cmdCalendar2_raised_focus_img.Visible = True
28300             .cmdCalendar2_raised_focus_dots_img.Visible = False
28310           End Select
28320           .cmdCalendar2_raised_img.Visible = False
28330           .cmdCalendar2_raised_semifocus_dots_img.Visible = False
28340           .cmdCalendar2_sunken_focus_dots_img.Visible = False
28350           .cmdCalendar2_raised_img_dis.Visible = False
28360         End If
28370       End Select
28380     Case "MouseUp"
28390       Select Case intNum
            Case 1
28400         .cmdCalendar1_raised_focus_dots_img.Visible = True
28410         .cmdCalendar1_raised_img.Visible = False
28420         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
28430         .cmdCalendar1_raised_focus_img.Visible = False
28440         .cmdCalendar1_sunken_focus_dots_img.Visible = False
28450         .cmdCalendar1_raised_img_dis.Visible = False
28460         blnCalendar1_MouseDown = False
28470       Case 2
28480         .cmdCalendar2_raised_focus_dots_img.Visible = True
28490         .cmdCalendar2_raised_img.Visible = False
28500         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
28510         .cmdCalendar2_raised_focus_img.Visible = False
28520         .cmdCalendar2_sunken_focus_dots_img.Visible = False
28530         .cmdCalendar2_raised_img_dis.Visible = False
28540         blnCalendar2_MouseDown = False
28550       End Select
28560     Case "LostFocus"
28570       Select Case intNum
            Case 1
28580         .cmdCalendar1_raised_img.Visible = True
28590         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
28600         .cmdCalendar1_raised_focus_img.Visible = False
28610         .cmdCalendar1_raised_focus_dots_img.Visible = False
28620         .cmdCalendar1_sunken_focus_dots_img.Visible = False
28630         .cmdCalendar1_raised_img_dis.Visible = False
28640         blnCalendar1_Focus = False
28650       Case 2
28660         .cmdCalendar2_raised_img.Visible = True
28670         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
28680         .cmdCalendar2_raised_focus_img.Visible = False
28690         .cmdCalendar2_raised_focus_dots_img.Visible = False
28700         .cmdCalendar2_sunken_focus_dots_img.Visible = False
28710         .cmdCalendar2_raised_img_dis.Visible = False
28720         blnCalendar2_Focus = False
28730       End Select
28740     End Select

28750   End With

EXITP:
28760   Exit Sub

ERRH:
28770   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
28780   Case Else
28790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28800   End Select
28810   Resume EXITP

End Sub

Public Sub UserPath_After_CA(frm As Access.Form)

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "UserPath_After_CA"

28910   With frm
28920     Select Case .UserReportPath_chk
          Case True
28930       .UserReportPath_chk_lbl1.FontBold = True
28940       .UserReportPath_chk_lbl1_dim.FontBold = True
28950       .UserReportPath_chk_lbl1_dim_hi.FontBold = True
28960       .UserReportPath_chk_lbl2.FontBold = True
28970       .UserReportPath_chk_lbl2_dim.FontBold = True
28980       .UserReportPath_chk_lbl2_dim_hi.FontBold = True
28990     Case False
29000       .UserReportPath_chk_lbl1.FontBold = False
29010       .UserReportPath_chk_lbl1_dim.FontBold = False
29020       .UserReportPath_chk_lbl1_dim_hi.FontBold = False
29030       .UserReportPath_chk_lbl2.FontBold = False
29040       .UserReportPath_chk_lbl2_dim.FontBold = False
29050       .UserReportPath_chk_lbl2_dim_hi.FontBold = False
29060     End Select
29070   End With

EXITP:
29080   Exit Sub

ERRH:
29090   Select Case ERR.Number
        Case Else
29100     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29110   End Select
29120   Resume EXITP

End Sub

Public Sub PrintWordExcel_Handler_CA(strProc As String, blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean, frm As Access.Form)

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "PrintWordExcel_Handler_CA"

        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long

29210   With frm

29220     lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
29230     intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
29240     strEvent = Mid(strProc, (intPos01 + 1))
29250     strCtlName = Left(strProc, (intPos01 - 1))

29260     Select Case strEvent
          Case "GotFocus"
29270       Select Case strCtlName
            Case "cmdPrintAll"
29280         blnPrintAll_Focus = True
29290         .cmdPrintAll_box01.Visible = True
29300         .cmdPrintAll_box02.Visible = True
29310         Select Case .chkAssetList
              Case True
29320           .cmdPrintAll_box03.Visible = True
29330         Case False
29340           .cmdPrintAll_box03.Visible = False
29350         End Select
29360         .cmdPrintAll_box04.Visible = True
29370       Case "cmdWordAll"
29380         blnWordAll_Focus = True
29390         .cmdWordAll_box01.Visible = True
29400         .cmdWordAll_box02.Visible = True
29410         Select Case .chkAssetList
              Case True
29420           .cmdWordAll_box03.Visible = True
29430         Case False
29440           .cmdWordAll_box03.Visible = False
29450         End Select
29460         .cmdWordAll_box04.Visible = True
29470       Case "cmdExcelAll"
29480         blnExcelAll_Focus = True
29490         .cmdExcelAll_box01.Visible = True
29500         .cmdExcelAll_box02.Visible = True
29510         Select Case .chkAssetList
              Case True
29520           .cmdExcelAll_box03.Visible = True
29530         Case False
29540           .cmdExcelAll_box03.Visible = False
29550         End Select
29560         .cmdExcelAll_box04.Visible = True
29570       End Select
29580     Case "MouseMove"
29590       Select Case strCtlName
            Case "cmdPrintAll"
29600         If gblnPrintAll = False Then
29610           .cmdPrintAll_box01.Visible = True
29620           .cmdPrintAll_box02.Visible = True
29630           Select Case .chkAssetList
                Case True
29640             .cmdPrintAll_box03.Visible = True
29650           Case False
29660             .cmdPrintAll_box03.Visible = False
29670           End Select
29680           .cmdPrintAll_box04.Visible = True
29690           If blnWordAll_Focus = False Then
29700             .cmdWordAll_box01.Visible = False
29710             .cmdWordAll_box02.Visible = False
29720             .cmdWordAll_box03.Visible = False
29730             .cmdWordAll_box04.Visible = False
29740           End If
29750           If blnExcelAll_Focus = False Then
29760             .cmdExcelAll_box01.Visible = False
29770             .cmdExcelAll_box02.Visible = False
29780             .cmdExcelAll_box03.Visible = False
29790             .cmdExcelAll_box04.Visible = False
29800           End If
29810         End If
29820       Case "cmdWordAll"
29830         .cmdWordAll_box01.Visible = True
29840         .cmdWordAll_box02.Visible = True
29850         Select Case .chkAssetList
              Case True
29860           .cmdWordAll_box03.Visible = True
29870         Case False
29880           .cmdWordAll_box03.Visible = False
29890         End Select
29900         .cmdWordAll_box04.Visible = True
29910         If blnPrintAll_Focus = False Then
29920           .cmdPrintAll_box01.Visible = False
29930           .cmdPrintAll_box02.Visible = False
29940           .cmdPrintAll_box03.Visible = False
29950           .cmdPrintAll_box04.Visible = False
29960         End If
29970         If blnExcelAll_Focus = False Then
29980           .cmdExcelAll_box01.Visible = False
29990           .cmdExcelAll_box02.Visible = False
30000           .cmdExcelAll_box03.Visible = False
30010           .cmdExcelAll_box04.Visible = False
30020         End If
30030       Case "cmdExcelAll"
30040         .cmdExcelAll_box01.Visible = True
30050         .cmdExcelAll_box02.Visible = True
30060         Select Case .chkAssetList
              Case True
30070           .cmdExcelAll_box03.Visible = True
30080         Case False
30090           .cmdExcelAll_box03.Visible = False
30100         End Select
30110         .cmdExcelAll_box04.Visible = True
30120         If blnPrintAll_Focus = False Then
30130           .cmdPrintAll_box01.Visible = False
30140           .cmdPrintAll_box02.Visible = False
30150           .cmdPrintAll_box03.Visible = False
30160           .cmdPrintAll_box04.Visible = False
30170         End If
30180         If blnWordAll_Focus = False Then
30190           .cmdWordAll_box01.Visible = False
30200           .cmdWordAll_box02.Visible = False
30210           .cmdWordAll_box03.Visible = False
30220           .cmdWordAll_box04.Visible = False
30230         End If
30240       End Select
30250     Case "LostFocus"
30260       Select Case strCtlName
            Case "cmdPrintAll"
30270         .cmdPrintAll_box01.Visible = False
30280         .cmdPrintAll_box02.Visible = False
30290         .cmdPrintAll_box03.Visible = False
30300         .cmdPrintAll_box04.Visible = False
30310         blnPrintAll_Focus = False
30320       Case "cmdWordAll"
30330         .cmdWordAll_box01.Visible = False
30340         .cmdWordAll_box02.Visible = False
30350         .cmdWordAll_box03.Visible = False
30360         .cmdWordAll_box04.Visible = False
30370         blnWordAll_Focus = False
30380       Case "cmdExcelAll"
30390         .cmdExcelAll_box01.Visible = False
30400         .cmdExcelAll_box02.Visible = False
30410         .cmdExcelAll_box03.Visible = False
30420         .cmdExcelAll_box04.Visible = False
30430         blnExcelAll_Focus = False
30440       End Select
30450     End Select
30460   End With

EXITP:
30470   Exit Sub

ERRH:
30480   Select Case ERR.Number
        Case Else
30490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
30500   End Select
30510   Resume EXITP

End Sub
