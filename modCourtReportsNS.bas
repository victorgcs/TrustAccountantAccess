Attribute VB_Name = "modCourtReportsNS"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCourtReportsNS"

'VGC 09/14/2017: CHANGES!

' ** Array: arr_varNSRpt().
Private lngNSRpts As Long, arr_varNSRpt As Variant
'Private Const CR_ID     As Integer = 0
Private Const CR_NUM    As Integer = 1
'Private Const CR_CAT    As Integer = 2
Private Const CR_CON    As Integer = 3
Private Const CR_DIV    As Integer = 4
'Private Const CR_DIVTXT As Integer = 5
Private Const CR_DIVTTL As Integer = 6
'Private Const CR_GRP    As Integer = 7
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
'Private Const C_RID   As Integer = 0
Private Const C_RNAM  As Integer = 1
'Private Const C_CAP   As Integer = 2
Private Const C_CAPN  As Integer = 3

Private blnExcel As Boolean, blnAllCancel As Boolean
Private strThisProc As String ', strCaseNum As String
' **

Public Function NSBuildCourtReportData(strReportNumber As String) As Integer
' ** Return Codes:
' **   0  Success.
' **  -1  Canceled.
' **  -9  Error.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "NSBuildCourtReportData"

        'Dim rsxDataIn As ADODB.Recordset, rsxDataOut As ADODB.Recordset  ' ** Early binding.
        Dim rsxDataIn As Object, rsxDataOut As Object                     ' ** Late binding.
        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim datTmp01 As Date
        Dim intReportNumber As Integer
        Dim intX As Integer
        Dim intRetVal As Integer

110     intRetVal = 0

120     If glngTaxCode_Distribution = 0& Then
130       glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
140     End If

        ' ** Delete the data from the tmpCourtReport table.
150     Set dbs = CurrentDb
160     With dbs
170       Set qdf = .QueryDefs("qryCourtReport_02")
180       qdf.Execute
190       .Close
200     End With
210     Set qdf = Nothing
220     Set dbs = Nothing

        'Set rsxDataOut = New ADODB.Recordset             ' ** Early binding.
230     Set rsxDataOut = CreateObject("ADODB.Recordset")  ' ** Late binding.
240   On Error Resume Next
250     rsxDataOut.Open "tmpCourtReportData", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTable
260     If ERR.Number <> 0 Then
270       Select Case ERR.Number
          Case -2147217838  ' ** Data source object is already initialized.
280   On Error GoTo ERRH
            ' ** For now, just let it go, since I think that means it's already available.
290       Case Else
300         intRetVal = -9
310         zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320   On Error GoTo ERRH
330       End Select
340     Else
350   On Error GoTo ERRH
360     End If

370     If intRetVal = 0 Then

          ' ** Build dummy records with zero in the amount to insure that all report sections are displayed.
380       intX = 1
390       Do While intX <= 9
400         With rsxDataOut
410           .AddNew
420           .Fields("ReportNumber") = intX * 10
430           .Fields("ReportCategory") = NSCourtReportCategory(intX * 10)
440           .Fields("ReportGroup") = NSCourtReportGroup(intX * 10)
450           .Fields("ReportDivision") = NSCourtReportDivision(intX * 10)
460           .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(intX * 10)
470           .Fields("ReportDivisionText") = NSCourtReportDivisionText(intX * 10)
480           .Fields("ReportGroupText") = NSCourtReportGroupText(intX * 10)
490           .Fields("accountno") = gstrAccountNo
500           .Fields("date") = gdatEndDate
510           .Fields("journaltype") = "Miscellaneous"
520           .Fields("Amount") = 0
530           .Fields("revcode_ID") = 0
540           .Fields("revcode_DESC") = "Dummy entry"
550           .Fields("revcode_TYPE") = 1
560           .Fields("revcode_SORTORDER") = 0
570           .Update
580         End With
590         intX = intX + 1
600       Loop

610       If strReportNumber = "0" Or strReportNumber = "0A" Then
620         intRetVal = NSGetCourtReportData  ' ** Function: Below.
            ' ** Return Codes:
            ' **   0  Success.
            ' **  -1  Canceled.
            ' **  -9  Error.
630       End If

640     Else
          ' ** rsxDataOut failed to open.
650     End If

660     If intRetVal = 0 Then

          'Set rsxDataIn = New ADODB.Recordset             ' ** Early binding.
670       Set rsxDataIn = CreateObject("ADODB.Recordset")  ' ** Late binding.
680       rsxDataIn.Open "qryCourtReport_NS_00_32", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect

          ' ** Loop through data processing records for requested account.
690       Do While rsxDataIn.EOF = False
700         If Trim(rsxDataIn.Fields("accountno")) = gstrAccountNo Then
              ' ** I have no explanation for why this Trim() is necessary! VGC 02/27/2013.
710           intReportNumber = rsxDataIn.Fields("Reportnumber")  ' ** Do this because it recalcs for each time it is refernced.
              ' ** Find the right date to use.
720           datTmp01 = NSCourtReportDate(intReportNumber, rsxDataIn.Fields("transdate"), rsxDataIn.Fields("assetdate"), rsxDataIn.Fields("journaltype"))

              ' ** If the date for the transaction is within range, build a report record.
730           If datTmp01 >= gdatStartDate And datTmp01 < gdatEndDate + 1 _
                  Or (rsxDataIn.Fields("journaltype") = "Sold" And _
                  rsxDataIn.Fields("transdate") >= gdatStartDate And rsxDataIn.Fields("transdate") < gdatEndDate + 1 And _
                  rsxDataIn.Fields("assetdate") < gdatStartDate) Then
                ' ** If the journal type is misc create 2 transactions.
740             With rsxDataOut
750               Select Case rsxDataIn.Fields("journaltype")
                    ' ** Investments Made - Report 5; Case 50 = "Investments Made"
                  Case "Misc."
760                 If rsxDataIn.Fields("icash") > 0 Then
770                   .AddNew
780                   .Fields("ReportNumber") = 70  ' ** Receipts of Income.
790                   .Fields("ReportCategory") = NSCourtReportCategory(70)
800                   .Fields("ReportGroup") = NSCourtReportGroup(70)
810                   .Fields("ReportDivision") = NSCourtReportDivision(70)
820                   .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(70)
830                   .Fields("ReportDivisionText") = NSCourtReportDivisionText(70)
840                   .Fields("ReportGroupText") = NSCourtReportGroupText(70)
850                   .Fields("Amount") = rsxDataIn.Fields("icash")
860                   .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
870                   .Fields("date") = datTmp01
880                   .Fields("journaltype") = "Miscellaneous"
890                   .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
900                   .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
910                   .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
920                   .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
930                   .Update
940                 ElseIf rsxDataIn.Fields("icash") < 0 Then
                      ' ** Disbursements of Income - Report 8/8A; Case 80 = "Disbursements of Income"
                      '"(journaltype = ""Misc."" And icash < 0)"
950                   .AddNew
960                   .Fields("ReportNumber") = 80  ' ** Disbursements of Income.
970                   .Fields("ReportCategory") = NSCourtReportCategory(80)
980                   .Fields("ReportGroup") = NSCourtReportGroup(80)
990                   .Fields("ReportDivision") = NSCourtReportDivision(80)
1000                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(80)
1010                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(80)
1020                  .Fields("ReportGroupText") = NSCourtReportGroupText(80)
1030                  .Fields("Amount") = rsxDataIn.Fields("icash")
1040                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
1050                  .Fields("date") = datTmp01
1060                  .Fields("journaltype") = "Miscellaneous"
1070                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
1080                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
1090                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
1100                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
1110                  .Update
1120                End If
1130                If rsxDataIn.Fields("pcash") > 0 Then
1140                  .AddNew
1150                  .Fields("ReportNumber") = 10  ' ** Receipts of Principal.
1160                  .Fields("ReportCategory") = NSCourtReportCategory(10)
1170                  .Fields("ReportGroup") = NSCourtReportGroup(10)
1180                  .Fields("ReportDivision") = NSCourtReportDivision(10)
1190                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(10)
1200                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(10)
1210                  .Fields("ReportGroupText") = NSCourtReportGroupText(10)
1220                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
1230                  .Fields("date") = datTmp01
1240                  .Fields("journaltype") = "Miscellaneous"
1250                  .Fields("Amount") = rsxDataIn.Fields("pcash")
1260                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
1270                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
1280                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
1290                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
1300                  .Update
1310                ElseIf rsxDataIn.Fields("pcash") < 0 Then
1320                  .AddNew
1330                  .Fields("ReportNumber") = 30  ' ** Disbursements of Principal
1340                  .Fields("ReportCategory") = NSCourtReportCategory(30)
1350                  .Fields("ReportGroup") = NSCourtReportGroup(30)
1360                  .Fields("ReportDivision") = NSCourtReportDivision(30)
1370                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(30)
1380                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(30)
1390                  .Fields("ReportGroupText") = NSCourtReportGroupText(30)
1400                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
1410                  .Fields("date") = datTmp01
1420                  .Fields("journaltype") = "Miscellaneous"
1430                  .Fields("Amount") = rsxDataIn.Fields("pcash")
1440                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
1450                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
1460                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
1470                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
1480                  .Update
1490                End If
1500              Case "Sold"
1510                If rsxDataIn.Fields("GainLoss") = 0 Then
                      ' ** Changes in Investment Holdings - Report 6; Case 60 = "Changes in Investment Holdings"
1520                  .AddNew  ' ** Force a Changes in Investment Holdings record.
1530                  .Fields("ReportNumber") = 60
1540                  .Fields("ReportCategory") = NSCourtReportCategory(60)
1550                  .Fields("ReportGroup") = NSCourtReportGroup(60)
1560                  .Fields("ReportDivision") = NSCourtReportDivision(60)
1570                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(60)
1580                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(60)
1590                  .Fields("ReportGroupText") = NSCourtReportGroupText(60)
1600                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
1610                  .Fields("date") = rsxDataIn.Fields("transdate")
1620                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
1630                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
1640                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
1650                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
1660                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
1670                  .Fields("Amount") = rsxDataIn.Fields("cost")
1680                  .Update
1690                Else
                      ' ** Gains (Losses) on Sale or Other Dispositions - Report 2; Case 20 = "Gain (Loss) on Sale or Other Distributions"
1700                  .AddNew  ' ** Force a Gains (Losses) on Sale or Other Dispositions record.
1710                  .Fields("ReportNumber") = 20
1720                  .Fields("ReportCategory") = NSCourtReportCategory(20)
1730                  .Fields("ReportGroup") = NSCourtReportGroup(20)
1740                  .Fields("ReportDivision") = NSCourtReportDivision(20)
1750                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(20)
1760                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(20)
1770                  .Fields("ReportGroupText") = NSCourtReportGroupText(20)
1780                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
1790                  .Fields("date") = rsxDataIn.Fields("transdate")
1800                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
1810                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
1820                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
1830                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
1840                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
1850                  .Fields("Amount") = (rsxDataIn.Fields("pcash") + rsxDataIn.Fields("Cost"))
1860                  .Update
1870                End If
                    ' ** Write a receipt of interest record for any sales record that gets here.
1880                If rsxDataIn.Fields("icash") > 0 Then
1890                  If (rsxDataIn.Fields("pcash") = 0 And rsxDataIn.Fields("icash") > 0 And _
                          (rsxDataIn.Fields("icash") = -rsxDataIn.Fields("cost"))) Then
                        ' ** Skip it.
1900                  Else
1910                    .AddNew  ' ** Force an INTEREST of income record.
1920                    .Fields("ReportNumber") = 70
1930                    .Fields("ReportCategory") = NSCourtReportCategory(70)
1940                    .Fields("ReportGroup") = NSCourtReportGroup(70)
1950                    .Fields("ReportDivision") = NSCourtReportDivision(70)
1960                    .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(70)
1970                    .Fields("ReportDivisionText") = NSCourtReportDivisionText(70)
1980                    .Fields("ReportGroupText") = NSCourtReportGroupText(70)
1990                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2000                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2010                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2020                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2030                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2040                    If datTmp01 < gdatStartDate Then  ' ** Override date for sales that span years.
2050                      .Fields("date") = rsxDataIn.Fields("transdate")
2060                    Else
2070                      .Fields("date") = datTmp01
2080                    End If
                        ' ** Icash on Sold journaltype should be reported as Interest.
2090                    .Fields("journaltype") = "Interest"
2100                    .Fields("Amount") = rsxDataIn.Fields("icash")
2110                    .Update
2120                  End If
2130                End If
                    ' ** Only add a Sold Record without the icash if the transdate (set earlier)
                    ' ** is in the reporting period.
2140                If datTmp01 >= gdatStartDate And datTmp01 < gdatEndDate + 1 And intReportNumber <> 60 And intReportNumber <> 20 Then
2150                  .AddNew
2160                  .Fields("ReportNumber") = intReportNumber
2170                  .Fields("ReportCategory") = NSCourtReportCategory(intReportNumber)
2180                  .Fields("ReportGroup") = NSCourtReportGroup(intReportNumber)
2190                  .Fields("ReportDivision") = NSCourtReportDivision(intReportNumber)
2200                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(intReportNumber)
2210                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(intReportNumber)
2220                  .Fields("ReportGroupText") = NSCourtReportGroupText(intReportNumber)
2230                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2240                  .Fields("date") = datTmp01
2250                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
2260                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2270                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2280                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2290                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2300                  .Fields("Amount") = rsxDataIn.Fields("gainloss")
2310                  .Update
2320                End If
2330              Case "Liability"
2340                If ((rsxDataIn.Fields("cost") < 0) And (rsxDataIn.Fields("cost") = (-rsxDataIn.Fields("pcash")))) Then
                      ' ** Information for Investments Made - Report 5; Case 50 = "Investments Made"  'VGC 02/16/2013: PER RICH!
2350                  .AddNew  ' ** Force an Investments Made record.
2360                  .Fields("ReportNumber") = 50
2370                  .Fields("ReportCategory") = NSCourtReportCategory(50)
2380                  .Fields("ReportGroup") = NSCourtReportGroup(50)
2390                  .Fields("ReportDivision") = NSCourtReportDivision(50)
2400                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(50)
2410                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(50)
2420                  .Fields("ReportGroupText") = NSCourtReportGroupText(50)
2430                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2440                  .Fields("date") = rsxDataIn.Fields("transdate")
2450                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
2460                  .Fields("Amount") = rsxDataIn.Fields("cost")
2470                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2480                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2490                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2500                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2510                  .Update
2520                End If
2530                If ((rsxDataIn.Fields("cost") > 0) And (rsxDataIn.Fields("cost") = (-rsxDataIn.Fields("pcash")))) Then
                      ' ** Changes in Investment Holdings - Report 6; Case 60 = "Changes in Investment Holdings"
2540                  .AddNew  ' ** Force a Changes in Investment Holdings record.
2550                  .Fields("ReportNumber") = 60
2560                  .Fields("ReportCategory") = NSCourtReportCategory(60)
2570                  .Fields("ReportGroup") = NSCourtReportGroup(60)
2580                  .Fields("ReportDivision") = NSCourtReportDivision(60)
2590                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(60)
2600                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(60)
2610                  .Fields("ReportGroupText") = NSCourtReportGroupText(60)
2620                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2630                  .Fields("date") = rsxDataIn.Fields("transdate")
2640                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
2650                  .Fields("Amount") = rsxDataIn.Fields("cost")
2660                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2670                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2680                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2690                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2700                  .Update
2710                End If
2720                If rsxDataIn.Fields("icash") < 0 Then
                      ' ** Disbursements of Income - Report 8/8A; Case 80 = "Disbursements of Income"
2730                  .AddNew  ' ** Force a Disbursement of income record.
2740                  .Fields("ReportNumber") = 80
2750                  .Fields("ReportCategory") = NSCourtReportCategory(80)
2760                  .Fields("ReportGroup") = NSCourtReportGroup(80)
2770                  .Fields("ReportDivision") = NSCourtReportDivision(80)
2780                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(80)
2790                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(80)
2800                  .Fields("ReportGroupText") = NSCourtReportGroupText(80)
2810                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
2820                  .Fields("date") = datTmp01
2830                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
2840                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
2850                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
2860                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
2870                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
2880                  .Fields("Amount") = rsxDataIn.Fields("icash")
2890                  .Update
2900                End If
2910                If (rsxDataIn.Fields("cost") <> 0 And rsxDataIn.Fields("pcash") = 0) Then
                      ' ** Receipts of Principal - Report 1; Case 10 = "Receipts of Principal"
2920                  .AddNew  ' ** Force a Receipts of Principal record.
2930                  .Fields("ReportNumber") = 10
2940                  .Fields("ReportCategory") = NSCourtReportCategory(10)
2950                  .Fields("ReportGroup") = NSCourtReportGroup(10)
2960                  .Fields("ReportDivision") = NSCourtReportDivision(10)
2970                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(10)
2980                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(10)
2990                  .Fields("ReportGroupText") = NSCourtReportGroupText(10)
3000                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3010                  .Fields("date") = datTmp01
3020                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3030                  .Fields("Amount") = rsxDataIn.Fields("cost")
3040                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3050                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3060                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3070                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3080                  .Update
3090                End If
                    'If (rsxDataIn.Fields("pcash") = -rsxDataIn.Fields("cost")) Then  ' ** VGC 03/04/2013: NEW! (i hope it's right)
                    '  ' ** Receipts of Income - Report 7; Case 70 = "Receipts of Income"
                    '  .AddNew  ' ** Force a Receipts of Income record.
                    '  .Fields("ReportNumber") = 70
                    '  .Fields("ReportCategory") = NSCourtReportCategory(70)
                    '  .Fields("ReportGroup") = NSCourtReportGroup(70)
                    '  .Fields("ReportDivision") = NSCourtReportDivision(70)
                    '  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(70)
                    '  .Fields("ReportDivisionText") = NSCourtReportDivisionText(70)
                    '  .Fields("ReportGroupText") = NSCourtReportGroupText(70)
                    '  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
                    '  .Fields("date") = datTmp01
                    '  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
                    '  .Fields("Amount") = rsxDataIn.Fields("cost")
                    '  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
                    '  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
                    '  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
                    '  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
                    '  .Update
                    'End If
3100              Case "Paid"
3110                If rsxDataIn.Fields("icash") <> 0 Then
3120                  If rsxDataIn.Fields("taxcode") <> glngTaxCode_Distribution Then  '<> "Distribution"
                        ' ** Disbursements of Income - Report 8/8A; Case 80 = "Disbursements of Income"
3130                    .AddNew  ' ** Force a Disbursement of income record.
3140                    .Fields("ReportNumber") = 80
3150                    .Fields("ReportCategory") = NSCourtReportCategory(80)
3160                    .Fields("ReportGroup") = NSCourtReportGroup(80)
3170                    .Fields("ReportDivision") = NSCourtReportDivision(80)
3180                    .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(80)
3190                    .Fields("ReportDivisionText") = NSCourtReportDivisionText(80)
3200                    .Fields("ReportGroupText") = NSCourtReportGroupText(80)
3210                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3220                    .Fields("date") = datTmp01
3230                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3240                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3250                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3260                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3270                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3280                    .Fields("Amount") = rsxDataIn.Fields("icash")
3290                    .Update
3300                  Else  ' ** rsxDataIn.Fields("taxcode") = 11.
                        ' ** Distributions of Income - Report 9; Case 90 = "Distributions of Income"
3310                    .AddNew  ' ** Force a Distribution of income record.
3320                    .Fields("ReportNumber") = 90
3330                    .Fields("ReportCategory") = NSCourtReportCategory(90)
3340                    .Fields("ReportGroup") = NSCourtReportGroup(90)
3350                    .Fields("ReportDivision") = NSCourtReportDivision(90)
3360                    .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(90)
3370                    .Fields("ReportDivisionText") = NSCourtReportDivisionText(90)
3380                    .Fields("ReportGroupText") = NSCourtReportGroupText(90)
3390                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3400                    .Fields("date") = datTmp01
3410                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3420                    .Fields("Amount") = rsxDataIn.Fields("icash")
3430                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3440                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3450                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3460                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3470                    .Update
3480                  End If
3490                End If
3500                If rsxDataIn.Fields("pcash") <> 0 Then
3510                  If rsxDataIn.Fields("taxcode") <> glngTaxCode_Distribution Then
3520                    .AddNew  ' ** Force a Disbursement of principal record.
3530                    .Fields("ReportNumber") = 30
3540                    .Fields("ReportCategory") = NSCourtReportCategory(30)
3550                    .Fields("ReportGroup") = NSCourtReportGroup(30)
3560                    .Fields("ReportDivision") = NSCourtReportDivision(30)
3570                    .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(30)
3580                    .Fields("ReportDivisionText") = NSCourtReportDivisionText(30)
3590                    .Fields("ReportGroupText") = NSCourtReportGroupText(30)
3600                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3610                    .Fields("date") = datTmp01
3620                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3630                    .Fields("Amount") = rsxDataIn.Fields("pcash")
3640                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3650                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3660                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3670                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3680                    .Update
3690                  Else  ' ** rsxDataIn.Fields("taxcode") = 11.  '= "Distribution"
3700                    .AddNew  ' ** Force a Disbtribution of principal record.
3710                    .Fields("ReportNumber") = 40
3720                    .Fields("ReportCategory") = NSCourtReportCategory(40)
3730                    .Fields("ReportGroup") = NSCourtReportGroup(40)
3740                    .Fields("ReportDivision") = NSCourtReportDivision(40)
3750                    .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(40)
3760                    .Fields("ReportDivisionText") = NSCourtReportDivisionText(40)
3770                    .Fields("ReportGroupText") = NSCourtReportGroupText(40)
3780                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3790                    .Fields("date") = datTmp01
3800                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
3810                    .Fields("Amount") = rsxDataIn.Fields("pcash")
3820                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
3830                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
3840                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
3850                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
3860                    .Update
3870                  End If
3880                End If
3890              Case "Purchase"
                    ' ** Investments Made - Report 5; Case 50 = "Investments Made"
3900                .AddNew  ' ** Force an Investments Made record.
3910                .Fields("ReportNumber") = 50
3920                .Fields("ReportCategory") = NSCourtReportCategory(50)
3930                .Fields("ReportGroup") = NSCourtReportGroup(50)
3940                .Fields("ReportDivision") = NSCourtReportDivision(50)
3950                .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(50)
3960                .Fields("ReportDivisionText") = NSCourtReportDivisionText(50)
3970                .Fields("ReportGroupText") = NSCourtReportGroupText(50)
3980                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
3990                .Fields("date") = rsxDataIn.Fields("transdate")
4000                .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4010                .Fields("Amount") = _
                      ((IIf(rsxDataIn.Fields("icash") = 0, rsxDataIn.Fields("pcash"), _
                      IIf(rsxDataIn.Fields("pcash") = 0, rsxDataIn.Fields("icash"), _
                      rsxDataIn.Fields("pcash")))) * -1)
4020                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4030                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4040                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4050                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4060                .Update
4070                If (rsxDataIn.Fields("icash") < 0) And (rsxDataIn.Fields("pcash") = -(rsxDataIn.Fields("cost"))) Then
                      ' ** Receipts of Income - Report 7; Case 70 = "Receipts of Income"
4080                  .AddNew  ' ** Force a Receipts of Income record.
4090                  .Fields("ReportNumber") = 70
4100                  .Fields("ReportCategory") = NSCourtReportCategory(70)
4110                  .Fields("ReportGroup") = NSCourtReportGroup(70)
4120                  .Fields("ReportDivision") = NSCourtReportDivision(70)
4130                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(70)
4140                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(70)
4150                  .Fields("ReportGroupText") = NSCourtReportGroupText(70)
4160                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4170                  .Fields("date") = datTmp01
4180                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4190                  .Fields("Amount") = rsxDataIn.Fields("icash")
4200                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4210                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4220                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4230                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4240                  .Update
4250                End If
4260                If intReportNumber <> 50 Then
4270                  .AddNew
4280                  .Fields("ReportNumber") = intReportNumber
4290                  .Fields("ReportCategory") = NSCourtReportCategory(intReportNumber)
4300                  .Fields("ReportGroup") = NSCourtReportGroup(intReportNumber)
4310                  .Fields("ReportDivision") = NSCourtReportDivision(intReportNumber)
4320                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(intReportNumber)
4330                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(intReportNumber)
4340                  .Fields("ReportGroupText") = NSCourtReportGroupText(intReportNumber)
4350                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4360                  .Fields("date") = datTmp01
4370                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4380                  .Fields("Amount") = NSCourtReportDollars(intReportNumber, _
                        rsxDataIn.Fields("pcash"), _
                        rsxDataIn.Fields("icash"), _
                        rsxDataIn.Fields("cost"), _
                        rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
4390                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4400                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4410                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4420                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4430                  .Update
4440                End If
4450              Case "Deposit"
4460                If InStr(Nz(rsxDataIn.Fields("jcomment"), vbNullString), "stock split") > 0 Then
                      ' ** Changes in Investment Holdings - Report 6; Case 60 = "Changes in Investment Holdings"
4470                  .AddNew  ' ** Force a Changes in Investment Holdings record.
4480                  .Fields("ReportNumber") = 60
4490                  .Fields("ReportCategory") = NSCourtReportCategory(60)
4500                  .Fields("ReportGroup") = NSCourtReportGroup(60)
4510                  .Fields("ReportDivision") = NSCourtReportDivision(60)
4520                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(60)
4530                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(60)
4540                  .Fields("ReportGroupText") = NSCourtReportGroupText(60)
4550                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4560                  .Fields("date") = rsxDataIn.Fields("transdate")
4570                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4580                  .Fields("Amount") = rsxDataIn.Fields("cost")
4590                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4600                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4610                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4620                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4630                  .Update
4640                End If
4650                If intReportNumber <> 60 Then
4660                  .AddNew
4670                  .Fields("ReportNumber") = intReportNumber
4680                  .Fields("ReportCategory") = NSCourtReportCategory(intReportNumber)
4690                  .Fields("ReportGroup") = NSCourtReportGroup(intReportNumber)
4700                  .Fields("ReportDivision") = NSCourtReportDivision(intReportNumber)
4710                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(intReportNumber)
4720                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(intReportNumber)
4730                  .Fields("ReportGroupText") = NSCourtReportGroupText(intReportNumber)
4740                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4750                  .Fields("date") = datTmp01
4760                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4770                  .Fields("Amount") = NSCourtReportDollars(intReportNumber, _
                        rsxDataIn.Fields("pcash"), _
                        rsxDataIn.Fields("icash"), _
                        rsxDataIn.Fields("cost"), _
                        rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
4780                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4790                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4800                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
4810                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
4820                  .Update
4830                End If
4840              Case Else
4850                .AddNew
4860                .Fields("ReportNumber") = intReportNumber
4870                .Fields("ReportCategory") = NSCourtReportCategory(intReportNumber)
4880                .Fields("ReportGroup") = NSCourtReportGroup(intReportNumber)
4890                .Fields("ReportDivision") = NSCourtReportDivision(intReportNumber)
4900                .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(intReportNumber)
4910                .Fields("ReportDivisionText") = NSCourtReportDivisionText(intReportNumber)
4920                .Fields("ReportGroupText") = NSCourtReportGroupText(intReportNumber)
4930                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
4940                .Fields("date") = datTmp01
4950                .Fields("journaltype") = rsxDataIn.Fields("journaltype")
4960                .Fields("Amount") = NSCourtReportDollars(intReportNumber, _
                      rsxDataIn.Fields("pcash"), _
                      rsxDataIn.Fields("icash"), _
                      rsxDataIn.Fields("cost"), _
                      rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
4970                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
4980                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
4990                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5000                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5010                .Update
5020                If rsxDataIn.Fields("journaltype") = "Received" And _
                        rsxDataIn.Fields("pcash") > 0 And rsxDataIn.Fields("icash") > 0 Then
                      ' ** Force a 2nd entry for the PCash side of the transaction.
5030                  .AddNew
5040                  .Fields("ReportNumber") = 70  ' ** Receipts of Income.
5050                  .Fields("ReportCategory") = NSCourtReportCategory(70)
5060                  .Fields("ReportGroup") = NSCourtReportGroup(70)
5070                  .Fields("ReportDivision") = NSCourtReportDivision(70)
5080                  .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(70)
5090                  .Fields("ReportDivisionText") = NSCourtReportDivisionText(70)
5100                  .Fields("ReportGroupText") = NSCourtReportGroupText(70)
5110                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
5120                  .Fields("date") = datTmp01
5130                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
5140                  .Fields("Amount") = NSCourtReportDollars(70, _
                        rsxDataIn.Fields("pcash"), _
                        rsxDataIn.Fields("icash"), _
                        rsxDataIn.Fields("cost"), _
                        rsxDataIn.Fields("journaltype"))  ' ** Function: Below.
5150                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
5160                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
5170                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5180                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5190                  .Update
5200                End If
5210              End Select  ' ** journaltype.
5220            End With  ' ** rsxDataOut.
5230          End If  ' ** datTmp01.
5240        End If  ' ** gstrAccountNo.
5250        rsxDataIn.MoveNext
5260      Loop

5270      rsxDataOut.Close

5280    End If

EXITP:
5290    Set qdf = Nothing
5300    Set dbs = Nothing
5310    Set rsxDataIn = Nothing
5320    Set rsxDataOut = Nothing
5330    NSBuildCourtReportData = intRetVal
5340    Exit Function

ERRH:
5350    intRetVal = -9  ' ** Error.
5360    Select Case ERR.Number
        Case Else
5370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5380    End Select
5390    Resume EXITP

End Function

Public Function NSCourtReportCategory(intCourtReport As Integer) As String

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportCategory"

        Dim strRetVal As String

5410    Select Case intCourtReport
        Case 5
5420      strRetVal = "Principal on Hand at Beginning of Period"
5430    Case 10
5440      strRetVal = "Receipts of Principal"
5450    Case 20
5460      strRetVal = "Gain (Loss) on Sale or Other Distributions"
5470    Case 30
5480      strRetVal = "Disbursements of Principal"
5490    Case 40
5500      strRetVal = "Distributions of Principal"
5510    Case 50
5520      strRetVal = "Investments Made"
5530    Case 60
5540      strRetVal = "Changes in Investment Holdings"
5550    Case 69
5560      strRetVal = "Income on Hand at Beginning of Period"
5570    Case 70
5580      strRetVal = "Receipts of Income"
5590    Case 80
5600      strRetVal = "Disbursements of Income"
5610    Case 90
5620      strRetVal = "Distributions of Income"
5630    Case Else
5640      strRetVal = "Unknown"
5650    End Select

EXITP:
5660    NSCourtReportCategory = strRetVal
5670    Exit Function

ERRH:
5680    strRetVal = vbNullString
5690    Select Case ERR.Number
        Case Else
5700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5710    End Select
5720    Resume EXITP

End Function

Public Function NSCourtReportDataID(strComment As String, curPCash As Currency, curICash As Currency, curCost As Currency, curGainLoss As Currency, strTaxCode As String, strJournalType As String) As Integer
' ** The code returned is the Court Report that the data belongs to.
' ** PROBLEM: RECEIVED items with both ICash and PCash only pass
' **          through here once, giving them a ReportNumber of 10 for
' **          one side of the cash only, and ignoring ReportNumber 70!

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportDataID"

        Dim lngTaxcode As Long
        Dim intRetVal As Integer

5810    intRetVal = 0  ' ** Force to 0 just in case.

5820    If glngTaxCode_Distribution = 0& Then
5830      glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
5840    End If
5850    If Trim(strTaxCode) <> vbNullString Then
5860      lngTaxcode = Val(strTaxCode)
5870    Else
5880      lngTaxcode = 0&
5890    End If

        ' ** Receipts of Principal.
5900    If (strJournalType = "Received" And curPCash > 0 And curICash = 0) Or (strJournalType = "Misc." And curPCash > 0) Or _
            (strJournalType = "Cost Adj." And curCost > 0) Or (strJournalType = "Deposit" And Not (strComment Like "*stock split*")) Then
          ' ** Skip RECEIVED with both curICash and curPCash at this step.
5910      intRetVal = 10
5920    Else
          ' ** Sold or Received Gains/Losses.
5930      If (strJournalType = "Sold" And curGainLoss <> 0) Or (strJournalType = "Received" And strComment Like "long term capital gain") Then
5940        intRetVal = 20
5950      Else
            ' ** Disbursements of Principal.
5960        If ((strJournalType = "Paid" And curPCash <> 0 And lngTaxcode <> glngTaxCode_Distribution) Or _
                (strJournalType = "Misc." And curPCash < 0) Or (strJournalType = "Cost Adj." And curCost < 0) Or _
                (strJournalType = "Withdrawn" And lngTaxcode <> glngTaxCode_Distribution)) Then    '####  TAXCODE  ####
5970          intRetVal = 30  '<> "Distribution"
5980        Else
              ' ** Paid or Withdrawn Distributions
5990          If ((strJournalType = "Paid" And curPCash <> 0 And lngTaxcode = glngTaxCode_Distribution) Or _
                  (strJournalType = "Withdrawn" And lngTaxcode = glngTaxCode_Distribution)) Then    '####  TAXCODE  ####
6000            intRetVal = 40  '= "Distribution"
6010          Else
                ' ** Purchases.
6020            If (strJournalType = "Purchase") Then
6030              intRetVal = 50
6040            Else
                  ' ** Change in Investment Holdings.
6050              If ((strJournalType = "Sold" And curGainLoss = 0) Or _
                      (strJournalType = "Deposit" And strComment Like "*stock split*") Or (strJournalType = "Liability")) Then
6060                intRetVal = 60
6070              Else
                    ' ** Receipts of Income.
6080                If (curICash > 0 And (strJournalType = "Dividend" Or strJournalType = "Misc." Or strJournalType = "Interest" Or _
                        (strJournalType = "Received" And curPCash = 0))) Then
                      ' ** Skip RECEIVED with both curICash and curPCash at this step.
6090                  intRetVal = 70
6100                Else
                      ' ** Receipts of both Principal and Income.
6110                  If (strJournalType = "Received" And curPCash > 0 And curICash > 0) Then
                        ' ** So, it appears this one needs 2 records: one for ReportNumber = 10 and one for ReportNumber = 70.
                        ' ** HOW DO WE DO THAT HERE?
                        ' ** Start with a 10, and handle the rest above.
6120                    intRetVal = 10
6130                  Else
                        ' ** Disbursements of Income.
6140                    If ((curICash <> 0 And strJournalType = "Paid" And lngTaxcode <> glngTaxCode_Distribution) Or _
                            (strJournalType = "Liability" And curICash < 0) Or _
                            (strJournalType = "Misc." And curICash < 0)) Then    '####  TAXCODE  ####
6150                      intRetVal = 80  '<> "Distribution"
6160                    Else
                          ' ** Distributions of Income.
6170                      If (curICash <> 0) And (strJournalType = "Paid") And (lngTaxcode = glngTaxCode_Distribution) Then    '####  TAXCODE  ####
6180                        intRetVal = 90  '= "Distribution"
6190                      End If
6200                    End If
6210                  End If
6220                End If
6230              End If
6240            End If
6250          End If
6260        End If
6270      End If
6280    End If

EXITP:
6290    NSCourtReportDataID = intRetVal
6300    Exit Function

ERRH:
6310    intRetVal = 0
6320    Select Case ERR.Number
        Case Else
6330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6340    End Select
6350    Resume EXITP

End Function

Public Function NSCourtReportDate(intCourtReport As Integer, datTransDate As Date, datAssetDate As Date, Optional varJournaltype As Variant) As Date

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportDate"

        Dim datRetVal As Date

6410    Select Case intCourtReport
        Case 10
6420      datRetVal = datTransDate
6430    Case 20
6440      datRetVal = datAssetDate
6450    Case 30
6460      datRetVal = datTransDate
6470    Case 40
6480      datRetVal = datTransDate
6490    Case 50
6500      datRetVal = datTransDate  '01/13/09, per Rich: changed from datAssetDate.
6510    Case 60
          'If IsMissing(varJournaltype) = True Then
          '  datRetVal = datAssetDate
          'Else
          '  If varJournaltype <> "Liability" Then
          '    datRetVal = datAssetDate
          '  Else
6520      datRetVal = datTransDate  '01/13/09, per Rich: changed from datAssetDate.
          '  End If
          'End If
6530    Case 70
6540      datRetVal = datTransDate
6550    Case 80
6560      datRetVal = datTransDate
6570    Case 90
6580      datRetVal = datTransDate
6590    Case Else
6600      datRetVal = #1/1/1900#
6610    End Select

EXITP:
6620    NSCourtReportDate = datRetVal
6630    Exit Function

ERRH:
6640    datRetVal = #1/1/1900#
6650    Select Case ERR.Number
        Case Else
6660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6670    End Select
6680    Resume EXITP

End Function

Public Function NSCourtReportDivision(intCourtReport As Integer) As Integer

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportDivision"

        Dim intRetVal As Integer

6710    Select Case intCourtReport
        Case 5, 10, 20, 30, 40
6720      intRetVal = 20
6730    Case 50, 60
6740      intRetVal = 40
6750    Case 69, 70, 80, 90
6760      intRetVal = 60
6770    Case Else
6780      intRetVal = 0
6790    End Select

EXITP:
6800    NSCourtReportDivision = intRetVal
6810    Exit Function

ERRH:
6820    intRetVal = 0
6830    Select Case ERR.Number
        Case Else
6840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6850    End Select
6860    Resume EXITP

End Function

Public Function NSCourtReportDivisionText(intCourtReport As Integer) As String

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportDivisionText"

        Dim strRetVal As String

6910    Select Case intCourtReport
        Case 5, 10, 20, 30, 40
6920      strRetVal = "Principal on Hand at End of Period"
6930    Case 50, 60
6940      strRetVal = ""
6950    Case 69, 70, 80, 90
6960      strRetVal = "Income on Hand at End of Period"
6970    Case Else
6980      strRetVal = "Unknown"
6990    End Select

EXITP:
7000    NSCourtReportDivisionText = strRetVal
7010    Exit Function

ERRH:
7020    strRetVal = vbNullString
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Function

Public Function NSCourtReportDivisionTitle(intCourtReport As Integer) As String

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportDivisionTitle"

        Dim strRetVal As String

7110    Select Case intCourtReport
        Case 5, 10, 20, 30, 40
7120      strRetVal = "Principal"
7130    Case 50, 60
7140      strRetVal = "For Information"
7150    Case 69, 70, 80, 90
7160      strRetVal = "Income"
7170    Case Else
7180      strRetVal = "Unknown"
7190    End Select

EXITP:
7200    NSCourtReportDivisionTitle = strRetVal
7210    Exit Function

ERRH:
7220    strRetVal = vbNullString
7230    Select Case ERR.Number
        Case Else
7240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7250    End Select
7260    Resume EXITP

End Function

Public Function NSCourtReportDollars(intCourtReport As Integer, curPCash As Currency, curICash As Currency, curCost As Currency, strJournalType As String) As Double

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportDollars"

        Dim dblRetVal As Double

7310    Select Case intCourtReport
        Case 10
7320      dblRetVal = curCost + curPCash
7330    Case 20
7340      dblRetVal = curCost + curPCash
7350    Case 30
7360      If strJournalType = "Withdrawn" _
              Or strJournalType = "Cost Adj." Then
7370        dblRetVal = curCost
7380      Else
7390        dblRetVal = curPCash
7400      End If
7410    Case 40
7420      If strJournalType = "Withdrawn" Then
7430        dblRetVal = curCost
7440      Else
7450        dblRetVal = curPCash
7460      End If
7470    Case 50
7480      dblRetVal = (curPCash + curICash) * -1
7490    Case 60
7500      dblRetVal = curCost
7510    Case 70
7520      dblRetVal = curICash
7530    Case 80
7540      dblRetVal = curICash
7550    Case 90
7560      dblRetVal = curICash
7570    Case Else
7580      dblRetVal = -999999#
7590    End Select

EXITP:
7600    NSCourtReportDollars = dblRetVal
7610    Exit Function

ERRH:
7620    dblRetVal = 0#
7630    Select Case ERR.Number
        Case Else
7640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7650    End Select
7660    Resume EXITP

End Function

Public Function NSCourtReportGroup(intCourtReport As Integer) As Integer

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportGroup"

        Dim intRetVal As Integer

7710    Select Case intCourtReport
        Case 5, 10, 20
7720      intRetVal = 10
7730    Case 30, 40
7740      intRetVal = 20
7750    Case 50, 60
7760      intRetVal = 30
7770    Case 69, 70
7780      intRetVal = 40
7790    Case 80, 90
7800      intRetVal = 50
7810    Case Else
7820      intRetVal = 0
7830    End Select

EXITP:
7840    NSCourtReportGroup = intRetVal
7850    Exit Function

ERRH:
7860    intRetVal = 0
7870    Select Case ERR.Number
        Case Else
7880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7890    End Select
7900    Resume EXITP

End Function

Public Function NSCourtReportGroupText(intCourtReport As Integer) As String

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportGroupText"

        Dim strRetVal As String

8010    Select Case intCourtReport
        Case 5, 10, 20
8020      strRetVal = vbNullString
8030    Case 30, 40
8040      strRetVal = vbNullString
8050    Case 50, 60
8060      strRetVal = vbNullString
8070    Case 69, 70
8080      strRetVal = vbNullString
8090    Case 80, 90
8100      strRetVal = vbNullString
8110    Case Else
8120      strRetVal = "Unknown"
8130    End Select

EXITP:
8140    NSCourtReportGroupText = strRetVal
8150    Exit Function

ERRH:
8160    strRetVal = vbNullString
8170    Select Case ERR.Number
        Case Else
8180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8190    End Select
8200    Resume EXITP

End Function

Public Function NSGetCourtReportData() As Integer
' ** Return Codes:
' **   0  Success.
' **  -1  Canceled.
' **  -9  Error.

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "NSGetCourtReportData"

        'Dim rsxDataOut As ADODB.Recordset  ' ** Early binding.
        Dim rsxDataOut As Object            ' ** Late binding.
        Dim frm As Access.Form
        Dim dblCashAssets_Beg As Double
        Dim dblNetIncome As Double
        Dim blnShowForm As Boolean
        Dim intRetVal As Integer

8310    intRetVal = 0

8320    DoCmd.Hourglass False

        ' ** Leave these as they are.
        ' **   gstrCrtRpt_Ordinal
        ' **   gstrCrtRpt_Version

8330    blnShowForm = True
8340    Set frm = Forms("frmRpt_CourtReports_NS")
8350    If gblnPrintAll = True Then
8360      If IsNull(frm.Ordinal) = False And IsNull(frm.Version) = False And _
              IsNull(frm.CashAssets_Beg) = False And IsNull(frm.NetIncome) = False Then
8370        blnShowForm = False
8380        If gstrCrtRpt_Ordinal = vbNullString Then gstrCrtRpt_Ordinal = frm.Ordinal
8390        If gstrCrtRpt_Version = vbNullString Then gstrCrtRpt_Version = frm.Version
8400        If gstrCrtRpt_CashAssets_Beg = vbNullString Then gstrCrtRpt_CashAssets_Beg = frm.CashAssets_Beg
8410        If gstrCrtRpt_NetIncome = vbNullString Then gstrCrtRpt_NetIncome = frm.NetIncome
8420      End If
8430    End If

8440    If blnShowForm = True Then
8450      gblnMessage = True  ' ** If this returns False, the dialog was canceled.
8460      DoCmd.OpenForm "frmRpt_CourtReports_NS_Input", , , , , acDialog, "frmRpt_CourtReports_NS"
8470      If gblnMessage = False Then
8480        intRetVal = -1  ' ** Canceled.
8490      Else
8500        frm.Ordinal = gstrCrtRpt_Ordinal
8510        frm.Version = gstrCrtRpt_Version
8520        frm.CashAssets_Beg = gstrCrtRpt_CashAssets_Beg
8530        frm.NetIncome = gstrCrtRpt_NetIncome
8540      End If
8550    End If  ' ** blnShowForm.

8560    If intRetVal = 0 Then

8570      DoCmd.Hourglass True

8580      dblCashAssets_Beg = CDbl(gstrCrtRpt_CashAssets_Beg)
8590      dblNetIncome = CDbl(gstrCrtRpt_NetIncome)

          'Set rsxDataOut = New ADODB.Recordset             ' ** Early binding.
8600      Set rsxDataOut = CreateObject("ADODB.Recordset")  ' ** Late binding.
8610      rsxDataOut.Open "tmpCourtReportData", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTable

          ' ** Get the beginning pricinpal/proerty balance for the report.
8620      With rsxDataOut
8630        .AddNew
8640        .Fields("ReportNumber") = 5
8650        .Fields("ReportCategory") = NSCourtReportCategory(5)
8660        .Fields("ReportGroup") = NSCourtReportGroup(5)
8670        .Fields("ReportDivision") = NSCourtReportDivision(5)
8680        .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(5)
8690        .Fields("ReportDivisionText") = NSCourtReportDivisionText(5)
8700        .Fields("ReportGroupText") = NSCourtReportGroupText(5)
8710        .Fields("accountno") = gstrAccountNo
8720        .Fields("date") = gdatStartDate
8730        .Fields("journaltype") = "Entered"
8740        .Fields("Amount") = dblCashAssets_Beg
8750        .Fields("revcode_ID") = 0
8760        .Fields("revcode_DESC") = "Dummy entry"
8770        .Fields("revcode_TYPE") = 1
8780        .Fields("revcode_SORTORDER") = 0
8790        .Update
8800      End With

          ' ** Get the beginning income balance for the report.
8810      With rsxDataOut
8820        .AddNew
8830        .Fields("ReportNumber") = 69
8840        .Fields("ReportCategory") = NSCourtReportCategory(69)
8850        .Fields("ReportGroup") = NSCourtReportGroup(69)
8860        .Fields("ReportDivision") = NSCourtReportDivision(69)
8870        .Fields("ReportDivisionTitle") = NSCourtReportDivisionTitle(69)
8880        .Fields("ReportDivisionText") = NSCourtReportDivisionText(69)
8890        .Fields("ReportGroupText") = NSCourtReportGroupText(69)
8900        .Fields("accountno") = gstrAccountNo
8910        .Fields("date") = gdatStartDate
8920        .Fields("journaltype") = "Entered"
8930        .Fields("Amount") = dblNetIncome
8940        .Fields("revcode_ID") = 0
8950        .Fields("revcode_DESC") = "Dummy entry"
8960        .Fields("revcode_TYPE") = 1
8970        .Fields("revcode_SORTORDER") = 0
8980        .Update
8990      End With

          'End If
9000    End If

9010    DoCmd.Hourglass False

EXITP:
9020    Set frm = Nothing
9030    Set rsxDataOut = Nothing
9040    NSGetCourtReportData = intRetVal
9050    Exit Function

ERRH:
9060    DoCmd.Hourglass False
9070    intRetVal = -9  ' ** Error.
9080    Select Case ERR.Number
        Case Else
9090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9100    End Select
9110    Resume EXITP

End Function

Public Function NSNum(ByVal strConst As String) As Integer

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "NSNum"

        Dim lngX As Long
        Dim intRetVal As Integer

9210    If lngNSRpts = 0& Or IsEmpty(arr_varNSRpt) = True Then
9220      NSCourtReportLoad  ' ** Function: Below.
9230    End If

9240    For lngX = 0& To (lngNSRpts - 1&)
9250      If arr_varNSRpt(CR_CON, lngX) = strConst Then
9260        intRetVal = arr_varNSRpt(CR_NUM, lngX)
9270        Exit For
9280      End If
9290    Next

EXITP:
9300    NSNum = intRetVal
9310    Exit Function

ERRH:
9320    intRetVal = 0
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Function

Public Function NSCourtReportLoad() As Boolean
' ** VGC 05/23/2011: NO CHANGES WERE MADE TO THIS AFTER COPYING FROM FLORIDA!

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "NSCourtReportLoad"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngX As Long
        Dim blnRetVal As Boolean

9410    blnRetVal = True

9420    Set dbs = CurrentDb
9430    Set qdf = dbs.QueryDefs("qryCourtReport_NS_20")
9440    Set rst = qdf.OpenRecordset
9450    With rst
9460      .MoveLast
9470      lngNSRpts = .RecordCount
9480      .MoveFirst
9490      arr_varNSRpt = .GetRows(lngNSRpts)
          ' *******************************************************
          ' ** Array: arr_varNSRpt()
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
9500      .Close
9510    End With
9520    dbs.Close

        ' ** Set Null text fields to vbNullString to simplify comparisons.
9530    For lngX = 0& To (lngNSRpts - 1&)
9540      If IsEmpty(arr_varNSRpt(CR_GRPTXT, lngX)) = True Then
9550        arr_varNSRpt(CR_GRPTXT, lngX) = vbNullString
9560      ElseIf IsNull(arr_varNSRpt(CR_GRPTXT, lngX)) = True Then
9570        arr_varNSRpt(CR_GRPTXT, lngX) = vbNullString
9580      End If
9590      If IsEmpty(arr_varNSRpt(CR_SCHED, lngX)) = True Then
9600        arr_varNSRpt(CR_SCHED, lngX) = vbNullString
9610      ElseIf IsNull(arr_varNSRpt(CR_SCHED, lngX)) = True Then
9620        arr_varNSRpt(CR_SCHED, lngX) = vbNullString
9630      End If
9640    Next

        ' ** Change Schedule letters and/or date source.
9650    For lngX = 0& To (lngNSRpts - 1&)
9660      Select Case arr_varNSRpt(CR_CON, lngX)
          Case "CRPT_RECEIPTS"
9670        arr_varNSRpt(CR_SCHED, lngX) = "A"  ' ** No change.
9680      Case "CRPT_DISBURSEMENTS"
9690        arr_varNSRpt(CR_SCHED, lngX) = "B"
9700      Case "CRPT_DISTRIBUTIONS"
9710        arr_varNSRpt(CR_SCHED, lngX) = "C"
9720      Case "CRPT_GAINS"
9730        arr_varNSRpt(CR_SCHED, lngX) = "D"
9740        arr_varNSRpt(CR_DATE, lngX) = "AssetDate"  ' ** VGC 10/12/2008: Change to match NS version, which uses assetdate.
9750      Case "CRPT_ON_HAND_ENDL", "CRPT_Cash_END", "CRPT_NON_Cash_END", "CRPT_ON_HAND_END"
9760        arr_varNSRpt(CR_SCHED, lngX) = "E"
9770      Case "CRPT_LOSSES"
9780        arr_varNSRpt(CR_SCHED, lngX) = vbNullString
9790      End Select
9800    Next

9810    For lngX = 0& To (lngNSRpts - 1&)

9820      If arr_varNSRpt(CR_DIVTTL, lngX) = "CHARGES" And CRPT_DIV_CHARGES = 0 Then
9830        CRPT_DIV_CHARGES = arr_varNSRpt(CR_DIV, lngX)  ' ** 20
9840      ElseIf arr_varNSRpt(CR_DIVTTL, lngX) = "CREDITS" And CRPT_DIV_CREDITS = 0 Then
9850        CRPT_DIV_CREDITS = arr_varNSRpt(CR_DIV, lngX)  ' ** 40
9860      ElseIf arr_varNSRpt(CR_DIVTTL, lngX) = "ADDITIONAL INFORMATION" And CRPT_DIV_ADDL = 0 Then
9870        CRPT_DIV_ADDL = arr_varNSRpt(CR_DIV, lngX)     ' ** 60
9880      End If

9890      If arr_varNSRpt(CR_CON, lngX) = "CRPT_ON_HAND_BEGL" Then       '      2
9900        CRPT_ON_HAND_BEGL = arr_varNSRpt(CR_NUM, lngX)
9910      ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_Cash_BEG" Then       '      3
9920        CRPT_CASH_BEG = arr_varNSRpt(CR_NUM, lngX)
9930      ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_NON_Cash_BEG" Then   '      4
9940        CRPT_NON_CASH_BEG = arr_varNSRpt(CR_NUM, lngX)
9950      ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_ON_HAND_BEG" Then    '5     5
9960        CRPT_ON_HAND_BEG = arr_varNSRpt(CR_NUM, lngX)
9970      ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_ADDL_PROP" Then      '10   10
9980        CRPT_ADDL_PROP = arr_varNSRpt(CR_NUM, lngX)
9990      ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_RECEIPTS" Then       '20   20
10000       CRPT_RECEIPTS = arr_varNSRpt(CR_NUM, lngX)
10010     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_GAINS" Then          '30   30
10020       CRPT_GAINS = arr_varNSRpt(CR_NUM, lngX)
10030     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_OTH_CHG" Then        '     40
10040       CRPT_OTH_CHG = arr_varNSRpt(CR_NUM, lngX)
10050     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_NET_INCOME" Then     '40   50
10060       CRPT_NET_INCOME = arr_varNSRpt(CR_NUM, lngX)
10070     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_DISBURSEMENTS" Then  '50   60
10080       CRPT_DISBURSEMENTS = arr_varNSRpt(CR_NUM, lngX)
10090     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_LOSSES" Then         '60   70
10100       CRPT_LOSSES = arr_varNSRpt(CR_NUM, lngX)
10110     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_DISTRIBUTIONS" Then  '80   80
10120       CRPT_DISTRIBUTIONS = arr_varNSRpt(CR_NUM, lngX)
10130     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_OTH_CRED" Then       '     90
10140       CRPT_OTH_CRED = arr_varNSRpt(CR_NUM, lngX)
10150     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_NET_LOSS" Then       '70  100
10160       CRPT_NET_LOSS = arr_varNSRpt(CR_NUM, lngX)
10170     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_ON_HAND_ENDL" Then   '    107
10180       CRPT_ON_HAND_ENDL = arr_varNSRpt(CR_NUM, lngX)
10190     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_Cash_END" Then       '    108
10200       CRPT_CASH_END = arr_varNSRpt(CR_NUM, lngX)
10210     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_NON_Cash_END" Then   '    109
10220       CRPT_NON_CASH_END = arr_varNSRpt(CR_NUM, lngX)
10230     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_ON_HAND_END" Then    '90  110
10240       CRPT_ON_HAND_END = arr_varNSRpt(CR_NUM, lngX)
10250     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_INVEST_INFO" Then    '100 120
10260       CRPT_INVEST_INFO = arr_varNSRpt(CR_NUM, lngX)
10270     ElseIf arr_varNSRpt(CR_CON, lngX) = "CRPT_CHANGES" Then        '110 130
10280       CRPT_CHANGES = arr_varNSRpt(CR_NUM, lngX)
10290     End If

10300   Next

EXITP:
10310   Set rst = Nothing
10320   Set qdf = Nothing
10330   Set dbs = Nothing
10340   NSCourtReportLoad = blnRetVal
10350   Exit Function

ERRH:
10360   blnRetVal = False
10370   Select Case ERR.Number
        Case Else
10380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10390   End Select
10400   Resume EXITP

End Function

Public Sub WordAll_NS(frm As Access.Form)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "WordAll_NS"

        Dim strDocName As String
        Dim blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String
        Dim lngX As Long

        ' ** Access: 14272506  Very Light Red
        ' ** Access: 12295153  Medium Red
        ' ** Word:   16770233  Very Light Blue
        ' ** Word:   16434048  Medium Blue
        ' ** Excel:  14677736  Very Light Green
        ' ** Excel:  5952646   Medium Green

10510   With frm
10520     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NS.

10530       DoCmd.Hourglass True
10540       DoEvents

10550       .cmdWordAll_box01.Visible = True
10560       .cmdWordAll_box02.Visible = True
10570       If .chkAssetList = True Then
10580         .cmdWordAll_box03.Visible = True
10590       End If
10600       .cmdWordAll_box04.Visible = True
10610       DoEvents

10620       blnExcel = False
10630       blnAutoStart = .chkOpenWord
10640       Beep
10650       DoCmd.Hourglass False
10660       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Word" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

10670       If msgResponse = vbOK Then

10680         DoCmd.Hourglass True
10690         DoEvents

10700         blnAllCancel = False
10710         .AllCancelSet11_NS blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NS.
10720         gblnPrintAll = True
10730         strThisProc = "cmdWordAll_Click"

              ' ** NSBuildCourtReportData() calls NSGetCourtReportData(), which opens frmRpt_CourtReports_NS_Input.
              ' ** The input form, frmRpt_CourtReports_NS_Input, populates these variables.
              ' **   gstrCrtRpt_Ordinal
              ' **   gstrCrtRpt_Version
              ' **   gstrCrtRpt_CashAssets_Beg
              ' **   gstrCrtRpt_NetIncome

              ' ** Get the Summary inputs first.
10740         DoCmd.Hourglass False
10750         gblnMessage = True
10760         DoCmd.OpenForm "frmRpt_CourtReports_NS_Input", , , , , acDialog, "frmRpt_CourtReports_NS"
10770         If gblnMessage = False Then
10780           blnAllCancel = True
10790           .AllCancelSet11_NS blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NS.
10800         Else
10810           frm.Ordinal = gstrCrtRpt_Ordinal
10820           frm.Version = gstrCrtRpt_Version
10830           frm.CashAssets_Beg = gstrCrtRpt_CashAssets_Beg
10840           frm.NetIncome = gstrCrtRpt_NetIncome
10850         End If
10860         DoEvents

10870         If blnAllCancel = False Then

10880           DoCmd.Hourglass True
10890           DoEvents

10900           If blnAllCancel = False Then
                  ' ** Summary of Account.
10910             .cmdWord00.SetFocus
10920             DoEvents
                  ' ** This also calls:
                  ' ** Property on Hand at Beginning of Account.
10930             .cmdWord00_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
10940             DoEvents
10950           End If
10960           If blnAllCancel = False Then
                  ' ** Receipts of Principal.
10970             .cmdWord01.SetFocus
10980             DoEvents
10990             .cmdWord01_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11000             DoEvents
11010           End If
11020           If blnAllCancel = False Then
                  ' ** Gain(Loss) On Sale or Disposition.
11030             .cmdWord02.SetFocus
11040             DoEvents
11050             .cmdWord02_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11060             DoEvents
11070           End If
11080           If blnAllCancel = False Then
                  ' ** Disbursements of Principal.
11090             .cmdWord03.SetFocus
11100             DoEvents
11110             .cmdWord03_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11120             DoEvents
11130           End If
11140           If blnAllCancel = False Then
                  ' ** Distribution of Principal to Beneficiaries.
11150             .cmdWord04.SetFocus
11160             DoEvents
11170             .cmdWord04_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11180             DoEvents
11190           End If
11200           If blnAllCancel = False Then
                  ' ** Information for Investments Made.
11210             .cmdWord05.SetFocus
11220             DoEvents
11230             .cmdWord05_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11240             DoEvents
11250           End If
11260           If blnAllCancel = False Then
                  ' ** Change in Investment Holdings.
11270             .cmdWord06.SetFocus
11280             DoEvents
11290             .cmdWord06_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11300             DoEvents
11310           End If
11320           If blnAllCancel = False Then
                  ' ** Receipts of Income.
11330             .cmdWord07.SetFocus
11340             DoEvents
11350             .cmdWord07_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11360             DoEvents
11370           End If
11380           If blnAllCancel = False Then
                  ' ** Disbursements of Income.
11390             .cmdWord08.SetFocus
11400             DoEvents
11410             .cmdWord08_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11420             DoEvents
11430           End If
11440           If blnAllCancel = False Then
                  ' ** Distributions of Income.
11450             .cmdWord09.SetFocus
11460             DoEvents
11470             .cmdWord09_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11480             DoEvents
11490           End If
11500           If blnAllCancel = False Then
                  ' ** Property on Hand at Close of Account.
11510             .cmdWord10.SetFocus
11520             DoEvents
11530             .cmdWord10_Click  ' ** Form Procedure:  frmRpt_CourtReports_NS.
11540             DoEvents
11550           End If

11560           DoCmd.Hourglass True
11570           DoEvents

11580           .cmdWordAll.SetFocus

11590           gblnPrintAll = False
11600           Beep

11610           If lngFiles > 0& Then

11620             DoCmd.Hourglass False

11630             strTmp01 = CStr(lngFiles) & " documents were created."
11640             If .chkOpenWord = True Then
11650               strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
11660               msgResponse = MsgBox(strTmp01, vbInformation + vbOKCancel, "Reports Exported")
11670             Else
11680               msgResponse = MsgBox(strTmp01, vbInformation + vbOKOnly, "Reports Exported")
11690             End If

11700             .cmdWordAll_box01.Visible = False
11710             .cmdWordAll_box02.Visible = False
11720             .cmdWordAll_box03.Visible = False
11730             .cmdWordAll_box04.Visible = False

11740             If .chkOpenWord = True And msgResponse = vbOK Then
11750               DoCmd.Hourglass True
11760               DoEvents
11770               For lngX = 0& To (lngFiles - 1&)
11780                 strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
11790                 OpenExe strDocName  ' ** Module Function: modShellFuncs.
11800                 DoEvents
11810                 If lngX < (lngFiles - 1&) Then
11820                   ForcePause 2  ' ** Module Function: modCodeUtilities.
11830                 End If
11840               Next
11850               Beep
11860             End If

11870           Else
11880             DoCmd.Hourglass False
11890             MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
11900             .cmdWordAll_box01.Visible = False
11910             .cmdWordAll_box02.Visible = False
11920             .cmdWordAll_box03.Visible = False
11930             .cmdWordAll_box04.Visible = False
11940           End If  ' ** lngFiles.

11950         Else
11960           .cmdWordAll_box01.Visible = False
11970           .cmdWordAll_box02.Visible = False
11980           .cmdWordAll_box03.Visible = False
11990           .cmdWordAll_box04.Visible = False
12000           gblnPrintAll = False
12010         End If  ' ** blnAllCancel.

12020       Else
12030         .cmdWordAll_box01.Visible = False
12040         .cmdWordAll_box02.Visible = False
12050         .cmdWordAll_box03.Visible = False
12060         .cmdWordAll_box04.Visible = False
12070         gblnPrintAll = False
12080       End If  ' ** msgResponse.

12090       DoCmd.Hourglass False
12100     End If  ' ** Validate.
12110   End With

EXITP:
12120   Set frm = Nothing
12130   Exit Sub

ERRH:
12140   gblnPrintAll = False
12150   DoCmd.Hourglass False
12160   Select Case ERR.Number
        Case Else
12170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12180   End Select
12190   Resume EXITP

End Sub

Public Sub ExcelAll_NS(frm As Access.Form)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "ExcelAll_NS"

        Dim strDocName As String
        Dim blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String
        Dim lngX As Long

        ' ** Access: 14272506  Very Light Red
        ' ** Access: 12295153  Medium Red
        ' ** Word:   16770233  Very Light Blue
        ' ** Word:   16434048  Medium Blue
        ' ** Excel:  14677736  Very Light Green
        ' ** Excel:  5952646   Medium Green

12210   With frm
12220     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NS.

12230       .cmdExcelAll_box01.Visible = True
12240       .cmdExcelAll_box02.Visible = True
12250       If .chkAssetList = True Then
12260         .cmdExcelAll_box03.Visible = True
12270       End If
12280       .cmdExcelAll_box04.Visible = True
12290       DoEvents

12300       blnExcel = True
12310       blnAutoStart = .chkOpenExcel
12320       Beep
12330       DoCmd.Hourglass False
12340       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Excel" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

12350       If msgResponse = vbOK Then

12360         DoCmd.Hourglass True
12370         DoEvents

12380         gblnPrintAll = True
12390         blnAllCancel = False
12400         .AllCancelSet11_NS blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NS.
12410         strThisProc = "cmdExcelAll_Click"

12420         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12430           DoCmd.Hourglass False
12440           msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                  "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                  "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                  "You may close Excel before proceding, then click Retry." & vbCrLf & _
                  "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
                ' ** ... Otherwise Trust Accountant will do it for you.
12450           If msgResponse <> vbRetry Then
12460             blnAllCancel = True
12470             .AllCancelSet11_NS blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NS.
12480           End If
12490         End If

12500         If blnAllCancel = False Then

                ' ** NSBuildCourtReportData() calls NSGetCourtReportData(), which opens frmRpt_CourtReports_NS_Input.
                ' ** The input form, frmRpt_CourtReports_NS_Input, populates these variables.
                ' **   gstrCrtRpt_Ordinal
                ' **   gstrCrtRpt_Version
                ' **   gstrCrtRpt_CashAssets_Beg
                ' **   gstrCrtRpt_NetIncome

                ' ** Get the Summary inputs first.
12510           DoCmd.Hourglass False
12520           gblnMessage = True
12530           DoCmd.OpenForm "frmRpt_CourtReports_NS_Input", , , , , acDialog, "frmRpt_CourtReports_NS"
12540           If gblnMessage = False Then
12550             blnAllCancel = True
12560             .AllCancelSet11_NS blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NS.
12570           Else
12580             frm.Ordinal = gstrCrtRpt_Ordinal
12590             frm.Version = gstrCrtRpt_Version
12600             frm.CashAssets_Beg = gstrCrtRpt_CashAssets_Beg
12610             frm.NetIncome = gstrCrtRpt_NetIncome
12620           End If
12630           DoEvents

12640           If blnAllCancel = False Then

12650             DoCmd.Hourglass True
12660             DoEvents

12670             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12680               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
12690             End If
12700             DoEvents

12710             If blnAllCancel = False Then
                    ' ** Summary of Account.
12720               .cmdExcel00.SetFocus
12730               DoEvents
                    ' ** This also calls:
                    ' ** Property on Hand at Beginning of Account.
12740               .cmdExcel00_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
12750               DoEvents
12760             End If
12770             If blnAllCancel = False Then
                    ' ** Additional Property Received.
12780               .cmdExcel01.SetFocus
12790               DoEvents
12800               .cmdExcel01_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
12810               DoEvents
12820             End If
12830             If blnAllCancel = False Then
                    ' ** Receipts.
12840               .cmdExcel02.SetFocus
12850               DoEvents
12860               .cmdExcel02_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
12870               DoEvents
12880             End If
12890             If blnAllCancel = False Then
                    ' ** Gains on Sales.
12900               .cmdExcel03.SetFocus
12910               DoEvents
12920               .cmdExcel03_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
12930               DoEvents
12940             End If
12950             If blnAllCancel = False Then
                    ' ** Disbursements.
12960               .cmdExcel04.SetFocus
12970               DoEvents
12980               .cmdExcel04_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
12990               DoEvents
13000             End If
13010             If blnAllCancel = False Then
                    ' ** Losses on Sales.
13020               .cmdExcel05.SetFocus
13030               DoEvents
13040               .cmdExcel05_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
13050               DoEvents
13060             End If
13070             If blnAllCancel = False Then
                    ' ** Distributions.
13080               .cmdExcel06.SetFocus
13090               DoEvents
13100               .cmdExcel06_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
13110               DoEvents
13120             End If
13130             If blnAllCancel = False Then
                    ' ** Property on Hand at Close of Account.
13140               .cmdExcel07.SetFocus
13150               DoEvents
13160               .cmdExcel07_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
13170               DoEvents
13180             End If
13190             If blnAllCancel = False Then
                    ' ** Information for Investments Made.
13200               .cmdExcel08.SetFocus
13210               DoEvents
13220               .cmdExcel08_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
13230               DoEvents
13240             End If
13250             If blnAllCancel = False Then
                    ' ** Change in Investment Holdings.
13260               .cmdExcel09.SetFocus
13270               DoEvents
13280               .cmdExcel09_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
13290               DoEvents
13300             End If
13310             If blnAllCancel = False Then
                    ' ** Other Charges.
13320               .cmdExcel10.SetFocus
13330               DoEvents
13340               .cmdExcel10_Click  ' ** Form Procedure: frmRpt_CourtReports_NS.
13350               DoEvents
13360             End If

13370             DoCmd.Hourglass True
13380             DoEvents

13390             .cmdExcelAll.SetFocus

13400             gblnPrintAll = False
13410             Beep

13420             If lngFiles > 0& Then

13430               DoCmd.Hourglass False

13440               strTmp01 = CStr(lngFiles) & " documents were created."
13450               If .chkOpenExcel = True Then
13460                 strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
13470               End If

13480               MsgBox strTmp01, vbInformation + vbOKOnly, "Reports Exported"

13490               .cmdExcelAll_box01.Visible = False
13500               .cmdExcelAll_box02.Visible = False
13510               .cmdExcelAll_box03.Visible = False
13520               .cmdExcelAll_box04.Visible = False

13530               If .chkOpenExcel = True Then
13540                 DoCmd.Hourglass True
13550                 DoEvents
13560                 For lngX = 0& To (lngFiles - 1&)
13570                   strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
13580                   OpenExe strDocName  ' ** Module Function: modShellFuncs.
13590                   DoEvents
13600                   If lngX < (lngFiles - 1&) Then
13610                     ForcePause 2  ' ** Module Function: modCodeUtilities.
13620                   End If
13630                 Next
13640               End If

13650             Else
13660               DoCmd.Hourglass False
13670               MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
13680               .cmdExcelAll_box01.Visible = False
13690               .cmdExcelAll_box02.Visible = False
13700               .cmdExcelAll_box03.Visible = False
13710               .cmdExcelAll_box04.Visible = False
13720             End If    ' ** lngFiles.

13730           End If  ' ** blnAllCancel.

13740         Else
13750           .cmdWordAll_box01.Visible = False
13760           .cmdWordAll_box02.Visible = False
13770           .cmdWordAll_box03.Visible = False
13780           .cmdWordAll_box04.Visible = False
13790           gblnPrintAll = False
13800         End If  ' ** blnAllCancel.

13810         DoCmd.Hourglass False

13820       Else
13830         .cmdWordAll_box01.Visible = False
13840         .cmdWordAll_box02.Visible = False
13850         .cmdWordAll_box03.Visible = False
13860         .cmdWordAll_box04.Visible = False
13870         gblnPrintAll = False
13880       End If  ' ** msgResponse.

13890     End If  ' ** Validate.
13900   End With

EXITP:
13910   Exit Sub

ERRH:
13920   gblnPrintAll = False
13930   DoCmd.Hourglass False
13940   Select Case ERR.Number
        Case Else
13950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13960   End Select
13970   Resume EXITP

End Sub

Public Sub FileArraySet_NS(arr_varTmp00 As Variant)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "FileArraySet_NS"

14010   arr_varFile = arr_varTmp00
14020   lngFiles = UBound(arr_varFile, 2) + 1&

EXITP:
14030   Exit Sub

ERRH:
14040   Select Case ERR.Number
        Case Else
14050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14060   End Select
14070   Resume EXITP

End Sub

Public Sub AllCancelSet2_NS(blnCancel As Boolean)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet2_NS"

14110   blnAllCancel = blnCancel

EXITP:
14120   Exit Sub

ERRH:
14130   Select Case ERR.Number
        Case Else
14140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14150   End Select
14160   Resume EXITP

End Sub

Public Sub SetUserReportPath_NS(frm As Access.Form)

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "SetUserReportPath_NS"

        Dim blnEnable As Boolean

14210   With frm
14220     blnEnable = True
14230     Select Case IsNull(.UserReportPath)
          Case True
14240       blnEnable = False
14250     Case False
14260       If Trim(.UserReportPath) = vbNullString Then
14270         blnEnable = False
14280       End If
14290     End Select
14300     Select Case blnEnable
          Case True
14310       .UserReportPath.BorderColor = CLR_LTBLU2
14320       .UserReportPath.BackStyle = acBackStyleNormal
14330       .UserReportPath.Enabled = True  ' ** It remains locked.
14340       .UserReportPath_chk.Enabled = True
14350       .UserReportPath_chk.Locked = False
14360       .UserReportPath_chk_lbl1.Visible = True
14370       .UserReportPath_chk_lbl1_dim.Visible = False
14380       .UserReportPath_chk_lbl1_dim_hi.Visible = False
14390       .UserReportPath_chk_lbl2.Visible = True
14400       .UserReportPath_chk_lbl2_dim.Visible = False
14410       .UserReportPath_chk_lbl2_dim_hi.Visible = False
14420     Case False
14430       .UserReportPath = vbNullString
14440       .UserReportPath.BorderColor = WIN_CLR_DISR
14450       .UserReportPath.BackStyle = acBackStyleTransparent
14460       .UserReportPath.Enabled = False
14470       .UserReportPath_chk.Enabled = False
14480       .UserReportPath_chk.Locked = False
14490       .UserReportPath_chk_lbl1.Visible = False
14500       .UserReportPath_chk_lbl1_dim.Visible = True
14510       .UserReportPath_chk_lbl1_dim_hi.Visible = True
14520       .UserReportPath_chk_lbl2.Visible = False
14530       .UserReportPath_chk_lbl2_dim.Visible = True
14540       .UserReportPath_chk_lbl2_dim_hi.Visible = True
14550     End Select
14560   End With

EXITP:
14570   Exit Sub

ERRH:
14580   DoCmd.Hourglass False
14590   Select Case ERR.Number
        Case Else
14600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14610   End Select
14620   Resume EXITP

End Sub

Public Sub AssetList_Word_NS(strRptPeriod As String, frm As Access.Form)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetList_Word_NS"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRptSuffix As String, strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim blnExcel As Boolean
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngX As Long, lngE As Long
        Dim intRetVal_BuildAssetListInfo As Integer

14710   blnContinue = True
14720   blnUseSavedPath = False

14730   With frm

14740     DoCmd.Hourglass True
14750     DoEvents

14760     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NS.

14770       gstrFormQuerySpec = .Name
14780       blnExcel = False

14790       strRptSuffix = vbNullString
14800       Select Case strRptPeriod
            Case "Beginning"
14810         intRetVal_BuildAssetListInfo = .BuildAssetListInfo_NS("01/01/1900", (.DateEnd - 1), strRptPeriod, strRptSuffix, THIS_PROC)  ' ** Form Function: frmRpt_CourtReports_NS.
14820       Case "Ending"
14830         intRetVal_BuildAssetListInfo = .BuildAssetListInfo_NS(.DateStart, .DateEnd, strRptPeriod, strRptSuffix, THIS_PROC)  ' ** Form Function: frmRpt_CourtReports_NS.
14840       End Select
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry, e.g., date.
            ' **   -9  Error.

14850       If intRetVal_BuildAssetListInfo = -2 Then
14860         Set dbs = CurrentDb
14870         With dbs
                ' ** Empty tmpAssetList2.
14880           Set qdf = dbs.QueryDefs("qryCourtReport_03")
14890           qdf.Execute
14900           .Close
14910         End With
14920       ElseIf intRetVal_BuildAssetListInfo < 0 Then
14930         blnContinue = False
14940       End If

14950       If blnContinue = True Then

              ' ** strRptSuffix should return among "_00B", "_00BA", "_00D", or "_00DA".
14960         If strRptSuffix <> vbNullString Then

14970           gdatStartDate = .DateStart
14980           gdatEndDate = .DateEnd
14990           gstrAccountNo = .cmbAccounts.Column(0)
15000           gstrAccountName = .cmbAccounts.Column(3)

15010           lngCaps = 0&
15020           arr_varCap = Empty

15030           Set dbs = CurrentDb
15040           With dbs
                  ' ** tblReport, captions of Court Reports, by specified [CrtTyp].
15050             Set qdf = .QueryDefs("qryCourtReport_15")
15060             With qdf.Parameters
15070               ![CrtTyp] = "NS"
15080             End With
15090             Set rst = qdf.OpenRecordset
15100             With rst
15110               .MoveLast
15120               lngCaps = .RecordCount
15130               .MoveFirst
15140               arr_varCap = .GetRows(lngCaps)
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
15150               .Close
15160             End With
15170             .Close
15180           End With

15190           If IsNull(.UserReportPath) = False Then
15200             If .UserReportPath <> vbNullString Then
15210               If .UserReportPath_chk = True Then
15220                 If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
15230                   blnUseSavedPath = True
15240                 End If
15250               End If
15260             End If
15270           End If

15280           strRptCap = vbNullString: strRptPathFile = vbNullString
15290           strRptPath = .UserReportPath
15300           strRptName = "rptCourtRptNS" & strRptSuffix

15310           For lngX = 0& To (lngCaps - 1&)
15320             If arr_varCap(C_RNAM, lngX) = strRptName Then
15330               strRptCap = arr_varCap(C_CAPN, lngX)
15340               Exit For
15350             End If
15360           Next

15370           Select Case blnUseSavedPath
                Case True
15380             strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
15390           Case False
15400             DoCmd.Hourglass False
15410             strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
15420           End Select

15430           If strRptPathFile <> vbNullString Then
15440             DoCmd.Hourglass True
15450             DoEvents
15460             If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
15470               Kill strRptPathFile
15480             End If
15490             Select Case blnExcel
                  Case True
15500               blnAutoStart = .chkOpenExcel
15510             Case False
15520               blnAutoStart = .chkOpenWord
15530             End Select
15540             If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
15550             Select Case gblnPrintAll
                  Case True
15560               lngFiles = lngFiles + 1&
15570               lngE = lngFiles - 1&
15580               ReDim Preserve arr_varFile(F_ELEMS, lngE)
15590               arr_varFile(F_RNAM, lngE) = strRptName
15600               arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
15610               arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
15620               FileArraySet_NS arr_varFile  ' ** Procedure: Above.
15630               DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
15640             Case False
15650               DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
15660             End Select
                  'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
15670             strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
15680             If strRptPath <> .UserReportPath Then
15690               .UserReportPath = strRptPath
15700               SetUserReportPath_NS frm  ' ** Procedure: Above.
15710             End If
15720           Else
15730             blnContinue = False
15740           End If

15750         End If  ' ** strRptSuffix.

15760       Else
15770         blnContinue = False
15780         DoCmd.Hourglass False
15790         MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
15800       End If  ' ** blnContinue.

15810     End If  ' ** Validate.

15820     DoCmd.Hourglass False

15830   End With

EXITP:
15840   Set rst = Nothing
15850   Set qdf = Nothing
15860   Set dbs = Nothing
15870   Exit Sub

ERRH:
15880   DoCmd.Hourglass False
15890   Select Case ERR.Number
        Case Else
15900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15910   End Select
15920   Resume EXITP

End Sub

Public Sub AssetList_Excel_NS(strRptPeriod As String, lngCaps As Long, arr_varCap As Variant, frm As Access.Form)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetList_Excel_NS"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strQry As String, strMacro As String
        Dim strRptSuffix As String, strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim strLastAssetType As String
        Dim blnNoData As Boolean, blnExcel As Boolean
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngRecs As Long
        Dim lngX As Long, lngE As Long
        Dim intRetVal_BuildAssetListInfo As Integer

16010   blnContinue = True
16020   blnUseSavedPath = False

16030   With frm

16040     DoCmd.Hourglass True
16050     DoEvents

16060     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NS.

16070       gstrFormQuerySpec = .Name
16080       blnNoData = False
16090       strRptSuffix = vbNullString
16100       blnExcel = True

16110       Select Case strRptPeriod
            Case "Beginning"
16120         intRetVal_BuildAssetListInfo = .BuildAssetListInfo_NS("01/01/1900", (.DateStart - 1), strRptPeriod, strRptSuffix, THIS_PROC)  ' ** Form Function: frmRpt_CourtReports_NS.
16130       Case "Ending"
16140         intRetVal_BuildAssetListInfo = .BuildAssetListInfo_NS(.DateStart, .DateEnd, strRptPeriod, strRptSuffix, THIS_PROC)  ' ** Form Function: frmRpt_CourtReports_NS.
16150       End Select
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry, e.g., date.
            ' **   -9  Error.

16160       If intRetVal_BuildAssetListInfo = -2 Then
16170         blnNoData = True
16180         Set dbs = CurrentDb
16190         With dbs
                ' ** Empty tmpAssetList2.
16200           Set qdf = dbs.QueryDefs("qryCourtReport_03")
16210           qdf.Execute
16220           .Close
16230         End With
16240       ElseIf intRetVal_BuildAssetListInfo < 0 Then
16250         blnContinue = False
16260       End If

16270       If blnContinue = True Then

              ' ** strRptSuffix should return among "_00B", "_00BA", "_00D", or "_00DA".
16280         If strRptSuffix <> vbNullString Then

16290           gdatStartDate = .DateStart
16300           gdatEndDate = .DateEnd
16310           gstrAccountNo = .cmbAccounts.Column(0)
16320           gstrAccountName = .cmbAccounts.Column(3)

16330           lngCaps = 0&
16340           arr_varCap = Empty

16350           Set dbs = CurrentDb
16360           With dbs
                  ' ** tblReport, captions of Court Reports, by specified [CrtTyp].
16370             Set qdf = .QueryDefs("qryCourtReport_15")
16380             With qdf.Parameters
16390               ![CrtTyp] = "NS"
16400             End With
16410             Set rst = qdf.OpenRecordset
16420             With rst
16430               .MoveLast
16440               lngCaps = .RecordCount
16450               .MoveFirst
16460               arr_varCap = .GetRows(lngCaps)
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
16470               .Close
16480             End With
16490             .Close
16500           End With

16510           If IsNull(.UserReportPath) = False Then
16520             If .UserReportPath <> vbNullString Then
16530               If .UserReportPath_chk = True Then
16540                 If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
16550                   blnUseSavedPath = True
16560                 End If
16570               End If
16580             End If
16590           End If

16600           If blnNoData = False Then
16610             Set dbs = CurrentDb
16620             With dbs

                    ' ** Empty tmpCourtReportData2.
16630               Set qdf = .QueryDefs("qryCourtReport_NS_00_B_06")
16640               qdf.Execute

16650               Select Case strRptSuffix
                    Case "_00B", "_00D"
                      ' ** Append qryCourtReport_NS_00_B_06 to tmpCourtReportData2.
16660                 Set qdf = .QueryDefs("qryCourtReport_NS_00_B_08a")
16670               Case "_00BA", "_00DA"
16680                 Set qdf = .QueryDefs("qryCourtReport_NS_00_B_08b")
16690               End Select
16700               qdf.Execute

                    ' ** tmpCourtReportData2, sorted.
16710               Set qdf = .QueryDefs("qryCourtReport_NS_00_B_10")
16720               Set rst = qdf.OpenRecordset
16730               With rst
16740                 .MoveLast
16750                 lngRecs = .RecordCount
16760                 .MoveFirst
16770                 strLastAssetType = vbNullString
16780                 For lngX = 1& To lngRecs
16790                   If ![assettype] <> strLastAssetType Then
16800                     .Edit
16810                     ![sort2] = 1&
16820                     .Update
16830                     strLastAssetType = ![assettype]
16840                   Else
16850                     .Edit
16860                     ![sort2] = 2&
16870                     .Update
16880                   End If
16890                   If lngX < lngRecs Then .MoveNext
16900                 Next
16910                 .Close
16920               End With

16930               .Close
16940             End With  ' ** dbs.
16950           End If  ' ** blnNoData.

16960           Select Case strRptPeriod
                Case "Beginning"
16970             strQry = "qryCourtReport_NS_00_B_24b"
16980           Case "Ending"
16990             strQry = "qryCourtReport_NS_00_B_24a"
17000           End Select

17010           strRptCap = vbNullString: strRptPathFile = vbNullString
17020           strRptPath = .UserReportPath
17030           strRptName = "rptCourtRptNS" & strRptSuffix

17040           For lngX = 0& To (lngCaps - 1&)
17050             If arr_varCap(C_RNAM, lngX) = strRptName Then
17060               strRptCap = arr_varCap(C_CAPN, lngX)
17070               Exit For
17080             End If
17090           Next

17100           strMacro = "mcrExcelExport_CR_NS" & Mid(strRptName, InStr(strRptName, "_"))
17110           If blnNoData = True Then
17120             strMacro = strMacro & "_nd"
17130           End If

17140           Select Case blnUseSavedPath
                Case True
17150             strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
17160           Case False
17170             DoCmd.Hourglass False
17180             strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
17190           End Select

17200           If strRptPathFile <> vbNullString Then
17210             DoCmd.Hourglass True
17220             DoEvents
17230             If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
17240               Kill strRptPathFile
17250             End If
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
17260             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
17270             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NS_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NS_xxx.xls") = True Then
17280               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NS_xxx.xls") = True Then
17290                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NS_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
17300               ElseIf FileExists(strRptPath & LNK_SEP & "CourtReport_NS_xxx.xls") = True Then
17310                 Name (strRptPath & LNK_SEP & "CourtReport_NS_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
17320               End If
17330               DoEvents
17340               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
17350                 DoEvents
17360                 Select Case blnExcel
                      Case True
17370                   blnAutoStart = .chkOpenExcel
17380                 Case False
17390                   blnAutoStart = .chkOpenWord
17400                 End Select
17410                 If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
17420                 Select Case gblnPrintAll
                      Case True
17430                   lngFiles = lngFiles + 1&
17440                   lngE = lngFiles - 1&
17450                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
17460                   arr_varFile(F_RNAM, lngE) = strRptName
17470                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
17480                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
17490                   FileArraySet_NS arr_varFile  ' ** Procedure: Above.
17500                 Case False
17510                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
17520                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
17530                   End If
17540                   DoEvents
17550                   If blnAutoStart = True Then
17560                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
17570                   End If
17580                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
17590               End If
17600             End If
17610             strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
17620             If strRptPath <> .UserReportPath Then
17630               .UserReportPath = strRptPath
17640               SetUserReportPath_NS frm  ' ** Procedure: Above.
17650             End If
17660           Else
17670             blnContinue = False
17680           End If

17690         End If  ' ** strRptSuffix.

17700       Else
17710         blnContinue = False
17720         DoCmd.Hourglass False
17730         MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
17740       End If  ' ** blnContinue.

17750     End If  ' ** Validate.

17760     DoCmd.Hourglass False

17770   End With

EXITP:
17780   Set rst = Nothing
17790   Set qdf = Nothing
17800   Set dbs = Nothing
17810   Exit Sub

ERRH:
17820   DoCmd.Hourglass False
17830   Select Case ERR.Number
        Case Else
17840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17850   End Select
17860   Resume EXITP

End Sub

Public Sub Calendar_Handler_NS(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_NS"

        Dim strEvent As String, strCtlName As String
        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer, intNum As Integer
        Dim blnRetVal As Boolean

17910   With frm

17920     strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
17930     strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.
17940     intNum = Val(Right(strCtlName, 1))

17950     Select Case strEvent
          Case "Click"
17960       Select Case intNum
            Case 1
17970         datStartDate = Date
17980         datEndDate = 0
17990         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
18000         If blnRetVal = True Then
18010           .DateStart = datStartDate
18020         Else
18030           .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
18040         End If
18050         .DateStart.SetFocus
18060       Case 2
18070         datStartDate = Date
18080         datEndDate = 0
18090         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
18100         If blnRetVal = True Then
18110           .DateEnd = datStartDate
18120         Else
18130           .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
18140         End If
18150         .DateEnd.SetFocus
18160         Cancel = 0
18170         .DateEnd_Exit Cancel  ' ** Form Procedure: frmRpt_CourtReports_NS.
18180         If Cancel = 0 Then
18190           .cmbAccounts.SetFocus
18200         End If
18210       End Select
18220     Case "GotFocus"
18230       Select Case intNum
            Case 1
18240         blnCalendar1_Focus = True
18250         .cmdCalendar1_raised_semifocus_dots_img.Visible = True
18260         .cmdCalendar1_raised_img.Visible = False
18270         .cmdCalendar1_raised_focus_img.Visible = False
18280         .cmdCalendar1_raised_focus_dots_img.Visible = False
18290         .cmdCalendar1_sunken_focus_dots_img.Visible = False
18300         .cmdCalendar1_raised_img_dis.Visible = False
18310       Case 2
18320         blnCalendar2_Focus = True
18330         .cmdCalendar2_raised_semifocus_dots_img.Visible = True
18340         .cmdCalendar2_raised_img.Visible = False
18350         .cmdCalendar2_raised_focus_img.Visible = False
18360         .cmdCalendar2_raised_focus_dots_img.Visible = False
18370         .cmdCalendar2_sunken_focus_dots_img.Visible = False
18380         .cmdCalendar2_raised_img_dis.Visible = False
18390       End Select
18400     Case "MouseDown"
18410       Select Case intNum
            Case 1
18420         blnCalendar1_MouseDown = True
18430         .cmdCalendar1_sunken_focus_dots_img.Visible = True
18440         .cmdCalendar1_raised_img.Visible = False
18450         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
18460         .cmdCalendar1_raised_focus_img.Visible = False
18470         .cmdCalendar1_raised_focus_dots_img.Visible = False
18480         .cmdCalendar1_raised_img_dis.Visible = False
18490       Case 2
18500         blnCalendar2_MouseDown = True
18510         .cmdCalendar2_sunken_focus_dots_img.Visible = True
18520         .cmdCalendar2_raised_img.Visible = False
18530         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
18540         .cmdCalendar2_raised_focus_img.Visible = False
18550         .cmdCalendar2_raised_focus_dots_img.Visible = False
18560         .cmdCalendar2_raised_img_dis.Visible = False
18570       End Select
18580     Case "MouseMove"
18590       Select Case intNum
            Case 1
18600         If blnCalendar1_MouseDown = False Then
18610           Select Case blnCalendar1_Focus
                Case True
18620             .cmdCalendar1_raised_focus_dots_img.Visible = True
18630             .cmdCalendar1_raised_focus_img.Visible = False
18640           Case False
18650             .cmdCalendar1_raised_focus_img.Visible = True
18660             .cmdCalendar1_raised_focus_dots_img.Visible = False
18670           End Select
18680           .cmdCalendar1_raised_img.Visible = False
18690           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
18700           .cmdCalendar1_sunken_focus_dots_img.Visible = False
18710           .cmdCalendar1_raised_img_dis.Visible = False
18720         End If
18730       Case 2
18740         If blnCalendar2_MouseDown = False Then
18750           Select Case blnCalendar2_Focus
                Case True
18760             .cmdCalendar2_raised_focus_dots_img.Visible = True
18770             .cmdCalendar2_raised_focus_img.Visible = False
18780           Case False
18790             .cmdCalendar2_raised_focus_img.Visible = True
18800             .cmdCalendar2_raised_focus_dots_img.Visible = False
18810           End Select
18820           .cmdCalendar2_raised_img.Visible = False
18830           .cmdCalendar2_raised_semifocus_dots_img.Visible = False
18840           .cmdCalendar2_sunken_focus_dots_img.Visible = False
18850           .cmdCalendar2_raised_img_dis.Visible = False
18860         End If
18870       End Select
18880     Case "MouseUp"
18890       Select Case intNum
            Case 1
18900         .cmdCalendar1_raised_focus_dots_img.Visible = True
18910         .cmdCalendar1_raised_img.Visible = False
18920         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
18930         .cmdCalendar1_raised_focus_img.Visible = False
18940         .cmdCalendar1_sunken_focus_dots_img.Visible = False
18950         .cmdCalendar1_raised_img_dis.Visible = False
18960         blnCalendar1_MouseDown = False
18970       Case 2
18980         .cmdCalendar2_raised_focus_dots_img.Visible = True
18990         .cmdCalendar2_raised_img.Visible = False
19000         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
19010         .cmdCalendar2_raised_focus_img.Visible = False
19020         .cmdCalendar2_sunken_focus_dots_img.Visible = False
19030         .cmdCalendar2_raised_img_dis.Visible = False
19040         blnCalendar2_MouseDown = False
19050       End Select
19060     Case "LostFocus"
19070       Select Case intNum
            Case 1
19080         .cmdCalendar1_raised_img.Visible = True
19090         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
19100         .cmdCalendar1_raised_focus_img.Visible = False
19110         .cmdCalendar1_raised_focus_dots_img.Visible = False
19120         .cmdCalendar1_sunken_focus_dots_img.Visible = False
19130         .cmdCalendar1_raised_img_dis.Visible = False
19140         blnCalendar1_Focus = False
19150       Case 2
19160         .cmdCalendar2_raised_img.Visible = True
19170         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
19180         .cmdCalendar2_raised_focus_img.Visible = False
19190         .cmdCalendar2_raised_focus_dots_img.Visible = False
19200         .cmdCalendar2_sunken_focus_dots_img.Visible = False
19210         .cmdCalendar2_raised_img_dis.Visible = False
19220         blnCalendar2_Focus = False
19230       End Select
19240     End Select

19250   End With

EXITP:
19260   Exit Sub

ERRH:
19270   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
19280   Case Else
19290     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19300   End Select
19310   Resume EXITP

End Sub
