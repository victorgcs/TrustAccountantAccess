Attribute VB_Name = "modCourtReportsFL"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCourtReportsFL"

'VGC 09/13/2017: CHANGES!

' ** Array: arr_varFLRpt().
Private lngFLRpts As Long, arr_varFLRpt As Variant
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
'Private Const C_RID   As Integer = 0
Private Const C_RNAM  As Integer = 1
'Private Const C_CAP   As Integer = 2
Private Const C_CAPN  As Integer = 3

Private blnExcel As Boolean, blnAllCancel As Boolean
Private strCaseNum As String, strThisProc As String
' **

Public Function FLRpt(ByVal intRptNum As Integer) As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "FLRpt"

        Dim blnRetVal As Boolean

110     Select Case intRptNum
        Case CRPT_ADDL_PROP, CRPT_OTH_CHG, CRPT_NET_INCOME, CRPT_LOSSES, CRPT_NET_LOSS, CRPT_OTH_CRED, _
            CRPT_INVEST_INFO, CRPT_CHANGES, CRPT_ON_HAND_BEGL, CRPT_CASH_BEG, CRPT_NON_CASH_BEG, _
            CRPT_ON_HAND_ENDL, CRPT_CASH_END, CRPT_NON_CASH_END
120       blnRetVal = False
130     Case Else
140       blnRetVal = True
150     End Select

EXITP:
160     FLRpt = blnRetVal
170     Exit Function

ERRH:
180     blnRetVal = False
190     Select Case ERR.Number
        Case Else
200       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
210     End Select
220     Resume EXITP

End Function

Public Function FLNum(ByVal strConst As String) As Integer

300   On Error GoTo ERRH

        Const THIS_PROC As String = "FLNum"

        Dim lngX As Long
        Dim intRetVal As Integer

310     For lngX = 0& To (lngFLRpts - 1&)
320       If arr_varFLRpt(CR_CON, lngX) = strConst Then
330         intRetVal = arr_varFLRpt(CR_NUM, lngX)
340         Exit For
350       End If
360     Next

EXITP:
370     FLNum = intRetVal
380     Exit Function

ERRH:
390     intRetVal = 0
400     Select Case ERR.Number
        Case Else
410       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
420     End Select
430     Resume EXITP

End Function

Public Function FLCourtReportLoad() As Boolean

500   On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportLoad"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngX As Long
        Dim blnRetVal As Boolean

510     blnRetVal = True

520     Set dbs = CurrentDb
530     Set qdf = dbs.QueryDefs("qryCourtReport_FL_20")
540     Set rst = qdf.OpenRecordset
550     With rst
560       .MoveLast
570       lngFLRpts = .RecordCount
580       .MoveFirst
590       arr_varFLRpt = .GetRows(lngFLRpts)
          ' *******************************************************
          ' ** Array: arr_varFLRpt()
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
600       .Close
610     End With
620     dbs.Close

        ' ** Set Null text fields to vbNullString to simplify comparisons.
630     For lngX = 0& To (lngFLRpts - 1&)
640       If IsEmpty(arr_varFLRpt(CR_GRPTXT, lngX)) = True Then
650         arr_varFLRpt(CR_GRPTXT, lngX) = vbNullString
660       ElseIf IsNull(arr_varFLRpt(CR_GRPTXT, lngX)) = True Then
670         arr_varFLRpt(CR_GRPTXT, lngX) = vbNullString
680       End If
690       If IsEmpty(arr_varFLRpt(CR_SCHED, lngX)) = True Then
700         arr_varFLRpt(CR_SCHED, lngX) = vbNullString
710       ElseIf IsNull(arr_varFLRpt(CR_SCHED, lngX)) = True Then
720         arr_varFLRpt(CR_SCHED, lngX) = vbNullString
730       End If
740     Next

        ' ** Change Schedule letters and/or date source.
750     For lngX = 0& To (lngFLRpts - 1&)
760       Select Case arr_varFLRpt(CR_CON, lngX)
          Case "CRPT_RECEIPTS"
770         arr_varFLRpt(CR_SCHED, lngX) = "A"  ' ** No change.
780       Case "CRPT_DISBURSEMENTS"
790         arr_varFLRpt(CR_SCHED, lngX) = "B"
800       Case "CRPT_DISTRIBUTIONS"
810         arr_varFLRpt(CR_SCHED, lngX) = "C"
820       Case "CRPT_GAINS"
830         arr_varFLRpt(CR_SCHED, lngX) = "D"
840         arr_varFLRpt(CR_DATE, lngX) = "AssetDate"  ' ** VGC 10/12/2008: Change to match NS version, which uses assetdate.
850       Case "CRPT_ON_HAND_ENDL", "CRPT_Cash_END", "CRPT_NON_Cash_END", "CRPT_ON_HAND_END"
860         arr_varFLRpt(CR_SCHED, lngX) = "E"
870       Case "CRPT_LOSSES"
880         arr_varFLRpt(CR_SCHED, lngX) = vbNullString
890       End Select
900     Next

910     For lngX = 0& To (lngFLRpts - 1&)

920       If arr_varFLRpt(CR_DIVTTL, lngX) = "CHARGES" And CRPT_DIV_CHARGES = 0 Then
930         CRPT_DIV_CHARGES = arr_varFLRpt(CR_DIV, lngX)  ' ** 20
940       ElseIf arr_varFLRpt(CR_DIVTTL, lngX) = "CREDITS" And CRPT_DIV_CREDITS = 0 Then
950         CRPT_DIV_CREDITS = arr_varFLRpt(CR_DIV, lngX)  ' ** 40
960       ElseIf arr_varFLRpt(CR_DIVTTL, lngX) = "ADDITIONAL INFORMATION" And CRPT_DIV_ADDL = 0 Then
970         CRPT_DIV_ADDL = arr_varFLRpt(CR_DIV, lngX)     ' ** 60
980       End If

990       If arr_varFLRpt(CR_CON, lngX) = "CRPT_ON_HAND_BEGL" Then       '      2
1000        CRPT_ON_HAND_BEGL = arr_varFLRpt(CR_NUM, lngX)
1010      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_Cash_BEG" Then       '      3
1020        CRPT_CASH_BEG = arr_varFLRpt(CR_NUM, lngX)
1030      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_NON_Cash_BEG" Then   '      4
1040        CRPT_NON_CASH_BEG = arr_varFLRpt(CR_NUM, lngX)
1050      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_ON_HAND_BEG" Then    '5     5
1060        CRPT_ON_HAND_BEG = arr_varFLRpt(CR_NUM, lngX)
1070      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_ADDL_PROP" Then      '10   10
1080        CRPT_ADDL_PROP = arr_varFLRpt(CR_NUM, lngX)
1090      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_RECEIPTS" Then       '20   20
1100        CRPT_RECEIPTS = arr_varFLRpt(CR_NUM, lngX)
1110      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_GAINS" Then          '30   30
1120        CRPT_GAINS = arr_varFLRpt(CR_NUM, lngX)
1130      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_OTH_CHG" Then        '     40
1140        CRPT_OTH_CHG = arr_varFLRpt(CR_NUM, lngX)
1150      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_NET_INCOME" Then     '40   50
1160        CRPT_NET_INCOME = arr_varFLRpt(CR_NUM, lngX)
1170      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_DISBURSEMENTS" Then  '50   60
1180        CRPT_DISBURSEMENTS = arr_varFLRpt(CR_NUM, lngX)
1190      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_LOSSES" Then         '60   70
1200        CRPT_LOSSES = arr_varFLRpt(CR_NUM, lngX)
1210      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_DISTRIBUTIONS" Then  '80   80
1220        CRPT_DISTRIBUTIONS = arr_varFLRpt(CR_NUM, lngX)
1230      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_OTH_CRED" Then       '     90
1240        CRPT_OTH_CRED = arr_varFLRpt(CR_NUM, lngX)
1250      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_NET_LOSS" Then       '70  100
1260        CRPT_NET_LOSS = arr_varFLRpt(CR_NUM, lngX)
1270      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_ON_HAND_ENDL" Then   '    107
1280        CRPT_ON_HAND_ENDL = arr_varFLRpt(CR_NUM, lngX)
1290      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_Cash_END" Then       '    108
1300        CRPT_CASH_END = arr_varFLRpt(CR_NUM, lngX)
1310      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_NON_Cash_END" Then   '    109
1320        CRPT_NON_CASH_END = arr_varFLRpt(CR_NUM, lngX)
1330      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_ON_HAND_END" Then    '90  110
1340        CRPT_ON_HAND_END = arr_varFLRpt(CR_NUM, lngX)
1350      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_INVEST_INFO" Then    '100 120
1360        CRPT_INVEST_INFO = arr_varFLRpt(CR_NUM, lngX)
1370      ElseIf arr_varFLRpt(CR_CON, lngX) = "CRPT_CHANGES" Then        '110 130
1380        CRPT_CHANGES = arr_varFLRpt(CR_NUM, lngX)
1390      End If

1400    Next

EXITP:
1410    Set rst = Nothing
1420    Set qdf = Nothing
1430    Set dbs = Nothing
1440    FLCourtReportLoad = blnRetVal
1450    Exit Function

ERRH:
1460    blnRetVal = False
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Function

Public Function FLCourtReportCategory(ByVal intCourtReport As Integer) As String
' ** HOW COME SO MANY OF THE LINES CALLING THIS USE THE REPORT NUMBER ALONE WITHOUT MULTIPLYING BY 10?
' ** THEY'LL ALL COME BACK 'Unknown'?

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportCategory"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

1610    strRetVal = vbNullString

1620    blnFound = False
1630    For lngX = 0& To (lngFLRpts - 1&)
1640      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
1650        strRetVal = arr_varFLRpt(CR_CAT, lngX)
1660        blnFound = True
1670        Exit For
1680      End If
1690    Next
1700    If blnFound = False Then strRetVal = "Unknown"

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
1710    FLCourtReportCategory = strRetVal
1720    Exit Function

ERRH:
1730    strRetVal = RET_ERR
1740    Select Case ERR.Number
        Case Else
1750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1760    End Select
1770    Resume EXITP

End Function

Public Function FLCourtReportDataID(ByVal strComment As String, ByVal curPCash As Currency, ByVal curICash As Currency, ByVal curCost As Currency, ByVal curGainLoss As Currency, ByVal strTaxCode As String, ByVal strJournalType As String) As Integer
' ** The code returned is the Court Report that the data belongs to.
' ** Used in queries:
' **   qryCourtReport - Summary-1-FL
' **   qryCourtReport_FL_21

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportDataID"

        Dim lngTaxcode As Long
        Dim intRetVal As Integer

1810    intRetVal = 0  ' ** Force to 0 just in case.

1820    If glngTaxCode_Distribution = 0& Then
1830      glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
1840    End If
1850    If Trim(strTaxCode) <> vbNullString Then
1860      lngTaxcode = Val(strTaxCode)
1870    Else
1880      lngTaxcode = 0&
1890    End If

        ' ** Additional Property Received.
1900    If (strJournalType = "Deposit" And Not (strComment Like "*stock split*")) _
            Or (strJournalType = "Cost Adj." And curCost > 0) Then
1910      intRetVal = CRPT_ADDL_PROP
1920    Else
          ' ** Receipts.
1930      If (strJournalType = "Received" And (curPCash > 0 Or curICash > 0)) Or (strJournalType = "Misc." And (curPCash + curICash > 0)) _
              Or (strJournalType = "Dividend" And curICash > 0) Or (strJournalType = "Purchase" And curICash > 0) _
              Or (strJournalType = "Sold" And curICash > 0) Or (strJournalType = "Interest" And curICash > 0) Then
1940        intRetVal = CRPT_RECEIPTS
1950      Else
            ' ** Gains on Sales.
1960        If (strJournalType = "Sold" And curGainLoss > 0) Then
1970          intRetVal = CRPT_GAINS
1980        Else
              ' ** Net Income is entered by hand (intRetVal = CRPT_NET_INCOME).
              ' ** Disbursements.
              ' ** 07/17/08: Added negative Purchases, per Rich.
              ' **           Moved negative Cost Adjustments to Losses, per Rich.
1990          If ((strJournalType = "Paid" And curPCash <> 0 And lngTaxcode <> glngTaxCode_Distribution) _
                  Or (strJournalType = "Paid" And curICash <> 0 And lngTaxcode <> glngTaxCode_Distribution) _
                  Or (strJournalType = "Misc." And ((curPCash + curICash < 0) Or ((curICash = -curPCash) And (curCost = 0)))) _
                  Or (strJournalType = "Withdrawn" And lngTaxcode <> glngTaxCode_Distribution) _
                  Or (strJournalType = "Liability" And curICash < 0) _
                  Or (strJournalType = "Purchase" And curICash < 0 And curPCash < 0)) Then  '<> "Distribution"
                ' ** Now includes simple transfers between Income and Principal.
2000            intRetVal = CRPT_DISBURSEMENTS    '####  TAXCODE  ####
2010          Else
                ' ** Losses on Sales.
                ' ** 07/17/08: Added negative Cost Adjustments, per Rich.
2020            If (strJournalType = "Sold" And curGainLoss < 0) Or (strJournalType = "Cost Adj." And curCost < 0) Then
2030              intRetVal = CRPT_LOSSES
2040            Else
                  ' ** Net Loss is entered by hand (intRetVal = CRPT_NET_LOSS).
                  ' ** Distributions.
2050              If ((strJournalType = "Paid" And curPCash <> 0 And lngTaxcode = glngTaxCode_Distribution) _
                      Or (strJournalType = "Paid" And curICash <> 0 And lngTaxcode = glngTaxCode_Distribution) _
                      Or (strJournalType = "Withdrawn" And lngTaxcode = glngTaxCode_Distribution)) Then  '= "Distribution"
2060                intRetVal = CRPT_DISTRIBUTIONS    '####  TAXCODE  ####
2070              Else
                    ' ** Property on Hand added separately  (intRetVal = CRPT_ON_HAND_END).
                    ' ** Information for Investments Made.
2080                If (strJournalType = "Purchase") Then
2090                  intRetVal = CRPT_INVEST_INFO
2100                Else
                      ' ** Changes in Investment Holdings.
2110                  If (strJournalType = "Sold" And curGainLoss = 0) _
                          Or (strJournalType = "Deposit" And strComment Like "*stock split*") _
                          Or (strJournalType = "Purchase" And curICash <= 0) _
                          Or (strJournalType = "Liability") Then
2120                    intRetVal = CRPT_CHANGES
2130                  End If
2140                End If
2150              End If
2160            End If
2170          End If
2180        End If
2190      End If
2200    End If

EXITP:
2210    FLCourtReportDataID = intRetVal
2220    Exit Function

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Function

Public Function FLCourtReportDate(ByVal intCourtReport As Integer, ByVal datTransDate As Date, ByVal datAssetDate As Date) As Date

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportDate"

        Dim lngX As Long, blnFound As Boolean
        Dim datRetVal As Date

2310    blnFound = False
2320    For lngX = 0& To (lngFLRpts - 1&)
2330      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
2340        Select Case arr_varFLRpt(CR_DATE, lngX)
            Case "TransDate"
2350          datRetVal = datTransDate
2360        Case "AssetDate"
2370          datRetVal = datAssetDate
2380        End Select
2390        blnFound = True
2400        Exit For
2410      End If
2420    Next
2430    If blnFound = False Then datRetVal = #1/1/1900#

EXITP:
2440    FLCourtReportDate = datRetVal
2450    Exit Function

ERRH:
2460    datRetVal = #1/1/1900#
2470    Select Case ERR.Number
        Case Else
2480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2490    End Select
2500    Resume EXITP

End Function

Public Function FLCourtReportDivision(ByVal intCourtReport As Integer) As Integer

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportDivision"

        Dim lngX As Long, blnFound As Boolean
        Dim intRetVal As Integer

2610    blnFound = False
2620    For lngX = 0& To (lngFLRpts - 1&)
2630      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
2640        intRetVal = arr_varFLRpt(CR_DIV, lngX)
2650        blnFound = True
2660        Exit For
2670      End If
2680    Next
2690    If blnFound = False Then intRetVal = 0

EXITP:
2700    FLCourtReportDivision = intRetVal
2710    Exit Function

ERRH:
2720    intRetVal = 0
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Function

Public Function FLCourtReportDivisionText(ByVal intCourtReport As Integer) As String

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportDivisionText"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

2810    blnFound = False
2820    For lngX = 0& To (lngFLRpts - 1&)
2830      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
2840        strRetVal = arr_varFLRpt(CR_DIVTXT, lngX)
2850        blnFound = True
2860        Exit For
2870      End If
2880    Next
2890    If blnFound = False Then strRetVal = "Unknown"

EXITP:
2900    FLCourtReportDivisionText = strRetVal
2910    Exit Function

ERRH:
2920    strRetVal = RET_ERR
2930    Select Case ERR.Number
        Case Else
2940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2950    End Select
2960    Resume EXITP

End Function

Public Function FLCourtReportDivisionTitle(ByVal intCourtReport As Integer) As String

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportDivisionTitle"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

3010    blnFound = False
3020    For lngX = 0& To (lngFLRpts - 1&)
3030      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
3040        strRetVal = arr_varFLRpt(CR_DIVTTL, lngX)
3050        blnFound = True
3060        Exit For
3070      End If
3080    Next
3090    If blnFound = False Then strRetVal = "Unknown"

EXITP:
3100    FLCourtReportDivisionTitle = strRetVal
3110    Exit Function

ERRH:
3120    strRetVal = RET_ERR
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Function

Public Function FLCourtReportGroup(ByVal intCourtReport As Integer) As Integer

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportGroup"

        Dim lngX As Long, blnFound As Boolean
        Dim intRetVal As Integer

3210    blnFound = False
3220    For lngX = 0& To (lngFLRpts - 1&)
3230      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
3240        intRetVal = arr_varFLRpt(CR_GRP, lngX)
3250        blnFound = True
3260        Exit For
3270      End If
3280    Next
3290    If blnFound = False Then intRetVal = 0

EXITP:
3300    FLCourtReportGroup = intRetVal
3310    Exit Function

ERRH:
3320    intRetVal = 0
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Function

Public Function FLCourtReportGroupText(ByVal intCourtReport As Integer) As String

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportGroupText"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

3410    blnFound = False
3420    For lngX = 0& To (lngFLRpts - 1&)
3430      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
3440        strRetVal = arr_varFLRpt(CR_GRPTXT, lngX)
3450        blnFound = True
3460        Exit For
3470      End If
3480    Next
3490    If blnFound = False Then strRetVal = "Unknown"

EXITP:
3500    FLCourtReportGroupText = strRetVal
3510    Exit Function

ERRH:
3520    strRetVal = RET_ERR
3530    Select Case ERR.Number
        Case Else
3540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3550    End Select
3560    Resume EXITP

End Function

Public Function FLCourtReportSchedule(ByVal intCourtReport As Integer) As String

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportSchedule"

        Dim lngX As Long, blnFound As Boolean
        Dim strRetVal As String

3610    blnFound = False
3620    For lngX = 0& To (lngFLRpts - 1&)
3630      If arr_varFLRpt(CR_NUM, lngX) = intCourtReport Then
3640        strRetVal = arr_varFLRpt(CR_SCHED, lngX)
3650        blnFound = True
3660        Exit For
3670      End If
3680    Next
3690    If blnFound = False Then strRetVal = vbNullString

EXITP:
3700    FLCourtReportSchedule = strRetVal
3710    Exit Function

ERRH:
3720    strRetVal = RET_ERR
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Function

Private Function FLCourtReportCaseNum() As String

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportCaseNum"

        Dim strRetVal As String

3810    strRetVal = vbNullString
3820  On Error Resume Next

3830    strRetVal = DLookup("[CaseNum]", "account", "[accountno] = '" & gstrAccountNo & "'")
3840  On Error GoTo ERRH

EXITP:
3850    FLCourtReportCaseNum = strRetVal
3860    Exit Function

ERRH:
3870    strRetVal = RET_ERR
3880    Select Case ERR.Number
        Case Else
3890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3900    End Select
3910    Resume EXITP

End Function

Public Function FLCourtReportDollars(ByVal intCourtReport As Integer, ByVal curPCash As Currency, ByVal curICash As Currency, ByVal curCost As Currency, ByVal strJournalType As String) As Double

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "FLCourtReportDollars"

        Dim dblRetVal As Double

4010    Select Case intCourtReport
        Case CRPT_ADDL_PROP      ' ** WAS: 10   NOW: 10
4020      dblRetVal = curCost + curPCash
4030    Case CRPT_RECEIPTS       ' ** WAS: 20   NOW: 20
4040      If (strJournalType = "Sold" And curICash > 0) Or strJournalType = "Purchase" And curICash > 0 And curPCash + curCost = 0 Then
4050        dblRetVal = curICash
4060      Else
4070        dblRetVal = curICash + curPCash
4080      End If
4090    Case CRPT_GAINS          ' ** WAS: 30   NOW: 30
4100      dblRetVal = curPCash + curCost
4110    Case CRPT_OTH_CHG        ' ** WAS:      NOW: 40
          ' ** Unknown.
4120    Case CRPT_NET_INCOME     ' ** WAS: 40   NOW: 50
          ' ** Nothing.
4130    Case CRPT_DISBURSEMENTS  ' ** WAS: 50   NOW: 60
4140      If strJournalType = "Withdrawn" Then
4150        dblRetVal = curCost * -1
4160      Else
4170        dblRetVal = (curICash + curPCash) * -1
4180      End If
4190    Case CRPT_LOSSES         ' ** WAS: 60   NOW: 70
4200      dblRetVal = (curPCash + curCost) * -1
4210    Case CRPT_DISTRIBUTIONS  ' ** WAS: 80   NOW: 80
4220      If strJournalType = "Withdrawn" Then
4230        dblRetVal = curCost * -1
4240      Else
4250        dblRetVal = (curICash + curPCash) * -1
4260      End If
4270    Case CRPT_OTH_CRED       ' ** WAS:      NOW: 90
          ' ** Unknown.
4280    Case CRPT_NET_LOSS       ' ** WAS: 70   NOW: 100
          ' ** Nothing.
4290    Case CRPT_ON_HAND_END    ' ** WAS: 90   NOW: 110
          ' ** Nothing.
4300    Case CRPT_INVEST_INFO    ' ** WAS: 100  NOW: 120
4310      dblRetVal = (curICash + curPCash) * -1
4320    Case CRPT_CHANGES        ' ** WAS: 110  NOW: 130
4330      dblRetVal = curCost
4340    Case Else
4350      dblRetVal = -999999
4360    End Select

EXITP:
4370    FLCourtReportDollars = dblRetVal
4380    Exit Function

ERRH:
4390    dblRetVal = 0#
4400    Select Case ERR.Number
        Case Else
4410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4420    End Select
4430    Resume EXITP

End Function

Public Function FLBuildCourtReportData(ByVal strReportNumber As String, strControlName As String) As Integer
' ** Called by:
' **
' ** Return Codes:
' **   0  Success.
' **  -1  Canceled.
' **  -9  Error.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "FLBuildCourtReportData"

        'Dim rsxDataIn As ADODB.Recordset, rsxDataOut As ADODB.Recordset  ' ** Early binding.
        Dim rsxDataIn As Object, rsxDataOut As Object                     ' ** Late binding.
        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim intReportSection As Integer
        Dim datTmp01 As Date
        Dim lngX As Long
        Dim intRetVal As Integer

4510    intRetVal = 0  ' ** Success.

4520    FLCourtReportLoad  ' ** Function: Above.

4530    strCaseNum = FLCourtReportCaseNum  ' ** Function: Above.

        ' ** Delete the data from the tmpCourtReport table.
4540    Set dbs = CurrentDb
4550    With dbs
4560      Set qdf = .QueryDefs("qryCourtReport_02")
4570      qdf.Execute
4580      .Close
4590    End With
4600    Set qdf = Nothing
4610    Set dbs = Nothing

        'Set rsxDataOut = New ADODB.Recordset             ' ** Early binding.
4620    Set rsxDataOut = CreateObject("ADODB.Recordset")  ' ** Late binding.
4630  On Error Resume Next
4640    rsxDataOut.Open "tmpCourtReportData", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTable
4650    If ERR.Number <> 0 Then
4660      Select Case ERR.Number
          Case -2147217838  ' ** Data source object is already initialized.
4670  On Error GoTo ERRH
            ' ** For now, just let it go, since I think that means it's already available.
4680      Case Else
4690        intRetVal = -9
4700        zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4710  On Error GoTo ERRH
4720      End Select
4730    Else
4740  On Error GoTo ERRH
4750    End If

4760    If intRetVal = 0 Then

          ' ** Build dummy records with zero in the amount to insure that all report sections are displayed.
          ' ** One for each 10's section in tblCourtReport. Report 5 gets added later in FLGetCourtReportData(), below.
4770      For lngX = 0& To (lngFLRpts - 1&)
4780        With rsxDataOut
4790          .AddNew
4800          .Fields("ReportNumber") = arr_varFLRpt(CR_NUM, lngX)
4810          .Fields("ReportCategory") = FLCourtReportCategory(CInt(arr_varFLRpt(CR_NUM, lngX)))  ' ** Function: Above.
4820          .Fields("ReportGroup") = FLCourtReportGroup(CInt(arr_varFLRpt(CR_NUM, lngX)))  ' ** Function: Above.
4830          .Fields("ReportDivision") = FLCourtReportDivision(CInt(arr_varFLRpt(CR_NUM, lngX)))  ' ** Function: Above.
4840          .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CInt(arr_varFLRpt(CR_NUM, lngX)))  ' ** Function: Above.
4850          .Fields("ReportDivisionText") = FLCourtReportDivisionText(CInt(arr_varFLRpt(CR_NUM, lngX)))  ' ** Function: Above.
4860          .Fields("ReportGroupText") = FLCourtReportGroupText(CInt(arr_varFLRpt(CR_NUM, lngX)))  ' ** Function: Above.
4870          .Fields("accountno") = gstrAccountNo
4880          .Fields("date") = gdatEndDate
4890          .Fields("journaltype") = "Miscellaneous"
4900          .Fields("shareface") = 0
4910          .Fields("Description") = "Dummy"
4920          .Fields("Amount") = 0
4930          .Fields("Amount_Inc") = 0
4940          .Fields("Amount_Prin") = 0
4950          .Fields("Amount_Cost") = 0
4960          .Fields("revcode_ID") = 0
4970          .Fields("revcode_DESC") = "Dummy entry"
4980          .Fields("revcode_TYPE") = 1
4990          .Fields("revcode_SORTORDER") = 0
5000          .Fields("ReportSchedule") = FLCourtReportSchedule(CInt(arr_varFLRpt(CR_NUM, lngX)))  ' ** Function: Above.
5010          .Fields("CaseNum") = strCaseNum
5020          .Update
5030        End With
5040      Next

          ' ** When called from the TAReports CommandBar, these variables should already be filled.
          ' **   gstrCrtRpt_Ordinal
          ' **   gstrCrtRpt_Version
          ' **   gstrCrtRpt_CashAssets_Beg
          ' **   gstrCrtRpt_NetIncome
          ' **   gstrCrtRpt_NetLoss
          ' **   gstrCrtRpt_CashAssets_End

          ' ** Summary.
5050      If strReportNumber = "0" Then  'Or strReportNumber = "0A" Then
            ' ** Get entered data.
5060        intRetVal = FLGetCourtReportData(THIS_PROC & "^" & strControlName)  ' ** Function: Below.
            ' ** Return Codes:
            ' **   0  Success.
            ' **  -1  Canceled.
            ' **  -9  Error.
5070      End If

5080    Else
          ' ** rsxDataOut failed to open.
5090    End If  ' ** intRetVal.

5100    If intRetVal = 0 Then

          'Set rsxDataIn = New ADODB.Recordset             ' ** Early binding.
5110      Set rsxDataIn = CreateObject("ADODB.Recordset")  ' ** Late binding.

5120      Select Case strReportNumber
          Case "4", "4A", "4B", "4C"  ' ** WHY ISN'T 4B INCLUDED (Guaradian), WHEN 4C IS (also Guardian)?
            'Report 4: instead of 21, use qryCourtReport_FL_04e
5130        rsxDataIn.Open "qryCourtReport_FL_04e", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
5140      Case Else
5150        rsxDataIn.Open "qryCourtReport_FL_21", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
5160      End Select

          'SetDateSpecificSQL(strAccountno As String, strOption As String, strActiveFormName As String, Optional varStartDate As Variant, Optional varEndDate As Variant)
          'Called by blnBuildAssetListInfo(), WHICH HAPPENS AFTER FLBuildCourtReportData()!!!!!!!!!!!!!!!!!!
          'To use qryCourtReport_05:
          '1st, strOption must be other than "StatementTransactions".
          '2nd, varStartDate must be included.
          'To use qryMaxBalDates:
          'Anything other than above.
          'Now I have

          ' ** Loop through data, processing records for requested account.
5170      Do While rsxDataIn.EOF = False
5180        If Trim(rsxDataIn.Fields("accountno")) = gstrAccountNo Then
              ' ** I have no explanation for why this Trim() is necessary! VGC 02/27/2013.
5190          intReportSection = rsxDataIn.Fields("Reportnumber")  ' ** Do this because it recalcs each time it's referenced.
              ' ** Find the right date to use.
5200          datTmp01 = FLCourtReportDate(intReportSection, rsxDataIn.Fields("transdate"), rsxDataIn.Fields("assetdate"))  ' ** Function: Above.
              ' ** If the date for the transaction is within range, build a report record.
              'USES assetdate!    ####
5210          If datTmp01 >= gdatStartDate And datTmp01 < (gdatEndDate + 1) Then
                ' ** If the journal type is MISC, create 2 transactions.
5220            If rsxDataIn.Fields("Reportnumber") = 0 And rsxDataIn.Fields("journaltype") = "Misc." Then
                  ' ** It'll only come back '0' if FLCourtReportDataID(), below, doesn't find a good match.
5230              With rsxDataOut
5240                .AddNew
5250                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
5260                .Fields("date") = datTmp01
5270                .Fields("journaltype") = "Miscellaneous"
5280                .Fields("shareface") = rsxDataIn.Fields("shareface")
5290                .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                      rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
5300                .Fields("Amount") = rsxDataIn.Fields("icash")
5310                .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
5320                .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
5330                .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
5340                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
5350                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
5360                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5370                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5380                .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
5390                .Fields("CaseNum") = strCaseNum
5400                If rsxDataIn.Fields("icash") > 0 Then
                      ' ** Does this '7' have any relation to Report 70 (Net Loss), or is it just for sort order?
5410                  .Fields("ReportNumber") = 7
5420                  .Fields("ReportCategory") = FLCourtReportCategory(7)  ' ** This'll come back 'Unknown'!
5430                  .Fields("ReportGroup") = FLCourtReportGroup(7)  ' ** This'll come back 0!
5440                  .Fields("ReportDivision") = FLCourtReportDivision(7)  ' ** This'll come back 0!
5450                  .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(7)  ' ** This'll come back 'Unknown'!
5460                  .Fields("ReportDivisionText") = FLCourtReportDivisionText(7)  ' ** This'll come back 'Unknown'!
5470                  .Fields("ReportGroupText") = FLCourtReportGroupText(7)  ' ** This'll come back 'Unknown'!
5480                Else
                      ' ** Does this '8' have any relation to Report 80 (Distributions), or is it just for sort order?
5490                  .Fields("ReportNumber") = 8
5500                  .Fields("ReportCategory") = FLCourtReportCategory(8)  ' ** This'll come back 'Unknown'!
5510                  .Fields("ReportGroup") = FLCourtReportGroup(8)  ' ** This'll come back 0!
5520                  .Fields("ReportDivision") = FLCourtReportDivision(8)  ' ** This'll come back 0!
5530                  .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(8)  ' ** This'll come back 'Unknown'!
5540                  .Fields("ReportDivisionText") = FLCourtReportDivisionText(8)  ' ** This'll come back 'Unknown'!
5550                  .Fields("ReportGroupText") = FLCourtReportGroupText(8)  ' ** This'll come back 'Unknown'!
5560                End If
5570                .Update
5580                .AddNew
5590                .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
5600                .Fields("date") = datTmp01
5610                .Fields("journaltype") = rsxDataIn.Fields("journaltype")
5620                .Fields("shareface") = rsxDataIn.Fields("shareface")
5630                .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                      rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
5640                .Fields("Amount") = rsxDataIn.Fields("icash")
5650                .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
5660                .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
5670                .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
5680                .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
5690                .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
5700                .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
5710                .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
5720                .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
5730                .Fields("CaseNum") = strCaseNum
5740                If rsxDataIn.Fields("pcash") > 0 Then
                      ' ** Does this '1' have any relation to Report 10 (Additional Property Received), or is it just for sort order?
5750                  .Fields("ReportNumber") = 1
5760                  .Fields("ReportCategory") = FLCourtReportCategory(1)  ' ** This'll come back 'Unknown'!
5770                  .Fields("ReportGroup") = FLCourtReportGroup(1)  ' ** This'll come back 0!
5780                  .Fields("ReportDivision") = FLCourtReportDivision(1)  ' ** This'll come back 0!
5790                  .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(1)  ' ** This'll come back 'Unknown'!
5800                  .Fields("ReportDivisionText") = FLCourtReportDivisionText(1)  ' ** This'll come back 'Unknown'!
5810                  .Fields("ReportGroupText") = FLCourtReportGroupText(1)  ' ** This'll come back 'Unknown'!
5820                Else
                      ' ** Does this '3' have any relation to Report 30 (Gains on Sales), or is it just for sort order?
5830                  .Fields("ReportNumber") = 3
5840                  .Fields("ReportCategory") = FLCourtReportCategory(3)  ' ** This'll come back 'Unknown'!
5850                  .Fields("ReportGroup") = FLCourtReportGroup(3)  ' ** This'll come back 0!
5860                  .Fields("ReportDivision") = FLCourtReportDivision(3)  ' ** This'll come back 0!
5870                  .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(3)  ' ** This'll come back 'Unknown'!
5880                  .Fields("ReportDivisionText") = FLCourtReportDivisionText(3)  ' ** This'll come back 'Unknown'!
5890                  .Fields("ReportGroupText") = FLCourtReportGroupText(3)  ' ** This'll come back 'Unknown'!
5900                End If
5910                .Update
5920              End With
5930            Else
                  ' ** Handle journal type special cases here.
5940              With rsxDataOut
5950                Select Case rsxDataIn.Fields("journaltype")

                    Case "Purchase"
                      ' ** Always add the purchase record.
                      ' ** intReportSection is already a 10's number.
5960                  .AddNew
5970                  .Fields("ReportNumber") = intReportSection
5980                  .Fields("ReportCategory") = FLCourtReportCategory(intReportSection)
5990                  .Fields("ReportGroup") = FLCourtReportGroup(intReportSection)
6000                  .Fields("ReportDivision") = FLCourtReportDivision(intReportSection)
6010                  .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(intReportSection)
6020                  .Fields("ReportDivisionText") = FLCourtReportDivisionText(intReportSection)
6030                  .Fields("ReportGroupText") = FLCourtReportGroupText(intReportSection)
6040                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
6050                  .Fields("date") = datTmp01
6060                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
6070                  .Fields("shareface") = rsxDataIn.Fields("shareface")
6080                  .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                        rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
                      ' ** 07/17/2008: Added this special case Purchase (correctly, I hope!), per Rich.
6090                  If rsxDataIn.Fields("icash") <> 0 And rsxDataIn.Fields("pcash") <> 0 Then
6100                    .Fields("Amount") = FLCourtReportDollars(intReportSection, 0@, rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
6110                  Else
6120                    .Fields("Amount") = FLCourtReportDollars(intReportSection, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
6130                  End If
6140                  .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
6150                  .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
6160                  .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
6170                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
6180                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
6190                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
6200                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
6210                  .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
6220                  .Fields("CaseNum") = strCaseNum
6230                  .Update

6240                Case "Sold"
6250                  If rsxDataIn.Fields("icash") > 0 Then
                        ' ** Add Sold into Interest.
                        ' ** intReportSection is already a 10's number.
6260                    .AddNew
6270                    .Fields("ReportNumber") = intReportSection
6280                    .Fields("ReportCategory") = FLCourtReportCategory(intReportSection)
6290                    .Fields("ReportGroup") = FLCourtReportGroup(intReportSection)
6300                    .Fields("ReportDivision") = FLCourtReportDivision(intReportSection)
6310                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(intReportSection)
6320                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(intReportSection)
6330                    .Fields("ReportGroupText") = FLCourtReportGroupText(intReportSection)
6340                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
6350                    .Fields("date") = datTmp01
6360                    .Fields("journaltype") = "Interest"
6370                    .Fields("shareface") = rsxDataIn.Fields("shareface")
6380                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                          rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
6390                    .Fields("Amount") = FLCourtReportDollars(intReportSection, 0, rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
6400                    .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
6410                    .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
6420                    .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
6430                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
6440                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
6450                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
6460                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
6470                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
6480                    .Fields("CaseNum") = strCaseNum
6490                    .Update
6500                  End If
                      ' ** Always add the sold record.
6510                  .AddNew
6520                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
6530                  .Fields("date") = datTmp01
6540                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
6550                  .Fields("shareface") = rsxDataIn.Fields("shareface")
6560                  .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                        rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
6570                  .Fields("Amount") = FLCourtReportDollars(CRPT_GAINS, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                        rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
6580                  .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
6590                  .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
6600                  .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
6610                  Select Case .Fields("Amount")
                      Case Is > 0
6620                    .Fields("ReportNumber") = CRPT_GAINS
6630                    .Fields("ReportCategory") = FLCourtReportCategory(CRPT_GAINS)
6640                    .Fields("ReportGroup") = FLCourtReportGroup(CRPT_GAINS)
6650                    .Fields("ReportDivision") = FLCourtReportDivision(CRPT_GAINS)
6660                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_GAINS)
6670                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_GAINS)
6680                    .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_GAINS)
6690                    .Fields("Amount") = FLCourtReportDollars(CRPT_GAINS, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
6700                  Case Is = 0
6710                    .Fields("ReportNumber") = CRPT_CHANGES
6720                    .Fields("ReportCategory") = FLCourtReportCategory(CRPT_CHANGES)
6730                    .Fields("ReportGroup") = FLCourtReportGroup(CRPT_CHANGES)
6740                    .Fields("ReportDivision") = FLCourtReportDivision(CRPT_CHANGES)
6750                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_CHANGES)
6760                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_CHANGES)
6770                    .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_CHANGES)
                        ' ** Reset the date for this report.
6780                    .Fields("date") = FLCourtReportDate(CRPT_CHANGES, rsxDataIn.Fields("transdate"), rsxDataIn.Fields("assetdate"))
                        ' ** Reset the amount for the report.
6790                    .Fields("Amount") = FLCourtReportDollars(CRPT_CHANGES, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
6800                  Case Is < 0
6810                    .Fields("ReportNumber") = CRPT_LOSSES
6820                    .Fields("ReportCategory") = FLCourtReportCategory(CRPT_LOSSES)
6830                    .Fields("ReportGroup") = FLCourtReportGroup(CRPT_LOSSES)
6840                    .Fields("ReportDivision") = FLCourtReportDivision(60)
6850                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_LOSSES)
6860                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_LOSSES)
6870                    .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_LOSSES)
6880                    .Fields("Amount") = FLCourtReportDollars(CRPT_LOSSES, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
6890                  End Select
6900                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
6910                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
6920                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
6930                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
6940                  .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
6950                  .Fields("CaseNum") = strCaseNum
6960                  .Update

6970                Case "Liability"
6980                  If rsxDataIn.Fields("icash") < 0 Then
                        ' ** Add Icash as a disbursement.
6990                    .AddNew
7000                    .Fields("ReportNumber") = CRPT_DISBURSEMENTS
7010                    .Fields("ReportCategory") = FLCourtReportCategory(CRPT_DISBURSEMENTS)
7020                    .Fields("ReportGroup") = FLCourtReportGroup(CRPT_DISBURSEMENTS)
7030                    .Fields("ReportDivision") = FLCourtReportDivision(CRPT_DISBURSEMENTS)
7040                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_DISBURSEMENTS)
7050                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_DISBURSEMENTS)
7060                    .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_DISBURSEMENTS)
7070                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
7080                    .Fields("date") = datTmp01
7090                    .Fields("journaltype") = "Liability"
7100                    .Fields("shareface") = rsxDataIn.Fields("shareface")
7110                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                          rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
7120                    .Fields("Amount") = rsxDataIn.Fields("icash") * -1
7130                    .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
7140                    .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
7150                    .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
7160                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
7170                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
7180                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
7190                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
7200                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
7210                    .Fields("CaseNum") = strCaseNum
7220                    .Update
7230                  Else
                        ' ** 07/17/2008: Changed the line above from 'End If' to 'Else'. It appeared to be adding it twice, per Rich.
                        ' ** Always add the liability record.
                        ' ** intReportSection is already a 10's number.
7240                    .AddNew
7250                    .Fields("ReportNumber") = intReportSection
7260                    .Fields("ReportCategory") = FLCourtReportCategory(intReportSection)
7270                    .Fields("ReportGroup") = FLCourtReportGroup(intReportSection)
7280                    .Fields("ReportDivision") = FLCourtReportDivision(intReportSection)
7290                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(intReportSection)
7300                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(intReportSection)
7310                    .Fields("ReportGroupText") = FLCourtReportGroupText(intReportSection)
7320                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
7330                    .Fields("date") = datTmp01
7340                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
7350                    .Fields("shareface") = rsxDataIn.Fields("shareface")
7360                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                          rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
7370                    .Fields("Amount") = FLCourtReportDollars(intReportSection, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
7380                    .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
7390                    .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
7400                    .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
7410                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
7420                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
7430                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
7440                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
7450                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
7460                    .Fields("CaseNum") = strCaseNum
7470                    .Update
7480                  End If

7490                Case "Misc."
7500                  If rsxDataIn.Fields("cost") = 0 And rsxDataIn.Fields("icash") = (-(rsxDataIn.Fields("pcash"))) Then
                        ' ** Simple transfer between Income and Principal.
7510                    .AddNew  ' ** Force a Disbursements record.
7520                    .Fields("ReportNumber") = 60
7530                    .Fields("ReportCategory") = FLCourtReportCategory(60)
7540                    .Fields("ReportGroup") = FLCourtReportGroup(60)
7550                    .Fields("ReportDivision") = FLCourtReportDivision(60)
7560                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(60)
7570                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(60)
7580                    .Fields("ReportGroupText") = FLCourtReportGroupText(60)
7590                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
7600                    .Fields("date") = datTmp01
7610                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
7620                    .Fields("shareface") = rsxDataIn.Fields("shareface")
7630                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                          rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
7640                    .Fields("Amount") = FLCourtReportDollars(60, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
7650                    .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
7660                    .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
7670                    .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
7680                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
7690                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
7700                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
7710                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
7720                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
7730                    .Fields("CaseNum") = strCaseNum
7740                    .Update
7750                  Else
7760                    .AddNew
7770                    .Fields("ReportNumber") = intReportSection
7780                    .Fields("ReportCategory") = FLCourtReportCategory(intReportSection)
7790                    .Fields("ReportGroup") = FLCourtReportGroup(intReportSection)
7800                    .Fields("ReportDivision") = FLCourtReportDivision(intReportSection)
7810                    .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(intReportSection)
7820                    .Fields("ReportDivisionText") = FLCourtReportDivisionText(intReportSection)
7830                    .Fields("ReportGroupText") = FLCourtReportGroupText(intReportSection)
7840                    .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
7850                    .Fields("date") = datTmp01
7860                    .Fields("journaltype") = rsxDataIn.Fields("journaltype")
7870                    .Fields("shareface") = rsxDataIn.Fields("shareface")
7880                    .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                          rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
7890                    .Fields("Amount") = FLCourtReportDollars(intReportSection, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                          rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
7900                    .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
7910                    .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
7920                    .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
7930                    .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
7940                    .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
7950                    .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
7960                    .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
7970                    .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
7980                    .Fields("CaseNum") = strCaseNum
7990                    .Update
8000                  End If

8010                Case Else  ' ** All other journaltypes go through here.
                      ' ** intReportSection is already a 10's number.
8020                  .AddNew
8030                  .Fields("ReportNumber") = intReportSection
8040                  .Fields("ReportCategory") = FLCourtReportCategory(intReportSection)
8050                  .Fields("ReportGroup") = FLCourtReportGroup(intReportSection)
8060                  .Fields("ReportDivision") = FLCourtReportDivision(intReportSection)
8070                  .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(intReportSection)
8080                  .Fields("ReportDivisionText") = FLCourtReportDivisionText(intReportSection)
8090                  .Fields("ReportGroupText") = FLCourtReportGroupText(intReportSection)
8100                  .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
8110                  .Fields("date") = datTmp01
8120                  .Fields("journaltype") = rsxDataIn.Fields("journaltype")
8130                  .Fields("shareface") = rsxDataIn.Fields("shareface")
8140                  .Fields("Description") = fncTransactionDesc(rsxDataIn.Fields("RecurringItem"), rsxDataIn.Fields("Description"), _
                        rsxDataIn.Fields("Rate"), rsxDataIn.Fields("Due"), rsxDataIn.Fields("jComment"))
8150                  .Fields("Amount") = FLCourtReportDollars(intReportSection, rsxDataIn.Fields("pcash"), rsxDataIn.Fields("icash"), _
                        rsxDataIn.Fields("cost"), rsxDataIn.Fields("journaltype"))
8160                  .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
8170                  .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
8180                  .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
8190                  .Fields("revcode_ID") = rsxDataIn.Fields("revcode_ID")
8200                  .Fields("revcode_DESC") = rsxDataIn.Fields("revcode_DESC")
8210                  .Fields("revcode_TYPE") = rsxDataIn.Fields("revcode_TYPE")
8220                  .Fields("revcode_SORTORDER") = rsxDataIn.Fields("revcode_SORTORDER")
8230                  .Fields("ReportSchedule") = rsxDataIn.Fields("ReportSchedule")
8240                  .Fields("CaseNum") = strCaseNum
8250                  .Update
8260                End Select
8270              End With
8280            End If
8290          End If
8300        End If
8310        rsxDataIn.MoveNext
8320      Loop
8330      rsxDataIn.Close

          ' ** Get asset information.
8340      rsxDataIn.Open "qryCourtReport_FL_22", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
8350      intReportSection = CRPT_ON_HAND_END
          ' ** Loop through data, processing records for requested account.
8360      Do While rsxDataIn.EOF = False
8370        If Trim(rsxDataIn.Fields("accountno")) = gstrAccountNo Then
8380          With rsxDataOut
                ' ** intReportSection is already a 10's number.
8390            .AddNew
8400            .Fields("ReportNumber") = intReportSection
8410            .Fields("ReportCategory") = FLCourtReportCategory(intReportSection)
8420            .Fields("ReportGroup") = FLCourtReportGroup(intReportSection)
8430            .Fields("ReportDivision") = FLCourtReportDivision(intReportSection)
8440            .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(intReportSection)
8450            .Fields("ReportDivisionText") = FLCourtReportDivisionText(intReportSection)
8460            .Fields("ReportGroupText") = FLCourtReportGroupText(intReportSection)
8470            .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
8480            .Fields("date") = gdatStartDate  ' ** Use this date to get by the filter on the report.
8490            .Fields("journaltype") = "Asset"
8500            .Fields("Description") = rsxDataIn.Fields("totdesc")
                ' Market value  .Fields("Amount") = rsxDataIn.Fields("MarketValueCurrentX") * rsxDataIn.Fields("TotalShareface")
8510            .Fields("Amount") = rsxDataIn.Fields("TotalCost")
8520            .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
8530            .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
8540            .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
8550            .Fields("revcode_ID") = 0
8560            .Fields("revcode_DESC") = "Dummy entry"
8570            .Fields("revcode_TYPE") = 1
8580            .Fields("revcode_SORTORDER") = 0
8590            .Fields("ReportSchedule") = FLCourtReportSchedule(intReportSection)
8600            .Fields("CaseNum") = strCaseNum
8610            .Update
8620          End With
8630        End If
8640        rsxDataIn.MoveNext
8650      Loop

8660      rsxDataIn.Close
8670      rsxDataIn.Open "account", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTableDirect
8680      intReportSection = CRPT_ON_HAND_END
          ' ** Loop through data, processing records for requested account.
8690      Do While rsxDataIn.EOF = False
8700        If Trim(rsxDataIn.Fields("accountno")) = gstrAccountNo Then
8710          With rsxDataOut
                ' ** intReportSection is already a 10's number.
8720            .AddNew
8730            .Fields("ReportNumber") = intReportSection
8740            .Fields("ReportCategory") = FLCourtReportCategory(intReportSection)
8750            .Fields("ReportGroup") = FLCourtReportGroup(intReportSection)
8760            .Fields("ReportDivision") = FLCourtReportDivision(intReportSection)
8770            .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(intReportSection)
8780            .Fields("ReportDivisionText") = FLCourtReportDivisionText(intReportSection)
8790            .Fields("ReportGroupText") = FLCourtReportGroupText(intReportSection)
8800            .Fields("accountno") = Trim(rsxDataIn.Fields("accountno"))
8810            .Fields("date") = gdatStartDate  ' ** Use this date to get by the filter on the report.
8820            .Fields("journaltype") = "Asset"
8830            .Fields("Description") = "Account info"
8840            .Fields("Amount") = rsxDataIn.Fields("pcash") + IIf(IsNull(rsxDataIn.Fields("icash")), 0, rsxDataIn.Fields("icash"))
8850            .Fields("Amount_Inc") = rsxDataIn.Fields("icash")
8860            .Fields("Amount_Prin") = rsxDataIn.Fields("pcash")
8870            .Fields("Amount_Cost") = rsxDataIn.Fields("cost")
8880            .Fields("revcode_ID") = 0
8890            .Fields("revcode_DESC") = "Dummy entry"
8900            .Fields("revcode_TYPE") = 1
8910            .Fields("revcode_SORTORDER") = 0
8920            .Fields("ReportSchedule") = FLCourtReportSchedule(intReportSection)
8930            .Fields("CaseNum") = strCaseNum
8940            .Update
8950          End With
8960        End If
8970        rsxDataIn.MoveNext
8980      Loop

8990      rsxDataOut.Close

9000    End If  ' ** intRetVal

EXITP:
9010    Set rsxDataIn = Nothing
9020    Set rsxDataOut = Nothing
9030    Set qdf = Nothing
9040    Set dbs = Nothing
9050    FLBuildCourtReportData = intRetVal
9060    Exit Function

ERRH:
9070    intRetVal = -9
9080    Select Case ERR.Number
        Case Else
9090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9100    End Select
9110    Resume EXITP

End Function

Public Function FLGetCourtReportData(Optional varType As Variant) As Integer
' ** Called by:
' **   FLBuildCourtReportData(), above
' **   frmRpt_CourtReports_FL.SummaryNew_FL()
' ** Return Codes:
' **   0  Success.
' **  -1  Canceled.
' **  -9  Error.

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "FLGetCourtReportData"

        'Dim rsxDataOut As ADODB.Recordset  ' ** Early binding.
        Dim rsxDataOut As Object            ' ** Late binding.
        Dim frm As Access.Form
        Dim dblNetIncome As Double, dblNetLoss As Double
        Dim dblCashAssets_Beg As Double, dblCashAssets_End As Double
        Dim blnShowForm As Boolean, strSource As String, strControlName As String
        Dim intRetVal As Integer

9210    intRetVal = 0

        ' ** When called from the TAReports CommandBar, these variables should already be filled.
        ' **   gstrCrtRpt_Ordinal
        ' **   gstrCrtRpt_Version
        ' **   gstrCrtRpt_CashAssets_Beg
        ' **   gstrCrtRpt_NetIncome
        ' **   gstrCrtRpt_NetLoss
        ' **   gstrCrtRpt_CashAssets_End

        ' ** Get ordinal and version info.
9220    If gblnCrtRpt_Zero = False Then
          ' ** If report is open in preview, don't call this form.

9230      If IsMissing(varType) = True Then
9240        blnShowForm = True
9250      Else
            ' ** 'SummaryNew' ^ strControlName
            ' ** 'FLBuildCourtReportData' ^ strControlName
            ' ** 'cmdPrintAll_Click' ^ 'cmdPrintAll'
9260        strSource = Left(varType, (InStr(varType, "^") - 1))
9270        strControlName = Mid(varType, (InStr(varType, "^") + 1))
            ' ** If it's 'Word' or 'Excel', then this get's called twice.
            ' ** First time is from 'FLBuildCourtReportData', second time 'SummaryNew'.
            ' ** Suppress the 'SummaryNew' for 'Word' and 'Excel'.
9280        If strControlName = "cmdPrintAll" Or strControlName = "cmdWordAll" Or strControlName = "cmdExcelAll" Then
9290          blnShowForm = True
9300        ElseIf InStr(strControlName, "Word") > 0 Or InStr(strControlName, "Excel") > 0 Then
9310          If strSource = "SummaryNew" Then
9320            blnShowForm = False
9330          Else
9340            blnShowForm = True
9350          End If
9360        Else
9370          If gblnPrintAll = True Then
9380            blnShowForm = False
9390          Else
9400            blnShowForm = True
9410          End If
9420        End If
9430      End If

9440      Set frm = Forms("frmRpt_CourtReports_FL")
9450      If gblnPrintAll = True Then
            ' ** The 2nd window was suppressed at one time, but now it's back.
9460        If strControlName = "cmdWord00_Click" Or strControlName = "cmdExcel00_Click" Then
9470          If IsNull(frm.Ordinal) = False And IsNull(frm.Version) = False And IsNull(frm.CashAssets_Beg) = False Then
9480            blnShowForm = False
9490            If gstrCrtRpt_Ordinal = vbNullString Then gstrCrtRpt_Ordinal = frm.Ordinal
9500            If gstrCrtRpt_Version = vbNullString Then gstrCrtRpt_Version = frm.Version
9510            If gstrCrtRpt_CashAssets_Beg = vbNullString Then gstrCrtRpt_CashAssets_Beg = frm.CashAssets_Beg
9520          End If
9530        End If
9540      End If

9550      If blnShowForm = True Then
9560        DoCmd.Hourglass False
            ' ** Leave these as they are.
            ' **   gstrCrtRpt_Ordinal
            ' **   gstrCrtRpt_Version
9570        gblnMessage = True  ' ** If this returns False, the dialog was canceled.
9580        DoCmd.OpenForm "frmRpt_CourtReports_FL_Input", , , , , acDialog, "frmRpt_CourtReports_FL"
9590        DoEvents
9600        If gblnMessage = False Then
9610          intRetVal = -1  ' ** Canceled.
9620        Else
9630          frm.Ordinal = gstrCrtRpt_Ordinal
9640          frm.Version = gstrCrtRpt_Version
9650          If (strControlName = "cmdPrintAll" Or strControlName = "cmdWordAll" Or strControlName = "cmdExcelAll") Or _
                  (gblnPrintAll = True And (strControlName = "cmdWord00_Click" Or strControlName = "cmdExcel00_Click")) Then
9660            Select Case strControlName
                Case "cmdPrintAll"
9670              frm.cmdPrint00.SetFocus
9680            Case "cmdWordAll"
9690              frm.cmdWord00.SetFocus
9700            Case "cmdExcelAll"
9710              frm.cmdExcel00.SetFocus
9720            End Select
9730            DoEvents
9740          End If
9750        End If
9760        DoEvents
9770      End If

9780    End If

9790    If intRetVal = 0 Then

9800      DoCmd.Hourglass True

9810      If Trim(gstrCrtRpt_CashAssets_Beg) = vbNullString Then gstrCrtRpt_CashAssets_Beg = "0"
9820      If Trim(gstrCrtRpt_NetIncome) = vbNullString Then gstrCrtRpt_NetIncome = "0"
9830      If Trim(gstrCrtRpt_NetLoss) = vbNullString Then gstrCrtRpt_NetLoss = "0"
9840      If Trim(gstrCrtRpt_CashAssets_End) = vbNullString Then gstrCrtRpt_CashAssets_End = "0"

9850      If IsNothing(frm) = True Then  ' ** Module Function: modUtilities.
9860        Set frm = Forms("frmRpt_CourtReports_FL")
9870      End If

          ' ** Get Beginning amounts.
9880      dblCashAssets_Beg = CDbl(gstrCrtRpt_CashAssets_Beg)
9890      frm.CashAssets_Beg = dblCashAssets_Beg
9900      dblCashAssets_End = CDbl(gstrCrtRpt_CashAssets_End)
9910      frm.CashAssets_End = dblCashAssets_End
9920      dblNetIncome = CDbl(gstrCrtRpt_NetIncome)
9930      dblNetLoss = CDbl(gstrCrtRpt_NetLoss)

          'Set rsxDataOut = New ADODB.Recordset             ' ** Early binding.
9940      Set rsxDataOut = CreateObject("ADODB.Recordset")  ' ** Late binding.
9950      rsxDataOut.Open "tmpCourtReportData", CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTable

9960      With rsxDataOut
            ' ** Get the beginning Cash Assets amount for the report.
9970        .AddNew
9980        .Fields("ReportNumber") = CRPT_CASH_BEG
9990        .Fields("ReportCategory") = FLCourtReportCategory(CRPT_CASH_BEG)
10000       .Fields("ReportGroup") = FLCourtReportGroup(CRPT_CASH_BEG)
10010       .Fields("ReportDivision") = FLCourtReportDivision(CRPT_CASH_BEG)
10020       .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_CASH_BEG)
10030       .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_CASH_BEG)
10040       .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_CASH_BEG)
10050       .Fields("accountno") = gstrAccountNo
10060       .Fields("date") = gdatStartDate
10070       .Fields("journaltype") = "Entered"
10080       .Fields("shareface") = 0
10090       .Fields("Description") = "Beginning Cash Assets"
10100       .Fields("Amount") = dblCashAssets_Beg
10110       .Fields("revcode_ID") = 0
10120       .Fields("revcode_DESC") = "Dummy entry"
10130       .Fields("revcode_TYPE") = 1
10140       .Fields("revcode_SORTORDER") = 0
10150       .Fields("ReportSchedule") = FLCourtReportSchedule(CRPT_CASH_BEG)
10160       .Fields("CaseNum") = strCaseNum
10170       .Update
            ' ** Get the Net Income Amount for the report.
10180       .AddNew
10190       .Fields("ReportNumber") = CRPT_NET_INCOME
10200       .Fields("ReportCategory") = FLCourtReportCategory(CRPT_NET_INCOME)
10210       .Fields("ReportGroup") = FLCourtReportGroup(CRPT_NET_INCOME)
10220       .Fields("ReportDivision") = FLCourtReportDivision(CRPT_NET_INCOME)
10230       .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_NET_INCOME)
10240       .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_NET_INCOME)
10250       .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_NET_INCOME)
10260       .Fields("accountno") = gstrAccountNo
10270       .Fields("date") = gdatStartDate
10280       .Fields("journaltype") = "Entered"
10290       .Fields("shareface") = 0
10300       .Fields("Description") = "Net Income"
10310       .Fields("Amount") = dblNetIncome
10320       .Fields("revcode_ID") = 0
10330       .Fields("revcode_DESC") = "Dummy entry"
10340       .Fields("revcode_TYPE") = 1
10350       .Fields("revcode_SORTORDER") = 0
10360       .Fields("ReportSchedule") = FLCourtReportSchedule(CRPT_NET_INCOME)
10370       .Fields("CaseNum") = strCaseNum
10380       .Update
            ' ** Get the Net Loss Amount for the report.
10390       .AddNew
10400       .Fields("ReportNumber") = CRPT_NET_LOSS
10410       .Fields("ReportCategory") = FLCourtReportCategory(CRPT_NET_LOSS)
10420       .Fields("ReportGroup") = FLCourtReportGroup(CRPT_NET_LOSS)
10430       .Fields("ReportDivision") = FLCourtReportDivision(CRPT_NET_LOSS)
10440       .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_NET_LOSS)
10450       .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_NET_LOSS)
10460       .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_NET_LOSS)
10470       .Fields("accountno") = gstrAccountNo
10480       .Fields("date") = gdatStartDate
10490       .Fields("journaltype") = "Entered"
10500       .Fields("shareface") = 0
10510       .Fields("Description") = "Net Loss"
10520       .Fields("Amount") = dblNetLoss
10530       .Fields("revcode_ID") = 0
10540       .Fields("revcode_DESC") = "Dummy entry"
10550       .Fields("revcode_TYPE") = 1
10560       .Fields("revcode_SORTORDER") = 0
10570       .Fields("ReportSchedule") = FLCourtReportSchedule(CRPT_NET_LOSS)
10580       .Fields("CaseNum") = strCaseNum
10590       .Update
            ' ** Get the ending Cash Assets amount for the report.
10600       .AddNew
10610       .Fields("ReportNumber") = CRPT_CASH_END
10620       .Fields("ReportCategory") = FLCourtReportCategory(CRPT_CASH_END)
10630       .Fields("ReportGroup") = FLCourtReportGroup(CRPT_CASH_END)
10640       .Fields("ReportDivision") = FLCourtReportDivision(CRPT_CASH_END)
10650       .Fields("ReportDivisionTitle") = FLCourtReportDivisionTitle(CRPT_CASH_END)
10660       .Fields("ReportDivisionText") = FLCourtReportDivisionText(CRPT_CASH_END)
10670       .Fields("ReportGroupText") = FLCourtReportGroupText(CRPT_CASH_END)
10680       .Fields("accountno") = gstrAccountNo
10690       .Fields("date") = gdatStartDate
10700       .Fields("journaltype") = "Entered"
10710       .Fields("shareface") = 0
10720       .Fields("Description") = "Ending Cash Assets"
10730       .Fields("Amount") = dblCashAssets_End
10740       .Fields("revcode_ID") = 0
10750       .Fields("revcode_DESC") = "Dummy entry"
10760       .Fields("revcode_TYPE") = 1
10770       .Fields("revcode_SORTORDER") = 0
10780       .Fields("ReportSchedule") = FLCourtReportSchedule(CRPT_CASH_END)
10790       .Fields("CaseNum") = strCaseNum
10800       .Update
10810     End With

10820   End If

EXITP:
10830   Set frm = Nothing
10840   Set rsxDataOut = Nothing
10850   FLGetCourtReportData = intRetVal
10860   Exit Function

ERRH:
10870   intRetVal = -9  ' ** Error.
10880   Select Case ERR.Number
        Case 13 ' ** Type mismatch.
10890     MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
10900   Case Else
10910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10920   End Select
10930   Resume EXITP

End Function

Public Sub WordAll_FL(blnAccountChecked As Boolean, frm As Access.Form)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "WordAll_FL"

        Dim strOrd As String, strVer As String
        Dim strCashBeg As String, dblCashBeg As Double
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

11010   With frm
11020     If .Validate = True Then  ' ** Function: Below.

11030       DoCmd.Hourglass True
11040       DoEvents

11050       .cmdWordAll_box01.Visible = True
11060       .cmdWordAll_box02.Visible = True
11070       If .chkAssetList = True Then
11080         .cmdWordAll_box03.Visible = True
11090       End If
11100       .cmdWordAll_box04.Visible = True
11110       DoEvents

11120       blnExcel = False
11130       blnAutoStart = .chkOpenWord
11140       Beep
11150       DoCmd.Hourglass False
11160       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Word" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

11170       If msgResponse = vbOK Then

11180         DoCmd.Hourglass True
11190         DoEvents

11200         blnAllCancel = False
11210         .AllCancelSet1_FL blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_FL.
11220         blnAccountChecked = False
11230         gblnPrintAll = True
11240         strThisProc = "cmdWordAll_Click"

11250         ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
11260         DoEvents

              ' ** Get the Summary inputs first.
11270         intRetVal = FLGetCourtReportData(strThisProc & "^" & "cmdWordAll")  ' ** Function: Above.
11280         If intRetVal <> 0 Then
11290           blnAllCancel = True
11300           .AllCancelSet1_FL blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_FL.
11310         Else
                ' ** Save these for later.
11320           strOrd = gstrCrtRpt_Ordinal
11330           strVer = gstrCrtRpt_Version
11340           strCashBeg = gstrCrtRpt_CashAssets_Beg
11350           dblCashBeg = Nz(Forms("frmRpt_CourtReports_FL").CashAssets_Beg, 0)
11360         End If
11370         DoEvents

11380         If blnAllCancel = False Then
                ' ** Summary of Account.
11390           .cmdWord00.SetFocus
11400           .cmdWord00_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
11410           DoEvents
11420         End If
11430         DoCmd.Hourglass True
11440         DoEvents
11450         If blnAllCancel = False Then
                ' ** Receipts.
11460           .cmdWord01.SetFocus
11470           .cmdWord01_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
11480           DoEvents
11490         End If
11500         If blnAllCancel = False Then
                ' ** Disbursements.
11510           .cmdWord02.SetFocus
11520           .cmdWord02_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
11530           DoEvents
11540         End If
11550         If blnAllCancel = False And .opgType <> .opgType_optGuard.OptionValue Then
                ' ** Distributions.
11560           .cmdWord03.SetFocus
11570           .cmdWord03_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
11580           DoEvents
11590         End If
11600         If blnAllCancel = False Then
                ' ** Capital Transactions and Adjustments.
11610           .cmdWord04.SetFocus
11620           .cmdWord04_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
11630           DoEvents
11640         End If
11650         If blnAllCancel = False Then
                ' ** Assets on Hand at Close of Accounting Period.
11660           .cmdWord05.SetFocus
11670           .cmdWord05_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
11680           DoEvents
11690         End If

11700         DoCmd.Hourglass True
11710         DoEvents

11720         .cmdWordAll.SetFocus

11730         gblnPrintAll = False
11740         Beep

11750         If lngFiles > 0& Then

11760           DoCmd.Hourglass False

11770           strTmp01 = CStr(lngFiles) & " documents were created."
11780           If .chkOpenWord = True Then
11790             strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
11800             msgResponse = MsgBox(strTmp01, vbInformation + vbOKCancel, "Reports Exported")
11810           Else
11820             msgResponse = MsgBox(strTmp01, vbInformation + vbOKOnly, "Reports Exported")
11830           End If

11840           .cmdWordAll_box01.Visible = False
11850           .cmdWordAll_box02.Visible = False
11860           .cmdWordAll_box03.Visible = False
11870           .cmdWordAll_box04.Visible = False

11880           If .chkOpenWord = True And msgResponse = vbOK Then
11890             DoCmd.Hourglass True
11900             DoEvents
11910             For lngX = 0& To (lngFiles - 1&)
11920               strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
11930               OpenExe strDocName  ' ** Module Function: modShellFuncs.
11940               DoEvents
11950               If lngX < (lngFiles - 1&) Then
11960                 ForcePause 2  ' ** Module Function: modCodeUtilities.
11970               End If
11980             Next
11990             Beep
12000           End If

12010         Else
12020           DoCmd.Hourglass False
12030           MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
12040           .cmdWordAll_box01.Visible = False
12050           .cmdWordAll_box02.Visible = False
12060           .cmdWordAll_box03.Visible = False
12070           .cmdWordAll_box04.Visible = False
12080         End If  ' ** lngFiles.

12090       Else
12100         .cmdWordAll_box01.Visible = False
12110         .cmdWordAll_box02.Visible = False
12120         .cmdWordAll_box03.Visible = False
12130         .cmdWordAll_box04.Visible = False
12140       End If  ' ** msgResponse.

12150       DoCmd.Hourglass False
12160     End If  ' ** Validate.
12170   End With

EXITP:
12180   Exit Sub

ERRH:
12190   gblnPrintAll = False
12200   DoCmd.Hourglass False
12210   Select Case ERR.Number
        Case Else
12220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12230   End Select
12240   Resume EXITP

End Sub

Public Sub ExcelAll_FL(blnAccountChecked As Boolean, frm As Access.Form)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "ExcelAll_FL"

        Dim strOrd As String, strVer As String
        Dim strCashBeg As String, dblCashBeg As Double
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

12310   With frm
12320     If .Validate = True Then  ' ** Function: Below.

12330       DoCmd.Hourglass True
12340       DoEvents

12350       .cmdExcelAll_box01.Visible = True
12360       .cmdExcelAll_box02.Visible = True
12370       If .chkAssetList = True Then
12380         .cmdExcelAll_box03.Visible = True
12390       End If
12400       .cmdExcelAll_box04.Visible = True
12410       DoEvents

12420       blnExcel = True
12430       blnAutoStart = .chkOpenExcel
12440       Beep
12450       DoCmd.Hourglass False
12460       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Excel" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

12470       If msgResponse = vbOK Then

12480         DoCmd.Hourglass True
12490         DoEvents

12500         blnAllCancel = False
12510         .AllCancelSet1_FL blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_FL.
12520         blnAccountChecked = False
12530         gblnPrintAll = True
12540         strThisProc = "cmdExcelAll_Click"

12550         ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

12560         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12570           DoCmd.Hourglass False
12580           msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                  "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                  "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                  "You may close Excel before proceding, then click Retry." & vbCrLf & _
                  "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
                ' ** ... Otherwise Trust Accountant will do it for you.
12590           If msgResponse <> vbRetry Then
12600             blnAllCancel = True
12610             .AllCancelSet1_FL blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_FL.
12620           End If
12630         End If

12640         If blnAllCancel = False Then

12650           DoCmd.Hourglass True
12660           DoEvents

12670           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12680             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
12690           End If
12700           DoEvents

                ' ** Get the Summary inputs first.
12710           intRetVal = FLGetCourtReportData(strThisProc & "^" & "cmdExcelAll")  ' ** Module Function: modCourtReportFL.
12720           If intRetVal <> 0 Then
12730             blnAllCancel = True
12740             .AllCancelSet1_FL blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_FL.
12750           Else
                  ' ** Save these for later.
12760             strOrd = gstrCrtRpt_Ordinal
12770             strVer = gstrCrtRpt_Version
12780             strCashBeg = gstrCrtRpt_CashAssets_Beg
12790             dblCashBeg = Nz(Forms("frmRpt_CourtReports_FL").CashAssets_Beg, 0)
12800           End If

12810           If blnAllCancel = False Then
                  ' ** Summary of Account.
12820             .cmdExcel00.SetFocus
12830             .cmdExcel00_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
12840             DoEvents
12850           End If
12860           If blnAllCancel = False Then
                  ' ** Receipts.
12870             .cmdExcel01.SetFocus
12880             .cmdExcel01_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
12890             DoEvents
12900           End If
12910           If blnAllCancel = False Then
                  ' ** Disbursements.
12920             .cmdExcel02.SetFocus
12930             .cmdExcel02_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
12940             DoEvents
12950           End If
12960           If blnAllCancel = False And .opgType <> .opgType_optGuard.OptionValue Then
                  ' ** Distributions.
12970             .cmdExcel03.SetFocus
12980             .cmdExcel03_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
12990             DoEvents
13000           End If
13010           If blnAllCancel = False Then
                  ' ** Capital Transactions and Adjustments.
13020             .cmdExcel04.SetFocus
13030             .cmdExcel04_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
13040             DoEvents
13050           End If
13060           If blnAllCancel = False Then
                  ' ** Assets on Hand at Close of Accounting Period.
13070             .cmdExcel05.SetFocus
13080             .cmdExcel05_Click  ' ** Form Procedure: frmRpt_CourtReports_FL.
13090             DoEvents
13100           End If

13110           DoCmd.Hourglass True
13120           DoEvents

13130           .cmdExcelAll.SetFocus

13140           gblnPrintAll = False
13150           Beep

13160           If lngFiles > 0& Then

13170             DoCmd.Hourglass False

13180             strTmp01 = CStr(lngFiles) & " documents were created."
13190             If .chkOpenExcel = True Then
13200               strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
13210             End If

13220             MsgBox strTmp01, vbInformation + vbOKOnly, "Reports Exported"

13230             .cmdExcelAll_box01.Visible = False
13240             .cmdExcelAll_box02.Visible = False
13250             .cmdExcelAll_box03.Visible = False
13260             .cmdExcelAll_box04.Visible = False

13270             If .chkOpenExcel = True Then
13280               DoCmd.Hourglass True
13290               DoEvents
13300               For lngX = 0& To (lngFiles - 1&)
13310                 strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
13320                 OpenExe strDocName  ' ** Module Function: modShellFuncs.
13330                 DoEvents
13340                 If lngX < (lngFiles - 1&) Then
13350                   ForcePause 2  ' ** Module Function: modCodeUtilities.
13360                 End If
13370               Next
13380             End If

13390           Else
13400             DoCmd.Hourglass False
13410             MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
13420             .cmdExcelAll_box01.Visible = False
13430             .cmdExcelAll_box02.Visible = False
13440             .cmdExcelAll_box03.Visible = False
13450             .cmdExcelAll_box04.Visible = False
13460           End If  ' ** lngFiles.

13470         End If  ' ** blnAllCancel.

13480       Else
13490         .cmdExcelAll_box01.Visible = False
13500         .cmdExcelAll_box02.Visible = False
13510         .cmdExcelAll_box03.Visible = False
13520         .cmdExcelAll_box04.Visible = False
13530       End If  ' ** msgResponse.

13540       DoCmd.Hourglass False
13550     End If  ' ** Validate.
13560   End With

EXITP:
13570   Exit Sub

ERRH:
13580   gblnPrintAll = False
13590   DoCmd.Hourglass False
13600   Select Case ERR.Number
        Case Else
13610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13620   End Select
13630   Resume EXITP

End Sub

Public Sub FileArraySet_FL(arr_varTmp00 As Variant)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "FileArraySet_FL"

13710   arr_varFile = arr_varTmp00
13720   lngFiles = UBound(arr_varFile, 2) + 1&

EXITP:
13730   Exit Sub

ERRH:
13740   Select Case ERR.Number
        Case Else
13750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13760   End Select
13770   Resume EXITP

End Sub

Public Sub AllCancelSet2_FL(blnCancel As Boolean)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet2_FL"

13810   blnAllCancel = blnCancel

EXITP:
13820   Exit Sub

ERRH:
13830   Select Case ERR.Number
        Case Else
13840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13850   End Select
13860   Resume EXITP

End Sub

Public Function SummaryNew_FL(strControlName As String, blnRebuildTable As Boolean, blnIsSummary As Boolean, frm As Access.Form) As Integer
' ** New for Summary, Report 0.
' ** Return Codes:
' **    0  Success.
' **   -1  Canceled.
' **   -2  No data.
' **   -3  Missing entry.
' **   -9  Error.

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "SummaryNew_FL"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnGuardian4Empty As Boolean, blnGuardian6Empty As Boolean
        Dim blnContinue As Boolean
        Dim intRetVal_BuildCourtReportData As Integer, intRetVal_BuildAssetListInfo As Integer
        Dim intRetVal As Integer, blnRetVal As Boolean

13910   intRetVal_BuildCourtReportData = 0: intRetVal_BuildAssetListInfo = 0
13920   blnRetVal = False: intRetVal = 0
13930   blnContinue = True

        '      AssetList, 2, 4, 6, 3, 7
        'Line:      1     2  3  4  5  6

13940   DoCmd.Hourglass True

13950   Set dbs = CurrentDb
13960   With dbs

13970     ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

13980     gdblCrtRpt_IncTot = 0#
13990     gdblCrtRpt_PrinTot = 0#
14000     gdblCrtRpt_CostTot = 0#

          ' ** Empty tmpCourtReportData3.
14010     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_01b")
14020     qdf.Execute
          ' ** Empty tmpCourtReportData4.
14030     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_01c")
14040     qdf.Execute
          ' ** Empty tmpCourtReportData5.
14050     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_01d")
14060     qdf.Execute
          ' ** Empty tmpCourtReportData6.
14070     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_01e")
14080     qdf.Execute

14090     frm.chkAssetList_Start = True
          'THIS IS THE 2ND HIT!
          'DateStart = 01/01/1900
          'DateEnd = 01/01/2009  WHICH GETS CHANGED TO  12/31/2008
          'WHEN CALIFORNIA GOES THROUGH THE 2ND TIME, RETURNING -2, IT CONTINUES ANYWAY,
          'INVOKING A PRINT ROUTINE ESPECIALLY FOR THAT SITUATION. THAT IS, THE ONLY
          'TIME THAT PARTICULAR OpenReport GETS HIT IS IF THERE IS NO PREVIOUS DATA!
14100     gvarCrtRpt_FL_SpecData = CLng(1)
14110     intRetVal_BuildAssetListInfo = frm.BuildAssetListInfo_FL(#1/1/1900#, (CDate(frm.DateStart) - 1), "Beginning", strControlName, THIS_PROC & "1")  ' ** Form Function: frmRpt_CourtReports_FL.
          ' ** Return Codes:
          ' **    0  Success.
          ' **   -2  No data.
          ' **   -3  Missing entry.
          ' **   -9  Error.
          ' ** Beginning Balance.
14120     If intRetVal_BuildAssetListInfo = 0 Then

14130       DoCmd.Hourglass False
14140       intRetVal_BuildCourtReportData = FLGetCourtReportData(THIS_PROC & "^" & strControlName)  ' ** Function: Above.
14150       DoCmd.Hourglass True
14160       If intRetVal_BuildCourtReportData = 0 Then

14170         If gstrCrtRpt_CashAssets_Beg = vbNullString And Val(Forms("frmRpt_CourtReports_FL").CashAssets_Beg) <> 0 Then
                ' ** I don't know where it's losing this!
14180           gstrCrtRpt_CashAssets_Beg = CStr(Forms("frmRpt_CourtReports_FL").CashAssets_Beg)
14190         End If

14200         Set qdf = .QueryDefs("qryCourtReport_FL_00_New_01_03")
14210         Set rst = qdf.OpenRecordset
14220         If rst.BOF = True And rst.EOF = True Then
14230           rst.Close
                ' ** Append an empty record to tmpCourtReportData3.
14240           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_02a")
14250           qdf.Execute
14260         Else
14270           rst.Close
                ' ** 1. Starting Balance: Append qryCourtReport_FL_00_New_01_03 to tmpCourtReportData3.
14280           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_02")  ' ** Now includes gstrCrtRpt_CashAssets_Beg.
14290           qdf.Execute dbFailOnError
                'AT THIS STAGE, THE VALUES IN SECTION I. ARE TODAY'S TOTALS, THAT IS, ENDING VALUES THROUGH TODAY.
                'THAT MEANS THAT THEY ARE NEVER THE CORRECT VALUES FOR A STARTING BALANCE,
                'AND THAT THEIR ONLY USE IS FOR CALCULATING OTHER THINGS.
14300         End If
14310       Else
              ' ** Return Codes:
              ' **   0  Success.
              ' **  -1  Canceled.
              ' **  -9  Error.
14320         blnContinue = False
14330         blnAllCancel = False
14340         AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
14350       End If
14360       frm.chkAssetList_Start = False

14370     ElseIf intRetVal_BuildAssetListInfo = -2 Then
            ' ** A new account may indeed have no beginning balance!

14380       If gstrCrtRpt_CashAssets_Beg = vbNullString Then 'And Val(Forms("frmRpt_CourtReports_FL").CashAssets_Beg) <> 0 Then
14390         gstrCrtRpt_CashAssets_Beg = "0"
14400       End If
14410       If IsNull(Forms("frmRpt_CourtReports_FL").CashAssets_Beg) = True Then
14420         Forms("frmRpt_CourtReports_FL").CashAssets_Beg = 0
14430       End If

            ' ** Append an empty record to tmpCourtReportData3.
14440       Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_02a")
14450       qdf.Execute

14460     Else
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry.
            ' **   -9  Error.
14470       blnContinue = False
14480       blnAllCancel = True
14490       AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
14500     End If

          ' ** Receipts.
14510     If blnContinue = True And blnAllCancel = False Then
14520       intRetVal_BuildAssetListInfo = frm.BuildAssetListInfo_FL(frm.DateStart, frm.DateEnd, "Ending", strControlName, THIS_PROC & "2")  ' ** Form Function: frmRpt_CourtReports_FL.
14530       If intRetVal_BuildAssetListInfo = 0 Then
14540         intRetVal_BuildCourtReportData = FLBuildCourtReportData("2", strControlName)  ' ** Function: Above.
              ' ** Return Codes:
              ' **   0  Success.
              ' **  -1  Canceled.
              ' **  -9  Error.
14550         If intRetVal_BuildCourtReportData = 0 Then
14560           blnRebuildTable = False
                ' ** Empty tmpCourtReportData2.
14570           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_01a")
14580           qdf.Execute
                ' ** Append qryCourtReport_FL_02m to tmpCourtReportData2.
14590           Set qdf = .QueryDefs("qryCourtReport_FL_02mb")
14600           qdf.Execute
                ' ** Append qryCourtReport_FL_02oa to tmpCourtReportData2.
14610           Set qdf = .QueryDefs("qryCourtReport_FL_02ob")
14620           qdf.Execute
                ' ** Append qryCourtReport_FL_02on to tmpCourtReportData2.
14630           Set qdf = .QueryDefs("qryCourtReport_FL_02pb")
14640           qdf.Execute
14650           Select Case frm.chkGroupBy_IncExpCode
                Case True
                  ' ** qryCourtReport_FL_00_New_02_02g, grouped and summed, for group by Inc/Exp; total.
14660             Set qdf = .QueryDefs("qryCourtReport_FL_00_New_02_02h")
14670           Case False
                  ' ** qryCourtReport_FL_00_New_02_01, grouped and summed, with Income_Tot, Principal_Tot, Cost_Tot.
14680             Set qdf = .QueryDefs("qryCourtReport_FL_00_New_02_02")
14690           End Select
14700           Set rst = qdf.OpenRecordset
14710           If rst.BOF = True And rst.EOF = True Then
14720             rst.Close
                  ' ** Append an empty record to tmpCourtReportData3.
14730             Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_03a")
14740             qdf.Execute
14750           Else
14760             rst.Close
14770             Select Case frm.chkGroupBy_IncExpCode
                  Case True
                    ' ** 2. Receipts: Append qryCourtReport_FL_00_New_02_02g to tmpCourtReportData3, for group by Inc/Exp; items.
14780               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_03g")
14790               qdf.Execute
                    ' ** 2. Receipts: Append qryCourtReport_FL_00_New_02_02h to tmpCourtReportData3, for group by Inc/Exp; total.
14800               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_03h")
14810             Case False
                    ' ** 2. Receipts: Append qryCourtReport_FL_00_New_02_02 to tmpCourtReportData3.
14820               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_03")
14830             End Select
14840             qdf.Execute
14850           End If
14860         Else
                ' ** Return Codes:
                ' **   0  Success.
                ' **  -1  Canceled.
                ' **  -9  Error.
14870           blnContinue = False
14880           blnAllCancel = True
14890           AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
14900         End If
14910       Else
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry.
              ' **   -9  Error.
14920         blnContinue = False
14930         blnAllCancel = True
14940         AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
14950       End If
14960     End If

14970     If blnContinue = True And blnAllCancel = False Then

14980       Select Case frm.opgType
            Case frm.opgType_optRep.OptionValue

              ' ** Disbursements.
14990         intRetVal_BuildAssetListInfo = frm.BuildAssetListInfo_FL(frm.DateStart, frm.DateEnd, "Ending", strControlName, THIS_PROC & "3")  ' ** Form Function: frmRpt_CourtReports_FL.
15000         If intRetVal_BuildAssetListInfo = 0 Then
15010           intRetVal_BuildCourtReportData = FLBuildCourtReportData("4", strControlName)  ' ** Function: Above.
15020           If intRetVal_BuildCourtReportData = 0 Then
15030             blnRebuildTable = False
15040             Select Case frm.chkGroupBy_IncExpCode
                  Case True
                    ' ** qryCourtReport_FL_00_New_03_02b, grouped and summed; for group by Inc/Exp.
15050               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_03_02c")
15060             Case False
                    ' ** qryCourtReport_FL_00_New_03_01, grouped and summed, with Income_Tot, Principal_Tot, Cost_Tot; 'III', 'B'.
15070               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_03_02")
15080             End Select
15090             Set rst = qdf.OpenRecordset
15100             If rst.BOF = True And rst.EOF = True Then
                    ' ** Append an empty record to tmpCourtReportData3.
15110               rst.Close
15120               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04a")
15130               qdf.Execute
15140             Else
15150               rst.Close
15160               Select Case frm.chkGroupBy_IncExpCode
                    Case True
                      ' ** 3. Disbursements: append qryCourtReport_FL_00_New_03_02b to tmpCourtReportData3; w/ revcode_SORTORDER, .._DESC; items.
15170                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04c")
15180                 qdf.Execute
                      ' ** 3. Disbursements: append qryCourtReport_FL_00_New_03_02c to tmpCourtReportData3; w/ revcode_SORTORDER, .._DESC; total.
15190                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04d")
15200               Case False
                      ' ** 3. Disbursements: Append qryCourtReport_FL_00_New_03_02 to tmpCourtReportData3, by Rep.
15210                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04")
15220               End Select
15230               qdf.Execute
15240             End If
15250           Else
                  ' ** Return Codes:
                  ' **   0  Success.
                  ' **  -1  Canceled.
                  ' **  -9  Error.
15260             blnContinue = False
15270             blnAllCancel = True
15280             AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
15290           End If
15300         Else
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry.
                ' **   -9  Error.
15310           blnContinue = False
15320           blnAllCancel = True
15330           AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
15340         End If

              ' ** Distributions.
15350         If blnContinue = True And blnAllCancel = False Then
15360           intRetVal_BuildAssetListInfo = frm.BuildAssetListInfo_FL(frm.DateStart, frm.DateEnd, "Ending", strControlName, THIS_PROC & "4")  ' ** Form Function: frmRpt_CourtReports_FL.
15370           If intRetVal_BuildAssetListInfo = 0 Then
15380             If blnIsSummary = True Then
15390               intRetVal_BuildCourtReportData = 0
15400             Else
15410               intRetVal_BuildCourtReportData = FLBuildCourtReportData("6", strControlName)  ' ** Function: Above.
15420             End If

15430             If intRetVal_BuildCourtReportData = 0 Then
15440               blnRebuildTable = False
                    ' ** qryCourtReport_FL_00_New_04_01, grouped and summed, with Income_Tot, Principal_Tot, Cost_Tot; 'IV', 'C'.
15450               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_04_02")
15460               Set rst = qdf.OpenRecordset
15470               If rst.BOF = True And rst.EOF = True Then
15480                 rst.Close
                      ' ** Append an empty record to tmpCourtReportData3.
15490                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05a")
15500                 qdf.Execute
15510               Else
15520                 rst.Close
                      ' ** 4. Distributions: Append qryCourtReport_FL_00_New_04_02 to tmpCourtReportData3.
15530                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05")
15540                 qdf.Execute
15550               End If
15560             Else
                    ' ** Return Codes:
                    ' **   0  Success.
                    ' **  -1  Canceled.
                    ' **  -9  Error.
15570               blnContinue = False
15580               blnAllCancel = True
15590               AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
15600             End If
15610           Else
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -3  Missing entry.
                  ' **   -9  Error.
15620             blnContinue = False
15630             blnAllCancel = True
15640             AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
15650           End If
15660         End If

15670       Case frm.opgType_optGuard.OptionValue
              ' ** Combine Disbursements (4) and Distributions (6) into one.

15680         blnGuardian4Empty = False: blnGuardian6Empty = False

15690         intRetVal_BuildAssetListInfo = frm.BuildAssetListInfo_FL(frm.DateStart, frm.DateEnd, "Ending", strControlName, THIS_PROC & "5")  ' ** Form Function: frmRpt_CourtReports_FL.
15700         If intRetVal_BuildAssetListInfo = 0 Then

                ' ** Disbursements.
15710           intRetVal_BuildCourtReportData = FLBuildCourtReportData("4", strControlName)  ' ** Function: Above.
15720           If intRetVal_BuildCourtReportData = 0 Then
15730             blnRebuildTable = False
15740             Select Case frm.chkGroupBy_IncExpCode
                  Case True
                    ' ** qryCourtReport_FL_00_New_03_02b, grouped and summed; for group by Inc/Exp.
15750               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_03_02c")
15760             Case False
                    ' ** qryCourtReport_FL_00_New_03_01, grouped and summed, with Income_Tot, Principal_Tot, Cost_Tot; 'III', 'B'.
15770               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_03_02")
15780             End Select
15790             Set rst = qdf.OpenRecordset
15800             If rst.BOF = True And rst.EOF = True Then
15810               rst.Close
15820               blnGuardian4Empty = True
15830             Else
15840               rst.Close
15850               Select Case frm.chkGroupBy_IncExpCode
                    Case True
                      ' ** Append qryCourtReport_FL_00_New_03_02b to tmpCourtReportData4; w/ revcode_SORTORDER, .._DESC; items.
15860                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04k")
15870                 qdf.Execute
                      ' ** Append qryCourtReport_FL_00_New_03_02c to tmpCourtReportData4; w/ revcode_SORTORDER, .._DESC; total.
15880                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04l")
15890               Case False
                      ' ** Append qryCourtReport_FL_00_New_03_02 to tmpCourtReportData4.
15900                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04b")
15910               End Select
15920               qdf.Execute
15930             End If

                  ' ** Distributions.
15940             If blnIsSummary = True Then
15950               intRetVal_BuildCourtReportData = 0
15960             Else
15970               intRetVal_BuildCourtReportData = FLBuildCourtReportData("6", strControlName)  ' ** Function: Above.
15980             End If

15990             If intRetVal_BuildCourtReportData = 0 Then
16000               blnRebuildTable = False
16010               Select Case frm.chkGroupBy_IncExpCode
                    Case True
                      ' ** qryCourtReport_FL_00_New_04_02b, grouped and summed; for group by Inc/Exp; Guardian.
16020                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_04_02c")
16030               Case False
                      ' ** qryCourtReport_FL_00_New_04_01, grouped and summed, with Income_Tot, Principal_Tot, Cost_Tot; 'IV', 'C'.
16040                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_04_02")
16050               End Select
16060               Set rst = qdf.OpenRecordset
16070               If rst.BOF = True And rst.EOF = True Then
                      ' ** Append an empty record to tmpCourtReportData3.
16080                 rst.Close
16090                 blnGuardian6Empty = True
16100               Else
16110                 rst.Close
16120                 Select Case frm.chkGroupBy_IncExpCode
                      Case True
                        ' ** Append qryCourtReport_FL_00_New_04_02b to tmpCourtReportData4; w/ revcode_SORTORDER, .._DESC; items; Guardian.
16130                   Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05e")
16140                   qdf.Execute
                        ' ** Append qryCourtReport_FL_00_New_04_02c to tmpCourtReportData4; w/ revcode_SORTORDER, .._DESC; total; Guardian.
16150                   Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05f")
16160                 Case False
                        ' ** Append qryCourtReport_FL_00_New_04_02 to tmpCourtReportData4.
16170                   Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05b")
16180                 End Select
16190                 qdf.Execute
16200               End If
16210               If blnGuardian4Empty = True And blnGuardian6Empty = True Then
                      ' ** Append an empty record to tmpCourtReportData3.
16220                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05g")
16230                 qdf.Execute
16240               Else
16250                 If blnGuardian4Empty = True Then
16260                   Select Case frm.chkGroupBy_IncExpCode
                        Case True
                          ' ** 4. Disbributions: Append qryCourtReport_FL_00_New_04_02i to tmpCourtReportData3; Guardian, Items.
16270                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05m")
16280                     qdf.Execute
                          ' ** 4. Disbributions: Append qryCourtReport_FL_00_New_04_02j to tmpCourtReportData3; Guardian, Totals.
16290                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05n")
16300                   Case False
                          ' ** 4. Disbributions: Append qryCourtReport_FL_00_New_04_02 to tmpCourtReportData3.
16310                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05h")
16320                   End Select
16330                   qdf.Execute
16340                 ElseIf blnGuardian6Empty = True Then
16350                   Select Case frm.chkGroupBy_IncExpCode
                        Case True
                          ' ** 3. Disbursements: Append qryCourtReport_FL_00_New_03_02i to tmpCourtReportData3; items; Guardian.
16360                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04i")
16370                     qdf.Execute
                          ' ** 3. Disbursements: Append qryCourtReport_FL_00_New_03_02j to tmpCourtReportData3; total; Guardian.
16380                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04j")
16390                   Case False
                          ' ** 3. Disbursements: Append qryCourtReport_FL_00_New_03_02 to tmpCourtReportData3.
16400                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_04h")
16410                   End Select
16420                   qdf.Execute
16430                 Else
                        ' ** Both records are now in tmpCourtReportData4.
16440                   Select Case frm.chkGroupBy_IncExpCode
                        Case True
                          ' ** 3./4. Disbursements and Distributions, Rev/Exp: Append qryCourtReport_FL_00_New_07_05k to tmpCourtReportData3; Guardian.
16450                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05l")
16460                   Case False
                          ' ** 3./4. Disbributions and Disbursements: Append qryCourtReport_FL_00_New_07_05i to tmpCourtReportData3.
16470                     Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_05j")
16480                   End Select
16490                   qdf.Execute
16500                 End If
16510               End If
16520             Else
                    ' ** Return Codes:
                    ' **   0  Success.
                    ' **  -1  Canceled.
                    ' **  -9  Error.
16530               blnContinue = False
16540               blnAllCancel = True
16550               AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
16560             End If

16570           Else
                  ' ** Return Codes:
                  ' **   0  Success.
                  ' **  -1  Canceled.
                  ' **  -9  Error.
16580             blnContinue = False
16590             blnAllCancel = True
16600             AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
16610           End If
16620         Else
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry.
                ' **   -9  Error.
16630           blnContinue = False
16640           blnAllCancel = True
16650           AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
16660         End If

16670       End Select

16680     End If  ' ** blnContinue.

          ' ** Capital Transactions and Adjustments.
16690     If blnContinue = True And blnAllCancel = False Then
16700       intRetVal_BuildAssetListInfo = frm.BuildAssetListInfo_FL(frm.DateStart, frm.DateEnd, "Ending", strControlName, THIS_PROC & "6")  ' ** Form Function: frmRpt_CourtReports_FL.
16710       If intRetVal_BuildAssetListInfo = 0 Then

16720         If blnIsSummary = True Then
16730           intRetVal_BuildCourtReportData = 0
16740         Else
16750           intRetVal_BuildCourtReportData = FLBuildCourtReportData("3", strControlName)  ' ** Function: Above.
16760         End If

16770         If intRetVal_BuildCourtReportData = 0 Then
16780           blnRebuildTable = False
16790           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_05_03")
16800           Set rst = qdf.OpenRecordset
16810           If rst.BOF = True And rst.EOF = True Then
16820             rst.Close
                  ' ** Append an empty record to tmpCourtReportData3.
16830             Select Case frm.opgType
                  Case frm.opgType_optRep.OptionValue
16840               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_06a")  ' ** 'V'.
16850             Case frm.opgType_optGuard.OptionValue
16860               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_06b")  ' ** 'IV'.
16870             End Select
16880             qdf.Execute
16890           Else
16900             rst.Close
                  ' ** 5. Cap Trans: Append qryCourtReport_FL_New_00_05_03 to tmpCourtReportData3.
16910             Select Case frm.opgType
                  Case frm.opgType_optRep.OptionValue
16920               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_06")  ' ** 'V'.
16930             Case frm.opgType_optGuard.OptionValue
16940               Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_06c")  ' ** 'IV'.
16950             End Select
16960             qdf.Execute
16970           End If
16980         Else
                ' ** Return Codes:
                ' **   0  Success.
                ' **  -1  Canceled.
                ' **  -9  Error.
16990           blnContinue = False
17000           blnAllCancel = True
17010           AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
17020         End If
17030       Else
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry.
              ' **   -9  Error.
17040         blnContinue = False
17050         blnAllCancel = True
17060         AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
17070       End If
17080     End If

          ' ** Assets on Hand at Close of Accounting Period.
17090     If blnContinue = True And blnAllCancel = False Then
17100       intRetVal_BuildAssetListInfo = frm.BuildAssetListInfo_FL(frm.DateStart, frm.DateEnd, "Ending", strControlName, THIS_PROC & "7")  ' ** Form Function: frmRpt_CourtReports_FL.
17110       If intRetVal_BuildAssetListInfo = 0 Then
17120         Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_07")
17130         Set rst = qdf.OpenRecordset
17140         If rst.BOF = True And rst.EOF = True Then
17150           rst.Close
                ' ** Append an empty record to tmpCourtReportData3.
17160           Select Case frm.opgType
                Case frm.opgType_optRep.OptionValue
17170             Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_10a")  ' ** 'VI'.
17180           Case frm.opgType_optGuard.OptionValue
17190             Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_10b")  ' ** 'V'.
17200           End Select
17210           qdf.Execute
17220         Else
17230           rst.Close
                ' ** 6. First 5 lines: Append qryCourtReport_FL_00_New_07_07 to tmpCourtReportData3.
17240           Select Case frm.opgType
                Case frm.opgType_optRep.OptionValue
17250             Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_07b")  ' ** 'VI'.
17260           Case frm.opgType_optGuard.OptionValue
17270             Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_07c")  ' ** 'V'.
17280           End Select
17290           qdf.Execute
17300         End If
17310       Else
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry.
              ' **   -9  Error.
17320         blnContinue = False
17330         blnAllCancel = True
17340         AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
17350       End If
17360     End If

17370     .Close
17380   End With

17390   If blnContinue = True And blnAllCancel = False Then
17400     If gstrCrtRpt_CashAssets_Beg = vbNullString And Val(Forms("frmRpt_CourtReports_FL").CashAssets_Beg) <> 0 Then
            ' ** I don't know where it's losing this!
17410       gstrCrtRpt_CashAssets_Beg = CStr(Forms("frmRpt_CourtReports_FL").CashAssets_Beg)
17420     End If
17430   End If

        ' ** Problems may come from either intRetVal_BuildAssetListInfo or intRetVal_BuildCourtReportData.
17440   If intRetVal_BuildCourtReportData <> 0 Or intRetVal_BuildAssetListInfo <> 0 Then
17450     If intRetVal_BuildCourtReportData = -1 Or intRetVal_BuildAssetListInfo = -3 Then
            ' ** Canceled or Missing Entry get priority.
17460       intRetVal = IIf(intRetVal_BuildCourtReportData <> -1, intRetVal_BuildAssetListInfo, intRetVal_BuildCourtReportData)
17470     Else
17480       If intRetVal_BuildAssetListInfo = -9 Or intRetVal_BuildCourtReportData = -9 Then
              ' ** Next in line are Errors.
17490         intRetVal = IIf(intRetVal_BuildCourtReportData <> -9, intRetVal_BuildAssetListInfo, intRetVal_BuildCourtReportData)
17500       Else
              ' ** All that's left is No Data from BuildAssetListInfo_FL().
17510         intRetVal = intRetVal_BuildAssetListInfo
17520       End If
17530     End If
17540   End If

17550   DoCmd.Hourglass False

EXITP:
        ' ** Check some Public variables in case there's been an error.
17560   blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
17570   blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
17580   Set rst = Nothing
17590   Set qdf = Nothing
17600   Set dbs = Nothing
17610   SummaryNew_FL = intRetVal
17620   Exit Function

ERRH:
17630   intRetVal = -9  ' ** Error.
17640   Select Case ERR.Number
        Case Else
17650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17660   End Select
17670   Resume EXITP

End Function

Public Sub SendToFile_FL(strReportNumber As String, strProc As String, blnRebuildTable As Boolean, blnIsSummary As Boolean, frm As Access.Form, Optional varExcel As Variant)

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "SendToFile_FL"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strQry As String, strMacro As String
        Dim strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim strOrd As String, strVer As String, strCashBeg As String, dblCashBeg As Double
        Dim blnExcel As Boolean, blnNoTmp2 As Boolean, blnNoData As Boolean
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String, curTmp02 As Currency, curTmp03 As Currency
        Dim curTmp04 As Currency, curTmp05 As Currency, curTmp06 As Currency, curTmp07 As Currency
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean
        Dim intRetVal_SummaryNew As Integer
        Dim intRetVal_BuildCourtReportData As Integer, intRetVal_BuildAssetListInfo As Integer

17710   blnContinue = True
17720   blnUseSavedPath = False

17730   With frm

17740     DoCmd.Hourglass True
17750     DoEvents

17760     If IsMissing(varExcel) = True Then
17770       blnExcel = False
17780     Else
17790       blnExcel = varExcel
17800     End If

17810     If blnExcel = True Then
17820       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
17830         DoCmd.Hourglass False
17840         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
17850         If msgResponse <> vbRetry Then
17860           blnContinue = False
17870           blnAllCancel = True
17880           AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
17890         End If
17900       End If
17910     End If

17920     If blnContinue = True Then

17930       DoCmd.Hourglass True
17940       DoEvents

17950       intRetVal_BuildCourtReportData = 0: intRetVal_BuildAssetListInfo = 0: intRetVal_SummaryNew = 0
17960       blnNoTmp2 = False

            ' ** Validate the dates.
17970       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_FL.

17980         DoEvents
17990         ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
18000         DoEvents

              ' ** Set global variables for report headers.
18010         gdatStartDate = .DateStart.Value
18020         gdatEndDate = .DateEnd.Value
18030         gstrAccountNo = .cmbAccounts.Column(0)
18040         gstrAccountName = .cmbAccounts.Column(3)

18050         If .chkGroupBy_IncExpCode = True Then
18060           gblnUseReveuneExpenseCodes = True
18070         Else
18080           gblnUseReveuneExpenseCodes = False
18090         End If

18100         Set dbs = CurrentDb
18110         With dbs
                ' ** qryCourtReport_15a (tblReport, captions of Court Reports, with rpt_caption_newx),
                ' ** by specified GlobalVarGet(), FormRef('TypeShort'), [CrtTyp].
18120           Set qdf = .QueryDefs("qryCourtReport_15")
18130           With qdf.Parameters
18140             ![CrtTyp] = "FL"
18150           End With
18160           Set rst = qdf.OpenRecordset
18170           With rst
18180             .MoveLast
18190             lngCaps = .RecordCount
18200             .MoveFirst
18210             arr_varCap = .GetRows(lngCaps)
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
18220             .Close
18230           End With
18240           .Close
18250         End With
18260         DoEvents

18270         Set dbs = CurrentDb
              ' ** Empty tmpCourtReportData3.
18280         Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01b")
18290         qdf.Execute
              ' ** Empty tmpCourtReportData4.
18300         Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01c")
18310         qdf.Execute
              ' ** Empty tmpCourtReportData5.
18320         Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01d")
18330         qdf.Execute
              ' ** Empty tmpCourtReportData6.
18340         Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01e")
18350         qdf.Execute
18360         dbs.Close
18370         DoEvents

18380         If IsNull(.UserReportPath) = False Then
18390           If .UserReportPath <> vbNullString Then
18400             If .UserReportPath_chk = True Then
18410               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
18420                 blnUseSavedPath = True
18430               End If
18440             End If
18450           End If
18460         End If
18470         DoEvents

18480         If strReportNumber <> "7" Then
                ' ** Build a new summary report table for everything except
                ' ** report 7, the property on hand report.
18490           intRetVal_BuildCourtReportData = FLBuildCourtReportData(strReportNumber, strProc)  ' ** Function: Above.
18500           DoEvents
18510           If intRetVal_BuildCourtReportData = 0 Then
18520             blnRebuildTable = False

                  ' ** Save these for later.
18530             strOrd = gstrCrtRpt_Ordinal
18540             strVer = gstrCrtRpt_Version
18550             strCashBeg = gstrCrtRpt_CashAssets_Beg
18560             dblCashBeg = Nz(Forms("frmRpt_CourtReports_FL").CashAssets_Beg, 0)

18570             If strReportNumber = "0" Then
18580               intRetVal_BuildAssetListInfo = .BuildAssetListInfo_FL(.DateStart, .DateEnd, "Ending", strProc)  ' ** Form Function: frmRpt_CourtReports_Fl.
18590               Set dbs = CurrentDb
18600               With dbs
18610                 Set qdf = .QueryDefs("qryCourtReport_FL_00_New_10_51")
18620                 Set rst = qdf.OpenRecordset
18630                 With rst
18640                   .MoveFirst
18650                   curTmp05 = ![ICash]
18660                   curTmp06 = ![PCash]
18670                   curTmp07 = ![TotalCost]
18680                   .Close
18690                 End With
18700                 .Close
18710               End With
18720               intRetVal_BuildAssetListInfo = .BuildAssetListInfo_FL(.DateStart, .DateEnd, "Ending", strProc, "SummaryNew0")  ' ** Form Function: frmRpt_CourtReports_Fl.
18730             Else
18740               intRetVal_BuildAssetListInfo = .BuildAssetListInfo_FL(.DateStart, .DateEnd, "Ending", strProc)  ' ** Form Function: frmRpt_CourtReports_Fl.
18750             End If
18760             DoEvents

18770             If intRetVal_BuildAssetListInfo = 0 Then
                    ' ** Property On Hand = Sum([TotalCost])+IIf(IsNull([icash]),0,[icash])+[pcash].

                    ' ** Disbursements.
18780               If Left(strReportNumber, 1) = "4" Then
18790                 Set dbs = CurrentDb
18800                 Select Case .opgType
                      Case .opgType_optRep.OptionValue
18810                   intRetVal_BuildCourtReportData = FLBuildCourtReportData(strReportNumber, strProc)  ' ** Function: Above.
18820                   DoEvents
18830                 Case .opgType_optGuard.OptionValue
18840                   intRetVal_BuildCourtReportData = FLBuildCourtReportData("4", strProc)  ' ** Function: Above.
18850                   DoEvents
18860                   If intRetVal_BuildCourtReportData = 0 Then
                          ' ** Empty tmpCourtReportData2.
18870                     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01a")
18880                     qdf.Execute
                          ' ** Append qryCourtReport_FL_04a to tmpCourtReportData2 (for rptCourtRptFL_04).
18890                     Set qdf = dbs.QueryDefs("qryCourtReport_FL_04f")
18900                     qdf.Execute
18910                     DoEvents
18920                     intRetVal_BuildCourtReportData = FLBuildCourtReportData("6", strProc)  ' ** Function: Above.
18930                     DoEvents
18940                     If intRetVal_BuildCourtReportData = 0 Then
                            ' ** Append qryCourtReport_FL_06m to tmpCourtReportData2 (for rptCourtRptFL_06).
18950                       Set qdf = dbs.QueryDefs("qryCourtReport_FL_04g")
18960                       qdf.Execute
18970                     Else
18980                       blnContinue = False
18990                       blnAllCancel = True
19000                       AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
19010                     End If
19020                   Else
19030                     blnContinue = False
19040                     blnAllCancel = True
19050                     AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
19060                   End If
19070                 End Select
19080                 dbs.Close
19090               End If
19100               DoEvents

                    ' ** Capital Transactions and Adjustments.
19110               If strReportNumber = "3" Then
19120                 Set dbs = CurrentDb
                      ' ** Append report title to tmpCourtReportData5.
19130                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03da")
19140                 qdf.Execute
19150                 DoEvents
                      ' ** Append report period to tmpCourtReportData5.
19160                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03ea")
19170                 qdf.Execute
19180                 DoEvents
                      ' ** Append report author to tmpCourtReportData5.
19190                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03ec")
19200                 qdf.Execute
19210                 DoEvents
                      ' ** Append qryCourtReport_FL_03f, w/column names, to tmpCourtReportData5.
19220                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03fa")
19230                 qdf.Execute
19240                 DoEvents
                      ' ** Append qryCourtReport_FL_03g, grouped and summed, to tmpCourtReportData5.
19250                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03ga")
19260                 qdf.Execute
19270                 DoEvents
                      ' ** Append qryCourtReport_FL_03h, w/ Net Gain/Loss, to tmpCourtReportData5.
19280                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03ha")
19290                 qdf.Execute
19300                 DoEvents
                      ' ** Append qryCourtReport_FL_03i, grouped and summed, to tmpCourtReportData5.
19310                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03ia")
19320                 qdf.Execute
19330                 DoEvents
                      ' ** Append qryCourtReport_FL_03ib, Net Gain/Loss, to tmpCourtReportData5.
19340                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_03ic")
19350                 qdf.Execute
19360                 dbs.Close
19370               End If
19380               DoEvents

19390               If intRetVal_BuildCourtReportData = 0 Then
19400                 blnRebuildTable = False
19410                 If strReportNumber = "0" Then  'Or strReportNumber = "0A" Then
                        ' ** Reassign these.
19420                   gstrCrtRpt_Ordinal = strOrd
19430                   gstrCrtRpt_Version = strVer
19440                   gstrCrtRpt_CashAssets_Beg = strCashBeg
19450                   .CashAssets_Beg = dblCashBeg
19460                   intRetVal_SummaryNew = SummaryNew_FL(strProc, blnRebuildTable, blnIsSummary, frm)  ' ** Function: Above.
19470                   If intRetVal_SummaryNew <> 0 Then
                          ' ** Return Codes:
                          ' **    0  Success.
                          ' **   -1  Canceled.
                          ' **   -2  No data.
                          ' **   -3  Missing entry.
                          ' **   -9  Error.
19480                     blnContinue = False
19490                     blnAllCancel = True
19500                     AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
19510                   Else
19520                     If IsNull(gvarCrtRpt_FL_SpecData) = False Then
19530                       Set dbs = CurrentDb
19540                       With dbs

                              ' ** VGC 11/27/2009: I HAVE BEEN UNABLE TO DIVINE WHAT MY ORIGINAL INTENT WAS.
                              ' ** THOUGH I KNOW, AND RICH CONCURS, THAT THIS FLORIDA REPORT DID GIVE
                              ' ** CORRECT INFORMATION AT ONE TIME, I HAVE NO IDEA NOW WHAT HAPPENED TO IT.
                              ' ** SINCE I CAN'T MAKE HEADS-NOR-TAILS OUT OF THE DATA BEING COLLECTED, AND
                              ' ** SECTION I. AND SECTION VI. JUST AREN'T SHOWING GOOD DATA, I'M JUST GOING
                              ' ** TO HAVE TO GO GRAB IT SOMEHOW, HERE, AND PUT IT INTO SECTIONS I. AND VI.
                              ' ** WE NEED A GOOD OPENING BALANCE SET, AND
                              ' ** WE NEED A GOOD CLOSING BALANCE SET.

19550                         Select Case frm.opgType
                              Case frm.opgType_optRep.OptionValue
                                ' ** tmpCourtReportData3, grouped and summed, for Sections II, III, IV, V. (SKIPS INC/EXP DETAILS!)
19560                           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_10_52")
19570                         Case frm.opgType_optGuard.OptionValue
                                ' ** tmpCourtReportData3, grouped and summed, for Sections II, III, IV. (SKIPS INC/EXP DETAILS!)
19580                           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_10_53")
19590                         End Select
19600                         Set rst = qdf.OpenRecordset
19610                         With rst
19620                           .MoveFirst
19630                           curTmp02 = ![ICash]      '  $3,426.26
19640                           curTmp03 = ![PCash]      '   ($953.82)
19650                           curTmp04 = ![Total]      '  $2,472.44
19660                           .Close
19670                         End With
19680                         DoEvents

                              ' ** Update tmpCourtReportData3, Section I, by specified [inctot], [printot], [costot].
19690                         Set qdf = .QueryDefs("qryCourtReport_FL_00_New_10_54")
19700                         With qdf.Parameters
19710                           ![inctot] = 0@ + Val(gstrCrtRpt_CashAssets_Beg)
19720                           ![printot] = (((curTmp05 + curTmp06 + curTmp07) - curTmp04) - Val(gstrCrtRpt_CashAssets_Beg))
19730                           ![costot] = 0@
19740                         End With
19750                         qdf.Execute
19760                         DoEvents

19770                         Select Case frm.opgType
                              Case frm.opgType_optRep.OptionValue
                                ' ** Update tmpCourtReportData3, Section VI, by specified [inctot], [printot], [costot].
19780                           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_10_55")
19790                         Case frm.opgType_optGuard.OptionValue
                                ' ** Update tmpCourtReportData3, Section V, by specified [inctot], [printot], [costot].
19800                           Set qdf = .QueryDefs("qryCourtReport_FL_00_New_10_56")
19810                         End Select
19820                         With qdf.Parameters
19830                           ![inctot] = (curTmp02 + Val(gstrCrtRpt_CashAssets_Beg))
19840                           ![printot] = ((curTmp05 + curTmp06 + curTmp07) - (curTmp02 + Val(gstrCrtRpt_CashAssets_Beg)))
19850                           ![costot] = curTmp07
19860                         End With
19870                         qdf.Execute
19880                         DoEvents

19890                         .Close
19900                       End With
19910                     End If
19920                   End If
19930                 End If
19940               Else
19950                 blnContinue = False
19960                 blnAllCancel = True
19970                 AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
19980               End If
19990               DoEvents

                    ' ** Set up sources for the NEW FL REPORT's.
20000               Select Case strReportNumber
                    Case "0"
20010                 .GetBal_Beg gstrAccountNo  ' ** Form Procedure: frmRpt_CourtReports_Fl.
20020                 DoEvents
20030                 .GetBal_End gstrAccountNo  ' ** Form Procedure: frmRpt_CourtReports_Fl.
20040                 DoEvents
20050               Case "2", "2A"
20060                 Set dbs = CurrentDb
20070                 With dbs
                        ' ** Empty tmpCourtReportData2.
20080                   Set qdf = .QueryDefs("qryCourtReport_FL_00_New_07_01a")
20090                   qdf.Execute
                        ' ** Append qryCourtReport_FL_02m to tmpCourtReportData2.
20100                   Set qdf = .QueryDefs("qryCourtReport_FL_02mb")
20110                   qdf.Execute
20120                   DoEvents
                        ' ** Append qryCourtReport_FL_02oa to tmpCourtReportData2.
20130                   Set qdf = .QueryDefs("qryCourtReport_FL_02ob")
20140                   qdf.Execute
20150                   DoEvents
                        ' ** Append qryCourtReport_FL_02on to tmpCourtReportData2.
20160                   Set qdf = .QueryDefs("qryCourtReport_FL_02pb")
20170                   qdf.Execute
20180                   .Close
20190                 End With
20200               End Select
20210               DoEvents
20220               If Right(strReportNumber, 1) = "A" Or Right(strReportNumber, 1) = "B" Then
20230                 If strReportNumber = "4A" Then
                        ' ** RecordSource =
20240                   Select Case .opgType
                        Case .opgType_optRep.OptionValue
20250                     strTmp01 = "rptCourtRptFL_04A"
20260                   Case .opgType_optGuard.OptionValue
20270                     strTmp01 = "rptCourtRptFL_04C"
20280                   End Select
20290                 Else
20300                   strTmp01 = "rptCourtRptFL_" & Right("00" & strReportNumber, 3)
20310                 End If
20320               Else
20330                 If strReportNumber = "4" Then
20340                   Select Case .opgType
                        Case .opgType_optRep.OptionValue
                          ' ** RecordSource = qryCourtReport_FL_04a.
20350                     strTmp01 = "rptCourtRptFL_04"
20360                   Case .opgType_optGuard.OptionValue
                          ' ** RecordSource = qryCourtReport_FL_04h.
20370                     strTmp01 = "rptCourtRptFL_04B"
20380                   End Select
20390                 ElseIf strReportNumber = "0" Then
20400                   Select Case .chkGroupBy_IncExpCode
                        Case True
20410                     strTmp01 = "rptCourtRptFL_00A"
20420                   Case False
20430                     strTmp01 = "rptCourtRptFL_00"
20440                   End Select
20450                 Else
20460                   strTmp01 = "rptCourtRptFL_" & Right("00" & strReportNumber, 2)
20470                 End If
20480               End If
20490               DoEvents

20500             Else
                    ' ** Return codes:
                    ' **    0  Success.
                    ' **   -2  No data.
                    ' **   -3  Missing entry.
                    ' **   -9  Error.
20510               blnContinue = False
20520               blnAllCancel = True
20530               AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
20540             End If
20550             DoEvents

20560             If blnExcel = True And blnContinue = True And blnAllCancel = False Then
20570               If strReportNumber <> "0B" Then
20580                 If blnNoTmp2 = True Then
                        ' ** No data to roll back.

20590                 End If
20600                 strQry = vbNullString: blnNoData = False
20610                 Select Case strReportNumber
                      Case "0"  ' ** Summary of Account.
20620                   Select Case .chkGroupBy_IncExpCode
                        Case True
20630                     strQry = "qryCourtReport_FL_00f"
20640                   Case False
20650                     strQry = "qryCourtReport_FL_00e"
20660                   End Select
20670                 Case "2"  ' ** Receipts.
20680                   strQry = "qryCourtReport_FL_02l"
20690                   varTmp00 = DCount("*", strQry)
20700                   If IsNull(varTmp00) = True Then
20710                     blnNoData = True
20720                     strQry = "qryCourtReport_FL_02ld"  ' ** For Export - No Data.
20730                   Else
20740                     If varTmp00 = 0 Then
20750                       blnNoData = True
20760                       strQry = "qryCourtReport_FL_02ld"  ' ** For Export - No Data.
20770                     End If
20780                   End If
20790                 Case "2A"  ' ** Receipts, Rev/Exp.
20800                   strQry = "qryCourtReport_FL_02la"
20810                   varTmp00 = DCount("*", strQry)
20820                   If IsNull(varTmp00) = True Then
20830                     blnNoData = True
20840                     strQry = "qryCourtReport_FL_02ld"  ' ** For Export - No Data.
20850                   Else
20860                     If varTmp00 = 0 Then
20870                       blnNoData = True
20880                       strQry = "qryCourtReport_FL_02ld"  ' ** For Export - No Data.
20890                     End If
20900                   End If
20910                 Case "3"  ' ** Capital Transactions and Adjustments.
20920                   strQry = "qryCourtReport_FL_03n"
20930                   varTmp00 = DCount("*", strQry)
20940                   If IsNull(varTmp00) = True Then
20950                     blnNoData = True
20960                     strQry = "qryCourtReport_FL_03q"  ' ** For Export - No Data.
20970                   Else
20980                     If varTmp00 = 0 Then
20990                       blnNoData = True
21000                       strQry = "qryCourtReport_FL_03q"  ' ** For Export - No Data.
21010                     End If
21020                   End If
21030                 Case "4"  ' ** Disbursements.
21040                   Select Case .opgType
                        Case .opgType_optRep.OptionValue
21050                     strQry = "qryCourtReport_FL_04q"    ' ** Report RecordSource: qryCourtReport_FL_04a. OK
21060                     varTmp00 = DCount("*", strQry)
21070                     If IsNull(varTmp00) = True Then
21080                       blnNoData = True
21090                       strQry = "qryCourtReport_FL_04qc"  ' ** For Export - No Data.
21100                     Else
21110                       If varTmp00 = 0 Then
21120                         blnNoData = True
21130                         strQry = "qryCourtReport_FL_04qc"  ' ** For Export - No Data.
21140                       End If
21150                     End If
21160                   Case .opgType_optGuard.OptionValue
21170                     strQry = "qryCourtReport_FL_04y"    ' ** Report RecordSource: qryCourtReport_FL_04h. OK
21180                     varTmp00 = DCount("*", strQry)
21190                     If IsNull(varTmp00) = True Then
21200                       blnNoData = True
21210                       strQry = "qryCourtReport_FL_04zc"  ' ** For Export - No Data.
21220                     Else
21230                       If varTmp00 = 0 Then
21240                         blnNoData = True
21250                         strQry = "qryCourtReport_FL_04zc"  ' ** For Export - No Data.
21260                       End If
21270                     End If
21280                   End Select
21290                 Case "4A"  ' ** Disbursements, Rev/Exp.
21300                   Select Case .opgType
                        Case .opgType_optRep.OptionValue
21310                     strQry = "qryCourtReport_FL_04_Ae"  ' ** Report RecordSource: qryCourtReport_FL_04a. OK
21320                     varTmp00 = DCount("*", strQry)
21330                     If IsNull(varTmp00) = True Then
21340                       blnNoData = True
21350                       strQry = "qryCourtReport_FL_04_Aec"  ' ** For Export - No Data.
21360                     Else
21370                       If varTmp00 = 0 Then
21380                         blnNoData = True
21390                         strQry = "qryCourtReport_FL_04_Aec"  ' ** For Export - No Data.
21400                       End If
21410                     End If
21420                   Case .opgType_optGuard.OptionValue
21430                     Set dbs = CurrentDb
21440                     With dbs
                            ' ** Append qryCourtReport_FL_04_Af to tmpCourtReportData6; report title.
21450                       Set qdf = .QueryDefs("qryCourtReport_FL_04_Ag")
21460                       qdf.Execute
21470                       DoEvents
                            ' ** Append qryCourtReport_FL_04_Ah to tmpCourtReportData6; report period.
21480                       Set qdf = .QueryDefs("qryCourtReport_FL_04_Ai")
21490                       qdf.Execute
21500                       DoEvents
                            ' ** Append qryCourtReport_FL_04_Aha to tmpCourtReportData6; report author.
21510                       Set qdf = .QueryDefs("qryCourtReport_FL_04_Aia")
21520                       qdf.Execute
21530                       DoEvents
                            ' ** Append qryCourtReport_FL_04_Aj (tmpCourtReportData2) to tmpCourtReportData6.
21540                       Set qdf = .QueryDefs("qryCourtReport_FL_04_Ak")
21550                       qdf.Execute
21560                       DoEvents
                            ' ** Append qryCourtReport_FL_04_Al (tmpCourtReportData2, grouped and summed by Rev/Exp) to tmpCourtReportData6.
21570                       Set qdf = .QueryDefs("qryCourtReport_FL_04_Am")
21580                       qdf.Execute
21590                       DoEvents
                            ' ** Append qryCourtReport_FL_04_An (tmpCourtReportData2, grouped and summed) to tmpCourtReportData6.
21600                       Set qdf = .QueryDefs("qryCourtReport_FL_04_Ao")
21610                       qdf.Execute
21620                       DoEvents
                            ' ** Update tmpCourtReportData6 with DLookups() to qryCourtReport_FL_04_Ap; uniqueidx.
21630                       Set qdf = .QueryDefs("qryCourtReport_FL_04_Aq")
21640                       qdf.Execute
21650                       DoEvents
                            ' ** Update tmpCourtReportData6 with DLookups() to qryCourtReport_FL_04_Ar; uniqueidy.
21660                       Set qdf = .QueryDefs("qryCourtReport_FL_04_As")
21670                       qdf.Execute
21680                       .Close
21690                     End With
21700                     DoEvents
21710                     strQry = "qryCourtReport_FL_04_Au"  ' ** Report RecordSource: qryCourtReport_FL_04h.
21720                     varTmp00 = DCount("*", strQry)
21730                     If IsNull(varTmp00) = True Then
21740                       blnNoData = True
21750                       strQry = "qryCourtReport_FL_04_Ax"  ' ** For Export - No Data.
21760                     Else
21770                       If varTmp00 = 0 Then
21780                         blnNoData = True
21790                         strQry = "qryCourtReport_FL_04_Ax"  ' ** For Export - No Data.
21800                       End If
21810                     End If
21820                   End Select
21830                 Case "6"  ' ** Distributions.
21840                   strQry = "qryCourtReport_FL_06i"
21850                   varTmp00 = DCount("*", strQry)
21860                   If IsNull(varTmp00) = True Then
21870                     blnNoData = True
21880                     strQry = "qryCourtReport_FL_06l"  ' ** For Export - No Data.
21890                   Else
21900                     If varTmp00 = 0 Then
21910                       blnNoData = True
21920                       strQry = "qryCourtReport_FL_06l"  ' ** For Export - No Data.
21930                     End If
21940                   End If
21950                 End Select
21960                 DoEvents

21970                 strRptCap = vbNullString: strRptPathFile = vbNullString
21980                 strRptPath = .UserReportPath
21990                 strRptName = strTmp01

22000                 strMacro = "mcrExcelExport_CR_FL" & Mid(strRptName, InStr(strRptName, "_"))
22010                 If blnNoData = True Then
22020                   strMacro = strMacro & "_nd"
22030                 End If

                      ' ** New caps have 'Rep' or 'Grdn' in caption.
22040                 For lngX = 0& To (lngCaps - 1&)
22050                   If arr_varCap(C_RNAM, lngX) = strRptName Then
22060                     strRptCap = arr_varCap(C_CAPN, lngX)
22070                     Exit For
22080                   End If
22090                 Next
22100                 DoEvents

22110                 Select Case blnUseSavedPath
                      Case True
22120                   strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
22130                 Case False
22140                   DoCmd.Hourglass False
22150                   strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
22160                 End Select
22170                 DoEvents

22180                 If strRptPathFile <> vbNullString Then
22190                   DoCmd.Hourglass True
22200                   DoEvents
22210                   Select Case blnExcel
                        Case True
22220                     blnAutoStart = .chkOpenExcel
22230                   Case False
22240                     blnAutoStart = .chkOpenWord
22250                   End Select
22260                   If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
22270                   If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
22280                     Kill strRptPathFile
22290                   End If
22300                   DoEvents
22310                   If strQry <> vbNullString Then
22320                     FLCourtReportLoad  ' ** Module Function: modCourReportsFL.
                          ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                          ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
22330                     DoCmd.RunMacro strMacro
                          ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                          ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
22340                     If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_FL_xxx.xls") = True Or _
                              FileExists(strRptPath & LNK_SEP & "CourtReport_FL_xxx.xls") = True Then
22350                       If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_FL_xxx.xls") = True Then
22360                         Name (CurrentAppPath & LNK_SEP & "CourtReport_FL_xxx.xls") As (strRptPathFile)
                              ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
22370                       Else
22380                         Name (strRptPath & LNK_SEP & "CourtReport_FL_xxx.xls") As (strRptPathFile)
                              ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
22390                       End If
22400                       DoEvents
22410                       If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
22420                         DoEvents
22430                         Select Case gblnPrintAll
                              Case True
22440                           lngFiles = lngFiles + 1&
22450                           lngE = lngFiles - 1&
22460                           ReDim Preserve arr_varFile(F_ELEMS, lngE)
22470                           arr_varFile(F_RNAM, lngE) = strRptName
22480                           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
22490                           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
22500                           FileArraySet_FL arr_varFile  ' ** Procedure: Above.
22510                         Case False
22520                           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22530                             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22540                           End If
22550                           DoEvents
22560                           If blnAutoStart = True Then
22570                             OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
22580                           End If
22590                         End Select
                              'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                              '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                              'End If
                              'DoEvents
                              'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
22600                       End If
22610                     End If
22620                   Else
22630                     DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
22640                   End If
22650                   strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
22660                   If strRptPath <> .UserReportPath Then
22670                     .UserReportPath = strRptPath
22680                     .SetUserReportPath  ' ** Form Procedure: frmRpt_CourtReports_Fl.
22690                   End If
22700                 Else
22710                   blnContinue = False
22720                 End If

22730               End If  ' ** 0B.
22740               DoEvents

22750             ElseIf blnContinue = True And blnAllCancel = False Then
                    ' ** Word, with or without Asset List.

22760               If strReportNumber <> "0B" Then

22770                 strRptCap = vbNullString: strRptPathFile = vbNullString
22780                 strRptPath = .UserReportPath
22790                 strRptName = strTmp01

22800                 For lngX = 0& To (lngCaps - 1&)
22810                   If arr_varCap(C_RNAM, lngX) = strRptName Then
22820                     strRptCap = arr_varCap(C_CAPN, lngX)
22830                     Exit For
22840                   End If
22850                 Next
22860                 DoEvents

22870                 Select Case blnUseSavedPath
                      Case True
22880                   strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
22890                 Case False
22900                   DoCmd.Hourglass False
22910                   strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
22920                 End Select
22930                 DoEvents

22940                 If strRptPathFile <> vbNullString Then
22950                   DoCmd.Hourglass True
22960                   DoEvents
22970                   Select Case blnExcel
                        Case True
22980                     blnAutoStart = .chkOpenExcel
22990                   Case False
23000                     blnAutoStart = .chkOpenWord
23010                   End Select
23020                   If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
23030                   If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
23040                     Kill strRptPathFile
23050                   End If
23060                   DoEvents
23070                   Select Case gblnPrintAll
                        Case True
23080                     lngFiles = lngFiles + 1&
23090                     lngE = lngFiles - 1&
23100                     ReDim Preserve arr_varFile(F_ELEMS, lngE)
23110                     arr_varFile(F_RNAM, lngE) = strRptName
23120                     arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
23130                     arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23140                     FileArraySet_FL arr_varFile  ' ** Procedure: Above.
23150                     DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
                          'Debug.Print "'OutputTo: 1A  " & CStr(lngFiles) & "  " & strRptName
23160                   Case False
23170                     DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
                          'Debug.Print "'OutputTo: 1B"
23180                   End Select
                        'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
23190                   strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23200                   If strRptPath <> .UserReportPath Then
23210                     .UserReportPath = strRptPath
23220                     .SetUserReportPath  ' ** Form Procedure: frmRpt_CourtReports_Fl.
23230                   End If
23240                 Else
23250                   blnContinue = False
23260                 End If
23270                 DoEvents

                      'OutputTo: 1A  1  rptCourtRptFL_00A
                      'OutputTo: 2A  2  rptCourtRptFL_00B
                      'OutputTo: 1A  3  rptCourtRptFL_02A
                      'OutputTo: 1A  4  rptCourtRptFL_04C
                      'OutputTo: 1A  5  rptCourtRptFL_03
                      'OutputTo: 3A  6  rptCourtRptFL_07

23280               End If  ' ** Not 0B.

23290             End If

23300             If ((strReportNumber = "0" And .chkAssetList = True) Or (strReportNumber = "0B")) And blnContinue = True And blnAllCancel = False Then

23310               If strReportNumber = "0B" And gblnPrintAll = True Then
                      ' ** For cmdWordAll and cmdExcelAll, see if this sets the focus on the Asset List button.
23320                 Select Case blnExcel
                      Case True
23330                   .cmdExcel00B.SetFocus
23340                 Case False
23350                   .cmdWord00B.SetFocus
23360                 End Select
23370                 DoEvents
23380               End If

23390               .chkAssetList_Start = True
23400               intRetVal_BuildAssetListInfo = .BuildAssetListInfo_FL(#1/1/1900#, (CDate(.DateStart) - 1), "Beginning", strProc, THIS_PROC)  ' ** Form Function: frmRpt_CourtReports_Fl.
23410               DoEvents

23420               If intRetVal_BuildAssetListInfo = 0 Then
23430                 If blnExcel = True Then

23440                   strQry = "qryCourtReport_FL_07y"

23450                   strRptCap = vbNullString: strRptPathFile = vbNullString
23460                   strRptPath = .UserReportPath
23470                   strRptName = "rptCourtRptFL_00B"

23480                   strMacro = "mcrExcelExport_CR_FL" & Mid(strRptName, InStr(strRptName, "_"))

23490                   For lngX = 0& To (lngCaps - 1&)
23500                     If arr_varCap(C_RNAM, lngX) = strRptName Then
23510                       strRptCap = arr_varCap(C_CAPN, lngX)
23520                       Exit For
23530                     End If
23540                   Next
23550                   DoEvents

23560                   Select Case blnUseSavedPath
                        Case True
23570                     strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
23580                   Case False
23590                     DoCmd.Hourglass False
23600                     strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
23610                   End Select
23620                   DoEvents

23630                   If strRptPathFile <> vbNullString Then
23640                     DoCmd.Hourglass True
23650                     DoEvents
23660                     If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
23670                       Kill strRptPathFile
23680                     End If
23690                     DoEvents
                          ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                          ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
23700                     DoCmd.RunMacro strMacro
                          ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                          ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
23710                     If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_FL_xxxx.xls") = True Or _
                              FileExists(strRptPath & LNK_SEP & "CourtReport_FL_xxxx.xls") = True Then  ' ** 4 X's!
23720                       If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_FL_xxxx.xls") = True Then
23730                         Name (CurrentAppPath & LNK_SEP & "CourtReport_FL_xxxx.xls") As (strRptPathFile)
                              ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
23740                       Else
23750                         Name (strRptPath & LNK_SEP & "CourtReport_FL_xxxx.xls") As (strRptPathFile)
                              ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
23760                       End If
23770                       DoEvents
23780                       If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
23790                         DoEvents
23800                         Select Case blnExcel
                              Case True
23810                           blnAutoStart = .chkOpenExcel
23820                         Case False
23830                           blnAutoStart = .chkOpenWord
23840                         End Select
23850                         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
23860                         Select Case gblnPrintAll
                              Case True
23870                           lngFiles = lngFiles + 1&
23880                           lngE = lngFiles - 1&
23890                           ReDim Preserve arr_varFile(F_ELEMS, lngE)
23900                           arr_varFile(F_RNAM, lngE) = strRptName
23910                           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
23920                           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23930                           FileArraySet_FL arr_varFile  ' ** Procedure: Above.
23940                         Case False
23950                           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
23960                             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
23970                           End If
23980                           DoEvents
23990                           If blnAutoStart = True Then
24000                             OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
24010                           End If
24020                         End Select
                              'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                              '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                              'End If
                              'DoEvents
                              'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
24030                       End If
24040                     End If
24050                     strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24060                     If strRptPath <> .UserReportPath Then
24070                       .UserReportPath = strRptPath
24080                       .SetUserReportPath  ' ** Form Procedure: frmRpt_CourtReports_Fl.
24090                     End If
24100                   Else
24110                     blnContinue = False
24120                   End If
24130                   DoEvents

24140                 Else
                        ' ** Word, with Asset List.

24150                   strRptCap = vbNullString: strRptPathFile = vbNullString
24160                   strRptPath = .UserReportPath
24170                   strRptName = "rptCourtRptFL_00B"

24180                   For lngX = 0& To (lngCaps - 1&)
24190                     If arr_varCap(C_RNAM, lngX) = strRptName Then
24200                       strRptCap = arr_varCap(C_CAPN, lngX)
24210                       Exit For
24220                     End If
24230                   Next
24240                   DoEvents

24250                   Select Case blnUseSavedPath
                        Case True
24260                     strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
24270                   Case False
24280                     DoCmd.Hourglass False
24290                     strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
24300                   End Select
24310                   DoEvents

24320                   If strRptPathFile <> vbNullString Then
24330                     DoCmd.Hourglass True
24340                     DoEvents
24350                     If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
24360                       Kill strRptPathFile
24370                     End If
24380                     DoEvents
24390                     Select Case blnExcel
                          Case True
24400                       blnAutoStart = .chkOpenExcel
24410                     Case False
24420                       blnAutoStart = .chkOpenWord
24430                     End Select
24440                     If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
24450                     Select Case gblnPrintAll
                          Case True
24460                       lngFiles = lngFiles + 1&
24470                       lngE = lngFiles - 1&
24480                       ReDim Preserve arr_varFile(F_ELEMS, lngE)
24490                       arr_varFile(F_RNAM, lngE) = strRptName
24500                       arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
24510                       arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24520                       FileArraySet_FL arr_varFile  ' ** Procedure: Above.
24530                       DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
                            'Debug.Print "'OutputTo: 2A  " & CStr(lngFiles) & "  " & strRptName
24540                     Case False
24550                       DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
                            'Debug.Print "'OutputTo: 2B"
24560                     End Select
                          'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
24570                     strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24580                     If strRptPath <> .UserReportPath Then
24590                       .UserReportPath = strRptPath
24600                       .SetUserReportPath  ' ** Form Procedure: frmRpt_CourtReports_Fl.
24610                     End If
24620                   Else
24630                     blnContinue = False
24640                   End If
24650                   DoEvents

24660                 End If
24670               Else
                      ' ** Return codes:
                      ' **    0  Success.
                      ' **   -2  No data.
                      ' **   -3  Missing entry.
                      ' **   -9  Error.
24680                 blnContinue = False
24690               End If

24700               .chkAssetList_Start = False

24710             End If
24720           Else
24730             blnContinue = False
24740             blnAllCancel = True
24750             AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
24760           End If
24770         Else
                ' ** Assets on Hand at Close of Accounting Period; non-current end date, 7.
                ' ** Assets on Hand at Close of Accounting Period; current end date, 7A.

24780           FLCourtReportLoad  ' ** Module Function: modCourReportsFL.
24790           intRetVal_BuildAssetListInfo = .BuildAssetListInfo_FL(.DateStart, .DateEnd, "Ending", strProc, THIS_PROC)  ' ** Form Function: frmRpt_CourtReports_Fl.
24800           DoEvents

24810           If intRetVal_BuildAssetListInfo = 0 Then
24820             If blnExcel = True Then

24830               blnNoData = False
24840               strQry = "qryCourtReport_FL_07y"
24850               varTmp00 = DCount("*", strQry)  ' ** I know it's unlikely.
24860               If IsNull(varTmp00) = True Then
24870                 blnNoData = True
24880                 strQry = "qryCourtReport_FL_04_Ax"  ' ** For Export - No Data.
24890               Else
24900                 If varTmp00 = 0 Then
24910                   blnNoData = True
24920                   strQry = "qryCourtReport_FL_04_Ax"  ' ** For Export - No Data.
24930                 End If
24940               End If

24950               strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
24960               strRptPath = .UserReportPath
24970               strRptName = "rptCourtRptFL_07"

24980               strMacro = "mcrExcelExport_CR_FL" & Mid(strRptName, InStr(strRptName, "_"))
24990               If blnNoData = True Then
25000                 strMacro = strMacro & "_nd"
25010               End If

25020               For lngX = 0& To (lngCaps - 1&)
25030                 If arr_varCap(C_RNAM, lngX) = strRptName Then
25040                   strRptCap = arr_varCap(C_CAPN, lngX)
25050                   Exit For
25060                 End If
25070               Next
25080               DoEvents

25090               Select Case blnUseSavedPath
                    Case True
25100                 strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
25110               Case False
25120                 DoCmd.Hourglass False
25130                 strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
25140               End Select

25150               If strRptPathFile <> vbNullString Then
25160                 DoCmd.Hourglass True
25170                 DoEvents
25180                 If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
25190                   Kill strRptPathFile
25200                 End If
25210                 DoEvents
                      ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                      ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
25220                 DoCmd.RunMacro strMacro
                      ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                      ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
25230                 If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_FL_xxxx.xls") = True Or _
                          FileExists(strRptPath & LNK_SEP & "CourtReport_FL_xxxx.xls") = True Then  ' ** 4 X's!
25240                   If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_FL_xxxx.xls") = True Then
25250                     Name (CurrentAppPath & LNK_SEP & "CourtReport_FL_xxxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
25260                   Else
25270                     Name (strRptPath & LNK_SEP & "CourtReport_FL_xxxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
25280                   End If
25290                   DoEvents
25300                   If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
25310                     DoEvents
25320                     Select Case blnExcel
                          Case True
25330                       blnAutoStart = .chkOpenExcel
25340                     Case False
25350                       blnAutoStart = .chkOpenWord
25360                     End Select
25370                     If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
25380                     Select Case gblnPrintAll
                          Case True
25390                       lngFiles = lngFiles + 1&
25400                       lngE = lngFiles - 1&
25410                       ReDim Preserve arr_varFile(F_ELEMS, lngE)
25420                       arr_varFile(F_RNAM, lngE) = strRptName
25430                       arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
25440                       arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
25450                       FileArraySet_FL arr_varFile  ' ** Procedure: Above.
25460                     Case False
25470                       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
25480                         EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
25490                       End If
25500                       DoEvents
25510                       If blnAutoStart = True Then
25520                         OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
25530                       End If
25540                     End Select
                          'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                          '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                          'End If
                          'DoEvents
                          'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
25550                   End If
25560                 End If
25570                 strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
25580                 If strRptPath <> .UserReportPath Then
25590                   .UserReportPath = strRptPath
25600                   .SetUserReportPath  ' ** Form Procedure: frmRpt_CourtReports_Fl.
25610                 End If
25620               Else
25630                 blnContinue = False
25640               End If
25650               DoEvents

25660             Else

25670               strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
25680               strRptPath = .UserReportPath
25690               strRptName = "rptCourtRptFL_07"

25700               For lngX = 0& To (lngCaps - 1&)
25710                 If arr_varCap(C_RNAM, lngX) = strRptName Then
25720                   strRptCap = arr_varCap(C_CAPN, lngX)
25730                   Exit For
25740                 End If
25750               Next
25760               DoEvents

25770               Select Case blnUseSavedPath
                    Case True
25780                 strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
25790               Case False
25800                 DoCmd.Hourglass False
25810                 strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
25820               End Select
25830               DoEvents

25840               If strRptPathFile <> vbNullString Then
25850                 DoCmd.Hourglass True
25860                 DoEvents
25870                 Select Case blnExcel
                      Case True
25880                   blnAutoStart = .chkOpenExcel
25890                 Case False
25900                   blnAutoStart = .chkOpenWord
25910                 End Select
25920                 If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
25930                 If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
25940                   Kill strRptPathFile
25950                 End If
25960                 DoEvents
25970                 Select Case gblnPrintAll
                      Case True
25980                   lngFiles = lngFiles + 1&
25990                   lngE = lngFiles - 1&
26000                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
26010                   arr_varFile(F_RNAM, lngE) = strRptName
26020                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
26030                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
26040                   FileArraySet_FL arr_varFile  ' ** Procedure: Above.
26050                   DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
                        'Debug.Print "'OutputTo: 3A  " & CStr(lngFiles) & "  " & strRptName
26060                 Case False
26070                   DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
                        'Debug.Print "'OutputTo: 3B"
26080                 End Select
                      'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
26090                 strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
26100                 If strRptPath <> .UserReportPath Then
26110                   .UserReportPath = strRptPath
26120                   .SetUserReportPath  ' ** Form Procedure: frmRpt_CourtReports_Fl.
26130                 End If
26140               Else
26150                 blnContinue = False
26160               End If

26170             End If
26180           Else
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -3  Missing entry.
                  ' **   -9  Error.
26190             blnContinue = False
26200             blnAllCancel = True
26210             AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
26220           End If
26230           DoEvents

26240           If blnContinue = False And intRetVal_BuildAssetListInfo = -2 Then
26250             DoCmd.Hourglass False
26260             MsgBox "There is no data for the report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
26270           End If

26280         End If

26290       Else
26300         blnAllCancel = True
26310         AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
26320       End If

26330     End If  ' ** blnContinue.

26340   End With

26350   DoCmd.Hourglass False

EXITP:
        ' ** Check some Public variables in case there's been an error.
26360   blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
26370   blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
26380   Set rst = Nothing
26390   Set qdf = Nothing
26400   Set dbs = Nothing
26410   Exit Sub

ERRH:
26420   DoCmd.Hourglass False
26430   Select Case ERR.Number
        Case 70  ' ** Permission denied.
26440     Beep
26450     MsgBox "Trust Accountant is unable to save the file." & vbCrLf & vbCrLf & _
            "If the program to which you're exporting is open," & vbCrLf & _
            "please close it and try again.", vbInformation + vbOKOnly, "Save Failed"
26460   Case 2501  ' ** The '|' action was Canceled.
          '  ' ** Do nothing.
26470   Case 2302  ' ** Access can't save the output data to the file you've selected.
26480     blnAllCancel = True
26490     AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
26500     Beep
26510     MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
26520   Case Else
26530     blnAllCancel = True
26540     AllCancelSet2_FL blnAllCancel  ' ** Procedure: Above.
26550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26560   End Select
26570   Resume EXITP

End Sub

Public Sub Calendar_Handler_FL(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_FL"

        Dim strEvent As String, strCtlName As String
        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer, intNum As Integer
        Dim blnRetVal As Boolean

26610   With frm

26620     strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
26630     strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.
26640     intNum = Val(Right(strCtlName, 1))

26650     Select Case strEvent
          Case "Click"
26660       Select Case intNum
            Case 1
26670         datStartDate = Date
26680         datEndDate = 0
26690         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
26700         If blnRetVal = True Then
26710           .DateStart = datStartDate
26720         Else
26730           .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
26740         End If
26750         .DateStart.SetFocus
26760       Case 2
26770         datStartDate = Date
26780         datEndDate = 0
26790         blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
26800         If blnRetVal = True Then
26810           .DateEnd = datStartDate
26820         Else
26830           .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
26840         End If
26850         .DateEnd.SetFocus
26860         Cancel = 0
26870         .DateEnd_Exit Cancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
26880         If Cancel = 0 Then
26890           .cmbAccounts.SetFocus
26900         End If
26910       End Select
26920     Case "GotFocus"
26930       Select Case intNum
            Case 1
26940         blnCalendar1_Focus = True
26950         .cmdCalendar1_raised_semifocus_dots_img.Visible = True
26960         .cmdCalendar1_raised_img.Visible = False
26970         .cmdCalendar1_raised_focus_img.Visible = False
26980         .cmdCalendar1_raised_focus_dots_img.Visible = False
26990         .cmdCalendar1_sunken_focus_dots_img.Visible = False
27000         .cmdCalendar1_raised_img_dis.Visible = False
27010       Case 2
27020         blnCalendar2_Focus = True
27030         .cmdCalendar2_raised_semifocus_dots_img.Visible = True
27040         .cmdCalendar2_raised_img.Visible = False
27050         .cmdCalendar2_raised_focus_img.Visible = False
27060         .cmdCalendar2_raised_focus_dots_img.Visible = False
27070         .cmdCalendar2_sunken_focus_dots_img.Visible = False
27080         .cmdCalendar2_raised_img_dis.Visible = False
27090       End Select
27100     Case "MouseDown"
27110       Select Case intNum
            Case 1
27120         blnCalendar1_MouseDown = True
27130         .cmdCalendar1_sunken_focus_dots_img.Visible = True
27140         .cmdCalendar1_raised_img.Visible = False
27150         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
27160         .cmdCalendar1_raised_focus_img.Visible = False
27170         .cmdCalendar1_raised_focus_dots_img.Visible = False
27180         .cmdCalendar1_raised_img_dis.Visible = False
27190       Case 2
27200         blnCalendar2_MouseDown = True
27210         .cmdCalendar2_sunken_focus_dots_img.Visible = True
27220         .cmdCalendar2_raised_img.Visible = False
27230         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
27240         .cmdCalendar2_raised_focus_img.Visible = False
27250         .cmdCalendar2_raised_focus_dots_img.Visible = False
27260         .cmdCalendar2_raised_img_dis.Visible = False
27270       End Select
27280     Case "MouseMove"
27290       Select Case intNum
            Case 1
27300         If blnCalendar1_MouseDown = False Then
27310           Select Case blnCalendar1_Focus
                Case True
27320             .cmdCalendar1_raised_focus_dots_img.Visible = True
27330             .cmdCalendar1_raised_focus_img.Visible = False
27340           Case False
27350             .cmdCalendar1_raised_focus_img.Visible = True
27360             .cmdCalendar1_raised_focus_dots_img.Visible = False
27370           End Select
27380           .cmdCalendar1_raised_img.Visible = False
27390           .cmdCalendar1_raised_semifocus_dots_img.Visible = False
27400           .cmdCalendar1_sunken_focus_dots_img.Visible = False
27410           .cmdCalendar1_raised_img_dis.Visible = False
27420         End If
27430       Case 2
27440         If blnCalendar2_MouseDown = False Then
27450           Select Case blnCalendar2_Focus
                Case True
27460             .cmdCalendar2_raised_focus_dots_img.Visible = True
27470             .cmdCalendar2_raised_focus_img.Visible = False
27480           Case False
27490             .cmdCalendar2_raised_focus_img.Visible = True
27500             .cmdCalendar2_raised_focus_dots_img.Visible = False
27510           End Select
27520           .cmdCalendar2_raised_img.Visible = False
27530           .cmdCalendar2_raised_semifocus_dots_img.Visible = False
27540           .cmdCalendar2_sunken_focus_dots_img.Visible = False
27550           .cmdCalendar2_raised_img_dis.Visible = False
27560         End If
27570       End Select
27580     Case "MouseUp"
27590       Select Case intNum
            Case 1
27600         .cmdCalendar1_raised_focus_dots_img.Visible = True
27610         .cmdCalendar1_raised_img.Visible = False
27620         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
27630         .cmdCalendar1_raised_focus_img.Visible = False
27640         .cmdCalendar1_sunken_focus_dots_img.Visible = False
27650         .cmdCalendar1_raised_img_dis.Visible = False
27660         blnCalendar1_MouseDown = False
27670       Case 2
27680         .cmdCalendar2_raised_focus_dots_img.Visible = True
27690         .cmdCalendar2_raised_img.Visible = False
27700         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
27710         .cmdCalendar2_raised_focus_img.Visible = False
27720         .cmdCalendar2_sunken_focus_dots_img.Visible = False
27730         .cmdCalendar2_raised_img_dis.Visible = False
27740         blnCalendar2_MouseDown = False
27750       End Select
27760     Case "LostFocus"
27770       Select Case intNum
            Case 1
27780         .cmdCalendar1_raised_img.Visible = True
27790         .cmdCalendar1_raised_semifocus_dots_img.Visible = False
27800         .cmdCalendar1_raised_focus_img.Visible = False
27810         .cmdCalendar1_raised_focus_dots_img.Visible = False
27820         .cmdCalendar1_sunken_focus_dots_img.Visible = False
27830         .cmdCalendar1_raised_img_dis.Visible = False
27840         blnCalendar1_Focus = False
27850       Case 2
27860         .cmdCalendar2_raised_img.Visible = True
27870         .cmdCalendar2_raised_semifocus_dots_img.Visible = False
27880         .cmdCalendar2_raised_focus_img.Visible = False
27890         .cmdCalendar2_raised_focus_dots_img.Visible = False
27900         .cmdCalendar2_sunken_focus_dots_img.Visible = False
27910         .cmdCalendar2_raised_img_dis.Visible = False
27920         blnCalendar2_Focus = False
27930       End Select
27940     End Select

27950   End With

EXITP:
27960   Exit Sub

ERRH:
27970   Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
27980   Case Else
27990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28000   End Select
28010   Resume EXITP

End Sub
