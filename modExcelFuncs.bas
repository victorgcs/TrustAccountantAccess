Attribute VB_Name = "modExcelFuncs"
Option Compare Database
Option Explicit

'VGC 03/23/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const IsDev = 0  ' ** 0 = release, -1 = development.
' ** Also in:
' **   frmXAdmin_Misc
' **   modAutonumberFieldFuncs
' **   modVersionDocFuncs
' **   zz_mod_MDEPrepFuncs

#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:
'modExcelFuncs
'frmAccountContacts
'frmCurrency
'frmCurrency_Rate
'frmErrorLog
'frmMasterBalance
'frmRpt_AccountBalance
'frmRpt_AccountProfile
'frmRpt_AccountReviews
'frmRpt_ArchivedTransactions
'frmRpt_AssetHistory
'frmRpt_CapitalGainAndLoss
'frmRpt_CashControl
'frmRpt_CourtReports_CA
'frmRpt_CourtReports_FL
'frmRpt_CourtReports_NS
'frmRpt_CourtReports_NY
'frmRpt_Holdings
'frmRpt_IncomeExpense
'frmRpt_IncomeStatement
'frmRpt_Locations
'frmRpt_Maturity
'frmRpt_NewClosedAccounts
'frmRpt_PurchasedSold
'frmRpt_StatementOfCondition
'frmRpt_TaxIncomeDeductions
'frmRpt_TaxLot
'frmRpt_TransactionsByType
'frmRpt_UnrealizedGainAndLoss
'frmStatementParameters
'frmTransaction_Audit

' ** See modShellFuncs:
' **   OpenExe()

'gstrFormQuerySpec = "frmRpt_CourtReports_NY"
'glngTaxCode_Distribution = 14&
'gstrReportQuerySpec

' ** See modProcessFuncs:
' **   EXE_IsRunning()
' **   EXE_Terminate()

' ** Forms using UserReportPath (14):  (x form error, x keydown error trap, x open file error trap)
' **   oox frmErrorLog
' **   oxx frmMasterBalance
' **   xxx frmRpt_CapitalGainAndLoss
' **   xxx frmRpt_CashControl
' **   xxx frmRpt_CourtReports_CA
' **   xxx frmRpt_CourtReports_FL
' **   xxx frmRpt_CourtReports_NS
' **   xxx frmRpt_IncomeExpense
' **   xxx frmRpt_IncomeStatement
' **   xxx frmRpt_TaxIncomeDeductions
' **   xxx frmRpt_TransactionsByType
' **   xxx frmStatementParameters
' **   xxx frmTransaction_Audit
' **   oxx frmXAdmin_FileInfo

Private Const THIS_NAME As String = "modExcelFuncs"
' **

Public Function IncomeExpense_MinUniqueID(dbs As DAO.Database) As Variant
' ** The whole thing was about an extra long Ledger description!
' ** Called by:
' **   IncomeExpense_Export(), Below.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "IncomeExpense_MinUniqueID"

        Dim qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngMins As Long, arr_varMin() As Variant
        Dim strAccountNo As String
        Dim lngRecs As Long, lngRecNo As Long
        Dim varTmp00 As Variant, lngTmp01 As Long
        Dim lngX As Long, lngE As Long
        Dim varRetVal As Variant

        ' ** Array: arr_varMin().
        Const M_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const M_ACTNO As Integer = 0

110     varRetVal = 0

120     lngMins = 0&
130     ReDim arr_varMin(M_ELEMS, 0)

140     With dbs

          ' ** Empty tmpIncomeExpenseReports2.
150       Set qdf = .QueryDefs("qryIncomeExpenseReports_27a")
160       qdf.Execute
170       Set qdf = Nothing

180       Set qdf = .QueryDefs("qryIncomeExpenseReports_25")
190       Set rst = qdf.OpenRecordset
200       With rst
210         If .BOF = True And .EOF = True Then
220           varRetVal = 0
230         Else
240           .MoveLast
250           lngRecs = .RecordCount
260           .MoveFirst
270           If lngRecs > 500& Then
280             .Move 500
290             lngRecNo = .AbsolutePosition
300             lngMins = lngMins + 1&
310             lngE = lngMins - 1&
320             ReDim Preserve arr_varMin(M_ELEMS, lngE)
330             arr_varMin(M_ACTNO, lngE) = ![accountno]
340             varTmp00 = lngRecs / 500&
350             If lngRecs Mod 500& = 0 Then
360               lngTmp01 = varTmp00
370             Else
380               lngTmp01 = Int(varTmp00) + 1&
390               For lngX = 2& To lngTmp01
400                 If lngRecNo + 500& < lngRecs Then
410                   .Move 500
420                   lngRecNo = .AbsolutePosition
430                   lngMins = lngMins + 1&
440                   lngE = lngMins - 1&
450                   ReDim Preserve arr_varMin(M_ELEMS, lngE)
460                   arr_varMin(M_ACTNO, lngE) = ![accountno]
470                 Else
480                   .MoveLast
490                   lngMins = lngMins + 1&
500                   lngE = lngMins - 1&
510                   ReDim Preserve arr_varMin(M_ELEMS, lngE)
520                   arr_varMin(M_ACTNO, lngE) = ![accountno]
530                   Exit For
540                 End If
550               Next
560             End If
570           Else
580             .MoveLast
590             strAccountNo = ![accountno]
600           End If
610         End If
620         .Close
630       End With  ' ** rst.
640       Set rst = Nothing
650       Set qdf = Nothing

          ' ** The arr_varMin() array holds accountno's that
          ' ** divide the whole into approximately 500-record groups.

660       If lngMins > 0& Then
670         strAccountNo = vbNullString
680         For lngX = 0& To (lngMins - 1&)
690           If lngX = 0& Then
                ' ** Append qryIncomeExpenseReports_27b (qryIncomeExpenseReports_25,
                ' ** grouped by accountno, revcode_DESC, Min(uniqueid), for first
                ' ** group, by specified [actno]) to tmpIncomeExpenseReports2.
700             Set qdf = .QueryDefs("qryIncomeExpenseReports_27c")
710             With qdf.Parameters
720               ![actno] = arr_varMin(M_ACTNO, lngX)
730             End With
740           Else
                ' ** Append qryIncomeExpenseReports_27d (qryIncomeExpenseReports_25, grouped
                ' ** by accountno, revcode_DESC, with Min(uniqueid), for all the othergroups,
                ' ** by specified [actno1], [actno2]) to tmpIncomeExpenseReports2.
750             Set qdf = .QueryDefs("qryIncomeExpenseReports_27e")
760             With qdf.Parameters
770               ![actno1] = arr_varMin(M_ACTNO, lngX - 1&)
780               ![actno2] = arr_varMin(M_ACTNO, lngX)
790             End With
800           End If
810           qdf.Execute dbFailOnError
820           Set qdf = Nothing
830         Next
840       End If

850     End With  ' ** dbs.

EXITP:
860     Set rst = Nothing
870     Set qdf = Nothing
880     IncomeExpense_MinUniqueID = varRetVal
890     Exit Function

ERRH:
900     varRetVal = 0
910     Select Case ERR.Number
        Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Function

Public Function IncomeExpense_Export(strQry1 As String, strPathFile As String, strRptPath As String, strMode As String) As Boolean
' ** The whole thing was about an extra long Ledger description!
' ** Called by:
' **   frmRpt_IncomeExpense:
' **     cmdRevIncExp_IncomeExcel_Click()
' **     cmdRevIncExp_ExpenseExcel_Click()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "IncomeExpense_Export"

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

      #If IsDev Then
        Dim xlApp As Excel.Application, wbk As Excel.Workbook, wks As Excel.Worksheet, rng As Excel.Range  ' ** Early Binding.
        Dim fnt As Excel.Font, bdr As Excel.Border
      #Else
        Dim xlApp As Object, wbk As Object, wks As Object, rng As Object, fnt As Object, bdr As Object     ' ** Late Binding.
      #End If
        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset
        Dim strLastCell As String, strLastCol As String, strLastRow As String
        Dim strRng_All As String, strRng_Header As String, strRng_Title As String
        Dim strRng_Text1 As String, strRng_Text2 As String, strRng_Values As String, strRng_ThisCell As String
        Dim lngCols As Long, lngThisRow As Long, lngThisCol As Long
        Dim strMacro As String, strSheetName As String
        Dim lngRecs As Long, lngFlds As Long
        Dim blnExcelOpen As Boolean
        Dim lngX As Long, lngY As Long

      #End If
        Dim blnRetVal As Boolean

      #If NoExcel Then
        ' ** Skip.
      #Else

1010    blnRetVal = True

1020    DoCmd.Hourglass True  ' ** Make sure it's still running.
1030    DoEvents

        'totdesc: Trim(IIf(IsNull([RecurringItem]) = True,"",IIf([journaltype]="Received",[RecurringItem],IIf([journaltype]="Paid",[RecurringItem],[RecurringItem]))) & IIf(IsNull([assetno]) = True,"",IIf(IsNull([assetdate]) = True,"",Format([assetdate],"mm/dd/yyyy") & " ") & IIf([shareface]-CLng([shareface])=0,Format([shareface],"#,##0"),Format([shareface],"#,##0.000")) & " " & CStr(Nz([Description],"")) & IIf([rate]>0," " & Format([rate],"#,##0.000%"),"")) & IIf(IsNull([due]) = True,"","  Due " & Format([due],"mm/dd/yyyy")) & "  " & Nz([Jcomment],""))

1040    Set dbs = CurrentDb
1050    With dbs

1060      If strMode = "Expense" Then
            ' ** It's assumed that, because it's here, the error involved the Min(uniqueid) problem,
            ' ** so use the alternate method to get Min(uniqueid).
1070        IncomeExpense_MinUniqueID dbs  ' ** Function: Above.
1080        DoCmd.DeleteObject acQuery, "qryIncomeExpenseReports_29"  ' ** qryIncomeExpenseReports_29_norm.
1090        .QueryDefs.Refresh
1100        DoEvents
1110        DoCmd.CopyObject , "qryIncomeExpenseReports_29", acQuery, "qryIncomeExpenseReports_29_tmp"
1120        .QueryDefs.Refresh
1130        DoEvents
1140        DoCmd.DeleteObject acQuery, "qryIncomeExpenseReports_31"  ' ** qryIncomeExpenseReports_31_norm.
1150        .QueryDefs.Refresh
1160        DoEvents
1170        DoCmd.CopyObject , "qryIncomeExpenseReports_31", acQuery, "qryIncomeExpenseReports_31_tmp"
1180        .QueryDefs.Refresh
1190        DoEvents
1200      End If

1210      Set qdf1 = .QueryDefs(strQry1)
1220      With qdf1

1230        lngFlds = .Fields.Count

1240        Select Case lngFlds
            Case 6&
1250          Select Case strMode
              Case "Income"
                ' ** Income: Dummy single record, with 6 fields; For Export.
                ' ** Called by modExcelFuncs.IncomeExpense_Export();
                ' ** qryIncomeExpenseReports_81_06_05 -> Microsoft Excell 2003 format; for Income, 6 fields.
1260            strMacro = "mcrExcelExport_IncomeExpense_01_06"
1270          Case "Expense"
                ' ** Expense: Dummy single record, with 6 fields; For Export.
                ' ** Called by modExcelFuncs.IncomeExpense_Export();
                ' ** qryIncomeExpenseReports_81_06_10 -> Microsoft Excell 2003 format; for Expense, 6 fields.
1280            strMacro = "mcrExcelExport_IncomeExpense_02_06"
1290          End Select
1300        Case 8&
1310          Select Case strMode
              Case "Income"
                ' ** Income: Dummy single record, with 8 fields; For Export.
                ' ** Called by modExcelFuncs.IncomeExpense_Export();
                ' ** qryIncomeExpenseReports_81_08_05 -> Microsoft Excell 2003 format; for Income, 8 fields.
1320            strMacro = "mcrExcelExport_IncomeExpense_01_08"
1330          Case "Expense"
                ' ** Expense: Dummy single record, with 8 fields; For Export.
                ' ** Called by modExcelFuncs.IncomeExpense_Export();
                ' ** qryIncomeExpenseReports_81_08_10 -> Microsoft Excell 2003 format; for Expense, 8 fields.
1340            strMacro = "mcrExcelExport_IncomeExpense_02_08"
1350          End Select
1360        Case 9&
1370          Select Case strMode
              Case "Income"
                ' ** Income: Dummy single record, with 9 fields; For Export.
                ' ** Called by modExcelFuncs.IncomeExpense_Export();
                ' ** qryIncomeExpenseReports_81_09_05 -> Microsoft Excell 2003 format; for 9 Income, fields.
1380            strMacro = "mcrExcelExport_IncomeExpense_01_09"
1390          Case "Expense"
                ' ** Expense: Dummy single record, with 9 fields; For Export.
                ' ** Called by modExcelFuncs.IncomeExpense_Export();
                ' ** qryIncomeExpenseReports_81_09_10 -> Microsoft Excell 2003 format; for 9 Expense, fields.
1400            strMacro = "mcrExcelExport_IncomeExpense_02_09"
1410          End Select
1420        End Select

            ' ** Export a query with one dummy record, just so we have something to export to.
1430        DoCmd.RunMacro strMacro
1440        DoEvents

1450        If FileExists(CurrentAppPath & LNK_SEP & "IncomeExpense_xxx.xls") = True Or _
                FileExists(strRptPath & LNK_SEP & "IncomeExpense_xxx.xls") = True Then  ' ** Module Function: modFileUtilities.

1460          If FileExists(CurrentAppPath & LNK_SEP & "IncomeExpense_xxx.xls") = True Then  ' ** Module Function: modFileUtilities.
1470            Name (CurrentAppPath & LNK_SEP & "IncomeExpense_xxx.xls") As (strPathFile)
                ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
1480          Else
1490            Name (strRptPath & LNK_SEP & "IncomeExpense_xxx.xls") As (strPathFile)
                ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
1500          End If
1510          DoEvents

1520          strSheetName = strMode & " Report"

      #If IsDev Then
1530          Set xlApp = New Excel.Application              ' ** Early Binding.
      #Else
1540          Set xlApp = CreateObject("Excel.Application")  ' ** Late Binding.
      #End If
1550          blnExcelOpen = True

1560          With xlApp

1570            .Visible = False
1580            .DisplayAlerts = False
1590            .Interactive = False

1600            Set wbk = xlApp.Workbooks.Open(strPathFile)
1610            With wbk
1620              If .Worksheets.Count > 0 Then
1630                Set wks = .Worksheets(1)
1640                With wks

1650                  .Name = strSheetName

1660                  strLastCell = .Cells.SpecialCells(xlCellTypeLastCell).Address  '$H$22205
1670                  strLastCell = Rem_Dollar(strLastCell)  ' ** Module Function: modStringFuncs.
1680                  strRng_All = "A1:" & strLastCell
1690                  Set rng = .Range(strRng_All)
1700                  lngCols = rng.Columns.Count
1710                  Set rng = Nothing

1720                  strLastCol = Left(strLastCell, 1)  ' ** Assumes single-letter address (26 or fewer columns).
1730                  strLastRow = Mid(strLastCell, 2)
1740                  strRng_Header = "A1:" & strLastCol & "1"
1750                  strRng_Title = "A2:B3"

1760                  strRng_Text1 = vbNullString: strRng_Text2 = vbNullString
1770                  If lngCols = lngFlds Then

1780                    Select Case lngFlds
                        Case 6&
                          ' ** Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
1790                      strRng_Text1 = "A4:B" & strLastRow
1800                      strRng_Values = "C4:E" & strLastRow
1810                      strRng_Text2 = "F4:F" & strLastRow
1820                    Case 8&
                          ' ** Account Num;Name;Date;Journal Type;Description;Income Cash;Principal Cash;Cost;
1830                      strRng_Text1 = "A4:E" & strLastRow
1840                      strRng_Values = "F4:H" & strLastRow
1850                    Case 9&
                          ' ** Account Num;Name;Type;Date;Journal Type;Description;Income Cash;Principal Cash;Cost;
1860                      strRng_Text1 = "A4:F" & strLastRow
1870                      strRng_Values = "G4:I" & strLastRow
1880                    End Select

1890                    Set rst = qdf1.OpenRecordset
1900                    If rst.BOF = True And rst.EOF = True Then
1910                      blnRetVal = False
1920                      DoCmd.Hourglass False
1930                      MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
                          ' ** Delete dummy Excel file.
1940                    Else
1950                      rst.MoveLast
1960                      lngRecs = rst.RecordCount
1970                      rst.MoveFirst

1980                      lngThisRow = 3&
1990                      For lngX = 1& To lngRecs
2000                        lngThisRow = lngThisRow + 1&
2010                        lngThisCol = 0&
2020                        For lngY = 1& To lngCols
2030                          lngThisCol = lngThisCol + 1&
2040                          strRng_ThisCell = Chr(64 + lngThisCol) & CStr(lngThisRow)
2050                          Set rng = .Range(strRng_ThisCell)
2060                          rng.Value = rst.Fields(lngThisCol)
2070                        Next  ' ** lngY.
2080                        If lngX < lngRecs Then rst.MoveNext
2090                      Next  ' ** lngX.

2100                    End If
2110                    rst.Close
2120                    Set rst = Nothing

2130                  Else
                        ' ** Columns don't match fields.
2140                    blnRetVal = False
2150                    Beep
2160                    MsgBox "The number of Excel columns does not match the number of fields exported.", _
                          vbInformation + vbOKOnly, "Column/Field Mismatch"
2170                  End If  ' ** lngCols/lngFlds.

2180                End With  ' ** wks.
2190                Set wks = Nothing
2200              End If  ' ** Count.
2210              .Save  ' ** wbk.Close SaveChanges:=True
2220              .Close
2230            End With  ' ** wbk.
2240            Set wbk = Nothing

2250            .DisplayAlerts = True
2260            .Interactive = True
2270            .Quit

2280          End With  ' ** xlApp.

2290          If blnRetVal = True Then
2300            Stop

                'If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
                '  DoEvents
                '  If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                '    EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                '  End If
                '  DoEvents
                '  OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
                'End If

2310          End If  ' ** blnRetVal.

2320        Else
              ' ** Problem exporting dummy record.
2330          blnRetVal = False

2340        End If  ' ** FileExists().

2350      End With  ' ** qdf1.
2360      Set qdf1 = Nothing

2370      If strMode = "Expense" Then
2380        DoCmd.DeleteObject acQuery, "qryIncomeExpenseReports_29"  ' ** qryIncomeExpenseReports_29_tmp
2390        .QueryDefs.Refresh
2400        DoEvents
2410        DoCmd.CopyObject , "qryIncomeExpenseReports_29", acQuery, "qryIncomeExpenseReports_29_norm."
2420        .QueryDefs.Refresh
2430        DoEvents
2440        DoCmd.DeleteObject acQuery, "qryIncomeExpenseReports_31"  ' ** qryIncomeExpenseReports_31_tmp
2450        .QueryDefs.Refresh
2460        DoEvents
2470        DoCmd.CopyObject , "qryIncomeExpenseReports_31", acQuery, "qryIncomeExpenseReports_31_norm."
2480        .QueryDefs.Refresh
2490        DoEvents
2500      End If

2510      .Close
2520    End With  ' ** dbs.

      #End If

EXITP:
      #If NoExcel Then
        ' ** Skip these.
      #Else
2530    Set xlApp = Nothing
2540    Set wbk = Nothing
2550    Set wks = Nothing
2560    Set rng = Nothing
2570    Set fnt = Nothing
2580    Set bdr = Nothing
2590    Set rst = Nothing
2600    Set qdf1 = Nothing
2610    Set qdf2 = Nothing
2620    Set dbs = Nothing
      #End If
2630    IncomeExpense_Export = blnRetVal
2640    Exit Function

ERRH:
2650    DoCmd.Hourglass False
2660    blnRetVal = False
      #If NoExcel Then
        ' ** Skip.
      #Else
2670    If blnExcelOpen = True Then
2680      wbk.Close
2690      xlApp.Quit
2700    End If
      #End If
2710    Select Case ERR.Number
        Case Else
2720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2730    End Select
2740    Resume EXITP

End Function

Public Function Qry_Chk() As Boolean
' ** The whole thing was about an extra long Ledger description!
' ** Called by:
' **   {not called}

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngQrys As Long, arr_varQry As Variant
        Dim lngFlds As Long, arr_varFld() As Variant
        Dim lngAllFlds As Long, arr_varAllFld() As Variant
        Dim lngMaxFlds As Long, blnMaxFlds As Boolean, lngMaxLen As Long
        Dim strFind As String
        Dim blnFound As Boolean
        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngW As Long, lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        'Const Q_DID  As Integer = 0
        'Const Q_DNAM As Integer = 1
        Const Q_QID  As Integer = 2
        Const Q_QNAM As Integer = 3
        'Const Q_DSC  As Integer = 4
        'Const Q_SORT As Integer = 5

        ' ** Array: arr_varFld().
        Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const F_NAM As Integer = 0
        Const F_TYP As Integer = 1
        'Const F_REQ As Integer = 2

        ' ** Array: arr_varAllFld().
        Const A_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const A_FNAM As Integer = 0
        Const A_FTYP As Integer = 1
        Const A_QIDS As Integer = 2
        Const A_CNT  As Integer = 3

2810    blnRetVal = True

2820    Set dbs = CurrentDb
2830    With dbs
2840      Set qdf = .QueryDefs("qryIncomeExpenseReports_80")
2850      Set rst = qdf.OpenRecordset
2860      With rst
2870        .MoveLast
2880        lngQrys = .RecordCount
2890        .MoveFirst
2900        arr_varQry = .GetRows(lngQrys)
            ' ******************************************************
            ' ** Array: arr_varQry()
            ' **
            ' **   Field  Element  Name                 Constant
            ' **   =====  =======  ===================  ==========
            ' **     1       0     dbs_id               Q_DID
            ' **     2       1     dbs_name             Q_DNAM
            ' **     3       2     qry_id               Q_QID
            ' **     4       3     qry_name             Q_QNAM
            ' **     5       4     qry_description      Q_DSC
            ' **     6       5     sort                 Q_SORT
            ' **
            ' ******************************************************
2910        .Close
2920      End With
2930      Set rst = Nothing
2940      Set qdf = Nothing
2950      .Close
2960    End With
2970    Set dbs = Nothing

2980    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
2990    DoEvents

3000    Debug.Print "'QRYS: " & CStr(lngQrys)
3010    DoEvents

3020    lngMaxFlds = 0&: blnMaxFlds = False

3030    For lngW = 0& To (lngQrys - 1&)

3040      lngFlds = 0&
3050      ReDim arr_varFld(F_ELEMS, 0)
3060      strFind = arr_varQry(Q_QNAM, lngW)

3070      blnRetVal = Tbl_Fld_List(arr_varFld, strFind)  ' ** Module Function: modXAdminFuncs.
3080      If blnRetVal = True Then
3090        If IsEmpty(arr_varFld(F_NAM, 0)) = False Then
3100          lngFlds = UBound(arr_varFld, 2)
3110          lngFlds = lngFlds + 1&
3120          For lngX = 0& To (lngFlds - 1&)
3130            blnFound = False
3140            For lngY = 0& To (lngAllFlds - 1&)
3150              If arr_varAllFld(A_FNAM, lngY) = arr_varFld(F_NAM, lngX) And arr_varAllFld(A_FTYP, lngY) = arr_varFld(F_TYP, lngX) Then
3160                blnFound = True
3170                If arr_varAllFld(A_FTYP, lngY) = arr_varFld(F_TYP, lngX) Then
3180                  arr_varAllFld(A_CNT, lngY) = arr_varAllFld(A_CNT, lngY) + 1&
3190                  arr_varAllFld(A_QIDS, lngY) = arr_varAllFld(A_QIDS, lngY) & ";" & CStr(arr_varQry(Q_QID, lngW))
3200                End If
3210              End If
3220            Next  ' ** lngY.
3230            If blnFound = False Then
3240              lngAllFlds = lngAllFlds + 1&
3250              lngE = lngAllFlds - 1&
3260              ReDim Preserve arr_varAllFld(A_ELEMS, lngE)
3270              arr_varAllFld(A_FNAM, lngE) = arr_varFld(F_NAM, lngX)
3280              arr_varAllFld(A_FTYP, lngE) = arr_varFld(F_TYP, lngX)
3290              arr_varAllFld(A_QIDS, lngE) = arr_varQry(Q_QID, lngW)
3300              arr_varAllFld(A_CNT, lngE) = CLng(1)
3310            End If
3320          Next  ' ** lngX.
3330          If lngMaxFlds = 0& Then
3340            lngMaxFlds = lngFlds
3350          Else
3360            If lngFlds <> lngMaxFlds Then
3370              blnMaxFlds = True
3380              lngMaxFlds = lngFlds
3390            End If
3400          End If
3410        Else
              ' ** Array empty!
3420          Stop
3430        End If
3440      End If

3450    Next  ' ** lngW.

        'blnMaxFlds = True means they don't all have the same number of fields.

3460    Debug.Print "'FLDS: " & CStr(lngAllFlds)
3470    DoEvents

3480    lngMaxLen = 0&
3490    For lngX = 0& To (lngAllFlds - 1&)
3500      If Len(arr_varAllFld(A_FNAM, lngX)) > lngMaxLen Then lngMaxLen = Len(arr_varAllFld(A_FNAM, lngX))
3510    Next

3520    For lngX = 0& To (lngAllFlds - 1&)
3530      Debug.Print "'" & Left(arr_varAllFld(A_FNAM, lngX) & String(lngMaxLen, " "), lngMaxLen) & "  " & _
            Tbl_Fld_Type(arr_varAllFld(A_FTYP, lngX)) & "  " & CStr(arr_varAllFld(A_CNT, lngX))  ' ** Module Function: modXAdminFuncs.
3540    Next

3550    For lngW = 0& To (lngQrys - 1&)
3560      strFind = CStr(arr_varQry(Q_QID, lngW))
3570      strTmp01 = vbNullString: strTmp02 = vbNullString
3580      For lngX = 0& To (lngAllFlds - 1&)
3590        strTmp01 = arr_varAllFld(A_QIDS, lngX)
3600        intPos01 = InStr(strTmp01, strFind)
3610        If intPos01 > 0 Then
3620          intLen = Len(strTmp01)
3630          If intPos01 = 1 Then
3640            If intLen > Len(strFind) Then
3650              If Mid(strTmp01, (intPos01 + Len(strFind)), 1) = ";" Then
3660                strTmp02 = strTmp02 & arr_varAllFld(A_FNAM, lngX) & ";"
3670              Else
                    ' ** Not one of the droids we're looking for.
3680              End If
3690            Else
3700              strTmp02 = strTmp02 & arr_varAllFld(A_FNAM, lngX) & ";"
3710            End If
3720          Else
3730            If Mid(strTmp01, (intPos01 - 1), 1) = ";" Then
3740              If intPos01 + intLen <= Len(strTmp02) Then
3750                If Mid(strTmp01, (intPos01 + intLen), 1) = ";" Then
3760                  strTmp02 = strTmp02 & arr_varAllFld(A_FNAM, lngX) & ";"
3770                Else
                      ' ** Not one of the droids we're looking for.
3780                End If
3790              Else
3800                strTmp02 = strTmp02 & arr_varAllFld(A_FNAM, lngX) & ";"
3810              End If
3820            Else
                  ' ** Not one of the droids we're looking for.
3830            End If
3840          End If
3850        End If
3860      Next  ' ** lngX.
3870      Debug.Print "'" & arr_varQry(Q_QNAM, lngW) & "  FLDS: " & CharCnt(strTmp02, ";")  ' ** Module Function: modStringFuncs.
3880      Debug.Print "'  " & strTmp02
3890    Next  ' ** lngW.

3900    Debug.Print "'DONE!  " & THIS_PROC & "()"

3910    Beep

        'QRYS: 16
        'FLDS: 10
        'Account Num     dbText  4
        'Name            dbText  4
        'Type            dbText  14
        'Date            dbText  4
        'Journal Type    dbText  4
        'Description     dbText  16
        'Income Cash     dbCurrency  16
        'Principal Cash  dbCurrency  16
        'Cost            dbCurrency  16
        'SubTitle        dbText  12

        'qryIncomeExpenseReports_23  FLDS: 9
        '  Account Num;Name;Type;Date;Journal Type;Description;Income Cash;Principal Cash;Cost;
        'qryIncomeExpenseReports_37  FLDS: 9
        '  Account Num;Name;Type;Date;Journal Type;Description;Income Cash;Principal Cash;Cost;

        'qryIncomeExpenseReports_23_plain  FLDS: 8
        '  Account Num;Name;Date;Journal Type;Description;Income Cash;Principal Cash;Cost;
        'qryIncomeExpenseReports_37_plain  FLDS: 8
        '  Account Num;Name;Date;Journal Type;Description;Income Cash;Principal Cash;Cost;

        'qryIncomeExpenseReports_55_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_55c_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_55d_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_55e_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_55f_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_55g_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_65_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_65c_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_65d_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_65e_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_65f_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;
        'qryIncomeExpenseReports_65g_all  FLDS: 6
        '  Type;Description;Income Cash;Principal Cash;Cost;SubTitle;

EXITP:
3920    Set rst = Nothing
3930    Set qdf = Nothing
3940    Set dbs = Nothing
3950    Qry_Chk = blnRetVal
3960    Exit Function

ERRH:
3970    Select Case ERR.Number
        Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Function

Public Function Excel_IncExp(strPathFile As String, strMode As String) As Boolean
' ** Called by:
' **   frmRpt_IncomeExpense:
' **     cmdRevIncExp_IncomeExcel_Click()
' **     cmdRevIncExp_ExpenseExcel_Click()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Excel_IncExp"

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

      #If IsDev Then
        Dim xlApp As Excel.Application, wbk As Excel.Workbook, wks As Excel.Worksheet, rng As Excel.Range  ' ** Early Binding.
        Dim fnt As Excel.Font, bdr As Excel.Border
      #Else
        Dim xlApp As Object, wbk As Object, wks As Object, rng As Object, fnt As Object, bdr As Object     ' ** Late Binding.
      #End If
        Dim strSheetName As String, blnChkDetail As Boolean  ', strPathFile As String
        Dim strLastCell As String, strLastCol As String, strLastRow As String, strRng_All As String
        Dim strRng_Header As String, strRng_Title As String, strRng_Text As String, strRng_Values As String
        Dim lngCols As Long
        Dim blnExcelOpen As Boolean
        Dim lngX As Long, lngY As Long

      #End If
        Dim blnRetVal As Boolean

      #If NoExcel Then
        ' ** Skip.
      #Else

4110    blnRetVal = True

4120    If strPathFile <> vbNullString Then

4130      strSheetName = vbNullString
4140      Select Case strMode
          Case "Income"
4150        strSheetName = "Income Report"
4160      Case "Expense"
4170        strSheetName = "Expense Report"
4180      End Select
4190      If InStr(strPathFile, "Detail") > 0 Then
4200        blnChkDetail = True
4210        strSheetName = strSheetName & " - Detailed"
4220      Else
4230        blnChkDetail = False
4240      End If

      #If IsDev Then
4250      Set xlApp = New Excel.Application              ' ** Early Binding.
      #Else
4260      Set xlApp = CreateObject("Excel.Application")  ' ** Late Binding.
      #End If
4270      blnExcelOpen = True

4280      xlApp.Visible = False
4290      xlApp.DisplayAlerts = False
4300      xlApp.Interactive = False
4310      Set wbk = xlApp.Workbooks.Open(strPathFile)
4320      With wbk
4330        If .Worksheets.Count > 0 Then
4340          Set wks = .Worksheets(1)
4350          With wks

4360            .Name = strSheetName

4370            strLastCell = .Cells.SpecialCells(xlCellTypeLastCell).Address  '$H$22205
4380            strLastCell = Rem_Dollar(strLastCell)  ' ** Module Function: modStringFuncs.
4390            strRng_All = "A1:" & strLastCell
4400            Set rng = .Range(strRng_All)
4410            lngCols = rng.Columns.Count
4420            Set rng = Nothing

4430            strLastCol = Left(strLastCell, 1)  ' ** Assumes single-letter address (26 or fewer columns).
4440            strLastRow = Mid(strLastCell, 2)
4450            strRng_Header = "A1:" & strLastCol & "1"
4460            strRng_Title = "A2:B3"
4470            strRng_Text = "A4:" & Chr(Asc(strLastCol) - 3) & strLastRow  ' ** Assumes 3 value columns.
4480            strRng_Values = Chr(Asc(strLastCol) - 2) & "4:" & strLastCell

                ' ** Column Headers.
4490            Set rng = .Range(strRng_Header)
4500            With rng
4510              .RowHeight = 13.5  ' ** Points.
4520              .Interior.Color = 12632256  ' 192 192 192
4530              Set fnt = .Font
4540              With fnt
4550                .Name = "Arial"
4560                .Size = 10
4570                .Bold = False
4580                .Color = 0&
4590              End With
4600              Set fnt = Nothing
4610              For lngY = 1& To 5&
4620                Select Case lngY
                    Case 1&
4630                  Set bdr = .Borders(xlEdgeLeft)
4640                Case 2&
4650                  Set bdr = .Borders(xlEdgeTop)
4660                Case 3&
4670                  Set bdr = .Borders(xlEdgeBottom)
4680                Case 4&
4690                  Set bdr = .Borders(xlEdgeRight)
4700                Case 5&
4710                  Set bdr = .Borders(xlInsideVertical)
4720                End Select
4730                With bdr
4740                  .Color = 0&
4750                  .LineStyle = xlContinuous
4760                  If lngY < 5& Then
4770                    .Weight = xlMedium
4780                  Else
4790                    .Weight = xlThin
4800                  End If
4810                End With  ' ** bdr.
4820                Set bdr = Nothing
4830              Next  ' ** bdr.
4840              .HorizontalAlignment = xlCenter
4850              .VerticalAlignment = xlBottom
4860              .WrapText = False
4870            End With  ' ** rng
4880            Set rng = Nothing

                ' ** Column Widths.
4890            For lngX = 1& To lngCols
4900              Set rng = .Range(Chr(64& + lngX) & "1:" & Chr(64& + lngX) & strLastRow)
4910              With rng
4920                Select Case blnChkDetail
                    Case True
                      ' **      A        B     C     D         E             F            G             H          I
                      ' ** Account Num  Name  Type  Date  Journal Type  Description  Income Cash  Principal Cash  Cost
                      ' ** ===========  ====  ====  ====  ============  ===========  ===========  ==============  ====
                      ' **      1        2     3     4         5             6            7             8          9
4930                  Select Case lngX
                      Case 2&
4940                    .ColumnWidth = 22  ' ** Font-based.
4950                  Case 3&
4960                    .ColumnWidth = 30
4970                  Case 4&
4980                    .ColumnWidth = 10
4990                  Case 5&
5000                    .ColumnWidth = 12
5010                  Case 6&
5020                    .ColumnWidth = 75
5030                  Case Else
5040                    .ColumnWidth = 15
5050                  End Select
5060                Case False
                      ' **      A        B     C         D             E            F             G          H
                      ' ** Account Num  Name  Date  Journal Type  Description  Income Cash  Principal Cash  Cost
                      ' ** ===========  ====  ====  ============  ===========  ===========  ==============  ====
                      ' **      1        2     3         4             5            6             7          8
5070                  Select Case lngX
                      Case 2&
5080                    .ColumnWidth = 22  ' ** Font-based.
5090                  Case 5&
5100                    .ColumnWidth = 75
5110                  Case Else
5120                    .ColumnWidth = 15
5130                  End Select
5140                End Select
5150              End With
5160            Next
5170            Set rng = Nothing

                ' ** Report Title and Period.
5180            Set rng = .Range(strRng_Title)
5190            With rng
5200              .RowHeight = 13.5  ' ** Points.
5210              Set fnt = .Font
5220              With fnt
5230                .Name = "Arial"
5240                .Size = 10
5250                .Bold = False
5260                .Color = 0&
5270              End With
5280              Set fnt = Nothing
5290              .HorizontalAlignment = xlLeft
5300              .VerticalAlignment = xlBottom
5310            End With
5320            Set rng = Nothing

                ' ** Report Text.
5330            Set rng = .Range(strRng_Text)
5340            With rng
5350              .RowHeight = 13.5  ' ** Points.
5360              Set fnt = .Font
5370              With fnt
5380                .Name = "Arial"
5390                .Size = 10
5400                .Bold = False
5410                .Color = 0&
5420              End With
5430              Set fnt = Nothing
5440              .HorizontalAlignment = xlLeft
5450              .VerticalAlignment = xlBottom
5460            End With
5470            Set rng = Nothing

                ' ** Report Values.
5480            Set rng = .Range(strRng_Values)
5490            With rng
5500              Set fnt = .Font
5510              With fnt
5520                .Name = "Arial"
5530                .Size = 10
5540                .Bold = False
5550                .Color = 0&
5560              End With
5570              Set fnt = Nothing
5580              .HorizontalAlignment = xlRight
5590              .VerticalAlignment = xlBottom
5600            End With
5610            Set rng = Nothing

5620          End With  ' ** wks.
5630          Set wks = Nothing
5640        End If  ' ** Count.
5650        .Save  ' ** wbk.Close SaveChanges:=True
5660        .Close
5670      End With  ' ** wbk.
5680      Set wbk = Nothing
5690      xlApp.DisplayAlerts = True
5700      xlApp.Interactive = True
5710      xlApp.Quit
5720    End If  ' ** vbNullString.

5730    Beep

      #End If

        ' ** XlCellType enumeration:
        ' **   -4175  xlCellTypeSameValidation        Cells having the same validation criteria.
        ' **   -4174  xlCellTypeAllValidation         Cells having validation criteria.
        ' **   -4173  xlCellTypeSameFormatConditions  Cells having the same format.
        ' **   -4172  xlCellTypeAllFormatConditions   Cells of any format.
        ' **   -4144  xlCellTypeComments              Cells containing notes.
        ' **   -4123  xlCellTypeFormulas              Cells containing formulas.
        ' **       2  xlCellTypeConstants             Cells containing constants.
        ' **       4  xlCellTypeBlanks                Empty cells.
        ' **      11  xlCellTypeLastCell              The last cell in the used range.
        ' **      12  xlCellTypeVisible               All visible cells.

        ' ** XlSpecialCellsValue enumeration:
        ' **    1  xlNumbers
        ' **    2  xlTextValues
        ' **    4  xlLogical
        ' **   16  xlErrors

        ' ** XlDirection enumeration:  (Excel 2007)
        ' **   -4162  xlUp       Up.
        ' **   -4161  xlToRight  To right.
        ' **   -4159  xlToLeft   To left.
        ' **   -4121  xlDown     Down.

        ' ** Borders enumeration:
        ' **    5  xlDiagonalDown
        ' **    6  xlDiagonalUp
        ' **    7  xlEdgeLeft
        ' **    8  xlEdgeTop
        ' **    9  xlEdgeBottom
        ' **   10  xlEdgeRight
        ' **   11  xlInsideVertical
        ' **   12  xlInsideHorizontal

        ' ** HorizontalAlignment enumeration:
        ' **   -4152  xlRight
        ' **   -4131  xlLeft
        ' **   -4130  xlJustify
        ' **   -4117  xlDistributed
        ' **   -4108  xlCenter

        ' ** VerticalAlignment enumeration:
        ' **   -4160  xlTop
        ' **   -4130  xlJustify
        ' **   -4117  xlDistributed
        ' **   -4108  xlCenter
        ' **   -4107  xlBottom

        ' ** XlLineStyle enumeration:
        ' **   -4142  xlLineStyleNone  No line.
        ' **   -4126  xlGray75         75% gray pattern.
        ' **   -4125  xlGray50         50% gray pattern.
        ' **   -4124  xlGray25         25% gray pattern.
        ' **   -4119  xlDouble         Double line.
        ' **   -4118  xlDot            Dotted line.
        ' **   -4115  xlDash           Dashed line.
        ' **   -4105  xlAutomatic      Excel applies automatic settings, such as a color, to the specified object.
        ' **       1  xlContinuous     Continuous line.
        ' **       4  xlDashDot        Alternating dashes and dots.
        ' **       5  xlDashDotDot     Dash followed by two dots.
        ' **      13  xlSlantDashDot   Slanted dashes.
        ' **      17  xlGray16         16% gray pattern.
        ' **      18  xlGray8          8% gray pattern.

        ' ** XlBorderWeight enumeration:
        ' **   -4138  xlMedium    Medium.
        ' **       1  xlHairline  Hairline (thinnest border).
        ' **       2  xlThin      Thin.
        ' **       4  xlThick     Thick (widest border).

EXITP:
      #If NoExcel Then
        ' ** Skip
      #Else
5740    Set bdr = Nothing
5750    Set fnt = Nothing
5760    Set rng = Nothing
5770    Set wks = Nothing
5780    Set wbk = Nothing
5790    Set xlApp = Nothing
      #End If
5800    Excel_IncExp = blnRetVal
5810    Exit Function

ERRH:
5820    blnRetVal = False
      #If NoExcel Then
        ' ** Skip
      #Else
5830    If blnExcelOpen = True Then
5840      wbk.Close
5850      xlApp.Quit
5860    End If
      #End If
5870    Select Case ERR.Number
        Case Else
5880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5890    End Select
5900    Resume EXITP

End Function

Public Function Excel_NameOnly(strPathFile As String, strMode As String) As Boolean
' ** Change the name of the exported worksheet.
' ** Called by:
' **

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "Excel_NameOnly"

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

      #If IsDev Then
        Dim xlApp As Excel.Application, wbk As Excel.Workbook, wks As Excel.Worksheet  ' ** Early Binding.
      #Else
        Dim xlApp As Object, wbk As Object, wks As Object                              ' ** Late Binding.
      #End If
        Dim strSheetName As String
        Dim blnExcelOpen As Boolean

      #End If
        Dim blnRetVal As Boolean

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

6010    blnRetVal = True

6020    If strPathFile <> vbNullString Then

6030      strSheetName = vbNullString
6040      Select Case strMode
          Case "Account Balance"
6050        strSheetName = "Account Balance"
6060      Case "Account Profile"
6070        strSheetName = "Account Profile"
6080      Case "Account Profiles"
6090        strSheetName = "Account Profiles"
6100      Case "Account Reviews"
6110        strSheetName = "Account Reviews"
6120      Case "Accounts New"
6130        strSheetName = "New Accounts"
6140      Case "Accounts Closed"
6150        strSheetName = "Closed Accounts"
6160      Case "Archived"
6170        strSheetName = "Archived Transactions"
6180      Case "Assets Purchased"
6190        strSheetName = "Purchased Assets"
6200      Case "Assets Sold"
6210        strSheetName = "Sold Assets"
6220      Case "Expense"
6230        strSheetName = "Expense Report"
6240      Case "Expense Summary"
6250        strSheetName = "Expense Report - Summary"
6260      Case "History"
6270        strSheetName = "Asset History"
6280      Case "Holdings"
6290        strSheetName = "Holdings"
6300      Case "Holdings All"
6310        strSheetName = "Holdings - All"
6320      Case "Income"
6330        strSheetName = "Income Report"
6340      Case "Income Summary"
6350        strSheetName = "Income Report - Summary"
6360      Case "Locations"
6370        strSheetName = "Locations"
6380      Case "Maturity"
6390        strSheetName = "Security Maturity"
6400      Case "Unrealized"
6410        strSheetName = "Unrealized Gain & Loss"
6420      End Select
6430      If InStr(strPathFile, "Detail") > 0 Then
6440        strSheetName = strSheetName & " - Detailed"
6450      End If

      #If IsDev Then
6460      Set xlApp = New Excel.Application              ' ** Early Binding.
      #Else
6470      Set xlApp = CreateObject("Excel.Application")  ' ** Late Binding.
      #End If
6480      blnExcelOpen = True

6490      xlApp.Visible = False
6500      xlApp.DisplayAlerts = False
6510      xlApp.Interactive = False
6520      Set wbk = xlApp.Workbooks.Open(strPathFile)
6530      With wbk
6540        If .Worksheets.Count > 0 Then
6550          Set wks = .Worksheets(1)
6560          With wks

6570            .Name = strSheetName

6580          End With  ' ** wks.
6590          Set wks = Nothing
6600        End If  ' ** Count.
6610        .Save  ' ** wbk.Close SaveChanges:=True
6620        .Close
6630      End With  ' ** wbk.
6640      Set wbk = Nothing
6650      DoEvents
6660      xlApp.DisplayAlerts = True
6670      xlApp.Interactive = True
6680      xlApp.Quit

6690    End If  ' ** vbNullString.

      #End If

EXITP:
      #If NoExcel Then
        ' ** Skip.
      #Else
6700    Set wks = Nothing
6710    Set wbk = Nothing
6720    Set xlApp = Nothing
      #End If
6730    Excel_NameOnly = blnRetVal
6740    Exit Function

ERRH:
6750    blnRetVal = False
      #If NoExcel Then
        ' ** Skip the whole function.
      #Else
6760    If blnExcelOpen = True Then
6770      wbk.Close
6780      xlApp.Quit
6790    End If
      #End If
6800    Select Case ERR.Number
        Case Else
6810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6820    End Select
6830    Resume EXITP

End Function

Public Function Excel_Trans(strPathFile As String) As Boolean
' ** Export statement transactions.
' ** Called by:
' **   frmStatementParameters:
' **     cmdTransactionsExcel_Click()

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "Excel_Trans"

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

      #If IsDev Then
        Dim xlApp As Excel.Application, wbk As Excel.Workbook, wks As Excel.Worksheet, rng As Excel.Range              ' ** Early Binding.
        Dim fnt As Excel.Font, bdr As Excel.Border, cel As Excel.Range
      #Else
        Dim xlApp As Object, wbk As Object, wks As Object, rng As Object, fnt As Object, bdr As Object, cel As Object  ' ** Late Binding.
      #End If
        Dim strSheetName As String, blnChkDetail As Boolean
        Dim strLastCell As String, strLastCol As String, strLastRow As String, strRng_All As String
        Dim strRng_Header As String, strRng_Title As String
        Dim strRng_Text1 As String, strRng_Text2 As String, strRng_Text3 As String, strRng_Text4 As String, strRng_Values As String
        Dim strRng_Dates1 As String, strRng_Dates2 As String, strRng_Dates3 As String, strRng_Dates4 As String
        Dim strRng_Integers1 As String, strRng_Integers2 As String, strRng_Decimals As String, strRng_Percents As String
        Dim lngCols As Long
        Dim blnExcelOpen As Boolean
        Dim lngX As Long, lngY As Long

      #End If
        Dim blnRetVal As Boolean

      #If NoExcel Then
        ' ** Skip.
      #Else

6910    blnRetVal = True

        'strPathFile = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Misc_Reports\
        'strPathFile = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Misc_Reports\

6920    If strPathFile <> vbNullString Then

6930      strSheetName = vbNullString
6940      If InStr(strPathFile, "Date") > 0 Then
6950        strSheetName = "Transactions - By Date"
6960      ElseIf InStr(strPathFile, "Type") > 0 Then
6970        strSheetName = "Transactions - By Type"
6980      End If
          'If InStr(strPathFile, "Detail") > 0 Then
          '  blnChkDetail = True
          '  strSheetName = strSheetName & " - Detailed"
          'Else
6990      blnChkDetail = False
          'End If

      #If IsDev Then
7000      Set xlApp = New Excel.Application              ' ** Early Binding.
      #Else
7010      Set xlApp = CreateObject("Excel.Application")  ' ** Late Binding.
      #End If
7020      blnExcelOpen = True

7030      xlApp.Visible = False
7040      xlApp.DisplayAlerts = False
7050      xlApp.Interactive = False
7060      Set wbk = xlApp.Workbooks.Open(strPathFile)
7070      With wbk
7080        If .Worksheets.Count > 0 Then
7090          Set wks = .Worksheets(1)
7100          With wks

7110            .Name = strSheetName

7120            strLastCell = .Cells.SpecialCells(xlCellTypeLastCell).Address  '$H$22205
7130            strLastCell = Rem_Dollar(strLastCell)  ' ** Module Function: modStringFuncs.
7140            strRng_All = "A1:" & strLastCell
7150            Set rng = .Range(strRng_All)
7160            lngCols = rng.Columns.Count
7170            Set rng = Nothing

7180            strLastCol = Left(strLastCell, 1)  ' ** Assumes single-letter address (26 or fewer columns).
7190            strLastRow = Mid(strLastCell, 2)
7200            strRng_Header = "A1:" & strLastCol & "1"
7210            strRng_Title = "A2:B3"
                ' ** All Cols:
                ' **   A - P
                ' ** Text Cols:
                ' **   A, B, D, G, O
7220            strRng_Text1 = "A4:B" & strLastRow
7230            strRng_Text2 = "A4:D" & strLastRow
7240            strRng_Text3 = "A4:G" & strLastRow
7250            strRng_Text4 = "A4:O" & strLastRow
                ' ** Value Cols:
                ' **   J, K, L
7260            strRng_Values = "J4:L" & strLastRow
                ' ** Date Cols:
                ' **   C, E, I, N
7270            strRng_Dates1 = "C4:C" & strLastRow
7280            strRng_Dates2 = "E4:E" & strLastRow
7290            strRng_Dates3 = "I4:I" & strLastRow
7300            strRng_Dates4 = "N4:N" & strLastRow
                ' ** Integer Cols:
                ' **   M, P
7310            strRng_Integers1 = "M4:M" & strLastRow
7320            strRng_Integers2 = "P4:P" & strLastRow
                ' ** Decimal Cols:
                ' **   F
7330            strRng_Decimals = "F4:F" & strLastRow
                ' ** Percent Cols:
                ' **   H
7340            strRng_Percents = "H4:H" & strLastRow

                ' ** Column Headers.
7350            Set rng = .Range(strRng_Header)
7360            With rng
7370              .RowHeight = 13.5  ' ** Points.
7380              .Interior.Color = 12632256  ' 192 192 192
7390              Set fnt = .Font
7400              With fnt
7410                .Name = "Arial"
7420                .Size = 10
7430                .Bold = False
7440                .Color = 0&
7450              End With
7460              Set fnt = Nothing
7470              For lngY = 1& To 5&
7480                Select Case lngY
                    Case 1&
7490                  Set bdr = .Borders(xlEdgeLeft)
7500                Case 2&
7510                  Set bdr = .Borders(xlEdgeTop)
7520                Case 3&
7530                  Set bdr = .Borders(xlEdgeBottom)
7540                Case 4&
7550                  Set bdr = .Borders(xlEdgeRight)
7560                Case 5&
7570                  Set bdr = .Borders(xlInsideVertical)
7580                End Select
7590                With bdr
7600                  .Color = 0&
7610                  .LineStyle = xlContinuous
7620                  If lngY < 5& Then
7630                    .Weight = xlMedium
7640                  Else
7650                    .Weight = xlThin
7660                  End If
7670                End With  ' ** bdr.
7680                Set bdr = Nothing
7690              Next  ' ** bdr.
7700              .HorizontalAlignment = xlCenter
7710              .VerticalAlignment = xlBottom
7720              .WrapText = False
7730            End With  ' ** rng
7740            Set rng = Nothing

                ' ** Column Widths.
7750            For lngX = 1& To lngCols
7760              Set rng = .Range(Chr(64& + lngX) & "1:" & Chr(64& + lngX) & strLastRow)
7770              With rng
                    ' ** Width is font-based, and dependent on column header.
                    ' **      A             B              C                 D              E           F            G        H       I           J              K         L          M              N          O          P
                    ' ** Account Num  Account Name  Transaction Date  Transaction Type  Trade Date  Share/Face  Description  Rate  Due Date  Income Cash  Principal Cash  Cost  Journal Number  Date Posted  User ID  Asset Number
                    ' ** ===========  ============  ================  ================  ==========  ==========  ===========  ====  ========  ===========  ==============  ====  ==============  ===========  =======  ============
                    ' **      1             2              3                 4              5           6            7        8       9           10             11        12         13             14         15         16
7780                Select Case lngX
                    Case 2&
7790                  .ColumnWidth = 35  ' ** Account Name.
7800                Case 4&
7810                  .ColumnWidth = 16  ' ** Transaction Type.
7820                Case 5&, 9&, 14&
7830                  .ColumnWidth = 12  ' ** Trade Date, Due Date, Date Posted.
7840                Case 7&
7850                  .ColumnWidth = 75  ' ** Description.
7860                Case 8&
7870                  .ColumnWidth = 10  ' ** Rate.
7880                Case Else
7890                  .ColumnWidth = 15
7900                End Select
7910              End With
7920            Next
7930            Set rng = Nothing

                ' ** Report Title and Period.
7940            Set rng = .Range(strRng_Title)
7950            With rng
7960              .RowHeight = 13.5  ' ** Points.
7970              Set fnt = .Font
7980              With fnt
7990                .Name = "Arial"
8000                .Size = 10
8010                .Bold = False
8020                .Color = 0&
8030              End With  ' ** fnt.
8040              Set fnt = Nothing
8050              .HorizontalAlignment = xlLeft
8060              .VerticalAlignment = xlBottom
8070            End With  ' ** rng.
8080            Set rng = Nothing

                ' ** Report Text.
8090            For lngX = 1& To 4&
8100              Select Case lngX
                  Case 1&
8110                Set rng = .Range(strRng_Text1)
8120              Case 2&
8130                Set rng = .Range(strRng_Text2)
8140              Case 3&
8150                Set rng = .Range(strRng_Text3)
8160              Case 4&
8170                Set rng = .Range(strRng_Text4)
8180              End Select
8190              With rng
8200                .RowHeight = 13.5  ' ** Points.
8210                Set fnt = .Font
8220                With fnt
8230                  .Name = "Arial"
8240                  .Size = 10
8250                  .Bold = False
8260                  .Color = 0&
8270                End With  ' ** fnt.
8280                Set fnt = Nothing
8290                .HorizontalAlignment = xlLeft
8300                .VerticalAlignment = xlBottom
8310                If lngX = 1& Then
                      ' ** If accountno signals 'Number as Text' error, dismiss it.
8320                  For Each cel In rng
8330                    With cel
8340                      .Select
8350                      If xlApp.ErrorCheckingOptions.NumberAsText Then
8360                        xlApp.ErrorCheckingOptions.NumberAsText = False
8370                      End If
8380                    End With  ' ** cel.
8390                  Next  ' ** cel.
8400                End If
8410              End With  ' ** rng.
8420              Set rng = Nothing
8430            Next  ' ** lngX.

                ' ** Report Values.
8440            Set rng = .Range(strRng_Values)
8450            With rng
8460              Set fnt = .Font
8470              With fnt
8480                .Name = "Arial"
8490                .Size = 10
8500                .Bold = False
8510                .Color = 0&
8520              End With  ' ** fnt.
8530              Set fnt = Nothing
8540              .HorizontalAlignment = xlRight
8550              .VerticalAlignment = xlBottom
8560              For Each cel In rng
8570                With cel
8580                  .Select
8590                  If xlApp.ErrorCheckingOptions.NumberAsText Then
8600                    If Trim(.Value) <> vbNullString Then
8610                      xlApp.WorksheetFunction.Trim (.Value)
8620                      .Value = .Value + 0
8630                    End If
8640                  End If
8650                End With  ' ** cel.
8660              Next  ' ** cel.
8670              .NumberFormat = "$#,##0.00;($#,##0.00)"
8680            End With  ' ** rng.
8690            Set rng = Nothing

                ' ** Report Dates.
8700            For lngX = 1& To 4&
8710              Select Case lngX
                  Case 1&
8720                Set rng = .Range(strRng_Dates1)
8730              Case 2&
8740                Set rng = .Range(strRng_Dates2)
8750              Case 3&
8760                Set rng = .Range(strRng_Dates3)
8770              Case 4&
8780                Set rng = .Range(strRng_Dates4)
8790              End Select
8800              With rng
8810                Set fnt = .Font
8820                With fnt
8830                  .Name = "Arial"
8840                  .Size = 10
8850                  .Bold = False
8860                  .Color = 0&
8870                End With  ' ** fnt.
8880                Set fnt = Nothing
8890                .HorizontalAlignment = xlLeft
8900                .VerticalAlignment = xlBottom
8910                .NumberFormat = "mm/dd/yyyy"  ' ** TextDate error only identifies 2-digit year.
8920              End With  ' ** rng.
8930              Set rng = Nothing
8940            Next  ' ** lngX.

                ' ** Report Integers.
8950            For lngX = 1& To 2&
8960              Select Case lngX
                  Case 1&
8970                Set rng = .Range(strRng_Integers1)
8980              Case 2&
8990                Set rng = .Range(strRng_Integers2)
9000              End Select
9010              With rng
9020                Set fnt = .Font
9030                With fnt
9040                  .Name = "Arial"
9050                  .Size = 10
9060                  .Bold = False
9070                  .Color = 0&
9080                End With  ' ** fnt.
9090                Set fnt = Nothing
9100                .HorizontalAlignment = xlLeft
9110                .VerticalAlignment = xlBottom
9120                For Each cel In rng
9130                  With cel
9140                    .Select
9150                    If xlApp.ErrorCheckingOptions.NumberAsText Then
9160                      If Trim(.Value) <> vbNullString Then
9170                        xlApp.WorksheetFunction.Trim (.Value)
9180                        .Value = .Value + 0
9190                      End If
9200                    End If
9210                  End With  ' ** cel.
9220                Next  ' ** cel.
9230                .NumberFormat = "#0"
9240              End With  ' ** rng.
9250              Set rng = Nothing
9260            Next  ' ** lngX.

                ' ** Report Decimals.
9270            Set rng = .Range(strRng_Decimals)
9280            With rng
9290              Set fnt = .Font
9300              With fnt
9310                .Name = "Arial"
9320                .Size = 10
9330                .Bold = False
9340                .Color = 0&
9350              End With  ' ** fnt.
9360              Set fnt = Nothing
9370              .HorizontalAlignment = xlRight
9380              .VerticalAlignment = xlBottom
9390              For Each cel In rng
9400                With cel
9410                  .Select
9420                  If xlApp.ErrorCheckingOptions.NumberAsText Then
9430                    If Trim(.Value) <> vbNullString Then
9440                      xlApp.WorksheetFunction.Trim (.Value)
9450                      .Value = .Value + 0
9460                    End If
9470                  End If
9480                End With  ' ** cel.
9490              Next  ' ** cel.
9500              .NumberFormat = "#,##0.0000;-#,##0.0000"
9510            End With  ' ** rng.
9520            Set rng = Nothing

                ' ** Report Percents.
9530            Set rng = .Range(strRng_Percents)
9540            With rng
9550              Set fnt = .Font
9560              With fnt
9570                .Name = "Arial"
9580                .Size = 10
9590                .Bold = False
9600                .Color = 0&
9610              End With
9620              Set fnt = Nothing
9630              .HorizontalAlignment = xlRight
9640              .VerticalAlignment = xlBottom
9650              For Each cel In rng
9660                With cel
9670                  .Select
9680                  If xlApp.ErrorCheckingOptions.NumberAsText Then
9690                    If Trim(.Value) <> vbNullString Then
9700                      xlApp.WorksheetFunction.Trim (.Value)
9710                      .Value = .Value + 0
9720                    End If
9730                  End If
9740                End With  ' ** cel.
9750              Next  ' ** cel.
9760              .NumberFormat = "#0.0000%"
9770            End With  ' ** rng.
9780            Set rng = Nothing
9790            .Range("A2").Select
9800          End With  ' ** wks.
9810          Set wks = Nothing
9820        End If  ' ** Count.
9830        .Save  ' ** wbk.Close SaveChanges:=True
9840        .Close
9850      End With  ' ** wbk.
9860      Set wbk = Nothing
9870      xlApp.DisplayAlerts = True
9880      xlApp.Interactive = True
9890      xlApp.Quit
9900    End If  ' ** vbNullString.

9910    Beep

      #End If

        ' ** XlCellType enumeration:
        ' **   -4175  xlCellTypeSameValidation        Cells having the same validation criteria.
        ' **   -4174  xlCellTypeAllValidation         Cells having validation criteria.
        ' **   -4173  xlCellTypeSameFormatConditions  Cells having the same format.
        ' **   -4172  xlCellTypeAllFormatConditions   Cells of any format.
        ' **   -4144  xlCellTypeComments              Cells containing notes.
        ' **   -4123  xlCellTypeFormulas              Cells containing formulas.
        ' **       2  xlCellTypeConstants             Cells containing constants.
        ' **       4  xlCellTypeBlanks                Empty cells.
        ' **      11  xlCellTypeLastCell              The last cell in the used range.
        ' **      12  xlCellTypeVisible               All visible cells.

        ' ** XlSpecialCellsValue enumeration:
        ' **    1  xlNumbers
        ' **    2  xlTextValues
        ' **    4  xlLogical
        ' **   16  xlErrors

        ' ** XlDirection enumeration:  (Excel 2007)
        ' **   -4162  xlUp       Up.
        ' **   -4161  xlToRight  To right.
        ' **   -4159  xlToLeft   To left.
        ' **   -4121  xlDown     Down.

        ' ** Borders enumeration:
        ' **    5  xlDiagonalDown
        ' **    6  xlDiagonalUp
        ' **    7  xlEdgeLeft
        ' **    8  xlEdgeTop
        ' **    9  xlEdgeBottom
        ' **   10  xlEdgeRight
        ' **   11  xlInsideVertical
        ' **   12  xlInsideHorizontal

        ' ** HorizontalAlignment enumeration:
        ' **   -4152  xlRight
        ' **   -4131  xlLeft
        ' **   -4130  xlJustify
        ' **   -4117  xlDistributed
        ' **   -4108  xlCenter

        ' ** VerticalAlignment enumeration:
        ' **   -4160  xlTop
        ' **   -4130  xlJustify
        ' **   -4117  xlDistributed
        ' **   -4108  xlCenter
        ' **   -4107  xlBottom

        ' ** XlLineStyle enumeration:
        ' **   -4142  xlLineStyleNone  No line.
        ' **   -4126  xlGray75         75% gray pattern.
        ' **   -4125  xlGray50         50% gray pattern.
        ' **   -4124  xlGray25         25% gray pattern.
        ' **   -4119  xlDouble         Double line.
        ' **   -4118  xlDot            Dotted line.
        ' **   -4115  xlDash           Dashed line.
        ' **   -4105  xlAutomatic      Excel applies automatic settings, such as a color, to the specified object.
        ' **       1  xlContinuous     Continuous line.
        ' **       4  xlDashDot        Alternating dashes and dots.
        ' **       5  xlDashDotDot     Dash followed by two dots.
        ' **      13  xlSlantDashDot   Slanted dashes.
        ' **      17  xlGray16         16% gray pattern.
        ' **      18  xlGray8          8% gray pattern.

        ' ** XlBorderWeight enumeration:
        ' **   -4138  xlMedium    Medium.
        ' **       1  xlHairline  Hairline (thinnest border).
        ' **       2  xlThin      Thin.
        ' **       4  xlThick     Thick (widest border).

EXITP:
      #If NoExcel Then
        ' ** Skip.
      #Else
9920    Set bdr = Nothing
9930    Set fnt = Nothing
9940    Set rng = Nothing
9950    Set wks = Nothing
9960    Set wbk = Nothing
9970    Set xlApp = Nothing
      #End If
9980    Excel_Trans = blnRetVal
9990    Exit Function

ERRH:
10000   blnRetVal = False
      #If NoExcel Then
        ' ** Skip.
      #Else
10010   If blnExcelOpen = True Then
10020     wbk.Close
10030     xlApp.Quit
10040   End If
      #End If
10050   Select Case ERR.Number
        Case Else
10060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10070   End Select
10080   Resume EXITP

End Function

Public Function Excel_AcctCon(strPathFile As String, strQryName2 As String) As Boolean
' ** Export account contacts.
' ** Called by:
' **   frmAccountContacts:
' **     cmdExcel_Click()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "Excel_AcctCon"

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

      #If IsDev Then
        Dim xlApp As Excel.Application, wbk As Excel.Workbook, wks As Excel.Worksheet, rng As Excel.Range  ' ** Early Binding.
        Dim fnt As Excel.Font, bdr As Excel.Border
      #Else
        Dim xlApp As Object, wbk As Object, wks As Object, rng As Object, fnt As Object, bdr As Object     ' ** Late Binding.
      #End If
        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strSheetName As String ', blnChkDetail As Boolean  ', strPathFile As String
        Dim strLastCell As String, strLastCol As String, strLastRow As String, strRng_All As String
        Dim lngCols As Long
        Dim lngFlds As Long, arr_varFld() As Variant
        Dim blnExcelOpen As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngE As Long

      #End If
        Dim blnRetVal As Boolean

      #If NoExcel Then
        ' ** Skip.
      #Else

        ' ** Array: arr_varFld().
        Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const F_COL  As Integer = 0
        Const F_FNAM As Integer = 1

10110   blnRetVal = True

10120   If strPathFile <> vbNullString And strQryName2 <> vbNullString Then

10130     strTmp01 = vbNullString: strTmp02 = vbNullString
10140     Set dbs = CurrentDb
10150     With dbs
10160       Set qdf = .QueryDefs(strQryName2)
10170       strTmp01 = qdf.SQL
10180       Set qdf = Nothing
10190       .Close
10200     End With
10210     Set dbs = Nothing

10220     lngFlds = 0&
10230     ReDim arr_varFld(F_ELEMS, 0)

10240     intPos01 = 0
10250     If strTmp01 <> vbNullString Then
10260       intPos01 = InStr(strTmp01, "SELECT ")  ' ** Should be 1.
10270       If intPos01 > 0 Then
10280         strTmp01 = Trim(Mid(strTmp01, 7))
10290         intPos01 = InStr(strTmp01, "FROM ")
10300         strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
10310         strTmp01 = Rem_CRLF(strTmp01)  ' ** Module Function: modStringFuncs.
10320         intPos01 = InStr(strTmp01, ",")
10330         Do While intPos01 > 0
10340           strTmp02 = Trim(Left(strTmp01, (intPos01 - 1)))
10350           strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))
10360           intPos01 = InStr(strTmp02, ".")
10370           If intPos01 > 0 Then
10380             strTmp02 = Mid(strTmp02, (intPos01 + 1))
10390           End If
10400           If Left(strTmp02, 1) = "[" Then strTmp02 = Mid(strTmp02, 2)
10410           If Right(strTmp02, 1) = "]" Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
10420           lngFlds = lngFlds + 1&
10430           lngE = lngFlds - 1&
10440           ReDim Preserve arr_varFld(F_ELEMS, lngE)
10450           arr_varFld(F_COL, lngE) = lngFlds
10460           arr_varFld(F_FNAM, lngE) = strTmp02
10470           intPos01 = InStr(strTmp01, ",")
10480           If intPos01 > 0 Then
10490             If Mid(strTmp01, (intPos01 - 4), 5) = "Last," Then
                    ' ** "Contact Last, First"
10500               intPos01 = InStr((intPos01 + 1), strTmp01, ",")
10510             ElseIf Mid(strTmp01, (intPos01 - 4), 5) = "City," Then
                    ' ** "City, State, Zip"
10520               intPos01 = InStr((intPos01 + 1), strTmp01, ",")
10530             End If
10540             If Mid(strTmp01, (intPos01 - 5), 6) = "State," Then
                    ' ** "City, State, Zip"
10550               intPos01 = InStr((intPos01 + 1), strTmp01, ",")
10560             End If
10570           End If
10580           If intPos01 = 0 Then
10590             strTmp02 = strTmp01
10600             intPos01 = InStr(strTmp02, ".")
10610             If intPos01 > 0 Then
10620               strTmp02 = Mid(strTmp02, (intPos01 + 1))
10630             End If
10640             If Left(strTmp02, 1) = "[" Then strTmp02 = Mid(strTmp02, 2)
10650             If Right(strTmp02, 1) = "]" Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
10660             lngFlds = lngFlds + 1&
10670             lngE = lngFlds - 1&
10680             ReDim Preserve arr_varFld(F_ELEMS, lngE)
10690             arr_varFld(F_COL, lngE) = lngFlds
10700             arr_varFld(F_FNAM, lngE) = strTmp02
10710             Exit Do
10720           End If
10730         Loop
10740       Else
10750         blnRetVal = False
10760       End If
10770     Else
10780       blnRetVal = False
10790     End If

10800     If blnRetVal = True Then

10810       strSheetName = "Account Contacts"

      #If IsDev Then
10820       Set xlApp = New Excel.Application              ' ** Early Binding.
      #Else
10830       Set xlApp = CreateObject("Excel.Application")  ' ** Late Binding.
      #End If
10840       blnExcelOpen = True

10850       xlApp.Visible = False
10860       xlApp.DisplayAlerts = False
10870       xlApp.Interactive = False
10880       Set wbk = xlApp.Workbooks.Open(strPathFile)
10890       With wbk
10900         If .Worksheets.Count > 0 Then
10910           Set wks = .Worksheets(1)
10920           With wks

10930             .Name = strSheetName

10940             strLastCell = .Cells.SpecialCells(xlCellTypeLastCell).Address  '$H$22205
10950             strLastCell = Rem_Dollar(strLastCell)  ' ** Module Function: modStringFuncs.
10960             strRng_All = "A1:" & strLastCell
10970             Set rng = .Range(strRng_All)
10980             lngCols = rng.Columns.Count
10990             Set rng = Nothing

11000             strLastCol = Left(strLastCell, 1)  ' ** Assumes single-letter address (26 or fewer columns).
11010             strLastRow = Mid(strLastCell, 2)

11020             For lngX = 1& To lngCols
11030               Set rng = .Range(Chr(64& + lngX) & "1:" & Chr(64& + lngX) & strLastRow)
11040               With rng
                      ' **      A             B           C        D            D               E          F         E       G     H     I     I            G                 G              H        H          I        I         H       H      J     J      K
                      ' ** Account Num  Account Name  Contact #  Name  Contact Last, First  Address 1  Address_2  Address  City  State  Zip  Zip +  City, State, Zip  City, State, Zip +  Phone 1  Phone 1 +  Phone 2  Phone 2 +  Phone  Phone +  Fax  Fax +  Email
                      ' ** ===========  ============  =========  ====  ===================  =========  =========  =======  ====  =====  ===  =====  ================  ==================  =======  =========  =======  =========  =====  =======  ===  =====  =====
                      ' **      1             2           3        4            4               5          6         5       7     8     9     9            7                 7              10       10         11       11        10      10     12    12     13
                      ' ** ? Application.ActiveSheet.Range("K1:K10").ColumnWidth
11050                 Select Case arr_varFld(F_FNAM, (lngX - 1&))
                      Case "Account Num"
                        ' ** As exported: 13.57
11060                   .ColumnWidth = 15  ' ** Font-based.
11070                 Case "Account Name"
                        ' ** As exported: 43.57
11080                   .ColumnWidth = 45
11090                 Case "Contact #"
                        ' ** As exported: 10.14
11100                   .ColumnWidth = 11
11110                 Case "Name", "Contact Last, First"
                        ' ** As exported: 39.14
11120                   .ColumnWidth = 40
11130                 Case "Address 1", "Address_2"
                        ' ** As exported: 30.29
11140                   .ColumnWidth = 32
11150                 Case "Address"
                        ' ** As exported: 47.86
11160                   .ColumnWidth = 50
11170                 Case "City"
                        ' ** As exported: 21.57
11180                   .ColumnWidth = 22
11190                 Case "State"
                        ' ** As exported: 5.71
11200                   .ColumnWidth = 6
11210                 Case "Zip", "Zip +"
                        ' ** As exported: 12.71
11220                   .ColumnWidth = 13
11230                 Case "City, State, Zip", "City, State, Zip +"
                        ' ** As exported: 39.14
11240                   .ColumnWidth = 40
11250                 Case "Phone 1", "Phone 1 +", "Phone 2", "Phone 2 +"
                        ' ** As exported: 15.43
11260                   .ColumnWidth = 16
11270                 Case "Phone", "Phone +"
                        ' ** As exported: 26
11280                   .ColumnWidth = 26
11290                 Case "Fax", "Fax +"
                        ' ** As exported: 15.43
11300                   .ColumnWidth = 16
11310                 Case "Email"
                        ' ** As exported: 30.29
11320                   .ColumnWidth = 32
11330                 End Select
11340               End With
11350             Next
11360             Set rng = Nothing

11370             .Range("A2").Select

11380           End With  ' ** wks.
11390           Set wks = Nothing
11400         End If  ' ** Count.
11410         .Save  ' ** wbk.Close SaveChanges:=True
11420         .Close
11430       End With  ' ** wbk.
11440       Set wbk = Nothing
11450       xlApp.DisplayAlerts = True
11460       xlApp.Interactive = True
11470       xlApp.Quit

11480     End If  ' ** blnRetVal.
11490   Else
11500     blnRetVal = False
11510   End If  ' ** vbNullString.

      #End If

EXITP:
      #If NoExcel Then
        ' ** Skip.
      #Else
11520   Set bdr = Nothing
11530   Set fnt = Nothing
11540   Set rng = Nothing
11550   Set wks = Nothing
11560   Set wbk = Nothing
11570   Set xlApp = Nothing
11580   Set qdf = Nothing
11590   Set dbs = Nothing
      #End If
11600   Excel_AcctCon = blnRetVal
11610   Exit Function

ERRH:
11620   blnRetVal = False
      #If NoExcel Then
        ' ** Skip.
      #Else
11630   If blnExcelOpen = True Then
11640     wbk.Close
11650     xlApp.Quit
11660   End If
      #End If
11670   Select Case ERR.Number
        Case Else
11680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11690   End Select
11700   Resume EXITP

End Function

Public Function Excel_Court(strPathFile As String) As Boolean
' ** Export court reports.
' ** Called by:
' **

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "Excel_Court"

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

      #If IsDev Then
        Dim xlApp As Excel.Application, wbk As Excel.Workbook, wks As Excel.Worksheet, rng As Excel.Range              ' ** Early Binding.
        Dim fnt As Excel.Font, bdr As Excel.Border, cel As Excel.Range, rng2 As Excel.Range
      #Else
        Dim xlApp As Object, wbk As Object, wks As Object, rng As Object, fnt As Object, bdr As Object, cel As Object  ' ** Late Binding.
        Dim rng2 As Object
      #End If
        Dim strSheetName As String
        Dim strRptTitle As String, strCourtState As String, strRptStateTitle As String
        Dim strLastCell As String, strLastCol As String, strLastRow As String, strRng_All As String
        Dim strRng_Header As String, strRng_Title As String, strRng_RptTitle As String, strRng_RptTitleRight As String
        Dim strRng_Text1 As String, strRng_Text2 As String, strRng_Text3 As String, strRng_Text4 As String, strRng_Text5 As String
        Dim strRng_Values1 As String, strRng_Values2 As String, strRng_Values3 As String, strRng_Values4 As String
        Dim strRng_Values5 As String, strRng_Values6 As String
        Dim strRng_Dates1 As String, strRng_Decimals1 As String, strRng_Decimals2 As String
        Dim lngTxtCnt As Long, lngValCnt As Long, lngDatCnt As Long, lngIntCnt As Long, lngDecCnt As Long, lngPctCnt As Long
        Dim blnColHeads As Boolean, blnColWidths As Boolean, blnRptTitlePeriod As Boolean
        Dim lngTrimTxt As Long
        Dim strRng_Header2 As String, strFirstRow2 As String, strLastRow2 As String
        Dim lngRowHeight As Long, lngRowHeight2 As Long
        Dim lngCols As Long
        Dim blnExcelOpen As Boolean, blnGrouped As Boolean, blnChkDates As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long

      #End If
        Dim blnRetVal As Boolean

      #If NoExcel Then
        ' ** Skip.
      #Else

11810   blnRetVal = True

        'strPathFile = "C:\VictorGCS_Clients\TrustAccountant\Clients\MasterTrust\CourtReport_CA_Disbursements_00187_110801_To_121231.xls"
        'strPathFile = "C:\VictorGCS_Clients\TrustAccountant\Clients\MasterTrust\CourtReport_CA_Distributions_00187_110801_To_121231.xls"
        'strPathFile = "C:\VictorGCS_Clients\TrustAccountant\Clients\MasterTrust\CourtReport_CA_Receipts_00187_110801_To_121231.xls"

11820   If strPathFile <> vbNullString Then

11830     intPos01 = InStr(strPathFile, "CourtReport_")
11840     intPos01 = InStr(intPos01, strPathFile, "_")
11850     strCourtState = Mid(strPathFile, (intPos01 + 1), 2)

      #If IsDev Then
11860     Set xlApp = New Excel.Application              ' ** Early Binding.
      #Else
11870     Set xlApp = CreateObject("Excel.Application")  ' ** Late Binding.
      #End If
11880     blnExcelOpen = True

11890     xlApp.Visible = False
11900     xlApp.DisplayAlerts = False
11910     xlApp.Interactive = False
11920     Set wbk = xlApp.Workbooks.Open(strPathFile)
11930     With wbk
11940       If .Worksheets.Count > 0 Then
11950         Set wks = .Worksheets(1)
11960         With wks

11970           strRng_Header2 = vbNullString: strFirstRow2 = vbNullString: strLastRow2 = vbNullString
11980           blnGrouped = False: blnChkDates = False: lngRowHeight2 = 0&

11990           strLastCell = .Cells.SpecialCells(xlCellTypeLastCell).Address  '$H$22205
12000           strLastCell = Rem_Dollar(strLastCell)  ' ** Module Function: modStringFuncs.
12010           strRng_All = "A1:" & strLastCell
12020           Set rng = .Range(strRng_All)
12030           lngCols = rng.Columns.Count
12040           Set rng = Nothing

12050           strLastCol = Left(strLastCell, 1)  ' ** Assumes single-letter address (26 or fewer columns).
12060           strLastRow = Mid(strLastCell, 2)
12070           strRng_Header = "A1:" & strLastCol & "1"
12080           strRng_Title = "A2:B3"
12090           lngRowHeight = .StandardHeight

12100           strRng_RptTitle = "B2:B2"
12110           Set rng = .Range(strRng_RptTitle)
12120           strRptTitle = rng.Value
12130           strRng_RptTitleRight = "C2:" & strLastCol & "2"  ' ** To the right of title cell.

12140           If InStr(strRptTitle, "Grouped") > 0 Then
12150             blnGrouped = True
12160           ElseIf InStr(strPathFile, "Grouped") > 0 Then
12170             blnGrouped = True
12180           End If

12190           If Left(strRptTitle, Len("Summary of Account")) = "Summary of Account" Then
                  ' ** CA: 'Summary of Account - First And Final'
                  ' ** CA: 'Summary of Account - Grouped - First And Final'
                  ' ** FL: 'Summary of Account - First And Final'
                  ' ** FL: 'Summary of Account - Grouped - First And Final'
12200             If Len(strRptTitle) > Len("Summary of Account") Then
12210               strRptTitle = Left(strRptTitle, Len("Summary of Account"))
12220             End If
12230           ElseIf Left(strRptTitle, Len("Summary Statement")) = "Summary Statement" Then
                  ' ** NY 'Summary Statement - First And Final Account'
12240             If Len(strRptTitle) > Len("Summary Statement") Then
12250               strRptTitle = "Summary Statement"
12260             End If
12270           End If

12280           strSheetName = strRptTitle
                'Debug.Print "'  '" & strRptTitle & "'  '" & strSheetName & "'"

12290           strRptStateTitle = strCourtState & ": " & strRptTitle

12300           Select Case strRptStateTitle
                Case "CA: Summary of Account"
                  ' ***********************
                  ' ** California:
                  ' ***********************
                  ' ** Rpt: Summary of Account  (18)
                  ' ** Rpt: Summary of Account  (18) - Grouped
                  ' ** =======================================
                  ' ** All Cols:
                  ' **   A - H
                  ' ** Text Cols:
                  ' **   A, B, C, D, E, F
12310             strRng_Text1 = "A4:D" & strLastRow
12320             strRng_Text2 = "E4:E" & strLastRow
12330             strRng_Text3 = "F4:F" & strLastRow
12340             lngTrimTxt = 0&
12350             lngTxtCnt = 3&
                  ' ** Value Cols:
                  ' **   G, H
12360             strRng_Values1 = "G4:G" & strLastRow
12370             strRng_Values2 = "H4:H" & strLastRow
12380             lngValCnt = 2&
                  ' ** Date Cols:
                  ' ** {none}
12390             lngDatCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
12400             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' ** {none}
12410             lngDecCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
12420             lngPctCnt = 0&
                  ' ** Column Headers.
12430             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
12440             blnColWidths = True
                  ' ** Report Title and Period.
12450             blnRptTitlePeriod = True
12460           Case "FL: Summary of Account"
                  ' ***********************
                  ' ** Florida:
                  ' ***********************
                  ' ** Rpt: Summary of Account  (18) - Personal Representative
                  ' ** Rpt: Summary of Account  (18) - Guardian of Property
                  ' ** Rpt: Summary of Account  (18) - Grouped - Personal Representative
                  ' ** Rpt: Summary of Account  (18) - Grouped - Guardian of Property
                  ' ** =================================================================
12470             Select Case blnGrouped
                  Case True
                    ' ** All Cols:
                    ' **   A - K
                    ' ** Text Cols:
                    ' **   A, B, C, D, E
12480               strRng_Text1 = "A4:B" & strLastRow
12490               strRng_Text2 = "C4:C" & strLastRow
12500               strRng_Text3 = "D4:D" & strLastRow
12510               strRng_Text4 = "E4:E" & strLastRow
12520               lngTrimTxt = 0&
12530               lngTxtCnt = 4&
                    ' ** Value Cols:
                    ' **   F, G, H, I, J, K
12540               strRng_Values1 = "F4:F" & strLastRow
12550               strRng_Values2 = "G4:G" & strLastRow
12560               strRng_Values3 = "H4:H" & strLastRow
12570               strRng_Values4 = "I4:I" & strLastRow
12580               strRng_Values5 = "J4:J" & strLastRow
12590               strRng_Values6 = "K4:K" & strLastRow
12600               lngValCnt = 6&
12610             Case False
                    ' ** All Cols:
                    ' **   A - G
                    ' ** Text Cols:
                    ' **   A, B, C, D
12620               strRng_Text1 = "A4:B" & strLastRow
12630               strRng_Text2 = "C4:C" & strLastRow
12640               strRng_Text3 = "D4:D" & strLastRow
12650               lngTrimTxt = 0&
12660               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   E, F, G
12670               strRng_Values1 = "E4:E" & strLastRow
12680               strRng_Values2 = "F4:F" & strLastRow
12690               strRng_Values3 = "G4:G" & strLastRow
12700               lngValCnt = 3&
12710             End Select
                  ' ** Date Cols:
                  ' ** {none}
12720             lngDatCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
12730             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' ** {none}
12740             lngDecCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
12750             lngPctCnt = 0&
                  ' ** Column Headers.
12760             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
12770             blnColWidths = True
                  ' ** Report Title and Period.
12780             blnRptTitlePeriod = True
12790           Case "NS: Summary of Account"
                  ' ***********************
                  ' ** National Standard:
                  ' ***********************
                  ' ** Rpt: Summary of Account  (18)
                  ' ** Rpt: Summary of Account  (18) - Grouped
                  ' ** =======================================
12800             Select Case blnGrouped
                  Case True
                    ' ** All Cols:
                    ' **   A - H
                    ' ** Text Cols:
                    ' **   A, B, C, D, E, F
12810               strRng_Text1 = "A4:B" & strLastRow
12820               strRng_Text2 = "C4:D" & strLastRow
12830               strRng_Text3 = "E4:E" & strLastRow
12840               strRng_Text4 = "F4:F" & strLastRow
12850               lngTrimTxt = 0&
12860               lngTxtCnt = 4&
                    ' ** Value Cols:
                    ' **   G, H
12870               strRng_Values1 = "G4:G" & strLastRow
12880               strRng_Values2 = "H4:H" & strLastRow
12890               lngValCnt = 2&
12900             Case False
                    ' ** All Cols:
                    ' **   A - F
                    ' ** Text Cols:
                    ' **   A, B, C, D, E
12910               strRng_Text1 = "A4:B" & strLastRow
12920               strRng_Text2 = "C4:D" & strLastRow
12930               strRng_Text3 = "E4:E" & strLastRow
12940               lngTrimTxt = 0&
12950               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   F
12960               strRng_Values1 = "F4:F" & strLastRow
12970               lngValCnt = 1&
12980             End Select
                  ' ** Date Cols:
                  ' ** {none}
12990             lngDatCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
13000             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' ** {none}
13010             lngDecCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
13020             lngPctCnt = 0&
                  ' ** Column Headers.
13030             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
13040             blnColWidths = True
                  ' ** Report Title and Period.
13050             blnRptTitlePeriod = True
13060           Case "NY: Summary Statement"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Summary Statement  (17)
                  ' ** Rpt: Summary Statement  (17) - Grouped
                  ' ** =======================================
13070             Select Case blnGrouped
                  Case True
                    ' ** All Cols:
                    ' **   A - J
                    ' ** Text Cols:
                    ' **   A, B, C, D, E, F, G, I
13080               strRng_Text1 = "A4:C" & strLastRow
13090               strRng_Text2 = "D4:E" & strLastRow
13100               strRng_Text3 = "F4:F" & strLastRow
13110               strRng_Text4 = "G4:G" & strLastRow
13120               strRng_Text5 = "I4:I" & strLastRow
13130               lngTrimTxt = 0&
13140               lngTxtCnt = 5&
                    ' ** Value Cols:
                    ' **   H, J
13150               strRng_Values1 = "H4:H" & strLastRow
13160               strRng_Values2 = "J4:J" & strLastRow
13170               lngValCnt = 2&
13180             Case False
                    ' ** All Cols:
                    ' **   A - H
                    ' ** Text Cols:
                    ' **   A, B, C, D, E, F, G
13190               strRng_Text1 = "A4:C" & strLastRow
13200               strRng_Text2 = "D4:E" & strLastRow
13210               strRng_Text3 = "F4:F" & strLastRow
13220               strRng_Text4 = "G4:G" & strLastRow
13230               lngTrimTxt = 0&
13240               lngTxtCnt = 4&
                    ' ** Value Cols:
                    ' **   H
13250               strRng_Values1 = "H4:H" & strLastRow
13260               lngValCnt = 1&
13270             End Select
                  ' ** Date Cols:
                  ' ** {none}
13280             lngDatCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
13290             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' ** {none}
13300             lngDecCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
13310             lngPctCnt = 0&
                  ' ** Column Headers.
13320             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
13330             blnColWidths = True
                  ' ** Report Title and Period.
13340             blnRptTitlePeriod = True
13350           Case "CA: Property on Hand at Beginning of Accounting Period", "CA: Property on Hand at Close of Accounting Period - Schedule E", _
                    "FL: Assets on Hand at Beginning of Accounting Period", "FL: Assets on Hand at Close of Accounting Period - Schedule E", _
                    "FL: Assets on Hand at Close of Accounting Period - Schedule D", "NS: Property on Hand at Beginning of Accounting Period", _
                    "NS: Property on Hand at Close of Accounting Period"
                  ' ***********************
                  ' ** California:
                  ' ***********************
                  ' ** Rpt: Property on Hand at Beginning of Accounting Period  (50)
                  ' ** Rpt: Property on Hand at Close of Accounting Period - Schedule E  (59)
                  ' ** ======================================================================
                  ' ***********************
                  ' ** Florida:
                  ' ***********************
                  ' ** Rpt: Assets on Hand at Beginning of Accounting Period  (48) - Personal Representative
                  ' ** Rpt: Assets on Hand at Beginning of Accounting Period  (48) - Guardian of Property
                  ' ** Rpt: Assets on Hand at Close of Accounting Period - Schedule E  (57) - Personal Representative
                  ' ** Rpt: Assets on Hand at Close of Accounting Period - Schedule D  (57) - Guardian of Property
                  ' ** ==============================================================================================
                  ' ***********************
                  ' ** National Standard:
                  ' ***********************
                  ' ** Rpt: Property on Hand at Beginning of Accounting Period  (50)
                  ' ** Rpt: Property on Hand at Close of Accounting Period  (46)
                  ' ** =============================================================
13360             If Len(strRptTitle) > 31 Then
13370               If InStr(strSheetName, "Beginning") > 0 Then
13380                 Select Case strCourtState
                      Case "CA", "NS"
13390                   strSheetName = "Prop on Hand at Beginning"  '(25)
13400                 Case "FL"
13410                   strSheetName = "Assets on Hand at Beginning"  '(27)
13420                 End Select
13430               ElseIf InStr(strSheetName, "Close") > 0 Then
13440                 Select Case strCourtState
                      Case "CA"
13450                   strSheetName = "Prop on Hand at Close - Sch E"  '(29)
13460                 Case "NS"
13470                   strSheetName = "Prop on Hand at Close"  '(21)
13480                 Case "FL"
13490                   If InStr(strPathFile, "_Rep_") > 0 Then
13500                     strSheetName = "Assets on Hand at Close - Sch E"  '(31)  IT WILL TAKE 31!
13510                   ElseIf InStr(strPathFile, "_Grdn_") > 0 Then
13520                     strSheetName = "Assets on Hand at Close - Sch D"  '(31)
13530                   End If
13540                 End Select
13550               End If
13560             End If
                  ' ** All Cols:
                  ' **   A - H
                  ' ** Text Cols:
                  ' **   A, B, C, D, F
13570             strRng_Text1 = "A4:D" & strLastRow
13580             strRng_Text2 = "F4:F" & strLastRow
13590             lngTrimTxt = 0&
13600             lngTxtCnt = 2&
                  ' ** Value Cols:
                  ' **   G, H
13610             strRng_Values1 = "G4:G" & strLastRow
13620             strRng_Values2 = "H4:H" & strLastRow
13630             lngValCnt = 2&
                  ' ** Date Cols:
                  ' ** {none}
13640             lngDatCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
13650             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' **   E
13660             strRng_Decimals1 = "E4:E" & strLastRow
13670             lngDecCnt = 1&
                  ' ** Percent Cols:
                  ' ** {none}
13680             lngPctCnt = 0&
                  ' ** Column Headers.
13690             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
13700             blnColWidths = True
                  ' ** Report Title and Period.
13710             blnRptTitlePeriod = True
13720           Case "CA: Additional Property Received", "CA: Information for Investments Made", "CA: Change in Investment Holdings", _
                    "CA: Other Charges", "CA: Other Credits", _
                    "NS: Information for Investments Made", "NS: Change in Investment Holdings"
                  ' ***********************
                  ' ** California:
                  ' ***********************
                  ' ** Rpt: Additional Property Received  (28)
                  ' ** Rpt: Information for Investments Made  (32)
                  ' ** Rpt: Change in Investment Holdings  (29)
                  ' ** Rpt: Other Charges  (13)
                  ' ** Rpt: Other Credits  (13)
                  ' ** ===========================================
                  ' ***********************
                  ' ** National Standard:
                  ' ***********************
                  ' ** Rpt: Information for Investments Made  (32)
                  ' ** Rpt: Change in Investment Holdings  (29)
                  ' ** ===========================================
13730             If Len(strRptTitle) > 31 Then
13740               Select Case strSheetName
                    Case "Information For Investments Made"
13750                 strSheetName = "Info For Investments Made"  '(25)
13760               End Select
13770             End If
                  ' ** All Cols:
                  ' **   A - G
                  ' ** Text Cols:
                  ' **   A, B, C, F
13780             strRng_Text1 = "A4:C" & strLastRow
13790             strRng_Text2 = "F4:F" & strLastRow
13800             lngTrimTxt = 2&
13810             lngTxtCnt = 2&
                  ' ** Value Cols:
                  ' **   G
13820             strRng_Values1 = "G4:G" & strLastRow
13830             lngValCnt = 1&
                  ' ** Date Cols:
                  ' **   D
13840             strRng_Dates1 = "D4:D" & strLastRow
13850             lngDatCnt = 1&
                  ' ** Integer Cols:
                  ' ** {none}
13860             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' **   E
13870             strRng_Decimals1 = "E4:E" & strLastRow
13880             lngDecCnt = 1&
                  ' ** Percent Cols:
                  ' ** {none}
13890             lngPctCnt = 0&
                  ' ** Column Headers.
13900             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
13910             blnColWidths = True
                  ' ** Report Title and Period.
13920             blnRptTitlePeriod = True
13930           Case "CA: Receipts - Schedule A", "FL: Receipts - Schedule A", _
                    "CA: Receipts - Grouped - Schedule A", "FL: Receipts - Grouped - Schedule A", _
                    "NS: Receipts of Principal", "NS: Receipts of Principal - Grouped", _
                    "NS: Receipts of Income", "NS: Receipts of Income - Grouped"
                  ' ***********************
                  ' ** California:
                  ' ***********************
                  ' ** Rpt: Receipts - Schedule A  (21)
                  ' ** Rpt: Receipts - Grouped - Schedule A  (31)
                  ' ** ==========================================
                  ' ***********************
                  ' ** Florida:
                  ' ***********************
                  ' ** Rpt: Receipts - Schedule A  (21) - Personal Representative
                  ' ** Rpt: Receipts - Schedule A  (21) - Guardian of Property
                  ' ** Rpt: Receipts - Grouped - Schedule A  (31) - Personal Representative
                  ' ** Rpt: Receipts - Grouped - Schedule A  (31) - Guardian of Property
                  ' ** ====================================================================
                  ' ***********************
                  ' ** National Standard:
                  ' ***********************
                  ' ** Rpt: Receipts of Principal  (21)
                  ' ** Rpt: Receipts of Principal - Grouped  (31)
                  ' ** Rpt: Receipts of Income  (18)
                  ' ** Rpt: Receipts of Income - Grouped  (28)
                  ' ** ==========================================
13940             Select Case strCourtState
                  Case "CA"
                    ' ** All Cols:
                    ' **   A - H
                    ' ** Text Cols:
                    ' **   A, B, C, D, G
13950               strRng_Text1 = "A4:C" & strLastRow
13960               strRng_Text2 = "D4:D" & strLastRow
13970               strRng_Text3 = "G4:G" & strLastRow
13980               lngTrimTxt = 3&
13990               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   H
14000               strRng_Values1 = "H4:H" & strLastRow
14010               lngValCnt = 1&
                    ' ** Date Cols:
                    ' **   E
14020               strRng_Dates1 = "E4:E" & strLastRow
14030               lngDatCnt = 1&
                    ' ** Decimal Cols:
                    ' **   F
14040               strRng_Decimals1 = "F4:F" & strLastRow
14050               lngDecCnt = 1&
14060             Case "FL"
                    ' ** All Cols:
                    ' **   A - I
                    ' ** Text Cols:
                    ' **   A, B, C, D, G
14070               strRng_Text1 = "A4:C" & strLastRow
14080               strRng_Text2 = "D4:D" & strLastRow
14090               strRng_Text3 = "G4:G" & strLastRow
14100               lngTrimTxt = 3&
14110               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   H, I
14120               strRng_Values1 = "H4:H" & strLastRow
14130               strRng_Values2 = "I4:I" & strLastRow
14140               lngValCnt = 2&
                    ' ** Date Cols:
                    ' **   E
14150               strRng_Dates1 = "E4:E" & strLastRow
14160               lngDatCnt = 1&
                    ' ** Decimal Cols:
                    ' **   F
14170               strRng_Decimals1 = "F4:F" & strLastRow
14180               lngDecCnt = 1&
14190             Case "NS"
14200               If InStr(strPathFile, "_Principal_") > 0 Then
14210                 Select Case blnGrouped
                      Case True
                        ' ** All Cols:
                        ' **   A - I
                        ' ** Text Cols:
                        ' **   A, B, C, D, G
14220                   strRng_Text1 = "A4:C" & strLastRow
14230                   strRng_Text2 = "D4:D" & strLastRow
14240                   strRng_Text3 = "G4:G" & strLastRow
14250                   lngTrimTxt = 3&
14260                   lngTxtCnt = 3&
                        ' ** Value Cols:
                        ' **   H, I
14270                   strRng_Values1 = "H4:H" & strLastRow
14280                   strRng_Values2 = "I4:I" & strLastRow
14290                   lngValCnt = 2&
                        ' ** Date Cols:
                        ' **   E
14300                   strRng_Dates1 = "E4:E" & strLastRow
14310                   lngDatCnt = 1&
                        ' ** Decimal Cols:
                        ' **   F
14320                   strRng_Decimals1 = "F4:F" & strLastRow
14330                   lngDecCnt = 1&
14340                 Case False
                        ' ** All Cols:
                        ' **   A - G
                        ' ** Text Cols:
                        ' **   A, B, C, F
14350                   strRng_Text1 = "A4:C" & strLastRow
14360                   strRng_Text2 = "F4:F" & strLastRow
14370                   lngTrimTxt = 2&
14380                   lngTxtCnt = 2&
                        ' ** Value Cols:
                        ' **   G
14390                   strRng_Values1 = "G4:G" & strLastRow
14400                   lngValCnt = 1&
                        ' ** Date Cols:
                        ' **   D
14410                   strRng_Dates1 = "D4:D" & strLastRow
14420                   lngDatCnt = 1&
                        ' ** Decimal Cols:
                        ' **   E
14430                   strRng_Decimals1 = "E4:E" & strLastRow
14440                   lngDecCnt = 1&
14450                 End Select
14460               ElseIf InStr(strPathFile, "_Income_") > 0 Then
14470                 Select Case blnGrouped
                      Case True
                        ' ** All Cols:
                        ' **   A - J
                        ' ** Text Cols:
                        ' **   A, B, C, D, E, H
14480                   strRng_Text1 = "A4:C" & strLastRow
14490                   strRng_Text2 = "D4:D" & strLastRow
14500                   strRng_Text3 = "E4:E" & strLastRow
14510                   strRng_Text4 = "H4:H" & strLastRow
14520                   lngTrimTxt = 4&
14530                   lngTxtCnt = 4&
                        ' ** Value Cols:
                        ' **   I, J
14540                   strRng_Values1 = "I4:I" & strLastRow
14550                   strRng_Values2 = "J4:J" & strLastRow
14560                   lngValCnt = 2&
                        ' ** Date Cols:
                        ' **   F
14570                   strRng_Dates1 = "F4:F" & strLastRow
14580                   lngDatCnt = 1&
                        ' ** Decimal Cols:
                        ' **   G
14590                   strRng_Decimals1 = "G4:G" & strLastRow
14600                   lngDecCnt = 1&
14610                 Case False
                        ' ** All Cols:
                        ' **   A - I
                        ' ** Text Cols:
                        ' **   A, B, C, D, G
14620                   strRng_Text1 = "A4:C" & strLastRow
14630                   strRng_Text2 = "D4:D" & strLastRow
14640                   strRng_Text3 = "G4:G" & strLastRow
14650                   lngTrimTxt = 3&
14660                   lngTxtCnt = 3&
                        ' ** Value Cols:
                        ' **   H, I
14670                   strRng_Values1 = "H4:H" & strLastRow
14680                   strRng_Values2 = "I4:I" & strLastRow
14690                   lngValCnt = 2&
                        ' ** Date Cols:
                        ' **   E
14700                   strRng_Dates1 = "E4:E" & strLastRow
14710                   lngDatCnt = 1&
                        ' ** Decimal Cols:
                        ' **   F
14720                   strRng_Decimals1 = "F4:F" & strLastRow
14730                   lngDecCnt = 1&
14740                 End Select
14750               End If
14760             End Select
                  ' ** Integer Cols:
                  ' ** {none}
14770             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
14780             lngPctCnt = 0&
                  ' ** Column Headers.
14790             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
14800             blnColWidths = True
                  ' ** Report Title and Period.
14810             blnRptTitlePeriod = True
14820           Case "CA: Gains on Sale or Other Dispositions - Schedule B", "CA: Losses on Sale or Other Dispositions - Schedule D", _
                    "NS: Gains (Losses) on Sale or Other Dispositions"
                  ' ***********************
                  ' ** California:
                  ' ***********************
                  ' ** Rpt: Gains on Sale or Other Dispositions - Schedule B  (48)
                  ' ** Rpt: Losses on Sale or Other Dispositions - Schedule D  (49)
                  ' ** ============================================================
                  ' ***********************
                  ' ** National Standard:
                  ' ***********************
                  ' ** Rpt: Gains (Losses) on Sale or Other Dispositions  (44)
                  ' ** =======================================================
14830             If Len(strRptTitle) > 31 Then
14840               Select Case strSheetName
                    Case "Gains (Losses) on Sale or Other Dispositions"
14850                 strSheetName = "Gains (Losses) on Sale"  '(22)
14860               Case "Gains on Sale or Other Dispositions - Schedule B"
14870                 strSheetName = "Gains on Sale - Schedule B"  '(26)
14880               Case "Losses on Sale or Other Dispositions - Schedule D"
14890                 strSheetName = "Losses on Sale - Schedule D"  '(27)
14900               End Select
14910             End If
                  ' ** All Cols:
                  ' **   A - I
                  ' ** Text Cols:
                  ' **   A, B, C, F
14920             strRng_Text1 = "A4:C" & strLastRow
14930             strRng_Text2 = "F4:F" & strLastRow
14940             lngTrimTxt = 0&
14950             lngTxtCnt = 2&
                  ' ** Value Cols:
                  ' **   G, H, I
14960             strRng_Values1 = "G4:G" & strLastRow
14970             strRng_Values2 = "H4:H" & strLastRow
14980             strRng_Values3 = "I4:I" & strLastRow
14990             lngValCnt = 3&
                  ' ** Date Cols:
                  ' **   D
15000             strRng_Dates1 = "D4:D" & strLastRow
15010             lngDatCnt = 1&
                  ' ** Integer Cols:
                  ' ** {none}
15020             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' ** {none}
15030             lngDecCnt = 0&
                  ' ** Percent Cols:
                  ' **   E
15040             strRng_Decimals1 = "E4:E" & strLastRow
15050             lngDecCnt = 1&
                  ' ** Column Headers.
15060             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
15070             blnColWidths = True
                  ' ** Report Title and Period.
15080             blnRptTitlePeriod = True
15090           Case "CA: Disbursements - Schedule C", "FL: Disbursements - Schedule B", _
                    "CA: Disbursements - Grouped - Schedule C", "FL: Disbursements - Grouped - Schedule B", _
                    "NS: Disbursements of Principal", "NS: Disbursements of Income", _
                    "NS: Disbursements of Principal - Grouped", "NS: Disbursements of Income - Grouped"
                  ' ***********************
                  ' ** California:
                  ' ***********************
                  ' ** Rpt: Disbursements - Schedule C  (26)
                  ' ** Rpt: Disbursements - Grouped - Schedule C  (36)
                  ' ** ===============================================
                  ' ***********************
                  ' ** Florida:
                  ' ***********************
                  ' ** Rpt: Disbursements - Schedule B  (26) - Personal Representative
                  ' ** Rpt: Disbursements - Grouped - Schedule B  (36) - Personal Representative
                  ' ** =========================================================================
                  ' ***********************
                  ' ** National Standard:
                  ' ***********************
                  ' ** Rpt: Disbursements of Principal  (26)
                  ' ** Rpt: Disbursements of Principal - Grouped  (36)
                  ' ** Rpt: Disbursements of Income  (23)
                  ' ** Rpt: Disbursements of Income - Grouped  (33)
                  ' ** ===============================================
15100             If Len(strRptTitle) > 31 Then
15110               Select Case strSheetName
                    Case "Disbursements of Principal - Grouped"
15120                 strSheetName = "Disburse of Principal - Grouped"  '(31)
15130               Case "Disbursements of Income - Grouped"
15140                 strSheetName = "Disburse of Income - Grouped"  '(28)
15150               Case "Disbursements - Grouped - Schedule C"
15160                 strSheetName = "Disbursements - Grouped - Sch C"  '(31)
15170               Case "Disbursements - Grouped - Schedule B"
15180                 strSheetName = "Disbursements - Grouped - Sch B"  '(31)
15190               End Select
15200             End If
15210             Select Case blnGrouped
                  Case True
15220               Select Case strCourtState
                    Case "CA"
                      ' ** All Cols:
                      ' **   A - G
                      ' ** Text Cols:
                      ' **   A, B, C, D, F
15230                 strRng_Text1 = "A4:B" & strLastRow
15240                 strRng_Text2 = "C4:D" & strLastRow
15250                 strRng_Text3 = "F4:F" & strLastRow
15260                 lngTrimTxt = 3&
15270                 lngTxtCnt = 3&
                      ' ** Value Cols:
                      ' **   G
15280                 strRng_Values1 = "G4:G" & strLastRow
15290                 lngValCnt = 1&
                      ' ** Date Cols:
                      ' **   E
15300                 strRng_Dates1 = "E4:E" & strLastRow
15310                 lngDatCnt = 1&
                      ' ** Decimal Cols:
                      ' ** {none}
15320                 lngDecCnt = 0&
15330               Case "FL"
                      ' ** All Cols:
                      ' **   A - H
                      ' ** Text Cols:
                      ' **   A, B, C, F
15340                 strRng_Text1 = "A4:B" & strLastRow
15350                 strRng_Text2 = "C4:C" & strLastRow
15360                 strRng_Text3 = "F4:F" & strLastRow
15370                 lngTrimTxt = 3&
15380                 lngTxtCnt = 3&
                      ' ** Value Cols:
                      ' **   G, H
15390                 strRng_Values1 = "G4:G" & strLastRow
15400                 strRng_Values2 = "H4:H" & strLastRow
15410                 lngValCnt = 2&
                      ' ** Date Cols:
                      ' **   D
15420                 strRng_Dates1 = "D4:D" & strLastRow
15430                 lngDatCnt = 1&
                      ' ** Decimal Cols:
                      ' **   E
15440                 strRng_Decimals1 = "E4:E" & strLastRow
15450                 lngDecCnt = 1&
15460               Case "NS"
                      ' ** All Cols:
                      ' **   A - G
                      ' ** Text Cols:
                      ' **   A, B, C, D, F
15470                 strRng_Text1 = "A4:B" & strLastRow
15480                 strRng_Text2 = "C4:C" & strLastRow
15490                 strRng_Text3 = "D4:D" & strLastRow
15500                 strRng_Text4 = "F4:F" & strLastRow
15510                 lngTrimTxt = 4&
15520                 lngTxtCnt = 4&
                      ' ** Value Cols:
                      ' **   G
15530                 strRng_Values1 = "G4:G" & strLastRow
15540                 lngValCnt = 1&
                      ' ** Date Cols:
                      ' **   E
15550                 strRng_Dates1 = "E4:E" & strLastRow
15560                 lngDatCnt = 1&
                      ' ** Decimal Cols:
                      ' ** {none}
15570                 lngDecCnt = 0&
15580               End Select
15590             Case False
15600               Select Case strCourtState
                    Case "CA"
                      ' ** All Cols:
                      ' **   A - F
                      ' ** Text Cols:
                      ' **   A, B, C, E
15610                 strRng_Text1 = "A4:B" & strLastRow
15620                 strRng_Text2 = "C4:C" & strLastRow
15630                 strRng_Text3 = "E4:E" & strLastRow
15640                 lngTrimTxt = 3&
15650                 lngTxtCnt = 3&
                      ' ** Value Cols:
                      ' **   F
15660                 strRng_Values1 = "F4:F" & strLastRow
15670                 lngValCnt = 1&
                      ' ** Date Cols:
                      ' **   D
15680                 strRng_Dates1 = "D4:D" & strLastRow
15690                 lngDatCnt = 1&
                      ' ** Decimal Cols:
                      ' ** {none}
15700                 lngDecCnt = 0&
15710               Case "FL"
                      ' ** All Cols:
                      ' **   A - G
                      ' ** Text Cols:
                      ' **   A, B, E
15720                 strRng_Text1 = "A4:B" & strLastRow
15730                 strRng_Text2 = "E4:E" & strLastRow
15740                 lngTrimTxt = 2&
15750                 lngTxtCnt = 2&
                      ' ** Value Cols:
                      ' **   F, G
15760                 strRng_Values1 = "F4:F" & strLastRow
15770                 strRng_Values2 = "G4:G" & strLastRow
15780                 lngValCnt = 2&
                      ' ** Date Cols:
                      ' **   C
15790                 strRng_Dates1 = "C4:C" & strLastRow
15800                 lngDatCnt = 1&
                      ' ** Decimal Cols:
                      ' **   D
15810                 strRng_Decimals1 = "D4:D" & strLastRow
15820                 lngDecCnt = 1&
15830               Case "NS"
                      ' ** All Cols:
                      ' **   A - F
                      ' ** Text Cols:
                      ' **   A, B, C, E
15840                 strRng_Text1 = "A4:B" & strLastRow
15850                 strRng_Text2 = "C4:C" & strLastRow
15860                 strRng_Text3 = "E4:E" & strLastRow
15870                 lngTrimTxt = 3&
15880                 lngTxtCnt = 3&
                      ' ** Value Cols:
                      ' **   F
15890                 strRng_Values1 = "F4:F" & strLastRow
15900                 lngValCnt = 1&
                      ' ** Date Cols:
                      ' **   D
15910                 strRng_Dates1 = "D4:D" & strLastRow
15920                 lngDatCnt = 1&
                      ' ** Decimal Cols:
                      ' ** {none}
15930                 lngDecCnt = 0&
15940               End Select
15950             End Select
                  ' ** Integer Cols:
                  ' ** {none}
15960             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
15970             lngPctCnt = 0&
                  ' ** Column Headers.
15980             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
15990             blnColWidths = True
                  ' ** Report Title and Period.
16000             blnRptTitlePeriod = True
16010           Case "CA: Distributions", "FL: Distributions - Schedule C", _
                    "NS: Distributions of Principal to Beneficiaries", "NS: Distributions of Income"
                  ' ***********************
                  ' ** California:
                  ' ***********************
                  ' ** Rpt: Distributions  (13)
                  ' ** ========================
                  ' ***********************
                  ' ** Florida:
                  ' ***********************
                  ' ** Rpt: Distributions - Schedule C  (26) - Personal Representative
                  ' ** ===============================================================
                  ' ***********************
                  ' ** National Standard:
                  ' ***********************
                  ' ** Rpt: Distributions of Principal to Beneficiaries  (43)
                  ' ** Rpt: Distributions of Income  (24)
                  ' ** ======================================================
16020             If Len(strRptTitle) > 31 Then
16030               Select Case strSheetName
                    Case "Distributions of Principal to Beneficiaries"
16040                 strSheetName = "Distributions of Principal"  '(26)
16050               End Select
16060             End If
16070             Select Case strCourtState
                  Case "CA"
                    ' ** All Cols:
                    ' **   A - F
                    ' ** Text Cols:
                    ' **   A, B, C, E
16080               strRng_Text1 = "A4:B" & strLastRow
16090               strRng_Text2 = "C4:C" & strLastRow
16100               strRng_Text3 = "E4:E" & strLastRow
16110               lngTrimTxt = 3&
16120               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   F
16130               strRng_Values1 = "F4:F" & strLastRow
16140               lngValCnt = 1&
                    ' ** Date Cols:
                    ' **   D
16150               strRng_Dates1 = "D4:D" & strLastRow
16160               lngDatCnt = 1&
16170             Case "FL"
                    ' ** All Cols:
                    ' **   A - G
                    ' ** Text Cols:
                    ' **   A, B, C, E
16180               strRng_Text1 = "A4:B" & strLastRow
16190               strRng_Text2 = "C4:C" & strLastRow
16200               strRng_Text3 = "E4:E" & strLastRow
16210               lngTrimTxt = 3&
16220               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   F, G
16230               strRng_Values1 = "F4:F" & strLastRow
16240               strRng_Values2 = "G4:G" & strLastRow
16250               lngValCnt = 2&
                    ' ** Date Cols:
                    ' **   D
16260               strRng_Dates1 = "D4:D" & strLastRow
16270               lngDatCnt = 1&
16280             Case "NS"
                    ' ** All Cols:
                    ' **   A - F
                    ' ** Text Cols:
                    ' **   A, B, C, E
16290               strRng_Text1 = "A4:B" & strLastRow
16300               strRng_Text2 = "C4:C" & strLastRow
16310               strRng_Text3 = "E4:E" & strLastRow
16320               lngTrimTxt = 3&
16330               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   F
16340               strRng_Values1 = "F4:F" & strLastRow
16350               lngValCnt = 1&
                    ' ** Date Cols:
                    ' **   D
16360               strRng_Dates1 = "D4:D" & strLastRow
16370               lngDatCnt = 1&
16380             End Select
                  ' ** Integer Cols:
                  ' ** {none}
16390             lngIntCnt = 0&
                  ' ** Decimal Cols:
                  ' ** {none}
16400             lngDecCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
16410             lngPctCnt = 0&
                  ' ** Column Headers.
16420             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
16430             blnColWidths = True
                  ' ** Report Title and Period.
16440             blnRptTitlePeriod = True
16450           Case "FL: Disbursements and Distributions - Schedule B", "FL: Disbursements and Distributions - Grouped - Schedule B"
                  ' ***********************
                  ' ** Florida:
                  ' ***********************
                  ' ** Rpt: Disbursements and Distributions - Schedule B  (44) - Guardian of Property
                  ' ** Rpt: Disbursements and Distributions - Grouped - Schedule B  (54) - Guardian of Property
                  ' ** ========================================================================================
16460             If Len(strRptTitle) > 31 Then
16470               strSheetName = "Disburse & Distribute - Sch B"  '(29)
16480             End If
16490             Select Case blnGrouped
                  Case True
                    ' ** All Cols:
                    ' **   A - G
                    ' ** Text Cols:
                    ' **   A, B, C, E
16500               strRng_Text1 = "A4:B" & strLastRow
16510               strRng_Text2 = "C4:C" & strLastRow
16520               strRng_Text3 = "E4:E" & strLastRow
16530               lngTrimTxt = 3&
16540               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   F, G
16550               strRng_Values1 = "F4:F" & strLastRow
16560               strRng_Values2 = "G4:G" & strLastRow
16570               lngValCnt = 2&
                    ' ** Date Cols:
                    ' **   D
16580               strRng_Dates1 = "D4:D" & strLastRow
16590               lngDatCnt = 1&
                    ' ** Decimal Cols:
                    ' ** {none}
16600               lngDecCnt = 0&
16610             Case False
                    ' ** All Cols:
                    ' **   A - G
                    ' ** Text Cols:
                    ' **   A, B, E
16620               strRng_Text1 = "A4:B" & strLastRow
16630               strRng_Text2 = "E4:E" & strLastRow
16640               lngTrimTxt = 2&
16650               lngTxtCnt = 2&
                    ' ** Value Cols:
                    ' **   F, G
16660               strRng_Values1 = "F4:F" & strLastRow
16670               strRng_Values2 = "G4:G" & strLastRow
16680               lngValCnt = 2&
                    ' ** Date Cols:
                    ' **   C
16690               strRng_Dates1 = "C4:C" & strLastRow
16700               lngDatCnt = 1&
                    ' ** Decimal Cols:
                    ' **   D
16710               strRng_Decimals1 = "D4:D" & strLastRow
16720               lngDecCnt = 1&
16730             End Select
                  ' ** Integer Cols:
                  ' ** {none}
16740             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
16750             lngPctCnt = 0&
                  ' ** Column Headers.
16760             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
16770             blnColWidths = True
                  ' ** Report Title and Period.
16780             blnRptTitlePeriod = True
16790           Case "FL: Capital Transactions and Adjustments - Schedule D", "FL: Capital Transactions and Adjustments - Schedule C"
                  ' ***********************
                  ' ** Florida:
                  ' ***********************
                  ' ** Rpt: Capital Transactions and Adjustments - Schedule D  (49) - Personal Representative
                  ' ** Rpt: Capital Transactions and Adjustments - Schedule C  (49) - Guardian of Property
                  ' ** ======================================================================================
16800             If Len(strRptTitle) > 31 Then
16810               If InStr(strPathFile, "_Rep_") > 0 Then
16820                 strSheetName = "Capital Trans & Adjust - Sch D"  '(30)
16830               ElseIf InStr(strPathFile, "_Grdn_") > 0 Then
16840                 strSheetName = "Capital Trans & Adjust - Sch C"  '(30)
16850               End If
16860             End If
                  ' ** All Cols:
                  ' **   A - J
                  ' ** Text Cols:
                  ' **   A, B, C, F
16870             strRng_Text1 = "A4:B" & strLastRow
16880             strRng_Text2 = "C4:C" & strLastRow
16890             strRng_Text3 = "F4:F" & strLastRow
16900             lngTrimTxt = 3&
16910             lngTxtCnt = 3&
                  ' ** Value Cols:
                  ' **   G, H, I, J
16920             strRng_Values1 = "G4:G" & strLastRow
16930             strRng_Values2 = "H4:H" & strLastRow
16940             strRng_Values3 = "I4:I" & strLastRow
16950             strRng_Values4 = "J4:J" & strLastRow
16960             lngValCnt = 4&
                  ' ** Date Cols:
                  ' **   D
16970             strRng_Dates1 = "D4:D" & strLastRow
16980             lngDatCnt = 1&
                  ' ** Decimal Cols:
                  ' **   E
16990             strRng_Decimals1 = "E4:E" & strLastRow
17000             lngDecCnt = 1&
                  ' ** Integer Cols:
                  ' ** {none}
17010             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
17020             lngPctCnt = 0&
                  ' ** Column Headers.
17030             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
17040             blnColWidths = True
                  ' ** Report Title and Period.
17050             blnRptTitlePeriod = True
17060           Case "NY: Statement of Principal Received - Schedule A"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of Principal Received - Schedule A  (44)
                  ' ** =======================================================
17070             If Len(strRptTitle) > 31 Then
17080               strSheetName = "Principal Received - Schedule A"  '(31)
17090             End If
                  ' ** Find the split between the 1st and 2nd reports.
17100             strRng_Header2 = vbNullString: strFirstRow2 = vbNullString: strLastRow2 = vbNullString
17110             strRng_Text4 = "E4:E" & strLastRow
17120             Set rng = .Range(strRng_Text4)
17130             With rng
17140               Set rng2 = .Find("Date", , xlValues, xlWhole, xlByColumns, xlNext, True)
17150               If IsNothing(rng2) = False Then  ' ** Module Function: modUtilities.
17160                 strTmp01 = rng2.Address  ' ** e.g., $E$32.
17170                 strTmp01 = Rem_Dollar(strTmp01)  ' ** Module Function: modStringFuncs.
17180                 strRng_Header2 = strTmp01 & ":" & "G" & Mid(strTmp01, 2)
17190                 strLastRow2 = strLastRow
17200                 strLastRow = CStr(Val(Mid(strTmp01, 2)) - 1)
17210                 strFirstRow2 = CStr(Val(Mid(strTmp01, 2)) + 1)
17220               End If
17230             End With
17240             Set rng2 = Nothing
17250             Set rng = Nothing
17260             strRng_Text4 = vbNullString
                  ' ** All Cols:
                  ' **   A - G
                  ' ** Text Cols:
                  ' **   A, B, C, D, F
                  '####################
                  'HERE IT IS!
                  '####################
17270             strRng_Text1 = "A4:C" & strLastRow2
17280             strRng_Text2 = "D4:D" & strLastRow2
17290             strRng_Text3 = "F4:F" & strLastRow2
17300             lngTrimTxt = 3&
17310             lngTxtCnt = 3&
                  ' ** Value Cols:
                  ' **   G
17320             strRng_Values1 = "G4:G" & strLastRow2
17330             lngValCnt = 1&
                  ' ** Date Cols:
                  ' **   E
17340             strRng_Dates1 = "E" & strFirstRow2 & ":E" & strLastRow2
17350             lngDatCnt = 1&
17360             blnChkDates = True
                  ' ** Decimal Cols:
                  ' **   E
17370             strRng_Decimals1 = "E4:E" & strLastRow
17380             lngDecCnt = 1&
                  ' ** Integer Cols:
                  ' ** {none}
17390             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
17400             lngPctCnt = 0&
                  ' ** Column Headers.
17410             blnColHeads = True
                  ' ** Column Widths.
17420             blnColWidths = True
                  ' ** Report Title and Period.
17430             blnRptTitlePeriod = True
17440           Case "NY: Statement of Increases on Sales, Liquidation or Distribution - Schedule A-1", _
                    "NY: Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability - Schedule B"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of Increases on Sales, Liquidation or Distribution - Schedule A-1  (75)
                  ' ** Rpt: Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability - Schedule B  (107)
                  ' ** =======================================================================================================================
17450             If Len(strRptTitle) > 31 Then
17460               If InStr(strRptTitle, "Increase") > 0 Then
17470                 strSheetName = "Increases on Sales - Sch A-1"  '(28)
17480               ElseIf InStr(strRptTitle, "Decrease") > 0 Then
17490                 strSheetName = "Decreases on Sales - Sch B"  '(26)
17500               End If
17510             End If
17520             strRng_Header2 = vbNullString
                  ' ** All Cols:
                  ' **   A - I
                  ' ** Text Cols:
                  ' **   A, B, C, F
17530             strRng_Text1 = "A4:C" & strLastRow
17540             strRng_Text2 = "F4:F" & strLastRow
17550             lngTrimTxt = 2&
17560             lngTxtCnt = 2&
                  ' ** Value Cols:
                  ' **   G, H, I
17570             strRng_Values1 = "G4:G" & strLastRow
17580             strRng_Values2 = "H4:H" & strLastRow
17590             strRng_Values3 = "I4:I" & strLastRow
17600             lngValCnt = 3&
                  ' ** Date Cols:
                  ' **   D
17610             strRng_Dates1 = "D4:D" & strLastRow
17620             lngDatCnt = 1&
                  ' ** Decimal Cols:
                  ' **   E
17630             strRng_Decimals1 = "E4:E" & strLastRow
17640             lngDecCnt = 1&
                  ' ** Integer Cols:
                  ' ** {none}
17650             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
17660             lngPctCnt = 0&
                  ' ** Column Headers.
17670             blnColHeads = True
17680             lngRowHeight2 = (3& * lngRowHeight)
17690             strRng_Header2 = strRng_Header
                  ' ** Column Widths.
17700             blnColWidths = True
                  ' ** Report Title and Period.
17710             blnRptTitlePeriod = True
17720           Case "NY: Statement of Administration Expenses Chargeable to Principal - Schedule C", _
                    "NY: Statement of Administration Expenses Chargeable to Principal - Grouped - Schedule C", _
                    "NY: Statement of Administration Expenses Chargeable to Income - Schedule C-2", _
                    "NY: Statement of Administration Expenses Chargeable to Income - Grouped - Schedule C-2"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of Administration Expenses Chargeable to Principal - Schedule C  (73)
                  ' ** Rpt: Statement of Administration Expenses Chargeable to Principal - Grouped - Schedule C  (83)
                  ' ** Rpt: Statement of Administration Expenses Chargeable to Income - Schedule C-2  (72)
                  ' ** Rpt: Statement of Administration Expenses Chargeable to Income - Grouped - Schedule C-2  (82)
                  ' ** ==============================================================================================
17730             If Len(strRptTitle) > 31 Then
17740               Select Case blnGrouped
                    Case True
17750                 If InStr(strRptTitle, "Principal") > 0 Then
17760                   strSheetName = "Admin Prin - Grouped - Sch C"  '(28)
17770                 ElseIf InStr(strRptTitle, "Income") > 0 Then
17780                   strSheetName = "Admin Inc - Sch C-2"  '(19)
17790                 End If
17800               Case False
17810                 If InStr(strRptTitle, "Principal") > 0 Then
17820                   strSheetName = "Admin Prin - Sch C"  '(18)
17830                 ElseIf InStr(strRptTitle, "Income") > 0 Then
17840                   strSheetName = "Admin Inc - Grouped - Sch C-2"  '(29)
17850                 End If
17860               End Select
17870             End If
17880             Select Case blnGrouped
                  Case True
                    ' ** All Cols:
                    ' **   A - H
                    ' ** Text Cols:
                    ' **   A, B, C, D, F
17890               strRng_Text1 = "A4:C" & strLastRow
17900               strRng_Text2 = "D4:D" & strLastRow
17910               strRng_Text3 = "F4:F" & strLastRow
17920               lngTrimTxt = 3&
17930               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   G, H
17940               strRng_Values1 = "G4:G" & strLastRow
17950               strRng_Values2 = "H4:H" & strLastRow
17960               lngValCnt = 2&
                    ' ** Date Cols:
                    ' **   E
17970               strRng_Dates1 = "E4:E" & strLastRow
17980               lngDatCnt = 1&
17990             Case False
                    ' ** All Cols:
                    ' **   A - F
                    ' ** Text Cols:
                    ' **   A, B, C, E
18000               strRng_Text1 = "A4:C" & strLastRow
18010               strRng_Text2 = "E4:E" & strLastRow
18020               lngTrimTxt = 2&
18030               lngTxtCnt = 2&
                    ' ** Value Cols:
                    ' **   F
18040               strRng_Values1 = "F4:F" & strLastRow
18050               lngValCnt = 1&
                    ' ** Date Cols:
                    ' **   D
18060               strRng_Dates1 = "D4:D" & strLastRow
18070               lngDatCnt = 1&
18080             End Select
                  ' ** Decimal Cols:
                  ' ** {none}
18090             lngDecCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
18100             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
18110             lngPctCnt = 0&
                  ' ** Column Headers.
18120             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
18130             blnColWidths = True
                  ' ** Report Title and Period.
18140             blnRptTitlePeriod = True
18150           Case "NY: Statement of Distributions of Principal - Schedule D", _
                    "NY: Statement of Distributions of Income - Schedule D-1"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of Distributions of Principal - Schedule D  (52)
                  ' ** Rpt: Statement of Distributions of Income - Schedule D-1  (51)
                  ' ** ===============================================================
18160             If Len(strRptTitle) > 31 Then
18170               If InStr(strRptTitle, "Principal") > 0 Then
18180                 strSheetName = "Distributions of Prin - Sch D"  '(29)
18190               ElseIf InStr(strRptTitle, "Income") > 0 Then
18200                 strSheetName = "Distributions of Inc - Sch D-1"  '(30)
18210               End If
18220             End If
                  ' ** All Cols:
                  ' **   A - F
                  ' ** Text Cols:
                  ' **   A, B, C, E
18230             strRng_Text1 = "A4:C" & strLastRow
18240             strRng_Text2 = "E4:E" & strLastRow
18250             lngTrimTxt = 2&
18260             lngTxtCnt = 2&
                  ' ** Value Cols:
                  ' **   F
18270             strRng_Values1 = "F4:F" & strLastRow
18280             lngValCnt = 1&
                  ' ** Date Cols:
                  ' **   D
18290             strRng_Dates1 = "D4:D" & strLastRow
18300             lngDatCnt = 1&
                  ' ** Decimal Cols:
                  ' ** {none}
18310             lngDecCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
18320             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
18330             lngPctCnt = 0&
                  ' ** Column Headers.
18340             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
18350             blnColWidths = True
                  ' ** Report Title and Period.
18360             blnRptTitlePeriod = True
18370           Case "NY: Statement of New Investments, Exchanges and Stock Distributions of Principal Assets - Schedule E"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of New Investments, Exchanges and Stock Distributions of Principal Assets - Schedule E  (96)
                  ' ** ===========================================================================================================
18380             If Len(strRptTitle) > 31 Then
18390               strSheetName = "New Investments - Schedule E"  '(28)
18400             End If
                  ' ** All Cols:
                  ' **   A - H
                  ' ** Text Cols:
                  ' **   A, B, C, D, G
18410             strRng_Text1 = "A4:C" & strLastRow
18420             strRng_Text2 = "D4:D" & strLastRow
18430             strRng_Text3 = "G4:G" & strLastRow
18440             lngTrimTxt = 3&
18450             lngTxtCnt = 3&
                  ' ** Value Cols:
                  ' **   H
18460             strRng_Values1 = "H4:H" & strLastRow
18470             lngValCnt = 1&
                  ' ** Date Cols:
                  ' **   E
18480             strRng_Dates1 = "E4:E" & strLastRow
18490             lngDatCnt = 1&
                  ' ** Decimal Cols:
                  ' **   F
18500             strRng_Decimals1 = "F4:F" & strLastRow
18510             lngDecCnt = 1&
                  ' ** Integer Cols:
                  ' ** {none}
18520             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
18530             lngPctCnt = 0&
                  ' ** Column Headers.
18540             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
18550             blnColWidths = True
                  ' ** Report Title and Period.
18560             blnRptTitlePeriod = True
18570           Case "NY: Statement of Principal Remaining on Hand - Schedule F", _
                    "NY: Statement of Income Remaining on Hand - Schedule F-1"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of Principal Remaining on Hand - Schedule F  (53)
                  ' ** Rpt: Statement of Income Remaining on Hand - Schedule F-1  (52)
                  ' ** ================================================================
18580             If Len(strRptTitle) > 31 Then
18590               If InStr(strRptTitle, "Principal") > 0 Then
18600                 strSheetName = "Principal on Hand - Schedule F"  '(30)
18610               ElseIf InStr(strRptTitle, "Income") > 0 Then
18620                 strSheetName = "Income on Hand - Schedule F-1"  '(29)
18630               End If
18640             End If
18650             If InStr(strRptTitle, "Principal") > 0 Then
                    ' ** All Cols:
                    ' **   A - J
                    ' ** Text Cols:
                    ' **   A, B, C, D, F
18660               strRng_Text1 = "A4:C" & strLastRow
18670               strRng_Text2 = "D4:D" & strLastRow
18680               strRng_Text3 = "F4:F" & strLastRow
18690               lngTrimTxt = 0&
18700               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   G, H, I, J
18710               strRng_Values1 = "G4:G" & strLastRow
18720               strRng_Values2 = "H4:H" & strLastRow
18730               strRng_Values3 = "I4:I" & strLastRow
18740               strRng_Values4 = "J4:J" & strLastRow
18750               lngValCnt = 4&
                    ' ** Decimal Cols:
                    ' **   E
18760               strRng_Decimals1 = "E4:E" & strLastRow
18770               lngDecCnt = 1&
18780             ElseIf InStr(strRptTitle, "Income") > 0 Then
                    ' ** All Cols:
                    ' **   A - E
                    ' ** Text Cols:
                    ' **   A, B, C, D
18790               strRng_Text1 = "A4:C" & strLastRow
18800               strRng_Text2 = "D4:D" & strLastRow
18810               lngTrimTxt = 0&
18820               lngTxtCnt = 2&
                    ' ** Value Cols:
                    ' **   E
18830               strRng_Values1 = "E4:E" & strLastRow
18840               lngValCnt = 1&
                    ' ** Decimal Cols:
                    ' ** {none}
18850               lngDecCnt = 0&
18860             End If
                  ' ** Date Cols:
                  ' ** {none}
18870             lngDatCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
18880             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
18890             lngPctCnt = 0&
                  ' ** Column Headers.
18900             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
18910             blnColWidths = True
                  ' ** Report Title and Period.
18920             blnRptTitlePeriod = True
18930           Case "NY: Statement of Income Received - Schedule AA-1"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of Income Received - Schedule AA-1  (44)
                  ' ** =======================================================
18940             If Len(strRptTitle) > 31 Then
18950               strSheetName = "Income Received - Schedule AA-1"  '(31)
18960             End If
                  ' ** All Cols:
                  ' **   A - E
                  ' ** Text Cols:
                  ' **   A, B, C, D
18970             strRng_Text1 = "A4:C" & strLastRow
18980             strRng_Text2 = "D4:D" & strLastRow
18990             lngTrimTxt = 0&
19000             lngTxtCnt = 2&
                  ' ** Value Cols:
                  ' **   E
19010             strRng_Values1 = "E4:E" & strLastRow
19020             lngValCnt = 1&
                  ' ** Date Cols:
                  ' ** {none}
19030             lngDatCnt = 0&
                  ' ** Decimal Cols:
                  ' ** {none}
19040             lngDecCnt = 0&
                  ' ** Integer Cols:
                  ' ** {none}
19050             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
19060             lngPctCnt = 0&
                  ' ** Column Headers.
19070             blnColHeads = True
19080             lngRowHeight2 = (2& * lngRowHeight)
19090             strRng_Header2 = strRng_Header
                  ' ** Column Widths.
19100             blnColWidths = True
                  ' ** Report Title and Period.
19110             blnRptTitlePeriod = True
19120           Case "NY: Statement of All Income Collected - Schedule A-2", _
                    "NY: Statement of All Income Collected - Grouped - Schedule A-2"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Statement of All Income Collected - Schedule A-2  (48)
                  ' ** Rpt: Statement of All Income Collected - Grouped - Schedule A-2  (58)
                  ' ** =====================================================================
19130             If Len(strRptTitle) > 31 Then
19140               Select Case blnGrouped
                    Case True
19150                 strSheetName = "Income Coll - Grouped - Sch A-2"  '(31)
19160               Case False
19170                 strSheetName = "Income Collected - Schedule A-2"  '(31)
19180               End Select
19190             End If
19200             Select Case blnGrouped
                  Case True
                    ' ** All Cols:
                    ' **   A - J
                    ' ** Text Cols:
                    ' **   A, B, C, D, E, H
19210               strRng_Text1 = "A4:C" & strLastRow
19220               strRng_Text2 = "D4:D" & strLastRow
19230               strRng_Text3 = "E4:E" & strLastRow
19240               strRng_Text4 = "H4:H" & strLastRow
19250               lngTrimTxt = 0&
19260               lngTxtCnt = 4&
                    ' ** Value Cols:
                    ' **   I, J
19270               strRng_Values1 = "I4:I" & strLastRow
19280               strRng_Values2 = "J4:J" & strLastRow
19290               lngValCnt = 2&
                    ' ** Date Cols:
                    ' **   F
19300               strRng_Dates1 = "F4:F" & strLastRow
19310               lngDatCnt = 1&
                    ' ** Decimal Cols:
                    ' **   G
19320               strRng_Decimals1 = "G4:G" & strLastRow
19330               lngDecCnt = 1&
19340             Case False
                    ' ** All Cols:
                    ' **   A - I
                    ' ** Text Cols:
                    ' **   A, B, C, D, G
19350               strRng_Text1 = "A4:C" & strLastRow
19360               strRng_Text2 = "D4:D" & strLastRow
19370               strRng_Text3 = "G4:G" & strLastRow
19380               lngTrimTxt = 0&
19390               lngTxtCnt = 3&
                    ' ** Value Cols:
                    ' **   H, I
19400               strRng_Values1 = "H4:H" & strLastRow
19410               strRng_Values2 = "I4:I" & strLastRow
19420               lngValCnt = 2&
                    ' ** Date Cols:
                    ' **   E
19430               strRng_Dates1 = "E4:E" & strLastRow
19440               lngDatCnt = 1&
                    ' ** Decimal Cols:
                    ' **   F
19450               strRng_Decimals1 = "F4:F" & strLastRow
19460               lngDecCnt = 1&
19470             End Select
                  ' ** Integer Cols:
                  ' ** {none}
19480             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
19490             lngPctCnt = 0&
                  ' ** Column Headers.
19500             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
19510             blnColWidths = True
                  ' ** Report Title and Period.
19520             blnRptTitlePeriod = True
19530           Case "NY: Property on Hand at Ending of Account Period"
                  ' ***********************
                  ' ** New York:
                  ' ***********************
                  ' ** Rpt: Property on Hand at Ending of Account Period  (44)
                  ' ** =====================================================================
19540             If Len(strRptTitle) > 31 Then
19550               strSheetName = "Property on Hand at Period End"  '(30)
19560             End If
                  ' ** All Cols:
                  ' **   A - J
                  ' ** Text Cols:
                  ' **   A, B, C, D, F
19570             strRng_Text1 = "A4:C" & strLastRow
19580             strRng_Text2 = "D4:D" & strLastRow
19590             strRng_Text3 = "F4:F" & strLastRow
19600             lngTrimTxt = 0&
19610             lngTxtCnt = 3&
                  ' ** Value Cols:
                  ' **   G, H, I, J
19620             strRng_Values1 = "G4:G" & strLastRow
19630             strRng_Values2 = "H4:H" & strLastRow
19640             strRng_Values3 = "I4:I" & strLastRow
19650             strRng_Values4 = "J4:J" & strLastRow
19660             lngValCnt = 4&
                  ' ** Date Cols:
                  ' ** {none}
19670             lngDatCnt = 0&
                  ' ** Decimal Cols:
                  ' **   E
19680             strRng_Decimals1 = "E4:E" & strLastRow
19690             lngDecCnt = 1&
                  ' ** Integer Cols:
                  ' ** {none}
19700             lngIntCnt = 0&
                  ' ** Percent Cols:
                  ' ** {none}
19710             lngPctCnt = 0&
                  ' ** Column Headers.
19720             blnColHeads = False  ' ** No change.
                  ' ** Column Widths.
19730             blnColWidths = True
                  ' ** Report Title and Period.
19740             blnRptTitlePeriod = True
19750           End Select  ' ** strRptTitle.

                'NY COURT RPTS: 17
                '"NY: Summary Statement"
                '"NY: Statement of Principal Received - Schedule A"
                '"NY: Statement of Increases on Sales, Liquidation or Distribution - Schedule A-1"
                '"NY: Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability - Schedule B"
                '"NY: Statement of Administration Expenses Chargeable to Principal - Schedule C"
                '"NY: Statement of Administration Expenses Chargeable to Principal - Grouped - Schedule C"
                '"NY: Statement of Administration Expenses Chargeable to Income - Schedule C-2"
                '"NY: Statement of Administration Expenses Chargeable to Income - Grouped - Schedule C-2"
                '"NY: Statement of Distributions of Principal - Schedule D"
                '"NY: Statement of Distributions of Income - Schedule D-1"
                '"NY: Statement of New Investments, Exchanges and Stock Distributions of Principal Assets - Schedule E"
                '"NY: Statement of Principal Remaining on Hand - Schedule F"
                '"NY: Statement of Income Remaining on Hand - Schedule F-1"
                '"NY: Statement of Income Received - Schedule AA-1"
                '"NY: Statement of All Income Collected - Schedule A-2"
                '"NY: Statement of All Income Collected - Grouped - Schedule A-2"
                '"NY: Property on Hand at Ending of Account Period"
                '****** strRptStateTitle: NY: Property on Hand at Close of Accounting Period

19760           .Activate
19770           .Name = strSheetName

                ' ** Column Headers.
19780           If blnColHeads = True And strRng_Header2 = vbNullString Then

19790           End If  ' ** blnColHeads, strRng_Header.

                ' ** Column Widths.
19800           If blnColWidths = True Then
                  ' ** Set this standard, then we'll bleed the title across multiple columns.
19810             Set rng = .Range("B1:B" & strLastRow)
19820             rng.ColumnWidth = 35  ' ** Account Name.
19830             Set rng = Nothing
19840           End If  ' ** blnColWidths.

                ' ** Report Text.
19850           For lngX = 1& To lngTxtCnt
19860             Select Case lngX
                  Case 1&
19870               Set rng = .Range(strRng_Text1)  'A4:C  MISSING LAST ROW: 4
19880             Case 2&
19890               Set rng = .Range(strRng_Text2)
19900             Case 3&
19910               Set rng = .Range(strRng_Text3)
19920             Case 4&
19930               Set rng = .Range(strRng_Text4)
19940             Case 5&
19950               Set rng = .Range(strRng_Text5)
19960             End Select
19970             With rng
19980               .RowHeight = 13.5  ' ** Points.
19990               Set fnt = .Font
20000               With fnt
20010                 .Name = "Arial"
20020                 .Size = 10
20030                 .Bold = False
20040                 .Color = 0&
20050               End With  ' ** fnt.
20060               Set fnt = Nothing
20070               .HorizontalAlignment = xlLeft
20080               .VerticalAlignment = xlBottom
20090               If lngX = 1& Then
                      ' ** If accountno signals 'Number as Text' error, dismiss it.
20100                 For Each cel In rng
20110                   With cel
20120                     .Select
20130                     xlApp.ErrorCheckingOptions.NumberAsText = True
20140                     If xlApp.ErrorCheckingOptions.NumberAsText = True Then
20150                       xlApp.ErrorCheckingOptions.NumberAsText = False
20160                     End If
20170                   End With  ' ** cel.
20180                 Next  ' ** cel.
20190               End If
20200               If lngX = lngTrimTxt Then
20210                 For Each cel In rng
20220                   With cel
20230                     .Select
20240                     If IsNull(.Value) = False Then
20250                       strTmp01 = CStr(.Value)
20260                       If strTmp01 <> vbNullString Then
20270                         If Left(strTmp01, 4) <> "    " Then  ' ** Don't trim total lines.
20280                           .Value = Trim(strTmp01)
20290                         End If
20300                       End If
20310                     End If
20320                   End With  ' ** cel.
20330                 Next  ' ** cel.
20340               End If
20350             End With  ' ** rng.
20360             Set rng = Nothing
20370           Next  ' ** lngTxtCnt: lngX.

                ' ** Report Values.
20380           For lngX = 1& To lngValCnt
20390             Select Case lngX
                  Case 1&
20400               Set rng = .Range(strRng_Values1)
20410             Case 2&
20420               Set rng = .Range(strRng_Values2)
20430             Case 3&
20440               Set rng = .Range(strRng_Values3)
20450             Case 4&
20460               Set rng = .Range(strRng_Values4)
20470             Case 5&
20480               Set rng = .Range(strRng_Values5)
20490             Case 6&
20500               Set rng = .Range(strRng_Values6)
20510             End Select
20520             With rng
20530               Set fnt = .Font
20540               With fnt
20550                 .Name = "Arial"
20560                 .Size = 10
20570                 .Bold = False
20580                 .Color = 0&
20590               End With  ' ** fnt.
20600               Set fnt = Nothing
20610               .HorizontalAlignment = xlRight
20620               .VerticalAlignment = xlBottom
20630               For Each cel In rng
20640                 With cel
20650                   .Select
20660                   If IsNull(.Value) = False Then
20670                     strTmp01 = .Value
20680                     strTmp01 = Trim(strTmp01)
20690                     If strTmp01 <> vbNullString Then
20700                       strTmp01 = Rem_Dollar(strTmp01)  ' ** Module Function: modStringFuncs.
20710                       If IsNumeric(strTmp01) = True Then
20720                         .Value = CDbl(strTmp01)
20730                       End If
20740                     End If
20750                   End If
20760                   If xlApp.ErrorCheckingOptions.NumberAsText Then
20770                     If Trim(.Value) <> vbNullString Then
20780                       xlApp.WorksheetFunction.Trim (.Value)
20790                       .Value = .Value + 0
20800                     End If
20810                   End If
20820                 End With  ' ** cel.
20830               Next  ' ** cel.
20840               .NumberFormat = "$#,##0.00;($#,##0.00)"
20850             End With  ' ** rng.
20860             Set rng = Nothing
20870           Next  ' ** lngValCnt: lngX.

                ' ** Report Dates.
20880           For lngX = 1& To lngDatCnt
20890             Select Case lngX
                  Case 1&
20900               Set rng = .Range(strRng_Dates1)
20910             End Select
20920             With rng
20930               Set fnt = .Font
20940               With fnt
20950                 .Name = "Arial"
20960                 .Size = 10
20970                 .Bold = False
20980                 .Color = 0&
20990               End With  ' ** fnt.
21000               Set fnt = Nothing
21010               .HorizontalAlignment = xlLeft
21020               .VerticalAlignment = xlBottom
21030               If blnChkDates = True Then
21040                 For Each cel In rng
21050                   With cel
21060                     If IsNull(.Value) = False Then
21070                       strTmp01 = .Value
21080                       If IsDate(strTmp01) = True Then
21090                         .Value = CDate(strTmp01)
21100                       End If
21110                     End If
21120                   End With  ' ** cel.
21130                 Next  ' ** cel.
21140                 Set cel = Nothing
21150               End If
21160               .NumberFormat = "mm/dd/yyyy"  ' ** TextDate error only identifies 2-digit year.
21170             End With  ' ** rng.
21180             Set rng = Nothing
21190           Next  ' ** lngDatCnt: lngX.

                ' ** Report Integers.
21200           For lngX = 1& To lngIntCnt

21210           Next  ' ** lngIntCnt: lngX.

                ' ** Report Decimals.
21220           For lngX = 1& To lngDecCnt
21230             Select Case lngX
                  Case 1&
21240               Set rng = .Range(strRng_Decimals1)
21250             Case 2&
21260               Set rng = .Range(strRng_Decimals2)
21270             End Select
21280             With rng
21290               Set fnt = .Font
21300               With fnt
21310                 .Name = "Arial"
21320                 .Size = 10
21330                 .Bold = False
21340                 .Color = 0&
21350               End With  ' ** fnt.
21360               Set fnt = Nothing
21370               .HorizontalAlignment = xlRight
21380               .VerticalAlignment = xlBottom
21390               For Each cel In rng
21400                 With cel
21410                   .Select
21420                   xlApp.ErrorCheckingOptions.NumberAsText = True
21430                   If xlApp.ErrorCheckingOptions.NumberAsText = True Then
21440                     strTmp01 = .Value
21450                     If Trim(strTmp01) <> vbNullString Then
21460                       strTmp01 = Trim(strTmp01)
21470                       If IsNumeric(strTmp01) = True Then
21480                         .Value = CDbl(strTmp01)
21490                       End If
21500                     End If
21510                   End If
21520                 End With  ' ** cel.
21530               Next  ' ** cel.
21540               .NumberFormat = "#,##0.0000;-#,##0.0000"
21550             End With  ' ** rng.
21560             Set rng = Nothing
21570           Next  ' ** lngDecCnt: lngX.

                ' ** Report Percents.
21580           For lngX = 1& To lngPctCnt

21590           Next  ' ** lngPctCnt: lngX.

                ' ** Column Headers, for 2nd reports or anomalies.
21600           If blnColHeads = True And strRng_Header2 <> vbNullString Then
21610             Set rng = .Range(strRng_Header2)
21620             With rng
21630               Select Case lngRowHeight2
                    Case 0&
21640                 .RowHeight = lngRowHeight  ' ** Points.
21650               Case Else
21660                 .RowHeight = lngRowHeight2
21670               End Select
21680               .Interior.Color = 12632256  ' 192 192 192
21690               Set fnt = .Font
21700               With fnt
21710                 .Name = "Arial"
21720                 .Size = 10
21730                 .Bold = False
21740                 .Color = 0&
21750               End With
21760               Set fnt = Nothing
21770               For lngY = 1& To 5&
21780                 Select Case lngY
                      Case 1&
21790                   Set bdr = .Borders(xlEdgeLeft)
21800                 Case 2&
21810                   Set bdr = .Borders(xlEdgeTop)
21820                 Case 3&
21830                   Set bdr = .Borders(xlEdgeBottom)
21840                 Case 4&
21850                   Set bdr = .Borders(xlEdgeRight)
21860                 Case 5&
21870                   Set bdr = .Borders(xlInsideVertical)
21880                 End Select
21890                 With bdr
21900                   .Color = 0&
21910                   .LineStyle = xlContinuous
                        'If lngY < 5& Then
                        '  .Weight = xlMedium
                        'Else
21920                   .Weight = xlThin
                        'End If
21930                 End With  ' ** bdr.
21940                 Set bdr = Nothing
21950               Next  ' ** bdr.
21960               .HorizontalAlignment = xlCenter
21970               .VerticalAlignment = xlBottom
21980               .WrapText = False
21990               If lngRowHeight2 <> 0& Then
22000                 For Each cel In rng
22010                   With cel
22020                     If IsNull(.Value) = False Then
22030                       strTmp01 = Trim(CStr(.Value))
22040                       If strTmp01 = "Proceeds or Distribution Value" Then
22050                         strTmp01 = "Proceeds or" & vbLf & "Distribution" & vbLf & "Value"
22060                         .Value = strTmp01
22070                       ElseIf strTmp01 = "Inventory Value" Then
22080                         strTmp01 = "Inventory" & vbLf & "Value"
22090                         .Value = strTmp01
22100                       ElseIf InStr(strTmp01, "00/00") > 0 Then
22110                         strTmp01 = "Inventory Value" & vbLf & Format(FormRef("StartDate"), "mm/dd/yyyy")
22120                         .Value = strTmp01
22130                       End If
22140                     End If
22150                   End With  ' ** cel.
22160                 Next  ' ** cel.
22170               End If  ' ** lngRowHeight2.
22180             End With  ' ** rng
22190             Set rng = Nothing
22200           End If  ' ** blnColHeads, strRng_Header2.

                ' ** Report Title and Period.
22210           If blnRptTitlePeriod = True Then
                  ' ** Allow the title to bleed across multiple columns.
22220             Set rng = .Range(strRng_RptTitle)
22230             rng.WrapText = False
22240             Set rng = Nothing
22250             Set rng = .Range(strRng_RptTitleRight)
22260             rng.ClearContents
22270             Set rng = Nothing
22280           End If  ' ** blnRptTitlePeriod.

22290           .Range("A2").Select

22300         End With  ' ** wks.
22310         Set wks = Nothing
22320       End If  ' ** Count.
22330       .Save  ' ** wbk.Close SaveChanges:=True
22340       .Close
22350     End With  ' ** wbk.
22360     Set wbk = Nothing
22370     xlApp.DisplayAlerts = True
22380     xlApp.Interactive = True
22390     xlApp.Quit

22400   Else
22410     blnRetVal = False
22420   End If  ' ** vbNullString.

      #End If

        ' ** XlFindLookin enumeration:
        ' **   -4123  xlFormulas
        ' **   -4144  xlComments
        ' **   -4163  xlValues

        ' ** XlLookAt enumeration:
        ' **   1  xlWhole
        ' **   2  xlPart

        ' ** XlSearchOrder enumeration:
        ' **   1  xlByRows
        ' **   2  xlByColumns

        ' ** XlSearchDirection enumeration:
        ' **   1  xlNext      (Default)
        ' **   2  xlPrevious

EXITP:
      #If NoExcel Then
        ' ** Skip.
      #Else
22430   Set xlApp = Nothing
22440   Set wbk = Nothing
22450   Set wks = Nothing
22460   Set rng = Nothing
22470   Set fnt = Nothing
22480   Set bdr = Nothing
22490   Set cel = Nothing
      #End If
22500   Excel_Court = blnRetVal
22510   Exit Function

ERRH:
22520   blnRetVal = False
      #If NoExcel Then
        ' ** Skip.
      #Else
22530   If blnExcelOpen = True Then
22540     wbk.Close
22550     xlApp.Quit
22560   End If
      #End If
22570   Select Case ERR.Number
        Case Else
22580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22590   End Select
22600   Resume EXITP

End Function

Public Function Excel_Holdings(strPathFile As String, blnAll As Boolean) As Boolean
' ** Export asset holdings report.
' ** Called by:
' **   frmRpt_Holdings:
' **     cmdExcel_Click()
'NAME, TRADE DATE, COST, LOCATION:
' SPECIFIED, DETAIL
' ALL, DETAIL

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "Excel_Holdings"

      #If NoExcel Then
        ' ** Skip the whole function.
      #Else

      #If IsDev Then
        Dim xlApp As Excel.Application, wbk As Excel.Workbook, wks As Excel.Worksheet, rng As Excel.Range              ' ** Early Binding.
        Dim fnt As Excel.Font, bdr As Excel.Border, cel As Excel.Range
      #Else
        Dim xlApp As Object, wbk As Object, wks As Object, rng As Object, fnt As Object, bdr As Object, cel As Object  ' ** Late Binding.
      #End If
        Dim strSheetName As String, blnChkDetail As Boolean
        Dim strLastCell As String, strLastCol As String, strLastRow As String, strRng_All As String
        Dim strRng_Header As String, strRng_Title As String
        Dim strRng_Text1 As String, strRng_Text2 As String
        Dim strRng_Dates1 As String
        Dim strRng_Decimals As String, strRng_Values As String
        Dim lngCols As Long
        Dim blnExcelOpen As Boolean
        Dim lngX As Long, lngY As Long

      #End If
        Dim blnRetVal As Boolean

      #If NoExcel Then
        ' ** Skip.
      #Else

22710   blnRetVal = True

22720   If strPathFile <> vbNullString Then

22730     strSheetName = vbNullString
22740     If InStr(strPathFile, "Detailed") > 0 Then
22750       strSheetName = "Holdings - Detailed"
22760       blnChkDetail = True
22770     ElseIf InStr(strPathFile, "Detailed") = 0 Then
22780       strSheetName = "Holdings"
22790       blnChkDetail = False
22800     End If
22810     Select Case blnAll
          Case True
22820       strSheetName = strSheetName & " - All"
22830     Case False
            ' ** As is.
22840     End Select

      #If IsDev Then
22850     Set xlApp = New Excel.Application              ' ** Early Binding.
      #Else
22860     Set xlApp = CreateObject("Excel.Application")  ' ** Late Binding.
      #End If
22870     blnExcelOpen = True

22880     xlApp.Visible = False
22890     xlApp.DisplayAlerts = False
22900     xlApp.Interactive = False
22910     Set wbk = xlApp.Workbooks.Open(strPathFile)
22920     With wbk
22930       If .Worksheets.Count > 0 Then
22940         Set wks = .Worksheets(1)
22950         With wks

22960           .Name = strSheetName

22970           strLastCell = .Cells.SpecialCells(xlCellTypeLastCell).Address  '$H$22205
22980           strLastCell = Rem_Dollar(strLastCell)  ' ** Module Function: modStringFuncs.
22990           strRng_All = "A1:" & strLastCell
23000           Set rng = .Range(strRng_All)
23010           lngCols = rng.Columns.Count
23020           Set rng = Nothing

23030           strLastCol = Left(strLastCell, 1)  ' ** Assumes single-letter address (26 or fewer columns).
23040           strLastRow = Mid(strLastCell, 2)
23050           strRng_Header = "A1:" & strLastCol & "1"
23060           strRng_Title = "A2:B3"
                ' ** All Cols:
                ' **   A - H
                ' ** Text Cols:
                ' **   A, B, C, D, H
23070           strRng_Text1 = "A4:D" & strLastRow
23080           strRng_Text2 = "H4:H" & strLastRow
                ' ** Date Cols:
                ' **    E
23090           strRng_Dates1 = "E4:E" & strLastRow
                ' ** Decimal Cols:
                ' **   F
23100           strRng_Decimals = "F4:F" & strLastRow
                ' ** Value Cols:
                ' **   G
23110           strRng_Values = "G4:G" & strLastRow

                ' ** Column Headers.
23120           Set rng = .Range(strRng_Header)
23130           With rng
23140             .RowHeight = 13.5  ' ** Points.
23150             .Interior.Color = 12632256  ' 192 192 192
23160             Set fnt = .Font
23170             With fnt
23180               .Name = "Arial"
23190               .Size = 10
23200               .Bold = False
23210               .Color = 0&
23220             End With
23230             Set fnt = Nothing
23240             For lngY = 1& To 5&
23250               Select Case lngY
                    Case 1&
23260                 Set bdr = .Borders(xlEdgeLeft)
23270               Case 2&
23280                 Set bdr = .Borders(xlEdgeTop)
23290               Case 3&
23300                 Set bdr = .Borders(xlEdgeBottom)
23310               Case 4&
23320                 Set bdr = .Borders(xlEdgeRight)
23330               Case 5&
23340                 Set bdr = .Borders(xlInsideVertical)
23350               End Select
23360               With bdr
23370                 .Color = 0&
23380                 .LineStyle = xlContinuous
23390                 If lngY < 5& Then
23400                   .Weight = xlMedium
23410                 Else
23420                   .Weight = xlThin
23430                 End If
23440               End With  ' ** bdr.
23450               Set bdr = Nothing
23460             Next  ' ** bdr.
23470             .HorizontalAlignment = xlCenter
23480             .VerticalAlignment = xlBottom
23490             .WrapText = False
23500           End With  ' ** rng
23510           Set rng = Nothing

23520           Set rng = .Range("H1")
23530           With rng
23540             .Value = "Location"
23550           End With
23560           Set rng = Nothing

                ' ** Column Widths.
23570           For lngX = 1& To lngCols
23580             Set rng = .Range(Chr(64& + lngX) & "1:" & Chr(64& + lngX) & strLastRow)
23590             With rng
                    ' ** Width is font-based, and dependent on column header.
                    ' **      A             B         C      D        E           F        G       H
                    ' ** Account Num  Account Name  CUSIP  Asset  Trade Date  Share/Face  Cost  Location
                    ' ** ===========  ============  =====  =====  ==========  ==========  ====  ========
                    ' **      1             2         3      4        5           6        7       8
23600               Select Case lngX
                    Case 2&, 8&
23610                 .ColumnWidth = 35  ' ** Account Name, Location.
23620               Case 3&
23630                 .ColumnWidth = 12  ' ** CUSIP.
23640               Case 4&
23650                 .ColumnWidth = 75  ' ** Asset.
23660               Case 5&
23670                 .ColumnWidth = 12  ' ** Trade Date
23680               Case Else
23690                 .ColumnWidth = 15
23700               End Select
23710             End With
23720           Next
23730           Set rng = Nothing

                ' ** Report Title and Period.
23740           Set rng = .Range(strRng_Title)
23750           With rng
23760             .RowHeight = 13.5  ' ** Points.
23770             Set fnt = .Font
23780             With fnt
23790               .Name = "Arial"
23800               .Size = 10
23810               .Bold = False
23820               .Color = 0&
23830             End With  ' ** fnt.
23840             Set fnt = Nothing
23850             .HorizontalAlignment = xlLeft
23860             .VerticalAlignment = xlBottom
23870           End With  ' ** rng.
23880           Set rng = Nothing

                ' ** Report Text.
23890           For lngX = 1& To 2&
23900             Select Case lngX
                  Case 1&
23910               Set rng = .Range(strRng_Text1)
23920             Case 2&
23930               Set rng = .Range(strRng_Text2)
23940             End Select
23950             With rng
23960               .RowHeight = 13.5  ' ** Points.
23970               Set fnt = .Font
23980               With fnt
23990                 .Name = "Arial"
24000                 .Size = 10
24010                 .Bold = False
24020                 .Color = 0&
24030               End With  ' ** fnt.
24040               Set fnt = Nothing
24050               .HorizontalAlignment = xlLeft
24060               .VerticalAlignment = xlBottom
24070               If lngX = 1& Then
                      ' ** If accountno signals 'Number as Text' error, dismiss it.
24080                 For Each cel In rng
24090                   With cel
24100                     .Select
24110                     If xlApp.ErrorCheckingOptions.NumberAsText Then
24120                       xlApp.ErrorCheckingOptions.NumberAsText = False
24130                     End If
24140                   End With  ' ** cel.
24150                 Next  ' ** cel.
24160               End If
24170             End With  ' ** rng.
24180             Set rng = Nothing
24190           Next  ' ** lngX.

                ' ** Report Values.
24200           Set rng = .Range(strRng_Values)
24210           With rng
24220             Set fnt = .Font
24230             With fnt
24240               .Name = "Arial"
24250               .Size = 10
24260               .Bold = False
24270               .Color = 0&
24280             End With  ' ** fnt.
24290             Set fnt = Nothing
24300             .HorizontalAlignment = xlRight
24310             .VerticalAlignment = xlBottom
24320             For Each cel In rng
24330               With cel
24340                 .Select
24350                 If xlApp.ErrorCheckingOptions.NumberAsText Then
24360                   If Trim(.Value) <> vbNullString Then
24370                     xlApp.WorksheetFunction.Trim (.Value)
24380                     .Value = .Value + 0
24390                   End If
24400                 End If
24410               End With  ' ** cel.
24420             Next  ' ** cel.
24430             .NumberFormat = "$#,##0.00;($#,##0.00)"
24440           End With  ' ** rng.
24450           Set rng = Nothing

                ' ** Report Dates.
24460           For lngX = 1& To 1&
24470             Select Case lngX
                  Case 1&
24480               Set rng = .Range(strRng_Dates1)
24490             End Select
24500             With rng
24510               Set fnt = .Font
24520               With fnt
24530                 .Name = "Arial"
24540                 .Size = 10
24550                 .Bold = False
24560                 .Color = 0&
24570               End With  ' ** fnt.
24580               Set fnt = Nothing
24590               .HorizontalAlignment = xlLeft
24600               .VerticalAlignment = xlBottom
24610               .NumberFormat = "mm/dd/yyyy"  ' ** TextDate error only identifies 2-digit year.
24620             End With  ' ** rng.
24630             Set rng = Nothing
24640           Next  ' ** lngX.

                ' ** Report Decimals.
24650           Set rng = .Range(strRng_Decimals)
24660           With rng
24670             Set fnt = .Font
24680             With fnt
24690               .Name = "Arial"
24700               .Size = 10
24710               .Bold = False
24720               .Color = 0&
24730             End With  ' ** fnt.
24740             Set fnt = Nothing
24750             .HorizontalAlignment = xlRight
24760             .VerticalAlignment = xlBottom
24770             For Each cel In rng
24780               With cel
24790                 .Select
24800                 If xlApp.ErrorCheckingOptions.NumberAsText Then
24810                   If Trim(.Value) <> vbNullString Then
24820                     xlApp.WorksheetFunction.Trim (.Value)
24830                     .Value = .Value + 0
24840                   End If
24850                 End If
24860               End With  ' ** cel.
24870             Next  ' ** cel.
24880             .NumberFormat = "#,##0.0000;-#,##0.0000"
24890           End With  ' ** rng.
24900           Set rng = Nothing

24910           .Range("A2").Select

24920         End With  ' ** wks.
24930         Set wks = Nothing
24940       End If  ' ** Count.
24950       .Save  ' ** wbk.Close SaveChanges:=True
24960       .Close
24970     End With  ' ** wbk.
24980     Set wbk = Nothing
24990     xlApp.DisplayAlerts = True
25000     xlApp.Interactive = True
25010     xlApp.Quit
25020   End If  ' ** vbNullString.

25030   Beep

      #End If

        ' ** XlCellType enumeration:
        ' **   -4175  xlCellTypeSameValidation        Cells having the same validation criteria.
        ' **   -4174  xlCellTypeAllValidation         Cells having validation criteria.
        ' **   -4173  xlCellTypeSameFormatConditions  Cells having the same format.
        ' **   -4172  xlCellTypeAllFormatConditions   Cells of any format.
        ' **   -4144  xlCellTypeComments              Cells containing notes.
        ' **   -4123  xlCellTypeFormulas              Cells containing formulas.
        ' **       2  xlCellTypeConstants             Cells containing constants.
        ' **       4  xlCellTypeBlanks                Empty cells.
        ' **      11  xlCellTypeLastCell              The last cell in the used range.
        ' **      12  xlCellTypeVisible               All visible cells.

        ' ** XlSpecialCellsValue enumeration:
        ' **    1  xlNumbers
        ' **    2  xlTextValues
        ' **    4  xlLogical
        ' **   16  xlErrors

        ' ** XlDirection enumeration:  (Excel 2007)
        ' **   -4162  xlUp       Up.
        ' **   -4161  xlToRight  To right.
        ' **   -4159  xlToLeft   To left.
        ' **   -4121  xlDown     Down.

        ' ** Borders enumeration:
        ' **    5  xlDiagonalDown
        ' **    6  xlDiagonalUp
        ' **    7  xlEdgeLeft
        ' **    8  xlEdgeTop
        ' **    9  xlEdgeBottom
        ' **   10  xlEdgeRight
        ' **   11  xlInsideVertical
        ' **   12  xlInsideHorizontal

        ' ** HorizontalAlignment enumeration:
        ' **   -4152  xlRight
        ' **   -4131  xlLeft
        ' **   -4130  xlJustify
        ' **   -4117  xlDistributed
        ' **   -4108  xlCenter

        ' ** VerticalAlignment enumeration:
        ' **   -4160  xlTop
        ' **   -4130  xlJustify
        ' **   -4117  xlDistributed
        ' **   -4108  xlCenter
        ' **   -4107  xlBottom

        ' ** XlLineStyle enumeration:
        ' **   -4142  xlLineStyleNone  No line.
        ' **   -4126  xlGray75         75% gray pattern.
        ' **   -4125  xlGray50         50% gray pattern.
        ' **   -4124  xlGray25         25% gray pattern.
        ' **   -4119  xlDouble         Double line.
        ' **   -4118  xlDot            Dotted line.
        ' **   -4115  xlDash           Dashed line.
        ' **   -4105  xlAutomatic      Excel applies automatic settings, such as a color, to the specified object.
        ' **       1  xlContinuous     Continuous line.
        ' **       4  xlDashDot        Alternating dashes and dots.
        ' **       5  xlDashDotDot     Dash followed by two dots.
        ' **      13  xlSlantDashDot   Slanted dashes.
        ' **      17  xlGray16         16% gray pattern.
        ' **      18  xlGray8          8% gray pattern.

        ' ** XlBorderWeight enumeration:
        ' **   -4138  xlMedium    Medium.
        ' **       1  xlHairline  Hairline (thinnest border).
        ' **       2  xlThin      Thin.
        ' **       4  xlThick     Thick (widest border).

EXITP:
      #If NoExcel Then
        ' ** Skip.
      #Else
25040   Set bdr = Nothing
25050   Set fnt = Nothing
25060   Set rng = Nothing
25070   Set wks = Nothing
25080   Set wbk = Nothing
25090   Set xlApp = Nothing
      #End If
25100   Excel_Holdings = blnRetVal
25110   Exit Function

ERRH:
25120   blnRetVal = False
      #If NoExcel Then
        ' ** Skip.
      #Else
25130   If blnExcelOpen = True Then
25140     wbk.Close
25150     xlApp.Quit
25160   End If
      #End If
25170   Select Case ERR.Number
        Case Else
25180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25190   End Select
25200   Resume EXITP

End Function
