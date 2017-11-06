Attribute VB_Name = "zz_mod_PennsvilleFuncs"
Option Compare Database
Option Explicit

'VGC 01/01/2015: CHANGES!

' ** No assetdate: Dividend, Interest, Misc., Paid
' ** Trade Date: assetdate
' ** Settle Date: Not used.
' ** Audit Date: transdate, only where Trade Date > Posted Date.
' ** Posted Date: transdate, except where Trade Date > Posted Date.

' ** REVID_INC:
' **   Dividend, Interest, Deposit, Purchase, Withdrawn, Sold, Received, Cost Adj.
' **   Misc. icash > 0
' ** REVID_EXP:
' **   Paid, Liability
' **   Misc. icash <= 0

Private Const THIS_NAME As String = "zz_mod_PennsvilleFuncs"
' **

Public Function PNB_RatesDates(varInput As Variant, strMode As String) As Variant

  Const THIS_PROC As String = "PNB_RatesDates"

  Dim intPos1 As Integer, intLen As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim intX As Integer
  Dim varRetVal As Variant

  varRetVal = 0

  If IsNull(varInput) = False Then
    If Trim(varInput) <> vbNullString Then
'BAYLAKE BANK WISCONSIN CERT OF DEP 5.1% 8/24/09
      strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
      Select Case strMode
      Case "r"
        intPos1 = InStr(varInput, "%")
        If intPos1 > 0 Then
          For intX = intPos1 To 1 Step -1
            If Mid(varInput, intX, 1) = " " Then
              strTmp01 = Left(varInput, intX)  ' ** Includes space.
              strTmp02 = Mid(varInput, (intX + 1))  ' ** Begins with rate.
              Exit For
            End If
          Next
          If strTmp01 <> vbNullString And strTmp02 <> vbNullString Then
            intPos1 = InStr(strTmp02, " ")  ' ** Space after rate.
            If intPos1 > 0 Then
              strTmp03 = Trim(Left(strTmp02, intPos1))  ' ** Just rate.
            Else
              strTmp03 = strTmp02  ' ** At end of line.
            End If
            strTmp03 = Left(strTmp03, (Len(strTmp03) - 1)) ' ** Without percent sign.
            varRetVal = (Val(strTmp03) / 100)
          End If
        End If
      Case "d", "dx"
        varTmp00 = varInput
        intPos1 = InStr(varTmp00, "/")
        If intPos1 = 0 Then
          varTmp00 = StringReplace(CStr(varTmp00), "-", "/")
          intPos1 = InStr(varTmp00, "/")
        End If
        If intPos1 > 0 Then
          For intX = intPos1 To 1 Step -1
            If Mid(varTmp00, intX, 1) = " " Then
              strTmp01 = Left(varTmp00, intX)  ' ** Includes space.
              strTmp02 = Mid(varTmp00, (intX + 1))  ' ** Begins with date.
              Exit For
            End If
          Next
          If strTmp01 <> vbNullString And strTmp02 <> vbNullString Then
            intPos1 = InStr(strTmp02, " ")  ' ** Space after date.
            If intPos1 > 0 Then
              strTmp03 = Trim(Left(strTmp02, intPos1))  ' ** Just date.
            Else
              strTmp03 = strTmp02  ' ** At end of line.
            End If
            If IsDate(strTmp03) = True Then
              varRetVal = CDate(strTmp03)
            End If
            If strMode = "dx" Then
              ' ** Replace the date within the description.
              intPos1 = InStr(strTmp02, " ")
              If intPos1 > 0 Then
                varRetVal = Trim(strTmp01) & " " & Format(varRetVal, "mm/dd/yyyy") & " " & Trim(Mid(strTmp02, intPos1))
                strTmp02 = Trim(Mid(strTmp02, intPos1))
              Else
                varRetVal = Trim(strTmp01) & " " & Format(varRetVal, "mm/dd/yyyy")
                strTmp02 = vbNullString
              End If
              If strTmp02 <> vbNullString Then
                intPos1 = InStr(strTmp02, "/")  ' ** Check for 2nd date.
                If intPos1 > 0 Then
                  strTmp01 = Left(varRetVal, (Len(varRetVal) - Len(strTmp02)))  ' ** Includes new date.
                  strTmp03 = vbNullString
                  For intX = intPos1 To 1 Step -1
                    If Mid(strTmp02, intX, 1) = " " Then
                      strTmp01 = strTmp01 & Left(strTmp02, intX)  ' ** Includes space.
                      strTmp02 = Mid(strTmp02, (intX + 1))  ' ** Begins with date.
                      intPos1 = InStr(strTmp02, " ")
                      If intPos1 > 0 Then
                        strTmp03 = Trim(Left(strTmp02, intPos1))  ' ** Just date.
                        strTmp02 = Trim(Mid(strTmp02, intPos1))
                      Else
                        strTmp03 = strTmp02  ' ** End of line.
                        strTmp02 = vbNullString
                      End If
                      Exit For
                    End If
                  Next
                  If strTmp03 <> vbNullString Then
                    If IsDate(strTmp03) = True Then
                      varRetVal = Trim(strTmp01 & " " & Format(CDate(strTmp03), "mm/dd/yyyy") & " " & strTmp02)
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End Select
    End If
  End If

  PNB_RatesDates = varRetVal

End Function

Public Function PNB_TransCodes() As Boolean

  Const THIS_PROC As String = "PNB_TransCodes"

  Dim blnRetVal As Boolean

  blnRetVal = True

'ADJ TO INCREASE BOOK VALUE
'ADJ TO DECREASE BOOK VALUE
'ADJ TO INCREASE SHARES
'ADJ TO DECREASE SHARES
'ADJ-INCREASE SHARES/UNITS
'AGENT FEES
'ASSET DISTRIBUTED
'ATTORNEY FEE
'BENEFICIARY DISTRIBUTION
'BROKERS FEE
'CASH DEPOSIT
'CASH DISTRIBUTION
'CUSTODIAL FEE
'DEPOSIT OF CASH
'DIVIDEND
'FEDERAL FIDUCIARY TAXES
'INT OTHER BANK
'INT OWN BANK
'INTEREST OTHER BANK
'INTEREST OWN BANK
'INVESTMENT MANAGEMENT FEE
'PURCHASE OF
'REQUISITION
'RETURNED ITEM
'SALE OF ASSET
'SAVINGS
'SURROGATES FEE
'TAXES PAID
'TRANSFER INCOME CASH TO PRINCIPAL CASH
'TRANSFER PRINCIPAL CASH TO INCOME CASH
'TRUSTEE FEE
'UTILITIES EXPENSE

'ADJ-DECREASE COST BASIS
'ADJ-INCREASE COST BASIS
'ACCOUNTANT FEE
'APPRAISAL FEE

  PNB_TransCodes = blnRetVal

End Function

Public Function PNB_TranNumHoles() As Boolean

  Const THIS_PROC As String = "PNB_TranNumHoles"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim lngTNums As Long, arr_varTNum() As Variant
  Dim lngLastTNum As Long, lngRecs As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTNum().
  Const T_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const T_TNUM As Integer = 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngTNums = 0&
  ReDim arr_varTNum(T_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    ' ** Empty tblMark_AutoNum3.
    Set qdf = .QueryDefs("zzz_qry_Pennsville_Trans_75_09")
    qdf.Execute
    Set qdf = Nothing

    'Set qdf = .QueryDefs("zzz_qry_Pennsville_Trans_75_10")
    'Set rst1 = qdf.OpenRecordset
    Set rst1 = .OpenRecordset("tblPennsville_Transaction_excel", dbOpenDynaset, dbConsistent)
    'rst1.Sort = "[Transaction Number]"
    rst1.Sort = "[Transaction Number]"
    Set rst2 = rst1.OpenRecordset
    With rst2
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      lngLastTNum = 0&
      For lngX = 1& To lngRecs
        If ![Transaction Number] <> (lngLastTNum + 1&) Then
          Do While ![Transaction Number] > (lngLastTNum + 1&)
            lngLastTNum = lngLastTNum + 1&
            lngTNums = lngTNums + 1&
            lngE = lngTNums - 1&
            ReDim Preserve arr_varTNum(T_ELEMS, lngE)
            arr_varTNum(T_TNUM, lngE) = lngLastTNum
          Loop
        End If
        lngLastTNum = ![Transaction Number]
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With
    rst1.Close
    Set rst1 = Nothing
    Set rst2 = Nothing
    Set qdf = Nothing

    If lngTNums > 0& Then

      Debug.Print "'HOLES: " & CStr(lngTNums)
      DoEvents

      Set rst1 = .OpenRecordset("tblMark_AutoNum3", dbOpenDynaset, dbConsistent)
      With rst1
        For lngX = 0& To (lngTNums - 1&)
          .AddNew
          ![unique_id] = arr_varTNum(T_TNUM, lngX)
          ![mark] = False
          ![Value] = Null
          ' ** ![autonum_id] : AutoNumber.
          .Update
        Next
        .Close
      End With
      Set rst1 = Nothing

    Else
      Debug.Print "'NONE FOUND!"
      DoEvents
    End If

    .Close
  End With

  Debug.Print "'DONE!"
  DoEvents

'HOLES: 759
'DONE!
  Beep

  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  PNB_TranNumHoles = blnRetVal

End Function

Public Function PNB_JNoHoles() As Boolean

  Const THIS_PROC As String = "PNB_JNoHoles"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngJNos As Long, arr_varJNo As Variant
  Dim lngRecs As Long
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    Set rst = .OpenRecordset("tblPennsville_Ledger]", dbOpenDynaset, dbConsistent)
    With rst
      .MoveLast
      lngJNos = .RecordCount
      .MoveFirst
      arr_varJNo = .GetRows(lngJNos)
      .Close
    End With
    Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbConsistent)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 0& To (lngJNos - 1)
        .Edit
        ![Value] = arr_varJNo(0, lngX)
        .Update
        If ((lngX < lngRecs) And (lngX < (lngJNos - 1))) Then .MoveNext
      Next
      .Close
    End With
    .Close
  End With

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  PNB_JNoHoles = blnRetVal

End Function

Public Function PNB_TaxLots() As Boolean

  Const THIS_PROC As String = "PNB_TaxLots"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset
  Dim lngTotDates As Long, datThisDate As Date
  Dim lngPT3ID As Long, lngPAAID As Long
  Dim dblShares As Double, dblICash As Double, dblPCash As Double, dblCost As Double
  Dim dblTmpShares As Double, dblTmpICash As Double, dblTmpPCash As Double, dblTmpCost As Double
  Dim strAccountNo As String, lngAssetNo As Long, datTransDate As Date, datAssetDate As Date
  Dim lngLots As Long, arr_varLot() As Variant
  Dim lngAdjs As Long, arr_varAdj() As Variant
  Dim strDesc As String, strRecur As String
  Dim lngRecs As Long, lngLotNum As Long, lngSetNum As Long
  Dim lngPT3ID_New As Long, lngJNum_New As Long, lngTNum_New As Long
  Dim blnSkip As Boolean
  Dim dblTmp01 As Double
  Dim lngW As Long, lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varLot().
  Const L_ELEMS As Integer = 14  ' ** Array's first-elemen UBound().
  Const L_ACTNO As Integer = 0
  Const L_ASTNO As Integer = 1
  Const L_TDATE As Integer = 2
  Const L_ADATE As Integer = 3
  Const L_LNUM  As Integer = 4
  Const L_CLOSD As Integer = 5
  Const L_SHRS  As Integer = 6
  Const L_ICSH  As Integer = 7
  Const L_PCSH  As Integer = 8
  Const L_COST  As Integer = 9
  Const L_PAAID As Integer = 10
  Const L_RSHRS As Integer = 11
  Const L_RICSH As Integer = 12
  Const L_RPCSH As Integer = 13
  Const L_RCOST As Integer = 14

  ' ** Array: arr_varAdj().
  Const A_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const A_ACTNO As Integer = 0
  Const A_ASTNO As Integer = 1
  Const A_TDATE As Integer = 2
  Const A_ADATE As Integer = 3
  Const A_COST  As Integer = 4

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    lngPT3ID_New = 0&: lngJNum_New = 0&: lngTNum_New = 0&
    lngSetNum = 11&  ' ** For zz_tbl_Pennsville_Tmp19.

    ' ** Update zz_tbl_Pennsville_Tmp18, for pt18_mark = False, all.
    Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_46_03")
    qdf1.Execute
    Set qdf1 = Nothing

    Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp18", dbOpenDynaset, dbReadOnly)
    With rst1
      .MoveLast
      lngTotDates = .RecordCount
      .Close
    End With
    Set rst1 = Nothing

    For lngW = 1& To lngTotDates

      ' ** zz_tbl_Pennsville_Tmp18, just pt18_mark = False, sorted.
      Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_46_01")
      Set rst1 = qdf1.OpenRecordset
      With rst1
        .MoveFirst
        datThisDate = ![pt18_transdate]
        .Close
      End With
      Set rst1 = Nothing
      Set qdf1 = Nothing

'If datThisDate = #10/1/2003# Then
'Stop
'End If
      ' ** tblPennsville_Transaction3, by specified [tdat].
      Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_47")
      With qdf1.Parameters
        ![tdat] = datThisDate
      End With
      Set rst1 = qdf1.OpenRecordset
      With rst1
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs

'If ![journalno] = 12 Or ![journalno] = 13 Then
'Stop
'End If
          lngPT3ID = ![pt3_id]
          strDesc = vbNullString: strRecur = vbNullString

          Select Case ![journaltype]
          Case "Dividend"
            ' ** zzz_qry_Pennsville_Trans_48_01_01 (tblPennsville_Transaction3,
            ' ** just 'Dividend'), with ledger_description, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_01_02")
          Case "Interest"
            ' ** zzz_qry_Pennsville_Trans_48_02_01 (tblPennsville_Transaction3,
            ' ** just 'Interest'), with ledger_description, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_02_02")
          Case "Misc."
            ' ** zzz_qry_Pennsville_Trans_48_03_01 (tblPennsville_Transaction3,
            ' ** just 'Misc.'), with ledger_description, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_03_02")
          Case "Received"
            ' ** zzz_qry_Pennsville_Trans_48_04_01 (tblPennsville_Transaction3,
            ' ** just 'Received'), with ledger_description, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_04_02")
          Case "Paid"
            ' ** zz_tbl_Pennsville_Tmp26, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_05_26")
          Case "Deposit", "Purchase"
            ' ** zzz_qry_Pennsville_Trans_48_06_01 (tblPennsville_Transaction3,
            ' ** just 'Deposit', 'Purchase'), with ledger_description, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_06_02")
          Case "Withdrawn", "Sold"
            ' ** zzz_qry_Pennsville_Trans_48_06_01 (tblPennsville_Transaction3,
            ' ** just 'Withdrawn', 'Sold'), with ledger_description, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_07_02")
          Case "Cost Adj."
            ' ** zzz_qry_Pennsville_Trans_48_08_01 (tblPennsville_Transaction3,
            ' ** just 'Cost Adj.'), with ledger_description, by specified [pt3id].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_48_08_02")
          Case "Liability"
            ' ** There are none.
            Stop
          End Select
          With qdf2.Parameters
            ![pt3id] = lngPT3ID
          End With
          Set rst2 = qdf2.OpenRecordset
          rst2.MoveFirst
          Select Case ![journaltype]
          Case "Misc.", "Received", "Paid"
            strRecur = rst2![RecurringItem]
          End Select
          If IsNull(rst2![ledger_description]) = False Then
            strDesc = rst2![ledger_description]
          End If
          rst2.Close
          Set rst2 = Nothing
          Set qdf2 = Nothing

          Select Case ![journaltype]
          Case "Withdrawn", "Sold"

            strAccountNo = ![accountno]
            lngAssetNo = ![assetno]
            datTransDate = ![transdate]
            datAssetDate = 0#

            dblShares = ![shareface]
            dblICash = ![ICash]
            dblPCash = ![PCash]
            dblCost = ![Cost]

            dblTmpShares = dblShares
            dblTmpICash = dblICash
            dblTmpPCash = dblPCash
            dblTmpCost = dblCost

            lngLots = 0&
            ReDim arr_varLot(L_ELEMS, 0)

            lngLotNum = 0&

            ' ** tblPennsville_ActiveAssets, by specified [actno], [astno], [tdat].
            Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_49_01")
            With qdf2.Parameters
              ![actno] = strAccountNo
              ![astno] = lngAssetNo
              ![tdat] = datTransDate
            End With
            Set rst2 = qdf2.OpenRecordset
            With rst2
              If .BOF = True And .EOF = True Then
                ' ** None available!
'Debug.Print "'NONE AVAILABLE!  ACTNO: " & strAccountNo & "  ASTNO: " & CStr(lngAssetNo) & "  TDATE: " & datThisDate & "  SHARES: " & CStr(rst1![shareface])
DoEvents
                Set rst3 = dbs.OpenRecordset("zz_tbl_Pennsville_Tmp19", dbOpenDynaset, dbAppendOnly)
                With rst3
                  .AddNew
                  ' ** ![pt19_id] : AutoNumber.
                  ![pt19_set] = lngSetNum
                  ![accountno] = strAccountNo
                  ![assetno] = lngAssetNo
                  ![transdate] = datThisDate
                  ![journalno] = rst1![journalno]
                  ![journaltype] = rst1![journaltype]
                  ![shareface] = rst1![shareface]
                  ![IsShort] = True
                  ![pt19_datemodified] = Now()
                  .Update
                  .Close
                End With
                Set rst3 = Nothing

              Else
                .MoveLast
                lngRecs = .RecordCount
                .MoveFirst
                For lngY = 1& To lngRecs
                  lngPAAID = ![paa_id]
                  If dblTmpShares = ![shareface] Then
                    ' ** Tax Lot Zeroed.
                    lngLotNum = lngLotNum + 1&
                    lngLots = lngLots + 1&
                    lngE = lngLots - 1&
                    ReDim Preserve arr_varLot(L_ELEMS, lngE)
                    arr_varLot(L_ACTNO, lngE) = strAccountNo
                    arr_varLot(L_ASTNO, lngE) = lngAssetNo
                    arr_varLot(L_TDATE, lngE) = datTransDate
                    arr_varLot(L_ADATE, lngE) = ![assetdate]
                    arr_varLot(L_LNUM, lngE) = lngLotNum
                    arr_varLot(L_CLOSD, lngE) = CBool(True)
                    arr_varLot(L_SHRS, lngE) = dblTmpShares
                    dblTmpShares = 0#
                    arr_varLot(L_ICSH, lngE) = dblTmpICash
                    dblTmpICash = 0#
                    arr_varLot(L_PCSH, lngE) = dblTmpPCash
                    dblTmpPCash = 0#
                    arr_varLot(L_COST, lngE) = dblTmpCost
                    dblTmpCost = 0#
                    arr_varLot(L_PAAID, lngE) = lngPAAID
                    arr_varLot(L_RSHRS, lngE) = 0#
                    arr_varLot(L_RICSH, lngE) = 0#
                    arr_varLot(L_RPCSH, lngE) = 0#
                    arr_varLot(L_RCOST, lngE) = 0#
                  ElseIf dblTmpShares < ![shareface] Then
                    ' ** Tax Lot covers it.
                    lngLotNum = lngLotNum + 1&
                    lngLots = lngLots + 1&
                    lngE = lngLots - 1&
                    ReDim Preserve arr_varLot(L_ELEMS, lngE)
                    arr_varLot(L_ACTNO, lngE) = strAccountNo
                    arr_varLot(L_ASTNO, lngE) = lngAssetNo
                    arr_varLot(L_TDATE, lngE) = datTransDate
                    arr_varLot(L_ADATE, lngE) = ![assetdate]
                    arr_varLot(L_LNUM, lngE) = lngLotNum
                    arr_varLot(L_CLOSD, lngE) = CBool(False)
                    arr_varLot(L_SHRS, lngE) = dblTmpShares
                    arr_varLot(L_ICSH, lngE) = dblTmpICash
                    arr_varLot(L_PCSH, lngE) = dblTmpPCash
                    arr_varLot(L_COST, lngE) = dblTmpCost
                    arr_varLot(L_PAAID, lngE) = lngPAAID
                    arr_varLot(L_RSHRS, lngE) = ![shareface] - dblTmpShares  ' ** Pos - Pos.
                    dblTmpShares = 0#
                    arr_varLot(L_RICSH, lngE) = ![ICash] + dblTmpICash  ' ** Neg + Pos
                    dblTmpICash = 0#
                    arr_varLot(L_RPCSH, lngE) = ![PCash] + dblTmpPCash  ' ** Neg + Pos
                    dblTmpPCash = 0#
                    arr_varLot(L_RCOST, lngE) = ![Cost] + dblTmpCost  ' ** Pos + Neg
                    dblTmpCost = 0#
                  Else
                    ' ** Partial coverage.
                    lngLotNum = lngLotNum + 1&
                    lngLots = lngLots + 1&
                    lngE = lngLots - 1&
                    ReDim Preserve arr_varLot(L_ELEMS, lngE)
                    arr_varLot(L_ACTNO, lngE) = strAccountNo
                    arr_varLot(L_ASTNO, lngE) = lngAssetNo
                    arr_varLot(L_TDATE, lngE) = datTransDate
                    arr_varLot(L_ADATE, lngE) = ![assetdate]
                    arr_varLot(L_LNUM, lngE) = lngLotNum
                    arr_varLot(L_CLOSD, lngE) = CBool(True)
                    arr_varLot(L_SHRS, lngE) = ![shareface]
                    dblTmpShares = dblTmpShares - ![shareface]
                    arr_varLot(L_ICSH, lngE) = ![ICash]
                    dblTmpICash = dblTmpICash + ![ICash]
                    arr_varLot(L_PCSH, lngE) = ![PCash]
                    dblTmpPCash = dblTmpPCash + ![PCash]
                    arr_varLot(L_COST, lngE) = ![Cost]
                    dblTmpCost = dblTmpCost + ![Cost]
                    arr_varLot(L_PAAID, lngE) = lngPAAID
                    arr_varLot(L_RSHRS, lngE) = 0#
                    arr_varLot(L_RICSH, lngE) = 0#
                    arr_varLot(L_RPCSH, lngE) = 0#
                    arr_varLot(L_RCOST, lngE) = 0#
                  End If
                  If dblTmpShares <= 0 Then
                    Exit For
                  Else
                    If lngY < lngRecs Then .MoveNext
                  End If
                Next  ' ** lngY.
              End If
              .Close
            End With
            Set rst2 = Nothing
            Set qdf2 = Nothing

          Case "Cost Adj."

          End Select

          If (![journaltype] <> "Withdrawn" And ![journaltype] <> "Sold") Then  'And ![journaltype] <> "Cost Adj.") Then

            Set rst2 = dbs.OpenRecordset("tblPennsville_Ledger", dbOpenDynaset, dbConsistent)
            With rst2
              .AddNew
              ' ** ![pl_id] : AutoNumber.
              ![pt3_id] = lngPT3ID
              ![journalno] = rst1![journalno]
              ![journaltype] = rst1![journaltype]
              ![assetno] = rst1![assetno]
              ![assettype] = Null
              ![transdate] = rst1![transdate]
              ![PostDate] = Null
              ![accountno] = rst1![accountno]
              ![shareface] = rst1![shareface]
              ![due] = rst1![due]
              ![rate] = rst1![rate]
              ![pershare] = rst1![pershare]
              ![ICash] = rst1![ICash]
              ![PCash] = rst1![PCash]
              ![Cost] = rst1![Cost]
              ![assetdate] = rst1![assetdate]
              If strDesc <> vbNullString Then
                ![description] = strDesc
              Else
                ![description] = Null
              End If
              ![posted] = Now()
              Select Case rst1![journaltype]
              Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received", "Cost Adj."
                ![taxcode] = TAXID_INC
              Case "Paid", "Liability"
                ![taxcode] = TAXID_DED
              Case "Misc."
                If rst1![ICash] > 0 Then
                  ![taxcode] = TAXID_INC
                ElseIf rst1![ICash] <= 0 Then
                  ![taxcode] = TAXID_DED
                End If
              End Select
              ![Location_ID] = CLng(1)
              Select Case rst1![journaltype]
              Case "Misc.", "Paid", "Received"
                If strRecur <> vbNullString Then
                  ![RecurringItem] = Left(strRecur, 50)
                Else
                  ![RecurringItem] = Null
                End If
              Case Else
                ![RecurringItem] = Null
              End Select
              Select Case rst1![journaltype]
              Case "Liability"
                Stop
                ![PurchaseDate] = Null
              Case Else
                ![PurchaseDate] = Null
              End Select
              ![ledger_HIDDEN] = False
              Select Case rst1![journaltype]
              Case "Dividend", "Interest", "Deposit", "Purchase", "Withdrawn", "Sold", "Received", "Cost Adj."
                ![revcode_ID] = REVID_INC
              Case "Paid", "Liability"
                ![revcode_ID] = REVID_EXP
              Case "Misc."
                If rst1![ICash] > 0 Then
                  ![revcode_ID] = REVID_INC
                ElseIf rst1![ICash] <= 0 Then
                  ![revcode_ID] = REVID_EXP
                End If
              End Select
              ![journal_USER] = "TAAdmin"
              ![CheckNum] = rst1![CheckNum]
              Select Case IsNull(rst1![CheckNum])
              Case True
                ![CheckPaid] = False
              Case False
                ![CheckPaid] = True
              End Select
              ![Transaction_Number] = rst1![Transaction_Number]
              ![Narrative] = rst1![Narrative]
              ![pl_lotnum] = 0&
              ![pt3_id_par] = 0&
              ![journalno_par] = 0&
              ![Transaction_Number_par] = 0&
              ![pl_datemodified] = Now()
              .Update
            End With  ' ** rst2.
            Set rst2 = Nothing

          Else
            ' ** 'Withdrawn', 'Sold', 'Cost Adj.'.

            If lngPT3ID_New = 0& Then
              lngPT3ID_New = DMax("[pt3_id]", "tblPennsville_Transaction3")
              lngJNum_New = DMax("[journalno]", "tblPennsville_Transaction3")
              lngTNum_New = DMax("[Transaction_Number]", "tblPennsville_Transaction3")
            End If

            Select Case ![journaltype]
            Case "Withdrawn", "Sold"
              Set rst2 = dbs.OpenRecordset("tblPennsville_Ledger", dbOpenDynaset, dbConsistent)
              With rst2
                For lngY = 0& To (lngLots - 1&)
                  .AddNew
                  ' ** ![pl_id] : AutoNumber.
                  If lngY = 0& Then
                    ![pt3_id] = lngPT3ID
                    ![journalno] = rst1![journalno]
                  Else
                    lngPT3ID_New = lngPT3ID_New + 1&
                    ![pt3_id] = lngPT3ID_New
                    lngJNum_New = lngJNum_New + 1&
                    ![journalno] = lngJNum_New
                  End If
                  ![journaltype] = rst1![journaltype]
                  ![assetno] = arr_varLot(L_ASTNO, lngY)
                  ![assettype] = Null
                  ![transdate] = arr_varLot(L_TDATE, lngY)
                  ![PostDate] = Null
                  ![accountno] = arr_varLot(L_ACTNO, lngY)
                  ![shareface] = arr_varLot(L_SHRS, lngY)
                  ![due] = rst1![due]
                  ![rate] = rst1![rate]
                  ![pershare] = rst1![pershare]
                  ![ICash] = arr_varLot(L_ICSH, lngY)
                  ![PCash] = arr_varLot(L_PCSH, lngY)
                  ![Cost] = arr_varLot(L_COST, lngY)
                  ![assetdate] = CDate(DateAdd("h", 9, arr_varLot(L_TDATE, lngY)))  ' ** 9:00 AM.
                  If strDesc <> vbNullString Then
                    ![description] = strDesc
                  Else
                    ![description] = Null
                  End If
                  ![posted] = Now()
                  ![taxcode] = TAXID_INC
                  ![Location_ID] = CLng(1)
                  ![RecurringItem] = Null
                  ![PurchaseDate] = arr_varLot(L_ADATE, lngY)
                  ![ledger_HIDDEN] = False
                  ![revcode_ID] = REVID_INC
                  ![journal_USER] = "TAAdmin"
                  ![CheckNum] = rst1![CheckNum]
                  ![CheckPaid] = False
                  If lngY = 0& Then
                    ![Transaction_Number] = rst1![Transaction_Number]
                  Else
                    lngTNum_New = lngTNum_New + 1&
                    ![Transaction_Number] = lngTNum_New
                  End If
                  ![Narrative] = rst1![Narrative]
                  ![pl_lotnum] = arr_varLot(L_LNUM, lngY)
                  ![pt3_id_par] = lngPT3ID
                  ![journalno_par] = rst1![journalno]
                  ![Transaction_Number_par] = rst1![Transaction_Number]
                  ![pl_datemodified] = Now()
                  .Update
                Next  ' ** lngY.
                .Close
              End With  ' ** rst2.
              Set rst2 = Nothing

              Set rst2 = dbs.OpenRecordset("tblPennsville_ActiveAssets", dbOpenDynaset, dbConsistent)
              With rst2
                .MoveFirst
                For lngY = 0& To (lngLots - 1&)
                  .FindFirst "[paa_id] = " & CStr(arr_varLot(L_PAAID, lngY))
                  Select Case .NoMatch
                  Case True
                    Stop
                  Case False
                    If arr_varLot(L_CLOSD, lngY) = True Then
                      .Edit
                      ![shareface] = 0#
                      ![ICash] = 0@
                      ![PCash] = 0@
                      ![Cost] = 0@
                      ![IsClosed] = True
                      ![paa_datemodified] = Now()
                      .Update
                    Else
                      .Edit
                      ![shareface] = arr_varLot(L_RSHRS, lngY)
                      ![ICash] = arr_varLot(L_RICSH, lngY)
                      ![PCash] = arr_varLot(L_RPCSH, lngY)
                      ![Cost] = arr_varLot(L_RCOST, lngY)
                      ![paa_datemodified] = Now()
                      .Update
                    End If
                  End Select
                Next  ' ** lngY.
                .Close
              End With  ' ** rst2.
              Set rst2 = Nothing

            Case "Cost Adj."

              blnSkip = True
              If blnSkip = False Then

                dblCost = ![Cost]  ' ** Can be neg or pos.

                ' ** zzz_qry_Pennsville_Trans_49_01 (tblPennsville_ActiveAssets, by specified
                ' ** [actno], [astno], [tdat]), grouped and summed, by accountno, assetno, with cnt.
                Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_49_02")
                With qdf2.Parameters
                  ![actno] = strAccountNo
                  ![astno] = lngAssetNo
                  ![tdat] = datTransDate
                End With
                Set rst2 = qdf2.OpenRecordset
                If rst2.BOF = True And rst2.EOF = True Then
                  Debug.Print "'ACTNO: " & strAccountNo & "  ASTNO: " & CStr(lngAssetNo) & "  " & _
                    "ADJ: " & CStr(dblTmp01) & "  COST: " & CStr(![Cost])
                  dblShares = 0#
                Else
                  rst2.MoveFirst
                  dblShares = rst2![shareface]
                End If
                rst2.Close
                Set rst2 = Nothing

                If dblShares > 0# Then

                  dblTmpCost = (dblCost / dblShares)  ' ** Pershare, neg or pos.

                  lngAdjs = 0
                  ReDim arr_varAdj(A_ELEMS, 0)

                  ' ** tblPennsville_ActiveAssets, by specified [actno], [astno], [tdat].
                  Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_49_01")
                  With qdf2.Parameters
                    ![actno] = strAccountNo
                    ![astno] = lngAssetNo
                    ![tdat] = datTransDate
                  End With
                  Set rst2 = qdf2.OpenRecordset
                  With rst2
                    If .BOF = True And .EOF = True Then
                      ' ** None available!
                      Stop
                    Else
                      .MoveLast
                      lngRecs = .RecordCount
                      .MoveFirst
                      For lngY = 1& To lngRecs
                        dblTmp01 = (dblTmpCost * ![shareface])
                        dblTmp01 = Round(dblTmp01, 2)
                        lngAdjs = lngAdjs + 1&
                        lngE = lngAdjs - 1&
                        ReDim Preserve arr_varAdj(A_ELEMS, lngE)
                        arr_varAdj(A_ACTNO, lngE) = strAccountNo
                        arr_varAdj(A_ASTNO, lngE) = lngAssetNo
                        arr_varAdj(A_TDATE, lngE) = datTransDate
                        arr_varAdj(A_ADATE, lngE) = ![assetdate]
                        arr_varAdj(A_COST, lngE) = dblTmp01
                        If dblTmp01 >= 0.01 Then
                          If dblTmp01 < ![Cost] Then
                            .Edit
                            ![Cost] = (![Cost] + dblTmp01)
                            ![averagepriceperunit] = (![Cost] / ![shareface])
                            ![priceperunit] = (![Cost] / ![shareface])
                            ![paa_datemodified] = Now()
                            .Update
                          Else
                            ' ** Adjustment will Zero Cost!
                            Beep
                            Debug.Print "'ACTNO: " & strAccountNo & "  ASTNO: " & CStr(lngAssetNo) & "  " & _
                              "ADJ: " & CStr(dblTmp01) & "  COST: " & CStr(![Cost])
                            Stop
                          End If
                        End If
                        If lngY < lngRecs Then .MoveNext
                      Next  ' ** lngY.
                    End If
                    .Close
                  End With  ' ** rst2.
                  Set rst2 = Nothing

                  If lngAdjs > 0& Then
                    Set rst2 = dbs.OpenRecordset("tblPennsville_Ledger", dbOpenDynaset, dbConsistent)
                    With rst2
                      For lngY = 0& To (lngAdjs - 1&)
                        .AddNew
                        ' ** ![pl_id] : AutoNumber.
                        If lngY = 0& Then
                          ![pt3_id] = lngPT3ID
                          ![journalno] = rst1![journalno]
                        Else
                          lngPT3ID_New = lngPT3ID_New + 1&
                          ![pt3_id] = lngPT3ID_New
                          lngJNum_New = lngJNum_New + 1&
                          ![journalno] = lngJNum_New
                        End If
                        ![journaltype] = rst1![journaltype]
                        ![assetno] = arr_varAdj(A_ASTNO, lngY)
                        ![assettype] = Null
                        ![transdate] = arr_varAdj(A_TDATE, lngY)
                        ![PostDate] = Null
                        ![accountno] = arr_varAdj(A_ACTNO, lngY)
                        ![shareface] = 0#
                        ![due] = Null
                        ![rate] = CDbl(0)
                        ![pershare] = CDbl(0)
                        ![ICash] = CCur(0)
                        ![PCash] = CCur(0)
                        ![Cost] = arr_varAdj(A_COST, lngY)
                        ![assetdate] = CDate(DateAdd("h", 9, arr_varAdj(A_TDATE, lngY)))  ' ** 9:00 AM.
                        If strDesc <> vbNullString Then
                          ![description] = strDesc
                        Else
                          ![description] = Null
                        End If
                        ![posted] = Now()
                        ![taxcode] = TAXID_INC
                        ![Location_ID] = CLng(1)
                        ![RecurringItem] = Null
                        ![PurchaseDate] = arr_varAdj(A_ADATE, lngY)
                        ![ledger_HIDDEN] = False
                        ![revcode_ID] = REVID_INC
                        ![journal_USER] = "TAAdmin"
                        ![CheckNum] = Null
                        ![CheckPaid] = False
                        If lngY = 0& Then
                          ![Transaction_Number] = rst1![Transaction_Number]
                        Else
                          lngTNum_New = lngTNum_New + 1&
                          ![Transaction_Number] = lngTNum_New
                        End If
                        ![Narrative] = rst1![Narrative]
                        ![pl_lotnum] = 0&
                        ![pt3_id_par] = lngPT3ID
                        ![journalno_par] = rst1![journalno]
                        ![Transaction_Number_par] = rst1![Transaction_Number]
                        ![pl_datemodified] = Now()
                        .Update
                      Next  ' ** lngY.
                      .Close
                    End With  ' ** rst2.
                    Set rst2 = Nothing
                  End If  ' ** lngAdjs.

                End If

              End If  ' ** blnSkip.

            End Select

          End If

          ' ** Update zz_tbl_Pennsville_Tmp18, for pt18_mark = True, by specified [tdat].
          Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_46_02")
          With qdf2.Parameters
            ![tdat] = datThisDate
          End With
          qdf2.Execute
          Set qdf2 = Nothing

          If lngX < lngRecs Then .MoveNext
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing
      Set qdf1 = Nothing

    Next  ' ** lngW.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Beep

  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  Set qdf1 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  PNB_TaxLots = blnRetVal

End Function

Public Function PNB_ResetTaxLots() As Boolean

  Const THIS_PROC As String = "PNB_ResetTaxLots"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim blnRetVal As Boolean

  blnRetVal = True

  DoCmd.DeleteObject acTable, "tblPennsville_ActiveAssets"
  DoEvents
  CurrentDb.TableDefs.Refresh
  CurrentDb.TableDefs.Refresh
  DoCmd.CopyObject , "tblPennsville_ActiveAssets", acTable, "tblPennsville_ActiveAssets_bak_new11"
  DoEvents
  CurrentDb.TableDefs.Refresh
  CurrentDb.TableDefs.Refresh

  Set dbs = CurrentDb
  With dbs

    ' ** Empty tblPennsville_Ledger.
    Set qdf = .QueryDefs("zzz_qry_Pennsville_Trans_46_04")
    qdf.Execute

    .Close
  End With


  Beep

  Set qdf = Nothing
  Set dbs = Nothing

  PNB_ResetTaxLots = blnRetVal

End Function

Public Function PNB_TaxLotQrys() As Boolean

  Const THIS_PROC As String = "PNB_TaxLotQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim strFrom As String, strTo As String, strDesc As String
  Dim strQryFrom As String, strQryTo As String
  Dim strSetFrom As String, strSetTo As String
  Dim strSQL As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  strFrom = "59"
  strTo = "60"
  strSetFrom = "8"
  strSetTo = "9"

  Set dbs = CurrentDb
  With dbs
    ' ** zzz_qry_Pennsville_Trans_54_03 - zzz_qry_Pennsville_Trans_54_21.
    For lngX = 3& To 21&
      strQryFrom = "zzz_qry_Pennsville_Trans_" & strFrom & "_" & Right("00" & CStr(lngX), 2)
      strQryTo = "zzz_qry_Pennsville_Trans_" & strTo & "_" & Right("00" & CStr(lngX), 2)
      DoCmd.CopyObject , strQryTo, acQuery, strQryFrom
      DoEvents
      .QueryDefs.Refresh
      .QueryDefs.Refresh
      Set qdf = .QueryDefs(strQryTo)
      With qdf
        strSQL = .SQL
        strSQL = StringReplace(strSQL, "Trans_" & strFrom & "_", "Trans_" & strTo & "_")  ' ** Module Functions: modStringFuncs.
        If InStr(strSQL, "CLng(" & strSetFrom & ") AS pt27_set") > 0 Then
          ' ** CLng(3) AS pt28_set
          strSQL = StringReplace(strSQL, "CLng(" & strSetFrom & ") AS pt27_set", "CLng(" & strSetTo & ") AS pt27_set")  ' ** Module Functions: modStringFuncs.
        ElseIf InStr(strSQL, "CLng(" & strSetFrom & ") AS pt28_set") > 0 Then
          ' ** CLng(3) AS pt28_set
          strSQL = StringReplace(strSQL, "CLng(" & strSetFrom & ") AS pt28_set", "CLng(" & strSetTo & ") AS pt28_set")  ' ** Module Functions: modStringFuncs.
        ElseIf InStr(strSQL, "CLng(" & strSetFrom & ") AS pt29_set") > 0 Then
          ' ** CLng(3) AS pt29_set
          strSQL = StringReplace(strSQL, "CLng(" & strSetFrom & ") AS pt29_set", "CLng(" & strSetTo & ") AS pt29_set")  ' ** Module Functions: modStringFuncs.
        End If
        If InStr(strSQL, "pt27_set)=" & strSetFrom) > 0 Then
          ' ** pt27_set)=6
          strSQL = StringReplace(strSQL, "pt27_set)=" & strSetFrom, "pt27_set)=" & strSetTo)  ' ** Module Functions: modStringFuncs.
        ElseIf InStr(strSQL, "pt28_set)=" & strSetFrom) > 0 Then
          ' ** pt28_set)=6
          strSQL = StringReplace(strSQL, "pt28_set)=" & strSetFrom, "pt28_set)=" & strSetTo)  ' ** Module Functions: modStringFuncs.
        ElseIf InStr(strSQL, "pt29_set)=" & strSetFrom) > 0 Then
          ' ** pt29_set)=6
          strSQL = StringReplace(strSQL, "pt29_set)=" & strSetFrom, "pt29_set)=" & strSetTo)  ' ** Module Functions: modStringFuncs.
        End If
        .SQL = strSQL
        strDesc = .Properties("Description")
        If InStr(strDesc, "_" & strFrom & "_") > 0 Then
          strDesc = StringReplace(strDesc, "_" & strFrom & "_", "_" & strTo & "_")  ' ** Module Functions: modStringFuncs.
          .Properties("Description") = strDesc
        End If
        DoEvents
        dbs.QueryDefs.Refresh
        strDesc = .Properties("Description")
        If InStr(strDesc, "Set " & strSetFrom) > 0 Then
          strDesc = StringReplace(strDesc, "Set " & strSetFrom, "Set " & strSetTo)  ' ** Module Functions: modStringFuncs.
          .Properties("Description") = strDesc
        End If
      End With
      DoEvents
      .QueryDefs.Refresh
      .QueryDefs.Refresh
    Next
    .Close
  End With

  Beep

  Set qdf = Nothing
  Set dbs = Nothing

  PNB_TaxLotQrys = blnRetVal

End Function

Public Function PNB_TaxLots2() As Boolean

  Const THIS_PROC As String = "PNB_TaxLots2"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, rst1 As DAO.Recordset
  Dim lngDates As Long, arr_varDate As Variant
  Dim lngShares As Long, arr_varShare() As Variant
  Dim lngLots As Long, arr_varLot() As Variant
  Dim lngRecs As Long
  Dim lngW As Long, lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varDate().
  Const D_ID    As Integer = 0
  Const D_TDATE As Integer = 1
  Const D_CNT   As Integer = 2
  Const D_MARK  As Integer = 3
  Const D_MDATE As Integer = 4

  ' ** Array: arr_varShare().
  Const S_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const S_ACTNO As Integer = 0
  Const S_ASTNO As Integer = 1
  Const S_TDATE As Integer = 2
  Const S_SHARS As Integer = 3
  Const S_ICASH As Integer = 4
  Const S_PCASH As Integer = 5
  Const S_COST  As Integer = 6
  Const S_CNT   As Integer = 7

  ' ** Array: arr_varLot().
  Const L_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const L_THISD As Integer = 0
  Const L_ACTNO As Integer = 1
  Const L_ASTNO As Integer = 2
  Const L_TDATE As Integer = 3
  Const L_SHARS As Integer = 4
  Const L_ICASH As Integer = 5
  Const L_PCASH As Integer = 6
  Const L_COST  As Integer = 7

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

'FOR EACH TRANS DAY, TOTAL ALL 'WITHDRAWN', 'SOLD',
'THEN TOTAL ALL TAX LOTS AVAILABLE ON THAT DAY, THOUGH
'THEY WON'T HAVE HAD PREVIOUS USE SUBTRACTED.

    ' ** Update zz_tbl_Pennsville_Tmp18, for pt18_mark = False, all.
    Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_46_03")
    qdf1.Execute
    Set qdf1 = Nothing

    Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp18", dbOpenDynaset, dbReadOnly)
    With rst1
      .MoveLast
      lngDates = .RecordCount
      .MoveFirst
      arr_varDate = .GetRows(lngDates)
      ' ******************************************************
      ' ** Array: arr_varDate()
      ' **
      ' **   Field  Element  Name                 Constant
      ' **   =====  =======  ===================  ==========
      ' **     1       0     pt18_id              D_ID
      ' **     2       1     pt18_transdate       D_TDATE
      ' **     3       2     pt18_count           D_CNT
      ' **     4       3     pt18_mark            D_MARK
      ' **     5       4     pt18_datemodified    D_MDATE
      ' **
      ' ******************************************************
      .Close
    End With
    Set rst1 = Nothing

    Debug.Print "'DATES: " & CStr(lngDates)
    DoEvents

    lngShares = 0&
    ReDim arr_varShare(S_ELEMS, 0)

    lngLots = 0&
    ReDim arr_varLot(L_ELEMS, 0)

    For lngW = 0& To (lngDates - 1&)

      ' ** tblPennsville_Transaction3, grouped and summed, by accountno,
      ' ** assetno, just 'Withdrawn', 'Sold', by  specified [tdat].
      Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_49_03")
      With qdf1.Parameters
        ![tdat] = arr_varDate(D_TDATE, lngW)
      End With
      Set rst1 = qdf1.OpenRecordset
      With rst1
        If .BOF = True And .EOF = True Then
          ' ** No 'Withdrawn', 'Sold' today.
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            lngShares = lngShares + 1&
            lngE = lngShares - 1&
            ReDim Preserve arr_varShare(S_ELEMS, lngE)
            arr_varShare(S_ACTNO, lngE) = ![accountno]
            arr_varShare(S_ASTNO, lngE) = ![assetno]
            arr_varShare(S_TDATE, lngE) = ![transdate]
            arr_varShare(S_SHARS, lngE) = ![shareface]
            arr_varShare(S_ICASH, lngE) = ![ICash]
            arr_varShare(S_PCASH, lngE) = ![PCash]
            arr_varShare(S_COST, lngE) = ![Cost]
            arr_varShare(S_CNT, lngE) = ![cnt]
            If lngX < lngRecs Then .MoveNext
          Next
        End If
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing
      Set qdf1 = Nothing

      ' ** zzz_qry_Pennsville_Trans_49_03 (tblPennsville_Transaction3, grouped and summed,
      ' ** by accountno, assetno, just 'Withdrawn', 'Sold', by  specified [tdat]), linked to
      ' ** tblPennsville_ActiveAssets, grouped and summed, by accountno, assetno, by specified [tdat].
      Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_49_04")
      With qdf1.Parameters
        ![tdat] = arr_varDate(D_TDATE, lngW)
      End With
      Set rst1 = qdf1.OpenRecordset
      With rst1
        If .BOF = True And .EOF = True Then
          ' ** None available.
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            lngLots = lngLots + 1&
            lngE = lngLots - 1&
            ReDim Preserve arr_varLot(L_ELEMS, lngE)
            arr_varLot(L_THISD, lngE) = arr_varDate(D_TDATE, lngW)
            arr_varLot(L_ACTNO, lngE) = ![accountno]
            arr_varLot(L_ASTNO, lngE) = ![assetno]
            arr_varLot(L_TDATE, lngE) = ![transdate]
            arr_varLot(L_SHARS, lngE) = ![shareface]
            arr_varLot(L_ICASH, lngE) = ![ICash]
            arr_varLot(L_PCASH, lngE) = ![PCash]
            arr_varLot(L_COST, lngE) = ![Cost]
            If lngX < lngRecs Then .MoveNext
          Next  ' ** lngX.
        End If
        .Close
      End With
      Set rst1 = Nothing
      Set qdf1 = Nothing

    Next  ' ** lngW.

    Debug.Print "'SHARES: " & CStr(lngShares)
    DoEvents

    Debug.Print "'LOTS: " & CStr(lngLots)
    DoEvents

    If lngShares > 0& Then
      Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp30", dbOpenDynaset, dbAppendOnly)
      With rst1
        For lngX = 0& To (lngShares - 1&)
          .AddNew
          ' ** ![pt30_id] : AutoNumber.
          ![accountno] = arr_varShare(S_ACTNO, lngX)
          ![assetno] = arr_varShare(S_ASTNO, lngX)
          ![transdate] = arr_varShare(S_TDATE, lngX)
          ![shareface] = arr_varShare(S_SHARS, lngX)
          ![ICash] = arr_varShare(S_ICASH, lngX)
          ![PCash] = arr_varShare(S_PCASH, lngX)
          ![Cost] = arr_varShare(S_COST, lngX)
          ![pt30_cnt] = arr_varShare(S_CNT, lngX)
          ![pt30_datemodified] = Now()
          .Update
        Next  ' ** lngX
        .Close
      End With
      Set rst1 = Nothing
    End If  ' ** lngShares.

    If lngLots > 0& Then
      Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp31", dbOpenDynaset, dbAppendOnly)
      With rst1
        For lngX = 0& To (lngLots - 1&)
          .AddNew
          ' ** ![pt31_id] : AutoNumber.
          ![pt31_thisdate] = arr_varLot(L_THISD, lngX)
          ![accountno] = arr_varLot(L_ACTNO, lngX)
          ![assetno] = arr_varLot(L_ASTNO, lngX)
          ![transdate] = arr_varLot(L_TDATE, lngX)
          ![shareface] = arr_varLot(L_SHARS, lngX)
          ![ICash] = arr_varLot(L_ICASH, lngX)
          ![PCash] = arr_varLot(L_PCASH, lngX)
          ![Cost] = arr_varLot(L_COST, lngX)
          ![pt31_datemodified] = Now()
          .Update
        Next  ' ** lngX
        .Close
      End With
      Set rst1 = Nothing
    End If  ' ** lngLots.

    .Close
  End With

  Debug.Print "'DONE!"

'DATES: 1711
'SHARES: 454
'LOTS: 995
'DONE!

  Beep

  Set rst1 = Nothing
  Set qdf1 = Nothing
  Set dbs = Nothing

  PNB_TaxLots2 = blnRetVal

End Function

Public Function PNB_TaxLots3() As Boolean

  Const THIS_PROC As String = "PNB_TaxLots3"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, rst1 As DAO.Recordset
  Dim lngShares As Long, arr_varShare As Variant
  Dim lngNewShares As Long, arr_varNewShare() As Variant
  Dim lngDateNegs As Long, arr_varDateNeg() As Variant
  Dim blnFound As Boolean
  Dim dblTmp01 As Double
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varShare().
  Const S_PTID   As Integer = 0
  Const S_ACTNO  As Integer = 1
  Const S_ASTNO  As Integer = 2
  Const S_TDATE  As Integer = 3
  Const S_SHARL  As Integer = 4
  Const S_SHARA  As Integer = 5
  Const S_ICASHL As Integer = 6
  Const S_ICASHA As Integer = 7
  Const S_PCASHL As Integer = 8
  Const S_PCASHA As Integer = 9
  Const S_COSTL  As Integer = 10
  Const S_COSTA  As Integer = 11

  ' ** Array: arr_varNewShare().
  Const N_ELEMS As Integer = 4  ' ** Array's first element UBound().
  Const N_ACTNO As Integer = 0
  Const N_ASTNO As Integer = 1
  Const N_SHARO As Integer = 2  ' original shares
  Const N_SHARR As Integer = 3  ' running total
  Const N_SHORT As Integer = 4  ' total short

  ' ** Array: arr_varDateNeg().
  Const D_ELEMS As Integer = 3  ' ** Array's first element UBound().
  Const D_ACTNO As Integer = 0
  Const D_ASTNO As Integer = 1
  Const D_TDATE As Integer = 2
  Const D_SHARS As Integer = 3

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

'I NEED RUNNING BALANCES!
'SO, .. WHAT?

    Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_64_01")
    Set rst1 = qdf1.OpenRecordset
    With rst1
      .MoveLast
      lngShares = .RecordCount
      .MoveFirst
      arr_varShare = .GetRows(lngShares)
      ' *************************************************
      ' ** Array: arr_varShare()
      ' **
      ' **   Field  Element  Name            Constant
      ' **   =====  =======  ==============  ==========
      ' **     1       0     pt30_id         S_PTID
      ' **     2       1     accountno       S_ACTNO
      ' **     3       2     assetno         S_ASTNO
      ' **     4       3     transdate       S_TDATE
      ' **     5       4     shareface       S_SHARL
      ' **     6       5     shareface_aa    S_SHARA
      ' **     7       6     icash           S_ICASHL
      ' **     8       7     icash_aa        S_ICASHA
      ' **     9       8     pcash           S_PCASHL
      ' **    10       9     pcash_aa        S_PCASHA
      ' **    11      10     cost            S_COSTL
      ' **    12      11     cost_aa         S_COSTA
      ' **
      ' *************************************************
      .Close
    End With
    Set rst1 = Nothing
    Set qdf1 = Nothing

    lngNewShares = 0&
    ReDim arr_varNewShare(N_ELEMS, 0)

    lngDateNegs = 0&
    ReDim arr_varDateNeg(D_ELEMS, 0)

    For lngX = 0& To (lngShares - 1&)
      blnFound = False
      For lngY = 0& To (lngNewShares - 1&)
        If arr_varNewShare(N_ACTNO, lngY) = arr_varShare(S_ACTNO, lngX) And arr_varNewShare(N_ASTNO, lngY) = arr_varShare(S_ASTNO, lngX) Then
          blnFound = True
          If Round(arr_varShare(S_SHARA, lngX), 4) > Round(arr_varNewShare(N_SHARO, lngY), 4) Then
            ' ** Additional deposits, purchases happened.
            arr_varNewShare(N_SHARR, lngY) = Round(arr_varNewShare(N_SHARR, lngY), 4) + (Round(arr_varShare(S_SHARA, lngX), 4) - Round(arr_varNewShare(N_SHARO, lngY), 4))
            arr_varNewShare(N_SHARO, lngY) = Round(arr_varShare(S_SHARA, lngX), 4)
          End If
          If Round(arr_varShare(S_SHARL, lngX), 4) <= Round(arr_varNewShare(N_SHARR, lngY), 4) Then
            arr_varNewShare(N_SHARR, lngY) = Round(arr_varNewShare(N_SHARR, lngY), 4) - Round(arr_varShare(S_SHARL, lngX), 4)
          Else
            dblTmp01 = Round(arr_varNewShare(N_SHARR, lngY), 4) - Round(arr_varShare(S_SHARL, lngX), 4)
            arr_varNewShare(N_SHORT, lngY) = Round(arr_varNewShare(N_SHORT, lngY) + dblTmp01, 4)
            arr_varNewShare(N_SHARR, lngY) = 0#
            lngDateNegs = lngDateNegs + 1&
            lngE = lngDateNegs - 1&
            ReDim Preserve arr_varDateNeg(D_ELEMS, lngE)
            arr_varDateNeg(D_ACTNO, lngE) = arr_varShare(S_ACTNO, lngX)
            arr_varDateNeg(D_ASTNO, lngE) = arr_varShare(S_ASTNO, lngX)
            arr_varDateNeg(D_TDATE, lngE) = arr_varShare(S_TDATE, lngX)
            arr_varDateNeg(D_SHARS, lngE) = Round(arr_varShare(S_SHARL, lngX), 4)
          End If
          Exit For
        End If
      Next  ' ** lngY.
      If blnFound = False Then
        lngNewShares = lngNewShares + 1&
        lngE = lngNewShares - 1&
        ReDim Preserve arr_varNewShare(N_ELEMS, lngE)
        arr_varNewShare(N_ACTNO, lngE) = arr_varShare(S_ACTNO, lngX)
        arr_varNewShare(N_ASTNO, lngE) = arr_varShare(S_ASTNO, lngX)
        arr_varNewShare(N_SHARO, lngE) = Round(arr_varShare(S_SHARA, lngX), 4)
        arr_varNewShare(N_SHARR, lngE) = Round(arr_varShare(S_SHARA, lngX), 4)
        arr_varNewShare(N_SHORT, lngE) = CDbl(0)
        If Round(arr_varShare(S_SHARL, lngX), 4) <= Round(arr_varNewShare(N_SHARR, lngE), 4) Then
          arr_varNewShare(N_SHARR, lngE) = Round(arr_varNewShare(N_SHARR, lngE), 4) - Round(arr_varShare(S_SHARL, lngX), 4)
        Else
          dblTmp01 = Round(arr_varNewShare(N_SHARR, lngE), 4) - Round(arr_varShare(S_SHARL, lngX), 4)
          arr_varNewShare(N_SHORT, lngE) = Round(arr_varNewShare(N_SHORT, lngE) + dblTmp01, 4)
          arr_varNewShare(N_SHARR, lngE) = 0#
          lngDateNegs = lngDateNegs + 1&
          lngE = lngDateNegs - 1&  ' ** lngE switches roles!
          ReDim Preserve arr_varDateNeg(D_ELEMS, lngE)
          arr_varDateNeg(D_ACTNO, lngE) = arr_varShare(S_ACTNO, lngX)
          arr_varDateNeg(D_ASTNO, lngE) = arr_varShare(S_ASTNO, lngX)
          arr_varDateNeg(D_TDATE, lngE) = arr_varShare(S_TDATE, lngX)
          arr_varDateNeg(D_SHARS, lngE) = Round(arr_varShare(S_SHARL, lngX), 4)
        End If
      End If
    Next  ' ** lngX

'I DON'T REALLY KNOW IF THIS'LL WORK!

    If lngNewShares > 0& Then
      ' ** arr_varNewShare().
      Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp32", dbOpenDynaset, dbAppendOnly)
      With rst1
        For lngX = 0& To (lngNewShares - 1&)
          .AddNew
          ' ** ![pt32_id] : AutoNumber.
          ![accountno] = arr_varNewShare(N_ACTNO, lngX)
          ![assetno] = arr_varNewShare(N_ASTNO, lngX)
          ![shareface_original] = Round(arr_varNewShare(N_SHARO, lngX), 4)
          ![shareface_running] = Round(arr_varNewShare(N_SHARR, lngX), 4)
          ![shareface_short] = Round(arr_varNewShare(N_SHORT, lngX), 4)
          ![pt32_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With
      Set rst1 = Nothing
    End If  ' ** lngNewShares.

    If lngDateNegs > 0& Then
      ' ** arr_varDateNeg().
      Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp33", dbOpenDynaset, dbAppendOnly)
      With rst1
        For lngX = 0& To (lngDateNegs - 1)
          .AddNew
          ' ** ![pt33_id] : AutoNumber.
          ![accountno] = arr_varDateNeg(D_ACTNO, lngX)
          ![assetno] = arr_varDateNeg(D_ASTNO, lngX)
          ![transdate] = arr_varDateNeg(D_TDATE, lngX)
          ![shareface] = Round(arr_varDateNeg(D_SHARS, lngX), 4)
          ![pt33_datemodified] = Now()
          .Update
        Next  ' ** lngDateNegs.
        .Close
      End With
      Set rst1 = Nothing
    End If  ' ** lngDateNegs.

    .Close
  End With

  Debug.Print "'DONE!"

  Beep

  Set rst1 = Nothing
  Set qdf1 = Nothing
  Set dbs = Nothing

  PNB_TaxLots3 = blnRetVal

End Function

Public Function PNB_TaxLots4() As Boolean

  Const THIS_PROC As String = "PNB_TaxLots4"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim lngAssetDates As Long, arr_varAssetDate() As Variant
  Dim lngNoneAvails As Long, arr_varNoneAvail() As Variant
  Dim strAccountNo As String, lngAssetNo As Long, datTransDate As Date
  Dim dblShareface As Double
  Dim lngRecs As Long, lngRecs2 As Long, lngLotNum As Long
  Dim dblTmpShares As Double, dblTmpICash As Double, dblTmpPCash As Double, dblTmpCost As Double
  Dim dblTmp01 As Double
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varAssetDate().
  Const A_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const A_ACTNO As Integer = 0
  Const A_ASTNO As Integer = 1
  Const A_TDATE As Integer = 2
  Const A_JNO   As Integer = 3
  Const A_ADATE As Integer = 4
  Const A_LNUM  As Integer = 5
  Const A_SHARS As Integer = 6

  ' ** Array: arr_varNoneAvail().
  Const N_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const N_ACTNO As Integer = 0
  Const N_ASTNO As Integer = 1
  Const N_TDATE As Integer = 2
  Const N_JNO   As Integer = 3
  Const N_SHORT As Integer = 4

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    lngNoneAvails = 0&
    ReDim arr_varNoneAvail(N_ELEMS, 0)

    Set qdf1 = .QueryDefs("zzz_qry_Pennsville_Trans_65_01")
    Set rst1 = qdf1.OpenRecordset
    With rst1
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs

        strAccountNo = ![accountno]
        lngAssetNo = ![assetno]
        datTransDate = ![transdate]
        dblShareface = ![shareface]
        dblTmpICash = ![ICash]
        dblTmpPCash = ![PCash]
        dblTmpCost = ![Cost]
        lngLotNum = 0&

        ' ** zzz_qry_Pennsville_Trans_65_03 (zzz_qry_Pennsville_Trans_65_02
        ' ** (tblPennsville_ActiveAssets), by specified [tdat]), by specified [actno], [astno].
        Set qdf2 = dbs.QueryDefs("zzz_qry_Pennsville_Trans_65_04")
        With qdf2.Parameters
          ![actno] = strAccountNo
          ![astno] = lngAssetNo
          ![tdat] = datTransDate
        End With
        Set rst2 = qdf2.OpenRecordset
        With rst2
          If .BOF = True And .EOF = True Then
            ' ** None available!
            lngNoneAvails = lngNoneAvails + 1&
            lngE = lngNoneAvails - 1&
            ReDim Preserve arr_varNoneAvail(N_ELEMS, lngE)
            arr_varNoneAvail(N_ACTNO, lngE) = strAccountNo
            arr_varNoneAvail(N_ASTNO, lngE) = lngAssetNo
            arr_varNoneAvail(N_TDATE, lngE) = datTransDate
            arr_varNoneAvail(N_JNO, lngE) = rst1![journalno]
            arr_varNoneAvail(N_SHORT, lngE) = dblShareface
          Else
            .MoveLast
            lngRecs2 = .RecordCount
            .MoveFirst
            dblTmpShares = dblShareface
            For lngY = 1& To lngRecs2
              lngLotNum = lngLotNum + 1&
              If Round(dblTmpShares, 4) <= Round(![shareface], 4) Then
                .Edit
                ![shareface] = Round(![shareface] - dblTmpShares, 4)
                dblTmp01 = dblTmpShares
                ![ICash] = ![ICash] + dblTmpICash
                ![PCash] = ![PCash] + dblTmpPCash
                ![Cost] = ![Cost] + dblTmpCost
                ![paa_datemodified] = Now()
                .Update
                lngAssetDates = lngAssetDates + 1&
                lngE = lngAssetDates - 1&
                ReDim Preserve arr_varAssetDate(A_ELEMS, lngE)
                arr_varAssetDate(A_ACTNO, lngE) = strAccountNo
                arr_varAssetDate(A_ASTNO, lngE) = lngAssetNo
                arr_varAssetDate(A_TDATE, lngE) = datTransDate
                arr_varAssetDate(A_JNO, lngE) = rst1![journalno]
                arr_varAssetDate(A_ADATE, lngE) = ![assetdate]
                arr_varAssetDate(A_LNUM, lngE) = lngLotNum
                arr_varAssetDate(A_SHARS, lngE) = dblTmp01
                dblTmpShares = 0#
                Exit For
              Else
                dblTmpShares = Round(dblTmpShares - ![shareface], 4)
                dblTmp01 = ![shareface]
                dblTmpICash = Round(dblTmpICash + ![ICash], 2)
                dblTmpPCash = Round(dblTmpPCash + ![PCash], 2)
                dblTmpCost = Round(dblTmpCost + ![Cost], 2)
                .Edit
                ![shareface] = 0#
                ![ICash] = 0@
                ![PCash] = 0@
                ![Cost] = 0@
                ![paa_datemodified] = Now()
                .Update
                lngAssetDates = lngAssetDates + 1&
                lngE = lngAssetDates - 1&
                ReDim Preserve arr_varAssetDate(A_ELEMS, lngE)
                arr_varAssetDate(A_ACTNO, lngE) = strAccountNo
                arr_varAssetDate(A_ASTNO, lngE) = lngAssetNo
                arr_varAssetDate(A_TDATE, lngE) = datTransDate
                arr_varAssetDate(A_JNO, lngE) = rst1![journalno]
                arr_varAssetDate(A_ADATE, lngE) = ![assetdate]
                arr_varAssetDate(A_LNUM, lngE) = lngLotNum
                arr_varAssetDate(A_SHARS, lngE) = dblTmp01
              End If
              If lngY < lngRecs2 Then .MoveNext
            Next  ' ** lngY.
            If Round(dblTmpShares, 4) > 0# Then
              lngNoneAvails = lngNoneAvails + 1&
              lngE = lngNoneAvails - 1&
              ReDim Preserve arr_varNoneAvail(N_ELEMS, lngE)
              arr_varNoneAvail(N_ACTNO, lngE) = strAccountNo
              arr_varNoneAvail(N_ASTNO, lngE) = lngAssetNo
              arr_varNoneAvail(N_TDATE, lngE) = datTransDate
              arr_varNoneAvail(N_JNO, lngE) = rst1![journalno]
              arr_varNoneAvail(N_SHORT, lngE) = dblTmpShares
            End If
          End If
          .Close
        End With  ' ** rst2.
        Set rst2 = Nothing
        Set qdf2 = Nothing
        If lngX < lngRecs Then .MoveNext
      Next  ' ** lngX.
      .Close
    End With  ' ** rst1.
    Set rst1 = Nothing
    Set qdf1 = Nothing

    If lngAssetDates > 0& Then
      ' ** arr_varAssetDate().
      Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp34", dbOpenDynaset, dbAppendOnly)
      With rst1
        For lngX = 0& To (lngAssetDates - 1&)
          .AddNew
          ' ** ![pt34_id] : AutoNumber.
          ![accountno] = arr_varAssetDate(A_ACTNO, lngX)
          ![assetno] = arr_varAssetDate(A_ASTNO, lngX)
          ![transdate] = arr_varAssetDate(A_TDATE, lngX)
          ![journalno] = arr_varAssetDate(A_JNO, lngX)
          ![assetdate] = arr_varAssetDate(A_ADATE, lngX)
          ![shareface] = arr_varAssetDate(A_SHARS, lngX)
          ![pt34_lotnum] = arr_varAssetDate(A_LNUM, lngX)
          ![pt34_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing
    End If  ' ** lngAssetDates

    If lngNoneAvails > 0& Then
      ' ** arr_varNoneAvail().
      Set rst1 = .OpenRecordset("zz_tbl_Pennsville_Tmp35", dbOpenDynaset, dbAppendOnly)
      With rst1
        For lngX = 0& To (lngNoneAvails - 1&)
          .AddNew
          ' ** ![pt35_id] : AutoNumber.
          ![accountno] = arr_varNoneAvail(N_ACTNO, lngX)
          ![assetno] = arr_varNoneAvail(N_ASTNO, lngX)
          ![transdate] = arr_varNoneAvail(N_TDATE, lngX)
          ![journalno] = arr_varNoneAvail(N_JNO, lngX)
          ![shareface_short] = arr_varNoneAvail(N_SHORT, lngX)
          ![pt35_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing
    End If  ' ** lngNoneAvails.

    .Close
  End With

  Debug.Print "'DONE!"

  Beep

  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf1 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  PNB_TaxLots4 = blnRetVal

End Function

Public Function IsAllUpperCase(varInput As Variant) As Boolean

  Const THIS_PROC As String = "IsAllUpperCase"

  Dim intLen As Integer
  Dim intX As Integer
  Dim blnRetVal As Boolean

  blnRetVal = False

  If IsNull(varInput) = False Then
    If Trim(varInput) <> vbNullString Then
      blnRetVal = True
      intLen = Len(varInput)
      For intX = 1 To intLen
        If Asc(Mid(varInput, intX, 1)) > 97 And Asc(Mid(varInput, intX, 1)) < 122 Then
          blnRetVal = False  ' ** Only 1 hit needed.
          Exit For
        End If
      Next
    End If
  End If

  IsAllUpperCase = blnRetVal

End Function

Public Function PNB_DivIntQyrs() As Boolean

  Const THIS_PROC As String = "PNB_DivIntQyrs"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strQryTmp01 As String, strSQLTmp01 As String
  Dim strDescTmp01 As String
  Dim lngNewQrys As Long, lngRecs As Long
  Dim intType As Integer, intPos1 As Integer
  Dim blnFound1 As Boolean
  Dim strTmp01 As String, strTmp02 As String
  Dim lngW As Long, lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const Q_NUM     As Integer = 0
  Const Q_Q05     As Integer = 1
  Const Q_Q06     As Integer = 2
  Const Q_Q07     As Integer = 3
  Const Q_Q08     As Integer = 4
  Const Q_Q09     As Integer = 5
  Const Q_Q10     As Integer = 6
  Const Q_Q11     As Integer = 7

  Const QSRC05 As String = "zzz_qry_xPennsville_Trans_27_02_05_03_05"
  Const QSRC06 As String = "zzz_qry_xPennsville_Trans_27_02_05_03_06"
  Const QSRC07 As String = "zzz_qry_xPennsville_Trans_27_02_05_03_07"
  Const QSRC08 As String = "zzz_qry_xPennsville_Trans_27_02_05_03_08"
  Const QSRC09 As String = "zzz_qry_xPennsville_Trans_27_02_05_03_09"
  Const QSRC10 As String = "zzz_qry_xPennsville_Trans_27_02_05_03_10"
  Const QSRC11 As String = "zzz_qry_xPennsville_Trans_27_02_05_03_11"

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  lngNewQrys = 0&

  Set dbs = CurrentDb
  With dbs
    For lngW = 6& To 34&
      ' ** zzz_qry_xPennsville_Trans_27_02_05_04_01 - zzz_qry_xPennsville_Trans_27_02_05_34_04.

      ' ** zzz_qry_xPennsville_Trans_27_02_05_07_04.
      strQryTmp01 = "zzz_qry_xPennsville_Trans_27_02_05_" & Right("00" & CStr(lngW), 2) & "_04"
      blnFound1 = False
      If QueryExists(strQryTmp01) = True Then
        blnFound1 = True
        Set qdf = .QueryDefs(strQryTmp01)
        ' ** .._27_02_05_07_02, linked to .._27_02_05_07_03, with CUSIP_new, Asset_Description_new_new; 18 (can't update .._Tmp51, .._Tmp53!).
        strDescTmp01 = qdf.Properties("Description")
        If InStr(strDescTmp01, "can't update") > 0 Then
          intType = 2
        Else
          intType = 1
        End If
      End If

      If blnFound1 = True Then

        lngQrys = lngQrys + 1&
        lngE = lngQrys - 1&
        ReDim Preserve arr_varQry(Q_ELEMS, lngE)
        arr_varQry(Q_NUM, lngE) = lngW
        strQryTmp01 = QSRC05
        strQryTmp01 = Left(strQryTmp01, 35) & Right("00" & CStr(lngW), 2) & Right(strQryTmp01, 3)  ' ** zzz_qry_xPennsville_Trans_27_02_05_
        arr_varQry(Q_Q05, lngE) = strQryTmp01
        strQryTmp01 = QSRC06
        strQryTmp01 = Left(strQryTmp01, 35) & Right("00" & CStr(lngW), 2) & Right(strQryTmp01, 3)
        arr_varQry(Q_Q06, lngE) = strQryTmp01
        strQryTmp01 = QSRC07
        strQryTmp01 = Left(strQryTmp01, 35) & Right("00" & CStr(lngW), 2) & Right(strQryTmp01, 3)
        arr_varQry(Q_Q07, lngE) = strQryTmp01
        strQryTmp01 = QSRC08
        strQryTmp01 = Left(strQryTmp01, 35) & Right("00" & CStr(lngW), 2) & Right(strQryTmp01, 3)
        arr_varQry(Q_Q08, lngE) = strQryTmp01
        strQryTmp01 = QSRC09
        strQryTmp01 = Left(strQryTmp01, 35) & Right("00" & CStr(lngW), 2) & Right(strQryTmp01, 3)
        arr_varQry(Q_Q09, lngE) = strQryTmp01
        strQryTmp01 = QSRC10
        strQryTmp01 = Left(strQryTmp01, 35) & Right("00" & CStr(lngW), 2) & Right(strQryTmp01, 3)
        arr_varQry(Q_Q10, lngE) = strQryTmp01
        strQryTmp01 = QSRC11
        strQryTmp01 = Left(strQryTmp01, 35) & Right("00" & CStr(lngW), 2) & Right(strQryTmp01, 3)
        arr_varQry(Q_Q11, lngE) = strQryTmp01

        Select Case intType
        Case 1
          DoCmd.CopyObject , arr_varQry(Q_Q05, lngE), acQuery, QSRC05
          lngNewQrys = lngNewQrys + 1&
          DoCmd.CopyObject , arr_varQry(Q_Q06, lngE), acQuery, QSRC06
          lngNewQrys = lngNewQrys + 1&
          DoCmd.CopyObject , arr_varQry(Q_Q07, lngE), acQuery, QSRC07
          lngNewQrys = lngNewQrys + 1&
          DoCmd.CopyObject , arr_varQry(Q_Q08, lngE), acQuery, QSRC08
          lngNewQrys = lngNewQrys + 1&
          DoCmd.CopyObject , arr_varQry(Q_Q09, lngE), acQuery, QSRC09
          lngNewQrys = lngNewQrys + 1&
          DoCmd.CopyObject , arr_varQry(Q_Q10, lngE), acQuery, QSRC10
          lngNewQrys = lngNewQrys + 1&
          DoCmd.CopyObject , arr_varQry(Q_Q11, lngE), acQuery, QSRC11
          lngNewQrys = lngNewQrys + 1&
          DoEvents
          .QueryDefs.Refresh
          .QueryDefs.Refresh
          DoEvents
        Case 2
          ' ** Skipped.
        End Select

        Select Case intType
        Case 1

          Set qdf = .QueryDefs(arr_varQry(Q_Q05, lngE))
          With qdf
            strSQLTmp01 = .SQL
            ' ** "zzz_qry_xPennsville_Trans" & "_27_02_05_" & "03" & "_04".
            strTmp01 = Mid(QSRC05, 26, 12) & "_04"                                ' ** Source SubQuery, DLookup() references.
            strTmp02 = Mid(QSRC05, 26, 10) & Right("00" & CStr(lngW), 2) & "_04"  ' ** New SubQuery, DLookup() references.
            strSQLTmp01 = StringReplace(strSQLTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .SQL = strSQLTmp01
            ' ** zz_tbl_Pennsville_Tmp53, with DLookups() to .._27_02_05_03_04; 5.
            strDescTmp01 = .Properties("Description")
            strDescTmp01 = StringReplace(strDescTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            lngRecs = 0&
            Set rst = .OpenRecordset
            With rst
              If .BOF = True And .EOF = True Then
                Stop
              Else
                .MoveLast
                lngRecs = .RecordCount
              End If
              .Close
            End With  ' ** rst.
            Set rst = Nothing
            intPos1 = InStr(strDescTmp01, ";")
            strDescTmp01 = Left(strDescTmp01, intPos1) & " " & CStr(lngRecs) & "."
            .Properties("Description") = strDescTmp01
          End With  ' ** qdf.
          Set qdf = Nothing
          .QueryDefs.Refresh

          Set qdf = .QueryDefs(arr_varQry(Q_Q06, lngE))
          With qdf
            strSQLTmp01 = .SQL
            strTmp01 = Mid(QSRC05, 26)                                                       ' ** Source Update reference.  PREVIOUS QUERY!
            strTmp02 = Mid(QSRC05, 26, 10) & Right("00" & CStr(lngW), 2) & Right(QSRC05, 3)  ' ** New Update reference.
            strSQLTmp01 = StringReplace(strSQLTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .SQL = strSQLTmp01
            ' ** Update .._27_02_05_03_05.
            strDescTmp01 = .Properties("Description")
            strDescTmp01 = StringReplace(strDescTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .Properties("Description") = strDescTmp01
          End With  ' ** qdf.
          .QueryDefs.Refresh
          qdf.Execute
          Set qdf = Nothing

          Set qdf = .QueryDefs(arr_varQry(Q_Q07, lngE))  ' ** No SQL changes.
          With qdf
            lngRecs = 0&
            Set rst = qdf.OpenRecordset
            With rst
              If .BOF = True And .EOF = True Then
                Stop
              Else
                .MoveLast
                lngRecs = .RecordCount
              End If
              .Close
            End With  ' ** rst.
            Set rst = Nothing
            strDescTmp01 = .Properties("Description")
            intPos1 = InStr(strDescTmp01, ";")
            strDescTmp01 = Left(strDescTmp01, intPos1) & " " & CStr(lngRecs) & "."
            .Properties("Description") = strDescTmp01
          End With  ' ** qdf.
          Set qdf = Nothing
          .QueryDefs.Refresh

          Set qdf = .QueryDefs(arr_varQry(Q_Q08, lngE))
          With qdf
            strSQLTmp01 = .SQL
            strTmp01 = Mid(QSRC07, 26)                                                       ' ** Source Update reference.  PREVIOUS QUERY!
            strTmp02 = Mid(QSRC07, 26, 10) & Right("00" & CStr(lngW), 2) & Right(QSRC07, 3)  ' ** New Update reference.
            strSQLTmp01 = StringReplace(strSQLTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .SQL = strSQLTmp01
            ' ** Update .._27_02_05_03_07.
            strDescTmp01 = .Properties("Description")
            strDescTmp01 = StringReplace(strDescTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .Properties("Description") = strDescTmp01
          End With  ' ** qdf.
          .QueryDefs.Refresh
          qdf.Execute
          Set qdf = Nothing

          Set qdf = .QueryDefs(arr_varQry(Q_Q09, lngE))
          With qdf
            strSQLTmp01 = .SQL
            strTmp01 = Mid(QSRC09, 26, 12) & "_04"                                ' ** Source reference 1.
            strTmp02 = Mid(QSRC09, 26, 10) & Right("00" & CStr(lngW), 2) & "_04"  ' ** New reference 1.
            strSQLTmp01 = StringReplace(strSQLTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            ' ** .._27_02_05_03_04, linked to .._27_02_05_03_02, with CUSIP_new, Asset_Description_new_new; 25.
            strDescTmp01 = .Properties("Description")
            strDescTmp01 = StringReplace(strDescTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            strTmp01 = Mid(QSRC09, 26, 12) & "_02"                                ' ** Source reference 2.
            strTmp02 = Mid(QSRC09, 26, 10) & Right("00" & CStr(lngW), 2) & "_02"  ' ** New reference 2.
            strSQLTmp01 = StringReplace(strSQLTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .SQL = strSQLTmp01
            strDescTmp01 = StringReplace(strDescTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            lngRecs = 0&
            Set rst = .OpenRecordset
            With rst
              If .BOF = True And .EOF = True Then
                Stop
              Else
                .MoveLast
                lngRecs = .RecordCount
              End If
              .Close
            End With  ' ** rst.
            Set rst = Nothing
            intPos1 = InStr(strDescTmp01, ";")
            strDescTmp01 = Left(strDescTmp01, intPos1) & " " & CStr(lngRecs) & "."
            .Properties("Description") = strDescTmp01
          End With  ' ** qdf.
          Set qdf = Nothing
          .QueryDefs.Refresh

          Set qdf = .QueryDefs(arr_varQry(Q_Q10, lngE))
          With qdf
            strSQLTmp01 = .SQL
            strTmp01 = Mid(QSRC09, 26)                                                       ' ** Source SubQuery, DLookup() references.  PREVIOUS QUERY!
            strTmp02 = Mid(QSRC09, 26, 10) & Right("00" & CStr(lngW), 2) & Right(QSRC09, 3)  ' ** New SubQuery, DLookup() references.
            strSQLTmp01 = StringReplace(strSQLTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .SQL = strSQLTmp01
            ' ** tblPennsville_Transaction_excel, with DLookups() to .._27_02_05_03_09; 25.
            strDescTmp01 = .Properties("Description")
            strDescTmp01 = StringReplace(strDescTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            lngRecs = 0&
            Set rst = .OpenRecordset
            With rst
              If .BOF = True And .EOF = True Then
                Stop
              Else
                .MoveLast
                lngRecs = .RecordCount
              End If
              .Close
            End With  ' ** rst.
            Set rst = Nothing
            intPos1 = InStr(strDescTmp01, ";")
            strDescTmp01 = Left(strDescTmp01, intPos1) & " " & CStr(lngRecs) & "."
            .Properties("Description") = strDescTmp01
          End With  ' ** qdf.
          Set qdf = Nothing
          .QueryDefs.Refresh

          Set qdf = .QueryDefs(arr_varQry(Q_Q11, lngE))
          With qdf
            strSQLTmp01 = .SQL
            strTmp01 = Mid(QSRC10, 26)                                                       ' ** Source Update reference.  PREVIOUS QUERY!
            strTmp02 = Mid(QSRC10, 26, 10) & Right("00" & CStr(lngW), 2) & Right(QSRC10, 3)  ' ** New Update reference.
            strSQLTmp01 = StringReplace(strSQLTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .SQL = strSQLTmp01
            ' ** Update .._27_02_05_03_10.
            strDescTmp01 = .Properties("Description")
            strDescTmp01 = StringReplace(strDescTmp01, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
            .Properties("Description") = strDescTmp01
          End With  ' ** qdf.
          .QueryDefs.Refresh
          qdf.Execute
          Set qdf = Nothing

        Case 2
          Debug.Print "'ALTERNATE SKIPPED: " & arr_varQry(Q_Q05, lngE)
          DoEvents
        End Select

      Else
        Debug.Print "'QRY SKIPPED: " & strQryTmp01
        DoEvents
      End If

    Next  ' ** lngW.

    .Close
  End With

  Debug.Print "'QRYS CREATED: " & CStr(lngNewQrys)
  Debug.Print "'DONE!"

'ALTERNATE SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_07_05
'ALTERNATE SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_14_05
'ALTERNATE SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_22_05
'ALTERNATE SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_31_05

'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_10_04
'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_13_04
'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_15_04
'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_17_04
'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_19_04
'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_21_04
'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_27_04
'QRY SKIPPED: zzz_qry_xPennsville_Trans_27_02_05_32_04

'QRYS CREATED: 133
'DONE!
  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  PNB_DivIntQyrs = blnRetVal

End Function

Public Function PNB_Outliers1() As Boolean

  Const THIS_PROC As String = "PNB_Outliers1"

  Dim blnRetVal As Boolean

  blnRetVal = True

'1:
'IIf([pte_id]=2752,2753,IIf([pte_id]=3913,3911,IIf([pte_id]=8937,8935,IIf([pte_id]=8960,8958,IIf([pte_id]=2000,1997,IIf([pte_id]=8973,8971,IIf([pte_id]=18576,18574,IIf([pte_id]=8186,8188,IIf([pte_id]=8999,8997,0)))))))))
'IIf([pte_id]=9002,9000,IIf([pte_id]=27,29,IIf([pte_id]=1303,1305,IIf([pte_id]=2336,2334,IIf([pte_id]=9034,9032,IIf([pte_id]=2018,2016,IIf([pte_id]=8215,8219,IIf([pte_id]=3338,3340,0))))))))
'IIf([pte_id]=8948,8944,IIf([pte_id]=5148,5146,IIf([pte_id]=5561,5555,IIf([pte_id]=4780,4781,IIf([pte_id]=5282,5283,0)))))
'2:
'IIf([pte_id]=8948,8946,IIf([pte_id]=5148,5144,IIf([pte_id]=5561,5557,IIf([pte_id]=4780,4783,IIf([pte_id]=5282,5285,0)))))
'3:
'IIf([pte_id]=5561,5559,IIf([pte_id]=4780,4785,IIf([pte_id]=5282,5287,0)))


'1:
'IIf([pte_id]=2752,2754,IIf([pte_id]=3913,3912,IIf([pte_id]=8937,8936,IIf([pte_id]=8960,8959,IIf([pte_id]=2000,1998,IIf([pte_id]=8973,8972,IIf([pte_id]=18576,18575,IIf([pte_id]=8186,8189,IIf([pte_id]=8999,8998,0)))))))))
'IIf([pte_id]=9002,9001,IIf([pte_id]=27,30,IIf([pte_id]=1303,1306,IIf([pte_id]=2336,2335,IIf([pte_id]=9034,9033,IIf([pte_id]=2018,2017,IIf([pte_id]=8215,8217,IIf([pte_id]=3338,3339,0))))))))
'IIf([pte_id]=8948,8945,IIf([pte_id]=5148,5145,IIf([pte_id]=5561,5556,IIf([pte_id]=4780,4782,IIf([pte_id]=5282,5284,0)))))


'2:
'IIf([pte_id]=8948,8947,IIf([pte_id]=5148,5147,IIf([pte_id]=5561,5558,IIf([pte_id]=4780,4784,IIf([pte_id]=5282,5286,0)))))
'3:
'IIf([pte_id]=5561,5560,IIf([pte_id]=4780,4786,IIf([pte_id]=5282,5288,0)))


'8236 - 8238
'8236 - 8237
'18220 - 18221
'18220 - 18222

'5282  - 5283,5285,5287
'  5282  - 5284,5286,5288

'4780  - 4781,4783,4785
'  4780  - 4782,4784,4786

'8948  - 8944,8946
'  8948  - 8945,8947

'5148  - 5146,5144
'  5148  - 5145,5147

'5561  - 5555,5557,5559
'  5561  - 5556,5558,5560

'2752  - 2753
'  2752  - 2754

'3913  - 3911
'  3913  - 3912

'8937  - 8935
'  8937  - 8936

'8960  - 8958
'  8960  - 8959

'2000  - 1997
'  2000  - 1998

'8973  - 8971
'  8973  - 8972

'18576 - 18574
'  18576 - 18575

'8186  - 8188
'  8186  - 8189

'8999  - 8997
'  8999  - 8998

'9002  - 9000
'  9002  - 9001

'27    - 29
'  27    - 30

'1303  - 1305
'  1303  - 1306

'2336  - 2334
'  2336  - 2335

'9034  - 9032
'  9034  - 9033

'2018  - 2016
'  2018  - 2017

'8215  - 8219
'  8215  - 8217

'3338  - 3340
'  3338  - 3339

'8236  -

'18220 -




  PNB_Outliers1 = blnRetVal

End Function

Public Function PNB_Qry_Import() As Boolean

  Const THIS_PROC As String = "PNB_Qry_Import"

  Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const Q_QNAM As Integer = 0

  blnRetVal = True

  strFile = "Trust_pennsville4.mdb"
  strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
  strPathFile = strPath & LNK_SEP & strFile

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)
  With wrk

    Set dbs = .OpenDatabase(strPathFile, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
    With dbs
      For Each qdf In .QueryDefs
        With qdf
          If InStr(.Name, "Pennsville") > 0 Then
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_QNAM, lngE) = .Name
          End If
        End With  ' ** qdf.
      Next  ' ** qdf.
      Set qdf = Nothing
      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    .Close
  End With  ' ** wrk.
  Set wrk = Nothing

  If lngQrys > 0& Then

    For lngX = 0& To (lngQrys - 1&)
      DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
    Next  ' ** lngX.

    CurrentDb.QueryDefs.Refresh

  Else
    Debug.Print "'NONE FOUND!"
  End If  ' ** lngQrys.

  Debug.Print "'DONE!"

  Beep

  Set qdf = Nothing
  Set dbs = Nothing
  Set wrk = Nothing

  PNB_Qry_Import = blnRetVal

End Function

Public Function PNB_Tbl_Delete() As Boolean

  Const THIS_PROC As String = "PNB_Tbl_Delete"

  Dim dbs As DAO.Database, tdf As DAO.TableDef
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const T_TNAM As Integer = 0

  blnRetVal = True

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    For Each tdf In .TableDefs
      With tdf
        If InStr(.Name, "Penns") > 0 Then
          lngTbls = lngTbls + 1&
          lngE = lngTbls - 1&
          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          arr_varTbl(T_TNAM, lngE) = .Name
        End If
      End With
    Next
    Set tdf = Nothing
    .Close
  End With
  Set dbs = Nothing

  Debug.Print "'TBLS TO DEL: " & CStr(lngTbls)
  DoEvents
  Stop

  For lngX = 0& To (lngTbls - 1&)
    Debug.Print "'" & arr_varTbl(T_TNAM, lngX)
    DoEvents
  Next

  If lngTbls > 0& Then
    For lngX = 0& To (lngTbls - 1&)
      DoCmd.DeleteObject acTable, arr_varTbl(T_TNAM, lngX)
      DoEvents
    Next
    CurrentDb.TableDefs.Refresh
  Else
    Debug.Print "'NONE FOUND!"
    DoEvents
  End If

'TBLS TO DEL: 85
'tblPennsvile_Account_Profile
'tblPennsville_Account_Profile_raw
'tblPennsville_ActiveAssets
'tblPennsville_Asset
'tblPennsville_Asset_Code
'tblPennsville_Asset_Code_Category1
'tblPennsville_Asset_Code_Category2
'tblPennsville_Asset_Code_Column
'tblPennsville_Asset_Code_raw
'tblPennsville_Asset_Holder
'tblPennsville_Asset_Holder_excel
'tblPennsville_Asset_Holder_excel_bak
'tblPennsville_AssetCusip_New
'tblPennsville_Contact_raw
'tblPennsville_Ledger
'tblPennsville_Ledger_tmp1
'tblPennsville_Map_Transaction
'tblPennsville_Transaction
'tblPennsville_Transaction_excel
'tblPennsville_Transaction_excel_bak
'tblPennsville_Transaction_excel_bak2
'tblPennsville_Transaction_Type
'tblPennsville_TransNum_New
'zz_tbl_Pennsville_Tmp01
'zz_tbl_Pennsville_Tmp02
'zz_tbl_Pennsville_Tmp03
'zz_tbl_Pennsville_Tmp04
'zz_tbl_Pennsville_Tmp05
'zz_tbl_Pennsville_Tmp06
'zz_tbl_Pennsville_Tmp07
'zz_tbl_Pennsville_Tmp08
'zz_tbl_Pennsville_Tmp09
'zz_tbl_Pennsville_Tmp10
'zz_tbl_Pennsville_Tmp11
'zz_tbl_Pennsville_Tmp12
'zz_tbl_Pennsville_Tmp13
'zz_tbl_Pennsville_Tmp14
'zz_tbl_Pennsville_Tmp15
'zz_tbl_Pennsville_Tmp16
'zz_tbl_Pennsville_Tmp17
'zz_tbl_Pennsville_Tmp18
'zz_tbl_Pennsville_Tmp19
'zz_tbl_Pennsville_Tmp20
'zz_tbl_Pennsville_Tmp21
'zz_tbl_Pennsville_Tmp22
'zz_tbl_Pennsville_Tmp23
'zz_tbl_Pennsville_Tmp24
'zz_tbl_Pennsville_Tmp25
'zz_tbl_Pennsville_Tmp26
'zz_tbl_Pennsville_Tmp27
'zz_tbl_Pennsville_Tmp28
'zz_tbl_Pennsville_Tmp29
'zz_tbl_Pennsville_Tmp30
'zz_tbl_Pennsville_Tmp31
'zz_tbl_Pennsville_Tmp32
'zz_tbl_Pennsville_Tmp33
'zz_tbl_Pennsville_Tmp34
'zz_tbl_Pennsville_Tmp35
'zz_tbl_Pennsville_Tmp40
'zz_tbl_Pennsville_Tmp41
'zz_tbl_Pennsville_Tmp42
'zz_tbl_Pennsville_Tmp43
'zz_tbl_Pennsville_Tmp44
'zz_tbl_Pennsville_Tmp45
'zz_tbl_Pennsville_Tmp46
'zz_tbl_Pennsville_Tmp47
'zz_tbl_Pennsville_Tmp48
'zz_tbl_Pennsville_Tmp49
'zz_tbl_Pennsville_Tmp50
'zz_tbl_Pennsville_Tmp51
'zz_tbl_Pennsville_Tmp52
'zz_tbl_Pennsville_Tmp53
'zz_tbl_Pennsville_Tmp54
'zz_tbl_Pennsville_Tmp55
'zz_tbl_Pennsville_Tmp56
'zz_tbl_Pennsville_Tmp57
'zz_tbl_Pennsville_Tmp58
'zz_tbl_Pennsville_Tmp59
'zz_tbl_Pennsville_Tmp60
'zz_tbl_Pennsville_Tmp61
'zz_tbl_Pennsville_Tmp62
'zz_tbl_Pennsville_xlsAccount_Profile_Extract
'zz_tbl_Pennsville_xlsAsset_Holder_Extract
'zz_tbl_Pennsville_xlsContact_Extract
'zz_tbl_Pennsville_xlsTransaction_Extract
'DONE!

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set tdf = Nothing
  Set dbs = Nothing

  PNB_Tbl_Delete = blnRetVal

End Function

Public Function PNB_Rpt_Delete() As Boolean

  Const THIS_PROC As String = "PNB_Rpt_Delete"

  Dim dbs As DAO.Database, cntr As DAO.Container, doc As DAO.Document
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const R_RNAM As Integer = 0

  blnRetVal = True

  lngRpts = 0&
  ReDim arr_varRpt(R_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    Set cntr = .Containers("Reports")
    With cntr
      For Each doc In .Documents
        With doc
          If InStr(.Name, "Pennsville") > 0 Then
            lngRpts = lngRpts + 1&
            lngE = lngRpts - 1&
            ReDim Preserve arr_varRpt(R_ELEMS, lngE)
            arr_varRpt(R_RNAM, lngE) = .Name
          End If
        End With
      Next
      Set doc = Nothing
    End With
    Set cntr = Nothing
    .Close
  End With
  Set dbs = Nothing

  Debug.Print "'RPTS TO DEL: " & CStr(lngRpts)
  DoEvents
  Stop

  For lngX = 0& To (lngRpts - 1&)
    Debug.Print "'" & arr_varRpt(R_RNAM, lngX)
    DoEvents
  Next

  If lngRpts > 0& Then
    For lngX = 0& To (lngRpts - 1&)
      DoCmd.DeleteObject acReport, arr_varRpt(R_RNAM, lngX)
      DoEvents
    Next
    CurrentDb.Containers("Reports").Documents.Refresh
  Else
    Debug.Print "'NONE FOUND!"
    DoEvents
  End If

'RPTS TO DEL: 10
'rptPennsville_AssetCodes
'rptPennsville_AssetCusip_New_SortCusip
'rptPennsville_AssetCusip_New_SortDesc
'rptPennsville_Assets_DDI
'rptPennsville_Assets_SortCusip
'rptPennsville_Assets_SortDesc
'rptPennsville_Assets_Statement
'rptPennsville_TransactionCodes
'rptPennsville_Transactions_Statement
'rptPennsville_Transactions_Statement_TA
'DONE!

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set doc = Nothing
  Set cntr = Nothing
  Set dbs = Nothing

  PNB_Rpt_Delete = blnRetVal

End Function
