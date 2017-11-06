Attribute VB_Name = "zz_mod_NegativeTaxLots"
Option Compare Database
Option Explicit

'VGC 05/09/2015: CHANGES!

'INCOME O/U:
'ANY WAY OF APPORTIONING IT NOW?
'zzz_qry_MasterTrust_40

Private Const THIS_NAME As String = "zz_mod_NegativeTaxLots"
' **

Public Function NetTaxLots_SetTmp() As Boolean

  Const THIS_PROC As String = "NetTaxLots_SetTmp"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    For lngX = 1& To 3&

      ' ** Update zzz_qry_MasterTrust_10_214_001_09_01r (tmpXAdmin_ActiveAssets_03,
      ' ** just 1st Tax Lot having sf2_before = Null, with DLookups() to
      ' ** zzz_qry_MasterTrust_10_214_001_09_01q (tmpXAdmin_ActiveAssets_03,
      ' ** just 1st Tax Lot having sf2_before <> Null, sf2_after = Null)).
      Set qdf = .QueryDefs("zzz_qry_MasterTrust_10_214_001_09_01s")
      qdf.Execute
      DoEvents
      Set qdf = Nothing

      ForcePause 1  ' ** Module Function: modCodeUtilities.

      ' ** Update zzz_qry_MasterTrust_10_214_001_09_01q (tmpXAdmin_ActiveAssets_03,
      ' ** just 1st Tax Lot having sf2_before <> Null, sf2_after = Null).
      Set qdf = .QueryDefs("zzz_qry_MasterTrust_10_214_001_09_01t")
      qdf.Execute
      DoEvents
      Set qdf = Nothing

      ForcePause 1  ' ** Module Function: modCodeUtilities.

      Beep

    Next  ' ** lngX.
    .Close
  End With  ' ** dbs.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  DoBeeps  ' ** Module Function: modWindowFunctions.

  Set qdf = Nothing
  Set dbs = Nothing

  NetTaxLots_SetTmp = blnRetVal

End Function

Public Function NegTaxLots_Qrys() As Boolean

  Const THIS_PROC As String = "NegTaxLots_Qrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, prp As DAO.Property
  Dim lngAccts As Long, arr_varAcct As Variant
  Dim lngNegs As Long, arr_varNeg As Variant
  Dim lngMaxActNoWidth As Long, lngMaxAstNoWidth As Long
  Dim strQry As String, strFirstQry As String, strLastQry As String, strNewQryBase As String
  Dim strDesc As String, strPartialDesc As String, strSQL As String
  Dim lngRecs As Long, lngNewQrys As Long, lngNegsThisAcct As Long
  Dim lngJournalNo As Long, datTransDate As Date, strAssetDate As String
  Dim dblShareNeg As Double, dblNegBal As Double
  Dim lngX As Long, lngY As Long, lngZ As Long
  Dim blnRetVal As Boolean

  Const strAcctNoQry As String = "zzz_qry_MasterTrust_06"
  Const strNegTaxLotQry As String = "zzz_qry_MasterTrust_05"
  Const strQryBase1 As String = "zzz_qry_MasterTrust_"
  Const strQryBase2 As String = "zzz_qry_MasterTrust_10_"

  ' ** Array: arr_varAcct().
  Const A_ACTNO As Integer = 0  'accountno
  Const A_ASTNO As Integer = 1  'assetno
  Const A_CNT   As Integer = 2  'cnt

  ' ** Array: arr_varNeg().
  Const N_ACTNO As Integer = 0  'accountno
  Const N_ASTNO As Integer = 1  'assetno
  Const N_TDAT  As Integer = 2  'transdate
  Const N_ADAT  As Integer = 3  'assetdate
  Const N_SHARE As Integer = 4  'shareface
  Const N_ICASH As Integer = 5  'icash
  Const N_PCASH As Integer = 6  'pcash
  Const N_COST  As Integer = 7  'cost

  blnRetVal = True

'WHAT ABOUT LEDGER ARCHIVE!!

  DoCmd.Hourglass True
  DoEvents

  Set dbs = CurrentDb
  With dbs

    Set qdf = .QueryDefs(strAcctNoQry)
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngAccts = .RecordCount
      .MoveFirst
      arr_varAcct = .GetRows(lngAccts)
      ' **********************************************
      ' ** Array: arr_varAcct()
      ' **
      ' **   Field  Element  Name         Constant
      ' **   =====  =======  ===========  ==========
      ' **     1       0     accountno    A_ACTNO
      ' **     2       1     assetno      A_ASTNO
      ' **     3       2     cnt          A_CNT
      ' **
      ' **********************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    ' ** These could come from the array!
    lngMaxActNoWidth = 3&
    lngMaxAstNoWidth = 3&
    lngNewQrys = 0&

    For lngX = 0& To (lngAccts - 1&)
 
     ' ** Ledger, just accountno = '00077', assetno = 1; 307.
      strDesc = "Ledger, just accountno = '" & arr_varAcct(A_ACTNO, lngX) & ", assetno = " & CStr(arr_varAcct(A_ASTNO, lngX)) & ";"

      ' ** zzz_qry_MasterTrust_10_077_001_01
      strQry = strQryBase2 & Right$(String(lngMaxActNoWidth, "0") & arr_varAcct(A_ACTNO, lngX), lngMaxActNoWidth)
      strQry = strQry & "_" & Right$(String(lngMaxAstNoWidth, "0") & arr_varAcct(A_ASTNO, lngX), lngMaxAstNoWidth)
      strQry = strQry & "_"
      strNewQryBase = strQry
      strPartialDesc = Mid$(strQry, Len(strQryBase1))
      strQry = strQry & "01"
      strFirstQry = strQry
      lngNegsThisAcct = 0&

      If QueryExists(strQry) = False Then  ' ** Module Function: modFileUtilities.

        strSQL = "SELECT ledger.journalno, ledger.journaltype, ledger.accountno, ledger.assetno, ledger.transdate, "
        strSQL = strSQL & "IIf(IsNull([PurchaseDate])=True,IIf(IsNull([assetdate])=True,Null,[assetdate]),[PurchaseDate]) AS assetdatex, "
        strSQL = strSQL & "ledger.shareface, ledger.icash, ledger.pcash, ledger.cost, ledger.description, ledger.assetdate, ledger.PurchaseDate "
        strSQL = strSQL & "FROM ledger "
        strSQL = strSQL & "WHERE (((ledger.journaltype) Not In ('Interest','Dividend','Received')) AND "
        strSQL = strSQL & "((ledger.accountno)='" & arr_varAcct(A_ACTNO, lngX) & "') AND "
        strSQL = strSQL & "((ledger.assetno)=" & CStr(arr_varAcct(A_ASTNO, lngX)) & ")) "
        strSQL = strSQL & "ORDER BY ledger.journalno;"
        Set qdf = .CreateQueryDef(strQry, strSQL)
On Error Resume Next
        qdf.Properties("Description") = strDesc
        If ERR.Number <> 0 Then
On Error GoTo 0
          Set prp = qdf.CreateProperty("Description", dbText, strDesc)
          qdf.Properties.Append prp
        Else
On Error GoTo 0
        End If
        .QueryDefs.Refresh
        .QueryDefs.Refresh
        DoEvents
        Set qdf = Nothing
        strSQL = vbNullString
        lngNewQrys = lngNewQrys + 1&

        ' ** Format the date fields.
        Qry_CheckBox strQry, "assetdatex", False, 2  ' ** Module Function: modQueryFunctions.
        DoEvents
        Qry_CheckBox strQry, "assetdate", False, 2  ' ** Module Function: modQueryFunctions.
        DoEvents
        Qry_CheckBox strQry, "PurchaseDate", False, 2  ' ** Module Function: modQueryFunctions.
        DoEvents
        Qry_CheckBox strQry, "transdate", False, 1  ' ** Module Function: modQueryFunctions.
        DoEvents

        ' ** Add the transaction count to the Description.
        Set qdf = .QueryDefs(strQry)
        Set rst = qdf.OpenRecordset
        With rst
          .MoveLast
          lngRecs = .RecordCount
          .Close
        End With
        qdf.Properties("Description") = strDesc & " " & CStr(lngRecs) & "."
        Set rst = Nothing
        Set qdf = Nothing

        .QueryDefs.Refresh
        .QueryDefs.Refresh
        DoEvents

      End If  ' ** QueryExists().

      For lngY = 2& To 8&

        Select Case lngY
        Case 2&
          ' ** zzz_qry_MasterTrust_10_077_001_02
          strDesc = "    .." & strPartialDesc & "01" & ", with .._dp/.._ws broken out;"
          strQry = strNewQryBase & "02"
        Case 3&
          ' ** zzz_qry_MasterTrust_10_077_001_03
          strDesc = "    .." & strPartialDesc & "02, grouped and summed, by accountno, assetno, assetdatex (Tax Lot);"
          strLastQry = strQry
          strQry = strNewQryBase & "03"
        Case 4&
          ' ** zzz_qry_MasterTrust_10_077_001_04
          strDesc = "    .." & strPartialDesc & "03, with totals;"
          strLastQry = strQry
          strQry = strNewQryBase & "04"
        Case 5&
          ' ** zzz_qry_MasterTrust_10_077_001_05
          strDesc = "    .." & strPartialDesc & "04, just shareface <> 0;"
          strLastQry = strQry
          strQry = strNewQryBase & "05"
        Case 6&
          ' ** zzz_qry_MasterTrust_10_077_001_06
          strDesc = "    .." & strPartialDesc & "05, linked to ActiveAssets, just discrepancies;"
          strLastQry = strQry  ' ** 05.
          strQry = strNewQryBase & "06"
        Case 7&
          ' ** zzz_qry_MasterTrust_10_077_001_07
          strDesc = "    .." & strPartialDesc & "05, grouped and summed, by accountno, assetno (w/o Neg);"
          ' ** strLastQry = strQry : Still 05.
          strQry = strNewQryBase & "07"
        Case 8&
          ' ** zzz_qry_MasterTrust_10_077_001_08
          strDesc = "    .._05, negative tax lots for accountno = '" & arr_varAcct(A_ACTNO, lngX) & "';"
          ' ** strLastQry = strQry : Still 05.
          strQry = strNewQryBase & "08"
        End Select

        If QueryExists(strQry) = False Then  ' ** Module Function: modFileUtilities.

          Select Case lngY
          Case 2&
            strSQL = "SELECT " & strFirstQry & ".journalno, " & strFirstQry & ".journaltype, " & strFirstQry & ".accountno, " & _
              strFirstQry & ".assetno, " & strFirstQry & ".transdate, " & strFirstQry & ".assetdatex, " & _
              "CDbl(IIf([journaltype] In ('Deposit','Purchase'),[shareface],IIf([journaltype]='Liability'," & _
              "IIf(IsNull([PurchaseDate])=True,[shareface],0),0))) AS shareface_dp, "
            strSQL = strSQL & "CDbl(IIf([journaltype] In ('Withdrawn','Sold'),[shareface],IIf([journaltype]='Liability'," & _
              "IIf(IsNull([PurchaseDate])=True,0,[shareface]),0))) AS shareface_ws, "
            strSQL = strSQL & "CCur(IIf([journaltype] In ('Deposit','Purchase','Dividend','Interest','Received'),[icash]," & _
              "IIf([journaltype]='Liability',IIf(IsNull([PurchaseDate])=True,[icash],0),IIf([journaltype]='Misc.'," & _
              "IIf([icash]+[pcash]>=0,[icash],0),0)))) AS icash_dp, "
            strSQL = strSQL & "CCur(IIf([journaltype] In ('Withdrawn','Sold','Paid'),[icash],IIf([journaltype]='Liability'," & _
              "IIf(IsNull([PurchaseDate])=True,0,[icash]),IIf([journaltype]='Misc.',IIf([icash]+[pcash]>=0,0,[icash]),0)))) AS icash_ws, "
            strSQL = strSQL & "CCur(IIf([journaltype] In ('Deposit','Purchase','Dividend','Interest','Received'),[pcash]," & _
              "IIf([journaltype]='Liability',IIf(IsNull([PurchaseDate])=True,[pcash],0),IIf([journaltype]='Misc.'," & _
              "IIf([icash]+[pcash]>=0,[pcash],0),0)))) AS pcash_dp, " & _
              "CCur(IIf([journaltype] In ('Withdrawn','Sold','Paid'),[pcash],IIf([journaltype]='Liability'," & _
              "IIf(IsNull([PurchaseDate])=True,0,[pcash]),IIf([journaltype]='Misc.',IIf([icash]+[pcash]>=0,0,[pcash]),0)))) AS pcash_ws, "
            strSQL = strSQL & "CCur(IIf([journaltype] In ('Deposit','Purchase','Dividend','Interest','Received'),[cost]," & _
              "IIf([journaltype]='Liability',IIf(IsNull([PurchaseDate])=True,[cost],0),IIf([journaltype]='Misc.'," & _
              "IIf([icash]+[pcash]>=0,[cost],0),IIf([journaltype]='Cost Adj.',IIf([cost]>0,[cost],0),0))))) AS cost_dp, "
            strSQL = strSQL & "CCur(IIf([journaltype] In ('Withdrawn','Sold','Paid'),[pcash],IIf([journaltype]='Liability'," & _
              "IIf(IsNull([PurchaseDate])=True,0,[pcash]),IIf([journaltype]='Misc.',IIf([icash]+[pcash]>=0,0,[pcash])," & _
              "IIf([journaltype]='Cost Adj.',IIf([cost]>0,0,[cost]),0))))) AS cost_ws, " & _
              strFirstQry & ".description, " & strFirstQry & ".assetdate, " & strFirstQry & ".PurchaseDate "
            strSQL = strSQL & "FROM " & strFirstQry & ";"
          Case 3&
            strSQL = "SELECT " & strLastQry & ".accountno, " & strLastQry & ".assetno, " & strLastQry & ".assetdatex, " & _
              "Sum(" & strLastQry & ".shareface_dp) AS shareface_dp, Sum(" & strLastQry & ".shareface_ws) AS shareface_ws, " & _
              "Sum(" & strLastQry & ".icash_dp) AS icash_dp, Sum(" & strLastQry & ".icash_ws) AS icash_ws, " & _
              "Sum(" & strLastQry & ".pcash_dp) AS pcash_dp, Sum(" & strLastQry & ".pcash_ws) AS pcash_ws, " & _
              "Sum(" & strLastQry & ".cost_dp) AS cost_dp, Sum(" & strLastQry & ".cost_ws) AS cost_ws "
            strSQL = strSQL & "FROM " & strLastQry & " "
            strSQL = strSQL & "GROUP BY " & strLastQry & ".accountno, " & strLastQry & ".assetno, " & strLastQry & ".assetdatex;"
          Case 4&
            strSQL = "SELECT " & strLastQry & ".accountno, " & strLastQry & ".assetno, " & strLastQry & ".assetdatex, " & _
              "CDbl(Round((Round(CDbl([shareface_dp]),4)-Round(CDbl([shareface_ws]),4)),4)) AS shareface_tot, " & _
              "CCur(Round((Round(CDbl([icash_dp]),2)+Round(CDbl([icash_ws]),2)),2)) AS icash_tot, " & _
              "CCur(Round((Round(CDbl([pcash_dp]),2)+Round(CDbl([pcash_ws]),2)),2)) AS pcash_tot, " & _
              "CCur(Round((Round(CDbl([cost_dp]),2)+Round(CDbl([cost_ws]),2)),2)) AS cost_tot "
            strSQL = strSQL & "FROM " & strLastQry & ";"
          Case 5&
            strSQL = "SELECT " & strLastQry & ".accountno, " & strLastQry & ".assetno, " & strLastQry & ".assetdatex, " & _
              strLastQry & ".shareface_tot, " & strLastQry & ".icash_tot, " & strLastQry & ".pcash_tot, " & strLastQry & ".cost_tot "
            strSQL = strSQL & "FROM " & strLastQry & " "
            strSQL = strSQL & "WHERE (((" & strLastQry & ".shareface_tot)<>0));"
          Case 6&
            strSQL = "SELECT " & strLastQry & ".accountno, " & strLastQry & ".assetno, " & strLastQry & ".assetdatex, " & _
              strLastQry & ".shareface_tot AS shareface_l, ActiveAssets.shareface AS shareface_aa, " & _
              "IIf(Round([shareface_tot],4)<>Round([shareface],4),'X',IIf([shareface]<0,'X','')) AS Sx, " & _
              strLastQry & ".icash_tot AS icash_l, ActiveAssets.icash AS icash_aa, " & strLastQry & ".pcash_tot AS pcash_l, " & _
              "ActiveAssets.pcash AS pcash_aa, " & strLastQry & ".cost_tot AS cost_l, ActiveAssets.cost AS cost_aa, " & _
              "IIf(Round(CDbl([cost_tot]),2)<>Round(CDbl([cost]),2),'X','') AS Cx "
            strSQL = strSQL & "FROM " & strLastQry & " LEFT JOIN ActiveAssets ON (" & strLastQry & ".accountno = ActiveAssets.accountno) AND " & _
              "(" & strLastQry & ".assetno = ActiveAssets.assetno) AND (" & strLastQry & ".assetdatex = ActiveAssets.assetdate) "
            strSQL = strSQL & "WHERE (((IIf(Round([shareface_tot],4)<>Round([shareface],4),'X',IIf([shareface]<0,'X','')))='X')) OR " & _
              "(((IIf(Round(CDbl([cost_tot]),2)<>Round(CDbl([cost]),2),'X',''))='X'));"
          Case 7&
            strSQL = "SELECT " & strLastQry & ".accountno, " & strLastQry & ".assetno, " & _
              "Sum(" & strLastQry & ".shareface_tot) AS shareface_tot, Sum(" & strLastQry & ".icash_tot) AS icash_tot, " & _
              "Sum(" & strLastQry & ".pcash_tot) AS pcash_tot, Sum(" & strLastQry & ".cost_tot) AS cost_tot "
            strSQL = strSQL & "FROM " & strLastQry & " "
            strSQL = strSQL & "WHERE (((" & strLastQry & ".shareface_tot) > 0)) "
            strSQL = strSQL & "GROUP BY " & strLastQry & ".accountno, " & strLastQry & ".assetno;"
          Case 8&
            strSQL = "SELECT " & strNegTaxLotQry & ".accountno, " & strNegTaxLotQry & ".assetno, " & strNegTaxLotQry & ".transdate, " & _
              strNegTaxLotQry & ".assetdate, " & strNegTaxLotQry & ".shareface, " & strNegTaxLotQry & ".icash, " & _
              strNegTaxLotQry & ".pcash, " & strNegTaxLotQry & ".cost "
            strSQL = strSQL & "FROM " & strNegTaxLotQry & " "
            strSQL = strSQL & "WHERE (((" & strNegTaxLotQry & ".accountno)='" & arr_varAcct(A_ACTNO, lngX) & "'));"
          End Select

          Set qdf = .CreateQueryDef(strQry, strSQL)
On Error Resume Next
          qdf.Properties("Description") = strDesc
          If ERR.Number <> 0 Then
On Error GoTo 0
            Set prp = qdf.CreateProperty("Description", dbText, strDesc)
            qdf.Properties.Append prp
          Else
On Error GoTo 0
          End If
          .QueryDefs.Refresh
          .QueryDefs.Refresh
          DoEvents
          Set qdf = Nothing
          strSQL = vbNullString
          lngNewQrys = lngNewQrys + 1&

          ' ** Add the transaction count to the Description.
          Set qdf = .QueryDefs(strQry)
          Set rst = qdf.OpenRecordset
          With rst
            If .BOF = True And .EOF = True Then
              lngRecs = 0&
            Else
              .MoveLast
              lngRecs = .RecordCount
              If lngY = 8& Then
                lngNegsThisAcct = lngRecs
              End If
            End If
            .Close
          End With
          qdf.Properties("Description") = strDesc & " " & CStr(lngRecs) & "."
          Set rst = Nothing
          Set qdf = Nothing

          .QueryDefs.Refresh
          .QueryDefs.Refresh
          DoEvents

        End If  ' ** QueryExists().

      Next  ' ** lngY.

      If lngNegsThisAcct > 0& Then

        strLastQry = strQry  ' ** 08.
        Set qdf = .QueryDefs(strLastQry)
        Set rst = qdf.OpenRecordset
        With rst
          .MoveLast
          lngNegs = .RecordCount
          .MoveFirst
          arr_varNeg = .GetRows(lngNegs)
          ' **********************************************
          ' ** Array: arr_varNeg()
          ' **
          ' **   Field  Element  Name         Constant
          ' **   =====  =======  ===========  ==========
          ' **     1       0     accountno    N_ACTNO
          ' **     2       1     assetno      N_ASTNO
          ' **     3       2     transdate    N_TDAT
          ' **     4       3     assetdate    N_ADAT
          ' **     5       4     shareface    N_SHARE
          ' **     6       5     icash        N_ICASH
          ' **     7       6     pcash        N_PCASH
          ' **     8       7     cost         N_COST
          ' **
          ' **********************************************
          .Close
        End With
        Set rst = Nothing
        Set qdf = Nothing

        strNewQryBase = strNewQryBase & "09_"

        For lngY = 0& To (lngNegs - 1&)

          ' ** zzz_qry_MasterTrust_10_077_001_09_01
          strAssetDate = Format(arr_varNeg(N_ADAT, lngY), "mm/dd/yyyy hh:nn:ss")
          strDesc = "        .." & strPartialDesc & "01, just assetdate = " & strAssetDate & ";"  ' 45, journalno >= 50, transdate = 12/31/2009."
          strQry = strNewQryBase & Right$("00" & CStr(lngY + 1&), 2)

          strSQL = "SELECT " & strFirstQry & ".journalno, " & strFirstQry & ".journaltype, " & strFirstQry & ".accountno, " & _
            strFirstQry & ".assetno, " & strFirstQry & ".transdate, " & strFirstQry & ".assetdatex, " & _
            strFirstQry & ".shareface, " & strFirstQry & ".icash, " & strFirstQry & ".pcash, " & strFirstQry & ".cost, " & _
            strFirstQry & ".description, " & strFirstQry & ".assetdate, " & strFirstQry & ".PurchaseDate "
          strSQL = strSQL & "FROM " & strFirstQry & " "
          strSQL = strSQL & "WHERE (((" & strFirstQry & ".assetdate)=#" & strAssetDate & "#)) OR " & _
            "(((" & strFirstQry & ".PurchaseDate)=#" & strAssetDate & "#));"

          Set qdf = .CreateQueryDef(strQry, strSQL)
On Error Resume Next
          qdf.Properties("Description") = strDesc
          If ERR.Number <> 0 Then
On Error GoTo 0
            Set prp = qdf.CreateProperty("Description", dbText, strDesc)
            qdf.Properties.Append prp
          Else
On Error GoTo 0
          End If
          .QueryDefs.Refresh
          .QueryDefs.Refresh
          DoEvents
          Set qdf = Nothing
          strSQL = vbNullString
          lngNewQrys = lngNewQrys + 1&

          ' ** Get the negative amount.
          dblShareNeg = arr_varNeg(N_SHARE, lngY)
          dblNegBal = dblShareNeg

'zzz_qry_MasterTrust_10_077_001_09_04
'  FLIP JOURNALNO'S 67012 WITH 67167
'zzz_qry_MasterTrust_10_081_001_09_06
'  FLIP JOURNALNO'S 67017 WITH 67069
'zzz_qry_MasterTrust_10_214_001_09_01
'  FLIP JOURNALNO'S 68638 WITH 68698
'  FLIP JOURNALNO'S 68696 WITH 68700
'DO WE NEED TO CHECK THIS IN OTHER TAX LOTS?

          ' ** Now start at the end and work backwards to find the
          ' ** first transaction that put the Tax Lot negative.
          Set qdf = .QueryDefs(strQry)
          Set rst = qdf.OpenRecordset
          With rst
            If .BOF = True And .EOF = True Then
              lngRecs = 0&
            Else
              .MoveLast
              lngRecs = .RecordCount
              For lngZ = lngRecs To 1& Step -1&
                If ![journaltype] = "Sold" Or ![journaltype] = "Withdrawn" Then
                  dblNegBal = dblNegBal + ![shareface]
                  If dblNegBal >= 0# Then
                    lngJournalNo = ![journalno]
                    datTransDate = ![transdate]
                    If dblNegBal > 0.1 Then
                      Debug.Print "'CHK JNO ORD!  " & strQry
                    End If
                    Exit For
                  End If
                End If
              Next
            End If
            .Close
          End With

          ' ** Add the transaction count, journalno, and transdate to the Description.
          ' ** 45, journalno >= 50, transdate = 12/31/2009."
          strDesc = strDesc & " " & CStr(lngRecs) & ", journalno >= " & CStr(lngJournalNo) & ", " & _
            "transdate = " & Format(datTransDate, "mm/dd/yyyy") & "."
          qdf.Properties("Description") = strDesc
          Set rst = Nothing
          Set qdf = Nothing

          .QueryDefs.Refresh
          .QueryDefs.Refresh
          DoEvents

        Next  ' ** lngY.

      End If  ' ** lngNegsThisAcct.

    Next  ' ** lngX.

    .Close
  End With  ' ** dbs.

  DoCmd.Hourglass False

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'NEW QRYS CREATED: " & CStr(lngNewQrys)
  Debug.Print "'DONE!  " & THIS_PROC & "()"
'CHK JNO ORD!  zzz_qry_MasterTrust_10_081_001_09_06
'CHK JNO ORD!  zzz_qry_MasterTrust_10_214_001_09_01
'NEW QRYS CREATED: 90
'DONE!  NegTaxLots_Qrys()

  DoBeeps  ' ** Module Function: modWindowFunctions.

  Set prp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NegTaxLots_Qrys = blnRetVal

End Function

Public Function NegTaxLots_Fix() As Boolean

  Const THIS_PROC As String = "NegTaxLots_Fix"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strAccountNo As String, lngAssetNo As Long
  Dim strDesc As String
  Dim strQry As String, strLastQry As String, strNewQryBase As String
  Dim strAssetDate As String, datAssetDate As Date, strTransDate As String, datTransDate As Date
  Dim strJournalNo As String, lngJournalNo As Long
  Dim lngAccts As Long, arr_varAcct As Variant
  Dim lngNegs As Long, arr_varNeg As Variant
  Dim lngTrans As Long, arr_varTran() As Variant
  Dim lngAssets As Long, arr_varAsset() As Variant
  Dim lngMaxActNoWidth As Long, lngMaxAstNoWidth As Long
  Dim dblShareFaceNeg As Double, dblShares As Double, dblShares2 As Double, dblPerShare As Double
  Dim dblICash As Double, dblPCash As Double, dblCost As Double
  Dim dblICashSum As Double, dblPCashSum As Double, dblCostSum As Double
  Dim lngLedgerUpdates As Long, lngLedgerAdds As Long, lngTaxLotUpdates As Long, lngTaxLotDels As Long
  Dim lngRecs As Long
  Dim blnFound As Boolean
  Dim intPos1 As Integer
  Dim strTmp00 As String, dblTmp01 As Double, dblTmp02 As Double, lngTmp03 As Long
  Dim lngV As Long, lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  Const strAcctNoQry As String = "zzz_qry_MasterTrust_06"
  Const strNegTaxLotQry As String = "zzz_qry_MasterTrust_05"
  Const strQryBase1 As String = "zzz_qry_MasterTrust_"
  Const strQryBase2 As String = "zzz_qry_MasterTrust_10_"

  ' ** Array: arr_varAcct().
  Const A_ACTNO As Integer = 0  'accountno
  Const A_ASTNO As Integer = 1  'assetno
  Const A_CNT   As Integer = 2  'cnt

  ' ** Array: arr_varNeg().
  Const N_ACTNO As Integer = 0  'accountno
  Const N_ASTNO As Integer = 1  'assetno
  Const N_TDAT  As Integer = 2  'transdate
  Const N_ADAT  As Integer = 3  'assetdate
  Const N_SHARE As Integer = 4  'shareface
  Const N_ICASH As Integer = 5  'icash
  Const N_PCASH As Integer = 6  'pcash
  Const N_COST  As Integer = 7  'cost

  ' ** Array: arr_varTran().
  Const T_ELEMS As Integer = 190  ' ** Array's first-element UBound().
  Const T_JNO    As Integer = 0
  Const T_JTYP   As Integer = 1
  Const T_ACTNO  As Integer = 2
  Const T_ASTNO  As Integer = 3
  Const T_TDAT   As Integer = 4
  Const T_ADAT   As Integer = 5
  Const T_PDAT   As Integer = 6
  Const T_SHR    As Integer = 7
  Const T_ICASH  As Integer = 8
  Const T_PCASH  As Integer = 9
  Const T_COST   As Integer = 10
  Const T_DESC   As Integer = 11
  Const T_PDAT02 As Integer = 12
  'Const T_PDAT03 As Integer = 13
  'Const T_PDAT04 As Integer = 14
  'Const T_PDAT05 As Integer = 15
  'Const T_PDAT06 As Integer = 16
  'Const T_PDAT07 As Integer = 17
  'Const T_PDAT08 As Integer = 18
  'Const T_PDAT09 As Integer = 19
  'Const T_PDAT10 As Integer = 20
  'Const T_PDAT11 As Integer = 21
  'Const T_PDAT12 As Integer = 22
  'Const T_PDAT13 As Integer = 23
  'Const T_PDAT14 As Integer = 24
  'Const T_PDAT15 As Integer = 25
  'Const T_PDAT16 As Integer = 26
  'Const T_PDAT17 As Integer = 27
  'Const T_PDAT18 As Integer = 28
  'Const T_PDAT19 As Integer = 29
  'Const T_PDAT20 As Integer = 30
  'Const T_PDAT21 As Integer = 31
  'Const T_PDAT22 As Integer = 32
  'Const T_PDAT23 As Integer = 33
  'Const T_PDAT24 As Integer = 34
  'Const T_PDAT25 As Integer = 35
  'Const T_PDAT26 As Integer = 36
  'Const T_PDAT27 As Integer = 37
  'Const T_PDAT28 As Integer = 38
  'Const T_PDAT29 As Integer = 39
  'Const T_PDAT30 As Integer = 40
  'Const T_PDAT31 As Integer = 41
  'Const T_PDAT32 As Integer = 42
  'Const T_PDAT33 As Integer = 43
  'Const T_PDAT34 As Integer = 44
  'Const T_PDAT35 As Integer = 45
  'Const T_PDAT36 As Integer = 46
  'Const T_PDAT37 As Integer = 47
  'Const T_PDAT38 As Integer = 48
  'Const T_PDAT39 As Integer = 49
  'Const T_PDAT40 As Integer = 50
  'Const T_PDAT41 As Integer = 51
  'Const T_PDAT42 As Integer = 52
  'Const T_PDAT43 As Integer = 53
  'Const T_PDAT44 As Integer = 54
  'Const T_PDAT45 As Integer = 55
  'Const T_PDAT46 As Integer = 56
  'Const T_PDAT47 As Integer = 57
  'Const T_PDAT48 As Integer = 58
  'Const T_PDAT49 As Integer = 59
  'Const T_PDAT50 As Integer = 60
  'Const T_PDAT51 As Integer = 61
  'Const T_PDAT52 As Integer = 62
  'Const T_PDAT53 As Integer = 63
  'Const T_PDAT54 As Integer = 64
  'Const T_PDAT55 As Integer = 65
  'Const T_PDAT56 As Integer = 66
  'Const T_PDAT57 As Integer = 67
  'Const T_PDAT58 As Integer = 68
  'Const T_PDAT59 As Integer = 69
  'Const T_PDAT60 As Integer = 70
  'Const T_PDAT61 As Integer = 71
  'Const T_PDAT62 As Integer = 72
  'Const T_PDAT63 As Integer = 73
  'Const T_PDAT64 As Integer = 74
  'Const T_PDAT65 As Integer = 75
  'Const T_PDAT66 As Integer = 76
  'Const T_PDAT67 As Integer = 77
  'Const T_PDAT68 As Integer = 78
  'Const T_PDAT69 As Integer = 79
  'Const T_PDAT70 As Integer = 80
  'Const T_PDAT71 As Integer = 81
  'Const T_PDAT72 As Integer = 82
  'Const T_PDAT73 As Integer = 83
  'Const T_PDAT74 As Integer = 84
  'Const T_PDAT75 As Integer = 85
  'Const T_PDAT76 As Integer = 86
  'Const T_PDAT77 As Integer = 87
  'Const T_PDAT78 As Integer = 88
  'Const T_PDAT79 As Integer = 89
  'Const T_PDAT80 As Integer = 90
  'Const T_PDAT81 As Integer = 91
  'Const T_PDAT82 As Integer = 92
  'Const T_PDAT83 As Integer = 93
  'Const T_PDAT84 As Integer = 94
  'Const T_PDAT85 As Integer = 95
  'Const T_PDAT86 As Integer = 96
  'Const T_PDAT87 As Integer = 97
  'Const T_PDAT88 As Integer = 98
  'Const T_PDAT89 As Integer = 99
  'Const T_PDAT90 As Integer = 100
  Const T_SHR02  As Integer = 101
  'Const T_SHR03  As Integer = 102
  'Const T_SHR04  As Integer = 103
  'Const T_SHR05  As Integer = 104
  'Const T_SHR06  As Integer = 105
  'Const T_SHR07  As Integer = 106
  'Const T_SHR08  As Integer = 107
  'Const T_SHR09  As Integer = 108
  'Const T_SHR10  As Integer = 109
  'Const T_SHR11  As Integer = 110
  'Const T_SHR12  As Integer = 111
  'Const T_SHR13  As Integer = 112
  'Const T_SHR14  As Integer = 113
  'Const T_SHR15  As Integer = 114
  'Const T_SHR16  As Integer = 115
  'Const T_SHR17  As Integer = 116
  'Const T_SHR18  As Integer = 117
  'Const T_SHR19  As Integer = 118
  'Const T_SHR20  As Integer = 119
  'Const T_SHR21  As Integer = 120
  'Const T_SHR22  As Integer = 121
  'Const T_SHR23  As Integer = 122
  'Const T_SHR24  As Integer = 123
  'Const T_SHR25  As Integer = 124
  'Const T_SHR26  As Integer = 125
  'Const T_SHR27  As Integer = 126
  'Const T_SHR28  As Integer = 127
  'Const T_SHR29  As Integer = 128
  'Const T_SHR30  As Integer = 129
  'Const T_SHR31  As Integer = 130
  'Const T_SHR32  As Integer = 131
  'Const T_SHR33  As Integer = 132
  'Const T_SHR34  As Integer = 133
  'Const T_SHR35  As Integer = 134
  'Const T_SHR36  As Integer = 135
  'Const T_SHR37  As Integer = 136
  'Const T_SHR38  As Integer = 137
  'Const T_SHR39  As Integer = 138
  'Const T_SHR40  As Integer = 139
  'Const T_SHR41  As Integer = 140
  'Const T_SHR42  As Integer = 141
  'Const T_SHR43  As Integer = 142
  'Const T_SHR44  As Integer = 143
  'Const T_SHR45  As Integer = 144
  'Const T_SHR46  As Integer = 145
  'Const T_SHR47  As Integer = 146
  'Const T_SHR48  As Integer = 147
  'Const T_SHR49  As Integer = 148
  'Const T_SHR50  As Integer = 149
  'Const T_SHR51  As Integer = 150
  'Const T_SHR52  As Integer = 151
  'Const T_SHR53  As Integer = 152
  'Const T_SHR54  As Integer = 153
  'Const T_SHR55  As Integer = 154
  'Const T_SHR56  As Integer = 155
  'Const T_SHR57  As Integer = 156
  'Const T_SHR58  As Integer = 157
  'Const T_SHR59  As Integer = 158
  'Const T_SHR60  As Integer = 159
  'Const T_SHR61  As Integer = 160
  'Const T_SHR62  As Integer = 161
  'Const T_SHR63  As Integer = 162
  'Const T_SHR64  As Integer = 163
  'Const T_SHR65  As Integer = 164
  'Const T_SHR66  As Integer = 165
  'Const T_SHR67  As Integer = 166
  'Const T_SHR68  As Integer = 167
  'Const T_SHR69  As Integer = 168
  'Const T_SHR70  As Integer = 169
  'Const T_SHR71  As Integer = 170
  'Const T_SHR72  As Integer = 171
  'Const T_SHR73  As Integer = 172
  'Const T_SHR74  As Integer = 173
  'Const T_SHR75  As Integer = 174
  'Const T_SHR76  As Integer = 175
  'Const T_SHR77  As Integer = 176
  'Const T_SHR78  As Integer = 177
  'Const T_SHR79  As Integer = 178
  'Const T_SHR80  As Integer = 179
  'Const T_SHR81  As Integer = 180
  'Const T_SHR82  As Integer = 181
  'Const T_SHR83  As Integer = 182
  'Const T_SHR84  As Integer = 183
  'Const T_SHR85  As Integer = 184
  'Const T_SHR86  As Integer = 185
  'Const T_SHR87  As Integer = 186
  'Const T_SHR88  As Integer = 187
  'Const T_SHR89  As Integer = 188
  'Const T_SHR90  As Integer = 189
  Const T_CNT    As Integer = 190

  Const lngTrans_PDAT_Beg As Long = 12&
  Const lngTrans_PDAT_End As Long = 100&
  Const lngTrans_SHR_Beg As Long = 101&
  Const lngTrans_SHR_End As Long = 189&

  ' ** Array: arr_varAsset().
  Const S_ELEMS As Integer = 187  ' ** Array's first-element UBound().
  Const S_ACTNO As Integer = 0
  Const S_ASTNO As Integer = 1
  Const S_TDAT  As Integer = 2
  Const S_ADAT  As Integer = 3
  Const S_SHR   As Integer = 4
  Const S_ICASH As Integer = 5
  Const S_PCASH As Integer = 6
  Const S_COST  As Integer = 7
  Const S_SHRN  As Integer = 8  'shareface_new.
  Const S_JNO02 As Integer = 9
  'Const S_JNO03 As Integer = 10
  'Const S_JNO04 As Integer = 11
  'Const S_JNO05 As Integer = 12
  'Const S_JNO06 As Integer = 13
  'Const S_JNO07 As Integer = 14
  'Const S_JNO08 As Integer = 15
  'Const S_JNO09 As Integer = 16
  'Const S_JNO10 As Integer = 17
  'Const S_JNO11 As Integer = 18
  'Const S_JNO12 As Integer = 19
  'Const S_JNO13 As Integer = 20
  'Const S_JNO14 As Integer = 21
  'Const S_JNO15 As Integer = 22
  'Const S_JNO16 As Integer = 23
  'Const S_JNO17 As Integer = 24
  'Const S_JNO18 As Integer = 25
  'Const S_JNO19 As Integer = 26
  'Const S_JNO20 As Integer = 27
  'Const S_JNO21 As Integer = 28
  'Const S_JNO22 As Integer = 29
  'Const S_JNO23 As Integer = 30
  'Const S_JNO24 As Integer = 31
  'Const S_JNO25 As Integer = 32
  'Const S_JNO26 As Integer = 33
  'Const S_JNO27 As Integer = 34
  'Const S_JNO28 As Integer = 35
  'Const S_JNO29 As Integer = 36
  'Const S_JNO30 As Integer = 37
  'Const S_JNO31 As Integer = 38
  'Const S_JNO32 As Integer = 39
  'Const S_JNO33 As Integer = 40
  'Const S_JNO34 As Integer = 41
  'Const S_JNO35 As Integer = 42
  'Const S_JNO36 As Integer = 43
  'Const S_JNO37 As Integer = 44
  'Const S_JNO38 As Integer = 45
  'Const S_JNO39 As Integer = 46
  'Const S_JNO40 As Integer = 47
  'Const S_JNO41 As Integer = 48
  'Const S_JNO42 As Integer = 49
  'Const S_JNO43 As Integer = 50
  'Const S_JNO44 As Integer = 51
  'Const S_JNO45 As Integer = 52
  'Const S_JNO46 As Integer = 53
  'Const S_JNO47 As Integer = 54
  'Const S_JNO48 As Integer = 55
  'Const S_JNO49 As Integer = 56
  'Const S_JNO50 As Integer = 57
  'Const S_JNO51 As Integer = 58
  'Const S_JNO52 As Integer = 59
  'Const S_JNO53 As Integer = 60
  'Const S_JNO54 As Integer = 61
  'Const S_JNO55 As Integer = 62
  'Const S_JNO56 As Integer = 63
  'Const S_JNO57 As Integer = 64
  'Const S_JNO58 As Integer = 65
  'Const S_JNO59 As Integer = 66
  'Const S_JNO60 As Integer = 67
  'Const S_JNO61 As Integer = 68
  'Const S_JNO61 As Integer = 69
  'Const S_JNO61 As Integer = 70
  'Const S_JNO61 As Integer = 71
  'Const S_JNO61 As Integer = 72
  'Const S_JNO61 As Integer = 73
  'Const S_JNO61 As Integer = 74
  'Const S_JNO61 As Integer = 75
  'Const S_JNO61 As Integer = 76
  'Const S_JNO71 As Integer = 77
  'Const S_JNO71 As Integer = 78
  'Const S_JNO71 As Integer = 79
  'Const S_JNO71 As Integer = 80
  'Const S_JNO71 As Integer = 81
  'Const S_JNO71 As Integer = 82
  'Const S_JNO71 As Integer = 83
  'Const S_JNO71 As Integer = 84
  'Const S_JNO71 As Integer = 85
  'Const S_JNO71 As Integer = 86
  'Const S_JNO81 As Integer = 87
  'Const S_JNO81 As Integer = 88
  'Const S_JNO81 As Integer = 89
  'Const S_JNO81 As Integer = 90
  'Const S_JNO81 As Integer = 91
  'Const S_JNO81 As Integer = 92
  'Const S_JNO81 As Integer = 93
  'Const S_JNO81 As Integer = 94
  'Const S_JNO81 As Integer = 95
  'Const S_JNO81 As Integer = 96
  'Const S_JNO91 As Integer = 97
  Const S_SHR02 As Integer = 98
  'Const S_SHR03 As Integer = 99
  'Const S_SHR04 As Integer = 100
  'Const S_SHR05 As Integer = 101
  'Const S_SHR06 As Integer = 102
  'Const S_SHR07 As Integer = 103
  'Const S_SHR08 As Integer = 104
  'Const S_SHR09 As Integer = 105
  'Const S_SHR10 As Integer = 106
  'Const S_SHR11 As Integer = 107
  'Const S_SHR12 As Integer = 108
  'Const S_SHR13 As Integer = 109
  'Const S_SHR14 As Integer = 110
  'Const S_SHR15 As Integer = 111
  'Const S_SHR16 As Integer = 112
  'Const S_SHR17 As Integer = 113
  'Const S_SHR18 As Integer = 114
  'Const S_SHR19 As Integer = 115
  'Const S_SHR20 As Integer = 116
  'Const S_SHR21 As Integer = 117
  'Const S_SHR22 As Integer = 118
  'Const S_SHR23 As Integer = 119
  'Const S_SHR24 As Integer = 120
  'Const S_SHR25 As Integer = 121
  'Const S_SHR26 As Integer = 122
  'Const S_SHR27 As Integer = 123
  'Const S_SHR28 As Integer = 124
  'Const S_SHR29 As Integer = 125
  'Const S_SHR30 As Integer = 126
  'Const S_SHR31 As Integer = 127
  'Const S_SHR32 As Integer = 128
  'Const S_SHR33 As Integer = 129
  'Const S_SHR34 As Integer = 130
  'Const S_SHR35 As Integer = 131
  'Const S_SHR36 As Integer = 132
  'Const S_SHR37 As Integer = 133
  'Const S_SHR38 As Integer = 134
  'Const S_SHR39 As Integer = 135
  'Const S_SHR40 As Integer = 136
  'Const S_SHR41 As Integer = 137
  'Const S_SHR42 As Integer = 138
  'Const S_SHR43 As Integer = 139
  'Const S_SHR44 As Integer = 140
  'Const S_SHR45 As Integer = 141
  'Const S_SHR46 As Integer = 142
  'Const S_SHR47 As Integer = 143
  'Const S_SHR48 As Integer = 144
  'Const S_SHR49 As Integer = 145
  'Const S_SHR50 As Integer = 146
  'Const S_SHR51 As Integer = 147
  'Const S_SHR52 As Integer = 148
  'Const S_SHR53 As Integer = 149
  'Const S_SHR54 As Integer = 150
  'Const S_SHR55 As Integer = 151
  'Const S_SHR56 As Integer = 152
  'Const S_SHR57 As Integer = 153
  'Const S_SHR58 As Integer = 154
  'Const S_SHR59 As Integer = 155
  'Const S_SHR60 As Integer = 156
  'Const S_SHR61 As Integer = 157
  'Const S_SHR62 As Integer = 158
  'Const S_SHR63 As Integer = 159
  'Const S_SHR64 As Integer = 160
  'Const S_SHR65 As Integer = 161
  'Const S_SHR66 As Integer = 162
  'Const S_SHR67 As Integer = 163
  'Const S_SHR68 As Integer = 164
  'Const S_SHR69 As Integer = 165
  'Const S_SHR70 As Integer = 166
  'Const S_SHR71 As Integer = 167
  'Const S_SHR72 As Integer = 168
  'Const S_SHR73 As Integer = 169
  'Const S_SHR74 As Integer = 170
  'Const S_SHR75 As Integer = 171
  'Const S_SHR76 As Integer = 172
  'Const S_SHR77 As Integer = 173
  'Const S_SHR78 As Integer = 174
  'Const S_SHR79 As Integer = 175
  'Const S_SHR80 As Integer = 176
  'Const S_SHR81 As Integer = 177
  'Const S_SHR82 As Integer = 178
  'Const S_SHR83 As Integer = 179
  'Const S_SHR84 As Integer = 180
  'Const S_SHR85 As Integer = 181
  'Const S_SHR86 As Integer = 182
  'Const S_SHR87 As Integer = 183
  'Const S_SHR88 As Integer = 184
  'Const S_SHR89 As Integer = 185
  'Const S_SHR90 As Integer = 186
  Const S_PERSH As Integer = 187

  Const lngAsset_JNO_Beg As Long = 9&
  Const lngAsset_JNO_End As Long = 97&
  Const lngAsset_SHR_Beg As Long = 98&
  Const lngAsset_SHR_End As Long = 186&

  Const lngMaxRecs As Long = 90&

  blnRetVal = True

  DoCmd.Hourglass True
  DoEvents

  Set dbs = CurrentDb
  With dbs

    Set qdf = .QueryDefs(strAcctNoQry)
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngAccts = .RecordCount
      .MoveFirst
      arr_varAcct = .GetRows(lngAccts)
      ' **********************************************
      ' ** Array: arr_varAcct()
      ' **
      ' **   Field  Element  Name         Constant
      ' **   =====  =======  ===========  ==========
      ' **     1       0     accountno    A_ACTNO
      ' **     2       1     assetno      A_ASTNO
      ' **     3       2     cnt          A_CNT
      ' **
      ' **********************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    ' ** These could come from the array!
    lngMaxActNoWidth = 3&
    lngMaxAstNoWidth = 3&
    lngLedgerUpdates = 0&: lngLedgerAdds = 0&: lngTaxLotUpdates = 0&: lngTaxLotDels = 0&

    For lngV = 0& To (lngAccts - 1&)
'If arr_varAcct(A_ACTNO, lngV) = "00292" Then

      strQry = strQryBase2 & Right$(String(lngMaxActNoWidth, "0") & arr_varAcct(A_ACTNO, lngV), lngMaxActNoWidth)
      strQry = strQry & "_" & Right$(String(lngMaxAstNoWidth, "0") & arr_varAcct(A_ASTNO, lngV), lngMaxAstNoWidth)
      strQry = strQry & "_"
      strNewQryBase = strQry
      strQry = strQry & "08"

      ' ** zzz_qry_MasterTrust_10_077_001_08
      strLastQry = strQry  ' ** 08.
      Set qdf = .QueryDefs(strLastQry)
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngNegs = .RecordCount
        .MoveFirst
        arr_varNeg = .GetRows(lngNegs)
        ' **********************************************
        ' ** Array: arr_varNeg()
        ' **
        ' **   Field  Element  Name         Constant
        ' **   =====  =======  ===========  ==========
        ' **     1       0     accountno    N_ACTNO
        ' **     2       1     assetno      N_ASTNO
        ' **     3       2     transdate    N_TDAT
        ' **     4       3     assetdate    N_ADAT
        ' **     5       4     shareface    N_SHARE
        ' **     6       5     icash        N_ICASH
        ' **     7       6     pcash        N_PCASH
        ' **     8       7     cost         N_COST
        ' **
        ' **********************************************
        .Close
      End With
      Set rst = Nothing
      Set qdf = Nothing

      ' ** zzz_qry_MasterTrust_10_077_001_09_01
      strNewQryBase = strNewQryBase & "09_"

      For lngW = 0& To (lngNegs - 1&)

        lngTrans = 0&
        ReDim arr_varTran(T_ELEMS, 0&)

        lngAssets = 0&
        ReDim arr_varAsset(S_ELEMS, 0&)

        strAccountNo = arr_varNeg(N_ACTNO, lngW)
        lngAssetNo = arr_varNeg(N_ASTNO, lngW)

        strQry = strNewQryBase & Right$("00" & CStr(lngW + 1&), 2)
        Set qdf = .QueryDefs(strQry)

        ' ** .._01, just assetdatex = 07/12/2012 17:22:36; 34, journalno >= 18414, transdate = 09/07/2012.
        strDesc = qdf.Properties("Description")
        intPos1 = InStr(strDesc, "=")
        strAssetDate = Trim$(Mid$(strDesc, (intPos1 + 1)))
        intPos1 = InStr(strAssetDate, ";")
        strJournalNo = Trim$(Mid$(strAssetDate, (intPos1 + 1)))
        strAssetDate = Trim$(Left$(strAssetDate, (intPos1 - 1)))
        datAssetDate = CDate(strAssetDate)
        intPos1 = InStr(strJournalNo, "=")
        strJournalNo = Trim$(Mid$(strJournalNo, (intPos1 + 1)))
        intPos1 = InStr(strJournalNo, ",")
        strTransDate = Trim$(Mid$(strJournalNo, (intPos1 + 1)))
        strJournalNo = Trim$(Left$(strJournalNo, (intPos1 - 1)))
        lngJournalNo = Val(strJournalNo)
        intPos1 = InStr(strTransDate, "=")
        strTransDate = Trim$(Mid$(strTransDate, (intPos1 + 1)))
        If Right$(strTransDate, 1) = "." Then strTransDate = Left$(strTransDate, (Len(strTransDate) - 1))
        datTransDate = CDate(strTransDate)
        Set qdf = Nothing

'WHAT ARE THE JOURNALNO AND TRANSDATE USED FOR?
'WHERE DO THEY COME FROM?
'I BELIEVE THE JOURNALNO IS THE FIRST ENTRY PUSHING BELOW ZERO!
'AND TRANSDATE WOULD BE ITS TRANSDATE?

'NEG TAX LOT:
'-2235
'WALK BACKWARDS TO ZERO!
'1. 2035 JNO: 79329  01/17/2013
'-200
'2. 200  JNO: 79327  01/17/2013
'0
'LEDGER RECS UPDATED: 1
'LEDGER RECS ADDED: 25
'TAX LOTS UPDATED: 1
'TAX LOTS DELETED: 26
'DONE!  NegTaxLots_Fix()

'CHK JNO ORD!  zzz_qry_MasterTrust_10_310_001_09_01
'NEW QRYS CREATED: 25
'DONE!  NegTaxLots_Qrys()

'LEDGER RECS UPDATED: 9
'LEDGER RECS ADDED: 37
'TAX LOTS UPDATED: 9
'TAX LOTS DELETED: 46
'DONE!  NegTaxLots_Fix()

        ' ** Ledger, by specified [actno], [astno], [jno], [adat], [pdat].
        Set qdf = .QueryDefs("zzz_qry_TaxLots_01")
        With qdf.Parameters
          ![actno] = strAccountNo
          ![astno] = lngAssetNo
          ![jno] = lngJournalNo
          ![adat] = datAssetDate
          ![pdat] = datAssetDate
        End With
        Set rst = qdf.OpenRecordset
        With rst
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            lngTrans = lngTrans + 1&
            lngE = lngTrans - 1&
            ReDim Preserve arr_varTran(T_ELEMS, lngE)
            ' ***************************************************
            ' ** Array: arr_varTran()
            ' **
            ' **   Field  Element  Name              Constant
            ' **   =====  =======  ================  ==========
            ' **     1       0     journalno         T_JNO
            ' **     2       1     journaltype       T_JTYP
            ' **     3       2     accountno         T_ACTNO
            ' **     4       3     assetno           T_ASTNO
            ' **     5       4     transdate         T_TDAT
            ' **     6       5     assetdate         T_ADAT
            ' **     7       6     PurchaseDate      T_PDAT
            ' **     8       7     shareface         T_SHR
            ' **     9       8     icash             T_ICASH
            ' **    10       9     pcash             T_PCASH
            ' **    11      10     cost              T_COST
            ' **    12      11     description       T_DESC
            ' **    13      12     PurchaseDate02    T_PDAT02
            ' **    ...
            ' **    47      46     PurchaseDate04    T_PDAT36
            ' **    48      47     shareface02       T_SHR02
            ' **    ...
            ' **    82      81     shareface36       T_SHR36
            ' **    83      82     cnt               T_CNT
            ' **
            ' ***************************************************
            arr_varTran(T_JNO, lngE) = ![journalno]
            arr_varTran(T_JTYP, lngE) = ![journaltype]
            arr_varTran(T_ACTNO, lngE) = ![accountno]
            arr_varTran(T_ASTNO, lngE) = ![assetno]
            arr_varTran(T_TDAT, lngE) = ![transdate]
            arr_varTran(T_ADAT, lngE) = ![assetdate]
            arr_varTran(T_PDAT, lngE) = ![PurchaseDate]
            arr_varTran(T_SHR, lngE) = ![shareface]
            arr_varTran(T_ICASH, lngE) = ![ICash]
            arr_varTran(T_PCASH, lngE) = ![PCash]
            arr_varTran(T_COST, lngE) = ![Cost]
            If IsNull(![description]) = False Then
              arr_varTran(T_DESC, lngE) = ![description]
            Else
              arr_varTran(T_DESC, lngE) = vbNullString
            End If
'arr_varTran(12 To 46, lngE)
            For lngY = lngTrans_PDAT_Beg To lngTrans_PDAT_End
              arr_varTran(lngY, lngE) = Null
            Next  ' ** lngY.
'arr_varTran(47 To 81, lngE)
            For lngY = lngTrans_SHR_Beg To lngTrans_SHR_End
              arr_varTran(lngY, lngE) = Null
            Next  ' ** lngY.
            arr_varTran(T_CNT, lngE) = CLng(0)
            If lngX < lngRecs Then .MoveNext
          Next  ' ** lngX.
          .Close
        End With
        Set rst = Nothing
        Set qdf = Nothing

        ' ** ActiveAssets, by specified [actno], [astno].
        Set qdf = .QueryDefs("zzz_qry_TaxLots_02")
        With qdf.Parameters
          ![actno] = strAccountNo
          ![astno] = lngAssetNo
        End With
        Set rst = qdf.OpenRecordset
        With rst
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            lngAssets = lngAssets + 1&
            lngE = lngAssets - 1&
            ReDim Preserve arr_varAsset(S_ELEMS, lngE)
            ' **************************************************
            ' ** Array: arr_varAsset()
            ' **
            ' **   Field  Element  Name             Constant
            ' **   =====  =======  ===============  ==========
            ' **     1       0     accountno        S_ACTNO
            ' **     2       1     assetno          S_ASTNO
            ' **     3       2     transdate        S_TDAT
            ' **     4       3     assetdate        S_ADAT
            ' **     5       4     shareface        S_SHR
            ' **     6       5     icash            S_ICASH
            ' **     7       6     pcash            S_PCASH
            ' **     8       7     cost             S_COST
            ' **     9       8     shareface_new    S_SHRN
            ' **    10       9     journalno02      S_JNO02
            ' **    ...
            ' **    44      43     journalno36      S_JNO36
            ' **    45      44     shareface02      S_SHR02
            ' **    ...
            ' **    79      78     shareface36      S_SHR36
            ' **    80      79     priceperunit     S_PERSH
            ' **
            ' **************************************************
            arr_varAsset(S_ACTNO, lngE) = ![accountno]
            arr_varAsset(S_ASTNO, lngE) = ![assetno]
            arr_varAsset(S_TDAT, lngE) = ![transdate]
            arr_varAsset(S_ADAT, lngE) = ![assetdate]
            arr_varAsset(S_SHR, lngE) = ![shareface]
            arr_varAsset(S_ICASH, lngE) = ![ICash]
            arr_varAsset(S_PCASH, lngE) = ![PCash]
            arr_varAsset(S_COST, lngE) = ![Cost]
            arr_varAsset(S_SHRN, lngE) = ![shareface]
'arr_varAsset(9 To 43)
            For lngY = lngAsset_JNO_Beg To lngAsset_JNO_End
              arr_varAsset(lngY, lngE) = Null
            Next  ' ** lngY.
'arr_varAsset(44 To 78)
            For lngY = lngAsset_SHR_Beg To lngAsset_SHR_End
              arr_varAsset(lngY, lngE) = Null
            Next  ' ** lngY.
            arr_varAsset(S_PERSH, lngE) = ![priceperunit]
            If lngX < lngRecs Then .MoveNext
          Next  ' ** lngX
          .Close
        End With
        Set rst = Nothing
        Set qdf = Nothing

        For lngX = 0& To (lngAssets - 1&)
          If arr_varAsset(S_SHR, lngX) < 0# Then
            dblShareFaceNeg = arr_varAsset(S_SHR, lngX)
            Exit For
          End If
        Next

'T_PDAT : 12 - 46
'T_SHR  : 47 - 81
'S_JNO  : 9 - 43
'S_SHR  : 44 - 78

        For lngX = 0& To (lngTrans - 1&)
          dblShares = Round(arr_varTran(T_SHR, lngX), 4)
          dblTmp01 = dblShares  ' ** Starts out same as shareface.
          dblTmp02 = 0#
          For lngY = 0& To (lngAssets - 1&)
            If arr_varAsset(S_SHR, lngY) > 0# Then  ' ** Skip the Neg!
              If arr_varAsset(S_SHRN, lngY) > 0# Then  ' ** Starts out same as shareface.
                If dblShares <= arr_varAsset(S_SHRN, lngY) Then
                  dblTmp02 = dblShares  ' ** shareface going to this Tax Lot.
                  arr_varAsset(S_SHRN, lngY) = Round((arr_varAsset(S_SHRN, lngY) - dblTmp02), 4)
                  dblShares = 0#
                  If arr_varAsset(S_SHRN, lngY) > 0# And arr_varAsset(S_SHRN, lngY) < 0.0001 Then arr_varAsset(S_SHRN, lngY) = 0#
                Else
                  dblTmp02 = arr_varAsset(S_SHRN, lngY)  ' ** shareface going to this Tax Lot.
                  dblShares = Round((dblShares - arr_varAsset(S_SHRN, lngY)), 4)  ' ** dblShares now reduced.
                  arr_varAsset(S_SHRN, lngY) = 0#
                  If dblShares > 0# And dblShares < 0.0001 Then dblShares = 0#
                End If
                dblShares = Round(dblShares, 4)  ' ** May be Zero.
                arr_varTran(T_CNT, lngX) = arr_varTran(T_CNT, lngX) + 1&
                blnFound = False
                For lngZ = lngAsset_JNO_Beg To lngAsset_JNO_End
                  If IsNull(arr_varAsset(lngZ, lngY)) = True Then
                    blnFound = True
                    arr_varAsset(lngZ, lngY) = arr_varTran(T_JNO, lngX)  ' ** S_JNO00
'arr_varAsset((lngZ + 35, lngY)
                    arr_varAsset((lngZ + (lngMaxRecs - 1&)), lngY) = dblTmp02           ' ** S_SHR00
                    Exit For
                  End If
                Next  ' ** lngZ
                If blnFound = False Then
                  Stop
                End If
                blnFound = False
'arr_varTran(12 To 46, lngX)
                For lngZ = lngTrans_PDAT_Beg To lngTrans_PDAT_End
                  If IsNull(arr_varTran(lngZ, lngX)) = True Then
                    blnFound = True
                    arr_varTran(lngZ, lngX) = arr_varAsset(S_ADAT, lngY)  ' ** T_PDAT00
'arr_varTran((lngZ + 35), lngX)
                    arr_varTran((lngZ + (lngMaxRecs - 1&)), lngX) = dblTmp02            ' ** T_SHR00
                    Exit For
                  End If
                Next  ' ** lngZ.
                If blnFound = False Then
                  Stop
                End If
              End If
            End If
            If dblShares = 0# Then
              Exit For
            End If
          Next  ' ** lngY.
        Next  ' ** lngX.

        ' ** The transaction has now been apportioned.
        ' ** Each JNo has to be redirected to the new PurchaseDate, and, if its
        ' ** apportionment had to be split, the original JNo must be so edited.

'T_PDAT : 12 - 46
'T_SHR  : 47 - 81
'S_JNO  : 9 - 43
'S_SHR  : 44 - 78

        For lngX = 0& To (lngTrans - 1&)

          strTmp00 = "Delta Data, Inc.: Negative Share/Face redirection; made " & Format(Date, "mm/dd/yyyy")
          If IsNull(arr_varTran(T_DESC, lngX)) = False Then
            strTmp00 = arr_varTran(T_DESC, lngX) & "; " & strTmp00
          End If

          If arr_varTran(T_CNT, lngX) = 1& Then
            ' ** All of it went to a single other Tax Lot.

            ' ** Update zzz_qry_TaxLots_03 (Ledger, by specified [jno], [pdat], [dsc]).
            Set qdf = .QueryDefs("zzz_qry_TaxLots_04")
            With qdf.Parameters
              ![jno] = arr_varTran(T_JNO, lngX)
              ![pdat] = arr_varTran(T_PDAT02, lngX)
              ![dsc] = strTmp00
            End With
            qdf.Execute
            Set qdf = Nothing
            lngLedgerUpdates = lngLedgerUpdates + 1&

          Else
            ' ** It took more than 1 other Tax Lot to absorb the Sold.

            ' ** These are the original values.
            dblShares = arr_varTran(T_SHR, lngX)
            dblICash = arr_varTran(T_ICASH, lngX)
            dblPCash = arr_varTran(T_PCASH, lngX)
            dblCost = arr_varTran(T_COST, lngX)

            dblTmp01 = 0#

            ' ** Get a price per share.
            dblPerShare = Round((dblCost / dblShares), 2)
            dblShares2 = arr_varTran(T_SHR02, lngX)  ' ** The first shareface value.
            dblICashSum = 0#: dblPCashSum = 0#: dblCostSum = 0#

            ' ** Update zzz_qry_TaxLots_05 (Ledger, by specified [jno], [pdat], [dsc], [sharf], [icsh], [pcsh], [cst]).
            Set qdf = .QueryDefs("zzz_qry_TaxLots_06")
            With qdf.Parameters
              ![jno] = arr_varTran(T_JNO, lngX)
              ![pdat] = arr_varTran(T_PDAT02, lngX)
              ![dsc] = strTmp00
              ![sharf] = dblShares2
              dblTmp01 = 0#
              If dblICash = 0# Then
                ![icsh] = 0#
              Else
                dblTmp01 = Round(((dblICash / dblShares) * dblShares2), 2)
                ![icsh] = dblTmp01
                dblICashSum = (dblICashSum + dblTmp01)
              End If
              dblTmp01 = 0#
              If dblPCash = 0# Then
                ![pcsh] = 0#
              Else
                dblTmp01 = Round(((dblPCash / dblShares) * dblShares2), 2)
                ![pcsh] = dblTmp01
                dblPCashSum = (dblPCashSum + dblTmp01)
              End If
              dblTmp01 = 0#
              dblTmp01 = Round((dblShares2 * dblPerShare), 2)
              ![cst] = dblTmp01
              dblCostSum = (dblCostSum + dblTmp01)
            End With
            qdf.Execute
            Set qdf = Nothing
            lngLedgerUpdates = lngLedgerUpdates + 1&

            strTmp00 = "Delta Data, Inc.: Negative Share/Face adjustment; made " & Format(Date, "mm/dd/yyyy")

            ' ** Create new transactions for the remainder
            For lngY = 2& To arr_varTran(T_CNT, lngX)
              lngTmp03 = (T_SHR02 + (lngY - 1&))
              dblShares2 = arr_varTran(lngTmp03, lngX)
              ' ** Append zzz_qry_TaxLots_07 (Ledger, as new 'Sold'/'Withdrawn' entry, by
              ' ** specified [jno], [pdat], [dsc], [sharf], [icsh], [pcsh], [cst]) to Ledger.
              Set qdf = .QueryDefs("zzz_qry_TaxLots_08")
              With qdf.Parameters
                ![jno] = arr_varTran(T_JNO, lngX)
                lngTmp03 = (T_PDAT02 + (lngY - 1&))
                ![pdat] = arr_varTran(lngTmp03, lngX)
                ![dsc] = strTmp00
                ![sharf] = dblShares2
                dblTmp01 = 0#
                If dblICash = 0# Then
                  ![icsh] = 0#
                Else
                  dblTmp01 = Round(((dblICash / dblShares) * dblShares2), 2)
                  ![icsh] = dblTmp01
                  dblICashSum = (dblICashSum + dblTmp01)
                End If
                dblTmp01 = 0#
                If dblPCash = 0# Then
                  ![pcsh] = 0#
                Else
                  dblTmp01 = Round(((dblPCash / dblShares) * dblShares2), 2)
                  ![pcsh] = dblTmp01
                  dblPCashSum = (dblPCashSum + dblTmp01)
                End If
                dblTmp01 = 0#
                dblTmp01 = Round((dblShares2 * dblPerShare), 2)
                ![cst] = dblTmp01
                dblCostSum = (dblCostSum + dblTmp01)
              End With
              qdf.Execute
              Set qdf = Nothing
              DoEvents
              lngLedgerAdds = lngLedgerAdds + 1&
            Next  ' ** lngY.

          End If  ' ** T_CNT.

        Next  ' ** lngX.

'T_PDAT : 12 - 46
'T_SHR  : 47 - 81
'S_JNO  : 9 - 43
'S_SHR  : 44 - 78

        ' ** Now delete and/or update ActiveAssets to reflect reassignment.
        For lngX = 0& To (lngAssets - 1&)
          If arr_varAsset(S_SHR, lngX) < 0# Then
            ' ** Delete the Negative!

            ' ** Delete ActiveAssets, by specified [actno], [astno], [adat].
            Set qdf = .QueryDefs("zzz_qry_TaxLots_09")
            With qdf.Parameters
              ![actno] = arr_varAsset(S_ACTNO, lngX)
              ![astno] = arr_varAsset(S_ASTNO, lngX)
              ![adat] = arr_varAsset(S_ADAT, lngX)
            End With
            qdf.Execute
            Set qdf = Nothing
            DoEvents
            lngTaxLotDels = lngTaxLotDels + 1&

          ElseIf arr_varAsset(S_SHRN, lngX) <> arr_varAsset(S_SHR, lngX) Then
            ' ** It looks like it's been used.

            blnFound = False
'arr_varAsset(9 To 43, lngX)
            For lngY = lngAsset_JNO_Beg To lngAsset_JNO_End
              If IsNull(arr_varAsset(lngY, lngX)) = False Then
                ' ** Aha! It HAS been used!
                blnFound = True
                Exit For
              End If
            Next  ' ** lngY.

            If blnFound = True Then
              If arr_varAsset(S_SHRN, lngX) = 0# Then
                ' ** All used up, delete it.

                ' ** Delete ActiveAssets, by specified [actno], [astno], [adat].
                Set qdf = .QueryDefs("zzz_qry_TaxLots_09")
                With qdf.Parameters
                  ![actno] = arr_varAsset(S_ACTNO, lngX)
                  ![astno] = arr_varAsset(S_ASTNO, lngX)
                  ![adat] = arr_varAsset(S_ADAT, lngX)
                End With
                qdf.Execute
                Set qdf = Nothing
                DoEvents
                lngTaxLotDels = lngTaxLotDels + 1&

              Else
                ' ** Partially used, edit it.

                dblShares = arr_varAsset(S_SHRN, lngX)
                dblPerShare = arr_varAsset(S_PERSH, lngX)

                ' ** Update zzz_qry_TaxLots_10 (ActiveAssets, by specified [actno], [astno], [adat], [sharf], [icsh], [pcsh], [cst]).
                Set qdf = .QueryDefs("zzz_qry_TaxLots_11")
                With qdf.Parameters
                  ![actno] = arr_varAsset(S_ACTNO, lngX)
                  ![astno] = arr_varAsset(S_ASTNO, lngX)
                  ![adat] = arr_varAsset(S_ADAT, lngX)
                  ![sharf] = dblShares
                  dblTmp01 = Round((dblShares * dblPerShare), 2)
                  If arr_varAsset(S_ICASH, lngX) = 0# Then
                    ![icsh] = 0#
                    ![pcsh] = (-dblTmp01)  ' ** ICash or PCash are negative.
                  Else
                    ![icsh] = (-dblTmp01)
                    ![pcsh] = 0#
                  End If
                  ![cst] = dblTmp01  '** Cost is positive.
                End With
                qdf.Execute
                Set qdf = Nothing
                lngTaxLotUpdates = lngTaxLotUpdates + 1&

              End If  ' ** S_SHRN.
            End If  ' ** blnFound.

          End If  ' ** S_SHR, S_SHRN.
        Next  ' ** lngX

      Next  ' ** lngW.

'Beep
'End If
    Next  ' ** lngV.

    .Close
  End With  ' ** dbs.

  DoCmd.Hourglass False

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'LEDGER RECS UPDATED: " & CStr(lngLedgerUpdates)
  Debug.Print "'LEDGER RECS ADDED: " & CStr(lngLedgerAdds)
  Debug.Print "'TAX LOTS UPDATED: " & CStr(lngTaxLotUpdates)
  Debug.Print "'TAX LOTS DELETED: " & CStr(lngTaxLotDels)
  Debug.Print "'DONE!  " & THIS_PROC & "()"
'LEDGER RECS UPDATED: 159
'LEDGER RECS ADDED: 54
'TAX LOTS UPDATED: 63
'TAX LOTS DELETED: 119
'DONE!  NegTaxLots_Fix()
'WITHOUT '00214'!!!!!!!!!!!!!!!!!!!!!!!!!!

  DoBeeps  ' ** Module Function: modWindowFunctions.

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NegTaxLots_Fix = blnRetVal

End Function

Public Function NegTaxLots_Match1() As Boolean

  Const THIS_PROC As String = "NegTaxLots_Match1"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngTrans As Long, arr_varTran() As Variant
  Dim lngElems As Long, arr_varElem() As Variant
  Dim dblDeposit As Double, lngDepElem As Long, dblBal As Double
  Dim lngRecs As Long
  Dim lngX As Long, lngE As Long
  Dim lngA As Long, lngB As Long, lngC As Long, lngD As Long, lngF As Long, lngG As Long, lngH As Long, lngI As Long, lngJ As Long, lngK As Long
  Dim lngL As Long, lngM As Long, lngN As Long, lngO As Long, lngP As Long, lngQ As Long, lngR As Long, lngS As Long, lngT As Long, lngU As Long
  Dim lngV As Long, lngW As Long, lngY As Long, lngZ As Long
  Dim blnFound As Boolean
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTran().
  Const T_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const T_JNO   As Integer = 0
  Const T_JTYP  As Integer = 1
  Const T_ACTNO As Integer = 2
  Const T_ASTNO As Integer = 3
  Const T_SHARE As Integer = 4

  ' ** Array: arr_varElem().
  Const E_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const E_VAL As Integer = 0

  blnRetVal = True

  lngTrans = 0&
  ReDim arr_varTran(T_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    Set qdf = .QueryDefs("zzz_qry_MasterTrust_10_214_001_09_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        lngTrans = lngTrans + 1&
        lngE = lngTrans - 1&
        ReDim Preserve arr_varTran(T_ELEMS, lngE)
        arr_varTran(T_JNO, lngE) = ![journalno]
        arr_varTran(T_JTYP, lngE) = ![journaltype]
        arr_varTran(T_ACTNO, lngE) = ![accountno]
        arr_varTran(T_ASTNO, lngE) = ![assetno]
        arr_varTran(T_SHARE, lngE) = ![shareface]
        ' ************************************************
        ' ** Array: arr_varTrans()
        ' **
        ' **   Field  Element  Name           Constant
        ' **   =====  =======  =============  ==========
        ' **     1       0     journalno      T_JNO
        ' **     2       1     journaltype    T_JTYP
        ' **     3       2     accountno      T_ACTNO
        ' **     4       3     assetno        T_ASTNO
        ' **     5       4     shareface      T_SHARE
        ' **
        ' ************************************************
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With
  Set dbs = Nothing

  lngDepElem = -1&
  dblDeposit = 0#: dblBal = 0#

  ' ** Get the initial Purchase/Deposit.
  For lngX = 0& To (lngTrans - 1&)
    If arr_varTran(T_JTYP, lngX) = "Deposit" Or arr_varTran(T_JTYP, lngX) = "Purchase" Then
      dblDeposit = arr_varTran(T_SHARE, lngX)
      lngDepElem = lngX
      Exit For
    End If
  Next  ' ** lngX

  ' ** 40 records.
  If lngDepElem >= 0& Then
    blnFound = False
    For lngA = 0& To (lngTrans - 1&)
      If lngA <> lngDepElem Then
        dblBal = dblDeposit
        dblBal = dblBal - arr_varTran(T_SHARE, lngA)
        If dblBal <= -0.0002 Then
          dblBal = dblBal + arr_varTran(T_SHARE, lngA)
          Exit For
        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
          blnFound = True
          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngA)
          Stop
        Else
          For lngB = 0& To (lngTrans - 1&)
            If lngB <> lngDepElem And lngB <> lngA Then
              dblBal = dblBal - arr_varTran(T_SHARE, lngB)
              If dblBal <= -0.0002 Then
                dblBal = dblBal + arr_varTran(T_SHARE, lngB)
                Exit For
              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                blnFound = True
                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngB)
                Stop
              Else
                For lngC = 0& To (lngTrans - 1&)
                  If lngC <> lngDepElem And lngC <> lngA And lngC <> lngB Then
                    dblBal = dblBal - arr_varTran(T_SHARE, lngC)
                    If dblBal <= -0.0002 Then
                      dblBal = dblBal + arr_varTran(T_SHARE, lngC)
                      Exit For
                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                      blnFound = True
                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngC)
                      Stop
                    Else
                      For lngD = 0& To (lngTrans - 1&)
                        If lngD <> lngDepElem And lngD <> lngA And lngD <> lngB And lngD <> lngC Then
                          dblBal = dblBal - arr_varTran(T_SHARE, lngD)
                          If dblBal <= -0.0002 Then
                            dblBal = dblBal + arr_varTran(T_SHARE, lngD)
                            Exit For
                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                            blnFound = True
                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngD)
                            Stop
                          Else
                            For lngE = 0& To (lngTrans - 1&)
                              If lngE <> lngDepElem And lngE <> lngA And lngE <> lngB And lngE <> lngC And lngE <> lngD Then
                                dblBal = dblBal - arr_varTran(T_SHARE, lngE)
                                If dblBal <= -0.0002 Then
                                  dblBal = dblBal + arr_varTran(T_SHARE, lngE)
                                  Exit For
                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                  blnFound = True
                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngE)
                                  Stop
                                Else
                                  For lngF = 0& To (lngTrans - 1&)
                                    If lngF <> lngDepElem And lngF <> lngA And lngF <> lngB And lngF <> lngC And lngF <> lngD And lngF <> lngE Then
                                      dblBal = dblBal - arr_varTran(T_SHARE, lngF)
                                      If dblBal <= -0.0002 Then
                                        dblBal = dblBal + arr_varTran(T_SHARE, lngF)
                                        Exit For
                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                        blnFound = True
                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngF)
                                        Stop
                                      Else
                                        For lngG = 0& To (lngTrans - 1&)
                                          If lngG <> lngDepElem And lngG <> lngA And lngG <> lngB And lngG <> lngC And lngG <> lngD And lngG <> lngE And _
                                              lngG <> lngF Then
                                            dblBal = dblBal - arr_varTran(T_SHARE, lngG)
                                            If dblBal <= -0.0002 Then
                                              dblBal = dblBal + arr_varTran(T_SHARE, lngG)
                                              Exit For
                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                              blnFound = True
                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngG)
                                              Stop
                                            Else
                                              For lngH = 0& To (lngTrans - 1&)
                                                If lngH <> lngDepElem And lngH <> lngA And lngH <> lngB And lngH <> lngC And lngH <> lngD And lngH <> lngE And _
                                                    lngH <> lngF And lngH <> lngG Then
                                                  dblBal = dblBal - arr_varTran(T_SHARE, lngH)
                                                  If dblBal <= -0.0002 Then
                                                    dblBal = dblBal + arr_varTran(T_SHARE, lngH)
                                                    Exit For
                                                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                    blnFound = True
                                                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngH)
                                                    Stop
                                                  Else
                                                    For lngI = 0& To (lngTrans - 1&)
                                                      If lngI <> lngDepElem And lngI <> lngA And lngI <> lngB And lngI <> lngC And lngI <> lngD And _
                                                          lngI <> lngE And lngI <> lngF And lngI <> lngG And lngI <> lngH Then
                                                        dblBal = dblBal - arr_varTran(T_SHARE, lngI)
                                                        If dblBal <= -0.0002 Then
                                                          dblBal = dblBal + arr_varTran(T_SHARE, lngI)
                                                          Exit For
                                                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                          blnFound = True
                                                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngI)
                                                          Stop
                                                        Else
                                                          For lngJ = 0& To (lngTrans - 1&)
                                                            If lngJ <> lngDepElem And lngJ <> lngA And lngJ <> lngB And lngJ <> lngC And lngJ <> lngD And _
                                                                lngJ <> lngE And lngJ <> lngF And lngJ <> lngG And lngJ <> lngH And lngJ <> lngI Then
                                                              dblBal = dblBal - arr_varTran(T_SHARE, lngJ)
                                                              If dblBal <= -0.0002 Then
                                                                dblBal = dblBal + arr_varTran(T_SHARE, lngJ)
                                                                Exit For
                                                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                blnFound = True
                                                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngJ)
                                                                Stop
                                                              Else
                                                                For lngK = 0& To (lngTrans - 1&)
                                                                  If lngK <> lngDepElem And lngK <> lngA And lngK <> lngB And lngK <> lngC And lngK <> lngD And _
                                                                      lngK <> lngE And lngK <> lngF And lngK <> lngG And lngK <> lngH And lngK <> lngI And _
                                                                      lngK <> lngJ Then
                                                                    dblBal = dblBal - arr_varTran(T_SHARE, lngK)
                                                                    If dblBal <= -0.0002 Then
                                                                      dblBal = dblBal + arr_varTran(T_SHARE, lngK)
                                                                      Exit For
                                                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                      blnFound = True
                                                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngK)
                                                                      Stop
                                                                    Else
                                                                      For lngL = 0& To (lngTrans - 1&)
                                                                        If lngL <> lngDepElem And lngL <> lngA And lngL <> lngB And lngL <> lngC And lngL <> lngD And _
                                                                            lngL <> lngE And lngL <> lngF And lngL <> lngG And lngL <> lngH And lngL <> lngI And _
                                                                            lngL <> lngJ And lngL <> lngK Then
                                                                          dblBal = dblBal - arr_varTran(T_SHARE, lngL)
                                                                          If dblBal <= -0.0002 Then
                                                                            dblBal = dblBal + arr_varTran(T_SHARE, lngL)
                                                                            Exit For
                                                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                            blnFound = True
                                                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngL)
                                                                            Stop
                                                                          Else
                                                                            For lngM = 0& To (lngTrans - 1&)
                                                                              If lngM <> lngDepElem And lngM <> lngA And lngM <> lngB And lngM <> lngC And lngM <> lngD And _
                                                                                  lngM <> lngE And lngM <> lngF And lngM <> lngG And lngM <> lngH And lngM <> lngI And _
                                                                                  lngM <> lngJ And lngM <> lngK And lngM <> lngL Then
                                                                                dblBal = dblBal - arr_varTran(T_SHARE, lngM)
                                                                                If dblBal <= -0.0002 Then
                                                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngM)
                                                                                  Exit For
                                                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                  blnFound = True
                                                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngM)
                                                                                  Stop
                                                                                Else
                                                                                  For lngN = 0& To (lngTrans - 1&)
                                                                                    If lngN <> lngDepElem And lngN <> lngA And lngN <> lngB And lngN <> lngC And _
                                                                                        lngN <> lngD And lngN <> lngE And lngN <> lngF And lngN <> lngG And _
                                                                                        lngN <> lngH And lngN <> lngI And lngN <> lngJ And lngN <> lngK And _
                                                                                        lngN <> lngL And lngN <> lngM Then
                                                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngN)
                                                                                      If dblBal <= -0.0002 Then
                                                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngN)
                                                                                        Exit For
                                                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                        blnFound = True
                                                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngN)
                                                                                        Stop
                                                                                      Else
                                                                                        For lngO = 0& To (lngTrans - 1&)
                                                                                          If lngO <> lngDepElem And lngO <> lngA And lngO <> lngB And lngO <> lngC And _
                                                                                              lngO <> lngD And lngO <> lngE And lngO <> lngF And lngO <> lngG And _
                                                                                              lngO <> lngH And lngO <> lngI And lngO <> lngJ And lngO <> lngK And _
                                                                                              lngO <> lngL And lngO <> lngM And lngO <> lngN Then
                                                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngO)
                                                                                            If dblBal <= -0.0002 Then
                                                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngO)
                                                                                              Exit For
                                                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                              blnFound = True
                                                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngO)
                                                                                              Stop
                                                                                            Else
                                                                                              For lngP = 0& To (lngTrans - 1&)
                                                                                                If lngP <> lngDepElem And lngP <> lngA And lngP <> lngB And lngP <> lngC And _
                                                                                                    lngP <> lngD And lngP <> lngE And lngP <> lngF And lngP <> lngG And _
                                                                                                    lngP <> lngH And lngP <> lngI And lngP <> lngJ And lngP <> lngK And _
                                                                                                    lngP <> lngL And lngP <> lngM And lngP <> lngN And lngP <> lngO Then
                                                                                                  dblBal = dblBal - arr_varTran(T_SHARE, lngP)
                                                                                                  If dblBal <= -0.0002 Then
                                                                                                    dblBal = dblBal + arr_varTran(T_SHARE, lngP)
                                                                                                    Exit For
                                                                                                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                    blnFound = True
                                                                                                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngP)
                                                                                                    Stop
                                                                                                  Else
                                                                                                    For lngQ = 0& To (lngTrans - 1&)
                                                                                                      If lngQ <> lngDepElem And lngQ <> lngA And lngQ <> lngB And lngQ <> lngC And _
                                                                                                          lngQ <> lngD And lngQ <> lngE And lngQ <> lngF And lngQ <> lngG And _
                                                                                                          lngQ <> lngH And lngQ <> lngI And lngQ <> lngJ And lngQ <> lngK And _
                                                                                                          lngQ <> lngL And lngQ <> lngM And lngQ <> lngN And lngQ <> lngO And _
                                                                                                          lngQ <> lngP Then
                                                                                                        dblBal = dblBal - arr_varTran(T_SHARE, lngQ)
                                                                                                        If dblBal <= -0.0002 Then
                                                                                                          dblBal = dblBal + arr_varTran(T_SHARE, lngQ)
                                                                                                          Exit For
                                                                                                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                          blnFound = True
                                                                                                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngQ)
                                                                                                          Stop
                                                                                                        Else
                                                                                                          For lngR = 0& To (lngTrans - 1&)
                                                                                                            If lngR <> lngDepElem And lngR <> lngA And lngR <> lngB And _
                                                                                                                lngR <> lngC And lngR <> lngD And lngR <> lngE And lngR <> lngF And _
                                                                                                                lngR <> lngG And lngR <> lngH And lngR <> lngI And lngR <> lngJ And _
                                                                                                                lngR <> lngK And lngR <> lngL And lngR <> lngM And lngR <> lngN And _
                                                                                                                lngR <> lngO And lngR <> lngP And lngR <> lngQ Then
                                                                                                              dblBal = dblBal - arr_varTran(T_SHARE, lngR)
                                                                                                              If dblBal <= -0.0002 Then
                                                                                                                dblBal = dblBal + arr_varTran(T_SHARE, lngR)
                                                                                                                Exit For
                                                                                                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                blnFound = True
                                                                                                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngR)
                                                                                                                Stop
                                                                                                              Else
                                                                                                                For lngS = 0& To (lngTrans - 1&)
                                                                                                                  If lngS <> lngDepElem And lngS <> lngA And lngS <> lngB And _
                                                                                                                      lngS <> lngC And lngS <> lngD And lngS <> lngE And _
                                                                                                                      lngS <> lngF And lngS <> lngG And lngS <> lngH And _
                                                                                                                      lngS <> lngI And lngS <> lngJ And lngS <> lngK And _
                                                                                                                      lngS <> lngL And lngS <> lngM And lngS <> lngN And _
                                                                                                                      lngS <> lngO And lngS <> lngP And lngS <> lngQ And _
                                                                                                                      lngS <> lngR Then
                                                                                                                    dblBal = dblBal - arr_varTran(T_SHARE, lngS)
                                                                                                                    If dblBal <= -0.0002 Then
                                                                                                                      dblBal = dblBal + arr_varTran(T_SHARE, lngS)
                                                                                                                      Exit For
                                                                                                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                      blnFound = True
                                                                                                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngS)
                                                                                                                      Stop
                                                                                                                    Else
                                                                                                                      For lngT = 0& To (lngTrans - 1&)
                                                                                                                        If lngT <> lngDepElem And lngT <> lngA And lngT <> lngB And _
                                                                                                                            lngT <> lngC And lngT <> lngD And lngT <> lngE And _
                                                                                                                            lngT <> lngF And lngT <> lngG And lngT <> lngH And _
                                                                                                                            lngT <> lngI And lngT <> lngJ And lngT <> lngK And _
                                                                                                                            lngT <> lngL And lngT <> lngM And lngT <> lngN And _
                                                                                                                            lngT <> lngO And lngT <> lngP And lngT <> lngQ And _
                                                                                                                            lngT <> lngR And lngT <> lngS Then
                                                                                                                          dblBal = dblBal - arr_varTran(T_SHARE, lngT)
                                                                                                                          If dblBal <= -0.0002 Then
                                                                                                                            dblBal = dblBal + arr_varTran(T_SHARE, lngT)
                                                                                                                            Exit For
                                                                                                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                            blnFound = True
                                                                                                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngT)
                                                                                                                            Stop
                                                                                                                          Else
                                                                                                                            For lngU = 0& To (lngTrans - 1&)
                                                                                                                              If lngU <> lngDepElem And lngU <> lngA And lngU <> lngB And _
                                                                                                                                  lngU <> lngC And lngU <> lngD And lngU <> lngE And _
                                                                                                                                  lngU <> lngF And lngU <> lngG And lngU <> lngH And _
                                                                                                                                  lngU <> lngI And lngU <> lngJ And lngU <> lngK And _
                                                                                                                                  lngU <> lngL And lngU <> lngM And lngU <> lngN And _
                                                                                                                                  lngU <> lngO And lngU <> lngP And lngU <> lngQ And _
                                                                                                                                  lngU <> lngR And lngU <> lngS And lngU <> lngT Then
                                                                                                                                dblBal = dblBal - arr_varTran(T_SHARE, lngU)
                                                                                                                                If dblBal <= -0.0002 Then
                                                                                                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngU)
                                                                                                                                  Exit For
                                                                                                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                  blnFound = True
                                                                                                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngU)
                                                                                                                                  Stop
                                                                                                                                Else
                                                                                                                                  For lngV = 0& To (lngTrans - 1&)
                                                                                                                                    If lngV <> lngDepElem And lngV <> lngA And _
                                                                                                                                        lngV <> lngB And lngV <> lngC And lngV <> lngD And _
                                                                                                                                        lngV <> lngE And lngV <> lngF And lngV <> lngG And _
                                                                                                                                        lngV <> lngH And lngV <> lngI And lngV <> lngJ And _
                                                                                                                                        lngV <> lngK And lngV <> lngL And lngV <> lngM And _
                                                                                                                                        lngV <> lngN And lngV <> lngO And lngV <> lngP And _
                                                                                                                                        lngV <> lngQ And lngV <> lngR And lngV <> lngS And _
                                                                                                                                        lngV <> lngT And lngV <> lngU Then
                                                                                                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngV)
                                                                                                                                      If dblBal <= -0.0002 Then
                                                                                                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngV)
                                                                                                                                        Exit For
                                                                                                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                        blnFound = True
                                                                                                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngV)
                                                                                                                                        Stop
                                                                                                                                      Else
                                                                                                                                        For lngW = 0& To (lngTrans - 1&)
                                                                                                                                          If lngW <> lngDepElem And lngW <> lngA And _
                                                                                                                                              lngW <> lngB And lngW <> lngC And _
                                                                                                                                              lngW <> lngD And lngW <> lngE And _
                                                                                                                                              lngW <> lngF And lngW <> lngG And _
                                                                                                                                              lngW <> lngH And lngW <> lngI And _
                                                                                                                                              lngW <> lngJ And lngW <> lngK And _
                                                                                                                                              lngW <> lngL And lngW <> lngM And _
                                                                                                                                              lngW <> lngN And lngW <> lngO And _
                                                                                                                                              lngW <> lngP And lngW <> lngQ And _
                                                                                                                                              lngW <> lngR And lngW <> lngS And _
                                                                                                                                              lngW <> lngT And lngW <> lngU And _
                                                                                                                                              lngW <> lngV Then
                                                                                                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngW)
                                                                                                                                            If dblBal <= -0.0002 Then
                                                                                                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngW)
                                                                                                                                              Exit For
                                                                                                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                              blnFound = True
                                                                                                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngW)
                                                                                                                                              Stop
                                                                                                                                            Else
                                                                                                                                              For lngX = 0& To (lngTrans - 1&)
                                                                                                                                                If lngX <> lngDepElem And lngX <> lngA And _
                                                                                                                                                    lngX <> lngB And lngX <> lngC And _
                                                                                                                                                    lngX <> lngD And lngX <> lngE And _
                                                                                                                                                    lngX <> lngF And lngX <> lngG And _
                                                                                                                                                    lngX <> lngH And lngX <> lngI And _
                                                                                                                                                    lngX <> lngJ And lngX <> lngK And _
                                                                                                                                                    lngX <> lngL And lngX <> lngM And _
                                                                                                                                                    lngX <> lngN And lngX <> lngO And _
                                                                                                                                                    lngX <> lngP And lngX <> lngQ And _
                                                                                                                                                    lngX <> lngR And lngX <> lngS And _
                                                                                                                                                    lngX <> lngT And lngX <> lngU And _
                                                                                                                                                    lngX <> lngV And lngX <> lngW Then
                                                                                                                                                  dblBal = dblBal - arr_varTran(T_SHARE, lngX)
                                                                                                                                                  If dblBal <= -0.0002 Then
                                                                                                                                                    dblBal = dblBal + arr_varTran(T_SHARE, lngX)
                                                                                                                                                    Exit For
                                                                                                                                                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                    blnFound = True
                                                                                                                                                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngX)
                                                                                                                                                    Stop
                                                                                                                                                  Else
                                                                                                                                                    For lngY = 0& To (lngTrans - 1&)
                                                                                                                                                      If lngY <> lngDepElem And lngY <> lngA And _
                                                                                                                                                          lngY <> lngB And lngY <> lngC And _
                                                                                                                                                          lngY <> lngD And lngY <> lngE And _
                                                                                                                                                          lngY <> lngF And lngY <> lngG And _
                                                                                                                                                          lngY <> lngH And lngY <> lngI And _
                                                                                                                                                          lngY <> lngJ And lngY <> lngK And _
                                                                                                                                                          lngY <> lngL And lngY <> lngM And _
                                                                                                                                                          lngY <> lngN And lngY <> lngO And _
                                                                                                                                                          lngY <> lngP And lngY <> lngQ And _
                                                                                                                                                          lngY <> lngR And lngY <> lngS And _
                                                                                                                                                          lngY <> lngT And lngY <> lngU And _
                                                                                                                                                          lngY <> lngV And lngY <> lngW And _
                                                                                                                                                          lngY <> lngX Then
                                                                                                                                                        dblBal = dblBal - arr_varTran(T_SHARE, lngY)
                                                                                                                                                        If dblBal <= -0.0002 Then
                                                                                                                                                          dblBal = dblBal + arr_varTran(T_SHARE, lngY)
                                                                                                                                                          Exit For
                                                                                                                                                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                          blnFound = True
                                                                                                                                                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngY)
                                                                                                                                                          Stop
                                                                                                                                                        Else
                                                                                                                                                          For lngZ = 0& To (lngTrans - 1&)
                                                                                                                                                            If lngZ <> lngDepElem And lngZ <> lngA And _
                                                                                                                                                                lngZ <> lngB And lngZ <> lngC And _
                                                                                                                                                                lngZ <> lngD And lngZ <> lngE And _
                                                                                                                                                                lngZ <> lngF And lngZ <> lngG And _
                                                                                                                                                                lngZ <> lngH And lngZ <> lngI And _
                                                                                                                                                                lngZ <> lngJ And lngZ <> lngK And _
                                                                                                                                                                lngZ <> lngL And lngZ <> lngM And _
                                                                                                                                                                lngZ <> lngN And lngZ <> lngO And _
                                                                                                                                                                lngZ <> lngP And lngZ <> lngQ And _
                                                                                                                                                                lngZ <> lngR And lngZ <> lngS And _
                                                                                                                                                                lngZ <> lngT And lngZ <> lngU And _
                                                                                                                                                                lngZ <> lngV And lngZ <> lngW And _
                                                                                                                                                                lngZ <> lngX And lngZ <> lngY Then
                                                                                                                                                              dblBal = dblBal - arr_varTran(T_SHARE, lngZ)
                                                                                                                                                              If dblBal <= -0.0002 Then
                                                                                                                                                                dblBal = dblBal + arr_varTran(T_SHARE, lngZ)
                                                                                                                                                                Exit For
                                                                                                                                                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                                blnFound = True
                                                                                                                                                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngZ)
                                                                                                                                                                Stop
                                                                                                                                                              Else
                                                                                                                                                                'To be continued.
                                                                                                                                                                lngElems = 26&
                                                                                                                                                                ReDim arr_varElem(E_ELEMS, lngElems)  ' ** We'll ignore Zero.
                                                                                                                                                                arr_varElem(E_VAL, 1) = lngA: arr_varElem(E_VAL, 2) = lngB
                                                                                                                                                                arr_varElem(E_VAL, 3) = lngC: arr_varElem(E_VAL, 4) = lngD
                                                                                                                                                                arr_varElem(E_VAL, 5) = lngE: arr_varElem(E_VAL, 6) = lngF
                                                                                                                                                                arr_varElem(E_VAL, 7) = lngG: arr_varElem(E_VAL, 8) = lngH
                                                                                                                                                                arr_varElem(E_VAL, 9) = lngI: arr_varElem(E_VAL, 10) = lngJ
                                                                                                                                                                arr_varElem(E_VAL, 11) = lngK: arr_varElem(E_VAL, 12) = lngL
                                                                                                                                                                arr_varElem(E_VAL, 13) = lngM: arr_varElem(E_VAL, 14) = lngN
                                                                                                                                                                arr_varElem(E_VAL, 15) = lngO: arr_varElem(E_VAL, 16) = lngP
                                                                                                                                                                arr_varElem(E_VAL, 17) = lngQ: arr_varElem(E_VAL, 18) = lngR
                                                                                                                                                                arr_varElem(E_VAL, 19) = lngS: arr_varElem(E_VAL, 20) = lngT
                                                                                                                                                                arr_varElem(E_VAL, 21) = lngU: arr_varElem(E_VAL, 22) = lngV
                                                                                                                                                                arr_varElem(E_VAL, 23) = lngW: arr_varElem(E_VAL, 24) = lngX
                                                                                                                                                                arr_varElem(E_VAL, 25) = lngY: arr_varElem(E_VAL, 26) = lngZ
                                                                                                                                                                blnRetVal = NegTaxLots_Match2(dblBal, lngTrans, arr_varTran, blnFound, lngDepElem, arr_varElem)
                                                                                                                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then
                                                                                                                                                                  Exit For
                                                                                                                                                                Else
                                                                                                                                                                  'To be continued.
                                                                                                                                                                End If
                                                                                                                                                              End If  ' ** dblBal.
                                                                                                                                                            End If  ' ** lngDepElem, lngA - lngY.
                                                                                                                                                          Next  ' ** lngZ.
                                                                                                                                                        End If  ' ** dblBal.
                                                                                                                                                      End If  ' ** lngDepElem, lngA - lngX.
                                                                                                                                                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                                    Next  ' ** lngY.
                                                                                                                                                  End If  ' ** dblBal.
                                                                                                                                                End If  ' ** lngDepElem, lngA - lngW.
                                                                                                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                              Next  ' ** lngX.
                                                                                                                                            End If  ' ** dblBal.
                                                                                                                                          End If  ' ** lngDepElem, lngA - lngV.
                                                                                                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                        Next  ' ** lngW.
                                                                                                                                      End If  ' ** dblBal.
                                                                                                                                    End If  ' ** lngDepElem, lngA - lngU.
                                                                                                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                  Next  ' ** lngV.
                                                                                                                                End If  ' ** dblBal.
                                                                                                                              End If  ' ** lngDepElem, lngA - lngT.
                                                                                                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                            Next  ' ** lngU.
                                                                                                                          End If  ' ** dblBal.
                                                                                                                        End If  ' ** lngDepElem, lngA - lngS.
                                                                                                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                      Next  ' ** lngT.
                                                                                                                    End If  ' ** dblBal.
                                                                                                                  End If  ' ** lngDepElem, lngA - lngR.
                                                                                                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                Next  ' ** lngS.
                                                                                                              End If  ' ** dblBal.
                                                                                                            End If  ' ** lngDepElem, lngA - lngQ.
                                                                                                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                          Next  ' ** lngR.
                                                                                                        End If  ' ** dblBal.
                                                                                                      End If  ' ** lngDepElem, lngA - lngP.
                                                                                                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                    Next  ' ** lngQ.
                                                                                                  End If  ' ** dblBal.
                                                                                                End If  ' ** lngDepElem, lngA - lngO.
                                                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                              Next  ' ** lngP.
                                                                                            End If  ' ** dblBal.
                                                                                          End If  ' ** lngDepElem, lngA - lngN.
                                                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                        Next  ' ** lngO.
                                                                                      End If  ' ** dblBal.
                                                                                    End If  ' ** lngDepElem, lngA - lngM.
                                                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                  Next  ' ** lngN.
                                                                                End If  ' ** dblBal.
                                                                              End If  ' ** lngDepElem, lngA - lngL.
                                                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                            Next  ' ** lngM.
                                                                          End If  ' ** dblBal.
                                                                        End If  ' ** lngDepElem, lngA, lngB, lngC, lngD, lngE, lngF, lngG, lngH, lngI, lngJ, lngK.
                                                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                      Next  ' ** lngL.
                                                                    End If  ' ** dblBal.
                                                                  End If  ' ** lngDepElem, lngA, lngB, lngC, lngD, lngE, lngF, lngG, lngH, lngI, lngJ.
                                                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                Next  ' ** lngK.
                                                              End If  ' ** dblBal.
                                                            End If  ' ** lngDepElem, lngA, lngB, lngC, lngD, lngE, lngF, lngG, lngH, lngI.
                                                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                          Next  ' ** lngJ.
                                                        End If  ' ** dblBal.
                                                      End If  ' ** lngDepElem, lngA, lngB, lngC, lngD, lngE, lngF, lngG, lngH.
                                                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                    Next  ' ** lngI.
                                                  End If  ' ** dblBal.
                                                End If  ' ** lngDepElem, lngA, lngB, lngC, lngD, lngE, lngF, lngG.
                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                              Next  ' ** lngH.
                                            End If  ' ** dblBal.
                                          End If  ' ** lngDepElem, lngA, lngB, lngC, lngD, lngE, lngF.
                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                        Next  ' ** lngG.
                                      End If  ' ** dblBal.
                                    End If  ' ** lngDepElem, lngA, lngB, lngC, lngD, lngE.
                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                  Next  ' ** lngF.
                                End If  ' ** dblBal.
                              End If  ' ** lngDepElem, lngA, lngB, lngC, lngD.
                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                            Next  ' ** lngE.
                          End If  ' ** dblBal.
                        End If  ' ** lngDepElem, lngA, lngB, lngC.
                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                      Next  ' ** lngD.
                    End If  ' ** dblBal.
                  End If  ' ** lngDepElem, lngA, lngB.
                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                Next  ' ** lngC.
              End If  ' ** dblBal.
            End If  ' ** lngDepElem, lngA.
            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
          Next  ' ** lngB.
        End If  ' ** dblBal.
      End If  ' ** lngDepElem.
      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
    Next  ' ** lngA.

  End If  ' ** lngDepElem.

'If dblBal <= -0.0002 Then
'  Exit For
'ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
'  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngX)
'  Stop
'End If

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  NegTaxLots_Match1 = blnRetVal

End Function

Public Function NegTaxLots_Match2(ByRef dblBal As Double, ByRef lngTrans As Long, ByRef arr_varTran As Variant, ByRef blnFound As Boolean, ByRef lngDepElem As Long, ByRef arr_varElem As Variant) As Boolean

  Const THIS_PROC As String = "NegTaxLots_Match2"

  Dim lngElems2 As Long, arr_varElem2() As Variant
  Dim blnElem As Boolean
  Dim lngX As Long
  Dim lngA1 As Long, lngB1 As Long, lngC1 As Long, lngD1 As Long, lngE1 As Long, lngF1 As Long, lngG1 As Long, lngH1 As Long, lngI1 As Long
  Dim lngJ1 As Long, lngK1 As Long, lngL1 As Long, lngM1 As Long, lngN1 As Long, lngO1 As Long, lngP1 As Long, lngQ1 As Long, lngR1 As Long
  Dim lngS1 As Long, lngT1 As Long, lngU1 As Long, lngV1 As Long, lngW1 As Long, lngX1 As Long, lngY1 As Long, lngZ1 As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTran().
  Const T_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const T_JNO   As Integer = 0
  Const T_JTYP  As Integer = 1
  Const T_ACTNO As Integer = 2
  Const T_ASTNO As Integer = 3
  Const T_SHARE As Integer = 4

  ' ** Array: arr_varElem().
  Const E_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const E_VAL As Integer = 0

  blnRetVal = True

  For lngA1 = 0& To (lngTrans - 1&)
    blnElem = False
    For lngX = 1& To 26&
      If lngA1 = arr_varElem(E_VAL, lngX) Then
        blnElem = True
        Exit For
      End If
    Next  ' ** lngX.
    If blnElem = False And lngA1 <> lngDepElem Then
      dblBal = dblBal - arr_varTran(T_SHARE, lngA1)
      If dblBal <= -0.0002 Then
        dblBal = dblBal + arr_varTran(T_SHARE, lngA1)
        Exit For
      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
        blnFound = True
        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngA1)
        Stop
      Else
        For lngB1 = 0& To (lngTrans - 1&)
          For lngX = 1& To 26&
            If lngB1 = arr_varElem(E_VAL, lngX) Then
              blnElem = True
              Exit For
            End If
          Next  ' ** lngX.
          If blnElem = False And lngB1 <> lngDepElem And lngB1 <> lngA1 Then
            dblBal = dblBal - arr_varTran(T_SHARE, lngB1)
            If dblBal <= -0.0002 Then
              dblBal = dblBal + arr_varTran(T_SHARE, lngB1)
              Exit For
            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
              blnFound = True
              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngB1)
              Stop
            Else
              For lngC1 = 0& To (lngTrans - 1&)
                For lngX = 1& To 26&
                  If lngC1 = arr_varElem(E_VAL, lngX) Then
                    blnElem = True
                    Exit For
                  End If
                Next  ' ** lngX.
                If blnElem = False And lngC1 <> lngDepElem And lngC1 <> lngA1 And lngC1 <> lngB1 Then
                  dblBal = dblBal - arr_varTran(T_SHARE, lngC1)
                  If dblBal <= -0.0002 Then
                    dblBal = dblBal + arr_varTran(T_SHARE, lngC1)
                    Exit For
                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                    blnFound = True
                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngC1)
                    Stop
                  Else
                    For lngD1 = 0& To (lngTrans - 1&)
                      For lngX = 1& To 26&
                        If lngD1 = arr_varElem(E_VAL, lngX) Then
                          blnElem = True
                          Exit For
                        End If
                      Next  ' ** lngX.
                      If blnElem = False And lngD1 <> lngDepElem And lngD1 <> lngA1 And lngD1 <> lngB1 And lngD1 <> lngC1 Then
                        dblBal = dblBal - arr_varTran(T_SHARE, lngD1)
                        If dblBal <= -0.0002 Then
                          dblBal = dblBal + arr_varTran(T_SHARE, lngD1)
                          Exit For
                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                          blnFound = True
                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngD1)
                          Stop
                        Else
                          For lngE1 = 0& To (lngTrans - 1&)
                            For lngX = 1& To 26&
                              If lngE1 = arr_varElem(E_VAL, lngX) Then
                                blnElem = True
                                Exit For
                              End If
                            Next  ' ** lngX.
                            If blnElem = False And lngE1 <> lngDepElem And lngE1 <> lngA1 And lngE1 <> lngB1 And lngE1 <> lngC1 And _
                                lngE1 <> lngD1 Then
                              dblBal = dblBal - arr_varTran(T_SHARE, lngE1)
                              If dblBal <= -0.0002 Then
                                dblBal = dblBal + arr_varTran(T_SHARE, lngE1)
                                Exit For
                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                blnFound = True
                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngE1)
                                Stop
                              Else
                                For lngF1 = 0& To (lngTrans - 1&)
                                  For lngX = 1& To 26&
                                    If lngF1 = arr_varElem(E_VAL, lngX) Then
                                      blnElem = True
                                      Exit For
                                    End If
                                  Next  ' ** lngX.
                                  If blnElem = False And lngF1 <> lngDepElem And lngF1 <> lngA1 And lngF1 <> lngB1 And lngF1 <> lngC1 And _
                                      lngF1 <> lngD1 And lngF1 <> lngE1 Then
                                    dblBal = dblBal - arr_varTran(T_SHARE, lngF1)
                                    If dblBal <= -0.0002 Then
                                      dblBal = dblBal + arr_varTran(T_SHARE, lngF1)
                                      Exit For
                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                      blnFound = True
                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngF1)
                                      Stop
                                    Else
                                      For lngG1 = 0& To (lngTrans - 1&)
                                        For lngX = 1& To 26&
                                          If lngG1 = arr_varElem(E_VAL, lngX) Then
                                            blnElem = True
                                            Exit For
                                          End If
                                        Next  ' ** lngX.
                                        If blnElem = False And lngG1 <> lngDepElem And lngG1 <> lngA1 And lngG1 <> lngB1 And lngG1 <> lngC1 And _
                                            lngG1 <> lngD1 And lngG1 <> lngE1 And lngG1 <> lngF1 Then
                                          dblBal = dblBal - arr_varTran(T_SHARE, lngG1)
                                          If dblBal <= -0.0002 Then
                                            dblBal = dblBal + arr_varTran(T_SHARE, lngG1)
                                            Exit For
                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                            blnFound = True
                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngG1)
                                            Stop
                                          Else
                                            For lngH1 = 0& To (lngTrans - 1&)
                                              For lngX = 1& To 26&
                                                If lngH1 = arr_varElem(E_VAL, lngX) Then
                                                  blnElem = True
                                                  Exit For
                                                End If
                                              Next  ' ** lngX.
                                              If blnElem = False And lngH1 <> lngDepElem And lngH1 <> lngA1 And lngH1 <> lngB1 And _
                                                  lngH1 <> lngC1 And lngH1 <> lngD1 And lngH1 <> lngE1 And lngH1 <> lngF1 And lngH1 <> lngG1 Then
                                                dblBal = dblBal - arr_varTran(T_SHARE, lngH1)
                                                If dblBal <= -0.0002 Then
                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngH1)
                                                  Exit For
                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                  blnFound = True
                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngH1)
                                                  Stop
                                                Else
                                                  For lngI1 = 0& To (lngTrans - 1&)
                                                    For lngX = 1& To 26&
                                                      If lngI1 = arr_varElem(E_VAL, lngX) Then
                                                        blnElem = True
                                                        Exit For
                                                      End If
                                                    Next  ' ** lngX.
                                                    If blnElem = False And lngI1 <> lngDepElem And lngI1 <> lngA1 And lngI1 <> lngB1 And _
                                                        lngI1 <> lngC1 And lngI1 <> lngD1 And lngI1 <> lngE1 And lngI1 <> lngF1 And _
                                                        lngI1 <> lngG1 And lngI1 <> lngH1 Then
                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngI1)
                                                      If dblBal <= -0.0002 Then
                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngI1)
                                                        Exit For
                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                        blnFound = True
                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngI1)
                                                        Stop
                                                      Else
                                                        For lngJ1 = 0& To (lngTrans - 1&)
                                                          For lngX = 1& To 26&
                                                            If lngJ1 = arr_varElem(E_VAL, lngX) Then
                                                              blnElem = True
                                                              Exit For
                                                            End If
                                                          Next  ' ** lngX.
                                                          If blnElem = False And lngJ1 <> lngDepElem And lngJ1 <> lngA1 And lngJ1 <> lngB1 And _
                                                              lngJ1 <> lngC1 And lngJ1 <> lngD1 And lngJ1 <> lngE1 And lngJ1 <> lngF1 And _
                                                              lngJ1 <> lngG1 And lngJ1 <> lngH1 And lngJ1 <> lngI1 Then
                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngJ1)
                                                            If dblBal <= -0.0002 Then
                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngJ1)
                                                              Exit For
                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                              blnFound = True
                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngJ1)
                                                              Stop
                                                            Else
                                                              For lngK1 = 0& To (lngTrans - 1&)
                                                                For lngX = 1& To 26&
                                                                  If lngK1 = arr_varElem(E_VAL, lngX) Then
                                                                    blnElem = True
                                                                    Exit For
                                                                  End If
                                                                Next  ' ** lngX.
                                                                If blnElem = False And lngK1 <> lngDepElem And lngK1 <> lngA1 And lngK1 <> lngB1 And _
                                                                    lngK1 <> lngC1 And lngK1 <> lngD1 And lngK1 <> lngE1 And lngK1 <> lngF1 And _
                                                                    lngK1 <> lngG1 And lngK1 <> lngH1 And lngK1 <> lngI1 And lngK1 <> lngJ1 Then
                                                                  dblBal = dblBal - arr_varTran(T_SHARE, lngK1)
                                                                  If dblBal <= -0.0002 Then
                                                                    dblBal = dblBal + arr_varTran(T_SHARE, lngK1)
                                                                    Exit For
                                                                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                    blnFound = True
                                                                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngK1)
                                                                    Stop
                                                                  Else
                                                                    For lngL1 = 0& To (lngTrans - 1&)
                                                                      For lngX = 1& To 26&
                                                                        If lngL1 = arr_varElem(E_VAL, lngX) Then
                                                                          blnElem = True
                                                                          Exit For
                                                                        End If
                                                                      Next  ' ** lngX.
                                                                      If blnElem = False And lngL1 <> lngDepElem And lngL1 <> lngA1 And _
                                                                          lngL1 <> lngB1 And lngL1 <> lngC1 And lngL1 <> lngD1 And _
                                                                          lngL1 <> lngE1 And lngL1 <> lngF1 And lngL1 <> lngG1 And _
                                                                          lngL1 <> lngH1 And lngL1 <> lngI1 And lngL1 <> lngJ1 And _
                                                                          lngL1 <> lngK1 Then
                                                                        dblBal = dblBal - arr_varTran(T_SHARE, lngL1)
                                                                        If dblBal <= -0.0002 Then
                                                                          dblBal = dblBal + arr_varTran(T_SHARE, lngL1)
                                                                          Exit For
                                                                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                          blnFound = True
                                                                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngL1)
                                                                          Stop
                                                                        Else
                                                                          For lngM1 = 0& To (lngTrans - 1&)
                                                                            For lngX = 1& To 26&
                                                                              If lngM1 = arr_varElem(E_VAL, lngX) Then
                                                                                blnElem = True
                                                                                Exit For
                                                                              End If
                                                                            Next  ' ** lngX.
                                                                            If blnElem = False And lngM1 <> lngDepElem And lngM1 <> lngA1 And _
                                                                                lngM1 <> lngB1 And lngM1 <> lngC1 And lngM1 <> lngD1 And _
                                                                                lngM1 <> lngE1 And lngM1 <> lngF1 And lngM1 <> lngG1 And _
                                                                                lngM1 <> lngH1 And lngM1 <> lngI1 And lngM1 <> lngJ1 And _
                                                                                lngM1 <> lngK1 And lngM1 <> lngL1 Then
                                                                              dblBal = dblBal - arr_varTran(T_SHARE, lngM1)
                                                                              If dblBal <= -0.0002 Then
                                                                                dblBal = dblBal + arr_varTran(T_SHARE, lngM1)
                                                                                Exit For
                                                                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                blnFound = True
                                                                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngM1)
                                                                                Stop
                                                                              Else
                                                                                For lngN1 = 0& To (lngTrans - 1&)
                                                                                  For lngX = 1& To 26&
                                                                                    If lngN1 = arr_varElem(E_VAL, lngX) Then
                                                                                      blnElem = True
                                                                                      Exit For
                                                                                    End If
                                                                                  Next  ' ** lngX.
                                                                                  If blnElem = False And lngN1 <> lngDepElem And lngN1 <> lngA1 And _
                                                                                      lngN1 <> lngB1 And lngN1 <> lngC1 And lngN1 <> lngD1 And _
                                                                                      lngN1 <> lngE1 And lngN1 <> lngF1 And lngN1 <> lngG1 And _
                                                                                      lngN1 <> lngH1 And lngN1 <> lngI1 And lngN1 <> lngJ1 And _
                                                                                      lngN1 <> lngK1 And lngN1 <> lngL1 And lngN1 <> lngM1 Then
                                                                                    dblBal = dblBal - arr_varTran(T_SHARE, lngN1)
                                                                                    If dblBal <= -0.0002 Then
                                                                                      dblBal = dblBal + arr_varTran(T_SHARE, lngN1)
                                                                                      Exit For
                                                                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                      blnFound = True
                                                                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngN1)
                                                                                      Stop
                                                                                    Else
                                                                                      For lngO1 = 0& To (lngTrans - 1&)
                                                                                        For lngX = 1& To 26&
                                                                                          If lngO1 = arr_varElem(E_VAL, lngX) Then
                                                                                            blnElem = True
                                                                                            Exit For
                                                                                          End If
                                                                                        Next  ' ** lngX.
                                                                                        If blnElem = False And lngO1 <> lngDepElem And _
                                                                                            lngO1 <> lngA1 And lngO1 <> lngB1 And lngO1 <> lngC1 And _
                                                                                            lngO1 <> lngD1 And lngO1 <> lngE1 And lngO1 <> lngF1 And _
                                                                                            lngO1 <> lngG1 And lngO1 <> lngH1 And lngO1 <> lngI1 And _
                                                                                            lngO1 <> lngJ1 And lngO1 <> lngK1 And lngO1 <> lngL1 And _
                                                                                            lngO1 <> lngM1 And lngO1 <> lngN1 Then
                                                                                          dblBal = dblBal - arr_varTran(T_SHARE, lngO1)
                                                                                          If dblBal <= -0.0002 Then
                                                                                            dblBal = dblBal + arr_varTran(T_SHARE, lngO1)
                                                                                            Exit For
                                                                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                            blnFound = True
                                                                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngO1)
                                                                                            Stop
                                                                                          Else
                                                                                            For lngP1 = 0& To (lngTrans - 1&)
                                                                                              For lngX = 1& To 26&
                                                                                                If lngP1 = arr_varElem(E_VAL, lngX) Then
                                                                                                  blnElem = True
                                                                                                  Exit For
                                                                                                End If
                                                                                              Next  ' ** lngX.
                                                                                              If blnElem = False And lngP1 <> lngDepElem And _
                                                                                                  lngP1 <> lngA1 And lngP1 <> lngB1 And lngP1 <> lngC1 And _
                                                                                                  lngP1 <> lngD1 And lngP1 <> lngE1 And lngP1 <> lngF1 And _
                                                                                                  lngP1 <> lngG1 And lngP1 <> lngH1 And lngP1 <> lngI1 And _
                                                                                                  lngP1 <> lngJ1 And lngP1 <> lngK1 And lngP1 <> lngL1 And _
                                                                                                  lngP1 <> lngM1 And lngP1 <> lngN1 And lngP1 <> lngO1 Then
                                                                                                dblBal = dblBal - arr_varTran(T_SHARE, lngP1)
                                                                                                If dblBal <= -0.0002 Then
                                                                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngP1)
                                                                                                  Exit For
                                                                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                  blnFound = True
                                                                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngP1)
                                                                                                  Stop
                                                                                                Else
                                                                                                  For lngQ1 = 0& To (lngTrans - 1&)
                                                                                                    For lngX = 1& To 26&
                                                                                                      If lngQ1 = arr_varElem(E_VAL, lngX) Then
                                                                                                        blnElem = True
                                                                                                        Exit For
                                                                                                      End If
                                                                                                    Next  ' ** lngX.
                                                                                                    If blnElem = False And lngQ1 <> lngDepElem And _
                                                                                                        lngQ1 <> lngA1 And lngQ1 <> lngB1 And lngQ1 <> lngC1 And _
                                                                                                        lngQ1 <> lngD1 And lngQ1 <> lngE1 And lngQ1 <> lngF1 And _
                                                                                                        lngQ1 <> lngG1 And lngQ1 <> lngH1 And lngQ1 <> lngI1 And _
                                                                                                        lngQ1 <> lngJ1 And lngQ1 <> lngK1 And lngQ1 <> lngL1 And _
                                                                                                        lngQ1 <> lngM1 And lngQ1 <> lngN1 And lngQ1 <> lngO1 And _
                                                                                                        lngQ1 <> lngP1 Then
                                                                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngQ1)
                                                                                                      If dblBal <= -0.0002 Then
                                                                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngQ1)
                                                                                                        Exit For
                                                                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                        blnFound = True
                                                                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngQ1)
                                                                                                        Stop
                                                                                                      Else
                                                                                                        For lngR1 = 0& To (lngTrans - 1&)
                                                                                                          For lngX = 1& To 26&
                                                                                                            If lngR1 = arr_varElem(E_VAL, lngX) Then
                                                                                                              blnElem = True
                                                                                                              Exit For
                                                                                                            End If
                                                                                                          Next  ' ** lngX.
                                                                                                          If blnElem = False And lngR1 <> lngDepElem And _
                                                                                                              lngR1 <> lngA1 And lngR1 <> lngB1 And lngR1 <> lngC1 And _
                                                                                                              lngR1 <> lngD1 And lngR1 <> lngE1 And lngR1 <> lngF1 And _
                                                                                                              lngR1 <> lngG1 And lngR1 <> lngH1 And lngR1 <> lngI1 And _
                                                                                                              lngR1 <> lngJ1 And lngR1 <> lngK1 And lngR1 <> lngL1 And _
                                                                                                              lngR1 <> lngM1 And lngR1 <> lngN1 And lngR1 <> lngO1 And _
                                                                                                              lngR1 <> lngP1 And lngR1 <> lngQ1 Then
                                                                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngR1)
                                                                                                            If dblBal <= -0.0002 Then
                                                                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngR1)
                                                                                                              Exit For
                                                                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                              blnFound = True
                                                                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngR1)
                                                                                                              Stop
                                                                                                            Else
                                                                                                              For lngS1 = 0& To (lngTrans - 1&)
                                                                                                                For lngX = 1& To 26&
                                                                                                                  If lngS1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                    blnElem = True
                                                                                                                    Exit For
                                                                                                                  End If
                                                                                                                Next  ' ** lngX.
                                                                                                                If blnElem = False And lngS1 <> lngDepElem And _
                                                                                                                    lngS1 <> lngA1 And lngS1 <> lngB1 And lngS1 <> lngC1 And _
                                                                                                                    lngS1 <> lngD1 And lngS1 <> lngE1 And lngS1 <> lngF1 And _
                                                                                                                    lngS1 <> lngG1 And lngS1 <> lngH1 And lngS1 <> lngI1 And _
                                                                                                                    lngS1 <> lngJ1 And lngS1 <> lngK1 And lngS1 <> lngL1 And _
                                                                                                                    lngS1 <> lngM1 And lngS1 <> lngN1 And lngS1 <> lngO1 And _
                                                                                                                    lngS1 <> lngP1 And lngS1 <> lngQ1 And lngS1 <> lngR1 Then
                                                                                                                  dblBal = dblBal - arr_varTran(T_SHARE, lngS1)
                                                                                                                  If dblBal <= -0.0002 Then
                                                                                                                    dblBal = dblBal + arr_varTran(T_SHARE, lngS1)
                                                                                                                    Exit For
                                                                                                                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                    blnFound = True
                                                                                                                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngS1)
                                                                                                                    Stop
                                                                                                                  Else
                                                                                                                    For lngT1 = 0& To (lngTrans - 1&)
                                                                                                                      For lngX = 1& To 26&
                                                                                                                        If lngT1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                          blnElem = True
                                                                                                                          Exit For
                                                                                                                        End If
                                                                                                                      Next  ' ** lngX.
                                                                                                                      If blnElem = False And lngT1 <> lngDepElem And _
                                                                                                                          lngT1 <> lngA1 And lngT1 <> lngB1 And lngT1 <> lngC1 And _
                                                                                                                          lngT1 <> lngD1 And lngT1 <> lngE1 And lngT1 <> lngF1 And _
                                                                                                                          lngT1 <> lngG1 And lngT1 <> lngH1 And lngT1 <> lngI1 And _
                                                                                                                          lngT1 <> lngJ1 And lngT1 <> lngK1 And lngT1 <> lngL1 And _
                                                                                                                          lngT1 <> lngM1 And lngT1 <> lngN1 And lngT1 <> lngO1 And _
                                                                                                                          lngT1 <> lngP1 And lngT1 <> lngQ1 And lngT1 <> lngR1 And _
                                                                                                                          lngT1 <> lngS1 Then
                                                                                                                        dblBal = dblBal - arr_varTran(T_SHARE, lngT1)
                                                                                                                        If dblBal <= -0.0002 Then
                                                                                                                          dblBal = dblBal + arr_varTran(T_SHARE, lngT1)
                                                                                                                          Exit For
                                                                                                                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                          blnFound = True
                                                                                                                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngT1)
                                                                                                                          Stop
                                                                                                                        Else
                                                                                                                          For lngU1 = 0& To (lngTrans - 1&)
                                                                                                                            For lngX = 1& To 26&
                                                                                                                              If lngU1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                                blnElem = True
                                                                                                                                Exit For
                                                                                                                              End If
                                                                                                                            Next  ' ** lngX.
                                                                                                                            If blnElem = False And lngU1 <> lngDepElem And _
                                                                                                                                lngU1 <> lngA1 And lngU1 <> lngB1 And lngU1 <> lngC1 And _
                                                                                                                                lngU1 <> lngD1 And lngU1 <> lngE1 And lngU1 <> lngF1 And _
                                                                                                                                lngU1 <> lngG1 And lngU1 <> lngH1 And lngU1 <> lngI1 And _
                                                                                                                                lngU1 <> lngJ1 And lngU1 <> lngK1 And lngU1 <> lngL1 And _
                                                                                                                                lngU1 <> lngM1 And lngU1 <> lngN1 And lngU1 <> lngO1 And _
                                                                                                                                lngU1 <> lngP1 And lngU1 <> lngQ1 And lngU1 <> lngR1 And _
                                                                                                                                lngU1 <> lngS1 And lngU1 <> lngT1 Then
                                                                                                                              dblBal = dblBal - arr_varTran(T_SHARE, lngU1)
                                                                                                                              If dblBal <= -0.0002 Then
                                                                                                                                dblBal = dblBal + arr_varTran(T_SHARE, lngU1)
                                                                                                                                Exit For
                                                                                                                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                blnFound = True
                                                                                                                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngU1)
                                                                                                                                Stop
                                                                                                                              Else
                                                                                                                                For lngV1 = 0& To (lngTrans - 1&)
                                                                                                                                  For lngX = 1& To 26&
                                                                                                                                    If lngV1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                                      blnElem = True
                                                                                                                                      Exit For
                                                                                                                                    End If
                                                                                                                                  Next  ' ** lngX.
                                                                                                                                  If blnElem = False And lngV1 <> lngDepElem And _
                                                                                                                                      lngV1 <> lngA1 And lngV1 <> lngB1 And lngV1 <> lngC1 And _
                                                                                                                                      lngV1 <> lngD1 And lngV1 <> lngE1 And lngV1 <> lngF1 And _
                                                                                                                                      lngV1 <> lngG1 And lngV1 <> lngH1 And lngV1 <> lngI1 And _
                                                                                                                                      lngV1 <> lngJ1 And lngV1 <> lngK1 And lngV1 <> lngL1 And _
                                                                                                                                      lngV1 <> lngM1 And lngV1 <> lngN1 And lngV1 <> lngO1 And _
                                                                                                                                      lngV1 <> lngP1 And lngV1 <> lngQ1 And lngV1 <> lngR1 And _
                                                                                                                                      lngV1 <> lngS1 And lngV1 <> lngT1 And lngV1 <> lngU1 Then
                                                                                                                                    dblBal = dblBal - arr_varTran(T_SHARE, lngV1)
                                                                                                                                    If dblBal <= -0.0002 Then
                                                                                                                                      dblBal = dblBal + arr_varTran(T_SHARE, lngV1)
                                                                                                                                      Exit For
                                                                                                                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                      blnFound = True
                                                                                                                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngV1)
                                                                                                                                      Stop
                                                                                                                                    Else
                                                                                                                                      For lngW1 = 0& To (lngTrans - 1&)
                                                                                                                                        For lngX = 1& To 26&
                                                                                                                                          If lngW1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                                            blnElem = True
                                                                                                                                            Exit For
                                                                                                                                          End If
                                                                                                                                        Next  ' ** lngX.
                                                                                                                                        If blnElem = False And lngW1 <> lngDepElem And _
                                                                                                                                            lngW1 <> lngA1 And lngW1 <> lngB1 And lngW1 <> lngC1 And _
                                                                                                                                            lngW1 <> lngD1 And lngW1 <> lngE1 And lngW1 <> lngF1 And _
                                                                                                                                            lngW1 <> lngG1 And lngW1 <> lngH1 And lngW1 <> lngI1 And _
                                                                                                                                            lngW1 <> lngJ1 And lngW1 <> lngK1 And lngW1 <> lngL1 And _
                                                                                                                                            lngW1 <> lngM1 And lngW1 <> lngN1 And lngW1 <> lngO1 And _
                                                                                                                                            lngW1 <> lngP1 And lngW1 <> lngQ1 And lngW1 <> lngR1 And _
                                                                                                                                            lngW1 <> lngS1 And lngW1 <> lngT1 And lngW1 <> lngU1 And _
                                                                                                                                            lngW1 <> lngV1 Then
                                                                                                                                          dblBal = dblBal - arr_varTran(T_SHARE, lngW1)
                                                                                                                                          If dblBal <= -0.0002 Then
                                                                                                                                            dblBal = dblBal + arr_varTran(T_SHARE, lngW1)
                                                                                                                                            Exit For
                                                                                                                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                            blnFound = True
                                                                                                                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngW1)
                                                                                                                                            Stop
                                                                                                                                          Else
                                                                                                                                            For lngX1 = 0& To (lngTrans - 1&)
                                                                                                                                              For lngX = 1& To 26&
                                                                                                                                                If lngX1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                                                  blnElem = True
                                                                                                                                                  Exit For
                                                                                                                                                End If
                                                                                                                                              Next  ' ** lngX.
                                                                                                                                              If blnElem = False And lngX1 <> lngDepElem And _
                                                                                                                                                  lngX1 <> lngA1 And lngX1 <> lngB1 And lngX1 <> lngC1 And _
                                                                                                                                                  lngX1 <> lngD1 And lngX1 <> lngE1 And lngX1 <> lngF1 And _
                                                                                                                                                  lngX1 <> lngG1 And lngX1 <> lngH1 And lngX1 <> lngI1 And _
                                                                                                                                                  lngX1 <> lngJ1 And lngX1 <> lngK1 And lngX1 <> lngL1 And _
                                                                                                                                                  lngX1 <> lngM1 And lngX1 <> lngN1 And lngX1 <> lngO1 And _
                                                                                                                                                  lngX1 <> lngP1 And lngX1 <> lngQ1 And lngX1 <> lngR1 And _
                                                                                                                                                  lngX1 <> lngS1 And lngX1 <> lngT1 And lngX1 <> lngU1 And _
                                                                                                                                                  lngX1 <> lngV1 And lngX1 <> lngW1 Then
                                                                                                                                                dblBal = dblBal - arr_varTran(T_SHARE, lngX1)
                                                                                                                                                If dblBal <= -0.0002 Then
                                                                                                                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngX1)
                                                                                                                                                  Exit For
                                                                                                                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                  blnFound = True
                                                                                                                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngX1)
                                                                                                                                                  Stop
                                                                                                                                                Else
                                                                                                                                                  For lngY1 = 0& To (lngTrans - 1&)
                                                                                                                                                    For lngX = 1& To 26&
                                                                                                                                                      If lngY1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                                                        blnElem = True
                                                                                                                                                        Exit For
                                                                                                                                                      End If
                                                                                                                                                    Next  ' ** lngX.
                                                                                                                                                    If blnElem = False And lngY1 <> lngDepElem And _
                                                                                                                                                        lngY1 <> lngA1 And lngY1 <> lngB1 And lngY1 <> lngC1 And _
                                                                                                                                                        lngY1 <> lngD1 And lngY1 <> lngE1 And lngY1 <> lngF1 And _
                                                                                                                                                        lngY1 <> lngG1 And lngY1 <> lngH1 And lngY1 <> lngI1 And _
                                                                                                                                                        lngY1 <> lngJ1 And lngY1 <> lngK1 And lngY1 <> lngL1 And _
                                                                                                                                                        lngY1 <> lngM1 And lngY1 <> lngN1 And lngY1 <> lngO1 And _
                                                                                                                                                        lngY1 <> lngP1 And lngY1 <> lngQ1 And lngY1 <> lngR1 And _
                                                                                                                                                        lngY1 <> lngS1 And lngY1 <> lngT1 And lngY1 <> lngU1 And _
                                                                                                                                                        lngY1 <> lngV1 And lngY1 <> lngW1 And lngY1 <> lngX1 Then
                                                                                                                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngY1)
                                                                                                                                                      If dblBal <= -0.0002 Then
                                                                                                                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngY1)
                                                                                                                                                        Exit For
                                                                                                                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                        blnFound = True
                                                                                                                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngY1)
                                                                                                                                                        Stop
                                                                                                                                                      Else
                                                                                                                                                        For lngZ1 = 0& To (lngTrans - 1&)
                                                                                                                                                          For lngX = 1& To 26&
                                                                                                                                                            If lngZ1 = arr_varElem(E_VAL, lngX) Then
                                                                                                                                                              blnElem = True
                                                                                                                                                              Exit For
                                                                                                                                                            End If
                                                                                                                                                          Next  ' ** lngX.
                                                                                                                                                          If blnElem = False And lngZ1 <> lngDepElem And _
                                                                                                                                                              lngZ1 <> lngA1 And lngZ1 <> lngB1 And lngZ1 <> lngC1 And _
                                                                                                                                                              lngZ1 <> lngD1 And lngZ1 <> lngE1 And lngZ1 <> lngF1 And _
                                                                                                                                                              lngZ1 <> lngG1 And lngZ1 <> lngH1 And lngZ1 <> lngI1 And _
                                                                                                                                                              lngZ1 <> lngJ1 And lngZ1 <> lngK1 And lngZ1 <> lngL1 And _
                                                                                                                                                              lngZ1 <> lngM1 And lngZ1 <> lngN1 And lngZ1 <> lngO1 And _
                                                                                                                                                              lngZ1 <> lngP1 And lngZ1 <> lngQ1 And lngZ1 <> lngR1 And _
                                                                                                                                                              lngZ1 <> lngS1 And lngZ1 <> lngT1 And lngZ1 <> lngU1 And _
                                                                                                                                                              lngZ1 <> lngV1 And lngZ1 <> lngW1 And lngZ1 <> lngX1 And _
                                                                                                                                                              lngZ1 <> lngY1 Then
                                                                                                                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngZ1)
                                                                                                                                                            If dblBal <= -0.0002 Then
                                                                                                                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngZ1)
                                                                                                                                                              Exit For
                                                                                                                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                              blnFound = True
                                                                                                                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngZ1)
                                                                                                                                                              Stop
                                                                                                                                                            Else
                                                                                                                                                              'To be continued.
                                                                                                                                                                lngElems2 = 26&
                                                                                                                                                                ReDim arr_varElem2(E_ELEMS, lngElems2)  ' ** We'll ignore Zero.
                                                                                                                                                                arr_varElem2(E_VAL, 1) = lngA1: arr_varElem2(E_VAL, 2) = lngB1
                                                                                                                                                                arr_varElem2(E_VAL, 3) = lngC1: arr_varElem2(E_VAL, 4) = lngD1
                                                                                                                                                                arr_varElem2(E_VAL, 5) = lngE1: arr_varElem2(E_VAL, 6) = lngF1
                                                                                                                                                                arr_varElem2(E_VAL, 7) = lngG1: arr_varElem2(E_VAL, 8) = lngH1
                                                                                                                                                                arr_varElem2(E_VAL, 9) = lngI1: arr_varElem2(E_VAL, 10) = lngJ1
                                                                                                                                                                arr_varElem2(E_VAL, 11) = lngK1: arr_varElem2(E_VAL, 12) = lngL1
                                                                                                                                                                arr_varElem2(E_VAL, 13) = lngM1: arr_varElem2(E_VAL, 14) = lngN1
                                                                                                                                                                arr_varElem2(E_VAL, 15) = lngO1: arr_varElem2(E_VAL, 16) = lngP1
                                                                                                                                                                arr_varElem2(E_VAL, 17) = lngQ1: arr_varElem2(E_VAL, 18) = lngR1
                                                                                                                                                                arr_varElem2(E_VAL, 19) = lngS1: arr_varElem2(E_VAL, 20) = lngT1
                                                                                                                                                                arr_varElem2(E_VAL, 21) = lngU1: arr_varElem2(E_VAL, 22) = lngV1
                                                                                                                                                                arr_varElem2(E_VAL, 23) = lngW1: arr_varElem2(E_VAL, 24) = lngX1
                                                                                                                                                                arr_varElem2(E_VAL, 25) = lngY1: arr_varElem2(E_VAL, 26) = lngZ1
                                                                                                                                                                blnRetVal = NegTaxLots_Match3(dblBal, lngTrans, arr_varTran, blnFound, lngDepElem, arr_varElem, arr_varElem2)
                                                                                                                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then
                                                                                                                                                                  Exit For
                                                                                                                                                                Else
                                                                                                                                                                  'To be continued.
                                                                                                                                                              End If
                                                                                                                                                            End If  ' ** dblBal
                                                                                                                                                          End If  ' ** lngDepElem, lngA1 - lngY1.
                                                                                                                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                                        Next  ' ** lngZ1.
                                                                                                                                                      End If  ' ** dblBal
                                                                                                                                                    End If  ' ** lngDepElem, lngA1 - lngX1.
                                                                                                                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                                  Next  ' ** lngY1.
                                                                                                                                                End If  ' ** dblBal
                                                                                                                                              End If  ' ** lngDepElem, lngA1 - lngW1.
                                                                                                                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                            Next  ' ** lngX1.
                                                                                                                                          End If  ' ** dblBal
                                                                                                                                        End If  ' ** lngDepElem, lngA1 - lngV1.
                                                                                                                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                      Next  ' ** lngW1.
                                                                                                                                    End If  ' ** dblBal
                                                                                                                                  End If  ' ** lngDepElem, lngA1 - lngU1.
                                                                                                                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                Next  ' ** lngV1.
                                                                                                                              End If  ' ** dblBal
                                                                                                                            End If  ' ** lngDepElem, lngA1 - lngT1.
                                                                                                                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                          Next  ' ** lngU1.
                                                                                                                        End If  ' ** dblBal
                                                                                                                      End If  ' ** lngDepElem, lngA1 - lngS1.
                                                                                                                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                    Next  ' ** lngT1.
                                                                                                                  End If  ' ** dblBal
                                                                                                                End If  ' ** lngDepElem, lngA1 - lngR1.
                                                                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                              Next  ' ** lngS1.
                                                                                                            End If  ' ** dblBal
                                                                                                          End If  ' ** lngDepElem, lngA1 - lngQ1.
                                                                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                        Next  ' ** lngR1.
                                                                                                      End If  ' ** dblBal
                                                                                                    End If  ' ** lngDepElem, lngA1 - lngP1.
                                                                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                  Next  ' ** lngQ1.
                                                                                                End If  ' ** dblBal
                                                                                              End If  ' ** lngDepElem, lngA1 - lngO1.
                                                                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                            Next  ' ** lngP1.
                                                                                          End If  ' ** dblBal
                                                                                        End If  ' ** lngDepElem, lngA1 - lngN1.
                                                                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                      Next  ' ** lngO1.
                                                                                    End If  ' ** dblBal
                                                                                  End If  ' ** lngDepElem, lngA1 - lngM1.
                                                                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                Next  ' ** lngN1.
                                                                              End If  ' ** dblBal
                                                                            End If  ' ** lngDepElem, lngA1 - lngL1.
                                                                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                          Next  ' ** lngM1.
                                                                        End If  ' ** dblBal
                                                                      End If  ' ** lngDepElem, lngA1 - lngK1.
                                                                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                    Next  ' ** lngL1.
                                                                  End If  ' ** dblBal
                                                                End If  ' ** lngDepElem, lngA1 - lngJ1.
                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                              Next  ' ** lngK1.
                                                            End If  ' ** dblBal
                                                          End If  ' ** lngDepElem, lngA1, lngB1, lngC1, lngD1, lngE1, lngF1, lngG1, lngH1, lngI1.
                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                        Next  ' ** lngJ1.
                                                      End If  ' ** dblBal
                                                    End If  ' ** lngDepElem, lngA1, lngB1, lngC1, lngD1, lngE1, lngF1, lngG1, lngH1.
                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                  Next  ' ** lngI1.
                                                End If  ' ** dblBal
                                              End If  ' ** lngDepElem, lngA1, lngB1, lngC1, lngD1, lngE1, lngF1, lngG1.
                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                            Next  ' ** lngH1.
                                          End If  ' ** dblBal
                                        End If  ' ** lngDepElem, lngA1, lngB1, lngC1, lngD1, lngE1, lngF1.
                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                      Next  ' ** lngG1.
                                    End If  ' ** dblBal
                                  End If  ' ** lngDepElem, lngA1, lngB1, lngC1, lngD1, lngE1.
                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                Next  ' ** lngF1.
                              End If  ' ** dblBal
                            End If  ' ** lngDepElem, lngA1, lngB1, lngC1, lngD1.
                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                          Next  ' ** lngE1.
                        End If  ' ** dblBal
                      End If  ' ** lngDepElem, lngA1, lngB1, lngC1.
                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                    Next  ' ** lngD1.
                  End If  ' ** dblBal
                End If  ' ** lngDepElem, lngA1, lngB1.
                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
              Next  ' ** lngC1.
            End If  ' ** dblBal
          End If  ' ** lngDepElem, lngA1.
          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
        Next  ' ** lngB1.
      End If  ' ** dblBal
    End If  ' ** lngDepElem.
    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
  Next  ' ** lngAA.

  NegTaxLots_Match2 = blnRetVal

End Function

Public Function NegTaxLots_Match3(ByRef dblBal As Double, ByRef lngTrans As Long, ByRef arr_varTran As Variant, ByRef blnFound As Boolean, ByRef lngDepElem As Long, ByRef arr_varElem As Variant, ByRef arr_varElem2 As Variant) As Boolean

  Const THIS_PROC As String = "NegTaxLots_Match3"

  Dim blnElem As Boolean
  Dim lngX As Long
  Dim lngA2 As Long, lngB2 As Long, lngC2 As Long, lngD2 As Long, lngE2 As Long, lngF2 As Long, lngG2 As Long, lngH2 As Long, lngI2 As Long
  Dim lngJ2 As Long, lngK2 As Long, lngL2 As Long, lngM2 As Long, lngN2 As Long, lngO2 As Long, lngP2 As Long, lngQ2 As Long, lngR2 As Long
  Dim lngS2 As Long, lngT2 As Long, lngU2 As Long, lngV2 As Long, lngW2 As Long, lngX2 As Long, lngY2 As Long, lngZ2 As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTran().
  Const T_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const T_JNO   As Integer = 0
  Const T_JTYP  As Integer = 1
  Const T_ACTNO As Integer = 2
  Const T_ASTNO As Integer = 3
  Const T_SHARE As Integer = 4

  ' ** Array: arr_varElem().
  Const E_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const E_VAL As Integer = 0

  blnRetVal = True

  For lngA2 = 0& To (lngTrans - 1&)
    blnElem = False
    For lngX = 1& To 26&
      If lngA2 = arr_varElem(E_VAL, lngX) Or lngA2 = arr_varElem2(E_VAL, lngX) Then
        blnElem = True
        Exit For
      End If
    Next  ' ** lngX.
    If blnElem = False And lngA2 <> lngDepElem Then
      dblBal = dblBal - arr_varTran(T_SHARE, lngA2)
      If dblBal <= -0.0002 Then
        dblBal = dblBal + arr_varTran(T_SHARE, lngA2)
        Exit For
      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
        blnFound = True
        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngA2)
        Stop
      Else
        For lngB2 = 0& To (lngTrans - 1&)
          For lngX = 1& To 26&
            If lngB2 = arr_varElem(E_VAL, lngX) Or lngB2 = arr_varElem2(E_VAL, lngX) Then
              blnElem = True
              Exit For
            End If
          Next  ' ** lngX.
          If blnElem = False And lngB2 <> lngDepElem And lngB2 <> lngA2 Then
            dblBal = dblBal - arr_varTran(T_SHARE, lngB2)
            If dblBal <= -0.0002 Then
              dblBal = dblBal + arr_varTran(T_SHARE, lngB2)
              Exit For
            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
              blnFound = True
              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngB2)
              Stop
            Else
              For lngC2 = 0& To (lngTrans - 1&)
                For lngX = 1& To 26&
                  If lngC2 = arr_varElem(E_VAL, lngX) Or lngC2 = arr_varElem2(E_VAL, lngX) Then
                    blnElem = True
                    Exit For
                  End If
                Next  ' ** lngX.
                If blnElem = False And lngC2 <> lngDepElem And lngC2 <> lngA2 And lngC2 <> lngB2 Then
                  dblBal = dblBal - arr_varTran(T_SHARE, lngC2)
                  If dblBal <= -0.0002 Then
                    dblBal = dblBal + arr_varTran(T_SHARE, lngC2)
                    Exit For
                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                    blnFound = True
                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngC2)
                    Stop
                  Else
                    For lngD2 = 0& To (lngTrans - 1&)
                      For lngX = 1& To 26&
                        If lngD2 = arr_varElem(E_VAL, lngX) Or lngD2 = arr_varElem2(E_VAL, lngX) Then
                          blnElem = True
                          Exit For
                        End If
                      Next  ' ** lngX.
                      If blnElem = False And lngD2 <> lngDepElem And lngD2 <> lngA2 And lngD2 <> lngB2 And lngD2 <> lngC2 Then
                        dblBal = dblBal - arr_varTran(T_SHARE, lngD2)
                        If dblBal <= -0.0002 Then
                          dblBal = dblBal + arr_varTran(T_SHARE, lngD2)
                          Exit For
                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                          blnFound = True
                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngD2)
                          Stop
                        Else
                          For lngE2 = 0& To (lngTrans - 1&)
                            For lngX = 1& To 26&
                              If lngE2 = arr_varElem(E_VAL, lngX) Or lngE2 = arr_varElem2(E_VAL, lngX) Then
                                blnElem = True
                                Exit For
                              End If
                            Next  ' ** lngX.
                            If blnElem = False And lngE2 <> lngDepElem And lngE2 <> lngA2 And lngE2 <> lngB2 And lngE2 <> lngC2 And _
                                lngE2 <> lngD2 Then
                              dblBal = dblBal - arr_varTran(T_SHARE, lngE2)
                              If dblBal <= -0.0002 Then
                                dblBal = dblBal + arr_varTran(T_SHARE, lngE2)
                                Exit For
                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                blnFound = True
                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngE2)
                                Stop
                              Else
                                For lngF2 = 0& To (lngTrans - 1&)
                                  For lngX = 1& To 26&
                                    If lngF2 = arr_varElem(E_VAL, lngX) Or lngF2 = arr_varElem2(E_VAL, lngX) Then
                                      blnElem = True
                                      Exit For
                                    End If
                                  Next  ' ** lngX.
                                  If blnElem = False And lngF2 <> lngDepElem And lngF2 <> lngA2 And lngF2 <> lngB2 And lngF2 <> lngC2 And _
                                      lngF2 <> lngD2 And lngF2 <> lngE2 Then
                                    dblBal = dblBal - arr_varTran(T_SHARE, lngF2)
                                    If dblBal <= -0.0002 Then
                                      dblBal = dblBal + arr_varTran(T_SHARE, lngF2)
                                      Exit For
                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                      blnFound = True
                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngF2)
                                      Stop
                                    Else
                                      For lngG2 = 0& To (lngTrans - 1&)
                                        For lngX = 1& To 26&
                                          If lngG2 = arr_varElem(E_VAL, lngX) Or lngG2 = arr_varElem2(E_VAL, lngX) Then
                                            blnElem = True
                                            Exit For
                                          End If
                                        Next  ' ** lngX.
                                        If blnElem = False And lngG2 <> lngDepElem And lngG2 <> lngA2 And lngG2 <> lngB2 And lngG2 <> lngC2 And _
                                            lngG2 <> lngD2 And lngG2 <> lngE2 And lngG2 <> lngF2 Then
                                          dblBal = dblBal - arr_varTran(T_SHARE, lngG2)
                                          If dblBal <= -0.0002 Then
                                            dblBal = dblBal + arr_varTran(T_SHARE, lngG2)
                                            Exit For
                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                            blnFound = True
                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngG2)
                                            Stop
                                          Else
                                            For lngH2 = 0& To (lngTrans - 1&)
                                              For lngX = 1& To 26&
                                                If lngH2 = arr_varElem(E_VAL, lngX) Or lngH2 = arr_varElem2(E_VAL, lngX) Then
                                                  blnElem = True
                                                  Exit For
                                                End If
                                              Next  ' ** lngX.
                                              If blnElem = False And lngH2 <> lngDepElem And lngH2 <> lngA2 And lngH2 <> lngB2 And _
                                                  lngH2 <> lngC2 And lngH2 <> lngD2 And lngH2 <> lngE2 And lngH2 <> lngF2 And lngH2 <> lngG2 Then
                                                dblBal = dblBal - arr_varTran(T_SHARE, lngH2)
                                                If dblBal <= -0.0002 Then
                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngH2)
                                                  Exit For
                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                  blnFound = True
                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngH2)
                                                  Stop
                                                Else
                                                  For lngI2 = 0& To (lngTrans - 1&)
                                                    For lngX = 1& To 26&
                                                      If lngI2 = arr_varElem(E_VAL, lngX) Or lngI2 = arr_varElem2(E_VAL, lngX) Then
                                                        blnElem = True
                                                        Exit For
                                                      End If
                                                    Next  ' ** lngX.
                                                    If blnElem = False And lngI2 <> lngDepElem And lngI2 <> lngA2 And lngI2 <> lngB2 And _
                                                        lngI2 <> lngC2 And lngI2 <> lngD2 And lngI2 <> lngE2 And lngI2 <> lngF2 And _
                                                        lngI2 <> lngG2 And lngI2 <> lngH2 Then
                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngI2)
                                                      If dblBal <= -0.0002 Then
                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngI2)
                                                        Exit For
                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                        blnFound = True
                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngI2)
                                                        Stop
                                                      Else
                                                        For lngJ2 = 0& To (lngTrans - 1&)
                                                          For lngX = 1& To 26&
                                                            If lngJ2 = arr_varElem(E_VAL, lngX) Or lngJ2 = arr_varElem2(E_VAL, lngX) Then
                                                              blnElem = True
                                                              Exit For
                                                            End If
                                                          Next  ' ** lngX.
                                                          If blnElem = False And lngJ2 <> lngDepElem And lngJ2 <> lngA2 And lngJ2 <> lngB2 And _
                                                              lngJ2 <> lngC2 And lngJ2 <> lngD2 And lngJ2 <> lngE2 And lngJ2 <> lngF2 And _
                                                              lngJ2 <> lngG2 And lngJ2 <> lngH2 And lngJ2 <> lngI2 Then
                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngJ2)
                                                            If dblBal <= -0.0002 Then
                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngJ2)
                                                              Exit For
                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                              blnFound = True
                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngJ2)
                                                              Stop
                                                            Else
                                                              For lngK2 = 0& To (lngTrans - 1&)
                                                                For lngX = 1& To 26&
                                                                  If lngK2 = arr_varElem(E_VAL, lngX) Or lngK2 = arr_varElem2(E_VAL, lngX) Then
                                                                    blnElem = True
                                                                    Exit For
                                                                  End If
                                                                Next  ' ** lngX.
                                                                If blnElem = False And lngK2 <> lngDepElem And lngK2 <> lngA2 And lngK2 <> lngB2 And _
                                                                    lngK2 <> lngC2 And lngK2 <> lngD2 And lngK2 <> lngE2 And lngK2 <> lngF2 And _
                                                                    lngK2 <> lngG2 And lngK2 <> lngH2 And lngK2 <> lngI2 And lngK2 <> lngJ2 Then
                                                                  dblBal = dblBal - arr_varTran(T_SHARE, lngK2)
                                                                  If dblBal <= -0.0002 Then
                                                                    dblBal = dblBal + arr_varTran(T_SHARE, lngK2)
                                                                    Exit For
                                                                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                    blnFound = True
                                                                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngK2)
                                                                    Stop
                                                                  Else
                                                                    For lngL2 = 0& To (lngTrans - 1&)
                                                                      For lngX = 1& To 26&
                                                                        If lngL2 = arr_varElem(E_VAL, lngX) Or lngL2 = arr_varElem2(E_VAL, lngX) Then
                                                                          blnElem = True
                                                                          Exit For
                                                                        End If
                                                                      Next  ' ** lngX.
                                                                      If blnElem = False And lngL2 <> lngDepElem And lngL2 <> lngA2 And _
                                                                          lngL2 <> lngB2 And lngL2 <> lngC2 And lngL2 <> lngD2 And _
                                                                          lngL2 <> lngE2 And lngL2 <> lngF2 And lngL2 <> lngG2 And _
                                                                          lngL2 <> lngH2 And lngL2 <> lngI2 And lngL2 <> lngJ2 And _
                                                                          lngL2 <> lngK2 Then
                                                                        dblBal = dblBal - arr_varTran(T_SHARE, lngL2)
                                                                        If dblBal <= -0.0002 Then
                                                                          dblBal = dblBal + arr_varTran(T_SHARE, lngL2)
                                                                          Exit For
                                                                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                          blnFound = True
                                                                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngL2)
                                                                          Stop
                                                                        Else
                                                                          For lngM2 = 0& To (lngTrans - 1&)
                                                                            For lngX = 1& To 26&
                                                                              If lngM2 = arr_varElem(E_VAL, lngX) Or lngM2 = arr_varElem2(E_VAL, lngX) Then
                                                                                blnElem = True
                                                                                Exit For
                                                                              End If
                                                                            Next  ' ** lngX.
                                                                            If blnElem = False And lngM2 <> lngDepElem And lngM2 <> lngA2 And _
                                                                                lngM2 <> lngB2 And lngM2 <> lngC2 And lngM2 <> lngD2 And _
                                                                                lngM2 <> lngE2 And lngM2 <> lngF2 And lngM2 <> lngG2 And _
                                                                                lngM2 <> lngH2 And lngM2 <> lngI2 And lngM2 <> lngJ2 And _
                                                                                lngM2 <> lngK2 And lngM2 <> lngL2 Then
                                                                              dblBal = dblBal - arr_varTran(T_SHARE, lngM2)
                                                                              If dblBal <= -0.0002 Then
                                                                                dblBal = dblBal + arr_varTran(T_SHARE, lngM2)
                                                                                Exit For
                                                                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                blnFound = True
                                                                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngM2)
                                                                                Stop
                                                                              Else
                                                                                For lngN2 = 0& To (lngTrans - 1&)
                                                                                  For lngX = 1& To 26&
                                                                                    If lngN2 = arr_varElem(E_VAL, lngX) Or lngN2 = arr_varElem2(E_VAL, lngX) Then
                                                                                      blnElem = True
                                                                                      Exit For
                                                                                    End If
                                                                                  Next  ' ** lngX.
                                                                                  If blnElem = False And lngN2 <> lngDepElem And lngN2 <> lngA2 And _
                                                                                      lngN2 <> lngB2 And lngN2 <> lngC2 And lngN2 <> lngD2 And _
                                                                                      lngN2 <> lngE2 And lngN2 <> lngF2 And lngN2 <> lngG2 And _
                                                                                      lngN2 <> lngH2 And lngN2 <> lngI2 And lngN2 <> lngJ2 And _
                                                                                      lngN2 <> lngK2 And lngN2 <> lngL2 And lngN2 <> lngM2 Then
                                                                                    dblBal = dblBal - arr_varTran(T_SHARE, lngN2)
                                                                                    If dblBal <= -0.0002 Then
                                                                                      dblBal = dblBal + arr_varTran(T_SHARE, lngN2)
                                                                                      Exit For
                                                                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                      blnFound = True
                                                                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngN2)
                                                                                      Stop
                                                                                    Else
                                                                                      For lngO2 = 0& To (lngTrans - 1&)
                                                                                        For lngX = 1& To 26&
                                                                                          If lngO2 = arr_varElem(E_VAL, lngX) Or lngO2 = arr_varElem2(E_VAL, lngX) Then
                                                                                            blnElem = True
                                                                                            Exit For
                                                                                          End If
                                                                                        Next  ' ** lngX.
                                                                                        If blnElem = False And lngO2 <> lngDepElem And _
                                                                                            lngO2 <> lngA2 And lngO2 <> lngB2 And lngO2 <> lngC2 And _
                                                                                            lngO2 <> lngD2 And lngO2 <> lngE2 And lngO2 <> lngF2 And _
                                                                                            lngO2 <> lngG2 And lngO2 <> lngH2 And lngO2 <> lngI2 And _
                                                                                            lngO2 <> lngJ2 And lngO2 <> lngK2 And lngO2 <> lngL2 And _
                                                                                            lngO2 <> lngM2 And lngO2 <> lngN2 Then
                                                                                          dblBal = dblBal - arr_varTran(T_SHARE, lngO2)
                                                                                          If dblBal <= -0.0002 Then
                                                                                            dblBal = dblBal + arr_varTran(T_SHARE, lngO2)
                                                                                            Exit For
                                                                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                            blnFound = True
                                                                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngO2)
                                                                                            Stop
                                                                                          Else
                                                                                            For lngP2 = 0& To (lngTrans - 1&)
                                                                                              For lngX = 1& To 26&
                                                                                                If lngP2 = arr_varElem(E_VAL, lngX) Or lngP2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                  blnElem = True
                                                                                                  Exit For
                                                                                                End If
                                                                                              Next  ' ** lngX.
                                                                                              If blnElem = False And lngP2 <> lngDepElem And _
                                                                                                  lngP2 <> lngA2 And lngP2 <> lngB2 And lngP2 <> lngC2 And _
                                                                                                  lngP2 <> lngD2 And lngP2 <> lngE2 And lngP2 <> lngF2 And _
                                                                                                  lngP2 <> lngG2 And lngP2 <> lngH2 And lngP2 <> lngI2 And _
                                                                                                  lngP2 <> lngJ2 And lngP2 <> lngK2 And lngP2 <> lngL2 And _
                                                                                                  lngP2 <> lngM2 And lngP2 <> lngN2 And lngP2 <> lngO2 Then
                                                                                                dblBal = dblBal - arr_varTran(T_SHARE, lngP2)
                                                                                                If dblBal <= -0.0002 Then
                                                                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngP2)
                                                                                                  Exit For
                                                                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                  blnFound = True
                                                                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngP2)
                                                                                                  Stop
                                                                                                Else
                                                                                                  For lngQ2 = 0& To (lngTrans - 1&)
                                                                                                    For lngX = 1& To 26&
                                                                                                      If lngQ2 = arr_varElem(E_VAL, lngX) Or lngQ2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                        blnElem = True
                                                                                                        Exit For
                                                                                                      End If
                                                                                                    Next  ' ** lngX.
                                                                                                    If blnElem = False And lngQ2 <> lngDepElem And _
                                                                                                        lngQ2 <> lngA2 And lngQ2 <> lngB2 And lngQ2 <> lngC2 And _
                                                                                                        lngQ2 <> lngD2 And lngQ2 <> lngE2 And lngQ2 <> lngF2 And _
                                                                                                        lngQ2 <> lngG2 And lngQ2 <> lngH2 And lngQ2 <> lngI2 And _
                                                                                                        lngQ2 <> lngJ2 And lngQ2 <> lngK2 And lngQ2 <> lngL2 And _
                                                                                                        lngQ2 <> lngM2 And lngQ2 <> lngN2 And lngQ2 <> lngO2 And _
                                                                                                        lngQ2 <> lngP2 Then
                                                                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngQ2)
                                                                                                      If dblBal <= -0.0002 Then
                                                                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngQ2)
                                                                                                        Exit For
                                                                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                        blnFound = True
                                                                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngQ2)
                                                                                                        Stop
                                                                                                      Else
                                                                                                        For lngR2 = 0& To (lngTrans - 1&)
                                                                                                          For lngX = 1& To 26&
                                                                                                            If lngR2 = arr_varElem(E_VAL, lngX) Or lngR2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                              blnElem = True
                                                                                                              Exit For
                                                                                                            End If
                                                                                                          Next  ' ** lngX.
                                                                                                          If blnElem = False And lngR2 <> lngDepElem And _
                                                                                                              lngR2 <> lngA2 And lngR2 <> lngB2 And lngR2 <> lngC2 And _
                                                                                                              lngR2 <> lngD2 And lngR2 <> lngE2 And lngR2 <> lngF2 And _
                                                                                                              lngR2 <> lngG2 And lngR2 <> lngH2 And lngR2 <> lngI2 And _
                                                                                                              lngR2 <> lngJ2 And lngR2 <> lngK2 And lngR2 <> lngL2 And _
                                                                                                              lngR2 <> lngM2 And lngR2 <> lngN2 And lngR2 <> lngO2 And _
                                                                                                              lngR2 <> lngP2 And lngR2 <> lngQ2 Then
                                                                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngR2)
                                                                                                            If dblBal <= -0.0002 Then
                                                                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngR2)
                                                                                                              Exit For
                                                                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                              blnFound = True
                                                                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngR2)
                                                                                                              Stop
                                                                                                            Else
                                                                                                              For lngS2 = 0& To (lngTrans - 1&)
                                                                                                                For lngX = 1& To 26&
                                                                                                                  If lngS2 = arr_varElem(E_VAL, lngX) Or lngS2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                    blnElem = True
                                                                                                                    Exit For
                                                                                                                  End If
                                                                                                                Next  ' ** lngX.
                                                                                                                If blnElem = False And lngS2 <> lngDepElem And _
                                                                                                                    lngS2 <> lngA2 And lngS2 <> lngB2 And lngS2 <> lngC2 And _
                                                                                                                    lngS2 <> lngD2 And lngS2 <> lngE2 And lngS2 <> lngF2 And _
                                                                                                                    lngS2 <> lngG2 And lngS2 <> lngH2 And lngS2 <> lngI2 And _
                                                                                                                    lngS2 <> lngJ2 And lngS2 <> lngK2 And lngS2 <> lngL2 And _
                                                                                                                    lngS2 <> lngM2 And lngS2 <> lngN2 And lngS2 <> lngO2 And _
                                                                                                                    lngS2 <> lngP2 And lngS2 <> lngQ2 And lngS2 <> lngR2 Then
                                                                                                                  dblBal = dblBal - arr_varTran(T_SHARE, lngS2)
                                                                                                                  If dblBal <= -0.0002 Then
                                                                                                                    dblBal = dblBal + arr_varTran(T_SHARE, lngS2)
                                                                                                                    Exit For
                                                                                                                  ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                    blnFound = True
                                                                                                                    Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngS2)
                                                                                                                    Stop
                                                                                                                  Else
                                                                                                                    For lngT2 = 0& To (lngTrans - 1&)
                                                                                                                      For lngX = 1& To 26&
                                                                                                                        If lngT2 = arr_varElem(E_VAL, lngX) Or lngT2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                          blnElem = True
                                                                                                                          Exit For
                                                                                                                        End If
                                                                                                                      Next  ' ** lngX.
                                                                                                                      If blnElem = False And lngT2 <> lngDepElem And _
                                                                                                                          lngT2 <> lngA2 And lngT2 <> lngB2 And lngT2 <> lngC2 And _
                                                                                                                          lngT2 <> lngD2 And lngT2 <> lngE2 And lngT2 <> lngF2 And _
                                                                                                                          lngT2 <> lngG2 And lngT2 <> lngH2 And lngT2 <> lngI2 And _
                                                                                                                          lngT2 <> lngJ2 And lngT2 <> lngK2 And lngT2 <> lngL2 And _
                                                                                                                          lngT2 <> lngM2 And lngT2 <> lngN2 And lngT2 <> lngO2 And _
                                                                                                                          lngT2 <> lngP2 And lngT2 <> lngQ2 And lngT2 <> lngR2 And _
                                                                                                                          lngT2 <> lngS2 Then
                                                                                                                        dblBal = dblBal - arr_varTran(T_SHARE, lngT2)
                                                                                                                        If dblBal <= -0.0002 Then
                                                                                                                          dblBal = dblBal + arr_varTran(T_SHARE, lngT2)
                                                                                                                          Exit For
                                                                                                                        ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                          blnFound = True
                                                                                                                          Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngT2)
                                                                                                                          Stop
                                                                                                                        Else
                                                                                                                          For lngU2 = 0& To (lngTrans - 1&)
                                                                                                                            For lngX = 1& To 26&
                                                                                                                              If lngU2 = arr_varElem(E_VAL, lngX) Or lngU2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                                blnElem = True
                                                                                                                                Exit For
                                                                                                                              End If
                                                                                                                            Next  ' ** lngX.
                                                                                                                            If blnElem = False And lngU2 <> lngDepElem And _
                                                                                                                                lngU2 <> lngA2 And lngU2 <> lngB2 And lngU2 <> lngC2 And _
                                                                                                                                lngU2 <> lngD2 And lngU2 <> lngE2 And lngU2 <> lngF2 And _
                                                                                                                                lngU2 <> lngG2 And lngU2 <> lngH2 And lngU2 <> lngI2 And _
                                                                                                                                lngU2 <> lngJ2 And lngU2 <> lngK2 And lngU2 <> lngL2 And _
                                                                                                                                lngU2 <> lngM2 And lngU2 <> lngN2 And lngU2 <> lngO2 And _
                                                                                                                                lngU2 <> lngP2 And lngU2 <> lngQ2 And lngU2 <> lngR2 And _
                                                                                                                                lngU2 <> lngS2 And lngU2 <> lngT2 Then
                                                                                                                              dblBal = dblBal - arr_varTran(T_SHARE, lngU2)
                                                                                                                              If dblBal <= -0.0002 Then
                                                                                                                                dblBal = dblBal + arr_varTran(T_SHARE, lngU2)
                                                                                                                                Exit For
                                                                                                                              ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                blnFound = True
                                                                                                                                Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngU2)
                                                                                                                                Stop
                                                                                                                              Else
                                                                                                                                For lngV2 = 0& To (lngTrans - 1&)
                                                                                                                                  For lngX = 1& To 26&
                                                                                                                                    If lngV2 = arr_varElem(E_VAL, lngX) Or lngV2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                                      blnElem = True
                                                                                                                                      Exit For
                                                                                                                                    End If
                                                                                                                                  Next  ' ** lngX.
                                                                                                                                  If blnElem = False And lngV2 <> lngDepElem And _
                                                                                                                                      lngV2 <> lngA2 And lngV2 <> lngB2 And lngV2 <> lngC2 And _
                                                                                                                                      lngV2 <> lngD2 And lngV2 <> lngE2 And lngV2 <> lngF2 And _
                                                                                                                                      lngV2 <> lngG2 And lngV2 <> lngH2 And lngV2 <> lngI2 And _
                                                                                                                                      lngV2 <> lngJ2 And lngV2 <> lngK2 And lngV2 <> lngL2 And _
                                                                                                                                      lngV2 <> lngM2 And lngV2 <> lngN2 And lngV2 <> lngO2 And _
                                                                                                                                      lngV2 <> lngP2 And lngV2 <> lngQ2 And lngV2 <> lngR2 And _
                                                                                                                                      lngV2 <> lngS2 And lngV2 <> lngT2 And lngV2 <> lngU2 Then
                                                                                                                                    dblBal = dblBal - arr_varTran(T_SHARE, lngV2)
                                                                                                                                    If dblBal <= -0.0002 Then
                                                                                                                                      dblBal = dblBal + arr_varTran(T_SHARE, lngV2)
                                                                                                                                      Exit For
                                                                                                                                    ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                      blnFound = True
                                                                                                                                      Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngV2)
                                                                                                                                      Stop
                                                                                                                                    Else
                                                                                                                                      For lngW2 = 0& To (lngTrans - 1&)
                                                                                                                                        For lngX = 1& To 26&
                                                                                                                                          If lngW2 = arr_varElem(E_VAL, lngX) Or lngW2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                                            blnElem = True
                                                                                                                                            Exit For
                                                                                                                                          End If
                                                                                                                                        Next  ' ** lngX.
                                                                                                                                        If blnElem = False And lngW2 <> lngDepElem And _
                                                                                                                                            lngW2 <> lngA2 And lngW2 <> lngB2 And lngW2 <> lngC2 And _
                                                                                                                                            lngW2 <> lngD2 And lngW2 <> lngE2 And lngW2 <> lngF2 And _
                                                                                                                                            lngW2 <> lngG2 And lngW2 <> lngH2 And lngW2 <> lngI2 And _
                                                                                                                                            lngW2 <> lngJ2 And lngW2 <> lngK2 And lngW2 <> lngL2 And _
                                                                                                                                            lngW2 <> lngM2 And lngW2 <> lngN2 And lngW2 <> lngO2 And _
                                                                                                                                            lngW2 <> lngP2 And lngW2 <> lngQ2 And lngW2 <> lngR2 And _
                                                                                                                                            lngW2 <> lngS2 And lngW2 <> lngT2 And lngW2 <> lngU2 And _
                                                                                                                                            lngW2 <> lngV2 Then
                                                                                                                                          dblBal = dblBal - arr_varTran(T_SHARE, lngW2)
                                                                                                                                          If dblBal <= -0.0002 Then
                                                                                                                                            dblBal = dblBal + arr_varTran(T_SHARE, lngW2)
                                                                                                                                            Exit For
                                                                                                                                          ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                            blnFound = True
                                                                                                                                            Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngW2)
                                                                                                                                            Stop
                                                                                                                                          Else
                                                                                                                                            For lngX2 = 0& To (lngTrans - 1&)
                                                                                                                                              For lngX = 1& To 26&
                                                                                                                                                If lngX2 = arr_varElem(E_VAL, lngX) Or lngX2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                                                  blnElem = True
                                                                                                                                                  Exit For
                                                                                                                                                End If
                                                                                                                                              Next  ' ** lngX.
                                                                                                                                              If blnElem = False And lngX2 <> lngDepElem And _
                                                                                                                                                  lngX2 <> lngA2 And lngX2 <> lngB2 And lngX2 <> lngC2 And _
                                                                                                                                                  lngX2 <> lngD2 And lngX2 <> lngE2 And lngX2 <> lngF2 And _
                                                                                                                                                  lngX2 <> lngG2 And lngX2 <> lngH2 And lngX2 <> lngI2 And _
                                                                                                                                                  lngX2 <> lngJ2 And lngX2 <> lngK2 And lngX2 <> lngL2 And _
                                                                                                                                                  lngX2 <> lngM2 And lngX2 <> lngN2 And lngX2 <> lngO2 And _
                                                                                                                                                  lngX2 <> lngP2 And lngX2 <> lngQ2 And lngX2 <> lngR2 And _
                                                                                                                                                  lngX2 <> lngS2 And lngX2 <> lngT2 And lngX2 <> lngU2 And _
                                                                                                                                                  lngX2 <> lngV2 And lngX2 <> lngW2 Then
                                                                                                                                                dblBal = dblBal - arr_varTran(T_SHARE, lngX2)
                                                                                                                                                If dblBal <= -0.0002 Then
                                                                                                                                                  dblBal = dblBal + arr_varTran(T_SHARE, lngX2)
                                                                                                                                                  Exit For
                                                                                                                                                ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                  blnFound = True
                                                                                                                                                  Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngX2)
                                                                                                                                                  Stop
                                                                                                                                                Else
                                                                                                                                                  For lngY2 = 0& To (lngTrans - 1&)
                                                                                                                                                    For lngX = 1& To 26&
                                                                                                                                                      If lngY2 = arr_varElem(E_VAL, lngX) Or lngY2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                                                        blnElem = True
                                                                                                                                                        Exit For
                                                                                                                                                      End If
                                                                                                                                                    Next  ' ** lngX.
                                                                                                                                                    If blnElem = False And lngY2 <> lngDepElem And _
                                                                                                                                                        lngY2 <> lngA2 And lngY2 <> lngB2 And lngY2 <> lngC2 And _
                                                                                                                                                        lngY2 <> lngD2 And lngY2 <> lngE2 And lngY2 <> lngF2 And _
                                                                                                                                                        lngY2 <> lngG2 And lngY2 <> lngH2 And lngY2 <> lngI2 And _
                                                                                                                                                        lngY2 <> lngJ2 And lngY2 <> lngK2 And lngY2 <> lngL2 And _
                                                                                                                                                        lngY2 <> lngM2 And lngY2 <> lngN2 And lngY2 <> lngO2 And _
                                                                                                                                                        lngY2 <> lngP2 And lngY2 <> lngQ2 And lngY2 <> lngR2 And _
                                                                                                                                                        lngY2 <> lngS2 And lngY2 <> lngT2 And lngY2 <> lngU2 And _
                                                                                                                                                        lngY2 <> lngV2 And lngY2 <> lngW2 And lngY2 <> lngX2 Then
                                                                                                                                                      dblBal = dblBal - arr_varTran(T_SHARE, lngY2)
                                                                                                                                                      If dblBal <= -0.0002 Then
                                                                                                                                                        dblBal = dblBal + arr_varTran(T_SHARE, lngY2)
                                                                                                                                                        Exit For
                                                                                                                                                      ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                        blnFound = True
                                                                                                                                                        Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngY2)
                                                                                                                                                        Stop
                                                                                                                                                      Else
                                                                                                                                                        For lngZ2 = 0& To (lngTrans - 1&)
                                                                                                                                                          For lngX = 1& To 26&
                                                                                                                                                            If lngZ2 = arr_varElem(E_VAL, lngX) Or lngZ2 = arr_varElem2(E_VAL, lngX) Then
                                                                                                                                                              blnElem = True
                                                                                                                                                              Exit For
                                                                                                                                                            End If
                                                                                                                                                          Next  ' ** lngX.
                                                                                                                                                          If blnElem = False And lngZ2 <> lngDepElem And _
                                                                                                                                                              lngZ2 <> lngA2 And lngZ2 <> lngB2 And lngZ2 <> lngC2 And _
                                                                                                                                                              lngZ2 <> lngD2 And lngZ2 <> lngE2 And lngZ2 <> lngF2 And _
                                                                                                                                                              lngZ2 <> lngG2 And lngZ2 <> lngH2 And lngZ2 <> lngI2 And _
                                                                                                                                                              lngZ2 <> lngJ2 And lngZ2 <> lngK2 And lngZ2 <> lngL2 And _
                                                                                                                                                              lngZ2 <> lngM2 And lngZ2 <> lngN2 And lngZ2 <> lngO2 And _
                                                                                                                                                              lngZ2 <> lngP2 And lngZ2 <> lngQ2 And lngZ2 <> lngR2 And _
                                                                                                                                                              lngZ2 <> lngS2 And lngZ2 <> lngT2 And lngZ2 <> lngU2 And _
                                                                                                                                                              lngZ2 <> lngV2 And lngZ2 <> lngW2 And lngZ2 <> lngX2 And _
                                                                                                                                                              lngZ2 <> lngY2 Then
                                                                                                                                                            dblBal = dblBal - arr_varTran(T_SHARE, lngZ2)
                                                                                                                                                            If dblBal <= -0.0002 Then
                                                                                                                                                              dblBal = dblBal + arr_varTran(T_SHARE, lngZ2)
                                                                                                                                                              Exit For
                                                                                                                                                            ElseIf dblBal < 0.0002 And dblBal > -0.0002 Then
                                                                                                                                                              blnFound = True
                                                                                                                                                              Debug.Print "'JOURNALNO: " & arr_varTran(T_JNO, lngZ2)
                                                                                                                                                              Stop
                                                                                                                                                            Else
                                                                                                                                                              'To be continued.
                                                                                                                                                            End If  ' ** dblBal
                                                                                                                                                          End If  ' ** lngDepElem, lngA2 - lngY2.
                                                                                                                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                                        Next  ' ** lngZ2.
                                                                                                                                                      End If  ' ** dblBal
                                                                                                                                                    End If  ' ** lngDepElem, lngA2 - lngX2.
                                                                                                                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                                  Next  ' ** lngY2.
                                                                                                                                                End If  ' ** dblBal
                                                                                                                                              End If  ' ** lngDepElem, lngA2 - lngW2.
                                                                                                                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                            Next  ' ** lngX2.
                                                                                                                                          End If  ' ** dblBal
                                                                                                                                        End If  ' ** lngDepElem, lngA2 - lngV2.
                                                                                                                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                      Next  ' ** lngW2.
                                                                                                                                    End If  ' ** dblBal
                                                                                                                                  End If  ' ** lngDepElem, lngA2 - lngU2.
                                                                                                                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                                Next  ' ** lngV2.
                                                                                                                              End If  ' ** dblBal
                                                                                                                            End If  ' ** lngDepElem, lngA2 - lngT2.
                                                                                                                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                          Next  ' ** lngU2.
                                                                                                                        End If  ' ** dblBal
                                                                                                                      End If  ' ** lngDepElem, lngA2 - lngS2.
                                                                                                                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                                    Next  ' ** lngT2.
                                                                                                                  End If  ' ** dblBal
                                                                                                                End If  ' ** lngDepElem, lngA2 - lngR2.
                                                                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                              Next  ' ** lngS2.
                                                                                                            End If  ' ** dblBal
                                                                                                          End If  ' ** lngDepElem, lngA2 - lngQ2.
                                                                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                        Next  ' ** lngR2.
                                                                                                      End If  ' ** dblBal
                                                                                                    End If  ' ** lngDepElem, lngA2 - lngP2.
                                                                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                                  Next  ' ** lngQ2.
                                                                                                End If  ' ** dblBal
                                                                                              End If  ' ** lngDepElem, lngA2 - lngO2.
                                                                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                            Next  ' ** lngP2.
                                                                                          End If  ' ** dblBal
                                                                                        End If  ' ** lngDepElem, lngA2 - lngN2.
                                                                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                      Next  ' ** lngO2.
                                                                                    End If  ' ** dblBal
                                                                                  End If  ' ** lngDepElem, lngA2 - lngM2.
                                                                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                                Next  ' ** lngN2.
                                                                              End If  ' ** dblBal
                                                                            End If  ' ** lngDepElem, lngA2 - lngL2.
                                                                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                          Next  ' ** lngM2.
                                                                        End If  ' ** dblBal
                                                                      End If  ' ** lngDepElem, lngA2 - lngK2.
                                                                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                                    Next  ' ** lngL2.
                                                                  End If  ' ** dblBal
                                                                End If  ' ** lngDepElem, lngA2 - lngJ2.
                                                                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                              Next  ' ** lngK2.
                                                            End If  ' ** dblBal
                                                          End If  ' ** lngDepElem, lngA2, lngB2, lngC2, lngD2, lngE2, lngF2, lngG2, lngH2, lngI2.
                                                          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                        Next  ' ** lngJ2.
                                                      End If  ' ** dblBal
                                                    End If  ' ** lngDepElem, lngA2, lngB2, lngC2, lngD2, lngE2, lngF2, lngG2, lngH2.
                                                    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                                  Next  ' ** lngI2.
                                                End If  ' ** dblBal
                                              End If  ' ** lngDepElem, lngA2, lngB2, lngC2, lngD2, lngE2, lngF2, lngG2.
                                              If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                            Next  ' ** lngH2.
                                          End If  ' ** dblBal
                                        End If  ' ** lngDepElem, lngA2, lngB2, lngC2, lngD2, lngE2, lngF2.
                                        If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                      Next  ' ** lngG2.
                                    End If  ' ** dblBal
                                  End If  ' ** lngDepElem, lngA2, lngB2, lngC2, lngD2, lngE2.
                                  If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                                Next  ' ** lngF2.
                              End If  ' ** dblBal
                            End If  ' ** lngDepElem, lngA2, lngB2, lngC2, lngD2.
                            If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                          Next  ' ** lngE2.
                        End If  ' ** dblBal
                      End If  ' ** lngDepElem, lngA2, lngB2, lngC2.
                      If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
                    Next  ' ** lngD2.
                  End If  ' ** dblBal
                End If  ' ** lngDepElem, lngA2, lngB2.
                If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
              Next  ' ** lngC2.
            End If  ' ** dblBal
          End If  ' ** lngDepElem, lngA2.
          If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
        Next  ' ** lngB2.
      End If  ' ** dblBal
    End If  ' ** lngDepElem.
    If dblBal <= -0.0002 Or blnFound = True Or blnRetVal = False Then Exit For
  Next  ' ** lngAA.

  NegTaxLots_Match3 = blnRetVal

End Function
