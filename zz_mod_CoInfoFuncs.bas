Attribute VB_Name = "zz_mod_CoInfoFuncs"
Option Compare Database
Option Explicit

'VGC 05/04/2015: CHANGES!

Private Const THIS_NAME As String = "zz_mod_CoInfoFuncs"
' **

Public Function CoInfo_HasCountry() As Boolean

  Const THIS_PROC As String = "CoInfo_HasCountry"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry As Variant
  Dim strSQL As String
  Dim lngQrysFound As Long
  Dim intPos1 As Integer
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_CID   As Integer = 0
  Const Q_QNAM  As Integer = 1
  Const Q_QTYP  As Integer = 2
  Const Q_FNAM  As Integer = 3
  Const Q_ADFLD As Integer = 4
  Const Q_SRC   As Integer = 5
  Const Q_COD   As Integer = 6
  Const Q_CTRY  As Integer = 7
  Const Q_PCOD  As Integer = 8
  Const Q_GCTRY As Integer = 9
  Const Q_GPCOD As Integer = 10
  Const Q_CCTRY As Integer = 11
  Const Q_CPCOD As Integer = 12
  Const Q_DAT   As Integer = 13

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_CompanyInformation_01, all records, all fields.
    Set qdf = .QueryDefs("zzz_qry_CompanyInformation_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngQrys = .RecordCount
      .MoveFirst
      arr_varQry = .GetRows(lngQrys)
      ' *********************************************************
      ' ** Array: arr_varQry()
      ' **
      ' **   Field  Element  Name                    Constant
      ' **   =====  =======  ======================  ==========
      ' **     1       0     coinfot_id              Q_CID
      ' **     2       1     qry_name                Q_QNAM
      ' **     3       2     qrytype_type            Q_QTYP
      ' **     4       3     fld_name                Q_FNAM
      ' **     5       4     coinfot_field           Q_ADFLD
      ' **     6       5     coinfot_source          Q_SRC
      ' **     7       6     coinfot_code            Q_COD
      ' **     8       7     has_country             Q_CTRY
      ' **     9       8     has_postalcode          Q_PCOD
      ' **    10       9     has_gscountry           Q_GCTRY
      ' **    11      10     has_gspostalcode        Q_GPCOD
      ' **    12      11     has_cocountry           Q_CCTRY
      ' **    13      12     has_copostalcode        Q_CPCOD
      ' **    14      13     coinfot_datemodified    Q_DAT
      ' **
      ' *********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    For lngX = 0& To (lngQrys - 1&)
      Set qdf = .QueryDefs(arr_varQry(Q_QNAM, lngX))
      With qdf
        arr_varQry(Q_QTYP, lngX) = .Type
        strSQL = .SQL
        intPos1 = InStr(strSQL, "CoInfo_Country")
        If intPos1 > 0 Then
          arr_varQry(Q_CTRY, lngX) = CBool(True)
        End If
        intPos1 = InStr(strSQL, "CoInfo_PostalCode")
        If intPos1 > 0 Then
          arr_varQry(Q_PCOD, lngX) = CBool(True)
        End If
        intPos1 = InStr(strSQL, "gsCoCountry")
        If intPos1 > 0 Then
          arr_varQry(Q_GCTRY, lngX) = CBool(True)
        End If
        intPos1 = InStr(strSQL, "gsCoPostalCode")
        If intPos1 > 0 Then
          arr_varQry(Q_GPCOD, lngX) = CBool(True)
        End If
        intPos1 = InStr(strSQL, "CompanyCountry")
        If intPos1 > 0 Then
          arr_varQry(Q_CCTRY, lngX) = CBool(True)
        End If
        intPos1 = InStr(strSQL, "CompanyPostalCode")
        If intPos1 > 0 Then
          arr_varQry(Q_CPCOD, lngX) = CBool(True)
        End If
      End With
    Next

    Set qdf = .QueryDefs("zzz_qry_CompanyInformation_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveFirst
      lngQrysFound = 0&
      For lngX = 0& To (lngQrys - 1&)
        If arr_varQry(Q_CTRY, lngX) = True Or arr_varQry(Q_PCOD, lngX) = True Or arr_varQry(Q_GCTRY, lngX) = True Or _
            arr_varQry(Q_GPCOD, lngX) = True Or arr_varQry(Q_CCTRY, lngX) = True Or arr_varQry(Q_CPCOD, lngX) = True Then
          .FindFirst "[coinfot_id] = " & CStr(arr_varQry(Q_CID, lngX))
          If .NoMatch = False Then
            .Edit
            ![has_country] = arr_varQry(Q_CTRY, lngX)
            ![has_postalcode] = arr_varQry(Q_PCOD, lngX)
            ![has_gscountry] = arr_varQry(Q_GCTRY, lngX)
            ![has_gspostalcode] = arr_varQry(Q_GPCOD, lngX)
            ![has_cocountry] = arr_varQry(Q_CCTRY, lngX)
            ![has_copostalcode] = arr_varQry(Q_CPCOD, lngX)
            ![coinfot_datemodified] = Now()
            .Update
            lngQrysFound = lngQrysFound + 1&
          Else
            Stop
          End If
        Else
          .FindFirst "[coinfot_id] = " & CStr(arr_varQry(Q_CID, lngX))
          If .NoMatch = True Then
            Stop
          End If
        End If
        Select Case IsNull(![qrytype_type])
        Case True
          .Edit
          ![qrytype_type] = arr_varQry(Q_QTYP, lngX)
          ![coinfot_datemodified] = Now()
          .Update
        Case False
          If ![qrytype_type] <> arr_varQry(Q_QTYP, lngX) Then
            .Edit
            ![qrytype_type] = arr_varQry(Q_QTYP, lngX)
            ![coinfot_datemodified] = Now()
            .Update
          End If
        End Select
      Next
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    .Close
  End With

  Debug.Print "'QRYS FOUND: " & CStr(lngQrysFound)
  DoEvents

'QRYS: 206
'QRYS FOUND: 104
'DONE!
  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  CoInfo_HasCountry = blnRetVal

End Function

Public Function CoInfo_CompanyAdd() As Boolean

  Const THIS_PROC As String = "CoInfo_CompanyAdd"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry As Variant
  Dim Q_QT As String, strSQL As String
  Dim strCompanySQL1 As String, strCompanySQL2 As String, strCompanySQL3 As String, strCompanySQL4 As String
  Dim lngCompanyQrys1 As Long, lngCompanyQrys2 As Long, lngAddFlds As Long, lngAdds As Long
  Dim blnSkip As Boolean
  Dim intPos1 As String, intPos2 As String
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_CID   As Integer = 0
  Const Q_QNAM  As Integer = 1
  Const Q_QTYP  As Integer = 2
  Const Q_FNAM  As Integer = 3
  Const Q_ADFLD As Integer = 4
  Const Q_SRC   As Integer = 5
  Const Q_COD   As Integer = 6
  Const Q_CTRY  As Integer = 7
  Const Q_PCOD  As Integer = 8
  Const Q_GCTRY As Integer = 9
  Const Q_GPCOD As Integer = 10
  Const Q_CCTRY As Integer = 11
  Const Q_CPCOD As Integer = 12
  Const Q_DAT   As Integer = 13

  ' ** Array: arr_varFldQry().
  Const F_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const F_QNAM As Integer = 0
  Const F_SRC  As Integer = 1

  blnRetVal = True

  Q_QT = Chr(34)

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** zz_tbl_CompanyInformation_01, all records, all fields.
    Set qdf = .QueryDefs("zzz_qry_CompanyInformation_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngQrys = .RecordCount
      .MoveFirst
      arr_varQry = .GetRows(lngQrys)
      ' *********************************************************
      ' ** Array: arr_varQry()
      ' **
      ' **   Field  Element  Name                    Constant
      ' **   =====  =======  ======================  ==========
      ' **     1       0     coinfot_id              Q_CID
      ' **     2       1     qry_name                Q_QNAM
      ' **     3       2     qrytype_type            Q_QTYP
      ' **     4       3     fld_name                Q_FNAM
      ' **     5       4     coinfot_field           Q_ADFLD
      ' **     6       5     coinfot_source          Q_SRC
      ' **     7       6     coinfot_code            Q_COD
      ' **     8       7     has_country             Q_CTRY
      ' **     9       8     has_postalcode          Q_PCOD
      ' **    10       9     has_gscountry           Q_GCTRY
      ' **    11      10     has_gspostalcode        Q_GPCOD
      ' **    12      11     has_cocountry           Q_CCTRY
      ' **    13      12     has_copostalcode        Q_CPCOD
      ' **    14      13     coinfot_datemodified    Q_DAT
      ' **
      ' *********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    ' ** qryAddEditAssets_03:
    ' ** SELECT DISTINCTROW masterasset.assetno AS assetnox, masterasset.cusip, masterasset.description AS description_masterasset,
    ' **   masterasset.shareface, masterasset.assettype, masterasset.rate, masterasset.due, masterasset.marketvalue,
    ' **   masterasset.marketvaluecurrent, masterasset.yield, masterasset.currentDate, masterasset.masterasset_TYPE,
    ' **   assettype.assettype_description, assettype.taxcode, assettype.multiplier, assettype.Interest, assettype.Dividend,
    ' **   "North Fork Bank" AS CompanyName, "Oak Plaza" AS CompanyAddress1, "100 Oak Street" AS CompanyAddress2,
    ' **   "North Fork" AS CompanyCity, "MN" AS CompanyState, "55114-5123" AS CompanyZip, "(612) 334-7800" AS CompanyPhone,
    ' **   "" AS CompanyCountry, "" AS CompanyPostalCode
    ' ** FROM assettype INNER JOIN masterasset ON assettype.assettype = masterasset.assettype
    ' ** ORDER BY masterasset.assettype, masterasset.description;

    ' ** "North Fork Bank" AS CompanyName, "Oak Plaza" AS CompanyAddress1, "100 Oak Street" AS CompanyAddress2, "North Fork" AS CompanyCity, "MN" AS CompanyState, "55114-5123" AS CompanyZip, "(612) 334-7800" AS CompanyPhone
    strCompanySQL1 = Q_QT & "North Fork Bank" & Q_QT & " AS CompanyName, " & Q_QT & "Oak Plaza" & _
      Q_QT & " AS CompanyAddress1, " & Q_QT & "100 Oak Street" & Q_QT & " AS CompanyAddress2, " & _
      Q_QT & "North Fork" & Q_QT & " AS CompanyCity, " & Q_QT & "MN" & Q_QT & " AS CompanyState, " & _
      Q_QT & "55114-5123" & Q_QT & " AS CompanyZip, " & Q_QT & "(612) 334-7800" & Q_QT & " AS CompanyPhone"

    ' ** qryAssetList_01_01:
    ' ** SELECT DISTINCTROW masterasset.assetno, masterasset.cusip, masterasset.description, masterasset.rate*100 AS Rate,
    ' **   masterasset.due, masterasset.shareface, masterasset.assettype, assettype.assettype_description, assettype.taxcode,
    ' **   assettype.multiplier, assettype.Interest, assettype.Dividend,
    ' **   'MyCompany' AS CompanyName, 'MyAddr1' AS CompanyAddress1, 'MyAddr2' AS CompanyAddress2, 'MyCity' AS CompanyCity,
    ' **   'MyState' AS CompanyState, 'MyZip' AS CompanyZip, 'MyPhone' AS CompanyPhone
    ' ** FROM assettype INNER JOIN masterasset ON assettype.assettype = masterasset.assettype
    ' ** WHERE (((Left([masterasset].[description],3))<>'HA-'))
    ' ** ORDER BY masterasset.assettype, masterasset.description, masterasset.cusip;

    ' ** 'MyCompany' AS CompanyName, 'MyAddr1' AS CompanyAddress1, 'MyAddr2' AS CompanyAddress2, 'MyCity' AS CompanyCity, 'MyState' AS CompanyState, 'MyZip' AS CompanyZip, 'MyPhone' AS CompanyPhone
    strCompanySQL2 = "'MyCompany' AS CompanyName, 'MyAddr1' AS CompanyAddress1, 'MyAddr2' AS CompanyAddress2, " & _
      "'MyCity' AS CompanyCity, 'MyState' AS CompanyState, 'MyZip' AS CompanyZip, 'MyPhone' AS CompanyPhone"

    ' ** qryCourtReport_06_archive_01:
    ' ** SELECT DISTINCTROW ledger.journalno, account.shortname, account.legalname, ledger.accountno, ledger.transdate,
    ' **   ledger.journaltype, CDate(IIf(IsNull([assetdate])=True,0,IIf(IsNull([PurchaseDate])=True,[assetdate],[PurchaseDate]))) AS assetdatex,
    ' **   ledger.shareface, masterasset.due, masterasset.rate, ledger.pershare, ledger.icash, ledger.pcash, ledger.cost,
    ' **   ([ledger.pcash]-([ledger.cost]*-1)) AS GainLoss, ledger.posted, masterasset.description,
    ' **   IIf(IsNull(ledger!description),"",ledger!description) AS jcommentx, masterasset.assetno, masterasset.yield,
    ' **   Balance.icash AS PreviousIcash, Balance.pcash AS PreviousPcash, Balance.cost AS PreviousCost, journaltype.sortOrder,
    ' **   ledger.RecurringItem, ledger.taxcode, ledger.revcode_ID,
    ' **   CoInfoGet('gsCoName') AS CompanyName, CoInfoGet('gsCoAddress1') AS CompanyAddress1, CoInfoGet('gsCoAddress2') AS CompanyAddress2,
    ' **   CoInfoGet('gsCoCity') AS CompanyCity, CoInfoGet('gsCoState') AS CompanyState, CoInfoGet('gsCoZip') AS CompanyZip,
    ' **   CoInfoGet('gsCoPhone') AS CompanyPhone,
    ' **   account.CaseNum
    ' ** FROM ((account LEFT JOIN qryCourtReport_05 ON account.accountno = qryCourtReport_05.accountno)
    ' **   LEFT JOIN Balance ON (qryCourtReport_05.accountno = Balance.accountno) AND
    ' **   (qryCourtReport_05.[balance date] = Balance.[balance date])) INNER JOIN ((ledger
    ' **   LEFT JOIN masterasset ON ledger.assetno = masterasset.assetno)
    ' **   LEFT JOIN journaltype ON ledger.journaltype = journaltype.journaltype) ON account.accountno = ledger.accountno
    ' ** WHERE (((ledger.ledger_HIDDEN)=False));

    ' ** CoInfoGet('gsCoName') AS CompanyName, CoInfoGet('gsCoAddress1') AS CompanyAddress1, CoInfoGet('gsCoAddress2') AS CompanyAddress2, CoInfoGet('gsCoCity') AS CompanyCity, CoInfoGet('gsCoState') AS CompanyState, CoInfoGet('gsCoZip') AS CompanyZip, CoInfoGet('gsCoPhone') AS CompanyPhone
    strCompanySQL3 = "CoInfoGet('gsCoName') AS CompanyName, CoInfoGet('gsCoAddress1') AS CompanyAddress1, " & _
      "CoInfoGet('gsCoAddress2') AS CompanyAddress2, CoInfoGet('gsCoCity') AS CompanyCity, CoInfoGet('gsCoState') AS CompanyState, " & _
      "CoInfoGet('gsCoZip') AS CompanyZip, CoInfoGet('gsCoPhone') AS CompanyPhone"

    ' ** qryStatementParameters_AssetList_23:
    ' ** SELECT ActiveAssets.assetno, masterasset.description AS MasterAssetDescription, masterasset.due, masterasset.rate,
    ' **   masterasset.currentDate, Sum(IIf(IsNull([ActiveAssets].[cost])=True,0,[ActiveAssets].[cost])) AS TotalCost1,
    ' **   Sum(IIf(IsNull([ActiveAssets].[shareface])=True,0,[ActiveAssets].[shareface]))*IIf([assettype].[assettype]='90',1,1)
    ' **   AS TotalShareface1, account.accountno, account.shortname, account.legalname,
    ' **   IIf(IsNull([masterasset].[assetno])=True,0,IIf(IsNull([masterasset].[marketvalue])=True,0,
    ' **   (IIf([masterasset].[assettype]='90',-1,1)*[masterasset].[marketvalue]))) AS MarketValueX,
    ' **   IIf(IsNull([masterasset].[assetno])=True,0,IIf(IsNull([masterasset].[marketvaluecurrent])=True,0,
    ' **   (IIf([masterasset].[assettype]='90',-1,1)*[masterasset].[marketvaluecurrent]))) AS MarketValueCurrentX,
    ' **   assettype.assettype AS assettypex, assettype.assettype_description,
    ' **   IIf(IsNull([ActiveAssets].[assetno])=True,'',CStr([masterasset].[Description]) &
    ' **   IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) &
    ' **   IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc,
    ' **   IIf(IsNull([masterasset].[yield])=True,0,[masterasset].[yield]) AS YieldX, account.icash AS icash1, account.pcash AS pcash1,
    ' **   NullIfNullStr(CoInfoGet('gsCoName')) AS CompanyName, NullIfNullStr(CoInfoGet('gsCoAddress1')) AS CompanyAddress1,
    ' **   NullIfNullStr(CoInfoGet('gsCoAddress2')) AS CompanyAddress2, NullIfNullStr(CoInfoGet('gsCoCity')) AS CompanyCity,
    ' **   NullIfNullStr(CoInfoGet('gsCoState')) AS CompanyState, NullIfNullStr(CoInfoGet('gsCoZip')) AS CompanyZip,
    ' **   NullIfNullStr(CoInfoGet('gsCoPhone')) AS CompanyPhone
    ' ** FROM (account INNER JOIN tmpRelatedAccount_01 ON account.accountno = tmpRelatedAccount_01.accountno)
    ' **   LEFT JOIN ((masterasset LEFT JOIN assettype ON masterasset.assettype = assettype.assettype)
    ' **   RIGHT JOIN ActiveAssets ON masterasset.assetno = ActiveAssets.assetno) ON account.accountno = ActiveAssets.accountno
    ' ** GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, masterasset.currentDate,
    ' **   account.accountno, account.shortname, account.legalname, IIf(IsNull([masterasset].[assetno])=True,0,
    ' **   IIf(IsNull([masterasset].[marketvalue])=True,0,(IIf([masterasset].[assettype]='90',-1,1)*[masterasset].[marketvalue]))),
    ' **   IIf(IsNull([masterasset].[assetno])=True,0,IIf(IsNull([masterasset].[marketvaluecurrent])=True,0,
    ' **   (IIf([masterasset].[assettype]='90',-1,1)*[masterasset].[marketvaluecurrent]))), assettype.assettype,
    ' **   assettype.assettype_description, IIf(IsNull([ActiveAssets].[assetno])=True,'',CStr([masterasset].[Description]) &
    ' **   IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & IIf([masterasset].[due] Is Not Null,'  Due ' &
    ' **   Format([masterasset].[due],'mm/dd/yyyy'))), IIf(IsNull([masterasset].[yield])=True,0,[masterasset].[yield]),
    ' **   account.icash, account.pcash;

    ' ** NullIfNullStr(CoInfoGet('gsCoName')) AS CompanyName, NullIfNullStr(CoInfoGet('gsCoAddress1')) AS CompanyAddress1, NullIfNullStr(CoInfoGet('gsCoAddress2')) AS CompanyAddress2, NullIfNullStr(CoInfoGet('gsCoCity')) AS CompanyCity, NullIfNullStr(CoInfoGet('gsCoState')) AS CompanyState, NullIfNullStr(CoInfoGet('gsCoZip')) AS CompanyZip, NullIfNullStr(CoInfoGet('gsCoPhone')) AS CompanyPhone
    strCompanySQL4 = "NullIfNullStr(CoInfoGet('gsCoName')) AS CompanyName, NullIfNullStr(CoInfoGet('gsCoAddress1')) AS CompanyAddress1, " & _
      "NullIfNullStr(CoInfoGet('gsCoAddress2')) AS CompanyAddress2, NullIfNullStr(CoInfoGet('gsCoCity')) AS CompanyCity, " & _
      "NullIfNullStr(CoInfoGet('gsCoState')) AS CompanyState, NullIfNullStr(CoInfoGet('gsCoZip')) AS CompanyZip, " & _
      "NullIfNullStr(CoInfoGet('gsCoPhone')) AS CompanyPhone"

    ' ** qryAddEditAssets_04:
    ' ** SELECT CLng(IIf(IsNull([assetnox])=False,[assetnox],0)) AS assetno, qryAddEditAssets_03.cusip,
    ' **   qryAddEditAssets_03.description_masterasset, qryAddEditAssets_03.shareface, qryAddEditAssets_03.assettype,
    ' **   qryAddEditAssets_03.rate, qryAddEditAssets_03.due, qryAddEditAssets_03.marketvalue,
    ' **   qryAddEditAssets_03.marketvaluecurrent, qryAddEditAssets_03.yield, qryAddEditAssets_03.currentDate,
    ' **   qryAddEditAssets_03.masterasset_TYPE, qryAddEditAssets_03.assettype_description, qryAddEditAssets_03.taxcode,
    ' **   qryAddEditAssets_03.multiplier, qryAddEditAssets_03.Interest, qryAddEditAssets_03.Dividend,
    ' **   qryAddEditAssets_03.CompanyName, qryAddEditAssets_03.CompanyAddress1, qryAddEditAssets_03.CompanyAddress2,
    ' **   qryAddEditAssets_03.CompanyCity, qryAddEditAssets_03.CompanyState, qryAddEditAssets_03.CompanyZip, qryAddEditAssets_03.CompanyPhone
    ' ** FROM qryAddEditAssets_03;

    ' ** qryAssetList: (FED BY CODE!)
    ' ** SELECT ActiveAssets.assetno, masterasset.description AS MasterAssetDescription, masterasset.due, masterasset.rate,
    ' **   Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, Sum(IIf(IsNull([ActiveAssets].[shareface]),0,
    ' **   [ActiveAssets].[shareface]))*IIf([assettype].[assettype]='90',1,1) AS TotalShareface, account.accountno, account.shortname,
    ' **   account.legalname, assettype.assettype, assettype.assettype_description,
    ' **   IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) &
    ' **   IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) &
    ' **   IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc,
    ' **   account.icash, account.pcash,
    ' **   'MasterTrust of California' AS CompanyName, 'P.O. Box 10338' AS CompanyAddress1, '' AS CompanyAddress2,
    ' **   'San Bernardino' AS CompanyCity, 'CA' AS CompanyState, '92423' AS CompanyZip, '9093824678' AS CompanyPhone,
    ' **   IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX,
    ' **   IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX,
    ' **   IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX, masterasset.currentDate
    ' ** FROM account LEFT JOIN ((masterasset LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) RIGHT JOIN
    ' **   ActiveAssets ON masterasset.assetno = ActiveAssets.assetno) ON account.accountno = ActiveAssets.accountno
    ' ** GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, account.accountno,
    ' **   account.shortname, account.legalname, assettype.assettype, assettype.assettype_description,
    ' **   IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) &
    ' **   IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) &
    ' **   IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), account.icash, account.pcash,
    ' **   IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), IIf(IsNull([masterasset].[marketvaluecurrent]),0,
    ' **   [masterasset].[marketvaluecurrent]), IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]), masterasset.currentDate,
    ' **   account.accountno
    ' ** HAVING (((account.accountno)='00230'));

    blnSkip = False
    If blnSkip = False Then

      lngCompanyQrys1 = 0&: lngCompanyQrys2 = 0&: lngAdds = 0&
      For lngX = 0& To (lngQrys - 1&)
        If arr_varQry(Q_FNAM, lngX) = "CompanyZip" Then
          If arr_varQry(Q_CTRY, lngX) = False And arr_varQry(Q_PCOD, lngX) = False And arr_varQry(Q_GCTRY, lngX) = False And _
              arr_varQry(Q_GPCOD, lngX) = False And arr_varQry(Q_CCTRY, lngX) = False And arr_varQry(Q_CPCOD, lngX) = False Then
            lngCompanyQrys1 = lngCompanyQrys1 + 1&
            strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
            Set qdf = .QueryDefs(arr_varQry(Q_QNAM, lngX))
            With qdf
              strSQL = .SQL
              intPos1 = InStr(strSQL, strCompanySQL1)
              If intPos1 > 0 Then
                ' ** "North Fork Bank" AS CompanyName, "Oak Plaza" AS CompanyAddress1, "100 Oak Street" AS CompanyAddress2, "North Fork" AS CompanyCity, "MN" AS CompanyState, "55114-5123" AS CompanyZip, "(612) 334-7800" AS CompanyPhone, "" AS CompanyCountry, "" AS CompanyPostalCode
                lngCompanyQrys2 = lngCompanyQrys2 + 1&
                strTmp01 = Left(strSQL, (intPos1 - 1&))
                strTmp02 = Mid(strSQL, intPos1)
                intPos2 = InStr(strTmp02, "CompanyPhone")
                strTmp03 = Mid(strTmp02, (intPos2 + Len("CompanyPhone")))
                strTmp02 = Left(strTmp02, ((intPos2 + Len("CompanyPhone")) - 1))
                strTmp02 = strTmp02 & ", " & Q_QT & Q_QT & " AS CompanyCountry, " & Q_QT & Q_QT & " AS CompanyPostalCode"
                strSQL = strTmp01 & strTmp02 & strTmp03
                .SQL = strSQL
                lngAdds = lngAdds + 1&
              Else
                intPos1 = InStr(strSQL, strCompanySQL2)
                If intPos1 > 0 Then
                  ' ** 'MyCompany' AS CompanyName, 'MyAddr1' AS CompanyAddress1, 'MyAddr2' AS CompanyAddress2, 'MyCity' AS CompanyCity, 'MyState' AS CompanyState, 'MyZip' AS CompanyZip, 'MyPhone' AS CompanyPhone, 'MyCountry' AS CompanyCountry, 'MyPostalCode' AS CompanyPostalCode
                  lngCompanyQrys2 = lngCompanyQrys2 + 1&
                  strTmp01 = Left(strSQL, (intPos1 - 1&))
                  strTmp02 = Mid(strSQL, intPos1)
                  intPos2 = InStr(strTmp02, "CompanyPhone")
                  strTmp03 = Mid(strTmp02, (intPos2 + Len("CompanyPhone")))
                  strTmp02 = Left(strTmp02, ((intPos2 + Len("CompanyPhone")) - 1))
                  strTmp02 = strTmp02 & ", 'MyCountry' AS CompanyCountry, 'MyPostalCode' AS CompanyPostalCode"
                  strSQL = strTmp01 & strTmp02 & strTmp03
                  .SQL = strSQL
                  lngAdds = lngAdds + 1&
                Else
                  intPos1 = InStr(strSQL, strCompanySQL3)
                  If intPos1 > 0 Then
                    ' ** CoInfoGet('gsCoName') AS CompanyName, CoInfoGet('gsCoAddress1') AS CompanyAddress1, CoInfoGet('gsCoAddress2') AS CompanyAddress2, CoInfoGet('gsCoCity') AS CompanyCity, CoInfoGet('gsCoState') AS CompanyState, CoInfoGet('gsCoZip') AS CompanyZip, CoInfoGet('gsCoPhone') AS CompanyPhone, CoInfoGet('gsCoCountry') AS CompanyCountry, CoInfoGet('gsCoPostalCode') AS CompanyPostalCode
                    lngCompanyQrys2 = lngCompanyQrys2 + 1&
                    If .Type = dbQAppend Then
                      ' ** INSERT INTO tmpUpdatedValues ( accountno, shortname, numCopies, PreviousTotalMarketValue, PreviousAccountValue,
                      ' **   CurrentTotalMarketValue, SumNegativePcash, SumNegativeIcash, SumNegativeCost, SumPositivePcash, SumPositiveCost,
                      ' **   SumPositiveIcash, Icash, CurrentICash, Pcash, CurrentPcash, Cost, CurrentCost, LastBalanceDate,
                      ' **   CompanyName, CompanyAddress1, CompanyAddress2, CompanyCity, CompanyState, CompanyZip, CompanyPhone,
                      ' **   CompanyCountry, CompanyPostalCode )
                      ' ** SELECT account.accountno, account.shortname, account.numCopies, Balance.TotalMarketValue AS PreviousTotalMarketValue,
                      ' **   Balance.AccountValue AS PreviousAccountValue, ([qryCurrentTotalMarketValue].[TotalMarketValue]+
                      ' **   [qryCurrentTotalMarketValue].[IcashAndPcash]) AS CurrentTotalMarketValue, qrySumDecreases.SumNegativePcash,
                      ' **   qrySumDecreases.SumNegativeIcash, qrySumDecreases.SumNegativeCost, qrySumIncreases.SumPositivePcash,
                      ' **   qrySumIncreases.SumPositiveCost, qrySumIncreases.SumPositiveIcash, qryTransRangeTotals.CurrentIcash AS Icash,
                      ' **   IIf(IsNull([qryTransRangeTotals].[CurrentIcash]),0,[qryTransRangeTotals].[CurrentIcash])+[balance].[icash]
                      ' **   AS CurrentICash, qryTransRangeTotals.CurrentPcash AS Pcash, IIf(IsNull([qryTransRangeTotals].[CurrentPcash]),0,
                      ' **   [qryTransRangeTotals].[CurrentPcash])+[balance].[pcash] AS CurrentPcash, qryTransRangeTotals.CurrentCost AS Cost,
                      ' **   IIf(IsNull([qryTransRangeTotals].[CurrentCost]),0,[qryTransRangeTotals].[CurrentCost])+[balance].[cost] AS
                      ' **   CurrentCost, qryTransRangeTotals.LastBalanceDate,
                      ' **   CoInfoGet('gsCoName') AS CompanyName, CoInfoGet('gsCoAddress1') AS CompanyAddress1,
                      ' **   CoInfoGet('gsCoAddress2') AS CompanyAddress2, CoInfoGet('gsCoCity') AS CompanyCity,
                      ' **   CoInfoGet('gsCoState') AS CompanyState, CoInfoGet('gsCoZip') AS CompanyZip, CoInfoGet('gsCoPhone') AS CompanyPhone,
                      ' **   CoInfoGet('gsCoCountry') AS CompanyCountry, CoInfoGet('gsCoPostalCode') AS CompanyPostalCode
                      ' ** FROM ((((((account LEFT JOIN qryCurrentTotalMarketValue ON account.accountno = qryCurrentTotalMarketValue.accountno)
                      ' **   LEFT JOIN qryMaxBalDates ON account.accountno = qryMaxBalDates.accountno) LEFT JOIN Balance ON
                      ' **   (qryMaxBalDates.[MaxOfbalance date] = Balance.[balance date]) AND (qryMaxBalDates.accountno = Balance.accountno))
                      ' **   LEFT JOIN qrySumDecreases ON account.accountno = qrySumDecreases.accountno) LEFT JOIN qrySumIncreases ON
                      ' **   account.accountno = qrySumIncreases.accountno) LEFT JOIN qryTransRangeTotals ON
                      ' **   account.accountno = qryTransRangeTotals.accountno) INNER JOIN qryStatementParameters_20 ON
                      ' **   account.accountno = qryStatementParameters_20.accountno;
                      intPos2 = InStr(strSQL, "CompanyPhone")
                      strTmp03 = Mid(strSQL, (intPos2 + Len("CompanyPhone")))
                      strTmp01 = Left(strSQL, ((intPos2 + Len("CompanyPhone")) - 1))
                      strTmp02 = ", CompanyCountry, CompanyPostalCode"
                      strSQL = strTmp01 & strTmp02 & strTmp03
                      intPos1 = InStr(strSQL, strCompanySQL3)
                      strTmp01 = Left(strSQL, (intPos1 - 1&))
                      strTmp02 = Mid(strSQL, intPos1)
                      intPos2 = InStr(strTmp02, "CompanyPhone")
                      strTmp03 = Mid(strTmp02, (intPos2 + Len("CompanyPhone")))
                      strTmp02 = Left(strTmp02, ((intPos2 + Len("CompanyPhone")) - 1))
                      strTmp02 = strTmp02 & ", CoInfoGet('gsCoCountry') AS CompanyCountry, CoInfoGet('gsCoPostalCode') AS CompanyPostalCode"
                      strSQL = strTmp01 & strTmp02 & strTmp03
                      .SQL = strSQL
                    Else
                      strTmp01 = Left(strSQL, (intPos1 - 1&))
                      strTmp02 = Mid(strSQL, intPos1)
                      intPos2 = InStr(strTmp02, "CompanyPhone")
                      strTmp03 = Mid(strTmp02, (intPos2 + Len("CompanyPhone")))
                      strTmp02 = Left(strTmp02, ((intPos2 + Len("CompanyPhone")) - 1))
                      strTmp02 = strTmp02 & ", CoInfoGet('gsCoCountry') AS CompanyCountry, CoInfoGet('gsCoPostalCode') AS CompanyPostalCode"
                      strSQL = strTmp01 & strTmp02 & strTmp03
                      .SQL = strSQL
                    End If
                    lngAdds = lngAdds + 1&
                  Else
                    intPos1 = InStr(strSQL, strCompanySQL4)
                    If intPos1 > 0 Then
                      ' ** NullIfNullStr(CoInfoGet('gsCoName')) AS CompanyName, NullIfNullStr(CoInfoGet('gsCoAddress1')) AS CompanyAddress1, NullIfNullStr(CoInfoGet('gsCoAddress2')) AS CompanyAddress2, NullIfNullStr(CoInfoGet('gsCoCity')) AS CompanyCity, NullIfNullStr(CoInfoGet('gsCoState')) AS CompanyState, NullIfNullStr(CoInfoGet('gsCoZip')) AS CompanyZip, NullIfNullStr(CoInfoGet('gsCoPhone')) AS CompanyPhone, NullIfNullStr(CoInfoGet('gsCoCountry')) AS CompanyCountry, NullIfNullStr(CoInfoGet('gsCoPostalCode')) AS CompanyPostalCode
                      lngCompanyQrys2 = lngCompanyQrys2 + 1&
                      strTmp01 = Left(strSQL, (intPos1 - 1&))
                      strTmp02 = Mid(strSQL, intPos1)
                      intPos2 = InStr(strTmp02, "CompanyPhone")
                      strTmp03 = Mid(strTmp02, (intPos2 + Len("CompanyPhone")))
                      strTmp02 = Left(strTmp02, ((intPos2 + Len("CompanyPhone")) - 1))
                      strTmp02 = strTmp02 & ", NullIfNullStr(CoInfoGet('gsCoCountry')) AS CompanyCountry, NullIfNullStr(CoInfoGet('gsCoPostalCode')) AS CompanyPostalCode"
                      strSQL = strTmp01 & strTmp02 & strTmp03
                      .SQL = strSQL
                      lngAdds = lngAdds + 1&
                    Else
                      intPos1 = InStr(strSQL, ".CompanyZip")
                      If intPos1 > 0 Then
                        If arr_varQry(Q_ADFLD, lngX) = False Then
                          arr_varQry(Q_ADFLD, lngX) = CBool(True)
                          strTmp01 = Rem_CRLF(Mid(strSQL, intPos1, 100))
                          strTmp01 = Trim(Mid(strTmp01, InStr(strTmp01, " ")))
                          strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
                          arr_varQry(Q_SRC, lngX) = strTmp01
                        Else
                          intPos2 = InStr(strSQL, ".CompanyPhone")
                          If intPos2 > 0 Then
                            If .Type = dbQAppend Then
                              ' ** INSERT INTO tmpAssetList2 ( assetno, MasterAssetDescription, due, rate, TotalCost, TotalShareface,
                              ' **   accountno, shortname, legalname, assettype, assettype_description, totdesc, icash, pcash, currentDate,
                              ' **   CompanyName, CompanyAddress1, CompanyAddress2, CompanyCity, CompanyState, CompanyZip, CompanyPhone,
                              ' **   CompanyCountry, CompanyPostalCode,
                              ' **   MarketValueX, MarketValueCurrentX, YieldX )
                              ' ** SELECT qryStatementParameters_AssetList_15c.assetno,
                              ' **   qryStatementParameters_AssetList_15c.MasterAssetDescription, qryStatementParameters_AssetList_15c.due,
                              ' **   qryStatementParameters_AssetList_15c.rate, qryStatementParameters_AssetList_15c.TotalCost,
                              ' **   qryStatementParameters_AssetList_15c.TotalShareface, qryStatementParameters_AssetList_15c.accountno,
                              ' **   qryStatementParameters_AssetList_15c.shortname, qryStatementParameters_AssetList_15c.legalname,
                              ' **   qryStatementParameters_AssetList_15c.assettype, qryStatementParameters_AssetList_15c.assettype_description,
                              ' **   qryStatementParameters_AssetList_15c.totdesc, qryStatementParameters_AssetList_15c.icash,
                              ' **   qryStatementParameters_AssetList_15c.pcash, qryStatementParameters_AssetList_15c.currentDate,
                              ' **   qryStatementParameters_AssetList_15c.CompanyName, qryStatementParameters_AssetList_15c.CompanyAddress1,
                              ' **   qryStatementParameters_AssetList_15c.CompanyAddress2, qryStatementParameters_AssetList_15c.CompanyCity,
                              ' **   qryStatementParameters_AssetList_15c.CompanyState, qryStatementParameters_AssetList_15c.CompanyZip,
                              ' **   qryStatementParameters_AssetList_15c.CompanyPhone, qryStatementParameters_AssetList_15c.CompanyCountry,
                              ' **   qryStatementParameters_AssetList_15c.CompanyPostalCode,
                              ' **   qryStatementParameters_AssetList_15c.MarketValueX, qryStatementParameters_AssetList_15c.MarketValueCurrentX,
                              ' **   qryStatementParameters_AssetList_15c.YieldX
                              ' ** FROM qryStatementParameters_AssetList_15c;
                              intPos2 = InStr(strSQL, "CompanyPhone")
                              strTmp01 = Left(strSQL, ((intPos2 + Len("CompanyPhone")) - 1))
                              strTmp03 = Mid(strSQL, (intPos2 + Len("CompanyPhone")))
                              strTmp02 = strTmp02 & ", CompanyCountry, CompanyPostalCode"
                              strSQL = strTmp01 & strTmp02 & strTmp03
                              intPos2 = InStr(strSQL, ".CompanyPhone")
                              strTmp01 = Left(strSQL, ((intPos2 + Len(".CompanyPhone")) - 1))
                              strTmp03 = Mid(strSQL, (intPos2 + Len(".CompanyPhone")))
                              strTmp02 = ", " & arr_varQry(Q_SRC, lngX) & ".CompanyCountry, " & arr_varQry(Q_SRC, lngX) & ".CompanyPostalCode"
                              strSQL = strTmp01 + strTmp02 + strTmp03
                              .SQL = strSQL
                              lngAdds = lngAdds + 1&
                            Else
                              strTmp01 = Left(strSQL, ((intPos2 + Len(".CompanyPhone")) - 1))
                              strTmp03 = Mid(strSQL, (intPos2 + Len(".CompanyPhone")))
                              strTmp02 = ", " & arr_varQry(Q_SRC, lngX) & ".CompanyCountry, " & arr_varQry(Q_SRC, lngX) & ".CompanyPostalCode"
                              strSQL = strTmp01 + strTmp02 + strTmp03
                              .SQL = strSQL
                              lngAdds = lngAdds + 1&
                            End If
                          Else
                            Stop
                          End If
                        End If
                      Else
                        intPos1 = InStr(strSQL, "AS CompanyZip")
                        If intPos1 > 0 Then
                          arr_varQry(Q_COD, lngX) = CBool(True)
                          'Debug.Print "'BY CODE? " & Rem_CRLF(Mid(strSQL, intPos1, 100))
                          'DoEvents
                        Else
                          Debug.Print "'QRY: " & .Name
                          DoEvents
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End With
            Set qdf = Nothing
          End If
        End If
      Next

    End If  ' ** blnSkip.

'dbQUpdate
'dbQMakeTable
'tmpUpdatedValues

    ' ** zz_tbl_CompanyInformation_01, all records, all fields.
    Set qdf = .QueryDefs("zzz_qry_CompanyInformation_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveFirst
      lngAddFlds = 0&
      For lngX = 0& To (lngQrys - 1&)
        If arr_varQry(Q_ADFLD, lngX) = True Then
          .FindFirst "[coinfot_id] = " & CStr(arr_varQry(Q_CID, lngX))
          If .NoMatch = False Then
            If ![coinfot_field] = False Then
              .Edit
              ![coinfot_field] = arr_varQry(Q_ADFLD, lngX)
              ![coinfot_source] = arr_varQry(Q_SRC, lngX)
              ![coinfot_datemodified] = Now()
              .Update
              lngAddFlds = lngAddFlds + 1&
            End If
          Else
            Stop
          End If
        ElseIf arr_varQry(Q_COD, lngX) = True Then
          .FindFirst "[coinfot_id] = " & CStr(arr_varQry(Q_CID, lngX))
          If .NoMatch = False Then
            If ![coinfot_code] = False Then
              .Edit
              ![coinfot_code] = arr_varQry(Q_COD, lngX)
              ![coinfot_datemodified] = Now()
              .Update
            End If
          Else
            Stop
          End If
        End If
      Next
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    .Close
  End With

  Debug.Print "'FLDS ADDED: " & CStr(lngAdds)
  DoEvents

  Debug.Print "'ADD FLDS: " & CStr(lngAddFlds)
  DoEvents

  Debug.Print "'CO QRYS1: " & CStr(lngCompanyQrys1)
  Debug.Print "'CO QRYS2: " & CStr(lngCompanyQrys2)
  DoEvents

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  CoInfo_CompanyAdd = blnRetVal

End Function

Public Function CoInfo_GroupBy() As Boolean

  Const THIS_PROC As String = "CoInfo_GroupBy"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strSQL As String
  Dim intPos1 As Integer, intPos2 As Integer
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const Q_QNAM As Integer = 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    For Each qdf In .QueryDefs
      With qdf
        strSQL = .SQL
        intPos1 = InStr(strSQL, "CompanyCountry")
        If intPos1 > 0 Then
          intPos2 = InStr(strSQL, "GROUP BY ")
          If intPos2 > 0 Then
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_QNAM, lngE) = .Name
          End If
        End If
      End With
    Next
    Set qdf = Nothing

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    For lngX = 0& To (lngQrys - 1&)
      Set qdf = arr_varQry(Q_QNAM, lngX)
      With qdf
        strSQL = .SQL
        intPos1 = InStr(strSQL, "GROUP BY")
        If intPos1 > 0 Then
          intPos2 = InStr(intPos1, strSQL, "CompanyCountry")
          If intPos2 = 0 Then
            Debug.Print "'QRY: " & arr_varQry(Q_QNAM, lngX)
            DoEvents
          End If
        End If
      End With
      Set qdf = Nothing
      'Debug.Print "'QRY: " & arr_varQry(Q_QNAM, lngX)
      'DoEvents
    Next

    .Close
  End With

'QRYS: 38
'QRY: qryAssetList
'QRY: qryAssetList_01_02
'QRY: qryAssetList_01_03
'QRY: qryAssetList_01_04
'QRY: qryAssetList_01_05
'QRY: qryAssetList_01_07
'QRY: qryCourtReport_08_09_archive_01
'QRY: qryCourtReport_08_10_archive_01
'QRY: qryCourtReport_CA_07z
'QRY: qryCourtReport_FL_07z
'QRY: qryFeeCalculations_01
'QRY: qryFeeCalculations_01b
'QRY: qryLocationReport_01
'QRY: qryLocationReport_02
'QRY: qryRpt_PortfolioModeling_01
'QRY: qryStatementParameters_AssetList_06a
'QRY: qryStatementParameters_AssetList_06b
'QRY: qryStatementParameters_AssetList_06c
'QRY: qryStatementParameters_AssetList_06d
'QRY: qryStatementParameters_AssetList_07a
'QRY: qryStatementParameters_AssetList_07aq
'QRY: qryStatementParameters_AssetList_07b
'QRY: qryStatementParameters_AssetList_07c
'QRY: qryStatementParameters_AssetList_07s
'QRY: qryStatementParameters_AssetList_07s_bak
'QRY: qryStatementParameters_AssetList_07sq
'QRY: qryStatementParameters_AssetList_07sq_bak
'QRY: qryStatementParameters_AssetList_07t
'QRY: qryStatementParameters_AssetList_07t_bak
'QRY: qryStatementParameters_AssetList_23
'QRY: qryStatementParameters_AssetList_27
'QRY: qryStatementParameters_AssetList_28
'QRY: qryStatementParameters_AssetList_28s
'QRY: qryStatementParameters_AssetList_28s_bak
'QRY: qryStatementParameters_Trans_02
'QRY: qryStatementParameters_Trans_02a
'QRY: qryStatementParameters_Trans_02b
'QRY: qryTaxReporting_02
'DONE!

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  CoInfo_GroupBy = blnRetVal

End Function

Public Function CoInfo_QryField() As Boolean

  Const THIS_PROC As String = "CoInfo_QryField"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry As Variant
  Dim lngFlds As Long, arr_varFld As Variant
  Dim lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_DID  As Integer = 0
  Const Q_QID  As Integer = 1
  Const Q_QNAM As Integer = 2

  ' ** Array: arr_varFld().
  Const F_FNAM As Integer = 0

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    Set qdf = .QueryDefs("zzz_qry_Form_Control_09_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngQrys = .RecordCount
      .MoveFirst
      arr_varQry = .GetRows(lngQrys)
      ' *********************************************
      ' ** Array: arr_varQry()
      ' **
      ' **   Field  Element  Name        Constant
      ' **   =====  =======  ==========  ==========
      ' **     1       0     dbs_id      Q_DID
      ' **     2       1     qry_id      Q_QID
      ' **     3       2     qry_name    Q_QNAM
      ' **
      ' *********************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    Set rst = .OpenRecordset("tblQuery_Field", dbOpenDynaset, dbConsistent)

    For lngX = 0& To (lngQrys - 1&)
      arr_varFld = Qry_FldList_rel(arr_varQry(Q_QNAM, lngX))  ' ** Module Function: modQueryFunctions.
      lngFlds = UBound(arr_varFld, 2) + 1
      With rst
        For lngY = 0& To (lngFlds - 1&)
          .AddNew
          ![dbs_id] = arr_varQry(Q_DID, lngX)
          ![qry_id] = arr_varQry(Q_QID, lngX)
          ' ** ![qryfld_id] : AutoNumber.
          ![qryfld_name] = arr_varFld(F_FNAM, lngY)
          ![datatype_db_type] = dbText
          ![ctltype_type] = acTextBox
          ![qryfld_format] = Null
          ![qryfld_datemodified] = Now()
          .Update
        Next
      End With
    Next
    rst.Close

    .Close
  End With

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  CoInfo_QryField = blnRetVal

End Function

Public Function CoInfo_RowSource_Alt() As Boolean
' ** Collect the alternate combo box RowSource's.

  Const THIS_PROC As String = "CoInfo_RowSource_Alt"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngLines As Long, lngDecLines As Long
  Dim strLine As String, strModName As String, strProcName As String, strCodeLine As String
  Dim strFormName As String, strControlName As String, strQryName As String
  Dim lngThisDbsID As Long
  Dim blnAddAll As Boolean, blnAdd As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim strTmp01 As String
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 13  ' ** Array's first-element UBound().
  Const Q_DID  As Integer = 0
  Const Q_VID  As Integer = 1
  Const Q_VNAM As Integer = 2
  Const Q_PID  As Integer = 3
  Const Q_PNAM As Integer = 4
  Const Q_LIN  As Integer = 5
  Const Q_COD  As Integer = 6
  Const Q_TXT  As Integer = 7
  Const Q_FID  As Integer = 8
  Const Q_FNAM As Integer = 9
  Const Q_CID  As Integer = 10
  Const Q_CNAM As Integer = 11
  Const Q_QID  As Integer = 12
  Const Q_QNAM As Integer = 13

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        If .Type = vbext_ct_Document Then
          strModName = .Name
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngX = lngDecLines To lngLines
              strLine = Trim(.Lines(lngX, 1))
              If strLine <> vbNullString Then
                If Left(strLine, 1) <> "'" Then
                  intPos1 = InStr(strLine, ".RowSource")
                  If intPos1 > 0 Then
                    intPos2 = InStr(strLine, "=")
                    If intPos2 > 0 Then
                      intPos3 = InStr(strLine, " ")
                      strCodeLine = Trim(Left(strLine, intPos3))
                      strTmp01 = Trim(Mid(strLine, intPos3))
                      If Left(strTmp01, 3) = "If " Or Right(strTmp01, 5) = " Then" Or _
                          InStr(strLine, "strCurSort = .lbxShortAccountName") > 0 Or InStr(strLine, "ctl.RowSource = strTmp00") > 0 Then
                        ' ** Skip.
                      Else
                        strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                        strControlName = Left(strLine, (intPos1 - 1))
                        strControlName = Trim(Mid(strControlName, intPos3))
                        If Left(strControlName, 1) = "." Then strControlName = Mid(strControlName, 2)
                        If Left(strControlName, 3) = "Me." Then strControlName = Mid(strControlName, 4)
                        intPos3 = InStr(strModName, "_")
                        strFormName = Mid(strModName, (intPos3 + 1))
                        strQryName = Trim(Mid(strLine, (intPos2 + 1)))
                        intPos3 = InStr(strQryName, "'")
                        If intPos3 > 0 Then strQryName = Trim(Left(strQryName, (intPos3 - 1)))
                        strQryName = Rem_Quotes(strQryName)  ' ** Module Function: modStringFuncs.
                        lngQrys = lngQrys + 1&
                        lngE = lngQrys - 1&
                        ReDim Preserve arr_varQry(Q_ELEMS, lngE)
                        ' ***************************************************
                        ' ** Array: arr_varQry()
                        ' **
                        ' **   Field  Element  Name              Constant
                        ' **   =====  =======  ================  ==========
                        ' **     1       0     dbs_id            Q_DID
                        ' **     2       1     vbcom_id          Q_VID
                        ' **     3       2     vbcom_name        Q_VNAM
                        ' **     4       3     vbcomproc_id      Q_PID
                        ' **     5       4     vbcomproc_name    Q_PNAM
                        ' **     6       5     vbcomproc_line    Q_LIN
                        ' **     7       6     vbcomproc_code    Q_COD
                        ' **     8       7     vbcomproc_raw     Q_TXT
                        ' **     9       8     frm_id            Q_FID
                        ' **    10       9     frm_name          Q_FNAM
                        ' **    11      10     ctl_id            Q_CID
                        ' **    12      11     ctl_name          Q_CNAM
                        ' **    13      12     qry_id            Q_QID
                        ' **    14      13     qry_name          Q_QNAM
                        ' **
                        ' ***************************************************
                        arr_varQry(Q_DID, lngE) = lngThisDbsID
                        arr_varQry(Q_VID, lngE) = Null
                        arr_varQry(Q_VNAM, lngE) = strModName
                        arr_varQry(Q_PID, lngE) = Null
                        arr_varQry(Q_PNAM, lngE) = strProcName
                        arr_varQry(Q_LIN, lngE) = lngX
                        arr_varQry(Q_COD, lngE) = strCodeLine
                        arr_varQry(Q_TXT, lngE) = strLine
                        arr_varQry(Q_FID, lngE) = Null
                        arr_varQry(Q_FNAM, lngE) = strFormName
                        arr_varQry(Q_CID, lngE) = Null
                        arr_varQry(Q_CNAM, lngE) = strControlName
                        arr_varQry(Q_QID, lngE) = Null
                        arr_varQry(Q_QNAM, lngE) = strQryName
                      End If
                    End If  ' ** intPos2.
                  End If  ' ** intPos1.
                End If  ' ** Remark.
              End If  ' ** vbNullString.
            Next  ' ** lngX.
          End With  ' ** cod.
        End If  ' ** Form/Report module.
      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  Debug.Print "'ALT QRYS: " & CStr(lngQrys)
  DoEvents
'Stop

  If lngQrys > 0& Then
    Set dbs = CurrentDb
    With dbs

      Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngQrys - 1&)
          .FindFirst "[dbs_id] = " & CStr(arr_varQry(Q_DID, lngX)) & " And [vbcom_name] = '" & arr_varQry(Q_VNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varQry(Q_VID, lngX) = ![vbcom_id]
          Else
            Stop
          End If
        Next
        .Close
      End With
      Set rst = Nothing
'Stop

      Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngQrys - 1&)
          .FindFirst "[dbs_id] = " & CStr(arr_varQry(Q_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varQry(Q_VID, lngX)) & " And " & _
            "[vbcomproc_name] = '" & arr_varQry(Q_PNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varQry(Q_PID, lngX) = ![vbcomproc_id]
          Else
            Stop
          End If
        Next
        .Close
      End With
      Set rst = Nothing
'Stop

      Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngQrys - 1&)
          .FindFirst "[dbs_id] = " & CStr(arr_varQry(Q_DID, lngX)) & " And [frm_name] = '" & arr_varQry(Q_FNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varQry(Q_FID, lngX) = ![frm_id]
          Else
            Stop
          End If
        Next
        .Close
      End With
      Set rst = Nothing
'Stop

      Set rst = .OpenRecordset("tblForm_Control", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngQrys - 1&)
          .FindFirst "[dbs_id] = " & CStr(arr_varQry(Q_DID, lngX)) & " And [frm_id] = " & CStr(arr_varQry(Q_FID, lngX)) & " And " & _
            "[ctl_name] = '" & arr_varQry(Q_CNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varQry(Q_CID, lngX) = ![ctl_id]
          Else
            Debug.Print "'" & arr_varQry(Q_CNAM, lngX)
            Stop
          End If
        Next
        .Close
      End With
      Set rst = Nothing
'Stop

      Set rst = .OpenRecordset("tblQuery", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngQrys - 1&)
          If arr_varQry(Q_QNAM, lngX) = "strSQL" Or arr_varQry(Q_QNAM, lngX) = "strTmp00" Or arr_varQry(Q_QNAM, lngX) = "strSortNow" Then
            arr_varQry(Q_QID, lngX) = 0&
          Else
            .FindFirst "[dbs_id] = " & CStr(arr_varQry(Q_DID, lngX)) & " And [qry_name] = '" & arr_varQry(Q_QNAM, lngX) & "'"
            If .NoMatch = False Then
              arr_varQry(Q_QID, lngX) = ![qry_id]
            Else
              Stop
            End If
          End If
        Next
        .Close
      End With
      Set rst = Nothing
'Stop

      Set rst = .OpenRecordset("zz_tbl_Form_Control_RowSource_Alt", dbOpenDynaset, dbConsistent)
      With rst
        blnAddAll = False: blnAdd = False
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        Else
          .MoveFirst
        End If
        For lngX = 0& To (lngQrys - 1&)
          blnAdd = False
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            .FindFirst "[dbs_id] = " & CStr(arr_varQry(Q_DID, lngX)) & " And [frm_id] = " & CStr(arr_varQry(Q_FID, lngX)) & " And " & _
              "[ctl_id] = " & CStr(arr_varQry(Q_CID, lngX)) & " And [vbcomproc_id] = " & CStr(arr_varQry(Q_PID, lngX)) & " And " & _
              "[vbcomproc_line] = " & CStr(arr_varQry(Q_LIN, lngX))
            If .NoMatch = True Then
              blnAdd = True
            End If
          End Select
          If blnAdd = True Then
            .AddNew
            ![dbs_id] = arr_varQry(Q_DID, lngX)
            ![frm_id] = arr_varQry(Q_FID, lngX)
            ![ctl_id] = arr_varQry(Q_CID, lngX)
            ' ** ![rowsrcalt_id] : AutoNumber.
            ![vbcom_id] = arr_varQry(Q_VID, lngX)
            ![vbcomproc_id] = arr_varQry(Q_PID, lngX)
            ![qry_id] = arr_varQry(Q_QID, lngX)
            ![frm_name] = arr_varQry(Q_FNAM, lngX)
            ![ctl_name] = arr_varQry(Q_CNAM, lngX)
            ![vbcom_name] = arr_varQry(Q_VNAM, lngX)
            ![vbcomproc_name] = arr_varQry(Q_PNAM, lngX)
            ![qry_name] = arr_varQry(Q_QNAM, lngX)
            ![vbcomproc_line] = arr_varQry(Q_LIN, lngX)
            ![vbcomproc_code] = arr_varQry(Q_COD, lngX)
            ![vbcomproc_raw] = arr_varQry(Q_TXT, lngX)
            ![rowsrcalt_datemodified] = Now()
            .Update
          Else
            If ![qry_name] <> arr_varQry(Q_QNAM, lngX) Then
              .Edit
              ![qry_name] = arr_varQry(Q_QNAM, lngX)
              ![rowsrcalt_datemodified] = Now()
              .Update
            End If
            If ![vbcomproc_code] <> arr_varQry(Q_COD, lngX) Then
              .Edit
              ![vbcomproc_code] = arr_varQry(Q_COD, lngX)
              ![rowsrcalt_datemodified] = Now()
              .Update
            End If
            If ![vbcomproc_raw] <> arr_varQry(Q_TXT, lngX) Then
              .Edit
              ![vbcomproc_raw] = arr_varQry(Q_TXT, lngX)
              ![rowsrcalt_datemodified] = Now()
              .Update
            End If
          End If
        Next
        .Close
      End With
      Set rst = Nothing

      .Close
    End With
  End If  ' ** lngQrys.

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  CoInfo_RowSource_Alt = blnRetVal

End Function

Public Function Rpt_Archive_Qry_Copy() As Boolean

  Const THIS_PROC As String = "CoInfo_RowSource_Alt"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strQryName As String, strQryNum As String, strSQL As String, strDesc As String
  Dim lngQrysCreated As Long
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const Q_QNAM1 As Integer = 0
  Const Q_DSC1  As Integer = 1
  Const Q_SQL1  As Integer = 2
  Const Q_QNAM2 As Integer = 3
  Const Q_DSC2  As Integer = 4
  Const Q_SQL2  As Integer = 5

  Const QRY_BASE As String = "qryRpt_ArchivedTransactions_"

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    lngQrys = 0&
    ReDim arr_varQry(Q_ELEMS, 0)

    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
          ' ** qryRpt_ArchivedTransactions_20_01, qryRpt_ArchivedTransactions_30_01_01.
          If Mid(.Name, (Len(QRY_BASE) + 1), 1) = "2" Or Mid(.Name, (Len(QRY_BASE) + 1), 1) = "3" Then
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_QNAM1, lngE) = .Name
            arr_varQry(Q_DSC1, lngE) = .Properties("Description")
            arr_varQry(Q_SQL1, lngE) = .SQL
            arr_varQry(Q_QNAM2, lngE) = Null
            arr_varQry(Q_DSC2, lngE) = Null
            arr_varQry(Q_SQL2, lngE) = Null
          End If
        End If
      End With
    Next
    Set qdf = Nothing

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    If lngQrys > 0& Then

      For lngX = 0& To (lngQrys - 1&)
        strQryName = arr_varQry(Q_QNAM1, lngX)
        strQryNum = Mid(strQryName, 29, 2)
        strQryNum = CStr(Val(strQryNum) + 20)  ' ** .._40.
        strQryName = Left(strQryName, 28) & strQryNum & Mid(strQryName, 31)
        arr_varQry(Q_QNAM2, lngX) = strQryName
        strSQL = arr_varQry(Q_SQL1, lngX)
        intPos1 = InStr(strSQL, "qryRpt_ArchivedTransactions_02")
        If intPos1 > 0 Then
          strSQL = StringReplace(strSQL, "qryRpt_ArchivedTransactions_02", "qryRpt_ArchivedTransactions_03")  ' ** Module Function: modStringFuncs.
        End If
        intPos1 = InStr(strSQL, "Transactions_2")
        If intPos1 > 0 Then
          strSQL = StringReplace(strSQL, "Transactions_2", "Transactions_4")  ' ** Module Function: modStringFuncs.
        End If
        intPos1 = InStr(strSQL, "Transactions_3")
        If intPos1 > 0 Then
          strSQL = StringReplace(strSQL, "Transactions_3", "Transactions_5")  ' ** Module Function: modStringFuncs.
        End If
        strTmp02 = "Archived Transaction Statement"
        intPos1 = InStr(strSQL, strTmp02)
        If intPos1 > 0 Then
          strTmp01 = Left(strSQL, (intPos1 - 1))
          intPos2 = intPos1 + Len(strTmp02)
          strTmp03 = Mid(strSQL, intPos2)
          strTmp02 = strTmp02 & " - All"
          strSQL = strTmp01 & strTmp02 & strTmp03
        End If
        arr_varQry(Q_SQL2, lngX) = strSQL
        strDesc = arr_varQry(Q_DSC1, lngX)
        intPos1 = InStr(strDesc, ".._02")
        If intPos1 > 0 Then
          strDesc = StringReplace(strDesc, ".._02", ".._03")  ' ** Module Function: modStringFuncs.
        End If
        intPos1 = InStr(strDesc, ".._2")
        If intPos1 > 0 Then
          strDesc = StringReplace(strDesc, ".._2", ".._4")  ' ** Module Function: modStringFuncs.
        End If
        intPos1 = InStr(strDesc, ".._3")
        If intPos1 > 0 Then
          strDesc = StringReplace(strDesc, ".._3", ".._5")  ' ** Module Function: modStringFuncs.
        End If
        arr_varQry(Q_DSC2, lngX) = strDesc
      Next

      lngQrysCreated = 0&
      For lngX = 0& To (lngQrys - 1&)
        Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngX), arr_varQry(Q_SQL2, lngX))
        With qdf
          strDesc = arr_varQry(Q_DSC2, lngX)
          Set prp = .CreateProperty("Description", dbText, strDesc)
On Error Resume Next
          .Properties.Append prp
          If ERR.Number <> 0 Then
On Error GoTo 0
            .Properties("Description") = strDesc
          Else
On Error GoTo 0
          End If
        End With
        lngQrysCreated = lngQrysCreated + 1&
      Next
      Set prp = Nothing
      Set qdf = Nothing

    End If  ' ** lngQrys.

    Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
    DoEvents

    .Close
  End With
  Set dbs = Nothing

'QRYS: 86
'QRYS CREATED: 86
'DONE!
  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set prp = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

End Function

Public Function Datasheet_Cols() As Boolean

  Const THIS_PROC As String = "Datasheet_Cols"

  Dim dat As Form, ctl As Control, dbs As DAO.Database, qdf As DAO.QueryDef, fld As DAO.Field
  Dim lngCols As Long, arr_varCol() As Variant
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varCol().
  Const C_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const C_CNAM As Integer = 0
  Const C_WDTH As Integer = 1

  blnRetVal = True

  lngCols = 0&
  ReDim arr_varCol(C_ELEMS, 0)

  Set dat = Screen.ActiveDatasheet
  With dat
    For Each ctl In .Controls
      With ctl
        lngCols = lngCols + 1&
        lngE = lngCols - 1&
        ReDim Preserve arr_varCol(C_ELEMS, lngE)
        arr_varCol(C_CNAM, lngE) = .Name
        arr_varCol(C_WDTH, lngE) = .ColumnWidth
      End With
    Next
  End With
  Set ctl = Nothing
  Set dat = Nothing

  Stop

  Set dbs = CurrentDb
  With dbs
    Set qdf = .QueryDefs("qryRpt_AccountProfile_10_02_10") 'qryRpt_AccountProfile_10_01_10") '"qryRpt_AccountProfile_11_01_09")
    With qdf
      For lngX = 0& To (lngCols - 1&)
        Set fld = .Fields(arr_varCol(C_CNAM, lngX))
        With fld
          .Properties("ColumnWidth") = arr_varCol(C_WDTH, lngX)
        End With
      Next
    End With
    Set fld = Nothing
    Set qdf = Nothing
    .Close
  End With
  Set dbs = Nothing

  'SETS COLUMN WIDTHS GREAT, BUT WON'T SAVE!
  'Set dat = Screen.ActiveDatasheet
  'With dat
  '  For lngX = 0& To (lngCols - 1&)
  '    Set ctl = .Controls(arr_varCol(C_CNAM, lngX))
  '    With ctl
  '      .ColumnWidth = arr_varCol(C_WDTH, lngX)
  '    End With
  '  Next
  'End With

  Debug.Print "'COLS: " & CStr(lngCols)
  DoEvents

' ** Datasheet Control properties:
'ColumnWidth
'ColumnOrder
'ColumnHidden
'Name
'ControlType
'ControlSource
'Enabled
'Locked
'Format
'Text
'SelStart
'SelLength
'SelText
'SmartTags
'DONE!
  Debug.Print "'DONE!"

  Beep

  Set fld = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set ctl = Nothing
  Set dat = Nothing

  Datasheet_Cols = blnRetVal

End Function

Public Function CtlTipText_Match() As Boolean

  Const THIS_PROC As String = "CtlTipText_Match"

  Dim frm1 As Access.Form, frm2 As Access.Form, ctl1 As Access.Control, ctl2 As Access.Control
  Dim lngTipsAdded As Long, lngStatsAdded As Long
  Dim varTmp00 As Variant
  Dim blnRetVal As Boolean

  blnRetVal = True

  Set frm1 = Forms(0)
  Set frm2 = Forms(1)

  lngTipsAdded = 0&: lngStatsAdded = 0&
  With frm1
    For Each ctl1 In .Controls
      With ctl1
On Error Resume Next
        varTmp00 = .ControlTipText
        If ERR.Number = 0 Then
On Error GoTo 0
          If IsNull(varTmp00) = False Then
            If varTmp00 <> vbNullString Then
              For Each ctl2 In frm2.Controls
                With ctl2
                  If .Name = ctl1.Name Then
                    If IsNull(.ControlTipText) = True Then
                      .ControlTipText = varTmp00
                      lngTipsAdded = lngTipsAdded + 1&
                      'Debug.Print "'" & .Name & "  " & varTmp00
                    Else
                      If .ControlTipText = vbNullString Then
                        .ControlTipText = varTmp00
                        lngTipsAdded = lngTipsAdded + 1&
                        'Debug.Print "'" & .Name & "  " & varTmp00
                      End If
                    End If
                    Exit For
                  End If
                End With
              Next
            End If
          End If
        Else
On Error GoTo 0
        End If
On Error Resume Next
        varTmp00 = .StatusBarText
        If ERR.Number = 0 Then
On Error GoTo 0
          If IsNull(varTmp00) = False Then
            If varTmp00 <> vbNullString Then
              For Each ctl2 In frm2.Controls
                With ctl2
                  If .Name = ctl1.Name Then
                    If IsNull(.StatusBarText) = True Then
                      .StatusBarText = varTmp00
                      lngStatsAdded = lngStatsAdded + 1&
                      'Debug.Print "'" & .Name & "  " & varTmp00
                    Else
                      If .StatusBarText = vbNullString Then
                        .StatusBarText = varTmp00
                        lngStatsAdded = lngStatsAdded + 1&
                        'Debug.Print "'" & .Name & "  " & varTmp00
                      End If
                    End If
                    Exit For
                  End If
                End With
              Next
            End If
          End If
        Else
On Error GoTo 0
        End If
      End With
    Next

  End With


  Debug.Print "'TIPS ADDED:  " & CStr(lngTipsAdded)
  Debug.Print "'STATS ADDED: " & CStr(lngStatsAdded)
  DoEvents

'TIPS ADDED:  3
'STATS ADDED: 3
'DONE!
'TIPS ADDED: 1
'DONE!
'TIPS ADDED: 14
'DONE!
  Debug.Print "'DONE!"

  Beep

  Set ctl1 = Nothing
  Set ctl2 = Nothing
  Set frm1 = Nothing
  Set frm2 = Nothing

  CtlTipText_Match = blnRetVal

End Function
