Attribute VB_Name = "zz_mod_MasterTrustFuncs"
Option Compare Database
Option Explicit

'VGC 03/27/2015: CHANGES!

Private Const THIS_NAME As String = "zz_mod_MasterTrustFuncs"
' **

Public Function MT_QryCopy1() As Boolean

  Const THIS_PROC As String = "MT_QryCopy1"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, prp As Object
  Dim lngAccts As Long, arr_varAcct As Variant
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strAccountNo As String, strActNo As String, strQryName As String, strSQL As String, strDesc As String
  Dim strQryTemplate As String, strDateTemplate As String, strNewDate As String
  Dim lngQrysCreated As Long, lngRecs As Long, lngAssets As Long
  Dim lngMults As Long, arr_varMult() As Variant
  Dim intPos1 As Integer
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varAcct().
  Const A_ACTNO  As Integer = 0
  Const A_BDAT   As Integer = 1
  Const A_ICASH  As Integer = 2
  Const A_PCASH  As Integer = 3
  Const A_COST   As Integer = 4
  Const A_MKTVAL As Integer = 5
  Const A_ACTVAL As Integer = 6

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const Q_QNAM1 As Integer = 0
  Const Q_SQL1  As Integer = 1
  Const Q_DSC1  As Integer = 2
  Const Q_TYP   As Integer = 3
  Const Q_QNAM2 As Integer = 4
  Const Q_SQL2  As Integer = 5
  Const Q_DSC2  As Integer = 6

  ' ** Array: arr_varMult().
  Const M_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const M_ACTNO As Integer = 0
  Const M_BDATE As Integer = 1
  Const M_CNT   As Integer = 2
  Const M_QNAM  As Integer = 3

  Const QRY_BASE As String = "zzz_qry_MasterTrust_05_"
  Const QRY_NEW As String = "zzz_qry_MasterTrust_08_"
  Const SEQ_OLD As String = "05"
  Const SEQ_NEW As String = "08"

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    ' ** zzz_qry_MasterTrust_02 (Balance, just TotalMarketValue <> cost), just 12/31/2014.
    Set qdf = .QueryDefs("zzz_qry_MasterTrust_03_05")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngAccts = .RecordCount
      .MoveFirst
      arr_varAcct = .GetRows(lngAccts)
      ' *****************************************************
      ' ** Array: arr_varAcct()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     accountno           A_ACTNO
      ' **     2       1     balance date        A_BDAT
      ' **     3       2     icash               A_ICASH
      ' **     4       3     pcash               A_PCASH
      ' **     5       4     cost                A_COST
      ' **     6       5     TotalMarketValue    A_MKTVAL
      ' **     7       6     AccountValue        A_ACTVAL
      ' **
      ' *****************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    lngQrys = 0&
    ReDim arr_varQry(Q_ELEMS, 0)

    strQryTemplate = "298"
    strDateTemplate = "12/31/2011"
    strNewDate = arr_varAcct(A_BDAT, 0)
    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, Len(QRY_BASE & strQryTemplate)) = (QRY_BASE & strQryTemplate) Then
          lngQrys = lngQrys + 1&
          lngE = lngQrys - 1&
          ReDim Preserve arr_varQry(Q_ELEMS, lngE)
          arr_varQry(Q_QNAM1, lngE) = .Name
          arr_varQry(Q_SQL1, lngE) = .SQL
          arr_varQry(Q_DSC1, lngE) = .Properties("Description")
          arr_varQry(Q_TYP, lngE) = .Type
          arr_varQry(Q_QNAM2, lngE) = Null
          arr_varQry(Q_SQL2, lngE) = Null
          arr_varQry(Q_DSC2, lngE) = Null
        End If
      End With
    Next  ' ** qdf.
    Set qdf = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'ACCTS: " & CStr(lngAccts)
    DoEvents

    lngMults = 0&
    ReDim arr_varMult(M_ELEMS, 0)

    Debug.Print "'|";

    lngQrysCreated = 0&
    For lngX = 0& To (lngAccts - 1&)

      strAccountNo = arr_varAcct(A_ACTNO, lngX)
      strActNo = Right(strAccountNo, 3)

      For lngY = 0& To (lngQrys - 1&)
        strQryName = QRY_NEW & strActNo & Right(arr_varQry(Q_QNAM1, lngY), 3)
        arr_varQry(Q_QNAM2, lngY) = strQryName
        strSQL = arr_varQry(Q_SQL1, lngY)
        If lngY = 0& Or lngY = 1& Then
          strSQL = StringReplace(strSQL, strDateTemplate, strNewDate)  ' ** Module Function: modStringFuncs.
          If strActNo <> strQryTemplate Then
            strSQL = StringReplace(strSQL, "00" & strQryTemplate, strAccountNo)  ' ** Module Function: modStringFuncs.
          End If
        Else
          If strActNo <> strQryTemplate Then
            strSQL = StringReplace(strSQL, "_" & SEQ_OLD & "_" & strQryTemplate & "_", "_" & SEQ_NEW & "_" & strActNo & "_")  ' ** Module Function: modStringFuncs.
          Else
            strSQL = StringReplace(strSQL, "_" & SEQ_OLD & "_", "_" & SEQ_NEW & "_")  ' ** Module Function: modStringFuncs.
          End If
        End If
        arr_varQry(Q_SQL2, lngY) = strSQL
        strDesc = arr_varQry(Q_DSC1, lngY)
        If lngY = 0& Or lngY = 1& Then
          If strActNo <> strQryTemplate Then
            strDesc = StringReplace(strDesc, "00" & strQryTemplate, strAccountNo)  ' ** Module Function: modStringFuncs.
          End If
          strDesc = StringReplace(strDesc, strDateTemplate, strNewDate)  ' ** Module Function: modStringFuncs.
        Else
          If strActNo <> strQryTemplate Then
            strDesc = StringReplace(strDesc, "_" & SEQ_OLD & "_" & strQryTemplate & "_", "_" & SEQ_NEW & "_" & strActNo & "_")  ' ** Module Function: modStringFuncs.
          Else
            strDesc = StringReplace(strDesc, "_" & SEQ_OLD & "_", "_" & SEQ_NEW & "_")  ' ** Module Function: modStringFuncs.
          End If
        End If
        Select Case lngY
        Case 4&, 5&, 8&, 9&, 10&
          ' ** Leave record count.
        Case Else
          intPos1 = InStr(strDesc, ";")
          If intPos1 > 0 Then
            strDesc = Left(strDesc, intPos1)
          End If
        End Select
        arr_varQry(Q_DSC2, lngY) = strDesc
      Next  ' ** lngY.

      For lngY = 0& To (lngQrys - 1&)
        Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngY), arr_varQry(Q_SQL2, lngY))
        With qdf
          Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC2, lngY))
On Error Resume Next
          .Properties.Append prp
          If ERR.Number <> 0 Then
On Error GoTo 0
            .Properties("Description") = arr_varQry(Q_DSC2, lngY)
          Else
On Error GoTo 0
          End If
        End With  ' ** qdf.
        Set prp = Nothing
        Set qdf = Nothing
        lngQrysCreated = lngQrysCreated + 1&
      Next  ' ** lngY.
      .QueryDefs.Refresh

      For lngY = 0& To (lngQrys - 1&)
        Select Case lngY
        Case 0&, 1&, 2&, 3&
          ' ** Get record counts.
          Set qdf = .QueryDefs(arr_varQry(Q_QNAM2, lngY))
          Set rst = qdf.OpenRecordset
          With rst
            If .BOF = True And .EOF = True Then
              lngRecs = 0&
            Else
              .MoveLast
              lngRecs = .RecordCount
            End If
            .Close
          End With
          strDesc = qdf.Properties("Description")
          strDesc = strDesc & " " & CStr(lngRecs) & "."
          qdf.Properties("Description") = strDesc
          Set rst = Nothing
          Set qdf = Nothing
        Case 6&, 7&
          ' ** Asset count.
          Set qdf = .QueryDefs(arr_varQry(Q_QNAM2, lngY))
          Set rst = qdf.OpenRecordset
          With rst
            '.MoveLast
            If .BOF = True And .EOF = True Then
              lngAssets = 0&
            Else
              .MoveLast
              lngAssets = .RecordCount
            End If
            .Close
          End With
          strDesc = qdf.Properties("Description")
          strDesc = strDesc & " " & CStr(lngAssets) & "."
          qdf.Properties("Description") = strDesc
          Set rst = Nothing
          Set qdf = Nothing
          If lngAssets > 1& And lngY = 6& Then
            strActNo = Mid(arr_varQry(Q_QNAM2, lngY), (Len(QRY_BASE) + 1), 3)
            strAccountNo = "00" & strActNo
            lngMults = lngMults + 1&
            lngE = lngMults - 1&
            ReDim Preserve arr_varMult(M_ELEMS, lngE)
            arr_varMult(M_ACTNO, lngE) = strAccountNo
            arr_varMult(M_BDATE, lngE) = CDate(DateAdd("yyyy", 3, CDate(strDateTemplate)))
            arr_varMult(M_CNT, lngE) = lngAssets
            arr_varMult(M_QNAM, lngE) = arr_varQry(Q_QNAM2, lngY)
            'Debug.Print "'MULTI: " & CStr(lngAssets) & "  " & arr_varQry(Q_QNAM2, lngY)
          End If
        End Select
      Next  ' ** lngY.

      For lngY = 0& To (lngQrys - 1&)
        arr_varQry(Q_QNAM2, lngY) = Null
        arr_varQry(Q_SQL2, lngY) = Null
        arr_varQry(Q_DSC2, lngY) = Null
      Next  ' ** lngY

      If ((lngX + 1&) Mod 100) = 0 Then
        Debug.Print "|  " & CStr(lngX + 1&)
        Debug.Print "'|";
      ElseIf ((lngX + 1&) Mod 10) = 0 Then
        Debug.Print "|";
      Else
        Debug.Print ".";
      End If
      DoEvents

    Next  ' ** lngX.
    Debug.Print
    DoEvents

    Debug.Print "'MULTS: " & CStr(lngMults)
    DoEvents

    If lngMults > 0& Then
      Set rst = .OpenRecordset("zz_tbl_MasterTrust_Balance_02", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngMults - 1&)
          .AddNew
          ' ** ![mtmulti_id] : AutoNumber.
          ![accountno] = arr_varMult(M_ACTNO, lngX)
          ![balance_date] = arr_varMult(M_BDATE, lngX)
          ![mtmulti_cnt] = arr_varMult(M_CNT, lngX)
          ![qry_name] = arr_varMult(M_QNAM, lngX)
          ![mtmulti_datemodified] = Now()
          .Update
          'Debug.Print "'MULTI: " & CStr(arr_varMult(M_CNT, lngX)) & "  " & arr_varMult(M_QNAM, lngX)
          'DoEvents
        Next
        .Close
      End With
      Set rst = Nothing
    End If

    .Close
  End With

  Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
  DoEvents

'ACCTS: 255
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  100
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  200
'|.........|.........|.........|.........|.........|.....
'MULTS: 65
'QRYS CREATED: 3060
'DONE!

'ACCTS: 268
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  100
'|.........|.........|.........|.........|.........|.........|.........|.........|.........|.........|  200
'|.........|.........|.........|.........|.........|.........|........
'MULTS: 64
'QRYS CREATED: 3216
'DONE!
  Beep
  Debug.Print "'DONE!"
  DoEvents

  Set prp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  MT_QryCopy1 = blnRetVal

End Function

Public Function MT_QryCopy2()

  Const THIS_PROC As String = "MT_QryCopy2"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, prp As Object
  Dim lngMults As Long, arr_varMult As Variant
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strQryTemplate As String, strActNo As String, strQryName As String, strSQL As String, strDesc As String, strQryNum As String
  Dim lngQrysCreated  As Long
  Dim intPos1 As Integer
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varMult().
  Const M_MID   As Integer = 0
  Const M_ACTNO As Integer = 1
  Const M_BDATE As Integer = 2
  Const M_CNT   As Integer = 3
  Const M_QRY   As Integer = 4

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const Q_QNAM1 As Integer = 0
  Const Q_SQL1  As Integer = 1
  Const Q_DSC1  As Integer = 2
  Const Q_TYP   As Integer = 3
  Const Q_QNAM2 As Integer = 4
  Const Q_SQL2  As Integer = 5
  Const Q_DSC2  As Integer = 6

  Const QRY_BASE As String = "zzz_qry_MasterTrust_08_"

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    ' ** zz_tbl_MasterTrust_Balance_02, w/o '00090', for 12/31/2014.
    Set qdf = .QueryDefs("zzz_qry_MasterTrust_11_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngMults = .RecordCount
      .MoveFirst
      arr_varMult = .GetRows(lngMults)
      ' *************************************************
      ' ** Array: arr_varMult().
      ' **
      ' **   Field  Element  Name            Constant
      ' **   =====  =======  ==============  ==========
      ' **     1       0     mtmulti_id      M_MID
      ' **     2       1     accountno       M_ACTNO
      ' **     3       2     balance_date    M_BDATE
      ' **     4       3     mtmulti_cnt     M_CNT
      ' **     5       4     qry_name        M_QRY
      ' **
      ' *************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'MULTS: " & CStr(lngMults)
    DoEvents

    lngQrys = 0&
    ReDim arr_varQry(Q_ELEMS, 0)

    strQryTemplate = "090"
    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, Len(QRY_BASE & strQryTemplate)) = (QRY_BASE & strQryTemplate) Then
          If Val(Right(.Name, 2)) >= 20 And Val(Right(.Name, 2)) <= 26 Then
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_QNAM1, lngE) = .Name
            arr_varQry(Q_SQL1, lngE) = .SQL
            arr_varQry(Q_DSC1, lngE) = .Properties("Description")
            arr_varQry(Q_TYP, lngE) = .Type
            arr_varQry(Q_QNAM2, lngE) = Null
            arr_varQry(Q_SQL2, lngE) = Null
            arr_varQry(Q_DSC2, lngE) = Null
          End If
        End If
      End With
    Next  ' ** qdf.
    Set qdf = Nothing

    lngQrysCreated = 0&
    For lngX = 0& To (lngMults - 1&)

      strActNo = Right(arr_varMult(M_ACTNO, lngX), 3)
      'strDesc = .QueryDefs(arr_varMult(M_QRY, lngX)).Properties("Description")
      'strDesc = strDesc & " " & CStr(arr_varMult(M_CNT, lngX)) & "."
      '.QueryDefs(arr_varMult(M_QRY, lngX)).Properties("Description") = strDesc
      'strQryName = arr_varMult(M_QRY, lngX)
      'strQryName = Left(strQryName, (Len(strQryName) - 1)) & "8"
      'strDesc = .QueryDefs(strQryName).Properties("Description")
      'strDesc = strDesc & " " & CStr(arr_varMult(M_CNT, lngX)) & "."
      '.QueryDefs(strQryName).Properties("Description") = strDesc
      strQryName = Left(arr_varMult(M_QRY, lngX), (Len(arr_varMult(M_QRY, lngX)) - 2))

      For lngY = 0& To (lngQrys - 1&)
        strQryNum = Right(arr_varQry(Q_QNAM1, lngY), 2)
        arr_varQry(Q_QNAM2, lngY) = strQryName & strQryNum
        strSQL = arr_varQry(Q_SQL1, lngY)
        strSQL = StringReplace(strSQL, "_" & strQryTemplate & "_", "_" & strActNo & "_")  ' ** Module Function: modStringFuncs.
        arr_varQry(Q_SQL2, lngY) = strSQL
        strDesc = arr_varQry(Q_DSC1, lngY)
        strDesc = StringReplace(strDesc, "_" & strQryTemplate & "_", "_" & strActNo & "_")  ' ** Module Function: modStringFuncs.
        intPos1 = InStr(strDesc, ";")
        If intPos1 > 0 Then
          strDesc = Left(strDesc, intPos1)
          Select Case strQryNum
          Case "20", "21", "22"
            strDesc = strDesc & " " & CStr(arr_varMult(M_CNT, lngX)) & "."
          Case Else
            strDesc = strDesc & " 1."
          End Select
        End If
        arr_varQry(Q_DSC2, lngY) = strDesc
      Next  ' ** lngY.

      For lngY = 0& To (lngQrys - 1&)
        Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngY), arr_varQry(Q_SQL2, lngY))
        With qdf
          Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC2, lngY))
On Error Resume Next
          .Properties.Append prp
          If ERR.Number <> 0 Then
On Error GoTo 0
            .Properties("Description") = arr_varQry(Q_DSC2, lngY)
          Else
On Error GoTo 0
          End If
        End With  ' ** qdf.
        Set qdf = Nothing
        lngQrysCreated = lngQrysCreated + 1&
      Next  ' ** lngY

      For lngY = 0& To (lngQrys - 1&)
        arr_varQry(Q_QNAM2, lngY) = Null
        arr_varQry(Q_SQL2, lngY) = Null
        arr_varQry(Q_DSC2, lngY) = Null
      Next  ' ** lngY

    Next  ' ** lngX.

    .Close
  End With

  Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
  DoEvents

'MULTS: 64
'QRYS CREATED: 448
'DONE!

'MULTS: 63
'QRYS CREATED: 441
'DONE!
  Beep
  Debug.Print "'DONE!"
  DoEvents

  Set prp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  MT_QryCopy2 = blnRetVal

End Function

Public Function MT_QryCopy3() As Boolean

  Const THIS_PROC As String = "MT_QryCopy3"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strAccountNo As String, strActNo As String, strQryName As String, strQryNum As String
  Dim lngQrysRun As Long
  Dim blnSkip As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const Q_ACTNO As Integer = 0
  Const Q_QNAM1 As Integer = 1
  Const Q_QNAM2 As Integer = 2

  Const QRY_BASE As String = "zzz_qry_MasterTrust_08_"

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    lngQrys = 0&
    ReDim arr_varQry(Q_ELEMS, 0)

    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
          If .Type = dbQAppend Then
            strActNo = Mid(.Name, (Len(QRY_BASE) + 1), 3)
            strAccountNo = "00" & strActNo
            strQryNum = Right(.Name, 2)
            If strQryNum = "12" Then
              lngQrys = lngQrys + 1&
              lngE = lngQrys - 1&
              ReDim Preserve arr_varQry(Q_ELEMS, lngE)
              arr_varQry(Q_ACTNO, lngE) = strAccountNo
              arr_varQry(Q_QNAM1, lngE) = .Name
              arr_varQry(Q_QNAM2, lngE) = Null
            Else
              For lngY = 0& To (lngQrys - 1&)
                If arr_varQry(Q_ACTNO, lngY) = strAccountNo Then
                  arr_varQry(Q_QNAM2, lngY) = .Name
                  Exit For
                End If
              Next
            End If
          End If
        End If
      End With
    Next  ' ** qdf.
    Set qdf = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    blnSkip = True
    If blnSkip = False Then
      Set rst = .OpenRecordset("zz_tbl_MasterTrust_Balance_03", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngQrys - 1&)
          .AddNew
          ' ** ![mtqry_id] : AutoNumber.
          ![accountno] = arr_varQry(Q_ACTNO, lngX)
          ![balance_date] = #12/31/2014#
          ![qry_name1] = arr_varQry(Q_QNAM1, lngX)
          ![qry_name2] = arr_varQry(Q_QNAM2, lngX)
          ![mtqry_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With
      Set rst = Nothing
    End If  ' ** blnSkip.

    For lngX = 0& To (lngQrys - 1&)
      If IsNull(arr_varQry(Q_QNAM2, lngX)) = False Then
        arr_varQry(Q_QNAM1, lngX) = arr_varQry(Q_QNAM2, lngX)
      End If
    Next  ' ** lngX.

    blnSkip = False
    If blnSkip = False Then
      lngQrysRun = 0&
      For lngX = 0& To (lngQrys - 1&)
        Set qdf = .QueryDefs(arr_varQry(Q_QNAM1, lngX))
        qdf.Execute
        DoEvents
        lngQrysRun = lngQrysRun + 1&
      Next
      Set qdf = Nothing
    End If  ' ** blnSkip.

    .Close
  End With

  Debug.Print "'QRYS RUN: " & CStr(lngQrysRun)
  DoEvents

  Beep
  Debug.Print "'DONE!"
  DoEvents

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  MT_QryCopy3 = blnRetVal

End Function

Public Function MT_FindBalance() As Boolean

  Const THIS_PROC As String = "MT_FindBalance"

  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim lngRefs As Long, arr_varRef As Variant
  Dim lngQryType As Long, strQryName As String
  Dim blnSkip As Boolean
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varRef().
  Const RF_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const RF_QRY As Integer = 0
  Const RF_SQL As Integer = 1

  blnRetVal = True

  blnSkip = True
  If blnSkip = False Then

    arr_varRef = Qry_FindStr_rel("Balance")  ' ** Module Function: modQueryFunctions.

    lngRefs = UBound(arr_varRef, 2) + 1&

    Debug.Print "'REFS: " & CStr(lngRefs)
    DoEvents

    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("zz_tbl_MasterTrust_Balance_04", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngRefs - 1&)
          .AddNew
          ' ** ![mtref_id] : AutoNumber.
          ![qry_name] = arr_varRef(RF_QRY, lngX)
          ![qry_sql] = arr_varRef(RF_SQL, lngX)
          ![mtref_datemodified] = Now()
           .Update
        Next
        .Close
      End With
      Set rst = Nothing
      .Close
    End With

  End If  ' ** blnSkip.

  blnSkip = False
  If blnSkip = False Then

    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("zz_tbl_MasterTrust_Balance_04", dbOpenDynaset, dbConsistent)
      rst.MoveLast
      lngRefs = rst.RecordCount
      rst.MoveFirst
      Debug.Print "'QRYS: " & CStr(lngRefs)
      DoEvents
      For lngX = 1& To lngRefs
        strQryName = rst![qry_name]
        lngQryType = .QueryDefs(strQryName).Type
        rst.Edit
        rst![qrytype_type] = lngQryType
        If InStr(strQryName, "MasterTrust") > 0 Then
          rst![IsMT] = True
        End If
        rst![mtref_datemodified] = Now()
        rst.Update
        If lngX < lngRefs Then rst.MoveNext
      Next
      rst.Close
      Set rst = Nothing
      .Close
    End With

  End If  ' ** blnSkip.

  Beep
  Debug.Print "'DONE!"

  Set rst = Nothing
  Set dbs = Nothing

  MT_FindBalance = blnRetVal

End Function

Public Function MT_QryCopy4() As Boolean

  Const THIS_PROC As String = "MT_QryCopy4"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strActNo As String, strAccountNo As String, strDesc As String
  Dim lngRecs As Long, lngNotFounds As Long
  Dim blnFound As Boolean, blnSkip As Boolean
  Dim intPos1 As Integer
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const Q_ACTNO As Integer = 0
  Const Q_QNAM1 As Integer = 1
  Const Q_CNT   As Integer = 2
  Const Q_FND   As Integer = 3

  Const QRY_BASE As String = "zzz_qry_MasterTrust_08_"

  blnRetVal = True

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    blnSkip = True
    If blnSkip = False Then

      For Each qdf In .QueryDefs
        With qdf
          If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
            strActNo = Mid(.Name, (Len(QRY_BASE) + 1), 3)
            strAccountNo = "00" & strActNo
            blnFound = False
            For lngY = 0& To (lngQrys - 1&)
              If arr_varQry(Q_ACTNO, lngY) = strAccountNo Then
                blnFound = True
                arr_varQry(Q_CNT, lngY) = arr_varQry(Q_CNT, lngY) + 1&
                Exit For
              End If
            Next
            If blnFound = False Then
              lngQrys = lngQrys + 1&
              lngE = lngQrys - 1&
              ReDim Preserve arr_varQry(Q_ELEMS, lngE)
              arr_varQry(Q_ACTNO, lngE) = strAccountNo
              arr_varQry(Q_QNAM1, lngE) = .Name
              arr_varQry(Q_CNT, lngE) = CLng(1)
              arr_varQry(Q_FND, lngE) = CBool(False)
            End If
          End If
        End With
      Next
      Set qdf = Nothing

      Set qdf = .QueryDefs("zzz_qry_MasterTrust_03_05")
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          blnFound = False
          For lngY = 0& To (lngQrys - 1&)
            If arr_varQry(Q_ACTNO, lngY) = ![accountno] Then
              blnFound = True
              arr_varQry(Q_FND, lngY) = CBool(True)
              Exit For
            End If
          Next
          If blnFound = False Then
            Debug.Print "'QRY NOT FOUND: " & ![accountno]
            DoEvents
            lngNotFounds = lngNotFounds + 1&
          End If
          If lngX < lngRecs Then .MoveNext
        Next
        .Close
      End With
      Set rst = Nothing
      Set qdf = Nothing

      Debug.Print "'NOT FOUND: " & CStr(lngNotFounds)
      DoEvents

    End If  ' ** blnSkip.

    blnSkip = False
    If blnSkip = False Then
      For Each qdf In .QueryDefs
        With qdf
          If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
            ' ** Union of .._05_136_01, .._05_136_02; 642.
            If Right(.Name, 2) = "07" Then
              strDesc = .Properties("Description")
              If Right(strDesc, 2) = "0." Then
                Debug.Print "'ZERO!  " & .Name
                DoEvents
              End If
              'intPos1 = InStr(strDesc, "_05_136_")
              'If intPos1 > 0 Then
              '  strActNo = Mid(.Name, (Len(QRY_BASE) + 1), 3)
              '  strDesc = StringReplace(strDesc, "_05_136_", "_08_" & strActNo & "_")  ' ** Module Function: modStringFuncs.
              '  .Properties("Description") = strDesc
              'End If
              'intPos1 = InStr(strDesc, "_15_")
              'If intPos1 > 0 Then
              '  strActNo = Mid(.Name, (Len(QRY_BASE) + 1), 3)
              '  strDesc = StringReplace(strDesc, ".._05_243_15", ".._08_" & strActNo)  ' ** Module Function: modStringFuncs.
              '  .Properties("Description") = strDesc
              'End If
            End If
          End If
        End With
      Next
    End If  ' ** blnSkip.

    .Close
  End With

  Beep
  Debug.Print "'DONE!"

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  MT_QryCopy4 = blnRetVal

End Function

Public Function AutoNum_Holes() As Boolean

  Const THIS_PROC As String = "AutoNum_Holes"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngHoles As Long, arr_varHole() As Variant
  Dim strTblName1 As String, strTblName2 As String, strFldName As String, strQryName1 As String, strQryName2 As String
  Dim lngLastID As Long, lngRecs As Long
  Dim blnSkip As Boolean
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varHole().
  Const H_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const H_ID As Integer = 0

  blnRetVal = True

  ' ** zzz_qry_MasterTrust_19_04 (Union of zzz_qry_MasterTrust_19_02 (LedgerArchive,
  ' ** just journalno), zzz_qry_MasterTrust_19_03 (Ledger, just journalno)), sorted.
  strQryName1 = "zzz_qry_MasterTrust_19_01"
  strFldName = "journalno"
  'strTblName1 = "LedgerArchive"
  'strTblName2 = "ledger"

  Set dbs = CurrentDb
  With dbs

    strQryName2 = "qryTmp_Table_Empty_077_tblMark"
    If QueryExists(strQryName2) = False Then  ' ** Module Function: modFileUtilities.
      Beep
      Debug.Print "'QRY MOVED!  " & strQryName2
    Else

      ' ** Empty tblMark.
      Set qdf = .QueryDefs(strQryName2)
      qdf.Execute
      Set qdf = Nothing

      lngHoles = 0&
      ReDim arr_varHole(H_ELEMS, 0)

      Set qdf = .QueryDefs(strQryName1)
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        lngLastID = 0&
        For lngX = 1& To lngRecs
          If .Fields(strFldName) <> lngLastID + 1& Then
            Do Until .Fields(strFldName) = (lngLastID + 1&)
              lngLastID = lngLastID + 1&
              lngHoles = lngHoles + 1&
              lngE = lngHoles - 1&
              ReDim Preserve arr_varHole(H_ELEMS, lngE)
              arr_varHole(H_ID, lngE) = lngLastID
            Loop
          End If
          lngLastID = .Fields(strFldName)
          If lngX < lngRecs Then .MoveNext
        Next
        .Close
      End With
      Set rst = Nothing

      Debug.Print "'HOLES: " & CStr(lngHoles) & "  QRY: " & strQryName1
      DoEvents

      If lngHoles > 0& Then
        Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbAppendOnly)
        With rst
          For lngX = 0& To (lngHoles - 1&)
            .AddNew
            ![unique_id] = arr_varHole(H_ID, lngX)
            ![mark] = False
            ![Value] = Null
            .Update
          Next
          .Close
        End With
        Set rst = Nothing
      End If

      blnSkip = True
      If blnSkip = False Then

        lngHoles = 0&
        ReDim arr_varHole(H_ELEMS, 0)

        ' ** lngLastID remains from first table.
        Set rst = .OpenRecordset(strTblName2, dbOpenDynaset, dbReadOnly)
        With rst
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            If .Fields(strFldName) <> lngLastID + 1& Then
              Do Until .Fields(strFldName) = (lngLastID + 1&)
                lngLastID = lngLastID + 1&
                lngHoles = lngHoles + 1&
                lngE = lngHoles - 1&
                ReDim Preserve arr_varHole(H_ELEMS, lngE)
                arr_varHole(H_ID, lngE) = lngLastID
              Loop
            End If
            lngLastID = .Fields(strFldName)
            If lngX < lngRecs Then .MoveNext
          Next
          .Close
        End With
        Set rst = Nothing

        Debug.Print "'HOLES: " & CStr(lngHoles) & "  TBL: " & strTblName2
        DoEvents

        If lngHoles > 0& Then
          Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbAppendOnly)
          With rst
            For lngX = 0& To (lngHoles - 1&)
              .AddNew
              ![unique_id] = arr_varHole(H_ID, lngX)
              ![mark] = False
              ![Value] = Null
              .Update
            Next
            .Close
          End With
          Set rst = Nothing
        End If

      End If  ' ** blnSkip.

    End If  ' ** QueryExists().

    .Close
  End With

  Beep
  Debug.Print "'DONE!"

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  AutoNum_Holes = blnRetVal

End Function
