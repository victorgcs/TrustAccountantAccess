Attribute VB_Name = "zz_mod_IndexDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_IndexDocFuncs"

'VGC 11/23/2016: CHANGES!

Private blnRetValx As Boolean
' **

Public Function QuikIdxDoc() As Boolean
  Const THIS_PROC As String = "QuikIdxDoc"
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Idx_ChkDocQrys(False) = True Then  ' ** Function: Below.
      blnRetValx = Idx_Doc  ' ** Function: Below.
      DoEvents
      DoBeeps  ' ** Module Function: modWindowsFuncs.
      Debug.Print "'FINISHED!"
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED Idx_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikIdxDoc = blnRetValx
End Function

Public Function Idx_Doc() As Boolean
' ** Document all indexes to tblIndex and tblIndex_Field.

  Const THIS_PROC As String = "Idx_Doc"

  Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef, idx As DAO.index, fld As DAO.Field
  Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset, rst4 As DAO.Recordset
  Dim strThisFile As String, strThisPath As String
  Dim lngDbs As Long, arr_varDb() As Variant
  Dim lngTdfs As Long, arr_varTmpTdf() As Variant, arr_varTdf As Variant
  Dim lngIdxs As Long, arr_varTmpIdx() As Variant, arr_varIdx As Variant
  Dim lngIdxFlds As Long, arr_varTmpIdxFld() As Variant, arr_varIdxFld As Variant
  Dim lngTblID As Long, lngFldID As Long
  Dim intIdxOrd As Integer, lngIdxsTot As Long
  Dim lngIdxXs As Long, arr_varIdxX As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim strPath As String, strFile As String
  Dim blnFound As Boolean
  Dim lngRecs As Long
  Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngDE As Long, lngTE As Long, lngIE As Long, lngFE As Long

  ' ** Array: arr_varDb().
  Const B_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const B_ID   As Integer = 0
  Const B_NAME As Integer = 1
  Const B_PATH As Integer = 2
  Const B_RCNT As Integer = 3  ' ** Relationship count.
  Const B_TCNT As Integer = 4  ' ** Table count.
  Const B_TARR As Integer = 5

  ' ** Array: arr_varTdf().
  Const T_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const T_DBSID As Integer = 0
  Const T_ID    As Integer = 1
  Const T_NAME  As Integer = 2
  Const T_SRC   As Integer = 3
  Const T_ICNT  As Integer = 4
  Const T_IARR  As Integer = 5
  Const T_DELEM As Integer = 6
  Const T_EXIST As Integer = 7

  ' ** Array: arr_varIdx().
  Const I_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const I_TBLID   As Integer = 0
  Const I_TELEM   As Integer = 1
  Const I_ID      As Integer = 2
  Const I_NAME    As Integer = 3
  Const I_PRIME   As Integer = 4
  Const I_UNIQUE  As Integer = 5
  Const I_IGNULLS As Integer = 6
  Const I_REQUIRE As Integer = 7
  Const I_FOREIGN As Integer = 8
  Const I_FCNT    As Integer = 9
  Const I_FARR    As Integer = 10

  ' ** Array: arr_varIdxFld().
  Const IF_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const IF_IDXID As Integer = 0
  Const IF_IELEM As Integer = 1
  Const IF_FLDID As Integer = 2
  Const IF_ID    As Integer = 3
  Const IF_NAME  As Integer = 4
  Const IF_ATTR  As Integer = 5

  ' ** Array: arr_varIdxX().
  Const IX_DID  As Integer = 0
  Const IX_DNAM As Integer = 1
  Const IX_TID  As Integer = 2
  Const IX_TNAM As Integer = 3
  Const IX_ID   As Integer = 4
  Const IX_NAM  As Integer = 5
  Const IX_FND  As Integer = 6

  Const D_ELEMS As Integer = 0  ' ** Array's first-element UBound().

  blnRetValx = True

  strThisFile = CurrentAppName  ' ** Module Function: modFileUtilities.
  strThisPath = CurrentAppPath & LNK_SEP  ' ** Module Function: modFileUtilities.

  TmpTable_Chk True  ' ** Module Function: modUtilities.

  lngDbs = 0&
  ReDim arr_varDb(B_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    ' ** Get a list of databases to document.
    Set rst1 = .OpenRecordset("tblDatabase", dbOpenDynaset, dbConsistent)
    With rst1
      If .BOF = True And .EOF = True Then
        ' ** Table is empty.
        Stop
      Else
        ' ** Get the database list.
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          If ![dbs_name] <> "TAJrnTmp.mdb" Then  ' ** Skip this for now.
            lngDbs = lngDbs + 1&
            lngDE = lngDbs - 1&
            ReDim Preserve arr_varDb(B_ELEMS, lngDE)
            arr_varDb(B_ID, lngDE) = ![dbs_id]
            arr_varDb(B_NAME, lngDE) = ![dbs_name]
            arr_varDb(B_PATH, lngDE) = ![dbs_path]
            arr_varDb(B_TCNT, lngDE) = CLng(0)
          End If
          If lngX < lngRecs Then .MoveNext
        Next
      End If
      .Close
    End With

    ' ** Get a list of tables to document.
    For lngDE = 0& To (lngDbs - 1&)
      lngTdfs = 0&
      ReDim arr_varTmpTdf(T_ELEMS, 0)
      Set rst1 = .OpenRecordset("tblDatabase_Table", dbOpenDynaset, dbReadOnly)
      With rst1
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          If ![dbs_id] = arr_varDb(B_ID, lngDE) Then
            lngTdfs = lngTdfs + 1&
            lngTE = lngTdfs - 1&
            ReDim Preserve arr_varTmpTdf(T_ELEMS, lngTE)
            arr_varTmpTdf(T_DBSID, lngTE) = ![dbs_id]
            arr_varTmpTdf(T_ID, lngTE) = ![tbl_id]
            arr_varTmpTdf(T_NAME, lngTE) = ![tbl_name]
            arr_varTmpTdf(T_SRC, lngTE) = ![tbl_sourcetablename]
            arr_varTmpTdf(T_ICNT, lngTE) = 0&
            arr_varTmpTdf(T_IARR, lngTE) = Empty
            arr_varTmpTdf(T_DELEM, lngTE) = lngDE
            arr_varTmpTdf(T_EXIST, lngTE) = CBool(False)
          End If
          If lngX < lngRecs Then .MoveNext
        Next
        .Close
      End With
      For lngX = 0& To (lngTdfs - 1&)
        If IsNull(arr_varTmpTdf(T_SRC, lngX)) = True Then
          arr_varTmpTdf(T_SRC, lngX) = arr_varTmpTdf(T_NAME, lngX)
        End If
      Next
      arr_varDb(B_TCNT, lngDE) = lngTdfs
      arr_varDb(B_TARR, lngDE) = arr_varTmpTdf
    Next

    .Close
  End With
  Set rst1 = Nothing
  Set dbs = Nothing

  'lngIdxsTot = 0&
  'For lngDE = 0& To (lngDbs - 1&)
  '  arr_varTdf = arr_varDb(B_TARR, lngDE)
  '  lngTdfs = arr_varDb(B_TCNT, lngDE)
  '  For lngTE = 0& To (lngTdfs - 1&)
  '    lngIdxsTot = lngIdxsTot + 1&
  '    If lngIdxsTot > 140& Then
  '      Debug.Print "'" & CStr(lngIdxsTot) & ". DB: " & arr_varDb(B_NAME, lngDE) & "  TBL: " & arr_varTdf(T_NAME, lngTE)
  '    End If
  '    If lngIdxsTot >= 145& Then Exit For
  '  Next
  '  If lngIdxsTot >= 145& Then Exit For
  'Next

  ' ** Collect the info.
  For lngDE = 0& To (lngDbs - 1&)
    Debug.Print IIf(lngDE = 0&, "'", vbNullString) & "IDX_DOC(): " & arr_varDb(B_NAME, lngDE)
    Debug.Print "'|";
    DoEvents
    If arr_varDb(B_NAME, lngDE) = strThisFile Then
      Set dbs = CurrentDb
    Else
      Set wrk = DBEngine.CreateWorkspace("Tmp", "superuser", TA_SEC, dbUseJet)
      Set dbs = wrk.OpenDatabase(arr_varDb(B_PATH, lngDE) & LNK_SEP & arr_varDb(B_NAME, lngDE), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
    End If
    With dbs

      lngIdxsTot = 0&
      arr_varTdf = arr_varDb(B_TARR, lngDE)
      lngTdfs = arr_varDb(B_TCNT, lngDE)
      .TableDefs.Refresh
      .TableDefs.Refresh

      ' ** Check if all the tables are currenty present.
      For lngTE = 0& To (lngTdfs - 1&)
        For Each tdf In .TableDefs
          With tdf
            If .Name = arr_varTdf(T_NAME, lngTE) Then
              arr_varTdf(T_EXIST, lngTE) = CBool(True)
              Exit For
            End If
          End With  ' ** tdf.
        Next  ' ** tdf.
      Next  ' ** lngTE.

      For lngTE = 0& To (lngTdfs - 1&)
        If arr_varTdf(T_EXIST, lngTE) = True Then
          lngIdxs = 0&
          ReDim arr_varTmpIdx(I_ELEMS, 0)
          Set tdf = .TableDefs(arr_varTdf(T_NAME, lngTE))
          With tdf
            arr_varTdf(T_ICNT, lngTE) = .Indexes.Count
            If .Indexes.Count > 0 Then
              ' ** Collect Index list.
              For Each idx In .Indexes
                lngIdxFlds = 0&
                ReDim arr_varIdxFld(IF_ELEMS, 0)
                With idx
                  lngIdxs = lngIdxs + 1&
                  lngIE = lngIdxs - 1&
                  ReDim Preserve arr_varTmpIdx(I_ELEMS, lngIE)
                  arr_varTmpIdx(I_TBLID, lngIE) = arr_varTdf(T_ID, lngTE)  '[tbl_id]
                  arr_varTmpIdx(I_TELEM, lngIE) = lngTE
                  arr_varTmpIdx(I_ID, lngIE) = CLng(0)  '[idx_id]
                  arr_varTmpIdx(I_NAME, lngIE) = .Name
                  arr_varTmpIdx(I_PRIME, lngIE) = .Primary
                  arr_varTmpIdx(I_UNIQUE, lngIE) = .Unique
                  arr_varTmpIdx(I_IGNULLS, lngIE) = .IgnoreNulls
                  arr_varTmpIdx(I_REQUIRE, lngIE) = .Required
                  arr_varTmpIdx(I_FOREIGN, lngIE) = .Foreign
                  arr_varTmpIdx(I_FCNT, lngIE) = .Fields.Count
                  ' ** Collect Index Field list.
                  For Each fld In .Fields
                    With fld
                      lngIdxFlds = lngIdxFlds + 1&
                      lngFE = lngIdxFlds - 1&
                      ReDim Preserve arr_varTmpIdxFld(IF_ELEMS, lngFE)
                      arr_varTmpIdxFld(IF_IDXID, lngFE) = CLng(0)  '[idx_id]
                      arr_varTmpIdxFld(IF_IELEM, lngFE) = lngIE
                      arr_varTmpIdxFld(IF_FLDID, lngFE) = CLng(0)  '[fld_id]
                      arr_varTmpIdxFld(IF_ID, lngFE) = CLng(0)  '[idxfld_id]
                      arr_varTmpIdxFld(IF_NAME, lngFE) = .Name
                      arr_varTmpIdxFld(IF_ATTR, lngFE) = .Attributes
                    End With  ' ** This field: fld.
                  Next        ' ** For each field: fld.
                End With  ' ** This index: idx.
                arr_varTmpIdx(I_FARR, lngIE) = arr_varTmpIdxFld
              Next        ' ** For each index: idx.
              arr_varTdf(T_IARR, lngTE) = arr_varTmpIdx
            End If
          End With  ' ** This table: tdf.
          ' ** Update tbl_idx_cnt in tblDatabase_Table, by specified [dbid], [tbid], [idxcnt].
          Set qdf = CurrentDb.QueryDefs("zz_qry_Index_02")
          With qdf.Parameters
            ![dbid] = arr_varDb(B_ID, lngDE)
            ![tbid] = arr_varTdf(T_ID, lngTE)
            ![idxcnt] = arr_varTdf(T_ICNT, lngTE)
          End With
          qdf.Execute
          lngIdxsTot = lngIdxsTot + arr_varTdf(T_ICNT, lngTE)
        End If  ' ** T_EXIST.
        If (lngTE + 1&) Mod 100 = 0 And lngTE <> 0& Then
          Debug.Print "| " & CStr(lngTE + 1&) & " of " & CStr(lngTdfs)
          Debug.Print "'|";
        ElseIf (lngTE + 1&) Mod 10 = 0 And lngTE <> 0& Then
          Debug.Print "|";
        Else
          Debug.Print ".";
        End If
        DoEvents
      Next  ' ** For each table: lngTE.

      ' ** Update dbs_idx_cnt in tblDatabase, by specified [dbid], [idxcnt].
      Set qdf = CurrentDb.QueryDefs("zz_qry_Index_03")
      With qdf.Parameters
        ![dbid] = arr_varDb(B_ID, lngDE)
        ![idxcnt] = lngIdxsTot
      End With
      qdf.Execute

      arr_varDb(B_TARR, lngDE) = arr_varTdf

      .Close
    End With
    Debug.Print vbNullString
    Debug.Print "'";
    DoEvents
  Next          ' ** For each database: lngDE.

  ' ** Index field attributes enumeration:
  ' **       1  dbDescending      The field is sorted in descending (Z to A or 100 to 0) order;
  ' **                            this option applies only  to a Field object in a Fields
  ' **                            collection of an Index object. If you omit this constant,
  ' **                            the field is sorted in ascending (A to Z or 0 to 100) order.
  ' **                            This is the default value for Index and TableDef fields
  ' **                            (Microsoft Jet workspaces only).
  ' **       1  dbFixedField      The field size is fixed (default for Numeric fields).
  ' **       2  dbVariableField   The field size is variable (Text fields only).
  ' **      16  dbAutoIncrField   The field value for new records is automatically incremented to
  ' **                            a unique Long integer that can't be changed (in a Microsoft Jet
  ' **                            workspace, supported only for Microsoft Jet database(.mdb) tables).
  ' **      32  dbUpdatableField  The field value can be changed.
  ' **    8192  dbSystemField     The field stores replication information for replicas; you can't
  ' **                            delete this type of field (Microsoft Jet workspaces only).
  ' **   32768  dbHyperlinkField  The field contains hyperlink information (Memo fields only).

  ' ** Now add the info to the tables.
  Set dbs = CurrentDb
  With dbs

    ' ** tblIndex, with add'l fields, all records (except TAJrnTmp.mdb).
    Set qdf = .QueryDefs("zz_qry_Index_05")
    Set rst1 = qdf.OpenRecordset
    With rst1
      If .BOF = True And .EOF = True Then
        lngIdxXs = 0&
      Else
        .MoveLast
        lngIdxXs = .RecordCount
        .MoveFirst
        arr_varIdxX = .GetRows(lngRecs)
        ' *********************************************
        ' ** Array: arr_varIdxX()
        ' **
        ' **   Field  Element  Name        Constant
        ' **   =====  =======  ==========  ==========
        ' **     1       0     dbs_id      IX_DID
        ' **     2       1     dbs_name    IX_DNAM
        ' **     3       2     tbl_id      IX_TID
        ' **     4       3     tbl_name    IX_TNAM
        ' **     5       4     idx_id      IX_ID
        ' **     6       5     idx_name    IX_NAM
        ' **     7       6     idx_fnd     IX_FND
        ' **
        ' *********************************************
      End If
      .Close
    End With

    For lngDE = 0& To (lngDbs - 1&)
      Debug.Print IIf(lngDE = 0&, vbNullString, "'") & "WRITING: " & arr_varDb(B_NAME, lngDE)
      DoEvents
      lngTdfs = arr_varDb(B_TCNT, lngDE)
      arr_varTdf = arr_varDb(B_TARR, lngDE)

      ' ** tblDatabase_Table by specified [dbid].
      Set qdf = .QueryDefs("zz_qry_Index_01")
      With qdf.Parameters
        ![dbid] = arr_varDb(B_ID, lngDE)
      End With
      Set rst1 = qdf.OpenRecordset()

      For lngTE = 0& To (lngTdfs - 1&)

        lngIdxs = arr_varTdf(T_ICNT, lngTE)
        arr_varIdx = arr_varTdf(T_IARR, lngTE)

        With rst1
          .FindFirst "[dbs_id] = " & CStr(arr_varDb(B_ID, lngDE)) & " " & _
            "AND [tbl_name] = '" & arr_varTdf(T_NAME, lngTE) & "'"
          If .NoMatch = True Then
            Stop
          Else
            lngTblID = ![tbl_id]
          End If
          arr_varTdf(T_ID, lngTE) = lngTblID
          If Compare_StringA_StringB(![tbl_name], "<>", arr_varTdf(T_NAME, lngTE)) = True Then  ' ** Module Function: modStringFuncs.
            .Edit
            ![tbl_name] = arr_varTdf(T_NAME, lngTE)  ' ** Maybe capitalization changed.
            ![tbl_datemodified] = Now()
            .Update
          End If
          If arr_varTdf(T_SRC, lngTE) <> vbNullString Then
            If IsNull(![tbl_sourcetablename]) = True Then
              .Edit
              ![tbl_sourcetablename] = arr_varTdf(T_SRC, lngTE)
              ![tbl_datemodified] = Now()
              .Update
            Else
              If Compare_StringA_StringB(![tbl_sourcetablename], "<>", arr_varTdf(T_SRC, lngTE)) = True Then
                .Edit
                ![tbl_sourcetablename] = arr_varTdf(T_SRC, lngTE)
                ![tbl_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![tbl_sourcetablename]) = False Then
              .Edit
              ![tbl_sourcetablename] = Null
              ![tbl_datemodified] = Now()
              .Update
            End If
          End If
          If ![tbl_idx_cnt] <> arr_varTdf(T_ICNT, lngTE) Then
            .Edit
            ![tbl_idx_cnt] = arr_varTdf(T_ICNT, lngTE)
            ![tbl_datemodified] = Now()
            .Update
          End If
        End With

        ' ** Get list of all indexes for this table.
        Set rst2 = dbs.OpenRecordset("tblIndex", dbOpenDynaset, dbConsistent)

        For lngIE = 0& To (lngIdxs - 1&)

          lngIdxFlds = arr_varIdx(I_FCNT, lngIE)
          arr_varIdxFld = arr_varIdx(I_FARR, lngIE)

          With rst2
            .FindFirst "[idx_name] = '" & arr_varIdx(I_NAME, lngIE) & "' And " & _
              "[dbs_id] = " & CStr(arr_varDb(B_ID, lngDE)) & " And [tbl_id] = " & CStr(arr_varTdf(T_ID, lngTE))
            If .NoMatch = True Then
              .AddNew
              ![dbs_id] = arr_varDb(B_ID, lngDE)
              ![tbl_id] = arr_varTdf(T_ID, lngTE)
              ![idx_name] = arr_varIdx(I_NAME, lngIE)
              ![idx_primary] = arr_varIdx(I_PRIME, lngIE)
              ![idx_unique] = arr_varIdx(I_UNIQUE, lngIE)
              ![idx_ignore_nulls] = arr_varIdx(I_IGNULLS, lngIE)
              ![idx_required] = arr_varIdx(I_REQUIRE, lngIE)
              ![idx_foreign] = arr_varIdx(I_FOREIGN, lngIE)
              ![idx_fld_cnt] = arr_varIdx(I_FCNT, lngIE)
              ![idx_datemodified] = Now()
              .Update
              .Bookmark = .LastModified
              arr_varIdx(I_ID, lngIE) = ![idx_id]
            Else
              arr_varIdx(I_ID, lngIE) = ![idx_id]
              If ![idx_primary] <> arr_varIdx(I_PRIME, lngIE) Then
                .Edit
                ![idx_primary] = arr_varIdx(I_PRIME, lngIE)
                ![idx_datemodified] = Now()
                .Update
              End If
              If ![idx_unique] <> arr_varIdx(I_UNIQUE, lngIE) Then
                .Edit
                ![idx_unique] = arr_varIdx(I_UNIQUE, lngIE)
                ![idx_datemodified] = Now()
                .Update
              End If
              If ![idx_ignore_nulls] <> arr_varIdx(I_IGNULLS, lngIE) Then
                .Edit
                ![idx_ignore_nulls] = arr_varIdx(I_IGNULLS, lngIE)
                ![idx_datemodified] = Now()
                .Update
              End If
              If ![idx_required] <> arr_varIdx(I_REQUIRE, lngIE) Then
                .Edit
                ![idx_required] = arr_varIdx(I_REQUIRE, lngIE)
                ![idx_datemodified] = Now()
                .Update
              End If
              If ![idx_foreign] <> arr_varIdx(I_FOREIGN, lngIE) Then
                .Edit
                ![idx_foreign] = arr_varIdx(I_FOREIGN, lngIE)
                ![idx_datemodified] = Now()
                .Update
              End If
              If ![idx_fld_cnt] <> arr_varIdx(I_FCNT, lngIE) Then
                .Edit
                ![idx_fld_cnt] = arr_varIdx(I_FCNT, lngIE)
                ![idx_datemodified] = Now()
                .Update
              End If
            End If
          End With

          ' ** Add idx_id to the index-field array.
          For lngX = 0& To (lngIdxFlds - 1&)
            arr_varIdxFld(IF_IDXID, lngX) = arr_varIdx(I_ID, lngIE)
          Next

          ' ** tblDatabase_Table_Field, by specified [tbid].
          Set qdf = .QueryDefs("zz_qry_Index_04")
          With qdf.Parameters
            ![tbid] = arr_varIdx(I_TBLID, lngIE)
          End With
          Set rst3 = qdf.OpenRecordset()

          ' ** Get list of all fields for this index.
          Set rst4 = dbs.OpenRecordset("tblIndex_Field", dbOpenDynaset, dbConsistent)

          intIdxOrd = 0
          For lngFE = 0& To (lngIdxFlds - 1&)

            With rst3
              .FindFirst "[dbs_id] = " & CStr(arr_varDb(B_ID, lngDE)) & " And " & _
                "[tbl_id] = " & CStr(arr_varIdx(I_TBLID, lngIE)) & " And " & _
                "[fld_name] = '" & arr_varIdxFld(IF_NAME, lngFE) & "'"
              If .NoMatch = True Then
                Stop
              Else
                lngFldID = ![fld_id]
              End If
            End With
            arr_varIdxFld(IF_FLDID, lngFE) = lngFldID

            intIdxOrd = intIdxOrd + 1
            With rst4
              .FindFirst "[idx_id] = " & CStr(arr_varIdxFld(IF_IDXID, lngFE)) & " And " & _
                "[fld_id] = " & CStr(arr_varIdxFld(IF_FLDID, lngFE))
              If .NoMatch = True Then
                .AddNew
                ![dbs_id] = arr_varDb(B_ID, lngDE)
                ![tbl_id] = arr_varIdx(I_TBLID, arr_varIdxFld(IF_IELEM, lngFE))
                ![idx_id] = arr_varIdxFld(IF_IDXID, lngFE)
                ![fld_id] = arr_varIdxFld(IF_FLDID, lngFE)
                ![idxfld_order] = intIdxOrd
                ![idxfld_name] = arr_varIdxFld(IF_NAME, lngFE)
                ![idxfld_attributes] = arr_varIdxFld(IF_ATTR, lngFE)
                ![idxfld_datemodified] = Now()
                .Update
                .Bookmark = .LastModified
                arr_varIdxFld(IF_ID, lngFE) = ![idxfld_id]
              Else
                If ![idxfld_order] <> intIdxOrd Then
                  .Edit
                  ![idxfld_order] = intIdxOrd
                  ![idxfld_datemodified] = Now()
                  .Update
                End If
                If ![idxfld_name] <> arr_varIdxFld(IF_NAME, lngFE) Then
                  .Edit
                  ![idxfld_name] = arr_varIdxFld(IF_NAME, lngFE)
                  ![idxfld_datemodified] = Now()
                  .Update
                End If
                If ![idxfld_attributes] <> arr_varIdxFld(IF_ATTR, lngFE) Then
                  .Edit
                  ![idxfld_attributes] = arr_varIdxFld(IF_ATTR, lngFE)
                  ![idxfld_datemodified] = Now()
                  .Update
                End If
              End If
            End With

          Next

          rst4.Close
          Set rst4 = Nothing

          rst3.Close
          Set rst3 = Nothing

        Next

        rst2.Close
        Set rst2 = Nothing

      Next

      rst1.Close
      Set rst1 = Nothing

    Next

    lngDels = 0&
    ReDim arr_varDel(D_ELEMS, 0)

    ' ** Check for indexes not longer present.
    Set rst1 = .OpenRecordset("tblIndex", dbOpenDynaset, dbReadOnly)
    With rst1
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngW = 1& To lngRecs
        blnFound = False
        For lngX = 0& To (lngDbs - 1&)
          If arr_varDb(B_ID, lngX) = ![dbs_id] And arr_varDb(B_TCNT, lngX) > 0& Then
            arr_varTdf = arr_varDb(B_TARR, lngX)
            lngTdfs = UBound(arr_varTdf, 2)
            For lngY = 0& To (lngTdfs - 1)
              If arr_varTdf(T_ID, lngY) = ![tbl_id] And arr_varTdf(T_ICNT, lngY) > 0& Then
                arr_varIdx = arr_varTdf(T_IARR, lngY)
                lngIdxs = (UBound(arr_varIdx, 2) + 1&)
                For lngZ = 0& To (lngIdxs - 1&)
                  If arr_varIdx(I_NAME, lngZ) = ![idx_name] Then
                    blnFound = True
                    Exit For
                  End If
                Next  ' ** lngIdxs: lngZ.
              End If  ' ** T_ID, T_ICNT.
              If blnFound = True Then
                Exit For
              End If
            Next  ' ** lngTdfs: lngY.
          End If  ' ** dbs_id, B_TCNT.
          If blnFound = True Then
            Exit For
          End If
        Next  ' ** lngDbs: lngX.
        If blnFound = False Then
          lngDels = lngDels + 1&
          lngDE = lngDels - 1&
          ReDim Preserve arr_varDel(D_ELEMS, lngDE)
          arr_varDel(0, lngDE) = ![idx_id]
        End If
        If lngW < lngRecs Then .MoveNext
      Next  ' ** lngRecs: lngW.
      .Close
    End With

    If lngDels > 0& Then
      For lngX = 0& To (lngDels - 1&)
        ' ** Delete tblIndex, by specified [idxid].
        Set qdf = .QueryDefs("zz_qry_Index_06")
        With qdf.Parameters
          ![idxid] = arr_varDel(0, lngX)
        End With
        qdf.Execute
      Next
      Debug.Print "'DELS: " & CStr(lngDels)
    End If

    .Close
  End With

  TmpTable_Chk False  ' ** Module Function: modUtilities.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set fld = Nothing
  Set idx = Nothing
  Set tdf = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  Set rst4 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set wrk = Nothing

  Idx_Doc = blnRetValx

End Function

Public Function TmpTable_Chk(blnLoad As Boolean) As Boolean
' ** NO LONGER DELETES TEMP TABLES; LEAVE THEM AROUND!

On Error GoTo ERRH

  Const THIS_PROC As String = "TmpTable_Chk"

  Dim dbs0 As DAO.Database, qdf0 As DAO.QueryDef
  Dim blnRetVal As Boolean

  blnRetVal = True

  If gstrTrustDataLocation = vbNullString Then
    IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
  End If

  Select Case blnLoad
  Case True

    ' ** Make sure all the referenced temporary tables are present
    ' ** while documenting, then delete them when finished.

    If TableExists("USysRibbons") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "USysRibbons", acTable, "zz_USysRibbons"
    End If
    If TableExists("tblDatabase_Table_Link_tmp01") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tblDatabase_Table_Link_tmp01", acTable, "zz_tbl_Database_Table_Link"
    End If
    If TableExists("tblDatabase_Table_Link_tmp02") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tblDatabase_Table_Link_tmp02", acTable, "zz_tbl_Database_Table_Link"
    End If
    If TableExists("tblDatabase_Table_Link_tmp03") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tblDatabase_Table_Link_tmp03", acTable, "zz_tbl_Database_Table_Link"
    End If

    CurrentDb.TableDefs.Refresh
    CurrentDb.TableDefs.Refresh

  Case False

    TableDelete "tblDatabase_Table_Link_tmp01"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp02"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp03"  ' ** Module Function: modFileUtilities.

    CurrentDb.TableDefs.Refresh
    CurrentDb.TableDefs.Refresh

  End Select

EXITP:
  Set qdf0 = Nothing
  Set dbs0 = Nothing
  TmpTable_Chk = blnRetVal
  Exit Function

ERRH:
  Select Case ERR.Number
  Case Else
    Beep
    MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
      "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()", _
      vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
  End Select
  Resume EXITP

End Function

Private Function Idx_ChkDocQrys(Optional varSkip As Variant) As Boolean

  Const THIS_PROC As String = "Idx_ChkDocQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngQrys As Long, arr_varQry As Variant
  Dim lngThisDbsID As Long, lngImpQrys As Long
  Dim blnSkip As Boolean
  Dim varTmp00 As Variant
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varQry().
  Const Q_DID  As Integer = 0
  Const Q_VID  As Integer = 1
  Const Q_QID  As Integer = 2
  Const Q_QDID As Integer = 3
  Const Q_VNAM As Integer = 4
  Const Q_QNAM As Integer = 5
  Const Q_TYP  As Integer = 6
  Const Q_DSC  As Integer = 7
  Const Q_SQL  As Integer = 8
  Const Q_FND  As Integer = 9
  Const Q_IMP  As Integer = 10

  blnRetValx = True

  Select Case IsMissing(varSkip)
  Case True
    blnSkip = True
  Case False
    blnSkip = varSkip
  End Select

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  If TableExists("zz_tbl_Database_Table_Link") = False Then  ' ** Module Function: modFileUtilities.
    Set dbs = CurrentDb
    With dbs
      ' ** Data-Definition: Create table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_01")
      qdf.Execute
      Set qdf = Nothing
      ' **    Data-Definition: Create index [contype_type] on table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_02")
      qdf.Execute
      Set qdf = Nothing
      ' **    Data-Definition: Create index [dbs_id_asof], [dbs_id], [tbl_id] Unique on table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_03")
      qdf.Execute
      Set qdf = Nothing
      ' **    Data-Definition: Create index [dbs_id_asof], [tbllnk_name] Unique on table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_04")
      qdf.Execute
      Set qdf = Nothing
      ' **    Data-Definition: Create index [dbs_id], [tbl_id], [tbllnk_id] PrimaryKey on table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_05")
      qdf.Execute
      Set qdf = Nothing
      ' **    Data-Definition: Create index [tbllnk_id] Unique on table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_06")
      qdf.Execute
      Set qdf = Nothing
      ' **    Data-Definition: Create index [tbllnk_name] on table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_07")
      qdf.Execute
      Set qdf = Nothing
      ' **    Data-Definition: Create index [tbllnk_sourcetablename] on table zz_tbl_Database_Table_Link.
      Set qdf = .QueryDefs("zz_qry_System_59_08")
      qdf.Execute
      Set qdf = Nothing
      Beep
      .Close
    End With
    Set dbs = Nothing
  End If

  'blnSkip = True
  If blnSkip = False Then

    strPath = gstrDir_Dev
    strFile = CurrentAppName  ' ** Module Function: modFileUtilities.
    strPathFile = strPath & LNK_SEP & strFile

    Set dbs = CurrentDb
    With dbs
      ' ** zz_tbl_Query_Documentation, by specified [vbnam].
      Set qdf = .QueryDefs("qryQuery_Documentation_01")
      With qdf.Parameters
        ![vbnam] = THIS_NAME
      End With
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngQrys = .RecordCount
        .MoveFirst
        arr_varQry = .GetRows(lngQrys)
        ' ****************************************************
        ' ** Array: arr_varQry()
        ' **
        ' **   Field  Element  Name               Constant
        ' **   =====  =======  =================  ==========
        ' **     1       0     dbs_id             Q_DID
        ' **     2       1     vbcom_id           Q_VID
        ' **     3       2     qry_id             Q_QID
        ' **     4       3     qrydoct_id         Q_QDID
        ' **     5       4     vbcom_name         Q_VNAM
        ' **     6       5     qry_name           Q_QNAM
        ' **     7       6     qrytype_type       Q_TYP
        ' **     8       7     qry_description    Q_DSC
        ' **     9       8     qry_sql            Q_SQL
        ' **    10       9     qry_found          Q_FND
        ' **    11      10     qry_import         Q_IMP
        ' **
        ' ****************************************************
        .Close
      End With
      Set rst = Nothing
      Set qdf = Nothing
      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'IDX DOC QRYS: " & CStr(lngQrys)
    DoEvents

  End If  ' ** blnSkip.

  'blnSkip = True
  'If blnSkip = False Then
  '  Set dbs = CurrentDb
  '  With dbs
  '    varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", "[vbcom_name] = '" & THIS_NAME & "'")
  '    If IsNull(varTmp00) = True Then
  '      Stop
  '    End If
  '    Set rst = .OpenRecordset("zz_tbl_Query_Documentation", dbOpenDynaset, dbAppendOnly)
  '    For lngX = 0& To (lngQrys - 1&)
  '      'Set qdf = .QueryDefs(arr_varQry(Q_QNAM, lngX))
  '      With rst
  '        .AddNew
  '        ' ** ![qrydoct_id] : AutoNumber.
  '        ![dbs_id] = lngThisDbsID
  '        ![vbcom_id] = varTmp00
  '        '![qry_id] =
  '        ![vbcom_name] = THIS_NAME
  '        ![qry_name] = arr_varQry(Q_QNAM, lngX)
  '        '![qrytype_type] = qdf.Type
  '        '![qry_description] = qdf.Properties("Description")
  '        '![qry_sql] = qdf.SQL
  '        ![qrydoct_datemodified] = Now()
  '        .Update
  '      End With
  '    Next
  '    rst.Close
  '    Set rst = Nothing
  '    .Close
  '  End With
  '  Set dbs = Nothing
  'End If

  'blnSkip = True
  If blnSkip = False Then

    For lngX = 0& To (lngQrys - 1&)
      If QueryExists(CStr(arr_varQry(Q_QNAM, lngX))) = True Then  ' ** Module Function: modFileUtilities.
        arr_varQry(Q_FND, lngX) = CBool(True)
      End If
    Next

    lngImpQrys = 0&
    For lngX = 0& To (lngQrys - 1&)
      If arr_varQry(Q_FND, lngX) = False Then
        lngImpQrys = lngImpQrys + 1&
      End If
    Next

    If lngImpQrys > 0& Then
      Debug.Print "'QRYS TO IMPORT: " & CStr(lngImpQrys)
      DoEvents
      For lngX = 0& To (lngQrys - 1&)
        If arr_varQry(Q_FND, lngX) = False Then
On Error Resume Next
          DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
          If ERR.Number <> 0 Then
On Error GoTo 0
            Debug.Print "'QRY MISSING!  " & arr_varQry(Q_QNAM, lngX)
          Else
On Error GoTo 0
          End If
          arr_varQry(Q_IMP, lngX) = CBool(True)
        End If
      Next
    Else
      Debug.Print "'ALL IDX DOC QRYS PRESENT!"
    End If

    Debug.Print "'DONE!"
    DoEvents

    Beep

  End If  ' ** blnSkip.

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Idx_ChkDocQrys = blnRetValx

End Function
