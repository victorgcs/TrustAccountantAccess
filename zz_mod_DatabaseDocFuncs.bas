Attribute VB_Name = "zz_mod_DatabaseDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_DatabaseDocFuncs"

'VGC 04/11/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const IsDemo = 0  ' ** 0 = new/upgrade; -1 = demo.
' ** Also in:
' **   modGlobConst
' **   modSecurityFunctions
' **   zz_mod_MDEPrepFuncs

' ** AcControlType enumeration:  (my own)
Public Const acNone              As Long = 99&
Public Const acDatasheetColumn   As Long = 115&
Public Const acEmptyCell         As Long = 127&
Public Const acWebBrowser        As Long = 128&
Public Const acNavigationControl As Long = 129&
Public Const acNavigationButton  As Long = 130&

Private blnRetValx As Boolean, blnWithJrnlTmp As Boolean
' **

Public Function QuikTblDoc() As Boolean
  Const THIS_PROC As String = "QuikTblDoc"
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Tbl_ChkDocQrys(False) = True Then  ' ** Function: Below.
      blnWithJrnlTmp = False
      blnRetValx = Tbl_Doc  ' ** Function: Below.
      blnRetValx = Tbl_Link_Chk  ' ** Function: Below.
      blnRetValx = Tbl_Fld_Doc  ' ** Function: Below.
      blnRetValx = Tbl_Link_Chk  ' ** Function: Below.
      blnRetValx = Tbl_Fld_DateFormat_Doc  ' ** Function: Below.
      blnRetValx = Tbl_Fld_RowSource_Doc  ' ** Function: Below.
      blnRetValx = Tbl_Link_Chk  ' ** Function: Below.
      blnRetValx = Tbl_RecCnt_Doc  ' ** Function: Below.
      blnRetValx = Tbl_Link_Chk  ' ** Function: Below.
      blnRetValx = Tbl_AutoNum_Doc  ' ** Function: Below.
      blnRetValx = Tbl_Link_Chk  ' ** Function: Below.
      blnRetValx = Tbl_Link_DocConn
      blnRetValx = Tbl_Link_Chk  ' ** Function: Below.
      blnRetValx = Tbl_DescCnt
      'blnRetValx = Tbl_Link_Doc
      'blnRetValx = References_Doc  ' ** Module Function: modXAdminFuncs.
      ' ** TO UNDO ABORTED DOC, RUN THIS IN THE IMMEDIATE WINDOW:
      ' **   Qry_TmpTables(False, True)  ' ** Function: Below.
      DoEvents
      DoBeeps  ' ** Module Function: modWindowFunctions.
      Debug.Print "'FINISHED!"
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED Tbl_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikTblDoc = blnRetValx
End Function

Public Function Tbl_Doc() As Boolean
' ** Document all tables to tblDatabase_Table.

  Const THIS_PROC As String = "Tbl_Doc"

  Dim wrk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim tdf As DAO.TableDef, prp As DAO.Property, fld As DAO.Field
  Dim strThisFile As String, strThisPath As String, lngThisDbsID As Long
  Dim strThatFile As String, strThatPath As String, lngThatDbsID As Long
  Dim lngDats As Long, arr_varDat As Variant
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngMults As Long, arr_varMult() As Variant
  Dim blnDocTbl As Boolean, blnFound As Boolean, blnDelete As Boolean
  Dim lngTblID As Long
  Dim lngRecs As Long
  Dim intPos1 As Integer
  Dim strTmp00 As String, lngTmp01 As Long, lngTmp02 As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varDat().
  Const D_DID  As Integer = 0
  Const D_DNAM As Integer = 1
  Const D_PATH As Integer = 2
  Const D_TBLS As Integer = 3
  Const D_RELS As Integer = 4
  Const D_IDXS As Integer = 5
  Const D_DAT  As Integer = 6

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 12  ' ** Array's first-element UBound().
  Const T_DID  As Integer = 0
  Const T_DNAM As Integer = 1
  Const T_TID  As Integer = 2
  Const T_TNAM As Integer = 3
  Const T_SRC  As Integer = 4
  Const T_DESC As Integer = 5
  Const T_HID  As Integer = 6
  Const T_FLDS As Integer = 7
  Const T_RELS As Integer = 8
  Const T_IDXS As Integer = 9
  Const T_ORD  As Integer = 10
  Const T_DAT  As Integer = 11
  Const T_FND  As Integer = 12

  ' ** Array: arr_varMult().
  Const M_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const M_TNAM  As Integer = 0
  Const M_CNT   As Integer = 1
  Const M_TID1  As Integer = 2
  Const M_TID2  As Integer = 3
  Const M_DID1  As Integer = 4
  Const M_DNAM1 As Integer = 5
  Const M_DID2  As Integer = 6
  Const M_DNAM2 As Integer = 7

  blnRetValx = True
  DoCmd.Hourglass True
  DoEvents

  strThisFile = Parse_File(CurrentDb.Name)  ' ** Module Function: modFileUtilities.
  strThisPath = Parse_Path(CurrentDb.Name) & LNK_SEP  ' ** Module Function: modFileUtilities.

  Qry_TmpTables True, blnWithJrnlTmp  ' ** Function: Below.

  DoCmd.Hourglass True
  DoEvents

  Set dbsLoc = CurrentDb
  With dbsLoc

' ** This needs to exclude TrustImport and TrstXAdm.
'PARAMETERS [wjrnl] Bit;
'SELECT tblDatabase.dbs_id, tblDatabase.dbs_name, tblDatabase.dbs_path, tblDatabase.dbs_tbl_cnt, tblDatabase.dbs_rel_cnt, tblDatabase.dbs_idx_cnt, tblDatabase.dbs_datemodified
'FROM tblDatabase
'WHERE (((IIf([dbs_name]<>'TAJrnTmp.mdb',-1,IIf([wjrnl]=False,0,-1)))=-1) AND ((IIf([dbs_name] In ('TrustImport.mdb','TrustImport.mde','TrstXAdm.mdb','TrstXAdm.mde'),0,-1))=-1))
'ORDER BY tblDatabase.dbs_id;

    ' ** tblDatabase, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngDats = .RecordCount
      .MoveFirst
      arr_varDat = .GetRows(lngDats)
      ' ******************************************************
      ' ** Array: arr_varDat()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ===========
      ' **     1       0     dbs_id              D_DID
      ' **     2       1     dbs_name            D_DNAM
      ' **     3       2     dbs_path            D_PATH
      ' **     4       3     dbs_tbl_cnt         D_TBLS
      ' **     5       4     dbs_rel_cnt         D_RELS
      ' **     6       5     dbs_idx_cnt         D_IDXS
      ' **     7       6     dbs_datemodified    D_DAT
      ' **
      ' ******************************************************
      .Close
    End With

    lngThisDbsID = 0&
    For lngX = 0& To (lngDats - 1&)
      If arr_varDat(D_DNAM, lngX) = strThisFile Then
        lngThisDbsID = arr_varDat(D_DID, lngX)
        Exit For
      End If
    Next

    ' ** Make sure paths are correct.
    For lngX = 0& To (lngDats - 1&)
      If Left(arr_varDat(D_DNAM, lngX), 11) = "TrustImport" Then
        If CurrentAppPath = gstrDir_Dev And Right(arr_varDat(D_PATH, lngX), 11) <> "TrustImport" Then  ' ** Module Functions: modFileUtilities.
          strTmp00 = Parse_Path(CurrentAppPath)  ' ** Module Functions: modFileUtilities.
          strTmp00 = Parse_Path(strTmp00)  ' ** Module Functions: modFileUtilities.
          strTmp00 = strTmp00 & LNK_SEP & "TrustImport"
          ' ** Update tblDatabase, by specified [dbid], [dbpat].
          Set qdf = .QueryDefs("zz_qry_Database_02")
          With qdf.Parameters
            ![dbid] = arr_varDat(D_DID, lngX)
            ![dbpat] = strTmp00
          End With
          qdf.Execute dbFailOnError
          arr_varDat(D_PATH, lngX) = strTmp00
        ElseIf CurrentAppPath = gstrDir_Def And Right(arr_varDat(D_PATH, lngX), 12) <> "Trust Import" Then  ' ** Module Functions: modFileUtilities.
          Stop
        ElseIf InStr(arr_varDat(D_PATH, lngX), "TrustAccountant") > 0 Then
          strTmp00 = Parse_Path(CurrentAppPath)  ' ** Module Functions: modFileUtilities.
          strTmp00 = Parse_Path(strTmp00)  ' ** Module Functions: modFileUtilities.
          strTmp00 = strTmp00 & LNK_SEP & "TrustImport"
          ' ** Update tblDatabase, by specified [dbid], [dbpat].
          Set qdf = .QueryDefs("zz_qry_Database_02")
          With qdf.Parameters
            ![dbid] = arr_varDat(D_DID, lngX)
            ![dbpat] = strTmp00
          End With
          qdf.Execute dbFailOnError
          arr_varDat(D_PATH, lngX) = strTmp00
        End If
      ElseIf Left(arr_varDat(D_DNAM, lngX), 8) = "TrstXAdm" Then
        If arr_varDat(D_PATH, lngX) <> CurrentAppPath Then  ' ** Module Function: modFileUtilities.
          strTmp00 = CurrentAppPath  ' ** Module Function: modFileUtilities.
          ' ** Update tblDatabase, by specified [dbid], [dbpat].
          Set qdf = .QueryDefs("zz_qry_Database_02")
          With qdf.Parameters
            ![dbid] = arr_varDat(D_DID, lngX)
            ![dbpat] = strTmp00
          End With
          qdf.Execute dbFailOnError
          arr_varDat(D_PATH, lngX) = strTmp00
        End If
      End If
    Next

    For lngY = 0& To (lngDats - 1&)
      If InStr(arr_varDat(D_DNAM, lngY), "TrustImport") > 0 Then
        If InStr(arr_varDat(D_PATH, lngY), "TrustAccountant") > 0 Then
          Stop
        End If
      ElseIf InStr(arr_varDat(D_DNAM, lngY), "TrustImport") = 0 Then
        If InStr(arr_varDat(D_PATH, lngY), "TrustImport") > 0 Then
          Stop
        End If
      End If
    Next

    lngMults = 0&
    ReDim arr_varMult(M_ELEMS, 0)

' ** This needs to exclude TrustImport and TrstXAdm.
'PARAMETERS [wjrnl] Bit;
'SELECT tblDatabase_Table.tbl_name, Count(tblDatabase_Table.tbl_id) AS cnt, Last(tblDatabase_Table.tbl_id) AS tbl_id1, First(tblDatabase_Table.tbl_id) AS tbl_id2, First(tblDatabase.dbs_id) AS dbs_id1, First(tblDatabase.dbs_name) AS dbs_name1, Last(tblDatabase.dbs_id) AS dbs_id2, Last(tblDatabase.dbs_name) AS dbs_name2
'FROM tblDatabase INNER JOIN tblDatabase_Table ON tblDatabase.dbs_id = tblDatabase_Table.dbs_id
'WHERE (((IIf([dbs_name]<>'TAJrnTmp.mdb',-1,IIf([wjrnl]=False,0,-1)))=-1) AND ((IIf(Left([tbl_name],4)='MSys',0,-1))=-1) AND ((IIf([dbs_name] In ('TrustImport.mdb','TrustImport.mde','TrstXAdm.mdb','TrstXAdm.mde'),0,-1))=-1))
'GROUP BY tblDatabase_Table.tbl_name
'HAVING (((Count(tblDatabase_Table.tbl_id))>1));

    ' ** tblDatabase_Table, grouped by tbl_name, with cnt > 1; tables
    ' ** appearing in more than 1 of TA's databases, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_02")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        ' ** Nothing documented yet.
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          lngMults = lngMults + 1&
          lngE = lngMults - 1&
          ReDim Preserve arr_varMult(M_ELEMS, lngE)
          ' **********************************************
          ' ** Array: arr_varMult()
          ' **
          ' **   Field  Element  Name         Constant
          ' **   =====  =======  ===========  ==========
          ' **     1       0     tbl_name     M_TNAM
          ' **     2       1     cnt          M_CNT
          ' **     3       2     tbl_id1      M_TID1
          ' **     4       3     tbl_id2      M_TID2
          ' **     5       4     dbs_id1      M_DID1
          ' **     6       5     dbs_name1    M_DNAM1
          ' **     7       6     dbs_id2      M_DID2
          ' **     8       7     dbs_name2    M_DNAM2
          ' **
          ' **********************************************
          For lngY = 0& To (.Fields.Count - 1)
            arr_varMult(lngY, lngE) = .Fields(lngY)
          Next
          If lngX < lngRecs Then .MoveNext
        Next
      End If
      .Close
    End With

    For lngY = 0& To (lngDats - 1&)
      If InStr(arr_varDat(D_DNAM, lngY), "TrustImport") > 0 Then
        If InStr(arr_varDat(D_PATH, lngY), "TrustAccountant") > 0 Then
          Stop
        End If
      ElseIf InStr(arr_varDat(D_DNAM, lngY), "TrustImport") = 0 Then
        If InStr(arr_varDat(D_PATH, lngY), "TrustImport") > 0 Then
          Stop
        End If
      End If
    Next

' ** This needs to exclude TrustImport and TrstXAdm.
'PARAMETERS [wjrnl] Bit;
'SELECT tblDatabase_Table.dbs_id, tblDatabase.dbs_name, tblDatabase_Table.tbl_id, tblDatabase_Table.tbl_name, tblDatabase_Table.tbl_sourcetablename, tblDatabase_Table.tbl_description, tblDatabase_Table.sec_hidden, tblDatabase_Table.tbl_fld_cnt, tblDatabase_Table.tbl_rel_cnt, tblDatabase_Table.tbl_idx_cnt, tblDatabase_Table.tbl_order, tblDatabase_Table.tbl_datemodified, CBool(False) AS tbl_fnd
'FROM tblDatabase INNER JOIN tblDatabase_Table ON tblDatabase.dbs_id = tblDatabase_Table.dbs_id
'WHERE (((IIf([dbs_name]<>'TAJrnTmp.mdb',-1,IIf([wjrnl]=False,0,-1)))=-1) AND ((IIf(Left([tbl_name],4)='MSys',0,-1))=-1) AND ((IIf([dbs_name] In ('TrustImport.mdb','TrustImport.mde','TrstXAdm.mdb','TrstXAdm.mde'),0,-1))=-1))
'ORDER BY tblDatabase_Table.dbs_id, tblDatabase_Table.tbl_name;

    ' ** tblDatabase_Table, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        ' ** Nothing documented yet.
        lngTbls = 0&
        ReDim arr_varTbl(T_ELEMS, 0)
      Else
        .MoveLast
        lngTbls = .RecordCount
        ReDim arr_varTbl(T_ELEMS, (lngTbls - 1&))
        .MoveFirst
        For lngX = 1& To lngTbls
          lngE = lngX - 1&
          For lngY = 0& To (.Fields.Count - 1&)
            arr_varTbl(lngY, lngE) = .Fields(lngY)
            ' ********************************************************
            ' ** Array: arr_varTbl()
            ' **
            ' **   Field  Element  Name                   Constant
            ' **   =====  =======  =====================  ==========
            ' **     1       0     dbs_id                 T_DID
            ' **     2       1     dbs_name               T_DNAM
            ' **     3       2     tbl_id                 T_TID
            ' **     4       3     tbl_name               T_TNAM
            ' **     5       4     tbl_sourcetablename    T_SRC
            ' **     6       5     tbl_description        T_DESC
            ' **     7       6     sec_hidden             T_HID
            ' **     8       7     tbl_fld_cnt            T_FLDS
            ' **     9       8     tbl_rel_cnt            T_RELS
            ' **    10       9     tbl_idx_cnt            T_IDXS
            ' **    11      10     tbl_order              T_ORD
            ' **    12      11     tbl_datemodified       T_DAT
            ' **    13      12     tbl_fnd                T_FND
            ' **
            ' ********************************************************
          Next

          For lngY = 0& To (lngDats - 1&)
            If InStr(arr_varDat(D_DNAM, lngY), "TrustImport") > 0 Then
              If InStr(arr_varDat(D_PATH, lngY), "TrustAccountant") > 0 Then
                Stop
              End If
            ElseIf InStr(arr_varDat(D_DNAM, lngY), "TrustImport") = 0 Then
              If InStr(arr_varDat(D_PATH, lngY), "TrustImport") > 0 Then
                Stop
              End If
            End If
          Next

          ' ** Cross-check the database that the table is from.
          ' ** (Yes, I'm trusting that the doubles are in correctly. It's all the others I want to check.)
          For Each tdf In dbsLoc.TableDefs
            If tdf.Name = ![tbl_name] Then
              blnFound = False
              For lngY = 0& To (lngMults - 1&)
                If arr_varMult(M_TNAM, lngY) = tdf.Name Then
                  blnFound = True
                  Exit For
                End If
              Next
              If blnFound = False Then
                If tdf.Connect = vbNullString Then
                  ' ** Should be listed with this database.
                  If ![dbs_id] <> lngThisDbsID Then
                    lngTmp01 = ![dbs_id]
                    lngTmp02 = ![tbl_id]
                    .Edit
                    ![dbs_id] = lngThisDbsID
On Error Resume Next
                    .Update
                    If ERR.Number <> 0 Then
                      Select Case ERR.Number
                      Case 3397  ' ** Cannot perform cascading operation.  There must be a related record in table 'tblIndex'.
                        strTmp00 = ERR.description
On Error GoTo 0
                        intPos1 = InStr(strTmp00, "'")
                        If intPos1 > 0 Then
                          strTmp00 = Mid$(strTmp00, (intPos1 + 1))
                          intPos1 = InStr(strTmp00, "'")
                          If intPos1 > 0 Then
                            strTmp00 = Left$(strTmp00, (intPos1 - 1))
                            Select Case strTmp00
                            Case "tblIndex", "tblRelation"
                              ' ** Delete tblIndex, by specified [dbsid], [tblid].
                              Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_30")
                              With qdf.Parameters
                                ![dbsid] = lngTmp01
                                ![tblid] = lngTmp02
                              End With
                              qdf.Execute
                              ' ** Delete tblRelation, by specified [dbsid], [tblid].
                              Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_31")
                              With qdf.Parameters
                                ![dbsid] = lngTmp01
                                ![tblid] = lngTmp02
                              End With
                              qdf.Execute
                              .Update
                            Case Else
                              DoCmd.Hourglass False
                              Debug.Print "'ERROR: " & CStr(ERR.Number) & "  " & ERR.description
                              Debug.Print "'" & tdf.Name
                              Stop
                            End Select
                          Else
                            DoCmd.Hourglass False
                            Stop
                          End If
                        Else
                          DoCmd.Hourglass False
                          Stop
                        End If
                      Case Else
                        DoCmd.Hourglass False
                        Debug.Print "'ERROR: " & CStr(ERR.Number) & "  " & ERR.description
On Error GoTo 0
                        Stop
                      End Select
                    Else
On Error GoTo 0
                    End If
                    arr_varTbl(T_DID, lngX) = lngThisDbsID
                  End If
                Else
                  ' ** Should be listed with the connected database.
                  strTmp00 = tdf.Connect
                  intPos1 = InStr(strTmp00, LNK_IDENT)
                  If intPos1 > 0 Then
                    ' ** I know. It's not likely a spreadsheet would come over with the same same as a table.
                    strThatFile = Mid$(strTmp00, (intPos1 + Len(LNK_IDENT)))
                    strThatPath = Parse_Path(strThatFile)  ' ** Module Function: modFileUtilities.
                    strThatFile = Parse_File(strThatFile)  ' ** Module Function: modFileUtilities.
                    For lngY = 0& To (lngDats - 1&)
                      If arr_varDat(D_DNAM, lngY) = strThatFile Then
                        lngThatDbsID = arr_varDat(D_DID, lngY)
                        Exit For
                      End If
                    Next
                    If ![dbs_id] <> lngThatDbsID Then
                      ' ** When I've moved a local table to one of the backends, or
                      ' ** switched backends, simply changing the dbs_id errors, because
                      ' ** all of its fields are still listed under its former database.
                      ' ** So, let's delete its fields first.
                      lngTblID = ![tbl_id]
                      ' ** Delete tblDatabase_Table_Field, by specified [tblid].
                      Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_29")
                      With qdf.Parameters
                        ![tblid] = lngTblID
                      End With
                      qdf.Execute
                      .Edit
                      ![dbs_id] = lngThatDbsID
                      .Update
                      arr_varTbl(T_DID, lngX) = lngThatDbsID
                    End If
                  End If
                End If
              End If
              Exit For
            End If
          Next
          If lngX < lngTbls Then .MoveNext
        Next
      End If
      .Close
    End With

    ' ** Set Null text fields to vbNullString to simplify comparisons.
    For lngX = 0& To (lngTbls - 1&)
      If IsNull(arr_varTbl(T_DESC, lngX)) = True Then
        arr_varTbl(T_DESC, lngX) = vbNullString
      End If
    Next

    .Close
  End With

  For lngX = 0& To (lngDats - 1&)
    If arr_varDat(D_DNAM, lngX) = strThisFile Then
      Set wrk = DBEngine.Workspaces(0)
    Else
On Error Resume Next
      Set wrk = CreateWorkspace("tmp", "Admin", "", dbUseJet)
      If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
        Set wrk = CreateWorkspace("tmp", "superuser", TA_SEC, dbUseJet)
        If ERR.Number <> 0 Then
On Error GoTo 0
          Set wrk = CreateWorkspace("tmp", "superuser", TA_SEC6, dbUseJet)
        Else
On Error GoTo 0
        End If
      Else
On Error GoTo 0
      End If
    End If
    With wrk
      If arr_varDat(D_DNAM, lngX) = strThisFile Then
        Set dbsLnk = .Databases(0)
      Else
        For lngY = 0& To (lngDats - 1&)
          If InStr(arr_varDat(D_DNAM, lngY), "TrustImport") > 0 Then
            If InStr(arr_varDat(D_PATH, lngY), "TrustAccountant") > 0 Then
              Stop
            End If
          ElseIf InStr(arr_varDat(D_DNAM, lngY), "TrustImport") = 0 Then
            If InStr(arr_varDat(D_PATH, lngY), "TrustImport") > 0 Then
              Stop
            End If
          End If
        Next
        Set dbsLnk = .OpenDatabase(arr_varDat(D_PATH, lngX) & LNK_SEP & arr_varDat(D_DNAM, lngX), False, False)  ' ** {pathfile}, {exclusive}, {read-only}
      End If
      With dbsLnk
        dbsLnk.TableDefs.Refresh
        dbsLnk.TableDefs.Refresh
        Set dbsLoc = CurrentDb
        ' ** tblDatabase_Table, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
        Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_01")
        With qdf.Parameters
          ![wjrnl] = blnWithJrnlTmp
        End With
        Set rst = qdf.OpenRecordset

        For Each tdf In .TableDefs
          With tdf
            blnDocTbl = True
            If .Connect <> vbNullString Then
              ' ** Only document tables local to the remote database.
              blnDocTbl = False
            End If
            If Left$(.Name, 4) = "MSys" Then
              ' ** Don't document Microsoft System tables, but do document User System tables: USys...
              blnDocTbl = False
            End If
            If blnDocTbl = True Then
              For lngY = 0& To (lngTbls - 1&)
                If arr_varTbl(T_DID, lngY) = arr_varDat(D_DID, lngX) And arr_varTbl(T_TNAM, lngY) = .Name Then
                  blnDocTbl = False
                  arr_varTbl(T_FND, lngY) = True
                  If arr_varTbl(T_FLDS, lngY) <> .Fields.Count Then
                    ' ** Update tblDatabase_Table, for tbl_fld_cnt, by specified [dbsid], [tblid].
                    Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_05")
                    With qdf.Parameters
                      ![dbsid] = arr_varTbl(T_DID, lngY)
                      ![tblid] = arr_varTbl(T_TID, lngY)
                      ![FldCnt] = tdf.Fields.Count
                    End With
                    qdf.Execute dbFailOnError
                  End If
                  blnFound = False
                  For Each prp In tdf.Properties
                    With prp
                      If .Name = "Description" Then
                        blnFound = True
                        strTmp00 = .Value
                        If Trim$(strTmp00) <> vbNullString Then
                          If arr_varTbl(T_DESC, lngY) <> strTmp00 Then
                            arr_varTbl(T_DESC, lngY) = strTmp00
                            ' ** Update tbl_description in tblDatabase_Table, by specified [dbsid], [tblid], [tbldesc].
                            Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_24")
                            With qdf.Parameters
                              ![dbsid] = arr_varTbl(T_DID, lngY)
                              ![tblid] = arr_varTbl(T_TID, lngY)
                              ![tbldesc] = strTmp00
                            End With
                            qdf.Execute dbFailOnError
                          End If
                          Exit For
                        Else
                          ' ** Has Description property, but nothing's in it.
                          If arr_varTbl(T_DESC, lngY) <> vbNullString Then
                            .Value = arr_varTbl(T_DESC, lngY)
                          End If
                        End If
                      End If
                    End With  ' ** prp.
                  Next
                  If blnFound = False Then
                    ' ** Doesn't have Description property.
                    If arr_varTbl(T_DESC, lngY) <> vbNullString Then
                      Tbl_Property_Add tdf, "Description", dbText, arr_varTbl(T_DESC, lngY) ' ** Function: Below.
                    End If
                  End If
                End If
              Next  ' ** For each TableDef: tdf.
            End If
            If blnDocTbl = True Then
              If Left$(.Name, 4) <> "~TMP" Then
For lngY = 0& To (lngTbls - 1&)

Next
                With rst
                  .AddNew
                  ![dbs_id] = arr_varDat(D_DID, lngX)
                  ![tbl_name] = tdf.Name
                  If tdf.SourceTableName <> vbNullString Then
                    ![tbl_sourcetablename] = tdf.SourceTableName
                  End If
                  For Each prp In tdf.Properties
                    With prp
                      If .Name = "Description" Then
                        If Trim$(.Value) <> vbNullString Then
                          rst![tbl_description] = CStr(.Value)
                          Exit For
                        End If
                      End If
                    End With
                  Next
                  ![tbl_fld_cnt] = tdf.Fields.Count
                  ![tbl_rel_cnt] = 0&
                  ![tbl_idx_cnt] = 0&
                  ![tbl_datemodified] = Now()
On Error Resume Next
                  .Update
                  If ERR.Number = 0 Then
On Error GoTo 0
                    Debug.Print "'TBL ADD: '" & tdf.Name & "'  IN  '" & arr_varDat(D_DNAM, lngX) & "'"
                  Else
                    Debug.Print "'CAN'T ADD! " & tdf.Name & "  " & CStr(ERR.Number) & "  " & ERR.description
Stop
'CAN'T ADD! LedgerArchive_Backup  3022  The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.
'CAN'T ADD! m_TBL_tmp01  3022  The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.
'CAN'T ADD! tblDatabase_Table_Link_tmp01  3022  The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.
'CAN'T ADD! tblDatabase_Table_Link_tmp02  3022  The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.
'CAN'T ADD! tblDatabase_Table_Link_tmp03  3022  The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.  Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again.
On Error GoTo 0
                  End If
                End With
              End If  ' ** ~TMP.
            End If
          End With
        Next
        rst.Close
        dbsLoc.Close
        .Close
      End With  ' ** This Database: dbsLnk.
      .Close
    End With  ' ** wrk.
  Next  ' ** For each Database: dbsLnk.

  Set dbsLoc = CurrentDb
  With dbsLoc

    ' ** Update zz_qry_Database_Table_06b (tblDatabase, with DLookups() to zz_qry_Database_Table_06a
    ' ** (tblDatabase_Table, grouped by dbs_id, with table cnt)).
    Set qdf = .QueryDefs("zz_qry_Database_Table_06c")
    qdf.Execute dbFailOnError

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

    ' ** Look for deleted tables.
    For lngX = 0& To (lngTbls - 1&)
      If arr_varTbl(T_FND, lngX) = False Then
        If Left(arr_varTbl(T_TNAM, lngX), 4) <> "MSys" And Left(arr_varTbl(T_TNAM, lngX), 4) <> "USys" And _
            arr_varTbl(T_TNAM, lngX) <> "xlExcelImport" And arr_varTbl(T_DNAM, lngX) <> "TrstXadm.mdb" Then
Select Case arr_varTbl(T_TNAM, lngX)
Case "m_TBL1", "m_TBL2", "zz_tbl_Database_Table_Link", "zz_tbl_Form_Doc", "zz_tbl_Form_Property", "zz_tbl_Form_Property_Value", _
    "zz_tbl_sql_code_01", "zz_tbl_sql_code_02", "zz_tbl_sql_code_03", "zz_tbl_sql_code_04", "zz_tbl_sql_code_05", _
    "zz_tbl_sql_code_06", "zz_tbl_VBComponent_KeyDown"
  ' ** Leave these in.
Case Else
          strTmp00 = arr_varTbl(T_DNAM, lngX)
          blnDelete = True
          DoCmd.Hourglass False
          Debug.Print "'DEL TBL? '" & arr_varTbl(T_TNAM, lngX) & "' IN '" & strTmp00 & "'"
Stop
          If blnDelete = True Then
            DoCmd.Hourglass True
            ' ** Delete tblDatabase_Table, by specified [tblid].
            Set qdf = .QueryDefs("zz_qry_Database_Table_28")
            With qdf.Parameters
              ![tblid] = arr_varTbl(T_TID, lngX)
            End With
            qdf.Execute dbFailOnError
          Else
            DoCmd.Hourglass True
          End If
End Select
        End If
      End If
    Next

    .Close
  End With

  DoEvents
  Qry_TmpTables False, blnWithJrnlTmp  ' ** Function: Below.

  DoCmd.Hourglass False
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set prp = Nothing
  Set fld = Nothing
  Set tdf = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set wrk = Nothing

  Tbl_Doc = blnRetValx

End Function

Public Function Tbl_Fld_Doc() As Boolean
' ** Document all fields to tblDatabase_Table_Field.

  Const THIS_PROC As String = "Tbl_Fld_Doc"

  Dim wrk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim tdf As DAO.TableDef, fld As DAO.Field, prp As DAO.Property
  Dim strThisFile As String, strThisPath As String
  Dim lngDats As Long, arr_varDat As Variant
  Dim lngTbls As Long, arr_varTbl As Variant
  Dim lngFlds As Long, arr_varFld() As Variant
  Dim blnDocFld As Boolean, strFld As String
  Dim blnSQL As Boolean, strSQL As String
  Dim blnDelete As Boolean
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  ' ** Array: arr_varDat().
  Const D_DID  As Integer = 0
  Const D_DNAM As Integer = 1
  Const D_PATH As Integer = 2
  Const D_TBLS As Integer = 3
  Const D_RELS As Integer = 4
  Const D_IDXS As Integer = 5
  Const D_DAT  As Integer = 6

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 11  ' ** Array's first-element UBound().
  Const T_DID  As Integer = 0
  Const T_DNAM As Integer = 1
  Const T_TID  As Integer = 2
  Const T_TNAM As Integer = 3
  Const T_SRC  As Integer = 4
  Const T_DESC As Integer = 5
  Const T_FLDS As Integer = 6
  Const T_RELS As Integer = 7
  Const T_IDXS As Integer = 8
  Const T_ORD  As Integer = 9
  Const T_DAT  As Integer = 10
  Const T_FND  As Integer = 11

  ' ** Array: arr_varFld().
  Const F_ELEMS As Integer = 17  ' ** Array's first-element UBound()
  Const F_DID  As Integer = 0
  Const F_DNAM As Integer = 1
  Const F_TID  As Integer = 2
  Const F_FID  As Integer = 3
  Const F_FNAM As Integer = 4
  Const F_DESC As Integer = 5
  Const F_REQ  As Integer = 6
  Const F_TYPE As Integer = 7
  Const F_SIZE As Integer = 8
  Const F_CTL  As Integer = 9
  Const F_DEF  As Integer = 10
  Const F_FRMT As Integer = 11
  Const F_DEC  As Integer = 12
  Const F_INP  As Integer = 13
  Const F_ZERO As Integer = 14
  Const F_ATTR As Integer = 15
  Const F_DATE As Integer = 16
  Const F_FND  As Integer = 17

  blnRetValx = True
  DoCmd.Hourglass True
  DoEvents

  strThisFile = Parse_File(CurrentDb.Name)  ' ** Module Function: modFileUtilities.
  strThisPath = Parse_Path(CurrentDb.Name) & LNK_SEP  ' ** Module Function: modFileUtilities.

  Qry_TmpTables True, blnWithJrnlTmp  ' ** Function: Below.

  DoCmd.Hourglass True
  DoEvents

  Set dbsLoc = CurrentDb
  With dbsLoc

    ' ** tblDatabase, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngDats = .RecordCount
      .MoveFirst
      arr_varDat = .GetRows(lngDats)
      ' *****************************************************
      ' ** Array: arr_varDat()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     dbs_id              D_DID
      ' **     2       1     dbs_name            D_DNAM
      ' **     3       2     dbs_path            D_PATH
      ' **     4       3     dbs_tbl_cnt         D_TBLS
      ' **     5       4     dbs_rel_cnt         D_RELS
      ' **     6       5     dbs_idx_cnt         D_IDXS
      ' **     7       6     dbs_datemodified    D_DAT
      ' **
      ' *****************************************************
      .Close
    End With

    ' ** tblDatabase_Table, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngTbls = .RecordCount
      .MoveFirst
      arr_varTbl = .GetRows(lngTbls)
      ' ********************************************************
      ' ** Array: arr_varTbl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     dbs_id                 T_DID
      ' **     2       1     dbs_name               T_DNAM
      ' **     3       2     tbl_id                 T_TID
      ' **     4       3     tbl_name               T_TNAM
      ' **     5       4     tbl_sourcetablename    T_SRC
      ' **     6       5     tbl_description        T_DESC
      ' **     7       6     tbl_fld_cnt            T_FLDS
      ' **     8       7     tbl_rel_cnt            T_RELS
      ' **     9       8     tbl_idx_cnt            T_IDXS
      ' **    10       9     tbl_order              T_ORD
      ' **    11      10     tbl_datemodified       T_DAT
      ' **    12      11     Found                  T_FND
      ' **
      ' ********************************************************
      .Close
    End With

    ' ** tblDatabase_Table_Field, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_Field_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        ' ** Nothing documented yet.
      Else
        .MoveLast
        lngFlds = .RecordCount
        ReDim arr_varFld(F_ELEMS, (lngFlds - 1&))
        .MoveFirst
        For lngX = 1& To lngFlds
          lngE = lngX - 1&
          For lngY = 0& To (.Fields.Count - 1&)
            arr_varFld(lngY, lngE) = .Fields(lngY)
            ' ********************************************************
            ' ** Array: arr_varFld()
            ' **
            ' **   Field  Element  Name                   Constant
            ' **   =====  =======  =====================  ==========
            ' **     1       0     dbs_id                 F_DID
            ' **     2       1     dbs_name               F_DNAM
            ' **     3       2     tbl_id                 F_TID
            ' **     4       3     fld_id                 F_FID
            ' **     5       4     fld_name               F_FNAM
            ' **     6       5     fld_description        F_DESC
            ' **     7       6     fld_required           F_REQ
            ' **     8       7     datatype_db_type       F_TYPE
            ' **     9       8     fld_size               F_SIZE
            ' **    10       9     dispctl_type           F_CTL
            ' **    11      10     fld_defaultvalue       F_DEF
            ' **    12      11     fld_format             F_FRMT
            ' **    13      12     fld_decimalplaces      F_DEC
            ' **    14      13     fld_inputmask          F_INP
            ' **    15      14     fld_allowzerolength    F_ZERO
            ' **    16      15     fld_attributes         F_ATTR
            ' **    17      16     fld_datemodified       F_DATE
            ' **    18      17     Found                  F_FND
            ' **
            ' ********************************************************
          Next
          arr_varFld(F_FND, lngE) = CBool(False)
          If lngX < lngFlds Then .MoveNext
        Next
      End If
      .Close
    End With
    ' ** Set text fields to vbNullString so that comparisons don't error.
    For lngX = 0& To (lngFlds - 1&)
      If IsNull(arr_varFld(F_DESC, lngX)) = True Then
        arr_varFld(F_DESC, lngX) = vbNullString
      End If
      If IsNull(arr_varFld(F_DEF, lngX)) = True Then
        arr_varFld(F_DEF, lngX) = vbNullString
      End If
      If IsNull(arr_varFld(F_FRMT, lngX)) = True Then
        arr_varFld(F_FRMT, lngX) = vbNullString
      End If
      If IsNull(arr_varFld(F_INP, lngX)) = True Then
        arr_varFld(F_INP, lngX) = vbNullString
      End If
    Next

    .Close
  End With

  For lngX = 0& To (lngDats - 1&)
    If arr_varDat(D_DNAM, lngX) = strThisFile Then
      Set wrk = DBEngine.Workspaces(0)
    Else
On Error Resume Next
      Set wrk = CreateWorkspace("tmp", "Admin", "", dbUseJet)
      If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
        Set wrk = CreateWorkspace("tmp", "superuser", TA_SEC, dbUseJet)
        If ERR.Number <> 0 Then
On Error GoTo 0
          Set wrk = CreateWorkspace("tmp", "superuser", TA_SEC6, dbUseJet)
        Else
On Error GoTo 0
        End If
      Else
On Error GoTo 0
      End If
    End If
    With wrk
      If arr_varDat(D_DNAM, lngX) = strThisFile Then
        Set dbsLnk = .Databases(0)
      Else
        Set dbsLnk = .OpenDatabase(arr_varDat(D_PATH, lngX) & LNK_SEP & arr_varDat(D_DNAM, lngX), False, False)  ' ** {pathfile}, {exclusive}, {read-only}
      End If
      With dbsLnk
        dbsLnk.TableDefs.Refresh
        dbsLnk.TableDefs.Refresh
        Set dbsLoc = CurrentDb
        ' ** tblDatabase_Table_Field, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
        Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_Field_01")
        With qdf.Parameters
          ![wjrnl] = blnWithJrnlTmp
        End With
        Set rst = qdf.OpenRecordset
        For lngY = 0& To (lngTbls - 1&)
          If arr_varTbl(T_DID, lngY) = arr_varDat(D_DID, lngX) Then
On Error Resume Next
            Set tdf = .TableDefs(arr_varTbl(T_TNAM, lngY))
            If ERR.Number = 0 Then
On Error GoTo 0
              With tdf
                For Each fld In .Fields
                  blnDocFld = True
                  With fld
                    For lngZ = 0& To (lngFlds - 1&)
                      If arr_varFld(F_DID, lngZ) = arr_varDat(D_DID, lngX) And arr_varFld(F_TID, lngZ) = arr_varTbl(T_TID, lngY) And _
                        arr_varFld(F_FNAM, lngZ) = .Name Then
                        blnDocFld = False
                        arr_varFld(F_FND, lngZ) = True
                        For Each prp In .Properties
                          With prp
                            blnSQL = False: strFld = vbNullString

                            ' ** Check text properties.
                            If .Name = "Description" Or .Name = "DefaultValue" Or .Name = "Format" Or .Name = "InputMask" Then
                              strFld = "fld_" & LCase$(.Name)
                              If Trim$(.Value) <> vbNullString Then
                                Select Case .Name
                                Case "Description"
                                  If arr_varFld(F_DESC, lngZ) <> .Value Then  ' ** Nulls have been changed to vbNullString in the array.
                                    blnSQL = True
                                  End If
                                Case "DefaultValue"
                                  If arr_varFld(F_DEF, lngZ) <> .Value Then
                                    blnSQL = True
                                  End If
                                Case "Format"
                                  If arr_varFld(F_FRMT, lngZ) <> .Value Then
                                    blnSQL = True
                                  End If
                                Case "InputMask"
                                  If arr_varFld(F_INP, lngZ) <> .Value Then
                                    blnSQL = True
                                  End If
                                End Select
                              End If
                            End If
                            If blnSQL = True Then
                              strSQL = "UPDATE tblDatabase_Table_Field " & _
                                "SET tblDatabase_Table_Field." & strFld & " = " & _
                                  Chr(34) & IIf(Left$(.Value, 1) = Chr(34) And Right$(.Value, 1) = Chr(34), _
                                  Mid$(Left$(.Value, (Len(.Value) - 1)), 2), .Value) & Chr(34) & ", " & _
                                "tblDatabase_Table_Field.fld_datemodified = Now() " & _
                                "WHERE (((tblDatabase_Table_Field.dbs_id)=" & CStr(arr_varFld(F_DID, lngZ)) & ") AND " & _
                                "((tblDatabase_Table_Field.tbl_id)=" & CStr(arr_varFld(F_TID, lngZ)) & ") AND " & _
                                "((tblDatabase_Table_Field.fld_id)=" & CStr(arr_varFld(F_FID, lngZ)) & "));"
                              dbsLoc.Execute strSQL, dbFailOnError
                            End If
                            blnSQL = False

                            ' ** Check Boolean properties.
                            If .Name = "Required" Or .Name = "AllowZeroLength" Then
                              strFld = "fld_" & LCase$(.Name)
                              Select Case .Name
                              Case "Required"
                                If arr_varFld(F_REQ, lngZ) <> .Value Then
                                  blnSQL = True
                                End If
                              Case "AllowZeroLength"
                                If arr_varFld(F_ZERO, lngZ) <> .Value Then
                                  blnSQL = True
                                End If
                              End Select
                            End If
                            If blnSQL = True Then
                              strSQL = "UPDATE tblDatabase_Table_Field " & _
                                "SET tblDatabase_Table_Field." & strFld & " = " & _
                                IIf(.Value = True, "True", "False") & ", " & _
                                "tblDatabase_Table_Field.fld_datemodified = Now()" & _
                                "WHERE (((tblDatabase_Table_Field.dbs_id)=" & CStr(arr_varFld(F_DID, lngZ)) & ") AND " & _
                                "((tblDatabase_Table_Field.tbl_id)=" & CStr(arr_varFld(F_TID, lngZ)) & ") AND " & _
                                "((tblDatabase_Table_Field.fld_id)=" & CStr(arr_varFld(F_FID, lngZ)) & "));"
                              dbsLoc.Execute strSQL, dbFailOnError
                            End If
                            blnSQL = False

                            ' ** Check numeric properties.
                            If .Name = "Type" Or .Name = "Size" Or .Name = "ControlType" Or .Name = "DecimalPlaces" Or .Name = "Attributes" Then
                              strFld = "fld_" & LCase$(.Name)
                              Select Case .Name
                              Case "Type"
                                strFld = "datatype_db_type"
                                If IsNull(.Value) = False Then
                                  ' ** Field property isn't Null.
                                  If IsNull(arr_varFld(F_TYPE, lngZ)) = False Then
                                    ' ** Documented property isn't Null.
                                    If arr_varFld(F_TYPE, lngZ) <> .Value Then
                                      blnSQL = True
                                    End If
                                  Else
                                    ' ** Documented property is Null.
                                    blnSQL = True
                                  End If
                                Else
                                  ' ** Field property is Null.
                                  If IsNull(arr_varFld(F_TYPE, lngZ)) = False Then
                                    ' ** Documented property isn't Null.
                                    blnSQL = True
                                  End If
                                End If
                              Case "Size"
                                If IsNull(.Value) = False Then
                                  If IsNull(arr_varFld(F_SIZE, lngZ)) = False Then
                                    If arr_varFld(F_SIZE, lngZ) <> .Value Then
                                      blnSQL = True
                                    End If
                                  Else
                                    blnSQL = True
                                  End If
                                Else
                                  If IsNull(arr_varFld(F_SIZE, lngZ)) = False Then
                                    blnSQL = True
                                  End If
                                End If
                              Case "ControlType"
                                strFld = "dispctl_type"
                                If IsNull(.Value) = False Then
                                  If IsNull(arr_varFld(F_CTL, lngZ)) = False Then
                                    If arr_varFld(F_CTL, lngZ) <> .Value Then
                                      blnSQL = True
                                    End If
                                  Else
                                    blnSQL = True
                                  End If
                                Else
                                  If IsNull(arr_varFld(F_CTL, lngZ)) = False Then
                                    blnSQL = True
                                  End If
                                End If
                              Case "DecimalPlaces"
                                If IsNull(.Value) = False Then
                                  If IsNull(arr_varFld(F_DEC, lngZ)) = False Then
                                    If arr_varFld(F_DEC, lngZ) <> .Value Then
                                      blnSQL = True
                                    End If
                                  Else
                                    blnSQL = True
                                  End If
                                Else
                                  If IsNull(arr_varFld(F_DEC, lngZ)) = False Then
                                    blnSQL = True
                                  End If
                                End If
                              Case "Attributes"
                                If IsNull(.Value) = False Then
                                  If IsNull(arr_varFld(F_ATTR, lngZ)) = False Then
                                    If arr_varFld(F_ATTR, lngZ) <> .Value Then
                                      blnSQL = True
                                    End If
                                  Else
                                    blnSQL = True
                                  End If
                                Else
                                  If IsNull(arr_varFld(F_ATTR, lngZ)) = False Then
                                    blnSQL = True
                                  End If
                                End If
                              End Select
                            End If
                            If blnSQL = True Then
                              strSQL = "UPDATE tblDatabase_Table_Field " & _
                                "SET tblDatabase_Table_Field." & strFld & " = " & _
                                IIf(IsNull(.Value) = True, "Null", CStr(.Value)) & ", " & _
                                "tblDatabase_Table_Field.fld_datemodified = Now() " & _
                                "WHERE (((tblDatabase_Table_Field.dbs_id)=" & CStr(arr_varFld(F_DID, lngZ)) & ") AND " & _
                                "((tblDatabase_Table_Field.tbl_id)=" & CStr(arr_varFld(F_TID, lngZ)) & ") AND " & _
                                "((tblDatabase_Table_Field.fld_id)=" & CStr(arr_varFld(F_FID, lngZ)) & "));"
                              dbsLoc.Execute strSQL, dbFailOnError
                            End If
                            blnSQL = False

                          End With
                        Next
                      End If
                    Next
                  End With
                  If blnDocFld = True Then
                    With rst
                      .AddNew
                      ![dbs_id] = arr_varTbl(T_DID, lngY)
                      ![tbl_id] = arr_varTbl(T_TID, lngY)
                      ![fld_name] = fld.Name
                      ![fld_required] = fld.Required
                      ![datatype_db_type] = fld.Type
                      ![fld_size] = fld.Size
                      For Each prp In fld.Properties
                        With prp
                          Select Case .Name
                          Case "Description"
                            If .Value <> vbNullString Then
                              rst![fld_description] = .Value
                            End If
                          Case "DisplayControl"
On Error Resume Next
                            rst![dispctl_type] = CInt(.Value)
On Error GoTo 0
                          Case "Format"
                            If .Value <> vbNullString Then
                              rst![fld_format] = .Value
                            End If
                          Case "DecimalPlaces"
                            rst![fld_decimalplaces] = .Value
                          Case "InputMask"
                            If .Value <> vbNullString Then
                              rst![fld_inputmask] = .Value
                            End If
                          End Select
                        End With
                      Next
                      If fld.DefaultValue <> vbNullString Then
                        ![fld_defaultvalue] = fld.DefaultValue
                      End If
                        ![fld_attributes] = fld.Attributes
                      ' ** Field Attribute constant enumeration:
                      ' **      1  dbDescending      The field is sorted in descending (Z to A or 100 to 0) order;
                      ' **                           this option applies only to a Field object in a Fields collection of an Index object.
                      ' **                           If you omit this constant, the field is sorted in ascending (A to Z or 0 to 100) order.
                      ' **                           This is the default value for Index and TableDef fields (Microsoft Jet workspaces only).
                      ' **      1  dbFixedField      The field size is fixed (default for Numeric fields).
                      ' **      2  dbVariableField   The field size is variable (Text fields only).
                      ' **     16  dbAutoIncrField   The field value for new records is automatically incremented to a unique Long integer
                      ' **                           that can't be changed (in a Microsoft Jet workspace, supported only for
                      ' **                           Microsoft Jet database(.mdb) tables).
                      ' **     32  dbUpdatableField  The field value can be changed.
                      ' **   8192  dbSystemField     The field stores replication information for replicas;
                      ' **                           you can't delete this type of field (Microsoft Jet workspaces only).
                      ' **  32768  dbHyperlinkField  The field contains hyperlink information (Memo fields only).
                      ![fld_allowzerolength] = fld.AllowZeroLength
                      ![fld_datemodified] = Now()
                      .Update
                    End With
                  End If
                Next
              End With
            Else
On Error GoTo 0
            End If
          End If
        Next
        rst.Close
        dbsLoc.Close
        .Close
      End With
      .Close
    End With
  Next

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  For lngX = 0& To (lngFlds - 1&)
    If arr_varFld(F_FND, lngX) = False Then
      For lngY = 0& To (lngTbls - 1&)
        If arr_varTbl(T_TID, lngY) = arr_varFld(F_TID, lngX) Then
          If Left(arr_varTbl(T_TNAM, lngY), 4) <> "MSys" And Left(arr_varTbl(T_TNAM, lngY), 4) <> "USys" And _
              arr_varTbl(T_TNAM, lngY) <> "xlExcelImport" And arr_varTbl(T_DNAM, lngY) <> "TrstXAdm.mdb" Then
Select Case arr_varTbl(T_TNAM, lngY)
Case "m_TBL1", "m_TBL2", "zz_tbl_Database_Table_Link", "zz_tbl_Form_Doc", "zz_tbl_Form_Property", "zz_tbl_Form_Property_Value", _
    "zz_tbl_sql_code_01", "zz_tbl_sql_code_02", "zz_tbl_sql_code_03", "zz_tbl_sql_code_04", "zz_tbl_sql_code_05", _
    "zz_tbl_sql_code_06", "zz_tbl_VBComponent_KeyDown"
  ' ** Leave these in.
Case Else
            blnDelete = True
            DoCmd.Hourglass False
            Debug.Print "'DEL FLD? " & arr_varFld(F_FNAM, lngX) & " " & CStr(arr_varFld(F_FID, lngX)) & "  IN  " & arr_varTbl(T_TNAM, lngY)
Stop
            If blnDelete = True Then
              DoCmd.Hourglass True
              Set dbsLoc = CurrentDb
              ' ** Delete tblDatabase_Table_Field, by specified [fldid].
              Set qdf = dbsLoc.QueryDefs("zz_qry_Database_Table_Field_02")
              With qdf.Parameters
                ![fldid] = arr_varFld(F_FID, lngX)
              End With
              qdf.Execute dbFailOnError
              dbsLoc.Close
            Else
              DoCmd.Hourglass True
            End If
            Exit For
End Select
          End If
        End If
      Next
    End If
  Next

  DoEvents
  Qry_TmpTables False, blnWithJrnlTmp  ' ** Function: Below.

  DoCmd.Hourglass False
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set prp = Nothing
  Set fld = Nothing
  Set tdf = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set wrk = Nothing

  Tbl_Fld_Doc = blnRetValx

End Function

Public Function Tbl_Fld_RowSource_Doc() As Boolean
' ** Document all table combo box Row Sources to tblDatabase_Table_Field_RowSource.

  Const THIS_PROC As String = "Tbl_Fld_RowSource_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim tdf As DAO.TableDef, fld As DAO.Field, prp As DAO.Property
  Dim lngDats As Long, arr_varDat() As Variant
  Dim lngLastDbID As Long, lngLastTblID As Long, strLastTable As String
  Dim strPathFile As String
  Dim lngFlds As Long, arr_varFld() As Variant
  Dim lngCombos As Long, lngChks As Long, lngTxts As Long
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngRecs As Long
  Dim blnFound As Boolean, blnAdd As Boolean, blnUpdated As Boolean, blnDelete As Boolean
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varDat().
  Const D_DID  As Integer = 0
  Const D_DNAM As Integer = 1
  Const D_PATH As Integer = 2
  Const D_TBLS As Integer = 3
  Const D_RELS As Integer = 4
  Const D_IDXS As Integer = 5
  Const D_DAT  As Integer = 6

  ' ** Array: arr_varFld().
  Const F_ELEMS As Integer = 16  ' ** Array's first-element UBound().
  Const F_DID       As Integer = 0
  Const F_DNAM      As Integer = 1
  Const F_TID       As Integer = 2
  Const F_TNAM      As Integer = 3
  Const F_FID       As Integer = 4
  Const F_FNAM      As Integer = 5
  Const F_DTATYP    As Integer = 6
  Const F_CTLTYP    As Integer = 7
  Const F_ROWSRCTYP As Integer = 8
  Const F_ROWSRC    As Integer = 9
  Const F_BNDCOL    As Integer = 10
  Const F_COLCNT    As Integer = 11
  Const F_COLHEAD   As Integer = 12
  Const F_COLWDTHS  As Integer = 13
  Const F_LSTROWS   As Integer = 14
  Const F_LSTWDTH   As Integer = 15
  Const F_LMTTOLST  As Integer = 16

  ' ** Array: arr_varDel().
  Const DEL_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const DEL_RID  As Integer = 0
  Const DEL_DNAM As Integer = 1
  Const DEL_TNAM As Integer = 2
  Const DEL_FNAM As Integer = 3

If TableExists("zz_tbl_RePost_Posting") = True Then
Stop
End If
  blnRetValx = True
  DoCmd.Hourglass True

  ' ** Get a list of databases to document.
  Set dbs = CurrentDb
  With dbs
    ' ** tblDatabase, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngDats = .RecordCount
      .MoveFirst
      arr_varDat = .GetRows(lngDats)
      ' *****************************************************
      ' ** Array: arr_varDat()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     dbs_id              D_DID
      ' **     2       1     dbs_name            D_DNAM
      ' **     3       2     dbs_path            D_PATH
      ' **     4       3     dbs_tbl_cnt         D_TBLS
      ' **     5       4     dbs_rel_cnt         D_RELS
      ' **     6       5     dbs_idx_cnt         D_IDXS
      ' **     7       6     dbs_datemodified    D_DAT
      ' **
      ' *****************************************************
      .Close
    End With
    .Close
  End With

  lngFlds = 0&
  ReDim arr_varFld(F_ELEMS, 0)

  For lngX = 0& To (lngDats - 1&)

    If arr_varDat(D_DNAM, lngX) = CurrentAppName Then  ' ** Module Function: modFileUtilities.
      Set dbs = CurrentDb
    Else
      strPathFile = arr_varDat(D_PATH, lngX)
      If Right$(strPathFile, 1) = LNK_SEP Then
        strPathFile = strPathFile & arr_varDat(D_DNAM, lngX)
      Else
        strPathFile = strPathFile & LNK_SEP & arr_varDat(D_DNAM, lngX)
      End If
      Set dbs = DBEngine.OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
    End If

    With dbs
      For Each tdf In .TableDefs
        With tdf
          If .Connect = vbNullString Then  ' ** Don't look at linked tables.
            For Each fld In .Fields
              blnFound = False
              With fld
                For Each prp In .Properties
                  Select Case prp.Name
                  Case "DisplayControl"
                    If IsNull(prp.Value) = False Then
                      blnFound = True
                      lngFlds = lngFlds + 1&
                      lngE = lngFlds - 1&
                      ReDim Preserve arr_varFld(F_ELEMS, lngE)
                      ' ***********************************************************
                      ' ** Array: arr_varFld()
                      ' **
                      ' **   Field  Element  Name                   Constant
                      ' **   =====  =======  =====================  =============
                      ' **     1       0     dbs_id                 F_DID
                      ' **     2       1     dbs_name               F_DNAM
                      ' **     3       2     tbl_id                 F_TID
                      ' **     4       3     tbl_name               F_TNAM
                      ' **     5       4     fld_id                 F_FID
                      ' **     6       5     fld_name               F_FNAM
                      ' **     7       6     datatype_db_type       F_DTATYP
                      ' **     8       7     ctltype_type           F_CTLTYP
                      ' **     9       8     rowsrctype_type        F_ROWSRCTYP
                      ' **    10       9     rowsrc_rowsource       F_ROWSRC
                      ' **    11      10     rowsrc_boundcolumn     F_BNDCOL
                      ' **    12      11     rowsrc_columncount     F_COLCNT
                      ' **    13      12     rowsrc_columnheads     F_COLHEAD
                      ' **    14      13     rowsrc_columnwidths    F_COLWDTHS
                      ' **    15      14     rowsrc_listrows        F_LSTROWS
                      ' **    16      15     rowsrc_listwidth       F_LSTWDTH
                      ' **    17      16     rowsrc_limittolist     F_LMTTOLST
                      ' **
                      ' ***********************************************************
                      arr_varFld(F_DID, lngE) = arr_varDat(D_DID, lngX)
                      arr_varFld(F_DNAM, lngE) = arr_varDat(D_DNAM, lngX)
                      arr_varFld(F_TID, lngE) = CLng(0)  ' ** We'll get it later.
                      arr_varFld(F_TNAM, lngE) = tdf.Name
                      arr_varFld(F_FID, lngE) = CLng(0)  ' ** We'll get it later.
                      arr_varFld(F_FNAM, lngE) = .Name
                      arr_varFld(F_DTATYP, lngE) = .Type
                      arr_varFld(F_CTLTYP, lngE) = prp.Value
                      Exit For
                    End If
                  End Select
                Next  ' ** prp.
                If blnFound = True Then
                  ' ** Only do this loop if it has a DisplayControl listed.
                  Select Case arr_varFld(F_CTLTYP, lngE)
                  Case acComboBox
                    For Each prp In .Properties
                      Select Case prp.Name
                      Case "RowSourceType"
                        arr_varFld(F_ROWSRCTYP, lngE) = prp.Value
                      Case "RowSource"
                        arr_varFld(F_ROWSRC, lngE) = prp.Value
                      Case "BoundColumn"
                        arr_varFld(F_BNDCOL, lngE) = prp.Value
                      Case "ColumnCount"
                        arr_varFld(F_COLCNT, lngE) = prp.Value
                      Case "ColumnHeads"
                        arr_varFld(F_COLHEAD, lngE) = prp.Value
                      Case "ColumnWidths"
                        arr_varFld(F_COLWDTHS, lngE) = prp.Value
                      Case "ListRows"
                        arr_varFld(F_LSTROWS, lngE) = prp.Value
                      Case "ListWidth"
                        arr_varFld(F_LSTWDTH, lngE) = prp.Value
                      Case "LimitToList"
                        arr_varFld(F_LMTTOLST, lngE) = prp.Value
                      End Select
                      ' ** These are not found in a table combo box.
                      ' **   rowsrc_autoexpand
                      ' **   rowsrc_multiselect
                      ' **   rowsrc_hasformref
                      ' **   rowsrc_formref
                    Next  ' ** prp.
                  Case acCheckBox
                    If arr_varFld(F_DTATYP, lngE) <> dbBoolean Then
                      Debug.Print "'TBL: " & arr_varFld(F_TNAM, lngE) & "  FLD: " & arr_varFld(F_FNAM, lngE)
                      Debug.Print "'  acCheckBox, BUT NOT BOOLEAN FIELD!"
                    End If
                    ' ** Nothing further.
                  Case acTextBox
                    ' ** Nothing further.
                  End Select
                End If
              End With  ' ** fld.
            Next  ' ** fld.
          End If  ' ** Local only.
        End With  ' ** tdf.
      Next  ' ** tdf.
      .Close
    End With

  Next  ' ** lngX.

  lngCombos = 0&: lngChks = 0&: lngTxts = 0&
  For lngX = 0& To (lngFlds - 1&)
    Select Case arr_varFld(F_CTLTYP, lngX)
    Case acComboBox
      lngCombos = lngCombos + 1&
    Case acTextBox
      lngTxts = lngTxts + 1&
    Case acCheckBox
      lngChks = lngChks + 1&
    Case Else
      Debug.Print "'CTLTYP: " & CStr(arr_varFld(F_CTLTYP, lngX))
    End Select
  Next

  Set dbs = CurrentDb
  With dbs

    ' ** tblDatabase_Table, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      lngLastDbID = 0&: lngLastTblID = 0&: strLastTable = vbNullString
      For lngX = 0& To (lngFlds - 1&)
        If arr_varFld(F_TNAM, lngX) <> strLastTable Or arr_varFld(F_DID, lngX) <> lngLastDbID Then
          .FindFirst "[dbs_id] = " & CStr(arr_varFld(F_DID, lngX)) & " And [tbl_name] = '" & arr_varFld(F_TNAM, lngX) & "'"
          If .NoMatch = False Then
            lngLastDbID = arr_varFld(F_DID, lngX)
            lngLastTblID = ![tbl_id]
          Else
            lngLastDbID = 0&
            lngLastTblID = 0&
            If (arr_varFld(F_TNAM, lngX) = "zz_tbl_Form_Control_01" And arr_varFld(F_DNAM, lngX) = "Trust.mdb") Or _
                (arr_varFld(F_TNAM, lngX) = "zz_tbl_Form_Shortcut_tmp01" And arr_varFld(F_DNAM, lngX) = "Trust.mdb") Or _
                (arr_varFld(F_TNAM, lngX) = "zz_tbl_m_TBL_tmp01" And arr_varFld(F_DNAM, lngX) = "Trust.mdb") Or _
                (arr_varFld(F_TNAM, lngX) = "zz_tbl_VBComponent_KeyDown" And arr_varFld(F_DNAM, lngX) = "Trust.mdb") Then
              ' ** Just ignore this!
            Else
              Debug.Print "'TBLX: " & arr_varFld(F_TNAM, lngX) & "  IN: " & arr_varFld(F_DNAM, lngX)
            End If
          End If
        End If
        arr_varFld(F_TID, lngX) = lngLastTblID
      Next
      .Close
    End With

    ' ** tblDatabase_Table_Field, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_Field_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      For lngX = 0& To (lngFlds - 1&)
        If arr_varFld(F_TID, lngX) > 0& Then
          .FindFirst "[tbl_id] = " & CStr(arr_varFld(F_TID, lngX)) & " And [fld_name] = '" & arr_varFld(F_FNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varFld(F_FID, lngX) = ![fld_id]
          Else
            Debug.Print "'FLDX: " & arr_varFld(F_FNAM, lngX) & "  IN: " & arr_varFld(F_TNAM, lngX)
          End If
        End If
      Next
      .Close
    End With

    ' ** tblDatabase_Table_Field_RowSource, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_Field_RowSource_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst

      lngE = -1&  ' ** To be sure it's not hanging around.
      For lngX = 0& To (lngFlds - 1&)
        If arr_varFld(F_TID, lngX) > 0& And arr_varFld(F_FID, lngX) > 0& Then
          Select Case arr_varFld(F_CTLTYP, lngX)
          Case acComboBox
            blnAdd = False: blnUpdated = False
            If .BOF = True And .EOF = True Then
              blnAdd = True
            Else
'rowsrc_id
              .FindFirst "[dbs_id] = " & CStr(arr_varFld(F_DID, lngX)) & " And " & _
                "[tbl_id] = " & CStr(arr_varFld(F_TID, lngX)) & " And " & _
                "[fld_id] = " & CStr(arr_varFld(F_FID, lngX))
              If .NoMatch = True Then
                blnAdd = True
              End If
            End If
            If blnAdd = True Then
              .AddNew
'dbs_id
              ![dbs_id] = arr_varFld(F_DID, lngX)
'tbl_id
              ![tbl_id] = arr_varFld(F_TID, lngX)
'fld_id
              ![fld_id] = arr_varFld(F_FID, lngX)
            Else
              .Edit
            End If
'ctltype_type
            If IsNull(![ctltype_type]) = True Then
              ![ctltype_type] = arr_varFld(F_CTLTYP, lngX): blnUpdated = True
            Else
              If ![ctltype_type] <> arr_varFld(F_CTLTYP, lngX) Then
                ![ctltype_type] = arr_varFld(F_CTLTYP, lngX): blnUpdated = True
              End If
            End If
            If arr_varFld(F_ROWSRCTYP, lngX) <> vbNullString Then
'rowsrctype_type
              If IsNull(![rowsrctype_type]) = True Then
                ![rowsrctype_type] = arr_varFld(F_ROWSRCTYP, lngX): blnUpdated = True
              Else
                If ![rowsrctype_type] <> arr_varFld(F_ROWSRCTYP, lngX) Then
                  ![rowsrctype_type] = arr_varFld(F_ROWSRCTYP, lngX): blnUpdated = True
                End If
              End If
              If arr_varFld(F_ROWSRCTYP, lngX) = "Table/Query" Then
                If arr_varFld(F_ROWSRC, lngX) = vbNullString Then
'qrytbltype_type
                  If IsNull(![qrytbltype_type]) = True Then
                    ![qrytbltype_type] = acNothing: blnUpdated = True
                  Else
                    If ![qrytbltype_type] <> acNothing Then
                      ![qrytbltype_type] = acNothing: blnUpdated = True
                    End If
                  End If
'rowsrc_rowsource
                  If IsNull(![rowsrc_rowsource]) = True Then
                    ![rowsrc_rowsource] = "{empty}": blnUpdated = True
                  Else
                    If ![rowsrc_rowsource] <> "{empty}" Then
                      ![rowsrc_rowsource] = "{empty}": blnUpdated = True
                    End If
                  End If
                Else
                  If Left$(arr_varFld(F_ROWSRC, lngX), 3) = "qry" Then
                    If IsNull(![qrytbltype_type]) = True Then
                      ![qrytbltype_type] = acQuery: blnUpdated = True
                    Else
                      If ![qrytbltype_type] <> acQuery Then
                        ![qrytbltype_type] = acQuery: blnUpdated = True
                      End If
                    End If
                  Else
                    If InStr(arr_varFld(F_ROWSRC, lngX), "SELECT") > 0 Then
                      If IsNull(![qrytbltype_type]) = True Then
                        ![qrytbltype_type] = acSQL: blnUpdated = True
                      Else
                        If ![qrytbltype_type] <> acSQL Then
                          ![qrytbltype_type] = acSQL: blnUpdated = True
                        End If
                      End If
                    Else
                      If IsNull(![qrytbltype_type]) = True Then
                        ![qrytbltype_type] = acTable: blnUpdated = True
                      Else
                        If ![qrytbltype_type] <> acTable Then
                          ![qrytbltype_type] = acTable: blnUpdated = True
                        End If
                      End If
                    End If
                  End If
                  If IsNull(![rowsrc_rowsource]) = True Then
                    ![rowsrc_rowsource] = arr_varFld(F_ROWSRC, lngX): blnUpdated = True
                  Else
                    If ![rowsrc_rowsource] <> arr_varFld(F_ROWSRC, lngX) Then
                      ![rowsrc_rowsource] = arr_varFld(F_ROWSRC, lngX): blnUpdated = True
                    End If
                  End If
                End If
              Else
                If IsNull(![qrytbltype_type]) = True Then
                  ![qrytbltype_type] = acNothing: blnUpdated = True
                Else
                  If ![qrytbltype_type] <> acNothing Then
                    ![qrytbltype_type] = acNothing: blnUpdated = True
                  End If
                End If
                If arr_varFld(F_ROWSRC, lngX) = vbNullString Then
                  If IsNull(![rowsrc_rowsource]) = True Then
                    ![rowsrc_rowsource] = "{empty}": blnUpdated = True
                  Else
                    If ![rowsrc_rowsource] <> "{empty}" Then
                      ![rowsrc_rowsource] = "{empty}": blnUpdated = True
                    End If
                  End If
                Else
                  If IsNull(![rowsrc_rowsource]) = True Then
                    ![rowsrc_rowsource] = arr_varFld(F_ROWSRC, lngX): blnUpdated = True
                  Else
                    If ![rowsrc_rowsource] <> arr_varFld(F_ROWSRC, lngX) Then
                      ![rowsrc_rowsource] = arr_varFld(F_ROWSRC, lngX): blnUpdated = True
                    End If
                  End If
                End If
              End If
              If ![rowsrc_rowsource] <> "{empty}" Then
'rowsrc_boundcolumn
                If IsNull(![rowsrc_boundcolumn]) = True Then
                  ![rowsrc_boundcolumn] = arr_varFld(F_BNDCOL, lngX): blnUpdated = True
                Else
                  If ![rowsrc_boundcolumn] <> arr_varFld(F_BNDCOL, lngX) Then
                    ![rowsrc_boundcolumn] = arr_varFld(F_BNDCOL, lngX): blnUpdated = True
                  End If
                End If
'rowsrc_columncount
                If IsNull(![rowsrc_columncount]) = True Then
                  ![rowsrc_columncount] = arr_varFld(F_COLCNT, lngX): blnUpdated = True
                Else
                  If ![rowsrc_columncount] <> arr_varFld(F_COLCNT, lngX) Then
                    ![rowsrc_columncount] = arr_varFld(F_COLCNT, lngX): blnUpdated = True
                  End If
                End If
                If arr_varFld(F_COLWDTHS, lngX) <> vbNullString Then
'rowsrc_columnwidths
                  If IsNull(![rowsrc_columnwidths]) = True Then
                    ![rowsrc_columnwidths] = arr_varFld(F_COLWDTHS, lngX): blnUpdated = True
                  Else
                    If ![rowsrc_columnwidths] <> arr_varFld(F_COLWDTHS, lngX) Then
                      ![rowsrc_columnwidths] = arr_varFld(F_COLWDTHS, lngX): blnUpdated = True
                    End If
                  End If
                End If
'rowsrc_listwidth
                If Right$(arr_varFld(F_LSTWDTH, lngX), 4) = "twip" Then
                  strTmp01 = Left$(arr_varFld(F_LSTWDTH, lngX), (Len(arr_varFld(F_LSTWDTH, lngX)) - 4))
                  If IsNull(![rowsrc_listwidth]) = True Then
                    ![rowsrc_listwidth] = Val(strTmp01): blnUpdated = True
                  Else
                    If ![rowsrc_listwidth] <> Val(strTmp01) Then
                      ![rowsrc_listwidth] = Val(strTmp01): blnUpdated = True
                    End If
                  End If
                Else
                  If IsNull(![rowsrc_listwidth]) = True Then
                    ![rowsrc_listwidth] = Val(arr_varFld(F_LSTWDTH, lngX)): blnUpdated = True
                  Else
                    If ![rowsrc_listwidth] <> Val(arr_varFld(F_LSTWDTH, lngX)) Then
                      ![rowsrc_listwidth] = Val(arr_varFld(F_LSTWDTH, lngX)): blnUpdated = True
                    End If
                  End If
                End If
              End If
'rowsrc_columnheads
              If IsNull(![rowsrc_columnheads]) = True Then
                ![rowsrc_columnheads] = arr_varFld(F_COLHEAD, lngX): blnUpdated = True
              Else
                If ![rowsrc_columnheads] <> arr_varFld(F_COLHEAD, lngX) Then
                  ![rowsrc_columnheads] = arr_varFld(F_COLHEAD, lngX): blnUpdated = True
                End If
              End If
'rowsrc_limittolist
              If IsNull(![rowsrc_limittolist]) = True Then
                ![rowsrc_limittolist] = arr_varFld(F_LMTTOLST, lngX): blnUpdated = True
              Else
                If ![rowsrc_limittolist] <> arr_varFld(F_LMTTOLST, lngX) Then
                  ![rowsrc_limittolist] = arr_varFld(F_LMTTOLST, lngX): blnUpdated = True
                End If
              End If
'rowsrc_listrows
              If IsNull(![rowsrc_listrows]) = True Then
                ![rowsrc_listrows] = arr_varFld(F_LSTROWS, lngX): blnUpdated = True
              Else
                If ![rowsrc_listrows] <> arr_varFld(F_LSTROWS, lngX) Then
                  ![rowsrc_listrows] = arr_varFld(F_LSTROWS, lngX): blnUpdated = True
                End If
              End If
            Else
              If IsNull(![rowsrctype_type]) = True Then
                ![rowsrctype_type] = "Table/Query": blnUpdated = True
              Else
                If ![rowsrctype_type] <> "Table/Query" Then
                  ![rowsrctype_type] = "Table/Query": blnUpdated = True
                End If
              End If
              If IsNull(![qrytbltype_type]) = True Then
                ![qrytbltype_type] = acNothing: blnUpdated = True
              Else
                If ![qrytbltype_type] <> acNothing Then
                  ![qrytbltype_type] = acNothing: blnUpdated = True
                End If
              End If
              If IsNull(![rowsrc_rowsource]) = True Then
                ![rowsrc_rowsource] = "{empty}": blnUpdated = True
              Else
                If ![rowsrc_rowsource] <> "{empty}" Then
                  ![rowsrc_rowsource] = "{empty}": blnUpdated = True
                End If
              End If
            End If
            ' ** These are not found in a table combo box.
            ' **   rowsrc_autoexpand
            ' **   rowsrc_multiselect
            ' **   rowsrc_hasformref
            ' **   rowsrc_formref
            If blnUpdated = True Then
              ![rowsrc_datemodified] = Now()
            End If
            .Update
          End Select
        End If  ' ** lngTblID, lngFldID.
      Next

      ' ** qrytbltype_type:
      ' **   acTable
      ' **   acQuery
      ' **   acSQL
      ' **   acNothing

      .Close
    End With  ' ** rst.

    lngDels = 0&
    ReDim arr_varDel(DEL_ELEMS, 0)

    Set rst = .OpenRecordset("tblDatabase_Table_Field_RowSource", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        blnFound = False
        For lngY = 0& To (lngFlds - 1&)
          If arr_varFld(F_CTLTYP, lngY) = acComboBox Then
            If arr_varFld(F_DID, lngY) = ![dbs_id] And arr_varFld(F_TID, lngY) = ![tbl_id] And arr_varFld(F_FID, lngY) = ![fld_id] Then
              blnFound = True
              Exit For
            End If
          End If
        Next
        If blnFound = False Then
          lngDels = lngDels + 1&
          lngE = lngDels - 1&
          ReDim Preserve arr_varDel(DEL_ELEMS, lngE)
          arr_varDel(DEL_RID, lngE) = ![rowsrc_id]
          varTmp00 = DLookup("[dbs_name]", "tblDatabase", "[dbs_id] = " & CStr(![dbs_id]))
          arr_varDel(DEL_DNAM, lngE) = varTmp00
          varTmp00 = DLookup("[tbl_name]", "tblDatabase_Table", "[tbl_id] = " & CStr(![tbl_id]))
          arr_varDel(DEL_TNAM, lngE) = varTmp00
          varTmp00 = DLookup("[fld_name]", "tblDatabase_Table_Field", "[fld_id] = " & CStr(![fld_id]))
          arr_varDel(DEL_FNAM, lngE) = varTmp00
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With  ' ** rst.

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    If lngDels > 0& Then
      For lngX = 0& To (lngDels - 1&)
        If arr_varDel(DEL_DNAM, lngX) = "TrustImport.mdb" Or arr_varDel(DEL_DNAM, lngX) = "TrustImport.mde" Or _
            arr_varDel(DEL_DNAM, lngX) = "TrstXAdm.mdb" Or arr_varDel(DEL_DNAM, lngX) = "TrstXAdm.mde" Then
          ' ** Let Trust Import handle these.
        Else
          blnDelete = True
          Debug.Print "'DEL COMBO? " & arr_varDel(DEL_FNAM, lngX) & "  IN  " & arr_varDel(DEL_TNAM, lngX) & ", " & arr_varDel(DEL_DNAM, lngX)
Stop
          If blnDelete = True Then
            ' ** Delete tblDatabase_Table_Field_RowSource, by specified [rsrcid].
            Set qdf = .QueryDefs("zz_qry_Database_Table_Field_RowSource_02")
            With qdf.Parameters
              ![rsrcid] = arr_varDel(DEL_RID, lngX)
            End With
            qdf.Execute
          End If
        End If
      Next
    End If

    .Close
  End With  ' ** dbs.

  Debug.Print "'COMBOS: " & CStr(lngCombos) & "  CHKS: " & CStr(lngChks) & "  TXTS: " & CStr(lngTxts)

' ** Property list of a field with a combo box:
'Value
'Attributes
'CollatingOrder
'Type
'Name
'OrdinalPosition
'Size
'SourceField
'SourceTable
'ValidateOnSet
'DataUpdatable
'ForeignName
'DefaultValue
'ValidationRule
'ValidationText
'Required
'AllowZeroLength
'FieldSize
'OriginalValue
'VisibleValue
'ColumnWidth
'ColumnOrder
'ColumnHidden
'DecimalPlaces
'DisplayControl
'RowSourceType
'RowSource
'BoundColumn
'ColumnCount
'ColumnHeads
'ListRows
'ListWidth
'LimitToList
'ColumnWidths
'GUID

If TableExists("zz_tbl_RePost_Posting") = True Then
Stop
End If
  DoCmd.Hourglass False
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set prp = Nothing
  Set fld = Nothing
  Set tdf = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Tbl_Fld_RowSource_Doc = blnRetValx

End Function

Public Function Tbl_Fld_DateFormat_Doc() As Boolean

  Const THIS_PROC As String = "Tbl_Fld_DateFormat_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngFlds As Long, arr_varFld As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim blnAdd As Boolean, blnFound As Boolean
  Dim lngRecs As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varFld().
  Const F_DID  As Integer = 0
  Const F_DNAM As Integer = 1
  Const F_TID  As Integer = 2
  Const F_TNAM As Integer = 3
  Const F_FID  As Integer = 4
  Const F_FNAM As Integer = 5
  Const F_TYP  As Integer = 6
  Const F_DEF  As Integer = 7
  Const F_FRMT As Integer = 8
  Const F_MASK As Integer = 9
  Const F_ATTR As Integer = 10
  Const F_DAT  As Integer = 11
  Const F_NOW  As Integer = 12
  Const F_DDEF As Integer = 13
  Const F_NDEF As Integer = 14

  blnRetValx = True
  DoCmd.Hourglass True
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** tblDatabase_Table_Field, just Date fields.
    Set qdf = .QueryDefs("zz_qry_Database_Table_Field_DateFormat_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngFlds = .RecordCount
      .MoveFirst
      arr_varFld = .GetRows(lngFlds)
      ' *****************************************************
      ' ** Array: arr_varFld()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     dbs_id              F_DID
      ' **     2       1     dbs_name            F_DNAM
      ' **     3       2     tbl_id              F_TID
      ' **     4       3     tbl_name            F_TNAM
      ' **     5       4     fld_id              F_FID
      ' **     6       5     fld_name            F_FNAM
      ' **     7       6     datatype_db_type    F_TYP
      ' **     8       7     fld_defaultvalue    F_DEF
      ' **     9       8     fld_format          F_FRMT
      ' **    10       9     fld_inputmask       F_MASK
      ' **    11      10     fld_attributes      F_ATTR
      ' **    12      11     datfrmt_date        F_DAT
      ' **    13      12     datfrmt_now         F_NOW
      ' **    14      13     datfrmt_datedef     F_DDEF
      ' **    15      14     datfrmt_nowdef      F_NDEF
      ' **
      ' *****************************************************
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Set qdf = Nothing

    Set rst = .OpenRecordset("tblDatabase_Table_Field_DateFormat", dbOpenDynaset, dbConsistent)
    With rst
      If .BOF = True And .EOF = True Then
        lngRecs = 0&
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
      End If
      For lngX = 0& To (lngFlds - 1&)
        blnAdd = False
        If lngRecs = 0& Then
          blnAdd = True
          .AddNew
        Else
          .FindFirst "[dbs_id] = " & CStr(arr_varFld(F_DID, lngX)) & " And " & _
            "[tbl_id] = " & CStr(arr_varFld(F_TID, lngX)) & " And " & _
            "[fld_id] = " & CStr(arr_varFld(F_FID, lngX))
          Select Case .NoMatch
          Case True
            blnAdd = True
            .AddNew
          Case False
            .Edit
          End Select
        End If
        If blnAdd = True Then
          ![dbs_id] = arr_varFld(F_DID, lngX)
          ![tbl_id] = arr_varFld(F_TID, lngX)
          ![fld_id] = arr_varFld(F_FID, lngX)
        End If
        ![datfrmt_date] = arr_varFld(F_DAT, lngX)
        ![datfrmt_now] = arr_varFld(F_NOW, lngX)
        ![datfrmt_datedef] = arr_varFld(F_DDEF, lngX)
        ![datfrmt_nowdef] = arr_varFld(F_NDEF, lngX)
        ![datfrmt_datemodified] = Now()
        .Update
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
      Next  ' ** lngX
      .Close
    End With  ' ** rst.
    Set rst = Nothing

    lngDels = 0&
    ReDim arr_varDel(0)

    Set rst = .OpenRecordset("tblDatabase_Table_Field_DateFormat", dbOpenDynaset, dbConsistent)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        blnFound = False
        For lngY = 0& To (lngFlds - 1&)
          If arr_varFld(F_DID, lngY) = ![dbs_id] And arr_varFld(F_TID, lngY) = ![tbl_id] And _
              arr_varFld(F_FID, lngY) = ![fld_id] Then
            blnFound = True
            Exit For
          End If
        Next  ' ** lngY.
        If blnFound = False Then
          lngDels = lngDels + 1&
          lngE = lngDels - 1&
          ReDim arr_varDel(lngE)
          arr_varDel(lngE) = ![datfrmt_id]
        End If
        If lngX < lngRecs Then .MoveNext
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.

    If lngDels > 0& Then
      For lngX = 0& To (lngDels - 1&)
        ' ** Delete tblDatabase_Table_Field_DateFormat, by specified [dfmtid].
        Set qdf = .QueryDefs("zz_qry_Database_Table_Field_DateFormat_02")
        With qdf.Parameters
          ![dfmtid] = arr_varDel(lngX)
        End With
        qdf.Execute
        Set qdf = Nothing
      Next
    End If

    .Close
  End With  ' ** dbs.

  DoCmd.Hourglass False

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Tbl_Fld_DateFormat_Doc = blnRetValx

End Function

Public Function Tbl_RecCnt_Doc() As Boolean
' ** Document the current number of records in each table to tblDataBase_Table_RecordCount.

  Const THIS_PROC As String = "Tbl_RecCnt_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim lngRecs As Long
  Dim lngDbsID As Long, lngTblID As Long
  Dim blnAdd As Boolean, blnSkip As Boolean
  Dim lngMSysCnt As Long, lngSysTmpCnt As Long

  blnRetValx = True
  DoCmd.Hourglass True
  DoEvents

  Qry_TmpTables True, blnWithJrnlTmp  ' ** Function: Below.

  DoCmd.Hourglass True
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** tblDatabase_Table_RecordCount, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_Table_RecCnt_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst1 = qdf.OpenRecordset

    lngMSysCnt = 0&: lngSysTmpCnt = 0&
    .TableDefs.Refresh
    .TableDefs.Refresh
    For Each tdf In .TableDefs
      lngRecs = 0&: blnAdd = False: blnSkip = False
      With tdf
        If Left$(.Name, 4) <> "MSys" And Left$(.Name, 4) <> "~TMP" Then  ' ** Skip those pesky system tables.
          Set rst2 = dbs.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
          With rst2
            If .BOF = True And .EOF = True Then
              lngRecs = 0&
            Else
              .MoveLast
              lngRecs = .RecordCount
            End If
            .Close
          End With
          If .Connect = vbNullString Then
            lngDbsID = Nz(DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & Parse_File(CurrentDb.Name) & "'"), 0)
            If lngDbsID = 0& Then Stop
          Else
            If Parse_File(.Connect) = "TAJrnTmp.mdb" And blnWithJrnlTmp = False Then  ' ** Module Function: modFileUtilities.
              blnSkip = True
            Else
              lngDbsID = Nz(DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & Parse_File(.Connect) & "'"), 0)
              If lngDbsID = 0& Then Stop
            End If
          End If
          If blnSkip = False Then
            lngTblID = Nz(DLookup("[tbl_id]", "tblDatabase_Table", "[tbl_name] = '" & .Name & "' And [dbs_id] = " & CStr(lngDbsID)), 0)
            If lngTblID = 0& Then
              ' ** Maybe it's got a different name here than in its home database.
              lngTblID = Nz(DLookup("[tbl_id]", "tblDatabase_Table", "[tbl_name] = '" & .SourceTableName & "' And [dbs_id] = " & CStr(lngDbsID)), 0)
              If lngTblID = 0& Then Stop
            End If
            With rst1
              If .BOF = True And .EOF = True Then
                blnAdd = True
              Else
                .FindFirst "[tbl_id] = " & CStr(lngTblID)
                If .NoMatch = False Then
                  .Edit
                  ![reccnt_count] = lngRecs
                  ![reccnt_datemodified] = Now()
                  .Update
                Else
                  blnAdd = True
                End If
              End If
              If blnAdd = True Then
                .AddNew
                ![dbs_id] = lngDbsID
                ![tbl_id] = lngTblID
                ![reccnt_count] = lngRecs
                ![reccnt_datemodified] = Now()
                .Update
              End If
            End With
          End If  ' ** blnSkip.
        Else
          If Left$(.Name, 4) = "MSys" Then
            lngMSysCnt = lngMSysCnt + 1&
          ElseIf Left$(.Name, 4) = "~TMP" Then
            lngSysTmpCnt = lngSysTmpCnt + 1&
            'Set rst2 = .OpenRecordset(.Name)
            'With rst2
            '  If .BOF = True And .EOF = True Then
            '    ' ** Empty.
            '  Else
            '    .MoveLast
            '    Debug.Print "'SYS TEMP TBL: " & .Name & "  RECS: " & CStr(.RecordCount)
            '  End If
            '  .Close
            'End With
          End If
        End If
      End With
    Next

    rst1.Close

    .Close
  End With

  DoEvents
  Qry_TmpTables False, blnWithJrnlTmp  ' ** Function: Below.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  DoCmd.Hourglass False
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rst1 = Nothing
  Set rst2 = Nothing
  Set tdf = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Tbl_RecCnt_Doc = blnRetValx

End Function

Public Function Tbl_AutoNum_Doc() As Boolean
' ** Document all AutoNumber fields to tblDatabase_AutoNumber.

  Const THIS_PROC As String = "Tbl_AutoNum_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, tdf As DAO.TableDef, fld As DAO.Field
  Dim lngDbsID As Long, lngTblID As Long, lngFldID As Long
  Dim strDbsName As String, strDbsPath As String, strTableName As String
  Dim lngFldFnds As Long, arr_varFldFnd() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngLastSeed As Long
  Dim blnAdd As Boolean, blnFound As Boolean, blnGetSeed As Boolean, blnSkip As Boolean
  Dim lngRecs As Long
  Dim varTmp00 As Variant
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varFldFnd().
  Const FF_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const FF_AID  As Integer = 0
  Const FF_DID  As Integer = 1
  Const FF_DNAM As Integer = 2
  Const FF_PATH As Integer = 3
  Const FF_TID  As Integer = 4
  Const FF_TNAM As Integer = 5
  Const FF_FID  As Integer = 6
  Const FF_FNAM As Integer = 7
  Const FF_LID  As Integer = 8
  Const FF_SEED As Integer = 9

  blnRetValx = True
  DoCmd.Hourglass True
  DoEvents

  blnGetSeed = False ' ** When True, it takes a very long time.

  Qry_TmpTables True, blnWithJrnlTmp  ' ** Function: Below.

  DoCmd.Hourglass True
  DoEvents

  lngFldFnds = 0&
  ReDim arr_varFldFnd(FF_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    ' ** tblDatabase_AutoNumber, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_AutoNumber_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset

    dbs.TableDefs.Refresh
    dbs.TableDefs.Refresh
    For Each tdf In .TableDefs
      blnSkip = False
      With tdf
        If Left$(.Name, 4) <> "MSys" And Left$(.Name, 4) <> "USys" Then
          If .Connect = vbNullString Then
            strDbsName = Parse_File(CurrentDb.Name)  ' ** Module Function: modFileUtilities.
            strDbsPath = Parse_Path(CurrentDb.Name)  ' ** Module Function: modFileUtilities.
            lngDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strDbsName & "'")
            strTableName = .Name
          Else
            If Parse_File(.Connect) = "TAJrnTmp.mdb" And blnWithJrnlTmp = False Then  ' ** Module Function: modFileUtilities.
              blnSkip = True
            Else
              strDbsName = Parse_File(.Connect)  ' ** Module Function: modFileUtilities.
              strDbsPath = Parse_Path(Mid$(.Connect, (InStr(.Connect, LNK_IDENT) + Len(LNK_IDENT))))  ' ** Module Function: modFileUtilities.
              lngDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strDbsName & "'")
              strTableName = .SourceTableName
            End If
          End If
          If blnSkip = False Then
            lngTblID = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And " & _
              "[tbl_name] = '" & strTableName & "'")  ' ** Module Function: modFileUtilities.
            blnAdd = False
            For Each fld In .Fields
              With fld
                If (.Attributes And dbAutoIncrField) > 0 Then
                  lngFldID = DLookup("[fld_id]", "tblDatabase_Table_Field", "[dbs_id] = " & CStr(lngDbsID) & " And " & _
                    "[tbl_id] = " & CStr(lngTblID) & " And [fld_name] = '" & .Name & "'")  ' ** Module Function: modFileUtilities.
                  With rst
                    lngFldFnds = lngFldFnds + 1&
                    lngE = lngFldFnds - 1&
                    ReDim Preserve arr_varFldFnd(FF_ELEMS, lngE)
                    arr_varFldFnd(FF_AID, lngE) = CLng(0)
                    arr_varFldFnd(FF_DID, lngE) = lngDbsID
                    arr_varFldFnd(FF_DNAM, lngE) = strDbsName
                    arr_varFldFnd(FF_PATH, lngE) = strDbsPath
                    arr_varFldFnd(FF_TID, lngE) = lngTblID
                    arr_varFldFnd(FF_TNAM, lngE) = tdf.Name
                    arr_varFldFnd(FF_FID, lngE) = lngFldID
                    arr_varFldFnd(FF_FNAM, lngE) = fld.Name
                    arr_varFldFnd(FF_SEED, lngE) = CLng(-1)
                    .FindFirst "[dbs_id] = " & CStr(lngDbsID) & " And [tbl_id] = " & CStr(lngTblID)
                    If .NoMatch = True Then
                      blnAdd = True
                      .AddNew
                    Else
                      .Edit
                    End If
                    If blnAdd = True Then
                      ![dbs_id] = lngDbsID
                      ![tbl_id] = lngTblID
                      ![autonum_lastid] = 0&
                      ![autonum_seed] = 0&
                      arr_varFldFnd(FF_LID, lngE) = 0&
                      arr_varFldFnd(FF_SEED, lngE) = 0&
                    Else
                      arr_varFldFnd(FF_AID, lngE) = ![autonum_id]
                      arr_varFldFnd(FF_LID, lngE) = ![autonum_lastid]
                      arr_varFldFnd(FF_SEED, lngE) = ![autonum_seed]
                    End If
                    ![fld_id] = lngFldID
                    ![autonum_datemodified] = Now()
                    .Update
                    If blnAdd = True Then
                      .Bookmark = .LastModified
                      arr_varFldFnd(FF_AID, lngE) = ![autonum_id]
                    End If
                  End With
                  'Debug.Print "'" & tdf.Name & ": " & .Name & " dbAutoIncrField"
                End If
              End With
            Next
          End If  ' ** blnSkip.
        End If
      End With
    Next

    rst.Close

    .Close
  End With

  If lngFldFnds > 0& Then

    ' ** Delete obsolete entries.
    Set dbs = CurrentDb
    With dbs

      lngDels = 0&
      ReDim arr_varDel(0, 0)

      ' ** tblDatabase_AutoNumber, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
      Set qdf = .QueryDefs("zz_qry_Database_AutoNumber_01")
      With qdf.Parameters
        ![wjrnl] = blnWithJrnlTmp
      End With
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          blnFound = False
          For lngY = 0& To (lngFldFnds - 1&)
            If arr_varFldFnd(FF_DID, lngY) = ![dbs_id] And _
                arr_varFldFnd(FF_TID, lngY) = ![tbl_id] And _
                arr_varFldFnd(FF_FID, lngY) = ![fld_id] Then
              blnFound = True
              Exit For
            End If
          Next
          If blnFound = False Then
            lngDels = lngDels + 1&
            lngE = lngDels - 1&
            ReDim Preserve arr_varDel(0, lngE)
            arr_varDel(0, lngE) = ![autonum_id]
          End If
          If lngX < lngRecs Then .MoveNext
        Next
      End With

      If lngDels > 0& Then
        For lngX = 0& To (lngDels - 1&)
          ' ** Delete tblDatabase_AutoNumber, by specified [autid].
          Set qdf = .QueryDefs("zz_qry_Database_AutoNumber_14")
          With qdf.Parameters
            ![autid] = arr_varDel(0, lngX)
          End With
          qdf.Execute
        Next
      End If

      ' ** Get their current last ID's.
      For lngX = 0& To (lngFldFnds - 1&)
        varTmp00 = DMax("[" & arr_varFldFnd(FF_FNAM, lngX) & "]", arr_varFldFnd(FF_TNAM, lngX))
        If IsNull(varTmp00) = False Then
          arr_varFldFnd(FF_LID, lngX) = varTmp00
        Else
          arr_varFldFnd(FF_LID, lngX) = 0&
        End If
        ' ** Update autonum_lastid in tblDatabase_AutoNumber, by specified [autid], [lstid].
        Set qdf = .QueryDefs("zz_qry_Database_AutoNumber_03")
        With qdf.Parameters
          ![autid] = arr_varFldFnd(FF_AID, lngX)
          ![lstid] = arr_varFldFnd(FF_LID, lngX)
        End With
        qdf.Execute
      Next

      ' ** Update zz_qry_Database_AutoNumber_08 (m_TBL, with DLookups() to zz_qry_Database_AutoNumber_07
      ' ** (zz_qry_Database_AutoNumber_06 (m_TBL, linked to zz_qry_Database_AutoNumber_05 (tblDatabase_AutoNumber,
      ' ** with dbs_name, tbl_name, fld_name), with Ax), just discrepancies), with mtbl_AUTONUMBER_new).
      Set qdf = .QueryDefs("zz_qry_Database_AutoNumber_09")
      qdf.Execute

      ' ** Update zz_qry_Database_AutoNumber_12 (tblTemplate_m_TBL, with DLookups() to zz_qry_Database_AutoNumber_11
      ' ** (zz_qry_Database_AutoNumber_10 (tblTemplate_m_TBL, linked to zz_qry_Database_AutoNumber_05 (tblDatabase_AutoNumber,
      ' ** with dbs_name, tbl_name, fld_name), with Ax), just discrepancies), with mtbl_AUTONUMBER_new).
      Set qdf = .QueryDefs("zz_qry_Database_AutoNumber_13")
     qdf.Execute

      .Close
    End With

    ' ** Now get their current Seed ID's.
    If blnGetSeed = True Then
      Debug.Print "'FLDS: " & CStr(lngFldFnds) & "  ";
      DoEvents
      For lngX = 0& To (lngFldFnds - 1&)
        blnFound = False
        DoCmd.Hourglass True
        If Parse_File(CurrentDb.Name) = arr_varFldFnd(FF_DNAM, lngX) Then blnFound = True  ' ** Module Function: modFileUtilities.
        lngLastSeed = ChangeSeed(CStr(arr_varFldFnd(FF_TNAM, lngX)), _
          CStr(arr_varFldFnd(FF_FNAM, lngX)), 0&, blnFound, True, _
          (arr_varFldFnd(FF_PATH, lngX) & LNK_SEP & arr_varFldFnd(FF_DNAM, lngX)))  ' ** Module Function: modAutonumberFieldFuncs.
        DoCmd.Hourglass True
        Set dbs = CurrentDb
        With dbs
          ' ** Update autonum_seed in tblDatabase_AutoNumber, by specified [autid], [lstsed].
          Set qdf = .QueryDefs("zz_qry_Database_AutoNumber_02")
          With qdf.Parameters
            ![autid] = arr_varFldFnd(FF_AID, lngX)
            ![lstsed] = lngLastSeed
          End With
          qdf.Execute
          arr_varFldFnd(FF_SEED, lngX) = lngLastSeed
          .Close
        End With
        DoEvents
        If (lngX + 1&) Mod 10 = 0 Then
          Debug.Print "|";
        Else
          Debug.Print ".";
        End If
        DoEvents
      Next
    End If

  End If

  DoEvents
  Qry_TmpTables False, blnWithJrnlTmp  ' ** Function: Below.

  DoCmd.Hourglass False
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set fld = Nothing
  Set tdf = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Tbl_AutoNum_Doc = blnRetValx

End Function

Public Function Tbl_Link_DocConn() As Boolean
' ** Document all linked-table Connect strings to tblDatabase_Table_Link.

  Const THIS_PROC As String = "Tbl_Link_DocConn"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim lngDbsID As Long, strDbsName As String, lngTblID As Long, strTblName As String
  Dim varTmp00 As Variant, strTmp01 As String

  blnRetValx = True
  DoCmd.Hourglass True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Set dbs = CurrentDb
  With dbs
    Set rst1 = .OpenRecordset("tblDatabase_Table_Link", dbOpenDynaset, dbConsistent)
    Set rst2 = .OpenRecordset("tblTemplate_Database_Table_Link", dbOpenDynaset, dbConsistent)
    For Each tdf In .TableDefs
      With tdf
        If Left$(.Name, 4) <> "~TMP" Then  ' ** Skip those pesky system tables!
          If .Connect <> vbNullString Then
If .Name = "tblForm1" Or .Name = "tblForm_Graphics1" Or .Name = "tblXAdmin_Graphics1" Or .Name = "tblXAdmin_Graphics_Format1" Or .Name = "tblXAdmin_Graphics_Group1" Then
  ' ** These are temporary.
Else
            strTmp01 = .Connect
            strDbsName = Parse_File(strTmp01)  ' ** Module Function: modFileUtilities
            varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strDbsName & "'")
            If IsNull(varTmp00) = False Then
              lngDbsID = CLng(varTmp00)
              strTblName = .Name
              If strTblName = "LedgerArchive" Then strTblName = "ledger"
              If strTblName = "tblDataTypeDb1" Then strTblName = "tblDataTypeDb"
              ' ** tblSecurity_Group_TI, tblSecurity_GroupUser_TI, tblSecurity_License_TI, tblSecurity_User_TI
              If Right$(strTblName, 3) = "_TI" Then strTblName = Left$(strTblName, (Len(strTblName) - 3))
              varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And [tbl_name] = '" & strTblName & "'")
              If IsNull(varTmp00) = False Then
                lngTblID = CLng(varTmp00)
                With rst1
                  .FindFirst "[dbs_id] = " & CStr(lngDbsID) & " And [tbl_id] = " & CStr(lngTblID)
                  If .NoMatch = False Then
                    If IsNull(![tbllnk_connect]) = True Then
                      .Edit
                      ![tbllnk_connect] = strTmp01
                      ![tbllnk_datemodified] = Now()
                      .Update
                    Else
                      If ![tbllnk_connect] <> strTmp01 Then
                        .Edit
                        ![tbllnk_connect] = strTmp01
                        ![tbllnk_datemodified] = Now()
                        .Update
                      End If
                    End If
                  Else
                    Debug.Print "'NOT FOUND IN rst1: " & tdf.Name
                  End If
                End With
                With rst2
                  .FindFirst "[dbs_id] = " & CStr(lngDbsID) & " And [tbl_id] = " & CStr(lngTblID)
                  If .NoMatch = False Then
                    If IsNull(![tbllnk_connect]) = True Then
                      .Edit
                      ![tbllnk_connect] = strTmp01
                      ![tbllnk_datemodified] = Now()
                      .Update
                    Else
                      If ![tbllnk_connect] <> strTmp01 Then
                        .Edit
                        ![tbllnk_connect] = strTmp01
                        ![tbllnk_datemodified] = Now()
                        .Update
                      End If
                    End If
                  Else
                    Debug.Print "'NOT FOUND IN rst2: " & tdf.Name
                  End If
                End With
              Else
                Beep
                Stop
              End If
            Else
              Beep
              Stop
            End If
End If
          End If
        End If
      End With
    Next
    rst1.Close
    rst2.Close
    .Close
  End With

  Set tdf = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set dbs = Nothing

  DoCmd.Hourglass False
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Tbl_Link_DocConn = blnRetValx

End Function

Public Function Tbl_Link_Doc() As Boolean
' ** Document all currently linked tables to tblDatabase_Table_Link.
' #######################################################################
' ## 4 TABLES NEED TO BE KEPT UP-TO-DATE AND IN-SYNC, ALL HANDLED HERE:
' ##   m_TBL
' ##   tblTemplate_m_TBL
' ##   tblDatabase_Table_Link
' ##   tblTemplate_Database_Table_Link
' #######################################################################

  Const THIS_PROC As String = "Tbl_Link_Doc"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strWhoRules As String
  Dim strDatabaseName As String, strTableName As String, strSourceTableName As String
  Dim lngDbsID As Long, lngTblID As Long
  Dim lngDtaID As Long, lngArchID As Long, lngAuxID As Long
  Dim lngLinks As Long, arr_varLink() As Variant
  Dim blnAdd As Boolean, blnAddAll As Boolean, blnFound As Boolean
  Dim blnWrongDB As Boolean, strCorrectDB As String
  Dim lngRecs As Long
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngEdits As Long, arr_varEdit() As Variant
  Dim lngAvails As Long, arr_varAvail() As Variant
  Dim blnSeedReset As Boolean, blnListOnly As Boolean
  Dim varTmp00 As Variant, lngTmp01 As Long, lngTmp02 As Long, lngTmp03 As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varLink().
  Const L_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const L_DBSID       As Integer = 0
  Const L_DBSNAM      As Integer = 1
  Const L_TBLID       As Integer = 2
  Const L_TBLNAM      As Integer = 3
  Const L_FND_MTBL    As Integer = 4  ' ** Found in m_TBL.
  Const L_ERR_MTBL    As Integer = 5  ' ** Found in m_TBL, but listed with wrong database.
  Const L_FND_T_MTBL  As Integer = 6  ' ** Found in tblTemplate_m_TBL.
  Const L_ERR_T_MTBL  As Integer = 7  ' ** Found in tblTemplate_m_TBL, but listed with wrong database.
  Const L_FND_DTLNK   As Integer = 8  ' ** Found in tblDatabase_Table_Link.
  Const L_FND_T_DTLNK As Integer = 9  ' ** Found in tblTemplate_Database_Table_Link

  ' ** Array: arr_varDel().
  Const D_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const D_DBSID       As Integer = 0
  Const D_DBSNAM      As Integer = 1
  Const D_TBLID       As Integer = 2
  Const D_TBLNAM      As Integer = 3
  Const D_DEL_MTBL    As Integer = 4  ' ** Delete from m_TBL.
  Const D_DEL_T_MTBL  As Integer = 5  ' ** Delete from tblTemplate_m_TBL.
  Const D_DEL_DTLNK   As Integer = 6  ' ** Delete from tblDatabase_Table_Link.
  Const D_DEL_T_DTLNK As Integer = 7  ' ** Delete from tblTemplate_Database_Table_Link
  Const D_DID          As Integer = 8
  Const D_ORD         As Integer = 9

  ' ** Array: arr_varEdit().
  Const E_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const E_TBLNAM As Integer = 0
  Const E_MTBL   As Integer = 1
  Const E_T_MTBL As Integer = 2
  Const E_MTBLID As Integer = 3
  Const E_DBSNAM As Integer = 4

  ' ** Array: arr_varAvail().
  Const A_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const A_ID As Integer = 0
  Const A_ORD As Integer = 1
  Const A_USED As Integer = 2

  blnRetValx = True

  strWhoRules = "CurrentLinks"  '"CurrentLinks", "tblDatabase_Table_Link", "tblTemplate_Database_Table_Link", "m_TBL", "tblTemplate_m_TBL"
  blnListOnly = False

  Set dbs = CurrentDb
  With dbs

    lngDtaID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TrustDta.mdb'")
    lngArchID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TrstArch.mdb'")
    lngAuxID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TrustAux.mdb'")

Select Case strWhoRules
Case "CurrentLinks"
'NOT FOUND: tblDocument
'NOT FOUND: tblDocument_Image
'NOT FOUND: tblDocumentAutoShapeType
'NOT FOUND: tblDocumentFieldKind
'NOT FOUND: tblDocumentFieldType
'NOT FOUND: tblDocumentInlineShapeType
'NOT FOUND: tblDocumentKind
'NOT FOUND: tblDocumentLinkType
'NOT FOUND: tblDocumentShapeType
'NOT FOUND: tblDocumentType
'NOT FOUND: tblDocumentWrapType
'NOT FOUND: tblImage
'NOT FOUND: tblParameterDirection
'NOT FOUND: tblPictureAlignmentType
'NOT FOUND: tblPictureSizeMode
'NOT FOUND: tblPictureType

    lngLinks = 0&
    ReDim arr_varLink(L_ELEMS, 0)

    For Each tdf In .TableDefs
      With tdf
        If .Connect <> vbNullString Then
          strDatabaseName = Parse_File(.Connect)  ' ** Module Function: modFileUtilities.
          lngDbsID = IIf(strDatabaseName = "TrustDta.mdb", lngDtaID, _
            IIf(strDatabaseName = "TrstArch.mdb", lngArchID, IIf(strDatabaseName = "TrustAux.mdb", lngAuxID, 0&)))
          strTableName = .Name
          strSourceTableName = .SourceTableName
          lngLinks = lngLinks + 1&
          lngE = lngLinks - 1&
          ReDim Preserve arr_varLink(L_ELEMS, lngE)
          If strTableName = "LedgerArchive" Then
            varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And [tbL_name] = '" & strSourceTableName & "'")
          Else
            varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And [tbL_name] = '" & strTableName & "'")
          End If
          If IsNull(varTmp00) = False Then
            lngTblID = CLng(varTmp00)
            arr_varLink(L_DBSID, lngE) = lngDbsID
            arr_varLink(L_DBSNAM, lngE) = strDatabaseName
            arr_varLink(L_TBLID, lngE) = lngTblID
            arr_varLink(L_TBLNAM, lngE) = strTableName
            arr_varLink(L_FND_MTBL, lngE) = CBool(False)
            arr_varLink(L_ERR_MTBL, lngE) = CBool(False)
            arr_varLink(L_FND_T_MTBL, lngE) = CBool(False)
            arr_varLink(L_ERR_T_MTBL, lngE) = CBool(False)
            arr_varLink(L_FND_DTLNK, lngE) = CBool(False)
            arr_varLink(L_FND_T_DTLNK, lngE) = CBool(False)
          Else
            Beep
            Stop
          End If
        End If
      End With
    Next

    lngDels = 0&
    ReDim arr_varDel(D_ELEMS, 0)

    lngEdits = 0&
    ReDim arr_varEdit(E_ELEMS, 0)

    Set rst = .OpenRecordset("m_TBL", dbOpenDynaset, dbConsistent)
    With rst
      blnAddAll = False
      If .BOF = True And .EOF = True Then
        ' ** Not likely!
        blnAddAll = True
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          strTableName = ![mtbl_NAME]
          blnFound = False: blnWrongDB = False: strCorrectDB = vbNullString
          For lngY = 0& To (lngLinks - 1&)
            If arr_varLink(L_TBLNAM, lngY) = strTableName Then
              If arr_varLink(L_DBSNAM, lngY) = "TrustDta.mdb" And ![mtbl_DTA] = True Then
                blnFound = True
              ElseIf arr_varLink(L_DBSNAM, lngY) = "TrstArch.mdb" And ![mtbl_ARCH] = True Then
                blnFound = True
              ElseIf arr_varLink(L_DBSNAM, lngY) = "TrustAux.mdb" And ![mtbl_AUX] = True Then
                blnFound = True
              Else
                ' ** Found, but listed in the wrong database.
                blnFound = True
                blnWrongDB = True
                strCorrectDB = arr_varLink(L_DBSNAM, lngY)
              End If
            End If
            If blnFound = True And blnWrongDB = False Then
              arr_varLink(L_FND_MTBL, lngY) = CBool(True)
              arr_varLink(L_ERR_MTBL, lngY) = CBool(False)
              Exit For
            ElseIf blnFound = True And blnWrongDB = True Then
              arr_varLink(L_FND_MTBL, lngY) = CBool(True)
              arr_varLink(L_ERR_MTBL, lngY) = CBool(True)
              Exit For
            End If
          Next
          If blnFound = False Then
            lngDels = lngDels + 1&
            lngE = lngDels - 1&
            ReDim Preserve arr_varDel(D_ELEMS, lngE)
            If ![mtbl_DTA] = True Then
              arr_varDel(D_DBSID, lngE) = lngDtaID
              arr_varDel(D_DBSNAM, lngE) = "TrustDta.mdb"
            ElseIf ![mtbl_ARCH] = True Then
              arr_varDel(D_DBSID, lngE) = lngArchID
              arr_varDel(D_DBSNAM, lngE) = "TrstArch.mdb"
            ElseIf ![mtbl_AUX] = True Then
              arr_varDel(D_DBSID, lngE) = lngAuxID
              arr_varDel(D_DBSNAM, lngE) = "TrustAux.mdb"
            End If
            arr_varDel(D_TBLID, lngE) = CLng(0)
            arr_varDel(D_TBLNAM, lngE) = strTableName
            arr_varDel(D_DEL_MTBL, lngE) = CBool(True)
            arr_varDel(D_DEL_T_MTBL, lngE) = CBool(False)
            arr_varDel(D_DEL_DTLNK, lngE) = CBool(False)
            arr_varDel(D_DEL_T_DTLNK, lngE) = CBool(False)
            arr_varDel(D_DID, lngE) = ![mtbl_ID]
            arr_varDel(D_ORD, lngE) = ![mtbl_ORDER]
          ElseIf blnWrongDB = True Then
            lngEdits = lngEdits + 1&
            lngE = lngEdits - 1&
            ReDim Preserve arr_varEdit(E_ELEMS, lngE)
            arr_varEdit(E_TBLNAM, lngE) = strTableName
            arr_varEdit(E_MTBL, lngE) = CBool(True)
            arr_varEdit(E_T_MTBL, lngE) = CBool(False)
            arr_varEdit(E_MTBLID, lngE) = ![mtbl_ID]
            arr_varEdit(E_DBSNAM, lngE) = strCorrectDB
          End If
          If lngX < lngRecs Then .MoveNext
        Next  ' ** m_TBL: lngX.
      End If
      .Close
    End With
    If blnAddAll = True Then
      ' ** All the L_FND's will be false.
    End If

    Set rst = .OpenRecordset("tblDatabase_Table_Link", dbOpenDynaset, dbConsistent)
    With rst
      blnAddAll = False
      If .BOF = True And .EOF = True Then
        ' ** Not likely!
        blnAddAll = True
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          lngDbsID = ![dbs_id]
          strDatabaseName = IIf(lngDbsID = lngDtaID, "TrustDta.mdb", _
            IIf(lngDbsID = lngArchID, "TrstArch.mdb", IIf(lngDbsID = lngAuxID, "TrustAux.mdb", vbNullString)))
          lngTblID = ![tbl_id]
          strTableName = DLookup("[tbl_name]", "tblDatabase_Table", "[tbl_id] = " & CStr(lngTblID))
          blnFound = False
          For lngY = 0& To (lngLinks - 1&)
            If arr_varLink(L_TBLID, lngY) = lngTblID And arr_varLink(L_DBSID, lngY) = lngDbsID Then
              blnFound = True
              arr_varLink(L_FND_DTLNK, lngY) = CBool(True)
              Exit For
            End If
          Next
          If blnFound = False Then
            For lngY = 0& To (lngDels - 1&)
              If arr_varDel(D_DBSID, lngY) = lngDbsID And arr_varDel(D_TBLID, lngY) = lngTblID Then
                blnFound = True
                arr_varDel(D_DEL_DTLNK, lngY) = CBool(True)
                Exit For
              End If
            Next
            If blnFound = False Then
              lngDels = lngDels + 1&
              lngE = lngDels - 1&
              ReDim Preserve arr_varDel(D_ELEMS, lngE)
              arr_varDel(D_DBSID, lngE) = lngDbsID
              arr_varDel(D_DBSNAM, lngE) = strDatabaseName
              arr_varDel(D_TBLID, lngE) = lngTblID
              arr_varDel(D_TBLNAM, lngE) = strTableName
              arr_varDel(D_DEL_MTBL, lngE) = CBool(False)
              arr_varDel(D_DEL_T_MTBL, lngE) = CBool(False)
              arr_varDel(D_DEL_DTLNK, lngE) = CBool(True)
              arr_varDel(D_DEL_T_DTLNK, lngE) = CBool(False)
              arr_varDel(D_DID, lngE) = ![tbllnk_id]
              arr_varDel(D_ORD, lngE) = CLng(0)
            End If
          End If
          If lngX < lngRecs Then .MoveNext
        Next
      End If
      .Close
    End With
    If blnAddAll = True Then
      ' ** All the L_FND's will be false.
    End If

    If blnListOnly = True Then
      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      ' ** Summarize what I've got so far...
      If lngDels > 0& Then
        lngTmp01 = 0&
        For lngX = 0& To (lngDels - 1&)
          If arr_varDel(D_DEL_MTBL, lngX) = True Then
            Debug.Print "'DEL FROM m_TBL: " & arr_varDel(D_TBLNAM, lngX)
            lngTmp01 = lngTmp01 + 1&
          End If
        Next
        If lngTmp01 > 0& Then
          Debug.Print ""
        End If
        lngTmp02 = 0&
        For lngX = 0& To (lngDels - 1&)
          If arr_varDel(D_DEL_DTLNK, lngX) = True Then
            Debug.Print "'DEL FROM tblDatabase_Table_Link: " & arr_varDel(D_TBLNAM, lngX)
            lngTmp02 = lngTmp02 + 1&
          End If
        Next
        If lngTmp02 > 0& Then
          Debug.Print ""
        End If
      End If
      lngTmp03 = 0&
      For lngX = 0& To (lngLinks - 1&)
        If arr_varLink(L_FND_MTBL, lngX) = False Then
          Debug.Print "'ADD TO m_TBL: " & arr_varLink(L_TBLNAM, lngX)
          lngTmp03 = lngTmp03 + 1&
        End If
      Next
      If lngTmp03 > 0& Then
        Debug.Print ""
      End If
      For lngX = 0& To (lngLinks - 1&)
        If arr_varLink(L_FND_DTLNK, lngX) = False Then
          Debug.Print "'ADD TO tblDatabase_Table_Link: " & arr_varLink(L_TBLNAM, lngX)
        End If
      Next
    End If

    If blnListOnly = False Then

      lngAvails = 0&
      ReDim arr_varAvail(A_ELEMS, 0)

      ' ** Delete from m_TBL.
      For lngX = 0& To (lngDels - 1&)
        If arr_varDel(D_DEL_MTBL, lngX) = True Then
          lngAvails = lngAvails + 1&
          lngE = lngAvails - 1&
          ReDim Preserve arr_varAvail(A_ELEMS, lngE)
          arr_varAvail(A_ID, lngE) = arr_varDel(D_DID, lngX)
          arr_varAvail(A_ORD, lngE) = arr_varDel(D_ORD, lngX)
          arr_varAvail(A_USED, lngE) = CBool(False)
          ' ** Delete m_TBL, by specified [mtblid].
          Set qdf = .QueryDefs("qrySystemUpdate_11i_m_TBL")
          With qdf.Parameters
            ![mtblid] = arr_varDel(D_DID, lngX)
          End With
          qdf.Execute dbFailOnError
          Debug.Print "'DELETED mtbl_ID: " & arr_varDel(D_DID, lngX) & "  TBL: " & arr_varDel(D_TBLNAM, lngX)
        End If
      Next

      ' ** Edit m_Tbl.
      For lngX = 0& To (lngEdits - 1&)
        ' ** Update m_TBL, by specified [mtblid], [isdta], [isarch], [isaux].
        Set qdf = .QueryDefs("qrySystemUpdate_11o_m_TBL")
        With qdf.Parameters
          ![mtblid] = arr_varEdit(E_MTBLID, lngX)
          Select Case arr_varEdit(E_DBSNAM, lngX)
          Case "TrustDta.mdb"
            ![isdta] = CBool(True)
            ![IsArch] = CBool(False)
            ![isaux] = CBool(False)
          Case "TrstArch.mdb"
            ![isdta] = CBool(False)
            ![IsArch] = CBool(True)
            ![isaux] = CBool(False)
          Case "TrustAux.mdb"
            ![isdta] = CBool(False)
            ![IsArch] = CBool(False)
            ![isaux] = CBool(True)
          End Select
        End With
        qdf.Execute
      Next

      blnSeedReset = False  ' ** I did it manually.
      For lngX = 0& To (lngLinks - 1&)
        If arr_varLink(L_FND_MTBL, lngX) = False Then
          blnFound = False
          For lngY = 0& To (lngAvails - 1&)
            If arr_varAvail(A_USED, lngY) = False Then
              ' ** Append new record to m_TBL, using missing AutoNumber, by specified
              ' ** [mtblid], [tblnam], [autonum], [ord], [newrecs], [isdta], [isarch], [isaux].
              Set qdf = .QueryDefs("qrySystemUpdate_11j_m_TBL")
              With qdf.Parameters
                ![mtblid] = arr_varAvail(A_ID, lngY)
                ![tblnam] = arr_varLink(L_TBLNAM, lngX)
                varTmp00 = DLookup("[fld_id]", "tblDatabase_AutoNumber", "[dbs_id] = " & CStr(arr_varLink(L_DBSID, lngX)) & " And " & _
                  "[tbl_id] = " & CStr(arr_varLink(L_TBLID, lngX)))
                If IsNull(varTmp00) = False Then
                  ![autonum] = DLookup("[fld_name]", "tblDatabase_Table_Field", "[fld_id] = " & CStr(varTmp00))
                Else
                  ![autonum] = Null  ' ** Designated 'Value', so I'll have to see if that accepts a Null.
                End If
                ![Ord] = arr_varAvail(A_ORD, lngY)
                ![newrecs] = False  ' ** I'll have to set these manually!
                Select Case arr_varLink(L_DBSID, lngX)
                Case lngDtaID
                  ![isdta] = True
                  ![IsArch] = False
                  ![isaux] = False
                Case lngArchID
                  ![isdta] = False
                  ![IsArch] = True
                  ![isaux] = False
                Case lngAuxID
                  ![isdta] = False
                  ![IsArch] = False
                  ![isaux] = True
                End Select
              End With
              qdf.Execute dbFailOnError
              arr_varAvail(A_USED, lngY) = CBool(True)
              blnFound = True
              Debug.Print "'ADDED TO m_TBL  TBL: " & arr_varLink(L_TBLNAM, lngX)
              Exit For
            End If
          Next
          If blnFound = False Then
            If blnSeedReset = False Then
              ' ** Reset the Autonumber field to the next available.
              ChangeSeed_Ext "m_TBL"  ' ** Module Function: modAutonumberFieldFuncs.
              blnSeedReset = True
              ' ** I THINK ChangeSeed_Ext() RESETS THE DBS VARIABLE!
              Set dbs = CurrentDb
            End If
            ' ** Append new record to m_TBL, by specified [tblnam], [autonum], [ord], [newrecs], [isdta], [isarch], [isaux].
            Set qdf = .QueryDefs("qrySystemUpdate_11k_m_TBL")
            With qdf.Parameters
              ![tblnam] = arr_varLink(L_TBLNAM, lngX)
              varTmp00 = DLookup("[fld_id]", "tblDatabase_AutoNumber", "[dbs_id] = " & CStr(arr_varLink(L_DBSID, lngX)) & " And " & _
                "[tbl_id] = " & CStr(arr_varLink(L_TBLID, lngX)))
              If IsNull(varTmp00) = False Then
                ![autonum] = DLookup("[fld_name]", "tblDatabase_Table_Field", "[fld_id] = " & CStr(varTmp00))
              Else
                ![autonum] = Null  ' ** Designated 'Value', so I'll have to see if that accepts a Null.
              End If
              ![Ord] = (DMax("[mtbl_ORDER]", "m_TBL") + 1&)
              ![newrecs] = False  ' ** I'll have to set these manually!
              Select Case arr_varLink(L_DBSID, lngX)
              Case lngDtaID
                ![isdta] = True
                ![IsArch] = False
                ![isaux] = False
              Case lngArchID
                ![isdta] = False
                ![IsArch] = True
                ![isaux] = False
              Case lngAuxID
                ![isdta] = False
                ![IsArch] = False
                ![isaux] = True
              End Select
            End With
            qdf.Execute dbFailOnError
            Debug.Print "'ADDED TO m_TBL  TBL: " & arr_varLink(L_TBLNAM, lngX)
          End If
        End If
      Next

      ' *************************************
      ' ** Sync tblTemplate_m_TBL to m_TBL.
      ' *************************************

      ' ** Delete qrySystemUpdate_11b_m_TBL (tblTemplate_m_TBL, not in m_TBL) from tblTemplate_m_TBL.
      ' ** ####  m_TBL Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_11c_m_TBL")
      qdf.Execute dbFailOnError

      ' ** Append qrySystemUpdate_11a_m_TBL (m_TBL, not in tblTemplate_m_TBL) to tblTemplate_m_TBL.
      ' ** ####  m_TBL Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_11d_m_TBL")
      qdf.Execute dbFailOnError

      ' ** Update tblTemplate_m_TBL from m_TBL.
      ' ** ####  m_TBL Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_11e_m_TBL")
      qdf.Execute dbFailOnError

      ' ***************************************************
      ' ** Sync tblTemplate_Database_Table_Link to m_TBL.
      ' ***************************************************

      ' ** Delete qrySystemUpdate_11m_m_TBL (tblTemplate_Database_Table_Link,
      ' ** not in tblTemplate_m_TBL) from tblTemplate_Database_Table_Link.
      ' ** ####  tblTemplate_m_TBL Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_11n_m_TBL")
      qdf.Execute dbFailOnError

      ' ** Append qrySystemUpdate_11oc_m_TBL (qrySystemUpdate_11ob_m_TBL (qrySystemUpdate_11l_m_TBL (tblTemplate_m_TBL,
      ' ** not in tblTemplate_Database_Table_Link), with missing AutoNumber function),
      ' ** with tblMark, with tbllnk_id_new; Cartesian) to tblTemplate_Database_Table_Link.
      ' ** ####  tblTemplate_m_TBL Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_11oa_m_TBL")
      qdf.Execute dbFailOnError

      ' ********************************************************************
      ' ** Sync tblDatabase_Table_Link to tblTemplate_Database_Table_Link.
      ' ********************************************************************

      ' ** Delete qrySystemUpdate_12a_tblDatabase_Table_Link (tblDatabase_Table_Link,
      ' ** not in tblTemplate_Database_Table_Link) from tblDatabase_Table_Link.
      ' ** ####  tblTemplate_Database_Table_Link Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_12f_tblDatabase_Table_Link")
      qdf.Execute dbFailOnError

      ' ** Append qrySystemUpdate_12b_tblDatabase_Table_Link (tblTemplate_Database_Table_Link,
      ' ** not in tblDatabase_Table_Link) to tblDatabase_Table_Link.
      ' ** ####  tblTemplate_Database_Table_Link Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_12g_tblDatabase_Table_Link")
      qdf.Execute dbFailOnError

      ' ** Update tblDatabase_Table_Link from tblTemplate_Database_Table_Link.
      ' ** ####  tblTemplate_Database_Table_Link Takes Precedence  ####
      Set qdf = .QueryDefs("qrySystemUpdate_12h_tblDatabase_Table_Link")
      qdf.Execute dbFailOnError

      ' ** Reset the Autonumber field to the next available.
      ChangeSeed_Ext "m_TBL"  ' ** Module Function: modAutonumberFieldFuncs.
      ChangeSeed_Ext "tblDatabase_Table_Link"  ' ** Module Function: modAutonumberFieldFuncs.

    End If

Case "tblDatabase_Table_Link"

    lngLinks = 0&
    ReDim arr_varLink(L_ELEMS, 0)

    Set rst = .OpenRecordset("tblDatabase_Table_Link", dbOpenDynaset, dbConsistent)

    For Each tdf In .TableDefs
      blnAdd = False
      lngDbsID = 0&: lngTblID = 0&
      With tdf
        If .Connect <> vbNullString Then
          strDatabaseName = Parse_File(.Connect)  ' ** Module Function: modFileUtilities.
          lngDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strDatabaseName & "'")
          strTableName = .Name                   ' ** This is the table's local name.
          strSourceTableName = .SourceTableName  ' ** This is the table's name in its source database.
          lngTblID = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And [tbL_TBLNAM] = '" & strSourceTableName & "'")
          lngLinks = lngLinks + 1&
          lngE = lngLinks - 1&
          ReDim Preserve arr_varLink(L_ELEMS, lngE)
          arr_varLink(L_DBSID, lngE) = lngDbsID
          arr_varLink(L_TBLID, lngE) = lngTblID
          arr_varLink(L_TBLNAM, lngE) = strTableName
          If rst.BOF = True And rst.EOF = True Then
            blnAdd = True
          Else
            rst.FindFirst "[tbllnk_name] = '" & strTableName & "'"
            If rst.NoMatch = True Then
              blnAdd = True
            End If
          End If
          With rst
            If blnAdd = True Then
'Append new rec to tblDatabase_Table_Link
'USING MISSING AUTONUMBER!!!!!!!!!!!!!!!
              .AddNew
              ![dbs_id] = lngDbsID
              ![tbl_id] = lngTblID
              ![tbllnk_name] = strTableName
              ![tbllnk_sourcetablename] = strSourceTableName
              ![tbllnk_datemodified] = Now()
              .Update
            Else
              If ![tbl_id] <> lngTblID Then
                .Edit
                ![tbl_id] = lngTblID
                ![tbllnk_datemodified] = Now()
                .Update
              End If
              If ![tbllnk_sourcetablename] <> strSourceTableName Then
                .Edit
                ![tbllnk_sourcetablename] = strSourceTableName
                ![tbllnk_datemodified] = Now()
                .Update
              End If
            End If
          End With
        End If
      End With
    Next

    rst.Close

    If lngLinks > 0& Then
      lngDels = 0&
      ReDim arr_varDel(0)
      Set rst = .OpenRecordset("tblDatabase_Table_Link", dbOpenDynaset, dbConsistent)
      With rst
        If .BOF = True And .EOF = True Then
          ' ** What are ya doin here then?
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            blnFound = False
            For lngY = 0& To (lngLinks - 1&)
              If arr_varLink(L_DBSID, lngY) = ![dbs_id] And arr_varLink(L_TBLID, lngY) = ![tbl_id] Then
                blnFound = True
                Exit For
              End If
            Next
            If blnFound = False Then
              lngDels = lngDels + 1&
              ReDim Preserve arr_varDel(lngDels - 1&)
              arr_varDel(lngDels - 1&) = ![tbllnk_id]
            End If
            If lngX < lngRecs Then .MoveNext
          Next
        End If
        .Close
      End With
      If lngDels > 0& Then
        For lngX = 0& To (lngDels - 1&)
          ' ** Delete tblDatabase_Table_Link, by specified [lnkid].
          Set qdf = .QueryDefs("qrySystemUpdate_12o_tblDatabase_Table_Link")
          With qdf.Parameters
            ![lnkid] = arr_varDel(lngX)
          End With
          qdf.Execute
        Next
      End If
    End If

    ' ** Delete qrySystemUpdate_12b_tblDatabase_Table_Link (tblTemplate_Database_Table_Link,
    ' ** not in tblDatabase_Table_Link) from tblTemplate_Database_Table_Link.
    ' ** ####  tblDatabase_Table_Link Takes Precedence  ####
    Set qdf = .QueryDefs("qrySystemUpdate_12c_tblDatabase_Table_Link")
    qdf.Execute dbFailOnError

    ' ** Append qrySystemUpdate_12a_tblDatabase_Table_Link (tblDatabase_Table_Link,
    ' ** not in tblTemplate_Database_Table_Link) to tblTemplate_Database_Table_Link.
    ' ** ####  tblDatabase_Table_Link Takes Precedence  ####
    Set qdf = .QueryDefs("qrySystemUpdate_12d_tblDatabase_Table_Link")
    qdf.Execute dbFailOnError

    ' ** Update tblTemplate_Database_Table_Link from tblDatabase_Table_Link.
    ' ** ####  tblDatabase_Table_Link Takes Precedence  ####
    Set qdf = .QueryDefs("qrySystemUpdate_12e_tblDatabase_Table_Link")
    qdf.Execute dbFailOnError

'SYNC tblTemplate_m_TBL to tblDatabase_Table_Link

    ' ** Delete qrySystemUpdate_11a_m_TBL (m_TBL, not in tblTemplate_m_TBL) from m_TBL.
    ' ** ####  tblTemplate_m_TBL TAKES PRECEDENCE  ####
    Set qdf = .QueryDefs("qrySystemUpdate_11f_m_TBL")
    qdf.Execute dbFailOnError

    ' ** Append qrySystemUpdate_11c_m_TBL (tblTemplate_m_TBL, not in m_TBL) to m_TBL.
    ' ** ####  tblTemplate_m_TBL TAKES PRECEDENCE  ####
    Set qdf = .QueryDefs("qrySystemUpdate_11g_m_TBL")
    qdf.Execute dbFailOnError

    ' ** Update m_TBL from tblTemplate_m_TBL.
    ' ** ####  tblTemplate_m_TBL TAKES PRECEDENCE  ####
    Set qdf = .QueryDefs("qrySystemUpdate_11h_m_TBL")
    qdf.Execute dbFailOnError

Case Else
  Stop
End Select

    .Close
  End With

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set tdf = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Tbl_Link_Doc = blnRetValx

End Function

Public Function Tbl_Link_List() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "Tbl_Link_List"

  Dim dbs As DAO.Database, tdf As DAO.TableDef
  Dim lngLinks As Long, arr_varLink() As Variant
  Dim lngDbs As Long, arr_varDb() As Variant
  Dim lngMaxWidth As Long
  Dim blnFound As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varLink().
  Const L_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const L_NAM As Integer = 0
  Const L_FIL As Integer = 1
  Const L_CON As Integer = 2

  ' ** Array: arr_varDb().
  Const D_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const D_FIL As Integer = 0
  Const D_CON As Integer = 1

  blnRetValx = True

  lngLinks = 0&
  ReDim arr_varLink(L_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    For Each tdf In .TableDefs
      With tdf
        If .Connect <> vbNullString Then
          lngLinks = lngLinks + 1&
          lngE = lngLinks - 1&
          ReDim Preserve arr_varLink(L_ELEMS, lngE)
          arr_varLink(L_NAM, lngE) = .Name
          arr_varLink(L_FIL, lngE) = Parse_File(.Connect)  ' ** Module Function: modFileUtilities.
          arr_varLink(L_CON, lngE) = .Connect
        End If
      End With
    Next
    .Close
  End With

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  lngMaxWidth = 0&
  For lngX = 0& To (lngLinks - 1&)
    If Len(arr_varLink(L_NAM, lngX)) > lngMaxWidth Then lngMaxWidth = Len(arr_varLink(L_NAM, lngX))
  Next

  lngDbs = 0&
  ReDim arr_varDb(D_ELEMS, 0)

  For lngX = 0& To (lngLinks - 1&)
    blnFound = False
    For lngY = 0& To (lngDbs - 1&)
      If arr_varDb(D_FIL, lngY) = arr_varLink(L_FIL, lngX) And _
          arr_varDb(D_CON, lngY) = arr_varLink(L_CON, lngX) Then
        blnFound = True
        Exit For
      End If
    Next
    If blnFound = False Then
      lngDbs = lngDbs + 1&
      lngE = lngDbs - 1&
      ReDim Preserve arr_varDb(D_ELEMS, lngE)
      arr_varDb(D_FIL, lngE) = arr_varLink(L_FIL, lngX)
      arr_varDb(D_CON, lngE) = arr_varLink(L_CON, lngX)
    End If
  Next

  Debug.Print "'TOT LINKED: " & CStr(lngLinks) & "  TOT DBS: " & CStr(lngDbs)

  For lngX = 0& To (lngDbs - 1&)
    For lngY = 0& To (lngLinks - 1&)
      If arr_varLink(L_FIL, lngY) = arr_varDb(D_FIL, lngX) And _
          arr_varLink(L_CON, lngY) = arr_varDb(D_CON, lngX) Then
        Debug.Print "'LNK TBL: " & Left$(arr_varLink(L_NAM, lngY) & Space(lngMaxWidth), lngMaxWidth) & " " & _
          arr_varLink(L_FIL, lngY) & "  " & _
          Mid$(arr_varLink(L_CON, lngY), (InStr(arr_varLink(L_CON, lngY), LNK_IDENT) + Len(LNK_IDENT)))
      End If
    Next
  Next

  Beep

EXITP:
  Set tdf = Nothing
  Set dbs = Nothing
  Tbl_Link_List = blnRetValx
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

Public Function Tbl_Link_Chk() As Boolean

  Const THIS_PROC As String = "Tbl_Link_Chk"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, tdf As DAO.TableDef
  Dim lngDuos As Long, arr_varDuo As Variant
  Dim strAuxPathFile As String, strPath As String
  Dim blnMisLinked As Boolean
  Dim lngRecs As Long
  Dim intPos1 As Integer
  Dim lngX As Long

  ' ** Array: arr_varDuo().
  Const D_DID  As Integer = 0
  Const D_DNAM As Integer = 1
  Const D_TID  As Integer = 2
  Const D_TNAM As Integer = 3

  blnRetValx = True
  blnMisLinked = False

  Set dbs = CurrentDb
  With dbs

    ' ** zz_qry_Database_Table_Link_03 (zz_qry_Database_Table_Link_02 (tblDatabase_Table, linked to
    ' ** zz_qry_Database_Table_Link_01 (tblDatabase_Table, just those in TrustDta.mdb), just those
    ' ** tables in TrustAux.mdb also in TrustDta.mdb), linked to tblTemplate_m_TBL), just needed fields.
    Set qdf = .QueryDefs("zz_qry_Database_Table_Link_04")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngDuos = .RecordCount
      .MoveFirst
      arr_varDuo = .GetRows(lngDuos)
      .Close
    End With

    strAuxPathFile = CurrentBackendPath  ' ** Module Function: modFileUtilities.
    strAuxPathFile = strAuxPathFile & LNK_SEP & gstrFile_AuxDataName

    For lngX = 0& To (lngDuos - 1&)
      Set tdf = .TableDefs(arr_varDuo(D_TNAM, lngX))
      With tdf
        If InStr(.Connect, arr_varDuo(D_DNAM, lngX)) = 0 Then
          ' ** It's linked to TrustDta.mdb, instead of TrustAux.mdb!
          blnMisLinked = True
          TableDelete CStr(arr_varDuo(D_TNAM, lngX))  ' ** Module Function: modFileUtilities.
          dbs.TableDefs.Refresh
          dbs.TableDefs.Refresh
          DoCmd.TransferDatabase acLink, "Microsoft Access", strAuxPathFile, acTable, arr_varDuo(D_TNAM, lngX), arr_varDuo(D_TNAM, lngX)
          Debug.Print "'RE-LINKED! " & arr_varDuo(D_TNAM, lngX)
        End If
      End With
    Next

    ' ** tblDatabase, all records, with or without TAJrnTmp.mdb, by specified [wjrnl].
    Set qdf = .QueryDefs("zz_qry_Database_01")
    With qdf.Parameters
      ![wjrnl] = blnWithJrnlTmp
    End With
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngDuos = .RecordCount
      .MoveFirst
      strPath = vbNullString
      For lngX = 1& To lngDuos
        If Left(![dbs_name], 6) = "Trust." Then
          strPath = ![dbs_path]
        ElseIf Left(![dbs_name], 8) = "TrstXAdm" Then
          If ![dbs_path] <> strPath Then
            Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
            Beep
            Debug.Print "'PATH WRONG! " & ![dbs_name]
            DoEvents
            If strPath = vbNullString Then
              Stop
            Else
              .Edit
              ![dbs_path] = strPath
              ![dbs_datemodified] = Now()
              .Update
            End If
          End If
        ElseIf Left(![dbs_name], 11) = "TrustImport" Then
          If Right(![dbs_path], 11) <> "TrustImport" And Right(![dbs_path], 12) <> "Trust Import" Then
            Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
            Beep
            Debug.Print "'PATH WRONG! " & ![dbs_name]
            DoEvents
            If strPath = vbNullString Then
              Stop
            Else
              intPos1 = InStr(strPath, "Trust")
              If InStr(strPath, "TrustAccountant") > 0 Or InStr(strPath, "Trust Accountant") > 0 Then
                strPath = Left$(strPath, (intPos1 - 1)) & "TrustImport"
              Else
                Stop
              End If
              .Edit
              ![dbs_path] = strPath
              ![dbs_datemodified] = Now()
              .Update
            End If
          End If
        End If
        If lngX < lngDuos Then .MoveNext
      Next
      .Close
    End With

    .Close
  End With

  Set tdf = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  If blnMisLinked = True Then
    Beep
    Beep
  End If

  Tbl_Link_Chk = blnRetValx

End Function

Public Function Tbl_Link_Add() As Boolean
' ** Add a new table to the linked list.
' ** 1. Run the Database Docs to put the table in tblDatabase_Table, etc.
' ** 2. Link the new table to here.
' ** 3. Then run this.

  Const THIS_PROC As String = "Tbl_Link_Add"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, qdf3 As DAO.QueryDef, qdf4 As DAO.QueryDef
  Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset, rst4 As DAO.Recordset
  Dim strNewTable As String, strNewSourceTable As String, strNewConnect As String, strNewAutoNum As String
  Dim strNewDbs As String, strNewDbsPath As String
  Dim gstrFile_IApp As String, gstrFile_Admin As String
  Dim lngDbsID As Long, lngTblID As Long, lngFldID As Long
  Dim lngTblOrd As Long, lngMTblID As Long, lngTblLnkID As Long
  Dim lngThisDbsID As Long
  Dim lngRecs1 As Long, lngRecs2 As Long, lngRecs3 As Long, lngRecs4 As Long
  Dim blnFound As Boolean, blnNewRecs As Boolean
  Dim varTmp00 As Variant

  blnRetValx = True

  strNewTable = "tblQuery_Documentation"

  blnNewRecs = False

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.

    Set dbs = CurrentDb
    With dbs

      blnFound = False: lngDbsID = 0&
      strNewSourceTable = vbNullString: strNewConnect = vbNullString: strNewDbs = vbNullString
      For Each tdf In .TableDefs
        With tdf
          If .Name = strNewTable Then
            If .Connect <> vbNullString Then
              strNewSourceTable = .SourceTableName
              strNewConnect = .Connect
              strNewDbs = Parse_File(strNewConnect)  ' ** Module Function: modFileUtilities.
              If strNewDbs = gstrFile_DataName Or strNewDbs = gstrFile_ArchDataName Or strNewDbs = gstrFile_AuxDataName Or _
                  strNewDbs = (gstrFile_App & "." & gstrExt_AppDev) Or strNewDbs = (gstrFile_App & "." & gstrExt_AppRun) Or _
                  strNewDbs = (gstrFile_IApp & "." & gstrExt_AppDev) Or strNewDbs = (gstrFile_IApp & "." & gstrExt_AppRun) Or _
                  strNewDbs = (gstrFile_Admin & "." & gstrExt_AppDev) Or strNewDbs = (gstrFile_Admin & "." & gstrExt_AppRun) Then
                varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strNewDbs & "'")
                If IsNull(varTmp00) = False Then
                  blnFound = True
                  lngDbsID = CLng(varTmp00)
                Else
                  blnRetValx = False
                  Beep
                  Debug.Print "'DBS NOT FOUND!"
                End If
              Else
                blnRetValx = False
                Beep
                Debug.Print "'DBS NOT ONE OF RECOGNIZED 3!"
              End If
            Else
              blnRetValx = False
              Beep
              Debug.Print "'TBL NOT A LINKED TBL!"
            End If
            Exit For
          End If
        End With
      Next

      .Close
    End With  ' ** dbs.

    If blnFound = True Then
      blnFound = False: lngTblID = 0&
      If strNewTable = strNewSourceTable Then
        varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And [tbl_name] = '" & strNewTable & "'")
      Else
        varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And [tbl_name] = '" & strNewSourceTable & "'")
      End If
      If IsNull(varTmp00) = False Then
        blnFound = True
        lngTblID = CLng(varTmp00)
      Else
        Debug.Print "'RUNNING Tbl_Doc()..."
        DoEvents
        blnRetValx = Tbl_Doc  ' ** Function: Above.
        Set dbs = CurrentDb
        With dbs
          .TableDefs.Refresh
          .TableDefs.Refresh
          DoEvents
          .Close
        End With  ' ** dbs.
        Set dbs = Nothing
        If blnRetValx = True Then
          varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[tbl_name] = '" & strNewTable & "'")
          If IsNull(varTmp00) = False Then
            blnFound = True
            lngTblID = CLng(varTmp00)
          End If
        End If
      End If
    ElseIf blnFound = False And blnRetValx = True Then
      varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(lngDbsID) & " And [tbl_name] = '" & strNewTable & "'")
      If IsNull(varTmp00) = False Then
        ' ** Documented, but not yet linked.
        Set dbs = CurrentDb
        With dbs
          blnFound = True
          lngTblID = CLng(varTmp00)
          varTmp00 = DLookup("[dbs_id]", "tblDatabase_Table", "[tbl_id] = " & CStr(lngTblID))
          lngDbsID = CLng(varTmp00)
          varTmp00 = DLookup("[dbs_name]", "tblDatabase", "[dbs_id] = " & CStr(lngDbsID))
          strNewDbs = CStr(varTmp00)
          varTmp00 = DLookup("[dbs_path]", "tblDatabase", "[dbs_id] = " & CStr(lngDbsID))
          strNewDbsPath = CStr(varTmp00)
'CHECK THIS OUT!
          'strNewSourceTable = strNewTable
          DoCmd.TransferDatabase acLink, "Microsoft Access", (strNewDbsPath & LNK_SEP & strNewDbs), _
            acTable, strNewSourceTable, strNewTable
          .TableDefs.Refresh
          .TableDefs.Refresh
          strNewConnect = .TableDefs(strNewTable).Connect
          DoEvents
          .Close
        End With  ' ** dbs.
        Set dbs = Nothing
      Else
        blnRetValx = False
        Beep
        Debug.Print "'TBL NOT FOUND, PERIOD!!"
      End If
    End If  ' ** blnFound.

    Set dbs = CurrentDb
    With dbs

      If blnFound = True Then
        lngFldID = 0&: strNewAutoNum = vbNullString
        Set tdf = .TableDefs(strNewTable)
        With tdf
          For Each fld In .Fields
            With fld
              If (.Attributes And dbAutoIncrField) > 0 Then
                strNewAutoNum = .Name
                blnFound = False
                varTmp00 = DLookup("[fld_id]", "tblDatabase_Table_Field", "[dbs_id] = " & CStr(lngDbsID) & " And " & _
                  "[tbl_id] = " & CStr(lngTblID) & " And [fld_name] = '" & .Name & "'")  ' ** Module Function: modFileUtilities.
                If IsNull(varTmp00) = False Then
                  blnFound = True
                  lngFldID = CLng(varTmp00)
                Else
                  blnRetValx = False
                  Beep
                  Debug.Print "'FLD NOT FOUND: " & strNewAutoNum
                End If
              End If
            End With
          Next
        End With
      ElseIf blnRetValx = True Then
        blnRetValx = False
        Beep
        Debug.Print "'TBL STILL NOT FOUND IN tblDatabase_Table!"
      End If  ' ** blnFound.

      If blnFound = True Then

        If lngThisDbsID = 1& Then
          lngTblOrd = (DMax("[mtbl_ORDER]", "m_TBL") + 1&)
          lngMTblID = 0&
        End If
        varTmp00 = Empty: blnFound = False
        varTmp00 = vbNullString

        If lngThisDbsID = 1& Then
          Set rst1 = .OpenRecordset("m_TBL", dbOpenDynaset, dbConsistent)
          Set rst2 = .OpenRecordset("tblTemplate_m_TBL", dbOpenDynaset, dbConsistent)
        End If
        ' ** tblDatabase_Table_Link, by CurrentAppName().
        Set qdf3 = .QueryDefs("zz_qry_Database_Table_Link_40")
        Set rst3 = qdf3.OpenRecordset
        ' ** tblTemplate_Database_Table_Link, by CurrentAppName().
        Set qdf4 = .QueryDefs("zz_qry_Database_Table_Link_41")
        Set rst4 = qdf4.OpenRecordset
        If lngThisDbsID = 1& Then
          With rst1
            .MoveLast
            lngRecs1 = .RecordCount
            .MoveFirst
            .FindFirst "[mtbl_NAME] = '" & strNewTable & "'"
            If .NoMatch = False Then
              blnFound = True
              varTmp00 = .Name
            End If
          End With
        End If
        If blnFound = False Then
          If lngThisDbsID = 1& Then
            With rst2
              .MoveLast
              lngRecs2 = .RecordCount
              .MoveFirst
              .FindFirst "[mtbl_NAME] = '" & strNewTable & "'"
              If .NoMatch = False Then
                blnFound = True
                varTmp00 = .Name
              End If
            End With
          End If
          If blnFound = False Then
            With rst3
              .MoveLast
              lngRecs3 = .RecordCount
              .MoveFirst
              .FindFirst "[tbllnk_name] = '" & strNewTable & "'"
              If .NoMatch = False Then
                blnFound = True
                varTmp00 = .Name
              End If
            End With
            If blnFound = False Then
              With rst4
                .MoveLast
                lngRecs4 = .RecordCount
                .MoveFirst
                .FindFirst "[tbllnk_name] = '" & strNewTable & "'"
                If .NoMatch = False Then
                  blnFound = True
                  varTmp00 = .Name
                End If
              End With
            End If
          End If
        End If
        If blnFound = True Then
          blnRetValx = False
          Beep
          Debug.Print "'TBL ALREADY FOUND: " & varTmp00
        End If
      End If  ' ** blnFound.

      If blnRetValx = True Then
        If lngThisDbsID <> 1& Then
          lngRecs1 = lngRecs3
          lngRecs2 = lngRecs3
        End If
        If lngRecs1 = lngRecs2 And lngRecs1 = lngRecs3 And lngRecs1 = lngRecs4 Then

          If lngThisDbsID = 1& Then
            ' ** Table 1: m_TBL.
            With rst1
              .AddNew
              ' ** ![mtbl_ID] = {AutoNumber}
              ![mtbl_NAME] = strNewTable
              If lngFldID > 0& Then
                ![mtbl_AUTONUMBER] = strNewAutoNum
              End If
              ![mtbl_ORDER] = lngTblOrd
              ![mtbl_NEWRecs] = blnNewRecs
              ![mtbl_ACTIVE] = CBool(True)
              Select Case strNewDbs
              Case gstrFile_DataName
                ![mtbl_DTA] = CBool(True)
                ![mtbl_ARCH] = CBool(False)
                ![mtbl_AUX] = CBool(False)
              Case gstrFile_ArchDataName
                ![mtbl_ARCH] = CBool(True)
                ![mtbl_DTA] = CBool(False)
                ![mtbl_AUX] = CBool(False)
              Case gstrFile_AuxDataName
                ![mtbl_AUX] = CBool(True)
                ![mtbl_DTA] = CBool(False)
                ![mtbl_ARCH] = CBool(False)
              End Select
              .Update
              .Bookmark = .LastModified
              lngMTblID = ![mtbl_ID]
            End With
          End If

          If lngThisDbsID = 1& Then
            ' ** Table 2: tblTemplate_m_TBL.
            With rst2
              .AddNew
              ![mtbl_ID] = lngMTblID
              ![mtbl_NAME] = strNewTable
              If lngFldID > 0& Then
                ![mtbl_AUTONUMBER] = strNewAutoNum
              End If
              ![mtbl_ORDER] = lngTblOrd
              ![mtbl_NEWRecs] = blnNewRecs
              ![mtbl_ACTIVE] = CBool(True)
              Select Case strNewDbs
              Case gstrFile_DataName
                ![mtbl_DTA] = CBool(True)
                ![mtbl_ARCH] = CBool(False)
                ![mtbl_AUX] = CBool(False)
              Case gstrFile_ArchDataName
                ![mtbl_ARCH] = CBool(True)
                ![mtbl_DTA] = CBool(False)
                ![mtbl_AUX] = CBool(False)
              Case gstrFile_AuxDataName
                ![mtbl_AUX] = CBool(True)
                ![mtbl_DTA] = CBool(False)
                ![mtbl_ARCH] = CBool(False)
              End Select
              .Update
            End With
          End If

          ' ** Table 3: tblDatabase_Table_Link.
          With rst3
            .AddNew
            ![dbs_id] = lngDbsID
            ![tbl_id] = lngTblID
            ![dbs_id_asof] = lngThisDbsID
            ' ** ![tbllnk_id] = {AutoNumber}
            ![tbllnk_name] = strNewTable
            ![tbllnk_sourcetablename] = strNewSourceTable
            ![contype_type] = dbCJet
            ![tbllnk_connect] = strNewConnect
            ![tbllnk_versionadded] = "2.2.20"  'AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
            ![tbllnk_datemodified] = Now()
            .Update
            .Bookmark = .LastModified
            lngTblLnkID = ![tbllnk_id]
          End With

          ' ** Table 4: tblTemplate_Database_Table_Link
          With rst4
            .AddNew
            ![dbs_id] = lngDbsID
            ![tbl_id] = lngTblID
            ![dbs_id_asof] = lngThisDbsID
            ![tbllnk_id] = lngTblLnkID
            ![tbllnk_name] = strNewTable
            ![tbllnk_sourcetablename] = strNewSourceTable
            ![contype_type] = dbCJet
            ![tbllnk_connect] = strNewConnect
            ![tbllnk_versionadded] = "2.2.20"  'AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
            ![tbllnk_datemodified] = Now()
            .Update
            .Bookmark = .LastModified
            lngTblLnkID = 0&
          End With

          Debug.Print "'LINK ADDED: " & strNewTable

        Else
          blnRetValx = False
          Beep
          Debug.Print "'LINK TBLS DON'T MATCH!"
        End If  ' ** lngRecs.
      End If  ' ** blnRetValx.

      .Close
    End With  ' ** dbs.

  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If  ' ** EmptyDatabase.

  Set fld = Nothing
  Set tdf = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  Set rst4 = Nothing
  Set qdf3 = Nothing
  Set qdf4 = Nothing
  Set dbs = Nothing

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Tbl_Link_Add = blnRetValx

End Function

Public Function Tbl_DescCnt() As Boolean
' ** Put record count for System Doc tables into their Description,
' ** update tblDatabase_Table, and run TableDescriptionUpdate2()
' ** to update source table descriptions as well.

  Const THIS_PROC As String = "Tbl_DescCnt"

  Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, tdf As DAO.TableDef, prp As DAO.Property
  Dim lngBases As Long, arr_varBase() As Variant
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim strThisFile As String, strAlert As String
  Dim lngRecs As Long
  Dim lngPos1 As Long, lngPos2 As Long, lngLen As Long
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varBase().
  Const B_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const B_NAM As Integer = 0
  Const B_LEN As Integer = 1
  Const B_LFT As Integer = 2

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 8  ' ** Array's first-element UBound().
  Const T_DID    As Integer = 0
  Const T_DNAM   As Integer = 1
  Const T_TID    As Integer = 2
  Const T_TNAM   As Integer = 3
  Const T_RID    As Integer = 4
  Const T_CNT    As Integer = 5
  Const T_DSCOLD As Integer = 6
  Const T_DSCNEW As Integer = 7
  Const T_NOPROP As Integer = 8

  blnRetValx = True

  ' ** First, update the table record counts.
  Tbl_RecCnt_Doc  ' ** Function: Above.
  DoEvents

  lngBases = 11&
  lngE = lngBases - 1&
  ReDim arr_varBase(B_ELEMS, lngE)

  ' ** tblDatabase & tblDatabase_..
  arr_varBase(B_NAM, 0) = "tblDatabase"
  ' ** tblDocument & tblDocument_..
  arr_varBase(B_NAM, 1) = "tblDocument"
  ' ** tblForm & tblForm_..
  arr_varBase(B_NAM, 2) = "tblForm"
  ' ** tblIndex & tblIndex_..
  arr_varBase(B_NAM, 3) = "tblIndex"
  ' ** tblMacro & tblMacro_..
  arr_varBase(B_NAM, 4) = "tblMacro"
  ' ** tblQuery & tblQuery_..
  arr_varBase(B_NAM, 5) = "tblQuery"
  ' ** tblRelation & tblRelation_..
  arr_varBase(B_NAM, 6) = "tblRelation"
  ' ** tblReport & tblReport_..
  arr_varBase(B_NAM, 7) = "tblReport"
  ' ** tblSystemColor & tblSystemColor_..
  arr_varBase(B_NAM, 8) = "tblSystemColor"
  ' ** tblVBComponent & tblVBComponent_..
  arr_varBase(B_NAM, 9) = "tblVBComponent"
  ' ** tblVersion & tblVersion_..
  arr_varBase(B_NAM, 10) = "tblVersion"

  For lngX = 0& To (lngBases - 1&)
    arr_varBase(B_LEN, lngX) = Len(arr_varBase(B_NAM, lngX))
    arr_varBase(B_LFT, lngX) = (arr_varBase(B_NAM, lngX) & "_")
  Next

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    For lngX = 0& To (lngBases - 1&)
      ' ** zz_qry_Database_Table_RecCnt_02b (zz_qry_Database_Table_RecCnt_02a
      ' ** (tblDatabase_Table_Link, just those linked to CurrentAppName(), with
      ' ** tbl_name_len), linked to tblDatabase_Table_RecordCount,
      ' ** with tbl_name_base, reccnt_id, reccnt_count).
      Set qdf = .QueryDefs("zz_qry_Database_Table_RecCnt_02")
      Set rst = qdf.OpenRecordset
      With rst
        If .BOF = True And .EOF = True Then
          Stop
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngY = 1& To lngRecs
            If ![tbl_name_base] = arr_varBase(B_NAM, lngX) Then
              If ![tbl_name] = arr_varBase(B_NAM, lngX) Or Left$(![tbl_name], (arr_varBase(B_LEN, lngX) + 1)) = arr_varBase(B_LFT, lngX) Then
                If InStr(![tbl_name], "_Staging") = 0 Then  ' ** Skip Staging tables.
                  lngTbls = lngTbls + 1&
                  lngE = lngTbls - 1&
                  ReDim Preserve arr_varTbl(T_ELEMS, lngE)
                  arr_varTbl(T_DID, lngE) = ![dbs_id]
                  arr_varTbl(T_DNAM, lngE) = ![dbs_name]
                  arr_varTbl(T_TID, lngE) = ![tbl_id]
                  arr_varTbl(T_TNAM, lngE) = ![tbl_name]
                  arr_varTbl(T_RID, lngE) = ![reccnt_id]
                  arr_varTbl(T_CNT, lngE) = ![reccnt_count]
                  arr_varTbl(T_DSCOLD, lngE) = vbNullString
                  arr_varTbl(T_DSCNEW, lngE) = vbNullString
                  arr_varTbl(T_NOPROP, lngE) = CBool(False)
                End If
              End If
            End If
            If lngY < lngRecs Then .MoveNext
          Next
        End If
        .Close
      End With
      Set rst = Nothing
      Set qdf = Nothing
    Next

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    'For lngX = 0& To (lngTbls - 1&)
    '  'Debug.Print "'" & arr_varTbl(T_TNAM, lngX)
    '  If IsNull(arr_varTbl(T_CNT, lngX)) = True Then
    '    Debug.Print "'NULL CNT: " & arr_varTbl(T_TNAM, lngX)
    '  Else
    '    If arr_varTbl(T_CNT, lngX) = 0& Then
    '      Debug.Print "'ZERO CNT: " & arr_varTbl(T_TNAM, lngX)
    '    End If
    '  End If
    'Next
'ZERO CNT: tblDatabase_Table_Field_DateFormat
'ZERO CNT: tblQuery_FormRef

    For lngX = 0& To (lngTbls - 1&)
      strTmp01 = vbNullString
      Set tdf = .TableDefs(arr_varTbl(T_TNAM, lngX))
      With tdf
        For Each prp In .Properties
          With prp
            If .Name = "Description" Then
              strTmp01 = .Value
              Exit For
            End If
          End With
        Next
        If strTmp01 = vbNullString Then
          arr_varTbl(T_NOPROP, lngX) = CBool(True)
          varTmp00 = DLookup("[tbl_description]", "tblDatabase_Table", "[tbl_id] = " & CStr(arr_varTbl(T_TID, lngX)))
          If IsNull(varTmp00) = False Then
            strTmp01 = Trim(varTmp00)
          Else
            Stop
          End If
        End If
      End With
      If strTmp01 <> vbNullString Then
        arr_varTbl(T_DSCOLD, lngX) = strTmp01
      Else
        Debug.Print "'TBL DESC NOT FOUND! " & arr_varTbl(T_TNAM, lngX)
      End If
    Next

    For lngX = 0& To (lngTbls - 1&)
      strTmp01 = arr_varTbl(T_DSCOLD, lngX)
      lngLen = Len(strTmp01)
      lngPos1 = 0&: strTmp02 = vbNullString
      strAlert = vbNullString
      For lngY = lngLen To 1& Step -1&
        If Mid$(strTmp01, lngY, 1) = ";" Then
          lngPos1 = lngY
          Exit For
        End If
      Next
      If lngPos1 > 0& Then
        strTmp02 = Trim$(Mid$(strTmp01, lngPos1))
        If Mid$(strTmp02, 2, 1) = " " Then
          strTmp02 = Mid$(strTmp02, 3)
          If Right$(strTmp02, 2) = "##" Then
            lngPos2 = InStr(strTmp02, "##")
            varTmp00 = strTmp02
            strTmp02 = Trim$(Left$(strTmp02, (lngPos2 - 1&)))  ' ** Strip the extra spaces and alert from the description.
            strAlert = Mid(varTmp00, (Len(strTmp02) + 1))
          End If
          If Right$(strTmp02, 1) = "." Or Right$(strTmp02, 1) = "!" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
          If IsNumeric(strTmp02) = True Then
            strTmp02 = Left$(strTmp01, lngPos1) & " " & CStr(arr_varTbl(T_CNT, lngX)) & "."
            arr_varTbl(T_DSCNEW, lngX) = strTmp02 & strAlert
          Else
            Debug.Print "'DESC WRONG?  TBL: " & arr_varTbl(T_TNAM, lngX) & "  DESC: '" & strTmp01 & "'"
          End If
        Else
          Debug.Print "'DESC WRONG?  TBL: " & arr_varTbl(T_TNAM, lngX) & "  DESC: '" & strTmp01 & "'"
        End If
      Else
        Debug.Print "'DESC WRONG?  TBL: " & arr_varTbl(T_TNAM, lngX) & "  DESC: '" & strTmp01 & "'"
      End If
    Next

    .Close
  End With  ' ** dbs.
  DoEvents

'For lngX = 0& To (lngTbls - 1&)
'  If arr_varTbl(T_DSCNEW, lngX) = vbNullString Then
'    Debug.Print "'NO DESC! " & arr_varTbl(T_DSCNEW, lngX)
'  End If
'Next

  strThisFile = CurrentAppName  ' ** Module Function: modFileUtilities.

  For lngX = 0& To (lngTbls - 1&)
    If arr_varTbl(T_DNAM, lngX) = strThisFile Then
      Set wrk = DBEngine.Workspaces(0)
    Else
      Set wrk = CreateWorkspace("tmp", "superuser", TA_SEC, dbUseJet)
    End If
    With wrk
      If arr_varTbl(T_DNAM, lngX) = strThisFile Then
        Set dbs = .Databases(0)
      Else
        varTmp00 = DLookup("[dbs_path]", "tblDatabase", "[dbs_name] = '" & arr_varTbl(T_DNAM, lngX) & "'")
        varTmp00 = varTmp00 & LNK_SEP & arr_varTbl(T_DNAM, lngX)
        Set dbs = .OpenDatabase(varTmp00, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
      End If
      With dbs

        Set tdf = .TableDefs(arr_varTbl(T_TNAM, lngX))
        With tdf
          If arr_varTbl(T_NOPROP, lngX) = False Then
            .Properties("Description") = arr_varTbl(T_DSCNEW, lngX)
          Else
            Debug.Print "'NO PROPERTY: " & arr_varTbl(T_TNAM, lngX)
          End If
        End With
        Set tdf = Nothing

        .TableDefs.Refresh
        .TableDefs.Refresh

        ' ** Update tblDatabase_Table tbl_description, by specified [tblid], [dsc].
        Set qdf = .QueryDefs("zz_qry_Database_Table_RecCnt_03")
        With qdf.Parameters
          ![tblid] = arr_varTbl(T_TID, lngX)
          ![dsc] = arr_varTbl(T_DSCNEW, lngX)
        End With
        qdf.Execute dbFailOnError
        Set qdf = Nothing

        .Close
      End With  ' ** dbs.
      .Close
    End With  ' ** wrk.
    DoEvents
  Next  ' ** lngX.

  ' ** Update table descriptions.
  TableDescriptionUpdate  ' ** Module Function: modFileUtilities.

  Set prp = Nothing
  Set tdf = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set wrk = Nothing

  Beep

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Tbl_DescCnt = blnRetValx

End Function

Public Function Tbl_AutoNum_Holes() As Boolean
' ** Look for holes in a table's AutoNumber sequence.

  Const THIS_PROC As String = "Tbl_AutoNum_Holes"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, idx As DAO.index
  Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset, qdf As DAO.QueryDef
  Dim strTable As String, strField As String, strField2 As String, strIndex As String
  Dim strTableSave As String
  Dim lngRecs As Long
  Dim lngAutoID As Long, lngThisDbsID As Long
  Dim lngANs As Long, arr_varAN() As Variant
  Dim datDatePrev As Date, datDateNext As Date
  Dim blnFound As Boolean, blnSpace_fld As Boolean, blnAdvance As Boolean
  Dim varTmp00 As Variant
  Dim lngX As Long, lngE As Long

  Const AN_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const AN_TBL As Integer = 0
  Const AN_FLD As Integer = 1
  Const AN_VAL As Integer = 2
  Const AN_PRV As Integer = 3
  Const AN_NXT As Integer = 4

  blnRetValx = True

  strTable = "tblPreference_Control"
  strField = "prefctl_id"  ' ** If vbNullString, then it looks for the AutoNumber field.
  strTableSave = "tblMark"

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngANs = 0&
  ReDim arr_varAN(AN_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    strField2 = vbNullString
    blnSpace_fld = False
    If strField = vbNullString Then
      blnFound = False
    Else
      blnFound = True
    End If

    Set tdf = .TableDefs(strTable)

      ' ** First, if no field is specified, find the AutoNumber field.
    If strField = vbNullString Then
      With tdf
        For Each fld In .Fields
          With fld
            If CBool(.Attributes And dbAutoIncrField) = True Then
              blnFound = True
              strField = .Name
              If InStr(strField, " ") > 0 Then  ' ** We may not need this check in this function.
                blnSpace_fld = True
                strField2 = "[" & strField & "]"
              End If
              Exit For
            End If
            ' ** Field Attribute constant enumeration:
            ' **      1  dbDescending      The field is sorted in descending (Z to A or 100 to 0) order;
            ' **                           this option applies only to a Field object in a Fields collection of an Index object.
            ' **                           If you omit this constant, the field is sorted in ascending (A to Z or 0 to 100) order.
            ' **                           This is the default value for Index and TableDef fields (Microsoft Jet workspaces only).
            ' **      1  dbFixedField      The field size is fixed (default for Numeric fields).
            ' **      2  dbVariableField   The field size is variable (Text fields only).
            ' **     16  dbAutoIncrField   The field value for new records is automatically incremented to a unique Long integer
            ' **                           that can't be changed (in a Microsoft Jet workspace, supported only for
            ' **                           Microsoft Jet database(.mdb) tables).
            ' **     32  dbUpdatableField  The field value can be changed.
            ' **   8192  dbSystemField     The field stores replication information for replicas;
            ' **                           you can't delete this type of field (Microsoft Jet workspaces only).
            ' **  32768  dbHyperlinkField  The field contains hyperlink information (Memo fields only).
          End With
        Next
      End With  ' ** tdf.
    End If

    If blnFound = True Then

      ' ** Next, find the AutoNumber index.
'If False Then
      With tdf
        blnFound = False: strIndex = vbNullString
        For Each idx In .Indexes
          With idx
            Set fld = .Fields(0)  ' ** We want the AutoNumber field to be the first field.
            If fld.Name = strField Then
              ' ** Check to make sure it's a unique index.
              If .Unique = True Then
                blnFound = True
                strIndex = .Name
              Else
                ' ** Just doc this and let it keep checking.
                Debug.Print "'AUTO NOT UNIQUE: " & tdf.Name & "  IDX: " & .Name
              End If
            End If
          End With
        Next
      End With  ' ** tdf.
'End If

      If blnFound = True Then

        ' ** Open the table and sort it by the strField.
        Set rst1 = .OpenRecordset(strTable)
        ' ** Recordset Type enumeration:
        ' **    1  dbOpenTable        Opens a table-type Recordset object (Microsoft Jet workspaces only).
        ' **    2  dbOpenDynaset      Opens a dynaset-type Recordset object, which is similar to an ODBC keyset cursor.
        ' **    4  dbOpenSnapshot     Opens a snapshot-type Recordset object, which is similar to an ODBC static cursor.
        ' **    8  dbOpenForwardOnly  Opens a forward-only-type Recordset object.
        ' **   16  dbOpenDynamic      Opens a dynamic-type Recordset object, which is similar to an ODBC dynamic cursor.
        ' **                          (ODBCDirect workspaces only)
        ' ** Recordset Option enumeration:
        ' **      8  dbAppendOnly      Allows users to append new records to the Recordset, but prevents them from editing
        ' **                           or deleting existing records (Microsoft Jet dynaset-type Recordset only).
        ' **     64  dbSQLPassThrough  Passes an SQL statement to a Microsoft Jet-connected ODBC data source for processing
        ' **                           (Microsoft Jet snapshot-type Recordset only).
        ' **    512  dbSeeChanges      Generates a run-time error if one user is changing data that another user is editing
        ' **                           (Microsoft Jet dynaset-type Recordset only). This is useful in applications where multiple
        ' **                           users have simultaneous read/write access to the same data.
        ' **      1  dbDenyWrite       Prevents other users from modifying or adding records (Microsoft Jet Recordset objects only).
        ' **      2  dbDenyRead        Prevents other users from reading data in a table (Microsoft Jet table-type Recordset only).
        ' **    256  dbForwardOnly     Creates a forward-only Recordset (Microsoft Jet snapshot-type Recordset only).
        ' **                           It is provided only for backward compatibility, and you should use the dbOpenForwardOnly
        ' **                           constant in the type argument instead of using this option.
        ' **      4  dbReadOnly        Prevents users from making changes to the Recordset (Microsoft Jet only).
        ' **                           The dbReadOnly constant in the lockedits argument replaces this option,
        ' **                           which is provided only for backward compatibility.
        ' **   1024  dbRunAsync        Runs an asynchronous query (ODBCDirect workspaces only).
        ' **   2048  dbExecDirect      Runs a query by skipping SQLPrepare and directly calling SQLExecDirect (ODBCDirect
        ' **                           workspaces only). Use this option only when you’re not opening a Recordset based
        ' **                           on a parameter query. For more information, see the "Microsoft ODBC 3.0 Programmer’s Reference."
        ' **     16  dbInconsistent    Allows inconsistent updates (Microsoft Jet dynaset-type and snapshot-type Recordset objects only).
        ' **     32  dbConsistent      Allows only consistent updates (Microsoft Jet dynaset-type and snapshot-type Recordset objects only).

        With rst1
          .sort = (strField)
          Set rst2 = .OpenRecordset  ' ** Ordered recordset.
        End With

        ' ** Finally, look for holes in the sequence.
        lngAutoID = 0&: datDatePrev = 1
        With rst2
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          lngAutoID = .Fields(strField)
          If lngAutoID = 1& Then
            lngAutoID = lngAutoID - 1&
          Else
            If lngAutoID > 1& Then
              lngAutoID = lngAutoID - 1&
            End If
          End If
          blnAdvance = False
          For lngX = 1& To lngRecs
            If .Fields(strField) <> lngAutoID + 1& Then
              blnAdvance = True
              Do While blnAdvance = True
                ' ** Advance the count alone until it meshes with the current record.
                If .Fields(strField) <> lngAutoID + 1& Then
                  'datDateNext = ![posted]
                  lngANs = lngANs + 1&
                  lngE = lngANs - 1&
                  ReDim Preserve arr_varAN(AN_ELEMS, lngE)
                  arr_varAN(AN_TBL, lngE) = strTable
                  arr_varAN(AN_FLD, lngE) = strField
                  lngAutoID = lngAutoID + 1&
                  arr_varAN(AN_VAL, lngE) = lngAutoID
                  'arr_varAN(AN_PRV, lngE) = datDatePrev
                  'arr_varAN(AN_NXT, lngE) = datDateNext
                Else
                  lngAutoID = .Fields(strField)
                  blnAdvance = False
                End If
              Loop
              'datDatePrev = ![posted]
            Else
              lngAutoID = .Fields(strField)
              'datDatePrev = ![posted]
              blnAdvance = False
            End If
            If lngX < lngRecs Then .MoveNext
          Next
          .Close
        End With
        rst1.Close

        If lngANs > 0& Then
          If strTableSave = "tblMark" Then
            varTmp00 = DLookup("[qry_name]", "tblQuery", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
              "[qry_description] = 'Empty tblMark.'")
            ' ** Empty tblMark.
            Set qdf = .QueryDefs(varTmp00)
            qdf.Execute
          End If
          Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
          Debug.Print "'HOLES: " & CStr(lngANs) & "  TBL: " & strTable
          Set rst1 = .OpenRecordset(strTableSave, dbOpenDynaset, dbAppendOnly)
          With rst1
            For lngX = 0& To (lngANs - 1&)
              .AddNew
              ![unique_id] = arr_varAN(AN_VAL, lngX)
              '![mark] = False 'arr_varAN(AN_PRV, lngX)
              '![posted_next] = arr_varAN(AN_NXT, lngX)
              .Update
              'Debug.Print "'  " & CStr(arr_varAN(AN_VAL, lngX))
            Next
            .Close
          End With
        Else
          Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
          Debug.Print "'NO HOLES! " & strTable
        End If

      Else
        Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
        Debug.Print "'NO AUTO INDEX: " & strTable
      End If

    Else
      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      Debug.Print "'AUTO NOT FOUND: " & strTable
    End If

    .Close
  End With

  Beep

  Set fld = Nothing
  Set idx = Nothing
  Set tdf = Nothing
  Set qdf = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set dbs = Nothing

  Tbl_AutoNum_Holes = blnRetValx

End Function

Public Function Tbl_ColBestFit() As Boolean

  Const THIS_PROC As String = "Tbl_ColBestFit"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, fld As DAO.Field, prp As Property
  Dim frm As Form, ctl1 As Control, ctl2 As Control
  Dim obj As Object
  Dim strFind As String
  Dim intPrps As Integer
  Dim strCtlName As String
  Dim intX As Integer
  Dim blnRetVal As Boolean

  blnRetVal = True

  strFind = "tblForm_Graphics"

' ** Width:
' **     Popup    Twips   Pixels  Font-Based Twips-Per-Number
' **   =========  ======  ======  =============================
' **    15.6667      -1           Standard Width
' **     8          750     50    93.75
' **    11.3333    1050     70    92.6473313156803
' **    10.5        975     65    92.8571428571429
' **    47.6667    4320    288    90.6293072522327
' **    15.3333    1410     94    91.9567216450471
' **    55         4980    332    90.5454545454545

  Set obj = Application.Screen.ActiveDatasheet
  With obj
    intX = -1
    For Each ctl1 In .Controls
      With ctl1
        'If .ControlType = acDatasheetColumn Then
          If .ControlType <> acDatasheetColumn Then
            Debug.Print "'" & .Name & "  CTLTYP: " & .ControlType
          End If
          If Left$(.Name, 8) = "ctl_name" Then
            'Debug.Print "'" & Left$(CStr(.ColumnWidth) & "    ", 4) & "  " & .Name
            .ColumnWidth = 4980
            '.SizeToFit
            'intPrps = .Properties.Count
            'intX = -1
            'For Each prp In .Properties
            '  intX = intX + 1
            '  With prp
            '    Debug.Print "'" & Left$(CStr(intX) & "." & "    ", 4) & " " & ctl1.Properties(intX).Name
            '  End With
            '  If intX >= 200 Then
            '    Exit For
            '  End If
            'Next
          End If
        'End If
      End With
    Next
  End With

'Datasheet Column Properties:
'0.   ColumnWidth
'1.   ColumnOrder
'2.   ColumnHidden
'3.   DefaultValue
'4.   Name
'5.   ControlType
'6.   ControlSource
'7.   Enabled
'8.   Locked
'9.   Format
'10.  Text
'11.  SelStart
'12.  SelLength
'13.  SelText

'Control Properties:
'1.   ColumnWidth
'2.   ColumnOrder
'3.   ColumnHidden
'4.   DefaultValue
'5.   Name
'6.   ControlType
'7.   ControlSource
'8.   Enabled
'9.   Locked
'10.  Format
'11.  Text
'12.  SelStart
'13.  SelLength
'14.  SelText
'15.  SmartTags
'16.  AggregateType

  Beep

  Set ctl1 = Nothing
  Set ctl2 = Nothing
  Set frm = Nothing
  Set obj = Nothing
  Set prp = Nothing
  Set fld = Nothing
  Set tdf = Nothing
  Set dbs = Nothing

  Tbl_ColBestFit = blnRetVal

End Function

Public Function Dbs_Properties_Doc() As Boolean

  Const THIS_PROC As String = "Dbs_Properties_Doc"

  Dim wrk As DAO.Workspace, dbs As DAO.Database, prp As DAO.Property, cntr As DAO.Container, doc As DAO.Document
  Dim intWrkType As Integer
  Dim blnLocal As Boolean

  blnRetValx = True
  blnLocal = False

  intWrkType = 0
  If blnLocal = True Then
    Set wrk = DBEngine.Workspaces(0)
  Else
On Error Resume Next
    Set wrk = CreateWorkspace("tmpBE", "admin", "", dbUseJet)
    If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
      Set wrk = CreateWorkspace("tmpBE", "superuser", TA_SEC, dbUseJet)
      If ERR.Number <> 0 Then
On Error GoTo 0
        Set wrk = CreateWorkspace("tmpBE", "superuser", TA_SEC6, dbUseJet)
        intWrkType = 3
      Else
On Error GoTo 0
        intWrkType = 2
      End If
    Else
On Error GoTo 0
      intWrkType = 1
    End If
  End If

  With wrk

    If blnLocal = True Then
      Set dbs = .Databases(0)
    Else
      Set dbs = .OpenDatabase("C:\VictorGCS_Clients\TrustAccountant\NewWorking\TestDatabase\TrustSec.mdw", False, True)  ' ** {pathfile}, {exclusive}, {read-only}  '## OK
    End If

    With dbs

      'For Each prp In .Properties
      '  With prp
      '    Debug.Print "'" & .Name
      '  End With
      'Next
'TrustSec.mdw Properties:
'Name
'Connect
'Transactions
'Updatable
'CollatingOrder
'QueryTimeout
'Version
'RecordsAffected
'ReplicaID
'DesignMasterID
'Connection

      'For Each cntr In .Containers
      '  With cntr
      '    Debug.Print "'" & .Name
      '  End With
      'Next
'TrustSec.mdw Containers:
'Databases
'Relationships
'Tables

      Set cntr = .Containers("Databases")
      With cntr
        For Each doc In .Documents
          With doc
            Debug.Print "'DOC: " & .Name
            For Each prp In .Properties
              With prp
                Debug.Print "'PRP: " & .Name
              End With
            Next
          End With
        Next
      End With
' ** Databases Container MSysDb Document Properties
'Name
'Owner
'UserName
'Permissions
'AllPermissions
'Container
'DateCreated
'LastUpdated

      .Close
    End With

    .Close
  End With

'dbs.Containers("Databases").Documents("MSysDb").Properties("AccessVersion") = 08.50

' ** Databases Container Documents:
'AccessLayout
'MSysDb
'SummaryInfo
'UserDefined

' ** AccessLayout Document Properties:
'Name
'Owner
'UserName
'Permissions
'AllPermissions
'Container
'DateCreated
'LastUpdated
'KeepLocal

' ** MSysDb Document Properties:
'Name
'Owner
'UserName
'Permissions
'AllPermissions
'Container
'DateCreated
'LastUpdated
'AccessVersion
'Build
'AppTitle
'AppIcon

' ** SummaryInfo Document Properties:
'Name
'Owner
'UserName
'Permissions
'AllPermissions
'Container
'DateCreated
'LastUpdated
'Title
'Author
'Company
'Manager

' ** UserDefined Document Properties:
'Name
'Owner
'UserName
'Permissions
'AllPermissions
'Container
'DateCreated
'LastUpdated
'ReplicateProject
'AppVersion
'AppDate

' ** Database Containers:
'DataAccessPages
'Databases
'Forms
'Modules
'Relationships
'Reports
'Scripts
'SysRel
'Tables

' ** Database Properties:
'Name
'Connect
'Transactions
'Updatable
'CollatingOrder
'QueryTimeout
'Version
'RecordsAffected
'ReplicaID
'DesignMasterID
'Connection
'AccessVersion
'Build
'AppTitle
'AppIcon

  Beep

  Set doc = Nothing
  Set cntr = Nothing
  Set prp = Nothing
  Set dbs = Nothing
  Set wrk = Nothing

  Dbs_Properties_Doc = blnRetValx

End Function

Public Function DecPlace_Doc() As Boolean
' ** Populate tblDecimalPlaceDb.

  Const THIS_PROC As String = "DecPlace_Doc"

  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim intX As Integer

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs
    Set rst = .OpenRecordset("tblDecimalPlaceDb", dbOpenDynaset, dbAppendOnly)
    With rst
      For intX = 0 To 255
        .AddNew
        ![decplace_number] = intX
        If intX < 255 Then
          ![decplace_type] = CStr(intX)
        Else
          ![decplace_type] = "Auto"
        End If
        .Update
      Next
      .Close
    End With
    .Close
  End With

  Beep

  Set rst = Nothing
  Set dbs = Nothing

  DecPlace_Doc = blnRetValx

End Function

Private Function Tbl_Property_Add(tdf As DAO.TableDef, strName As String, varType As Variant, varValue As Variant) As Integer

On Error GoTo ERRH

  Const THIS_PROC As String = "Tbl_Property_Add"

  Dim prp As Object

  Const conPropNotFoundError As Long = 3270&

  blnRetValx = True

  tdf.Properties(strName) = varValue

EXITP:
  Set prp = Nothing
  Tbl_Property_Add = blnRetValx
  Exit Function

ERRH:
  Select Case ERR.Number
  Case conPropNotFoundError
    Set prp = tdf.CreateProperty(strName, varType, varValue)
    tdf.Properties.Append prp
    Resume        ' ** Execution resumes with the statement that caused the error.
    'Resume Next  ' ** Esecution resumes with the statement immediately following the statement that caused the error.
  Case Else
    blnRetValx = False
    Resume EXITP  ' ** Execution resumes at line specified.
  End Select

End Function

Public Function Tbl_Tmp_List() As Boolean

  Const THIS_PROC As String = "Tbl_Tmp_List"

  Dim dbs As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim blnTblList As Boolean, blnQryList As Boolean
  Dim blnWithZZs As Boolean, blnWithApp As Boolean, blnWithTemplate As Boolean, blnWithNonTmp As Boolean
  Dim strTblName As String, strQryName As String, strTblDesc As String
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngRecs As Long
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const T_NAM As Integer = 0
  Const T_DSC As Integer = 1
  Const T_QRY As Integer = 2
  Const T_ZZT As Integer = 3

  blnRetValx = True

  blnTblList = True
  blnQryList = False

  blnWithApp = False
  blnWithTemplate = False
  blnWithZZs = False
  blnWithNonTmp = True

  Set dbs = CurrentDb
  With dbs

    If blnTblList = True Then

      For Each tdf In .TableDefs
        With tdf
          strTblName = .Name
          If Left$(strTblName, 3) = "tmp" And blnWithApp = True Then
            Debug.Print "'TBL: " & strTblName
          End If
          If Left$(strTblName, 11) = "tblTemplate" And blnWithTemplate = True Then
            Debug.Print "'TBL: " & strTblName
          End If
          If Left$(strTblName, 2) = "zz" And blnWithZZs = True Then
            Debug.Print "'TBL: " & strTblName
          End If
        End With
      Next

      If blnWithNonTmp = True Then
        lngTbls = 0&
        ReDim arr_varTbl(T_ELEMS, 0)
        Set rst = .OpenRecordset("tblDatabase_Table", dbOpenDynaset, dbReadOnly)
        With rst
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            If ![dbs_id] <> 5& Then  ' ** Skip TAJrnTmp.mdb.
              If IsNull(![tbl_description]) = False Then
                strTblDesc = ![tbl_description]
                If Left$(strTblDesc, 18) = "End-User Temporary" And Left$(![tbl_name], 3) <> "tmp" And _
                    Left$(![tbl_name], 12) <> "tblTemplate_" And Left$(![tbl_name], 2) <> "zz" Then
                  lngTbls = lngTbls + 1&
                  lngE = lngTbls - 1&
                  ReDim Preserve arr_varTbl(T_ELEMS, lngE)
                  arr_varTbl(T_NAM, lngE) = ![tbl_name]
                  arr_varTbl(T_DSC, lngE) = strTblDesc
                  arr_varTbl(T_QRY, lngE) = vbNullString
                  'If Left$(![tbl_name], 2) = "zz" Then
                  '  arr_varTbl(T_ZZT, lngE) = CBool(True)
                  'Else
                    arr_varTbl(T_ZZT, lngE) = CBool(False)
                  'End If
                End If
              End If
            End If
            If lngX < lngRecs Then .MoveNext
          Next
        End With
      End If

    End If

    If blnQryList = True Or blnWithNonTmp = True Then
      For Each qdf In .QueryDefs
        With qdf
          strQryName = .Name
          If Left$(strQryName, 19) = "qryTmp_Table_Empty_" Then
            If blnWithApp = True Then
              If Mid$(strQryName, 23, 1) = "_" Then  ' ** 3-digit numbers now!
                If Mid$(strQryName, 24, 3) = "tmp" Then
                  Debug.Print "'     QRY: " & strQryName
                End If
              Else
                If Mid$(strQryName, 25, 3) = "tmp" Then
                  Debug.Print "'     QRY: " & strQryName
                End If
              End If
            End If
            If blnWithTemplate = True Then
              If Mid$(strQryName, 24, 12) = "tblTemplate_" Then
                Debug.Print "'     QRY: " & strQryName
              End If
            End If
            If blnWithZZs = True Then

            End If
            If blnWithNonTmp = True Then
              For lngX = 0& To (lngTbls - 1&)
                If Right$(strQryName, Len(arr_varTbl(T_NAM, lngX))) = arr_varTbl(T_NAM, lngX) Then
                  If (arr_varTbl(T_ZZT, lngX) = False And InStr(strQryName, "_zz_") = 0) Or _
                      (arr_varTbl(T_ZZT, lngX) = True And InStr(strQryName, "_zz_") > 0) Then
                    If (InStr(arr_varTbl(T_NAM, lngX), "tblTemplate_") = 0 And InStr(strQryName, "tblTemplate_") = 0) Or _
                        (InStr(arr_varTbl(T_NAM, lngX), "tblTemplate_") > 0 And InStr(strQryName, "tblTemplate_") > 0) Then
                      If arr_varTbl(T_QRY, lngX) = vbNullString Then
                        arr_varTbl(T_QRY, lngX) = strQryName
                      Else
                        Debug.Print "'2 QRYS: " & strQryName
                      End If
                      Exit For
                    End If
                  End If
                ElseIf arr_varTbl(T_NAM, lngX) = "journal map" And Right$(strQryName, Len(arr_varTbl(T_NAM, lngX))) = "Journal_Map" Then
                  If arr_varTbl(T_QRY, lngX) = vbNullString Then
                    arr_varTbl(T_QRY, lngX) = strQryName
                  Else
                    Debug.Print "'2 QRYS: " & strQryName
                  End If
                  Exit For
                End If
              Next
            End If
          End If
        End With
      Next
    End If

    .Close
  End With

  If blnWithNonTmp = True Then
    For lngX = 0& To (lngTbls - 1&)
      Debug.Print "'TBL: " & arr_varTbl(T_NAM, lngX)
      Debug.Print "'     QRY: " & arr_varTbl(T_QRY, lngX)
    Next
  End If

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Debug.Print "'DONE! " & THIS_PROC & "()"

  Set qdf = Nothing
  Set tdf = Nothing
  Set rst = Nothing
  Set dbs = Nothing

  Beep

  Tbl_Tmp_List = blnRetValx

End Function

Public Function Qry_TmpTables(blnLoad As Boolean, Optional varJrnlTmp As Variant) As Boolean
' ** No longer deletes Temp tables; leave them around!
' ** Called by:
' **   zz_mod_DatabaseDocFuncs:
' **     Tbl_Doc()
' **     Tbl_Fld_Doc()
' **     Tbl_RecCnt_Doc()
' **     Tbl_AutoNum_Doc()
' **   zz_mod_IndexDocFuncs:
' **     Idx_Doc()
' **   zz_mod_QueryDocFuncs. (This):
' **     Qry_Doc()
' **     Qry_Parm_Doc()

  Const THIS_PROC As String = "Qry_TmpTables"

  Dim dbs0 As DAO.Database, qdf0 As DAO.QueryDef
  Dim blnJrnlTmp As Boolean

  blnRetValx = True

  If gstrTrustDataLocation = vbNullString Then
    IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
  End If

  If IsMissing(varJrnlTmp) = True Then
    blnJrnlTmp = True
  Else
    blnJrnlTmp = CBool(varJrnlTmp)
  End If

  Select Case blnLoad
  Case True

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    ' ** Make sure all the referenced temporary tables are present
    ' ** while documenting, then delete them when finished.
    If TableExists("tmpIncomeExpenseReports") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tmpIncomeExpenseReports", acTable, "zz_tmpIncomeExpenseReports"
    End If
    If TableExists("tmpUpdatedValues") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tmpUpdatedValues", acTable, "zz_tmpUpdatedValues"
    End If
    If TableExists("USysRibbons") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "USysRibbons", acTable, "zz_USysRibbons"
    End If
    If TableExists("tmp_ActiveAssets") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tmp_ActiveAssets", acTable, "tblTemplate_ActiveAssets"
    End If
    If TableExists("tmp_Journal") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tmp_Journal", acTable, "tblTemplate_Journal"
    End If
    If TableExists("tmp_Ledger") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tmp_Ledger", acTable, "tblTemplate_Ledger"
    End If
    If TableExists("tmp_m_REVCODE") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tmp_m_REVCODE", acTable, "tblTemplate_m_REVCODE"
    End If
    If TableExists("tmp_RecurringItems") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tmp_RecurringItems", acTable, "tblTemplate_RecurringItems"
    End If
    If TableExists("LedgerArchive_Backup") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "LedgerArchive_Backup", acTable, "tblTemplate_LedgerArchive"
    End If
    If TableExists("m_TBL_tmp01") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "m_TBL_tmp01", acTable, "tblTemplate_m_TBL"
    End If
    If TableExists("tblDatabase_Table_Link_tmp01") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tblDatabase_Table_Link_tmp01", acTable, "tblTemplate_Database_Table_Link"
    End If
    If TableExists("tblDatabase_Table_Link_tmp02") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tblDatabase_Table_Link_tmp02", acTable, "tblTemplate_Database_Table_Link"
    End If
    If TableExists("tblDatabase_Table_Link_tmp03") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "tblDatabase_Table_Link_tmp03", acTable, "tblTemplate_Database_Table_Link"
    End If

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    If blnJrnlTmp = True Then
      'RePost_TmpDB_Link  ' ** Module Function: modRePostFuncs.
      'RePost_TmpDB_Link_RT True  ' ** Module Function: modRePostFuncs.
    End If

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

  Case False

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    'TableDelete "tmp_RecurringItems"  ' ** Module Function: modFileUtilities.
    'TableDelete "tmp_m_REVCODE"  ' ** Module Function: modFileUtilities.
    'TableDelete "tmp_ActiveAssets"  ' ** Module Function: modFileUtilities.
    'TableDelete "tmp_Journal"  ' ** Module Function: modFileUtilities.
    'TableDelete "tmp_Ledger"  ' ** Module Function: modFileUtilities.
    'TableDelete "tmpIncomeExpenseReports"  ' ** Module Function: modFileUtilities.
    'TableDelete "tmpRevCodeEdit"  ' ** Module Function: modFileUtilities.
    'TableDelete "tmpUpdatedValues"  ' ** Module Function: modFileUtilities.
    'TableDelete "USysRibbons"  ' ** Module Function: modFileUtilities.
    TableDelete "LedgerArchive_Backup"  ' ** Module Function: modFileUtilities.

    TableDelete "account1"  ' ** Module Function: modFileUtilities.
    TableDelete "account2"  ' ** Module Function: modFileUtilities.
    TableDelete "account3"  ' ** Module Function: modFileUtilities.
    TableDelete "ActiveAssets1"  ' ** Module Function: modFileUtilities.
    TableDelete "ActiveAssets2"  ' ** Module Function: modFileUtilities.
    TableDelete "adminofficer1"  ' ** Module Function: modFileUtilities.
    TableDelete "adminofficer2"  ' ** Module Function: modFileUtilities.
    TableDelete "adminofficer3"  ' ** Module Function: modFileUtilities.
    TableDelete "Balance1"  ' ** Module Function: modFileUtilities.
    TableDelete "Balance2"  ' ** Module Function: modFileUtilities.
    TableDelete "ledger1"  ' ** Module Function: modFileUtilities.
    TableDelete "ledger2"  ' ** Module Function: modFileUtilities.
    TableDelete "Location1"  ' ** Module Function: modFileUtilities.
    TableDelete "Location2"  ' ** Module Function: modFileUtilities.
    TableDelete "masterasset1"  ' ** Module Function: modFileUtilities.
    TableDelete "masterasset2"  ' ** Module Function: modFileUtilities.
    TableDelete "tblAccountBalance1"  ' ** Module Function: modFileUtilities.
    TableDelete "zz_tbl_TmpLedger_Test"  ' ** Module Function: modFileUtilities.
    TableDelete "zz_tmpLedger_01"  ' ** Module Function: modFileUtilities.
    TableDelete "m_TBL_tmp01"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp01"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp02"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp03"  ' ** Module Function: modFileUtilities.

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    If blnJrnlTmp = True Then
      'RePost_TmpDB_Unlink  ' ** Module Function: modRePostFuncs.
      'RePost_TmpDB_Link_RT False  ' ** Module Function: modRePostFuncs.
    End If

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

  End Select

  Set qdf0 = Nothing
  Set dbs0 = Nothing

'zz_tmpIncomeExpenseReports
'zz_tmpUpdatedValues
'zz_USysRibbons
'tblTemplate_ActiveAssets
'tblTemplate_Database_Table_Link
'tblTemplate_Journal
'tblTemplate_Ledger
'tblTemplate_LedgerArchive
'tblTemplate_m_REVCODE
'tblTemplate_m_TBL
'tblTemplate_RecurringItems

  Qry_TmpTables = blnRetValx

End Function

Private Function Tbl_ChkDocQrys(Optional varSkip As Variant) As Boolean

  Const THIS_PROC As String = "Tbl_ChkDocQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngQrys As Long, arr_varQry As Variant
  Dim lngImpQrys As Long
  Dim blnSkip As Boolean
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

  'blnSkip = True
  If blnSkip = False Then

    Set dbs = CurrentDb
    With dbs
      ' ** tblQuery_Documentation, by specified [vbnam].
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
        ' **     4       3     qrydoc_id          Q_QDID
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
      .Close
    End With
    Set dbs = Nothing

    strPath = gstrDir_Dev
    strFile = "TrstXAdm.mdb" 'CurrentAppName  ' ** Module Function: modFileUtilities.
    strPathFile = strPath & LNK_SEP & strFile

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'DB DOC QRYS: " & CStr(lngQrys)
    DoEvents

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
          DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
          arr_varQry(Q_IMP, lngX) = CBool(True)
        End If
      Next
    Else
      Debug.Print "'ALL DB DOC QRYS PRESENT!"
    End If

    Debug.Print "'DONE!"
    DoEvents

    Beep

    Set rst = Nothing
    Set qdf = Nothing
    Set dbs = Nothing

  End If  ' ** blnSkip.

  Tbl_ChkDocQrys = blnRetValx

End Function
