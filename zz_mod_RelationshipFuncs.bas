Attribute VB_Name = "zz_mod_RelationshipFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_RelationshipFuncs"

'VGC 10/17/2014: CHANGES!

' ** DbRelation enumeration:
' **          0  dbRelationEnforce        The relationship is enforced (referential integrity). {my own}
' **          1  dbRelationUnique         The relationship is one-to-one.
' **          2  dbRelationDontEnforce    The relationship isn't enforced (no referential integrity).
' **          4  dbRelationInherited      The relationship exists in a non-current database that contains the two linked tables.
' **        256  dbRelationUpdateCascade  Updates will cascade.
' **       4096  dbRelationDeleteCascade  Deletions will cascade.
' **   16777216  dbRelationLeft           In Design view, display a LEFT JOIN as the default join type. Microsoft Access only.
' **   33554432  dbRelationRight          In Design view, display a RIGHT JOIN as the default join type. Microsoft Access only.

' ** Relationship Names:
' ** Maximum saved length seems to be 64, so must be unique within that.

' ** Right now, this can remain private.
Private Const dbRelationEnforce As Long = 0&

Private blnRetValx As Boolean, blnWithJrnlTmp As Boolean
' **

Public Function QuikRelDoc() As Boolean
  Const THIS_PROC As String = "QuikRelDoc"
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Rel_ChkDocQrys = True Then  ' ** Function: Below.
      blnWithJrnlTmp = False
      blnRetValx = Rel_Doc  ' ** Function: Below.
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED Rel_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikRelDoc = blnRetValx
End Function

Private Function Rel_Doc() As Boolean
' ** Document all relationships to tblRelation and tblRelation_Field.
' ** Called by:
' **   QuikRelDoc(), Above

  Const THIS_PROC As String = "Rel_Doc"

  Dim wrk As DAO.Workspace, dbsFE As DAO.Database, dbsBE As DAO.Database
  Dim qdf As DAO.QueryDef, tdf As DAO.TableDef
  Dim rstDB As DAO.Recordset, rstRel As DAO.Recordset, rstRelFld As DAO.Recordset
  Dim Rel As DAO.Relation, fld As DAO.Field
  Dim strThisFile As String, strThisPath As String, blnThisDbs As Boolean
  Dim strThatFile As String, strThatPath As String
  Dim lngDbs As Long, arr_varDb() As Variant, lngDBID_AsOf As Long, lngDBID1 As Long, lngDBID2 As Long
  Dim lngTblID1 As Long, lngTblID2 As Long, lngRelID As Long
  Dim lngRelCnt As Long, intRelOrd As Integer
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim blnFound As Boolean, blnFound2 As Boolean
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  Const D_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const D_DID  As Integer = 0
  Const D_DNAM As Integer = 1
  Const D_PATH As Integer = 2
  Const D_TBLS As Integer = 3
  Const D_RELS As Integer = 4
  Const D_IDXS As Integer = 5
  Const D_DAT  As Integer = 6

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strThisFile = Parse_File(CurrentDb.Name)  ' ** Module Function: modFileUtilities.
  strThisPath = Parse_Path(CurrentDb.Name)  ' ** Module Function: modFileUtilities.

  Set dbsFE = CurrentDb
  With dbsFE

    lngDbs = 0&
    ReDim arr_varDb(D_ELEMS, 0)

    ' ** Get list of databases to document.
    Set rstDB = .OpenRecordset("tblDatabase", dbOpenDynaset, dbConsistent)
    With rstDB

      If .BOF = True And .EOF = True Then
        ' ** Table is empty!
        Stop
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          If Left(![dbs_name], 11) <> "TrustImport" And Left(![dbs_name], 8) <> "TrstXAdm" Then  ' ** Let Trust Import handle these.
            If blnWithJrnlTmp = False And ![dbs_name] = "TAJrnTmp.mdb" Then
              ' ** Skip it.
            Else
              lngDbs = lngDbs + 1&
              lngE = lngDbs - 1&
              ReDim Preserve arr_varDb(D_ELEMS, lngE)
              ' *****************************************************
              ' ** Array: arr_varDB()
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
              arr_varDb(D_DID, lngE) = ![dbs_id]
              arr_varDb(D_DNAM, lngE) = ![dbs_name]
              arr_varDb(D_PATH, lngE) = ![dbs_path]
              arr_varDb(D_TBLS, lngE) = ![dbs_tbl_cnt]
              arr_varDb(D_RELS, lngE) = ![dbs_rel_cnt]
              arr_varDb(D_IDXS, lngE) = ![dbs_idx_cnt]
              arr_varDb(D_DAT, lngE) = ![dbs_datemodified]
            End If
          End If
          If lngX < lngRecs Then .MoveNext
        Next
        .MoveFirst

        For lngX = 0& To (lngDbs - 1&)
          If ![dbs_id] <> arr_varDb(D_DID, lngX) Then
            .FindFirst "[dbs_id] = " & CStr(arr_varDb(D_DID, lngX))
            If .NoMatch = True Then
              Stop
            End If
          End If
          If arr_varDb(D_DNAM, lngX) = strThisFile Then
            strTmp01 = strThisPath
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_DataName Then
            strTmp01 = Parse_Path(CurrentBackendPathFile("m_VD"))  ' ** Module Function: modFileUtilities.
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_ArchDataName Then
            strTmp01 = Parse_Path(CurrentBackendPathFile("m_VA"))  ' ** Module Function: modFileUtilities.
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_AuxDataName Then
            strTmp01 = Parse_Path(CurrentBackendPathFile("m_VX"))  ' ** Module Function: modFileUtilities.
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_RePostDataName Then
            ' ** Leave it as-is.
            strTmp01 = ![dbs_path]
          Else
            Stop
          End If
            If arr_varDb(D_PATH, lngX) <> strTmp01 Then
              .Edit
              ![dbs_path] = strThisPath
              ![dbs_datemodified] = Now()
              .Update
              arr_varDb(D_PATH, lngX) = strThisPath
            End If
        Next
      End If

      .Close
    End With  ' ** rstDB.

    ' ** Empty tblRelation, by specified [dbid].
    Set qdf = dbsFE.QueryDefs("zz_qry_Relation_01")
    With qdf.Parameters
      ![dbidao] = lngThisDbsID
    End With
    qdf.Execute

    .Close
  End With  ' ** dbsFE.
  Set rstDB = Nothing
  Set qdf = Nothing
  Set dbsFE = Nothing

  ' ** Reset the Autonumber field to 1.
  ChangeSeed_Ext "tblRelation"  ' ** Module Function: modAutonumberFieldFuncs.
  ChangeSeed_Ext "tblRelation_Field"  ' ** Module Function: modAutonumberFieldFuncs.

  Set dbsFE = CurrentDb
  Set rstRel = dbsFE.OpenRecordset("tblRelation", dbOpenDynaset, dbConsistent)
  Set rstRelFld = dbsFE.OpenRecordset("tblRelation_Field", dbOpenDynaset, dbAppendOnly)

  For lngX = 0& To (lngDbs - 1&)

    If arr_varDb(D_DNAM, lngX) = strThisFile Then
      Set dbsBE = CurrentDb
      blnThisDbs = True
    Else
      Set wrk = DBEngine.CreateWorkspace("Tmp", "Superuser", TA_SEC, dbUseJet)
      Set dbsBE = wrk.OpenDatabase(arr_varDb(D_PATH, lngX) & LNK_SEP & arr_varDb(D_DNAM, lngX), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
      blnThisDbs = False
    End If
    With dbsBE

      lngRelCnt = 0&
      For Each Rel In .Relations
        If Left(Rel.Table, 4) <> "MSys" And Left(Rel.ForeignTable, 4) <> "MSys" And _
            Left(Rel.Table, 4) <> "~TMP" And Left(Rel.ForeignTable, 4) <> "~TMP" Then  ' ** Skip those pesky system tables!

          ' ** If the dbs_id's don't match, make sure we don't put it where the ForeignTable isn't available.
          ' ** Or, non-matching relationships can only go in the frontend!
          blnFound = False
          For Each tdf In .TableDefs
            If tdf.Name = Rel.Table And IIf(blnThisDbs = False, IIf(tdf.Connect = vbNullString, True, False), True) Then
              blnFound = True
              Exit For
            End If
          Next
          blnFound2 = False
          For Each tdf In .TableDefs
            If tdf.Name = Rel.ForeignTable And IIf(blnThisDbs = False, IIf(tdf.Connect = vbNullString, True, False), True) Then
              blnFound2 = True
              Exit For
            End If
          Next
          If blnFound = False Or blnFound2 = False Then
            If Parse_File(.Name) <> Parse_File(CurrentDb.Name) Then  ' ** Module Functions: modFileUtilities.
              ' ** Don't document this; it doesn't belong!
              blnFound = False
              Debug.Print "'REL ERR: " & Parse_File(.Name) & "  " & Rel.Table & " -> " & Rel.ForeignTable & " IN " & arr_varDb(D_DNAM, lngX)
            End If
          End If

          If blnFound = True Then

            lngRelCnt = lngRelCnt + 1&

            ' ** Check for each table in tblDatabase_Table.
            lngDBID1 = 0&: lngDBID2 = 0&: lngTblID1 = 0&: lngTblID2 = 0&
            For lngY = 1& To 2&
              ' ** Once for each side of the relationship, in case it spans databases.

'IF THIS IS A BACKEND DATABASE, WHAT THE HELL'S IT DOING LOOKING FOR A CONNECT STRING!!!!!!!!!!!!!!!
'AND LATER, I USE THATFILE, WHICH DOESN'T GET SET!!!!!!!!!

              Select Case lngY
              Case 1&
                strThatPath = dbsBE.TableDefs(Rel.Table).Connect
              Case 2&
                strThatPath = dbsBE.TableDefs(Rel.ForeignTable).Connect
              End Select

              strThatFile = vbNullString
              If strThatPath <> vbNullString Then
                strThatPath = Mid$(strThatPath, (InStr(strThatPath, LNK_IDENT) + Len(LNK_IDENT)))
                strThatFile = Parse_File(strThatPath)  ' ** Module Function: modFileUtilities.
                strThatPath = Parse_Path(strThatPath)  ' ** Module Function: modFileUtilities.
              Else
                strThatFile = Parse_File(dbsBE.Name)  ' ** Module Function: modFileUtilities.
              End If

              ' ** tblDatabase_Table, by specified [dbid], [tblnam].
              Set qdf = dbsFE.QueryDefs("zz_qry_Relation_02")
              With qdf.Parameters
                Select Case blnThisDbs
                Case True
                  ' ** This DB.
                  Select Case lngY
                  Case 1&
                    ' ** The relation is local to this DB, but its tables may not be.
                    blnFound = False
                    For Each tdf In dbsBE.TableDefs
                      With tdf
                        If .Name = Rel.Table Then
                          If .Connect = vbNullString Then
                            ' ** Yes, this table is local.
                            blnFound = True
                            lngDBID1 = arr_varDb(D_DID, lngX)
                            Exit For
                          Else
                            ' ** This table is linked.
                            For lngZ = 0& To (lngDbs - 1)
                              If arr_varDb(D_DNAM, lngZ) = Parse_File(.Connect) Then  ' ** Module Function: modFileUtilities.
                                blnFound = True
                                lngDBID1 = arr_varDb(D_DID, lngZ)
                                Exit For
                              End If
                            Next
                          End If
                        End If
                      End With
                    Next
                    If blnFound = False Then
                      Stop
                    End If
                    ![dbid] = lngDBID1  ' ** Query Parameter.
                    If Rel.Table = "tblDataTypeDb1" Then
                      ![tblnam] = "tblDataTypeDb"  ' ** Query Parameter.
                    Else
                      ![tblnam] = Rel.Table  ' ** Query Parameter.
                    End If
                  Case 2&
                    ' ** The relation is local to this DB, but its tables may not be.
                    blnFound = False
                    For Each tdf In dbsBE.TableDefs
                      With tdf
                        If .Name = Rel.ForeignTable Then
                          If .Connect = vbNullString Then
                            ' ** Yes, this table is local.
                            blnFound = True
                            lngDBID2 = arr_varDb(D_DID, lngX)
                            Exit For
                          Else
                            ' ** This table is linked.
                            For lngZ = 0& To (lngDbs - 1)
                              If arr_varDb(D_DNAM, lngZ) = Parse_File(.Connect) Then  ' ** Module Function: modFileUtilities.
                                varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(arr_varDb(D_DID, lngZ)) & " And " & _
                                  "[tbl_name] = '" & Rel.ForeignTable & "'")
                                Select Case IsNull(varTmp00)
                                Case True
                                  ' ** The Foreign Table isn't in the same database as the Table.
                                  varTmp00 = DLookup("[dbs_id]", "tblDatabase_Table", "[tbl_name] = '" & Rel.ForeignTable & "'")
                                  If IsNull(varTmp00) = True Then
                                    Stop
                                  Else
                                    blnFound = True
                                    lngDBID2 = varTmp00
                                  End If
                                Case False
                                  blnFound = True
                                  lngDBID2 = arr_varDb(D_DID, lngZ)
                                End Select
                                Exit For
                              End If
                            Next
                          End If
                        End If
                      End With
                    Next
                    If blnFound = False Then
                      ' ** The ForeignTable was not found in the Table's database.
                      ' ** 2 ways to get it:
                      ' **   1. Come back here to see if the ForeignTable has a Connect string, then get its database from that.
                      ' **   2. Just DLookup() on the table name.
                      varTmp00 = DLookup("[dbs_id]", "tblDatabase_Table", "[tbl_name] = '" & Rel.ForeignTable & "'")
                      If IsNull(varTmp00) = True Then
                        Stop
                      Else
                        blnFound = True
                        lngDBID2 = varTmp00
                      End If
                    End If
                    ![dbid] = lngDBID2  ' ** Query Parameter.
                    If Rel.ForeignTable = "tblDataTypeDb1" Then
                      ![tblnam] = "tblDataTypeDb"  ' ** Query Parameter.
                    Else
                      ![tblnam] = Rel.ForeignTable  ' ** Query Parameter.
                    End If
                  End Select
                Case False
                  ' ** Other DBs.
                  Select Case lngY
                  Case 1&
                    blnFound = False
                    For lngZ = 0& To (lngDbs - 1&)
                      If arr_varDb(D_DNAM, lngZ) = strThatFile Then
                        blnFound = True
                        lngDBID1 = arr_varDb(D_DID, lngZ)
                        ![dbid] = lngDBID1  ' ** Query Parameter.
                        Exit For
                      End If
                    Next
                    If blnFound = False Then
                      Stop
                    End If
                    If Rel.Table = "tblDataTypeDb1" Then
                      ![tblnam] = "tblDataTypeDb"  ' ** Query Parameter.
                    Else
                      ![tblnam] = Rel.Table  ' ** Query Parameter.
                    End If
                  Case 2&
                    blnFound = False
                    For lngZ = 0& To (lngDbs - 1&)
                      If arr_varDb(D_DNAM, lngZ) = strThatFile Then
                        varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = " & CStr(arr_varDb(D_DID, lngZ)) & " And " & _
                          "[tbl_name] = '" & Rel.ForeignTable & "'")
                        Select Case IsNull(varTmp00)
                        Case True
                          ' ** The Foreign Table isn't in the same database as the Table.
                          varTmp00 = DLookup("[dbs_id]", "tblDatabase_Table", "[tbl_name] = '" & Rel.ForeignTable & "'")
                          If IsNull(varTmp00) = True Then
                            Stop
                          Else
                            blnFound = True
                            lngDBID2 = varTmp00
                            ![dbid] = lngDBID2  ' ** Query Parameter.
                          End If
                        Case False
                          blnFound = True
                          lngDBID2 = arr_varDb(D_DID, lngZ)
                          ![dbid] = lngDBID2  ' ** Query Parameter.
                        End Select
                        Exit For
                      End If
                    Next
                    If blnFound = False Then
                      Stop
                    End If
                    If Rel.ForeignTable = "tblDataTypeDb1" Then
                      ![tblnam] = "tblDataTypeDb"  ' ** Query Parameter.
                    Else
                      ![tblnam] = Rel.ForeignTable  ' ** Query Parameter.
                    End If
                  End Select
                End Select  ' ** blnThisDbs.
              End With  ' ** Parameters.

              Set rstDB = qdf.OpenRecordset()
              With rstDB
                If .BOF = True And .EOF = True Then
                  ' ** Add it.
                  Stop
                Else
                  Select Case lngY
                  Case 1&
                    lngTblID1 = ![tbl_id]
                  Case 2&
                    lngTblID2 = ![tbl_id]
                  End Select
                End If
                .Close
              End With  ' ** rstDB.
              Set rstDB = Nothing
              Set qdf = Nothing

              If (lngY = 1& And (lngDBID1 = 0& Or lngTblID1 = 0&)) Or (lngY = 2& And (lngDBID2 = 0& Or lngTblID2 = 0&)) Then
                Stop
              End If

            Next  ' ** lngY.

If lngDBID2 = 0& Then
Stop
End If

            ' ** Add the relation.
            If (Rel.Attributes And dbRelationInherited) > 0 Then
              ' ** It's a non-local relationship
            Else
              With rstRel
                .AddNew
                ![dbs_id_asof] = lngThisDbsID
                ![dbs_id1] = lngDBID1
                ![tbl_id1] = lngTblID1  'rel.Table
                ![dbs_id2] = lngDBID2
                ![tbl_id2] = lngTblID2  'rel.ForeignTable
                strTmp01 = Rel.Name
                intPos1 = InStr(strTmp01, "].")
                If intPos1 > 0 Then
                  ' ** An inherited relationship?
                  ' ** [C:\VictorGCS_Clients\TrustAccountant\NewWorking\TrustDta.mdb].{10931695-4206-4DA4-8D2D-3221A47C277E}  '## OK
                  strTmp01 = Mid$(strTmp01, (intPos1 + 2))
                End If
                ![rel_name] = strTmp01  ' ** Maximum saved length seems to be 61, so must be unique within that.
                ![rel_attributes] = Rel.Attributes
                ![rel_asof_listed] = blnThisDbs
                ![rel_datemodified] = Now()
On Error Resume Next
                .Update
If ERR.Number <> 0 Then
On Error GoTo 0
  Beep
  Debug.Print "'REL DUPED? " & Rel.Name & "  IN  " & arr_varDb(D_DNAM, lngX)
  DoEvents
Else
On Error GoTo 0
                .Bookmark = .LastModified
                lngRelID = ![rel_id]
End If
              End With  ' ** rstRel.
If lngRelID = 0& Then
  Stop
Else
  rstRel.FindFirst "[rel_id] = " & CStr(lngRelID)
  If rstRel.NoMatch = True Then
    Stop
  End If
End If

If lngRelID > 0& Then
              ' ** Add the relation fields.
              intRelOrd = 0
              For Each fld In Rel.Fields
                With rstRelFld
                  intRelOrd = intRelOrd + 1
                  .AddNew
                  ![dbs_id_asof] = lngThisDbsID
                  ![dbs_id] = lngDBID1
                  ![tbl_id] = lngTblID1
                  ![rel_id] = lngRelID
                  ![fld_id] = DLookup("[fld_id]", "tblDatabase_Table_Field", "[dbs_id] = " & CStr(lngDBID1) & _
                    " And [tbl_id] = " & CStr(lngTblID1) & " And [fld_name] = '" & fld.Name & "'")
                  ![relfld_order] = intRelOrd
                  ![relfld_name] = fld.Name
                  ![relfld_foreignname] = fld.ForeignName
                  ![relfld_datemodified] = Now()
                  .Update
                End With  ' ** rstRelFld.
              Next  ' ** fld.

              With rstRel
                .FindFirst "[rel_id] = " & CStr(lngRelID)
                If .NoMatch = False Then
                  .Edit
                  ![rel_fld_cnt] = CLng(intRelOrd)
                  ![rel_datemodified] = Now()
                  .Update
                End If
              End With  ' ** rstRel

End If

            End If  ' ** Local relationships only.

          End If  ' ** blnFound.

        End If  ' ** MSys.
      Next  ' ** rel.

      .Close
    End With  ' ** dbsBE.
    Set dbsBE = Nothing

    If arr_varDb(D_DNAM, lngX) <> strThisFile Then
      wrk.Close
      Set wrk = Nothing
    End If

    ' ** Update tblDatabase, for dbs_rel_cnt, by specified [dbid], [relcnt].
    Set qdf = dbsFE.QueryDefs("zz_qry_Relation_03")
    With qdf.Parameters
      ![dbid] = arr_varDb(D_DID, lngX)
      ![relcnt] = lngRelCnt
    End With
    qdf.Execute
    Set qdf = Nothing

    ' ** Update zz_qry_Relation_04e (tblDatabase_Table, with DLookups() to zz_qry_Relation_04d
    ' ** (zz_qry_Relation_04c (Union of zz_qry_Relation_04a (tblRelation, grouped by dbs_id1,
    ' ** tbl_id1, with cnt), zz_qry_Relation_04b (tblRelation, grouped by dbs_id, tbl_id2,
    ' ** with cnt)), grouped and summed by dbs_id, tbl_id, with cnt), by specified [dbid]).
    Set qdf = dbsFE.QueryDefs("zz_qry_Relation_04f")
    With qdf.Parameters
      ![dbid] = arr_varDb(D_DID, lngX)
    End With
    qdf.Execute
    Set qdf = Nothing

  Next
  rstRelFld.Close
  rstRel.Close
  dbsFE.Close

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  DoBeeps  ' ** Module Function: modWindowFunctions.
  Debug.Print "'FINISHED!  " & THIS_PROC & "()"

  Beep

  Set fld = Nothing
  Set Rel = Nothing
  Set tdf = Nothing
  Set rstDB = Nothing
  Set rstRel = Nothing
  Set rstRelFld = Nothing
  Set qdf = Nothing
  Set dbsFE = Nothing
  Set dbsBE = Nothing
  Set wrk = Nothing

  Rel_Doc = blnRetValx

End Function

Public Function Rel_Regen() As Boolean
' ** Regenerate all relationships.
' ** DOES NOT INCLUDE TAJrnTmp.mdb, tblRelation_View, tblRelation_View_Window.
' ** Maximum name length: 64 chars.

  Const THIS_PROC As String = "Rel_Regen"

  Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim fld As DAO.Field, tdf As DAO.TableDef
  Dim rel1 As DAO.Relation, rel2 As DAO.Relation
  Dim lngRels As Long, arr_varRel As Variant
  Dim lngFlds As Long, arr_varFld As Variant
  Dim lngRelTbls As Long, arr_varRelTbl As Variant
  Dim lngRelCnt As Long, strRel As String
  Dim strLastTblName As String, strLastRelName As String
  Dim lngJTmpDbsID As Long
  Dim blnFound As Boolean, blnJustList As Boolean, blnJustSysNames As Boolean, blnSkip As Boolean
  Dim intPos1 As Integer
  Dim lngTmp00 As Long, strTmp01 As String, strTmp02 As String, lngTmp03 As Long, blnTmp04 As Boolean
  Dim lngX As Long, lngY As Long, lngZ As Long

  ' ** Array: arr_varRel().
  Const R_ASID  As Integer = 0
  Const R_ASNAM As Integer = 1
  Const R_RID   As Integer = 2
  Const R_DID1  As Integer = 3
  Const R_DNAM1 As Integer = 4
  Const R_TID1  As Integer = 5
  Const R_TNAM1 As Integer = 6
  Const R_DID2  As Integer = 7
  Const R_DNAM2 As Integer = 8
  Const R_TID2  As Integer = 9
  Const R_TNAM2 As Integer = 10
  Const R_RNAM  As Integer = 11
  Const R_ATTR  As Integer = 12
  Const R_FLDS  As Integer = 13
  Const R_FARR  As Integer = 14

  ' ** Array: arr_varFld().
  Const FDF_RID   As Integer = 0
  Const FDF_DID   As Integer = 1
  Const FDF_DNAM  As Integer = 2
  Const FDF_TID1  As Integer = 3
  Const FDF_TNAM1 As Integer = 4
  Const FDF_TID2  As Integer = 5
  Const FDF_TNAM2 As Integer = 6
  Const FDF_FORD  As Integer = 7
  Const FDF_RFID  As Integer = 8
  Const FDF_FID1  As Integer = 9
  Const FDF_FNAM1 As Integer = 10
  Const FDF_FID2  As Integer = 11
  Const FDF_FNAM2 As Integer = 12
  Const FDF_DID2  As Integer = 14

  ' ** Array: arr_varRelTbl().
  Const RT_DID   As Integer = 0
  Const RT_TID   As Integer = 1
  Const RT_TNAM  As Integer = 2
  Const RT_ASID  As Integer = 3
  Const RT_ASNAM As Integer = 4
  Const RT_RELS  As Integer = 5
  Const RT_ARR_F As Integer = 6
  Const RT_ARR_L As Integer = 7
  Const RT_DID2  As Integer = 8

  blnRetValx = True

  blnJustSysNames = True  ' ** True: Just list/regenerate relationships with a system name.
  blnJustList = False  ' ** True: Just list the array results, don't regenerate them.

  Set dbs = CurrentDb
  With dbs

    ' ** Get a list of currently documented relationships

    Select Case blnJustSysNames
    Case True
      ' ** JUST SYSTEM NAME:
      ' ** Just relationships having system-name, with table names.
      Set qdf = .QueryDefs("zz_qry_Relation_23b")
    Case False
      ' ** WITHOUT NON-CONTIGUOUS:
      ' **   Also without tblRelation_View, tblRelation_View_Window, or anything in TAJrnTmp.mdb.
      ' ** All relationships, with table names (w/o TAJrnTmp.mdb).
      Set qdf = .QueryDefs("zz_qry_Relation_23a")  ' ** NEEDS TO BE IN dbs_id1 ORDER!
    End Select

    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRels = .RecordCount
      .MoveFirst
      .Sort = "dbs_id1, tbl_name1, tbl_name2"
      Set rst2 = .OpenRecordset
      rst2.MoveLast
      lngRels = rst2.RecordCount
      rst2.MoveFirst
      arr_varRel = rst2.GetRows(lngRels)
      ' ***************************************************
      ' ** Array: arr_varRel()
      ' **
      ' **   Field  Element  Name              Constant
      ' **   =====  =======  ================  ==========
      ' **     1       0     dbs_id_asof       R_ASID
      ' **     2       1     dbs_name_asof     R_ASNAM
      ' **     3       2     rel_id            R_RID
      ' **     4       3     dbs_id1           R_DID1
      ' **     5       4     dbs_name1         R_DNAM1
      ' **     6       5     tbl_id1           R_TID1
      ' **     7       6     tbl_name1         R_TNAM1
      ' **     8       7     dbs_id2           R_DID2
      ' **     9       8     dbs_name2         R_DNAM2
      ' **    10       9     tbl_id2           R_TID2
      ' **    11      10     tbl_name2         R_TNAM2
      ' **    12      11     rel_name          R_RNAM
      ' **    13      12     rel_attributes    R_ATTR
      ' **    14      13     rel_fld_cnt       R_FLDS
      ' **    15      14     rel_fld_arr       R_FARR
      ' **    16      15     DontEnforce
      ' **
      ' ***************************************************
      rst2.Close
      .Close
    End With  ' ** rst1.
    Set rst1 = Nothing
    Set rst2 = Nothing

    ' ** Get lists of relationship fields
    For lngX = 0& To (lngRels - 1&)

      Select Case blnJustSysNames
      Case True
        ' ** JUST SYSTEM NAME:
        ' ** zz_qry_Relation_13t (zz_qry_Relation_13s (tblRelation, linked to tblRelation_Field, with add'l fields,
        ' ** just relationships having system-name), with add'l foreign table fields), by specified [relid].
        'Set qdf = .QueryDefs("zz_qry_Relation_13u")
        Set qdf = .QueryDefs("zz_qry_Relation_25c")
      Case False
        ' ** WITHOUT NON-CONTIGUOUS:
        ' **   Also without tblRelation_View, tblRelation_View_Window, or anything in TAJrnTmp.mdb.
        ' ** zz_qry_Relation_13b (zz_qry_Relation_13a (tblRelation, linked to tblRelation_Field,
        ' ** with add'l fields), with add'l foreign table fields), by specified [relid].
        Set qdf = .QueryDefs("zz_qry_Relation_24c")  ' ** NEEDS TO BE IN dbs_id ORDER!
      End Select

      With qdf.Parameters
        ![relid] = CLng(arr_varRel(R_RID, lngX))
      End With
      Set rst1 = qdf.OpenRecordset
      With rst1
        .MoveLast
        lngFlds = .RecordCount
        .MoveFirst
        .Sort = "dbs_id, tbl_name1, tbl_name2, relfld_order, fld_name1, fld_name2"
        Set rst2 = .OpenRecordset
        rst2.MoveLast
        lngFlds = rst2.RecordCount
        rst2.MoveFirst
        arr_varFld = rst2.GetRows(lngFlds)
        ' **************************************************
        ' ** Array: arr_varFld()
        ' **
        ' **   Field  Element  Name            Constant
        ' **   =====  =======  ==============  ===========
        ' **     1       0     rel_id          FDF_RID
        ' **     2       1     dbs_id          FDF_DID
        ' **     3       2     dbs_name        FDF_DNAM
        ' **     4       3     tbl_id1         FDF_TID1
        ' **     5       4     tbl_name1       FDF_TNAM1
        ' **     6       5     tbl_id2         FDF_TID2
        ' **     7       6     tbl_name2       FDF_TNAM2
        ' **     8       7     relfld_order    FDF_FORD
        ' **     9       8     relfld_id       FDF_RFID
        ' **    10       9     fld_id1         FDF_FID1
        ' **    11      10     fld_name1       FDF_FNAM1
        ' **    12      11     fld_id2         FDF_FID2
        ' **    13      12     fld_name2       FDF_FNAM2
        ' **    14      13     DontEnforce
        ' **    15      14     dbs_id2         FDF_DID2
        ' **
        ' **************************************************
      End With
      arr_varRel(R_FARR, lngX) = arr_varFld
      Set qdf = Nothing
      lngFlds = 0&
      arr_varFld = Empty
    Next
    Set rst1 = Nothing
    Set rst2 = Nothing

    ' ** Get list of tables with relationships
    Select Case blnJustSysNames
    Case True
      ' ** JUST SYSTEM NAME:
      ' ** tblRelation, grouped by tbl_id, just relationships having system-name, with rel_cnt.
      Set qdf = .QueryDefs("zz_qry_Relation_26b")
    Case False
      ' ** WITHOUT NON-CONTIGUOUS:
      ' **   Also without tblRelation_View, tblRelation_View_Window, or anything in TAJrnTmp.mdb.
      ' ** tblRelation, grouped by tbl_id, with rel_cnt.
      Set qdf = .QueryDefs("zz_qry_Relation_26a")  ' ** NEEDS TO BE IN dbs_id1 ORDER!
    End Select

    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRelTbls = .RecordCount
      .MoveFirst
      .Sort = "dbs_id1, tbl_name"
      Set rst2 = .OpenRecordset
      rst2.MoveLast
      lngRelTbls = rst2.RecordCount
      rst2.MoveFirst
      arr_varRelTbl = rst2.GetRows(lngRelTbls)
      ' **************************************************
      ' ** Array: arr_varRelTbl()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ==========
      ' **     1       0     dbs_id           RT_DID
      ' **     2       1     tbl_id           RT_TID
      ' **     3       2     tbl_name         RT_TNAM
      ' **     4       3     dbs_id_asof      RT_ASID
      ' **     5       4     dbs_name_asof    RT_ASNAM
      ' **     6       5     rel_cnt          RT_RELS
      ' **     7       6     rel_arr_first    RT_ARR_F
      ' **     8       7     rel_arr_last     RT_ARR_L
      ' **     9       8     dbs_id2          RT_DID2
      ' **
      ' **************************************************
      .Close
    End With
    Set rst1 = Nothing
    Set rst2 = Nothing

    lngJTmpDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TAJrnTmp.mdb'")

    ' ** Put the first and last arr_varRel() array elements, for each table, into the arr_varRelTbl() array for that table.
    strLastTblName = vbNullString
    For lngX = 0& To (lngRelTbls - 1&)
      'If (((arr_varRelTbl(RT_DID, lngX) = lngJTmpDbsID) And (arr_varRelTbl(RT_TID, lngX) = 439&) And _
      '    (arr_varRelTbl(RT_TNAM, lngX) = "tblPricing_Import")) Or _
      '    ((arr_varRelTbl(RT_DID, lngX) = lngJTmpDbsID) And (arr_varRelTbl(RT_TID, lngX) = 650&) And _
      '    (arr_varRelTbl(RT_DID2, lngX) = lngJTmpDbsID) And (arr_varRelTbl(RT_TNAM, lngX) = "tblRelation_View")) Or _
      '    ((arr_varRelTbl(RT_DID, lngX) = lngJTmpDbsID) And (arr_varRelTbl(RT_TID, lngX) = 342&) And _
      '    (arr_varRelTbl(RT_TNAM, lngX) = "zz_tbl_RePost_Account")) Or _
      '    ((arr_varRelTbl(RT_DID, lngX) = lngJTmpDbsID) And (arr_varRelTbl(RT_TID, lngX) = 348&) And _
      '    (arr_varRelTbl(RT_TNAM, lngX) = "zz_tbl_RePost_MasterAsset"))) Then
      '  For lngY = 0& To (lngRels - 1&)
      '    If ((arr_varRel(R_TNAM1, lngY) = arr_varRelTbl(RT_TNAM, lngX)) And _
      '        (arr_varRel(R_DID2, lngY) = arr_varRelTbl(RT_DID2, lngX))) Then
      '      arr_varRelTbl(RT_ARR_F, lngX) = lngY
      '      Exit For
      '    End If
      '  Next
      '  If arr_varRelTbl(RT_TNAM, lngX) = "tblPricing_Import" Then
      '    arr_varRelTbl(RT_ARR_L, lngX - 1&) = 69&
      '  Else
      '    arr_varRelTbl(RT_ARR_L, lngX - 1&) = (lngY - 1&)
      '  End If
      '  strLastTblName = arr_varRelTbl(RT_TNAM, lngX)
      '  If lngX = (lngRelTbls - 1&) Then
      '    arr_varRelTbl(RT_ARR_L, lngX) = CLng(UBound(arr_varRel, 2))
      '  End If
      'Else
        For lngY = 0& To (lngRels - 1&)
          If ((arr_varRel(R_TNAM1, lngY) = arr_varRelTbl(RT_TNAM, lngX)) And _
              (arr_varRel(R_DID2, lngY) = arr_varRelTbl(RT_DID2, lngX))) Then
            If arr_varRelTbl(RT_TNAM, lngX) <> strLastTblName Then
              arr_varRelTbl(RT_ARR_F, lngX) = lngY
              If lngX > 0& Then
                arr_varRelTbl(RT_ARR_L, lngX - 1&) = lngY - 1&
              End If
              strLastTblName = arr_varRelTbl(RT_TNAM, lngX)
            End If
          End If
        Next
        If lngX = (lngRelTbls - 1&) Then
          arr_varRelTbl(RT_ARR_L, lngX) = CLng(UBound(arr_varRel, 2))
        End If
      'End If
    Next  ' ** lngX.

    .Close
  End With

'For lngX = 0& To (lngRels - 1&)
'  Debug.Print "'" & Left(CStr(lngX + 1&) & "  ", 2) & ": " & CStr(arr_varRel(R_DID2, lngX)) & "  " & _
'    arr_varRel(R_TNAM1, lngX) & " -> " & arr_varRel(R_TNAM2, lngX)
'Next
'lngTmp00 = 0&
'For lngX = 0& To (lngRelTbls - 1&)
'  If Len(arr_varRelTbl(RT_TNAM, lngX)) > lngTmp00 Then
'    lngTmp00 = Len(arr_varRelTbl(RT_TNAM, lngX))
'  End If
'Next
'For lngX = 0& To (lngRelTbls - 1&)
'  Debug.Print "'" & Left(CStr(lngX + 1) & "  ", 2) & ": " & CStr(arr_varRelTbl(RT_DID2, lngX)) & "  " & Left(arr_varRelTbl(RT_TNAM, lngX) & Space(lngTmp00), lngTmp00) & "  " & _
'    CStr(arr_varRelTbl(RT_ARR_F, lngX)) & " - " & CStr(arr_varRelTbl(RT_ARR_L, lngX))
'Next

blnSkip = False
If blnSkip = False Then

  Select Case blnJustList
  Case True

    For lngX = 0& To (lngRelTbls - 1&)
      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      Debug.Print "'" & Left$(CStr(arr_varRelTbl(RT_ARR_F, lngX)) & "   ", 3) & " - " & _
        Left$(CStr(arr_varRelTbl(RT_ARR_L, lngX)) & "   ", 3) & "  " & arr_varRelTbl(RT_TNAM, lngX)
    Next

  Case False  ' ** blnJustList.

    strLastRelName = vbNullString
    blnTmp04 = False

    ' ** For each table with one or more relationships.
    For lngX = 0& To (lngRelTbls - 1&)

      ' **************************************************
      ' ** Array: arr_varRelTbl()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ==========
      ' **     1       0     dbs_id           RT_DID
      ' **     2       1     tbl_id           RT_TID
      ' **     3       2     tbl_name         RT_TNAM
      ' **     4       3     dbs_id_asof      RT_ASID
      ' **     5       4     rel_cnt          RT_RELS
      ' **     6       5     rel_arr_first    RT_ARR_F
      ' **     7       6     rel_arr_last     RT_ARR_L
      ' **
      ' **************************************************

      strTmp01 = DLookup("[dbs_name]", "tblDatabase", "[dbs_id] = " & CStr(arr_varRelTbl(RT_DID, lngX)))
      If strTmp01 = Parse_File(CurrentDb.Name) Then  ' ** Module Function: modFileUtilities.
        Set wrk = DBEngine.Workspaces(0)
        Set dbs = wrk.Databases(0)
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
        ' ** If this is a database linked to here, get its linked path.
        ' ** If it doesn't show up here, take the path from tblDatabase.
        strTmp02 = vbNullString
        For Each tdf In CurrentDb.TableDefs
          With tdf
            If InStr(.Connect, strTmp01) > 0 Then
              strTmp02 = Parse_Path(Mid$(.Connect, (InStr(.Connect, LNK_IDENT) + Len(LNK_IDENT))))  ' ** Module Function: modFileUtilities.
              Exit For
            End If
          End With
        Next
        If strTmp02 = vbNullString Then
          strTmp02 = DLookup("[dbs_path]", "tblDatabase", "[dbs_id] = " & CStr(arr_varRelTbl(RT_DID, lngX)))
          If Right$(strTmp02, 1) = LNK_SEP Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
        End If
        Set dbs = wrk.OpenDatabase(strTmp02 & LNK_SEP & strTmp01, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
      End If

      With dbs

If .TableDefs(arr_varRelTbl(RT_TNAM, lngX)).Connect <> vbNullString Then
  Debug.Print "'REL ON LINKED! " & arr_varRelTbl(RT_TNAM, lngX)
Else

        ' ** Delete all relationships for this table.
        .Relations.Refresh
        lngTmp00 = .Relations.Count
        For lngY = (lngTmp00 - 1&) To 0 Step -1&
          If .Relations(lngY).Table = arr_varRelTbl(RT_TNAM, lngX) Then
            .Relations.Delete .Relations(lngY).Name
            .Relations.Refresh
          End If
        Next

        ' ** Now re-create the relationships.
        For lngY = arr_varRelTbl(RT_ARR_F, lngX) To arr_varRelTbl(RT_ARR_L, lngX)

          ' ** Define the relationship.
          lngFlds = arr_varRel(R_FLDS, lngY)
          arr_varFld = arr_varRel(R_FARR, lngY)
          If lngFlds <> (UBound(arr_varFld, 2) + 1&) Then
            Stop
          End If

          ' ** Rel_Doc(), above, records the Attributes from this database.
          ' ** Therefore, linked tables include the dbRelationInherited attribute,
          ' ** which we don't want when re-creating the relationship.
          lngTmp00 = arr_varRel(R_ATTR, lngY)
          If (lngTmp00 And dbRelationInherited) > 0 Then
            lngTmp00 = lngTmp00 - dbRelationInherited
          End If

          If (arr_varRel(R_TNAM1, lngY) & arr_varRel(R_TNAM2, lngY)) <> strLastRelName Then
            lngTmp03 = 1&
            strTmp01 = (arr_varRel(R_TNAM1, lngY) & arr_varRel(R_TNAM2, lngY))
            strLastRelName = strTmp01
          Else
            lngTmp03 = lngTmp03 + 1&
            strTmp01 = strLastRelName & CStr(lngTmp03)
          End If

          ' ** Cross-database relationships should be only used here, in the frontend!
          blnFound = True
          If arr_varRel(R_DID1, lngY) <> arr_varRel(R_DID2, lngY) Then
            'Debug.Print "'DBS: " & Parse_File(dbs.Name) & "  DB1: " & CStr(arr_varRel(R_DID1, lngY)) & "  DB2: " & CStr(arr_varRel(R_DID2, lngY))
            If Parse_File(dbs.Name) <> Parse_File(CurrentDb.Name) Then
              blnFound = False
            End If
          End If

          If blnFound = True Then
If Len(strTmp01) > 64 Then
Stop

End If
            Set rel1 = .CreateRelation(strTmp01, arr_varRel(R_TNAM1, lngY), arr_varRel(R_TNAM2, lngY), lngTmp00)
            For lngZ = 0& To (lngFlds - 1&)
              Set fld = rel1.CreateField(arr_varFld(FDF_FNAM1, lngZ))
              fld.ForeignName = arr_varFld(FDF_FNAM2, lngZ)
              rel1.Fields.Append fld
              Set fld = Nothing
            Next

            ' ** Add the fields.
            lngZ = 0&
            Do While True
              blnFound = False
              lngZ = lngZ + 1&
              For Each rel2 In .Relations
                With rel2
                  If rel2.Name = rel1.Name Then
                    blnFound = True
                    Exit For
                  End If
                End With
              Next
              If blnFound = True Then
                strTmp01 = rel1.Name
                strTmp02 = strTmp01
                If lngZ = 1& Then
                  strTmp02 = strTmp01 & CStr(lngZ + 1&)
                Else
                  strTmp02 = Left$(strTmp01, (Len(strTmp01) - 1))
                  strTmp02 = strTmp02 & CStr(lngZ + 1&)
                End If
                rel1.Name = strTmp02
              Else
                Exit Do
              End If
            Loop

            .Relations.Append rel1
            .Relations.Refresh

            Set rel2 = Nothing
            Set rel2 = Nothing

          End If  ' ** blnFound.

        Next  ' ** lngY.

End If

        .Close
      End With  ' ** dbs.
      wrk.Close

    Next  ' ** lngX.

  End Select  ' ** blnJustList.
End If  ' ** blnSkip.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set fld = Nothing
  Set tdf = Nothing
  Set rel1 = Nothing
  Set rel2 = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set wrk = Nothing

  Rel_Regen = blnRetValx

End Function

Public Function Rel_Find() As Boolean

  Const THIS_PROC As String = "Rel_Find"

  Dim dbs As DAO.Database, Rel As DAO.Relation
  Dim lngRels As Long, arr_varRel() As Variant
  Dim strTmp01 As String
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varRel().
  Const R_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const R_NAM As Integer = 0

  blnRetValx = True

  lngRels = 0&
  ReDim arr_varRel(R_ELEMS, 0)

  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{192D707E-3234-4EBE-AFEB-D0944590B69F}"
'1.    REL NAME: {192D707E-3234-4EBE-AFEB-D0944590B69F} VS. {078F5DD8-F844-4E3F-8758-FD279E4307E6}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{8FB51C62-98DB-4F4F-87F4-51E49F76B33E}"
'2.    REL NAME: {8FB51C62-98DB-4F4F-87F4-51E49F76B33E} VS. {07ED4229-4692-4469-BCD1-C7D32F82D1C8}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{66CD0D7D-E8D7-4F7C-8FD5-EAE86A0410BD}"
'3.    REL NAME: {66CD0D7D-E8D7-4F7C-8FD5-EAE86A0410BD} VS. {16C57B2D-9CA0-45D6-91A7-35E9F9B6992C}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{1686A562-0FB4-4AAA-87C5-11403775E0F8}"
'4.    REL NAME: {1686A562-0FB4-4AAA-87C5-11403775E0F8} VS. {1A60E3D6-3728-44BD-A966-8B84DD081E9E}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{F2B54758-722E-4002-AAD4-D8E3A0838E29}"
'5.    REL NAME: {F2B54758-722E-4002-AAD4-D8E3A0838E29} VS. {20DE711B-80F3-4104-9CD9-4CCEA9A09395}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{52142AE8-34A2-45E4-A708-6D35C8067B50}"
'6.    REL NAME: {52142AE8-34A2-45E4-A708-6D35C8067B50} VS. {2EAD89C9-7AA6-4F31-A0A6-A17EFC37CFC7}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{AF297DC5-64EC-4A88-94EA-BB6BC27BB801}"
'7.    REL NAME: {AF297DC5-64EC-4A88-94EA-BB6BC27BB801} VS. {3F3780FA-6FF1-4297-B157-94E42A5122D6}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{CE3D3C13-C663-42DD-B1DC-6F3FC02DDC50}"
'8.    REL NAME: {CE3D3C13-C663-42DD-B1DC-6F3FC02DDC50} VS. {4D819915-F224-4DD0-B49E-0F594709F5E5}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{908F6029-4D49-45F7-B9F1-3C59340C6A8C}"
'9.    REL NAME: {908F6029-4D49-45F7-B9F1-3C59340C6A8C} VS. {65964CB8-8BC5-418A-B7DB-D8A97613D2E2}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{84CBF3E6-8677-4BE5-97C1-03013CC6B5AB}"
'10.   REL NAME: {84CBF3E6-8677-4BE5-97C1-03013CC6B5AB} VS. {9D9CCAC9-C628-4084-AC04-8E9A73526C66}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{1987E810-C423-41E6-8BF6-D2DA8D21F67A}"
'11.   REL NAME: {1987E810-C423-41E6-8BF6-D2DA8D21F67A} VS. {C416DD5C-7185-420F-91AE-A3968852704B}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{D5FBAD6F-69E9-452E-A43E-FFBBDC2888EE}"
'12.   REL NAME: {D5FBAD6F-69E9-452E-A43E-FFBBDC2888EE} VS. {D91B1324-861A-4684-99A9-AB84A1BF5155}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{B61414BF-D709-4A7A-8EB1-C2C5432F731A}"
'13.   REL NAME: {B61414BF-D709-4A7A-8EB1-C2C5432F731A} VS. {F107371B-03AD-4F0E-848A-AA8F1CC3D660}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{75314DFE-B06B-4F03-945C-881A5CC6DF30}"
'14.   REL NAME: {75314DFE-B06B-4F03-945C-881A5CC6DF30} VS. {F70E7DEA-C180-4609-8453-F08A19F8C6B0}
  lngRels = lngRels + 1&
  lngE = lngRels - 1&
  ReDim Preserve arr_varRel(R_ELEMS, lngE)
  arr_varRel(R_NAM, lngE) = "{F067A0E4-DE9C-47A3-965A-9366629D99D1}"
'15.   REL NAME: {F067A0E4-DE9C-47A3-965A-9366629D99D1} VS. {F7562173-4366-419C-AE8B-24EAB22F1CE2}

  Set dbs = CurrentDb
  With dbs
For lngX = 0& To (lngRels - 1&)
    For Each Rel In .Relations
      strTmp01 = vbNullString
      With Rel
        If .Name = arr_varRel(R_NAM, lngX) Then
        'If .ForeignTable = "tblRelation_View" Then
          Debug.Print "'" & CStr(lngX + 1&) & ". REL: " & .Name & "  TBL1: " & .Table & "  TBL2: " & .ForeignTable
          Debug.Print "' ATTR: " & CStr(.Attributes)
          If Rel_Attr(.Attributes, "dbRelationDontEnforce") = True Then strTmp01 = strTmp01 & "dbRelationDontEnforce, "
          If Rel_Attr(.Attributes, "dbRelationEnforce") = True Then strTmp01 = strTmp01 & "dbRelationEnforce, "
          If Rel_Attr(.Attributes, "dbRelationInherited") = True Then strTmp01 = strTmp01 & "dbRelationInherited, "
          If Rel_Attr(.Attributes, "dbRelationUpdateCascade") = True Then strTmp01 = strTmp01 & "dbRelationUpdateCascade, "
          If Rel_Attr(.Attributes, "dbRelationDeleteCascade") = True Then strTmp01 = strTmp01 & "dbRelationDeleteCascade, "
          If Rel_Attr(.Attributes, "dbRelationLeft") = True Then strTmp01 = strTmp01 & "dbRelationLeft, "
          If Rel_Attr(.Attributes, "dbRelationRight") = True Then strTmp01 = strTmp01 & "dbRelationRight, "
          If Rel_Attr(.Attributes, "dbRelationUnique") = True Then strTmp01 = strTmp01 & "dbRelationUnique, "
          strTmp01 = Trim$(strTmp01)
          If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
          Debug.Print "'   " & strTmp01
          Exit For
        End If
      End With
    Next
Next
    .Close
  End With

  Beep

  Set Rel = Nothing
  Set dbs = Nothing

  Rel_Find = blnRetValx

End Function

Public Function Rel_Create() As Boolean
' ** Create a Relationship manually; if the table names are too long, for example.

  Const THIS_PROC As String = "Rel_Create"

  Dim dbs As DAO.Database, Rel As DAO.Relation

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs
    Set Rel = .CreateRelation("tblAssetTypeNew_AssetTypetblAssetTypeNew_AssetTypeGrouping", _
      "zz_tbl_AssetTypeNew_AssetType", "zz_tbl_AssetTypeNew_AssetTypeGrouping", _
      dbRelationDeleteCascade + dbRelationUpdateCascade + dbRelationUnique)
    With Rel
      .Fields.Append .CreateField("assettype", dbText)
      .Fields![assettype].ForeignName = "assettype"
    End With
    .Relations.Append Rel
    .Close
  End With

  Set Rel = Nothing
  Set dbs = Nothing

  Beep

  Rel_Create = blnRetValx

End Function

Public Function Rel_List() As Boolean

  Const THIS_PROC As String = "Rel_List"

  Dim dbs As DAO.Database, Rel As DAO.Relation
  Dim lngRels As Long, lngMaxLen As Long
  Dim lngX As Long

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs
    lngRels = .Relations.Count
    Debug.Print "'RELS: " & CStr(lngRels)
    DoEvents
    lngMaxLen = 0&
    For lngX = 0& To (lngRels - 1&)
      If Len(.Relations(lngX).Table) > lngMaxLen Then lngMaxLen = Len(.Relations(lngX).Table)
    Next
    For lngX = 0& To (lngRels - 1&)
      Set Rel = .Relations(lngX)
      With Rel
        Debug.Print "'" & Left(CStr(lngX + 1&) & "   ", 3) & " " & Left(.Table & Space(lngMaxLen), lngMaxLen) & "  ->  " & .ForeignTable
      End With
      Set Rel = Nothing
      If (lngX + 1&) Mod 100& = 0 Then
        Stop
      End If
    Next
    .Close
  End With

'RELS: 392
'1   m_REVCODE_TYPE                     ->  TaxCode
'2   TaxCode                            ->  journal
'3   RecurringItems                     ->  journal
'4   masterasset                        ->  asset
'5   journal                            ->  tblJournal_Memo
'6   tblPricing_AppraiseColumnIDQuote   ->  tblPricing_AppraiseColumn
'7   masterasset                        ->  ActiveAssets
'8   tblCheckMemoType                   ->  tblCheckMemo
'9   AssetType                          ->  masterasset
'10  tblCheckReconcileSourceType        ->  tblCheckReconcile_Item
'11  tblPricing_Import                  ->  tblPricing_FileType
'12  Schedule                           ->  account
'13  tblCheckReconcileEntryType         ->  tblCheckReconcile_Item
'14  tblCheckReconcile_Account          ->  tblCheckReconcile_Check
'15  RecurringItems                     ->  ledger
'16  adminofficer                       ->  account
'17  masterasset                        ->  tblPricing_MasterAsset_History
'18  Location                           ->  ledger
'19  m_REVCODE                          ->  ledger
'20  m_REVCODE                          ->  journal
'21  journaltype                        ->  ledger
'22  tblDataTypeDb1                     ->  tblPricing_AppraiseItemType
'23  TaxCode_Type                       ->  TaxCode
'24  accounttype                        ->  account
'25  TaxCode                            ->  ledger
'26  journal                            ->  tblJournal_Import
'27  tblCheckReconcile_Account          ->  tblCheckReconcile_Item
'28  tblPricing_AppraiseRowType         ->  tblPricing_AppraiseColumn
'29  account                            ->  Balance
'30  RecurringType                      ->  RecurringItems
'31  tblPricing_AppraiseRowType         ->  tblPricing_AppraiseFile
'32  account                            ->  ledger
'33  HiddenType                         ->  LedgerHidden
'34  tblPricing_AppraiseFile            ->  tblPricing_AppraiseColumn
'35  tblPricing_AppraiseSectionType     ->  tblPricing_AppraiseColumnDataType
'36  tblPricing_AppraiseColumnDataType  ->  tblPricing_AppraiseColumn
'37  AssetType                          ->  AssetTypeGrouping
'38  Location                           ->  ActiveAssets
'39  accounttype                        ->  AccountTypeGrouping
'40  account                            ->  asset
'41  InvestmentObjective                ->  account
'42  Location                           ->  journal
'43  journaltype                        ->  journal
'44  tblDataTypeDb1                     ->  tblPricing_AppraiseColumnDataType
'45  journaltype                        ->  RecurringType
'46  tblPricing_AppraiseSectionType     ->  tblPricing_AppraiseItemType
'47  TaxCode                            ->  AssetType
'48  Schedule                           ->  ScheduleDetail
'49  account                            ->  PortfolioModel
'50  m_REVCODE_TYPE                     ->  m_REVCODE
'51  account                            ->  journal
'52  account                            ->  ActiveAssets
'53  tblPricing_AppraiseItemType        ->  tblPricing_AppraiseColumn
'54  AssetType                          ->  tblPricing_MasterAsset_History
'55  tblXAdmin_Graphics                 ->  tblXAdmin_ExportQry
'56  tblQueryTableType                  ->  tblReport_RecordSource
'57  tblSectionType                     ->  tblForm_Section
'58  tblVBComponent                     ->  tblVBComponent_Shortcut
'59  tblObjectType                      ->  tblObject_Image
'60  tblDeclarationType                 ->  tblVBComponent_Declaration
'61  tblCommandBarHyperlinkType         ->  tblCommandBar_Control
'62  tblGroupKeepTogetherType           ->  tblReport_Specification
'63  tblDatabase_Table_Field            ->  tblIndex_Field
'64  tblVersion_File                    ->  tblVersion_NotFound
'65  tblForm                            ->  tblForm_Subform
'66  tblQueryTableType                  ->  tblForm_RecordSource
'67  tblControlType                     ->  tblDatabase_Table_Field_RowSource
'68  tblProcedureType                   ->  tblProcedureSubType
'69  tblReport                          ->  tblReport_RecordSource
'70  tblVBComponent                     ->  tblVBComponent_API
'71  tblForm_Control                    ->  tblForm_Control_Group
'72  tblReport                          ->  tblReport_Specification
'73  tblCommandBarButtonStyle           ->  tblCommandBar_Control
'74  tblDatabase                        ->  tblRelation
'75  tblReportOrientationType           ->  tblReport_Specification
'76  tblObjectType                      ->  tblQueryTableType
'77  tblSectionType                     ->  tblReport_Control
'78  tblProcedureType                   ->  tblVBComponent_API
'79  tblMacroAction                     ->  tblMacroActionArgument
'80  tblXAdmin_Graphics                 ->  tblXAdmin_ExportTbl
'81  tblForm_Control                    ->  tblForm_Control_RowSource
'82  tblQueryType                       ->  tblQuery
'83  tblReport_List_Report              ->  tblReport_List_Report_Alt
'84  tblForm                            ->  tblForm_Shortcut
'85  tblVBComponent_Procedure           ->  tblReport_VBComponent
'86  tblSystemColor_Base                ->  tblSystemColor_Section
'87  tblReport_Control_Specification_A  ->  tblReport_Control_Specification_B
'88  tblDatabase_Table_Field            ->  tblImport_Field
'89  tblControlType                     ->  tblDatabase_Table_Field
'90  tblTransactionForm                 ->  tblTransactionForm_Option
'91  tblGridlineStyle                   ->  tblForm_Specification_B
'92  tblFontWeight                      ->  tblReport_Control_Specification_A
'93  tblDataTypeVb                      ->  tblVBComponent_API
'94  tblJournalType                     ->  tblJournal_Field
'95  tblMacro                           ->  tblMacro_Row
'96  tblObjectType                      ->  tblXAdmin_Graphics_Image
'97  tblReport_List_Report_Alt_Control  ->  tblReport_List_Staging
'98  tblDataTypeVb                      ->  tblVBComponent_Procedure_Parameter
'99  tblPictureAlignmentType            ->  tblReport_Control_Specification_B
'100 tblObjectType                      ->  tblSystemColor_Control
'101 tblControlType                     ->  tblForm_Shortcut_Detail
'102 tblDatabase_Table                  ->  tblRelation
'103 tblVersion_File                    ->  tblVersion_Extra
'104 tblReport                          ->  tblReport_List_Report
'105 tblReport_NameBase                 ->  tblReport_NameBase_Report
'106 tblPictureAlignmentType            ->  tblObject_Image
'107 tblPictureAlignmentType            ->  tblForm_Specification_B
'108 tblVBAType                         ->  tblReference
'109 tblPicturePageType                 ->  tblReport_Specification
'110 tblSpecialEffect                   ->  tblForm_Section
'111 tblReport_List_Report_Alt_Control  ->  tblReport_List_Sort
'112 tblDocumentShapeType               ->  tblDocument_Document_Image
'113 tblObjectType                      ->  tblVBComponent_Procedure_Parameter
'114 tblVBComponent_Procedure           ->  tblVBComponent_Declaration_Detail
'115 tblParameterDataType               ->  tblQuery_Parameter
'116 tblVBComponent_Procedure           ->  tblVBComponent_Procedure_Detail
'117 tblForm                            ->  tblTransactionForm
'118 tblDocumentAutoShapeType           ->  tblDocument_Document_Image
'119 tblReportSortOrder                 ->  tblReport_Group
'120 tblKeyDownType                     ->  tblForm_Shortcut
'121 tblVersion_Table                   ->  tblVersion_Field
'122 tblCycle                           ->  tblForm_Specification_A
'123 tblButtonType                      ->  tblMsgBoxStyleType
'124 tblDatabase                        ->  tblDatabase_Table
'125 tblDAOType                         ->  tblVBComponent_Procedure_Parameter
'126 tblControlType                     ->  tblForm_Control
'127 tblReportListControlType           ->  tblReport_List_Control
'128 tblQuery                           ->  tblQuery_Parameter
'129 tblXAdmin_Graphics                 ->  tblXAdmin_Graphics_Control
'130 tblRecordLock                      ->  tblForm_Specification_B
'131 tblRecordsetType                   ->  tblForm_Specification_B
'132 tblControlBorderStyle              ->  tblForm_Control_Specification_A
'133 tblReport_List_Report_Alt          ->  tblReport_List_Report_Alt_Control
'134 tblPictureType                     ->  tblForm_Specification_B
'135 tblForm_Section                    ->  tblSystemColor_Section
'136 tblControlBorderStyle              ->  tblReport_Control_Specification_A
'137 tblDocument                        ->  tblDocument_Document_Image
'138 tblPictureType                     ->  tblObject_Image
'139 tblControlType                     ->  tblForm_Control_Group_Item
'140 tblBackStyle                       ->  tblReport_Control_Specification_A
'141 tblDataTypeDb                      ->  tblPreference_Control
'142 tblForm_Control                    ->  tblForm_Subform
'143 tblMacroActionArgument             ->  tblMacro_Row_Argument
'144 tblDatabase                        ->  tblVBComponent
'145 tblVBComponent_Event               ->  tblVBComponent_Procedure
'146 tblReport                          ->  tblReport_VBComponent
'147 tblProcedureSubType                ->  tblVBComponent_Procedure
'148 tblJournalType                     ->  tblTransactionForm_Option
'149 tblCommandBarPosition              ->  tblCommandBar
'150 tblDatabase_Table_Field            ->  tblDatabase_Table_Field_DateFormat
'151 tblDataTypeDb                      ->  tblVersion_Field
'152 tblDisplayType                     ->  tblForm_Control_Specification_A
'153 tblOLETypeAllowed                  ->  tblForm_Control_Specification_A
'154 tblReport                          ->  tblReport_Section
'155 tblForm_Control                    ->  tblReport_List_Control
'156 tblXAdmin_Graphics                 ->  tblForm_Graphics
'157 tblAutoActivateType                ->  tblForm_Control_Specification_A
'158 tblDecimalPlaceDb                  ->  tblDatabase_Table_Field
'159 tblControlType                     ->  tblForm_Shortcut
'160 tblScriptingType                   ->  tblVBComponent_Procedure_Parameter
'161 tblScopeType                       ->  tblVBComponent_Procedure
'162 tblScopeType                       ->  tblVBComponent_API
'163 tblReport_Control                  ->  tblReport_Subform
'164 tblScrollBarAlignment              ->  tblForm_Control_Specification_B
'165 tblDatabase_Table                  ->  tblDatabase_Table_Link
'166 tblForm                            ->  tblPreference_VBComponent
'167 tblForm_Control                    ->  tblForm_Graphics
'168 tblMacroAction                     ->  tblMacro_Row
'169 tblXAdmin_Graphics                 ->  tblXAdmin_Graphics_Image
'170 tblForm_Control                    ->  tblPreference_Control
'171 tblForm_Control                    ->  tblJournal_Field
'172 tblGridlineStyle                   ->  tblForm_Specification_A
'173 tblSectionType                     ->  tblForm_Control
'174 tblDefaultView                     ->  tblForm_Specification_A
'175 tblDisplayWhenType                 ->  tblForm_Section
'176 tblTextAlignType                   ->  tblReport_Control_Specification_B
'177 tblVersion_Directory               ->  tblVersion_File
'178 tblSpecialEffect                   ->  tblReport_Section
'179 tblRowSourceType                   ->  tblQueryTableType
'180 tblQuery                           ->  tblQuery_RecordSource
'181 tblDocument_Image                  ->  tblDocument_Document_Image
'182 tblForm                            ->  tblForm_RecordSource
'183 tblQuery                           ->  tblQuery_SourceChain
'184 tblForm_Section                    ->  tblForm_Control_Group
'185 tblDatabase                        ->  tblReport
'186 tblRelation                        ->  tblRelation_Field
'187 tblReport_Section                  ->  tblSystemColor_Section
'188 tblReport_List_Report              ->  tblReport_NameBase
'189 tblDocumentFieldType               ->  tblDocument_Document_Image
'190 tblDocumentInlineShapeType         ->  tblDocument_Document_Image
'191 tblReport_Section                  ->  tblReport_Control
'192 tblReport                          ->  tblReport_NameBase_Report
'193 tblSystemColor_Base                ->  tblSystemColor
'194 tblDatabase_Table_Field            ->  tblDatabase_Table_Field_RowSource
'195 tblObjectType                      ->  tblForm_Control
'196 tblForm                            ->  tblReport_List
'197 tblBorderWidth                     ->  tblReport_Control_Specification_A
'198 tblVBComponent_Declaration_Family  ->  tblVBComponent_Declaration
'199 tblObjectType                      ->  tblForm
'200 tblIndex                           ->  tblIndex_Field
'201 tblFormBorderStyle                 ->  tblReport_Specification
'202 tblDocumentLinkType                ->  tblDocument_Document_Image
'203 tblDocumentFieldKind               ->  tblDocument_Document_Image
'204 tblMinMaxButton                    ->  tblReport_Specification
'205 tblDatabase                        ->  tblXAdmin_ExportTbl
'206 tblCellEffect                      ->  tblForm_Specification_A
'207 tblObjectType                      ->  tblReport
'208 tblForm_Specification_A            ->  tblForm_Specification_B
'209 tblPictureType                     ->  tblReport_Specification
'210 tblDataTypeVb                      ->  tblVBComponent_Property
'211 tblVerbType                        ->  tblForm_Control_Specification_B
'212 tblObjectType                      ->  tblSectionType
'213 tblXAdmin_Graphics_Type            ->  tblXAdmin_Graphics
'214 tblVBComponent_Procedure           ->  tblVBComponent_Procedure_Parameter
'215 tblForm                            ->  tblForm_Specification_A
'216 tblForm_Section                    ->  tblForm_Control
'217 tblBorderWidth                     ->  tblForm_Control_Specification_A
'218 tblQueryType                       ->  tblXAdmin_Graphics
'219 tblMacro                           ->  tblMacro_Text
'220 tblDataTypeVb                      ->  tblVBComponent_Declaration
'221 tblForm_Section                    ->  tblJournal_Field
'222 tblDatabase                        ->  tblReport_NameBase
'223 tblForm_Graphics                   ->  tblCheckReconcile_Form_Graphics
'224 tblObjectType                      ->  tblReport_Control
'225 tblScopeType                       ->  tblVBComponent_Declaration
'226 tblDatabase                        ->  tblDatabase_Table_Link
'227 tblDatabase                        ->  tblXAdmin_ExportQry
'228 tblDatabase_Table                  ->  tblDatabase_Table_RecordCount
'229 tblRecordLock                      ->  tblReport_Specification
'230 tblReportKeepTogetherType          ->  tblReport_Group
'231 tblFontWeight                      ->  tblForm_Specification_A
'232 tblObjectType                      ->  tblReport_Section
'233 tblVBComponent_Property            ->  tblVBComponent_Event
'234 tblObjectType                      ->  tblXAdmin_Graphics_Control
'235 tblReport_List                     ->  tblReport_List_Control
'236 tblCommandBarComboStyle            ->  tblCommandBar_Control
'237 tblSpecialEffect                   ->  tblReport_Control_Specification_B
'238 tblUnderlineStyle                  ->  tblForm_Specification_A
'239 tblQuery                           ->  tblQuery_Field
'240 tblXAdmin_Graphics                 ->  tblTreeView_Icon
'241 tblSectionType                     ->  tblReport_Section
'242 tblDatabase_Table_Field            ->  tblRelation_Field
'243 tblReport_Control                  ->  tblSystemColor_Control
'244 tblDatabase                        ->  tblReference
'245 tblForm                            ->  tblForm_Section
'246 tblMultiSelectType                 ->  tblForm_Control_RowSource
'247 tblObjectType                      ->  tblVBComponent
'248 tblLineSlantType                   ->  tblForm_Control_Specification_A
'249 tblPictureSizeMode                 ->  tblReport_Specification
'250 tblRowSourceType                   ->  tblForm_Control_RowSource
'251 tblForm_Control                    ->  tblSystemColor_Control
'252 tblForm_Control                    ->  tblVBComponent_Shortcut
'253 tblVBComponent_Declaration         ->  tblVBComponent_Declaration_Detail
'254 tblKeyCode                         ->  tblForm_Shortcut
'255 tblForm_Control                    ->  tblImport_Field
'256 tblDatabase_Table                  ->  tblIndex
'257 tblDatabase_Table                  ->  tblDatabase_Table_Field
'258 tblScrollBarC                      ->  tblForm_Control_Specification_B
'259 tblQuery                           ->  tblQuery_FormRef
'260 tblDatabase                        ->  tblMacro
'261 tblDataTypeDb                      ->  tblReport_Group
'262 tblXAdmin_Graphics_Format          ->  tblXAdmin_Graphics
'263 tblCommandBarType                  ->  tblCommandBar
'264 tblVBComponent                     ->  tblVBComponent_Procedure
'265 tblBackStyle                       ->  tblForm_Control_Specification_A
'266 tblVersion_File                    ->  tblVersion_Table
'267 tblPictureType                     ->  tblForm_Control_Specification_B
'268 tblConnectionType                  ->  tblDatabase_Table_Link
'269 tblDatabase_Table_Field            ->  tblDatabase_AutoNumber
'270 tblSpecialEffect                   ->  tblForm_Control_Specification_B
'271 tblVBComponent_API                 ->  tblVBComponent_Declaration
'272 tblForm_Control_Specification_A    ->  tblForm_Control_Specification_B
'273 tblForm                            ->  tblForm_Control
'274 tblForm_Control                    ->  tblForm_Control_Specification_A
'275 tblVersion                         ->  tblVersion_Release
'276 tblDateGroupingType                ->  tblReport_Specification
'277 tblDataTypeVb                      ->  tblVBComponent_Procedure
'278 tblReport                          ->  tblReport_Subform
'279 tblDataTypeDb                      ->  tblQuery_Field
'280 tblObjectType                      ->  tblForm_Section
'281 tblConnectionType                  ->  tblXAdmin_Graphics
'282 tblForm_Control                    ->  tblImport_Field
'283 tblControlType                     ->  tblReport_Control
'284 tblPictureType                     ->  tblReport_Control_Specification_B
'285 tblMacro_Row                       ->  tblMacro_Row_Argument
'286 tblForm_Control                    ->  tblForm_Control_Group_Item
'287 tblVBComponent_Procedure           ->  tblPreference_VBComponent
'288 tblPictureSizeMode                 ->  tblReport_Control_Specification_B
'289 tblDatabase_Table                  ->  tblRelation
'290 tblFormBorderStyle                 ->  tblForm_Specification_A
'291 tblReport_List_Control             ->  tblReport_List_Report_Alt_Control
'292 tblDatabase                        ->  tblQuery
'293 tblPictureSizeMode                 ->  tblObject_Image
'294 tblObjectType                      ->  tblSystemColor_Section
'295 tblTaxCodeType                     ->  tblTaxCode
'296 tblScrollBarF                      ->  tblForm_Specification_B
'297 tblForceNewPageType                ->  tblForm_Section
'298 tblDataTypeVb                      ->  tblMacroActionArgument
'299 tblNewRowOrColType                 ->  tblReport_Section
'300 tblGridlineBehavior                ->  tblForm_Specification_A
'301 tblPictureSizeMode                 ->  tblForm_Specification_B
'302 tblKeyDownType                     ->  tblVBComponent_Shortcut
'303 tblStyleType                       ->  tblForm_Control_Specification_B
'304 tblReferenceType                   ->  tblReference
'305 tblDatabase                        ->  tblObject_Image
'306 tblSectionBaseType                 ->  tblSectionType
'307 tblQueryTableType                  ->  tblForm_Control_RowSource
'308 tblReport                          ->  tblReport_Group
'309 tblDatabase                        ->  tblForm
'310 tblForm_Control_Group              ->  tblForm_Control_Group_Item
'311 tblPictureAlignmentType            ->  tblForm_Control_Specification_B
'312 tblDataTypeDb                      ->  tblDatabase_Table_Field
'313 tblKeyCode                         ->  tblKeyboard_Monitor
'314 tblPictureAlignmentType            ->  tblReport_Specification
'315 tblVersionType                     ->  tblVersion_Release
'316 tblSystemColor_Base                ->  tblSystemColor_Control
'317 tblVBComponent                     ->  tblVBComponent_Declaration
'318 tblNewRowOrColType                 ->  tblForm_Section
'319 tblCommandBarControlType           ->  tblCommandBar_Control
'320 tblParameterDirection              ->  tblQuery_Parameter
'321 tblForceNewPageType                ->  tblReport_Section
'322 tblControlType                     ->  tblQuery_Field
'323 tblProcedureType                   ->  tblVBComponent_Procedure
'324 tblSystemColorType                 ->  tblSystemColor
'325 tblComponentType                   ->  tblVBComponent
'326 tblVBComponent_Procedure           ->  tblVBComponent_MessageBox
'327 tblReportGroupOnType               ->  tblReport_Group
'328 tblDatabase                        ->  tblCommandBar
'329 tblJournalType                     ->  tblJournalSubType
'330 tblMinMaxButton                    ->  tblForm_Specification_B
'331 tblTextAlignType                   ->  tblForm_Control_Specification_B
'332 tblRowSourceType                   ->  tblDatabase_Table_Field_RowSource
'333 tblPreference_Control              ->  tblPreference_User
'334 tblQueryTableType                  ->  tblDatabase_Table_Field_RowSource
'335 tblControlType                     ->  tblObject_Image
'336 tblVBComponent_Procedure           ->  tblPreference_VBComponent
'337 tblPictureSizeMode                 ->  tblForm_Control_Specification_B
'338 tblFontWeight                      ->  tblForm_Control_Specification_A
'339 tblForm_Control                    ->  tblForm_Shortcut
'340 tblCommandBar                      ->  tblCommandBar_Control
'341 tblCommandBarProtection            ->  tblCommandBar
'342 tblUpdateOptionType                ->  tblForm_Control_Specification_B
'343 tblXAdmin_Graphics_Group           ->  tblXAdmin_Graphics
'344 tblViewsAllowed                    ->  tblForm_Specification_B
'345 tblReport                          ->  tblReport_Control
'346 tblForm_Shortcut                   ->  tblForm_Shortcut_Detail
'347 tblReport_Control                  ->  tblReport_Control_Specification_A
'348 tblVBComponent_Procedure           ->  tblVBComponent_Procedure_Detail
'349 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'350 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'351 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'352 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'353 tblPageHeaderFooterType            ->  tblReport_Specification
'354 tblPageHeaderFooterType            ->  tblReport_Specification
'355 tblAcctCon_Field                   ->  tblAcctCon_Report_Field
'356 tblXAdmin_Relation                 ->  tblXAdmin_Relation_Field
'357 tblSecurity_Group                  ->  tblSecurity_GroupUser
'358 tblAcctCon_Section                 ->  tblAcctCon_Section_Option
'359 tblImportExport_Specifications     ->  tblImportExport_Columns
'360 tblAcctCon_Version_Section_Option  ->  tblAcctCon_Version_Group2
'361 tblAcctCon_Version_Section_Option  ->  tblAcctCon_Version_Group1
'362 tblAcctCon_Section_Option_Group    ->  tblAcctCon_Version_Group3
'363 tblAcctCon_Section_Option_Group    ->  tblAcctCon_Version_Group2
'364 tblAcctCon_Report                  ->  tblAcctCon_Report_Field
'365 tblAcctCon_Group                   ->  tblAcctCon_Group_Type
'366 tblAcctCon_Version_Section_Option  ->  tblAcctCon_Version_Group3
'367 tblAcctCon_Section_Group           ->  tblAcctCon_Section_Option_Group
'368 tblAcctCon_Version                 ->  tblAcctCon_Version_Option
'369 tblAcctCon_Version_Section         ->  tblAcctCon_Version_Section_Option
'370 tblAcctCon_Report                  ->  tblAcctCon_Report_Version
'371 tblAcctCon_Option                  ->  tblAcctCon_Section_Option
'372 tblAcctCon_Version                 ->  tblAcctCon_Report_Version
'373 tblAcctCon_Version                 ->  tblAcctCon_Version_Section
'374 tblAcctCon_Version_Option          ->  tblAcctCon_Version_Section_Option
'375 tblAcctCon_Section_Option_Group    ->  tblAcctCon_Version_Group1
'376 tblAcctCon_Format                  ->  tblAcctCon_Group_Type_Format
'377 tblAcctCon_Field                   ->  tblAcctCon_Group_Type_Format_Field
'378 tblAcctCon_Section                 ->  tblAcctCon_Version_Section
'379 tblSecurity_User                   ->  tblSecurity_GroupUser
'380 tblAcctCon_Group_Type_Format       ->  tblAcctCon_Group_Type_Format_Field
'381 tblRelation_View                   ->  tblRelation_View_Window
'382 tblAcctCon_Group_Type              ->  tblAcctCon_Group_Type_Format
'383 tblAcctCon_Section                 ->  tblAcctCon_Section_Group
'384 tblAcctCon_Group                   ->  tblAcctCon_Section_Group
'385 tblAcctCon_Section_Option          ->  tblAcctCon_Section_Option_Group
'386 tblAcctCon_Option                  ->  tblAcctCon_Version_Option
'387 journaltype                        ->  tblJournalType
'388 m_REVCODE_TYPE                     ->  tblTaxCode
'389 TaxCode_Type                       ->  tblTaxCodeType
'390 TaxCode                            ->  tblTaxCode
'391 tblDatabase                        ->  tblRelation_View

  Set Rel = Nothing
  Set dbs = Nothing

  Beep

  Rel_List = blnRetValx

End Function

Private Function Rel_ChkDocQrys() As Boolean
' ** Called by:
' **   QuikRelDoc(), Above

  Const THIS_PROC As String = "Rel_ChkDocQrys"

  blnRetValx = True

  DoCmd.Hourglass True
  DoEvents

'zz_qry_Relation_01
'zz_qry_Relation_02
'zz_qry_Relation_03
'zz_qry_Relation_04a
'zz_qry_Relation_04b
'zz_qry_Relation_04c
'zz_qry_Relation_04d
'zz_qry_Relation_04e
'zz_qry_Relation_04f

'10/17/2014:
'zz_qry_Relation_01
'zz_qry_Relation_02
'zz_qry_Relation_03
'zz_qry_Relation_04e
'zz_qry_Relation_04d
'zz_qry_Relation_04c
'zz_qry_Relation_04a
'zz_qry_Relation_04b
'zz_qry_Relation_04f

  DoCmd.Hourglass False
  DoEvents

  Beep

  Rel_ChkDocQrys = blnRetValx

End Function
