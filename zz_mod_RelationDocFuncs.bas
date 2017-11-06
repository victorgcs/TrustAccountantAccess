Attribute VB_Name = "zz_mod_RelationDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_RelationshipFuncs"

'VGC 11/21/2016: CHANGES!

' ** DbRelation enumeration:
' **          0  dbRelationEnforce        The relationship is enforced (referential integrity). {my constant}
' **          1  dbRelationUnique         The relationship is one-to-one.
' **          2  dbRelationDontEnforce    The relationship isn't enforced (no referential integrity).
' **          4  dbRelationInherited      The relationship exists in a non-current database that contains the two linked tables.
' **        256  dbRelationUpdateCascade  Updates will cascade.
' **       4096  dbRelationDeleteCascade  Deletions will cascade.
' **   16777216  dbRelationLeft           In Design view, display a LEFT JOIN as the default join type. Microsoft Access only.
' **   33554432  dbRelationRight          In Design view, display a RIGHT JOIN as the default join type. Microsoft Access only.

' ** Relationship Names:
' ** Maximum saved length seems to be 61, so must be unique within that.

' ** Right now, this can remain private.
Private Const dbRelationEnforce As Long = 0&

Private Const gstrFile_IApp As String = "TrustImport"

Private blnRetValx As Boolean, blnWithJrnlTmp As Boolean
' **

Public Function QuikRelDoc() As Boolean
  Const THIS_PROC As String = "QuikRelDoc"
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Rel_ChkDocQrys(False) = True Then  ' ** Function: Below.
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
  Dim Rel As Relation, fld As Object
  Dim strThisFile As String, strThisPath As String
  Dim strThatFile As String, strThatPath As String
  Dim lngDbs As Long, arr_varDb() As Variant, lngDBID_AsOf As Long, lngDBID1 As Long, lngDBID2 As Long
  Dim lngTblID1 As Long, lngTblID2 As Long, lngRelID As Long
  Dim lngRelCnt As Long, intRelOrd As Integer
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim blnThisDbs As Boolean, blnTrustDbs As Boolean, blnImportDbs As Boolean
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
          If blnWithJrnlTmp = False And ![dbs_name] = "TAJrnTmp.mdb" Then
            ' ** Skip it.
          ElseIf CurrentAppName = "Trust.mdb" And ![dbs_name] = "TrstXAdm.mdb" Then  ' ** Module Function: modFileUtilities.
            ' ** Skip this.
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
          If lngX < lngRecs Then .MoveNext
        Next
        .MoveFirst

        For lngX = 0& To (lngDbs - 1&)
          blnFound2 = True
          If ![dbs_id] <> arr_varDb(D_DID, lngX) Then
            .FindFirst "[dbs_id] = " & CStr(arr_varDb(D_DID, lngX))
            If .NoMatch = True Then
              Stop
            End If
          End If
          If arr_varDb(D_DNAM, lngX) = strThisFile Then
            strTmp01 = strThisPath
          ElseIf arr_varDb(D_DNAM, lngX) = (gstrFile_App & "." & gstrExt_AppDev) Or _
              arr_varDb(D_DNAM, lngX) = (gstrFile_App & "." & gstrExt_AppRun) Then
            strTmp01 = Parse_Path(CurrentBackendPathFile("m_VP"))  ' ** Module Function: modFileUtilities.
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_DataName Then
            strTmp01 = Parse_Path(CurrentBackendPathFile("m_VD"))  ' ** Module Function: modFileUtilities.
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_ArchDataName Then
            strTmp01 = Parse_Path(CurrentBackendPathFile("m_VA"))  ' ** Module Function: modFileUtilities.
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_AuxDataName Then
            strTmp01 = Parse_Path(CurrentBackendPathFile("m_VX"))  ' ** Module Function: modFileUtilities.
          ElseIf arr_varDb(D_DNAM, lngX) = (gstrFile_IApp & "." & gstrExt_AppDev) Or _
              arr_varDb(D_DNAM, lngX) = (gstrFile_IApp & "." & gstrExt_AppRun) Then
            strTmp01 = CurrentBackendPathFile("m_VI")
            If strTmp01 <> vbNullString Then
              strTmp01 = Parse_Path(CurrentBackendPathFile("m_VI"))  ' ** Module Function: modFileUtilities.
            Else
              blnFound2 = False
            End If
          ElseIf arr_varDb(D_DNAM, lngX) = gstrFile_RePostDataName Then
            ' ** Leave it as-is.
            strTmp01 = ![dbs_path]
          ElseIf arr_varDb(D_DNAM, lngX) = "FileSpec.mdb" Then
            strTmp01 = CurrentBackendPathFile("m_VI")
            If strTmp01 <> vbNullString Then
              strTmp01 = Parse_Path(CurrentBackendPathFile("m_VI"))  ' ** Module Function: modFileUtilities.
            Else
              strTmp01 = Parse_Path(CurrentAppPath)  ' ** Module Function: modFileUtilities.
              strTmp01 = strTmp01 & LNK_SEP & "Trust Import"
            End If
          Else
            Stop
          End If
          If arr_varDb(D_PATH, lngX) <> strTmp01 And blnFound2 = True Then
            .Edit
            ![dbs_path] = strTmp01
            ![dbs_datemodified] = Now()
            .Update
            arr_varDb(D_PATH, lngX) = strTmp01
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

  ' ** Reset the Autonumber field.
  ChangeSeed_Ext "tblRelation"  ' ** Module Function: modAutonumberFieldFuncs.
  ChangeSeed_Ext "tblRelation_Field"  ' ** Module Function: modAutonumberFieldFuncs.

  Set dbsFE = CurrentDb
  Set rstRel = dbsFE.OpenRecordset("tblRelation", dbOpenDynaset, dbConsistent)
  Set rstRelFld = dbsFE.OpenRecordset("tblRelation_Field", dbOpenDynaset, dbAppendOnly)

  For lngX = 0& To (lngDbs - 1&)

    If arr_varDb(D_DNAM, lngX) = strThisFile Then
      Set dbsBE = CurrentDb
      blnThisDbs = True
      blnTrustDbs = False
      blnImportDbs = False
    Else
      Set wrk = DBEngine.CreateWorkspace("Tmp", "Superuser", TA_SEC, dbUseJet)
      Set dbsBE = wrk.OpenDatabase(arr_varDb(D_PATH, lngX) & LNK_SEP & arr_varDb(D_DNAM, lngX), False, True)  ' ** {pathfile}, {exclusive}, {read-only}
      blnThisDbs = False
      If Left(arr_varDb(D_DNAM, lngX), 6) = "Trust." Then
        blnTrustDbs = True
      Else
        blnTrustDbs = False
      End If
      If Left(arr_varDb(D_DNAM, lngX), 12) = "TrustImport." Then
        blnImportDbs = True
      Else
        blnImportDbs = False
      End If
    End If
    With dbsBE

      lngRelCnt = 0&
      For Each Rel In .Relations
'FOR EACH RELATIONSHIP IN THAT DATABASE,...
        If Left(Rel.Table, 4) <> "MSys" And Left(Rel.ForeignTable, 4) <> "MSys" And _
            Left(Rel.Table, 4) <> "~TMP" And Left(Rel.ForeignTable, 4) <> "~TMP" Then  ' ** Skip those pesky system tables!

          ' ** If the dbs_id's don't match, make sure we don't put it where the ForeignTable isn't available.
          ' ** Or, non-matching relationships can only go in the frontend!
          blnFound = False
          For Each tdf In .TableDefs
            If tdf.Name = Rel.Table And IIf(blnThisDbs = False, IIf(tdf.Connect = vbNullString, True, False), True) Then
'IF THE TABLE IS LISTED, AND IT'S NOT THIS DATABASE, AND IT'S NOT LINKED TO THAT DATABASE,...
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
              blnFound = False
              If blnTrustDbs = True Or blnImportDbs = True Then
                ' ** These inherited relationships are OK, just don't document them here.
              Else
                ' ** Don't document this; it doesn't belong!
                Debug.Print "'REL ERR: " & Parse_File(.Name) & "  " & Rel.Table & " -> " & Rel.ForeignTable & " IN " & arr_varDb(D_DNAM, lngX)
              End If
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
                    strTmp01 = Rel.Table
                    Select Case strTmp01
                    Case "tblSecurity_Group_TI"
                      strTmp01 = "tblSecurity_Group"
                    Case "tblSecurity_GroupUser_TI"
                      strTmp01 = "tblSecurity_GroupUser"
                    Case "tblSecurity_User_TI"
                      strTmp01 = "tblSecurity_User"
                    End Select
                    For Each tdf In dbsBE.TableDefs
                      With tdf

                        If .Name = strTmp01 Then
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
                    If strTmp01 = "tblDataTypeDb1" Then
                      ![tblnam] = "tblDataTypeDb"  ' ** Query Parameter.
                    Else
                      ![tblnam] = strTmp01  ' ** Query Parameter.
                    End If
                  Case 2&
                    ' ** The relation is local to this DB, but its tables may not be.
                    blnFound = False
                    strTmp01 = Rel.ForeignTable
                    Select Case strTmp01
                    Case "tblSecurity_Group_TI"
                      strTmp01 = "tblSecurity_Group"
                    Case "tblSecurity_GroupUser_TI"
                      strTmp01 = "tblSecurity_GroupUser"
                    Case "tblSecurity_User_TI"
                      strTmp01 = "tblSecurity_User"
                    End Select
                    For Each tdf In dbsBE.TableDefs
                      With tdf
                        If .Name = strTmp01 Then
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
                                  "[tbl_name] = '" & strTmp01 & "'")
                                Select Case IsNull(varTmp00)
                                Case True
                                  ' ** The Foreign Table isn't in the same database as the Table.
                                  varTmp00 = DLookup("[dbs_id]", "tblDatabase_Table", "[tbl_name] = '" & strTmp01 & "'")
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
                      varTmp00 = DLookup("[dbs_id]", "tblDatabase_Table", "[tbl_name] = '" & strTmp01 & "'")
                      If IsNull(varTmp00) = True Then
                        Stop
                      Else
                        blnFound = True
                        lngDBID2 = varTmp00
                      End If
                    End If
                    ![dbid] = lngDBID2  ' ** Query Parameter.
                    If strTmp01 = "tblDataTypeDb1" Then
                      ![tblnam] = "tblDataTypeDb"  ' ** Query Parameter.
                    Else
                      ![tblnam] = strTmp01  ' ** Query Parameter.
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
  Dim blnFound As Boolean, blnJustList As Boolean, blnJustSysNames As Boolean
  Dim intPos1 As Integer
  Dim lngTmp00 As Long, strTmp01 As String, strTmp02 As String, lngTmp03 As Long, blnTmp04 As Boolean
  Dim lngX As Long, lngY As Long, lngZ As Long

  ' ** Array: arr_varRel().
  Const R_ID      As Integer = 0
  Const R_DBSID1  As Integer = 1
  Const R_DBSNAM1 As Integer = 2
  Const R_TBLID1  As Integer = 3
  Const R_TBLNAM1 As Integer = 4
  Const R_DBSID2  As Integer = 5
  Const R_DBSNAM2 As Integer = 6
  Const R_TBLID2  As Integer = 7
  Const R_TBLNAM2 As Integer = 8
  Const R_NAM     As Integer = 9
  Const R_ATTR    As Integer = 10
  Const R_FLDS    As Integer = 11
  Const R_FLDARR  As Integer = 12

  ' ** Array: arr_varFld().
  Const FDF_RELID   As Integer = 0
  Const FDF_DBSID   As Integer = 1
  Const FDF_DBSNAM  As Integer = 2
  Const FDF_TBLID1  As Integer = 3
  Const FDF_TBLNAM1 As Integer = 4
  Const FDF_TBLID2  As Integer = 5
  Const FDF_TBLNAM2 As Integer = 6
  Const FDF_FLDORD  As Integer = 7
  Const FDF_RFLDID  As Integer = 8
  Const FDF_FLDID1  As Integer = 9
  Const FDF_FLDNAM1 As Integer = 10
  Const FDF_FLDID2  As Integer = 11
  Const FDF_FLDNAM2 As Integer = 12
  Const FDF_DBSID2  As Integer = 14

  ' ** Array: arr_varRelTbl().
  Const RT_DBSID  As Integer = 0
  Const RT_TBLID  As Integer = 1
  Const RT_TBLNAM As Integer = 2
  Const RT_DBSIDA As Integer = 3
  Const RT_RELCNT As Integer = 4
  Const RT_ARR_F  As Integer = 5
  Const RT_ARR_L  As Integer = 6
  Const RT_DBSID2 As Integer = 7

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
      Set qdf = .QueryDefs("zz_qry_Relation_12s")
      'Set qdf = .QueryDefs("zz_qry_Relation_12x")
    Case False
      ' ** WITHOUT NON-CONTIGUOUS:
      ' **   Also without tblRelation_View, tblRelation_View_Window, or anything in TAJrnTmp.mdb.
      ' ** All relationships, with table names (w/o TAJrnTmp.mdb).
      Set qdf = .QueryDefs("zz_qry_Relation_12")  ' ** NEEDS TO BE IN dbs_id1 ORDER!
    End Select

    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRels = .RecordCount
      .MoveFirst
      .sort = "dbs_id1, tbl_name1, tbl_name2"
      Set rst2 = .OpenRecordset
      rst2.MoveLast
      lngRels = rst2.RecordCount
      rst2.MoveFirst
      arr_varRel = rst2.GetRows(lngRels)
      ' ****************************************************
      ' ** Array: arr_varRel()
      ' **
      ' **   Field  Element  Name              Constant
      ' **   =====  =======  ================  ===========
      ' **     1       0     rel_id            R_ID
      ' **     2       1     dbs_id1           R_DBSID1
      ' **     3       2     dbs_name1         R_DBSNAM1
      ' **     4       3     tbl_id1           R_TBLID1
      ' **     5       4     tbl_name1         R_TBLNAM1
      ' **     6       5     dbs_id2           R_DBSID2
      ' **     7       6     dbs_name2         R_DBSNAM2
      ' **     8       7     tbl_id2           R_TBLID2
      ' **     9       8     tbl_name2         R_TBLNAM2
      ' **    10       9     rel_name          R_NAM
      ' **    11      10     rel_attributes    R_ATTR
      ' **    12      11     rel_fld_cnt       R_FLDS
      ' **    13      12     rel_fld_arr       R_FLDARR
      ' **    14      13     DontEnforce
      ' **
      ' ****************************************************
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
        Set qdf = .QueryDefs("zz_qry_Relation_13u")
        'Set qdf = .QueryDefs("zz_qry_Relation_13z")
      Case False
        ' ** WITHOUT NON-CONTIGUOUS:
        ' **   Also without tblRelation_View, tblRelation_View_Window, or anything in TAJrnTmp.mdb.
        ' ** zz_qry_Relation_13b (zz_qry_Relation_13a (tblRelation, linked to tblRelation_Field,
        ' ** with add'l fields), with add'l foreign table fields), by specified [relid].
        Set qdf = .QueryDefs("zz_qry_Relation_13")  ' ** NEEDS TO BE IN dbs_id ORDER!
      End Select

      With qdf.Parameters
        ![relid] = arr_varRel(R_ID, lngX)
      End With
      Set rst1 = qdf.OpenRecordset
      With rst1
        .MoveLast
        lngFlds = .RecordCount
        .MoveFirst
        .sort = "dbs_id, tbl_name1, tbl_name2, relfld_order, fld_name1, fld_name2"
        Set rst2 = .OpenRecordset
        rst2.MoveLast
        lngFlds = rst2.RecordCount
        rst2.MoveFirst
        arr_varFld = rst2.GetRows(lngFlds)
        ' ****************************************************
        ' ** Array: arr_varFld()
        ' **
        ' **   Field  Element  Name            Constant
        ' **   =====  =======  ==============  =============
        ' **     1       0     rel_id          FDF_RELID
        ' **     2       1     dbs_id          FDF_DBSID
        ' **     3       2     dbs_name        FDF_DBSNAM
        ' **     4       3     tbl_id1         FDF_TBLID1
        ' **     5       4     tbl_name1       FDF_TBLNAM1
        ' **     6       5     tbl_id2         FDF_TBLID2
        ' **     7       6     tbl_name2       FDF_TBLNAM2
        ' **     8       7     relfld_order    FDF_FLDORD
        ' **     9       8     relfld_id       FDF_RFLDID
        ' **    10       9     fld_id1         FDF_FLDID1
        ' **    11      10     fld_name1       FDF_FLDNAM1
        ' **    12      11     fld_id2         FDF_FLDID2
        ' **    13      12     fld_name2       FDF_FLDNAM2
        ' **    14      13     DontEnforce
        ' **    15      14     dbs_id2         FDF_DBSID2
        ' **
        ' ****************************************************
      End With
      arr_varRel(R_FLDARR, lngX) = arr_varFld
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
      'tblRelation, grouped by tbl_id, just relationships having system-name, with rel_cnt.
      Set qdf = .QueryDefs("zz_qry_Relation_14s")
      'Set qdf = .QueryDefs("zz_qry_Relation_14x")
    Case False
      ' ** WITHOUT NON-CONTIGUOUS:
      ' **   Also without tblRelation_View, tblRelation_View_Window, or anything in TAJrnTmp.mdb.
      ' ** tblRelation, grouped by tbl_id, with rel_cnt.
      Set qdf = .QueryDefs("zz_qry_Relation_14")  ' ** NEEDS TO BE IN dbs_id1 ORDER!
    End Select

    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRelTbls = .RecordCount
      .MoveFirst
      .sort = "dbs_id1, tbl_name"
      Set rst2 = .OpenRecordset
      rst2.MoveLast
      lngRelTbls = rst2.RecordCount
      rst2.MoveFirst
      arr_varRelTbl = rst2.GetRows(lngRelTbls)
      ' ***************************************************
      ' ** Array: arr_varRelTbl()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ===========
      ' **     1       0     dbs_id           RT_DBSID
      ' **     2       1     tbl_id           RT_TBLID
      ' **     3       2     tbl_name         RT_TBLNAM
      ' **     4       3     dbs_id_asof      RT_DBSIDA
      ' **     5       4     rel_cnt          RT_RELCNT
      ' **     6       5     rel_arr_first    RT_ARR_F
      ' **     7       6     rel_arr_last     RT_ARR_L
      ' **     8       7     dbs_id2          RT_DBSID2
      ' **
      ' ***************************************************
      .Close
    End With
    Set rst1 = Nothing
    Set rst2 = Nothing

    ' ** Put the first and last arr_varRel() array elements, for each table, into the arr_varRelTbl() array for that table.
    strLastTblName = vbNullString
    For lngX = 0& To (lngRelTbls - 1&)
      If (((arr_varRelTbl(RT_DBSID, lngX) = 5&) And (arr_varRelTbl(RT_TBLID, lngX) = 439&) And _
          (arr_varRelTbl(RT_TBLNAM, lngX) = "tblPricing_Import")) Or _
          ((arr_varRelTbl(RT_DBSID, lngX) = 5&) And (arr_varRelTbl(RT_TBLID, lngX) = 650&) And _
          (arr_varRelTbl(RT_DBSID2, lngX) = 5&) And (arr_varRelTbl(RT_TBLNAM, lngX) = "tblRelation_View")) Or _
          ((arr_varRelTbl(RT_DBSID, lngX) = 5&) And (arr_varRelTbl(RT_TBLID, lngX) = 342&) And _
          (arr_varRelTbl(RT_TBLNAM, lngX) = "zz_tbl_RePost_Account")) Or _
          ((arr_varRelTbl(RT_DBSID, lngX) = 5&) And (arr_varRelTbl(RT_TBLID, lngX) = 348&) And _
          (arr_varRelTbl(RT_TBLNAM, lngX) = "zz_tbl_RePost_MasterAsset"))) Then
        For lngY = 0& To (lngRels - 1&)
          If ((arr_varRel(R_TBLNAM1, lngY) = arr_varRelTbl(RT_TBLNAM, lngX)) And _
              (arr_varRel(R_DBSID2, lngY) = arr_varRelTbl(RT_DBSID2, lngX))) Then
            arr_varRelTbl(RT_ARR_F, lngX) = lngY
            Exit For
          End If
        Next
        If arr_varRelTbl(RT_TBLNAM, lngX) = "tblPricing_Import" Then
          arr_varRelTbl(RT_ARR_L, lngX - 1&) = 69&
        Else
          arr_varRelTbl(RT_ARR_L, lngX - 1&) = (lngY - 1&)
        End If
        strLastTblName = arr_varRelTbl(RT_TBLNAM, lngX)
        If lngX = (lngRelTbls - 1&) Then
          arr_varRelTbl(RT_ARR_L, lngX) = CLng(UBound(arr_varRel, 2))
        End If
      Else
        For lngY = 0& To (lngRels - 1&)
          If ((arr_varRel(R_TBLNAM1, lngY) = arr_varRelTbl(RT_TBLNAM, lngX)) And _
              (arr_varRel(R_DBSID2, lngY) = arr_varRelTbl(RT_DBSID2, lngX))) Then
            If arr_varRelTbl(RT_TBLNAM, lngX) <> strLastTblName Then
              arr_varRelTbl(RT_ARR_F, lngX) = lngY
              If lngX > 0& Then
                arr_varRelTbl(RT_ARR_L, lngX - 1&) = lngY - 1&
              End If
              strLastTblName = arr_varRelTbl(RT_TBLNAM, lngX)
            End If
          End If
        Next
        If lngX = (lngRelTbls - 1&) Then
          arr_varRelTbl(RT_ARR_L, lngX) = CLng(UBound(arr_varRel, 2))
        End If
      End If
    Next

    .Close
  End With

  Select Case blnJustList
  Case True

    For lngX = 0& To (lngRelTbls - 1&)
      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      Debug.Print "'" & Left$(CStr(arr_varRelTbl(RT_ARR_F, lngX)) & "   ", 3) & " - " & _
        Left$(CStr(arr_varRelTbl(RT_ARR_L, lngX)) & "   ", 3) & "  " & arr_varRelTbl(RT_TBLNAM, lngX)
    Next

'0   - 1    tblConnectionType
'2   - 2    tblDatabase
'3   - 14   tblObjectType
'15  - 16   tblQueryType
'17  - 23   tblReport
'24  - 24   journal

'JUST SYSTEM NAME:
'0   - 1    tblCheckReconcile_Account
'2   - 2    tblCheckReconcileEntryType
'3   - 3    tblCheckReconcileSourceType
'4   - 4    tblForm_Control_Group
'5   - 5    tblImportExport_Specifications
'6   - 6    tblRelation_View
'7   - 7    tblSecurity_Group
'8   - 8    tblSecurity_User
'9   - 9    tblVersion
'10  - 10   tblVersionType
'11  - 11   tblXAdmin_Customer
'12  - 12   zz_tbl_Client_Directory
'13  - 13   zz_tbl_Client_File
'14  - 14   zz_tbl_Client_Table
'15  - 15   zz_tbl_Dev_Directory
'16  - 16   zz_tbl_Report_VBComponent_01
'17  - 17   zz_tbl_Report_VBComponent_02
'18  - 18   tblConnectionType
'19  - 19   tblDatabase
'20  - 20   tblObjectType
'21  - 21   tblQueryType
'22  - 22   tblReport
'23  - 28   account
'29  - 30   AccountType
'31  - 31   adminofficer
'32  - 34   AssetType
'35  - 35   HiddenType
'36  - 36   InvestmentObjective
'37  - 37   journal
'38  - 40   journaltype
'41  - 43   Location
'44  - 45   m_REVCODE
'46  - 47   m_REVCODE_TYPE
'48  - 50   masterasset
'51  - 52   RecurringItems
'53  - 53   RecurringType
'54  - 55   Schedule
'56  - 58   TaxCode
'59  - 59   TaxCode_Type
'60  - 61   tblDataTypeDb
'62  - 62   tblPricing_AppraiseColumnDataType
'63  - 63   tblPricing_AppraiseColumnIDQuote
'64  - 64   tblPricing_AppraiseFile
'65  - 65   tblPricing_AppraiseItemType
'66  - 67   tblPricing_AppraiseRowType
'68  - 69   tblPricing_AppraiseSectionType
'70  - 70   tblPricing_Import
'71  - 71   tblRelation_View
'72  - 74   zz_tbl_RePost_Account
'75  - 75   zz_tbl_RePost_MasterAsset

'LOOKS GOOD!
'0   - 0    tblCheckReconcileEntryType
'1   - 1    tblCheckReconcileSource
'2   - 2    tblImportExport_Specifications
'3   - 3    tblSecurity_Group
'4   - 4    tblSecurity_User
'5   - 5    tblVersion
'6   - 6    tblVersionType
'7   - 7    tblXAdmin_Customer
'8   - 8    zz_tbl_AssetTypeNew_AssetType
'9   - 9    zz_tbl_Client_Directory
'10  - 10   zz_tbl_Client_File
'11  - 11   zz_tbl_Client_Table
'12  - 12   zz_tbl_Dev_Directory
'13  - 13   zz_tbl_Report_VBComponent_01
'14  - 14   zz_tbl_Report_VBComponent_02
'15  - 20   account
'21  - 22   AccountType
'23  - 23   adminofficer
'24  - 26   AssetType
'27  - 27   HiddenType
'28  - 28   InvestmentObjective
'29  - 29   journal
'30  - 32   journaltype
'33  - 35   Location
'36  - 37   m_REVCODE
'38  - 39   m_REVCODE_TYPE
'40  - 42   masterasset
'43  - 43   RecurringType
'44  - 45   Schedule
'46  - 48   TaxCode
'49  - 49   TaxCode_Type
'50  - 51   tblDataTypeDb
'52  - 52   tblPricing_AppraiseColumnDataType
'53  - 53   tblPricing_AppraiseColumnIDQuote
'54  - 54   tblPricing_AppraiseFile
'55  - 55   tblPricing_AppraiseItemType
'56  - 57   tblPricing_AppraiseRowType
'58  - 59   tblPricing_AppraiseSectionType
'60  - 60   tblPricing_Import
'61  - 62   tblBackStyle
'63  - 64   tblBorderWidth
'65  - 65   tblButtonType
'66  - 66   tblCellEffect
'67  - 67   tblComponentType
'68  - 68   tblConnectionType
'69  - 70   tblControlBorderStyle
'71  - 76   tblControlType
'77  - 77   tblCycle
'78  - 78   tblDAOType
'79  - 79   tblDatabase
'80  - 85   tblDatabase_Table
'86  - 49   tblDatabase_Table_Field
'50  - 92   tblDataTypeDb
'93  - 96   tblDataTypeVb
'97  - 97   tblDateGroupingType
'98  - 98   tblDecimalPlaceDb
'99  - 99   tblDefaultView
'100 - 100  tblDisplayWhenType
'101 - 101  tblDocument
'102 - 102  tblDocument_Image
'103 - 103  tblDocumentAutoShapeType
'104 - 104  tblDocumentFieldKind
'105 - 105  tblDocumentFieldType
'106 - 106  tblDocumentInlineShapeType
'107 - 107  tblDocumentLinkType
'108 - 108  tblDocumentShapeType
'109 - 111  tblFontWeight
'112 - 113  tblForceNewPageType
'114 - 121  tblForm
'122 - 129  tblForm_Control
'130 - 130  tblForm_Control_Specification_A
'131 - 133  tblForm_Section
'134 - 134  tblForm_Specification_A
'135 - 135  tblFormBorderStyle
'136 - 136  tblGridlineBehavior
'137 - 138  tblGridlineStyle
'139 - 139  tblGroupKeepTogetherType
'140 - 140  tblIndex
'141 - 142  tblJournalType
'143 - 144  tblKeyCode
'145 - 145  tblLineSlantType
'146 - 147  tblMacro
'148 - 148  tblMacro_Row
'149 - 150  tblMacroAction
'151 - 151  tblMacroActionArgument
'152 - 152  tblMinMaxButton
'153 - 156  tblMsgBoxStyleType
'157 - 157  tblMultiSelectType
'158 - 159  tblNewRowOrColType
'160 - 170  tblObjectType
'171 - 172  tblPageHeaderFooterType
'173 - 173  tblParameterDataType
'174 - 174  tblParameterDirection
'175 - 178  tblPictureAlignmentType
'179 - 179  tblPicturePageType
'180 - 183  tblPictureSizeMode
'184 - 187  tblPictureType
'188 - 188  tblPreference_Control
'189 - 189  tblProcedureSubType
'190 - 191  tblProcedureType
'192 - 195  tblQuery
'196 - 199  tblQueryTableType
'200 - 200  tblQueryType
'201 - 202  tblRecordLock
'203 - 203  tblRecordsetType
'204 - 204  tblReferenceType
'205 - 205  tblRelation
'206 - 212  tblReport
'213 - 214  tblReport_Control
'215 - 216  tblReport_Section
'217 - 217  tblReportGroupOnType
'218 - 218  tblReportKeepTogetherType
'219 - 219  tblReportOrientationType
'220 - 220  tblReportSortOrder
'221 - 223  tblRowSourceType
'224 - 224  tblScopeType
'225 - 225  tblScriptingType
'226 - 226  tblScrollBarAlignment
'227 - 227  tblScrollBarC
'228 - 228  tblScrollBarF
'229 - 232  tblSpecialEffect
'233 - 233  tblStyleType
'234 - 236  tblSystemColor_Base
'237 - 237  tblSystemColorType
'238 - 238  tblTaxCodeType
'239 - 240  tblTextAlignType
'241 - 241  tblTransactionForm
'242 - 242  tblUnderlineStyle
'243 - 243  tblVBAType
'244 - 244  tblVBComponent
'245 - 245  tblVBComponent_Event
'246 - 248  tblVBComponent_Procedure
'249 - 249  tblVBComponent_Property
'250 - 250  tblVersion_Directory
'251 - 251  tblVersion_File
'252 - 252  tblVersion_Table
'253 - 253  tblViewsAllowed
'254 - 256  tblXAdmin_Graphics
'257 - 257  tblXAdmin_Graphics_Format
'258 - 258  tblXAdmin_Graphics_Type

'lngRelTbls = 0&


  Case False  ' ** blnJustList.

    strLastRelName = vbNullString
    blnTmp04 = False

    ' ** For each table with one or more relationships.
    For lngX = 0& To (lngRelTbls - 1&)

      ' ***************************************************
      ' ** Array: arr_varRelTbl()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ===========
      ' **     1       0     dbs_id           RT_DBSID
      ' **     2       1     tbl_id           RT_TBLID
      ' **     3       2     tbl_name         RT_TBLNAM
      ' **     4       3     dbs_id_asof      RT_DBSIDA
      ' **     5       4     rel_cnt          RT_RELCNT
      ' **     6       5     rel_arr_first    RT_ARR_F
      ' **     7       6     rel_arr_last     RT_ARR_L
      ' **
      ' ***************************************************

'tblPricing_AppraiseColumnDataType SHOWED UP IN TrustAux.mdb,
'WHEN IT'S REALLY IN TrustDta.mdb!!!

      strTmp01 = DLookup("[dbs_name]", "tblDatabase", "[dbs_id] = " & CStr(arr_varRelTbl(RT_DBSIDA, lngX)))
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
          strTmp02 = DLookup("[dbs_path]", "tblDatabase", "[dbs_id] = " & CStr(arr_varRelTbl(RT_DBSID, lngX)))
          If Right$(strTmp02, 1) = LNK_SEP Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
        End If
        Set dbs = wrk.OpenDatabase(strTmp02 & LNK_SEP & strTmp01, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
      End If

      With dbs

        ' ** Delete all relationships for this table.
        .Relations.Refresh
        lngTmp00 = .Relations.Count
        For lngY = (lngTmp00 - 1&) To 0 Step -1&
          If .Relations(lngY).Table = arr_varRelTbl(RT_TBLNAM, lngX) Then
            .Relations.Delete .Relations(lngY).Name
            .Relations.Refresh
          End If
        Next

        ' ** Now re-create the relationships.
        For lngY = arr_varRelTbl(RT_ARR_F, lngX) To arr_varRelTbl(RT_ARR_L, lngX)

          ' ** Define the relationship.
          lngFlds = arr_varRel(R_FLDS, lngY)
          arr_varFld = arr_varRel(R_FLDARR, lngY)
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

          If (arr_varRel(R_TBLNAM1, lngY) & arr_varRel(R_TBLNAM2, lngY)) <> strLastRelName Then
            lngTmp03 = 1&
            strTmp01 = (arr_varRel(R_TBLNAM1, lngY) & arr_varRel(R_TBLNAM2, lngY))
            strLastRelName = strTmp01
          Else
            lngTmp03 = lngTmp03 + 1&
            strTmp01 = strLastRelName & CStr(lngTmp03)
          End If

          ' ** Cross-database relationships should be only used here, in the frontend!
          blnFound = True
          If arr_varRel(R_DBSID1, lngY) <> arr_varRel(R_DBSID2, lngY) Then
            'Debug.Print "'DBS: " & Parse_File(dbs.Name) & "  DB1: " & CStr(arr_varRel(R_DBSID1, lngY)) & "  DB2: " & CStr(arr_varRel(R_DBSID2, lngY))
            If Parse_File(dbs.Name) <> Parse_File(CurrentDb.Name) Then
              blnFound = False
            End If
          End If

'rel_id = 25
'rel_id = 62
          If blnFound = True Then
If Len(strTmp01) > 64 Then
Stop

End If
            Set rel1 = .CreateRelation(strTmp01, arr_varRel(R_TBLNAM1, lngY), arr_varRel(R_TBLNAM2, lngY), lngTmp00)
            For lngZ = 0& To (lngFlds - 1&)
              Set fld = rel1.CreateField(arr_varFld(FDF_FLDNAM1, lngZ))
              fld.ForeignName = arr_varFld(FDF_FLDNAM2, lngZ)
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

        .Close
      End With  ' ** dbs.
      wrk.Close

    Next  ' ** lngX.

  End Select  ' ** blnJustList.

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

Public Function Rel_Regen2() As Boolean
' ** Regenerate all relationships.
' ** Maximum name length: 64 chars.

  Const THIS_PROC As String = "Rel_Regen2"

  Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim fld As DAO.Field, tdf As DAO.TableDef, Rel As DAO.Relation
  Dim rel1 As DAO.Relation, rel2 As DAO.Relation
  Dim lngRels As Long, arr_varRel As Variant
  Dim lngFlds As Long, arr_varFld As Variant
  Dim lngRelTbls As Long, arr_varRelTbl As Variant
  Dim lngRelCnt As Long, strRel As String
  Dim strLastTblName As String, strLastRelName As String
  Dim blnFound As Boolean, blnJustList As Boolean, blnJustSysNames As Boolean
  Dim blnJustMissing As Boolean
  Dim intPos1 As Integer
  Dim lngTmp00 As Long, strTmp01 As String, strTmp02 As String, lngTmp03 As Long, blnTmp04 As Boolean
  Dim lngX As Long, lngY As Long, lngZ As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varRel().
  Const R_ID      As Integer = 0
  Const R_DBSID1  As Integer = 1
  Const R_DBSNAM1 As Integer = 2
  Const R_TBLID1  As Integer = 3
  Const R_TBLNAM1 As Integer = 4
  Const R_DBSID2  As Integer = 5
  Const R_DBSNAM2 As Integer = 6
  Const R_TBLID2  As Integer = 7
  Const R_TBLNAM2 As Integer = 8
  Const R_NAM     As Integer = 9
  Const R_ATTR    As Integer = 10
  Const R_FLDS    As Integer = 11
  Const R_FLDARR  As Integer = 12

  ' ** Array: arr_varFld().
  Const FDF_RELID   As Integer = 0
  Const FDF_DBSID   As Integer = 1
  Const FDF_DBSNAM  As Integer = 2
  Const FDF_TBLID1  As Integer = 3
  Const FDF_TBLNAM1 As Integer = 4
  Const FDF_TBLID2  As Integer = 5
  Const FDF_TBLNAM2 As Integer = 6
  Const FDF_FLDORD  As Integer = 7
  Const FDF_RFLDID  As Integer = 8
  Const FDF_FLDID1  As Integer = 9
  Const FDF_FLDNAM1 As Integer = 10
  Const FDF_FLDID2  As Integer = 11
  Const FDF_FLDNAM2 As Integer = 12
  Const FDF_DBSID2  As Integer = 14

  ' ** Array: arr_varRelTbl().
  Const RT_DBSID  As Integer = 0
  Const RT_TBLID  As Integer = 1
  Const RT_TBLNAM As Integer = 2
  Const RT_DBSIDA As Integer = 3
  Const RT_RELCNT As Integer = 4
  Const RT_ARR_F  As Integer = 5
  Const RT_ARR_L  As Integer = 6
  Const RT_DBSID2 As Integer = 7

  blnRetVal = True

  blnJustSysNames = False  ' ** True: List/Regenerate relationships with a system name; False: List/Regenerate all relationships.
  blnJustList = False      ' ** True: List the array results, don't regenerate them; False: Don't list results.
  blnJustMissing = True    ' ** True: List/Regenerate missing relationships only.

  Set dbs = CurrentDb
  With dbs

    ' ** Get a list of currently documented relationships

    Select Case blnJustSysNames
    Case True
      ' ** tblRelation, just TrstXAdm.mdb relationships having a system name, with table names.
      Set qdf = .QueryDefs("qryRelation_02")
    Case False
      ' ** tblRelation, all TrstXAdm.mdb relationships, with table names.
      Set qdf = .QueryDefs("qryRelation_01")
    End Select

    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRels = .RecordCount
      .MoveFirst
      .sort = "dbs_id1, tbl_name1, tbl_name2"
      Set rst2 = .OpenRecordset
      rst2.MoveLast
      lngRels = rst2.RecordCount
      rst2.MoveFirst
      arr_varRel = rst2.GetRows(lngRels)
      ' ****************************************************
      ' ** Array: arr_varRel()
      ' **
      ' **   Field  Element  Name              Constant
      ' **   =====  =======  ================  ===========
      ' **     1       0     rel_id            R_ID
      ' **     2       1     dbs_id1           R_DBSID1
      ' **     3       2     dbs_name1         R_DBSNAM1
      ' **     4       3     tbl_id1           R_TBLID1
      ' **     5       4     tbl_name1         R_TBLNAM1
      ' **     6       5     dbs_id2           R_DBSID2
      ' **     7       6     dbs_name2         R_DBSNAM2
      ' **     8       7     tbl_id2           R_TBLID2
      ' **     9       8     tbl_name2         R_TBLNAM2
      ' **    10       9     rel_name          R_NAM
      ' **    11      10     rel_attributes    R_ATTR
      ' **    12      11     rel_fld_cnt       R_FLDS
      ' **    13      12     rel_fld_arr       R_FLDARR
      ' **    14      13     DontEnforce
      ' **
      ' ****************************************************
      rst2.Close
      .Close
    End With  ' ** rst1.
    Set rst1 = Nothing
    Set rst2 = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    Debug.Print "'RELS GOAL: " & CStr(lngRels)
    DoEvents

    ' ** Get lists of relationship fields

    For lngX = 0& To (lngRels - 1&)

      Select Case blnJustSysNames
      Case True
        ' ** qryRelation_03 (tblRelation, linked to tblRelation_Field, with add'l fields),
        ' ** having a system name, with add'l foreign table fields, by specified [relid].
        Set qdf = .QueryDefs("qryRelation_05")
      Case False
        ' ** qryRelation_03 (tblRelation, linked to tblRelation_Field, with add'l fields),
        ' ** with add'l foreign table fields, by specified [relid].
        Set qdf = .QueryDefs("qryRelation_04")
      End Select

      With qdf.Parameters
        ![relid] = arr_varRel(R_ID, lngX)
      End With
      Set rst1 = qdf.OpenRecordset
      With rst1
        .MoveLast
        lngFlds = .RecordCount
        .MoveFirst
        .sort = "dbs_id, tbl_name1, tbl_name2, relfld_order, fld_name1, fld_name2"
        Set rst2 = .OpenRecordset
        rst2.MoveLast
        lngFlds = rst2.RecordCount
        rst2.MoveFirst
        arr_varFld = rst2.GetRows(lngFlds)
        ' ****************************************************
        ' ** Array: arr_varFld()
        ' **
        ' **   Field  Element  Name            Constant
        ' **   =====  =======  ==============  =============
        ' **     1       0     rel_id          FDF_RELID
        ' **     2       1     dbs_id          FDF_DBSID
        ' **     3       2     dbs_name        FDF_DBSNAM
        ' **     4       3     tbl_id1         FDF_TBLID1
        ' **     5       4     tbl_name1       FDF_TBLNAM1
        ' **     6       5     tbl_id2         FDF_TBLID2
        ' **     7       6     tbl_name2       FDF_TBLNAM2
        ' **     8       7     relfld_order    FDF_FLDORD
        ' **     9       8     relfld_id       FDF_RFLDID
        ' **    10       9     fld_id1         FDF_FLDID1
        ' **    11      10     fld_name1       FDF_FLDNAM1
        ' **    12      11     fld_id2         FDF_FLDID2
        ' **    13      12     fld_name2       FDF_FLDNAM2
        ' **    14      13     DontEnforce
        ' **    15      14     dbs_id2         FDF_DBSID2
        ' **
        ' ****************************************************
      End With
      arr_varRel(R_FLDARR, lngX) = arr_varFld
      Set qdf = Nothing
      lngFlds = 0&
      arr_varFld = Empty
    Next
    Set rst1 = Nothing
    Set rst2 = Nothing

    ' ** Get list of tables with relationships

    Select Case blnJustSysNames
    Case True
      ' ** qryRelation_02 (tblRelation, just TrustAux.mdb relationships having
      ' ** a system name, with table names), grouped by tbl_id1, with rel_cnt.
      Set qdf = .QueryDefs("qryRelation_07")
    Case False
      ' ** qryRelation_01 (tblRelation, all TrustAux.mdb relationships,
      ' ** with table names), grouped by tbl_id1, with rel_cnt.
      Set qdf = .QueryDefs("qryRelation_06")
    End Select

    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRelTbls = .RecordCount
      .MoveFirst
      .sort = "[dbs_id], [tbl_name]"
      Set rst2 = .OpenRecordset
      rst2.MoveLast
      lngRelTbls = rst2.RecordCount
      rst2.MoveFirst
      arr_varRelTbl = rst2.GetRows(lngRelTbls)
      ' ***************************************************
      ' ** Array: arr_varRelTbl()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ===========
      ' **     1       0     dbs_id           RT_DBSID
      ' **     2       1     tbl_id           RT_TBLID
      ' **     3       2     tbl_name         RT_TBLNAM
      ' **     4       3     dbs_id_asof      RT_DBSIDA
      ' **     5       4     rel_cnt          RT_RELCNT
      ' **     6       5     rel_arr_first    RT_ARR_F
      ' **     7       6     rel_arr_last     RT_ARR_L
      ' **     8       7     dbs_id2          RT_DBSID2
      ' **
      ' ***************************************************
      .Close
    End With
    Set rst1 = Nothing
    Set rst2 = Nothing

    ' ** Put the first and last arr_varRel() array elements, for each table, into the arr_varRelTbl() array for that table.
    strLastTblName = vbNullString
    For lngX = 0& To (lngRelTbls - 1&)
      For lngY = 0& To (lngRels - 1&)
        If ((arr_varRel(R_TBLNAM1, lngY) = arr_varRelTbl(RT_TBLNAM, lngX)) And _
            (arr_varRel(R_DBSID2, lngY) = arr_varRelTbl(RT_DBSID2, lngX))) Then
          If arr_varRelTbl(RT_TBLNAM, lngX) <> strLastTblName Then
            arr_varRelTbl(RT_ARR_F, lngX) = lngY
            If lngX > 0& Then
              arr_varRelTbl(RT_ARR_L, lngX - 1&) = lngY - 1&
            End If
            strLastTblName = arr_varRelTbl(RT_TBLNAM, lngX)
          End If
        End If
      Next
      If lngX = (lngRelTbls - 1&) Then
        arr_varRelTbl(RT_ARR_L, lngX) = CLng(UBound(arr_varRel, 2))
      End If
    Next

    .Close
  End With

  Select Case blnJustList
  Case True

    For lngX = 0& To (lngRelTbls - 1&)
      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      Debug.Print "'" & Left$(CStr(lngX + 1&) & "   ", 3) & ": " & Left$(CStr(arr_varRelTbl(RT_ARR_F, lngX)) & "   ", 3) & " - " & _
        Left$(CStr(arr_varRelTbl(RT_ARR_L, lngX)) & "   ", 3) & "  " & arr_varRelTbl(RT_TBLNAM, lngX)
    Next

  Case False  ' ** blnJustList.

    strLastRelName = vbNullString
    blnTmp04 = False

    ' ** For each table with one or more relationships.
    For lngX = 0& To (lngRelTbls - 1&)

      ' ***************************************************
      ' ** Array: arr_varRelTbl()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ===========
      ' **     1       0     dbs_id           RT_DBSID
      ' **     2       1     tbl_id           RT_TBLID
      ' **     3       2     tbl_name         RT_TBLNAM
      ' **     4       3     dbs_id_asof      RT_DBSIDA
      ' **     5       4     rel_cnt          RT_RELCNT
      ' **     6       5     rel_arr_first    RT_ARR_F
      ' **     7       6     rel_arr_last     RT_ARR_L
      ' **
      ' ***************************************************

      strTmp01 = DLookup("[dbs_name]", "tblDatabase", "[dbs_id] = " & CStr(arr_varRelTbl(RT_DBSIDA, lngX)))
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
            Set wrk = CreateWorkspace("tmp", "superuser", TA_SEC2, dbUseJet)
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
          strTmp02 = DLookup("[dbs_path]", "tblDatabase", "[dbs_id] = " & CStr(arr_varRelTbl(RT_DBSID, lngX)))
          If Right$(strTmp02, 1) = LNK_SEP Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
        End If
        Set dbs = wrk.OpenDatabase(strTmp02 & LNK_SEP & strTmp01, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
      End If

      With dbs

        If blnJustMissing = False Then
          ' ** Delete all relationships for this table.
          .Relations.Refresh
          lngTmp00 = .Relations.Count
          For lngY = (lngTmp00 - 1&) To 0 Step -1&
            If .Relations(lngY).Table = arr_varRelTbl(RT_TBLNAM, lngX) Then
              .Relations.Delete .Relations(lngY).Name
              .Relations.Refresh
            End If
          Next
        End If

        ' ** Now re-create the relationships.
        For lngY = arr_varRelTbl(RT_ARR_F, lngX) To arr_varRelTbl(RT_ARR_L, lngX)

          ' ** Define the relationship.
          lngFlds = arr_varRel(R_FLDS, lngY)
          arr_varFld = arr_varRel(R_FLDARR, lngY)
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

          If (arr_varRel(R_TBLNAM1, lngY) & arr_varRel(R_TBLNAM2, lngY)) <> strLastRelName Then
            lngTmp03 = 1&
            strTmp01 = (arr_varRel(R_TBLNAM1, lngY) & arr_varRel(R_TBLNAM2, lngY))
            strLastRelName = strTmp01
          Else
            lngTmp03 = lngTmp03 + 1&
            strTmp01 = strLastRelName & CStr(lngTmp03)
          End If

          blnFound = False
          For Each Rel In .Relations
            With Rel
              If .Table = arr_varRel(R_TBLNAM1, lngY) And .ForeignTable = arr_varRel(R_TBLNAM2, lngY) Then
                If .Name = strTmp01 Then
                  If arr_varRel(R_FLDS, lngY) = .Fields.Count Then
                    blnFound = True
                  Else
                    Stop
                  End If
                  Exit For
                End If
              End If
            End With  ' ** rel.
          Next  ' ** rel.
          If blnFound = False Then
            Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
            'Debug.Print "'NOT FOUND! " & strTmp01 & "  "
          End If

          If blnFound = False Then
If Len(strTmp01) > 64 Then
Stop

End If
            Set rel1 = .CreateRelation(strTmp01, arr_varRel(R_TBLNAM1, lngY), arr_varRel(R_TBLNAM2, lngY), lngTmp00)
            For lngZ = 0& To (lngFlds - 1&)
              Set fld = rel1.CreateField(arr_varFld(FDF_FLDNAM1, lngZ))
              fld.ForeignName = arr_varFld(FDF_FLDNAM2, lngZ)
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

        .Close
      End With  ' ** dbs.
      wrk.Close

    Next  ' ** lngX.

  End Select  ' ** blnJustList.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

'1  : 0   - 1    tblBackStyle
'2  : 2   - 3    tblBorderWidth
'3  : 4   - 4    tblButtonType
'4  : 5   - 5    tblCellEffect
'5  : 6   - 6    tblComponentType
'6  : 7   - 8    tblConnectionType
'7  : 9   - 10   tblControlBorderStyle
'8  : 11  - 18   tblControlType
'9  : 19  - 19   tblCycle
'10 : 20  - 20   tblDAOType
'11 : 21  - 22   tblDatabase
'12 : 23  - 28   tblDatabase_Table
'13 : 29  - 33   tblDatabase_Table_Field
'14 : 34  - 38   tblDataTypeDb
'15 : 39  - 44   tblDataTypeVb
'16 : 45  - 45   tblDateGroupingType
'17 : 46  - 46   tblDecimalPlaceDb
'18 : 47  - 47   tblDeclarationType
'19 : 48  - 48   tblDefaultView
'20 : 49  - 49   tblDisplayWhenType
'21 : 50  - 50   tblDocument
'22 : 51  - 51   tblDocument_Image
'23 : 52  - 52   tblDocumentAutoShapeType
'24 : 53  - 53   tblDocumentFieldKind
'25 : 54  - 54   tblDocumentFieldType
'26 : 55  - 55   tblDocumentInlineShapeType
'27 : 56  - 56   tblDocumentLinkType
'28 : 57  - 57   tblDocumentShapeType
'29 : 58  - 60   tblFontWeight
'30 : 61  - 62   tblForceNewPageType
'31 : 63  - 70   tblForm
'32 : 71  - 85   tblForm_Control
'33 : 86  - 86   tblForm_Control_Group
'34 : 87  - 87   tblForm_Control_Specification_A
'35 : 88  - 91   tblForm_Section
'36 : 92  - 92   tblForm_Specification_A
'37 : 93  - 93   tblFormBorderStyle
'38 : 94  - 94   tblGridlineBehavior
'39 : 95  - 96   tblGridlineStyle
'40 : 97  - 97   tblGroupKeepTogetherType
'41 : 98  - 98   tblIndex
'42 : 99  - 101  tblJournalType
'43 : 102 - 103  tblKeyCode
'44 : 104 - 104  tblLineSlantType
'45 : 105 - 106  tblMacro
'46 : 107 - 107  tblMacro_Row
'47 : 108 - 109  tblMacroAction
'48 : 110 - 110  tblMacroActionArgument
'49 : 111 - 111  tblMinMaxButton
'50 : 112 - 115  tblMsgBoxStyleType
'51 : 116 - 116  tblMultiSelectType
'52 : 117 - 118  tblNewRowOrColType
'53 : 119 - 131  tblObjectType
'54 : 132 - 133  tblPageHeaderFooterType
'55 : 134 - 134  tblParameterDataType
'56 : 135 - 135  tblParameterDirection
'57 : 136 - 139  tblPictureAlignmentType
'58 : 140 - 140  tblPicturePageType
'59 : 141 - 144  tblPictureSizeMode
'60 : 145 - 148  tblPictureType
'61 : 149 - 149  tblPreference_Control
'62 : 150 - 150  tblProcedureSubType
'63 : 151 - 153  tblProcedureType
'64 : 154 - 158  tblQuery
'65 : 159 - 162  tblQueryTableType
'66 : 163 - 164  tblQueryType
'67 : 165 - 166  tblRecordLock
'68 : 167 - 167  tblRecordsetType
'69 : 168 - 168  tblReferenceType
'70 : 169 - 169  tblRelation
'71 : 170 - 170  tblRelation_View
'72 : 171 - 177  tblReport
'73 : 178 - 179  tblReport_Control
'74 : 180 - 181  tblReport_Section
'75 : 182 - 182  tblReportGroupOnType
'76 : 183 - 183  tblReportKeepTogetherType
'77 : 184 - 184  tblReportOrientationType
'78 : 185 - 185  tblReportSortOrder
'79 : 186 - 188  tblRowSourceType
'80 : 189 - 191  tblScopeType
'81 : 192 - 192  tblScriptingType
'82 : 193 - 193  tblScrollBarAlignment
'83 : 194 - 194  tblScrollBarC
'84 : 195 - 195  tblScrollBarF
'85 : 196 - 199  tblSectionType
'86 : 200 - 203  tblSpecialEffect
'87 : 204 - 204  tblStyleType
'88 : 205 - 207  tblSystemColor_Base
'89 : 208 - 208  tblSystemColorType
'90 : 209 - 209  tblTaxCodeType
'91 : 210 - 211  tblTextAlignType
'92 : 212 - 212  tblTransactionForm
'93 : 213 - 213  tblUnderlineStyle
'94 : 214 - 214  tblVBAType
'95 : 215 - 217  tblVBComponent
'96 : 218 - 218  tblVBComponent_API
'97 : 219 - 219  tblVBComponent_Declaration
'98 : 220 - 220  tblVBComponent_Declaration_Family
'99 : 221 - 221  tblVBComponent_Event
'100: 222 - 227  tblVBComponent_Procedure
'101: 228 - 228  tblVBComponent_Property
'102: 229 - 229  tblVersion
'103: 230 - 230  tblVersion_Directory
'104: 231 - 233  tblVersion_File
'105: 234 - 234  tblVersion_Table
'106: 235 - 235  tblVersionType
'107: 236 - 236  tblViewsAllowed
'108: 237 - 239  tblXAdmin_Graphics
'109: 240 - 240  tblXAdmin_Graphics_Format
'110: 241 - 241  tblXAdmin_Graphics_Type
'DONE!  Rel_Regen()

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set Rel = Nothing
  Set fld = Nothing
  Set tdf = Nothing
  Set rel1 = Nothing
  Set rel2 = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set wrk = Nothing

  Rel_Regen2 = blnRetVal

End Function

Public Function Rel_Find() As Boolean

  Const THIS_PROC As String = "Rel_Find"

  Dim dbs As DAO.Database, Rel As DAO.Relation
  Dim lngRels As Long, arr_varRel() As Variant
  Dim strTmp00 As String
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
      strTmp00 = vbNullString
      With Rel
        If .Name = arr_varRel(R_NAM, lngX) Then
        'If .ForeignTable = "tblRelation_View" Then
          Debug.Print "'" & CStr(lngX + 1&) & ". REL: " & .Name & "  TBL1: " & .Table & "  TBL2: " & .ForeignTable
          Debug.Print "' ATTR: " & CStr(.Attributes)
          If Rel_Attr(.Attributes, "dbRelationDontEnforce") = True Then strTmp00 = strTmp00 & "dbRelationDontEnforce, "
          If Rel_Attr(.Attributes, "dbRelationEnforce") = True Then strTmp00 = strTmp00 & "dbRelationEnforce, "
          If Rel_Attr(.Attributes, "dbRelationInherited") = True Then strTmp00 = strTmp00 & "dbRelationInherited, "
          If Rel_Attr(.Attributes, "dbRelationUpdateCascade") = True Then strTmp00 = strTmp00 & "dbRelationUpdateCascade, "
          If Rel_Attr(.Attributes, "dbRelationDeleteCascade") = True Then strTmp00 = strTmp00 & "dbRelationDeleteCascade, "
          If Rel_Attr(.Attributes, "dbRelationLeft") = True Then strTmp00 = strTmp00 & "dbRelationLeft, "
          If Rel_Attr(.Attributes, "dbRelationRight") = True Then strTmp00 = strTmp00 & "dbRelationRight, "
          If Rel_Attr(.Attributes, "dbRelationUnique") = True Then strTmp00 = strTmp00 & "dbRelationUnique, "
          strTmp00 = Trim$(strTmp00)
          If Right$(strTmp00, 1) = "," Then strTmp00 = Left$(strTmp00, (Len(strTmp00) - 1))
          Debug.Print "'   " & strTmp00
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

'WHERE ARE THESE?
'271 tblDatabase_Table                  ->  tblXAdmin_ExportTbl
'272 tblDatabase                        ->  tblRelation_View
'276 tblQuery                           ->  tblXAdmin_ExportQry

'RELS: 292
'1   tblBackStyle                       ->  tblForm_Control_Specification_A
'2   tblBackStyle                       ->  tblReport_Control
'3   tblBorderWidth                     ->  tblForm_Control_Specification_A
'4   tblBorderWidth                     ->  tblReport_Control
'5   tblButtonType                      ->  tblMsgBoxStyleType
'6   tblCellEffect                      ->  tblForm_Specification_A
'7   tblComponentType                   ->  tblVBComponent
'8   tblConnectionType                  ->  tblDatabase_Table_Link
'9   tblConnectionType                  ->  tblXAdmin_Graphics
'10  tblControlBorderStyle              ->  tblForm_Control_Specification_A
'11  tblControlBorderStyle              ->  tblReport_Control
'12  tblControlType                     ->  tblDatabase_Table_Field
'13  tblControlType                     ->  tblDatabase_Table_Field_RowSource
'14  tblControlType                     ->  tblForm_Control
'15  tblControlType                     ->  tblForm_Shortcut
'16  tblControlType                     ->  tblObject_Image
'17  tblControlType                     ->  tblReport_Control
'18  tblCycle                           ->  tblForm_Specification_A
'19  tblDAOType                         ->  tblVBComponent_Procedure_Parameter
'20  tblDatabase_Table_Field            ->  tblDatabase_AutoNumber
'21  tblDatabase_Table_Field            ->  tblDatabase_Table_Field_DateFormat
'22  tblDatabase_Table_Field            ->  tblDatabase_Table_Field_RowSource
'23  tblDatabase_Table_Field            ->  tblIndex_Field
'24  tblDatabase_Table_Field            ->  tblRelation_Field
'25  tblDatabase_Table                  ->  tblDatabase_Table_Field
'26  tblDatabase_Table                  ->  tblDatabase_Table_Link
'27  tblDatabase_Table                  ->  tblDatabase_Table_RecordCount
'28  tblDatabase_Table                  ->  tblIndex
'29  tblDatabase_Table                  ->  tblRelation
'30  tblDatabase_Table                  ->  tblRelation
'31  tblDatabase                        ->  tblDatabase_Table
'32  tblDataTypeDb                      ->  tblDatabase_Table_Field
'33  tblDataTypeDb                      ->  tblReport_Group
'34  tblDataTypeVb                      ->  tblMacroActionArgument
'35  tblDataTypeVb                      ->  tblVBComponent_API
'36  tblDataTypeVb                      ->  tblVBComponent_Procedure
'37  tblDataTypeVb                      ->  tblVBComponent_Procedure_Parameter
'38  tblDataTypeVb                      ->  tblVBComponent_Property
'39  tblDateGroupingType                ->  tblReport_Specification
'40  tblDecimalPlaceDb                  ->  tblDatabase_Table_Field
'41  tblDefaultView                     ->  tblForm_Specification_A
'42  tblDisplayWhenType                 ->  tblForm_Section
'43  tblDocument_Image                  ->  tblDocument_Document_Image
'44  tblDocumentAutoShapeType           ->  tblDocument_Document_Image
'45  tblDocumentFieldKind               ->  tblDocument_Document_Image
'46  tblDocumentFieldType               ->  tblDocument_Document_Image
'47  tblDocumentInlineShapeType         ->  tblDocument_Document_Image
'48  tblDocumentLinkType                ->  tblDocument_Document_Image
'49  tblDocumentShapeType               ->  tblDocument_Document_Image
'50  tblDocument                        ->  tblDocument_Document_Image
'51  tblFontWeight                      ->  tblForm_Control_Specification_A
'52  tblFontWeight                      ->  tblForm_Specification_A
'53  tblFontWeight                      ->  tblReport_Control
'54  tblForceNewPageType                ->  tblForm_Section
'55  tblForceNewPageType                ->  tblReport_Section
'56  tblForm_Control_Specification_A    ->  tblForm_Control_Specification_B
'57  tblForm_Control                    ->  tblForm_Control_RowSource
'58  tblForm_Control                    ->  tblForm_Control_Specification_A
'59  tblForm_Control                    ->  tblForm_Graphics
'60  tblForm_Control                    ->  tblForm_Shortcut
'61  tblForm_Control                    ->  tblForm_Subform
'62  tblForm_Control                    ->  tblJournal_Field
'63  tblForm_Control                    ->  tblPreference_Control
'64  tblForm_Control                    ->  tblReport_List
'65  tblForm_Control                    ->  tblReport_List
'66  tblForm_Control                    ->  tblReport_List
'67  tblForm_Control                    ->  tblReport_List
'68  tblForm_Control                    ->  tblReport_List
'69  tblForm_Control                    ->  tblSystemColor_Control
'70  tblForm_Section                    ->  tblForm_Control
'71  tblForm_Section                    ->  tblJournal_Field
'72  tblForm_Section                    ->  tblSystemColor_Section
'73  tblForm_Specification_A            ->  tblForm_Specification_B
'74  tblFormBorderStyle                 ->  tblForm_Specification_A
'75  tblForm                            ->  tblForm_Control
'76  tblForm                            ->  tblForm_Graphics
'77  tblForm                            ->  tblForm_RecordSource
'78  tblForm                            ->  tblForm_Section
'79  tblForm                            ->  tblForm_Shortcut
'80  tblForm                            ->  tblForm_Specification_A
'81  tblForm                            ->  tblForm_Subform
'82  tblForm                            ->  tblReport_List
'83  tblForm                            ->  tblTransactionForm
'84  tblGridlineBehavior                ->  tblForm_Specification_A
'85  tblGridlineStyle                   ->  tblForm_Specification_A
'86  tblGridlineStyle                   ->  tblForm_Specification_B
'87  tblGroupKeepTogetherType           ->  tblReport_Specification
'88  tblIndex                           ->  tblIndex_Field
'89  tblJournalType                     ->  tblJournal_Field
'90  tblJournalType                     ->  tblJournalSubType
'91  tblKeyCode                         ->  tblForm_Shortcut
'92  tblKeyCode                         ->  tblKeyboard_Monitor
'93  tblLineSlantType                   ->  tblForm_Control_Specification_A
'94  tblMacro_Row                       ->  tblMacro_Row_Argument
'95  tblMacroActionArgument             ->  tblMacro_Row_Argument
'96  tblMacroAction                     ->  tblMacro_Row
'97  tblMacroAction                     ->  tblMacroActionArgument
'98  tblMacro                           ->  tblMacro_Row
'99  tblMacro                           ->  tblMacro_Text
'100 tblMinMaxButton                    ->  tblForm_Specification_B
'101 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'102 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'103 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'104 tblMsgBoxStyleType                 ->  tblVBComponent_MessageBox
'105 tblMultiSelectType                 ->  tblForm_Control_RowSource
'106 tblNewRowOrColType                 ->  tblForm_Section
'107 tblNewRowOrColType                 ->  tblReport_Section
'108 tblObjectType                      ->  tblForm
'109 tblObjectType                      ->  tblForm_Control
'110 tblObjectType                      ->  tblForm_Section
'111 tblObjectType                      ->  tblObject_Image
'112 tblObjectType                      ->  tblQueryTableType
'113 tblObjectType                      ->  tblReport
'114 tblObjectType                      ->  tblReport_Control
'115 tblObjectType                      ->  tblReport_Section
'116 tblObjectType                      ->  tblSystemColor_Control
'117 tblObjectType                      ->  tblSystemColor_Section
'118 tblObjectType                      ->  tblVBComponent
'119 tblObjectType                      ->  tblVBComponent_Procedure_Parameter
'120 tblPageHeaderFooterType            ->  tblReport_Specification
'121 tblPageHeaderFooterType            ->  tblReport_Specification
'122 tblParameterDataType               ->  tblQuery_Parameter
'123 tblParameterDirection              ->  tblQuery_Parameter
'124 tblPictureAlignmentType            ->  tblForm_Control_Specification_B
'125 tblPictureAlignmentType            ->  tblForm_Specification_B
'126 tblPictureAlignmentType            ->  tblObject_Image
'127 tblPictureAlignmentType            ->  tblReport_Specification
'128 tblPicturePageType                 ->  tblReport_Specification
'129 tblPictureSizeMode                 ->  tblForm_Control_Specification_B
'130 tblPictureSizeMode                 ->  tblForm_Specification_B
'131 tblPictureSizeMode                 ->  tblObject_Image
'132 tblPictureSizeMode                 ->  tblReport_Specification
'133 tblPictureType                     ->  tblForm_Control_Specification_B
'134 tblPictureType                     ->  tblForm_Specification_B
'135 tblPictureType                     ->  tblObject_Image
'136 tblPictureType                     ->  tblReport_Specification
'137 tblPreference_Control              ->  tblPreference_User
'138 tblProcedureSubType                ->  tblVBComponent_Procedure
'139 tblProcedureType                   ->  tblProcedureSubType
'140 tblProcedureType                   ->  tblVBComponent_API
'141 tblProcedureType                   ->  tblVBComponent_Procedure
'142 tblQueryTableType                  ->  tblDatabase_Table_Field_RowSource
'143 tblQueryTableType                  ->  tblForm_Control_RowSource
'144 tblQueryTableType                  ->  tblForm_RecordSource
'145 tblQueryTableType                  ->  tblReport_RecordSource
'146 tblQuery                           ->  tblQuery_Field
'147 tblQuery                           ->  tblQuery_FormRef
'148 tblQuery                           ->  tblQuery_Parameter
'149 tblQuery                           ->  tblQuery_RecordSource
'150 tblQuery                           ->  tblQuery_SourceChain
'151 tblQueryType                       ->  tblQuery
'152 tblQueryType                       ->  tblXAdmin_Graphics
'153 tblRecordLock                      ->  tblForm_Specification_B
'154 tblRecordLock                      ->  tblReport_Specification
'155 tblRecordsetType                   ->  tblForm_Specification_B
'156 tblReferenceType                   ->  tblReference
'157 tblRelation                        ->  tblRelation_Field
'158 tblReport_Control                  ->  tblReport_Subform
'159 tblReport_Control                  ->  tblSystemColor_Control
'160 tblReport_Section                  ->  tblReport_Control
'161 tblReport_Section                  ->  tblSystemColor_Section
'162 tblReportGroupOnType               ->  tblReport_Group
'163 tblReportKeepTogetherType          ->  tblReport_Group
'164 tblReportOrientationType           ->  tblReport_Specification
'165 tblReportSortOrder                 ->  tblReport_Group
'166 tblReport                          ->  tblReport_Control
'167 tblReport                          ->  tblReport_Group
'168 tblReport                          ->  tblReport_RecordSource
'169 tblReport                          ->  tblReport_Section
'170 tblReport                          ->  tblReport_Specification
'171 tblReport                          ->  tblReport_Subform
'172 tblReport                          ->  tblReport_VBComponent
'173 tblRowSourceType                   ->  tblDatabase_Table_Field_RowSource
'174 tblRowSourceType                   ->  tblForm_Control_RowSource
'175 tblRowSourceType                   ->  tblQueryTableType
'176 tblScopeType                       ->  tblVBComponent_API
'177 tblScopeType                       ->  tblVBComponent_Procedure
'178 tblScriptingType                   ->  tblVBComponent_Procedure_Parameter
'179 tblScrollBarAlignment              ->  tblForm_Control_Specification_B
'180 tblScrollBarC                      ->  tblForm_Control_Specification_B
'181 tblScrollBarF                      ->  tblForm_Specification_B
'182 tblSpecialEffect                   ->  tblForm_Control_Specification_B
'183 tblSpecialEffect                   ->  tblForm_Section
'184 tblSpecialEffect                   ->  tblReport_Control
'185 tblSpecialEffect                   ->  tblReport_Section
'186 tblStyleType                       ->  tblForm_Control_Specification_B
'187 tblSystemColor_Base                ->  tblSystemColor
'188 tblSystemColor_Base                ->  tblSystemColor_Control
'189 tblSystemColor_Base                ->  tblSystemColor_Section
'190 tblSystemColorType                 ->  tblSystemColor
'191 tblTaxCodeType                     ->  tblTaxCode
'192 tblTextAlignType                   ->  tblForm_Control_Specification_B
'193 tblTextAlignType                   ->  tblReport_Control
'194 tblTransactionForm                 ->  tblTransactionForm_Option
'195 tblUnderlineStyle                  ->  tblForm_Specification_A
'196 tblVBAType                         ->  tblReference
'197 tblVBComponent_Event               ->  tblVBComponent_Procedure
'198 tblVBComponent_Procedure           ->  tblReport_VBComponent
'199 tblVBComponent_Procedure           ->  tblVBComponent_MessageBox
'200 tblVBComponent_Procedure           ->  tblVBComponent_Procedure_Parameter
'201 tblVBComponent_Property            ->  tblVBComponent_Event
'202 tblVBComponent                     ->  tblVBComponent_API
'203 tblVBComponent                     ->  tblVBComponent_Procedure
'204 tblVersion_Directory               ->  tblVersion_File
'205 tblVersion_File                    ->  tblVersion_Table
'206 tblVersion_Table                   ->  tblVersion_Field
'207 tblViewsAllowed                    ->  tblForm_Specification_B
'208 tblXAdmin_Graphics_Format          ->  tblXAdmin_Graphics
'209 tblXAdmin_Graphics_Type            ->  tblXAdmin_Graphics
'210 tblXAdmin_Graphics                 ->  tblForm_Graphics
'211 tblXAdmin_Graphics                 ->  tblXAdmin_ExportQry
'212 tblXAdmin_Graphics                 ->  tblXAdmin_ExportTbl
'213 account                            ->  ActiveAssets
'214 account                            ->  asset
'215 account                            ->  Balance
'216 account                            ->  journal
'217 account                            ->  ledger
'218 account                            ->  PortfolioModel
'219 AccountType                        ->  account
'220 AccountType                        ->  AccountTypeGrouping
'221 adminofficer                       ->  account
'222 AssetType                          ->  AssetTypeGrouping
'223 AssetType                          ->  masterasset
'224 AssetType                          ->  tblPricing_MasterAsset_History
'225 HiddenType                         ->  LedgerHidden
'226 InvestmentObjective                ->  account
'227 journal                            ->  tblJournal_Import
'228 journal                            ->  tblJournal_Memo
'229 journaltype                        ->  journal
'230 journaltype                        ->  ledger
'231 journaltype                        ->  RecurringType
'232 Location                           ->  ActiveAssets
'233 Location                           ->  journal
'234 Location                           ->  ledger
'235 m_REVCODE_TYPE                     ->  m_REVCODE
'236 m_REVCODE_TYPE                     ->  TaxCode
'237 m_REVCODE                          ->  journal
'238 m_REVCODE                          ->  ledger
'239 masterasset                        ->  ActiveAssets
'240 masterasset                        ->  asset
'241 masterasset                        ->  tblPricing_MasterAsset_History
'242 RecurringItems                     ->  journal
'243 RecurringItems                     ->  ledger
'244 RecurringType                      ->  RecurringItems
'245 Schedule                           ->  account
'246 Schedule                           ->  ScheduleDetail
'247 TaxCode_Type                       ->  TaxCode
'248 TaxCode                            ->  AssetType
'249 TaxCode                            ->  journal
'250 TaxCode                            ->  ledger
'251 tblDataTypeDb1                     ->  tblPricing_AppraiseColumnDataType
'252 tblDataTypeDb1                     ->  tblPricing_AppraiseItemType
'253 tblPricing_AppraiseColumnDataType  ->  tblPricing_AppraiseColumn
'254 tblPricing_AppraiseColumnIDQuote   ->  tblPricing_AppraiseColumn
'255 tblPricing_AppraiseFile            ->  tblPricing_AppraiseColumn
'256 tblPricing_AppraiseItemType        ->  tblPricing_AppraiseColumn
'257 tblPricing_AppraiseRowType         ->  tblPricing_AppraiseColumn
'258 tblPricing_AppraiseRowType         ->  tblPricing_AppraiseFile
'259 tblPricing_AppraiseSectionType     ->  tblPricing_AppraiseColumnDataType
'260 tblPricing_AppraiseSectionType     ->  tblPricing_AppraiseItemType
'261 tblPricing_Import                  ->  tblPricing_FileType
'262 journaltype                        ->  tblJournalType
'263 m_REVCODE_TYPE                     ->  tblTaxCode
'264 TaxCode_Type                       ->  tblTaxCodeType
'265 TaxCode                            ->  tblTaxCode
'266 tblCheckReconcile_Account          ->  tblCheckReconcile_Check
'267 tblCheckReconcile_Account          ->  tblCheckReconcile_Item
'268 tblCheckReconcileEntryType         ->  tblCheckReconcile_Item
'269 tblCheckReconcileSourceType        ->  tblCheckReconcile_Item
'270 tblConnectionType                  ->  tblXAdmin_Import
'271 tblDatabase_Table                  ->  tblXAdmin_ExportTbl
'272 tblDatabase                        ->  tblRelation_View
'273 tblForm_Control_Group              ->  tblForm_Control_Group_Item
'274 tblImportExport_Specifications     ->  tblImportExport_Columns
'275 tblObjectType                      ->  tblXAdmin_Import
'276 tblQuery                           ->  tblXAdmin_ExportQry
'277 tblQueryType                       ->  tblXAdmin_Import
'278 tblRelation_View                   ->  tblRelation_View_Window
'279 tblReport                          ->  zz_tbl_Report_VBComponent_01
'280 tblSecurity_Group                  ->  tblSecurity_GroupUser
'281 tblSecurity_User                   ->  tblSecurity_GroupUser
'282 tblVBComponent_Declaration_Family  ->  tblVBComponent_Declaration
'283 tblVBComponent_Declaration         ->  tblVBComponent_Declaration_Detail
'284 tblVersion                         ->  tblVersion_Release
'285 tblVersionType                     ->  tblVersion_Release
'286 tblXAdmin_Customer                 ->  tblXAdmin_Import
'287 zz_tbl_Client_Directory            ->  zz_tbl_Client_File
'288 zz_tbl_Client_File                 ->  zz_tbl_Client_Table
'289 zz_tbl_Client_Table                ->  zz_tbl_Client_Field
'290 zz_tbl_Dev_Directory               ->  zz_tbl_Dev_File
'291 zz_tbl_Report_VBComponent_01       ->  zz_tbl_Report_VBComponent_02
'292 zz_tbl_Report_VBComponent_02       ->  zz_tbl_Report_VBComponent_03

  Set Rel = Nothing
  Set dbs = Nothing

  Beep

  Rel_List = blnRetValx

End Function

Public Function Rel_QryImport() As Boolean

  Const THIS_PROC As String = "Rel_QryImport"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry As Variant
  Dim strPathFile As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_DID  As Integer = 0
  Const Q_DNAM As Integer = 1
  Const Q_QID  As Integer = 2
  Const Q_QNAM As Integer = 3

  blnRetVal = True

  DoCmd.Hourglass True
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** tblQuery, just 'zz_qry_Relation_..' queries in Trust.mdb.
    Set qdf = .QueryDefs("zz_qry_Relation_30x")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngQrys = .RecordCount
      .MoveFirst
      arr_varQry = .GetRows(lngQrys)
      ' ***************************************************
      ' ** Array: arr_varQry()
      ' **
      ' **   Field  Element  Name              Constant
      ' **   =====  =======  ================  ==========
      ' **     1       0     dbs_id            Q_DID
      ' **     2       1     dbs_name          Q_DNAM
      ' **     3       2     qry_id            Q_QID
      ' **     4       3     qry_name          Q_QNAM
      ' **
      ' ***************************************************
      .Close
    End With
    .Close
  End With
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  strPathFile = gstrDir_Dev & LNK_SEP & "Trust.mdb"

  For lngX = 0& To (lngQrys - 1&)
    DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
  Next

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  DoCmd.Hourglass False

  Beep

  Rel_QryImport = blnRetVal

End Function

Private Function Rel_ChkDocQrys(Optional varSkip As Variant) As Boolean
' ** Called by:
' **   QuikRelDoc(), Above

  Const THIS_PROC As String = "Rel_ChkDocQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strPath1 As String, strFile1 As String, strPathFile1 As String, strPath2 As String, strFile2 As String, strPathFile2 As String
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

  'blnSkip = True
  If blnSkip = False Then

    strPath1 = gstrDir_Dev
    strFile1 = CurrentAppName  ' ** Module Function: modFileUtilities.
    strPathFile1 = strPath1 & LNK_SEP & strFile1

    strPath2 = CurrentAppPath  ' ** Module Function: modFileUtilities.
    strFile2 = "TrstXAdm - Copy (6).mdb"
    strPathFile2 = strPath2 & LNK_SEP & strFile2

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

    Debug.Print "'REL DOC QRYS: " & CStr(lngQrys)
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
  'End If  ' ** blnSkip.

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
          DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile1, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
          If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
            DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile2, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
            If ERR.Number <> 0 Then
On Error GoTo 0
              Debug.Print "'QRY MISSING!  " & arr_varQry(Q_QNAM, lngX)
            Else
On Error GoTo 0
              arr_varQry(Q_IMP, lngX) = CBool(True)
            End If
          Else
On Error GoTo 0
          End If
          arr_varQry(Q_IMP, lngX) = CBool(True)
        End If
      Next
    Else
      Debug.Print "'ALL REL DOC QRYS PRESENT!"
    End If

    Debug.Print "'DONE!"
    DoEvents

    Beep

  End If  ' ** blnSkip.

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rel_ChkDocQrys = blnRetValx

End Function
