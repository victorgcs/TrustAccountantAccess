Attribute VB_Name = "zz_mod_QuerySQLDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_QuerySQLDocFuncs"

'VGC 02/05/2017: CHANGES!

' ** This applies to data collected via VBA_Find_Proc() in zz_mod_Module_Format_Funcs.

' ** Qry_RemExpr_rel() in modQueryFunctions.

Private Const QTE As String = """"  ' ** Chr(34), Double-Quote.
Private Const APO As String = "'"   ' ** Chr(39), Apostrophe, Single-Quote.
Private Const OCT As String = "#"   ' ** Chr(35), Octothorp, Pound-Sign, Hash-Mark.

Private blnJrnlTmp As Boolean
Private blnSQLDebug As Boolean, blnRetValx As Boolean

' ** Array: arr_varSQLTerm() [Functions: VBA_IsSQL(), VBA_IsSQLT_TERMs()].
Private Const SQLT_ELEMS As Integer = 1  ' ** Array's first-element UBound().
Private Const SQLT_TERM     As Integer = 0
Private Const SQLT_PRIORITY As Integer = 1

Private Const SQL_PRI1 As Integer = 1
Private Const SQL_PRI2 As Integer = 2
Private Const SQL_PRI3 As Integer = 3

Private Const SQL_NONE   As Integer = 1
Private Const SQL_HASCMD As Integer = 2
Private Const SQL_HASVRB As Integer = 4
Private Const SQL_HASTRM As Integer = 8
Private Const SQL_HASVCD As Integer = 16

Private blnModSqlVars As Boolean
' ** Array: arr_varModSQL().
Private Const MSQL_ELEMS As Integer = 4  ' ** Array's first-element UBound().
Private Const MSQL_RESP     As Integer = 0
Private Const MSQL_LINE_TXT As Integer = 1
Private Const MSQL_LINE_NUM As Integer = 2
Private Const MSQL_DOCD     As Integer = 3
Private Const MSQL_MOD_ELEM As Integer = 4
' **

Public Function QuikQryDoc() As Boolean
  Const THIS_PROC As String = "QuikQryDoc"
  Dim strRetVal As String
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Qry_ChkDocQrys(False) = True Then  ' ** Function: Below.
      blnJrnlTmp = False
      blnRetValx = Qry_Doc  ' ** Function: Below.
      blnRetValx = Qry_Parm_Doc  ' ** Function: Below.
Stop
      strRetVal = Qry_Tbl_Doc  ' ** Function: Below.
      'blnRetValx = Qry_FldDoc  ' ** Function: Below.
      blnRetValx = Qry_ChkDocQrys2  ' ** Function: Below.
      ' ** TO UNDO ABORTED DOC, RUN THIS IN THE IMMEDIATE WINDOW:
      ' **   Qry_TmpTables(False)  ' ** Function: Below.
      DoEvents
      DoBeeps  ' ** Module Function: modWindowFunctions.
      Debug.Print "'FINISHED!"
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED Qry_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikQryDoc = blnRetValx
End Function

Private Function Qry_Doc() As Boolean
' ** Called by:
' **   QuikQryDoc(), Above

  Const THIS_PROC As String = "Qry_Doc"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, prp As DAO.Property
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngQryxs As Long, arr_varQryx() As Variant
  Dim lngQryID As Long
  Dim blnAdd As Boolean, blnFound As Boolean
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim strQryName As String
  Dim lngTmp00 As Long, strTmp01 As String
  Dim intPos1 As Integer
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const Q_DID  As Integer = 0
  Const Q_DNAM As Integer = 1
  Const Q_QID  As Integer = 2
  Const Q_QNAM As Integer = 3

  blnRetValx = True

  DoCmd.Hourglass True
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  If gstrTrustDataLocation = vbNullString Then
    IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
  End If

If TableExists("_~xusr") = False Then
Stop
End If

  Debug.Print "'LINK DONE"  ' ** To reset the screen after linking.
  DoEvents

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  lngQryxs = 0&
  ReDim arr_varQryx(Q_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    ' ** Empty tblQuery_Staging1.
    Set qdf1 = .QueryDefs("zz_qry_Query_01a")
    qdf1.Execute

    For Each qdf1 In .QueryDefs
      lngQryxs = lngQryxs + 1&
      lngE = lngQryxs - 1&
      ReDim Preserve arr_varQryx(Q_ELEMS, lngE)
      arr_varQryx(Q_DID, lngE) = lngThisDbsID
      arr_varQryx(Q_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
      arr_varQryx(Q_QID, lngE) = lngQryxs  ' ** Really just the index number.
      arr_varQryx(Q_QNAM, lngE) = qdf1.Name
    Next

    ' ** Binary Sort arr_varQryx() array.
    For lngX = UBound(arr_varQryx, 2) To 1& Step -1&
      For lngY = 0& To (lngX - 1&)
        If arr_varQryx(Q_QNAM, lngY) > arr_varQryx(Q_QNAM, (lngY + 1)) Then
          strTmp01 = arr_varQryx(Q_QNAM, lngY)
          lngTmp00 = arr_varQryx(Q_QID, lngY)
          arr_varQryx(Q_QNAM, lngY) = arr_varQryx(Q_QNAM, (lngY + 1))
          arr_varQryx(Q_QID, lngY) = arr_varQryx(Q_QID, (lngY + 1))
          arr_varQryx(Q_QNAM, (lngY + 1)) = strTmp01
          arr_varQryx(Q_QID, (lngY + 1)) = lngTmp00
        End If
      Next
    Next

    Set rst1 = .OpenRecordset("tblQuery", dbOpenDynaset, dbConsistent)
    Set rst2 = .OpenRecordset("tblQuery_Staging1", dbOpenDynaset, dbAppendOnly)

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'QRYS: " & CStr(lngQryxs)
    Debug.Print "'|";
    DoEvents

    For lngX = 0& To (lngQryxs - 1&)
      Set qdf1 = .QueryDefs(arr_varQryx(Q_QNAM, lngX))
      With qdf1
        rst2.AddNew
        rst2![dbsx_id] = lngThisDbsID
        rst2![qryx_name] = .Name
        rst2![qryx_datemodified] = Now()
        rst2.Update
        If Left$(.Name, 1) <> "~" Then  ' ** Skip those pesky system queries.
          strQryName = .Name
          Select Case strQryName
          Case Else
            ' ** Nothing else yet.
          End Select
          lngQrys = lngQrys + 1&
          lngE = lngQrys - 1&
          ReDim Preserve arr_varQry(Q_ELEMS, lngE)
          arr_varQry(Q_DID, lngE) = lngThisDbsID
          arr_varQry(Q_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
          arr_varQry(Q_QNAM, lngE) = .Name
          With rst1
            blnAdd = False
            If .BOF = True And .EOF = True Then
              blnAdd = True
              .AddNew
              ![dbs_id] = arr_varQry(Q_DID, lngE)
            Else
              .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [qry_name] = '" & qdf1.Name & "'"
              If .NoMatch = False Then
                .Edit
              Else
                blnAdd = True
                .AddNew
                ![dbs_id] = arr_varQry(Q_DID, lngE)
              End If
            End If
            If blnAdd = True Then
              ![qry_name] = qdf1.Name
            End If
            ![qrytype_type] = qdf1.Type
            For Each prp In qdf1.Properties
              If prp.Name = "Description" Then
                If prp.Value <> vbNullString Then
                  ![qry_description] = prp.Value
                End If
                Exit For
              End If
            Next
            ![qry_sql] = qdf1.SQL
            If qdf1.Parameters.Count > 0 Then
              ![qry_param] = True
              If Left$(qdf1.SQL, 10) = "PARAMETERS" Then
                ![qry_param_clause] = True
              Else
                ![qry_param_clause] = False
              End If
              ![qry_paramcnt] = qdf1.Parameters.Count
            Else
              ![qry_param] = False
              ![qry_paramcnt] = 0&
              ![qry_formrefcnt] = 0&
            End If
            If InStr(qdf1.SQL, "[Forms]") > 0 Then
              ![qry_formref] = True
              intPos1 = InStr(qdf1.SQL, "[Forms]")
              Do While intPos1 > 0
                ![qry_formrefcnt] = ![qry_formrefcnt] + 1&
                intPos1 = InStr((intPos1 + 1), qdf1.SQL, "[Forms]")
              Loop
            ElseIf InStr(qdf1.SQL, "[Reports]") > 0 Then
              ![qry_formref] = True
              intPos1 = InStr(qdf1.SQL, "[Reports]")
              Do While intPos1 > 0
                ![qry_formrefcnt] = ![qry_formrefcnt] + 1&
                intPos1 = InStr((intPos1 + 1), qdf1.SQL, "[Reports]")
              Loop
            Else
              ![qry_formref] = False
              ![qry_formrefcnt] = 0&
            End If
            ![qry_tblcnt] = 0&
            ![qry_fldcnt] = qdf1.Fields.Count
            ![qry_datemodified] = Now()
            If blnAdd = False Then
              lngQryID = ![qry_id]
            End If
            .Update
            If blnAdd = True Then
              .Bookmark = .LastModified
              lngQryID = ![qry_id]
            End If
            arr_varQry(Q_QID, lngE) = lngQryID
          End With
        End If
      End With  ' ** qdf1.
      If ((lngX + 1&) Mod 100&) = 0 And lngX > 0& Then
        Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngQryxs)
        Debug.Print "'|";
        DoEvents
      ElseIf ((lngX + 1&) Mod 10&) = 0 And lngX > 0& Then
        Debug.Print "|";
        DoEvents
      Else
        Debug.Print ".";
        DoEvents
      End If
    Next  ' ** lngX.
    Debug.Print "  " & CStr(lngQryxs) & "  (PASS 1)"
    DoEvents

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    If lngQrys > 0& Then
      lngDels = 0&
      ReDim arr_varDel(0)
      With rst1
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        Debug.Print "|";
        DoEvents
        For lngY = 1& To lngRecs
          If ![dbs_id] = lngThisDbsID Then
            blnFound = False
            For lngZ = 0& To (lngQrys - 1&)
              If arr_varQry(Q_QNAM, lngZ) = ![qry_name] Then
                blnFound = True
                Exit For
              End If
            Next
            If blnFound = False Then
              lngDels = lngDels + 1&
              lngE = lngDels - 1&
              ReDim Preserve arr_varDel(lngE)
              arr_varDel(lngE) = ![qry_id]
            End If
          End If
          If ((lngY Mod 100) = 0) Then
            Debug.Print "|  " & CStr(lngY) & " of " & CStr(lngRecs)
            Debug.Print "|";
            DoEvents
          ElseIf ((lngY Mod 10) = 0) Then
            Debug.Print "|";
            DoEvents
          Else
            Debug.Print ".";
            DoEvents
          End If
          If lngY < lngRecs Then .MoveNext
        Next
      End With
    End If
    Debug.Print "  " & CStr(lngRecs) & "  (PASS 2)"
    DoEvents

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    If lngDels > 0& Then
      Debug.Print "'DELS: " & CStr(lngDels)
      DoEvents
      For lngY = 0& To (lngDels - 1&)
        ' ** Delete tblQuery, by specified [qid].
        Set qdf1 = .QueryDefs("zz_qry_Query_01b")
        With qdf1.Parameters
          ![qid] = arr_varDel(lngY)
        End With
        qdf1.Execute
      Next
    End If

    rst1.Close

    ' ** Update tblQuery, with  qry_formref = False for zz_qry's, by specified CurrentAppName().
    Set qdf1 = .QueryDefs("zz_qry_Query_06")
    qdf1.Execute

    rst2.Close
    .Close
  End With  ' ** dbs.

If TableExists("_~xusr") = False Then
Stop
End If

  DoCmd.Hourglass False
  DoEvents

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'QRYS: " & CStr(lngQrys)

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set prp = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf1 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  Qry_Doc = blnRetValx

End Function

Private Function Qry_Parm_Doc() As Boolean
' ** Called by:
' **   QuikQryDoc(), Above

  Const THIS_PROC As String = "Qry_Parm_Doc"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim prm As DAO.Parameter
  Dim lngQryID As Long, lngQryParmID As Long
  Dim lngParms As Long, arr_varParm() As Variant
  Dim strParms As String, lngParmCnt As Long, strThisParm As String, lngThisParmCnt As Long, strSQL As String
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim strQryName As String
  Dim lngDels As Long, arr_varDel() As Variant
  Dim blnAdd As Boolean, blnHasParmClause As Boolean, blnInClause As Boolean, blnFound As Boolean
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp01 As String, blnTmp02 As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varParm().
  Const P_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const P_DID  As Integer = 0
  Const P_QID  As Integer = 1
  Const P_PID  As Integer = 2
  Const P_PNAM As Integer = 3
  Const P_ORD  As Integer = 4

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  If gstrTrustDataLocation = vbNullString Then
    IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
  End If

  Debug.Print "'LINK DONE"  ' ** To reset the screen after linking.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    lngParms = 0&
    ReDim arr_varParm(P_ELEMS, 0)

    Set rst2 = .OpenRecordset("tblQuery_Parameter", dbOpenDynaset, dbConsistent)

    ' ** tblQuery, just parameter queries, by specified CurrentAppName().
    Set qdf1 = .QueryDefs("zz_qry_Query_Parameter_02")
    Set rst1 = qdf1.OpenRecordset
    With rst1
      If .BOF = True And .EOF = True Then
        ' ** How silly!
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        Debug.Print "'QRYS: " & CStr(lngRecs)
        Debug.Print "'|";
        DoEvents
        For lngX = 1& To lngRecs
          If ![dbs_id] <> lngThisDbsID Then  ' ** I should already be qualified.
            strQryName = ![qry_name]
            Select Case strQryName
            Case Else
              ' ** Nothing else yet.
            End Select
            lngQryID = ![qry_id]
            strSQL = ![qry_sql]  ' ** Whole query.
            If ![param_spec] = True Then
              ' ** PARAMETERS are specified.
              blnHasParmClause = True
              If ![qry_param_clause] = False Then
                .Edit
                ![qry_param_clause] = True
                ![qry_datemodified] = Now()
                .Update
              End If
              intPos1 = InStr(strSQL, ";")
              strParms = Left$(strSQL, intPos1)  ' ** Just PARAMETERS clause.
            Else
              ' ** PARAMETERS are not specified.
              blnHasParmClause = False
              strParms = vbNullString
            End If
            lngY = 0&
            Set qdf2 = dbs.QueryDefs(![qry_name])
            With qdf2
              lngParmCnt = DCount("[qryparam_id]", "tblQuery_Parameter", "[qry_id] = " & CStr(lngQryID))
              For Each prm In .Parameters
                blnAdd = False: blnInClause = False
                lngY = lngY + 1&
                With prm
                  strThisParm = .Name
                  lngThisParmCnt = DCount("[qryparam_id]", "tblQuery_Parameter", "[qry_id] = " & CStr(lngQryID) & " And " & _
                    "[qryparam_name] = '" & strThisParm & "'")
                  If lngThisParmCnt = 0 Then
                    ' ** Parameters haven't yet been documented.
                    blnAdd = True
                  Else
                    ' ** There are already parameters for this query in the table.
                    rst2.FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [qry_id] = " & CStr(lngQryID) & " And " & _
                      "[qryparam_order] = " & CStr(lngY)
                    If rst2.NoMatch = True Then
                      blnAdd = True
                    End If
                  End If
                  With rst2
                    If blnAdd = True Then
                      lngParms = lngParms + 1&
                      lngE = lngParms - 1&
                      ReDim Preserve arr_varParm(P_ELEMS, lngE)
                      arr_varParm(P_DID, lngE) = lngThisDbsID
                      arr_varParm(P_QID, lngE) = lngQryID
                      arr_varParm(P_PID, lngE) = CLng(0)
                      arr_varParm(P_PNAM, lngE) = prm.Name
                      arr_varParm(P_ORD, lngE) = lngY
                      .AddNew
                      ![dbs_id] = lngThisDbsID
                      ![qry_id] = lngQryID
                      ![qryparam_name] = prm.Name
                      ![pdatatype_type] = prm.Type
                      ![parmdir_type] = prm.Direction  ' ** Not really applicable to non-ODBC sources.
                      If blnHasParmClause = True Then
                        intPos1 = InStr(strParms, prm.Name)
                        If intPos1 > 0 Then
                          blnInClause = True
                          intPos2 = InStr(intPos1, strParms, ",")
                          strTmp01 = vbNullString
                          If intPos2 = 0 Then
                            ' ** It's the last or only parameter.
                            If intPos1 = 1 Then
                              strTmp01 = strParms
                            Else
                              strTmp01 = Mid$(strParms, intPos1)
                            End If
                            If Right$(strTmp01, 1) = ";" Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                            If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                          Else
                            If intPos1 = 1 Then
                              strTmp01 = Left$(strParms, (intPos2 - 1))
                            Else
                              strTmp01 = Mid$(strParms, intPos1, ((intPos2 - intPos1) + 1))
                            End If
                            ![qryparam_sql] = strTmp01
                          End If
                        Else
                          ' ** A parameter not specified; shouldn't happen unless
                          ' ** there's a problem with the underlying table.
                          intPos1 = InStr(strSQL, prm.Name)
                          If intPos1 > 0 Then
                            strTmp01 = Mid$(strSQL, intPos1, Len(prm.Name))
                            ![qryparam_sql] = strTmp01
                          Else
                            strTmp01 = "{where is it?}"
                          End If
                          ![qryparam_sql] = strTmp01
                        End If
                      Else
                        intPos1 = InStr(strSQL, prm.Name)
                        If intPos1 > 0 Then
                          strTmp01 = Mid$(strSQL, intPos1, Len(prm.Name))
                        Else
                          strTmp01 = "{where is it?}"
                        End If
                        ![qryparam_sql] = strTmp01
                      End If
                      ![qryparam_order] = lngY
                      ![qryparam_clause] = blnInClause
                      ![qryparam_datemodified] = Now()
                      .Update
                      .Bookmark = .LastModified
                      arr_varParm(P_PID, lngE) = ![qryparam_id]
                    Else
                      lngParms = lngParms + 1&
                      lngE = lngParms - 1&
                      ReDim Preserve arr_varParm(P_ELEMS, lngE)
                      arr_varParm(P_DID, lngE) = ![dbs_id]
                      arr_varParm(P_QID, lngE) = lngQryID
                      arr_varParm(P_PID, lngE) = ![qryparam_id]
                      arr_varParm(P_PNAM, lngE) = prm.Name
                      arr_varParm(P_ORD, lngE) = lngY
                      ' ** Already in the table.
                      blnTmp02 = False
                      If ![qryparam_name] <> prm.Name Then
                        .Edit
                        ![qryparam_name] = prm.Name
                        ![qryparam_datemodified] = Now()
                        .Update
                      End If
                      If ![pdatatype_type] <> prm.Type Then
                        .Edit
                        ![pdatatype_type] = prm.Type
                        ![qryparam_datemodified] = Now()
                        .Update
                      End If
                      If ![parmdir_type] <> prm.Direction Then
                        .Edit
                        ![parmdir_type] = prm.Direction
                        ![qryparam_datemodified] = Now()
                        .Update
                      End If
                      If blnHasParmClause = True Then
                        intPos1 = InStr(strParms, prm.Name)
                        If intPos1 > 0 Then
                          blnInClause = True
                          intPos2 = InStr(intPos1, strParms, ",")
                          strTmp01 = vbNullString
                          If intPos2 = 0 Then
                            ' ** It's the last or only parameter.
                            If intPos1 = 1 Then
                              strTmp01 = strParms
                            Else
                              strTmp01 = Mid$(strParms, intPos1)
                            End If
                            If Right$(strTmp01, 1) = ";" Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                            If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                          Else
                            If intPos1 = 1 Then
                              strTmp01 = Left$(strParms, (intPos2 - 1))
                            Else
                              strTmp01 = Mid$(strParms, intPos1, ((intPos2 - intPos1) + 1))
                            End If
                            If Right$(strTmp01, 1) = ";" Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                            If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                          End If
                          If IsNull(![qryparam_sql]) = True Then
                            .Edit
                            ![qryparam_sql] = strTmp01
                            ![qryparam_datemodified] = Now()
                            .Update
                            blnTmp02 = True
                          Else
                            If ![qryparam_sql] <> strTmp01 Then
                              .Edit
                              ![qryparam_sql] = strTmp01
                              ![qryparam_datemodified] = Now()
                              .Update
                              blnTmp02 = True
                            End If
                          End If
                        Else
                          ' ** A parameter not specified; shouldn't happen unless
                          ' ** there's a problem with the underlying table.
                          intPos1 = InStr(strSQL, prm.Name)
                          If intPos1 > 0 Then
                            strTmp01 = Mid$(strSQL, intPos1, Len(prm.Name))
                          Else
                            strTmp01 = "{where is it?}"
                          End If
                          If IsNull(![qryparam_sql]) = True Then
                            .Edit
                            ![qryparam_sql] = strTmp01
                            ![qryparam_datemodified] = Now()
                            .Update
                            blnTmp02 = True
                          Else
                            If ![qryparam_sql] <> strTmp01 Then
                              .Edit
                              ![qryparam_sql] = strTmp01
                              ![qryparam_datemodified] = Now()
                              .Update
                              blnTmp02 = True
                            End If
                          End If
                        End If
                      Else
                        intPos1 = InStr(strSQL, prm.Name)
                        If intPos1 > 0 Then
                          strTmp01 = Mid$(strSQL, intPos1, Len(prm.Name))
                        Else
                          strTmp01 = "{where is it?}"
                        End If
                        If IsNull(![qryparam_sql]) = True Then
                          .Edit
                          ![qryparam_sql] = strTmp01
                          ![qryparam_datemodified] = Now()
                          .Update
                          blnTmp02 = True
                        Else
                          If ![qryparam_sql] <> strTmp01 Then
                            .Edit
                            ![qryparam_sql] = strTmp01
                            ![qryparam_datemodified] = Now()
                            .Update
                            blnTmp02 = True
                          End If
                        End If
                      End If
                    End If  ' ** blnAdd.
                    If blnAdd = True Then
                      .Bookmark = .LastModified
                    End If
                    If ![qryparam_clause] <> blnInClause Then
                      .Edit
                      ![qryparam_clause] = blnInClause
                      ![qryparam_datemodified] = Now()
                      .Update
                    End If
                  End With  ' ** rst2.
                End With  ' ** This parameter: prm, lngY.
              Next  ' ** For each parameter: prm.
            End With
            If ![qry_paramcnt] <> lngY Then
              .Edit
              ![qry_paramcnt] = lngY
              ![qry_datemodified] = Now()
              .Update
            End If
            Set qdf2 = Nothing
            If lngParmCnt > 0& Then
              ' ** Parameters had already been added to the table.
              If lngParmCnt > lngY Then
                ' ** Delete tblQuery_Parameter, by specified [qid], [pord].
                Set qdf2 = dbs.QueryDefs("zz_qry_Query_Parameter_01a")
                With qdf2.Parameters
                  ![qid] = lngQryID
                  ![pord] = lngY
                End With
                qdf2.Execute
              End If
            End If
          End If
          If ((lngX Mod 100) = 0) Then
            Debug.Print "|  " & CStr(lngX) & " of " & CStr(lngRecs)
            Debug.Print "'|";
            DoEvents
          ElseIf ((lngX Mod 10) = 0) Then
            Debug.Print "|";
            DoEvents
          Else
            Debug.Print ".";
            DoEvents
          End If
          If lngX < lngRecs Then .MoveNext
        Next
        Debug.Print "  " & CStr(lngRecs) & "  (PASS 1)"
        DoEvents
        'lngQryParmID = ![qryparam_id]

      End If
      .Close
    End With  ' ** zz_qry_Query_Parameter_02: rst1.

    rst2.Close  ' ** tblQuery_Parameter.

    If lngParms > 0& Then

      lngDels = 0&
      ReDim arr_varDel(0)

      Set rst2 = .OpenRecordset("tblQuery_Parameter", dbOpenDynaset, dbConsistent)
      With rst2
        If .BOF = True And .EOF = True Then
          ' ** Aww, skip it!
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          Debug.Print "'QRYS: " & CStr(lngRecs)
          Debug.Print "|";
          DoEvents
          For lngX = 1& To lngRecs
            If ![dbs_id] = lngThisDbsID Then
              blnFound = False
              For lngY = 0& To (lngParms - 1&)
                If arr_varParm(P_PID, lngY) = ![qryparam_id] Then
                  blnFound = True
                  Exit For
                End If
              Next
              If blnFound = False Then
                lngDels = lngDels + 1&
                lngE = lngDels - 1&
                ReDim Preserve arr_varDel(lngE)
                arr_varDel(lngE) = ![qryparam_id]
              End If
            End If
            If ((lngX Mod 100) = 0) Then
              Debug.Print "|  " & CStr(lngX) & " of " & CStr(lngRecs)
              Debug.Print "'|";
              DoEvents
            ElseIf ((lngX Mod 10) = 0) Then
              Debug.Print "|";
              DoEvents
            Else
              Debug.Print ".";
              DoEvents
            End If
            If lngX < lngRecs Then .MoveNext
          Next
        End If
      End With
      Debug.Print "  " & CStr(lngRecs) & "  (PASS 2)"
      DoEvents

      If lngDels > 0& Then
        Debug.Print "'DELS: " & CStr(lngDels)
        DoEvents
        For lngX = 0& To (lngDels - 1&)
          ' ** Delete tblQuery_Parameter, by specified [qpid].
          Set qdf1 = .QueryDefs("zz_qry_Query_Parameter_01b")
          With qdf1.Parameters
            ![qpid] = arr_varDel(lngX)
          End With
          qdf1.Execute
        Next
      End If

    End If


    .Close
  End With  ' ** dbs.

  ' ** zz_qry_Query_02:
  ' ** PARAMETERS [parm01] Bit, [parm02] Byte, [parm03] Short, [parm04] Long, [parm05] Currency,
  ' **   [parm06] IEEESingle, [parm07] IEEEDouble, [parm08] DateTime, [parm09] Binary,
  ' **   [parm10] Text ( 255 ), [parm11] LongBinary, [parm12] Text, [parm13] Guid, [parm14] Value;
  ' ** NOTE: Though 'Value' is on the dropdown list, it returns a 10, dbText.
  ' **       I could find no mention of this choice in any of the documentation.

  ' ** Direction property:
  ' **   Sets or returns a value that indicates whether a Parameter object represents an input parameter,
  ' **   an output parameter, both, or the return value from the procedure (ODBCDirect workspaces only).

  ' ** DbParamaterDirection enumeration:
  ' **   1 dbParamInput        Passes information to the procedure. (Default)
  ' **   2 dbParamOutput       Returns information from the procedure as in an output parameter in SQL.
  ' **   3 dbParamInputOutput  Passes information both to and from the procedure.
  ' **   4 dbParamReturnValue  Passes the return value from a procedure.

  ' *******************************************************************************************************************
  ' ** Equivalent ANSI SQL Data Types:
  ' ** The following table lists ANSI SQL data types, their equivalent Microsoft Jet database engine SQL
  ' ** data types, and their valid synonyms. It also lists the equivalent Microsoft® SQL Server™ data types.
  ' **
  ' **   ANSI SQL                    Microsoft Jet        Synonym                      Microsoft SQL
  ' **   data type                   SQL data type                                     Server data type
  ' **   ==========================  ===================  ===========================  ==============================
  ' **   BIT, BIT VARYING            BINARY (See Notes)   VARBINARY, BINARY VARYING    BINARY, VARBINARY
  ' **                                                    BIT VARYING
  ' **   Not supported               BIT (See Notes)      BOOLEAN, LOGICAL,            BIT
  ' **                                                    LOGICAL1, YESNO
  ' **   Not supported               TINYINT              INTEGER1, BYTE               TINYINT
  ' **   Not supported               COUNTER (See Notes)  AUTOINCREMENT                (See Notes)
  ' **   Not supported               MONEY                CURRENCY                     MONEY
  ' **   DATE, TIME, TIMESTAMP       DATETIME             DATE, TIME (See Notes)       DATETIME
  ' **   Not supported               UNIQUEIDENTIFIER     GUID                         UNIQUEIDENTIFIER
  ' **   DECIMAL                     DECIMAL              NUMERIC, DEC                 DECIMAL
  ' **   REAL                        REAL                 SINGLE, FLOAT4, IEEESINGLE   REAL
  ' **   DOUBLE PRECISION,           FLOAT                DOUBLE, FLOAT8,              FLOAT
  ' **   FLOAT                                            IEEEDOUBLE, NUMBER
  ' **                                                    (See Notes)
  ' **   SMALLINT                    SMALLINT             SHORT, INTEGER2              SMALLINT
  ' **   INTEGER                     INTEGER              LONG, INT, INTEGER4          INTEGER
  ' **   INTERVAL                    Not supported                                     Not supported
  ' **   Not supported               IMAGE                LONGBINARY, GENERAL,         IMAGE
  ' **                                                    OLEOBJECT
  ' **   Not supported               TEXT (See Notes)     LONGTEXT, LONGCHAR, MEM0,    TEXT
  ' **                                                    NOTE, NTEXT (See Notes)
  ' **   CHARACTER,                  CHAR (See Notes)     TEXT(n), ALPHANUMERIC,       CHAR, VARCHAR, NCHAR, NVARCHAR
  ' **   CHARACTER VARYING,                               CHARACTER, STRING, VARCHAR,
  ' **   NATIONAL CHARACTER,                              CHARACTER VARYING, NCHAR,
  ' **   NATIONAL CHARACTER VARYING                       NATIONAL CHARACTER,
  ' **                                                    NATIONAL CHAR,
  ' **                                                    NATIONAL CHARACTER VARYING,
  ' **                                                    NATIONAL CHAR VARYING
  ' **                                                     (See Notes)
  ' *******************************************************************************************************************

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set prm = Nothing
  Set rst2 = Nothing
  Set rst1 = Nothing
  Set qdf1 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  Qry_Parm_Doc = blnRetValx

End Function

Private Function Qry_Tbl_Doc() As String
' ** This table documentation is based on the actual SQL code.
' ** With much Sturm und Drang, the parsing seems to be accurate.
' ** Called by:
' **   QuikQryDoc(), Above

  Const THIS_PROC As String = "Qry_Tbl_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim strQryName As String, lngQryID As Long
  Dim strSQL As String
  Dim blnAdd As Boolean, blnByOrd As Boolean, blnByID As Boolean
Dim blnSkip As Boolean
  Dim lngSources As Long, arr_varSource() As Variant
  Dim lngQDDLs As Long
  Dim lngThisDbsID As Long, lngNotThisDbsID1 As Long, lngNotThisDbsID2 As Long, lngRecs As Long
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer, intPos4 As Integer, intLen As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String, strTmp06 As String
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim strRetVal As String

  ' ** Array: arr_varSource().
  Const S_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const S_DID As Integer = 0
  Const S_NAM As Integer = 1
  Const S_QID As Integer = 2
  Const S_TID As Integer = 3
  Const S_ORD As Integer = 4

  strRetVal = vbNullString

'WHAT IS THIS? 2  'MSysAccessStorage'
'WHAT IS THIS? 2  'MSysAccessXML'
'WHAT IS THIS? 2  'MSysACEs'
'WHAT IS THIS? 2  'MSysObjects'
'WHAT IS THIS? 2  'MSysQueries'
'WHAT IS THIS? 2  'MSysRelationships'

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
  varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TrstXAdm.mdb'")  ' ** Module Function: modFileUtilities.
  Select Case IsNull(varTmp00)
  Case True
    varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TrstXAdm.mde'")  ' ** Module Function: modFileUtilities.
    lngNotThisDbsID1 = varTmp00
  Case False
    lngNotThisDbsID1 = varTmp00
  End Select
  varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TrustImport.mdb'")  ' ** Module Function: modFileUtilities.
  Select Case IsNull(varTmp00)
  Case True
    varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = 'TrustImport.mde'")  ' ** Module Function: modFileUtilities.
    lngNotThisDbsID2 = varTmp00
  Case False
    lngNotThisDbsID2 = varTmp00
  End Select

  If gstrTrustDataLocation = vbNullString Then
    IniFile_GetDataLoc  ' ** Module Function: modStartupFuncs.
  End If

  If TableExists("tmpUpdatedValues") = False Then  ' ** Module Function: modFileUtilities.
    DoCmd.CopyObject , "tmpUpdatedValues", acTable, "zz_tmpUpdatedValues"
  End If
  If TableExists("zz_tbl_RePost_Posting") = False Then  ' ** Module Function: modFileUtilities.
    'DoCmd.CopyObject , "zz_tbl_RePost_Posting", acTable, "zz_tbl_RePost_Posting_Bak"
  End If
  If TableExists("LedgerArchive_Backup") = False Then  ' ** Module Function: modFileUtilities.
    DoCmd.CopyObject , "LedgerArchive_Backup", acTable, "tblTemplate_LedgerArchive"
  End If

  Application.SetOption "Show System Objects", True  ' ** Do show system objects.

'Add DLookups() ?
  Set dbs = CurrentDb
  With dbs

    ' ** dbQueryDefType enumeration:
    ' **     0  dbQSelect          Select
    ' **    16  dbQCrosstab        Crosstab
    ' **    32  dbQDelete          Delete
    ' **    48  dbQUpdate          Update
    ' **    64  dbQAppend          Append
    ' **    80  dbQMakeTable       Make-Table
    ' **    96  dbQDDL             Data-definition
    ' **   112  dbQSQLPassThrough  Pass-through (Microsoft Jet workspaces only)
    ' **   128  dbQSetOperation    Union
    ' **   144  dbQSPTBulk         Used with dbQSQLPassThrough to specify a query that doesn't return records.
    ' **                           (Microsoft Jet workspaces only)
    ' **   160  dbQCompound        Compound
    ' **   224  dbQProcedure       Procedure (ODBCDirect workspaces only)
    ' **   240  dbQAction          Action

    ' ** Current Trust Accountant census (07/13/2009):
    ' ** qrytype_type      cnt
    ' ** ===============  =====
    ' ** dbQSelect        2125
    ' ** dbQDelete         189
    ' ** dbQUpdate         206
    ' ** dbQAppend         206
    ' ** dbQMakeTable       20
    ' ** dbQSetOperation    97

    ' ** Empty tblQuery_RecordSource, by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_Query_RecordSource_01")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    qdf.Execute
    Set qdf = Nothing

    ' ** Empty tblQuery_SourceChain, by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_Query_SourceChain_01")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    qdf.Execute
    Set qdf = Nothing

    .Close
  End With  ' ** dbs.

  Debug.Print "'RESETTING AUTONUMBER!"
  DoEvents

  ' ** Reset the Autonumber fields to 1.
  ChangeSeed_Ext "tblQuery_RecordSource"  ' ** Module Function: modAutonumberFieldFuncs.
  ChangeSeed_Ext "tblQuery_SourceChain"  ' ** Module Function: modAutonumberFieldFuncs.

  Set dbs = CurrentDb
  With dbs

blnSkip = False
If blnSkip = False Then
    Set rst2 = .OpenRecordset("tblQuery_RecordSource", dbOpenDynaset, dbConsistent)

    Set rst1 = .OpenRecordset("tblQuery", dbOpenDynaset, dbReadOnly)
    With rst1
      If .BOF = True And .EOF = True Then
        ' ** Get outta here!
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst

        Debug.Print "'QRYS: " & CStr(lngRecs)
        Debug.Print "'|";
        DoEvents
        For lngX = 1& To lngRecs
          If ![dbs_id] = lngThisDbsID Then

            strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString: strTmp05 = vbNullString
            intPos1 = 0: intPos2 = 0: intPos3 = 0: intPos4 = 0
            strRetVal = vbNullString
            strQryName = ![qry_name]
            lngQryID = ![qry_id]
            strSQL = ![qry_sql]

            Select Case ![qrytype_type]
            Case dbQSelect  ' ** Select.
              strTmp02 = Trim$(Rem_CRLF(strSQL))  ' ** Module Function: modStringFuncs.
              intPos1 = InStr(strTmp02, " FROM ")  'THIS COULD BE A ' from ' WITHIN A TEXT FIELD ASSIGNMENT!
              If intPos1 > 0 Then
                If Compare_StringA_StringB(" FROM ", "=", Mid$(strTmp02, intPos1, 6)) = False Then
                  intPos1 = InStr((intPos1 + 4), strTmp02, " FROM ")
                  If intPos1 > 0 Then
                    If Compare_StringA_StringB(" FROM ", "=", Mid$(strTmp02, intPos1, 6)) = False Then
                      intPos1 = InStr((intPos1 + 4), strTmp02, " FROM ")
                    End If
                  End If
                End If
              End If

              If intPos1 > 0 Then
                strTmp03 = Trim$(Mid$(strTmp02, (intPos1 + 5)))
                intPos2 = InStr(strTmp03, " FROM ")
                If intPos2 > 0 Then
                  If Compare_StringA_StringB(" FROM ", "=", Mid$(strTmp03, intPos2, 6)) = False Then
                    intPos2 = InStr((intPos2 + 4), strTmp03, " FROM ")
                    If intPos2 > 0 Then
                      If Compare_StringA_StringB(" FROM ", "=", Mid$(strTmp03, intPos2, 6)) = False Then
                        intPos2 = InStr((intPos2 + 4), strTmp03, " FROM ")
                      End If
                    End If
                  End If
                End If
                If intPos2 > 0 Then
                  ' ** Includes a subquery. ONLY ONE?
                  strTmp06 = Trim$(Mid$(strTmp03, (intPos2 + 5)))
                End If
                ' ** First, finish off the standard part of the SQL.
                intPos2 = InStr(strTmp03, " WHERE ")
                If intPos2 > 0 Then
                  strTmp03 = Trim$(Left$(strTmp03, intPos2))
                End If
                intPos2 = InStr(strTmp03, " ORDER BY ")
                If intPos2 > 0 Then
                  strTmp03 = Trim$(Left$(strTmp03, intPos2))
                End If
                intPos2 = InStr(strTmp03, " GROUP BY ")
                If intPos2 > 0 Then
                  strTmp03 = Trim$(Left$(strTmp03, intPos2))
                End If
                intPos2 = InStr(strTmp03, " HAVING ")
                If intPos2 > 0 Then
                  strTmp03 = Trim$(Left$(strTmp03, intPos2))
                End If
                intPos2 = InStr(strTmp03, "JOIN")
                If intPos2 > 0 Then
                  ' ** Query JOIN verbs: 'INNER JOIN', 'LEFT JOIN', 'RIGHT JOIN'
                  ' ** Table locations: left of 1st JOIN verb, right of JOIN verb thereafter.
                  ' ** strTmp03 is everything to the right of FROM, without any criteria or sort.
                  strTmp04 = Trim$(Left$(strTmp03, (intPos2 - 1)))  ' ** 1st table.
                  If Right$(strTmp04, 6) = " INNER" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                  If Right$(strTmp04, 6) = " RIGHT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                  If Right$(strTmp04, 5) = " LEFT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 4)))
                  If Left$(strTmp04, 1) = "(" Then strTmp04 = Mid$(strTmp04, 2)
                  If InStr(strTmp04, ",") > 0 Then
                    ' ** May be Cartesian.
                    intPos3 = InStr(strTmp04, ",")
                    strTmp05 = Trim$(Left$(strTmp04, (intPos3 - 1)))
                    strTmp04 = Trim$(Mid$(strTmp04, (intPos3 + 1)))
                    intPos3 = InStr(strTmp05, " AS ")
                    If intPos3 > 0 Then
                      strTmp05 = Trim$(Left$(strTmp05, intPos3))
                    End If
                    If Left$(strTmp05, 1) = "[" And Right$(strTmp05, 1) = "]" Then
                      strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                      strTmp05 = Mid$(strTmp05, 2)
                    End If
                    intPos3 = InStr(strTmp05, "DISTINCTROW")
                    If intPos3 > 0 Then
                      If intPos3 = 1 Then
                        strTmp05 = Trim$(Mid$(strTmp05, 12))
                      Else
                        strTmp05 = Trim$(Left$(strTmp05, (intPos3 - 1))) & Trim$(Mid$(strTmp05, (intPos3 + 12)))
                      End If
                    End If
                    If Trim$(strTmp01) = vbNullString Then
                      strTmp01 = strTmp05 & "^"
                    Else
                      If Right$(strTmp01, 1) = "^" Then
                        strTmp01 = strTmp01 & strTmp05 & "^"
                      Else
                        strTmp01 = strTmp01 & "^" & strTmp05 & "^"
                      End If
                    End If
                    strTmp05 = vbNullString
                  End If
                  Do While Left$(strTmp04, 1) = "("
                    strTmp04 = Mid$(strTmp04, 2)
                  Loop
                  intPos3 = InStr(strTmp04, " AS ")
                  If intPos3 > 0 Then
                    strTmp04 = Trim$(Left$(strTmp04, intPos3))
                  End If
                  If Left$(strTmp04, 1) = "[" And Right$(strTmp04, 1) = "]" Then
                    strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                    strTmp04 = Mid$(strTmp04, 2)
                  End If
                  intPos3 = InStr(strTmp04, "DISTINCTROW")
                  If intPos3 > 0 Then
                    If intPos3 = 1 Then
                      strTmp04 = Trim$(Mid$(strTmp04, 12))
                    Else
                      strTmp04 = Trim$(Left$(strTmp04, (intPos3 - 1))) & Trim$(Mid$(strTmp04, (intPos3 + 12)))
                    End If
                  End If
                  If Trim$(strTmp01) = vbNullString Then
                    strTmp01 = Trim$(strTmp04) & "^"  ' ** strTmp01 now has 1st table in it.
                  Else
                    If Right$(strTmp01, 1) = "^" Then
                      strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp04)) & "^"
                    Else
                      strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp04)) & "^"
                    End If
                  End If
                  ' ** For the rest of the JOIN's, tables are on the right.
                  strTmp05 = Trim$(Mid$(strTmp03, (intPos2 + 4)))   ' ** Everything right of 1st JOIN.
                  Do While Left$(strTmp05, 1) = "("
                    strTmp05 = Mid$(strTmp05, 2)
                  Loop
                  intPos3 = InStr(strTmp05, " JOIN ")
                  If intPos3 > 0 Then
                    Do While intPos3 > 0
                      strTmp04 = Trim$(Left$(strTmp05, intPos3))  ' ** Table will be at the left, before any ON's.
                      Do While Left$(strTmp04, 1) = "("
                        strTmp04 = Mid$(strTmp04, 2)
                      Loop
                      intPos4 = InStr(strTmp04, " ON ")
                      If intPos4 > 0 Then
                        strTmp04 = Trim$(Left$(strTmp04, intPos4))
                      End If
                      intPos4 = InStr(strTmp04, "DISTINCTROW")
                      If intPos4 > 0 Then
                        If intPos4 = 1 Then
                          strTmp04 = Trim$(Mid$(strTmp04, 12))
                        Else
                          strTmp04 = Trim$(Left$(strTmp04, (intPos4 - 1))) & Trim$(Mid$(strTmp04, (intPos4 + 12)))
                        End If
                      End If
                      Do While Right$(strTmp04, 1) = ")"
                        strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                      Loop
                      If Right$(strTmp04, 6) = " INNER" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                      If Right$(strTmp04, 6) = " RIGHT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                      If Right$(strTmp04, 5) = " LEFT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 4)))
                      If Right$(strTmp04, 1) = ";" Then strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                      intPos4 = InStr(strTmp04, " AS ")
                      If intPos4 > 0 Then
                        strTmp04 = Trim$(Left$(strTmp04, intPos4))
                      End If
                      Do While Left$(strTmp04, 1) = "("
                        strTmp04 = Mid$(strTmp04, 2)
                      Loop
                      If Left$(strTmp04, 1) = "[" And Right$(strTmp04, 1) = "]" Then
                        strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                        strTmp04 = Mid$(strTmp04, 2)
                      End If
                      intPos4 = InStr(strTmp04, "DISTINCTROW")
                      If intPos4 > 0 Then
                        If intPos4 = 1 Then
                          strTmp04 = Trim$(Mid$(strTmp04, 12))
                        Else
                          strTmp04 = Trim$(Left$(strTmp04, (intPos4 - 1))) & Trim$(Mid$(strTmp04, (intPos4 + 12)))
                        End If
                      End If
                      If Trim$(strTmp01) = vbNullString Then
                        strTmp01 = Trim$(strTmp04) & "^"
                      Else
                        If Right$(strTmp01, 1) = "^" Then
                          strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp04)) & "^"
                        Else
                          strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp04)) & "^"
                        End If
                      End If
                      strTmp05 = Trim$(Mid$(strTmp05, (intPos3 + 5)))
                      intPos3 = InStr(strTmp05, " JOIN ")
                      If intPos3 = 0 Then
                        intPos4 = InStr(strTmp05, " ON ")
                        If intPos4 > 0 Then
                          strTmp05 = Trim$(Left$(strTmp05, intPos4))
                        End If
                        intPos4 = InStr(strTmp05, "DISTINCTROW")
                        If intPos4 > 0 Then
                          If intPos4 = 1 Then
                            strTmp05 = Trim$(Mid$(strTmp05, 12))
                          Else
                            strTmp05 = Trim$(Left$(strTmp05, (intPos4 - 1))) & Trim$(Mid$(strTmp05, (intPos4 + 12)))
                          End If
                        End If
                        Do While Right$(strTmp05, 1) = ")"
                          strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                        Loop
                        If Right$(strTmp05, 1) = ";" Then strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                        intPos4 = InStr(strTmp05, " AS ")
                        If intPos4 > 0 Then
                          strTmp05 = Trim$(Left$(strTmp05, intPos4))
                        End If
                        Do While Left$(strTmp05, 1) = "("
                          strTmp05 = Mid$(strTmp05, 2)
                        Loop
                        If Left$(strTmp05, 1) = "[" And Right$(strTmp05, 1) = "]" Then
                          strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                          strTmp05 = Mid$(strTmp05, 2)
                        End If
                        intPos4 = InStr(strTmp05, "DISTINCTROW")
                        If intPos4 > 0 Then
                          If intPos4 = 1 Then
                            strTmp05 = Trim$(Mid$(strTmp05, 12))
                          Else
                            strTmp05 = Trim$(Left$(strTmp05, (intPos4 - 1))) & Trim$(Mid$(strTmp05, (intPos4 + 12)))
                          End If
                        End If
                        If Trim$(strTmp01) = vbNullString Then
                          strTmp01 = Trim$(strTmp05) & "^"
                        Else
                          If Right$(strTmp01, 1) = "^" Then
                            strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp05)) & "^"
                          Else
                            strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp05)) & "^"
                          End If
                        End If
                      End If
                    Loop
                  Else
                    If Right$(strTmp05, 6) = " INNER" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 5)))
                    If Right$(strTmp05, 6) = " RIGHT" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 5)))
                    If Right$(strTmp05, 5) = " LEFT" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 4)))
                    If Right$(strTmp05, 1) = ";" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 1)))
                    intPos2 = InStr(strTmp05, " ON")
                    If intPos2 > 0 Then
                      strTmp05 = Trim$(Left$(strTmp05, intPos2))
                    End If
                    intPos2 = InStr(strTmp05, "ON ")
                    If intPos2 > 0 Then
                      strTmp05 = Trim$(Left$(strTmp05, (intPos2 - 1)))
                    End If
                    intPos2 = InStr(strTmp05, " AS ")
                    If intPos2 > 0 Then
                      strTmp05 = Trim$(Left$(strTmp05, intPos2))
                    End If
                    Do While Left$(strTmp05, 1) = "("
                      strTmp05 = Mid$(strTmp05, 2)
                    Loop
                    If Left$(strTmp05, 1) = "[" And Right$(strTmp05, 1) = "]" Then
                      strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                      strTmp05 = Mid$(strTmp05, 2)
                    End If
                    intPos2 = InStr(strTmp05, "DISTINCTROW")
                    If intPos2 > 0 Then
                      If intPos2 = 1 Then
                        strTmp05 = Trim$(Mid$(strTmp05, 12))
                      Else
                        strTmp05 = Trim$(Left$(strTmp05, (intPos2 - 1))) & Trim$(Mid$(strTmp05, (intPos2 + 12)))
                      End If
                    End If
                    If Trim$(strTmp01) = vbNullString Then
                      strTmp01 = Trim$(strTmp05) & "^"
                    Else
                      If Right$(strTmp01, 1) = "^" Then
                        strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp05)) & "^"
                      Else
                        strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp05)) & "^"
                      End If
                    End If
                  End If
                Else
                  If Right$(strTmp03, 1) = ";" Then strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                  intPos2 = InStr(strTmp03, ",")
                  If intPos2 > 0 Then
                    ' ** May be Cartesian.
                    strTmp05 = Trim$(Mid$(strTmp03, (intPos2 + 1)))
                    strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                    intPos3 = InStr(strTmp05, " AS ")
                    If intPos3 > 0 Then
                      strTmp05 = Trim$(Left$(strTmp05, intPos3))
                      If strTmp05 = strTmp03 Then
                        Do While Left$(strTmp03, 1) = "("
                          strTmp03 = Mid$(strTmp03, 2)
                        Loop
                        intPos3 = InStr(strTmp03, " AS ")
                        If intPos3 > 0 Then
                          strTmp03 = Trim$(Left$(strTmp03, intPos3))
                        End If
                        If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                          strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                          strTmp03 = Mid$(strTmp03, 2)
                        End If
                        intPos3 = InStr(strTmp03, "DISTINCTROW")
                        If intPos3 > 0 Then
                          If intPos3 = 1 Then
                            strTmp03 = Trim$(Mid$(strTmp03, 12))
                          Else
                            strTmp03 = Trim$(Left$(strTmp03, (intPos3 - 1))) & Trim$(Mid$(strTmp03, (intPos3 + 12)))
                          End If
                        End If
                        If Trim$(strTmp01) = vbNullString Then
                          strTmp01 = Trim$(strTmp03) & "^"
                        Else
                          If Right$(strTmp01, 1) = "^" Then
                            strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                          Else
                            strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp03)) & "^"
                          End If
                        End If
                      Else
                        'Stop
                        Debug.Print "'ERR QRYX 1: " & ![qry_name]
                      End If
                    Else
                      ' ** May be Cartesian!
                      Do While intPos2 > 0
                        Do While Left$(strTmp03, 1) = "("
                          strTmp03 = Mid$(strTmp03, 2)
                        Loop
                        intPos3 = InStr(strTmp03, " AS ")
                        If intPos3 > 0 Then
                          strTmp03 = Trim$(Left$(strTmp03, intPos3))
                        End If
                        If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                          strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                          strTmp03 = Mid$(strTmp03, 2)
                        End If
                        intPos3 = InStr(strTmp03, "DISTINCTROW")
                        If intPos3 > 0 Then
                          If intPos3 = 1 Then
                            strTmp03 = Trim$(Mid$(strTmp03, 12))
                          Else
                            strTmp03 = Trim$(Left$(strTmp03, (intPos3 - 1))) & Trim$(Mid$(strTmp03, (intPos3 + 12)))
                          End If
                        End If
                        If Trim$(strTmp01) = vbNullString Then
                          strTmp01 = Trim$(strTmp03) & "^"
                        Else
                          If Right$(strTmp01, 1) = "^" Then
                            strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                          Else
                            strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp03)) & "^"
                          End If
                        End If
                        strTmp03 = strTmp05
                        intPos2 = InStr(strTmp03, ",")
                        If intPos2 > 0 Then
                          ' ** May be Cartesian.
                          strTmp05 = Trim$(Mid$(strTmp03, (intPos2 + 1)))
                          If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                            strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                            strTmp03 = Mid$(strTmp03, 2)
                          End If
                          strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                        Else
                          If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                            strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                            strTmp03 = Mid$(strTmp03, 2)
                          End If
                          intPos2 = InStr(strTmp03, "DISTINCTROW")
                          If intPos3 > 0 Then
                            If intPos3 = 1 Then
                              strTmp03 = Trim$(Mid$(strTmp03, 12))
                            Else
                              strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1))) & Trim$(Mid$(strTmp03, (intPos2 + 12)))
                            End If
                          End If
                          If Trim$(strTmp01) = vbNullString Then
                            strTmp01 = Trim$(strTmp03) & "^"
                          Else
                            If Right$(strTmp01, 1) = "^" Then
                              strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                            Else
                              strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp03)) & "^"
                            End If
                          End If
                        End If
                      Loop
                    End If
                  Else
                    Do While Left$(strTmp03, 1) = "("
                      strTmp03 = Mid$(strTmp03, 2)
                    Loop
                    intPos3 = InStr(strTmp03, " AS ")
                    If intPos3 > 0 Then
                      strTmp03 = Trim$(Left$(strTmp03, intPos3))
                    End If
                    intPos3 = InStr(strTmp03, "DISTINCTROW")
                    If intPos3 > 0 Then
                      If intPos3 = 1 Then
                        strTmp03 = Trim$(Mid$(strTmp03, 12))
                      Else
                        strTmp03 = Trim$(Left$(strTmp03, (intPos3 - 1))) & Trim$(Mid$(strTmp03, (intPos3 + 12)))
                      End If
                    End If
                    If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                      strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                      strTmp03 = Mid$(strTmp03, 2)
                    End If
                    If Trim$(strTmp01) = vbNullString Then
                      strTmp01 = Trim$(strTmp03) & "^"
                    Else
                      If Right$(strTmp01, 1) = "^" Then
                        strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                      Else
                        strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp03)) & "^"
                      End If
                    End If
                  End If
                End If
                If strTmp06 <> vbNullString Then
                  intPos2 = InStr(strTmp06, ")")
                  If intPos2 > 0 Then
                    strTmp05 = Trim$(Left$(strTmp06, (intPos2 - 1)))
                    intPos2 = InStr(strTmp05, " AS ")
                    If intPos2 > 0 Then
                      strTmp05 = Trim$(Left$(strTmp05, intPos2))
                    End If
                    intPos2 = InStr(strTmp05, "DISTINCTROW")
                    If intPos2 > 0 Then
                      If intPos2 = 1 Then
                        strTmp05 = Trim$(Mid$(strTmp05, 12))
                      Else
                        strTmp05 = Trim$(Left$(strTmp05, (intPos2 - 1))) & Trim$(Mid$(strTmp05, (intPos2 + 12)))
                      End If
                    End If
                    If Left$(strTmp05, 1) = "[" And Right$(strTmp05, 1) = "]" Then
                      strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                      strTmp05 = Mid$(strTmp05, 2)
                    End If
                    Do While Left$(strTmp05, 1) = "("
                      strTmp05 = Mid$(strTmp05, 2)
                    Loop
                    strTmp03 = strTmp05
                    If Trim$(strTmp01) = vbNullString Then
                      strTmp01 = Trim$(strTmp03) & "^"
                    Else
                      If Right$(strTmp01, 1) = "^" Then
                        strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                      Else
                        strTmp01 = Trim$(Trim$(strTmp01) & "^" & Trim$(strTmp03)) & "^"
                      End If
                    End If
                  Else
                    Stop
                  End If
                End If
                If Right$(strTmp01, 1) = "^" Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                strRetVal = strTmp01
              Else
                Stop
              End If

            Case dbQCrosstab        ' ** Crosstab.
              Stop

            Case dbQDelete  ' ** Delete.
              intPos1 = InStr(strSQL, "FROM")
              If intPos1 > 0 Then
                intPos2 = InStr(intPos1, strSQL, "WHERE")
                If intPos2 > 0 Then
                  strTmp01 = Mid$(strSQL, (intPos1 + 4))
                  strTmp01 = Trim$(Left$(strTmp01, (InStr(strTmp01, "WHERE") - 1)))
                  strTmp01 = Rem_CRLF(strTmp01)  ' ** Module Function: modStringFuncs.
                Else
                  strTmp01 = Rem_CRLF(Trim$(Mid$(strSQL, (intPos1 + 4))))  ' ** Module Function: modStringFuncs.
                End If
                strTmp01 = Trim$(strTmp01)
                intPos1 = InStr(strTmp01, "LEFT JOIN")
                If intPos1 > 0 Then
                  strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))
                End If
                If Right$(strTmp01, 1) = ";" Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                If Left$(strTmp01, 1) = "[" And Right$(strTmp01, 1) = "]" Then
                  strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                  strTmp01 = Mid$(strTmp01, 2)
                End If
                strRetVal = strTmp01
              Else
                Stop
              End If

            Case dbQUpdate  ' ** Update.
              strTmp02 = Trim$(Rem_CRLF(strSQL))  ' ** Module Function: modStringFuncs.
              intPos1 = InStr(strTmp02, "UPDATE ")
              If intPos1 > 0 Then
                strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 6)))
                intPos2 = InStr(strTmp02, " SET ")
                If intPos2 > 0 Then
                  strTmp03 = Trim$(Left$(strTmp02, intPos2))
                  intPos2 = InStr(strTmp03, " JOIN ")
                  If intPos2 > 0 Then
                    strTmp03 = Trim$(Left$(strTmp03, intPos2))
                    If Right$(strTmp03, 6) = " INNER" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 5)))
                    If Right$(strTmp03, 6) = " RIGHT" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 5)))
                    If Right$(strTmp03, 5) = " LEFT" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 4)))
                    If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                      strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                      strTmp03 = Mid$(strTmp03, 2)
                    End If
                    intPos3 = InStr(strTmp03, "INSERT INTO ")
                    If intPos3 > 0 Then
                      strTmp03 = Trim$(Mid$(strTmp03, 12))
                    End If
                    intPos3 = InStr(strTmp03, "INTO ")
                    If intPos3 > 0 Then
                      strTmp03 = Trim$(Mid$(strTmp03, 5))
                    End If
                    intPos3 = InStr(strTmp03, "DISTINCTROW ")
                    If intPos3 > 0 Then
                      If intPos3 = 1 Then
                        strTmp03 = Trim$(Mid$(strTmp03, 12))
                      Else
                        strTmp03 = Trim$(Left$(strTmp03, (intPos3 - 1)) & Trim$(Mid$(strTmp03, (intPos3 + 12))))
                      End If
                    End If
                    strTmp01 = strTmp03
                    strRetVal = strTmp01
                  Else
                    intPos3 = InStr(strTmp03, ",")
                    If intPos3 > 0 Then
                      ' ** May be Cartesian.
                      strTmp04 = Trim$(Mid$(strTmp03, (intPos3 + 1)))
                      strTmp03 = Trim$(Left$(strTmp03, (intPos3 - 1)))
                      intPos3 = InStr(strTmp04, " AS ")
                      If intPos3 > 0 Then
                        strTmp04 = Trim$(Left$(strTmp04, intPos3))
                        If strTmp04 = strTmp03 Then
                          intPos3 = InStr(strTmp03, "INSERT INTO ")
                          If intPos3 > 0 Then
                            strTmp03 = Trim$(Mid$(strTmp03, 12))
                          End If
                          intPos3 = InStr(strTmp03, "INTO ")
                          If intPos3 > 0 Then
                            strTmp03 = Trim$(Mid$(strTmp03, 5))
                          End If
                          If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                            strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                            strTmp03 = Mid$(strTmp03, 2)
                          End If
                          intPos3 = InStr(strTmp03, "DISTINCTROW ")
                          If intPos3 > 0 Then
                            If intPos3 = 1 Then
                              strTmp03 = Trim$(Mid$(strTmp03, 12))
                            Else
                              strTmp03 = Trim$(Left$(strTmp03, (intPos3 - 1)) & Trim$(Mid$(strTmp03, (intPos3 + 12))))
                            End If
                          End If
                          strTmp01 = strTmp03
                          strRetVal = strTmp01
                        Else
                          Stop
                        End If
                      Else
                        ' ** qryReport_List_65w; Cartesian.
                        If strTmp03 = "qryReport_List_65v" And strTmp04 = "tblReportListControlType" Then
                          strTmp03 = "qryReport_List_65v, tblReportListControlType"
                        Else
                          Stop

                        End If
                      End If
                    Else
                      intPos3 = InStr(strTmp03, "INSERT INTO ")
                      If intPos3 > 0 Then
                        strTmp03 = Trim$(Mid$(strTmp03, 12))
                      End If
                      intPos3 = InStr(strTmp03, "INTO ")
                      If intPos3 > 0 Then
                        strTmp03 = Trim$(Mid$(strTmp03, 5))
                      End If
                      If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                        strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                        strTmp03 = Mid$(strTmp03, 2)
                      End If
                      intPos3 = InStr(strTmp03, "DISTINCTROW ")
                      If intPos3 > 0 Then
                        If intPos3 = 1 Then
                          strTmp03 = Trim$(Mid$(strTmp03, 12))
                        Else
                          strTmp03 = Trim$(Left$(strTmp03, (intPos3 - 1)) & Trim$(Mid$(strTmp03, (intPos3 + 12))))
                        End If
                      End If
                      strTmp01 = strTmp03
                      strRetVal = strTmp03
                    End If
                  End If
                Else
                  Stop
                End If
              Else
                Stop
              End If

            Case dbQAppend  ' ** Append.
              strTmp02 = Trim$(Rem_CRLF(strSQL))  ' ** Module Function: modStringFuncs.
              intPos1 = InStr(strTmp02, "INSERT INTO ")
              If intPos1 > 0 Then
                strTmp02 = Mid$(strTmp02, intPos1)
                intPos2 = InStr(strTmp02, "(")
                If intPos2 > 0 Then
                  strTmp03 = Trim$(Left$(strTmp02, (intPos2 - 1)))
                  intPos3 = InStr(strTmp03, "INSERT INTO ")
                  If intPos3 > 0 Then
                    strTmp03 = Trim$(Mid$(strTmp03, 12))
                  End If
                  intPos3 = InStr(strTmp03, "INTO ")
                  If intPos3 > 0 Then
                    strTmp03 = Trim$(Mid$(strTmp03, 5))
                  End If
                  If Left$(strTmp03, 1) = "[" And Right$(strTmp03, 1) = "]" Then
                    strTmp03 = Left$(strTmp03, (Len(strTmp03) - 1))
                    strTmp03 = Mid$(strTmp03, 2)
                  End If
                  strTmp01 = strTmp03
                  strRetVal = strTmp03
                Else
                  Stop
                End If
              Else
                Stop
              End If

            Case dbQMakeTable  ' ** Make-Table.
              strTmp02 = Trim$(Rem_CRLF(strSQL))  ' ** Module Function: modStringFuncs.
              intPos1 = InStr(strTmp02, "INTO ")
              If intPos1 > 0 Then
                strTmp02 = Mid$(strTmp02, intPos1)
                intPos2 = InStr(strTmp02, " FROM ")
                If intPos2 > 0 Then
                  strTmp03 = Trim$(Left$(strTmp02, intPos2))
                  intPos3 = InStr(strTmp03, "INSERT INTO ")
                  If intPos3 > 0 Then
                    strTmp03 = Trim$(Mid$(strTmp03, 12))
                  End If
                  intPos3 = InStr(strTmp03, "INTO ")
                  If intPos3 > 0 Then
                    strTmp03 = Trim$(Mid$(strTmp03, 5))
                  End If
                  strTmp01 = strTmp03
                  strRetVal = strTmp01
                Else
                  Stop
                End If
              Else
                Stop
              End If

            Case dbQDDL             ' ** Data-definition.
              ' ** Since I use Data-Definition Language queries mainly to create new tables,
              ' ** the tables listed may not exist, so we won't track these.
              lngQDDLs = lngQDDLs + 1&

            Case dbQSQLPassThrough  ' ** Pass-through (Microsoft Jet workspaces only).
              Stop

            Case dbQSetOperation  ' ** Union.
              If InStr(strSQL, "SELECT") = 0 Then
                intPos1 = InStr(strSQL, "UNION")
                If intPos1 > 0 Then
                  strTmp02 = strSQL
                  Do While intPos1 > 0
                    strTmp03 = Trim$(Left$(strTmp02, (intPos1 - 1)))
                    If Left$(strTmp03, 9) = "ALL TABLE" Then
                      strTmp03 = Trim$(Mid$(strTmp03, 10))
                    End If
                    If Left$(strTmp03, 5) = "TABLE" Then
                      strTmp03 = Trim$(Mid$(strTmp03, 6))
                    End If
                    intPos2 = InStr(strTmp03, "ORDER BY")
                    If intPos2 > 0 Then
                      strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                    End If
                    strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                    strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 5)))
                    intPos1 = InStr(strTmp02, "UNION")
                    If intPos1 = 0 Then
                      strTmp03 = strTmp02
                      If Left$(strTmp03, 9) = "ALL TABLE" Then
                        strTmp03 = Trim$(Mid$(strTmp03, 10))
                      End If
                      If Left$(strTmp03, 5) = "TABLE" Then
                        strTmp03 = Trim$(Mid$(strTmp03, 6))
                      End If
                      intPos2 = InStr(strTmp03, "ORDER BY")
                      If intPos2 > 0 Then
                        strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                      End If
                      strTmp01 = Trim$(strTmp01) & Trim$(strTmp03)
                    End If
                  Loop
                  If Right$(strTmp01, 1) = "^" Then strTmp01 = Trim$(Left$(strTmp01, (Len(strTmp01) - 1)))
                  strTmp01 = Trim$(Rem_CRLF(strTmp01))  ' ** Module Function: modStringFuncs.
                  If Right$(strTmp01, 1) = ";" Then strTmp01 = Trim$(Left$(strTmp01, (Len(strTmp01) - 1)))
                  intPos2 = InStr(strTmp01, " ^")
                  If intPos2 > 0 Then
                    ' ** I'm having a helluva time getting that space to go away before the caret.
                    Do While intPos2 > 0
                      strTmp01 = Left$(strTmp01, (intPos2 - 1)) & Mid$(strTmp01, (intPos2 + 1))
                      intPos2 = InStr(strTmp01, " ^")
                    Loop
                  End If
                  strRetVal = strTmp01
                Else
                  Stop
                End If
              Else
                ' ** Union containing at least 1 SELECT statment.
                intPos1 = InStr(strSQL, "UNION")
                If intPos1 > 0 Then
                  strTmp02 = strSQL
                  strTmp02 = Rem_CRLF(strTmp02)  ' ** Module Function: modStringFuncs.
                  intPos1 = InStr(strTmp02, "UNION")  ' ** Find it again, because Rem_CRLF() replaces 2 characters with 1.
                  Do While intPos1 > 0
                    strTmp03 = Trim$(Left$(strTmp02, (intPos1 - 1)))  ' ** Everything left of UNION.
                    If Left$(strTmp03, 6) = "SELECT" Then
                      intPos2 = InStr(strTmp03, "FROM")
                      If intPos2 > 0 Then
                        strTmp03 = Trim$(Mid$(strTmp03, (intPos2 + 4)))
                        intPos2 = InStr(strTmp03, "WHERE")
                        If intPos2 > 0 Then
                          strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                        End If
                        intPos2 = InStr(strTmp03, "ORDER BY")
                        If intPos2 > 0 Then
                          strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                        End If
                        intPos2 = InStr(strTmp03, "GROUP BY")
                        If intPos2 > 0 Then
                          strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                        End If
                        intPos2 = InStr(strTmp03, "HAVING")
                        If intPos2 > 0 Then
                          strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                        End If
                        intPos2 = InStr(strTmp03, "JOIN")
                        If intPos2 > 0 Then
                          ' ** Query JOIN verbs: 'INNER JOIN', 'LEFT JOIN', 'RIGHT JOIN'
                          ' ** Table locations: left of 1st JOIN verb, right of JOIN verb thereafter.
                          ' ** strTmp03 is everything to the right of FROM, without any criteria or sort.
                          strTmp04 = Trim$(Left$(strTmp03, (intPos2 - 1)))  ' ** 1st table.
                          If Right$(strTmp04, 6) = " INNER" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                          If Right$(strTmp04, 6) = " RIGHT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                          If Right$(strTmp04, 5) = " LEFT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 4)))
                          strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp04)) & "^"  ' ** strTmp01 now has 1st table in it.
                          ' ** For the rest of the JOIN's, tables are on the right.
                          strTmp05 = Trim$(Mid$(strTmp03, (intPos2 + 4)))   ' ** Everything right of 1st JOIN.
                          Do While Left$(strTmp05, 1) = "("
                            strTmp05 = Mid$(strTmp05, 2)
                          Loop
                          intPos3 = InStr(strTmp05, " JOIN ")
                          If intPos3 > 0 Then
                            Do While intPos3 > 0
                              strTmp04 = Trim$(Left$(strTmp05, intPos3))  ' ** Table will be at the left, before any ON's.
                              Do While Left$(strTmp04, 1) = "("
                                strTmp04 = Mid$(strTmp04, 2)
                              Loop
                              intPos4 = InStr(strTmp04, " ON ")
                              If intPos4 > 0 Then
                                strTmp04 = Trim$(Left$(strTmp04, intPos4))
                              End If
                              Do While Right$(strTmp04, 1) = ")"
                                strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                              Loop
                              If Right$(strTmp04, 6) = " INNER" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                              If Right$(strTmp04, 6) = " RIGHT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                              If Right$(strTmp04, 5) = " LEFT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 4)))
                              If Right$(strTmp04, 1) = ";" Then strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                              strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp04)) & "^"
                              strTmp05 = Trim$(Mid$(strTmp05, (intPos3 + 5)))
                              intPos3 = InStr(strTmp05, " JOIN ")
                              If intPos3 = 0 Then
                                intPos4 = InStr(strTmp05, " ON ")
                                If intPos4 > 0 Then
                                  strTmp05 = Trim$(Left$(strTmp05, intPos4))
                                End If
                                Do While Right$(strTmp05, 1) = ")"
                                  strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                                Loop
                                If Right$(strTmp05, 1) = ";" Then strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                                strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp05)) & "^"
                              End If
                            Loop
                          Else
                            If Right$(strTmp05, 6) = " INNER" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 5)))
                            If Right$(strTmp05, 6) = " RIGHT" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 5)))
                            If Right$(strTmp05, 5) = " LEFT" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 4)))
                            If Right$(strTmp05, 1) = ";" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 1)))
                            strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp05)) & "^"
                          End If
                        Else
                          If Right$(strTmp03, 6) = " INNER" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 5)))
                          If Right$(strTmp03, 6) = " RIGHT" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 5)))
                          If Right$(strTmp03, 5) = " LEFT" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 4)))
                          strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                        End If
                      Else
                        Stop
                      End If
                    Else
                      ' ** Mixture of TABLE and SELECT.
                      If Left$(strTmp03, 9) = "TABLE ALL" Then strTmp03 = Trim$(Mid$(strTmp03, 10))
                      If Left$(strTmp03, 5) = "TABLE" Then strTmp03 = Trim$(Mid$(strTmp03, 6))
                      strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                    End If
                    strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 5)))  ' ** Should be everything right of UNION.
                    intPos1 = InStr(strTmp02, " UNION ")
                    If intPos1 = 0 Then
                      If Left$(strTmp02, 6) = "SELECT" Then
                        strTmp03 = strTmp02
                        strTmp02 = vbNullString
                        intPos2 = InStr(strTmp03, " FROM ")
                        If intPos2 > 0 Then
                          strTmp03 = Trim$(Mid$(strTmp03, (intPos2 + 5)))
                          intPos2 = InStr(strTmp03, "WHERE")
                          If intPos2 > 0 Then
                            strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                          End If
                          intPos2 = InStr(strTmp03, "ORDER BY")
                          If intPos2 > 0 Then
                            strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                          End If
                          intPos2 = InStr(strTmp03, "GROUP BY")
                          If intPos2 > 0 Then
                            strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                          End If
                          intPos2 = InStr(strTmp03, "HAVING")
                          If intPos2 > 0 Then
                            strTmp03 = Trim$(Left$(strTmp03, (intPos2 - 1)))
                          End If
                          intPos2 = InStr(strTmp03, "JOIN")
                          If intPos2 > 0 Then
                            ' ** Query JOIN verbs: 'INNER JOIN', 'LEFT JOIN', 'RIGHT JOIN'
                            ' ** Table locations: left of 1st JOIN verb, right of JOIN verb thereafter.
                            ' ** strTmp03 is everything to the right of FROM, without any criteria or sort.
                            strTmp04 = Trim$(Left$(strTmp03, (intPos2 - 1)))  ' ** 1st table.
                            If Right$(strTmp04, 6) = " INNER" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                            If Right$(strTmp04, 6) = " RIGHT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                            If Right$(strTmp04, 5) = " LEFT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 4)))
                            If Right$(strTmp04, 1) = ";" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 1)))
                            strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp04)) & "^"  ' ** strTmp01 now has 1st table in it.
                            ' ** For the rest of the JOIN's, tables are on the right.
                            strTmp05 = Trim$(Mid$(strTmp03, (intPos2 + 4)))   ' ** Everything right of 1st JOIN.
                            Do While Left$(strTmp05, 1) = "("
                              strTmp05 = Mid$(strTmp05, 2)
                            Loop
                            intPos3 = InStr(strTmp05, " JOIN ")
                            If intPos3 > 0 Then
                              Do While intPos3 > 0
                                strTmp04 = Trim$(Left$(strTmp05, intPos3))  ' ** Table will be at the left, before any ON's.
                                Do While Left$(strTmp04, 1) = "("
                                  strTmp04 = Mid$(strTmp04, 2)
                                Loop
                                intPos4 = InStr(strTmp04, " ON ")
                                If intPos4 > 0 Then
                                  strTmp04 = Trim$(Left$(strTmp04, intPos4))
                                End If
                                Do While Right$(strTmp04, 1) = ")"
                                  strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                                Loop
                                If Right$(strTmp04, 6) = " INNER" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                                If Right$(strTmp04, 6) = " RIGHT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 5)))
                                If Right$(strTmp04, 5) = " LEFT" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 4)))
                                If Right$(strTmp04, 1) = ";" Then strTmp04 = Left$(strTmp04, (Len(strTmp04) - 1))
                                If Right$(strTmp04, 1) = ";" Then strTmp04 = Trim$(Left$(strTmp04, (Len(strTmp04) - 1)))
                                strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp04)) & "^"
                                strTmp05 = Trim$(Mid$(strTmp05, (intPos3 + 5)))
                                intPos3 = InStr(strTmp05, " JOIN ")
                                If intPos3 = 0 Then
                                  intPos4 = InStr(strTmp05, " ON ")
                                  If intPos4 > 0 Then
                                    strTmp05 = Trim$(Left$(strTmp05, intPos4))
                                  End If
                                  Do While Right$(strTmp05, 1) = ")"
                                    strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                                  Loop
                                  If Right$(strTmp05, 1) = ";" Then strTmp05 = Left$(strTmp05, (Len(strTmp05) - 1))
                                  strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp05)) & "^"
                                End If
                              Loop
                            Else
                              If Right$(strTmp05, 6) = " INNER" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 5)))
                              If Right$(strTmp05, 6) = " RIGHT" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 5)))
                              If Right$(strTmp05, 5) = " LEFT" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 4)))
                              If Right$(strTmp05, 1) = ";" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 1)))
                              If Right$(strTmp05, 1) = ";" Then strTmp05 = Trim$(Left$(strTmp05, (Len(strTmp05) - 1)))
                              strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp05)) & "^"
                            End If
                          Else
                            If Right$(strTmp03, 6) = " INNER" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 5)))
                            If Right$(strTmp03, 6) = " RIGHT" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 5)))
                            If Right$(strTmp03, 5) = " LEFT" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 4)))
                            If Right$(strTmp03, 1) = ";" Then strTmp03 = Trim$(Left$(strTmp03, (Len(strTmp03) - 1)))
                            strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                          End If
                        Else
                          Stop
                        End If
                      Else
                        ' ** Check for TABLE, TABLE ALL.
                        strTmp03 = strTmp02
                        strTmp02 = vbNullString
                        If Left$(strTmp03, 10) = "TABLE ALL " Then
                          strTmp03 = Trim$(Mid$(strTmp03, 10))
                        End If
                        If Left$(strTmp03, 6) = "TABLE " Then
                          strTmp03 = Trim$(Mid$(strTmp03, 6))
                        End If
                        strTmp01 = Trim$(Trim$(strTmp01) & Trim$(strTmp03)) & "^"
                      End If
                    End If
                  Loop
                  If Right$(strTmp01, 1) = "^" Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                  strRetVal = strTmp01
                Else
                  Stop
                End If
              End If

            Case dbQSPTBulk         ' ** Used with dbQSQLPassThrough to specify a query that doesn't return records. (Microsoft Jet workspaces only).
              Stop

            Case dbQCompound        ' ** Compound.
              Stop

            Case dbQProcedure       ' ** Procedure (ODBCDirect workspaces only).
              Stop

            Case dbQAction          ' ** Action.
              Stop

            Case Else
              Debug.Print "'QRY TYPE? " & strQryName

            End Select  ' ** qrytype_type.

            ' ** All the tables for 1 query are now in strRetVal.
            If Trim$(strRetVal) <> vbNullString Then

              lngSources = 0&
              ReDim arr_varSource(S_ELEMS, 0)

              ' ** Various special cases.
              blnAdd = False
              intPos1 = InStr(strRetVal, "tblConnectionType, tblXAdmin_ExportQry")
              If intPos1 > 0 Then
                blnAdd = True
                intLen = Len("tblConnectionType, tblXAdmin_ExportQry")
                strTmp01 = Mid$(strRetVal, intPos1, intLen)
                intPos2 = InStr(strTmp01, ",")
                strTmp01 = Left$(strTmp01, (intPos2 - 1)) & "^" & Mid$(strTmp01, (intPos2 + 2))
              Else
                intPos1 = InStr(strRetVal, "tblQueryType, (tblXAdmin_ExportTbl")
                If intPos1 > 0 Then
                  blnAdd = True
                  intLen = Len("tblQueryType, (tblXAdmin_ExportTbl")
                  strTmp01 = Mid$(strRetVal, intPos1, intLen)
                  intPos2 = InStr(strTmp01, ",")
                  strTmp01 = Left$(strTmp01, (intPos2 - 1)) & "^" & Mid$(strTmp01, (intPos2 + 3))
                Else
                  intPos1 = InStr(strRetVal, "tblQueryType, tblXAdmin_Graphics")
                  If intPos1 > 0 Then
                    blnAdd = True
                    intLen = Len("tblQueryType, tblXAdmin_Graphics")
                    strTmp01 = Mid$(strRetVal, intPos1, intLen)
                    intPos2 = InStr(strTmp01, ",")
                    strTmp01 = Left$(strTmp01, (intPos2 - 1)) & "^" & Mid$(strTmp01, (intPos2 + 2))
                  Else
                    intPos1 = InStr(strRetVal, "(tblVBComponent_Declaration")
                    If intPos1 > 0 Then
                      strRetVal = Mid$(strRetVal, 2)
'WHAT IS THIS? 2  '(tblVBComponent_Declaration'
                    Else
                      intPos1 = InStr(strRetVal, "((tblVBComponent_Declaration_Detail")
                      If intPos1 > 0 Then
                        strRetVal = Mid$(strRetVal, 3)
'WHAT IS THIS? 2  '((tblVBComponent_Declaration_Detail'
                      Else
                        intPos1 = InStr(strRetVal, "(tblVBComponent_Declaration_Detail")
                        If intPos1 > 0 Then
                          strRetVal = Mid$(strRetVal, 2)
'WHAT IS THIS? 2  '(tblVBComponent_Declaration_Detail'
                        Else
                          intPos1 = InStr(strRetVal, "qryReport_List_65v, tblReportListControlType")
                          If intPos1 > 0 Then
                            blnAdd = True
                            intPos2 = InStr(strRetVal, ",")
                            strTmp01 = Left$(strRetVal, (intPos2 - 1)) & "^" & Mid$(strRetVal, (intPos2 + 2))
                          Else
                            If Right(strRetVal, 1) = ";" Then
                              strRetVal = Trim(Left(strRetVal, (Len(strRetVal) - 1)))
'WHAT IS THIS? 2  'qryPortfolioModeling_Select_05;'
                            Else
                              ' ** Next?
                            End If
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If

              If blnAdd = True Then
                If intPos1 > 1 Then
                  If (intPos1 + intLen) > Len(strRetVal) Then
                    strRetVal = Left$(strRetVal, (intPos1 - 1)) & strTmp01 & IIf(Right$(strRetVal, 1) = "^", "^", vbNullString)
                  Else
                    strRetVal = Left$(strRetVal, (intPos1 - 1)) & strTmp01 & Mid$(strRetVal, (intPos1 + intLen))
                  End If
                Else
                  If (intPos1 + intLen) > Len(strRetVal) Then
                    strRetVal = strTmp01 & IIf(Right$(strRetVal, 1) = "^", "^", vbNullString)
                  Else
                    strRetVal = strTmp01 & Mid$(strRetVal, (intPos1 + intLen))
                  End If
                End If
              End If
              blnAdd = False

              intPos1 = InStr(strRetVal, "^")
              If intPos1 > 0 Then
                Do While intPos1 > 0
                  lngSources = lngSources + 1&
                  lngE = lngSources - 1&
                  ReDim Preserve arr_varSource(S_ELEMS, lngE)
                  arr_varSource(S_NAM, lngE) = Left$(strRetVal, (intPos1 - 1))
                  If InStr(arr_varSource(S_NAM, lngE), " WHERE") > 0 Then
                    arr_varSource(S_NAM, lngE) = Left(arr_varSource(S_NAM, lngE), (InStr(arr_varSource(S_NAM, lngE), " WHERE") - 1))
                  End If
                  If arr_varSource(S_NAM, lngE) = "(tblForm_Control" Then
                    arr_varSource(S_NAM, lngE) = Mid$(arr_varSource(S_NAM, lngE), 2)
                    'Debug.Print "'QRY: " & strQryName & "  '(tblForm_Control'"
                    strRetVal = Mid$(strRetVal, 2)
                  End If
                  arr_varSource(S_DID, lngE) = lngThisDbsID
                  arr_varSource(S_QID, lngE) = CLng(0)
                  arr_varSource(S_TID, lngE) = CLng(0)
                  arr_varSource(S_ORD, lngE) = lngSources
                  strRetVal = Mid$(strRetVal, (intPos1 + 1))
                  intPos1 = InStr(strRetVal, "^")
                  If intPos1 = 0 Then
                    lngSources = lngSources + 1&
                    lngE = lngSources - 1&
                    ReDim Preserve arr_varSource(S_ELEMS, lngE)
                    arr_varSource(S_NAM, lngE) = strRetVal
                    If InStr(arr_varSource(S_NAM, lngE), " WHERE") > 0 Then
                      arr_varSource(S_NAM, lngE) = Left(arr_varSource(S_NAM, lngE), (InStr(arr_varSource(S_NAM, lngE), " WHERE") - 1))
                    End If
                    If arr_varSource(S_NAM, lngE) = "(tblForm_Control" Then
                      arr_varSource(S_NAM, lngE) = Mid$(arr_varSource(S_NAM, lngE), 2)
                      'Debug.Print "'QRY: " & strQryName & "  '(tblForm_Control'"
                      strRetVal = "tblForm_Control"
                    End If
                    arr_varSource(S_DID, lngE) = lngThisDbsID
                    arr_varSource(S_QID, lngE) = CLng(0)
                    arr_varSource(S_TID, lngE) = CLng(0)
                    arr_varSource(S_ORD, lngE) = lngSources
                  End If
                Loop
              Else
                lngSources = lngSources + 1&
                lngE = lngSources - 1&
                arr_varSource(S_NAM, lngE) = strRetVal
                If InStr(arr_varSource(S_NAM, lngE), " WHERE") > 0 Then
                  arr_varSource(S_NAM, lngE) = Left(arr_varSource(S_NAM, lngE), (InStr(arr_varSource(S_NAM, lngE), " WHERE") - 1))
                End If
                If arr_varSource(S_NAM, lngE) = "(tblForm_Control" Then
                  arr_varSource(S_NAM, lngE) = Mid$(arr_varSource(S_NAM, lngE), 2)
                  'Debug.Print "'QRY: " & strQryName & "  '(tblForm_Control'"
                  strRetVal = "tblForm_Control"
                ElseIf arr_varSource(S_NAM, lngE) = "(tblVBComponent_Declaration_Detail" Then
                  arr_varSource(S_NAM, lngE) = Mid$(arr_varSource(S_NAM, lngE), 2)
                  strRetVal = "tblVBComponent_Declaration_Detail"
'WHAT IS THIS? 2  '(tblVBComponent_Declaration_Detail'
                ElseIf arr_varSource(S_NAM, lngE) = "((tblDatabase" Then
                  arr_varSource(S_NAM, lngE) = Mid$(arr_varSource(S_NAM, lngE), 3)
                  strRetVal = "tblDatabase"
'WHAT IS THIS? 2  '((tblDatabase'
                End If
                arr_varSource(S_DID, lngE) = lngThisDbsID
                arr_varSource(S_QID, lngE) = CLng(0)
                arr_varSource(S_TID, lngE) = CLng(0)
                arr_varSource(S_ORD, lngE) = lngSources
              End If

              For lngY = 0& To (lngSources - 1&)
On Error Resume Next
                varTmp00 = DLookup("[qry_id]", "tblQuery", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
                  "[qry_name] = '" & arr_varSource(S_NAM, lngY) & "'")
                If ERR.Number <> 0 Then
On Error GoTo 0
                  Debug.Print "'ERR QRYX 2: " & arr_varSource(S_NAM, lngY)
                Else
On Error GoTo 0
                  If IsNull(varTmp00) = True Then
                    If arr_varSource(S_NAM, lngY) = "LedgerArchive" Then
                      varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = 3 And [tbl_name] = 'ledger'")
                    ElseIf arr_varSource(S_NAM, lngY) = "tblDataTypeDb1" Then
                      varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = 2 And [tbl_name] = 'tblDataTypeDb'")
                    Else
                      varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[tbl_name] = '" & arr_varSource(S_NAM, lngY) & "' And " & _
                        "[dbs_id] <> " & CStr(lngNotThisDbsID1) & " And [dbs_id] <> " & CStr(lngNotThisDbsID2))
                    End If
                    If IsNull(varTmp00) = True Then
                      If IsNumeric(Right(arr_varSource(S_NAM, lngY), 1)) = True Then
                        ' ** Lots of client queries are going to come up with additional tables, like 'ActiveAssets1'.
                        ' ** How should I handle them?
                        ' ** Just give them a TAJrnTmp.mdb number?
                        ' ** Will it matter if Ledger1, Ledger2, Ledger3 all point to the same table?
                        strTmp01 = Left(arr_varSource(S_NAM, lngY), (Len(arr_varSource(S_NAM, lngY)) - 1))
                        varTmp00 = DLookup("[tbl_id]", "tblDatabase_Table", "[dbs_id] = 5 And [tbl_name] = '" & arr_varSource(S_NAM, lngY) & "'")
                        If IsNull(varTmp00) = True Then
                          Debug.Print "'WHAT IS THIS? 1  '" & arr_varSource(S_NAM, lngY) & "'"
                          Stop
                        End If
                      Else
                        Debug.Print "'WHAT IS THIS? 2  '" & arr_varSource(S_NAM, lngY) & "'"
                        Stop
                      End If
                    Else
                      arr_varSource(S_TID, lngY) = CLng(varTmp00)
                    End If
                  Else
                    arr_varSource(S_QID, lngY) = CLng(varTmp00)
                  End If
                End If
              Next

              With rst2
                For lngY = 0& To (lngSources - 1&)
                  If arr_varSource(S_QID, lngY) > 0& Or arr_varSource(S_TID, lngY) > 0& Then
                    blnAdd = False: blnByOrd = False: blnByID = False
                    If .BOF = True And .EOF = True Then
                      blnAdd = True
                    Else
                      .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [qry_id] = " & CStr(lngQryID) & " And " & _
                        "[qryrecsrc_order] = " & CStr(arr_varSource(S_ORD, lngY))
                        ' ** Index: qry_id, qryrecsrc_ord : Unique
                      If .NoMatch = True Then
                        If arr_varSource(S_QID, lngY) <> 0& Then
                          ' ** qry_id, qry_id_recsrc : Unique
                          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [qry_id] = " & CStr(lngQryID) & " And " & _
                            "[qry_id_recsrc] = " & CStr(arr_varSource(S_QID, lngY))
                          If .NoMatch = True Then
                            blnAdd = True
                          Else
                            blnByID = True
                          End If
                        ElseIf arr_varSource(S_TID, lngY) <> 0& Then
                          ' ** qry_id, tbl_id_recsrc : Unique
                          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [qry_id] = " & CStr(lngQryID) & " And " & _
                            "[qry_id_recsrc] = " & CStr(arr_varSource(S_TID, lngY))
                          If .NoMatch = True Then
                            blnAdd = True
                          Else
                            blnByID = True
                          End If
                        Else
                          blnAdd = True
                        End If
                      Else
                        blnByOrd = True
                      End If
                    End If
                    If blnAdd = True Then
                      .AddNew
                      ![dbs_id] = arr_varSource(S_DID, lngY)
                      ![qry_id] = lngQryID
                      ![qryrecsrc_order] = arr_varSource(S_ORD, lngY)
                      ![qryrecsrc_name] = arr_varSource(S_NAM, lngY)
                      If arr_varSource(S_QID, lngY) <> 0& Then
                        ![qry_id_recsrc] = arr_varSource(S_QID, lngY)
                      ElseIf arr_varSource(S_TID, lngY) <> 0& Then
                        ![tbl_id_recsrc] = arr_varSource(S_TID, lngY)
                      End If
                      ![qryrecsrc_datemodified] = Now()
On Error Resume Next
                      .Update
                      If ERR.Number <> 0 Then
On Error GoTo 0
                        If arr_varSource(S_QID, lngY) <> 0& Then
                          varTmp00 = DLookup("[qryrecsrc_order]", "tblQuery_RecordSource", _
                            "[dbs_id] = " & CStr(arr_varSource(S_DID, lngY)) & " And [qry_id] = " & CStr(lngQryID) & " And " & _
                            "[qry_id_recsrc] = " & CStr(arr_varSource(S_QID, lngY)))
                          If IsNull(varTmp00) = True Then
                            Stop
                          Else
                            .CancelUpdate dbUpdateRegular
                          End If
                        ElseIf arr_varSource(S_TID, lngY) <> 0& Then
                          varTmp00 = DLookup("[qryrecsrc_order]", "tblQuery_RecordSource", _
                            "[dbs_id] = " & CStr(arr_varSource(S_DID, lngY)) & " And [qry_id] = " & CStr(lngQryID) & " And " & _
                            "[tbl_id_recsrc] = " & CStr(arr_varSource(S_TID, lngY)))
                          If IsNull(varTmp00) = True Then
                            Stop
                          Else
                            .CancelUpdate dbUpdateRegular
                          End If
                        Else
                          Stop
                        End If
                      Else
On Error GoTo 0
                      End If
                    Else
                      If ![qryrecsrc_name] <> arr_varSource(S_NAM, lngY) Then
                        .Edit
                        ![qryrecsrc_name] = arr_varSource(S_NAM, lngY)
                        ![qryrecsrc_datemodified] = Now()
                        .Update
                      End If
                      If ![qryrecsrc_order] <> arr_varSource(S_ORD, lngY) Then
                        .Edit
                        ![qryrecsrc_order] = arr_varSource(S_ORD, lngY)
                        ![qryrecsrc_datemodified] = Now()
                        .Update
                      End If
                      If arr_varSource(S_QID, lngY) <> 0& Then
                        If ![qry_id_recsrc] <> arr_varSource(S_QID, lngY) Then
                          .Edit
                          ![qry_id_recsrc] = arr_varSource(S_QID, lngY)
                          ![tbl_id_recsrc] = Null
                          ![qryrecsrc_datemodified] = Now()
                          .Update
                        End If
                      ElseIf arr_varSource(S_TID, lngY) <> 0& Then
                        If ![tbl_id_recsrc] <> arr_varSource(S_TID, lngY) Then
                          .Edit
                          ![tbl_id_recsrc] = arr_varSource(S_TID, lngY)
                          ![qry_id_recsrc] = Null
                          ![qryrecsrc_datemodified] = Now()
                          .Update
                        End If
                      End If
                    End If
                  End If
                Next

              End With  ' ** rst2.

            End If  ' ** vbNullString.

          End If  ' ** lngThisDbsID.

          If ((lngX Mod 100) = 0) Then
            Debug.Print "|  " & CStr(lngX) & " of " & CStr(lngRecs)
            Debug.Print "'|";
            DoEvents
          ElseIf ((lngX Mod 10) = 0) Then
            Debug.Print "|";
            DoEvents
          Else
            Debug.Print ".";
            DoEvents
          End If

          If lngX < lngRecs Then .MoveNext
        Next  ' ** lngRecs: lngX.

      End If  ' ** BOF/EOF.
      .Close
    End With  ' ** rst1.
    Set rst1 = Nothing
    Set qdf = Nothing
    Debug.Print "  " & CStr(lngRecs) & "  (PASS 1)"
    DoEvents

    Debug.Print "'UPDATE:  ";
    DoEvents

    ' ** Update zz_qry_Query_RecordSource_04 (tblQuery, with DLookups() to zz_qry_Query_RecordSource_03
    ' ** (zz_qry_Query_RecordSource_02 (tblQuery, linked to tblQuery_RecordSource, by specified CurrentAppName()),
    ' ** grouped by qry_id, with cnt), with qry_tblcnt_new).
    Set qdf = .QueryDefs("zz_qry_Query_RecordSource_05")
    qdf.Execute
    Set qdf = Nothing
    Debug.Print "| ";
    DoEvents

    ' ** Append zz_qry_Query_RecordSource_08 (zz_qry_Query_RecordSource_07 (zz_qry_Query_RecordSource_06
    ' ** (tblQuery_RecordSource, with has_qry_id_srctbl, has_tbl_id_srctbl, by specified CurrentAppName()),
    ' ** grouped and summed by qry_id), just has_qry_id_srctbl = 0) to tblQuery_SourceChain.
    Set qdf = .QueryDefs("zz_qry_Query_RecordSource_09")
    qdf.Execute
    Set qdf = Nothing
    Debug.Print "| ";
    DoEvents

    ' ** Append zz_qry_Query_SourceChain_20_01b (zz_qry_Query_SourceChain_20_01a (tblQuery_RecordSource,
    ' ** linked to tblQuery_SourceChain, whose parent is qrysrc_level = 0, by specified CurrentAppName()),
    ' ** grouped, with cnt) to tblQuery_SourceChain, as qrysrc_level = 1.
    Set qdf = .QueryDefs("zz_qry_Query_SourceChain_20_01c")
    qdf.Execute
    Set qdf = Nothing
    Debug.Print "| ";
    DoEvents

    ' ** Append zz_qry_Query_SourceChain_20_02b (zz_qry_Query_SourceChain_20_02a (tblQuery_RecordSource,
    ' ** linked to tblQuery_SourceChain, whose parent is qrysrc_level = 1, by specified CurrentAppName()),
    ' ** grouped, with cnt) to tblQuery_SourceChain, as qrysrc_level = 2.
    Set qdf = .QueryDefs("zz_qry_Query_SourceChain_20_02c")
    qdf.Execute
    Set qdf = Nothing
    Debug.Print "| ";
    DoEvents
End If  ' ** blnSkip.

    For lngX = 3& To 11&
      ' ** Append zz_qry_Query_SourceChain_20_03c (zz_qry_Query_SourceChain_20_03b (zz_qry_Query_SourceChain_20_03a
      ' ** (tblQuery_RecordSource, linked to tblQuery_SourceChain, whose parent is qrysrc_level = 2, by specified CurrentAppName()),
      ' ** grouped, with cnt), not in tblQuery_SourceChain) to tblQuery_SourceChain, as qrysrc_level = 3.
      Set qdf = .QueryDefs("zz_qry_Query_SourceChain_20_" & Right$("00" & CStr(lngX), 2) & "d")
      qdf.Execute
    Set qdf = Nothing
      Debug.Print "| ";
      DoEvents
    Next

    ' ** Append zz_qry_Query_SourceChain_34a (zz_qry_Query_SourceChain_33a (zz_qry_Query_SourceChain_32a
    ' ** (zz_qry_Query_SourceChain_31 (zz_qry_Query_SourceChain_30 (tblQuery_RecordSource, not in
    ' ** tblQuery_SourceChain, by specified CurrentAppName()), grouped by qry_id_srctbl) qry_id_srctbl,
    ' ** linked to tblQuery_RecordSource, for their parents), just those with tbl_id_srctbl), grouped by
    ' ** qry_id, qry_id_parent, not in tblQuery_SourceChain) to tblQuery_SourceChain, as qrysrc_level = 0.
    Set qdf = .QueryDefs("zz_qry_Query_SourceChain_34b")
    qdf.Execute
    Set qdf = Nothing
    Debug.Print "| ";
    DoEvents

    ' ** zz_qry_Query_SourceChain_35b (zz_qry_Query_SourceChain_35a (zz_qry_Query_SourceChain_33b
    ' ** (zz_qry_Query_SourceChain_32a (zz_qry_Query_SourceChain_31 (zz_qry_Query_SourceChain_30
    ' ** (tblQuery_RecordSource, not in tblQuery_SourceChain, by specified CurrentAppName()), grouped
    ' ** by qry_id_srctbl), qry_id_srctbl, linked to tblQuery_RecordSource, for their parents), just
    ' ** those with qry_id_srctbl), grouped by qry_id, qry_id_srctbl, not in tblQuery_SourceChain),
    ' ** linked to tblQuery_SourceChain, for parent level), with parent in tblQuery_SourceChain.
    lngX = DCount("[qry_id]", "zz_qry_Query_SourceChain_35c")

    Do While lngX > 0
      ' ** Append zz_qry_Query_SourceChain_35c (zz_qry_Query_SourceChain_35b (zz_qry_Query_SourceChain_35a
      ' ** (zz_qry_Query_SourceChain_33b (zz_qry_Query_SourceChain_32a (zz_qry_Query_SourceChain_31
      ' ** (zz_qry_Query_SourceChain_30 (tblQuery_RecordSource, not in tblQuery_SourceChain, by specified
      ' ** CurrentAppName()), grouped by qry_id_srctbl), qry_id_srctbl, linked to tblQuery_RecordSource,
      ' ** for their parents), just those with qry_id_srctbl), grouped by qry_id, qry_id_srctbl, not in
      ' ** tblQuery_SourceChain), linked to tblQuery_SourceChain, for parent level), with parent in
      ' ** tblQuery_SourceChain.) to tblQuery_SourceChain, as qrysrc_level = 1, 2 (repeat).
      Set qdf = .QueryDefs("zz_qry_Query_SourceChain_35d")
      qdf.Execute
      Set qdf = Nothing
      Debug.Print "| ";
      DoEvents
      ' ** zz_qry_Query_SourceChain_35b (zz_qry_Query_SourceChain_35a (zz_qry_Query_SourceChain_33b
      ' ** (zz_qry_Query_SourceChain_32a (zz_qry_Query_SourceChain_31 (zz_qry_Query_SourceChain_30
      ' ** (tblQuery_RecordSource, not in tblQuery_SourceChain, by specified CurrentAppName()), grouped
      ' ** by qry_id_srctbl), qry_id_srctbl, linked to tblQuery_RecordSource, for their parents), just
      ' ** those with qry_id_srctbl), grouped by qry_id, qry_id_srctbl, not in tblQuery_SourceChain),
      ' ** linked to tblQuery_SourceChain, for parent level), with parent in tblQuery_SourceChain.
      lngX = DCount("[qry_id]", "zz_qry_Query_SourceChain_35c")
    Loop

    .Close
  End With
  Debug.Print
  DoEvents

  'If TableExists("account1") = True Then  ' ** Module Function: modFileUtilities.
  '  DoCmd.DeleteObject acTable, "account1"
  'End If
  'If TableExists("tmpUpdatedValues") = True Then  ' ** Module Function: modFileUtilities.
  '  DoCmd.DeleteObject acTable, "tmpUpdatedValues"
  'End If
  If TableExists("zz_tbl_RePost_Posting") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "zz_tbl_RePost_Posting"
  End If
  If TableExists("LedgerArchive_Backup") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "LedgerArchive_Backup"
  End If

  Application.SetOption "Show System Objects", False  ' ** Hide system objects.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Debug.Print "'DONE!  " & THIS_PROC & "()  DDLS: " & CStr(lngQDDLs)
  DoEvents

  Beep

  Qry_Tbl_Doc = strRetVal

End Function

Public Function SQL_Assemble() As Boolean
' ** Not currently called.

  Const THIS_PROC As String = "SQL_Assemble"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf3 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset
  Dim lngRecs As Long
  Dim lngFirstLine_VBCom_ID As Long, lngFirstLine_LineNum As Long, lngFirstLine_SQL_ID As Long
  Dim lngLast_LineNum As Long, lngLast_ALine As Long
  Dim strSQL As String
  Dim strUniqueID As String
  Dim lngX As Long

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs

    Set rst1 = .OpenRecordset("zz_tbl_sql_code_04", dbOpenDynaset)

    Set qdf1 = .QueryDefs("zz_qry_SQL_Code_41a")  ' ** zz_tbl_sql_code_01.
    Set rst2 = qdf1.OpenRecordset
    With rst2
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      lngFirstLine_SQL_ID = 0&: lngLast_LineNum = 0&: lngLast_ALine = 0&
      For lngX = 1& To lngRecs
        If ![sql1_aline] = -1& Then
          ' ** Save previous SQL to zz_tbl_sql_code_04.
          If lngFirstLine_SQL_ID > 0& Then
            With rst1
              .AddNew
              ![sql4_src] = 1&
              ![sql_id] = lngFirstLine_SQL_ID
              If strSQL <> vbNullString Then
                ![sql4_sql_complete] = strSQL
              Else
                Stop
              End If
              ![sql4_err] = False
              strUniqueID = Right$("000000" & CStr(lngFirstLine_VBCom_ID), 6) & _
                Right$("0000000" & CStr(lngFirstLine_LineNum), 7) & _
                Right$("000000" & CStr(lngFirstLine_SQL_ID), 6)
              ![sql4_uniqueid] = strUniqueID
              ![sql4_datemodified] = Now()
              .Update
            End With
          End If
          lngFirstLine_VBCom_ID = ![vbcom_id]
          lngFirstLine_LineNum = ![sql1_linenum]
          lngFirstLine_SQL_ID = ![sql1_id]
          lngLast_LineNum = ![sql1_linenum]
          lngLast_ALine = ![sql1_aline]
          strSQL = ![sql1_sql_edit]
        Else
          If ![sql1_aline] = lngFirstLine_LineNum Then
            lngLast_ALine = ![sql1_aline]  ' ** Now the same as lngLast_LineNum.
            lngLast_LineNum = ![sql1_linenum]
            strSQL = strSQL & vbCrLf & ![sql1_sql_edit]
          Else
            blnRetValx = False
            Set rst3 = dbs.OpenRecordset("zz_tbl_sql_code_05", dbOpenDynaset, dbAppendOnly)
            With rst3
              .AddNew
              ![vbcom_id] = rst2![vbcom_id]
              ![vbcom_id_expected] = lngFirstLine_VBCom_ID
              ![sql1_id] = rst2![sql1_id]
              '![sql1_id_expected] = lngFirstLine_SQL_ID  ' ** IRRELEVANT!
              ![sql1_aline] = rst2![sql1_aline]
              ![sql1_aline_expected] = lngFirstLine_LineNum
              ![sql1_linenum] = rst2![sql1_linenum]
              ![sql1_sql_edit] = rst2![sql1_sql_edit]
              ![sql1_procedure] = rst2![sql1_procedure]
              .Update
              .Close
            End With
            strUniqueID = "zz_" & Right$("0000" & CStr(rst2![sql1_id]), 4)
            'Debug.Print "'" & CStr(![sql1_id]) & " " & ![sql1_assign] & " " & CStr(![sql1_aline]) & " SQL: " & ![sql1_sql_edit]
            'Stop
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    With rst1
      .AddNew
      ![sql4_src] = 1&
      ![sql_id] = lngFirstLine_SQL_ID
      If strSQL <> vbNullString Then
        ![sql4_sql_complete] = strSQL
      Else
        Stop
      End If
      ![sql4_err] = False
      strUniqueID = Right$("000000" & CStr(lngFirstLine_VBCom_ID), 6) & _
        Right$("0000000" & CStr(lngFirstLine_LineNum), 7) & _
        Right$("000000" & CStr(lngFirstLine_SQL_ID), 6)
      ![sql4_uniqueid] = strUniqueID
      ![sql4_datemodified] = Now()
On Error Resume Next
      .Update
      If ERR.Number <> 0 Then
On Error GoTo 0
        ![sql4_uniqueid] = strUniqueID & "_x"
        .Update
      Else
On Error GoTo 0
      End If
      .Close
    End With

    .Close
  End With

  If blnRetValx = False Then
    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    Debug.Print "'MISMATCHES FOUND!"
  End If

  Beep

  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  Set qdf1 = Nothing
  Set qdf3 = Nothing
  Set dbs = Nothing

  SQL_Assemble = blnRetValx

End Function

Public Function SQL_HasAll(varInput As Variant) As Variant
' ** Not currently called.

  Const THIS_PROC As String = "SQL_HasAll"

  Dim intHit As Integer
  Dim varTmp00 As Variant
  Dim varRetVal As Variant

  varRetVal = Null

  intHit = 0

  If IsNull(varInput) = False Then
    varRetVal = varInput
    ' ** Form reference.
    varTmp00 = SQL_HasFrm(varRetVal)  ' ** Function: Below.
    If IsNull(varTmp00) = False Then
      varRetVal = varTmp00: intHit = intHit + 1: varTmp00 = Null
    End If
    ' ** Recordset assignment.
    varTmp00 = SQL_HasRst(varRetVal)  ' ** Function: Below.
    If IsNull(varTmp00) = False Then
      varRetVal = varTmp00: intHit = intHit + 1: varTmp00 = Null
    End If
    ' ** Variable reference.
    varTmp00 = SQL_HasVar(varRetVal)  ' ** Function: Below.
    If IsNull(varTmp00) = False Then
      varRetVal = varTmp00: intHit = intHit + 1: varTmp00 = Null
    End If
    ' ** Function call.
    varTmp00 = SQL_HasFnc(varRetVal)  ' ** Function: Below.
    If IsNull(varTmp00) = False Then
      varRetVal = varTmp00: intHit = intHit + 1: varTmp00 = Null
    End If
    ' ** SQL execution.
    varTmp00 = SQL_HasExe(varRetVal)  ' ** Function: Below.
    If IsNull(varTmp00) = False Then
      varRetVal = varTmp00: intHit = intHit + 1: varTmp00 = Null
    End If
    ' ** Recordset constant.
    varTmp00 = SQL_HasCon(varRetVal)  ' ** Function: Below.
    If IsNull(varTmp00) = False Then
      varRetVal = varTmp00: intHit = intHit + 1: varTmp00 = Null
    End If
    ' ** Field reference.
    varTmp00 = SQL_HasFld(varRetVal)  ' ** Function: Below.
    If IsNull(varTmp00) = False Then
      varRetVal = varTmp00: intHit = intHit + 1: varTmp00 = Null
    End If
    If intHit = 0 Then varRetVal = Null
  End If

  SQL_HasAll = varRetVal

End Function

Private Function SQL_HasFrm(varInput As Variant) As Variant
' ** Replace references to form controls with defaults,
' ** so that the SQL can be tested.
' ** Called by:
' **   SQL_HasAll(), Above

  Const THIS_PROC As String = "SQL_HasFrm"

  Dim intCnt As Integer, intCnt2 As Integer
  Dim intHit As Integer, intHit2 As Integer
  Dim strChk As String, strRep As String
  Dim intPos1 As Integer
  Dim intX As Integer, intY As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  intCnt = 0: intHit = 0: intCnt2 = 0: intHit2 = 0

  If IsNull(varInput) = False Then

    intPos1 = InStr(varInput, "Me.")
    If intPos1 > 0 Then

      varRetVal = varInput

      ' ** Count the number of times 'Me.' appears in the line.
      Do While intPos1 > 0
        intCnt = intCnt + 1
        intPos1 = InStr((intPos1 + 1), varInput, "Me.")
      Loop

      ' ** Though 'Me.' may appear multiple times, they could be for the same reference.
      ' ** If they're different references, succeeding passes will do nothing.
      For intX = 1 To intCnt

        ' ** Standard Account Number,
        ' ** preceded by {'" &} and followed by {& "'}.
        strRep = "11"  ' ** Default to Account 11.
        For intY = 1 To 8
          Select Case intY
          Case 1
            strChk = " & Me.accountno & "
            '"WHERE (((ActiveAssets.accountno) = '" & Me.accountno & "'));"
          Case 2
            strChk = " & Me.cmbAccounts & "
            '"WHERE (((ledger.accountno)='" & Me.cmbAccounts & "') AND "
          Case 3
            strChk = " & Me.dividendAccountNo & "
            '"WHERE (((qryJournal_Dividend_02b.accountno)='" & Me.dividendAccountNo & "')) "
          Case 4
            strChk = " & Me.interestAccountNo & "
            '"WHERE (((qryJournal_Interest_02b.accountno)='" & Me.interestAccountNo & "')) "
          Case 5
            strChk = " & Me.miscAccountNo & "
            '...
          Case 6
            strChk = " & Me.purchaseAccountNo & "
            '...
          Case 7
            strChk = " & Me.saleAccountNo & "
            '...
          Case 8
            strChk = " & Trim$(Me.cmbAccounts) & "
            '"WHERE account.accountno= '" & Trim$(Me.cmbAccounts) & "';"
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Non-standard Account Number,
        ' ** followed by {& "'}, but without opening {'" &}.
        strRep = "11"  ' ** Default to Account 11.
        For intY = 1 To 1
          Select Case intY
          Case 1
            strChk = "Trim(Me.cmbAccounts) & "
            'Trim(Me.cmbAccounts) & "' ORDER BY [balance date] DESC;", dbOpenSnapshot)
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
              varRetVal = strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Non-standard Account Number,
        ' ** preceded by {" &} and followed by {& "}, with single-quotes included in replacement.
        strRep = SQLFormatStr("11", dbText)
        For intY = 1 To 1
          Select Case intY
          Case 1
            strChk = " & SQLFormatStr(Me.interestAccountNo, dbText) & "
            '"WHERE (((qryJournal_Purchase_07.accountno) = " & SQLFormatStr(Me.interestAccountNo, dbText) & "));"
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Non-standard Account Number,
        ' ** includes closing paren from an OpenRecordset method.
        strRep = "11"  ' ** Default to Account 11.
        For intY = 1 To 2
          Select Case intY
          Case 1
            strChk = "left(Me.accountno, 2) & "
            'left(Me.accountno, 2) & "'")
          Case 2
            strChk = " & left$(Me.accountno, 2) & "
            '"WHERE accounttype.accounttype = '" & left$(Me.accountno, 2) & "'")
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 = 1 Then
            varRetVal = strRep & "'"  ' ** Strip closing paren.
            intHit = intHit + 1
          ElseIf intPos1 > 1 Then
            varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & "'"  ' ** Strip closing paren.
            intHit = intHit + 1
          End If
        Next

        ' ** Standard Date,
        ' ** preceded by {#" &} and followed by {& "#}.
        For intY = 1 To 6
          Select Case intY
          Case 1
            strRep = "1/1/2007"  ' ** Default to a good demo year.
            strChk = " & Me.DateStart & "
            '"((ledger.transdate) Between #" & Me.DateStart & "# And #" & Me.DateEnd & "#) AND "
          Case 2
            strRep = "1/1/2007"
            strChk = " & Me.TransDateStart & "
            '"AND ((ledger.transdate) BETWEEN #" & Me.TransDateStart & "# AND #" & Me.TransDateEnd & "#));"
          Case 3
            strRep = "12/31/2007"  ' ** Default to a good demo year.
            strChk = " & Me.DateEnd & "
            '"((ledger.transdate) Between #" & Me.DateStart & "# And #" & Me.DateEnd & "#) AND "
          Case 4
            strRep = "12/31/2007"
            strChk = " & Me.TransDateEnd & "
            '"AND ((ledger.transdate) BETWEEN #" & Me.TransDateStart & "# AND #" & Me.TransDateEnd & "#));"
          Case 5
            strRep = "12/31/2007"
            strChk = " & CStr(Me.DateEnd) & "
            '"(ledger.transdate)<=#" & CStr(Me.DateEnd) & "#) AND ((ledger.pcash)>=0)) "
          Case 6
            strRep = "12/31/2007"
            strChk = " & CStr(Me.DateEnd) & "
            'Set rstFrom = CurrentDb.OpenRecordset("SELECT * FROM ledger WHERE ledger.transdate <= #" & CStr(Me.DateEnd) & "#"
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = OCT) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = OCT) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Alternate Standard Date,
        ' ** preceded by {#" &} and followed by {& "#}.
        strRep = "05/31/2008"  ' ** Default to a good demo date.
        For intY = 1 To 2
          Select Case intY
          Case 1
            strChk = " & Me.newDate & "
            '"masterasset.currentDate = #" & Me.newDate & "# WHERE Not IsNull(masterasset.marketvalue);"
          Case 2
            strChk = " & Format(Me.transdate, " & Chr(34) & "MM/DD/YYYY" & Chr(34) & ") & "
            '"WHERE transdate = #" & Format(Me.transdate, "MM/DD/YYYY") & "# "
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = OCT) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = OCT) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Non-standard Date,
        ' ** followed by {& "'}, but without opening {#" &}.
        strRep = "12/31/2007"
        strChk = "Me.DateEnd & "
        'Me.DateEnd & "#)<=365,'ST','LT')) AS [Holding Period], masterasset.marketvaluecurrent, [masterasset].[currentdate] "
        intPos1 = InStr(varRetVal, strChk)
        If intPos1 = 1 Then
          varRetVal = strRep & Mid$(varRetVal, (Len(strChk) + 2))
          intHit = intHit + 1
        End If

        ' ** Standard CUSIP (Committee of Uniform Security Identification Procedures),
        ' ** preceded by {'" &} and followed by {& "'}.
        strRep = "001957109"  ' ** Default to AT&T.
        strChk = " & Me.cusip & "
        '"WHERE (((masterasset.cusip) = '" & Me.cusip & "'));"
        intPos1 = InStr(varRetVal, strChk)
        If intPos1 > 0 Then
          If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
              (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
            varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
            intHit = intHit + 1
          End If
        End If

        ' ** Discretionary vs. Non-Discretionary Accounts,
        ' ** within an IIF() statement.
        strRep = "No"  ' ** Default to option 1, All Accounts (frmRpt_StatementOfCondition).
        strChk = " & IIf(Me.opgAccountType.Value = 2, " & QTE & "Yes" & QTE & ", " & QTE & "No" & QTE & ") & "
        '"WHERE account.discretion = '" & IIf(Me.opgAccountType.Value = 2, "Yes", "No") & "' "
        intPos1 = InStr(varRetVal, strChk)
        If intPos1 > 0 Then
          If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
              (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
            varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
            intHit = intHit + 1
          End If
        End If

        ' ** Standard Asset Number,
        ' ** preceded by {" &} and followed by {& "}.
        strRep = "1"  ' ** Default to AT&T.
        For intY = 1 To 5
          Select Case intY
          Case 1
            strChk = " & Me.purchaseAssetNo & "
            '"WHERE (((masterasset.assetno) = " & Me.purchaseAssetNo & ") AND "
          Case 2
            strChk = " & CStr(Me.assetno) & "
            '"AND assetno = " & CStr(Me.assetno) & " "
          Case 3
            strChk = " & Me.Assetlist & "
            '"WHERE (((masterasset.assetno) = " & Me.Assetlist & "));"
          Case 4
            strChk = " & Me.cmbAssets & "
            '"HAVING (((ActiveAssets.assetno) = " & Me.cmbAssets & "));"
          Case 5
            strChk = " & CStr(Me.cmbAsset.Value) & "
            'rstAccounts![accountno] & "' AND assetno = " & CStr(Me.cmbAsset.Value) & " ORDER BY assetdate;", dbOpenSnapshot)
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Non-standard Asset Number,
        ' ** preceded by {" &}, but without closing {& "}.
        strRep = "1"  ' ** Default to AT&T.
        For intY = 1 To 1
          Select Case intY
          Case 1
            strChk = " & CStr(Me.interestAssetNo)"
            'Set rst = dbs.OpenRecordset("SELECT SUM(shareface) AS sumsf FROM activeAssets WHERE assetno = " & CStr(Me.interestAssetNo)
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Non-standard Asset Number,
        ' ** followed by {& "}, but without opening {" &}, .
        strRep = "1"  ' ** Default to AT&T.
        For intY = 1 To 1
          Select Case intY
          Case 1
            strChk = "CStr(Me.cmbAsset.Value) & "
            'CStr(Me.cmbAsset.Value) & ";", dbOpenSnapshot)
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
              varRetVal = strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Schedule_ID,
        ' ** preceded by {" &} and followed by {& "}.
        strRep = "1"  ' ** Default to 1, Estate Fee Schedule.
        strChk = " & Me.txtScheduleId & "
        '"WHERE (((ScheduleDetail.[Schedule_ID]) = " & Me.txtScheduleId & "));"
        intPos1 = InStr(varRetVal, strChk)
        If intPos1 > 0 Then
          If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
              (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
            varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
            intHit = intHit + 1
          End If
        End If

        ' ** Standard Location_ID.
        ' ** preceded by {" &} and followed by {& "}.
        strRep = "1"  ' ** Default to 1, North Fork Bank.
        strChk = " & Me.Location_ID & "
        '"WHERE [ActiveAssets].[Location_ID] = " & Me.Location_ID & ";"
        intPos1 = InStr(varRetVal, strChk)
        If intPos1 > 0 Then
          If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
              (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
            varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
            intHit = intHit + 1
          End If
        End If

        ' ** Miscellaneous others,
        ' ** preceded by {'" &} and followed by {& "'}.
        For intY = 1 To 2
          Select Case intY
          Case 1
            strRep = "Reinvest"  ' ** Default to Reinvest, previously only option. (VGC 02/15/09: added 'LTCG', 'StockSplit'.)
            strChk = " & Me.journalSubtype & "
            '"AND journalSubtype = '" & Me.journalSubtype & "' "
          Case 2
            strRep = "Received"  ' ** Default reinvestment Received.
            strChk = " & Me.journaltype & "
            '"AND journalType = '" & Me.journaltype & "' "
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Miscellaneous others,
        ' ** preceded by {" &} and followed by {& "}.
        For intY = 1 To 2
          Select Case intY
          Case 1
            strRep = "100"  ' ** Arbitrary (unsure of sign).
            strChk = " & CStr(Me.ICash) & "
            '"AND icash = " & CStr(Me.ICash) & " "
          Case 2
            strRep = "100"  ' ** Arbitrary (unsure of sign).
            strChk = " & CStr(Me.Cost) & "
            '"AND cost = " & CStr(Me.Cost) & ";"
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit = intHit + 1
            End If
          End If
        Next

        ' ** Simple trimming.
        For intY = 1 To 2
          Select Case intY
          Case 1
            strChk = "Me.cmbRecurringItems.RowSource = "
            'Me.cmbRecurringItems.RowSource = "SELECT RecurringItems.Name, RecurringItems.Type "
          Case 2
            strChk = "Me.cmbTaxCodes.RowSource = "
            'Me.cmbTaxCodes.RowSource = "SELECT DISTINCTROW taxcode.taxcode_description, taxcode.taxcode "
          End Select
          If Left$(varRetVal, Len(strChk)) = strChk Then
            varRetVal = Mid$(varRetVal, (Len(strChk) + 1))
            intHit = intHit + 1
          End If
        Next

      Next

      If intHit > 0 Then
        If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
        If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
        varRetVal = Trim(varRetVal)
      Else
        varRetVal = Null
      End If

    End If

    intPos1 = InStr(varInput, ".Form.")
    If intPos1 > 0 Then

      varRetVal = varInput

      ' ** Count the number of times '.Form.' appears in the line.
      Do While intPos1 > 0
        intCnt2 = intCnt2 + 1
        intPos1 = InStr((intPos1 + 1), varInput, ".Form.")
      Loop

      ' ** Though '.Form.' may appear multiple times, they could be for the same reference.
      ' ** If they're different references, succeeding passes will do nothing.
      For intX = 1 To intCnt2

        ' ** Standard Account Number,
        ' ** preceded by {'" &} and followed by {& "'}.
        strRep = "11"  ' ** Default to Account 11.
        For intY = 1 To 1
          Select Case intY
          Case 1
            strChk = " & Form_frmJournal.frmJournal_Sub4_Sold.Form.saleAccountNo & "
            '"AND ((ActiveAssets.accountno) = '" & Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleAccountNo & "')) "
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit2 = intHit2 + 1
            End If
          End If
        Next

        ' ** Standard Asset Number,
        ' ** preceded by {" &} and followed by {& "}.
        strRep = "1"  ' ** Default to AT&T.
        For intY = 1 To 2
          Select Case intY
          Case 1
            strChk = " & Form_frmJournal.frmJournal_Sub4_Sold.Form.saleAssetno.Column(2) & "
            '"WHERE (((ActiveAssets.assetno) = " & Forms("frmJournal").frmJournal_Sub4_Sold.Form.saleAssetno.Column(2) & ") "
          Case 2
            strChk = " & CStr(Form_frmJournal.frmJournal_Sub1_Dividend.Form.dividendAssetNo) & "
            '"WHERE (((masterasset.assetno) = " & CStr(Forms("frmJournal").frmJournal_Sub1_Dividend.Form.dividendAssetNo) & "));"
          End Select
          intPos1 = InStr(varRetVal, strChk)
          If intPos1 > 0 Then
            If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
                (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
              varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
              intHit2 = intHit2 + 1
            End If
          End If
        Next

      Next

      If intHit2 > 0 Then
        If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
        If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
        varRetVal = Trim(varRetVal)
      ElseIf intHit = 0 Then
        varRetVal = Null
      End If

    End If

  End If

  SQL_HasFrm = varRetVal

End Function

Private Function SQL_HasRst(varInput As Variant) As Variant
' ** Trim references to recordset assignments,
' ** so that the SQL can be tested.
' ** Called by:
' **   SQL_HasAll(), Above

  Const THIS_PROC As String = "SQL_HasRst"

  Dim intHit As Integer
  Dim strChk As String
  Dim intPos1 As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  intHit = 0

  If IsNull(varInput) = False Then

    strChk = ".OpenRecordset("
    intPos1 = InStr(varInput, strChk)
    If intPos1 > 0 Then

      varRetVal = varInput

      ' ** Simple trimming.
'Set rst = dbs.OpenRecordset("SELECT accounttype.* FROM accounttype WHERE accounttype.accounttype = '"
'Set rst = dbs.OpenRecordset("SELECT accounttype.* FROM accounttype "
'Set rst = dbs.OpenRecordset("SELECT COUNT(*) AS NumRecs FROM masterasset WHERE Left(description,3) = 'HA-' "
'Set rst = dbs.OpenRecordset("SELECT SUM(shareface) AS sumsf FROM activeAssets WHERE assetno = "
'Set rstBD = dbs.OpenRecordset("SELECT MAX([balance date]) AS MaxBD FROM balance "
'Set rstBalanceCheck = CurrentDb.OpenRecordset("SELECT balance.* FROM balance WHERE accountno = '"
'Set rst = dbs.OpenRecordset("SELECT COUNT(*) AS TranCount FROM ledger WHERE ledger.transdate > #"
'Set rstMasterAst = dbs.OpenRecordset("SELECT masterasset.* FROM masterasset WHERE assetno = "
'Set rstActiveAst = dbs.OpenRecordset("SELECT activeassets.* FROM activeassets WHERE accountno = '"
      If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
        varRetVal = Mid$(varRetVal, (intPos1 + Len(strChk)))
        intHit = intHit + 1
      End If

      If intHit > 0 Then
        If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
        If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
        varRetVal = Trim(varRetVal)
      Else
        varRetVal = Null
      End If

    End If

  End If

  SQL_HasRst = varRetVal

End Function

Private Function SQL_HasVar(varInput As Variant) As Variant
' ** Replace references to variables with defaults,
' ** so that the SQL can be tested.
' ** Called by:
' **   SQL_HasAll(), Above

  Const THIS_PROC As String = "SQL_HasVar"

  Dim intHit As Integer
  Dim strChk As String, strRep As String
  Dim intPos1 As Integer
  Dim intY As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  intHit = 0

  If IsNull(varInput) = False Then

    varRetVal = varInput

    ' ** Standard Account Number,
    ' ** preceded by {'" &} and followed by {& "'}.
    strRep = "11"  ' ** Default to Account 11.
    For intY = 1 To 3
      Select Case intY
      Case 1
        strChk = " & strAccountNo & "
        '"WHERE (((qryJournal_Dividend_02b.accountno)='" & strAccountNo & "')) "
        '"WHERE journaltype = 'Paid' AND PrintCheck = True AND AccountNo = '" & strAccountNo & "';"
      Case 2
        strChk = " & strAccountNumber & "
        '"WHERE (((Balance.accountno) = '" & strAccountNumber & "') AND ((Balance.[balance date])=#" & strEndDate & "#));"
        '"VALUES ('" & strAccountNumber & "', #" & strEndDate & "#, " & strIcash & ", " & strPcash & ", " & strCost & ", "
        '"WHERE (((Balance.accountno) = '" & strAccountNumber & "') AND ((Balance.[balance date])=#" & strEndDate & "#));"
      Case 3
        strChk = " & strAcctNo & "
        '"WHERE (((account.accountno) = '" & strAcctNo & "')) "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Related Accounts,
    ' ** preceded by {'" &} and followed by {& "'}.
    strRep = "11,20"  ' ** Default to Demo example.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strChk = " & strRelatedAccounts & "
        '"SELECT qryRelatedAssetList.assetno, '" & strRelatedAccounts & "' AS accountno, "
        '"GROUP BY qryRelatedAssetList.assetno, '" & strRelatedAccounts & "', qryRelatedAssetList.MasterAssetDescription, "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Account Name,
    ' ** preceded by {'" &} and followed by {& "'}.
    strRep = "William B. Johnson Trust"  ' ** Default to Account 11.
    For intY = 1 To 2
      Select Case intY
      Case 1
        strChk = " & " & Chr(34) & "*" & Chr(34) & " & strAcctName & " & Chr(34) & "*" & Chr(34) & " & "
        '"WHERE (account.shortname Like '" & "*" & strAcctName & "*" & "') "
        strRep = "*" & strRep & "*"
      Case 2
        strChk = " & strAccountShortName & "
        '"AND account.ShortName = '" & strAccountShortName & "';"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Standard Date,
    ' ** preceded by {#" &} and followed by {& "#}.
    For intY = 1 To 5
      Select Case intY
      Case 1
        strRep = "1/1/2006"
        strChk = " & CStr(datBeginDate) & "
        '"assetdate >= #" & CStr(datBeginDate) & "# AND journaltype IN ('Deposit','Purchase','Withdrawn','Sold') "
      Case 2
        strRep = "12/31/2006"  ' ** Default to a good demo Balance table year.
        strChk = " & strEndDate & "
        '"WHERE (((Balance.accountno) = '" & strAccountNumber & "') AND ((Balance.[balance date])=#" & strEndDate & "#));"
        '"VALUES ('" & strAccountNumber & "', #" & strEndDate & "#, " & strIcash & ", " & strPcash & ", " & strCost & ", "
      Case 3
        strRep = "12/31/2006"
        strChk = " & Format$(datStatementDate, " & Chr(34) & "mm/dd/yyyy" & Chr(34) & ") & "
        '"WHERE (((Journal.transdate)<=#" & Format$(datStatementDate, "mm/dd/yyyy") & "#));"
      Case 4
        strRep = "12/31/2006"
        strChk = " & Format(datStatementDate, " & Chr(34) & "mm/dd/yyyy" & Chr(34) & ") & "
        '"AND ledger.posted > #" & Format(datStatementDate, "mm/dd/yyyy") & "# "
      Case 5
        strRep = "12/31/2006"
        strChk = " & strCurrentDate & "
        '"#" & strCurrentDate & "# AS currentDate," & CoInfo & " "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = OCT) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = OCT) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Date,
    ' ** followed by {& "#}, but without opening {#" &}.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strRep = "12/31/2007"
        strChk = "CStr(datDate) & "
        'CStr(datDate) & "#", dbOpenSnapshot)
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = OCT) Then
          varRetVal = strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Standard Asset Number,
    ' ** preceded by {" &} and followed by {& "}.
    strRep = "1"  ' ** Default to AT&T.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strChk = " & CStr(varAssetNo) & "
        '"WHERE (((masterasset.assetno) = " & CStr(varAssetNo) & "));"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** User Name,
    ' ** preceded by {'" &} and followed by {& "'}.
    strRep = "TADemo"  ' ** Default to Demo user.
    strChk = " & strCurrentUser & "
    '"SELECT qryJournal_Purchase_07.*, '" & strCurrentUser & "' AS journal_USER "
    intPos1 = InStr(varRetVal, strChk)
    If intPos1 > 0 Then
      If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
          (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
        varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
        intHit = intHit + 1
      End If
    End If

    ' ** Standard Balance table values,
    ' ** preceded by {" &} and followed by {& "}.
    For intY = 1 To 5
      Select Case intY
      Case 1
        strRep = "-488.88"  ' ** Just use Demo values.
        strChk = " & strIcash & "
        '"VALUES ('" & strAccountNumber & "', #" & strEndDate & "#, " & strIcash & ", " & strPcash & ", " & strCost & ", "
      Case 2
        strRep = "6175.56"
        strChk = " & strPcash & "
        '"VALUES ('" & strAccountNumber & "', #" & strEndDate & "#, " & strIcash & ", " & strPcash & ", " & strCost & ", "
      Case 3
        strRep = "108224.4297"
        strChk = " & strCost & "
        '"VALUES ('" & strAccountNumber & "', #" & strEndDate & "#, " & strIcash & ", " & strPcash & ", " & strCost & ", "
      Case 4
        strRep = "198828.3622"
        strChk = " & strTotalMarketValue & "
        'strCost & ", Balance.TotalMarketValue = " & strTotalMarketValue & ", Balance.AccountValue = " & strAccountValue & " "
      Case 5
        strRep = "198828.3622"
        strChk = " & strAccountValue & "
        'strCost & ", Balance.TotalMarketValue = " & strTotalMarketValue & ", Balance.AccountValue = " & strAccountValue & " "
        'strTotalMarketValue & ", " & strAccountValue & ");"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Balance table values,
    ' ** followed by {& "}, but without opening {" &}.
    For intY = 1 To 2
      Select Case intY
      Case 1
        strRep = "108224.4297"
        strChk = "strCost & "
        'strCost & ", Balance.TotalMarketValue = " & strTotalMarketValue & ", Balance.AccountValue = " & strAccountValue & " "
      Case 2
        strRep = "198828.3622"
        strChk = "strTotalMarketValue & "
        'strTotalMarketValue & ", " & strAccountValue & ");"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Standard Fee Calculations table values.
    ' ** preceded by {" &} and followed by {& "}.
    For intY = 1 To 2
      Select Case intY
      Case 1
        strRep = "0.0055"
        strChk = " & CStr(dblRate) & "
        '"VALUES ('" & rst2![accountno] & "', " & CStr(dblRate) & ", " & CStr(dblRemainder) & ", "
      Case 2
        strRep = "181239.9338"
        strChk = " & CStr(dblRemainder) & "
        '"VALUES ('" & rst2![accountno] & "', " & CStr(dblRate) & ", " & CStr(dblRemainder) & ", "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Trim$(Left$(varRetVal, (intPos1 - 2))) & strRep & Trim$(Mid$(varRetVal, (intPos1 + Len(strChk) + 1)))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Fee Calculations table values.
    ' ** followed by {& "}, but without opening {" &}.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strRep = CStr(181239.9338 * 0.0055)
        strChk = "CStr(dblRemainder * dblRate) & "
        'CStr(dblRemainder * dblRate) & ");"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 = 1 Then
        If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Standard Reinvestment value,
    ' ** preceded by {" &} and followed by {& "}.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strRep = "-2.00"
        strChk = " & strPrice & "
        '"SELECT DISTINCTROW account.accountno, account.shortname, Sum([journal map].icash/" & strPrice & ") AS total_shareface, "
        'strPrice = Str(Form_frmMap_Reinvest_DivInt_Price.txtPrice * -1)
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Trim$(Left$(varRetVal, (intPos1 - 2))) & strRep & Trim$(Mid$(varRetVal, (intPos1 + Len(strChk) + 1)))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Boolean value,
    ' ** preceded by {" &} and followed by {& "}.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strRep = " True "
        strChk = " & strGLWhere & "
        '"AND CDate(Format(ActiveAssets.assetdate,'mm/dd/yyyy')) <= #" & Me.DateEnd & "#) AND " & strGLWhere & ";"
        'strGLWhere will return either True or False, no data involved.
        'strGLWhere = "((([masterasset].[marketvaluecurrent] * [activeassets].[shareface] * " & _
        '  "IIf([assettype] = '90', -1, 1)) - ([activeassets].[cost])) <> 0)"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Trim$(Left$(varRetVal, (intPos1 - 2))) & strRep & Trim$(Mid$(varRetVal, (intPos1 + Len(strChk) + 1)))
          intHit = intHit + 1
        End If
      End If
    Next

    If intHit > 0 Then
      If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
      If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
      varRetVal = Trim(varRetVal)
    Else
      varRetVal = Null
    End If

  End If

  SQL_HasVar = varRetVal

End Function

Private Function SQL_HasFnc(varInput As Variant) As Variant
' ** Replace references to functions with defaults,
' ** so that the SQL can be tested.
' ** Called by:
' **   SQL_HasAll(), Above

  Const THIS_PROC As String = "SQL_HasFnc"

  Dim intHit As Integer
  Dim strChk As String, strRep As String
  Dim intPos1 As Integer
  Dim intY As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  intHit = 0

  If IsNull(varInput) = False Then

    varRetVal = varInput

    ' ** Standard Company info,
    ' ** preceded by {" &} and followed by {& "}. (Single-quotes included with replacement.)
    strRep = CoInfo(False)  ' ** Default to Demo company.  ' ** Module Function: modQueryFunctions.
    strChk = " & CoInfo & "
    '"Sum([journal map].icash) AS total_icash, " & CoInfo & ", "
    intPos1 = InStr(varRetVal, strChk)
    If intPos1 > 0 Then
      If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
          (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
        varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
        intHit = intHit + 1
      End If
    End If

    ' ** Non-standard Company info,
    ' ** preceded by {" &}, but without closing {& "}.
    For intY = 1 To 3
      Select Case intY
      Case 1
        strRep = CoInfo(False)  ' ** Module Function: modQueryFunctions.
        strChk = " & CoInfo(False)"
        '"Nz([taxcode].[taxcode_description],'Unspecified') AS taxcode_description, masterasset.assettype, " & CoInfo(False)
      Case 2
        strRep = CoInfo(True)  ' ** Default to Demo company, GROUP BY version.  ' ** Module Function: modQueryFunctions.
        strChk = " & CoInfo(True)"
        '"Nz([taxcode].[taxcode_description],'Unspecified') AS taxcode_description, masterasset.assettype, " & CoInfo(True)
      Case 3
        strRep = CoInfo(False)  ' ** Module Function: modQueryFunctions.
        strChk = " & CoInfo"
        '" assettype.*," & CoInfo
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) Then
          varRetVal = Trim$(Left$(varRetVal, (intPos1 - 2))) & " " & Trim$(strRep)
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Company info,
    ' ** followed by {& "}, but without opening {" &}.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strRep = CoInfo(True)  ' ** Module Function: modQueryFunctions.
        strChk = "CoInfo(True) & "
        'CoInfo(True) & ";"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Trim$(strRep) & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Company info,
    ' ** on a standalone line.
    strRep = CoInfo(False)  ' ** Module Function: modQueryFunctions.
    strChk = "CoInfo(False)"
    'CoInfo(False)
    intPos1 = InStr(varRetVal, strChk)
    If intPos1 > 0 Then
      If Len(varRetVal) = Len(strChk) Then
        varRetVal = Trim$(strRep)
        intHit = intHit + 1
      End If
    End If

    ' ** Now() Date,
    ' ** preceded by {#" &} and followed by {& "#}.
    strRep = Format$(Now(), "m/d/yyyy")
    strChk = " & Format$(Now(), " & Chr(34) & "m/d/yyyy" & Chr(34) & ") & "
    '"VALUES ('" & Me.accountno & "', #" & Format$(Now(), "m/d/yyyy") & "#, 0, 0, 0, 0, 0);"
    intPos1 = InStr(varRetVal, strChk)
    If intPos1 > 0 Then
      If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = OCT) And _
          (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = OCT) Then
        varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
        intHit = intHit + 1
      End If
    End If

    If intHit > 0 Then
      If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
      If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
      varRetVal = Trim(varRetVal)
    Else
      varRetVal = Null
    End If

  End If

  SQL_HasFnc = varRetVal

End Function

Private Function SQL_HasExe(varInput As Variant) As Variant
' ** Trim references to SQL execution,
' ** so that the SQL can be tested.
' ** Called by:
' **   SQL_HasAll(), Above

  Const THIS_PROC As String = "SQL_HasExe"

  Dim intHit As Integer
  Dim strChk As String
  Dim intPos1 As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  intHit = 0

  If IsNull(varInput) = False Then

    varRetVal = varInput

    ' ** Simple trimming.
    strChk = "dbs.Execute "
    'dbs.Execute "UPDATE tmpAccountInfo SET assetno = NULL, MasterAssetDescription = NULL, "
    intPos1 = InStr(varRetVal, strChk)
    If intPos1 > 0 Then
      If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
        varRetVal = Trim$(Mid$(varRetVal, (Len(strChk) + 1)))
        intHit = intHit + 1
      End If
    End If

    If intHit > 0 Then
      If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
      If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
      varRetVal = Trim(varRetVal)
    Else
      varRetVal = Null
    End If

  End If

  SQL_HasExe = varRetVal

End Function

Private Function SQL_HasCon(varInput As Variant) As Variant
' ** Trim references to recordset constants,
' ** so that the SQL can be tested.
' ** Called by:
' **   SQL_HasAll(), Above

  Const THIS_PROC As String = "SQL_HasCon"

  Dim intHit As Integer
  Dim strChk As String
  Dim intPos1 As Integer
  Dim intY As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  intHit = 0

  If IsNull(varInput) = False Then

    varRetVal = varInput

    For intY = 1 To 2
      Select Case intY
      Case 1
        strChk = ", dbOpenDynaset)"
        '"ORDER BY ledger.transdate, ledger.accountno, ledger.assetno", dbOpenDynaset)
      Case 2
        strChk = ", dbOpenSnapshot)"
        '"ORDER BY assetdate DESC;", dbOpenSnapshot)
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) Then
          varRetVal = Left$(varRetVal, (intPos1 - 1))
          intHit = intHit + 1
        End If
      End If
    Next

    If intHit > 0 Then
      If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
      If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
      varRetVal = Trim(varRetVal)
    Else
      varRetVal = Null
    End If

  End If

  SQL_HasCon = varRetVal

End Function

Private Function SQL_HasFld(varInput As Variant) As Variant
' ** Replace references to recordset fields with defaults,
' ** so that the SQL can be tested.
' ** Called by:
' **   SQL_HasAll(), Above

  Const THIS_PROC As String = "SQL_HasFld"

  Dim intHit As Integer
  Dim strChk As String, strRep As String
  Dim intPos1 As Integer
  Dim intY As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  intHit = 0

  If IsNull(varInput) = False Then

    varRetVal = varInput

    ' ** Standard Account Number,
    ' ** preceded by {'" &} and followed by {& "'}.
    strRep = "11"  ' ** Default to Account 11.
    For intY = 1 To 3
      Select Case intY
      Case 1
        strChk = " & rst![accountno] & "
        '"WHERE (((account.accountno)='" & rst![accountno] & "'));"
      Case 2
        strChk = " & rst1![accountno] & "
        '"WHERE (((qryFeeCalculations_04.accountno) = '" & rst1![accountno] & "'));"
      Case 3
        strChk = " & rst2![accountno] & "
        '"VALUES ('" & rst2![accountno] & "', " & CStr(rst2![rate]) & ", " & CStr(rst2![Amount]) & ", "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE And Mid$(varRetVal, (intPos1 - 2), 1) = APO) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Account Number,
    ' ** followed by {& "'}, but without opening {'" &}.
    strRep = "11"  ' ** Default to Account 11.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strChk = "rstAccounts![accountno] & "
        'rstAccounts![accountno] & "' AND assetno = " & CStr(Me.cmbAsset.Value) & " ORDER BY assetdate;", dbOpenSnapshot)
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE And Mid$(varRetVal, (intPos1 + Len(strChk) + 1), 1) = APO) Then
          varRetVal = strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Account info,
    ' ** preceded by {" &} and followed by {& "}. (Where appropriate, single-quotes included with replacement.)
    For intY = 1 To 5
      Select Case intY
      Case 1
        strRep = SQLFormatStr("William B. Johnson Trust", dbText)  ' ** Default to Account 11.
        strChk = " & SQLFormatStr(rstTmpAccountInfo![shortname], dbText) & "
        '"shortname = " & SQLFormatStr(rstTmpAccountInfo![shortname], dbText) & ", "
      Case 2
        strRep = SQLFormatStr("North Fork Bank Trustee for the William B. Johnson Trust under Agreement Dated 11-12-98", dbText)
        strChk = " & SQLFormatStr(rstTmpAccountInfo![legalname], dbText) & "
        '"legalname = " & SQLFormatStr(rstTmpAccountInfo![legalname], dbText) & ", "
      Case 3
        strRep = SQLFormatStr("11", dbText)
        strChk = " & SQLFormatStr(rstTmpAccountInfo![accountno], dbText) & "
        '"WHERE accountno = " & SQLFormatStr(rstTmpAccountInfo![accountno], dbText) & ";"
      Case 4
        strRep = SQLFormatStr(859.12, dbCurrency)
        strChk = " & SQLFormatStr(rstTmpAccountInfo![ICash], dbCurrency) & "
        '"icash = " & SQLFormatStr(rstTmpAccountInfo![ICash], dbCurrency) & ", "
      Case 5
        strRep = SQLFormatStr(6106#, dbCurrency)
        strChk = " & SQLFormatStr(rstTmpAccountInfo![PCash], dbCurrency) & "
        '"pcash = " & SQLFormatStr(rstTmpAccountInfo![PCash], dbCurrency) & " "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Standard Asset Number,
    ' ** preceded by {" &} and followed by {& "}.
    strRep = "1"  ' ** Default to AT&T.
    For intY = 1 To 1
      Select Case intY
      Case 1
        strChk = " & CStr(rstMaster![assetno]) & "
        '"WHERE (((ActiveAssets.assetno) = " & CStr(rstMaster![assetno]) & ")) "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Asset Number,
    ' ** entirely within a function.
    For intY = 1 To 2
      Select Case intY
      Case 1
        strRep = SQLFormatStr(1, dbLong)  ' ** Default to AT&T.
        strChk = "SQLFormatStr(varAssetNo, rst1![assetno].Type)  ' ** Module Function: modUtilities."
        'SQLFormatStr(varAssetNo, rst1![assetno].Type)  ' ** Module Function: modUtilities.
      Case 2
        strRep = SQLFormatStr(1, dbLong)  ' ** Default to AT&T.
        strChk = "SQLFormatStr(varAssetNo, rst1![assetno].Type)"  ' ** In case it's already stripped.
        'SQLFormatStr(varAssetNo, rst1![assetno].Type)  ' ** Module Function: modUtilities.
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 = 1 Then
        varRetVal = strRep
        intHit = intHit + 1
      End If
    Next

    ' ** Standard Fee Calculations table values.
    ' ** preceded by {" &} and followed by {& "}.
    For intY = 1 To 2
      Select Case intY
      Case 1
        strRep = "0.0055"  ' ** Default to Demo Fee Calculations for Account 11.
        strChk = " & CStr(rst2![rate]) & "
        '"VALUES ('" & rst2![accountno] & "', " & CStr(rst2![rate]) & ", " & CStr(rst2![Amount]) & ", "
      Case 2
        strRep = "81067.1026"
        strChk = " & CStr(rst2![Amount]) & "
        '"VALUES ('" & rst2![accountno] & "', " & CStr(rst2![rate]) & ", " & CStr(rst2![Amount]) & ", "
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If (Mid$(varRetVal, (intPos1 - 1), 1) = QTE) And _
            (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = Left$(varRetVal, (intPos1 - 2)) & strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard Fee Calculations table values.
    ' ** entirely within a function.
    For intY = 1 To 2
      Select Case intY
      Case 1
        strRep = CStr(0.0055 * 81067.1026)
        strChk = "CStr(rst2![Amount] * rst2![rate]) & "
        'CStr(rst2![Amount] * rst2![rate]) & ");"
      Case 2
        strRep = CStr(181239.9338 * 0.0055)
        strChk = "CStr(dblRemainder * rst2![rate]) & "
        'CStr(dblRemainder * rst2![rate]) & ");"
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 > 0 Then
        If intPos1 = 1 And (Mid$(varRetVal, (intPos1 + Len(strChk)), 1) = QTE) Then
          varRetVal = strRep & Mid$(varRetVal, (intPos1 + Len(strChk) + 1))
          intHit = intHit + 1
        End If
      End If
    Next

    ' ** Non-standard CUSIP (Committee of Uniform Security Identification Procedures),
    ' ** entirely  within a function.
    For intY = 1 To 2
      Select Case intY
      Case 1
        strRep = SQLFormatStr("001957109", dbText)  ' ** Default to AT&T.
        strChk = "SQLFormatStr(rst1![cusip], rst1![cusip].Type)  ' ** Module Function: modUtilities."
        'SQLFormatStr(rst1![cusip], rst1![cusip].Type)  ' ** Module Function: modUtilities.
      Case 2
        strRep = SQLFormatStr("001957109", dbText)  ' ** Default to AT&T.
        strChk = "SQLFormatStr(rst1![cusip], rst1![cusip].Type)"  ' ** In case it's already stripped.
        'SQLFormatStr(rst1![cusip], rst1![cusip].Type)  ' ** Module Function: modUtilities.
      End Select
      intPos1 = InStr(varRetVal, strChk)
      If intPos1 = 1 Then
        varRetVal = strRep
        intHit = intHit + 1
      End If
    Next

    If intHit > 0 Then
      If Left$(varRetVal, 1) = QTE Then varRetVal = Mid$(varRetVal, 2)
      If Right$(varRetVal, 1) = QTE Then varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
      varRetVal = Trim(varRetVal)
    Else
      varRetVal = Null
    End If

  End If

  SQL_HasFld = varRetVal

End Function

Public Function SQL_LineNums() As Boolean
' ** Not currently called.

  Const THIS_PROC As String = "SQL_LineNums"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngRecs As Long, strTable As String, strQry As String
  Dim lngFirstLine As Long, lngLastLine As Long
  Dim lngX As Long, lngY As Long

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs

    For lngX = 1& To 1&  '3&
      strTable = vbNullString: strQry = vbNullString
      Select Case lngX
      Case 1&
        'strTable = "zz_tbl_sql_code_01"
        strQry = "zz_qry_SQL_Code_03a"
      Case 2&
        strTable = "zz_tbl_sql_code_02"
      Case 3&
        strTable = "zz_tbl_sql_code_03"
      End Select
      If strTable <> vbNullString Then
        Set rst = .OpenRecordset(strTable, dbOpenDynaset)
      Else
        Set qdf = .QueryDefs(strQry)
        Set rst = qdf.OpenRecordset
      End If
      With rst
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        lngFirstLine = 0&: lngLastLine = 0&
        For lngY = 1& To lngRecs
          If ![sql1_aline] > 0 Then
            ' ** I think I should leave these alone, since I may have updated them manually or via a query.
            lngLastLine = ![sql1_linenum]
          Else
            ' ** BTW, [sql1_assign] is irrelevant here, since that applies to variables, etc.
            If ![sql1_linenum] <> lngLastLine + 1& Then
              lngFirstLine = ![sql1_linenum]
              lngLastLine = lngFirstLine
              If ![sql1_aline] <> -1& Then
                .Edit
                ![sql1_aline] = -1&
                ![sql1_IsUpd] = True
                ![sql1_datemodified] = Now()
                .Update
              End If
            Else
              ' ** For now, it will be assumed that adjascent lines are
              ' ** continuations of previous lines. These subsequent lines
              ' ** may also have an assignment, e.g., strSQL = strSQL & "...".
              If ![sql1_aline] <> lngFirstLine Then
                lngLastLine = ![sql1_linenum]
                .Edit
                ![sql1_aline] = lngFirstLine
                ![sql1_IsUpd] = True
                ![sql1_datemodified] = Now()
                .Update
              End If
            End If
          End If
          If lngY < lngRecs Then .MoveNext
        Next
        .Close
      End With
    Next
    .Close
  End With

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  SQL_LineNums = blnRetValx

End Function

Public Function SQL_FirstLine(varInput As Variant) As String
' ** Return the first line of a multi-line SQL statement.
' ** Applies to zz_tbl_sql_code_04.
' ** Not currently called.

  Const THIS_PROC As String = "SQL_FirstLine"

  Dim intPos1 As Integer
  Dim strRetVal As String

  strRetVal = vbNullString

  If IsNull(varInput) = False Then
    intPos1 = InStr(varInput, vbCr)
    If intPos1 = 0 Then intPos1 = InStr(varInput, vbLf)
    If intPos1 > 0 Then
      strRetVal = Trim(Left$(varInput, (intPos1 - 1)))
    Else
      ' ** Evidently only 1 line.
      strRetVal = Trim(varInput)
    End If
  End If

  SQL_FirstLine = strRetVal

End Function

Public Function SQL_LineCnt(varInput As Variant) As Long
' ** Return the number of lines in a complete SQL statement.
' ** Applies to zz_tbl_sql_code_04.
' ** Not currently called.

  Const THIS_PROC As String = "SQL_LineCnt"

  Dim lngCnt_CRLF As Long
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp01 As String
  Dim lngRetVal As Long

  lngRetVal = 0&

  lngCnt_CRLF = 0&

  If IsNull(varInput) = False Then
    intPos1 = InStr(varInput, vbCr)
    If intPos1 = 0 Then intPos1 = InStr(varInput, vbLf)
    If intPos1 > 0 Then
      Do While intPos1 > 0
        lngCnt_CRLF = lngCnt_CRLF + 1&
        intPos2 = InStr((intPos1 + 1), varInput, vbCr)
        If intPos2 = 0 Then intPos2 = InStr((intPos1 + 1), varInput, vbLf)
        If intPos2 = 0 Then intPos1 = 0 Else intPos1 = intPos2
      Loop
      'lngCnt_CRLF = lngCnt_CRLF + 1&  ' ** Count the last line.
    End If
    lngRetVal = lngCnt_CRLF
  End If

  SQL_LineCnt = lngRetVal

End Function

Public Function VBA_Find_Proc() As Boolean
' ** Check module procedures...
' **
' ** Used with some others, not sure which.

  Const THIS_PROC As String = "VBA_Find_Proc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim rptao As AccessObject, prj As CurrentProject
  Dim frm As Access.Form, rpt As Access.Report
Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strLine As String, lngLineNum As Long
  Dim strModName As String, lngModLines As Long, lngModDecLines As Long
  Dim lngMods As Long, arr_varMod() As Variant
  Dim strProcName As String, lngProcLines As Long
  Dim lngProcs As Long, arr_varProc() As Variant
  Dim lngProcStart As Long, lngProcEnd As Long
  Dim blnProcStartFound As Boolean
  Dim strProcKind As String, lngProcKind As Long
  Dim arr_varWord() As Variant, lngDims As Long
  Dim lngSQLCodes As Long, arr_varSQLCode() As Variant
  Dim lngDocLines As Long, arr_varDocLine() As Variant
  Dim lngViews As Long, arr_varView() As Variant
  Dim lngErrs1 As Long, arr_varErr1() As Variant, lngErrs2 As Long, arr_varErr2() As Variant
  Dim blnCaseErrOn1 As Boolean, blnCaseErrOn2 As Boolean
  Dim blnNumbered As Boolean, blnHasScope As Boolean, blnIsExplicit As Boolean, blnSkip As Boolean
  Dim blnRemarkFound As Boolean, intRemPos As Integer
  Dim blnHasThisProc As Boolean, blnHasErrLbl As Boolean, blnHasExitLbl As Boolean
  Dim blnHasDispName As Boolean, blnHasRptCall As Boolean, lngDispNamesRpt As Long, lngDispNamesFrm As Long
  Dim lngNoThis As Long, lngNoErr As Long, lngNoExit As Long
  Dim intProcHasSQL As Integer, intLastProcHasSQL As Integer, intLastModHasSQL As Integer
  Dim blnFound As Boolean, blnPreviousWritten As Boolean
Dim lngParseSQLs As Long, arr_varParseSQL() As Variant
Dim lngVars As Long, arr_varVar() As Variant
Dim blnEOL As Boolean, lngSQLBegElem As Long, blnSQLChk As Boolean
Dim blnHasMsgBox As Boolean, blnMsgContinues As Boolean, lngMsgBoxLineNum As Long, lngMsgBoxTmp0 As Long, blnSkipIt As Boolean
Dim lngModSQLs As Long, arr_varModSQL() As Variant
Dim lngVBComID As Long, lngFrmID As Long
Dim lngElemA As Long
  Dim lngElemM As Long, lngElemP As Long, lngElems As Long, lngElemE1 As Long, lngElemE2 As Long
  Dim lngElemD1 As Long, lngElemD2 As Long, lngElemV1 As Long, lngElemV2 As Long, lngElemQ As Long
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer, intLen As Integer
  Dim lngTmp00 As Long, strTmp01 As String, strTmp02 As String, strTmp03 As String, lngTmp04 As Long
  Dim arr_varTmpArr05 As Variant, arr_varTmpArr06 As Variant, arr_varTmpArr07 As Variant
  Dim lngX As Long, lngY As Long, lngZ As Long, lngV As Long, lngW As Long

  ' ** Array: arr_varMod().
  Const MOD_ELEMS As Integer = 17  ' ** Array's first-element UBound().
  Const MOD_NAME     As Integer = 0
  Const MOD_TYPE     As Integer = 1
  Const MOD_OBJ      As Integer = 2
  Const MOD_PROCS    As Integer = 3
  Const MOD_THIS     As Integer = 4
  Const MOD_OPDAT    As Integer = 5
  Const MOD_OPEXP    As Integer = 6
  Const MOD_DISPNAME As Integer = 7
  Const MOD_DISPCNT  As Integer = 8
  Const MOD_RPTCALL  As Integer = 9
  Const MOD_TAG      As Integer = 10
  Const MOD_SUBFORM  As Integer = 11
  Const MOD_OBJ_OPEN As Integer = 12
  Const MOD_OOP_ELEM As Integer = 13
  Const MOD_HAS_SQL  As Integer = 14
  Const MOD_SQL_ARR  As Integer = 15
  Const MOD_PROC_ARR As Integer = 16
  Const MOD_MSQL_ARR As Integer = 17

  ' ** Array: arr_varProc().
  Const PROC_ELEMS As Integer = 16  ' ** Array's first-element UBound().
  Const PROC_NAME     As Integer = 0
  Const PROC_KIND     As Integer = 1
  Const PROC_KINDNAME As Integer = 2
  Const PROC_START    As Integer = 3
  Const PROC_END      As Integer = 4
  Const PROC_THIS     As Integer = 5
  Const PROC_ERR_LBL  As Integer = 6
  Const PROC_EXIT_LBL As Integer = 7
  Const PROC_SCOPE    As Integer = 8
  Const PROC_TYPED    As Integer = 9
  Const PROC_DIMS     As Integer = 10
  Const PROC_HAS_SQL  As Integer = 11
  Const PROC_MOD_ELEM As Integer = 12
  Const PROC_DOC_ARR  As Integer = 13
  Const PROC_VEW_ARR  As Integer = 14
  Const PROC_VAR_ARR  As Integer = 15
  Const PROC_PAR_ARR  As Integer = 16

  ' ** Array: arr_varWord().
  Const WORD_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const WRD_LINE As Integer = 0
  Const WRD_1 As Integer = 1
  Const WRD_2 As Integer = 2
  Const WRD_3 As Integer = 3
  Const WRD_4 As Integer = 4
  Const WRD_5 As Integer = 5
  Const WRD_6 As Integer = 6

  Const lngWords As Long = 6&

  ' ** Array: arr_varSQLCode().
  Const SQL_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const SQL_RESP      As Integer = 0
  Const SQL_LINE_NUM  As Integer = 1
  Const SQL_LINE_TXT  As Integer = 2
  Const SQL_DOCD      As Integer = 3
  Const SQL_PROC_ELEM As Integer = 4

  ' ** Array: arr_varDocLine().
  Const DOC_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const DOC_LINE_NUM As Integer = 0
  Const DOC_LINE_TXT As Integer = 1
  Const DOC_SHOWN    As Integer = 2

  ' ** Array: arr_varView().
  Const VEW_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const VEW_LINE_NUM As Integer = 0
  Const VEW_LINE_TXT As Integer = 1

  ' ** Array: arr_varParseSQL().
  Const PAR_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const PAR_VEW_ELEM As Integer = 0
  Const PAR_RAW_LINE As Integer = 1
  Const PAR_BEG_ELEM As Integer = 2
  Const PAR_ASSIGN   As Integer = 3
  Const PAR_EDIT1    As Integer = 4
  Const PAR_EDIT2    As Integer = 5

  ' ** Array: arr_varVar().
  Const VAR_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const VAR_NAME As Integer = 0

  ' ** Array: arr_varErr1().
  Const ERR_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const ERR_MOD_NAME  As Integer = 0
  Const ERR_PROC_NAME As Integer = 1
  Const ERR_LINE_NUM  As Integer = 2
  Const ERR_LINE_TXT  As Integer = 3
  Const ERR_EQUALS    As Integer = 4
  Const ERR_NUM       As Integer = 5
  Const ERR_REM       As Integer = 6

  blnRetValx = True
blnSQLChk = True
blnSQLDebug = False

  lngMods = 0&
  ReDim arr_varMod(MOD_ELEMS, 0)
  ' ********************************************************
  ' ** Array: arr_varMod()
  ' **
  ' **   Element  Description               Constant
  ' **   =======  ========================  ==============
  ' **      0     Module Name               MOD_NAME
  ' **      1     Module Type               MOD_TYPE
  ' **      2     Object Type               MOD_OBJ
  ' **      3     Number Of Procedures      MOD_PROCS
  ' **      4     Has THIS_NAME             MOD_THIS
  ' **      5     Option Compare            MOD_OPDAT
  ' **      6     Option Explicit           MOD_OPEXP
  ' **      7     Has Disp{Obj}Name()       MOD_DISPNAME
  ' **      8     Disp{Obj}Name() Count     MOD_DISPCNT
  ' **      9     Has RptCallFrm Line       MOD_RPTCALL
  ' **     10     Object's Tag              MOD_TAG
  ' **     11     Is Subform/Subreport      MOD_SUBFORM
  ' **     12     Has {Object}_Open() Proc  MOD_OBJ_OPEN
  ' **     13     {Object}_Open() Element   MOD_OOP_ELEM
  ' **     14     Has SQL Response          MOD_HAS_SQL
  ' **     15     arr_varSQLCode() Array    MOD_SQL_ARR
  ' **     16     arr_varProc() Array       MOD_PROC_ARR
  ' **
  ' ********************************************************

  lngProcs = 0&
  ReDim arr_varProc(PROC_ELEMS, 0)  ' ** ReDim after every module.
  ' **********************************************************
  ' ** Array: arr_varProc()
  ' **
  ' **   Element  Description                Constant
  ' **   =======  =========================  ===============
  ' **      0     Procedure Name             PROC_NAME
  ' **      1     Procedure Kind             PROC_KIND
  ' **      2     Procedure Kind Name        PROC_KINDNAME
  ' **      3     Start Line                 PROC_START
  ' **      4     End Line                   PROC_END
  ' **      5     Has THIS_PROC              PROC_THIS
  ' **      6     Has Error Label            PROC_ERR_LBL
  ' **      7     Has Exit Label             PROC_EXIT_LBL
  ' **      8     Has Scope                  PROC_SCOPE
  ' **      9     Explicit Type              PROC_TYPED
  ' **     10     Count of Dim's             PROC_DIMS
  ' **     11     VBA_IsSQL() Response       PROC_HAS_SQL
  ' **     12     arr_varMod() Element       PROC_MOD_ELEM
  ' **     13     arr_varDocLine() Array     PROC_DOC_ARR
  ' **     14     arr_varView() Array        PROC_VEW_ARR
  ' **     15     arr_varVar() Array         PROC_VAR_ARR
  ' **     16     arr_varParseSQL() Array    PROC_PAR_ARR
  ' **
  ' **********************************************************

  lngErrs1 = 0&
  ReDim arr_varErr1(ERR_ELEMS, 0)  ' ** Once for entire run of all modules.
  ' *********************************************************
  ' ** Array: arr_varErr1()
  ' **
  ' **   Element  Description               Constant
  ' **   =======  ========================  ===============
  ' **      0     Module Name               ERR_MOD_NAME
  ' **      1     Procedure Name            ERR_PROC_NAME
  ' **      2     Line Number               ERR_LINE_NUM
  ' **      3     Line                      ERR_LINE_TXT
  ' **
  ' *********************************************************

  If blnSQLChk = True Then
    Set dbs = CurrentDb
    With dbs
      ' ** Update zz_tbl_sql_code_01, set sql1_Chk = False.
      Set qdf = .QueryDefs("zz_qry_SQL_Code_05")
      qdf.Execute
      .Close
    End With
  End If

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc

        strModName = .Name
        If Left$(strModName, 2) <> "z_" And Left$(strModName, 2) <> "zz_" And _
            Right$(strModName, 4) <> "_bak" And Right$(strModName, 4) <> "_old" Then  ' ** Skip these.

' ** Limit to specified object(s).
'If strModName = "Form_Statement Parameters" Then
If Left$(strModName, 5) = "Form_" Then
'If left$(strModName, 7) = "Report_" Then
'If left$(strModName, 5) <> "Form_" And left$(strModName, 7) <> "Report_" Then

          lngProcs = 0&
          ReDim arr_varProc(PROC_ELEMS, 0)  ' ** ReDim after every module.
          ' **********************************************************
          ' ** Array: arr_varProc()
          ' **
          ' **   Element  Description                Constant
          ' **   =======  =========================  ===============
          ' **      0     Procedure Name             PROC_NAME
          ' **      1     Procedure Kind             PROC_KIND
          ' **      2     Procedure Kind Name        PROC_KINDNAME
          ' **      3     Start Line                 PROC_START
          ' **      4     End Line                   PROC_END
          ' **      5     Has THIS_PROC              PROC_THIS
          ' **      6     Has Error Label            PROC_ERR_LBL
          ' **      7     Has Exit Label             PROC_EXIT_LBL
          ' **      8     Has Scope                  PROC_SCOPE
          ' **      9     Explicit Type              PROC_TYPED
          ' **     10     Count of Dim's             PROC_DIMS
          ' **     11     VBA_IsSQL() Response       PROC_HAS_SQL
          ' **     12     arr_varMod() Element       PROC_MOD_ELEM
          ' **     13     arr_varDocLine() Array     PROC_DOC_ARR
          ' **     14     arr_varView() Array        PROC_VEW_ARR
          ' **     15     arr_varVar() Array         PROC_VAR_ARR
          ' **     16     arr_varParseSQL() Array    PROC_PAR_ARR
          ' **
          ' **********************************************************

          ' ** Only 1 record.
          ReDim arr_varWord(WORD_ELEMS, 0)  ' ** ReDim after every procedure.
          ' ********************************************************
          ' ** Array: arr_varWord()
          ' **
          ' **   Element  Description                   Constant
          ' **   =======  ============================  ==========
          ' **      0     Procedure Declaration Line    WRD_LINE
          ' **      1     1st Word                      WRD_1
          ' **      2     2nd Word                      WRD_2
          ' **      3     3rd Word                      WRD_3
          ' **      4     4th Word                      WRD_4
          ' **      5     5th Word                      WRD_5
          ' **      6     6th Word                      WRD_6
          ' **
          ' ********************************************************

          lngSQLCodes = 0&
          ReDim arr_varSQLCode(SQL_ELEMS, 0)
          ' ********************************************************
          ' ** Array: arr_varSQLCode()
          ' **
          ' **   Element  Description              Constant
          ' **   =======  =======================  ===============
          ' **      0     VBA_IsSQL() Response     SQL_RESP
          ' **      1     Line Number              SQL_LINE_NUM
          ' **      2     Line with SQL Code       SQL_LINE_TXT
          ' **      3     Actual SQL Line T/F      SQL_DOCD
          ' **      4     arr_varProc() Element    SQL_PROC_ELEM
          ' **
          ' ********************************************************

          lngModSQLs = 0&
          ReDim arr_varModSQL(MSQL_ELEMS, 0)
          ' ********************************************************
          ' ** Array: arr_varModSQL()
          ' **
          ' **   Element  Description              Constant
          ' **   =======  =======================  ===============
          ' **      0     VBA_IsSQL() Response     MSQL_RESP
          ' **      1     Line Number              MSQL_LINE_TXT
          ' **      2     Line with SQL Code       MSQL_LINE_NUM
          ' **      3     Actual SQL Line T/F      MSQL_DOCD
          ' **      4     arr_varMod() Element     MSQL_MOD_ELEM
          ' **
          ' ********************************************************

          lngMods = lngMods + 1&
          lngElemM = lngMods - 1&
          ReDim Preserve arr_varMod(MOD_ELEMS, lngElemM)
          arr_varMod(MOD_NAME, lngElemM) = strModName
          arr_varMod(MOD_TYPE, lngElemM) = .Type
          ' **   vbext_ComponentType enumeration:
          ' **       1  vbext_ct_StdModule        Standard Module
          ' **       2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
          ' **       3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm in the VBA Editor.
          ' **      11  vbext_ct_ActiveXDesigner
          ' **     100  vbext_ct_Document         Module behind Form, Report, or Excel Worksheet.
          If Left$(.Name, 5) = "Form_" Then
            arr_varMod(MOD_OBJ, lngElemM) = acForm
          ElseIf Left$(.Name, 7) = "Report_" Then
            arr_varMod(MOD_OBJ, lngElemM) = acReport
          Else
            arr_varMod(MOD_OBJ, lngElemM) = acModule
          End If
          ' ** AcObjectType enumeration:
          ' **  -1  acDefault
          ' **   0  acTable            Table
          ' **   1  acQuery            Query
          ' **   2  acForm             Form
          ' **   3  acReport           Report
          ' **   4  acMacro            Macro
          ' **   5  acModule           Module
          ' **   7  acServerView       Server View
          ' **   8  acDiagram          Database Diagram
          ' **   9  acStoredProcedure  Stored Procedure
          ' **  10  acFunction         Function
          arr_varMod(MOD_PROCS, lngElemM) = 0&
          arr_varMod(MOD_THIS, lngElemM) = CBool(False)
          arr_varMod(MOD_OPDAT, lngElemM) = CBool(False)
          arr_varMod(MOD_OPEXP, lngElemM) = CBool(False)
          arr_varMod(MOD_DISPNAME, lngElemM) = CBool(False)
          arr_varMod(MOD_DISPCNT, lngElemM) = CLng(0)
          arr_varMod(MOD_TAG, lngElemM) = vbNullString
          arr_varMod(MOD_SUBFORM, lngElemM) = CBool(False)
          arr_varMod(MOD_OBJ_OPEN, lngElemM) = CBool(False)
          arr_varMod(MOD_OOP_ELEM, lngElemM) = CLng(-1)
          arr_varMod(MOD_HAS_SQL, lngElemM) = CInt(0)

          Set cod = .CodeModule
          With cod

            lngModLines = .CountOfLines
            lngModDecLines = .CountOfDeclarationLines
            strProcName = vbNullString
            lngProcLines = 0&
            lngProcStart = 0&: lngProcEnd = 0&
            blnProcStartFound = False
            strProcKind = vbNullString: lngProcKind = -1
            blnHasScope = False: blnIsExplicit = False
            blnHasThisProc = False: blnHasErrLbl = False: blnHasExitLbl = False
            blnRemarkFound = False: blnHasDispName = False: blnHasRptCall = False: lngDispNamesRpt = 0&: lngDispNamesFrm = 0&
            lngDims = 0&
            blnCaseErrOn1 = False: blnCaseErrOn2 = False
            intLastModHasSQL = 0: intProcHasSQL = 0: intLastProcHasSQL = 0
            blnPreviousWritten = False
            blnModSqlVars = False

            For lngX = 1& To lngModLines

              lngLineNum = lngX
              blnNumbered = False
              strLine = vbNullString

              If lngLineNum <= lngModDecLines Then
                ' ** Declaration section.
                strLine = Trim$(.Lines(lngLineNum, 1))
                If strLine <> vbNullString Then
                  If Left$(strLine, 1) <> "'" Then
                    If InStr(strLine, "THIS_NAME") > 0 Then
                      arr_varMod(MOD_THIS, lngElemM) = CBool(True)
                    ElseIf InStr(strLine, "Option Compare Database") > 0 Then
                      arr_varMod(MOD_OPDAT, lngElemM) = CBool(True)
                    ElseIf InStr(strLine, "Option Explicit") > 0 Then
                      arr_varMod(MOD_OPEXP, lngElemM) = CBool(True)
                    Else
                      ' ** Check for module-level SQL variables.
                      If blnSQLChk = True Then
                        intProcHasSQL = VBA_IsSQL(cod, lngLineNum)  ' ** Function: Below.
                        If intProcHasSQL > SQL_NONE Then
                          blnModSqlVars = True
                          lngModSQLs = lngModSQLs + 1&
                          lngElemA = lngModSQLs - 1&
                          ReDim Preserve arr_varModSQL(MSQL_ELEMS, lngElemA)
                          ' ********************************************************
                          ' ** Array: arr_varModSQL()
                          ' **
                          ' **   Element  Description              Constant
                          ' **   =======  =======================  ===============
                          ' **      0     VBA_IsSQL() Response     MSQL_RESP
                          ' **      1     Line Number              MSQL_LINE_TXT
                          ' **      2     Line with SQL Code       MSQL_LINE_NUM
                          ' **      3     Actual SQL Line T/F      MSQL_DOCD
                          ' **      4     arr_varMod() Element     MSQL_MOD_ELEM
                          ' **
                          ' ********************************************************
                          arr_varModSQL(MSQL_RESP, lngElemA) = intProcHasSQL
                          arr_varModSQL(MSQL_LINE_TXT, lngElemA) = strLine
                          arr_varModSQL(MSQL_LINE_NUM, lngElemA) = lngLineNum
                          arr_varModSQL(MSQL_DOCD, lngElemA) = CBool(False)
                          arr_varModSQL(MSQL_MOD_ELEM, lngElemA) = lngElemM
                        End If
                      End If
                    End If
                  End If  ' ** Not a remark.
                End If  ' ** Not a blank line.
              Else
                ' ** Procedure section.

                If .ProcOfLine(lngLineNum, vbext_pk_Proc) <> vbNullString Then
                  ' ** Returns name of procedure that the specified line is in.
                  ' ** Doesn't care if type of procedure is incorrect.

                  If .ProcOfLine(lngLineNum, vbext_pk_Proc) <> strProcName Then
                    ' ** A new procedure.
                    ' ** It'll continue to hit within this section until
                    ' ** strProcName gets set to the new procedure.

                    If strProcName = vbNullString Then blnPreviousWritten = True

                    If blnPreviousWritten = False Then
                      ' ** The very first time this is hit, strProcName hasn't been set to the
                      ' ** new procedure, and blnProcStartFound = True from the previous one.

                      ' ** Record gathered info from the previous procedure.
                      arr_varProc(PROC_THIS, lngElemP) = blnHasThisProc
                      arr_varProc(PROC_ERR_LBL, lngElemP) = blnHasErrLbl
                      arr_varProc(PROC_EXIT_LBL, lngElemP) = blnHasExitLbl
                      arr_varProc(PROC_DIMS, lngElemP) = lngDims
                      arr_varProc(PROC_HAS_SQL, lngElemP) = intLastProcHasSQL
                      If Not (intLastProcHasSQL And intLastModHasSQL) Then
                        intLastModHasSQL = intLastModHasSQL + intLastProcHasSQL
                      End If

                      ' ** Reset everything for this new procedure.
                      ReDim arr_varWord(WORD_ELEMS, 0)
                      For lngY = 0& To lngWords: arr_varWord(lngY, 0) = vbNullString: Next
                      strProcKind = vbNullString: lngProcKind = -1
                      blnProcStartFound = False
                      blnHasScope = False: blnIsExplicit = False
                      lngProcStart = 0&: lngProcLines = 0&: lngProcEnd = 0&
                      blnHasThisProc = False: blnHasErrLbl = False: blnHasExitLbl = False
                      blnRemarkFound = False
                      intProcHasSQL = 0: intLastProcHasSQL = 0
                      lngDims = 0&
                      blnCaseErrOn1 = False: blnCaseErrOn2 = False
                      blnPreviousWritten = True
                      ' ** From here on out, while we await blnProcStartFound = True,
                      ' ** this section won't be hit.

                    Else
                      ' ** OK, now the previous info's been saved, and blnProcStartFound = False.
                      ' ** At this point, strProcName is still the previous procedure.

                      If blnProcStartFound = True Then
                        ' ** We don't want this to be hit until blnProcStartFound = True.
                        strProcName = .ProcOfLine(lngLineNum, vbext_pk_Proc)
                        lngProcs = lngProcs + 1&
                        lngElemP = lngProcs - 1&
                        ReDim Preserve arr_varProc(PROC_ELEMS, lngElemP)
                        arr_varProc(PROC_NAME, lngElemP) = strProcName
                        arr_varProc(PROC_KIND, lngElemP) = lngProcKind
                        arr_varProc(PROC_KINDNAME, lngElemP) = strProcKind
                        arr_varProc(PROC_START, lngElemP) = lngProcStart
                        arr_varProc(PROC_END, lngElemP) = lngProcEnd
                        arr_varProc(PROC_SCOPE, lngElemP) = blnHasScope
                        arr_varProc(PROC_TYPED, lngElemP) = blnIsExplicit
                        arr_varProc(PROC_MOD_ELEM, lngElemP) = lngElemM
                        blnPreviousWritten = False
                      Else
                        ' ** It'll just pass through here until it finds the procedure start.
                      End If

                    End If  ' ** blnPreviousWritten.

                  End If  ' ** New procedure.

                  strLine = Trim$(.Lines(lngLineNum, 1))

                  If strLine <> vbNullString Then
                    If blnProcStartFound = True Then
                      ' ** We're within a procedure.
                      If Left$(strLine, 1) <> "'" Then
                        ' ** Not a remark.

'SEARCH HERE!

                        ' ** Look for SQL.
                        If blnSQLChk = True Then
                          intProcHasSQL = VBA_IsSQL(cod, lngLineNum)  ' ** Function: Below.
                          If intProcHasSQL > SQL_NONE Then
                            If (intProcHasSQL And intLastProcHasSQL) = 0 Then
                              If (intProcHasSQL And SQL_HASCMD) <> 0 And (intLastProcHasSQL And SQL_HASCMD) = 0 Then
                                intLastProcHasSQL = intLastProcHasSQL + SQL_HASCMD
                              End If
                              If (intProcHasSQL And SQL_HASVRB) <> 0 And (intLastProcHasSQL And SQL_HASVRB) = 0 Then
                                intLastProcHasSQL = intLastProcHasSQL + SQL_HASVRB
                              End If
                              If (intProcHasSQL And SQL_HASTRM) <> 0 And (intLastProcHasSQL And SQL_HASTRM) = 0 Then
                                intLastProcHasSQL = intLastProcHasSQL + SQL_HASTRM
                              End If
                            End If
                            lngSQLCodes = lngSQLCodes + 1&
                            lngElems = lngSQLCodes - 1&
                            ReDim Preserve arr_varSQLCode(SQL_ELEMS, lngElems)
                            ' ********************************************************
                            ' ** Array: arr_varSQLCode()
                            ' **
                            ' **   Element  Description              Constant
                            ' **   =======  =======================  ===============
                            ' **      0     VBA_IsSQL() Response     SQL_RESP
                            ' **      1     Line Number              SQL_LINE_NUM
                            ' **      2     Line with SQL Code       SQL_LINE_TXT
                            ' **      3     Actual SQL Line T/F      SQL_DOCD
                            ' **      4     arr_varProc() Element    SQL_PROC_ELEM
                            ' **
                            ' ********************************************************
                            arr_varSQLCode(SQL_RESP, lngElems) = intProcHasSQL
                            arr_varSQLCode(SQL_LINE_TXT, lngElems) = strLine
                            arr_varSQLCode(SQL_LINE_NUM, lngElems) = lngLineNum
                            arr_varSQLCode(SQL_DOCD, lngElems) = CBool(False)
                            arr_varSQLCode(SQL_PROC_ELEM, lngElems) = lngElemP
                          End If
                        End If

                        ' ** Look for my DispFrmName() and DispRptName() functions.
                        If InStr(strLine, "DispFrmName") > 0 Or InStr(strLine, "DispRptName") > 0 Then
                          blnHasDispName = CBool(True)
                          If InStr(strLine, "DispFrmName") > 0 Then
                            lngDispNamesFrm = lngDispNamesFrm + 1&
                          Else
                            lngDispNamesRpt = lngDispNamesRpt + 1&
                          End If
                        End If

                        ' ** Count the Dim, Static, and Const statements.
                        If Left$(strLine, 4) = "Dim " Or Left$(strLine, 7) = "Static " Or _
                           Left$(strLine, 6) = "Const " Then
                          lngDims = lngDims + 1&
                        End If

                        '' ** Look for gstrReportCallingForm handling.
                        'If InStr(strLine, "If gstrReportCallingForm <> vbNullString Then") > 0 Then
                        '  blnHasRptCall = True
                        'End If

                        '' ** Look for documented errors.
                        'If InStr(strLine, " Err = ") > 0 Or InStr(strLine, " Err.Number = ") > 0 Then
                        '  lngErrs1 = lngErrs1 + 1&
                        '  lngElemE1 = lngErrs1 - 1&
                        '  ReDim Preserve arr_varErr1(ERR_ELEMS, lngElemE1)
                        '  arr_varErr1(ERR_MOD_NAME, lngElemE1) = strModName
                        '  arr_varErr1(ERR_PROC_NAME, lngElemE1) = strProcName
                        '  arr_varErr1(ERR_LINE_NUM, lngElemE1) = lngLineNum
                        '  arr_varErr1(ERR_LINE_TXT, lngElemE1) = strLine
                        '  arr_varErr1(ERR_EQUALS, lngElemE1) = CBool(True)
                        '  arr_varErr1(ERR_NUM, lngElemE1) = CLng(0)
                        '  arr_varErr1(ERR_REM, lngElemE1) = vbNullString
                        'ElseIf InStr(strLine, " Select Case Err") > 0 Or InStr(strLine, " Select Case DataErr") > 0 Then
                        '  lngErrs1 = lngErrs1 + 1&
                        '  lngElemE1 = lngErrs1 - 1&
                        '  ReDim Preserve arr_varErr1(ERR_ELEMS, lngElemE1)
                        '  arr_varErr1(ERR_MOD_NAME, lngElemE1) = strModName
                        '  arr_varErr1(ERR_PROC_NAME, lngElemE1) = strProcName
                        '  arr_varErr1(ERR_LINE_NUM, lngElemE1) = lngLineNum
                        '  arr_varErr1(ERR_LINE_TXT, lngElemE1) = strLine
                        '  arr_varErr1(ERR_EQUALS, lngElemE1) = CBool(False)
                        '  arr_varErr1(ERR_NUM, lngElemE1) = CLng(0)
                        '  arr_varErr1(ERR_REM, lngElemE1) = vbNullString
                        '  If blnCaseErrOn1 = True Then blnCaseErrOn2 = True Else blnCaseErrOn1 = True
                        'ElseIf blnCaseErrOn1 = True Then
                        '  intPos1 = InStr(strLine, "Case ")
                        '  If intPos1 > 0 Then
                        '    strTmp02 = vbNullString
                        '    strTmp01 = Mid$(strLine, (intPos1 + 5))
                        '    intPos1 = InStr(strTmp01, " ")
                        '    If intPos1 > 0 Then
                        '      strTmp02 = Trim$(Mid$(strTmp01, intPos1))
                        '      strTmp01 = Trim$(Left$(strTmp01, intPos1))
                        '      If Right$(strTmp01, 1) = "," Then
                        '        strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
                        '      Else
                        '        intPos1 = InStr(strTmp02, "'")
                        '        If intPos1 > 0 Then
                        '          strTmp02 = Trim$(Mid$(strTmp02, intPos1))
                        '        Else
                        '          strTmp02 = vbNullString
                        '        End If
                        '      End If
                        '    End If
                        '    If strTmp01 <> "Else" Then
                        '      If IsNumeric(strTmp01) = True Then
                        '        lngErrs1 = lngErrs1 + 1&
                        '        lngElemE1 = lngErrs1 - 1&
                        '        ReDim Preserve arr_varErr1(ERR_ELEMS, lngElemE1)
                        '        arr_varErr1(ERR_MOD_NAME, lngElemE1) = strModName
                        '        arr_varErr1(ERR_PROC_NAME, lngElemE1) = strProcName
                        '        arr_varErr1(ERR_LINE_NUM, lngElemE1) = lngLineNum
                        '        arr_varErr1(ERR_LINE_TXT, lngElemE1) = strLine
                        '        arr_varErr1(ERR_EQUALS, lngElemE1) = CBool(False)
                        '        arr_varErr1(ERR_NUM, lngElemE1) = CLng(strTmp01)
                        '        arr_varErr1(ERR_REM, lngElemE1) = strTmp02
                        '      End If
                        '    End If
                        '  End If
                        '  If blnCaseErrOn2 = True Then
                        '    If InStr(strLine, " End Select") > 0 Then
                        '      blnCaseErrOn2 = False
                        '    End If
                        '  Else
                        '    If InStr(strLine, " End Select") > 0 Then
                        '      blnCaseErrOn1 = False
                        '    End If
                        '  End If
                        'End If

                        ' ** Look for THIS_PROC.
                        If InStr(strLine, "Const THIS_PROC as String") Then
                          ' ** Has THIS_PROC.
                          blnHasThisProc = True
                        Else
                          ' ** No THIS_PROC.
                          intPos1 = InStr(strLine, " ")
                          If intPos1 > 0 Then
                            strTmp01 = Left$(strLine, (intPos1 - 1))
                          Else
                            strTmp01 = strLine
                          End If
                          ' ** Look for error handler and exit labels.
                          If Right$(strTmp01, 1) = ":" Then
                            ' ** Is a label.
                            If InStr(strTmp01, "err") > 0 Or strTmp01 = "eh:" Or strTmp01 = "ER:" Or _
                               strTmp01 = "HANDBASKET:" Then  ' ** Common to this programmer.
                              ' ** Has error handler label.
                              blnHasErrLbl = True
                            ElseIf InStr(strTmp01, "exit") > 0 Or strTmp01 = "EX:" Then  ' ** Common to this programmer.
                              ' ** Has exit label.
                              blnHasExitLbl = True
                            End If
                          End If
                        End If

                      End If  ' ** Not a remark.
                    Else
                      ' ** Find the procedure start line.
                      If Left$(strLine, 1) <> "'" Then
                        ' ** It's not a remark.
                        intPos2 = InStr(strLine, " ")
                        If intPos2 > 0 Then
                          ' ** It's got a space, as a procedure declaration line will.
                          intPos1 = 1&
                          ' ** Collect the first 6 words in the line.
                          For lngY = 1& To lngWords
                            If arr_varWord(lngY, 0) = vbNullString Then
                              If lngY = 1& Then
                                arr_varWord(WRD_LINE, 0) = strLine
                              End If
                              arr_varWord(lngY, 0) = Trim$(Mid$(strLine, intPos1, (intPos2 - intPos1)))
                              If Left$(arr_varWord(lngY, 0), 1) = "'" Then
                                ' ** Remark encountered.
                                blnRemarkFound = True
                                intRemPos = intPos1 + IIf(lngY = 1&, 0, 1)
                                arr_varWord(lngY, 0) = vbNullString
                                Exit For
                              Else
                                If InStr(arr_varWord(lngY, 0), "(") > 0 Then
                                  arr_varWord(lngY, 0) = _
                                    Left$(arr_varWord(lngY, 0), (InStr(arr_varWord(lngY, 0), "(") - 1))
                                End If
                                intPos1 = intPos2
                              End If
                            End If
                            intPos2 = InStr((intPos2 + 1), strLine, " ")
                            If intPos2 = 0 Then
                              Exit For
                            End If
                          Next
                        End If  ' ** Has at least 1 space.
                      Else
                        ' ** It is a remark.
                        blnRemarkFound = True
                      End If  ' ** Is or isn't a remark.
                      ' ** Now pick up last word.
                      If blnRemarkFound = False And intPos1 > 0 Then
                        For lngY = 1& To lngWords
                          If arr_varWord(lngY, 0) = vbNullString Then
                            arr_varWord(lngY, 0) = Trim$(Mid$(strLine, intPos1))
                            If InStr(arr_varWord(lngY, 0), "(") > 0 Then
                              arr_varWord(lngY, 0) = _
                                Left$(arr_varWord(lngY, 0), (InStr(arr_varWord(lngY, 0), "(") - 1))
                            End If
                            Exit For
                          End If
                        Next
                      End If
                      If arr_varWord(WRD_2, 0) <> vbNullString Then
                        ' ** Procedures declaration will have at least 1 space.
                        blnHasScope = False: blnIsExplicit = False
                        For lngY = 1& To lngWords
                          strTmp01 = arr_varWord(lngY, 0)
                          If lngY < lngWords Then
                            strTmp02 = arr_varWord((lngY + 1&), 0)
                          Else
                            ' ** Should have already found the procedure name and quit the loop!
                            strTmp02 = "ERROR!"
                          End If
                          Select Case strTmp01
                          Case "Public", "Private", "Friend", "Static", "Global"
                            ' ** 1st word.
                            ' ** Scope.
                            blnHasScope = True
                          Case "Sub"
                            ' ** Usually 2nd word.
                            ' ** Procedure with no scope.
                            ' ** vbext_ProcKind enumeration:
                            ' **   vbext_pk_Proc  0 (both Sub and Function)
                            ' **   vbext_pk_Let   1
                            ' **   vbext_pk_Set   2
                            ' **   vbext_pk_Get   3
                            lngProcKind = vbext_pk_Proc
                            strProcKind = strTmp01
                            strTmp03 = "Sub"
                          Case "Function"
                            ' ** Usually 2nd word.
                            ' ** Function with no scope.
                            lngProcKind = vbext_pk_Proc
                            strProcKind = strTmp01
                          Case "Property"
                            ' ** Usually 2nd word.
                            ' ** Property with no scope.
                            ' ** strTmp02 is next word after "Property".
                            Select Case strTmp02
                            Case "Let"
                              lngProcKind = vbext_pk_Let
                              strProcKind = strTmp01
                            Case "Set"
                              lngProcKind = vbext_pk_Set
                              strProcKind = strTmp01
                            Case "Get"
                              lngProcKind = vbext_pk_Get
                              strProcKind = strTmp01
                            Case Else
                              Beep
                              blnRetValx = False
                              Debug.Print "'UNKNOWN PROPERTY TYPE! : " & lngLineNum & " " & strTmp01 & " " & strTmp02
                            End Select
                          Case "Let", "Set", "Get"
                            ' ** Usually 3rd Word.
                            If lngProcKind = -1& Then
                              Debug.Print "'WHY KIND NOT FOUND ON PREVIOUS LOOP? : " & lngY & _
                                " LINE: " & lngLineNum & " '" & strTmp01 & "'"
                            End If
                          Case Else
                            If strTmp01 = .ProcOfLine(lngLineNum, vbext_pk_Proc) Then
                              ' ** Usually 4th word.
                              If lngY < lngWords Then
                                For lngZ = (lngY + 1&) To lngWords
                                  arr_varWord(lngZ, 0) = vbNullString
                                Next
                              End If
                              If lngProcKind <> -1& Then
                                blnProcStartFound = True
                                If strProcKind = "Function" Or strProcKind = "Get" Then
                                  intPos1 = InStr(strLine, ")")
                                  If intPos1 > 0 Then
                                    intPos3 = InStr(strLine, "'")
                                    intPos2 = InStr((intPos1 + 1), strLine, ")")
                                    Do While intPos2 > 0
                                      If intPos3 = 0 Or (intPos3 > 0 And intPos2 < intPos3) Then
                                        intPos1 = intPos2
                                      ElseIf (intPos3 > 0 And intPos2 > intPos3) Then
                                        Exit Do
                                      End If
                                      intPos2 = InStr((intPos1 + 1), strLine, ")")
                                    Loop
                                    If (intPos3 = 0 And intPos1 < Len(strLine)) Or _
                                       (intPos3 > 0 And intPos1 < intPos3) Then
                                      ' ** Procedure explicitly typed.
                                      If Mid$(strLine, (intPos1 + 2), 2) = "As" Then
                                        blnIsExplicit = True
                                      Else
                                        intPos2 = InStr(strLine, "_")
                                        If (intPos2 > 0 And intPos3 = 0) Or (intPos2 > 0 And intPos2 < intPos3) Then
                                          Debug.Print "'MULTI-LINE: " & strModName & " " & _
                                            .ProcOfLine(lngLineNum, vbext_pk_Proc) & "()"
                                        ElseIf (intPos2 > 0 And intPos2 > intPos3) Then
                                          ' ** Procedure not explicitly typed.
                                        Else
                                          Debug.Print "'SQL WHAT? " & strModName & " " & _
                                            .ProcOfLine(lngLineNum, vbext_pk_Proc) & "() '" & strLine & "'"
                                        End If
                                      End If
                                    Else
                                      ' ** Procedure not explicitly typed.
                                    End If
                                  Else
                                    Debug.Print "'NO )! " & strModName & " " & _
                                      .ProcOfLine(lngLineNum, vbext_pk_Proc) & "() '" & strLine & "'"
                                  End If
                                End If
                                lngProcStart = lngLineNum
                                lngProcLines = .ProcCountLines(.ProcOfLine(lngLineNum, vbext_pk_Proc), lngProcKind)
                                lngZ = (lngProcStart + lngProcLines)
                                strTmp02 = "End " & strProcKind
                                strTmp03 = Trim$(.Lines(lngZ, 1))
                                intPos1 = InStr(strTmp03, strTmp02)
                                Do While intPos1 = 0
                                  lngZ = lngZ - 1&
                                  intPos1 = InStr(.Lines(lngZ, 1), strTmp02)
                                  If lngZ = lngLineNum Then
                                    lngZ = 0&
                                    Exit Do  ' ** Exit the loop!
                                  End If
                                Loop
                                If lngZ > 0& Then
                                  lngProcEnd = lngZ
                                Else
                                  Debug.Print "'STOP : PROC_END NOT FOUND! " & _
                                    .ProcOfLine(lngLineNum, vbext_pk_Proc) & " " & lngLineNum
                                End If
                              End If
                              Exit For
                            Else
                              ' ** Unknown 1st word; not a procedure declaration line.
                              'Beep
                              blnRetValx = False
                              Debug.Print "'UNKNOWN FIRST WORD! : lngLineNum = " & lngLineNum & ", lngY = " & lngY & _
                                " '" & arr_varWord(0, 0) & " " & arr_varWord(1, 0) & " " & arr_varWord(2, 0) & " " & _
                                arr_varWord(3, 0) & " " & arr_varWord(4, 0) & "'"
                            End If
                            Exit For
                          End Select
                        Next  ' ** For each word: lngY.
                      End If  ' ** Possible declaration line: arr_varWord(WRD_2, 0).

                    End If    ' ** blnProcStartFound.

                  End If    ' ** Not a blank line.

                End If  ' ** Is procedure.

              End If    ' ** Is declaraction or procedure.

            Next      ' ** For each code line: lngLineNum.

            ' ** Save the last procedure's info.
            If blnProcStartFound = True Then
              ' ** Record info from the previous procedure.
              arr_varProc(PROC_THIS, lngElemP) = blnHasThisProc
              arr_varProc(PROC_ERR_LBL, lngElemP) = blnHasErrLbl
              arr_varProc(PROC_EXIT_LBL, lngElemP) = blnHasExitLbl
              arr_varProc(PROC_DIMS, lngElemP) = lngDims
              If blnSQLChk = True Then
                arr_varProc(PROC_HAS_SQL, lngElemP) = intLastProcHasSQL
                If Not (intLastProcHasSQL And intLastModHasSQL) Then
                  intLastModHasSQL = intLastModHasSQL + intLastProcHasSQL
                End If
              End If
            End If

            If blnHasDispName = True Then
              arr_varMod(MOD_DISPNAME, lngElemM) = CBool(True)
              If lngDispNamesFrm > 0& Then
                arr_varMod(MOD_DISPCNT, lngElemM) = lngDispNamesFrm
              Else
                arr_varMod(MOD_DISPCNT, lngElemM) = lngDispNamesRpt
              End If
            End If

            If blnHasRptCall = True Then
              arr_varMod(MOD_RPTCALL, lngElemM) = CBool(True)
            End If

          End With  ' ** This CodeModule: cod.

          If lngProcs > 0& Then
            arr_varMod(MOD_PROCS, lngElemM) = lngProcs
            arr_varMod(MOD_PROC_ARR, lngElemM) = arr_varProc
            For lngY = 0& To (lngProcs - 1&)
              lngElemP = lngY
              If (arr_varMod(MOD_OBJ, lngElemM) = acForm And arr_varProc(PROC_NAME, lngElemP) = "Form_Open") Or _
                 (arr_varMod(MOD_OBJ, lngElemM) = acReport And arr_varProc(PROC_NAME, lngElemP) = "Report_Open") Then
                arr_varMod(MOD_OBJ_OPEN, lngElemM) = CBool(True)
                arr_varMod(MOD_OOP_ELEM, lngElemM) = lngElemP
                Exit For
              End If
            Next
            arr_varMod(MOD_HAS_SQL, lngElemM) = intLastModHasSQL
            If (intLastModHasSQL And SQL_NONE) = 0 Then
              arr_varMod(MOD_SQL_ARR, lngElemM) = arr_varSQLCode
            End If
            If blnModSqlVars = True Then
              arr_varMod(MOD_MSQL_ARR, lngElemM) = arr_varModSQL
            End If
          Else
            arr_varMod(MOD_PROCS, lngElemM) = 0&
          End If

End If  ' ** Limit to specified object(s).

        End If  ' ** Specified module.

      End With  ' ** This VBComponent: vbc.
    Next      ' ** For each VBComponent: vbc.
  End With  ' ** This ActiveProject: vbp.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  ' ************************************************************
  ' ** Report on what was found.
  ' ************************************************************

  ' **********************************************************
  ' ** Array: arr_varProc()
  ' **
  ' **   Element  Description                Constant
  ' **   =======  =========================  ===============
  ' **      0     Procedure Name             PROC_NAME
  ' **      1     Procedure Kind             PROC_KIND
  ' **      2     Procedure Kind Name        PROC_KINDNAME
  ' **      3     Start Line                 PROC_START
  ' **      4     End Line                   PROC_END
  ' **      5     Has THIS_PROC              PROC_THIS
  ' **      6     Has Error Label            PROC_ERR_LBL
  ' **      7     Has Exit Label             PROC_EXIT_LBL
  ' **      8     Has Scope                  PROC_SCOPE
  ' **      9     Explicit Type              PROC_TYPED
  ' **     10     Count of Dim's             PROC_DIMS
  ' **     11     VBA_IsSQL() Response       PROC_HAS_SQL
  ' **     12     arr_varMod() Element       PROC_MOD_ELEM
  ' **     13     arr_varDocLine() Array     PROC_DOC_ARR
  ' **     14     arr_varView() Array        PROC_VEW_ARR
  ' **     15     arr_varVar() Array         PROC_VAR_ARR
  ' **     16     arr_varParseSQL() Array    PROC_PAR_ARR
  ' **
  ' **********************************************************

  ' ** Load arr_varEvent() array.
  'VBA_Event_Load  ' ** Module Function: zz_mod_MDEPrepFuncs.

  'For lngX = 0& To (lngMods - 1&)
  '  lngElemM = lngX
  '  strModName = arr_varMod(MOD_NAME, lngElemM)
  '  arr_varTmpArr05 = arr_varMod(MOD_PROC_ARR, lngElemM)
  '  lngProcs = UBound(arr_varTmpArr05, 2) + 1&
  '  For lngY = 0& To (lngProcs - 1&)
  '    lngElemP = lngY
  '    If arr_varTmpArr05(PROC_KINDNAME, lngElemP) = "Sub" Then
  '      strProcName = arr_varTmpArr05(PROC_NAME, lngElemP)
  '      intPos1 = InStr(strProcName, "_")
  '      If intPos1 > 0 Then
  '        lngTmp00 = Len(strProcName)
  '        ' ** Look for last underscore, not first.
  '        For lngZ = lngTmp00 To 1& Step -1&
  '          If Mid$(strProcName, lngZ, 1) = "_" Then
  '            strTmp01 = Mid$(strProcName, (lngZ + 1))
  '            strTmp02 = Left$(strProcName, (lngZ - 1))
  '            arr_varTmpArr05(PROC_DOC_ARR, lngElemP) = strTmp01      ' ** strEvent.
  '            arr_varTmpArr05(PROC_VEW_ARR, lngElemP) = strTmp02      ' ** strCtl.
  '            arr_varTmpArr05(PROC_VAR_ARR, lngElemP) = CBool(True)  ' ** PROC_XEVENT.
  '            Exit For
  '          End If
  '        Next
  '      Else
  '        ' ** Not a potential event, since they all have an underscore.
  '        arr_varTmpArr05(PROC_VAR_ARR, lngElemP) = CBool(False)
  '      End If
  '    End If  ' ** Procedure is a Sub.
  '  Next  ' ** For each procedure: lngY, lngElemP.
  '  arr_varMod(MOD_PROC_ARR, lngElemM) = arr_varTmpArr05
  'Next  ' ** For each module: lngX, lngElemM.

  'For lngX = 0& To (lngMods - 1&)
  '  lngElemM = lngX
  '  strModName = arr_varMod(MOD_NAME, lngElemM)
  '  arr_varTmpArr05 = arr_varMod(MOD_PROC_ARR, lngElemM)
  '  lngProcs = UBound(arr_varTmpArr05, 2) + 1&
'Debug.Print "'" & strModName & "  " & CStr(lngProcs)
  '  For lngY = 0& To (lngProcs - 1&)
  '    lngElemP = lngY
  '    If arr_varTmpArr05(PROC_KINDNAME, lngElemP) = "Sub" Then
  '      strProcName = arr_varTmpArr05(PROC_NAME, lngElemP)
  '      If arr_varTmpArr05(PROC_VAR_ARR, lngElemP) = True Then  ' ** PROC_XEVENT.
  '        strTmp01 = arr_varTmpArr05(PROC_DOC_ARR, lngElemP)     ' ** strEvent.
  '        strTmp02 = arr_varTmpArr05(PROC_VEW_ARR, lngElemP)     ' ** strCtl.
  '        strTmp03 = vbNullString                               ' ** strProc.
  '        Select Case strTmp01
  '        Case "BeforeInsert", "AfterInsert", "BeforeUpdate", "AfterUpdate", "BeforeDelConfirm", "AfterDelConfirm"
  '          ' ** These are the only known events that don't begin with 'On'.
  '          strTmp03 = strTmp01
  '        Case Else
  '          ' *****************************************************
  '          ' ** Array: arr_varEvent()
  '          ' **
  '          ' **   Field  Element  Name                Constant
  '          ' **   =====  =======  ==================  ==========
  '          ' **     1       0     vbcom_event_name    EVT_NAME
  '          ' **     2       1     vbcom_frm           EVT_ISFRM
  '          ' **     3       2     vbcom_rpt           EVT_ISRPT
  '          ' **     4       3     vbcom_ctl           EVT_ISCTL
  '          ' **
  '          ' *****************************************************
  '          For lngZ = 0& To (lngEvents - 1&)
  '            If arr_varEvent(EVT_NAME, lngZ) = ("On" & strTmp01) Then
  '              strTmp03 = "On" & strTmp01
  '              Exit For
  '            End If
  '          Next
  '        End Select

  '        If strTmp03 <> vbNullString Then  ' ** strProc.
  '          ' ** Yes, it's an event procedure.
  '          'Debug.Print "'EVT: " & strTmp02 & "  " & strTmp01
  '        Else
  '          ' ** No, it's unknown.
  '          Debug.Print "'EVT? " & strTmp02 & "  " & strTmp01
  '        End If
  '      End If  ' ** Procedure is a potential event: PROC_XEVENT.
  '    End If  ' ** Procedure is a Sub.
  '  Next  ' ** For each procedure: lngY, lngElemP.
  'Next  ' ** For each module: lngX, lngElemM.

  ' ** Report on in-line SQL code.
  If blnSQLChk = True Then
    For lngX = 0& To (lngMods - 1&)
      lngElemM = lngX

      strModName = arr_varMod(MOD_NAME, lngElemM)
      If IsEmpty(arr_varMod(MOD_SQL_ARR, lngElemM)) = False Then
        arr_varTmpArr05 = arr_varMod(MOD_SQL_ARR, lngElemM)
        lngSQLCodes = UBound(arr_varTmpArr05, 2) + 1&
      Else
        lngSQLCodes = 0&
      End If
      If IsEmpty(arr_varMod(MOD_PROC_ARR, lngElemM)) = False Then
        arr_varTmpArr06 = arr_varMod(MOD_PROC_ARR, lngElemM)
        lngProcs = UBound(arr_varTmpArr06, 2) + 1&
      Else
        lngProcs = 0&
      End If

      For lngY = 0& To (lngProcs - 1&)
        lngElemP = lngY

        lngDocLines = 0&
        ReDim arr_varDocLine(DOC_ELEMS, 0)

'If arr_varTmpArr06(PROC_NAME, lngElemP) = "CommonAssetListCode" Then
' ** All procedures have been documented for SQL-related words in arr_varSQLCode(),
' ** but only this specified one will have its actual SQL lines collected.

        For lngZ = 0& To (lngSQLCodes - 1&)
          lngElems = lngZ
          If arr_varTmpArr05(SQL_PROC_ELEM, lngElems) = lngElemP Then
            strTmp02 = arr_varTmpArr05(SQL_LINE_TXT, lngElems)
            If Right$(strTmp02, 1) = "_" Then  ' ** Line-Continuation.
  
              intLen = Len(strTmp02)
              strTmp02 = Left$(strTmp02, intLen - 1)
              strTmp02 = strTmp02 & "{lc}"  ' ** Line-Continuation.
              ' ** Strip line number.
              intPos1 = InStr(strTmp02, " ")
              If intPos1 > 0 Then
                If IsNumeric(Trim$(Left$(strTmp02, (intPos1 - 1)))) Then
                  strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))
                End If
              End If
  
              With vbp
                Set vbc = .VBComponents(strModName)
                With vbc
                  Set cod = .CodeModule
                  With cod
  
                    ' ** Check previously doc'd lines.
                    blnFound = False
                    For lngV = 0& To (lngDocLines - 1&)
                      If arr_varDocLine(DOC_LINE_NUM, lngV) = arr_varTmpArr05(SQL_LINE_NUM, lngElems) Then
                        blnFound = True
                        Exit For
                      End If
                    Next
  
                    If blnFound = False Then
                      ' ** It's a new line to document.
  
                      lngTmp00 = arr_varTmpArr05(SQL_LINE_NUM, lngElems)
                      lngDocLines = lngDocLines + 1&
                      lngElemD1 = lngDocLines - 1&
                      ReDim Preserve arr_varDocLine(DOC_ELEMS, lngElemD1)
                      arr_varDocLine(DOC_LINE_NUM, lngElemD1) = lngTmp00
                      arr_varDocLine(DOC_LINE_TXT, lngElemD1) = strTmp02
                      arr_varDocLine(DOC_SHOWN, lngElemD1) = CBool(False)
                      ' ** OK, if it's got an LC, then get all the lines after it.
                      Do While True
                        lngTmp00 = lngTmp00 + 1&
                        blnFound = False
                        For lngV = 0& To (lngDocLines - 1&)
                          ' ** See if this next line has been doc'd.
                          If arr_varDocLine(DOC_LINE_NUM, lngV) = lngTmp00 Then
                            blnFound = True
                            Exit For
                          End If
                        Next
                        If blnFound = False Then
                          ' ** This next line hasn't been documented.
                          strLine = Trim$(.Lines(lngTmp00, 1))
                          If Right$(strLine, 1) = "_" Then
                            intLen = Len(strLine)
                            strTmp02 = Left$(strLine, intLen - 1)
                            strTmp02 = strTmp02 & "{lc}"  ' ** Line-Continuation.
                          Else
                            strTmp02 = strLine
                          End If
                          ' ** Strip line number.
                          intPos1 = InStr(strTmp02, " ")
                          If intPos1 > 0 Then
                            If IsNumeric(Trim$(Left$(strTmp02, (intPos1 - 1)))) Then
                              strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))
                            End If
                          End If
                          lngDocLines = lngDocLines + 1&
                          lngElemD2 = lngDocLines - 1&
                          ReDim Preserve arr_varDocLine(DOC_ELEMS, lngElemD2)
                          arr_varDocLine(DOC_LINE_NUM, lngElemD2) = lngTmp00
                          arr_varDocLine(DOC_LINE_TXT, lngElemD2) = strTmp02
                          arr_varDocLine(DOC_SHOWN, lngElemD1) = CBool(False)
                          If Right$(strTmp02, 4) <> "{lc}" Then Exit Do
                        Else
                          Exit Do
                        End If  ' ** blnFound.
                      Loop  ' ** Next line and lines with line-continuation.
  
                      ' ** We'll have to check the lines before, as well.
                      lngTmp00 = arr_varTmpArr05(SQL_LINE_NUM, lngElems)
                      Do While True
                        lngTmp00 = lngTmp00 - 1&
                        blnFound = False
                        For lngV = 0& To (lngDocLines - 1&)
                          ' ** See if this previous line has been doc'd.
                          If arr_varDocLine(DOC_LINE_NUM, lngV) = lngTmp00 Then
                            blnFound = True
                            Exit For
                          End If
                        Next
                        If blnFound = False Then
                          ' ** This previous line hasn't been documented.
                          strLine = Trim$(.Lines(lngTmp00, 1))
                          If Right$(strLine, 1) = "_" Then
                            intLen = Len(strLine)
                            strTmp02 = Left$(strLine, intLen - 1)
                            strTmp02 = strTmp02 & "{lc}"  ' ** Line-Continuation.
                            ' ** Strip line number.
                            intPos1 = InStr(strTmp02, " ")
                            If intPos1 > 0 Then
                              If IsNumeric(Trim$(Left$(strTmp02, (intPos1 - 1)))) Then
                                strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))
                              End If
                            End If
                            lngDocLines = lngDocLines + 1&
                            lngElemD2 = lngDocLines
                            ReDim Preserve arr_varDocLine(DOC_ELEMS, lngElemD2)
                            arr_varDocLine(DOC_LINE_NUM, lngElemD2) = lngTmp00
                            arr_varDocLine(DOC_LINE_TXT, lngElemD2) = strTmp02
                            arr_varDocLine(DOC_SHOWN, lngElemD1) = CBool(False)
                          Else
                            Exit Do
                          End If
                        Else
                          Exit Do
                        End If  ' ** blnFound.
                      Loop  ' ** Previous lines.
  
                    End If  ' ** blnFound.
  
                  End With  ' ** This CodeModule: cod.
                End With  ' ** This VBComponent: vbc.
              End With  ' ** This ActiveProject: vbp.
  
              ' ** Binary Sort arr_varDocLine() array.
              For lngV = UBound(arr_varDocLine, 2) To 1& Step -1&
                For lngW = 0& To (lngV - 1&)
                  If arr_varDocLine(DOC_LINE_NUM, lngW) > arr_varDocLine(DOC_LINE_NUM, (lngW + 1)) Then
                    lngTmp00 = arr_varDocLine(DOC_LINE_NUM, lngW)
                    strTmp01 = arr_varDocLine(DOC_LINE_TXT, lngW)
                    arr_varDocLine(DOC_LINE_NUM, lngW) = arr_varDocLine(DOC_LINE_NUM, (lngW + 1))
                    arr_varDocLine(DOC_LINE_TXT, lngW) = arr_varDocLine(DOC_LINE_TXT, (lngW + 1))
                    arr_varDocLine(DOC_LINE_NUM, (lngW + 1)) = lngTmp00
                    arr_varDocLine(DOC_LINE_TXT, (lngW + 1)) = strTmp01
                  End If
                Next
              Next
  
              arr_varTmpArr06(PROC_DOC_ARR, lngElemP) = arr_varDocLine
  
            End If  ' ** Has line-continuation.
          End If  ' ** SQL line in this procedure.
        Next  ' ** For each SQL line: lngElemS, lngZ.

'End If  ' ** Limit to specified procedure.

      Next  ' ** For each procedure: lngElemP, lngY.

      arr_varMod(MOD_PROC_ARR, lngElemM) = arr_varTmpArr06

      ' ** We're still in this module: arr_varMod()
      ' ** arr_varTmpArr05() is arr_varSQLCode(), and has all SQL-related lines in this module.
      ' ** arr_varTmpArr06() is arr_varProc(), and has all procedures in this module.

      ' ** Show all actual SQL lines.
      For lngY = 0& To (lngProcs - 1&)
        lngElemP = lngY

        arr_varTmpArr07 = arr_varTmpArr06(PROC_DOC_ARR, lngElemP)
        ' ** arr_varTmpArr07() is arr_varDocLine(), and has all actual SQL lines in this procedure.
        strProcName = arr_varTmpArr06(PROC_NAME, lngElemP)

        strTmp01 = vbNullString
        lngViews = 0&
        ReDim arr_varView(VEW_ELEMS, 0)

'If strProcName = "CommonAssetListCode" Then
'If lngElemP > 40& And lngElemP <= 45& Then

        For lngZ = 0& To (lngSQLCodes - 1&)
          ' ** If there are arr_varSQLCode() lines in this procedure,
          ' ** coordinate them with documented lines in arr_varDocLine().
          lngElems = lngZ
          If arr_varTmpArr05(SQL_PROC_ELEM, lngElems) = lngElemP Then
            ' ** This procedure has SQL-related lines.
            lngTmp00 = arr_varTmpArr05(SQL_LINE_NUM, lngElems)
            If strTmp01 = vbNullString Then
              strTmp01 = strProcName
              lngViews = lngViews + 1&
              lngElemV1 = lngViews - 1&
              ReDim Preserve arr_varView(VEW_ELEMS, lngElemV1)
              arr_varView(VEW_LINE_NUM, lngElemV1) = 0&
              arr_varView(VEW_LINE_TXT, lngElemV1) = "'" & strTmp01 & "():"
              'Debug.Print "'" & strTmp01 & "():"
            End If
            If IsEmpty(arr_varTmpArr07) = False Then
              ' ** This procedure has actual SQL lines.
              lngDocLines = UBound(arr_varTmpArr07, 2) + 1&
              For lngV = 0& To (lngDocLines - 1&)
                lngElemD1 = lngV
                lngTmp04 = arr_varTmpArr07(DOC_LINE_NUM, lngElemD1)
                If lngTmp00 = lngTmp04 Then
                  arr_varTmpArr05(SQL_DOCD, lngElems) = CBool(True)
                End If
                If arr_varTmpArr07(DOC_SHOWN, lngElemD1) = False Then
                  lngViews = lngViews + 1&
                  lngElemV1 = lngViews - 1&
                  ReDim Preserve arr_varView(VEW_ELEMS, lngElemV1)
                  arr_varView(VEW_LINE_NUM, lngElemV1) = lngTmp04
                  arr_varView(VEW_LINE_TXT, lngElemV1) = "'" & Left$(CStr(lngTmp04) & "     ", 4) & _
                    " :D " & arr_varTmpArr07(DOC_LINE_TXT, lngElemD1)
                  'Debug.Print "'" & Left$(CStr(lngTmp04) & "     ", 4) & _
                  '  " :D " & arr_varTmpArr07(DOC_LINE_TXT, lngElemD1)
                  arr_varTmpArr07(DOC_SHOWN, lngElemD1) = CBool(True)
                End If
              Next
            Else
              ' ** Only SQL-related lines in this procedure, nothing in arr_varDocLines().
              strTmp02 = arr_varTmpArr05(SQL_LINE_TXT, lngElems)
              ' ** Strip line-continuation.
              If Right$(strTmp02, 1) = "_" Then
                strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
              End If
              ' ** Strip line number.
              intPos1 = InStr(strTmp02, " ")
              If intPos1 > 0 Then
                If IsNumeric(Trim$(Left$(strTmp02, (intPos1 - 1)))) Then
                  strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))
                End If
              End If
              lngViews = lngViews + 1&
              lngElemV1 = lngViews - 1&
              ReDim Preserve arr_varView(VEW_ELEMS, lngElemV1)
              arr_varView(VEW_LINE_NUM, lngElemV1) = arr_varTmpArr05(SQL_LINE_NUM, lngElems)
              arr_varView(VEW_LINE_TXT, lngElemV1) = "'" & _
                Left$(CStr(arr_varTmpArr05(SQL_LINE_NUM, lngElems)) & "     ", 4) & " :S " & strTmp02
              arr_varTmpArr05(SQL_DOCD, lngElems) = CBool(True)
              'Debug.Print "'" & Left$(CStr(arr_varTmpArr05(SQL_LINE_NUM, lngElemS)) & "     ", 4) & _
              '  " :S " & strTmp02
            End If  ' ** Actual SQL lines in this procedure.
          End If  ' ** This procedure has SQL-related lines.
        Next  ' ** For each SQL-related line: lngElemS, lngZ.

        For lngZ = 0& To (lngSQLCodes - 1&)
          lngElems = lngZ
          If arr_varTmpArr05(SQL_PROC_ELEM, lngElems) = lngElemP Then
            If arr_varTmpArr05(SQL_DOCD, lngElems) = False Then
              ' ** Show any remaining that aren't actual SQL lines.
              strTmp02 = arr_varTmpArr05(SQL_LINE_TXT, lngElems)
              ' ** Strip line-continuation.
              If Right$(strTmp02, 1) = "_" Then
                strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
              End If
              ' ** Strip line number.
              intPos1 = InStr(strTmp02, " ")
              If intPos1 > 0 Then
                If IsNumeric(Trim$(Left$(strTmp02, (intPos1 - 1)))) Then
                  strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))
                End If
              End If
              lngViews = lngViews + 1&
              lngElemV1 = lngViews - 1&
              ReDim Preserve arr_varView(VEW_ELEMS, lngElemV1)
              arr_varView(VEW_LINE_NUM, lngElemV1) = arr_varTmpArr05(SQL_LINE_NUM, lngElems)
              arr_varView(VEW_LINE_TXT, lngElemV1) = "'" & _
                Left$(CStr(arr_varTmpArr05(SQL_LINE_NUM, lngElems)) & "     ", 4) & " :S " & strTmp02
              'Debug.Print "'" & Left$(CStr(arr_varTmpArr05(SQL_LINE_NUM, lngElemS)) & "     ", 4) & _
              '  " :S " & strTmp02
              arr_varTmpArr05(SQL_DOCD, lngElems) = CBool(True)
            End If
          End If
        Next

        If lngViews > 0& Then

          ' ** Binary Sort arr_varView() array.
          For lngV = UBound(arr_varView, 2) To 1& Step -1&
            For lngW = 0& To (lngV - 1&)
              If arr_varView(VEW_LINE_NUM, lngW) > arr_varView(VEW_LINE_NUM, (lngW + 1)) Then
                lngTmp00 = arr_varView(VEW_LINE_NUM, lngW)
                strTmp01 = arr_varView(VEW_LINE_TXT, lngW)
                arr_varView(VEW_LINE_NUM, lngW) = arr_varView(VEW_LINE_NUM, (lngW + 1))
                arr_varView(VEW_LINE_TXT, lngW) = arr_varView(VEW_LINE_TXT, (lngW + 1))
                arr_varView(VEW_LINE_NUM, (lngW + 1)) = lngTmp00
                arr_varView(VEW_LINE_TXT, (lngW + 1)) = strTmp01
              End If
            Next
          Next

          ' ** Print the Results.
          For lngV = 0& To (lngViews - 1&)
            'Debug.Print arr_varView(VEW_LINE_TXT, lngV)
          Next

          arr_varTmpArr06(PROC_VEW_ARR, lngElemP) = arr_varView
          arr_varMod(MOD_PROC_ARR, lngElemM) = arr_varTmpArr06

        End If

'End If

      Next  ' ** For each procedure: lngElemP, lngY.

'OK, now that I've got everything, I've got to parse it into genuine SQL!
'So far, it appears that "S" lines and "D" lines don't mix within the same SQL statement.

      ' ** We're still in this module: arr_varMod()
      ' ** arr_varTmpArr05() is arr_varSQLCode(), and has all SQL-related lines in this module.
      arr_varTmpArr06 = arr_varMod(MOD_PROC_ARR, lngElemM)
      strModName = arr_varMod(MOD_NAME, lngElemM)
      For lngY = 0& To (lngProcs - 1&)
        lngElemP = lngY
'If arr_varTmpArr06(PROC_NAME, lngElemP) = "cmdTransactionsPrint_Click" Then
        strProcName = arr_varTmpArr06(PROC_NAME, lngElemP)
        arr_varTmpArr07 = arr_varTmpArr06(PROC_VEW_ARR, lngElemP)
        If IsEmpty(arr_varTmpArr07) = False Then

          ' ** Get any SQL variables in this procedure.
          lngVars = 0&
          ReDim arr_varVar(VAR_ELEMS, 0)
          For lngZ = 0& To (lngSQLCodes - 1&)
            lngElems = lngZ
            If arr_varTmpArr05(SQL_PROC_ELEM, lngElems) = lngElemP Then
              ' ** SQL-related lines for this procedure.
              If ((arr_varTmpArr05(SQL_RESP, lngElems) And SQL_HASVRB) <> 0) Then
                ' ** SQL-related line with variable declaration. (SQL_HASVCD is line where used.)
                strLine = arr_varTmpArr05(SQL_LINE_TXT, lngElems)
                intPos1 = InStr(strLine, "sql")
                Do While intPos1 > 0
                  strTmp01 = vbNullString
                  intRemPos = 0&: blnEOL = False
                  intPos2 = InStr((intPos1 + 1), strLine, " ")
                  If intPos2 = 0& Then
                    blnEOL = True
                    For lngV = (intPos1 - 1&) To 0& Step -1&
                      If Mid$(strLine, lngV, 1) = " " Then
                        strTmp01 = Trim$(Mid$(strLine, (lngV + 1)))
                        Exit For
                      End If
                    Next
                  Else
                    strTmp01 = Left$(strLine, (intPos2 - 1&))
                    intRemPos = InStr(intPos2, strLine, "'")  ' ** Entire remarked lines have already been eliminated.
                    If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))  ' ** (And a remark is
                    For lngV = (intPos1 - 1&) To 0& Step -1&                                       ' ** always preceeded
                      If Mid$(strTmp01, lngV, 1) = " " Then                                         ' ** by a space.)
                        strTmp01 = Trim$(Mid$(strTmp01, (lngV + 1)))
                        Exit For
                      End If
                    Next
                  End If
                  If strTmp01 <> vbNullString Then
                    lngVars = lngVars + 1&
                    lngElemV2 = lngVars - 1&
                    ReDim Preserve arr_varVar(VAR_ELEMS, lngElemV2)
                    arr_varVar(VAR_NAME, lngElemV2) = Trim$(strTmp01)
                  End If
                  If blnEOL = True Then Exit Do
                  intPos1 = InStr((intPos1 + 1), strLine, "sql")
                  If intRemPos > 0 And intPos1 > intRemPos Then Exit Do
                Loop  ' ** For each SQL-related variable in this line.
              Else
'DO WE WANT TO DOC RECORDSET ASSIGNMENTS IN THE SAME WAY?
'THIS LOOP JUST COLLECTS VARIABLE NAMES. RECORDSET ASSIGNMENTS, AND .SQL, .RunSQL, CAN BE HANDLED LATER.
              End If  ' ** SQL-related variable declarations.
            End If  ' ** SQL-related lines in this procedure.
          Next  ' ** For each SQL-related line: lngElemS, lngZ.

          ' ** Parse the SQL code line.
          lngParseSQLs = 0&: lngSQLBegElem = -1&
          blnHasMsgBox = False: blnMsgContinues = False: lngMsgBoxLineNum = 0&: lngMsgBoxTmp0 = 0&
          ReDim arr_varParseSQL(PAR_ELEMS, 0)
          lngViews = UBound(arr_varTmpArr07, 2) + 1&
          For lngZ = 0& To (lngViews - 1&)
            lngElemV1 = lngZ
            strLine = arr_varTmpArr07(VEW_LINE_TXT, lngElemV1)

            ' ** "D" lines come from arr_varDocLine().
            ' ** "S" lines come from from arr_varSQLCode().
            ' ** Let's start with "D" lines.
            If Mid$(strLine, 8, 1) = "D" And Len(Trim$(strLine)) > 8 Then
              ' ** A line from arr_varDocLine().

              strTmp01 = Mid$(strLine, 10)  ' ** Strip off the line number, etc.
              blnFound = False

              ' ** First, check variables collected above.
              For lngV = 0& To (lngVars - 1&)
                lngElemV2 = lngV
                strTmp03 = arr_varVar(VAR_NAME, lngElemV2)
                intPos1 = InStr(strTmp01, strTmp03)
                If intPos1 > 0 Then
                  ' ** At this point I'm not checking for similar variables, which I'll have to do eventually!
                  blnFound = True
                  intPos2 = InStr(strTmp01, strTmp03 & " = " & strTmp03 & " & ")
                  If intPos2 > 0 Then
                    ' ** Concatenation.
                    ' **   e.g.: strSQL = strSQL & "
                    'Debug.Print "'OTHER VAR: " & strTmp03 & " : " & strTmp01
                    'Stop
                    strTmp02 = Mid$(strTmp01, (intPos2 + Len(strTmp03 & " = " & strTmp03 & " & ")))
                    If Right$(strTmp02, 7) = " & {lc}" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 7))
                    lngParseSQLs = lngParseSQLs + 1&
                    lngElemQ = lngParseSQLs - 1&
                    ReDim Preserve arr_varParseSQL(PAR_ELEMS, lngElemQ)
                    arr_varParseSQL(PAR_VEW_ELEM, lngElemQ) = lngElemV1
                    arr_varParseSQL(PAR_RAW_LINE, lngElemQ) = strLine
                    arr_varParseSQL(PAR_BEG_ELEM, lngElemQ) = CLng(-1)  ' ** Indicating the assignment line.
                    lngSQLBegElem = lngElemQ
                    arr_varParseSQL(PAR_ASSIGN, lngElemQ) = CBool(True)
                    arr_varParseSQL(PAR_EDIT1, lngElemQ) = Trim$(strTmp02)
                    arr_varParseSQL(PAR_EDIT2, lngElemQ) = vbNullString
                  Else
                    intPos2 = InStr(strTmp01, strTmp03 & " = ")
                    If intPos2 > 0 Then
                      ' ** Simple assignment.
                      'Debug.Print "'VAR: " & strTmp03 & " : " & strTmp01
                      strTmp02 = Mid$(strTmp01, (intPos2 + Len(strTmp03 & " = ")))
                      If Right$(strTmp02, 7) = " & {lc}" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 7))
                      lngParseSQLs = lngParseSQLs + 1&
                      lngElemQ = lngParseSQLs - 1&
                      ReDim Preserve arr_varParseSQL(PAR_ELEMS, lngElemQ)
                      arr_varParseSQL(PAR_VEW_ELEM, lngElemQ) = lngElemV1
                      arr_varParseSQL(PAR_RAW_LINE, lngElemQ) = strLine
                      arr_varParseSQL(PAR_BEG_ELEM, lngElemQ) = CLng(-1)  ' ** Indicating the assignment line.
                      lngSQLBegElem = lngElemQ
                      arr_varParseSQL(PAR_ASSIGN, lngElemQ) = CBool(True)
                      arr_varParseSQL(PAR_EDIT1, lngElemQ) = Trim$(strTmp02)
                      arr_varParseSQL(PAR_EDIT2, lngElemQ) = vbNullString
                    Else
                      ' ** It's got a variable, but not where expected.
                      Stop
                    End If

                  End If  ' ** Concatenation or simple assignment.

                Else
                  ' ** Line doesn't use one of the SQL-related variables.
                End If  ' ** Line uses SQL-related variable.

              Next  ' ** For each SQL-related variable in arr_varVar().

              If blnFound = False Then
                ' ** I need to stop MsgBox's and subsequent lines from showing up!
                blnSkipIt = False
                If InStr(strLine, "MsgBox ") > 0 Or InStr(strLine, "MsgBox(") > 0 Or InStr(strLine, "InputBox(") > 0 Or _
                    InStr(strLine, "strMessage =") > 0 Or InStr(strLine, ".AsOf_lbl") > 0 Or _
                    InStr(strLine, ".DateRange_lbl") > 0 Or InStr(strLine, ".Filter =") > 0 Then
                  blnHasMsgBox = True
                  lngMsgBoxLineNum = arr_varTmpArr07(VEW_LINE_NUM, lngElemV1)
                  lngMsgBoxTmp0 = arr_varTmpArr07(VEW_LINE_NUM, lngElemV1)
                  If Right$(strTmp01, 7) = " & {lc}" Then
                    blnMsgContinues = True
                  Else
                    blnMsgContinues = False
                  End If
                  blnSkipIt = True
                Else
                  If blnHasMsgBox = True Then
                    If blnMsgContinues = True Then
                      If lngMsgBoxTmp0 = 0& Then
                        If arr_varTmpArr07(VEW_LINE_NUM, lngElemV1) = lngMsgBoxLineNum + 1& Then
                          blnSkipIt = True
                          lngMsgBoxTmp0 = arr_varTmpArr07(VEW_LINE_NUM, lngElemV1)
                          If Right$(strTmp01, 7) = " & {lc}" Then
                            blnMsgContinues = True
                          Else
                            blnMsgContinues = False
                          End If
                        Else
                          blnMsgContinues = False
                          blnHasMsgBox = False: lngMsgBoxLineNum = 0&: lngMsgBoxTmp0 = 0&
                        End If
                      Else
                        If arr_varTmpArr07(VEW_LINE_NUM, lngElemV1) = lngMsgBoxTmp0 + 1& Then
                          blnSkipIt = True
                          lngMsgBoxTmp0 = arr_varTmpArr07(VEW_LINE_NUM, lngElemV1)
                          If Right$(strTmp01, 7) = " & {lc}" Then
                            blnMsgContinues = True
                          Else
                            blnMsgContinues = False
                          End If
                        Else
                          blnMsgContinues = False
                          blnHasMsgBox = False: lngMsgBoxLineNum = 0&: lngMsgBoxTmp0 = 0&
                        End If
                      End If
                    Else
                      blnHasMsgBox = False: lngMsgBoxLineNum = 0&: lngMsgBoxTmp0 = 0&
                    End If
                    If InStr(strTmp01, "vbOKOnly") > 0 Or InStr(strTmp01, "vbYesNo") > 0 Then
                      blnSkipIt = True
                      blnHasMsgBox = False: blnMsgContinues = False: lngMsgBoxLineNum = 0&: lngMsgBoxTmp0 = 0&
                    End If
                    If Left$(strTmp01, 16) = "rsDataIn.Fields(" Or Left$(strTmp01, 8) = ".Fields(" Or _
                        Left$(strTmp01, 17) = "rsxDataIn.Fields(" Then
                      blnSkipIt = True
                    End If
                  End If
                  If blnSkipIt = False Then
                    strTmp02 = strTmp01
                    If Right$(strTmp02, 7) = " & {lc}" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 7))
                    lngParseSQLs = lngParseSQLs + 1&
                    lngElemQ = lngParseSQLs - 1&
                    ReDim Preserve arr_varParseSQL(PAR_ELEMS, lngElemQ)
                    arr_varParseSQL(PAR_VEW_ELEM, lngElemQ) = lngElemV1
                    arr_varParseSQL(PAR_RAW_LINE, lngElemQ) = strLine
                    arr_varParseSQL(PAR_BEG_ELEM, lngElemQ) = lngSQLBegElem  ' ** This assumes it belongs to the last assignment.
                    ' ** Check here for Recordset assignments and the SQL commands: .SQL and .RunSQL.
                    If Left$(strLine, 4) = "Set " Or InStr(strLine, ".Execute") > 0 Or _
                        InStr(strLine, ".SQL =") > 0 Or InStr(strLine, ".RunSQL") > 0 Then
                      arr_varParseSQL(PAR_ASSIGN, lngElemQ) = CBool(True)
                    Else
                      arr_varParseSQL(PAR_ASSIGN, lngElemQ) = CBool(False)
                    End If
                    arr_varParseSQL(PAR_EDIT1, lngElemQ) = Trim$(strTmp02)
                    arr_varParseSQL(PAR_EDIT2, lngElemQ) = vbNullString
                  End If
                End If
              End If

            End If  ' ** Line is from arr_varDocLine().

          Next  ' ** For each line in arr_varView().

          For lngZ = 0& To (lngParseSQLs - 1&)
            strTmp01 = arr_varParseSQL(PAR_EDIT1, lngZ)
            intPos1 = InStr(strTmp01, Chr(34) & Chr(39))  ' ** "'
            If intPos1 > 0 Then
'OK, now we've got to figure out how to deal with variable replacements!

            End If
            intPos1 = InStr(strTmp01, Chr(39) & Chr(34))  ' ** '"
            If intPos1 > 0 Then

            End If

          Next

          'For lngZ = 0& To (lngParseSQLs - 1&)
          '  strTmp01 = arr_varParseSQL(PAR_EDIT1, lngZ)
          '  If Right$(strTmp01, 1) = Chr(34) Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
          '  If Left$(strTmp01, 1) = Chr(34) Then strTmp01 = Mid$(strTmp01, 2)
          '  arr_varParseSQL(PAR_EDIT2, lngZ) = Trim$(strTmp01)
          'Next

          arr_varTmpArr06(PROC_VAR_ARR, lngElemP) = arr_varVar
          arr_varTmpArr06(PROC_PAR_ARR, lngElemP) = arr_varParseSQL

          For lngZ = 0& To (lngParseSQLs - 1&)
            strTmp01 = IIf(arr_varParseSQL(PAR_ASSIGN, lngZ) = True, "A", "C")  ' ** Assignment line, Continuation line.
            If arr_varParseSQL(PAR_BEG_ELEM, lngZ) = -1 Then
              strTmp02 = "-1 "
            Else
              strTmp02 = " " & Left$(CStr(arr_varParseSQL(PAR_BEG_ELEM, lngZ)) & "  ", 2)
            End If
'At this point, the line collecting looks good.
'PAR_EDIT1 has the SQL (with quotes), and PAR_BEG_ELEM has the assignment line (-1 indicating the assignment line itself).
            'Debug.Print "'MOD: " & strModName & " PROC: " & arr_varTmpArr06(PROC_NAME, lngY) & "() " & _
            '  "TYP: " & strTmp01 & " ASGN ELEM: " & arr_varParseSQL(PAR_BEG_ELEM, lngZ)
            'Debug.Print "' SQL: " & arr_varParseSQL(PAR_EDIT1, lngZ)
          Next

        End If  ' ** Procedure has arr_varView().
'End If
      Next  ' ** For each procedure in this module: lngElemP, lngY.

      arr_varMod(MOD_PROC_ARR, lngElemM) = arr_varTmpArr06
      ' ** Just show oddities.
      For lngY = 0& To (lngProcs - 1&)
        lngElemP = lngY
        intProcHasSQL = arr_varTmpArr06(PROC_HAS_SQL, lngElemP)
        If (intProcHasSQL And SQL_NONE) <> 0 Then
          If (((intProcHasSQL And SQL_HASVRB) = 0) And ((intProcHasSQL And SQL_HASCMD) = 0) And _
             ((intProcHasSQL And SQL_HASTRM))) Or _
             (((intProcHasSQL And SQL_HASVRB)) And ((intProcHasSQL And SQL_HASCMD) = 0) And _
             ((intProcHasSQL And SQL_HASTRM) = 0)) Or _
             (((intProcHasSQL And SQL_HASVRB) = 0) And ((intProcHasSQL And SQL_HASCMD)) And _
             ((intProcHasSQL And SQL_HASTRM) = 0)) Then
            ' ** This shows which procs have a SQL variable, SQL command, or a SQL term.
            If intProcHasSQL And SQL_HASVRB Then
              'Debug.Print "'SQL VAR: " & arr_varTmpArr06(PROC_NAME, lngElemP)
            End If
            If intProcHasSQL And SQL_HASCMD Then
              'Debug.Print "'SQL CMD: " & arr_varTmpArr06(PROC_NAME, lngElemP)
            End If
            If intProcHasSQL And SQL_HASTRM Then
              'Debug.Print "'SQL TERM: " & arr_varTmpArr06(PROC_NAME, lngElemP)
            End If
          End If
        End If
      Next

      ' ** Show all SQL lines.
      For lngY = 0& To (lngSQLCodes - 1&)
        lngElems = lngY
        strTmp01 = arr_varTmpArr05(SQL_LINE_TXT, lngElems)
        If Right$(strTmp01, 1) = "_" Then
          intLen = Len(strTmp01)
          strTmp01 = Left$(strTmp01, intLen - 1)
          strTmp01 = strTmp01 & "{lc}"  ' ** Line-Continuation.
        End If
        strTmp02 = Left$(CStr(arr_varTmpArr05(SQL_LINE_NUM, lngElems)) & "     ", 5)
        'Debug.Print "'" & strTmp02 & ": " & strTmp01
      Next

      ' ** SQL DOC: Let's see what kind of data I've got.
      Set dbs = CurrentDb
      With dbs
        'Set rst = dbs.OpenRecordset("zz_tbl_sql_code_03", dbOpenDynaset)
        'Set rst = dbs.OpenRecordset("zz_tbl_sql_code_02", dbOpenDynaset)
        Set rst = dbs.OpenRecordset("zz_tbl_sql_code_01", dbOpenDynaset, dbConsistent)
        With rst
          If arr_varMod(MOD_NAME, lngElemM) <> "modGlobConst" Then
            arr_varTmpArr05 = arr_varMod(MOD_PROC_ARR, lngElemM)
            For lngY = 0& To UBound(arr_varTmpArr05, 2)  'arr_varMod()
              If IsEmpty(arr_varTmpArr05(PROC_PAR_ARR, lngY)) = False Then
                arr_varTmpArr06 = arr_varTmpArr05(PROC_PAR_ARR, lngY)
                arr_varTmpArr07 = arr_varTmpArr05(PROC_VEW_ARR, lngY)
                For lngZ = 0& To UBound(arr_varTmpArr06, 2)  'arr_varProc()
                  If IsNull(arr_varTmpArr06(PAR_EDIT1, lngZ)) = False Then
                    If arr_varTmpArr06(PAR_EDIT1, lngZ) <> vbNullString And _
                        arr_varTmpArr07(VEW_LINE_NUM, arr_varTmpArr06(PAR_VEW_ELEM, lngZ)) <> 0 Then
                      If Left$(arr_varTmpArr06(PAR_EDIT1, lngZ), 16) <> "rsDataIn.Fields(" And _
                          Left$(arr_varTmpArr06(PAR_EDIT1, lngZ), 8) <> ".Fields(" And _
                          Left$(arr_varTmpArr06(PAR_EDIT1, lngZ), 17) <> "rsxDataIn.Fields(" And _
                          arr_varTmpArr05(PROC_NAME, lngY) <> "FormRef" Then
                        lngVBComID = DLookup("[vbcom_id]", "tblVBComponent", "[vbcom_name] = '" & strModName & "'")
                        lngFrmID = DLookup("[frm_id]", "tblForm", "[frm_name] = '" & Mid$(strModName, 6) & "'")
                        lngTmp00 = arr_varTmpArr07(VEW_LINE_NUM, arr_varTmpArr06(PAR_VEW_ELEM, lngZ))
                        .FindFirst "[vbcom_id] = " & CStr(lngVBComID) & " And [sql1_linenum] = " & CStr(lngTmp00)
                        If .NoMatch = False Then
                          ' ** This module line number is already in table.
                          If IsNull(arr_varTmpArr06(PAR_RAW_LINE, lngZ)) = False Then
                            If arr_varTmpArr06(PAR_RAW_LINE, lngZ) <> vbNullString Then
                              If IsNull(![sql1_raw]) = False Then
                                If ![sql1_raw] <> arr_varTmpArr06(PAR_RAW_LINE, lngZ) Then
                                  ' ** A change in code, so replace everything.
                                  .Edit
                                  ![sql1_assign] = arr_varTmpArr06(PAR_ASSIGN, lngZ)
                                  ![sql1_aline] = arr_varTmpArr06(PAR_BEG_ELEM, lngZ)  ' ** Assignment line element.
                                  ![sql1_sql] = arr_varTmpArr06(PAR_EDIT1, lngZ)
                                  ![sql1_raw] = arr_varTmpArr06(PAR_RAW_LINE, lngZ)
                                  ![sql1_IsUpd] = True
                                  ![sql1_datemodified] = Now()
                                  .Update
                                Else
                                  ' ** Already correctly documented.
                                End If
                              End If
                            End If
                          End If
                          If ![sql1_assign] = False And arr_varTmpArr06(PAR_ASSIGN, lngZ) = True Then
                            ' ** A line may have been manually changed to True, so don't reverse that.
                            .Edit
                            ![sql1_assign] = arr_varTmpArr06(PAR_ASSIGN, lngZ)
                            ![sql1_IsUpd] = True
                            ![sql1_datemodified] = Now()
                            .Update
                          End If
                          If arr_varTmpArr06(PAR_BEG_ELEM, lngZ) > 0 And ![sql1_aline] <= 0 Then
'THIS ELEMENT DOESN'T SEEM TO HAVE VALUES MUCH!
                            .Edit
                            ![sql1_aline] = arr_varTmpArr06(PAR_BEG_ELEM, lngZ)
                            ![sql1_IsUpd] = True
                            ![sql1_datemodified] = Now()
                            .Update
                          End If
                          .Edit
                          ![sql1_Chk] = True
                          .Update
                        Else
                          .AddNew
                          ![vbcom_id] = lngVBComID
                          ![objtype_type] = acForm
                          ![frm_id] = lngFrmID
                          ![sql1_module] = strModName
                          ![sql1_procedure] = arr_varTmpArr05(PROC_NAME, lngY)
                          ![sql1_linenum] = lngTmp00
                          ![sql1_assign] = arr_varTmpArr06(PAR_ASSIGN, lngZ)
                          ![sql1_aline] = arr_varTmpArr06(PAR_BEG_ELEM, lngZ)  ' ** Assignment line element.
                          ![sql1_sql] = arr_varTmpArr06(PAR_EDIT1, lngZ)
                          If IsNull(arr_varTmpArr06(PAR_RAW_LINE, lngZ)) = False Then
                            If arr_varTmpArr06(PAR_RAW_LINE, lngZ) <> vbNullString Then
                              ![sql1_raw] = arr_varTmpArr06(PAR_RAW_LINE, lngZ)
                            End If
                          End If
                          ![sql1_IsUpd] = True
                          ![sql1_Chk] = True
                          ![sql1_datemodified] = Now()
                          .Update
                        End If
                      End If
                    End If
                  End If
                Next
              End If
            Next
          End If
          .Close
        End With
        .Close
      End With

    Next  ' ** For each module: lngElemM, lngX.
  End If  ' ** blnSQLChk.

  ' ** Report on procedure scope and explicit typing.
  For lngX = 0& To (lngMods - 1&)
    lngProcs = arr_varMod(MOD_PROCS, lngElemM)
    arr_varTmpArr05 = arr_varMod(MOD_PROC_ARR, lngElemM)
    For lngY = 0& To (lngProcs - 1&)
      lngElemP = lngY
      If arr_varTmpArr05(PROC_SCOPE, lngElemP) = False Then
        Debug.Print "'NO SCOPE: " & arr_varMod(MOD_NAME, lngElemM) & " " & arr_varTmpArr05(PROC_NAME, lngElemP) & "()"
      End If
      If (arr_varTmpArr05(PROC_KINDNAME, lngElemP) = "Function" Or arr_varTmpArr05(PROC_KINDNAME, lngElemP) = "Get") And _
         arr_varTmpArr05(PROC_TYPED, lngElemP) = False Then
        Debug.Print "'NO TYPE: " & arr_varMod(MOD_NAME, lngElemM) & " " & arr_varTmpArr05(PROC_KINDNAME, lngElemP) & _
          " " & arr_varTmpArr05(PROC_NAME, lngElemP) & "()"
      End If
    Next
  Next

  '' ** Get form and report tags.
  'For lngX = 0& To (lngMods - 1&)
  '  lngElemM = lngX
  '  If arr_varMod(MOD_OBJ, lngElemM) = acForm Then
  '    'If arr_varMod(MOD_DISPNAME, lngElemM) = False Then
  '      strTmp01 = Mid$(arr_varMod(MOD_NAME, lngElemM), 6)
  '      DoCmd.OpenForm strTmp01, acDesign, , , , acHidden
  '      Set frm = Forms(strTmp01)
  '      With frm
  '        If .Tag <> vbNullString Then
  '          arr_varMod(MOD_TAG, lngElemM) = .Tag
  '          If InStr(.Tag, "subform") > 0 Then
  '            arr_varMod(MOD_SUBFORM, lngElemM) = CBool(True)
  '          End If
  '        End If
  '      End With
  '      DoCmd.Close acForm, strTmp01, acSaveNo
  '    'End If
  '  ElseIf arr_varMod(MOD_OBJ, lngElemM) = acReport Then
  '    'If arr_varMod(MOD_DISPNAME, lngElemM) = False Then
  '      strTmp01 = Mid$(arr_varMod(MOD_NAME, lngElemM), 8)
  '      DoCmd.OpenReport strTmp01, acViewDesign
  '      Set rpt = Reports(strTmp01)
  '      With rpt
  '        If .Tag <> vbNullString Then
  '          arr_varMod(MOD_TAG, lngElemM) = .Tag
  '          If InStr(.Tag, "subform") > 0 Or InStr(.Tag, "subreport") > 0 Then
  '            arr_varMod(MOD_SUBFORM, lngElemM) = CBool(True)
  '          End If
  '        End If
  '      End With
  '      DoCmd.Close acReport, strTmp01, acSaveNo
  '    'End If
  '  End If
  'Next

  ' ** Report on DispFrmName()/DispRptName() call.
  'For lngX = 0& To (lngMods - 1&)
  '  lngElemM = lngX
  '  If arr_varMod(MOD_OBJ, lngElemM) = acForm And arr_varMod(MOD_DISPNAME, lngElemM) = False Then
  '    If arr_varMod(MOD_TAG, lngElemM) = vbNullString Then
  '      Debug.Print "'" & arr_varMod(MOD_NAME, lngElemM) & _
  '        IIf(arr_varMod(MOD_OBJ_OPEN, lngElemM) = False, " NO Form_Open()", "")
  '    ElseIf InStr(arr_varMod(MOD_TAG, lngElemM), "subform") = 0 Then
  '      Debug.Print "'" & arr_varMod(MOD_NAME, lngElemM) & _
  '        IIf(arr_varMod(MOD_OBJ_OPEN, lngElemM) = False, " NO Form_Open()", "")
  '    End If
  '  ElseIf arr_varMod(MOD_OBJ, lngElemM) = acReport And arr_varMod(MOD_DISPNAME, lngElemM) = False Then
  '    If arr_varMod(MOD_TAG, lngElemM) = vbNullString Then
  '      Debug.Print "'" & arr_varMod(MOD_NAME, lngElemM) & _
  '        IIf(arr_varMod(MOD_OBJ_OPEN, lngElemM) = False, " NO Report_Open()", "") & " " & _
  '        arr_varMod(MOD_DISPNAME, lngElemM)
  '    ElseIf InStr(arr_varMod(MOD_TAG, lngElemM), "subreport") = 0 Then
  '      Debug.Print "'" & arr_varMod(MOD_NAME, lngElemM) & _
  '        IIf(arr_varMod(MOD_OBJ_OPEN, lngElemM) = False, " NO Report_Open()", "") & " " & _
  '        arr_varMod(MOD_DISPNAME, lngElemM)
  '    End If
  '  End If
  '  If arr_varMod(MOD_DISPCNT, lngElemM) > 1& Then
  '    ' ** frmJournal OK!
  '    Debug.Print "'EXTRA DISPNAME: " & arr_varMod(MOD_NAME, lngElemM)
  '  End If
  'Next

  '' ** Collect names of ALL reports, with or without a module.
  'Set prj = Application.CurrentProject
  'For Each rptao In prj.AllReports
  '  With rptao
  '    blnFound = False
  '    ' ** Check report against all modules searched.
  '    For lngX = 0& To (lngMods - 1&)
  '      lngElemM = lngX
  '      If .Name = Mid$(arr_varMod(MOD_NAME, lngElemM), 8) Then
  '        If arr_varMod(MOD_SUBFORM, lngElemM) = False Then
  '          blnFound = True
  '          If arr_varMod(MOD_RPTCALL, lngElemM) = False Then
  '            'Debug.Print "'NO RPTCALL: " & .Name
  '          End If
  '          Exit For
  '        End If
  '      End If
  '    Next
  '    If blnFound = False Then
  '      DoCmd.OpenReport .Name, acViewDesign
  '      Set rpt = Reports(.Name)
  '      If .Name <> "zz_rptRelationships" Then
  '        With rpt
  '          If .Tag <> vbNullString Then
  '            If InStr(.Tag, "subform") = 0 And InStr(.Tag, "subreport") = 0 Then
  '              Debug.Print "'NO MOD: " & .Name
  '            End If
  '          Else
  '            Debug.Print "'NO MOD: " & .Name
  '          End If
  '        End With
  '      End If
  '      DoCmd.Close acReport, .Name, acSaveNo
  '    End If
  '  End With
  'Next

  '' ** Report on documented errors.
  'For lngX = 0& To (lngErrs1 - 1&)
  '  lngElemE1 = lngX
  '  strLine = arr_varErr1(ERR_LINE_TXT, lngElemE1)
  '  If arr_varErr1(ERR_EQUALS, lngElemE1) = True Then
  '    ' ** Err = ...
  '    intPos1 = InStr(strLine, " Err = ")
  '    If intPos1 > 0 Then
  '      strTmp01 = Trim$(Mid$(strLine, (intPos1 + 7)))
  '      intPos1 = InStr(strTmp01, " ")
  '      If intPos1 > 0 Then
  '        strTmp02 = Trim$(Left$(strTmp01, intPos1))
  '      Else
  '        strTmp02 = strTmp01
  '      End If
  '      If IsNumeric(strTmp02) = True Then
  '        arr_varErr1(ERR_NUM, lngElemE1) = CLng(Val(strTmp02))
  '      End If
  '      If intPos1 > 0 Then
  '        strTmp01 = Trim$(Mid$(strTmp01, intPos1))
  '        intPos1 = InStr(strTmp01, "'")
  '        If intPos1 > 0 Then
  '          arr_varErr1(ERR_REM, lngElemE1) = Mid$(strTmp01, intPos1)
  '        End If
  '      End If
  '    End If
  '    ' ** Err.Number = ...
  '    intPos1 = InStr(strLine, " Err.Number = ")
  '    If intPos1 > 0 Then
  '      strTmp01 = Trim$(Mid$(strLine, (intPos1 + 14)))
  '      intPos1 = InStr(strTmp01, " ")
  '      If intPos1 > 0 Then
  '        strTmp02 = Trim$(Left$(strTmp01, intPos1))
  '      Else
  '        strTmp02 = strTmp01
  '      End If
  '      If IsNumeric(strTmp02) = True Then
  '        arr_varErr1(ERR_NUM, lngElemE1) = CLng(Val(strTmp02))
  '      End If
  '      If intPos1 > 0 Then
  '        strTmp01 = Trim$(Mid$(strTmp01, intPos1))
  '        intPos1 = InStr(strTmp01, "'")
  '        If intPos1 > 0 Then
  '          arr_varErr1(ERR_REM, lngElemE1) = Mid$(strTmp01, intPos1)
  '        End If
  '      End If
  '    End If
  '  End If
  'Next

  'lngErrs2 = 0&
  'ReDim arr_varErr2(ERR_ELEMS, 0)

  'blnSkip = False
  'For lngX = 0& To (lngErrs1 - 1&)
  '  lngElemE1 = lngX
  '  If arr_varErr1(ERR_NUM, lngElemE1) <> 0& Then
  '    blnSkip = False
  '    For lngY = 0& To (lngErrs2 - 1&)
  '      lngElemE2 = lngY
  '      If arr_varErr2(ERR_NUM, lngElemE2) = arr_varErr1(ERR_NUM, lngElemE1) Then
  '        blnSkip = True  ' ** lngElemE2 should remain the found element.
  '        Exit For
  '      End If
  '    Next
  '    If blnSkip = False Then
  '      lngErrs2 = lngErrs2 + 1&
  '      lngElemE2 = lngErrs2 - 1&  ' ** lngElemE2 is the new element.
  '      ReDim Preserve arr_varErr2(ERR_ELEMS, lngElemE2)
  '      arr_varErr2(ERR_NUM, lngElemE2) = arr_varErr1(ERR_NUM, lngElemE1)
  '      arr_varErr2(ERR_REM, lngElemE2) = vbNullString
  '    Else
  '      blnSkip = False
  '    End If
  '    strTmp01 = arr_varErr1(ERR_REM, lngElemE1)
  '    If strTmp01 <> vbNullString Then
  '      If Left$(strTmp01, 1) = "'" Then
  '        If arr_varErr2(ERR_REM, lngElemE2) = vbNullString Then
  '          arr_varErr2(ERR_REM, lngElemE2) = arr_varErr1(ERR_REM, lngElemE1)
  '        ElseIf Len(arr_varErr1(ERR_REM, lngElemE1)) > Len(arr_varErr2(ERR_REM, lngElemE2)) Then
  '          arr_varErr2(ERR_REM, lngElemE2) = arr_varErr1(ERR_REM, lngElemE1)
  '        End If
  '      Else
  '        intPos1 = InStr(strTmp01, " ")
  '        If intPos1 = 0 Then
  '          If IsNumeric(strTmp01) = True Then
  '            For lngY = 0& To (lngErrs2 - 1&)
  '              lngElemE2 = lngY
  '              If arr_varErr2(ERR_NUM, lngElemE2) = CLng(strTmp01) Then
  '                blnSkip = True  ' ** lngElemE2 should remain the found element.
  '                Exit For
  '              End If
  '            Next
  '            If blnSkip = False Then
  '              lngErrs2 = lngErrs2 + 1&
  '              lngElemE2 = lngErrs2 - 1&  ' ** lngElemE2 is the new element.
  '              ReDim Preserve arr_varErr2(ERR_ELEMS, lngElemE2)
  '              arr_varErr2(ERR_NUM, lngElemE2) = CLng(strTmp01)
  '              arr_varErr2(ERR_REM, lngElemE2) = vbNullString
  '            Else
  '              blnSkip = False
  '            End If
  '          End If
  '        Else
  '          ' ** If it's a list of error numbers, see that each one gets in the arr_varErr2() array.
  '          blnSkip = False
  '          Do While intPos1 > 0
  '            strTmp02 = Trim$(Left$(strTmp01, intPos1))
  '            If Right$(strTmp02, 1) = "," Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
  '            strTmp01 = Trim$(Mid$(strTmp01, intPos1))
  '            If IsNumeric(strTmp02) = True Then
  '              For lngY = 0& To (lngErrs2 - 1&)
  '                lngElemE2 = lngY
  '                If arr_varErr2(ERR_NUM, lngElemE2) = CLng(strTmp02) Then
  '                  blnSkip = True  ' ** lngElemE2 should remain the found element.
  '                  Exit For
  '                End If
  '              Next
  '              If blnSkip = False Then
  '                lngErrs2 = lngErrs2 + 1&
  '                lngElemE2 = lngErrs2 - 1&  ' ** lngElemE2 is the new element.
  '                ReDim Preserve arr_varErr2(ERR_ELEMS, lngElemE2)
  '                arr_varErr2(ERR_NUM, lngElemE2) = CLng(strTmp02)
  '                arr_varErr2(ERR_REM, lngElemE2) = vbNullString
  '              Else
  '                blnSkip = False
  '              End If
  '            End If
  '            intPos1 = InStr(strTmp01, " ")
  '          Loop
  '          If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
  '          blnSkip = False
  '          If IsNumeric(strTmp01) = True Then
  '            For lngY = 0& To (lngErrs2 - 1&)
  '              lngElemE2 = lngY
  '              If arr_varErr2(ERR_NUM, lngElemE2) = CLng(strTmp01) Then
  '                blnSkip = True  ' ** lngElemE2 should remain the found element.
  '                Exit For
  '              End If
  '            Next
  '            If blnSkip = False Then
  '              lngErrs2 = lngErrs2 + 1&
  '              lngElemE2 = lngErrs2 - 1&  ' ** lngElemE2 is the new element.
  '              ReDim Preserve arr_varErr2(ERR_ELEMS, lngElemE2)
  '              arr_varErr2(ERR_NUM, lngElemE2) = CLng(strTmp01)
  '              arr_varErr2(ERR_REM, lngElemE2) = vbNullString
  '            Else
  '              blnSkip = False
  '            End If
  '          End If
  '        End If
  '      End If
  '    End If
  '  End If
  'Next

  '' ** Binary Sort arr_varErr2() array.
  '' ** Only used 2 elements.
  'For lngX = UBound(arr_varErr2, 2) To 1& Step -1&
  '  For lngY = 0& To (lngX - 1&)
  '    If arr_varErr2(ERR_NUM, lngY) > arr_varErr2(ERR_NUM, (lngY + 1)) Then
  '      lngTmp00 = arr_varErr2(ERR_NUM, lngY)
  '      strTmp01 = arr_varErr2(ERR_REM, lngY)
  '      arr_varErr2(ERR_NUM, lngY) = arr_varErr2(ERR_NUM, (lngY + 1))
  '      arr_varErr2(ERR_REM, lngY) = arr_varErr2(ERR_REM, (lngY + 1))
  '      arr_varErr2(ERR_NUM, (lngY + 1)) = lngTmp00
  '      arr_varErr2(ERR_REM, (lngY + 1)) = strTmp01
  '    End If
  '  Next
  'Next

  'Debug.Print "'LINES FOUND: " & lngErrs1
  'For lngX = 0& To (lngErrs2 - 1&)
  '  lngElemE2 = lngX
  '  If arr_varErr2(ERR_NUM, lngElemE2) <> 0 Then
  '    strTmp01 = Left$(CStr(arr_varErr2(ERR_NUM, lngElemE2)) & "     ", 5)
  '    Debug.Print "'" & strTmp01 & arr_varErr2(ERR_REM, lngElemE2)
  '  End If
  'Next

If False Then
  ' ** Report on Options, THIS_NAME, THIS_PROC, Error label, Exit label.
  For lngX = 0& To (lngMods - 1&)
    lngElemM = lngX
    If arr_varMod(MOD_THIS, lngElemM) = False Then
      Debug.Print "'NO THIS_NAME: " & arr_varMod(MOD_NAME, lngElemM) & _
        " PROCS: " & arr_varMod(MOD_PROCS, lngElemM)
    End If
    If arr_varMod(MOD_OPDAT, lngElemM) = False Then
      Debug.Print "'NO OPTION COMPARE: " & arr_varMod(MOD_NAME, lngElemM)
    End If
    If arr_varMod(MOD_OPEXP, lngElemM) = False Then
      Debug.Print "'NO OPTION EXPLICIT: " & arr_varMod(MOD_NAME, lngElemM)
    End If
    lngProcs = arr_varMod(MOD_PROCS, lngElemM)
    If lngProcs > 0& Then
      arr_varTmpArr05 = arr_varMod(MOD_PROC_ARR, lngElemM)
      lngNoThis = 0&: lngNoErr = 0&: lngNoExit = 0&
      For lngY = 0& To (lngProcs - 1&)
        lngElemP = lngY
        If arr_varTmpArr05(PROC_THIS, lngElemP) = False Then lngNoThis = lngNoThis + 1&
        If arr_varTmpArr05(PROC_ERR_LBL, lngElemP) = False Then lngNoErr = lngNoErr + 1&
        If arr_varTmpArr05(PROC_EXIT_LBL, lngElemP) = False Then lngNoExit = lngNoExit + 1&
      Next
      If lngNoThis = 0& And lngNoErr = 0& And lngNoExit = 0& Then
        'Debug.Print "'ALL HERE! " & arr_varMod(MOD_NAME, lngElemM)
      Else
        If lngNoErr > 0& Or lngNoExit > 0& Then
          'LEAVE THIS_PROC OUT FOR NOW!
          Debug.Print "'NO THIS: " & Left$(CStr(lngNoThis) & "   ", 3) & _
            " NO ERR: " & Left$(CStr(lngNoErr) & "   ", 3) & _
            " NO EXIT: " & Left$(CStr(lngNoExit) & "   ", 3) & _
            " PROCS: " & Left$(CStr(lngProcs) & "   ", 3) & " " & arr_varMod(MOD_NAME, lngElemM)
        End If
      End If
    Else
      'Debug.Print "'NO PROCS! " & arr_varMod(MOD_NAME, lngElemM)
    End If
  Next
End If

If False Then
  ' ** Add DispFrmName()/DispRptName() call.
  intPos2 = 0
  For lngX = 0& To (lngMods - 1&)
    lngElemM = lngX
    If arr_varMod(MOD_OBJ, lngElemM) <> acModule And arr_varMod(MOD_DISPNAME, lngElemM) = False And _
       arr_varMod(MOD_SUBFORM, lngElemM) = False Then
      With vbp
        Set vbc = .VBComponents(arr_varMod(MOD_NAME, lngElemM))
        With vbc
          Set cod = .CodeModule
          With cod
            lngLineNum = .CountOfDeclarationLines + 1&
            If arr_varMod(MOD_OBJ_OPEN, lngElemM) = False Then
              intPos2 = intPos2 + 1
              .InsertLines lngLineNum, ""
              .InsertLines lngLineNum, "End Sub"
              .InsertLines lngLineNum, ""
              .InsertLines lngLineNum, "160     Resume EXITP"
              .InsertLines lngLineNum, "150     End Select"
              .InsertLines lngLineNum, "140       zErrorHandler THIS_NAME, THIS_PROC, Err.Number, Erl"
              .InsertLines lngLineNum, "        Case Else"
              .InsertLines lngLineNum, "130     Select Case Err.Number"
              .InsertLines lngLineNum, "ERRH:"
              .InsertLines lngLineNum, ""
              .InsertLines lngLineNum, "120     Exit Sub"
              .InsertLines lngLineNum, "EXITP:"
              .InsertLines lngLineNum, ""
              If arr_varMod(MOD_OBJ, lngElemM) = acForm Then
                .InsertLines lngLineNum, "110     DispFrmName Me  ' ** Module Procedure: modFileUtilities."
              Else
                .InsertLines lngLineNum, "110     DispRptName Me  ' ** Module Procedure: modFileUtilities."
              End If
              .InsertLines lngLineNum, ""
              If arr_varMod(MOD_OBJ, lngElemM) = acForm Then
                .InsertLines lngLineNum, "        Const THIS_PROC As String = " & Chr(34) & "Form_Open" & Chr(34)
              Else
                .InsertLines lngLineNum, "        Const THIS_PROC As String = " & Chr(34) & "Report_Open" & Chr(34)
              End If
              .InsertLines lngLineNum, ""
              .InsertLines lngLineNum, "100   On Error GoTo ERRH"
              .InsertLines lngLineNum, ""
              If arr_varMod(MOD_OBJ, lngElemM) = acForm Then
                .InsertLines lngLineNum, "Private Sub Form_Open(Cancel As Integer)"
              Else
                .InsertLines lngLineNum, "Private Sub Report_Open(Cancel As Integer)"
              End If
              .InsertLines lngLineNum, ""
            Else
              arr_varTmpArr05 = arr_varMod(MOD_PROC_ARR, lngElemM)
              lngElemP = arr_varMod(MOD_OOP_ELEM, lngElemM)
              lngProcStart = arr_varTmpArr05(PROC_START, lngElemP)
              lngProcEnd = arr_varTmpArr05(PROC_END, lngElemP)
              lngDims = 0&: lngZ = 0&
              For lngY = (lngProcStart + 1&) To (lngProcEnd - 1&)
                lngLineNum = lngY
                strLine = Trim$(.Lines(lngLineNum, 1))
                If strLine <> vbNullString Then
                  ' ** OK, we want to be after On Error, Dim's, Static's, and Const's.
                  If arr_varTmpArr05(PROC_ERR_LBL, lngElemP) = False Then
                    If arr_varTmpArr05(PROC_DIMS, lngElemP) = 0& Then
                      ' ** Insert at the first non-Null statement line.
                      lngZ = lngLineNum
                    Else
                      If Left$(strLine, 4) = "Dim " Or Left$(strLine, 7) = "Static " Or _
                         Left$(strLine, 6) = "Const " Then
                        lngDims = lngDims + 1&
                        If lngDims = arr_varTmpArr05(PROC_DIMS, lngElemP) Then
                          ' ** 2 lines after last Dim.
                          lngZ = lngLineNum + 2&
                        End If
                      End If
                    End If
                  Else
                    If arr_varTmpArr05(PROC_DIMS, lngElemP) = 0& Then
                      If InStr(strLine, "On Error") > 0 Then
                        ' ** Insert 2 lines after On Error statement.
                        lngZ = lngLineNum + 2&
                      End If
                    Else
                      If Left$(strLine, 4) = "Dim " Or Left$(strLine, 7) = "Static " Or _
                         Left$(strLine, 6) = "Const " Then
                        lngDims = lngDims + 1&
                        If lngDims = arr_varTmpArr05(PROC_DIMS, lngElemP) Then
                          ' ** Insert 2 lines after last Dim.
                          lngZ = lngLineNum + 2&
                        End If
                      End If
                    End If
                  End If
                End If
                If lngZ > 0& Then Exit For
              Next
              If lngZ > 0& Then
                intPos2 = intPos2 + 1
                .InsertLines lngZ, ""
                If arr_varMod(MOD_OBJ, lngElemM) = acForm Then
                  .InsertLines lngZ, "102     DispFrmName Me  ' ** Module Procedure: modFileUtilities."
                Else
                  .InsertLines lngZ, "102     DispRptName Me  ' ** Module Procedure: modFileUtilities."
                End If
              Else
                Debug.Print "'INSERTION POINT NOT FOUND! " & arr_varMod(MOD_NAME, lngElemM) & ".Form_Open()"
              End If
            End If
          End With
        End With
      End With
    End If
  Next

  If intPos2 > 0 Then
    Debug.Print "'DispRptName/DispFrmName " & intPos2
  Else
    'Debug.Print "'NO INSERTIONS: " & intPos2
  End If
End If

  Beep

  Set frm = Nothing
  Set vbp = Nothing
  Set vbc = Nothing
  Set cod = Nothing
  Set rptao = Nothing
  Set prj = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_Find_Proc = blnRetValx

End Function

Private Function VBA_IsSQL(cod As CodeModule, lngLineNum As Long) As Integer
' ** Called by:
' **   VBA_Find_Proc(), Above

  Const THIS_PROC As String = "VBA_IsSQL"

  Dim strLine As String
  Dim lngQuotes As Long, arr_varQuote() As Variant
  Dim lngSingles As Long, arr_varSingle() As Variant
  Dim lngQuoteCnt As Long
  Dim lngSQLTerms As Long, arr_varSQLTerm As Variant
  Dim intThisChk As Integer
  Dim lngPos1 As Long, lngPos2 As Long, lngRemPos As Long
  Dim blnEOL As Boolean, blnDoModSqlVars As Boolean
  Dim strTmp01 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngElemQ As Long
  Dim lngElems As Long, lngElemN As Long, lngElemV As Long, lngElemX As Long
  Dim intRetVal As Integer

  Static strProc As String
  Static lngVars As Long, arr_varVar() As Variant
  Static lngModVars As Long, arr_varModVar() As Variant

  Const QUOT_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const QUOT_OP  As Integer = 0
  Const QUOT_CL  As Integer = 1
  Const QUOT_CHK As Integer = 2

  Const SNG_ELEMS  As Integer = 1  ' ** Array's first-element UBound().
  Const SNG_OP     As Integer = 0
  Const SNG_INQUOT As Integer = 1

  Const VAR_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const VAR_NAME As Integer = 0

  intRetVal = SQL_NONE

  lngQuotes = 0&
  lngQuoteCnt = 0&
  ReDim arr_varQuote(QUOT_ELEMS, 0)

  lngSingles = 0&
  ReDim arr_varSingle(SNG_ELEMS, 0)

  With cod
    strLine = Trim$(.Lines(lngLineNum, 1))
    If .ProcOfLine(lngLineNum, vbext_pk_Proc) = vbNullString Then
      ' ** Declaration section.
      If lngLineNum <= .CountOfDeclarationLines Then
        If strProc <> "DeclarationSection" Then
          strProc = "DeclarationSection"
          lngModVars = 0&
          ReDim arr_varModVar(VAR_ELEMS, 0)
        End If
      Else
        Stop
      End If
    Else
      If strProc <> .ProcOfLine(lngLineNum, vbext_pk_Proc) Then
        strProc = .ProcOfLine(lngLineNum, vbext_pk_Proc)
        lngVars = 0&
        ReDim arr_varVar(VAR_ELEMS, 0)
      End If
    End If
    If blnModSqlVars = True Then
      ' ** This won't be True until after the first time a variable is found, below.
      If lngModVars = 0& Then
        blnDoModSqlVars = True
      Else
        blnDoModSqlVars = False
      End If
    Else
      lngModVars = 0&
      ReDim arr_varModVar(VAR_ELEMS, 0)
      blnDoModSqlVars = False
    End If
    If strLine <> vbNullString Then  ' ** Not a blank.
      If Left$(strLine, 1) <> "'" Then  ' ** Not a remark.

        ' ** 1. DoCmd.RunSQL or assignment to .SQL property.
        If InStr(strLine, ".SQL") > 0 Or InStr(strLine, ".RunSQL") > 0 Then
          intRetVal = SQL_HASCMD
        End If

        ' ** 2. Variable with SQL in name.
        If Left$(strLine, 4) = "Dim " Or Left$(strLine, 7) = "Static " Or Left$(strLine, 6) = "Const " Or _
            Left$(strLine, 7) = "Private" Or Left$(strLine, 6) = "Public" Then
          lngPos1 = InStr(strLine, "sql")
          If lngPos1 > 0 Then
            If intRetVal = SQL_NONE Then
              intRetVal = SQL_HASVRB
            Else
              intRetVal = intRetVal + SQL_HASVRB
            End If
            Do While lngPos1 > 0
              strTmp01 = vbNullString
              lngRemPos = 0&: blnEOL = False
              lngPos2 = InStr((lngPos1 + 1), strLine, " ")
              If lngPos2 = 0& Then
                blnEOL = True
                For lngX = (lngPos1 - 1&) To 0& Step -1&
                  If Mid$(strLine, lngX, 1) = " " Then
                    strTmp01 = Trim$(Mid$(strLine, (lngX + 1)))
                    Exit For
                  End If
                Next
              Else
                strTmp01 = Left$(strLine, (lngPos2 - 1&))
                lngRemPos = InStr(lngPos2, strLine, "'")  ' ** Entire remarked lines have already been eliminated.
                If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))  ' ** (And a remark is
                For lngX = (lngPos1 - 1&) To 0& Step -1&                                       ' ** always preceeded
                  If Mid$(strTmp01, lngX, 1) = " " Then                                         ' ** by a space.)
                    strTmp01 = Trim$(Mid$(strTmp01, (lngX + 1)))
                    Exit For
                  End If
                Next
              End If
              If strTmp01 <> vbNullString Then
                lngVars = lngVars + 1&
                lngElemV = lngVars - 1&
                ReDim Preserve arr_varVar(VAR_ELEMS, lngElemV)
                arr_varVar(VAR_NAME, lngElemV) = Trim$(strTmp01)
                ' ** blnModSqlVars won't be True until after the first time a
                ' ** variable is found, so blnDoModSqlVars will also be False.
                If blnDoModSqlVars = False And strProc = "DeclarationSection" Then
                  blnModSqlVars = True
                  blnDoModSqlVars = True
                End If
                If blnDoModSqlVars = True Then
                  lngModVars = lngModVars + 1&
                  lngElemX = lngModVars - 1&
                  ReDim Preserve arr_varModVar(VAR_ELEMS, lngElemX)
                  arr_varModVar(VAR_NAME, lngElemX) = Trim$(strTmp01)
                End If
              End If
              If blnEOL = True Then Exit Do
              lngPos1 = InStr((lngPos1 + 1), strLine, "sql")
              If lngRemPos > 0 And lngPos1 > lngRemPos Then Exit Do
            Loop
          End If
        End If

        ' ** 3. Use of SQL variable.
        If lngVars > 0& Then
          For lngX = 0& To (lngVars - 1&)
            If InStr(strLine, arr_varVar(VAR_NAME, lngX)) > 0 Then
              If intRetVal = SQL_NONE Then
                intRetVal = SQL_HASVCD
              Else
                intRetVal = intRetVal + SQL_HASVCD
              End If
              Exit For
            End If
          Next
        End If
        If lngModVars > 0& And blnModSqlVars = True Then
          For lngX = 0& To (lngModVars - 1&)
            If InStr(strLine, arr_varModVar(VAR_NAME, lngX)) > 0 Then
              If intRetVal = SQL_NONE Then
                intRetVal = SQL_HASVCD
              Else
                intRetVal = intRetVal + SQL_HASVCD
              End If
              Exit For
            End If
          Next
        End If

        ' ** 4. Presence of SQL term.
        lngPos1 = InStr(strLine, Chr(34))  ' ** "
        If lngPos1 > 0& Then
          ' ** String lines can't be broken within opening and closing quotes.

          ' ** Find the quotes: 1 open, 2 close, 3 open, 4 close, etc.
          lngPos1 = InStr(strLine, Chr(34))
          Do While lngPos1 > 0&
            If lngQuoteCnt Mod 2& = 0& Then
              lngQuoteCnt = lngQuoteCnt + 1&
              lngQuotes = lngQuotes + 1&
              lngElemQ = lngQuotes - 1&
              ReDim Preserve arr_varQuote(QUOT_ELEMS, lngElemQ)
              arr_varQuote(QUOT_OP, lngElemQ) = lngPos1
              arr_varQuote(QUOT_CL, lngElemQ) = CInt(0)
              arr_varQuote(QUOT_CHK, lngElemQ) = SQL_NONE
            Else
              lngQuoteCnt = lngQuoteCnt + 1&
              lngElemQ = lngQuotes - 1&
              arr_varQuote(QUOT_CL, lngElemQ) = lngPos1
            End If
            lngPos1 = InStr((lngPos1 + 1&), strLine, Chr(34))
          Loop

          ' ** Check for Remarks.
          lngPos1 = InStr(strLine, "'")
          If lngPos1 > 0& Then
            Do While lngPos1 > 0&
              ' ** See if it's within quotes.
              ' ** If it is, there could be any number of them within the
              ' ** quotes. Pairing isn't checkable since one could be the
              ' ** close from a previous line or the open for the next.
              lngSingles = lngSingles + 1&
              lngElemN = lngSingles - 1&
              ReDim Preserve arr_varSingle(SNG_ELEMS, lngElemN)
              arr_varSingle(SNG_OP, lngElemN) = lngPos1
              arr_varSingle(SNG_INQUOT, lngElemN) = CBool(False)
              For lngX = 0& To (lngQuotes - 1&)
                lngElemQ = lngX
                If arr_varQuote(QUOT_OP, lngElemQ) < lngPos1 And arr_varQuote(QUOT_CL, lngElemQ) > lngPos1 Then
                  ' ** It's within quotes.
                  arr_varSingle(SNG_INQUOT, lngElemN) = CBool(True)
                  Exit For
                End If
              Next
              lngPos1 = InStr((lngPos1 + 1&), strLine, "'")
            Loop
          End If

          ' ** If there are singles, see if any of them initiate a remark.
          lngRemPos = 0&
          If lngSingles > 0& Then
            ' ** The first one found not within quotes will start a remark.
            For lngX = 0& To (lngSingles - 1&)
              lngElemN = lngX
              If arr_varSingle(SNG_INQUOT, lngElemN) = False Then
                lngRemPos = arr_varSingle(SNG_OP, lngElemN)
                Exit For
              End If
            Next
          End If

          ' ** Load arr_varSQLTerm() array with SQL words to look for.
          arr_varSQLTerm = VBA_IsSQLT_TERMs  ' ** Function: Below.
          lngSQLTerms = (UBound(arr_varSQLTerm, 2) + 1)

          If InStr(strLine, "MsgBox ") > 0 Or InStr(strLine, "MsgBox(") > 0 Then
            ' ** Skip it!
          Else
            ' ** Look for SQL terms.
            intThisChk = SQL_NONE
            For lngX = 0& To (lngSQLTerms - 1&)
              lngElems = lngX

              intThisChk = SQL_NONE
              If arr_varSQLTerm(SQLT_PRIORITY, lngElems) = SQL_PRI1 Then
                ' ** Check for the highest priority (most obvious) SQL terms.
                lngPos1 = InStr(strLine, arr_varSQLTerm(SQLT_TERM, lngElems))
                If lngPos1 > 0& Then
                  ' ** OK, where is the SQL term relative to the various quotes.
                  For lngY = 0& To (lngQuotes - 1&)
                    lngElemQ = lngY
                    If arr_varQuote(QUOT_OP, lngElemQ) < lngPos1 And arr_varQuote(QUOT_CL, lngElemQ) > lngPos1 Then
                      ' ** It's within quotes.
                      If lngRemPos = 0& Then
                        ' ** No remarks.
                        arr_varQuote(QUOT_CHK, lngElemQ) = SQL_HASTRM
                      Else
                        If lngPos1 < lngRemPos Then
                          ' ** To the left of a remark.
                          arr_varQuote(QUOT_CHK, lngElemQ) = SQL_HASTRM
                        End If
                      End If
                    End If  ' ** In quotes.
                  Next  ' ** For each quote.
                  For lngY = 0& To (lngQuotes - 1&)
                    lngElemQ = lngY
                    If arr_varQuote(QUOT_CHK, lngElemQ) <> SQL_NONE Then
                      ' ** If this term appears at least once within a pair of quotes, that's good enough.
                      intThisChk = arr_varQuote(QUOT_CHK, lngElemQ)
                      Exit For
                    End If
                  Next
                End If  ' ** Word found.
              End If  ' ** Only priority 1 words.

              If intThisChk <> SQL_NONE Then
                ' ** If this term appears anywhere within quotes, that's good enough.
                Exit For
              ElseIf arr_varSQLTerm(SQLT_PRIORITY, lngElems) = SQL_PRI2 Then
                ' ** Check the 2nd priority (more easily confused) SQL terms.
'2's: SELECT, DELETE, FROM, WHERE, BETWEEN, VALUES, ASC, DESC
                lngPos1 = InStr(strLine, arr_varSQLTerm(SQLT_TERM, lngElems))
                If lngPos1 > 0& Then
                  ' ** OK, where is the SQL term relative to the various quotes.
                  For lngY = 0& To (lngQuotes - 1&)
                    lngElemQ = lngY
                    If arr_varQuote(QUOT_OP, lngElemQ) < lngPos1 And arr_varQuote(QUOT_CL, lngElemQ) > lngPos1 Then
                      ' ** It's within quotes.
                      lngPos2 = InStr(strLine, "MsgBox")
                      If lngPos2 = 0& Or (lngPos2 > 0& And lngPos2 > lngPos1) Then
                        If lngRemPos = 0& Then
                          ' ** No remarks.
                          arr_varQuote(QUOT_CHK, lngElemQ) = SQL_HASTRM
                        Else
                          If lngPos1 < lngRemPos Then
                            ' ** To the left of a remark.
                            arr_varQuote(QUOT_CHK, lngElemQ) = SQL_HASTRM
                          End If
                        End If
                      End If
                    End If  ' ** In quotes.
                  Next  ' ** For each quote.
                  For lngY = 0& To (lngQuotes - 1&)
                    lngElemQ = lngY
                    If arr_varQuote(QUOT_CHK, lngElemQ) <> SQL_NONE Then
                      ' ** If this term appears at least once within a pair of quotes, that's good enough.
                      intThisChk = arr_varQuote(QUOT_CHK, lngElemQ)
                      Exit For
                    End If
                  Next
                End If  ' ** Word found.

              End If  ' ** Priority 2 words.

              If intThisChk <> SQL_NONE Then
                ' ** If this term appears anywhere within quotes, that's good enough.
                Exit For
              End If

            Next  ' ** For each SQL word.
          End If  ' ** Not a MsgBox.

          If intThisChk <> SQL_NONE Then
            If intRetVal = SQL_NONE Then
              intRetVal = intThisChk
            Else
              intRetVal = intRetVal + intThisChk
            End If
          End If

        End If

      End If
    End If
  End With

  VBA_IsSQL = intRetVal

End Function

Private Function VBA_IsSQLT_TERMs() As Variant
' ** SQL reserved words, with priority ranking.
' ** Called by:
' **   VBA_IsSQL(), Above

  Const THIS_PROC As String = "VBA_IsSQLT_TERMs"

  Dim lngSQLTerms As Long, arr_varSQLTerm() As Variant
  Dim lngElems As Long

  lngSQLTerms = 0&
  ReDim arr_varSQLTerm(SQLT_ELEMS, 0)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "DISTINCTROW"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "ORDER BY"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "GROUP BY"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "LEFT JOIN"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "INNER JOIN"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "RIGHT JOIN"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "INSERT INTO"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "UNION"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "SELECT DISTINCT"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(1)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "SELECT"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "DELETE"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "FROM"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "WHERE"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "BETWEEN"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "VALUES"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "ASC"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "DESC"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(2)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "UPDATE"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(3)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "SET"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(3)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "INTO"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(3)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "AS"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(3)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "ON"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(3)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "AND"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(3)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "OR"
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(3)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "Chr(34)"  ' "
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(4)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "Chr(35)"  ' #
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(4)

  lngSQLTerms = lngSQLTerms + 1&
  lngElems = (lngSQLTerms - 1&)
  ReDim Preserve arr_varSQLTerm(SQLT_ELEMS, lngElems)
  arr_varSQLTerm(SQLT_TERM, lngElems) = "Chr(38)"  ' &
  arr_varSQLTerm(SQLT_PRIORITY, lngElems) = CInt(4)

  VBA_IsSQLT_TERMs = arr_varSQLTerm

End Function

Public Function VBA_MsgBox_Parse(varInput As Variant, strInfo As String) As Variant
' ** Here to facilitate query documentation when zz_mod_ModuleDocFuncs not present.
' ** Called by:
' **   zz_qry_VBComponent_MsgBox_03

  Const THIS_PROC As String = "VBA_MsgBox_Parse"

  Dim strMsg As String, strSwitchTitle As String, strSwitch As String, strTitle1 As String
  Dim blnIsResponse As Boolean, blnIsIf As Boolean, strIfResponse As String
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp00 As String
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varInput) = False Then
    intPos1 = InStr(varInput, "MsgBox")
    If intPos1 > 0 Then
      intPos2 = InStr(varInput, "DEF_MSGBOX")
      If intPos2 > 0 Then
        If (intPos2 + 4) = intPos1 Then
          intPos1 = InStr((intPos1 + 1), varInput, "MsgBox")
        End If
      End If
    End If
    If intPos1 > 0 Then

      ' ** First, remove any remarks on the line.
      intPos2 = InStr(varInput, "' ** ")
      If intPos2 > 0 Then
        'MsgBox "The form must be complete to continue."  ' ** 3031: Not a valid password.
        strTmp00 = Trim(Left(varInput, (intPos2 - 1)))
      Else
        strTmp00 = Trim(varInput)
      End If

      blnIsResponse = False: blnIsIf = False: strIfResponse = vbNullString

      ' ** Then, get it into a standard format.
      If Left(strTmp00, 14) = "msgResponse = " Then
        'msgResponse = MsgBox("Account Number cannot have spaces!" & vbCrLf & "They have been removed.", vbExclamation + vbOKOnly, "Error")
        blnIsResponse = True
        strTmp00 = Trim(Mid(strTmp00, 15))
        If Mid$(strTmp00, 7, 1) = "(" Then
          strTmp00 = Trim$(Left$(strTmp00, 6) & " " & Mid$(strTmp00, 8))
          If Right$(strTmp00, 1) = ")" Then strTmp00 = Trim(Left$(strTmp00, (Len(strTmp00) - 1)))
        ElseIf Mid$(strTmp00, 8, 1) = "(" Then
          strTmp00 = Trim$(Left$(strTmp00, 7) & Mid$(strTmp00, 9))
          If Right$(strTmp00, 1) = ")" Then strTmp00 = Trim(Left$(strTmp00, (Len(strTmp00) - 1)))
        End If
      ElseIf Left(strTmp00, 3) = "If " Then
        'If MsgBox("Do you want to add this new account?", vbQuestion + vbYesNo + vbDefaultButton1, "Save New Account") = vbYes Then
        blnIsIf = True
        strTmp00 = Trim(Mid(strTmp00, 4))
        If Right$(strTmp00, 5) = " Then" Then
          strTmp00 = Trim$(Left$(strTmp00, (Len(strTmp00) - 4)))  ' ** Strip off 'Then'.
        End If
        If Right$(strTmp00, 4) = "vbOK" Or Right$(strTmp00, 4) = "vbNo" Then
          strIfResponse = Right$(strTmp00, 4)
          strTmp00 = Trim(Left$(strTmp00, (Len(strTmp00) - 4)))   ' ** Still has operator.
        ElseIf Right$(strTmp00, 5) = "vbYes" Then
          strIfResponse = Right$(strTmp00, 5)
          strTmp00 = Trim(Left$(strTmp00, (Len(strTmp00) - 5)))   ' ** Still has operator.
        ElseIf Right$(strTmp00, 8) = "vbCancel" Then
          strIfResponse = Right$(strTmp00, 8)
          strTmp00 = Trim(Left$(strTmp00, (Len(strTmp00) - 8)))   ' ** Still has operator.
        End If
        If Mid$(strTmp00, (Len(strTmp00) - 1), 1) = " " Then      ' ** =.
          strTmp00 = Trim$(Left$(strTmp00, (Len(strTmp00) - 1)))
        ElseIf Mid$(strTmp00, (Len(strTmp00) - 2), 1) = " " Then  ' ** <>.
          strTmp00 = Trim$(Left$(strTmp00, (Len(strTmp00) - 2)))
        End If
        If Mid$(strTmp00, 7, 1) = "(" Then
          strTmp00 = Trim$(Left$(strTmp00, 6) & " " & Mid$(strTmp00, 8))
          If Right$(strTmp00, 1) = ")" Then strTmp00 = Trim(Left$(strTmp00, (Len(strTmp00) - 1)))
        ElseIf Mid$(strTmp00, 8, 1) = "(" Then
          strTmp00 = Trim$(Left$(strTmp00, 7) & Mid$(strTmp00, 9))
          If Right$(strTmp00, 1) = ")" Then strTmp00 = Trim(Left$(strTmp00, (Len(strTmp00) - 1)))
        End If
      End If

      strMsg = vbNullString: strSwitchTitle = vbNullString: strSwitch = vbNullString: strTitle1 = vbNullString
      If InStr(strTmp00, ", vbInfo") > 0 Then
        strMsg = Trim$(Left$(strTmp00, (InStr(strTmp00, ", vbInfo") - 1)))
        strSwitchTitle = Trim$(Mid$(strTmp00, (InStr(strTmp00, ", vbInfo") + 1)))
      ElseIf InStr(strTmp00, ", vbQues") > 0 Then
        strMsg = Trim$(Left$(strTmp00, (InStr(strTmp00, ", vbQues") - 1)))
        strSwitchTitle = Trim$(Mid$(strTmp00, (InStr(strTmp00, ", vbQues") + 1)))
      ElseIf InStr(strTmp00, ", vbExcl") > 0 Then
        strMsg = Trim$(Left$(strTmp00, (InStr(strTmp00, ", vbExcl") - 1)))
        strSwitchTitle = Trim$(Mid$(strTmp00, (InStr(strTmp00, ", vbExcl") + 1)))
      ElseIf InStr(strTmp00, ", vbCrit") > 0 Then
        strMsg = Trim$(Left$(strTmp00, (InStr(strTmp00, ", vbCrit") - 1)))
        strSwitchTitle = Trim$(Mid$(strTmp00, (InStr(strTmp00, ", vbCrit") + 1)))
      Else
        'MsgBox strMsg, intStyle, strTitle11
        intPos1 = InStr(strTmp00, ",")
        If intPos1 > 0 Then
          strMsg = Trim$(Left$(strTmp00, (intPos1 - 1)))
          strSwitchTitle = Trim$(Mid$(strTmp00, (intPos1 + 1)))
        Else
          strMsg = strTmp00
        End If
      End If
      If strSwitchTitle <> vbNullString Then
        intPos1 = InStr(strSwitchTitle, ",")
        strSwitch = Trim$(Left$(strSwitchTitle, (intPos1 - 1)))
        strTitle1 = Trim$(Mid$(strSwitchTitle, (intPos1 + 1)))
      End If

      If InStr(strSwitch, "&") > 0 Then
        Debug.Print "'MSGBOX SWITCH AMPERSAND!  '" & strSwitch & "'"
      End If

      ' ** Now parse it into its components.
      Select Case strInfo
      Case "MSG"
'vbcommsg_text
        varRetVal = strMsg

      Case "SW1"
'vbcommsg_sw1
        If strSwitch <> vbNullString Then
          intPos1 = InStr(strSwitch, "+")
          If intPos1 > 0 Then
            strTmp00 = Trim$(Left$(strSwitch, (intPos1 - 1)))  ' ** Strip 2nd switch.
            If strTmp00 <> vbNullString Then
              varRetVal = strTmp00
            End If
          Else
            varRetVal = strSwitch
          End If
        End If

      Case "SW2"
'vbcommsg_sw2
        If strSwitch <> vbNullString Then
          intPos1 = InStr(strSwitch, "+")
          If intPos1 > 0 Then
            strTmp00 = Trim$(Mid$(strSwitch, (intPos1 + 1)))  ' ** Strip 1st switch.
            intPos1 = InStr(strTmp00, "+")
            If intPos1 > 0 Then
              strTmp00 = Trim$(Left$(strTmp00, (intPos1 - 1)))  ' ** Strip 3rd switch.
              If strTmp00 <> vbNullString Then
                varRetVal = strTmp00
              End If
            Else
              If strTmp00 <> vbNullString Then
                varRetVal = strTmp00
              End If
            End If
          Else
            ' ** Let it return Null.
          End If
        End If

      Case "SW3"
'vbcommsg_sw3
        If strSwitch <> vbNullString Then
          intPos1 = InStr(strSwitch, "+")
          If intPos1 > 0 Then
            strTmp00 = Trim$(Mid$(strSwitch, (intPos1 + 1)))  ' ** Strip 1st switch.
            intPos1 = InStr(strTmp00, "+")
            If intPos1 > 0 Then
              strTmp00 = Trim$(Mid$(strTmp00, (intPos1 + 1)))  ' ** Strip 2nd switch.
              intPos1 = InStr(strTmp00, "+")
              If intPos1 > 0 Then
                strTmp00 = Trim$(Left$(strTmp00, (intPos1 - 1)))  ' ** Strip 4th switch.
                If strTmp00 <> vbNullString Then
                  varRetVal = strTmp00
                End If
              Else
                If strTmp00 <> vbNullString Then
                  varRetVal = strTmp00
                End If
              End If
            Else
              ' ** Let it return Null.
            End If
          Else
            ' ** Let it return Null.
          End If
        End If

      Case "SW4"
'vbcommsg_sw4
        If strSwitch <> vbNullString Then
          intPos1 = InStr(strSwitch, "+")
          If intPos1 > 0 Then
            strTmp00 = Trim$(Mid$(strSwitch, (intPos1 + 1)))  ' ** Strip 1st switch.
            intPos1 = InStr(strTmp00, "+")
            If intPos1 > 0 Then
              strTmp00 = Trim$(Mid$(strTmp00, (intPos1 + 1)))  ' ** Strip 2nd switch.
              intPos1 = InStr(strTmp00, "+")
              If intPos1 > 0 Then
                strTmp00 = Trim$(Left$(strTmp00, (intPos1 + 1)))  ' ** Strip 3rd switch.
                If strTmp00 <> vbNullString Then
                  varRetVal = Trim$(Left$(strTmp00, (intPos1 + 1)))
                End If
              Else
                ' ** Let it return Null.
              End If
            Else
              ' ** Let it return Null.
            End If
          Else
            ' ** Let it return Null.
          End If
        End If

      Case "TITLE"
'vbcommsg_title
        If strTitle1 <> vbNullString Then
          varRetVal = strTitle1
        End If

      Case "ISRESP"
'vbcommsg_isresp
        varRetVal = blnIsResponse

      Case "ISIF"
'vbcommsg_isif
        varRetVal = blnIsIf

      Case "IFRESP"
'vbcommsg_ifresp
        If strIfResponse <> vbNullString Then
          varRetVal = strIfResponse
        End If

      End Select

    End If
  End If

  VBA_MsgBox_Parse = varRetVal

End Function

Public Function VBA_MsgBox_TitleParens(varInput As Variant, lngMsgID As Long) As Boolean
' ** Here to facilitate query documentation when zz_mod_ModuleDocFuncs not present.
' ** Just determine whether title is completely within parens.
' ** Called by:
' **   zz_qry_VBComponent_MsgBox_20

  Const THIS_PROC As String = "VBA_MsgBox_TitleParens"

  Dim lngOPCnt As Long, lngCPCnt As Long
  Dim lngParens As Long, arr_varParen() As Variant
  Dim lngLen As Long
  Dim blnFound As Boolean
  Dim strTmp00 As String
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varParen()
  Const P_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const P_OPEN As Integer = 0
  Const P_OPOS As Integer = 1
  Const P_CLOS As Integer = 2
  Const P_CPOS As Integer = 3

  blnRetVal = False

  If IsNull(varInput) = False Then
    strTmp00 = Trim(varInput)
    If strTmp00 <> vbNullString Then
      If Left$(strTmp00, 1) = "(" And Right$(strTmp00, 1) = ")" Then

        lngParens = 0&
        ReDim arr_varParen(P_ELEMS, 0)
        lngLen = Len(strTmp00)
        lngOPCnt = 0&: lngCPCnt = 0&

        For lngX = 1& To lngLen
          If Mid$(strTmp00, lngX, 1) = "(" Then
            lngOPCnt = lngOPCnt + 1&
            lngParens = lngParens + 1&
            lngE = lngParens - 1&
            ReDim Preserve arr_varParen(P_ELEMS, lngE)
            arr_varParen(P_OPEN, lngE) = Mid$(strTmp00, lngX, 1)
            arr_varParen(P_OPOS, lngE) = lngX
            arr_varParen(P_CLOS, lngE) = Null
            arr_varParen(P_CPOS, lngE) = CLng(0)
          ElseIf Mid$(strTmp00, lngX, 1) = ")" Then
            lngCPCnt = lngCPCnt + 1&
            blnFound = False
            For lngY = (lngParens - 1&) To 0& Step -1&
              If arr_varParen(P_CPOS, lngY) = 0& Then
                blnFound = True
                arr_varParen(P_CLOS, lngY) = Mid$(strTmp00, lngX, 1)
                arr_varParen(P_CPOS, lngY) = lngX
                Exit For
              End If
            Next
            If blnFound = False Then
              Select Case varInput
              Case "(" & Chr(34) & "Delete Stored Relationship View" & Chr(34) & " & Space(20)))"
                ' ** MISMATCHED PARENS! vbcommsg_id: 42341  ("Delete Stored Relationship View" & Space(20)))
                ' ** Line continued from previous, all there.
              Case "(" & Chr(34) & "Delete Obsolete Queries" & Chr(34) & " & Space(40)))"
                ' ** MISMATCHED PARENS! vbcommsg_id: 42691  ("Delete Obsolete Queries" & Space(40)))
                ' ** Line continued from previous, all there.
              Case "(Left$((" & Chr(34) & "DEL: " & Chr(34) & " & arr_varDel(D_QNM, lngX)) & Space(60), 60)))"
                ' ** MISMATCHED PARENS! vbcommsg_id: 42692  (Left$(("DEL: " & arr_varDel(D_QNM, lngX)) & Space(60), 60)))
                ' ** Line continued from previous, all there.
              Case "(" & Chr(34) & "New Tables Found" & Chr(34) & " & Space(40)))"
                ' ** MISMATCHED PARENS! vbcommsg_id: 42702  ("New Tables Found" & Space(40)))
                ' ** Line continued from previous, all there.
              Case "(" & Chr(34) & "Add " & Chr(34) & " & arr_varTbl(T_NAM, lngX) & Space(40)))"
                ' ** MISMATCHED PARENS! vbcommsg_id: 42704  ("Add " & arr_varTbl(T_NAM, lngX) & Space(40)))
                ' ** Line continued from previous, all there.
              Case "(" & Chr(34) & "Delete Obsolete Tables" & Chr(34) & " & Space(40)))"
                ' ** MISMATCHED PARENS! vbcommsg_id: 42705  ("Delete Obsolete Tables" & Space(40)))
                ' ** Line continued from previous, all there.
              Case "(Left$((" & Chr(34) & "DEL: " & Chr(34) & " & arr_varDel(D_NAM, lngX)) & Space(60), 60)))"
                ' ** MISMATCHED PARENS! vbcommsg_id: 42706  (Left$(("DEL: " & arr_varDel(D_NAM, lngX)) & Space(60), 60)))
                ' ** Line continued from previous, all there.
              Case Else
                Debug.Print "'MISMATCHED PARENS! vbcommsg_id: " & CStr(lngMsgID) & "  " & varInput
              End Select
            End If
          End If
        Next
        ' ** If the opening one's partner shows up before the final close,
        ' ** then the last close doesn't belong to the opening one.
        If arr_varParen(P_OPOS, 0) = 1& And arr_varParen(P_CPOS, 0) = lngLen Then
          blnRetVal = True
        Else
          If arr_varParen(P_CPOS, 0) = 0& Then
            Debug.Print "'MISSING CLOSE! vbcommsg_id: " & CStr(lngMsgID) & "  " & varInput
          End If
        End If
      End If
    End If
  End If

  VBA_MsgBox_TitleParens = blnRetVal

End Function

Public Function VBA_MsgBox_TitleNum(varInput As Variant, strInfo As String) As Variant
' ** Here to facilitate query documentation when zz_mod_ModuleDocFuncs not present.
' ** Called by:
' **   zz_qry_VBComponent_MsgBox_24

  Const THIS_PROC As String = "VBA_MsgBox_TitleNum"

  Dim lngLen As Long, lngPos1 As Long
  Dim strTmp00 As String, strTmp01 As String, strTmp02 As String
  Dim lngX As Long
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varInput) = False Then
    strTmp00 = Trim(varInput)
    If strTmp00 <> vbNullString Then
      strTmp01 = Right$(strTmp00, 2)
      Select Case strTmp01
      Case ("0" & Chr(34)), ("1" & Chr(34)), ("2" & Chr(34)), ("3" & Chr(34)), ("4" & Chr(34)), _
          ("5" & Chr(34)), ("6" & Chr(34)), ("7" & Chr(34)), ("8" & Chr(34)), ("9" & Chr(34))
        lngLen = Len(strTmp00)
        strTmp01 = vbNullString
        For lngX = (lngLen - 1) To 1& Step -1&  ' ** Skip last quotes.
          If Mid$(strTmp00, lngX, 1) = " " Then
            strTmp01 = Trim$(Mid$(strTmp00, lngX))
            Exit For
          ElseIf Mid$(strTmp00, lngX, 1) = Chr(34) Then
            strTmp01 = Trim$(Mid$(strTmp00, lngX))
            Exit For
          End If
        Next
        Select Case strInfo
        Case "NUM"
          If strTmp01 <> vbNullString Then
            If Right$(strTmp01, 1) = Chr(34) Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
            If Left$(strTmp01, 1) = Chr(34) Then strTmp01 = Mid$(strTmp01, 2)
            If strTmp01 = "File1" Then strTmp01 = "1"
            If strTmp01 = "File2" Then strTmp01 = "2"
            varRetVal = strTmp01
          End If
        Case "TITLE"
          If strTmp01 <> vbNullString Then
            If strTmp01 = ("File1" & Chr(34)) Or strTmp01 = ("File2" & Chr(34)) Then
              varRetVal = (Chr(34) & "Add File")
            Else
              lngPos1 = InStr(strTmp00, strTmp01)
              strTmp01 = Trim$(Left$(strTmp00, (lngPos1 - 1)))
              If Right$(strTmp01, 1) = "&" Then strTmp01 = Trim$(Left$(strTmp01, (Len(strTmp01) - 1)))
              varRetVal = strTmp01
            End If
          Else
            varRetVal = strTmp00
          End If
        End Select
      Case Else
        ' ** Don't care.
      End Select
    End If
  End If

  VBA_MsgBox_TitleNum = varRetVal

End Function

Public Function VBA_MsgBox_CrLf(varInput As Variant) As Variant
' ** Here to facilitate query documentation when zz_mod_ModuleDocFuncs not present.
' ** Called by:
' **   zz_qry_VBComponent_MsgBox_29

  Const THIS_PROC As String = "VBA_MsgBox_CrLf"

  Dim intPos1 As Integer, lngQCnt As Long
  Dim strTmp00 As String
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varInput) = False Then
    strTmp00 = Trim(varInput)
    If strTmp00 <> vbNullString Then
      intPos1 = InStr(strTmp00, "& vbCrLf")
      If intPos1 > 0 Then
        Do While intPos1 > 0
          strTmp00 = Left$(strTmp00, (intPos1 - 1)) & "{nl}" & Mid$(strTmp00, (intPos1 + 8))
          intPos1 = InStr(strTmp00, "& vbCrLf")
        Loop
      End If
      If InStr(strTmp00, Chr(34)) > 0 Then
        lngQCnt = CharCnt(strTmp00, Chr(34))  ' ** Module Function: modStringFuncs.
        If lngQCnt = 1& Then
          intPos1 = InStr(strTmp00, Chr(34))
          If intPos1 = 1 Then
            strTmp00 = Trim$(Mid$(strTmp00, 2))
          ElseIf intPos1 = Len(strTmp00) Then
            strTmp00 = Trim$(Left$(strTmp00, (intPos1 - 1)))
          Else
            strTmp00 = Trim$(Left$(strTmp00, (intPos1 - 1)) & Mid$(strTmp00, (intPos1 + 1)))
          End If
        Else
          intPos1 = InStr(strTmp00, Chr(34) & " & " & Chr(34))
          If intPos1 > 0 Then
            Do While intPos1 > 0
              strTmp00 = Trim$(Left$(strTmp00, (intPos1 - 1)) & Mid$(strTmp00, (intPos1 + 5)))
              intPos1 = InStr(strTmp00, Chr(34) & " & " & Chr(34))
            Loop
          End If
          intPos1 = InStr(strTmp00, Chr(34) & " {nl} & " & Chr(34))
          If intPos1 > 0 Then
            Do While intPos1 > 0
              strTmp00 = Trim$(Left$(strTmp00, (intPos1 - 1)) & " {nl} " & Mid$(strTmp00, (intPos1 + 10)))
              intPos1 = InStr(strTmp00, Chr(34) & " {nl} & " & Chr(34))
            Loop
          End If
          intPos1 = InStr(strTmp00, Chr(34) & " {nl} {nl} & " & Chr(34))
          If intPos1 > 0 Then
            Do While intPos1 > 0
              strTmp00 = Left$(strTmp00, (intPos1 - 1)) & " {nl} {nl} " & Mid$(strTmp00, (intPos1 + 15))
              intPos1 = InStr(strTmp00, Chr(34) & " {nl} {nl} & " & Chr(34))
            Loop
          End If
          lngQCnt = CharCnt(strTmp00, Chr(34))  ' ** Module Function: modStringFuncs.
          If lngQCnt = 2 Then
            If Left$(strTmp00, 1) = Chr(34) And Right$(strTmp00, 1) = Chr(34) Then
              strTmp00 = Trim$(Left$(Mid$(strTmp00, 2), (Len(Mid$(strTmp00, 2)) - 1)))
            End If
          End If

        End If
      End If
      strTmp00 = Rem_Spaces(strTmp00)  ' ** Module Function: modStringFuncs.
      varRetVal = strTmp00
    End If
  End If

  VBA_MsgBox_CrLf = varRetVal

End Function

Public Function Qry_ParseFormRef(varInput As Variant) As String
' ** Since I've removed all the Form and Report references, this function is moot.
' ** Not currently called.

  Const THIS_PROC As String = "Qry_ParseFormRef"

  Dim strSQL As String
  Dim lngRefs As Long, arr_varRef() As Variant
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp01 As String
  Dim lngX As Long, lngE As Long
  Dim strRetVal As String

  Const R_ELEMS As Integer = 2  ' ** Array's first-element UBound().

  strRetVal = vbNullString

  If IsNull(varInput) = False Then

    lngRefs = 0&
    ReDim arr_varRef(R_ELEMS, 0)

    strSQL = Trim(varInput)
    strSQL = Rem_CRLF(strSQL)  ' ** Module Functions: modStringFuncs.

    intPos1 = InStr(strSQL, "[Forms]")
    If intPos1 > 0 Then
      Do While intPos1 > 0
        lngRefs = lngRefs + 1&
        lngE = lngRefs - 1&
        ReDim Preserve arr_varRef(R_ELEMS, lngE)
        strTmp01 = Mid$(strSQL, intPos1)
        intPos2 = InStr(strTmp01, " ")
        If intPos2 > 0 Then
          strTmp01 = Trim$(Left$(strTmp01, intPos2))
          arr_varRef(0, lngE) = strTmp01
          intPos1 = InStr((intPos1 + 1), strSQL, "[Forms]")
        Else
          arr_varRef(0, lngE) = strTmp01
          intPos1 = 0
        End If
      Loop
    End If

    intPos1 = InStr(strSQL, "[Reports]")
    If intPos1 > 0 Then
      Do While intPos1 > 0
        lngRefs = lngRefs + 1&
        lngE = lngRefs - 1&
        ReDim Preserve arr_varRef(R_ELEMS, lngE)
        strTmp01 = Mid$(strSQL, intPos1)
        intPos2 = InStr(strTmp01, " ")
        If intPos2 > 0 Then
          strTmp01 = Trim$(Left$(strTmp01, intPos2))
          arr_varRef(0, lngE) = strTmp01
          intPos1 = InStr((intPos1 + 1), strSQL, "[Reports]")
        Else
          arr_varRef(0, lngE) = strTmp01
          intPos1 = 0
        End If
      Loop
    End If

    If lngRefs > 0& Then

      '[Reports]![rptAccountSummary]![accountno]));
      '[Forms]![frmMap_Reinvest_DivInt_Price].[txtprice]*-1
      '[Forms]![map]!pershare*[ActiveAssets]![shareface]*-1,"Currency"),^[forms]![Map]!txtPrice,^[forms]![Map].txtdate^[Forms]!Map!Assetlist));
      For lngX = 0& To (lngRefs - 1&)
        intPos1 = InStr(arr_varRef(0, lngX), "*")
        If intPos1 > 0 Then
          strTmp01 = Mid$(arr_varRef(0, lngX), intPos1)
          arr_varRef(0, lngX) = Left$(arr_varRef(0, lngX), (intPos1 - 1))
          intPos1 = InStr(strTmp01, "[Forms]")
          If intPos1 > 0 Then
            lngRefs = lngRefs + 1&
            lngE = lngRefs - 1&
            ReDim Preserve arr_varRef(R_ELEMS, lngE)
            arr_varRef(0, lngE) = Mid$(strTmp01, intPos1)
          End If
          intPos1 = InStr(strTmp01, "[Forms]")
          If intPos1 > 0 Then
            lngRefs = lngRefs + 1&
            lngE = lngRefs - 1&
            ReDim Preserve arr_varRef(R_ELEMS, lngE)
            arr_varRef(0, lngE) = Mid$(strTmp01, intPos1)
          End If
        End If
        If Right$(arr_varRef(0, lngX), 1) = ";" Then
          arr_varRef(0, lngX) = Left$(arr_varRef(0, lngX), (Len(arr_varRef(0, lngX)) - 1))
        End If
        Do While Right$(arr_varRef(0, lngX), 1) = ")"
          arr_varRef(0, lngX) = Left$(arr_varRef(0, lngX), (Len(arr_varRef(0, lngX)) - 1))
        Loop
        If Right$(arr_varRef(0, lngX), 1) = "," Then
          arr_varRef(0, lngX) = Left$(arr_varRef(0, lngX), (Len(arr_varRef(0, lngX)) - 1))
        End If
      Next

      strTmp01 = vbNullString
      For lngX = 0& To (lngRefs - 1&)
        strTmp01 = strTmp01 & arr_varRef(0, lngX) & "^"
      Next
      If Right$(strTmp01, 1) = "^" Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
      strRetVal = strTmp01

    End If

  End If

  Qry_ParseFormRef = strRetVal

End Function

Public Function Qry_FldDoc() As Boolean
' ** Not currently called.

  Const THIS_PROC As String = "Qry_FldDoc"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef
  Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset, prp As Object
  Dim lngFlds As Long, arr_varFld() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngDupes As Long, arr_varDupe() As Variant
  Dim lngRecs1 As Long, lngRecs2 As Long
  Dim lngThisDbsID As Long
  Dim blnAdd As Boolean, blnDelAll As Boolean, blnFound As Boolean
  Dim varFlds As Variant
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  Const dbUnknown As Long = 0&

  ' ** Array: arr_varFld().
  Const F_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const F_ID As Integer = 0
  Const F_NAM As Integer = 1

  ' ** Array: arr_varDupe().
  Const D_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const D_QID  As Integer = 0
  Const D_QNAM As Integer = 1
  Const D_FNAM As Integer = 2

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    Debug.Print "'|.";
    DoEvents

    lngDupes = 0&
    ReDim arr_varDupe(D_ELEMS, 0)

    Set rst2 = .OpenRecordset("tblQuery_Field", dbOpenDynaset, dbConsistent)

    ' ** tblQuery, just needed fields, by specified [dbid].
    Set qdf1 = .QueryDefs("zz_qry_Query_Field_01")
    With qdf1.Parameters
      ![dbid] = lngThisDbsID
    End With
    Set rst1 = qdf1.OpenRecordset
    rst1.MoveLast
    lngRecs1 = rst1.RecordCount
    rst1.MoveFirst
    For lngX = 1& To lngRecs1

      lngFlds = 0&
      ReDim arr_varFld(F_ELEMS, 0)
      blnAdd = False

      With rst2
        If .BOF = True And .EOF = True Then
          blnAdd = True
        Else
          .MoveFirst
          .FindFirst "[qry_id] = " & CStr(rst1![qry_id])
          If .NoMatch = True Then
            blnAdd = True
          End If
        End If
      End With

      Set qdf2 = .QueryDefs(rst1![qry_name])
      With qdf2
        varFlds = .Fields.Count
        If IsNull(varFlds) = False Then
          If varFlds > 0 Then
            For lngY = 0 To (varFlds - 1)
              If blnAdd = False Then
                rst2.MoveFirst
                rst2.FindFirst "[qry_id] = " & CStr(rst1![qry_id]) & " And [qryfld_name] = '" & .Fields(lngY).Name & "'"
                Select Case rst2.NoMatch
                Case True
                  blnAdd = True
                Case False
                  lngFlds = lngFlds + 1&
                  lngE = lngFlds - 1&
                  ReDim Preserve arr_varFld(F_ELEMS, lngE)
                  arr_varFld(F_ID, lngE) = rst2![qryfld_id]
                  arr_varFld(F_NAM, lngE) = .Fields(lngY).Name
                  If .Fields(lngY).Type > 0 Then
                    If rst2![datatype_db_type] <> .Fields(lngY).Type Then
                      rst2.Edit
                      rst2![datatype_db_type] = .Fields(lngY).Type
                      rst2![qryfld_datemodified] = Now()
                      rst2.Update
                    End If
                  Else
                    If rst2![datatype_db_type] <> dbUnknown Then
                      rst2.Edit
                      rst2![datatype_db_type] = dbUnknown
                      rst2![qryfld_datemodified] = Now()
                      rst2.Update
                    End If
                  End If
On Error Resume Next
                  Set prp = .Fields(lngY).Properties("DisplayControl")
                  If ERR.Number = 0 Then
On Error GoTo 0
                    Select Case IsNull(rst2![ctltype_type])
                    Case True
                      rst2.Edit
                      rst2![ctltype_type] = prp.Value
                      rst2![qryfld_datemodified] = Now()
                      rst2.Update
                    Case False
                      If rst2![ctltype_type] <> prp.Value Then
                        rst2.Edit
                        rst2![ctltype_type] = prp.Value
                        rst2![qryfld_datemodified] = Now()
                        rst2.Update
                      End If
                    End Select
                  Else
On Error GoTo 0
                    If IsNull(rst2![ctltype_type]) = False Then
                      rst2.Edit
                      rst2![ctltype_type] = Null
                      rst2![qryfld_datemodified] = Now()
                      rst2.Update
                    End If
                  End If
On Error Resume Next
                  Set prp = .Fields(lngY).Properties("Format")
                  If ERR.Number = 0 Then
On Error GoTo 0
                    Select Case IsNull(rst2![qryfld_format])
                    Case True
                      rst2.Edit
                      rst2![qryfld_format] = prp.Value
                      rst2![qryfld_datemodified] = Now()
                      rst2.Update
                    Case False
                      If rst2![qryfld_format] <> prp.Value Then
                        rst2.Edit
                        rst2![qryfld_format] = prp.Value
                        rst2![qryfld_datemodified] = Now()
                        rst2.Update
                      End If
                    End Select
                  Else
On Error GoTo 0
                    If IsNull(rst2![qryfld_format]) = False Then
                      rst2.Edit
                      rst2![qryfld_format] = Null
                      rst2![qryfld_datemodified] = Now()
                      rst2.Update
                    End If
                  End If
                End Select
              End If  ' ** blnAdd.
              If blnAdd = True Then
                rst2.AddNew
                rst2![qry_id] = rst1![qry_id]
                rst2![qryfld_name] = .Fields(lngY).Name  ' ** Takes up to 75 chars.
                rst2![datatype_db_type] = .Fields(lngY).Type
On Error Resume Next
                Set prp = .Fields(lngY).Properties("DisplayControl")
                If ERR.Number = 0 Then
On Error GoTo 0
                  rst2![ctltype_type] = prp.Value
                Else
On Error GoTo 0
                End If
On Error Resume Next
                Set prp = .Fields(lngY).Properties("Format")
                If ERR.Number = 0 Then
On Error GoTo 0
                  rst2![qryfld_format] = prp.Value
                Else
On Error GoTo 0
                End If
                rst2![qryfld_datemodified] = Now()
On Error Resume Next
                rst2.Update
                If ERR.Number = 0 Then
On Error GoTo 0
                  rst2.Bookmark = rst2.LastModified
                  lngFlds = lngFlds + 1&
                  lngE = lngFlds - 1&
                  ReDim Preserve arr_varFld(F_ELEMS, lngE)
                  arr_varFld(F_ID, lngE) = rst2![qryfld_id]
                  arr_varFld(F_NAM, lngE) = .Fields(lngY).Name
                Else
On Error GoTo 0
                  lngDupes = lngDupes + 1&
                  lngE = lngDupes - 1&
                  ReDim Preserve arr_varDupe(D_ELEMS, lngE)
                  arr_varDupe(D_QID, lngE) = rst1![qry_id]
                  arr_varDupe(D_QNAM, lngE) = rst1![qry_name]
                  arr_varDupe(D_FNAM, lngE) = .Fields(lngY).Name
                End If
              End If
            Next  ' ** varFlds: lngY.
          End If
        End If
      End With  ' ** qdf2.
      Set qdf2 = Nothing

      lngDels = 0&
      ReDim arr_varDel(0)
      blnDelAll = False

      ' ** tblQuery_Field, by specified [qid].
      Set qdf2 = .QueryDefs("zz_qry_Query_Field_03")
      With qdf2.Parameters
        ![qid] = rst1![qry_id]
      End With
      Set rst3 = qdf2.OpenRecordset
      With rst3
        If .BOF = True And .EOF = True Then
          ' ** Nothing to do.
        Else
          .MoveLast
          lngRecs2 = .RecordCount
          .MoveFirst
          If lngFlds > 0& Then
            For lngY = 1& To lngRecs2
              blnFound = False
              For lngZ = 0& To (lngFlds - 1&)
                If ![qryfld_id] = arr_varFld(F_ID, lngZ) Then
                  blnFound = True
                  Exit For
                End If
              Next  ' ** lngFlds: lngZ.
              If blnFound = False Then
                lngDels = lngDels + 1&
                ReDim Preserve arr_varDel(lngDels - 1&)
                arr_varDel(lngDels - 1&) = ![qryfld_id]
              End If
              If lngY < lngRecs2 Then .MoveNext
            Next  ' ** lngRecs2: lngY.
          Else
            ' ** Delete them all.
            blnDelAll = True
          End If
        End If
        .Close
      End With  ' ** rst3.
      Set rst3 = Nothing
      Set qdf2 = Nothing

      If lngDels > 0& Then
        For lngY = 0& To (lngDels - 1&)
          ' ** Delete tblQuery_Field, by specified [qfldid].
          Set qdf2 = .QueryDefs("zz_qry_Query_Field_05")
          With qdf2.Parameters
            ![qfldid] = arr_varDel(lngY)
          End With
          qdf2.Execute
          Set qdf2 = Nothing
        Next  ' ** lngDels: lngY.
      ElseIf blnDelAll = True Then
        ' ** Delete tblQuery_Field, by specified [qid].
        Set qdf2 = .QueryDefs("zz_qry_Query_Field_04")
        With qdf2.Parameters
          ![qid] = rst1![qry_id]
        End With
        qdf2.Execute
        Set qdf2 = Nothing
      End If

      If (lngX + 1&) Mod 100 = 0 And lngX <> 0& Then
        Debug.Print "| " & CStr(lngX + 1&) & " of " & CStr(lngRecs1)
        Debug.Print "'|";
      ElseIf (lngX + 1&) Mod 10 = 0 And lngX <> 0& Then
        Debug.Print "|";
      Else
        Debug.Print ".";
      End If
      DoEvents

      If lngX < lngRecs1 Then rst1.MoveNext
    Next  ' ** lngRecs1: lngX.
    rst1.Close
    rst2.Close

    .Close
  End With  ' ** dbs.

  If lngDupes > 0& Then
    Debug.Print "'DUPES!  " & CStr(lngDupes)
    For lngX = 0& To (lngDupes - 1&)
      Debug.Print "'FLD: " & arr_varDupe(D_FNAM, lngX) & "  IN  " & CStr(arr_varDupe(D_QID, lngX)) & "  " & arr_varDupe(D_QNAM, lngX)
      DoEvents
    Next
  End If

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set prp = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  Set qdf1 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  Qry_FldDoc = blnRetValx

End Function

Public Function Qry_CurrentAppName() As Boolean

  Const THIS_PROC As String = "Qry_CurrentAppName"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry As Variant
  Dim lngX As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_DID  As Integer = 0
  Const Q_DNAM As Integer = 1
  Const Q_QID  As Integer = 2
  Const Q_QNAM As Integer = 3
  Const Q_SQL  As Integer = 4
  Const Q_CAN  As Integer = 5
  Const Q_OLD  As Integer = 6
  Const Q_NEW  As Integer = 7

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    Set qdf = .QueryDefs("zz_qry_Query_04a")
    Set rst = qdf.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        Debug.Print "'NONE FOUND!"
      Else
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
        ' **     5       4     qry_sql           Q_SQL
        ' **     6       5     qry_curappnam     Q_CAN
        ' **     7       6     qry_curapp_old    Q_OLD
        ' **     8       7     qry_curapp_new    Q_NEW
        ' **
        ' ***************************************************
      End If
    End With
    Set rst = Nothing
    Set qdf = Nothing

    If lngQrys > 0& Then
      For lngX = 0& To (lngQrys - 1&)
        If QueryExists(CStr(arr_varQry(Q_QNAM, lngX))) = True Then  ' ** Module Function: modFileUtilities.
          Set qdf = .QueryDefs(arr_varQry(Q_QNAM, lngX))
          With qdf
            If .SQL = arr_varQry(Q_SQL, lngX) Then
              .SQL = StringReplace(CStr(arr_varQry(Q_SQL, lngX)), CStr(arr_varQry(Q_OLD, lngX)), CStr(arr_varQry(Q_NEW, lngX)))  ' ** Module Function: modStringFuncs.
            Else
              Debug.Print "'QRY DIFFERENT!  " & arr_varQry(Q_QNAM, lngX)
            End If
          End With
        End If
      Next  ' ** lngX.
    End If

    .Close
  End With

'QRY DIFFERENT!  qryAccountProfile_RelAccts_06
'QRY DIFFERENT!  qryBackupRestore_01
'QRY DIFFERENT!  qryPreferences_20x
'QRY DIFFERENT!  qryRelation_04
'DONE! Qry_CurrentAppName()

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Beep

  Qry_CurrentAppName = blnRetVal

End Function

Public Function Qry_TblChk1() As Boolean

  Const THIS_PROC As String = "Qry_TblChk1"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngFlds As Long
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngTblXs As Long, arr_varTblX() As Variant
  Dim lngAnoms As Long, arr_varAnom() As Variant
  Dim strSQL As String
  Dim lngQrys As Long
  Dim blnFound As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const T_TNAM  As Integer = 0
  Const T_ISQRY As Integer = 1
  Const T_FND   As Integer = 2

  ' ** Array: arr_varAnom().
  Const A_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const A_QNAM As Integer = 0
  Const A_TMP1 As Integer = 1

  blnRetValx = True

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)

  lngAnoms = 0&
  ReDim arr_varAnom(A_ELEMS, 0)

  Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.
  DoEvents

  Set dbs = CurrentDb
  With dbs
    lngQrys = .QueryDefs.Count
    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents
    Debug.Print "'|";
    lngTmp03 = 0&
    For Each qdf In .QueryDefs
      lngTmp03 = lngTmp03 + 1&
      lngTblXs = 0&
      ReDim arr_varTblX(T_ELEMS, 0)
      With qdf
        If Left(.Name, 1) <> "~" Then  ' ** Skip those pesky system queries.

          Select Case .Type
          Case dbQSelect
            lngFlds = .Fields.Count
            For lngX = 0& To (lngFlds - 1&)
              lngTblXs = lngTblXs + 1&
              lngE = lngTblXs - 1&
              ReDim Preserve arr_varTblX(T_ELEMS, lngE)
              arr_varTblX(T_TNAM, lngE) = .Fields(lngX).SourceTable
              arr_varTblX(T_ISQRY, lngE) = CBool(False)
              arr_varTblX(T_FND, lngE) = CBool(False)
            Next
          Case dbQDelete
            strSQL = .SQL
            If Left$(strSQL, 10) = "PARAMETERS" Then
              intPos1 = InStr(strSQL, ";")
              strSQL = Trim$(Mid$(strSQL, (intPos1 + 1)))
              If Left$(strSQL, 2) = vbCrLf Then strSQL = Trim$(Mid$(strSQL, 3))
            End If
            intPos1 = InStr(strSQL, "FROM ")
            strTmp01 = Trim$(Mid$(strSQL, (intPos1 + 4)))
            intPos1 = InStr(strTmp01, "WHERE ")
            If intPos1 > 0 Then
              strTmp01 = Left$(strTmp01, (intPos1 - 1))
              strTmp01 = Rem_CRLF(strTmp01)  ' ** Module Function: modStringFuncs.
            End If
            intPos1 = InStr(strTmp01, ";")
            If intPos1 > 0 Then strTmp01 = Left$(strTmp01, (intPos1 - 1))
            strTmp01 = Trim$(strTmp01)
            lngTblXs = lngTblXs + 1&
            lngE = lngTblXs - 1&
            ReDim Preserve arr_varTblX(T_ELEMS, lngE)
            arr_varTblX(T_TNAM, lngE) = strTmp01
            arr_varTblX(T_ISQRY, lngE) = CBool(False)
            arr_varTblX(T_FND, lngE) = CBool(False)
          Case dbQUpdate
            strSQL = .SQL
            If Left$(strSQL, 10) = "PARAMETERS" Then
              intPos1 = InStr(strSQL, ";")
              strSQL = Trim$(Mid$(strSQL, (intPos1 + 1)))
              If Left$(strSQL, 2) = vbCrLf Then strSQL = Trim$(Mid$(strSQL, 3))
            End If
            intPos1 = InStr(strSQL, "UPDATE ")
            strTmp01 = Trim$(Mid$(strSQL, (intPos1 + 6)))
            intPos1 = InStr(strTmp01, " SET ")
            strTmp01 = Trim$(Left$(strTmp01, intPos1))
            lngTblXs = lngTblXs + 1&
            lngE = lngTblXs - 1&
            ReDim Preserve arr_varTblX(T_ELEMS, lngE)
            arr_varTblX(T_TNAM, lngE) = strTmp01
            arr_varTblX(T_ISQRY, lngE) = CBool(False)
            arr_varTblX(T_FND, lngE) = CBool(False)
          Case dbQAppend
            strSQL = .SQL
            If Left$(strSQL, 10) = "PARAMETERS" Then
              intPos1 = InStr(strSQL, ";")
              strSQL = Trim$(Mid$(strSQL, (intPos1 + 1)))
              If Left$(strSQL, 2) = vbCrLf Then strSQL = Trim$(Mid$(strSQL, 3))
            End If
            intPos1 = InStr(strSQL, "(")
            strTmp01 = Trim$(Left$(strSQL, (intPos1 - 1)))
            intPos1 = InStr(strTmp01, " INTO ")
            strTmp01 = Trim$(Mid$(strTmp01, (intPos1 + 5)))
            lngTblXs = lngTblXs + 1&
            lngE = lngTblXs - 1&
            ReDim Preserve arr_varTblX(T_ELEMS, lngE)
            arr_varTblX(T_TNAM, lngE) = strTmp01
            arr_varTblX(T_ISQRY, lngE) = CBool(False)
            arr_varTblX(T_FND, lngE) = CBool(False)
            intPos1 = InStr(strSQL, "FROM ")
            strTmp01 = Trim$(Mid$(strSQL, (intPos1 + 4)))
            intPos1 = InStr(strTmp01, "WHERE ")
            If intPos1 > 0 Then strTmp01 = Left$(strTmp01, (intPos1 - 1))
            strTmp01 = Rem_CRLF(strTmp01)  ' ** Module Function: modStringFuncs.
            intPos1 = InStr(strTmp01, "GROUP BY ")
            If intPos1 > 0 Then strTmp01 = Left$(strTmp01, (intPos1 - 1))
            intPos1 = InStr(strTmp01, "HAVING ")
            If intPos1 > 0 Then strTmp01 = Left$(strTmp01, (intPos1 - 1))
            intPos1 = InStr(strTmp01, "ORDER BY ")
            If intPos1 > 0 Then strTmp01 = Left$(strTmp01, (intPos1 - 1))
            intPos1 = InStr(strTmp01, ";")
            If intPos1 > 0 Then strTmp01 = Left$(strTmp01, (intPos1 - 1))
            strTmp01 = Trim$(strTmp01)
            intPos1 = InStr(strTmp01, " ")
            If intPos1 > 0 Then
              strTmp02 = Trim$(Left$(strTmp01, intPos1))
              If Right$(strTmp02, 1) = "," Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
              strTmp01 = Trim$(Mid$(strTmp01, intPos1))
              lngTblXs = lngTblXs + 1&
              lngE = lngTblXs - 1&
              ReDim Preserve arr_varTblX(T_ELEMS, lngE)
              arr_varTblX(T_TNAM, lngE) = strTmp02
              arr_varTblX(T_ISQRY, lngE) = CBool(False)
              arr_varTblX(T_FND, lngE) = CBool(False)
              intPos1 = InStr(strTmp01, " JOIN ")
              If intPos1 > 0 Then
                Do While intPos1 > 0
                  strTmp01 = Trim$(Mid$(strTmp01, (intPos1 + 5)))
                  intPos1 = InStr(strTmp01, " ")
                  strTmp02 = Trim$(Left$(strTmp01, intPos1))
                  strTmp01 = Trim$(Mid$(strTmp01, intPos1))
                  lngTblXs = lngTblXs + 1&
                  lngE = lngTblXs - 1&
                  ReDim Preserve arr_varTblX(T_ELEMS, lngE)
                  arr_varTblX(T_TNAM, lngE) = strTmp02
                  arr_varTblX(T_ISQRY, lngE) = CBool(False)
                  arr_varTblX(T_FND, lngE) = CBool(False)
                  intPos1 = InStr(strTmp01, " JOIN ")
                Loop
              Else
                'CARTESIAN
                intPos1 = InStr(strTmp01, ",")
                If intPos1 > 0 Then
                  strTmp02 = Left$(strTmp01, (intPos1 - 1))
                  strTmp01 = Trim$(Mid$(strTmp01, (intPos1 + 1)))
                  lngTblXs = lngTblXs + 1&
                  lngE = lngTblXs - 1&
                  ReDim Preserve arr_varTblX(T_ELEMS, lngE)
                  arr_varTblX(T_TNAM, lngE) = strTmp01
                  arr_varTblX(T_ISQRY, lngE) = CBool(False)
                  arr_varTblX(T_FND, lngE) = CBool(False)
                  lngTblXs = lngTblXs + 1&
                  lngE = lngTblXs - 1&
                  ReDim Preserve arr_varTblX(T_ELEMS, lngE)
                  arr_varTblX(T_TNAM, lngE) = strTmp02
                  arr_varTblX(T_ISQRY, lngE) = CBool(False)
                  arr_varTblX(T_FND, lngE) = CBool(False)
               Else
                  Select Case .Name
                  Case "qryAccountHide_84", "qryAccountMenu_12d", "qryAccountMenu_12e", "qryPricing_07", "qryState_11", _
                      "qryStatementBalance_05", "qryStatementBalance_16"
                    ' ** qryAccountHide_84  'qryAccountHide_84a'
                    ' ** qryAccountMenu_12d  'qryAccountMenu_12b'
                    ' ** qryAccountMenu_12e  'qryAccountMenu_12c'
                    ' ** qryPricing_07  'qryPricing_07a'
                    ' ** qryState_11  'qryState_06'
                    ' ** qryStatementBalance_05  'qryStatementBalance_06b'
                    ' ** qryStatementBalance_16  'qryStatementBalance_06b'
                    lngTblXs = lngTblXs + 1&
                    lngE = lngTblXs - 1&
                    ReDim Preserve arr_varTblX(T_ELEMS, lngE)
                    arr_varTblX(T_TNAM, lngE) = strTmp01
                    arr_varTblX(T_ISQRY, lngE) = CBool(False)
                    arr_varTblX(T_FND, lngE) = CBool(False)
                   Case Else
                    lngAnoms = lngAnoms + 1&
                    lngE = lngAnoms - 1&
                    ReDim Preserve arr_varAnom(A_ELEMS, lngE)
                    arr_varAnom(A_QNAM, lngE) = .Name
                    If Trim$(strTmp01) = vbNullString Then strTmp01 = "{empty}"
                    arr_varAnom(A_TMP1, lngE) = strTmp01
                  End Select
                End If
              End If
            Else
              lngTblXs = lngTblXs + 1&
              lngE = lngTblXs - 1&
              ReDim Preserve arr_varTblX(T_ELEMS, lngE)
              arr_varTblX(T_TNAM, lngE) = strTmp01
              arr_varTblX(T_ISQRY, lngE) = CBool(False)
              arr_varTblX(T_FND, lngE) = CBool(False)
            End If
          Case dbQDDL
            ' ** Don't check.
          Case dbQSetOperation
            strSQL = .SQL
            intPos1 = InStr(strSQL, "ORDER BY ")
            If intPos1 > 0 Then
              strSQL = Trim$(Left$(strSQL, (intPos1 - 1)))
              If Right$(strSQL, 2) = vbCrLf Then strSQL = Left$(strSQL, (Len(strSQL) - 2))
            End If
            intPos1 = InStr(strSQL, ";")
            If intPos1 > 0 Then
              strTmp01 = Left$(strSQL, (intPos1 - 1))
            End If
            intPos1 = InStr(strTmp01, vbCrLf)
            Do While intPos1 > 0
              strTmp02 = Left$(strTmp01, (intPos1 - 1))
              strTmp01 = Mid$(strTmp01, (intPos1 + 2))
              intPos2 = InStr(strTmp02, "TABLE ")
              strTmp02 = Trim$(Mid$(strTmp02, (intPos2 + 5)))
              lngTblXs = lngTblXs + 1&
              lngE = lngTblXs - 1&
              ReDim Preserve arr_varTblX(T_ELEMS, lngE)
              arr_varTblX(T_TNAM, lngE) = strTmp02
              arr_varTblX(T_ISQRY, lngE) = CBool(False)
              arr_varTblX(T_FND, lngE) = CBool(False)
              intPos1 = InStr(strTmp01, vbCrLf)
            Loop
            intPos2 = InStr(strTmp01, "TABLE ")
            strTmp01 = Trim$(Mid$(strTmp01, (intPos2 + 5)))
            lngTblXs = lngTblXs + 1&
            lngE = lngTblXs - 1&
            ReDim Preserve arr_varTblX(T_ELEMS, lngE)
            arr_varTblX(T_TNAM, lngE) = strTmp01
            arr_varTblX(T_ISQRY, lngE) = CBool(False)
            arr_varTblX(T_FND, lngE) = CBool(False)
          End Select

          If lngTblXs > 0& Then
            For lngX = 0& To (lngTblXs - 1&)
              Select Case IsNull(arr_varTblX(T_TNAM, lngX))
              Case True
                arr_varTblX(T_TNAM, lngX) = "{empty}"
                arr_varTblX(T_FND, lngX) = CBool(True)  ' ** So it'll be skipped.
              Case False
                If Trim(arr_varTblX(T_TNAM, lngX)) = vbNullString Then
                  arr_varTblX(T_TNAM, lngX) = "{empty}"
                  arr_varTblX(T_FND, lngX) = CBool(True)  ' ** So it'll be skipped.
                End If
              End Select
              arr_varTblX(T_TNAM, lngX) = Rem_Parens(arr_varTblX(T_TNAM, lngX))  ' ** Module Function: modStringFuncs.
              arr_varTblX(T_TNAM, lngX) = Rem_Brackets(arr_varTblX(T_TNAM, lngX))  ' ** Module Function: modStringFuncs.
            Next
            If lngTblXs > 1& Then
              For lngX = 0& To (lngTblXs - 1&)
                If arr_varTblX(T_FND, lngX) = False Then
                  For lngY = (lngX + 1&) To (lngTblXs - 1&)
                    If arr_varTblX(T_TNAM, lngY) = arr_varTblX(T_TNAM, lngX) Then
                      arr_varTblX(T_FND, lngY) = CBool(True)
                    End If
                  Next
                End If
              Next
              For lngX = 0& To (lngTblXs - 1&)
                If arr_varTblX(T_FND, lngX) = False Then
                  If Left(arr_varTblX(T_TNAM, lngX), 3) = "qry" Or _
                      Left(arr_varTblX(T_TNAM, lngX), 6) = "zz_qry" Or _
                      Left(arr_varTblX(T_TNAM, lngX), 7) = "zzz_qry" Then
                    arr_varTblX(T_ISQRY, lngX) = CBool(True)
                  End If
                End If
              Next
              For lngX = 0& To (lngTblXs - 1&)
                If arr_varTblX(T_FND, lngX) = False And arr_varTblX(T_ISQRY, lngX) = False Then
                  blnFound = False
                  For lngY = 0& To (lngTbls - 1&)
                    If arr_varTbl(T_TNAM, lngY) = arr_varTblX(T_TNAM, lngX) Then
                      blnFound = True
                      Exit For
                    End If
                  Next
                  If blnFound = False Then
                    lngTbls = lngTbls + 1&
                    lngE = lngTbls - 1&
                    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
                    arr_varTbl(T_TNAM, lngE) = arr_varTblX(T_TNAM, lngX)
                    arr_varTbl(T_ISQRY, lngE) = CBool(False)
                    arr_varTbl(T_FND, lngE) = CBool(False)
                  End If
                End If
              Next
            Else
              If Left(arr_varTblX(T_TNAM, 0), 3) <> "qry" Or _
                  Left(arr_varTblX(T_TNAM, 0), 6) = "zz_qry" Or _
                  Left(arr_varTblX(T_TNAM, 0), 7) = "zzz_qry" Then
                ' ** Nope!
              Else
                blnFound = False
                For lngY = 0& To (lngTbls - 1&)
                  If arr_varTbl(T_TNAM, lngY) = arr_varTblX(T_TNAM, 0) Then
                    blnFound = True
                    Exit For
                  End If
                Next
                If blnFound = False Then
                  lngTbls = lngTbls + 1&
                  lngE = lngTbls - 1&
                  ReDim Preserve arr_varTbl(T_ELEMS, lngE)
                  arr_varTbl(T_TNAM, lngE) = arr_varTblX(T_TNAM, 0)
                  arr_varTbl(T_ISQRY, lngE) = CBool(False)
                  arr_varTbl(T_FND, lngE) = CBool(False)
                End If
              End If
            End If
          End If  ' ** lngTblXs.

        End If
      End With  ' ** qry.

      If lngTmp03 Mod 100 = 0 Then
        Debug.Print "|  " & CStr(lngTmp03) & " OF " & CStr(lngQrys)
        Debug.Print "'|";
      ElseIf lngTmp03 Mod 10 = 0 Then
        Debug.Print "|";
      Else
        Debug.Print ".";
      End If
      DoEvents

    Next  ' ** qry.
    Set qdf = Nothing
    Debug.Print
    DoEvents

    If lngTbls > 0& Then

      lngTmp03 = 0&
      For lngX = 0& To (lngTbls - 1&)
        If Left(arr_varTbl(T_TNAM, lngX), 3) = "qry" Or _
            Left(arr_varTbl(T_TNAM, lngX), 6) = "zz_qry" Or _
            Left(arr_varTbl(T_TNAM, lngX), 7) = "zzz_qry" Then
          arr_varTbl(T_ISQRY, lngX) = CBool(True)
        Else
          lngTmp03 = lngTmp03 + 1&
        End If
      Next

      Debug.Print "'WRITING!"
      DoEvents

      ' ** Empty tblMark.
      Set qdf = .QueryDefs("qryTmp_Table_Empty_113_tblMark")
      qdf.Execute
      Set qdf = Nothing

      Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngTbls - 1&)
          If arr_varTbl(T_ISQRY, lngX) = False Then
            .AddNew
            ![unique_id] = (lngX + 1&)
            ![mark] = False
            ![Value] = IIf(IsNull(arr_varTbl(T_TNAM, lngX)) = True, "{empty}", IIf(Trim(arr_varTbl(T_TNAM, lngX)) = vbNullString, "{empty}", arr_varTbl(T_TNAM, lngX)))
            .Update
          End If
        Next
        .Close
      End With

    End If

    .Close
  End With  ' ** dbs.

  If lngAnoms > 0& Then
    Debug.Print "'ANOMS: " & CStr(lngAnoms)
    DoEvents
    For lngX = 0& To (lngAnoms - 1&)
      Debug.Print "'" & arr_varAnom(A_QNAM, lngX) & "  '" & arr_varAnom(A_TMP1, lngX) & "'"
      DoEvents
    Next
  End If

  Debug.Print "'QRYS: " & CStr(lngQrys)
  Debug.Print "'TBLS: " & CStr(lngTmp03)
  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Qry_TblChk1 = blnRetValx

End Function

Public Function Qry_TblChk2() As Boolean

  Const THIS_PROC As String = "Qry_TblChk2"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngTbls As Long, arr_varTbl As Variant
  Dim lngNots As Long
  Dim lngX As Long

  ' ** Array: arr_varTbl().
  Const T_UID  As Integer = 0
  Const T_FND  As Integer = 1
  Const T_TNAM As Integer = 2

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_qry_Query_50 (tblMark), sorted.
    Set qdf = .QueryDefs("zz_qry_Query_51")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngTbls = .RecordCount
      .MoveFirst
      arr_varTbl = .GetRows(lngTbls)
      ' **********************************************
      ' ** arr_varTbl()
      ' **
      ' **   Field  Element  Name         Constant
      ' **   =====  =======  ===========  ==========
      ' **     1       0     unique_id    T_UID
      ' **     2       1     mark         T_FND
      ' **     3       2     value        T_TNAM
      ' **
      ' **********************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With
  Set dbs = Nothing

  For lngX = 0& To (lngTbls - 1&)
    arr_varTbl(T_FND, lngX) = TableExists(CStr(arr_varTbl(T_TNAM, lngX)))  ' ** Module Function: modFileUtilities.
  Next

  lngNots = 0&
  For lngX = 0& To (lngTbls - 1&)
    If arr_varTbl(T_FND, lngX) = False Then
      lngNots = lngNots + 1&
      Debug.Print "'NOT FOUND!  " & arr_varTbl(T_TNAM, lngX)
    End If
  Next

  If lngNots > 0& Then
    Debug.Print "'MISSING CNT: " & CStr(lngNots)
  Else
    Debug.Print "'ALL HERE!"
  End If

  Debug.Print "'DONE!  " & THIS_PROC & "()"
'QUERIES WITH THESE WILL ERROR DURING DOCUMENTATION!
'NOT FOUND!  LedgerArchive_Backup
'NOT FOUND!  tblDatabase_Table_Link_tmp01
'NOT FOUND!  zz_tbl_m_TBL_tmp01
'MISSING CNT: 3
'DONE!  Qry_TblChk2()

'CREATED IN Qry_TmpTables()!
'QRY: 'qryBackupRestore_02' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_01' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_02' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_03' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_04' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_05' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_06' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_07' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_08' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_09' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_10' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_11' LedgerArchive_Backup
'QRY: 'qryBackupRestore_03_12' LedgerArchive_Backup
'QRY: 'qryBackupRestore_05' LedgerArchive_Backup
'DONE!
'CREATED IN Qry_TmpTables()!
'QRY: 'zz_qry_Database_Table_Link_35b' tblDatabase_Table_Link_tmp01
'DONE!
'CREATED IN Qry_TmpTables()!
'QRY: 'zz_qry_Database_Table_Link_35a' zz_tbl_m_TBL_tmp01
'DONE!

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Qry_TblChk2 = blnRetValx

End Function

Public Function Qry_ChkDocQrys2() As Boolean

  Const THIS_PROC As String = "Qry_ChkDocQrys2"

  blnRetValx = True

  If TableExists("LedgerArchive_Backup") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "LedgerArchive_Backup"
    DoEvents
  End If
  If TableExists("m_TBL1") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "m_TBL1"
    DoEvents
  End If
  If TableExists("m_TBL2") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "m_TBL2"
    DoEvents
  End If
  If TableExists("zz_tbl_sql_code_01") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "zz_tbl_sql_code_01"
    DoCmd.DeleteObject acTable, "zz_tbl_sql_code_02"
    DoCmd.DeleteObject acTable, "zz_tbl_sql_code_03"
    DoCmd.DeleteObject acTable, "zz_tbl_sql_code_04"
    DoCmd.DeleteObject acTable, "zz_tbl_sql_code_05"
    DoCmd.DeleteObject acTable, "zz_tbl_sql_code_06"
    DoEvents
  End If
  If TableExists("m_TBL_tmp01") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "m_TBL_tmp01"
    DoEvents
  End If
  If TableExists("tblMark_AutoNum2") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "tblMark_AutoNum2"
    DoEvents
  End If
  If TableExists("tblMark_AutoNum3") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "tblMark_AutoNum3"
    DoEvents
  End If
  If TableExists("zz_tbl_Database_Table_Link") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "zz_tbl_Database_Table_Link"
    DoEvents
  End If
  If TableExists("zz_tbl_Form_Doc") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "zz_tbl_Form_Doc"
    DoEvents
  End If
  If TableExists("zz_tbl_Form_Property") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "zz_tbl_Form_Property"
    DoEvents
  End If
  If TableExists("zz_tbl_Form_Property_Value") = False Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "zz_tbl_Form_Property_Value"
    DoEvents
  End If
  If TableExists("zz_tbl_VBComponent_KeyDown") = True Then  ' ** Module Function: modFileUtilities.
    DoCmd.DeleteObject acTable, "zz_tbl_VBComponent_KeyDown"
    DoEvents
  End If
  CurrentDb.TableDefs.Refresh

  Beep

Qry_ChkDocQrys2 = blnRetValx

End Function

Private Function Qry_ChkDocQrys(blnSkip As Boolean) As Boolean

  Const THIS_PROC As String = "Qry_ChkDocQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, prp As Object
  Dim lngQrys As Long, arr_varQry As Variant
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngQrysToImport As Long, lngQrysImported As Long
  Dim strPath1 As String, strFile1 As String, strPathFile1 As String, strPath2 As String, strFile2 As String, strPathFile2 As String
  Dim blnCreate As Boolean
  Dim strTmp01 As String
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varQry().
  Const Q_DID   As Integer = 0
  Const Q_VID   As Integer = 1
  Const Q_QID   As Integer = 2
  Const Q_DOCID As Integer = 3
  Const Q_VNAM  As Integer = 4
  Const Q_QNAM  As Integer = 5
  Const Q_TYP   As Integer = 6
  Const Q_DSC   As Integer = 7
  Const Q_SQL   As Integer = 8
  Const Q_FND   As Integer = 9
  Const Q_IMP   As Integer = 10

  ' ** Array: arr_varTbl()
  Const T_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const T_QNAM As Integer = 0
  Const T_DSC  As Integer = 1
  Const T_TNAM As Integer = 2

  blnRetValx = True

  If TableExists("LedgerArchive_Backup") = False Then  ' ** Module Function: modFileUtilities.
    DoCmd.CopyObject , "LedgerArchive_Backup", acTable, "tblTemplate_LedgerArchive"
    DoEvents
    CurrentDb.TableDefs.Refresh
  End If
  If TableExists("m_TBL1") = False Then  ' ** Module Function: modFileUtilities.
    DoCmd.CopyObject , "m_TBL1", acTable, "tblTemplate_m_TBL2"
    DoEvents
    CurrentDb.TableDefs.Refresh
  End If
  If TableExists("m_TBL2") = False Then  ' ** Module Function: modFileUtilities.
    DoCmd.CopyObject , "m_TBL2", acTable, "tblTemplate_m_TBL2"
    DoEvents
    CurrentDb.TableDefs.Refresh
  End If
  If TableExists("zz_tbl_Form_Property") = False Then  ' ** Module Function: modFileUtilities.
    Set dbs = CurrentDb
    With dbs
      Set qdf = .QueryDefs("zz_qry_System_60_01")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_60_02")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_60_03")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_60_04")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_60_05")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_50_03")
      qdf.Execute
      .Close
    End With
    Set dbs = Nothing
  End If
  If TableExists("zz_tbl_Form_Property_Value") = False Then  ' ** Module Function: modFileUtilities.
    Set dbs = CurrentDb
    With dbs
      Set qdf = .QueryDefs("zz_qry_System_61_01")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_61_02")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_61_03")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_61_04")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_61_05")
      qdf.Execute
      Set qdf = Nothing
      .Close
    End With
    Set dbs = Nothing
  End If
  If TableExists("zz_tbl_VBComponent_KeyDown") = False Then  ' ** Module Function: modFileUtilities.
    Set dbs = CurrentDb
    With dbs
      Set qdf = .QueryDefs("zz_qry_System_62_01")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_62_02")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_62_03")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_62_04")
      qdf.Execute
      Set qdf = Nothing
      Set qdf = .QueryDefs("zz_qry_System_62_05")
      qdf.Execute
      Set qdf = Nothing
      .Close
    End With
    Set dbs = Nothing
  End If
  If TableExists("zz_tbl_sql_code_01") = False Then  ' ** Module Function: modFileUtilities.
    ' ** zz_tbl_sql_code_01 - zz_tbl_sql_code_06.
    lngTbls = 0&
    ReDim arr_varTbl(T_ELEMS, 0)
    Set dbs = CurrentDb
    With dbs
      ' ** zz_qry_System_52_01 - zz_qry_System_59_08.
      For Each qdf In .QueryDefs
        With qdf
          If Left(.Name, 14) = "zz_qry_System_" Then
            strTmp01 = Mid(.Name, 15, 2)
            If CLng(strTmp01) >= 52& And CLng(strTmp01) <= 59& Then
              lngTbls = lngTbls + 1&
              lngE = lngTbls - 1&
              ReDim Preserve arr_varTbl(T_ELEMS, lngE)
              arr_varTbl(T_QNAM, lngE) = .Name
              arr_varTbl(T_DSC, lngE) = .Properties("Description")
              arr_varTbl(T_TNAM, lngE) = Null
            End If
          End If
        End With
      Next
      For lngX = 0& To (lngTbls - 1&)
        If Right(arr_varTbl(T_QNAM, lngX), 3) = "_01" Then
          strTmp01 = GetLastWord(arr_varTbl(T_DSC, lngX))
          If Right(strTmp01, 1) = "." Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
          arr_varTbl(T_TNAM, lngX) = strTmp01
        End If
      Next
      strTmp01 = vbNullString
      For lngX = 0& To (lngTbls - 1&)
        If IsNull(arr_varTbl(T_TNAM, lngX)) = False Then
          strTmp01 = arr_varTbl(T_TNAM, lngX)
        Else
          arr_varTbl(T_TNAM, lngX) = strTmp01
        End If
      Next
      For lngX = 0& To (lngTbls - 1&)
        If strTmp01 <> Mid(arr_varTbl(T_QNAM, lngX), 15, 2) Then  ' ** zz_qry_System_52_01.
          blnCreate = False
          strTmp01 = Mid(arr_varTbl(T_QNAM, lngX), 15, 2)
          If TableExists(CStr(arr_varTbl(T_TNAM, lngX))) = False Then  ' ** Module Function: modFileUtilities.
            blnCreate = True
          End If
          If blnCreate = True Then
            Set qdf = .QueryDefs(arr_varTbl(T_QNAM, lngX))
            qdf.Execute
            Set qdf = Nothing
          End If
        End If
      Next
      .Close
    End With
    Set dbs = Nothing
    Beep
  End If

  If blnSkip = False Then

    strPath1 = "C:\Program Files\Delta Data\Trust Accountant"
    strFile1 = "TrstXAdm - Copy (6).mdb"
    strPathFile1 = strPath1 & LNK_SEP & strFile1

    strPath2 = "C:\VictorGCS_Clients\TrustAccountant\NewWorking"
    strFile2 = "Trust.mdb"
    strPathFile2 = strPath2 & LNK_SEP & strFile2

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
        ' ** arr_varQry()
        ' **
        ' **   Field  Element  Name               Constant
        ' **   =====  =======  =================  ==========
        ' **     1       0     dbs_id             Q_DID
        ' **     2       1     vbcom_id           Q_VID
        ' **     3       2     qry_id             Q_QID
        ' **     4       3     qrydoc_id          Q_DOCID
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
      End With  ' ** rst.
      Set rst = Nothing
      Set qdf = Nothing
      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

    For lngX = 0& To (lngQrys - 1&)
      If QueryExists(CStr(arr_varQry(Q_QNAM, lngX))) = True Then
        arr_varQry(Q_FND, lngX) = CBool(True)
      End If
    Next  ' ** lngX.

    lngQrysToImport = 0&
    For lngX = 0& To (lngQrys - 1&)
      If arr_varQry(Q_FND, lngX) = False Then
        lngQrysToImport = lngQrysToImport + 1&
      End If
    Next  ' ** lngX.

    Debug.Print "'QRYS TO IMPORT: " & CStr(lngQrysToImport)
    DoEvents

    If lngQrysToImport > 0& Then
      Set dbs = CurrentDb
      With dbs
        lngQrysImported = 0&
        For lngX = 0& To (lngQrys - 1&)
          If arr_varQry(Q_FND, lngX) = False Then
            Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM, lngX), arr_varQry(Q_SQL, lngX))
            With qdf
              Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC, lngX))
On Error Resume Next
              .Properties.Append prp
              If ERR.Number <> 0 Then
On Error GoTo 0
                .Properties("Description") = arr_varQry(Q_DSC, lngX)
              Else
On Error GoTo 0
              End If
              lngQrysImported = lngQrysImported + 1&
            End With
            Set qdf = Nothing
          End If
        Next  ' ** lngX.
        .QueryDefs.Refresh
        .Close
      End With
    End If

    Debug.Print "'QRYS IMPORTED: " & CStr(lngQrysImported)
    DoEvents

    Debug.Print "'DONE!  " & THIS_PROC & "()"
    DoEvents

    Beep

  End If  ' ** blnSkip.

  Set prp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Qry_ChkDocQrys = blnRetValx

End Function

Public Function Qry_ChkQryExist() As Boolean

  Const THIS_PROC As String = "Qry_ChkQryExist"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys1 As Long, arr_varQry1() As Variant
  Dim lngQrys2 As Long, arr_varQry2 As Variant
  Dim lngSysQrys As Long, lngQrysNotFound As Long, lngQrysNotDocumented As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varQry1().
  Const Q1_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const Q1_QID   As Integer = 0
  Const Q1_QNAM  As Integer = 1
  Const Q1_EXIST As Integer = 2

  ' ** Array: arr_varQry2().
  Const Q2_DID   As Integer = 0  'dbs_id
  Const Q2_DNAM  As Integer = 1  'dbs_name
  Const Q2_QID   As Integer = 2  'qry_id
  Const Q2_QNAM  As Integer = 3  'qry_name
  Const Q2_EXIST As Integer = 4  'qry_exist

  blnRetValx = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    lngQrys1 = 0&
    ReDim arr_varQry1(Q1_ELEMS, 0)

    lngSysQrys = 0&
    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, 1) <> "~" Then
          lngQrys1 = lngQrys1 + 1&
          lngE = lngQrys1 - 1&
          ReDim Preserve arr_varQry1(Q1_ELEMS, lngE)
          ' **********************************************
          ' ** Array: arr_varQry1()
          ' **
          ' **   Field  Element  Name         Constant
          ' **   =====  =======  ===========  ==========
          ' **     1       0     qry_id       Q1_QID
          ' **     2       1     qry_name     Q1_QNAM
          ' **     3       2     qry_exist    Q1_EXIST
          ' **
          ' **********************************************
          arr_varQry1(Q1_QID, lngE) = Null
          arr_varQry1(Q1_QNAM, lngE) = .Name
          arr_varQry1(Q1_EXIST, lngE) = CBool(False)
        Else
          lngSysQrys = lngSysQrys + 1&
        End If
      End With
    Next
    Set qdf = Nothing

    Debug.Print "'QRYS1:  " & CStr(lngQrys1)
    Debug.Print "'SYS QRYS:  " & CStr(lngSysQrys)
    DoEvents

    ' ** tblQuery, just Trust.mdb.
    Set qdf = .QueryDefs("zz_qry_Query_62")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngQrys2 = .RecordCount
      .MoveFirst
      arr_varQry2 = .GetRows(lngQrys2)
      ' **********************************************
      ' ** Array: arr_varQry2()
      ' **
      ' **   Field  Element  Name         Constant
      ' **   =====  =======  ===========  ==========
      ' **     1       0     dbs_id       Q2_DID
      ' **     2       1     dbs_name     Q2_DNAM
      ' **     3       2     qry_id       Q2_QID
      ' **     4       3     qry_name     Q2_QNAM
      ' **     5       4     qry_exist    Q2_EXIST
      ' **
      ' **********************************************
    End With
    Set rst = Nothing
    Set qdf = Nothing

    .Close
  End With

  Debug.Print "'QRYS2:  " & CStr(lngQrys2)
  DoEvents

  For lngX = 0& To (lngQrys2 - 1&)
    For lngY = 0& To (lngQrys1 - 1&)
      If arr_varQry1(Q1_EXIST, lngY) = False Then
        If arr_varQry1(Q1_QNAM, lngY) = arr_varQry2(Q2_QNAM, lngX) Then
          arr_varQry1(Q1_EXIST, lngY) = CBool(True)
          arr_varQry2(Q2_EXIST, lngX) = CBool(True)
          arr_varQry1(Q1_QID, lngY) = arr_varQry2(Q2_QID, lngX)
          Exit For
        End If
      End If
    Next  ' ** lngY.
  Next  ' ** lngX.

  lngQrysNotFound = 0&
  For lngX = 0& To (lngQrys2 - 1&)
    If arr_varQry2(Q2_EXIST, lngX) = False Then
      lngQrysNotFound = lngQrysNotFound + 1&
    End If
  Next  ' ** lngX.

  Debug.Print "'QRYS NOT FOUND!  " & CStr(lngQrysNotFound)
  DoEvents

  lngQrysNotDocumented = 0&
  For lngX = 0& To (lngQrys1 - 1&)
    If arr_varQry1(Q1_EXIST, lngX) = False Then
      lngQrysNotDocumented = lngQrysNotDocumented + 1&
    End If
  Next  ' ** lngX.

  Debug.Print "'QRYS NOT DOC'D:  " & CStr(lngQrysNotDocumented)
  DoEvents

  For lngX = 0& To (lngQrys1 - 1&)
    If arr_varQry1(Q1_EXIST, lngX) = False Then
      Debug.Print "'QRY1:  " & arr_varQry1(Q1_QNAM, lngX)
    End If
  Next  ' ** lngX.

'QRYS1:  9206
'SYS QRYS:  1
'QRYS2:  9203
'QRYS NOT FOUND!  0
'QRYS NOT DOC'D:  3
'QRY1:  zz_qry_Query_Field_02_01
'QRY1:  zz_qry_Query_Field_02_02
'QRY1:  zzz_qry_Query_01

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Qry_ChkQryExist = blnRetValx

End Function
