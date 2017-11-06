Attribute VB_Name = "zz_mod_DatabaseCompare"
Option Compare Database
Option Explicit

'VGC 09/19/2015: CHANGES!

Private Const THIS_NAME As String = "zz_mod_DatabaseCompare"
' **

Public Function DataComp_Tbls_Loc() As Boolean
' ** Compares tables local to this MDB to another MDB.
' ** Ignores linked tables.

  Const THIS_PROC As String = "DataComp_Tbls_Loc"

  Dim wrk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database
  Dim tdfLoc As DAO.TableDef, tdfLnk As DAO.TableDef, fldLoc As DAO.Field, fldLnk As DAO.Field, rst As DAO.Recordset
  Dim strSysPathFile As String
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngFlds As Long, arr_varFld() As Variant
  Dim lngNewTbls As Long, lngChangedTbls As Long, lngNewFlds As Long, lngChangedFlds As Long
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const T_DID  As Integer = 0
  Const T_TNAM As Integer = 1
  Const T_FLDS As Integer = 2
  Const T_FND  As Integer = 3
  Const T_CHG  As Integer = 4
  Const T_NOTE As Integer = 5

  ' ** Array: arr_varFld().
  Const F_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const F_DID   As Integer = 0
  Const F_TNAM  As Integer = 1
  Const F_TELEM As Integer = 2
  Const F_FNAM  As Integer = 3
  Const F_NOTE  As Integer = 4

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strSysPathFile = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
  strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
  strFile = "Trust - Copy (20).mdb"
  strPathFile = strPath & LNK_SEP & strFile

  DBEngine.SystemDB = strSysPathFile

  Set dbsLoc = CurrentDb

  Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)
  Set dbsLnk = wrk.OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)

  For Each tdfLnk In dbsLnk.TableDefs
    With tdfLnk
      If .Connect = vbNullString Then
        lngTbls = lngTbls + 1&
        lngE = lngTbls - 1&
        ReDim Preserve arr_varTbl(T_ELEMS, lngE)
        arr_varTbl(T_DID, lngE) = lngThisDbsID
        arr_varTbl(T_TNAM, lngE) = .Name
        arr_varTbl(T_FLDS, lngE) = .Fields.Count
        arr_varTbl(T_FND, lngE) = CBool(False)
        arr_varTbl(T_CHG, lngE) = CBool(False)
        arr_varTbl(T_NOTE, lngE) = vbNullString
      End If
    End With  ' ** tdfLnk.
  Next  ' ** tdfLnk.
  Set tdfLnk = Nothing

  Debug.Print "'LOCAL TBLS: " & CStr(lngTbls)
  DoEvents

  If lngTbls > 0& Then

    lngNewTbls = 0&: lngChangedTbls = 0&
    For lngX = 0& To (lngTbls - 1&)
      blnFound = False
      For Each tdfLoc In dbsLoc.TableDefs
        With tdfLoc
          If .Name = arr_varTbl(T_TNAM, lngX) Then
            blnFound = True
            arr_varTbl(T_FND, lngX) = CBool(True)
            If .Fields.Count <> arr_varTbl(T_FLDS, lngX) Then
              arr_varTbl(T_CHG, lngX) = CBool(True)
              arr_varTbl(T_NOTE, lngE) = "FLD CNT"
              lngChangedTbls = lngChangedTbls + 1&
            End If
            Exit For
          End If
        End With  ' ** tdfLoc.
      Next  ' ** tdfLoc.
      If blnFound = False Then
        lngNewTbls = lngNewTbls + 1&
      End If
    Next  ' ** lngX.
    Set tdfLoc = Nothing

    For Each tdfLoc In dbsLoc.TableDefs
      With tdfLoc
        If .Connect = vbNullString Then
          blnFound = False
          For lngX = 0& To (lngTbls - 1&)
            If arr_varTbl(T_TNAM, lngX) = .Name Then
              blnFound = True
              Exit For
            End If
          Next  ' ** lngX.
          If blnFound = False Then
            lngTbls = lngTbls + 1&
            lngE = lngTbls - 1&
            ReDim Preserve arr_varTbl(T_ELEMS, lngE)
            arr_varTbl(T_DID, lngE) = lngThisDbsID
            arr_varTbl(T_TNAM, lngE) = .Name
            arr_varTbl(T_FLDS, lngE) = .Fields.Count
            arr_varTbl(T_FND, lngE) = CBool(False)
            arr_varTbl(T_CHG, lngE) = CBool(False)
            arr_varTbl(T_NOTE, lngE) = "NEW TABLE HERE"
            lngNewTbls = lngNewTbls + 1&
          End If
        End If
      End With  ' ** tdfLoc.
    Next  ' ** tdfLoc.

    Debug.Print "'NEW TBLS: " & CStr(lngNewTbls)
    DoEvents

    lngFlds = 0&
    ReDim arr_varFld(F_ELEMS, 0)

    lngNewFlds = 0&: lngChangedFlds = 0&
    For lngX = 0& To (lngTbls - 1&)
      If arr_varTbl(T_FND, lngX) = True Then
        Set tdfLoc = dbsLoc.TableDefs(arr_varTbl(T_TNAM, lngX))
        Set tdfLnk = dbsLnk.TableDefs(arr_varTbl(T_TNAM, lngX))
        For Each fldLnk In tdfLnk.Fields
          blnFound = False
          lngE = -1&
          For Each fldLoc In tdfLoc.Fields
            With fldLoc
              If .Name = fldLnk.Name Then
                blnFound = True
                If .Type <> fldLnk.Type Then
                  lngFlds = lngFlds + 1&
                  lngE = lngFlds - 1&
                  ReDim Preserve arr_varFld(F_ELEMS, lngE)
                  arr_varFld(F_DID, lngE) = lngThisDbsID
                  arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
                  arr_varFld(F_TELEM, lngE) = lngX
                  arr_varFld(F_FNAM, lngE) = fldLnk.Name
                  arr_varFld(F_NOTE, lngE) = "TYPE"
                  If arr_varTbl(T_CHG, lngX) = False Then
                    arr_varTbl(T_CHG, lngX) = CBool(True)
                    lngChangedTbls = lngChangedTbls + 1&
                  End If
                  lngChangedFlds = lngChangedFlds + 1&
                End If
                If .Size <> fldLnk.Size Then
                  If lngE = -1& Then
                    lngFlds = lngFlds + 1&
                    lngE = lngFlds - 1&
                    ReDim Preserve arr_varFld(F_ELEMS, lngE)
                    arr_varFld(F_DID, lngE) = lngThisDbsID
                    arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
                    arr_varFld(F_TELEM, lngE) = lngX
                    arr_varFld(F_FNAM, lngE) = fldLnk.Name
                    arr_varFld(F_NOTE, lngE) = "SIZE"
                    lngChangedFlds = lngChangedFlds + 1&
                  Else
                    arr_varFld(F_NOTE, lngE) = arr_varFld(F_NOTE, lngE) & ";SIZE"
                  End If
                  If arr_varTbl(T_CHG, lngX) = False Then
                    arr_varTbl(T_CHG, lngX) = CBool(True)
                    lngChangedTbls = lngChangedTbls + 1&
                  End If
                End If
                If .Required <> fldLnk.Required Then
                  If lngE = -1& Then
                    lngFlds = lngFlds + 1&
                    lngE = lngFlds - 1&
                    ReDim Preserve arr_varFld(F_ELEMS, lngE)
                    arr_varFld(F_DID, lngE) = lngThisDbsID
                    arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
                    arr_varFld(F_TELEM, lngE) = lngX
                    arr_varFld(F_FNAM, lngE) = fldLnk.Name
                    arr_varFld(F_NOTE, lngE) = "REQ"
                    lngChangedFlds = lngChangedFlds + 1&
                  Else
                    arr_varFld(F_NOTE, lngE) = arr_varFld(F_NOTE, lngE) & ";REQ"
                  End If
                  If arr_varTbl(T_CHG, lngX) = False Then
                    arr_varTbl(T_CHG, lngX) = CBool(True)
                    lngChangedTbls = lngChangedTbls + 1&
                  End If
                End If
                '.Format
                '.DefaultValue
                'Index
                Exit For
              End If
            End With  ' ** fldLoc.
          Next  ' ** fldLoc.
          If blnFound = False Then
            lngFlds = lngFlds + 1&
            lngE = lngFlds - 1&
            ReDim Preserve arr_varFld(F_ELEMS, lngE)
            arr_varFld(F_DID, lngE) = lngThisDbsID
            arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
            arr_varFld(F_TELEM, lngE) = lngX
            arr_varFld(F_FNAM, lngE) = fldLnk.Name
            arr_varFld(F_NOTE, lngE) = "NEW FIELD"
            lngNewFlds = lngNewFlds + 1&
            If arr_varTbl(T_CHG, lngX) = False Then
              arr_varTbl(T_CHG, lngX) = CBool(True)
              lngChangedTbls = lngChangedTbls + 1&
            End If
          End If
        Next  ' ** fldLnk.
        Set fldLoc = Nothing
        Set fldLnk = Nothing
        For Each fldLoc In tdfLoc.Fields
          blnFound = False
          For Each fldLnk In tdfLnk.Fields
            With fldLnk
              If .Name = fldLoc.Name Then
                blnFound = True
                Exit For
              End If
            End With  ' ** fldLnk.
          Next  ' ** fldLnk.
          If blnFound = False Then
            lngFlds = lngFlds + 1&
            lngE = lngFlds - 1&
            ReDim Preserve arr_varFld(F_ELEMS, lngE)
            arr_varFld(F_DID, lngE) = lngThisDbsID
            arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
            arr_varFld(F_TELEM, lngE) = lngX
            arr_varFld(F_FNAM, lngE) = fldLoc.Name
            arr_varFld(F_NOTE, lngE) = "NEW FIELD HERE"
            lngNewFlds = lngNewFlds + 1&
            If arr_varTbl(T_CHG, lngX) = False Then
              arr_varTbl(T_CHG, lngX) = CBool(True)
              lngChangedTbls = lngChangedTbls + 1&
            End If
          End If
        Next  ' ** fldLoc.
        Set fldLoc = Nothing
        Set fldLnk = Nothing
      End If
    Next  ' ** lngX.
    Set tdfLoc = Nothing
    Set tdfLnk = Nothing

    For lngX = 0& To (lngTbls - 1&)
      If arr_varTbl(T_FND, lngX) = False Then
        'Debug.Print "'  NEW: " & arr_varTbl(T_TNAM, lngX)
        'DoEvents
      End If
    Next  ' ** lngX.
    DoEvents

    Debug.Print "'CHANGED TBLS: " & CStr(lngChangedTbls)
    For lngX = 0& To (lngTbls - 1&)
      If arr_varTbl(T_CHG, lngX) = True Then
        'Debug.Print "'  CHG: " & arr_varTbl(T_TNAM, lngX) & "  " & arr_varTbl(T_NOTE, lngX)
        'DoEvents
      End If
    Next  ' ** lngX.
    DoEvents

    Debug.Print "'NEW FLDS: " & CStr(lngNewFlds)
    For lngX = 0& To (lngFlds - 1&)
      If arr_varFld(F_NOTE, lngX) = "NEW FIELD" Then
        'Debug.Print "'  NEW FLD: " & arr_varFld(F_TNAM, lngX) & "  " & arr_varFld(F_FNAM, lngX)
        'DoEvents
      ElseIf arr_varFld(F_NOTE, lngX) = "NEW FIELD HERE" Then
        'Debug.Print "'  NEW FLD HERE: " & arr_varFld(F_TNAM, lngX) & "  " & arr_varFld(F_FNAM, lngX)
        'DoEvents
      End If
    Next  ' ** lngX.
    DoEvents

    Debug.Print "'CHANGED FLDS: " & CStr(lngChangedFlds)
    For lngX = 0& To (lngFlds - 1&)
      If Left(arr_varFld(F_NOTE, lngX), 9) <> "NEW FIELD" Then
        'Debug.Print "'  CHG FLD: " & arr_varFld(F_TNAM, lngX) & "  " & arr_varFld(F_FNAM, lngX) & "  " & arr_varFld(F_NOTE, lngX)
        'DoEvents
      End If
    Next  ' ** lngX.
    DoEvents

    Set rst = dbsLoc.OpenRecordset("zz_tbl_DataComp_01", dbOpenDynaset, dbConsistent)
    With rst
      For lngX = 0& To (lngTbls - 1&)
        If arr_varTbl(T_FND, lngX) = False Or arr_varTbl(T_CHG, lngX) = True Then
          .AddNew
          ' ** ![dc01_id] : AutoNumber.
          ![dbs_id] = arr_varTbl(T_DID, lngX)
          ![tbl_name] = arr_varTbl(T_TNAM, lngX)
          ![tbl_fld_cnt] = arr_varTbl(T_FLDS, lngX)
          ![tbl_found] = arr_varTbl(T_FND, lngX)
          ![tbl_change] = arr_varTbl(T_CHG, lngX)
          If arr_varTbl(T_NOTE, lngX) <> vbNullString Then
            ![tbl_note] = arr_varTbl(T_NOTE, lngX)
          End If
          ![dc01_datemodified] = Now()
          .Update
        End If
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    Set rst = Nothing

    Set rst = dbsLoc.OpenRecordset("zz_tbl_DataComp_02", dbOpenDynaset, dbConsistent)
    With rst
      For lngX = 0& To (lngFlds - 1&)
        .AddNew
        ' ** ![dc02_id] : AutoNumber.
        ![dbs_id] = arr_varFld(F_DID, lngX)
        ![tbl_name] = arr_varFld(F_TNAM, lngX)
        ![fld_name] = arr_varFld(F_FNAM, lngX)
        If arr_varFld(F_NOTE, lngX) <> vbNullString Then
          ![fld_note] = arr_varFld(F_NOTE, lngX)
        End If
        ![dc02_datemodified] = Now()
        .Update
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    Set rst = Nothing

  Else
    Debug.Print "'NONE FOUND!"
    DoEvents
  End If

  dbsLoc.Close
  dbsLnk.Close
  wrk.Close

'LOCAL TBLS: 261
'NEW TBLS: 19
'CHANGED TBLS: 2
'NEW FLDS: 10
'CHANGED FLDS: 0
'DONE!  DataComp_Tbls_Loc()
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  Beep

  Set rst = Nothing
  Set fldLoc = Nothing
  Set fldLnk = Nothing
  Set tdfLoc = Nothing
  Set tdfLnk = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set wrk = Nothing

  DataComp_Tbls_Loc = blnRetVal

End Function

Public Function DataComp_Tbls_Lnk() As Boolean
' ** Compares tables linked to this MDB with another MDB.
' ** Ignores local tables.

  Const THIS_PROC As String = "DataComp_Tbls_Lnk"

  Dim wrkLoc As DAO.Workspace, wrkLnk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, dbs As DAO.Database
  Dim tdfLoc As DAO.TableDef, tdfLnk As DAO.TableDef, fldLoc As DAO.Field, fldLnk As DAO.Field, rst As DAO.Recordset
  Dim strSysPathFile As String
  Dim strPath1 As String, strPath2 As String, strFile1 As String, strFile2 As String, strPathFile1 As String, strPathFile2 As String
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngFlds As Long, arr_varFld() As Variant
  Dim lngNewTbls As Long, lngChangedTbls As Long, lngNewFlds As Long, lngChangedFlds As Long
  Dim lngThisDbsID As Long, lngThatDbsID As Long
  Dim blnFound As Boolean
  Dim varTmp00 As Variant
  Dim lngW As Long, lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const T_DID  As Integer = 0
  Const T_TNAM As Integer = 1
  Const T_FLDS As Integer = 2
  Const T_FND  As Integer = 3
  Const T_CHG  As Integer = 4
  Const T_NOTE As Integer = 5

  ' ** Array: arr_varFld().
  Const F_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const F_DID   As Integer = 0
  Const F_TNAM  As Integer = 1
  Const F_TELEM As Integer = 2
  Const F_FNAM  As Integer = 3
  Const F_NOTE  As Integer = 4

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strSysPathFile = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
  strPath1 = CurrentAppPath  ' ** Module Function: modFileUtilities.
  strPath1 = strPath1 & LNK_SEP & "Database"
  strPath2 = strPath1

  DBEngine.SystemDB = strSysPathFile

  Set dbs = CurrentDb

  For lngW = 1& To 2&

    Select Case lngW
    Case 1&
      strFile1 = "TrstArch.mdb"
      varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strFile1 & "'")
      If IsNull(varTmp00) = False Then
        lngThatDbsID = varTmp00
      Else
        Stop
      End If
      strPathFile1 = strPath1 & LNK_SEP & strFile1
      strFile2 = "TrstArch_bak_WmB_new.mdb"
      strPathFile2 = strPath2 & LNK_SEP & strFile2
    Case 2&
      strFile1 = "TrustDta.mdb"
      varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strFile1 & "'")
      If IsNull(varTmp00) = False Then
        lngThatDbsID = varTmp00
      Else
        Stop
      End If
      strPathFile1 = strPath1 & LNK_SEP & strFile1
      strFile2 = "TrustDta_bak_WmB_new.mdb"
      strPathFile2 = strPath2 & LNK_SEP & strFile2
    End Select

    Set wrkLoc = CreateWorkspace("tmpDB1", "Superuser", TA_SEC, dbUseJet)
    Set dbsLoc = wrkLoc.OpenDatabase(strPathFile1, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
    Set wrkLnk = CreateWorkspace("tmpDB2", "Superuser", TA_SEC, dbUseJet)
    Set dbsLnk = wrkLoc.OpenDatabase(strPathFile2, False, True)  ' ** {pathfile}, {exclusive}, {read-only}

    lngTbls = 0&
    ReDim arr_varTbl(T_ELEMS, 0)

    For Each tdfLnk In dbsLnk.TableDefs
      With tdfLnk
        If Left(.Name, 4) <> "~TMP" And Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "USys" Then
          If .Connect = vbNullString Then
            lngTbls = lngTbls + 1&
            lngE = lngTbls - 1&
            ReDim Preserve arr_varTbl(T_ELEMS, lngE)
            arr_varTbl(T_DID, lngE) = lngThatDbsID
            arr_varTbl(T_TNAM, lngE) = .Name
            arr_varTbl(T_FLDS, lngE) = .Fields.Count
            arr_varTbl(T_FND, lngE) = CBool(False)
            arr_varTbl(T_CHG, lngE) = CBool(False)
            arr_varTbl(T_NOTE, lngE) = vbNullString
          End If
        End If
      End With  ' ** tdfLnk.
    Next  ' ** tdfLnk.
    Set tdfLnk = Nothing

    Debug.Print "'LINKED TBLS: " & CStr(lngTbls)
    Debug.Print "'  " & strPathFile1
    Debug.Print "'  " & strPathFile2
    DoEvents

    If lngTbls > 0& Then

      lngNewTbls = 0&: lngChangedTbls = 0&
      For lngX = 0& To (lngTbls - 1&)
        blnFound = False
        For Each tdfLoc In dbsLoc.TableDefs
          With tdfLoc
            If .Name = arr_varTbl(T_TNAM, lngX) Then
              blnFound = True
              arr_varTbl(T_FND, lngX) = CBool(True)
              If .Fields.Count <> arr_varTbl(T_FLDS, lngX) Then
                arr_varTbl(T_CHG, lngX) = CBool(True)
                arr_varTbl(T_NOTE, lngE) = "FLD CNT"
                lngChangedTbls = lngChangedTbls + 1&
              End If
              Exit For
            End If
          End With  ' ** tdfLoc.
        Next  ' ** tdfLoc.
        If blnFound = False Then
          lngNewTbls = lngNewTbls + 1&
        End If
      Next  ' ** lngX.
      Set tdfLoc = Nothing

      For Each tdfLoc In dbsLoc.TableDefs
        With tdfLoc
          If Left(.Name, 4) <> "~TMP" And Left(.Name, 4) <> "MSys" And Left(.Name, 4) <> "USys" Then
            If .Connect = vbNullString Then
              blnFound = False
              For lngX = 0& To (lngTbls - 1&)
                If arr_varTbl(T_TNAM, lngX) = .Name Then
                  blnFound = True
                  Exit For
                End If
              Next  ' ** lngX.
              If blnFound = False Then
                lngTbls = lngTbls + 1&
                lngE = lngTbls - 1&
                ReDim Preserve arr_varTbl(T_ELEMS, lngE)
                arr_varTbl(T_DID, lngE) = lngThatDbsID
                arr_varTbl(T_TNAM, lngE) = .Name
                arr_varTbl(T_FLDS, lngE) = .Fields.Count
                arr_varTbl(T_FND, lngE) = CBool(False)
                arr_varTbl(T_CHG, lngE) = CBool(False)
                arr_varTbl(T_NOTE, lngE) = "NEW TABLE HERE"
                lngNewTbls = lngNewTbls + 1&
              End If
            End If
          End If
        End With  ' ** tdfLoc.
      Next  ' ** tdfLoc.

      Debug.Print "'NEW TBLS: " & CStr(lngNewTbls)
      DoEvents

      lngFlds = 0&
      ReDim arr_varFld(F_ELEMS, 0)

      lngNewFlds = 0&: lngChangedFlds = 0&
      For lngX = 0& To (lngTbls - 1&)
        If arr_varTbl(T_FND, lngX) = True Then
          Set tdfLoc = dbsLoc.TableDefs(arr_varTbl(T_TNAM, lngX))
          Set tdfLnk = dbsLnk.TableDefs(arr_varTbl(T_TNAM, lngX))
          For Each fldLnk In tdfLnk.Fields
            blnFound = False
            lngE = -1&
            For Each fldLoc In tdfLoc.Fields
              With fldLoc
                If .Name = fldLnk.Name Then
                  blnFound = True
                  If .Type <> fldLnk.Type Then
                    lngFlds = lngFlds + 1&
                    lngE = lngFlds - 1&
                    ReDim Preserve arr_varFld(F_ELEMS, lngE)
                    arr_varFld(F_DID, lngE) = lngThatDbsID
                    arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
                    arr_varFld(F_TELEM, lngE) = lngX
                    arr_varFld(F_FNAM, lngE) = fldLnk.Name
                    arr_varFld(F_NOTE, lngE) = "TYPE"
                    If arr_varTbl(T_CHG, lngX) = False Then
                      arr_varTbl(T_CHG, lngX) = CBool(True)
                      lngChangedTbls = lngChangedTbls + 1&
                    End If
                    lngChangedFlds = lngChangedFlds + 1&
                  End If
                  If .Size <> fldLnk.Size Then
                    If lngE = -1& Then
                      lngFlds = lngFlds + 1&
                      lngE = lngFlds - 1&
                      ReDim Preserve arr_varFld(F_ELEMS, lngE)
                      arr_varFld(F_DID, lngE) = lngThatDbsID
                      arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
                      arr_varFld(F_TELEM, lngE) = lngX
                      arr_varFld(F_FNAM, lngE) = fldLnk.Name
                      arr_varFld(F_NOTE, lngE) = "SIZE"
                      lngChangedFlds = lngChangedFlds + 1&
                    Else
                      arr_varFld(F_NOTE, lngE) = arr_varFld(F_NOTE, lngE) & ";SIZE"
                    End If
                    If arr_varTbl(T_CHG, lngX) = False Then
                      arr_varTbl(T_CHG, lngX) = CBool(True)
                      lngChangedTbls = lngChangedTbls + 1&
                    End If
                  End If
                  If .Required <> fldLnk.Required Then
                    If lngE = -1& Then
                      lngFlds = lngFlds + 1&
                      lngE = lngFlds - 1&
                      ReDim Preserve arr_varFld(F_ELEMS, lngE)
                      arr_varFld(F_DID, lngE) = lngThatDbsID
                      arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
                      arr_varFld(F_TELEM, lngE) = lngX
                      arr_varFld(F_FNAM, lngE) = fldLnk.Name
                      arr_varFld(F_NOTE, lngE) = "REQ"
                      lngChangedFlds = lngChangedFlds + 1&
                    Else
                      arr_varFld(F_NOTE, lngE) = arr_varFld(F_NOTE, lngE) & ";REQ"
                    End If
                    If arr_varTbl(T_CHG, lngX) = False Then
                      arr_varTbl(T_CHG, lngX) = CBool(True)
                      lngChangedTbls = lngChangedTbls + 1&
                    End If
                  End If
                  '.Format
                  '.DefaultValue
                  'Index
                  Exit For
                End If
              End With  ' ** fldLoc.
            Next  ' ** fldLoc.
            If blnFound = False Then
              lngFlds = lngFlds + 1&
              lngE = lngFlds - 1&
              ReDim Preserve arr_varFld(F_ELEMS, lngE)
              arr_varFld(F_DID, lngE) = lngThatDbsID
              arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
              arr_varFld(F_TELEM, lngE) = lngX
              arr_varFld(F_FNAM, lngE) = fldLnk.Name
              arr_varFld(F_NOTE, lngE) = "NEW FIELD"
              lngNewFlds = lngNewFlds + 1&
              If arr_varTbl(T_CHG, lngX) = False Then
                arr_varTbl(T_CHG, lngX) = CBool(True)
                lngChangedTbls = lngChangedTbls + 1&
              End If
            End If
          Next  ' ** fldLnk.
          Set fldLoc = Nothing
          Set fldLnk = Nothing
          For Each fldLoc In tdfLoc.Fields
            blnFound = False
            For Each fldLnk In tdfLnk.Fields
              With fldLnk
                If .Name = fldLoc.Name Then
                  blnFound = True
                  Exit For
                End If
              End With  ' ** fldLnk.
            Next  ' ** fldLnk.
            If blnFound = False Then
              lngFlds = lngFlds + 1&
              lngE = lngFlds - 1&
              ReDim Preserve arr_varFld(F_ELEMS, lngE)
              arr_varFld(F_DID, lngE) = lngThatDbsID
              arr_varFld(F_TNAM, lngE) = arr_varTbl(T_TNAM, lngX)
              arr_varFld(F_TELEM, lngE) = lngX
              arr_varFld(F_FNAM, lngE) = fldLoc.Name
              arr_varFld(F_NOTE, lngE) = "NEW FIELD HERE"
              lngNewFlds = lngNewFlds + 1&
              If arr_varTbl(T_CHG, lngX) = False Then
                arr_varTbl(T_CHG, lngX) = CBool(True)
                lngChangedTbls = lngChangedTbls + 1&
              End If
            End If
          Next  ' ** fldLoc.
          Set fldLoc = Nothing
          Set fldLnk = Nothing
        End If
      Next  ' ** lngX.
      Set tdfLoc = Nothing
      Set tdfLnk = Nothing

      'For lngX = 0& To (lngTbls - 1&)
      '  If arr_varTbl(T_FND, lngX) = False Then
      '    Debug.Print "'  NEW: " & arr_varTbl(T_TNAM, lngX)
      '    DoEvents
      '  End If
      'Next  ' ** lngX.
      DoEvents

      Debug.Print "'CHANGED TBLS: " & CStr(lngChangedTbls)
      'For lngX = 0& To (lngTbls - 1&)
      '  If arr_varTbl(T_CHG, lngX) = True Then
      '    Debug.Print "'  CHG: " & arr_varTbl(T_TNAM, lngX) & "  " & arr_varTbl(T_NOTE, lngX)
      '    DoEvents
      '  End If
      'Next  ' ** lngX.
      DoEvents

      Debug.Print "'NEW FLDS: " & CStr(lngNewFlds)
      'For lngX = 0& To (lngFlds - 1&)
      '  If arr_varFld(F_NOTE, lngX) = "NEW FIELD" Then
      '    Debug.Print "'  NEW FLD: " & arr_varFld(F_TNAM, lngX) & "  " & arr_varFld(F_FNAM, lngX)
      '    DoEvents
      '  ElseIf arr_varFld(F_NOTE, lngX) = "NEW FIELD HERE" Then
      '    Debug.Print "'  NEW FLD HERE: " & arr_varFld(F_TNAM, lngX) & "  " & arr_varFld(F_FNAM, lngX)
      '    DoEvents
      '  End If
      'Next  ' ** lngX.
      DoEvents

      Debug.Print "'CHANGED FLDS: " & CStr(lngChangedFlds)
      'For lngX = 0& To (lngFlds - 1&)
      '  If Left(arr_varFld(F_NOTE, lngX), 9) <> "NEW FIELD" Then
      '    Debug.Print "'  CHG FLD: " & arr_varFld(F_TNAM, lngX) & "  " & arr_varFld(F_FNAM, lngX) & "  " & arr_varFld(F_NOTE, lngX)
      '    DoEvents
      '  End If
      'Next  ' ** lngX.
      DoEvents

      Set rst = dbs.OpenRecordset("zz_tbl_DataComp_01", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngTbls - 1&)
          If arr_varTbl(T_FND, lngX) = False Or arr_varTbl(T_CHG, lngX) = True Then
            .AddNew
            ' ** ![dc01_id] : AutoNumber.
            ![dbs_id] = arr_varTbl(T_DID, lngX)
            ![tbl_name] = arr_varTbl(T_TNAM, lngX)
            ![tbl_fld_cnt] = arr_varTbl(T_FLDS, lngX)
            ![tbl_found] = arr_varTbl(T_FND, lngX)
            ![tbl_change] = arr_varTbl(T_CHG, lngX)
            If arr_varTbl(T_NOTE, lngX) <> vbNullString Then
              ![tbl_note] = arr_varTbl(T_NOTE, lngX)
            End If
            ![dc01_datemodified] = Now()
            .Update
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      Set rst = dbs.OpenRecordset("zz_tbl_DataComp_02", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngFlds - 1&)
          .AddNew
          ' ** ![dc02_id] : AutoNumber.
          ![dbs_id] = arr_varFld(F_DID, lngX)
          ![tbl_name] = arr_varFld(F_TNAM, lngX)
          ![fld_name] = arr_varFld(F_FNAM, lngX)
          If arr_varFld(F_NOTE, lngX) <> vbNullString Then
            ![fld_note] = arr_varFld(F_NOTE, lngX)
          End If
          ![dc02_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

    Else
      Debug.Print "'NONE FOUND!"
      DoEvents
    End If

    dbsLoc.Close
    dbsLnk.Close
    wrkLoc.Close
    wrkLnk.Close

  Next  ' ** lngW.
  dbs.Close

'LINKED TBLS: 3
'  C:\Program Files\Delta Data\Trust Accountant\Database\TrstArch.mdb
'  C:\Program Files\Delta Data\Trust Accountant\Database\TrstArch_bak_WmB_new.mdb
'NEW TBLS: 0
'CHANGED TBLS: 0
'NEW FLDS: 0
'CHANGED FLDS: 0
'LINKED TBLS: 76
'  C:\Program Files\Delta Data\Trust Accountant\Database\TrustDta.mdb
'  C:\Program Files\Delta Data\Trust Accountant\Database\TrustDta_bak_WmB_new.mdb
'NEW TBLS: 0
'CHANGED TBLS: 0
'NEW FLDS: 0
'CHANGED FLDS: 0
'DONE!  DataComp_Tbls_Lnk()
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  Beep

  Set rst = Nothing
  Set fldLoc = Nothing
  Set fldLnk = Nothing
  Set tdfLoc = Nothing
  Set tdfLnk = Nothing
  Set dbs = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set wrkLoc = Nothing
  Set wrkLnk = Nothing

  DataComp_Tbls_Lnk = blnRetVal

End Function

Public Function DataComp_Qrys_Loc() As Boolean

  Const THIS_PROC As String = "DataComp_Qrys_Loc"

  Dim wrk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, rst As DAO.Recordset
  Dim qdfLoc As DAO.QueryDef, qdfLnk As DAO.QueryDef
  Dim strSysPathFile As String
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngNewQrys As Long, lngChangedQrys As Long
  Dim strDesc As String
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const Q_DID  As Integer = 0
  Const Q_QNAM As Integer = 1
  Const Q_TYP  As Integer = 2
  Const Q_SQL  As Integer = 3
  Const Q_DSC  As Integer = 4
  Const Q_FND  As Integer = 5
  Const Q_CHG  As Integer = 6
  Const Q_NOTE As Integer = 7

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strSysPathFile = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
  strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
  strFile = "Trust - Copy (20).mdb"
  strPathFile = strPath & LNK_SEP & strFile

  DBEngine.SystemDB = strSysPathFile

  Set dbsLoc = CurrentDb

  Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)
  Set dbsLnk = wrk.OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  lngX = 0&
  Debug.Print "'|";
  DoEvents

  For Each qdfLnk In dbsLnk.QueryDefs
    lngX = lngX + 1&
    strDesc = vbNullString
    With qdfLnk
      If Left(.Name, 7) <> "zz_qry_" And Left(.Name, 8) <> "zzz_qry_" Then
        lngQrys = lngQrys + 1&
        lngE = lngQrys - 1&
        ReDim Preserve arr_varQry(Q_ELEMS, lngE)
        arr_varQry(Q_DID, lngE) = lngThisDbsID
        arr_varQry(Q_QNAM, lngE) = .Name
        arr_varQry(Q_TYP, lngE) = .Type
        arr_varQry(Q_SQL, lngE) = .SQL
On Error Resume Next
        strDesc = .Properties("Description")
On Error GoTo 0
        If strDesc <> vbNullString Then
          arr_varQry(Q_DSC, lngE) = strDesc
        Else
          arr_varQry(Q_DSC, lngE) = vbNullString
        End If
        arr_varQry(Q_FND, lngE) = CBool(False)
        arr_varQry(Q_CHG, lngE) = CBool(False)
        arr_varQry(Q_NOTE, lngE) = vbNullString
      End If
    End With  ' ** qdfLnk.
    If (lngX Mod 1000&) = 0& Then
      Debug.Print "|  " & CStr(lngX)
      Debug.Print "'|";
    ElseIf (lngX Mod 100&) = 0& Then
      Debug.Print "|";
    ElseIf (lngX Mod 10&) = 0& Then
      Debug.Print ".";
    End If
    DoEvents
  Next  ' ** qdfLnk.
  Set qdfLnk = Nothing
  Debug.Print

  Debug.Print "'QRYS: " & CStr(lngQrys)
  DoEvents

  If lngQrys > 0& Then

    Debug.Print "'|";
    DoEvents

    lngNewQrys = 0&: lngChangedQrys = 0&
    For lngX = 0& To (lngQrys - 1&)
      blnFound = False
      strDesc = vbNullString
      For Each qdfLoc In dbsLoc.QueryDefs
        With qdfLoc
          If .Name = arr_varQry(Q_QNAM, lngX) Then
            blnFound = True
            arr_varQry(Q_FND, lngX) = CBool(True)
            If .Type <> arr_varQry(Q_TYP, lngX) Then
              arr_varQry(Q_CHG, lngX) = CBool(True)
              arr_varQry(Q_NOTE, lngX) = "DIFF TYPE;"
            End If
            If .SQL <> arr_varQry(Q_SQL, lngX) Then
              arr_varQry(Q_NOTE, lngX) = arr_varQry(Q_NOTE, lngX) & "DIFF SQL;"
            End If
On Error Resume Next
            strDesc = .Properties("Description")
On Error GoTo 0
            If strDesc <> arr_varQry(Q_DSC, lngX) Then
              arr_varQry(Q_NOTE, lngX) = arr_varQry(Q_NOTE, lngX) & "DIFF DESC;"
            End If
            Exit For
          End If
        End With  ' ** tdfLoc.
      Next  ' ** tdfLoc.
      If blnFound = False Then
        lngNewQrys = lngNewQrys + 1&
      End If
      If (lngX Mod 1000&) = 0& Then
        Debug.Print "|  " & CStr(lngX)
        Debug.Print "'|";
      ElseIf (lngX Mod 100&) = 0& Then
        Debug.Print "|";
      ElseIf (lngX Mod 10&) = 0& Then
        Debug.Print ".";
      End If
      DoEvents
    Next  ' ** lngX.
    Set qdfLoc = Nothing
    Debug.Print

    Debug.Print "'|";
    DoEvents

    lngX = 0&
    For Each qdfLoc In dbsLoc.QueryDefs
      lngX = lngX + 1&
      With qdfLoc
        If Left(.Name, 7) <> "zz_qry_" And Left(.Name, 8) <> "zzz_qry_" Then
            blnFound = False
          For lngX = 0& To (lngQrys - 1&)
            If arr_varQry(Q_QNAM, lngX) = .Name Then
              blnFound = True
              Exit For
            End If
          Next  ' ** lngX.
          If blnFound = False Then
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_DID, lngE) = lngThisDbsID
            arr_varQry(Q_QNAM, lngE) = .Name
            arr_varQry(Q_TYP, lngE) = .Type
            arr_varQry(Q_SQL, lngE) = Null
            arr_varQry(Q_DSC, lngE) = vbNullString
            arr_varQry(Q_FND, lngE) = CBool(False)
            arr_varQry(Q_CHG, lngE) = CBool(False)
            arr_varQry(Q_NOTE, lngE) = "NEW QRY HERE"
            lngNewQrys = lngNewQrys + 1&
          End If
        End If
      End With  ' ** tdfLoc.
      If (lngX Mod 1000&) = 0& Then
        Debug.Print "|  " & CStr(lngX)
        Debug.Print "'|";
      ElseIf (lngX Mod 100&) = 0& Then
        Debug.Print "|";
      ElseIf (lngX Mod 10&) = 0& Then
        Debug.Print ".";
      End If
      DoEvents
    Next  ' ** tdfLoc.
    Debug.Print

    Debug.Print "'NEW QRYS: " & CStr(lngNewQrys)
    DoEvents

    Debug.Print "'CHANGED QRYS: " & CStr(lngChangedQrys)
    DoEvents

    Set rst = dbsLoc.OpenRecordset("zz_tbl_DataComp_03", dbOpenDynaset, dbConsistent)
    With rst
      For lngX = 0& To (lngQrys - 1&)
        If arr_varQry(Q_FND, lngX) = False Or arr_varQry(Q_CHG, lngX) = True Then
          .AddNew
          ' ** ![dc03_id] : AutoNumber.
          ![dbs_id] = arr_varQry(Q_DID, lngX)
          ![qry_name] = arr_varQry(Q_QNAM, lngX)
          ![qrytype_type] = arr_varQry(Q_TYP, lngX)
          ![qry_sql] = arr_varQry(Q_SQL, lngX)
          If arr_varQry(Q_DSC, lngX) <> vbNullString Then
            ![qry_description] = arr_varQry(Q_DSC, lngX)
          End If
          ![qry_found] = arr_varQry(Q_FND, lngX)
          ![qry_change] = arr_varQry(Q_CHG, lngX)
          If arr_varQry(Q_NOTE, lngX) <> vbNullString Then
            ![qry_note] = arr_varQry(Q_NOTE, lngX)
          End If
          ![dc03_datemodified] = Now()
          .Update
        End If
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    Set rst = Nothing

  Else
    Debug.Print "'NONE FOUND!"
    DoEvents
  End If

  dbsLoc.Close
  dbsLnk.Close
  wrk.Close

'NEW QRYS: 35
'CHANGED QRYS: 0
'DONE!  DataComp_Qrys_Loc()
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  Beep

  Set rst = Nothing
  Set qdfLoc = Nothing
  Set qdfLnk = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set wrk = Nothing

  DataComp_Qrys_Loc = blnRetVal

End Function

Public Function DataComp_Frm_Loc() As Boolean

  Const THIS_PROC As String = "DataComp_Frm_Loc"

  Dim acApp As Access.Application, frmLoc As Access.Form, frmLnk As Access.Form, rst As DAO.Recordset
  Dim dbsLoc As DAO.Database, dbsLnk As DAO.Database, ctrLoc As DAO.Container, ctrLnk As DAO.Container
  Dim docLoc As DAO.Document, docLnk As DAO.Document, ctlLoc As Access.Control, ctlLnk As Access.Control
  Dim strSysPathFile As String
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngFrms As Long, arr_varFrm() As Variant
  Dim lngCtls As Long, arr_varCtl() As Variant
  Dim lngNewFrms As Long, lngChangedFrms As Long, lngNewCtls As Long, lngChangedCtls As Long
  Dim lngThisDbsID As Long
  Dim blnAccessOpen As Boolean, blnFound As Boolean, blnTypeDiff As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varFrm().
  Const F_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const F_DID  As Integer = 0
  Const F_FNAM As Integer = 1
  Const F_CTLS As Integer = 2
  Const F_CAP  As Integer = 3
  Const F_FND  As Integer = 4
  Const F_CHG  As Integer = 5
  Const F_NOTE As Integer = 6

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const C_DID  As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_FND  As Integer = 4
  Const C_CHG  As Integer = 5
  Const C_NOTE As Integer = 6

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strSysPathFile = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
  strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
  strFile = "Trust - Copy (20).mdb"
  strPathFile = strPath & LNK_SEP & strFile

  DBEngine.SystemDB = strSysPathFile

  Set acApp = New Access.Application

  blnAccessOpen = True

  With acApp

    .Visible = False

    .OpenCurrentDatabase strPathFile, False, TA_SEC

    lngFrms = 0&
    ReDim arr_varFrm(F_ELEMS, 0)

    Set dbsLnk = .CurrentDb
    With dbsLnk
      Set ctrLnk = .Containers("Forms")
      With ctrLnk
        For Each docLnk In .Documents
          With docLnk
            lngFrms = lngFrms + 1&
            lngE = lngFrms - 1&
            ReDim Preserve arr_varFrm(F_ELEMS, lngE)
            arr_varFrm(F_DID, lngE) = lngThisDbsID
            arr_varFrm(F_FNAM, lngE) = .Name
            arr_varFrm(F_CTLS, lngE) = 0&
            arr_varFrm(F_CAP, lngE) = vbNullString
            arr_varFrm(F_FND, lngE) = CBool(False)
            arr_varFrm(F_CHG, lngE) = CBool(False)
            arr_varFrm(F_NOTE, lngE) = vbNullString
          End With  ' ** docLnk.
        Next  ' ** docLnk.
      End With  ' ** ctrLnk.
      Set docLnk = Nothing
      Set ctrLnk = Nothing
    End With  ' ** dbsLnk.

    Debug.Print "'FRMS: " & CStr(lngFrms)
    DoEvents

    If lngFrms > 0& Then

      Set dbsLoc = CurrentDb
      Set ctrLoc = dbsLoc.Containers("Forms")

      lngNewFrms = 0&: lngChangedFrms = 0&
      For lngX = 0& To (lngFrms - 1&)
        blnFound = False
        For Each docLoc In ctrLoc.Documents
          With docLoc
            If .Name = arr_varFrm(F_FNAM, lngX) Then
              blnFound = True
              arr_varFrm(F_FND, lngX) = CBool(True)
              Exit For
            End If
          End With
        Next  ' ** docLoc.
        If blnFound = False Then
          lngNewFrms = lngNewFrms + 1&
        End If
      Next  ' ** lngX.
      Set docLoc = Nothing

      For Each docLoc In ctrLoc.Documents
        blnFound = False
        With docLoc
          For lngX = 0& To (lngFrms - 1&)
            If arr_varFrm(F_FNAM, lngX) = .Name Then
              blnFound = True
              Exit For
            End If
          Next  ' ** lngX.
          If blnFound = False Then
            lngFrms = lngFrms + 1&
            lngE = lngFrms - 1&
            ReDim Preserve arr_varFrm(F_ELEMS, lngE)
            arr_varFrm(F_DID, lngE) = lngThisDbsID
            arr_varFrm(F_FNAM, lngE) = .Name
            arr_varFrm(F_CTLS, lngE) = 0&
            arr_varFrm(F_CAP, lngE) = vbNullString
            arr_varFrm(F_FND, lngE) = CBool(False)
            arr_varFrm(F_CHG, lngE) = CBool(False)
            arr_varFrm(F_NOTE, lngE) = "NEW FRM HERE!"
            lngNewFrms = lngNewFrms + 1&
          End If
        End With  ' ** docLoc.
      Next  ' ** docLoc.
      Set docLoc = Nothing
      Set ctrLoc = Nothing

      Debug.Print "'NEW FRMS: " & CStr(lngNewFrms)
      DoEvents

      lngCtls = 0&
      ReDim arr_varCtl(C_ELEMS, 0)

      Debug.Print "'|";
      DoEvents

      lngNewCtls = 0&: lngChangedCtls = 0&
      For lngX = 0& To (lngFrms - 1&)
        If arr_varFrm(F_FND, lngX) = True Then
          .DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
          Set frmLnk = .Forms(arr_varFrm(F_FNAM, lngX))
          DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
          Set frmLoc = Forms(arr_varFrm(F_FNAM, lngX))
          With frmLnk
            If .Controls.Count <> frmLoc.Controls.Count Then
              arr_varFrm(F_CHG, lngX) = CBool(True)
              arr_varFrm(F_NOTE, lngX) = arr_varFrm(F_NOTE, lngX) & "CTL CNT DIFF;"
            End If
            If .RecordSource = vbNullString Then
              If frmLoc.RecordSource <> vbNullString Then
                arr_varFrm(F_CHG, lngX) = CBool(True)
                arr_varFrm(F_NOTE, lngX) = arr_varFrm(F_NOTE, lngX) & "REC SRC DIFF;"
              End If
            Else
              If frmLoc.RecordSource = vbNullString Then
                arr_varFrm(F_CHG, lngX) = CBool(True)
                arr_varFrm(F_NOTE, lngX) = arr_varFrm(F_NOTE, lngX) & "REC SRC DIFF;"
              Else
                If frmLoc.RecordSource <> .RecordSource Then
                  arr_varFrm(F_CHG, lngX) = CBool(True)
                  arr_varFrm(F_NOTE, lngX) = arr_varFrm(F_NOTE, lngX) & "REC SRC DIFF;"
                End If
              End If
            End If
            If .Caption = vbNullString Then
              If frmLoc.Caption <> vbNullString Then
                arr_varFrm(F_CHG, lngX) = CBool(True)
                arr_varFrm(F_NOTE, lngX) = arr_varFrm(F_NOTE, lngX) & "CAP DIFF;"
              End If
            Else
              arr_varFrm(F_CAP, lngX) = .Caption
              If frmLoc.Caption = vbNullString Then
                arr_varFrm(F_CHG, lngX) = CBool(True)
                arr_varFrm(F_NOTE, lngX) = arr_varFrm(F_NOTE, lngX) & "CAP DIFF;"
              Else
                If frmLoc.Caption <> .Caption Then
                  arr_varFrm(F_CHG, lngX) = CBool(True)
                  arr_varFrm(F_NOTE, lngX) = arr_varFrm(F_NOTE, lngX) & "CAP DIFF;"
                End If
              End If
            End If
            For Each ctlLnk In .Controls
              blnFound = False: blnTypeDiff = False
              For Each ctlLoc In frmLoc.Controls
                lngE = -1&
                With ctlLoc
                  If .Name = ctlLnk.Name Then
                    blnFound = True
                    If .ControlType <> ctlLnk.ControlType Then
                      blnTypeDiff = True
                      lngCtls = lngCtls + 1&
                      lngE = lngCtls - 1&
                      ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                      arr_varCtl(C_DID, lngE) = lngThisDbsID
                      arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                      arr_varCtl(C_CNAM, lngE) = .Name
                      arr_varCtl(C_CTYP, lngE) = .ControlType
                      arr_varCtl(C_FND, lngE) = CBool(True)
                      arr_varCtl(C_CHG, lngE) = CBool(True)
                      arr_varCtl(C_NOTE, lngE) = "CTL TYPE DIFF;"
                      lngChangedCtls = lngChangedCtls + 1&
                    End If
                    If blnTypeDiff = False Then
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel, acRectangle, acLine, acBoundObjectFrame, acImage, _
                          acCheckBox, acOptionGroup, acOptionButton, acToggleButton, acSubform, acCommandButton
                        If .Top <> ctlLnk.Top Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "TOP DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "TOP DIFF;"
                          End If
                        End If
                        If .Left <> ctlLnk.Left Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "LEFT DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "LEFT DIFF;"
                          End If
                        End If
                        If .Width <> ctlLnk.Width Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "WIDTH DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "WIDTH DIFF;"
                          End If
                        End If
                        If .Height <> ctlLnk.Height Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "HEIGHT DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "HEIGHT DIFF;"
                          End If
                        End If
                        If .Visible <> ctlLnk.Visible Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "VIS DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "VIS DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acBoundObjectFrame, acCheckBox, _
                          acOptionGroup, acOptionButton, acToggleButton, acSubform, acCommandButton
                        If .Enabled <> ctlLnk.Enabled Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "ENABLED DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "ENABLED DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acBoundObjectFrame, acCheckBox, _
                          acOptionGroup, acOptionButton, acToggleButton, acSubform
                        If .Locked <> ctlLnk.Locked Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "LOCK DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "LOCK DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acBoundObjectFrame, acCheckBox, _
                          acOptionGroup, acSubform, acCommandButton
                        If .TabStop <> ctlLnk.TabStop Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "TABSTOP DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "TABSTOP DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel, acCommandButton
                        If .ForeColor <> ctlLnk.ForeColor Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "FORECOLOR DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "FORECOLOR DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel, acRectangle, acBoundObjectFrame, acImage, acOptionGroup
                        If .BackColor <> ctlLnk.BackColor Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BACKCOLOR DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BACKCOLOR DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acLabel, acRectangle, acBoundObjectFrame, acImage, acOptionGroup
                        If .BackStyle <> ctlLnk.BackStyle Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BACKSTYLE DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BACKSTYLE DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel, acRectangle, acLine, acBoundObjectFrame, acImage, acOptionGroup
                        If .BorderColor <> ctlLnk.BorderColor Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BORDERCOLOR DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BORDERCOLOR DIFF;"
                          End If
                        End If
                        If .BorderStyle <> ctlLnk.BorderStyle Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BORDERSTYLE DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BORDERSTYLE DIFF;"
                          End If
                        End If
                        If .SpecialEffect <> ctlLnk.SpecialEffect Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "SPEC EFFECT DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SPEC EFFECT DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acCommandButton
                        If .Transparent <> ctlLnk.Transparent Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "TRANSPARENT DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "TRANSPARENT DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acLabel, acCommandButton
                        If .Caption = vbNullString Then
                          If ctlLnk.Caption <> vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CAP DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CAP DIFF;"
                            End If
                          End If
                        Else
                          If ctlLnk.Caption = vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CAP DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CAP DIFF;"
                            End If
                          Else
                            If .Caption <> ctlLnk.Caption Then
                              If lngE = -1& Then
                                lngCtls = lngCtls + 1&
                                lngE = lngCtls - 1&
                                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                                arr_varCtl(C_DID, lngE) = lngThisDbsID
                                arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                                arr_varCtl(C_CNAM, lngE) = .Name
                                arr_varCtl(C_CTYP, lngE) = .ControlType
                                arr_varCtl(C_FND, lngE) = CBool(True)
                                arr_varCtl(C_CHG, lngE) = CBool(True)
                                arr_varCtl(C_NOTE, lngE) = "CAP DIFF;"
                                lngChangedCtls = lngChangedCtls + 1&
                              Else
                                arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CAP DIFF;"
                              End If
                            End If
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acSubform
                        If ctlLnk.SourceObject = vbNullString Then
                          If .SourceObject <> vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "SRC OBJ DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SRC OBJ DIFF;"
                            End If
                          End If
                        Else
                          If .SourceObject = vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "SRC OBJ DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SRC OBJ DIFF;"
                            End If
                          Else
                            If .SourceObject <> ctlLnk.SourceObject Then
                              If lngE = -1& Then
                                lngCtls = lngCtls + 1&
                                lngE = lngCtls - 1&
                                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                                arr_varCtl(C_DID, lngE) = lngThisDbsID
                                arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                                arr_varCtl(C_CNAM, lngE) = .Name
                                arr_varCtl(C_CTYP, lngE) = .ControlType
                                arr_varCtl(C_FND, lngE) = CBool(True)
                                arr_varCtl(C_CHG, lngE) = CBool(True)
                                arr_varCtl(C_NOTE, lngE) = "SRC OBJ DIFF;"
                                lngChangedCtls = lngChangedCtls + 1&
                              Else
                                arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SRC OBJ DIFF;"
                              End If
                            End If
                          End If
                        End If
                        'LinkChildFields
                        'LinkMasterFields
                      End Select
                      Select Case .ControlType
                      Case acTextBox
                        If ctlLnk.ControlSource = vbNullString Then
                          If .ControlSource <> vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CTL SRC DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CTL SRC DIFF;"
                            End If
                          End If
                        Else
                          If .ControlSource = vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CTL SRC DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CTL SRC DIFF;"
                            End If
                          Else
                            If .ControlSource <> ctlLnk.ControlSource Then
                              If lngE = -1& Then
                                lngCtls = lngCtls + 1&
                                lngE = lngCtls - 1&
                                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                                arr_varCtl(C_DID, lngE) = lngThisDbsID
                                arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                                arr_varCtl(C_CNAM, lngE) = .Name
                                arr_varCtl(C_CTYP, lngE) = .ControlType
                                arr_varCtl(C_FND, lngE) = CBool(True)
                                arr_varCtl(C_CHG, lngE) = CBool(True)
                                arr_varCtl(C_NOTE, lngE) = "CTL SRC DIFF;"
                                lngChangedCtls = lngChangedCtls + 1&
                              Else
                                arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CTL SRC DIFF;"
                              End If
                            End If
                          End If
                        End If
                      End Select
                    End If  ' ** blnTypeDiff.
                  End If
                End With  ' ** ctlLoc.
              Next  ' ** ctlLoc.
              If blnFound = False Then
                lngCtls = lngCtls + 1&
                lngE = lngCtls - 1&
                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                arr_varCtl(C_DID, lngE) = lngThisDbsID
                arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                arr_varCtl(C_CNAM, lngE) = ctlLnk.Name
                arr_varCtl(C_CTYP, lngE) = ctlLnk.ControlType
                arr_varCtl(C_FND, lngE) = CBool(False)
                arr_varCtl(C_CHG, lngE) = CBool(False)
                arr_varCtl(C_NOTE, lngE) = "NEW CTL"
                lngNewCtls = lngNewCtls + 1&
              End If
            Next  ' ** ctlLnk.
            Set ctlLnk = Nothing
            Set ctlLoc = Nothing

          End With  ' ** frmLnk.
          Set frmLnk = Nothing
          Set frmLoc = Nothing
          .DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX)
          DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX)

        End If  ' ** F_FND.

        If ((lngX + 1&) Mod 100) = 0 Then
          Debug.Print "|  " & CStr(lngX + 1&)
          Debug.Print "'|";
        ElseIf ((lngX + 1&) Mod 10) = 0 Then
          Debug.Print "|";
        Else
          Debug.Print ".";
        End If
        DoEvents

        If (lngX + 1&) = 100 Or (lngX + 1&) = 200 Then
          Stop
        End If

      Next  ' ** lngX.
      Debug.Print
      Set frmLoc = Nothing
      Set frmLnk = Nothing

      Debug.Print "'|";
      DoEvents

      Set ctrLoc = dbsLoc.Containers("Forms")
      lngY = 0&
      For Each docLoc In ctrLoc.Documents
        lngY = lngY + 1&
        For lngX = 0& To (lngFrms - 1&)
          If arr_varFrm(F_FNAM, lngX) = docLoc.Name And arr_varFrm(F_FND, lngX) = True Then
            .DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
            Set frmLnk = .Forms(arr_varFrm(F_FNAM, lngX))
            DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
            Set frmLoc = .Forms(arr_varFrm(F_FNAM, lngX))
            With frmLoc
              For Each ctlLoc In .Controls
                blnFound = False
                For Each ctlLnk In frmLnk.Controls
                  With ctlLnk
                    If .Name = ctlLoc.Name Then
                      blnFound = True
                      Exit For
                    End If
                  End With  ' ** ctlLnk.
                Next  ' ** ctlLnk.
                If blnFound = False Then
                  arr_varFrm(F_CHG, lngX) = CBool(True)
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = lngThisDbsID
                  arr_varCtl(C_FNAM, lngE) = arr_varFrm(F_FNAM, lngX)
                  arr_varCtl(C_CNAM, lngE) = .Name
                  arr_varCtl(C_CTYP, lngE) = .ControlType
                  arr_varCtl(C_FND, lngE) = CBool(False)
                  arr_varCtl(C_CHG, lngE) = CBool(False)
                  arr_varCtl(C_NOTE, lngE) = "NEW CTL HERE!"
                  lngNewCtls = lngNewCtls + 1&
                End If
              Next  ' ** cltLoc.
            End With  ' ** frmLoc.
            Set ctlLoc = Nothing
            Set ctlLnk = Nothing
            Set frmLoc = Nothing
            Set frmLnk = Nothing
            DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX)
            .DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX)
            Exit For
          End If  ' ** F_FND.
        Next  ' ** lngX.
        If (lngY Mod 100) = 0 Then
          Debug.Print "|  " & CStr(lngY)
          Debug.Print "'|";
        ElseIf (lngY Mod 10) = 0 Then
          Debug.Print "|";
        Else
          Debug.Print ".";
        End If
        DoEvents
      Next  ' ** docLoc.
      Debug.Print
      Set docLoc = Nothing
      Set ctrLoc = Nothing

      For lngX = 0& To (lngFrms - 1&)
        If arr_varFrm(F_CHG, lngX) = True Then
          lngChangedFrms = lngChangedFrms + 1
        End If
      Next  ' ** lngX.

      Debug.Print "'CHANGED FRMS: " & CStr(lngChangedFrms)
      DoEvents

      Debug.Print "'NEW CTLS: " & CStr(lngNewCtls)
      DoEvents

      Debug.Print "'CHANGED CTLS:  " & CStr(lngChangedCtls)
      DoEvents

      Set rst = dbsLoc.OpenRecordset("zz_tbl_DataComp_04", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngFrms - 1&)
          If arr_varFrm(F_FND, lngX) = False Or arr_varFrm(F_CHG, lngX) = True Or arr_varFrm(F_NOTE, lngX) <> vbNullString Then
            .AddNew
            ' ** ![dc04_id] : AutoNumber.
            ![dbs_id] = arr_varFrm(F_DID, lngX)
            ![frm_name] = arr_varFrm(F_FNAM, lngX)
            ![frm_ctls] = arr_varFrm(F_CTLS, lngX)
            If arr_varFrm(F_CAP, lngX) <> vbNullString Then
              ![frm_caption] = arr_varFrm(F_CAP, lngX)
            End If
            ![frm_found] = arr_varFrm(F_FND, lngX)
            ![frm_changed] = arr_varFrm(F_CHG, lngX)
            If arr_varFrm(F_NOTE, lngX) <> vbNullString Then
              ![frm_note] = arr_varFrm(F_NOTE, lngX)
            End If
            ![dc04_datemodified] = Now()
            .Update
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      Set rst = dbsLoc.OpenRecordset("zz_tbl_DataComp_05", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngCtls - 1&)
          .AddNew
          ' ** ![dc05_id] : AutoNumber.
          ![dbs_id] = arr_varCtl(C_DID, lngX)
          ![frm_name] = arr_varCtl(C_FNAM, lngX)
          ![ctl_name] = arr_varCtl(C_CNAM, lngX)
          ![ctltype_type] = arr_varCtl(C_CTYP, lngX)
          ![ctl_found] = arr_varCtl(C_FND, lngX)
          ![ctl_changed] = arr_varCtl(C_CHG, lngX)
          ![ctl_note] = arr_varCtl(C_NOTE, lngX)
          ![dc05_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      dbsLoc.Close
      Set dbsLoc = Nothing

    Else
      Debug.Print "'NONE FOUND!"
      DoEvents
    End If  ' ** lngFrms.

    dbsLnk.Close
    Set dbsLnk = Nothing

    .Quit

  End With  ' ** acApp.
  Set acApp = Nothing

  blnAccessOpen = False

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  Beep

  Set rst = Nothing
  Set docLoc = Nothing
  Set docLnk = Nothing
  Set ctrLoc = Nothing
  Set ctrLnk = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set ctlLoc = Nothing
  Set ctlLnk = Nothing
  Set frmLoc = Nothing
  Set frmLnk = Nothing
  Set acApp = Nothing

  DataComp_Frm_Loc = blnRetVal

End Function

Public Function DataComp_Rpt_Loc() As Boolean

  Const THIS_PROC As String = "DataComp_Rpt_Loc"

  Dim acApp As Access.Application, rptLoc As Access.Report, rptLnk As Access.Report, rst As DAO.Recordset
  Dim dbsLoc As DAO.Database, dbsLnk As DAO.Database, ctrLoc As DAO.Container, ctrLnk As DAO.Container
  Dim docLoc As DAO.Document, docLnk As DAO.Document, ctlLoc As Access.Control, ctlLnk As Access.Control
  Dim strSysPathFile As String
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim lngCtls As Long, arr_varCtl() As Variant
  Dim lngNewRpts As Long, lngChangedRpts As Long, lngNewCtls As Long, lngChangedCtls As Long
  Dim lngThisDbsID As Long
  Dim blnAccessOpen As Boolean, blnFound As Boolean, blnTypeDiff As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const R_DID  As Integer = 0
  Const R_RNAM As Integer = 1
  Const R_CTLS As Integer = 2
  Const R_CAP  As Integer = 3
  Const R_FND  As Integer = 4
  Const R_CHG  As Integer = 5
  Const R_NOTE As Integer = 6

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const C_DID  As Integer = 0
  Const C_RNAM As Integer = 1
  Const C_CNAM As Integer = 2
  Const C_CTYP As Integer = 3
  Const C_FND  As Integer = 4
  Const C_CHG  As Integer = 5
  Const C_NOTE As Integer = 6

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strSysPathFile = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
  strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
  strFile = "Trust - Copy (20).mdb"
  strPathFile = strPath & LNK_SEP & strFile

  DBEngine.SystemDB = strSysPathFile

  Set acApp = New Access.Application

  blnAccessOpen = True

  With acApp

    .Visible = False

    .OpenCurrentDatabase strPathFile, False, TA_SEC

    lngRpts = 0&
    ReDim arr_varRpt(R_ELEMS, 0)

    Set dbsLnk = .CurrentDb
    With dbsLnk
      Set ctrLnk = .Containers("Reports")
      With ctrLnk
        For Each docLnk In .Documents
          With docLnk
            lngRpts = lngRpts + 1&
            lngE = lngRpts - 1&
            ReDim Preserve arr_varRpt(R_ELEMS, lngE)
            arr_varRpt(R_DID, lngE) = lngThisDbsID
            arr_varRpt(R_RNAM, lngE) = .Name
            arr_varRpt(R_CTLS, lngE) = 0&
            arr_varRpt(R_CAP, lngE) = vbNullString
            arr_varRpt(R_FND, lngE) = CBool(False)
            arr_varRpt(R_CHG, lngE) = CBool(False)
            arr_varRpt(R_NOTE, lngE) = vbNullString
          End With  ' ** docLnk.
        Next  ' ** docLnk.
      End With  ' ** ctrLnk.
      Set docLnk = Nothing
      Set ctrLnk = Nothing
    End With  ' ** dbsLnk.

    Debug.Print "'RPTS: " & CStr(lngRpts)
    DoEvents

    If lngRpts > 0& Then

      Set dbsLoc = CurrentDb
      Set ctrLoc = dbsLoc.Containers("Reports")

      lngNewRpts = 0&: lngChangedRpts = 0&
      For lngX = 0& To (lngRpts - 1&)
        blnFound = False
        For Each docLoc In ctrLoc.Documents
          With docLoc
            If .Name = arr_varRpt(R_RNAM, lngX) Then
              blnFound = True
              arr_varRpt(R_FND, lngX) = CBool(True)
              Exit For
            End If
          End With
        Next  ' ** docLoc.
        If blnFound = False Then
          lngNewRpts = lngNewRpts + 1&
        End If
      Next  ' ** lngX.
      Set docLoc = Nothing

      For Each docLoc In ctrLoc.Documents
        blnFound = False
        With docLoc
          For lngX = 0& To (lngRpts - 1&)
            If arr_varRpt(R_RNAM, lngX) = .Name Then
              blnFound = True
              Exit For
            End If
          Next  ' ** lngX.
          If blnFound = False Then
            lngRpts = lngRpts + 1&
            lngE = lngRpts - 1&
            ReDim Preserve arr_varRpt(R_ELEMS, lngE)
            arr_varRpt(R_DID, lngE) = lngThisDbsID
            arr_varRpt(R_RNAM, lngE) = .Name
            arr_varRpt(R_CTLS, lngE) = 0&
            arr_varRpt(R_CAP, lngE) = vbNullString
            arr_varRpt(R_FND, lngE) = CBool(False)
            arr_varRpt(R_CHG, lngE) = CBool(False)
            arr_varRpt(R_NOTE, lngE) = "NEW RPT HERE!"
            lngNewRpts = lngNewRpts + 1&
          End If
        End With  ' ** docLoc.
      Next  ' ** docLoc.
      Set docLoc = Nothing
      Set ctrLoc = Nothing

      Debug.Print "'NEW RPTS: " & CStr(lngNewRpts)
      DoEvents

      lngCtls = 0&
      ReDim arr_varCtl(C_ELEMS, 0)

      Debug.Print "'|";
      DoEvents

      lngNewCtls = 0&: lngChangedCtls = 0&
      For lngX = 0& To (lngRpts - 1&)
        If arr_varRpt(R_FND, lngX) = True Then
          .DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
          Set rptLnk = .Reports(arr_varRpt(R_RNAM, lngX))
          DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
          Set rptLoc = Reports(arr_varRpt(R_RNAM, lngX))
          With rptLnk
            If .Controls.Count <> rptLoc.Controls.Count Then
              arr_varRpt(R_CHG, lngX) = CBool(True)
              arr_varRpt(R_NOTE, lngX) = arr_varRpt(R_NOTE, lngX) & "CTL CNT DIFF;"
            End If
            If .RecordSource = vbNullString Then
              If rptLoc.RecordSource <> vbNullString Then
                arr_varRpt(R_CHG, lngX) = CBool(True)
                arr_varRpt(R_NOTE, lngX) = arr_varRpt(R_NOTE, lngX) & "REC SRC DIFF;"
              End If
            Else
              If rptLoc.RecordSource = vbNullString Then
                arr_varRpt(R_CHG, lngX) = CBool(True)
                arr_varRpt(R_NOTE, lngX) = arr_varRpt(R_NOTE, lngX) & "REC SRC DIFF;"
              Else
                If rptLoc.RecordSource <> .RecordSource Then
                  arr_varRpt(R_CHG, lngX) = CBool(True)
                  arr_varRpt(R_NOTE, lngX) = arr_varRpt(R_NOTE, lngX) & "REC SRC DIFF;"
                End If
              End If
            End If
            If .Caption = vbNullString Then
              If rptLoc.Caption <> vbNullString Then
                arr_varRpt(R_CHG, lngX) = CBool(True)
                arr_varRpt(R_NOTE, lngX) = arr_varRpt(R_NOTE, lngX) & "CAP DIFF;"
              End If
            Else
              arr_varRpt(R_CAP, lngX) = .Caption
              If rptLoc.Caption = vbNullString Then
                arr_varRpt(R_CHG, lngX) = CBool(True)
                arr_varRpt(R_NOTE, lngX) = arr_varRpt(R_NOTE, lngX) & "CAP DIFF;"
              Else
                If rptLoc.Caption <> .Caption Then
                  arr_varRpt(R_CHG, lngX) = CBool(True)
                  arr_varRpt(R_NOTE, lngX) = arr_varRpt(R_NOTE, lngX) & "CAP DIFF;"
                End If
              End If
            End If
            For Each ctlLnk In .Controls
              blnFound = False: blnTypeDiff = False
              For Each ctlLoc In rptLoc.Controls
                lngE = -1&
                With ctlLoc
                  If .Name = ctlLnk.Name Then
                    blnFound = True
                    If .ControlType <> ctlLnk.ControlType Then
                      blnTypeDiff = True
                      lngCtls = lngCtls + 1&
                      lngE = lngCtls - 1&
                      ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                      arr_varCtl(C_DID, lngE) = lngThisDbsID
                      arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                      arr_varCtl(C_CNAM, lngE) = .Name
                      arr_varCtl(C_CTYP, lngE) = .ControlType
                      arr_varCtl(C_FND, lngE) = CBool(True)
                      arr_varCtl(C_CHG, lngE) = CBool(True)
                      arr_varCtl(C_NOTE, lngE) = "CTL TYPE DIFF;"
                      lngChangedCtls = lngChangedCtls + 1&
                    End If
                    If blnTypeDiff = False Then
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel, acRectangle, acLine, acBoundObjectFrame, acImage, _
                          acCheckBox, acOptionGroup, acOptionButton, acToggleButton, acSubform
                        If .Top <> ctlLnk.Top Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "TOP DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "TOP DIFF;"
                          End If
                        End If
                        If .Left <> ctlLnk.Left Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "LEFT DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "LEFT DIFF;"
                          End If
                        End If
                        If .Width <> ctlLnk.Width Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "WIDTH DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "WIDTH DIFF;"
                          End If
                        End If
                        If .Height <> ctlLnk.Height Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "HEIGHT DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "HEIGHT DIFF;"
                          End If
                        End If
                        If .Visible <> ctlLnk.Visible Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "VIS DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "VIS DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel
                        If .ForeColor <> ctlLnk.ForeColor Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "FORECOLOR DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "FORECOLOR DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel, acRectangle, acBoundObjectFrame, acImage, acOptionGroup
                        If .BackColor <> ctlLnk.BackColor Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BACKCOLOR DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BACKCOLOR DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acLabel, acRectangle, acBoundObjectFrame, acImage, acOptionGroup
                        If .BackStyle <> ctlLnk.BackStyle Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BACKSTYLE DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BACKSTYLE DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acTextBox, acComboBox, acListBox, acLabel, acRectangle, acLine, acBoundObjectFrame, acImage, acOptionGroup
                        If .BorderColor <> ctlLnk.BorderColor Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BORDERCOLOR DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BORDERCOLOR DIFF;"
                          End If
                        End If
                        If .BorderStyle <> ctlLnk.BorderStyle Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "BORDERSTYLE DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "BORDERSTYLE DIFF;"
                          End If
                        End If
                        If .SpecialEffect <> ctlLnk.SpecialEffect Then
                          If lngE = -1& Then
                            lngCtls = lngCtls + 1&
                            lngE = lngCtls - 1&
                            ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                            arr_varCtl(C_DID, lngE) = lngThisDbsID
                            arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                            arr_varCtl(C_CNAM, lngE) = .Name
                            arr_varCtl(C_CTYP, lngE) = .ControlType
                            arr_varCtl(C_FND, lngE) = CBool(True)
                            arr_varCtl(C_CHG, lngE) = CBool(True)
                            arr_varCtl(C_NOTE, lngE) = "SPEC EFFECT DIFF;"
                            lngChangedCtls = lngChangedCtls + 1&
                          Else
                            arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SPEC EFFECT DIFF;"
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acLabel
                        If .Caption = vbNullString Then
                          If ctlLnk.Caption <> vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CAP DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CAP DIFF;"
                            End If
                          End If
                        Else
                          If ctlLnk.Caption = vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CAP DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CAP DIFF;"
                            End If
                          Else
                            If .Caption <> ctlLnk.Caption Then
                              If lngE = -1& Then
                                lngCtls = lngCtls + 1&
                                lngE = lngCtls - 1&
                                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                                arr_varCtl(C_DID, lngE) = lngThisDbsID
                                arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                                arr_varCtl(C_CNAM, lngE) = .Name
                                arr_varCtl(C_CTYP, lngE) = .ControlType
                                arr_varCtl(C_FND, lngE) = CBool(True)
                                arr_varCtl(C_CHG, lngE) = CBool(True)
                                arr_varCtl(C_NOTE, lngE) = "CAP DIFF;"
                                lngChangedCtls = lngChangedCtls + 1&
                              Else
                                arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CAP DIFF;"
                              End If
                            End If
                          End If
                        End If
                      End Select
                      Select Case .ControlType
                      Case acSubform
                        If ctlLnk.SourceObject = vbNullString Then
                          If .SourceObject <> vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "SRC OBJ DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SRC OBJ DIFF;"
                            End If
                          End If
                        Else
                          If .SourceObject = vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "SRC OBJ DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SRC OBJ DIFF;"
                            End If
                          Else
                            If .SourceObject <> ctlLnk.SourceObject Then
                              If lngE = -1& Then
                                lngCtls = lngCtls + 1&
                                lngE = lngCtls - 1&
                                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                                arr_varCtl(C_DID, lngE) = lngThisDbsID
                                arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                                arr_varCtl(C_CNAM, lngE) = .Name
                                arr_varCtl(C_CTYP, lngE) = .ControlType
                                arr_varCtl(C_FND, lngE) = CBool(True)
                                arr_varCtl(C_CHG, lngE) = CBool(True)
                                arr_varCtl(C_NOTE, lngE) = "SRC OBJ DIFF;"
                                lngChangedCtls = lngChangedCtls + 1&
                              Else
                                arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "SRC OBJ DIFF;"
                              End If
                            End If
                          End If
                        End If
                        'LinkChildFields
                        'LinkMasterFields
                      End Select
                      Select Case .ControlType
                      Case acTextBox
                        If ctlLnk.ControlSource = vbNullString Then
                          If .ControlSource <> vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CTL SRC DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CTL SRC DIFF;"
                            End If
                          End If
                        Else
                          If .ControlSource = vbNullString Then
                            If lngE = -1& Then
                              lngCtls = lngCtls + 1&
                              lngE = lngCtls - 1&
                              ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                              arr_varCtl(C_DID, lngE) = lngThisDbsID
                              arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                              arr_varCtl(C_CNAM, lngE) = .Name
                              arr_varCtl(C_CTYP, lngE) = .ControlType
                              arr_varCtl(C_FND, lngE) = CBool(True)
                              arr_varCtl(C_CHG, lngE) = CBool(True)
                              arr_varCtl(C_NOTE, lngE) = "CTL SRC DIFF;"
                              lngChangedCtls = lngChangedCtls + 1&
                            Else
                              arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CTL SRC DIFF;"
                            End If
                          Else
                            If .ControlSource <> ctlLnk.ControlSource Then
                              If lngE = -1& Then
                                lngCtls = lngCtls + 1&
                                lngE = lngCtls - 1&
                                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                                arr_varCtl(C_DID, lngE) = lngThisDbsID
                                arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                                arr_varCtl(C_CNAM, lngE) = .Name
                                arr_varCtl(C_CTYP, lngE) = .ControlType
                                arr_varCtl(C_FND, lngE) = CBool(True)
                                arr_varCtl(C_CHG, lngE) = CBool(True)
                                arr_varCtl(C_NOTE, lngE) = "CTL SRC DIFF;"
                                lngChangedCtls = lngChangedCtls + 1&
                              Else
                                arr_varCtl(C_NOTE, lngE) = arr_varCtl(C_NOTE, lngE) & "CTL SRC DIFF;"
                              End If
                            End If
                          End If
                        End If
                      End Select
                    End If  ' ** blnTypeDiff.
                  End If
                End With  ' ** ctlLoc.
              Next  ' ** ctlLoc.
              If blnFound = False Then
                lngCtls = lngCtls + 1&
                lngE = lngCtls - 1&
                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                arr_varCtl(C_DID, lngE) = lngThisDbsID
                arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                arr_varCtl(C_CNAM, lngE) = ctlLnk.Name
                arr_varCtl(C_CTYP, lngE) = ctlLnk.ControlType
                arr_varCtl(C_FND, lngE) = CBool(False)
                arr_varCtl(C_CHG, lngE) = CBool(False)
                arr_varCtl(C_NOTE, lngE) = "NEW CTL"
                lngNewCtls = lngNewCtls + 1&
              End If
            Next  ' ** ctlLnk.
            Set ctlLnk = Nothing
            Set ctlLoc = Nothing

          End With  ' ** rptLnk.
          Set rptLnk = Nothing
          Set rptLoc = Nothing
          .DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX)
          DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX)

        End If  ' ** R_FND.

        If ((lngX + 1&) Mod 100) = 0 Then
          Debug.Print "|  " & CStr(lngX + 1&)
          Debug.Print "'|";
        ElseIf ((lngX + 1&) Mod 10) = 0 Then
          Debug.Print "|";
        Else
          Debug.Print ".";
        End If
        DoEvents

        If (lngX + 1&) = 100 Or (lngX + 1&) = 200 Then
          Stop
        End If

      Next  ' ** lngX.
      Debug.Print
      Set rptLoc = Nothing
      Set rptLnk = Nothing

      Debug.Print "'|";
      DoEvents

      Set ctrLoc = dbsLoc.Containers("Reports")
      lngY = 0&
      For Each docLoc In ctrLoc.Documents
        lngY = lngY + 1&
        For lngX = 0& To (lngRpts - 1&)
          If arr_varRpt(R_RNAM, lngX) = docLoc.Name And arr_varRpt(R_FND, lngX) = True Then
            .DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
            Set rptLnk = .Reports(arr_varRpt(R_RNAM, lngX))
            DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
            Set rptLoc = .Reports(arr_varRpt(R_RNAM, lngX))
            With rptLoc
              For Each ctlLoc In .Controls
                blnFound = False
                For Each ctlLnk In rptLnk.Controls
                  With ctlLnk
                    If .Name = ctlLoc.Name Then
                      blnFound = True
                      Exit For
                    End If
                  End With  ' ** ctlLnk.
                Next  ' ** ctlLnk.
                If blnFound = False Then
                  arr_varRpt(R_CHG, lngX) = CBool(True)
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = lngThisDbsID
                  arr_varCtl(C_RNAM, lngE) = arr_varRpt(R_RNAM, lngX)
                  arr_varCtl(C_CNAM, lngE) = .Name
                  arr_varCtl(C_CTYP, lngE) = .ControlType
                  arr_varCtl(C_FND, lngE) = CBool(False)
                  arr_varCtl(C_CHG, lngE) = CBool(False)
                  arr_varCtl(C_NOTE, lngE) = "NEW CTL HERE!"
                  lngNewCtls = lngNewCtls + 1&
                End If
              Next  ' ** cltLoc.
            End With  ' ** rptLoc.
            Set ctlLoc = Nothing
            Set ctlLnk = Nothing
            Set rptLoc = Nothing
            Set rptLnk = Nothing
            DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX)
            .DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX)
            Exit For
          End If  ' ** R_FND.
        Next  ' ** lngX.
        If (lngY Mod 100) = 0 Then
          Debug.Print "|  " & CStr(lngY)
          Debug.Print "'|";
        ElseIf (lngY Mod 10) = 0 Then
          Debug.Print "|";
        Else
          Debug.Print ".";
        End If
        DoEvents
      Next  ' ** docLoc.
      Debug.Print
      Set docLoc = Nothing
      Set ctrLoc = Nothing

      For lngX = 0& To (lngRpts - 1&)
        If arr_varRpt(R_CHG, lngX) = True Then
          lngChangedRpts = lngChangedRpts + 1
        End If
      Next  ' ** lngX.

      Debug.Print "'CHANGED RPTS: " & CStr(lngChangedRpts)
      DoEvents

      Debug.Print "'NEW CTLS: " & CStr(lngNewCtls)
      DoEvents

      Debug.Print "'CHANGED CTLS:  " & CStr(lngChangedCtls)
      DoEvents

      Set rst = dbsLoc.OpenRecordset("zz_tbl_DataComp_06", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngRpts - 1&)
          If arr_varRpt(R_FND, lngX) = False Or arr_varRpt(R_CHG, lngX) = True Or arr_varRpt(R_NOTE, lngX) <> vbNullString Then
            .AddNew
            ' ** ![dc06_id] : AutoNumber.
            ![dbs_id] = arr_varRpt(R_DID, lngX)
            ![rpt_name] = arr_varRpt(R_RNAM, lngX)
            ![rpt_ctls] = arr_varRpt(R_CTLS, lngX)
            If arr_varRpt(R_CAP, lngX) <> vbNullString Then
              ![rpt_caption] = arr_varRpt(R_CAP, lngX)
            End If
            ![rpt_found] = arr_varRpt(R_FND, lngX)
            ![rpt_changed] = arr_varRpt(R_CHG, lngX)
            If arr_varRpt(R_NOTE, lngX) <> vbNullString Then
              ![rpt_note] = arr_varRpt(R_NOTE, lngX)
            End If
            ![dc06_datemodified] = Now()
            .Update
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      Set rst = dbsLoc.OpenRecordset("zz_tbl_DataComp_07", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngCtls - 1&)
          .AddNew
          ' ** ![dc07_id] : AutoNumber.
          ![dbs_id] = arr_varCtl(C_DID, lngX)
          ![rpt_name] = arr_varCtl(C_RNAM, lngX)
          ![ctl_name] = arr_varCtl(C_CNAM, lngX)
          ![ctltype_type] = arr_varCtl(C_CTYP, lngX)
          ![ctl_found] = arr_varCtl(C_FND, lngX)
          ![ctl_changed] = arr_varCtl(C_CHG, lngX)
          ![ctl_note] = arr_varCtl(C_NOTE, lngX)
          ![dc07_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      dbsLoc.Close
      Set dbsLoc = Nothing

    Else
      Debug.Print "'NONE FOUND!"
      DoEvents
    End If  ' ** lngRpts.

    dbsLnk.Close
    Set dbsLnk = Nothing

    .Quit

  End With  ' ** acApp.
  Set acApp = Nothing

  blnAccessOpen = False

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  Beep

  Set rst = Nothing
  Set docLoc = Nothing
  Set docLnk = Nothing
  Set ctrLoc = Nothing
  Set ctrLnk = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set ctlLoc = Nothing
  Set ctlLnk = Nothing
  Set rptLoc = Nothing
  Set rptLnk = Nothing
  Set acApp = Nothing

  DataComp_Rpt_Loc = blnRetVal

End Function

Public Function DataComp_Mod_Loc() As Boolean

  Const THIS_PROC As String = "DataComp_Mod_Loc"

  Dim acApp As Access.Application, vbpLoc As VBProject, vbpLnk As VBProject, vbcLoc As VBComponent, vbcLnk As VBComponent
  Dim codLoc As CodeModule, codLnk As CodeModule, dbs As DAO.Database, rst As DAO.Recordset
  Dim strSysPathFile As String
  Dim strPath As String, strFile As String, strPathFile As String
  Dim lngMods As Long, arr_varMod() As Variant
  Dim lngNewMods As Long, lngChangedMods As Long, lngDiffs As Long, strDiffs As String
  Dim lngThisDbsID As Long
  Dim blnAccessOpen As Boolean, blnFound As Boolean
  Dim lngTmp01 As Long, lngTmp02 As Long
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varMod().
  Const M_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const M_DID  As Integer = 0
  Const M_VNAM As Integer = 1
  Const M_VTYP As Integer = 2
  Const M_LINS As Integer = 3
  Const M_DECS As Integer = 4
  Const M_FND  As Integer = 5
  Const M_CHG  As Integer = 6
  Const M_NOTE As Integer = 7

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strSysPathFile = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
  strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
  strFile = "Trust - Copy (20).mdb"
  strPathFile = strPath & LNK_SEP & strFile

  DBEngine.SystemDB = strSysPathFile

  Set acApp = New Access.Application

  blnAccessOpen = True

  With acApp

    .Visible = False

    .OpenCurrentDatabase strPathFile, False, TA_SEC

    lngMods = 0&
    ReDim arr_varMod(M_ELEMS, 0)

    Set vbpLnk = .VBE.ActiveVBProject
    With vbpLnk
      For Each vbcLnk In .VBComponents
        With vbcLnk
          lngMods = lngMods + 1&
          lngE = lngMods - 1&
          ReDim Preserve arr_varMod(M_ELEMS, lngE)
          arr_varMod(M_DID, lngE) = lngThisDbsID
          arr_varMod(M_VNAM, lngE) = .Name
          arr_varMod(M_VTYP, lngE) = .Type
          Set codLnk = .CodeModule
          With codLnk
            arr_varMod(M_LINS, lngE) = .CountOfLines
            arr_varMod(M_DECS, lngE) = .CountOfDeclarationLines
          End With  ' ** codLnk.
          arr_varMod(M_FND, lngE) = CBool(False)
          arr_varMod(M_CHG, lngE) = CBool(False)
          arr_varMod(M_NOTE, lngE) = vbNullString
        End With  ' ** vbcLnk.
      Next  ' ** vbcLnk.
    End With  ' ** vbpLnk.
    Set codLnk = Nothing
    Set vbcLnk = Nothing

    Debug.Print "'MODS: " & CStr(lngMods)
    DoEvents

    If lngMods > 0& Then

      lngNewMods = 0&: lngChangedMods = 0&
      Set vbpLoc = Application.VBE.ActiveVBProject
      For lngX = 0& To (lngMods - 1&)
        blnFound = False
        For Each vbcLoc In vbpLoc.VBComponents
          With vbcLoc
            If .Name = arr_varMod(M_VNAM, lngX) Then
              blnFound = True
              arr_varMod(M_FND, lngX) = CBool(True)
              Set codLoc = .CodeModule
              With codLoc
                If .CountOfLines <> arr_varMod(M_LINS, lngX) Then
                  arr_varMod(M_CHG, lngX) = CBool(True)
                  arr_varMod(M_NOTE, lngX) = "LINE CNT DIFF;"
                  lngChangedMods = lngChangedMods + 1&
                End If
                If .CountOfLines = arr_varMod(M_LINS, lngX) And .CountOfDeclarationLines <> arr_varMod(M_DECS, lngX) Then
                  arr_varMod(M_CHG, lngX) = CBool(True)
                  arr_varMod(M_NOTE, lngX) = "DEC CNT DIFF;"
                  lngChangedMods = lngChangedMods + 1&
                End If
              End With  ' ** codLoc.
              Exit For
            End If
          End With
        Next  ' ** vbcLoc.
        Set codLoc = Nothing
        Set vbcLoc = Nothing
        If blnFound = False Then
          lngNewMods = lngNewMods + 1&
        End If
      Next  ' ** lngX.
      Set codLoc = Nothing
      Set vbcLoc = Nothing

      For Each vbcLoc In vbpLoc.VBComponents
        With vbcLoc
          blnFound = False
          For lngX = 0& To (lngMods - 1&)
            If arr_varMod(M_VNAM, lngX) = .Name Then
              blnFound = True
              Exit For
            End If
          Next  ' ** lngx.
          If blnFound = False Then
            lngMods = lngMods + 1&
            lngE = lngMods - 1&
            ReDim Preserve arr_varMod(M_ELEMS, lngE)
            arr_varMod(M_DID, lngE) = lngThisDbsID
            arr_varMod(M_VNAM, lngE) = .Name
            arr_varMod(M_VTYP, lngE) = .Type
            Set codLoc = .CodeModule
            With codLoc
              arr_varMod(M_LINS, lngE) = .CountOfLines
              arr_varMod(M_DECS, lngE) = .CountOfDeclarationLines
            End With  ' ** codLnk.
            arr_varMod(M_FND, lngE) = CBool(False)
            arr_varMod(M_CHG, lngE) = CBool(False)
            arr_varMod(M_NOTE, lngE) = "NEW MOD HERE!"
            lngNewMods = lngNewMods + 1&
          End If
        End With  ' ** vbcLoc.
      Next  ' ** vbcLoc
      Set codLoc = Nothing
      Set vbcLoc = Nothing

      Debug.Print "'NEW MODS: " & CStr(lngNewMods)
      DoEvents

      For lngX = 0& To (lngMods - 1&)
        If arr_varMod(M_FND, lngX) = True Then
          Set vbcLnk = vbpLnk.VBComponents(arr_varMod(M_VNAM, lngX))
          Set vbcLoc = vbpLoc.VBComponents(arr_varMod(M_VNAM, lngX))
          Set codLnk = vbcLnk.CodeModule
          Set codLoc = vbcLoc.CodeModule
          If codLoc.CountOfLines = codLnk.CountOfLines Then
            With codLoc
              lngDiffs = 0&: strDiffs = vbNullString
              For lngY = 1& To arr_varMod(M_LINS, lngX)
                If .Lines(lngY, 1) <> codLnk.Lines(lngY, 1) Then
                  lngDiffs = lngDiffs + 1&
                  strDiffs = strDiffs & CStr(lngY) & ","
                End If
                If lngDiffs > 20& Then
                  ' ** Too many.
                  lngDiffs = 0&
                  strDiffs = vbNullString
                  arr_varMod(M_NOTE, lngX) = arr_varMod(M_NOTE, lngX) & "MANY DIFFS;"
                  If arr_varMod(M_CHG, lngX) = False Then
                    arr_varMod(M_CHG, lngX) = CBool(True)
                    lngChangedMods = lngChangedMods + 1&
                  End If
                  Exit For
                End If
              Next  ' ** lngY.
              If lngDiffs > 0& Then
                arr_varMod(M_NOTE, lngX) = arr_varMod(M_NOTE, lngX) & strDiffs & ";"
                If arr_varMod(M_CHG, lngX) = False Then
                  arr_varMod(M_CHG, lngX) = CBool(True)
                  lngChangedMods = lngChangedMods + 1&
                End If
              End If
            End With  ' ** codLoc.
          Else
            lngTmp01 = codLoc.CountOfLines
            lngTmp02 = codLnk.CountOfLines
            If Abs(lngTmp01 - lngTmp02) <= 5 Then
              ' ** Let's see if there's just a couple of minor differences.
              lngDiffs = 0&: strDiffs = vbNullString
              If lngTmp01 > lngTmp02 Then
                With codLnk
                  ' ** This really isn't going to work!
                  For lngY = 1& To lngTmp02
                    If .Lines(lngY, 1) <> codLoc.Lines(lngY, 1) Then
                      lngDiffs = lngDiffs + 1&
                      strDiffs = strDiffs & CStr(lngY) & ","
                    End If
                    If lngDiffs > 20& Then
                      Exit For
                    End If
                  Next  ' ** lngY.
                End With  ' ** codLnk.
                If lngDiffs > 0& Then
                  arr_varMod(M_NOTE, lngX) = arr_varMod(M_NOTE, lngX) & CStr(lngDiffs) & ": " & strDiffs & ";"
                  If arr_varMod(M_CHG, lngX) = False Then
                    arr_varMod(M_CHG, lngX) = CBool(True)
                    lngChangedMods = lngChangedMods + 1&
                  End If
                End If
              Else
                With codLoc
                  ' ** This really isn't going to work!
                  For lngY = 1& To lngTmp01
                    If .Lines(lngY, 1) <> codLnk.Lines(lngY, 1) Then
                      lngDiffs = lngDiffs + 1&
                      strDiffs = strDiffs & CStr(lngY) & ","
                    End If
                    If lngDiffs > 20& Then
                      Exit For
                    End If
                  Next  ' ** lngY.
                End With  ' ** codLoc.
                If lngDiffs > 0& Then
                  arr_varMod(M_NOTE, lngX) = arr_varMod(M_NOTE, lngX) & CStr(lngDiffs) & ": " & strDiffs & ";"
                  If arr_varMod(M_CHG, lngX) = False Then
                    arr_varMod(M_CHG, lngX) = CBool(True)
                    lngChangedMods = lngChangedMods + 1&
                  End If
                End If
              End If
            Else
              ' ** Too much difference.
              arr_varMod(M_NOTE, lngX) = arr_varMod(M_NOTE, lngX) & "MANY DIFFS;"
              If arr_varMod(M_CHG, lngX) = False Then
                arr_varMod(M_CHG, lngX) = CBool(True)
                lngChangedMods = lngChangedMods + 1&
              End If
            End If
          End If
        End If  ' ** M_FND.
      Next  ' ** lngX.
      Set codLoc = Nothing
      Set codLnk = Nothing
      Set vbcLoc = Nothing
      Set vbcLnk = Nothing
      Set vbpLoc = Nothing
      Set vbpLnk = Nothing

      Debug.Print "'CHANGED MODS: " & CStr(lngChangedMods)
      DoEvents

      Set dbs = CurrentDb
      With dbs
        Set rst = .OpenRecordset("zz_tbl_DataComp_08", dbOpenDynaset, dbAppendOnly)
        With rst
          For lngX = 0& To (lngMods - 1&)
            If arr_varMod(M_FND, lngX) = False Or arr_varMod(M_CHG, lngX) = True Or arr_varMod(M_NOTE, lngX) <> vbNullString Then
              .AddNew
              '![dc08_id] : AutoNumber.
              ![dbs_id] = arr_varMod(M_DID, lngX)
              ![vbcom_name] = arr_varMod(M_VNAM, lngX)
              ![comtype_type] = arr_varMod(M_VTYP, lngX)
              ![vbcom_lines] = arr_varMod(M_LINS, lngX)
              ![vbcom_declines] = arr_varMod(M_DECS, lngX)
              ![vbcom_found] = arr_varMod(M_FND, lngX)
              ![vbcom_changed] = arr_varMod(M_CHG, lngX)
              ![vbcom_note] = arr_varMod(M_NOTE, lngX)
              ![dc08_datemodified] = Now()
              .Update
            End If
          Next  ' ** lngX.
          .Close
        End With  ' ** rst.
        .Close
      End With  ' ** dbs.
      Set rst = Nothing
      Set dbs = Nothing

    Else
      Debug.Print "'NONE FOUND!"
      DoEvents
    End If  ' ** lngMods.

    .Quit
  End With  ' ** acApp.
  Set acApp = Nothing

  blnAccessOpen = False

'MODS: 542
'NEW MODS: 2
'CHANGED MODS: 33
'DONE!  DataComp_Mod_Loc()
  Debug.Print "'DONE!  " & THIS_PROC & "()"
  Beep

  Set rst = Nothing
  Set dbs = Nothing
  Set codLoc = Nothing
  Set codLnk = Nothing
  Set vbcLoc = Nothing
  Set vbcLnk = Nothing
  Set vbpLoc = Nothing
  Set vbpLnk = Nothing
  Set acApp = Nothing

  DataComp_Mod_Loc = blnRetVal

End Function

Public Function FindTbls() As Boolean

  Const THIS_PROC As String = "FindTbls"

  Dim fso As Scripting.FileSystemObject, fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder, fsfds As Scripting.Folders
  Dim fsfls As Scripting.Files, fsfl As Scripting.File
  Dim wrk As DAO.Workspace, dbsLoc As DAO.Database, dbsLnk As DAO.Database, tdf As DAO.TableDef
  Dim strPath As String, strFile As String, strPathFile As String, strExt As String
  Dim strSysPathFile As String, strPathBase As String, strFileBase As String
  Dim lngFiles As Long, arr_varFile() As Variant
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim strTableName As String, lngHits As Long
  Dim lngThisDbsID As Long
  Dim blnSkip As Boolean
  Dim intLen As Integer
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varFile().
  Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const F_FNAM As Integer = 0
  Const F_PATH As Integer = 1
  Const F_FND  As Integer = 2

  ' ** Array: arr_varFile().
  Const T_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const T_TNAM As Integer = 0
  Const T_FNAM As Integer = 1
  Const T_PATH As Integer = 2

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strSysPathFile = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
  'strPathBase = CurrentAppPath  ' ** Module Function: modFileUtilities.
  'strPathBase = strPathBase & LNK_SEP & "Backups"
  strPathBase = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Ver2230_Bak"
  strFileBase = "Trust"
  intLen = Len(strFileBase)

  DBEngine.SystemDB = strSysPathFile

  lngFiles = 0&
  ReDim arr_varFile(F_ELEMS, 0)

  Set fso = CreateObject("Scripting.FileSystemObject")
  With fso

    Set fsfd1 = .GetFolder(strPathBase)
    With fsfd1

      ' ** First, get all the files in the root backups folder.
      Set fsfls = fsfd1.Files
      For Each fsfl In fsfls
        With fsfl
          strPath = .Path
          If InStr(strPath, ".mdb") > 0 Then strPath = Parse_Path(strPath)  ' ** Module Function: modFileUtilities.
          strFile = .Name
          strPathFile = strPath & LNK_SEP & strFile
          strExt = fso.GetExtensionName(strPathFile)
          If strExt = "mdb" And Left(strFile, intLen) = strFileBase Then
            lngFiles = lngFiles + 1&
            lngE = lngFiles - 1&
            ReDim Preserve arr_varFile(F_ELEMS, lngE)
            arr_varFile(F_FNAM, lngE) = strFile
            arr_varFile(F_PATH, lngE) = strPath
            arr_varFile(F_FND, lngE) = CBool(False)
          End If
        End With  ' ** fsfl.
      Next  ' ** fsfl.
      Set fsfl = Nothing
      Set fsfls = Nothing

      Debug.Print "'ROOT FILES: " & CStr(lngFiles)
      DoEvents

      ' ** Now get the files from the subfolders.
      Set fsfds = .SubFolders
      For Each fsfd2 In fsfds
        With fsfd2
          Set fsfls = .Files
          For Each fsfl In fsfls
            With fsfl
              strPath = .Path
              If InStr(strPath, ".mdb") > 0 Then strPath = Parse_Path(strPath)  ' ** Module Function: modFileUtilities.
              strFile = .Name
              strPathFile = strPath & LNK_SEP & strFile
              strExt = fso.GetExtensionName(strPathFile)
              If strExt = "mdb" And Left(strFile, intLen) = strFileBase Then
                lngFiles = lngFiles + 1&
                lngE = lngFiles - 1&
                ReDim Preserve arr_varFile(F_ELEMS, lngE)
                arr_varFile(F_FNAM, lngE) = strFile
                arr_varFile(F_PATH, lngE) = strPath
                arr_varFile(F_FND, lngE) = CBool(False)
              End If
            End With  ' ** fsfl.
          Next  ' ** fsfl.
          Set fsfl = Nothing
          Set fsfls = Nothing
        End With  ' ** fsfd2
      Next  ' ** fsfd2.
      Set fsfd2 = Nothing
      Set fsfds = Nothing

      Debug.Print "'TOT FILES: " & CStr(lngFiles)
      DoEvents

    End With  ' ** fsfd1.
    Set fsfd1 = Nothing
  End With  ' ** fso.
  Set fso = Nothing

  If lngFiles > 0& Then

    strTableName = "zz_tbl_DataComp_"
    intLen = Len(strTableName)

    lngTbls = 0&
    ReDim arr_varTbl(T_ELEMS, 0)

    Debug.Print "'|";
    DoEvents

    Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)
    For lngX = 0& To (lngFiles - 1&)
      strPath = arr_varFile(F_PATH, lngX)
      strFile = arr_varFile(F_FNAM, lngX)
      strPathFile = strPath & LNK_SEP & strFile
      Set dbsLnk = wrk.OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
      With dbsLnk
        For Each tdf In .TableDefs
          With tdf
            If Left(.Name, intLen) = strTableName Then
              Beep
              arr_varFile(F_FND, lngX) = CBool(True)
              lngTbls = lngTbls + 1&
              lngE = lngTbls - 1&
              ReDim Preserve arr_varTbl(T_ELEMS, lngE)
              arr_varTbl(T_TNAM, lngE) = .Name
              arr_varTbl(T_FNAM, lngE) = strFile
              arr_varTbl(T_PATH, lngE) = strPath
            End If
          End With  ' ** tdf.
        Next  ' ** tdf.
        Set tdf = Nothing
        .Close
      End With  ' ** dbsLnk.
      Set dbsLnk = Nothing
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
    wrk.Close
    Set wrk = Nothing
    Debug.Print

    Debug.Print "TBLS FOUND: " & CStr(lngTbls)
    DoEvents

    For lngX = 0& To (lngFiles - 1&)
      If arr_varFile(F_FND, lngX) = True Then
        lngHits = lngHits + 1&
        Debug.Print "'FND: " & arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FNAM, lngX)
      End If
    Next  ' ** lngX.

  End If  ' ** lngFiles.

  Debug.Print "'HITS: " & CStr(lngHits)
  DoEvents

'NONE FOUND ANYWHERE!
'zz_tbl_DataComp_01
'  ![dc01_id]
'  ![dbs_id]
'  ![tbl_name]
'  ![tbl_fld_cnt]
'  ![tbl_found]
'  ![tbl_change]
'  ![tbl_note]
'  ![dc01_datemodified]
'mm/dd/yyyy hh:nn:ss
'zz_tbl_DataComp_02
'  ![dc02_id]
'  ![dbs_id]
'  ![tbl_name]
'  ![fld_name]
'  ![fld_note]
'  ![dc02_datemodified]
'mm/dd/yyyy hh:nn:ss
'zz_tbl_DataComp_03
'  ![dc03_id]
'  ![dbs_id]
'  ![qry_name]
'  ![qrytype_type]
'  ![qry_sql]
'  ![qry_description]
'  ![qry_found]
'  ![qry_change]
'  ![qry_note]
'  ![dc03_datemodified]
'mm/dd/yyyy hh:nn:ss
'zz_tbl_DataComp_04
'  ![dc04_id]
'  ![dbs_id]
'  ![frm_name]
'  ![frm_ctls]
'  ![frm_caption]
'  ![frm_found]
'  ![frm_changed]
'  ![frm_note]
'  ![dc04_datemodified]
'mm/dd/yyyy hh:nn:ss
'zz_tbl_DataComp_05
'  ![dc05_id]
'  ![dbs_id]
'  ![frm_name]
'  ![ctl_name]
'  ![ctltype_type]
'  ![ctl_found]
'  ![ctl_changed]
'  ![ctl_note]
'  ![dc05_datemodified]
'mm/dd/yyyy hh:nn:ss
'zz_tbl_DataComp_06
'  ![dc06_id]
'  ![dbs_id]
'  ![rpt_name]
'  ![rpt_ctls]
'  ![rpt_caption]
'  ![rpt_found]
'  ![rpt_changed]
'  ![rpt_note]
'  ![dc06_datemodified]
'mm/dd/yyyy hh:nn:ss
'zz_tbl_DataComp_07
'  ![dc07_id]
'  ![dbs_id]
'  ![rpt_name]
'  ![ctl_name]
'  ![ctltype_type]
'  ![ctl_found]
'  ![ctl_changed]
'  ![ctl_note]
'  ![dc07_datemodified]
'mm/dd/yyyy hh:nn:ss
'zz_tbl_DataComp_08
'  ![dc08_id]
'  ![dbs_id]
'  ![vbcom_name]
'  ![comtype_type]
'  ![vbcom_lines]
'  ![vbcom_declines]
'  ![vbcom_found]
'  ![vbcom_changed]
'  ![vbcom_note]
'  ![dc08_datemodified]
'mm/dd/yyyy hh:nn:ss

  Beep
  Debug.Print "'DONE!"

  Set tdf = Nothing
  Set dbsLoc = Nothing
  Set dbsLnk = Nothing
  Set wrk = Nothing
  Set fsfl = Nothing
  Set fsfls = Nothing
  Set fsfd1 = Nothing
  Set fsfd2 = Nothing
  Set fsfds = Nothing
  Set fso = Nothing

  FindTbls = blnRetVal

End Function
