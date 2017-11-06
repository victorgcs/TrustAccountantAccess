Attribute VB_Name = "zz_mod_ModuleDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_ModuleDocFuncs"

'VGC 09/26/2017: CHANGES!

Public Const vbUndeclared As Integer = -3  ' ** (my own)

Private blnRetValx As Boolean  ' ** Universal replacement.
' **

Public Function QuikVBADoc() As Boolean
  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If VBA_ChkDocQrys(False) = True Then  ' ** Function: Below.
      blnRetValx = VBA_Component_Doc  ' ** Function: Below.
      blnRetValx = VBA_Proc_Doc_New  ' ** Function: Below.
      blnRetValx = VBA_MsgBox_Doc  ' ** Function: Below.
      blnRetValx = VBA_Component_API_Doc  ' ** Function: Below.
      blnRetValx = VBA_ExportAll  ' ** Module Function: zz_mod_MDEPrepFuncs.
      'blnRetValx = VBA_WinDialog_Doc  ' ** Function: Below.
      'blnRetValx = VBA_PublicVar_Doc  ' ** Function: Below.
      'blnRetValx = VBA_PublicUsage_Doc  ' ** Function: Below.
      DoEvents
      DoBeeps 4  ' ** Module Function: modWindowFunctions.
      Debug.Print "'FINISHED!"
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED VBA_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikVBADoc = blnRetValx
End Function

Private Function VBA_Component_Doc() As Boolean
' ** Document VBE components in this project to tblVBComponent.
' ** Called by:
' **   QuikVBADoc(), Above

  Const THIS_PROC As String = "VBA_Component_Doc"

  Dim vbp As VBProject, vbc As VBComponent
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngComs As Long, arr_varCom() As Variant
  Dim intType As Integer, strModName As String, strName As String
  Dim lngObjType As Long, lngThisDbsID As Long
  Dim lngRecs As Long
  Dim blnFound As Boolean, blnDelete As Boolean
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varCom().
  Const C_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const C_DID   As Integer = 0
  Const C_DNAM  As Integer = 1
  Const C_VID   As Integer = 2
  Const C_VNAM  As Integer = 3
  Const C_VTYP  As Integer = 4
  Const C_OTYP  As Integer = 5
  Const C_LINS  As Integer = 6
  Const C_VNAM2 As Integer = 7
  Const C_FND   As Integer = 8
  Const C_CHNG  As Integer = 9

  blnRetValx = True

  'Application.VBE.ActiveVBProject.VBComponents.Count = 201

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  Set rst = dbs.OpenRecordset("tblVBComponent", dbOpenDynaset, dbConsistent)

  lngComs = 0&
  ReDim arr_varCom(C_ELEMS, 0)

  ' ** Get existing list of VBA components.
  With rst
    If .BOF = True And .EOF = True Then
      ' ** Components not yet added to table.
    Else
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      ' ***************************************************
      ' ** Array: arr_varCom()
      ' **
      ' **   Field  Element  Name              Constant
      ' **   =====  =======  ================  ==========
      ' **     1       0     dbs_id            C_DID
      ' **     2       1     dbs_name          C_DNAM
      ' **     3       2     vbcom_id          C_VID
      ' **     4       3     vbcom_name        C_VNAM
      ' **     5       4     comtype_type      C_VTYP
      ' **     6       5     objtype_type      C_OTYP
      ' **     7       6     vbcom_lines       C_LINS
      ' **     8       7     vbcom_name2       C_VNAM2
      ' **     9       8     Found Yes/No      C_FND
      ' **    10       9     Changed Yes/No    C_CHNG
      ' **
      ' ***************************************************
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngComs = lngComs + 1&
          lngE = lngComs - 1&
          ReDim Preserve arr_varCom(C_ELEMS, lngE)
          arr_varCom(C_DID, lngE) = ![dbs_id]
          arr_varCom(C_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
          arr_varCom(C_VID, lngE) = ![vbcom_id]
          arr_varCom(C_VNAM, lngE) = ![vbcom_name]
          arr_varCom(C_VTYP, lngE) = ![comtype_type]
          arr_varCom(C_OTYP, lngE) = ![objtype_type]
          arr_varCom(C_LINS, lngE) = Nz(![vbcom_lines], 0&)
          arr_varCom(C_VNAM2, lngE) = ![vbcom_name2]
          arr_varCom(C_FND, lngE) = CBool(False)
          arr_varCom(C_CHNG, lngE) = CBool(False)
        End If
        If lngX < lngRecs Then .MoveNext
      Next
    End If
  End With

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        intType = .Type
        If Left$(strModName, 5) = "Form_" Then
          strName = Mid$(strModName, 6)
          lngObjType = acForm
        ElseIf Left$(strModName, 7) = "Report_" Then
          strName = Mid$(strModName, 8)
          lngObjType = acReport
        Else
          strName = strModName
          lngObjType = acModule
        End If
        blnFound = False
        For lngX = 0& To (lngComs - 1&)
          If arr_varCom(C_VNAM, lngX) = strModName Then
            blnFound = True
            arr_varCom(C_FND, lngX) = CBool(True)
            If arr_varCom(C_VTYP, lngX) <> intType Or arr_varCom(C_OTYP, lngX) <> lngObjType Or arr_varCom(C_VNAM2, lngX) <> strName Then
              arr_varCom(C_VTYP, lngX) = intType
              arr_varCom(C_OTYP, lngX) = lngObjType
              arr_varCom(C_VNAM2, lngX) = strName
              arr_varCom(C_CHNG, lngX) = CBool(True)
            End If
            If arr_varCom(C_LINS, lngX) <> .CodeModule.CountOfLines Then
              arr_varCom(C_LINS, lngX) = .CodeModule.CountOfLines
              arr_varCom(C_CHNG, lngX) = CBool(True)
            End If
            Exit For
          End If
        Next
        With rst
          If blnFound = False Then
            .AddNew
            ![dbs_id] = lngThisDbsID
            ![vbcom_name] = strModName
            ![comtype_type] = intType
            ' **   vbext_ComponentType enumeration:
            ' **       1  vbext_ct_StdModule        Standard Module
            ' **       2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
            ' **       3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm in the VBA Editor.
            ' **      11  vbext_ct_ActiveXDesigner
            ' **     100  vbext_ct_Document         Module behind Form, Report, or Excel Worksheet.
            If Left$(strModName, 5) = "Form_" Then
              strName = Mid$(strModName, 6)
              ![vbcom_name2] = strName
              ![objtype_type] = acForm
            ElseIf Left$(strModName, 7) = "Report_" Then
              strName = Mid$(strModName, 8)
              ![vbcom_name2] = strName
              ![objtype_type] = acReport
            Else
              strName = strModName
              ![vbcom_name2] = strName
              ![objtype_type] = acModule
            End If
            ![vbcom_lines] = vbc.CodeModule.CountOfLines
            ![vbcom_datemodified] = Now()
            .Update
            .Bookmark = .LastModified
            lngComs = lngComs + 1&
            lngE = lngComs - 1&
            ReDim Preserve arr_varCom(C_ELEMS, lngE)
            arr_varCom(C_DID, lngE) = lngThisDbsID
            arr_varCom(C_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
            arr_varCom(C_VID, lngE) = ![vbcom_id]
            arr_varCom(C_VNAM, lngE) = strModName
            arr_varCom(C_VTYP, lngE) = intType
            arr_varCom(C_LINS, lngE) = ![vbcom_lines]
            arr_varCom(C_VNAM2, lngE) = strName
            arr_varCom(C_FND, lngE) = CBool(True)
            arr_varCom(C_CHNG, lngE) = CBool(False)
          ElseIf arr_varCom(C_CHNG, lngX) = True Then
            .FindFirst "[vbcom_id] = " & CStr(arr_varCom(C_VID, lngX))
            If .NoMatch = False Then
              If ![comtype_type] <> arr_varCom(C_VTYP, lngX) Then
                .Edit
                ![comtype_type] = arr_varCom(C_VTYP, lngX)
                ![vbcom_datemodified] = Now()
                .Update
              End If
              If ![vbcom_name2] <> arr_varCom(C_VNAM2, lngX) Then
                .Edit
                ![vbcom_name2] = arr_varCom(C_VNAM2, lngX)
                ![vbcom_datemodified] = Now()
                .Update
              End If
              If ![objtype_type] <> arr_varCom(C_OTYP, lngX) Then
                .Edit
                ![objtype_type] = arr_varCom(C_OTYP, lngX)
                ![vbcom_datemodified] = Now()
                .Update
              End If
              If ![vbcom_lines] <> arr_varCom(C_LINS, lngX) Then
                .Edit
                ![vbcom_lines] = arr_varCom(C_LINS, lngX)
                ![vbcom_datemodified] = Now()
                .Update
              End If
              If IsNull(![vbcom_datemodified]) = True Then
                .Edit
                ![vbcom_datemodified] = Now()
                .Update
              End If
              arr_varCom(C_CHNG, lngX) = False
            Else
              Stop
            End If
          End If
        End With
      End With
    Next
  End With

  rst.Close

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  ' ** Only deletes from arr_varCom(), which contains only this database.
  For lngX = 0& To (lngComs - 1&)
    If arr_varCom(C_FND, lngX) = False And arr_varCom(C_DID, lngX) = lngThisDbsID Then
If InStr(arr_varCom(C_VNAM, lngX), "zz_") = 0 Then
      blnDelete = True
      Debug.Print "'DEL? " & Left$(arr_varCom(C_VID, lngX) & "   ", 3) & _
        " " & arr_varCom(C_VNAM, lngX)
Stop
      If blnDelete = True Then
        ' ** Delete tblVBComponent, by specified [compid].
        Set qdf = dbs.QueryDefs("zz_qry_VBComponent_01")
        With qdf.Parameters
          ![compid] = arr_varCom(C_VID, lngX)
        End With
        qdf.Execute dbFailOnError
      End If
Else
  Debug.Print "'NOT DELETED: " & Left$(arr_varCom(C_VID, lngX) & "   ", 3) & _
    " " & arr_varCom(C_VNAM, lngX)
End If
    End If
  Next

  dbs.Close

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_Component_Doc = blnRetValx

End Function

Public Function VBA_Proc_Doc_New() As Boolean
' ** Document all procedures (subs), functions, and properties
' ** to tblVBComponent_Procedure, and their parameters
' ** to tblVBComponent_Procedure_Parameter.
' ** Called by:
' **   QuikVBADoc(), Above

  Const THIS_PROC As String = "VBA_Proc_Doc_New"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstProc As DAO.Recordset, rstParam As DAO.Recordset, rst1 As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngModLines As Long, lngModDecLines As Long, lngBodyLine As Long
  Dim strModName As String, strProcName As String, strLine As String
  Dim lngProcs As Long, arr_varProc() As Variant
  Dim lngParms As Long, arr_varParm As Variant
  Dim strScopeType As String, strProcType As String, strThisProcSubType As String, strLastProcSubType As String
  Dim strReturnType As String, strFirstCodeLine As String, strLastCodeLine As String
  Dim strDeclareLine As String
  Dim lngThisDbsID As Long, lngVBComID As Long, lngDataType As Long, lngEventID As Long
  Dim lngAdds1 As Long, arr_varAdd1() As Variant, lngEdits1 As Long, lngAdds2 As Long, lngEdits2 As Long
  Dim lngRecs As Long, lngProcDels As Long, lngParmDels As Long
  Dim lngDels As Long, arr_varDel() As Variant
  Dim blnIsNew As Boolean, blnIsMulti As Boolean, blnFound As Boolean
  Dim blnAddAll As Boolean, blnAdd As Boolean, blnEdit As Boolean, blnDelete As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long, strTmp05 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varProc().
  Const P_ELEMS As Integer = 19  ' ** Array's first-element UBound().
  Const P_DID  As Integer = 0
  Const P_VID  As Integer = 1
  Const P_VNAM As Integer = 2
  Const P_PID  As Integer = 3
  Const P_PNAM As Integer = 4
  Const P_LBEG As Integer = 5
  Const P_LEND As Integer = 6
  Const P_CBEG As Integer = 7
  Const P_CEND As Integer = 8
  Const P_SCOP As Integer = 9
  Const P_PTYP As Integer = 10
  Const P_RTYP As Integer = 11
  Const P_DTYP As Integer = 12
  Const P_STYP As Integer = 13
  Const P_EVNT As Integer = 14
  Const P_MULT As Integer = 15
  Const P_DECL As Integer = 16
  Const P_MCNT As Integer = 17
  Const P_MARR As Integer = 18
  Const P_FND  As Integer = 19

  ' ** Array: arr_varParm().
  Const M_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const M_ORD   As Integer = 0
  Const M_MID   As Integer = 1
  Const M_MNAM  As Integer = 2
  Const M_MTYP  As Integer = 3
  Const M_DTYP  As Integer = 4
  Const M_SRC   As Integer = 5
  Const M_OPT   As Integer = 6
  Const M_PARR  As Integer = 7
  Const M_BARR  As Integer = 8
  Const M_DEF   As Integer = 9
  Const M_NOTYP As Integer = 10

  ' ** Array: arr_varAdd1().
  Const A_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const A_VNAM As Integer = 0
  Const A_PNAM As Integer = 1
  Const A_PTYP As Integer = 2
  Const A_STYP As Integer = 3

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  Set rstProc = dbs.OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbConsistent)
  Set rstParam = dbs.OpenRecordset("tblVBComponent_Procedure_Parameter", dbOpenDynaset, dbConsistent)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp

    strModName = vbNullString
    lngVBComID = 0&

    lngAdds1 = 0&: lngAdds2 = 0&
    ReDim arr_varAdd1(A_ELEMS, 0)
    lngEdits1 = 0&: lngEdits2 = 0&
    lngProcDels = 0&: lngParmDels = 0&

    For Each vbc In .VBComponents
      With vbc

        strModName = .Name
        strProcName = vbNullString

        varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
          "[vbcom_name] = '" & strModName & "'")
        If IsNull(varTmp00) = False Then
          lngVBComID = varTmp00
        Else
          Debug.Print "'NOT FOUND!  " & strModName
          Stop
        End If

        lngProcs = 0&
        ReDim arr_varProc(P_ELEMS, 0)

        Set cod = .CodeModule
        With cod

          lngModLines = .CountOfLines
          lngModDecLines = .CountOfDeclarationLines

          strFirstCodeLine = vbNullString: strLastCodeLine = vbNullString

          For lngX = 1& To lngModLines

            strLine = Trim(.Lines(lngX, 1))

            strTmp01 = .ProcOfLine(lngX, vbext_pk_Proc)
            If lngX <= lngModDecLines Then
              strTmp01 = "Declaration"
            End If

            strScopeType = vbNullString: strProcType = vbNullString: strThisProcSubType = vbNullString
            strDeclareLine = vbNullString: strReturnType = vbNullString

            ' ** Quite ridiculous! Quite ridiculous!
            If strTmp01 <> "Declaration" Then
              lngBodyLine = 0&: lngTmp02 = 0&: lngTmp03 = 0&: lngTmp04 = 0&
On Error Resume Next
              lngTmp02 = .ProcBodyLine(strTmp01, vbext_pk_Proc)
              If ERR.Number <> 0 Then
                ' ** Error: 35  ' ** Sub of Function not defined.
On Error GoTo 0
On Error Resume Next
                lngTmp02 = .ProcBodyLine(strTmp01, vbext_pk_Get)
                If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
                  ' ** No Get, how about Let and Set.
                  lngTmp03 = .ProcBodyLine(strTmp01, vbext_pk_Let)
                  If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
                    ' ** If it gets here, that means there isn't a Get or Let.
                    lngTmp04 = .ProcBodyLine(strTmp01, vbext_pk_Set)
                    If ERR.Number <> 0 Then
On Error GoTo 0
                      ' ** None of the above!
                      Stop
                    Else
On Error GoTo 0
                      ' ** Just a Set.
                      lngBodyLine = lngTmp04
                      strThisProcSubType = "Set"
                    End If
                  Else
On Error GoTo 0
On Error Resume Next
                    ' ** No Get, yes a Let, how about a Set.
                    lngTmp04 = .ProcBodyLine(strTmp01, vbext_pk_Set)
                    If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
                      ' ** No Get, no Set, just a Let.
                      lngBodyLine = lngTmp03
                      strThisProcSubType = "Let"
                    Else
On Error GoTo 0
On Error Resume Next
                      ' ** Yes, a Let and a Set.
                      If lngX = lngTmp03 Then
                        lngBodyLine = lngTmp03
                        strThisProcSubType = "Let"
                      ElseIf lngX = lngTmp04 Then
                        lngBodyLine = lngTmp04
                        strThisProcSubType = "Set"
                      ElseIf (lngX = (lngTmp03 - 1&)) And (strLine = vbNullString) Then
                        lngBodyLine = lngTmp03
                        strThisProcSubType = "Let"
                      ElseIf (lngX = (lngTmp04 - 1&)) And (strLine = vbNullString) Then
                        lngBodyLine = lngTmp04
                        strThisProcSubType = "Set"
                      Else
                        If lngTmp03 < lngTmp04 Then
                          If lngX > lngTmp03 And lngX < lngTmp04 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          ElseIf lngX > lngTmp04 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          End If
                        ElseIf lngTmp04 < lngTmp03 Then
                          If lngX > lngTmp04 And lngX < lngTmp03 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          ElseIf lngX > lngTmp03 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          End If
                        End If
                      End If
                    End If
                  End If
                Else
On Error GoTo 0
On Error Resume Next
                  ' ** There is a Get, see if there's a Let or a Set.
                  lngTmp03 = .ProcBodyLine(strTmp01, vbext_pk_Let)
                  If ERR.Number <> 0 Then
On Error GoTo 0
On Error Resume Next
                    ' ** No Let, but how about Set.
                    lngTmp04 = .ProcBodyLine(strTmp01, vbext_pk_Set)
                    If ERR.Number <> 0 Then
On Error GoTo 0
                      ' ** Just a Get.
                      lngBodyLine = lngTmp02
                      strThisProcSubType = "Get"
                    Else
On Error GoTo 0
                      ' ** Both a Get and a Set.
                      If lngX = lngTmp02 Then
                        lngBodyLine = lngTmp02
                        strThisProcSubType = "Get"
                      ElseIf lngX = lngTmp04 Then
                        lngBodyLine = lngTmp04
                        strThisProcSubType = "Set"
                      ElseIf (lngX = (lngTmp02 - 1&)) And (strLine = vbNullString) Then
                        lngBodyLine = lngTmp02
                        strThisProcSubType = "Get"
                      ElseIf (lngX = (lngTmp04 - 1&)) And (strLine = vbNullString) Then
                        lngBodyLine = lngTmp04
                        strThisProcSubType = "Set"
                      Else
                        If lngTmp02 < lngTmp04 Then
                          If lngX > lngTmp02 And lngX < lngTmp04 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          ElseIf lngX > lngTmp04 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          End If
                        ElseIf lngTmp04 < lngTmp02 Then
                          If lngX > lngTmp04 And lngX < lngTmp02 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          ElseIf lngX > lngTmp02 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          End If
                        End If
                      End If
                    End If
                  Else
On Error GoTo 0
On Error Resume Next
                    ' ** Now see if there's a Set.
                    lngTmp04 = .ProcBodyLine(strTmp01, vbext_pk_Set)
                    If ERR.Number <> 0 Then
On Error GoTo 0
                      ' ** No Set, but both Get and Let.
                      If lngX = lngTmp02 Then
                        lngBodyLine = lngTmp02
                        strThisProcSubType = "Get"
                      ElseIf lngX = lngTmp03 Then
                        lngBodyLine = lngTmp03
                        strThisProcSubType = "Let"
                      ElseIf lngX = (lngTmp02 - 1&) And strLine = vbNullString Then
                        lngBodyLine = lngTmp02
                        strThisProcSubType = "Get"
                      ElseIf lngX = (lngTmp03 - 1&) And strLine = vbNullString Then
                        lngBodyLine = lngTmp03
                        strThisProcSubType = "Let"
                      Else
                        If lngTmp02 < lngTmp03 Then
                          If lngX > lngTmp02 And lngX < lngTmp03 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          ElseIf lngX > lngTmp03 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          End If
                        ElseIf lngTmp03 < lngTmp02 Then
                          If lngX > lngTmp03 And lngX < lngTmp02 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          ElseIf lngX > lngTmp02 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          End If
                        End If
                      End If
                    Else
On Error GoTo 0
                      ' ** All 3 are present! (Unlikely)
                      If lngX = lngTmp02 Then
                        lngBodyLine = lngTmp02
                        strThisProcSubType = "Get"
                      ElseIf lngX = lngTmp03 Then
                        lngBodyLine = lngTmp03
                        strThisProcSubType = "Let"
                      ElseIf lngX = lngTmp04 Then
                        lngBodyLine = lngTmp04
                        strThisProcSubType = "Set"
                      ElseIf (lngX = (lngTmp02 - 1&)) And (strLine = vbNullString) Then
                        lngBodyLine = lngTmp02
                        strThisProcSubType = "Get"
                      ElseIf (lngX = (lngTmp03 - 1&)) And (strLine = vbNullString) Then
                        lngBodyLine = lngTmp03
                        strThisProcSubType = "Let"
                      ElseIf (lngX = (lngTmp04 - 1&)) And (strLine = vbNullString) Then
                        lngBodyLine = lngTmp04
                        strThisProcSubType = "Set"
                      Else
                        If lngTmp02 < lngTmp03 And lngTmp03 < lngTmp04 Then
                          '2 3 4
                          If lngX > lngTmp02 And lngX < lngTmp03 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          ElseIf lngX > lngTmp03 And lngX < lngTmp04 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          ElseIf lngX > lngTmp04 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          End If
                        ElseIf lngTmp02 < lngTmp04 And lngTmp04 < lngTmp03 Then
                          '2 4 3
                          If lngX > lngTmp02 And lngX < lngTmp04 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          ElseIf lngX > lngTmp04 And lngX < lngTmp03 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          ElseIf lngX > lngTmp03 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          End If
                        ElseIf lngTmp03 < lngTmp02 And lngTmp02 < lngTmp04 Then
                          '3 2 4
                          If lngX > lngTmp03 And lngX < lngTmp02 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          ElseIf lngX > lngTmp02 And lngX < lngTmp04 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          ElseIf lngX > lngTmp04 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          End If
                        ElseIf lngTmp03 < lngTmp04 And lngTmp04 < lngTmp02 Then
                          '3 4 2
                          If lngX > lngTmp03 And lngX < lngTmp04 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          ElseIf lngX > lngTmp04 And lngX < lngTmp02 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          ElseIf lngX > lngTmp02 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          End If
                        ElseIf lngTmp04 < lngTmp02 And lngTmp02 < lngTmp03 Then
                          '4 2 3
                          If lngX > lngTmp04 And lngX < lngTmp02 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          ElseIf lngX > lngTmp02 And lngX < lngTmp03 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          ElseIf lngX > lngTmp03 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          End If
                        ElseIf lngTmp04 < lngTmp03 And lngTmp03 < lngTmp02 Then
                          '4 3 2
                          If lngX > lngTmp04 And lngX < lngTmp03 Then
                            lngBodyLine = lngTmp04
                            strThisProcSubType = "Set"
                          ElseIf lngX > lngTmp03 And lngX < lngTmp02 Then
                            lngBodyLine = lngTmp03
                            strThisProcSubType = "Let"
                          ElseIf lngX > lngTmp02 Then
                            lngBodyLine = lngTmp02
                            strThisProcSubType = "Get"
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
'## CHECK!
'If strThisProcSubType = vbNullString Then
'  Stop
'End If
'##
              Else
On Error GoTo 0
                ' ** Sub or Function.
                lngBodyLine = lngTmp02
              End If
            Else
              ' ** Declaration.
              lngBodyLine = 0&
              strScopeType = "{unscoped}"
              strProcType = strTmp01
            End If  ' ** strTmp01.

'## CHECK!
'If lngBodyLine = 0& And strProcType <> "Declaration" Then
'  Stop
'End If
'##
            ' ** Make sure the declare line is complete.
            If strScopeType = vbNullString Then
              strDeclareLine = .Lines(lngBodyLine, 1)  ' ** This proc's declare line.
              blnIsMulti = False
              If Right$(strDeclareLine, 1) = "_" Then
                blnIsMulti = True
                strTmp05 = Trim(.Lines(lngBodyLine + 1&, 1))
                If Right$(strTmp05, 1) = "_" Then
                  strDeclareLine = Left$(strDeclareLine, (Len(strDeclareLine) - 1)) & " " & strTmp05
                  strTmp05 = Trim(.Lines(lngBodyLine + 2&, 1))
                  If Right$(strTmp05, 1) = "_" Then  ' ** I don't think any go over 3 lines.
                    Stop
                  Else
                    strDeclareLine = Left$(strDeclareLine, (Len(strDeclareLine) - 1)) & " " & strTmp05
                  End If
                Else
                  strDeclareLine = Left$(strDeclareLine, (Len(strDeclareLine) - 1)) & " " & strTmp05
                End If
              End If
              intPos1 = InStr(strDeclareLine, ")")
              If intPos1 > 0 Then
                intPos2 = InStr((intPos1 + 1), strDeclareLine, ")")  ' ** In case there's an array or something.
                If intPos2 = 0 Then
                  If intPos1 = Len(strDeclareLine) Then
                    ' ** All's well.
                  Else
                    strTmp05 = Trim$(Mid$(strDeclareLine, (intPos1 + 1)))
                    If Left$(strTmp05, 3) = "As " Then
                      intPos2 = InStr((intPos1 + 1), strDeclareLine, "'")
                      If intPos2 > 0 Then  ' ** Remark.
                        strDeclareLine = Trim$(Left$(strDeclareLine, (intPos2 - 1)))
                      Else
                        ' ** Leave it alone.
                      End If
                    Else
                      Debug.Print "'WHAT'S THERE?  " & strDeclareLine
                      Stop
                    End If
                  End If
                Else
                  ' ** A bit more complicated.
                  If intPos2 = Len(strDeclareLine) Then
                    ' ** Leave it alone.
                  Else
                    intPos3 = InStr((intPos2 + 1), strDeclareLine, ")")
                    If intPos3 > 0 Then
                      Debug.Print "'3RD PAREN!  " & strDeclareLine
                      Stop
                    Else
                      strTmp05 = Trim$(Mid$(strDeclareLine, (intPos2 + 1)))
                      If Left$(strTmp05, 3) = "As " Then
                        intPos3 = InStr((intPos2 + 1), strDeclareLine, "'")
                        If intPos3 > 0 Then  ' ** Remark.
                          strDeclareLine = Trim$(Left$(strDeclareLine, (intPos3 - 1)))
                        Else
                          ' ** Leave it alone.
                        End If
                      Else
                        Debug.Print "'WHAT'S THERE?  " & strDeclareLine
                        Stop
                      End If
                    End If
                  End If
                End If
              Else
                Stop
              End If
            End If

'## CHECK!
'If strTmp01 = vbNullString Then
'  Stop
'End If
'##
            blnIsNew = False
            ' ** strTmp01 is this proc name, strProcName is the last proc name.
            If strTmp01 <> strProcName Then
              blnIsNew = True
              strProcName = strTmp01
              strLastProcSubType = vbNullString
'## CHECK!
'If strFirstCodeLine <> vbNullString Or strLastCodeLine <> vbNullString Then
'  Stop
'End If
'##
            Else
              ' ** Proc names match, now check the subtypes.
              If strThisProcSubType <> vbNullString Then
                If strThisProcSubType <> strLastProcSubType Then
                  blnIsNew = True
                Else
                  ' ** Same property.
                End If
              Else
                ' ** Same procedure.
              End If
            End If
            strTmp01 = vbNullString
            lngTmp02 = 0&: lngTmp03 = 0&: lngTmp04 = 0&

            If blnIsNew = True Then

              If lngProcs > 0& Then
                arr_varProc(P_LEND, (lngProcs - 1&)) = (lngX - 1&)
                If strFirstCodeLine <> vbNullString Then
                  arr_varProc(P_CBEG, (lngProcs - 1&)) = CLng(strFirstCodeLine)
                Else
                  arr_varProc(P_CBEG, (lngProcs - 1&)) = Null
                End If
                If strLastCodeLine <> vbNullString Then
                  arr_varProc(P_CEND, (lngProcs - 1&)) = CLng(strLastCodeLine)
                Else
                  arr_varProc(P_CEND, (lngProcs - 1&)) = Null
                End If
'## CHECK!
'If (strFirstCodeLine = vbNullString Or strLastCodeLine = vbNullString) And _
'    arr_varProc(P_PNAM, (lngProcs - 1&)) <> "Declaration" And _
'    Left(arr_varProc(P_VNAM, (lngProcs - 1&)), 2) <> "zz" Then
'  Stop
'End If
'##
              End If  ' ** lngProcs.
              strFirstCodeLine = vbNullString: strLastCodeLine = vbNullString

              ' ** The Declaration section is already scoped.
              If strScopeType = vbNullString Then
                intPos1 = InStr(strDeclareLine, " ")  ' ** First space.
                If intPos1 > 0 Then
                  strScopeType = Trim$(Left$(strDeclareLine, intPos1))
                  If strScopeType <> "Public" And strScopeType <> "Private" Then
                    Debug.Print "'NO SCOPE!  " & strModName & "  " & strProcName
                    DoEvents
                    'Stop
                  End If
                Else
                  Stop
                End If
              End If

              ' ** The Declaration section is already typed.
              If strProcType = vbNullString Then
                intPos1 = InStr(strDeclareLine, " ")  ' ** First space.
                If intPos1 > 0 Then
                  strProcType = Trim$(Mid$(strDeclareLine, intPos1))
                  intPos1 = InStr(strProcType, " ")  ' ** Second space.
                  strProcType = Trim(Left$(strProcType, intPos1))
                  If strProcType <> "Function" And strProcType <> "Sub" And strProcType <> "Property" Then
                    Debug.Print "'SCOPE OR TYPE OFF!  " & strModName & "  " & strProcName
                    DoEvents
                    'Stop
                  End If
                Else
                  Stop
                End If
              End If

              If strProcType <> "Declaration" Then

                lngDataType = 0&
                intPos1 = InStr(strDeclareLine, ")")
                If intPos1 > 0 Then
                  If intPos1 = Len(strDeclareLine) Then
                    ' ** No return type.
                  Else
                    strTmp05 = Trim$(Mid$(strDeclareLine, (intPos1 + 1)))
                    If Left$(strTmp05, 3) = "As " Then
                      strReturnType = Trim$(Mid$(strTmp05, 3))
                      varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", _
                        "[datatype_vb_constant] = 'vb" & strReturnType & "'")
                      If IsNull(varTmp00) = False Then
                        lngDataType = CLng(varTmp00)
                      Else
                        If IsUC(strReturnType, True, True) = True Then  ' ** Module Function: modStringFuncs.
                          varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", _
                            "[datatype_vb_constant] = 'vbUserDefinedType'")
                          lngDataType = CLng(varTmp00)
                        Else
                          varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", _
                            "[datatype_vb_constant] = 'vbVariant'")
                          lngDataType = CLng(varTmp00)
                        End If
                      End If
                    Else
                      Debug.Print "'WHAT IS THIS?  " & strDeclareLine
                      Stop
                    End If
                  End If
                Else
                  Stop
                End If
             
                lngEventID = 0&
                intPos1 = InStr(strProcName, "_")
                If intPos1 > 0 Then
                  strTmp05 = Mid$(strProcName, intPos1)
                  varTmp00 = DLookup("[vbcom_event_id]", "tblVBComponent_Event", "[vbcom_event_ext] = '" & strTmp05 & "'")
                  If IsNull(varTmp00) = False Then
                    lngEventID = CLng(varTmp00)
                  End If
                End If

'## CHECK!
'If strProcType = "Property" And strThisProcSubType = vbNullString Then
'  Stop
'End If
'##
              End If  ' ** Declaration.

              lngProcs = lngProcs + 1&
              lngE = lngProcs - 1&
              ReDim Preserve arr_varProc(P_ELEMS, lngE)
              ' *********************************************************
              ' ** Array: arr_varProc()
              ' **
              ' **   FIELD  ELEMENT  NAME                    CONSTANT
              ' **   =====  =======  ======================  ==========
              ' **     1       0     dbs_id                  P_DID
              ' **     2       1     vbcom_id                P_VID
              ' **     3       2     vbcom_name              P_VNAM
              ' **     4       3     vbcomproc_id            P_PID
              ' **     5       4     vbcomproc_name          P_PNAM
              ' **     6       5     vbcomproc_line_beg      P_LBEG
              ' **     7       6     vbcomproc_line_end      P_LEND
              ' **     8       7     vbcomproc_code_beg      P_CBEG
              ' **     9       8     vbcomproc_code_end      P_CEND
              ' **    10       9     scopetype_type          P_SCOP
              ' **    11      10     proctype_type           P_PTYP
              ' **    12      11     vbcomproc_returntype    P_RTYP
              ' **    13      12     datatype_vb_type        P_DTYP
              ' **    14      13     procsubtype_type        P_STYP
              ' **    15      14     vbcom_event_id          P_EVNT
              ' **    16      15     vbcomproc_multiline     P_MULT
              ' **    17      16     Declare Line            P_DECL
              ' **    18      17     vbcomproc_params        P_MCNT
              ' **    19      18     arr_varParm()           P_MARR
              ' **    20      19     Found                   P_FND
              ' **
              ' *********************************************************
              arr_varProc(P_DID, lngE) = lngThisDbsID
              arr_varProc(P_VID, lngE) = lngVBComID
              arr_varProc(P_VNAM, lngE) = strModName
              arr_varProc(P_PID, lngE) = CLng(0)
              arr_varProc(P_PNAM, lngE) = strProcName
              If strProcName = "Declaration" Then
                arr_varProc(P_LBEG, lngE) = CLng(1)
              Else
                arr_varProc(P_LBEG, lngE) = lngBodyLine
              End If
              arr_varProc(P_LEND, lngE) = CLng(0)
              arr_varProc(P_CBEG, lngE) = Null
              arr_varProc(P_CEND, lngE) = Null
              arr_varProc(P_SCOP, lngE) = strScopeType
              arr_varProc(P_PTYP, lngE) = strProcType
              If strReturnType <> vbNullString Then
                arr_varProc(P_RTYP, lngE) = strReturnType
                arr_varProc(P_DTYP, lngE) = lngDataType
              Else
                arr_varProc(P_RTYP, lngE) = Null
                arr_varProc(P_DTYP, lngE) = Null
              End If
              If strThisProcSubType <> vbNullString Then
                arr_varProc(P_STYP, lngE) = strThisProcSubType
              Else
                arr_varProc(P_STYP, lngE) = Null
              End If
              If lngEventID > 0& Then
                arr_varProc(P_EVNT, lngE) = lngEventID
              Else
                arr_varProc(P_EVNT, lngE) = Null
              End If
              arr_varProc(P_MULT, lngE) = blnIsMulti
              If strProcType <> "Declaration" Then
                arr_varProc(P_DECL, lngE) = strDeclareLine
              Else
                arr_varProc(P_DECL, lngE) = Null
              End If
              arr_varProc(P_MCNT, lngE) = CLng(0)
              arr_varProc(P_MARR, lngE) = Empty
              arr_varProc(P_FND, lngE) = CBool(False)
              strLastProcSubType = strThisProcSubType

            Else

              If strProcType <> "Declaration" Then
                intPos1 = InStr(strLine, " ")
                If intPos1 > 0 Then
                  strTmp05 = Trim$(Left$(strLine, intPos1))
                  If IsNumeric(strTmp05) = True Then
                    If strFirstCodeLine = vbNullString Then
                      strFirstCodeLine = strTmp05
                    End If
                    strLastCodeLine = strTmp05
                  End If
                End If  ' ** intPos1
              End If  ' ** Declaration.

            End If  ' ** blnIsNew.

            If lngX = lngModLines Then
              arr_varProc(P_LEND, (lngProcs - 1&)) = lngX
              If strFirstCodeLine <> vbNullString Then
                arr_varProc(P_CBEG, (lngProcs - 1&)) = CLng(strFirstCodeLine)
              Else
                arr_varProc(P_CBEG, (lngProcs - 1&)) = Null
              End If
              If strLastCodeLine <> vbNullString Then
                arr_varProc(P_CEND, (lngProcs - 1&)) = CLng(strLastCodeLine)
              Else
                arr_varProc(P_CEND, (lngProcs - 1&)) = Null
              End If
            End If

          Next  ' ** lngX.

        End With  ' ** cod.
        Set cod = Nothing
      End With  ' ** vbc.
      Set vbc = Nothing

'## CHECK!
'If Left$(strModName, 2) <> "zz" Then
'  lngTmp02 = 0&: lngTmp03 = 0&
'  For lngX = 0& To (lngProcs - 1&)
'    If IsNull(arr_varProc(P_CBEG, lngX)) = True Then
'      If arr_varProc(P_PTYP, lngX) <> "Declaration" Then
'        lngTmp02 = lngTmp02 + 1&
'      End If
'    End If
'    If IsNull(arr_varProc(P_CEND, lngX)) = True Then
'      If arr_varProc(P_PTYP, lngX) <> "Declaration" Then
'        lngTmp03 = lngTmp03 + 1&
'      End If
'    End If
'  Next
'  If lngTmp02 > 0& Then
'    Debug.Print "'CODE NUM MISSING!  " & CStr(lngTmp02) & "  " & CStr(lngTmp03)
'    DoEvents
'    Stop
'  End If
'End If
'##

      If lngProcs > 0& Then

        blnAddAll = False
        With rstProc
          If .BOF = True And .EOF = True Then
            blnAddAll = True
          Else
            .MoveLast
            .MoveFirst
          End If
          For lngX = 0& To (lngProcs - 1&)
            blnAdd = False
            If blnAddAll = False Then
              If arr_varProc(P_PTYP, lngX) <> "Property" Then
                .FindFirst "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varProc(P_VID, lngX)) & " And " & _
                  "[vbcomproc_name] = '" & arr_varProc(P_PNAM, lngX) & "'"
              Else
                .FindFirst "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varProc(P_VID, lngX)) & " And " & _
                  "[vbcomproc_name] = '" & arr_varProc(P_PNAM, lngX) & "' And [procsubtype_type] = '" & arr_varProc(P_STYP, lngX) & "'"
              End If
              Select Case .NoMatch
              Case True
                blnAdd = True
              Case False
                ' ** Edit.
                arr_varProc(P_PID, lngX) = ![vbcomproc_id]
              End Select
            Else
              blnAdd = True
            End If
            If blnAdd = False Then
              blnEdit = False
              If arr_varProc(P_LBEG, lngX) > 0 Then
                If IsNull(![vbcomproc_line_beg]) = True Then
                  .Edit
                  ![vbcomproc_line_beg] = arr_varProc(P_LBEG, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![vbcomproc_line_beg] <> arr_varProc(P_LBEG, lngX) Then
                    .Edit
                    ![vbcomproc_line_beg] = arr_varProc(P_LBEG, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
              Else
                If IsNull(![vbcomproc_line_beg]) = False Then
                  .Edit
                  ![vbcomproc_line_beg] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If arr_varProc(P_LBEG, lngX) > 0 Then
                If IsNull(![vbcomproc_line_end]) = True Then
                  .Edit
                  ![vbcomproc_line_end] = arr_varProc(P_LEND, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![vbcomproc_line_end] <> arr_varProc(P_LEND, lngX) Then
                    .Edit
                    ![vbcomproc_line_end] = arr_varProc(P_LEND, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
              Else
                If IsNull(![vbcomproc_line_end]) = False Then
                  .Edit
                  ![vbcomproc_line_end] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If IsNull(arr_varProc(P_CBEG, lngX)) = False Then
                If IsNull(![vbcomproc_code_beg]) = True Then
                  .Edit
                  ![vbcomproc_code_beg] = arr_varProc(P_CBEG, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![vbcomproc_code_beg] <> arr_varProc(P_CBEG, lngX) Then
                    .Edit
                    ![vbcomproc_code_beg] = arr_varProc(P_CBEG, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
              Else
                If IsNull(![vbcomproc_code_beg]) = False Then
                  .Edit
                  ![vbcomproc_code_beg] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If IsNull(arr_varProc(P_CEND, lngX)) = False Then
                If IsNull(![vbcomproc_code_end]) = True Then
                  .Edit
                  ![vbcomproc_code_end] = arr_varProc(P_CEND, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![vbcomproc_code_end] <> arr_varProc(P_CEND, lngX) Then
                    .Edit
                    ![vbcomproc_code_end] = arr_varProc(P_CEND, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
              Else
                If IsNull(![vbcomproc_code_end]) = False Then
                  .Edit
                  ![vbcomproc_code_end] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If IsNull(![scopetype_type]) = True Then
                .Edit
                ![scopetype_type] = arr_varProc(P_SCOP, lngX)
                ![vbcomproc_datemodified] = Now()
                .Update
                blnEdit = True
              Else
                If ![scopetype_type] <> arr_varProc(P_SCOP, lngX) Then
                  .Edit
                  ![scopetype_type] = arr_varProc(P_SCOP, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If IsNull(![proctype_type]) = True Then
                .Edit
                ![proctype_type] = arr_varProc(P_PTYP, lngX)
                ![vbcomproc_datemodified] = Now()
                .Update
                blnEdit = True
              Else
                If ![proctype_type] <> arr_varProc(P_PTYP, lngX) Then
                  .Edit
                  ![proctype_type] = arr_varProc(P_PTYP, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If IsNull(arr_varProc(P_RTYP, lngX)) = False Then
                If IsNull(![vbcomproc_returntype]) = True Then
                  .Edit
                  ![vbcomproc_returntype] = arr_varProc(P_RTYP, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![vbcomproc_returntype] <> arr_varProc(P_RTYP, lngX) Then
                    .Edit
                    ![vbcomproc_returntype] = arr_varProc(P_RTYP, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
                If IsNull(![datatype_vb_type]) = True Then
                  .Edit
                  ![datatype_vb_type] = arr_varProc(P_DTYP, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![datatype_vb_type] <> arr_varProc(P_DTYP, lngX) Then
                    .Edit
                    ![datatype_vb_type] = arr_varProc(P_DTYP, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
              Else
                If IsNull(![vbcomproc_returntype]) = False Then
                  .Edit
                  ![vbcomproc_returntype] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
                If IsNull(![datatype_vb_type]) = False Then
                  .Edit
                  ![datatype_vb_type] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If IsNull(arr_varProc(P_STYP, lngX)) = False Then
                If IsNull(![procsubtype_type]) = True Then
                  .Edit
                  ![procsubtype_type] = arr_varProc(P_STYP, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![procsubtype_type] <> arr_varProc(P_STYP, lngX) Then
                    .Edit
                    ![procsubtype_type] = arr_varProc(P_STYP, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
              Else
                If IsNull(![procsubtype_type]) = False Then
                  .Edit
                  ![procsubtype_type] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If arr_varProc(P_EVNT, lngX) > 0 Then
                If IsNull(![vbcom_event_id]) = True Then
                  .Edit
                  ![vbcom_event_id] = arr_varProc(P_EVNT, lngX)
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                Else
                  If ![vbcom_event_id] <> arr_varProc(P_EVNT, lngX) Then
                    .Edit
                    ![vbcom_event_id] = arr_varProc(P_EVNT, lngX)
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdit = True
                  End If
                End If
              Else
                If IsNull(![vbcom_event_id]) = False Then
                  .Edit
                  ![vbcom_event_id] = Null
                  ![vbcomproc_datemodified] = Now()
                  .Update
                  blnEdit = True
                End If
              End If
              If ![vbcomproc_multiline] <> arr_varProc(P_MULT, lngX) Then
                .Edit
                ![vbcomproc_multiline] = arr_varProc(P_MULT, lngX)
                ![vbcomproc_datemodified] = Now()
                .Update
                blnEdit = True
              End If
              If blnEdit = True Then
                lngEdits1 = lngEdits1 + 1&
              End If
            Else
              .AddNew
              ![dbs_id] = arr_varProc(P_DID, lngX)
              ![vbcom_id] = arr_varProc(P_VID, lngX)
              ' ** ![vbcomproc_id] : AutoNumber.
              ![vbcomproc_name] = arr_varProc(P_PNAM, lngX)
              If arr_varProc(P_LBEG, lngX) > 0 Then
                ![vbcomproc_line_beg] = arr_varProc(P_LBEG, lngX)
              End If
              If arr_varProc(P_LEND, lngX) > 0 Then
                ![vbcomproc_line_end] = arr_varProc(P_LEND, lngX)
              End If
              If IsNull(arr_varProc(P_CBEG, lngX)) = False Then
                ![vbcomproc_code_beg] = arr_varProc(P_CBEG, lngX)
              End If
              If IsNull(arr_varProc(P_CEND, lngX)) = False Then
                ![vbcomproc_code_end] = arr_varProc(P_CEND, lngX)
              End If
              ![scopetype_type] = arr_varProc(P_SCOP, lngX)
              ![proctype_type] = arr_varProc(P_PTYP, lngX)
              If IsNull(arr_varProc(P_RTYP, lngX)) = False Then
                ![vbcomproc_returntype] = arr_varProc(P_RTYP, lngX)
              End If
              If IsNull(arr_varProc(P_DTYP, lngX)) = False Then
                ![datatype_vb_type] = arr_varProc(P_DTYP, lngX)
              End If
              If IsNull(arr_varProc(P_STYP, lngX)) = False Then
                ![procsubtype_type] = arr_varProc(P_STYP, lngX)
              End If
              If arr_varProc(P_EVNT, lngX) > 0 Then
                ![vbcom_event_id] = arr_varProc(P_EVNT, lngX)
              End If
              ![vbcomproc_multiline] = arr_varProc(P_MULT, lngX)
              ![vbcomproc_datemodified] = Now()
              .Update
              .Bookmark = .LastModified
              arr_varProc(P_PID, lngX) = ![vbcomproc_id]
              lngAdds1 = lngAdds1 + 1&
              lngE = lngAdds1 - 1&
              ReDim Preserve arr_varAdd1(A_ELEMS, lngE)
              arr_varAdd1(A_VNAM, lngE) = strModName
              arr_varAdd1(A_PNAM, lngE) = arr_varProc(P_PNAM, lngX)
              arr_varAdd1(A_PTYP, lngE) = arr_varProc(P_PTYP, lngX)
              arr_varAdd1(A_STYP, lngE) = arr_varProc(P_STYP, lngX)
            End If
            'Debug.Print "'" & arr_varProc(P_PNAM, lngX)
          Next  ' ** lngX.
        End With ' ** rstProc.

        ' ****************
        ' ** Parameters.
        ' ****************
        blnAddAll = False
        With rstParam
          If .BOF = True And .EOF = True Then
            blnAddAll = True
          Else
            .MoveLast
            .MoveFirst
          End If
          For lngX = 0& To (lngProcs - 1&)
            If arr_varProc(P_PID, lngX) = 0& Then
              Stop
            End If
          Next
          For lngX = 0& To (lngProcs - 1&)
            lngParms = 0&
            arr_varParm = Empty
            If IsNull(arr_varProc(P_DECL, lngX)) = False Then
              strTmp01 = arr_varProc(P_DECL, lngX)
              If Right$(strTmp01, 2) <> "()" Then
                arr_varParm = VBA_ProcParamSplit(strTmp01)  ' ** Function: Below.
                If arr_varParm(0, 0) > 0 Then
                  lngParms = UBound(arr_varParm, 2) + 1&
                  For lngY = 0& To (lngParms - 1&)
                    blnAdd = False
                    If blnAddAll = False Then
                      .FindFirst "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And " & _
                        "[vbcom_id] = " & CStr(arr_varProc(P_VID, lngX)) & " And " & _
                        "[vbcomproc_id] = " & CStr(arr_varProc(P_PID, lngX)) & " And " & _
                        "[vbcomparam_name] = '" & arr_varParm(M_MNAM, lngY) & "'"
                      Select Case .NoMatch
                      Case True
                        blnAdd = True
                      Case False
                        ' ** Edit.
                        arr_varParm(M_MID, lngY) = ![vbcomparam_id]
                      End Select
                    Else
                      blnAdd = True
                    End If
                    If blnAdd = False Then
                      blnEdit = False
                      ' ** ![dbs_id] : No edit.
                      ' ** ![vbcom_id] : No edit.
                      ' ** ![vbcomproc_id] : No edit.
                      ' ** ![vbcomparam_id] : AutoNumber.
                      If ![vbcomparam_order] <> arr_varParm(M_ORD, lngY) Then
                        ' ** Careful!
                        varTmp00 = DMax("[vbcomparam_order]", "tblVBComponent_Procedure_Parameter", _
                          "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And " & _
                          "[vbcom_id] = " & CStr(arr_varProc(P_VID, lngX)) & " And " & _
                          "[vbcomproc_id] = " & CStr(arr_varProc(P_PID, lngX)))
                        If IsNull(varTmp00) = False Then
                          lngTmp02 = varTmp00
                          ' ** See if another already has the new order.
                          varTmp00 = DLookup("[vbcomparam_id]", "tblVBComponent_Procedure_Parameter", _
                            "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And " & _
                            "[vbcom_id] = " & CStr(arr_varProc(P_VID, lngX)) & " And " & _
                            "[vbcomproc_id] = " & CStr(arr_varProc(P_PID, lngX)) & " And " & _
                            "[vbcomparam_order] = " & CStr(arr_varParm(M_ORD, lngY)))
                          If IsNull(varTmp00) = True Then
                            ' ** We're in luck! (Or maybe it was moved aside earlier.)
                            .Edit
                            ![vbcomparam_order] = arr_varParm(M_ORD, lngY)
                            ![vbcomparam_datemodified] = Now()
                            .Update
                          Else
                            ' ** A little shuffling...
                            lngTmp03 = varTmp00
                            lngTmp04 = ![vbcomparam_id]
                            ' ** Find the one using the order now.
                            .FindFirst "[vbcomparam_id] = " & CStr(lngTmp03)
                            If .NoMatch = False Then
                              .Edit
                              ![vbcomparam_order] = (lngTmp02 + 1&)  ' ** Move it out of the way.
                              ![vbcomparam_datemodified] = Now()
                              .Update
                              blnEdit = True
                              ' ** Now back to the one we're editing.
                              .FindFirst "[vbcomparam_id] = " & CStr(lngTmp04)
                              If .NoMatch = False Then
                                .Edit
                                ![vbcomparam_order] = arr_varParm(M_ORD, lngY)
                                ![vbcomparam_datemodified] = Now()
                                .Update
                              Else
                                Stop
                              End If
                            Else
                              Stop
                            End If
                          End If
                        Else
                          Stop
                        End If
                      End If
                      If Compare_StringA_StringB(![vbcomparam_name], "=", arr_varParm(M_MNAM, lngY)) = False Then  ' ** Module Function: modStringFuncs.
                        ' ** In case capitalization changed.
                        .Edit
                        ![vbcomparam_name] = arr_varParm(M_MNAM, lngY)
                        ![vbcomparam_datemodified] = Now()
                        .Update
                        blnEdit = True
                      End If
                      If IsNull(arr_varParm(M_MTYP, lngY)) = False Then
                        If IsNull(![vbcomparam_type]) = True Then
                          .Edit
                          ![vbcomparam_type] = arr_varParm(M_MTYP, lngY)
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        Else
                          If ![vbcomparam_type] <> arr_varParm(M_MTYP, lngY) Then
                            .Edit
                            ![vbcomparam_type] = arr_varParm(M_MTYP, lngY)
                            ![vbcomparam_datemodified] = Now()
                            .Update
                            blnEdit = True
                          End If
                        End If
                      Else
                        If IsNull(![vbcomparam_type]) = False Then
                          .Edit
                          ![vbcomparam_type] = Null
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        End If
                      End If
                      If IsNull(arr_varParm(M_DTYP, lngY)) = False Then
                        If IsNull(![datatype_vb_type]) = True Then
                          .Edit
                          ![datatype_vb_type] = arr_varParm(M_DTYP, lngY)
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        Else
                          If ![datatype_vb_type] <> arr_varParm(M_DTYP, lngY) Then
                            .Edit
                            ![datatype_vb_type] = arr_varParm(M_DTYP, lngY)
                            ![vbcomparam_datemodified] = Now()
                            .Update
                            blnEdit = True
                          End If
                        End If
                      Else
                        If IsNull(![datatype_vb_type]) = False Then
                          .Edit
                          ![datatype_vb_type] = Null
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        End If
                      End If
                      intPos1 = InStr(arr_varParm(M_MTYP, lngY), ".")
                      If intPos1 > 0 Then
                        strTmp01 = Left(arr_varParm(M_MTYP, lngY), (intPos1 - 1))
                        strTmp05 = Mid(arr_varParm(M_MTYP, lngY), (intPos1 + 1))
                        Select Case strTmp01
                        Case "Access"
                          ' ** Acces.Form, etc.
                          varTmp00 = DLookup("[objtype_type]", "tblObjectType", "[objtype_constant] = 'vb" & strTmp05 & "'")
                          If IsNull(varTmp00) = False Then
                            If IsNull(![objtype_type]) = True Then
                              .Edit
                              ![objtype_type] = varTmp00
                              ![vbcomparam_datemodified] = Now()
                              .Update
                              blnEdit = True
                            Else
                              If ![objtype_type] <> varTmp00 Then
                                .Edit
                                ![objtype_type] = varTmp00
                                ![vbcomparam_datemodified] = Now()
                                .Update
                                blnEdit = True
                              End If
                            End If
                          Else
                            ' ** Probably a control, or something else we aren't saving explicitly.
                            If IsNull(![objtype_type]) = False Then
                              .Edit
                              ![objtype_type] = Null
                              ![vbcomparam_datemodified] = Now()
                              .Update
                              blnEdit = True
                            End If
                          End If
                        Case "DAO"
                          ' ** DAO.Recordset, etc.
                          varTmp00 = DLookup("[daotype_id]", "tblDAOType", "[daotype_type] = '" & arr_varParm(M_MTYP, lngY) & "'")
                          If IsNull(varTmp00) = False Then
                            If IsNull(![daotype_type]) = True Then
                              .Edit
                              ![daotype_type] = arr_varParm(M_MTYP, lngY)
                              ![vbcomparam_datemodified] = Now()
                              .Update
                              blnEdit = True
                            Else
                              If ![daotype_type] <> arr_varParm(M_MTYP, lngY) Then
                                .Edit
                                ![daotype_type] = arr_varParm(M_MTYP, lngY)
                                ![vbcomparam_datemodified] = Now()
                                .Update
                                blnEdit = True
                              End If
                            End If
                          Else
                            Debug.Print "'NOT IN tblDAOType!  " & arr_varParm(M_MTYP, lngY)
                            DoEvents
                          End If
                        Case "Scripting"
                          ' ** Scripting.Folder, etc.
                          varTmp00 = DLookup("[scripttype_id]", "tblScriptingType", "[scripttype_type] = '" & arr_varParm(M_MTYP, lngY) & "'")
                          If IsNull(varTmp00) = False Then
                            If IsNull(![scripttype_type]) = True Then
                              .Edit
                              ![scripttype_type] = arr_varParm(M_MTYP, lngY)
                              ![vbcomparam_datemodified] = Now()
                              .Update
                              blnEdit = True
                            Else
                              If ![scripttype_type] <> arr_varParm(M_MTYP, lngY) Then
                                .Edit
                                ![scripttype_type] = arr_varParm(M_MTYP, lngY)
                                ![vbcomparam_datemodified] = Now()
                                .Update
                                blnEdit = True
                              End If
                            End If
                          Else
                            Debug.Print "'NOT IN tblScriptingType!  " & arr_varParm(M_MTYP, lngY)
                            DoEvents
                          End If
                        End Select
                      End If
                      If ![vbcomparam_optional] <> arr_varParm(M_OPT, lngY) Then
                        .Edit
                        ![vbcomparam_optional] = arr_varParm(M_OPT, lngY)
                        ![vbcomparam_datemodified] = Now()
                        .Update
                        blnEdit = True
                      End If
                      If IsNull(arr_varParm(M_SRC, lngY)) = False Then
                        If IsNull(![vbcomparam_explicit]) = True Then
                          .Edit
                          ![vbcomparam_explicit] = arr_varParm(M_SRC, lngY)
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        Else
                          If ![vbcomparam_explicit] <> arr_varParm(M_SRC, lngY) Then
                            .Edit
                            ![vbcomparam_explicit] = arr_varParm(M_SRC, lngY)
                            ![vbcomparam_datemodified] = Now()
                            .Update
                            blnEdit = True
                          End If
                        End If
                      Else
                        If IsNull(![vbcomparam_explicit]) = False Then
                          .Edit
                          ![vbcomparam_explicit] = Null
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        End If
                      End If
                      If IsNull(arr_varParm(M_DEF, lngY)) = False Then
                        If IsNull(![vbcomparam_default]) = True Then
                          .Edit
                          ![vbcomparam_default] = arr_varParm(M_DEF, lngY)
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        Else
                          If ![vbcomparam_default] <> arr_varParm(M_DEF, lngY) Then
                          .Edit
                            ![vbcomparam_default] = arr_varParm(M_DEF, lngY)
                            ![vbcomparam_datemodified] = Now()
                            .Update
                            blnEdit = True
                          End If
                        End If
                      Else
                        If IsNull(![vbcomparam_default]) = False Then
                          .Edit
                          ![vbcomparam_default] = Null
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        End If
                      End If
                      If arr_varParm(M_PARR, lngY) = True Then
                        If IsNull(![vbcomparam_special]) = True Then
                          .Edit
                          ![vbcomparam_special] = "ParamArray"
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        Else
                          If ![vbcomparam_special] <> "ParamArray" Then
                            .Edit
                            ![vbcomparam_special] = "ParamArray"
                            ![vbcomparam_datemodified] = Now()
                            .Update
                            blnEdit = True
                          End If
                        End If
                      ElseIf arr_varParm(M_BARR, lngY) = True Then
                        If IsNull(![vbcomparam_special]) = True Then
                          .Edit
                          ![vbcomparam_special] = "ByteArray"
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        Else
                          If ![vbcomparam_special] <> "ByteArray" Then
                            .Edit
                            ![vbcomparam_special] = "ByteArray"
                            ![vbcomparam_datemodified] = Now()
                            .Update
                            blnEdit = True
                          End If
                        End If
                      Else
                        If IsNull(![vbcomparam_special]) = False Then
                          .Edit
                          ![vbcomparam_special] = Null
                          ![vbcomparam_datemodified] = Now()
                          .Update
                          blnEdit = True
                        End If
                      End If
                      If blnEdit = True Then
                        lngEdits2 = lngEdits2 + 1&
                      End If
                    Else
                      .AddNew
                      ![dbs_id] = arr_varProc(P_DID, lngX)
                      ![vbcom_id] = arr_varProc(P_VID, lngX)
                      ![vbcomproc_id] = arr_varProc(P_PID, lngX)
                      ' ** ![vbcomparam_id] : AutoNumber.
                      ![vbcomparam_order] = arr_varParm(M_ORD, lngY)
                      ![vbcomparam_name] = arr_varParm(M_MNAM, lngY)
                      If IsNull(arr_varParm(M_MTYP, lngY)) = False Then
                        ![vbcomparam_type] = arr_varParm(M_MTYP, lngY)
                      Else
                        ![vbcomparam_type] = Null
                      End If
                      If IsNull(arr_varParm(M_DTYP, lngY)) = False Then
                        ![datatype_vb_type] = arr_varParm(M_DTYP, lngY)
                      Else
                        ![datatype_vb_type] = Null
                      End If
                      intPos1 = InStr(arr_varParm(M_MTYP, lngY), ".")
                      If intPos1 > 0 Then
                        strTmp01 = Left(arr_varParm(M_MTYP, lngY), (intPos1 - 1))
                        strTmp05 = Mid(arr_varParm(M_MTYP, lngY), (intPos1 + 1))
                        Select Case strTmp01
                        Case "Access"
                          ' ** Acces.Form, etc.
                          varTmp00 = DLookup("[objtype_type]", "tblObjectType", "[objtype_constant] = 'vb" & strTmp05 & "'")
                          If IsNull(varTmp00) = False Then
                            ![objtype_type] = varTmp00
                          Else
                            ![objtype_type] = Null
                          End If
                        Case "DAO"
                          varTmp00 = DLookup("[daotype_id]", "tblDAOType", "[daotype_type] = '" & arr_varParm(M_MTYP, lngY) & "'")
                          If IsNull(varTmp00) = False Then
                            ![daotype_type] = arr_varParm(M_MTYP, lngY)
                          Else
                            ![daotype_type] = Null
                            Debug.Print "'NOT IN tblDAOType!  " & arr_varParm(M_MTYP, lngY)
                            DoEvents
                          End If
                        Case "Scripting"
                          varTmp00 = DLookup("[scripttype_id]", "tblScriptingType", "[scripttype_type] = '" & arr_varParm(M_MTYP, lngY) & "'")
                          If IsNull(varTmp00) = False Then
                            ![scripttype_type] = arr_varParm(M_MTYP, lngY)
                          Else
                            ![scripttype_type] = Null
                            Debug.Print "'NOT IN tblScriptingType!  " & arr_varParm(M_MTYP, lngY)
                            DoEvents
                          End If
                        End Select
                      Else
                        ![objtype_type] = Null
                        ![daotype_type] = Null
                        ![scripttype_type] = Null
                      End If
                      ![vbcomparam_optional] = arr_varParm(M_OPT, lngY)
                      If IsNull(arr_varParm(M_SRC, lngY)) = False Then
                        ![vbcomparam_explicit] = arr_varParm(M_SRC, lngY)
                      Else
                        ![vbcomparam_explicit] = Null
                      End If
                      If IsNull(arr_varParm(M_DEF, lngY)) = False Then
                        ![vbcomparam_default] = arr_varParm(M_DEF, lngY)
                      Else
                        ![vbcomparam_default] = Null
                      End If
                      If arr_varParm(M_PARR, lngY) = True Then
                        ![vbcomparam_special] = "ParamArray"
                      ElseIf arr_varParm(M_BARR, lngY) = True Then
                        ![vbcomparam_special] = "ByteArray"
                      Else
                        ![vbcomparam_special] = Null
                      End If
                      ![vbcomparam_datemodified] = Now()
                      .Update
                      .Bookmark = .LastModified
                      arr_varParm(M_MID, lngY) = ![vbcomparam_id]
                      lngAdds2 = lngAdds2 + 1&
                    End If  ' ** blnAdd.
                  Next  ' ** lngY.
                End If  ' ** None found.
              End If  ' ** ().
            End If  ' ** Declaration.
            arr_varProc(P_MCNT, lngX) = lngParms
            If lngParms > 0& Then
              arr_varProc(P_MARR, lngX) = arr_varParm
            End If

          Next  ' ** lngX.
        End With  ' ** rstParam.

      End If  ' ** lngProcs.

      ' ** Now check for deleted procs and params.
      lngDels = 0&
      ReDim arr_varDel(0)

      ' ** tblVBComponent_Procedure, by specified [comid].
      Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Proc_19")
      With qdf.Parameters
        ![comid] = lngVBComID
      End With
      Set rst1 = qdf.OpenRecordset
      With rst1
        If .BOF = True And .EOF = True Then
          If lngProcs > 0& Then
            Stop
          End If
        Else
          .MoveLast
          lngRecs = .RecordCount  ' ** All procedures in this module.
          .MoveFirst
          For lngX = 1& To lngRecs
            blnFound = False
            For lngY = 0& To (lngProcs - 1&)
              If arr_varProc(P_PID, lngY) = ![vbcomproc_id] Then
                blnFound = True
                arr_varProc(P_FND, lngY) = CBool(True)
                Exit For
              End If
            Next
            If blnFound = False Then
              lngDels = lngDels + 1&
              lngE = lngDels - 1&
              ReDim Preserve arr_varDel(lngE)
              arr_varDel(lngE) = ![vbcomproc_id]  ' ** Just this procedure.
            End If
            If lngX < lngRecs Then .MoveNext
          Next  ' ** lngX.
        End If  ' ** BOF, EOF.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing
      Set qdf = Nothing

      If lngDels > 0& Then
        blnDelete = True
        Debug.Print "'DELETE " & CStr(lngDels) & " PROCS IN " & strModName & "?"
        'Stop
        If blnDelete = True Then
          For lngX = 0& To (lngDels - 1&)
            ' ** Delete tblVBComponent_Procedure, by specified [comprocid].
            Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Proc_01a")
            With qdf.Parameters
              ![comprocid] = arr_varDel(lngX)  ' ** Delete 1 procedure.
            End With
            qdf.Execute
            Set qdf = Nothing
            lngProcDels = lngProcDels + 1&
          Next  ' ** lngX.
        End If  ' ** blnDelete.
      End If  ' ** lngDels.

      lngDels = 0&
      ReDim arr_varDel(0)

      For lngX = 0& To (lngProcs - 1&)

        lngParms = arr_varProc(P_MCNT, lngX)

        ' ** tblVBComponent_Procedure_Parameter, by specified [comprocid].
        Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Proc_Param_02")
        With qdf.Parameters
          ![comprocid] = arr_varProc(P_PID, lngX)
        End With
        Set rst1 = qdf.OpenRecordset
        With rst1
          If .BOF = True And .EOF = True Then
            If lngParms > 0& Then
              Stop
            End If
          Else
            If lngParms > 0 Then
              arr_varParm = arr_varProc(P_MARR, lngX)
            End If
            .MoveLast
            lngRecs = .RecordCount  ' ** All parameters in this procedure.
            .MoveFirst
            For lngY = 1& To lngRecs
              blnFound = False
              If lngParms > 0& Then
                For lngZ = 0& To (lngParms - 1&)
                  If arr_varParm(M_MID, lngZ) = ![vbcomparam_id] Then
                    blnFound = True
                    Exit For
                  End If
                Next  ' ** lngZ.
              End If
              If blnFound = False Then
                lngDels = lngDels + 1&
                lngE = lngDels - 1&
                ReDim Preserve arr_varDel(lngE)
                arr_varDel(lngE) = ![vbcomparam_id]  ' ** Just this parameter.
              End If
              If lngY < lngRecs Then .MoveNext
            Next  ' ** lngY.
          End If  ' ** BOF, EOF.
          .Close
        End With  ' ** rst1.
        Set rst1 = Nothing
        Set qdf = Nothing

        If lngDels > 0& Then
          blnDelete = True
          Debug.Print "'DELETE " & CStr(lngDels) & " PRARAMS IN " & strModName & "." & arr_varProc(P_PNAM, lngX) & "()?"
          Stop
          If blnDelete = True Then
            For lngY = 0& To (lngDels - 1&)
              ' ** Delete tblVBComponent_Procedure_Parameter, by specified [paramid].
              Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Proc_Param_01a")
              With qdf.Parameters
                ![paramid] = arr_varDel(lngY)  ' ** Delete 1 parameter.
              End With
              qdf.Execute
              Set qdf = Nothing
              lngParmDels = lngParmDels + 1&
            Next  ' ** lngY.
          End If  ' ** blnDelete.
        End If

      Next  ' ** lngX.

    Next  ' ** vbc.
  End With  ' ** vbp.

  rstProc.Close
  rstParam.Close

  With dbs
    ' ** tblVBComponent_Procedure, just ends < begs.
    varTmp00 = DCount("*", "zz_qry_VBComponent_Proc_09b")
    If varTmp00 > 0 Then
      ' ** Update zz_qry_VBComponent_Proc_09b (tblVBComponent_Procedure,
      ' ** just ends < begs), just 'Declaration', for vbcompro_line_beg = 1.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Proc_09c")
      qdf.Execute
      Debug.Print "'BEG LINES OFF!  " & CStr(varTmp00)
      DoEvents
    End If
    .Close
  End With

  If lngEdits1 > 0& Then
    Debug.Print "'PROC EDITS: " & CStr(lngEdits1)
  End If
  If lngAdds1 > 0& Then
    Debug.Print "'PROC ADDS: " & CStr(lngAdds1)
    For lngX = 0& To (lngAdds1 - 1&)
      Debug.Print "'  NEW: " & arr_varAdd1(A_PNAM, lngX) & "  IN  " & arr_varAdd1(A_VNAM, lngX) & "  " & arr_varAdd1(A_PTYP, lngX) & _
        IIf(IsNull(arr_varAdd1(A_STYP, lngX)) = True, vbNullString, "  " & arr_varAdd1(A_STYP, lngX))
    Next
  End If
  If lngProcDels > 0& Then
    Debug.Print "'PROC DELS: " & CStr(lngProcDels)
  End If

  If lngEdits2 > 0& Then
    Debug.Print "'PARAM EDITS: " & CStr(lngEdits2)
  End If
  If lngAdds2 > 0& Then
    Debug.Print "'PARAM ADDS: " & CStr(lngAdds2)
  End If
  If lngParmDels > 0& Then
    Debug.Print "'PARAM DELS: " & CStr(lngParmDels)
  End If

  If lngEdits1 = 0& And lngAdds1 = 0& And lngEdits2 = 0& And lngAdds2 = 0& And lngProcDels = 0& And lngParmDels = 0& Then
    Debug.Print "'NO CHANGES!"
  ElseIf lngEdits1 = 0& And lngEdits2 = 0& Then
    Debug.Print "'NO EDITS!"
  ElseIf lngAdds1 = 0& And lngAdds2 = 0& Then
    Debug.Print "'NO ADDS!"
  End If

  Debug.Print "'DONE!  " & THIS_PROC & "()"

'NO CHANGES!
'DONE!  VBA_Proc_Doc_New()

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rstProc = Nothing
  Set rstParam = Nothing
  Set rst1 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_Proc_Doc_New = blnRetVal

End Function

Private Function VBA_ProcParamSplit(strDeclare As String) As Variant

  Const THIS_PROC As String = "VBA_ProcParamSplit"

  Dim lngParms As Long, arr_varParm() As Variant
  Dim strSource As String, lngDataType As Long
  Dim lngVars As Long, lngComs As Long, lngAs As Long
  Dim blnFound As Boolean, blnOption As Boolean, blnParamArray As Boolean, blnByteArray As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intLen As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, arr_varTmp05() As Variant
  Dim intX As Integer, lngY As Long, lngE As Long
  Dim arr_varRetVal As Variant

  ' ** Array: arr_varParm().
  Const M_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const M_ORD   As Integer = 0
  Const M_MID   As Integer = 1
  Const M_MNAM  As Integer = 2
  Const M_MTYP  As Integer = 3
  Const M_DTYP  As Integer = 4
  Const M_SRC   As Integer = 5
  Const M_OPT   As Integer = 6
  Const M_PARR  As Integer = 7
  Const M_BARR  As Integer = 8
  Const M_DEF   As Integer = 9
  Const M_NOTYP As Integer = 10

  If strDeclare <> vbNullString Then
    intPos1 = InStr(strDeclare, "(")
    If intPos1 > 0 Then
      strTmp01 = Mid$(strDeclare, intPos1)
      blnFound = False
      If Right$(strTmp01, 1) <> ")" Then
        intLen = Len(strTmp01)
        For intX = intLen To 1 Step -1
          If Mid$(strTmp01, intX, 1) = ")" Then
            blnFound = True
            strTmp01 = Left$(strTmp01, intX)
            Exit For
          End If
        Next
      Else
        blnFound = True
      End If
      If blnFound = True And strTmp01 <> "()" Then
        lngParms = 0&
        ReDim arr_varParm(M_ELEMS, 0)
        lngVars = 0&: lngComs = 0&: lngAs = 0&
        strTmp01 = Mid$(Left$(strTmp01, (Len(strTmp01) - 1)), 2)  ' ** Strip the parens.
        intPos1 = InStr(strTmp01, ",")
        If intPos1 > 0 Then
          ' ** Careful about defaults and commas in arrays or strings!
          lngComs = CharCnt(strTmp01, ",")  ' ** Module Function: modStringFuncs.
          lngAs = CharCnt(strTmp01, " As ", True)  ' ** Module Function: modStringFuncs.
          If lngAs = lngComs + 1& Then  ' ** There should always be 1 less comma than parameters.
            For lngY = 1& To lngAs
              lngDataType = vbUndeclared
              blnOption = False: blnParamArray = False: blnByteArray = False: strSource = vbNullString
              strTmp03 = vbNullString: strTmp04 = vbNullString
              If intPos1 > 0 Then
                strTmp02 = Left$(strTmp01, (intPos1 - 1))  ' ** This parameter.
                strTmp01 = Trim$(Mid$(strTmp01, (intPos1 + 1)))  ' ** Remaining parameters.
              Else
                strTmp02 = strTmp01  ' ** Last parameter.
                strTmp01 = vbNullString
              End If
              intPos2 = InStr(strTmp02, " ")
              If intPos2 > 0 Then
                strTmp03 = Trim$(Mid$(strTmp02, intPos2))
                strTmp02 = Trim$(Left$(strTmp02, intPos2))
                If strTmp02 = "Optional" Then
                  blnOption = True
                  intPos2 = InStr(strTmp03, " ")
                  If intPos2 > 0 Then
                    strTmp02 = Trim$(Left$(strTmp03, intPos2))
                    strTmp03 = Trim$(Mid$(strTmp03, intPos2))
                  Else
                    ' ** Untyped parameter!
                    strTmp02 = strTmp03
                    strTmp03 = vbNullString
                  End If
                End If
                If strTmp02 = "ParamArray" Then
                  blnParamArray = True
                  intPos2 = InStr(strTmp03, " ")
                  If intPos2 > 0 Then
                    strTmp02 = Trim$(Left$(strTmp03, intPos2))
                    strTmp03 = Trim$(Mid$(strTmp03, intPos2))
                  Else
                    ' ** Untyped parameter!
                    strTmp02 = strTmp03
                    strTmp03 = vbNullString
                  End If
                End If
                If strTmp02 = "ByteArray" Then
                  blnByteArray = True
                  intPos2 = InStr(strTmp03, " ")
                  If intPos2 > 0 Then
                    strTmp02 = Trim$(Left$(strTmp03, intPos2))
                    strTmp03 = Trim$(Mid$(strTmp03, intPos2))
                  Else
                    ' ** Untyped parameter!
                    strTmp02 = strTmp03
                    strTmp03 = vbNullString
                  End If
                End If
                If strTmp02 = "ByRef" Or strTmp02 = "ByVal" Then
                  strSource = strTmp02
                  intPos2 = InStr(strTmp03, " ")
                  If intPos2 > 0 Then
                    strTmp02 = Trim$(Left$(strTmp03, intPos2))
                    strTmp03 = Trim$(Mid$(strTmp03, intPos2))
                  Else
                    ' ** Untyped parameter!
                    strTmp02 = strTmp03
                    strTmp03 = vbNullString
                  End If
                End If
'## CHECK!
'If strTmp02 = "Optional" Or strTmp02 = "ParamArray" Or strTmp02 = "ByRef" Or strTmp02 = "ByVal" Then
'  ' ** At this point, none of these should be in strTmp02!
'  Debug.Print "'" & strTmp02
'  Stop
'End If
'##
                If strTmp03 <> vbNullString Then
                  If Left$(strTmp03, 3) = "As " Then
                    strTmp03 = Trim$(Mid$(strTmp03, 3))
                    ' ** Check for a default.
                    intPos2 = InStr(strTmp03, " ")
                    If intPos2 > 0 Then
                      strTmp04 = Trim$(Mid$(strTmp03, intPos2))
                      If Left$(strTmp04, 1) = "=" Then
                        strTmp04 = Trim$(Mid$(strTmp04, 2))
                      Else
                        Debug.Print "'WHAT IS THIS?  " & strTmp04
                        Stop
                      End If
                      strTmp03 = Trim$(Left$(strTmp03, intPos2))
                    End If
                    varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", "[datatype_vb_constant] = 'vb" & strTmp03 & "'")
                    Select Case IsNull(varTmp00)
                    Case True
                      If IsUC(strTmp03, True, True) = True Then  ' ** Module Function: modStringFuncs.
                        lngDataType = vbUserDefinedType
                      Else
                        lngDataType = vbVariant
                      End If
                    Case False
                      lngDataType = varTmp00
                    End Select
                  Else
                    Debug.Print "'WHAT IS THIS?  " & strTmp03
                    Stop
                  End If
                End If
                ' ** strTmp02 should be the parameter name, strTmp03 its data type, and strTmp04 any default.
                lngParms = lngParms + 1&
                lngE = lngParms - 1&
                ReDim Preserve arr_varParm(M_ELEMS, lngE)
                arr_varParm(M_ORD, lngE) = lngParms
                arr_varParm(M_MID, lngE) = CLng(0)
                arr_varParm(M_MNAM, lngE) = strTmp02
                If strTmp03 <> vbNullString Then
                  arr_varParm(M_MTYP, lngE) = strTmp03
                Else
                  arr_varParm(M_MTYP, lngE) = Null
                End If
                arr_varParm(M_DTYP, lngE) = lngDataType
                If strSource <> vbNullString Then
                  arr_varParm(M_SRC, lngE) = strSource
                Else
                  arr_varParm(M_SRC, lngE) = Null
                End If
                arr_varParm(M_OPT, lngE) = blnOption
                arr_varParm(M_PARR, lngE) = blnParamArray
                arr_varParm(M_BARR, lngE) = blnByteArray
                If strTmp04 <> vbNullString Then
                  arr_varParm(M_DEF, lngE) = strTmp04
                Else
                  arr_varParm(M_DEF, lngE) = Null
                End If
               If strTmp03 <> vbNullString Then
                  arr_varParm(M_NOTYP, lngE) = CBool(False)
                Else
                  arr_varParm(M_NOTYP, lngE) = CBool(True)
                End If
              Else
                ' ** Untyped parameter!
                lngParms = lngParms + 1&
                lngE = lngParms - 1&
                ReDim Preserve arr_varParm(M_ELEMS, lngE)
                arr_varParm(M_ORD, lngE) = lngParms
                arr_varParm(M_MID, lngE) = CLng(0)
                arr_varParm(M_MNAM, lngE) = strTmp02
                arr_varParm(M_MTYP, lngE) = "Variant"
                arr_varParm(M_DTYP, lngE) = vbVariant
                arr_varParm(M_SRC, lngE) = Null
                arr_varParm(M_OPT, lngE) = blnOption
                arr_varParm(M_PARR, lngE) = blnParamArray
                arr_varParm(M_BARR, lngE) = blnByteArray
                arr_varParm(M_DEF, lngE) = Null
                arr_varParm(M_NOTYP, lngE) = CBool(True)
              End If
              intPos1 = InStr(strTmp01, ",")
            Next  ' ** lngY.
          Else
            Debug.Print "'COMMAS DON'T ADD UP!"
            Debug.Print "'" & strDeclare
            DoEvents
            Stop
'Private Function GetData2(frm, strPricePathFile_TXT As String, dblProgBox_Width As Double,
'  dblProgBar_Len As Double, intPct As Integer, intCenteringSpaces) As Variant
          End If  ' ** lngComs, lngAs.
        Else
          lngVars = 1&
          lngDataType = vbUndeclared
          blnOption = False: blnParamArray = False: blnByteArray = False: strSource = vbNullString
          intPos2 = InStr(strTmp01, " ")
          strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString
          If intPos2 > 0 Then
            strTmp02 = Trim$(Left$(strTmp01, intPos2))
            strTmp03 = Trim$(Mid$(strTmp01, intPos2))
            If strTmp02 = "Optional" Then
              blnOption = True
              intPos2 = InStr(strTmp03, " ")
              If intPos2 > 0 Then
                strTmp02 = Trim$(Left$(strTmp03, intPos2))
                strTmp03 = Trim$(Mid$(strTmp03, intPos2))
              Else
                ' ** Untyped parameter!
                strTmp02 = strTmp03
                strTmp03 = vbNullString
              End If
            End If
            If strTmp02 = "ParamArray" Then
              blnParamArray = True
              intPos2 = InStr(strTmp03, " ")
              If intPos2 > 0 Then
                strTmp02 = Trim$(Left$(strTmp03, intPos2))
                strTmp03 = Trim$(Mid$(strTmp03, intPos2))
              Else
                ' ** Untyped parameter!
                strTmp02 = strTmp03
                strTmp03 = vbNullString
              End If
            End If
            If strTmp02 = "ByteArray" Then
              blnByteArray = True
              intPos2 = InStr(strTmp03, " ")
              If intPos2 > 0 Then
                strTmp02 = Trim$(Left$(strTmp03, intPos2))
                strTmp03 = Trim$(Mid$(strTmp03, intPos2))
              Else
                ' ** Untyped parameter!
                strTmp02 = strTmp03
                strTmp03 = vbNullString
              End If
            End If
            If strTmp02 = "ByRef" Or strTmp02 = "ByVal" Then
              strSource = strTmp02
              intPos2 = InStr(strTmp03, " ")
              If intPos2 > 0 Then
                strTmp02 = Trim$(Left$(strTmp03, intPos2))
                strTmp03 = Trim$(Mid$(strTmp03, intPos2))
              Else
                ' ** Untyped parameter!
                strTmp02 = strTmp03
                strTmp03 = vbNullString
              End If
            End If
'## CHECK !
'If strTmp02 = "Optional" Or strTmp02 = "ParamArray" Or strTmp02 = "ByRef" Or strTmp02 = "ByVal" Then
'  ' ** At this point, none of these should be in strTmp02!
'  Debug.Print "'" & strTmp02
'  Stop
'End If
'##
            If strTmp03 <> vbNullString Then
              If Left$(strTmp03, 3) = "As " Then
                strTmp03 = Trim$(Mid$(strTmp03, 3))
                ' ** Check for a default.
                intPos2 = InStr(strTmp03, " ")
                If intPos2 > 0 Then
                  strTmp04 = Trim$(Mid$(strTmp03, intPos2))
                  If Left$(strTmp04, 1) = "=" Then
                    strTmp04 = Trim$(Mid$(strTmp04, 2))
                  Else
                    Debug.Print "'WHAT IS THIS?  " & strTmp04
                    Stop
                  End If
                  strTmp03 = Trim$(Left$(strTmp03, intPos2))
                End If
                varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", "[datatype_vb_constant] = 'vb" & strTmp03 & "'")
                Select Case IsNull(varTmp00)
                Case True
                  If IsUC(strTmp03, True, True) = True Then  ' ** Module Function: modStringFuncs.
                    lngDataType = vbUserDefinedType
                  Else
                    lngDataType = vbVariant
                  End If
                Case False
                  lngDataType = varTmp00
                End Select
              Else
                Debug.Print "'WHAT IS THIS?  " & strTmp03
                Stop
              End If
            End If
            ' ** strTmp02 should be the parameter name, strTmp03 its data type, and strTmp04 any default.
            lngParms = lngParms + 1&
            lngE = lngParms - 1&
            ReDim Preserve arr_varParm(M_ELEMS, lngE)
            arr_varParm(M_ORD, lngE) = lngParms
            arr_varParm(M_MID, lngE) = CLng(0)
            arr_varParm(M_MNAM, lngE) = strTmp02
            If strTmp03 <> vbNullString Then
              arr_varParm(M_MTYP, lngE) = strTmp03
            Else
              arr_varParm(M_MTYP, lngE) = Null
            End If
            arr_varParm(M_DTYP, lngE) = lngDataType
            If strSource <> vbNullString Then
              arr_varParm(M_SRC, lngE) = strSource
            Else
              arr_varParm(M_SRC, lngE) = Null
            End If
            arr_varParm(M_OPT, lngE) = blnOption
            arr_varParm(M_PARR, lngE) = blnParamArray
            arr_varParm(M_BARR, lngE) = blnByteArray
            If strTmp04 <> vbNullString Then
              arr_varParm(M_DEF, lngE) = strTmp04
            Else
              arr_varParm(M_DEF, lngE) = Null
            End If
            If strTmp03 <> vbNullString Then
              arr_varParm(M_NOTYP, lngE) = CBool(False)
            Else
              arr_varParm(M_NOTYP, lngE) = CBool(True)
            End If
          Else
            ' ** Untyped parameter!
            lngParms = lngParms + 1&
            lngE = lngParms - 1&
            ReDim Preserve arr_varParm(M_ELEMS, lngE)
            arr_varParm(M_ORD, lngE) = lngParms
            arr_varParm(M_MID, lngE) = CLng(0)
            arr_varParm(M_MNAM, lngE) = strTmp01
            arr_varParm(M_MTYP, lngE) = "Variant"
            arr_varParm(M_DTYP, lngE) = vbVariant
            arr_varParm(M_SRC, lngE) = Null
            arr_varParm(M_OPT, lngE) = blnOption
            arr_varParm(M_PARR, lngE) = blnParamArray
            arr_varParm(M_BARR, lngE) = blnByteArray
            arr_varParm(M_DEF, lngE) = Null
            arr_varParm(M_NOTYP, lngE) = CBool(True)
          End If
        End If  ' ** Comma.
        arr_varRetVal = arr_varParm
      Else
        If strTmp01 = "()" Then
          ReDim arr_varTmp05(0, 0)
          arr_varTmp05(0, 0) = CLng(0)
          arr_varRetVal = arr_varTmp05
        Else
          Debug.Print "'" & strTmp01
          DoEvents
          Stop
        End If
      End If  ' ** blnFound.
    Else
      Stop
    End If  ' ** Paren.
  Else
    ReDim arr_varTmp05(0, 0)
    arr_varTmp05(0, 0) = CLng(0)
    arr_varRetVal = arr_varTmp05
  End If

'## CHECK!
'If lngParms > 0& Then
'  Debug.Print "'PARMS: " & CStr(lngParms)
'  intX = 0
'  For lngY = 0& To (lngParms - 1&)
'    If Len(arr_varParm(M_MNAM, lngY)) > intX Then intX = Len(arr_varParm(M_MNAM, lngY))
'  Next
'  For lngY = 0& To (lngParms - 1&)
'    Debug.Print "'" & CStr(arr_varParm(M_ORD, lngY)) & ". " & Left$(arr_varParm(M_MNAM, lngY) & Space(intX), intX) & "  " & _
'      arr_varParm(M_MTYP, lngY) & "  " & arr_varParm(M_DTYP, lngY) & _
'      "  SRC: " & Nz(arr_varParm(M_SRC, lngY), vbNullString) & _
'      "  OPT: " & arr_varParm(M_OPT, lngY) & _
'      "  PARR: " & arr_varParm(M_PARR, lngY) & _
'      "  DEF: " & Nz(arr_varParm(M_MTYP, lngY), vbNullString) & _
'      "  NOTYP: " & arr_varParm(M_NOTYP, lngY)
'    DoEvents
'  Next
'Else
'  Debug.Print "'NONE FOUND!"
'  DoEvents
'  Stop
'End If
'##

  VBA_ProcParamSplit = arr_varRetVal

End Function

Public Function VBA_ChkProps() As Boolean

  Const THIS_PROC As String = "VBA_ChkProps"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngModLines As Long, lngModDecLines As Long
  Dim strModName As String, strLine As String
  Dim strThisProcName As String, strLastProcName As String
  Dim strThisDeclare As String, strLastDeclare As String, strThisProcSubType As String, strLastProcSubType As String
  Dim strThisFirstCodeNum As String, strThisLastCodeNum As String, strLastFirstCodeNum As String, strLastLastCodeNum As String
  Dim lngThisDeclareLine As Long, lngLastDeclareLine As Long, lngThisEndLine As Long, lngLastEndLine As Long
  Dim lngProcs As Long, arr_varProc As Variant
  Dim blnSplit As Boolean, blnAddThis As Boolean, blnAddLast As Boolean
  Dim lngThisDbsID As Long
  Dim lngAdds As Long, lngEdits As Long
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varProc().
  Const P_DID  As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_VID  As Integer = 2
  Const P_VNAM As Integer = 3
  Const P_VTYP As Integer = 4
  Const P_PID  As Integer = 5
  Const P_PNAM As Integer = 6
  Const P_PTYP As Integer = 7
  Const P_PSUB As Integer = 8
  Const P_BEG  As Integer = 9
  Const P_END  As Integer = 10
  Const P_SCOP As Integer = 11
  Const P_RET  As Integer = 12
  Const P_DTYP As Integer = 13
  Const P_EVNT As Integer = 14

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
  lngEdits = 0&: lngAdds = 0&

  Set dbs = CurrentDb
  With dbs

    ' ** tblVBComponent_Procedure, just those with procsubtype_type.
    Set qdf = .QueryDefs("zz_qry_VBComponent_Proc_20a")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngProcs = .RecordCount
      .MoveFirst
      arr_varProc = .GetRows(lngProcs)
      ' *********************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   FIELD  ELEMENT  NAME                    CONSTANT
      ' **   =====  =======  ======================  ==========
      ' **     1       0     dbs_id                  P_DID
      ' **     2       1     dbs_name                P_DNAM
      ' **     3       2     vbcom_id                P_VID
      ' **     4       3     vbcom_name              P_VNAM
      ' **     5       4     comtype_type            P_VTYP
      ' **     6       5     vbcomproc_id            P_PID
      ' **     7       6     vbcomproc_name          P_PNAM
      ' **     8       7     proctype_type           P_PTYP
      ' **     9       8     procsubtype_type        P_PSUB
      ' **    10       9     vbcomproc_line_beg      P_BEG
      ' **    11      10     vbcomproc_line_end      P_END
      ' **    12      11     scopetype_type          P_SCOP
      ' **    13      12     vbcomproc_returntype    P_RET
      ' **    14      13     datatype_vb_type        P_DTYP
      ' **    15      14     vbcom_event_id          P_EVNT
      ' **
      ' *********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbConsistent)

  End With

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For lngX = 0& To (lngProcs - 1&)
      Set vbc = .VBComponents(arr_varProc(P_VNAM, lngX))
      With vbc
        strModName = .Name
        Set cod = .CodeModule
        With cod
          lngModLines = .CountOfLines
          lngModDecLines = .CountOfDeclarationLines
          strThisDeclare = vbNullString: strLastDeclare = vbNullString
          strThisProcSubType = vbNullString: strLastProcSubType = vbNullString
          strThisProcName = vbNullString: strLastProcName = vbNullString
          strThisFirstCodeNum = vbNullString: strThisLastCodeNum = vbNullString
          strLastFirstCodeNum = vbNullString: strLastLastCodeNum = vbNullString
          lngThisDeclareLine = 0&: lngLastDeclareLine = 0&: lngThisEndLine = 0&: lngLastEndLine = 0&
          blnSplit = False
          For lngY = arr_varProc(P_BEG, lngX) To arr_varProc(P_END, lngX)
            strLine = Trim(.Lines(lngY, 1))
            If strLine <> vbNullString Then
              If Left$(strLine, 1) <> "'" Then
                intPos1 = InStr(strLine, " ")  ' ** 1st space.
                If intPos1 > 0 Then
                  strTmp01 = Trim$(Left$(strLine, intPos1))
                  If IsNumeric(strTmp01) = True Then
                    If strThisDeclare <> vbNullString And strLastDeclare = vbNullString And _
                        strThisFirstCodeNum = vbNullString Then
                      strThisFirstCodeNum = strTmp01
                    ElseIf strThisDeclare <> vbNullString And strLastDeclare <> vbNullString And _
                        strThisFirstCodeNum = vbNullString And strLastFirstCodeNum <> vbNullString Then
                      strThisFirstCodeNum = strTmp01
                    Else
                      strThisLastCodeNum = strTmp01
                    End If
                  End If
                  strTmp01 = vbNullString
                  intPos2 = InStr((intPos1 + 1), strLine, " ")  ' ** 2nd space.
                  If intPos2 > 0 Then
                    strTmp01 = Trim$(Left$(strLine, intPos2))
                    If strTmp01 = "Public Property" Or strTmp01 = "Private Property" Then  ' ** There may never be a Private.
                      strTmp02 = Trim$(Mid$(strLine, intPos2))
                      intPos3 = InStr(strTmp02, " ")
                      strThisProcSubType = Trim$(Left$(strTmp02, intPos3))
                      strTmp02 = Trim$(Mid$(strTmp02, intPos3))
                      intPos3 = InStr(strTmp02, "(")
                      strThisProcName = Left$(strTmp02, (intPos3 - 1))
                      lngThisDeclareLine = lngY
                      If strLastDeclare <> vbNullString Then
                        strThisDeclare = strLine
                        If strLine <> strLastDeclare Then
                          If strLastProcName = strThisProcName And strLastProcSubType <> strThisProcSubType Then
                            ' ** Different procedure!
                            blnSplit = True
                          End If
                        End If
                      Else
                        strThisDeclare = strLine
                      End If
                    End If
                  Else
                    If strLine = "End Property" Then
                      lngThisEndLine = lngY
                      If lngLastEndLine <> 0& Then
                        If lngThisEndLine <> lngLastEndLine Then
                          ' ** This should be the end.
                        End If
                      Else
                        strLastDeclare = strThisDeclare
                        strThisDeclare = vbNullString
                        lngLastDeclareLine = lngThisDeclareLine
                        lngThisDeclareLine = 0&
                        strLastProcName = strThisProcName
                        strThisProcName = vbNullString
                        strLastProcSubType = strThisProcSubType
                        strThisProcSubType = vbNullString
                        lngLastEndLine = lngThisEndLine
                        lngThisEndLine = 0&
                        strLastFirstCodeNum = strThisFirstCodeNum
                        strThisFirstCodeNum = vbNullString
                        strLastLastCodeNum = strThisLastCodeNum
                        strThisLastCodeNum = vbNullString
                      End If
                    End If
                  End If
                End If
              End If  ' ** Remark.
            End If  ' ** vbNullString.
          Next  ' ** lngY
          If blnSplit = True Then
            If lngThisDeclareLine > 0& And lngLastDeclareLine > 0& And lngThisEndLine > 0& And lngLastEndLine > 0& Then
              ' ** Edit and AddNew.
              If strThisFirstCodeNum = vbNullString Or strThisLastCodeNum = vbNullString Or _
                  strLastFirstCodeNum = vbNullString Or strLastLastCodeNum = vbNullString Then
                Stop
              End If
              blnAddThis = False: blnAddLast = False
              With rst
                .FindFirst "[vbcomproc_id] = " & CStr(arr_varProc(P_PID, lngX))
                If .NoMatch = False Then
                  If ![procsubtype_type] = strThisProcSubType Then
                    blnAddLast = True
                    If ![vbcomproc_line_beg] <> lngThisDeclareLine Or ![vbcomproc_line_end] <> lngThisEndLine Or _
                        ![vbcomproc_code_beg] <> strThisFirstCodeNum Or ![vbcomproc_code_end] <> strThisLastCodeNum Then
                      .Edit
                      ![vbcomproc_line_beg] = lngThisDeclareLine
                      ![vbcomproc_line_end] = lngThisEndLine  ' ** Though the end usually gets the blank line after.
                      ![vbcomproc_code_beg] = strThisFirstCodeNum
                      ![vbcomproc_code_end] = strThisLastCodeNum
                      ![vbcomproc_datemodified] = Now()
                      .Update
                      lngEdits = lngEdits + 1&
                    End If
                  ElseIf ![procsubtype_type] = strLastProcSubType Then
                    blnAddThis = True
                    If ![vbcomproc_line_beg] <> lngLastDeclareLine Or ![vbcomproc_line_end] <> lngLastEndLine Or _
                        ![vbcomproc_code_beg] <> strLastFirstCodeNum Or ![vbcomproc_code_end] <> strLastLastCodeNum Then
                      .Edit
                      ![vbcomproc_line_beg] = lngLastDeclareLine
                      ![vbcomproc_line_end] = lngLastEndLine  ' ** Though the end usually gets the blank line after.
                      ![vbcomproc_code_beg] = strLastFirstCodeNum
                      ![vbcomproc_code_end] = strLastLastCodeNum
                      ![vbcomproc_datemodified] = Now()
                      .Update
                      lngEdits = lngEdits + 1&
                    End If
                  Else
                    Stop
                  End If
                  If blnAddThis = True Then
                    .AddNew
                    ![dbs_id] = lngThisDbsID
                    ![vbcom_id] = arr_varProc(P_VID, lngX)
                    ' ** ![vbcomproc_id] : AutoNumber.
                    ![vbcomproc_name] = arr_varProc(P_PNAM, lngX)
                    ![vbcomproc_line_beg] = lngThisDeclareLine
                    ![vbcomproc_line_end] = lngThisEndLine
                    ![vbcomproc_code_beg] = strThisFirstCodeNum
                    ![vbcomproc_code_end] = strThisLastCodeNum
                    ![scopetype_type] = arr_varProc(P_SCOP, lngX)
                    ![proctype_type] = arr_varProc(P_PTYP, lngX)
                    If strThisProcSubType = "Get" Then
                      ![vbcomproc_returntype] = GetLastWord(strThisDeclare)  ' ** Module Function: modStringFuncs.
                    ElseIf strThisProcSubType = "Let" Or strThisProcSubType = "Set" Then
                      ![vbcomproc_returntype] = Null
                    Else
                      Stop
                    End If
                    If strThisProcSubType = "Get" Then
                      varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", _
                        "[datatype_vb_constant] = 'vb" & GetLastWord(strThisDeclare) & "'")  ' ** Module Function: modStringFuncs.
                      If IsNull(varTmp00) = False Then
                        ![datatype_vb_type] = varTmp00
                      Else
                        Stop
                      End If
                    Else
                      ![datatype_vb_type] = Null
                    End If
                    ![procsubtype_type] = strThisProcSubType
                    ![vbcom_event_id] = Null
                    ![vbcomproc_multiline] = False
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    lngAdds = lngAdds + 1&
                  ElseIf blnAddLast = True Then
                    .AddNew
                    ![dbs_id] = lngThisDbsID
                    ![vbcom_id] = arr_varProc(P_VID, lngX)
                    ' ** ![vbcomproc_id] : AutoNumber.
                    ![vbcomproc_name] = arr_varProc(P_PNAM, lngX)
                    ![vbcomproc_line_beg] = lngLastDeclareLine
                    ![vbcomproc_line_end] = lngLastEndLine
                    ![vbcomproc_code_beg] = strLastFirstCodeNum
                    ![vbcomproc_code_end] = strLastLastCodeNum
                    ![scopetype_type] = arr_varProc(P_SCOP, lngX)
                    ![proctype_type] = arr_varProc(P_PTYP, lngX)
                    If strLastProcSubType = "Get" Then
                      ![vbcomproc_returntype] = GetLastWord(strLastDeclare)  ' ** Module Function: modStringFuncs.
                    ElseIf strLastProcSubType = "Let" Then
                      ![vbcomproc_returntype] = Null
                    Else
                      Stop
                    End If
                    varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", _
                      "[datatype_vb_constant] = 'vb" & GetLastWord(strLastDeclare) & "'")  ' ** Module Function: modStringFuncs.
                    If IsNull(varTmp00) = False Then
                      ![datatype_vb_type] = varTmp00
                    Else
                      Stop
                    End If
                    ![procsubtype_type] = strLastProcSubType
                    ![vbcom_event_id] = Null
                    ![vbcomproc_multiline] = False
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    lngAdds = lngAdds + 1&
                  End If
                Else
                  Stop
                End If
              End With  ' ** rst.
            Else
              Stop
            End If
          End If
        End With  ' ** cod.
        Set cod = Nothing
      End With  ' ** vbc.
      Set vbc = Nothing
    Next  ' ** lngX.
  End With  ' ** vbp.

  rst.Close
  dbs.Close

  If lngEdits = 0& And lngAdds = 0& Then
    Debug.Print "'NO CHANGES!"
  Else
    Debug.Print "'EDITS: " & CStr(lngEdits) & "  ADDS: " & CStr(lngAdds)
  End If
  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_ChkProps = blnRetVal

End Function

Private Function VBA_MsgBox_Doc() As Boolean
' ** Document all MsgBox() functions in the code to tblVBComponent_MessageBox.
' ** Called by:
' **   QuikVBADoc(), Above

  Const THIS_PROC As String = "VBA_MsgBox_Doc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strModName As String, strProcName As String
  Dim lngMsgs As Long, arr_varMsg() As Variant
  Dim lngThisDbsID As Long, lngComID As Long, lngComProcID As Long
  Dim strLineNum As String, lngLineNum As Long
  Dim strMsgBox As String
  Dim lngLines As Long, lngDecLines As Long, lngCnt As Long
  Dim strLine As String
  Dim strFind1 As String
  Dim blnLineCont As Boolean
  Dim blnAdd As Boolean, blnFound As Boolean, blnFound2 As Boolean
  Dim lngAdds As Long, lngEdits As Long
  Dim lngRecs As Long
  Dim lngDels As Long, arr_varDel() As Variant
  Dim intPos1 As Integer, intPos2 As Integer
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varMsg().
  Const M_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const M_DID As Integer = 0
  Const M_VID As Integer = 1
  Const M_PID As Integer = 2
  Const M_LIN As Integer = 3
  Const M_MSG As Integer = 4
  Const M_ERH As Integer = 5

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strFind1 = "MsgBox"

  lngMsgs = 0&
  ReDim arr_varMsg(M_ELEMS, 0)

  lngAdds = 0&: lngEdits = 0&

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        lngComID = DLookup("[vbcom_id]", "tblVBComponent", "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & strModName & "'")
        If lngComID = 0& Then
          Stop
        End If
        strProcName = vbNullString: strMsgBox = vbNullString
        blnLineCont = False
        If Left$(strModName, 7) <> "zz_mod_" Then  ' ** Skip mine.
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            If lngLines > lngDecLines Then
              For lngX = (lngDecLines + 1&) To lngLines
                If .ProcOfLine(lngX, vbext_pk_Proc) <> strProcName Then
                  strMsgBox = vbNullString
                  blnLineCont = False
                  blnFound2 = False
                End If
                strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
If strProcName <> "Qry_ChkParams" Then
                strLine = Trim$(.Lines(lngX, 1))
                If strLine <> vbNullString Then
                  If Left$(strLine, 1) <> "'" Then
                    If strLine = "ERRH:" Then blnFound2 = True
                    intPos1 = InStr(strLine, strFind1)
                    If intPos1 > 0 Then
                      intPos2 = InStr(strLine, "DEF_MSGBOX")
                      If intPos2 > 0 Then
                        If (intPos2 + 4) = intPos1 Then
                          intPos1 = InStr((intPos1 + 1), strLine, strFind1)
                        End If
                      End If
                      intPos2 = 0
                    End If
                    intPos2 = InStr(strLine, "VbMsgBoxResult")
                    If ((intPos2 > 0) And (intPos2 = (intPos1 - 2))) Then intPos1 = 0
                    If Left$(strLine, 7) = "Private" Then intPos1 = 0
                    If intPos1 > 0 Or (intPos1 = 0 And blnLineCont = True) Then
                      If blnLineCont = True Then
                        intPos1 = 1
                      End If
                      intPos2 = InStr(strLine, " '")  ' ** Check fo a remark elsewhere on the line.
                      If intPos2 > 0 And intPos2 < intPos1 Then
                        ' ** Skip it.
                        intPos2 = 0
                      Else
                        intPos2 = 0
                        If Right$(strLine, 1) = "_" Then
                          If strLineNum = vbNullString Then
                            ' ** 1st line of continuing MsgBox().
                            strLineNum = Trim$(Left$(strLine, InStr(strLine, " ")))
                            If Val(strLineNum) = 0 Then
                              strLineNum = "0"  ' ** Just in case there wasn't one, or something's goofy.
                              If strMsgBox = vbNullString Then
                                strMsgBox = Trim$(Left$(strLine, (Len(strLine) - 1)))
                              Else
                                strMsgBox = strMsgBox & " " & Trim$(Left$(strLine, (Len(strLine) - 1)))
                              End If
                            Else
                              If strMsgBox = vbNullString Then
                                strMsgBox = Trim$(Mid$(Left$(strLine, (Len(strLine) - 1)), InStr(strLine, " ")))
                              Else
                                strMsgBox = strMsgBox & " " & Trim$(Mid$(Left$(strLine, (Len(strLine) - 1)), InStr(strLine, " ")))
                              End If
                            End If
                          Else
                            ' ** Another continuing line.
                            strMsgBox = strMsgBox & " " & Trim$(Left$(strLine, (Len(strLine) - 1)))
                          End If
                          blnLineCont = True
                        Else
                          blnLineCont = False
                          ' ** 1-line MsgBox(), or last line of MsgBox().
                          If strMsgBox = vbNullString Then
                            ' ** 1-line MsgBox().
                            strLineNum = Left$(strLine, InStr(strLine, " "))
                            If Val(strLineNum) = 0 Then
                              strLineNum = "0"  ' ** Just in case there wasn't one, or something's goofy.
                              strMsgBox = strLine
                            Else
                              strMsgBox = Trim$(Mid$(strLine, InStr(strLine, " ")))
                            End If
                          Else
                            ' ** Last line of MsgBox().
                            strMsgBox = strMsgBox & " " & strLine
                          End If
If strLineNum = vbNullString Or IsNumeric(strLineNum) = False Or strLineNum = "0" Then
  Debug.Print "'MISSING LINENUM? " & strModName & "  " & strProcName & "()"
'Stop
End If
                          If strProcName = vbNullString Then strProcName = "{unknown}"
                          lngLineNum = CLng(strLineNum)
                          If IsNull(DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", _
                              "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngComID) & " And " & _
                              "[vbcomproc_name] = '" & strProcName & "'")) = False Then
                            lngComProcID = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", _
                              "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngComID) & " And " & _
                              "[vbcomproc_name] = '" & strProcName & "'")
                          Else: Stop
                          End If
If InStr(strMsgBox, "tblMsgBoxStyleType") = 0 Then
                          lngMsgs = lngMsgs + 1&
                          lngE = lngMsgs - 1&
                          ReDim Preserve arr_varMsg(M_ELEMS, lngE)
                          ' *****************************************************
                          ' ** Array: arr_varMsg()
                          ' **
                          ' **   Field  Element  Description         Constant
                          ' **   =====  =======  ==================  ==========
                          ' **     1       0     dbs_id              M_DID
                          ' **     2       1     vbcom_id            M_VID
                          ' **     3       2     vbcomproc_id        M_PID
                          ' **     4       3     vbcommsg_codeline   M_LIN
                          ' **     5       4     vbcommsg_raw        M_MSG
                          ' **     6       5     vbcommsg_errh       M_ERH
                          ' **
                          ' *****************************************************
                          arr_varMsg(M_DID, lngE) = lngThisDbsID
                          arr_varMsg(M_VID, lngE) = lngComID
                          arr_varMsg(M_PID, lngE) = lngComProcID
                          arr_varMsg(M_LIN, lngE) = lngLineNum
                          arr_varMsg(M_MSG, lngE) = strMsgBox
                          Select Case blnFound2
                          Case True
                            arr_varMsg(M_ERH, lngE) = lngX
                          Case False
                            arr_varMsg(M_ERH, lngE) = CLng(0)
                          End Select
                          lngComProcID = 0&: strLineNum = vbNullString: lngLineNum = 0&: strMsgBox = vbNullString
End If
                        End If
                      End If
                    End If
                  End If  ' ** Not remark: strLine.
                End If  ' ** Not blank: strLine.
End If  ' ** Qry_ChkParams.
              Next  ' * lngx.
            End If  ' ** lngDecLines.
          End With  ' ** cod
        End If  ' ** Not 'zz_mod_...".
      End With  ' ** vbc
    Next  ' * vbc
  End With  ' ** vbp.

'Debug.Print "'MSGS: " & CStr(lngMsgs)
  lngY = 0&
  For lngX = 0& To (lngMsgs - 1&)
    If arr_varMsg(M_LIN, lngX) = 0& Then
      Debug.Print "'" & arr_varMsg(M_MSG, lngX)
      DoEvents
      lngY = lngY + 1&
    End If
  Next
  Debug.Print "'NO LINE NUMS: " & CStr(lngY) & IIf(lngY > 0&, "!", vbNullString)
  If lngY > 0& Then
    Stop
    lngY = 0&
  End If

  Set dbs = CurrentDb
  With dbs

    ' ** Because searches to tblVBComponent_MessageBox using vbcommsg_codeline could very well be invalid,
    ' ** given that those code line numbers could change often, do it from scratch each time.

    ' ** Delete tblVBComponent_MessageBox, by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_01a")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    qdf.Execute dbFailOnError

    .Close
  End With

  ' ** Reset the Autonumber field.
  ChangeSeed_Ext "tblVBComponent_MessageBox"  ' ** Module Function: modAutonumberFieldFuncs.

  Set dbs = CurrentDb
  With dbs

    Set rst = dbs.OpenRecordset("tblVBComponent_MessageBox", dbOpenDynaset, dbConsistent)
    With rst

      For lngX = 0& To (lngMsgs - 1&)
        If arr_varMsg(M_LIN, lngX) > 0& Then
          blnAdd = False
          If .BOF = True And .EOF = True Then
            blnAdd = True
          Else
            ' *****************************************************
            ' ** Array: arr_varMsg()
            ' **
            ' **   Field  Element  Description         Constant
            ' **   =====  =======  ==================  ==========
            ' **     1       0     dbs_id              M_DID
            ' **     2       1     vbcom_id            M_VID
            ' **     3       2     vbcomproc_id        M_PID
            ' **     4       3     vbcommsg_codeline   M_LIN
            ' **     5       4     vbcommsg_raw        M_MSG
            ' **     6       5     vbcommsg_errh       M_ERH
            ' **
            ' *****************************************************
            .FindFirst "[dbs_id] = " & CStr(arr_varMsg(M_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varMsg(M_VID, lngX)) & " And " & _
              "[vbcomproc_id] = " & CStr(arr_varMsg(M_PID, lngX)) & " And [vbcommsg_codeline] = " & CStr(arr_varMsg(M_LIN, lngX))
            If .NoMatch = True Then
              blnAdd = True
            End If
          End If
          If blnAdd = True Then
            lngAdds = lngAdds + 1&
            .AddNew
            ![dbs_id] = arr_varMsg(M_DID, lngX)
            ![vbcom_id] = arr_varMsg(M_VID, lngX)
            ![vbcomproc_id] = arr_varMsg(M_PID, lngX)
            ![vbcommsg_codeline] = arr_varMsg(M_LIN, lngX)
            ![vbcommsg_errh] = arr_varMsg(M_ERH, lngX)
            '![vbcommsg_title] =
            '![vbcommsg_text] =
            '![vbcommsg_sw1] =
            '![mbtype_id1] =
            '![vbcommsg_sw2] =
            '![mbtype_id2] =
            '![vbcommsg_sw3] =
            '![mbtype_id3] =
            '![vbcommsg_sw4] =
            '![mbtype_id4] =
            ![vbcommsg_raw] = arr_varMsg(M_MSG, lngX)
            ![vbcommsg_datemodified] = Now()
            .Update
          Else
            If ![vbcommsg_raw] <> arr_varMsg(M_MSG, lngX) Then
              lngEdits = lngEdits + 1&
              .Edit
              ![vbcommsg_raw] = arr_varMsg(M_MSG, lngX)
              ![vbcommsg_datemodified] = Now()
              .Update
            End If
          End If
        End If
      Next
      .MoveFirst

      lngDels = 0&
      ReDim arr_varDel(0)
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          blnFound = False
          For lngY = 0& To (lngMsgs - 1&)
            If arr_varMsg(M_DID, lngY) = ![dbs_id] And arr_varMsg(M_VID, lngY) = ![vbcom_id] And _
                arr_varMsg(M_LIN, lngY) = ![vbcommsg_codeline] Then
              blnFound = True
              Exit For
            End If
          Next
          If blnFound = False Then
            lngDels = lngDels + 1&
            ReDim Preserve arr_varDel(lngDels - 1&)
            arr_varDel(lngDels - 1&) = ![vbcommsg_id]
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next

      .Close
    End With

    If lngDels > 0& Then
Debug.Print "'DELS: " & CStr(lngDels)
Stop
      For lngX = 0& To (lngDels - 1&)  ' ** Already qualified by lngThisDbsID.
        ' ** Delete tblVBComponent_MessageBox, by specified [msgid].
        Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_01b")
        With qdf.Parameters
          ![msgid] = arr_varDel(lngX)
        End With
        qdf.Execute dbFailOnError
      Next
    End If

    ' ** Update zz_qry_VBComponent_MsgBox_02 (tblVBComponent_MessageBox,
    ' ** with VBA_MsgBox_Parse(), by specified CurrentAppName()).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_03")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_04 (tblVBComponent_MessageBox, linked
    ' ** to tblmsgBoxType, with .._new fields, by specified CurrentAppName()).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_05")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_21 (zz_qry_VBComponent_MsgBox_20 (tblVBComponent_MessageBox, just
    ' ** vbcommsg_title within parens), with vbcommsg_title_new via VBA_MsgBox_Title(), by specified CurrentAppName()).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_22")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_24 (zz_qry_VBComponent_MsgBox_23 (tblVBComponent_MessageBox, just
    ' ** those with HasNum, by specifiedCurrentAppName()), with vbcommsg_title_new, vbcommsg_num_new).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_25")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_29 (zz_qry_VBComponent_MsgBox_28 (zz_qry_VBComponent_MsgBox_27
    ' ** (zz_qry_VBComponent_MsgBox_26 (tblVBComponent_MessageBox, with vbcommsg_text_new, sans 'MsgBox',
    ' ** by specified CurrentAppName()), with quote check), with vbcommsg_text_newx, and vbCrLf check),
    ' ** with vbcommsg_text_newy, via VBA_MsgBox_CrLf()).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_30")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_31 (tblVBComponent_MessageBox,
    ' ** just vbcommsg_title within quotes, by specified CurrentAppName()).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_32")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_33 (tblVBComponent_MessageBox,
    ' ** just vbcommsg_title with 1 quote, by specified CurrentAppName()).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_34")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_38 (zz_qry_VBComponent_MsgBox_37 (zz_qry_VBComponent_MsgBox_36
    ' ** (zz_qry_VBComponent_MsgBox_35 (tblVBComponent_MessageBox, just those with HasString, HasSpace in
    ' ** vbcommsg_title, by specified CurrentAppName()), with vbcommsg_title_new), with vbcommsg_title_newx),
    ' ** with vbcommsg_title_newy).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_39")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_42 (zz_qry_VBComponent_MsgBox_41 (zz_qry_VBComponent_MsgBox_40
    ' ** (tblVBComponent_MessageBox, just those with Space() in vbcommsg_raw, with switchtitle, by specified
    ' ** CurrentAppName()), just those with Space() in title), with vbcommsg_space_new).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_43")
    qdf.Execute

    ' ** Update zz_qry_VBComponent_MsgBox_45 (zz_qry_VBComponent_MsgBox_44 (tblVBComponent_MessageBox, just
    ' ** vbcommsg_space <> Null, with vbcommsg_space_new, by specified CurrentAppName()), with vbcommsg_space_newx).
    Set qdf = .QueryDefs("zz_qry_VBComponent_MsgBox_46")
    qdf.Execute

    .Close
  End With

  If lngAdds > 0& Or lngEdits > 0& Or lngDels > 0& Then
    Debug.Print "'MSGBOX ADDS: " & CStr(lngAdds) & "  EDITS: " & CStr(lngEdits) & "  DELS: " & CStr(lngDels)
  Else
    Debug.Print "'NO CHANGES!"
  End If

  ' ** VbMsgBoxButton enumeration:
  ' **         0  vbOKOnly               Display OK button only.
  ' **         0  vbDefaultButton1       First button is default.
  ' **         0  vbApplicationModal     Application modal; the user must respond to the message box before continuing work in the current application.
  ' **         1  vbOKCancel             Display OK and Cancel buttons.
  ' **         2  vbAbortRetryIgnore     Display Abort, Retry, and Ignore buttons.
  ' **         3  vbYesNoCancel          Display Yes, No, and Cancel buttons.
  ' **         4  vbYesNo                Display Yes and No buttons.
  ' **         5  vbRetryCancel          Display Retry and Cancel buttons.
  ' **        16  vbCritical             Display Critical Message icon.
  ' **        32  vbQuestion             Display Warning Query icon.
  ' **        48  vbExclamation          Display Warning Message icon.
  ' **        64  vbInformation          Display Information Message icon.
  ' **       256  vbDefaultButton2       Second button is default.
  ' **       512  vbDefaultButton3       Third button is default.
  ' **       768  vbDefaultButton4       Fourth button is default.
  ' **      4096  vbSystemModal          System modal; all applications are suspended until the user responds to the message box.
  ' **     16384  vbMsgBoxHelpButton     Adds Help button to the message box
  ' **     65536  vbMsgBoxSetForeground  Specifies the message box window as the foreground window
  ' **    524288  vbMsgBoxRight          Text is right aligned
  ' **   1048576  vbMsgBoxRtlReading     Specifies text should appear as right-to-left reading on Hebrew and Arabic systems

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_MsgBox_Doc = blnRetValx

End Function

Private Function VBA_Component_API_Doc() As Boolean
' ** Document all API calls in Trust Accountant to tblVBComponent_API.
' ** Called by:
' **   QuikVBADoc(), Above

  Const THIS_PROC As String = "VBA_Component_API_Doc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngAPIs As Long, arr_varAPI() As Variant
  Dim lngLines As Long, lngDecLines As Long
  Dim strLine As String
  Dim lngThisDbsID As Long
  Dim strModName As String, strProcName As String, strScope As String, strProcType As String, strReturnType As String
  Dim strAlias As String, strLibrary As String, strParams As String, blnMultiLine As Boolean
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varAPI().
  Const A_ELEMS As Integer = 12  ' ** Array's first-element UBound().
  Const A_DID   As Integer = 0
  Const A_DNAM  As Integer = 1
  Const A_VID   As Integer = 2
  Const A_VNAM  As Integer = 3
  Const A_ANAM  As Integer = 4
  Const A_PTYP  As Integer = 5   ' ** Procedure type: Sub or Function.
  Const A_SCOP  As Integer = 6   ' ** Public, Private, or none.
  Const A_PARS  As Integer = 7   ' ** Parameters.
  Const A_RTYP  As Integer = 8   ' ** Return type: Long, String, etc.
  Const A_MULT  As Integer = 9   ' ** Multi-line: True/False.
  Const A_LIB   As Integer = 10  ' ** Library name.
  Const A_ALIAS As Integer = 11  ' ** Alias.
  Const A_RAW   As Integer = 12

  Const DEC As String = "Declare "
  Const LC As String = "_"

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs
    ' ** Delete tblVBComponent_API, by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_VBComponent_API_01")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    qdf.Execute
    .Close
  End With

  lngAPIs = 0&
  ReDim arr_varAPI(A_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents  ' ** One-Based.
      With vbc
        strModName = .Name
        Set cod = .CodeModule
        With cod
          lngLines = .CountOfLines
          lngDecLines = .CountOfDeclarationLines
          For lngX = 1& To lngDecLines  ' ** One-Based.
            strLine = Trim$(.Lines(lngX, 1))
            strProcName = vbNullString: strProcType = vbNullString: strReturnType = vbNullString
            strScope = vbNullString: strLibrary = vbNullString: strAlias = vbNullString: strParams = vbNullString
            strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
            blnMultiLine = False
            If strLine <> vbNullString Then
              If Left$(strLine, 1) <> "'" Then
                intPos1 = InStr(strLine, DEC)
                If intPos1 > 0 Then
                  ' ** An API declaration.
                  If intPos1 > 1 Then
                    strTmp01 = Left$(strLine, (intPos1 - 1))  ' ** To the left of 'Declare', includes scope only.
                    strScope = Trim$(strTmp01)
                  Else
                    ' ** No scope!
                    strScope = "{none}"
                  End If
                  strTmp02 = Mid$(strLine, (intPos1 + Len(DEC)))  ' ** To the right of 'Declare', includes name, library,
                  If Right$(strTmp02, 1) = LC Then                ' ** alias (if any), params, and return type (if any).
                    ' ** Line continuation.
                    intPos1 = Len(strTmp02)
                    blnMultiLine = True
                    lngY = 0&
                    Do While intPos1 > 0
                      lngY = lngY + 1&
                      strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))  ' ** Strip off line-continuation character.
                      strTmp03 = Trim$(.Lines(lngX + lngY, 1))
                      strTmp02 = strTmp02 & strTmp03
                      strTmp03 = vbNullString: intPos1 = 0
                      If Right$(strTmp02, 1) = LC Then
                        intPos1 = Len(strTmp02)
                      End If
                    Loop
                  End If
                  intPos1 = InStr(strTmp02, " ")
                  strProcType = Trim$(Left$(strTmp02, (intPos1 - 1)))
                  strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))
                  strProcName = Left$(strTmp02, (InStr(strTmp02, " ") - 1))  ' ** Left of first space.
                  strTmp02 = Mid$(strTmp02, (InStr(strTmp02, " ") + 1))      ' ** Right of first space.
                  intPos1 = InStr(strTmp02, "(")
                  strParams = Mid$(strTmp02, intPos1)               ' ** Paren and right of paren.
                  strTmp02 = Trim$(Left$(strTmp02, (intPos1 - 1)))  ' ** Left of paren.
                  intPos1 = InStr(strParams, ")")
                  If intPos1 = Len(strParams) Then
                    ' ** No return type, so must be a Sub.
                  Else
                    strReturnType = Trim$(Mid$(strParams, (intPos1 + 1)))
                    intPos1 = InStr(strReturnType, "'")
                    If intPos1 > 0 Then strReturnType = Trim$(Left$(strReturnType, (intPos1 - 1)))
                    intPos1 = InStr(strReturnType, "As ")
                    If intPos1 > 0 Then
                      intPos1 = InStr(intPos1, strReturnType, ")")  ' ** Look for more, in case it's a parameter array.
                      If intPos1 > 0 Then
                        Do While intPos1 > 0
                          strReturnType = Trim$(Mid$(strReturnType, (intPos1 + 1)))
                          intPos1 = InStr(strReturnType, "As ")
                          intPos1 = InStr(intPos1, strReturnType, ")")
                        Loop
                      End If
                    End If
                  End If
                  intPos1 = 0
                  If Right$(strParams, 1) <> ")" Then
                    For lngY = Len(strParams) To 1 Step -1
                      If Mid$(strParams, lngY, 1) = ")" Then
                        intPos1 = lngY
                        Exit For
                      End If
                    Next
                  End If
                  If intPos1 > 0 Then strParams = Left$(strParams, intPos1)
                  intPos1 = InStr(strTmp02, "Lib ")  ' ** Should be 1.
                  strLibrary = Mid$(strTmp02, intPos1)
                  intPos1 = InStr(5, strLibrary, " ")
                  If intPos1 > 0 Then
                    strAlias = Trim$(Mid$(strLibrary, (intPos1 + 1)))
                    strLibrary = Trim$(Left$(strLibrary, (intPos1 - 1)))
                  Else
                    ' ** No Alias.
                  End If
                  lngAPIs = lngAPIs + 1&
                  lngE = lngAPIs - 1&
                  ReDim Preserve arr_varAPI(A_ELEMS, lngE)
                  ' ********************************************************
                  ' ** Array: arr_varAPI()
                  ' **
                  ' **   Field  Element  Name                   Constant
                  ' **   =====  =======  =====================  ==========
                  ' **     1       0     dbs_id                 A_DID
                  ' **     2       1     dbs_name               A_DNAM
                  ' **     3       2     vbcom_id               A_VID
                  ' **     4       3     vbcom_name             A_VNAM
                  ' **     5       4     vbcomapi_name          A_ANAM
                  ' **     6       5     proctype_type          A_PTYP
                  ' **     7       6     scopetype_type         A_SCOP
                  ' **     8       7     vbcomapi_parameters    A_PARS
                  ' **     9       8     vbcomapi_returntype    A_RTYP
                  ' **    10       9     vbcomapi_multiline     A_MULT
                  ' **    11      10     vbcomapi_library       A_LIB
                  ' **    12      11     vbcomapi_alias         A_ALIAS
                  ' **    13      12     vbcomapi_raw           A_RAW
                  ' **
                  ' ********************************************************
                  arr_varAPI(A_DID, lngE) = lngThisDbsID
                  arr_varAPI(A_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
                  arr_varAPI(A_VID, lngE) = CLng(0)
                  arr_varAPI(A_VNAM, lngE) = strModName
                  arr_varAPI(A_ANAM, lngE) = strProcName
                  arr_varAPI(A_PTYP, lngE) = strProcType
                  arr_varAPI(A_SCOP, lngE) = strScope
                  arr_varAPI(A_PARS, lngE) = strParams
                  arr_varAPI(A_RTYP, lngE) = strReturnType
                  arr_varAPI(A_MULT, lngE) = blnMultiLine
                  arr_varAPI(A_LIB, lngE) = strLibrary
                  arr_varAPI(A_ALIAS, lngE) = strAlias
                  arr_varAPI(A_RAW, lngE) = strLine
                End If
              End If  ' ** Remark.
            End If  ' ** vbNullString.
          Next  ' ** For each line: lngX.
        End With ' ** Code module: cod.
      End With  ' ** This component: vbc.
    Next  ' ** For each component: vbc.
  End With  ' ** vbp.

  If lngAPIs > 0& Then

    Set dbs = CurrentDb
    With dbs

      Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      With rst
        For lngX = 0& To (lngAPIs - 1&)
          .FindFirst "[dbs_id] = " & CStr(arr_varAPI(A_DID, lngX)) & " And [vbcom_name] = '" & arr_varAPI(A_VNAM, lngX) & "'"
          If .NoMatch = False Then
           arr_varAPI(A_VID, lngX) = ![vbcom_id]
          Else
            Debug.Print "'NOT FOUND! " & arr_varAPI(A_VNAM, lngX)
          End If
        Next
        .Close
      End With

      Set rst = .OpenRecordset("tblVBComponent_API", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 1& To 2&
          For lngY = 0& To (lngAPIs - 1&)
            If (lngX = 1& And InStr(arr_varAPI(A_VNAM, lngY), "zz_") = 0) Or (lngX = 2& And InStr(arr_varAPI(A_VNAM, lngY), "zz_") > 0) Then
              .AddNew
              ![dbs_id] = arr_varAPI(A_DID, lngY)
              ![vbcom_id] = arr_varAPI(A_VID, lngY)
              '![vbcom_name] = arr_varAPI(A_VNAM, lngY)
              If arr_varAPI(A_ANAM, lngY) <> vbNullString Then
                ![vbcomapi_name] = arr_varAPI(A_ANAM, lngY)
              Else
                ![vbcomapi_name] = "{unknown" & CStr(lngY) & "}"
              End If
              If arr_varAPI(A_PTYP, lngY) <> vbNullString Then
                ![proctype_type] = arr_varAPI(A_PTYP, lngY)
              Else
                ![proctype_type] = "{UNK" & CStr(lngY) & "}"
              End If
              ![scopetype_type] = arr_varAPI(A_SCOP, lngY)
              If arr_varAPI(A_PARS, lngY) <> vbNullString Then
                ![vbcomapi_parameters] = arr_varAPI(A_PARS, lngY)
              End If
              If arr_varAPI(A_RTYP, lngY) <> vbNullString Then
                intPos1 = InStr(arr_varAPI(A_RTYP, lngY), "'")  ' ** In case there's a remark after it.
                If intPos1 > 0 Then arr_varAPI(A_RTYP, lngY) = Trim$(Left$(arr_varAPI(A_RTYP, lngY), (intPos1 - 1)))
                '![vbcomapi_returntype] = arr_varAPI(A_RTYP, lngY)
                intPos1 = InStr(arr_varAPI(A_RTYP, lngY), "As ")
                If intPos1 > 0 Then
                  strTmp01 = Mid$(arr_varAPI(A_RTYP, lngY), (intPos1 + 3))
                Else
                  strTmp01 = arr_varAPI(A_RTYP, lngY)
                End If
                strTmp01 = "vb" & strTmp01
                varTmp00 = DLookup("[datatype_vb_type]", "tblDataTypeVb", "[datatype_vb_constant] = '" & strTmp01 & "'")
                Select Case IsNull(varTmp00)
                Case True
                  ![datatype_vb_type] = vbUndeclared
                Case False
                  ![datatype_vb_type] = varTmp00
                End Select
              Else
                ![datatype_vb_type] = vbUndeclared
              End If
              ![vbcomapi_multiline] = arr_varAPI(A_MULT, lngY)
              If arr_varAPI(A_LIB, lngY) <> vbNullString Then
                strTmp01 = arr_varAPI(A_LIB, lngY)
                intPos1 = InStr(strTmp01, " ")
                strTmp02 = Trim$(Mid$(strTmp01, (intPos1 + 1)))
                ![vbcomapi_library] = strTmp02
              End If
              If arr_varAPI(A_ALIAS, lngY) <> vbNullString Then
                strTmp01 = arr_varAPI(A_ALIAS, lngY)
                intPos1 = InStr(strTmp01, " ")
                strTmp02 = Trim$(Mid$(strTmp01, (intPos1 + 1)))
                ![vbcomapi_alias] = strTmp02
              End If
              ![vbcomapi_raw] = arr_varAPI(A_RAW, lngY)
              ![vbcomapi_datemodified] = Now()
              .Update
            End If
          Next  ' ** lngY.
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.

      .Close
    End With  ' ** dbs.

  End If

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'APIS: " & CStr(lngAPIs)
  'APIS: 203

  Debug.Print "'DONE! " & THIS_PROC & "()"

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_Component_API_Doc = blnRetVal

End Function

Public Function VBA_WinDialog_Doc() As Boolean
' ** Document usage of browse dialog Functions and Subs to tblVBComponent_Procedure_Detail.
' ** Called by:
' **

  Const THIS_PROC As String = "VBA_WinDialog_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngComps As Long, arr_varComp As Variant
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngHits As Long, arr_varHit() As Variant, lngHitXs As Long, arr_varHitX() As Variant
  Dim lngLine As Long, lngLastLine As Long, lngLines As Long, lngDecLines As Long, lngColumn As Long
  Dim strModName As String, strProcName As String, strLine As String
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim lngDbsID As Long, lngVBComID As Long, lngVBComProcID As Long, lngVBComps As Long, lngYStart As Long, lngYLoops As Long
  Dim blnFound As Boolean, blnDoc As Boolean, blnNoParens As Boolean
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngPos1 As Long, lngPos2 As Long, lngLen As Long
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
  Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varComp().
  Const C_DID   As Integer = 0
  Const C_DNAM  As Integer = 1
  Const C_VID   As Integer = 2
  Const C_VNAM  As Integer = 3
  Const C_LINES As Integer = 4

  ' ** Array: arr_varProc().
  Const P_DID  As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_VID  As Integer = 2
  Const P_VNAM As Integer = 3
  Const P_PID  As Integer = 4
  Const P_PNAM As Integer = 5
  Const P_SCOP As Integer = 6
  Const P_PTYP As Integer = 7
  Const P_RET  As Integer = 8
  Const P_RTYP As Integer = 9
  Const P_BEG  As Integer = 10
  Const P_END  As Integer = 11
  Const P_PCNT As Integer = 12

  ' ** Array: arr_varHit().
  Const H_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const H_DID   As Integer = 0
  Const H_VID   As Integer = 1
  Const H_PID   As Integer = 2
  Const H_LIN   As Integer = 3
  Const H_TXT   As Integer = 4
  Const H_SDID  As Integer = 5  ' ** Source dbs_id (where source proc is located).
  Const H_SCID  As Integer = 6  ' ** Source vbcom_id (where source proc is located).
  Const H_SPID  As Integer = 7  ' ** Source vbcomproc_id (what was found).

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    ' ** tblVBComponent, just needed fields, by CurrentAppName().
    Set qdf = .QueryDefs("zz_qry_VBComponent_WinDialog_04")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngComps = .RecordCount
      .MoveFirst
      arr_varComp = .GetRows(lngComps)
      ' ************************************************
      ' ** Array: arr_varComp()
      ' **
      ' **   Field  Element  Name           Constant
      ' **   =====  =======  =============  ==========
      ' **     1       0     dbs_id         C_DID
      ' **     2       1     dbs_name       C_DNAM
      ' **     3       2     vbcom_id       C_VID
      ' **     4       3     vbcom_name     C_VNAM
      ' **     5       4     vbcom_lines    C_LINES
      ' **
      ' ************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    lngHitXs = 0&
    ReDim arr_varHitX(H_ELEMS, 0)

    For lngW = 1& To 2&

' ** The parents in tblVBComponent_Procedure_Detail are the
' ** 3 Windows Dialog functions in modBrowseFilesAndFolders.
' ** The children in tblVBComponent_Procedure_Detail are the
' ** procedures that call the parent.
' ** The children should also have some children of their own,
' ** that is, they should also appear in the table as parents,
' ** with the various procedures that call them as detail.

      Select Case lngW
      Case 1&
        ' ** tblVBComponent_Procedure, just browse dialog functions, by CurrentAppName().
        Set qdf = .QueryDefs("zz_qry_VBComponent_WinDialog_02")  ' ** These are the 3 in modBrowseFilesAndFolders.
      Case 2&
        ' ** tblVBComponent_Procedure, linked to zz_qry_VBComponent_WinDialog_07a (zz_qry_VBComponent_WinDialog_06
        ' ** (zz_qry_VBComponent_WinDialog_05 (tblVBComponent_Procedure_Detail, linked to zz_qry_VBComponent_WinDialog_02
        ' ** (tblVBComponent_Procedure, just browse dialog functions, by specified CurrentAppName()), current browse
        ' ** function usage), just FileSaveDialog()), grouped by vbprocdet_name), just FileSaveDialog() calling procedures.
        Set qdf = .QueryDefs("zz_qry_VBComponent_WinDialog_08")
      End Select

      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngProcs = .RecordCount
        .MoveFirst
        arr_varProc = .GetRows(lngProcs)
        ' *********************************************************
        ' ** Array: arr_varProc()
        ' **
        ' **   Field  Element  Name                    Constant
        ' **   =====  =======  ======================  ==========
        ' **     1       0     dbs_id                  P_DID
        ' **     2       1     dbs_name                P_DNAM
        ' **     3       2     vbcom_id                P_VID
        ' **     4       3     vbcom_name              P_VNAM
        ' **     5       4     vbcomproc_id            P_PID
        ' **     6       5     vbcomproc_name          P_PNAM
        ' **     7       6     scopetype_type          P_SCOP
        ' **     8       7     proctype_type           P_PTYP
        ' **     9       8     vbcomproc_returntype    P_RET
        ' **    10       9     datatype_vb_type        P_RTYP
        ' **    11      10     vbcomproc_line_beg      P_BEG
        ' **    12      11     vbcomproc_line_end      P_END
        ' **    13      12     vbcomproc_params        P_PCNT
        ' **
        ' ********************************************************
        .Close
      End With
      Set rst = Nothing
      Set qdf = Nothing

      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      DoEvents

      Debug.Print "'TRACING " & CStr(lngProcs) & " FUNCTIONS:"
      DoEvents

      lngHits = 0&
      ReDim arr_varHit(H_ELEMS, 0)

      Set vbp = Application.VBE.ActiveVBProject
      With vbp

        strModName = vbNullString
        lngVBComps = .VBComponents.Count

        For lngX = 0& To (lngProcs - 1&)
          lngYStart = 1&: lngYLoops = 0&

          Select Case arr_varProc(P_SCOP, lngX)
          Case "Public"
            lngYLoops = lngVBComps
          Case "Private"
            lngYLoops = 1&
          End Select

          For lngY = 1& To lngYLoops

            blnFound = False
            Select Case arr_varProc(P_SCOP, lngX)
            Case "Public"
              For lngZ = lngYStart To (lngVBComps - 1&) 'Each vbc In .VBComponents
If InStr(.VBComponents(lngZ).Name, "error") > 0& Or InStr(strModName, "error") > 0 Then
'Stop
End If
                If .VBComponents(lngZ).Name <> strModName Then
                  blnFound = True
                  Set vbc = .VBComponents(lngZ)
                  lngYStart = lngZ
                  Exit For
                End If
              Next  ' ** vbc.
            Case "Private"
              blnFound = True
              Set vbc = .VBComponents(arr_varProc(P_VNAM, lngX))
            End Select

            If blnFound = True Then

              lngDbsID = 0&: lngVBComID = 0&: lngVBComProcID = 0&
              With vbc

                strModName = .Name
                For lngZ = 0& To (lngComps - 1&)
                  If arr_varComp(C_VNAM, lngZ) = strModName Then
                    lngVBComID = arr_varComp(C_VID, lngZ)
                    lngDbsID = arr_varComp(C_DID, lngZ)
                    Exit For
                  End If
                Next  ' ** lngZ.

                Set cod = .CodeModule
                With cod
                  lngLines = .CountOfLines
                  lngDecLines = .CountOfDeclarationLines
                  lngLine = (lngDecLines + 1&)
                  lngColumn = 1&
                  blnDoc = False
                  blnFound = .Find(arr_varProc(P_PNAM, lngX), lngLine, lngColumn, lngLines, -1, True, True, False)
                  ' ** object.Find(target, startline, startcol, endline, endcol [, wholeword] [, matchcase] [, patternsearch]) As Boolean

                  lngLines = .CountOfLines  ' ** The Find resets this to the found line!
                  If blnFound = True Then
                    Do While blnFound = True
                      strProcName = .ProcOfLine(lngLine, vbext_pk_Proc)
                      If strProcName <> "VBA_Chk_Events" Then  ' ** It'll find all sorts of things here!
                        If (strModName <> arr_varProc(P_VNAM, lngX)) Or (strModName = arr_varProc(P_VNAM, lngX) And _
                            (lngLine < arr_varProc(P_BEG, lngX) Or lngLine > arr_varProc(P_END, lngX))) Then  ' ** Make sure it doesn't find itself.
                          strLine = .Lines(lngLine, 1)
                          ' ** Check to make sure it's really a match, with no underscores to the left or right.
                          lngLen = Len(strLine)
                          lngPos1 = InStr(strLine, arr_varProc(P_PNAM, lngX))
                          If lngPos1 > 0 Then
                            lngPos2 = InStr(strLine, "'")
                            If lngPos2 > 0 And lngPos2 < lngPos1 Then
                              ' ** It's in a remark.
                            Else
                              strTmp01 = Mid$(strLine, (lngPos1 - 1&), 1)
                              Select Case strTmp01
                              Case " ", "("
                                ' ** OK.
                                blnDoc = True
                              Case Else
                                Debug.Print "'  " & strTmp01
                                Stop
                              End Select
                              strTmp01 = Mid$(strLine, (lngPos1 + Len(arr_varProc(P_PNAM, lngX))), 1)
                              Select Case strTmp01
                              Case " ", "(", ")"
                                ' ** OK.
                                blnDoc = True
                              Case Else
                                Debug.Print "'  " & strTmp01
                                Stop
                                blnDoc = False
                              End Select
                            End If  ' ** lngPos2.
                          End If  ' ** lngPos1.
                        End If  ' ** Not itself.
                      End If  ' ** Not "VBA_Chk_Events".

                      If blnDoc = True Then
                        strProcName = .ProcOfLine(lngLine, vbext_pk_Proc)
                        varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", "[vbcom_id] = " & CStr(lngVBComID) & " And " & _
                          "[vbcomproc_name] = '" & strProcName & "'")
                        If IsNull(varTmp00) = False Then
                          lngVBComProcID = varTmp00
                        End If
                        lngHits = lngHits + 1&
                        lngE = lngHits - 1&
                        ReDim Preserve arr_varHit(H_ELEMS, lngE)
                        ' ******************************************************
                        ' ** Array: arr_varHit()
                        ' **
                        ' **   Field  Element  Name                 Constant
                        ' **   =====  =======  ===================  ==========
                        ' **     1       0     dbs_id_det           H_DID
                        ' **     2       1     vbcom_id_det         H_VID
                        ' **     3       2     vbcomproc_id_det     H_PID
                        ' **     4       3     vbprocdet_linenum    H_LIN
                        ' **     5       4     vbprocdet_raw        H_TXT
                        ' **     6       5     dbs_id               H_SDID
                        ' **     7       6     vbcom_id             H_SCID
                        ' **     8       7     vbcomproc_id         H_SPID
                        ' **
                        ' ******************************************************
                        arr_varHit(H_DID, lngE) = lngThisDbsID
                        arr_varHit(H_VID, lngE) = lngVBComID
                        arr_varHit(H_PID, lngE) = lngVBComProcID
                        arr_varHit(H_LIN, lngE) = lngLine
                        arr_varHit(H_TXT, lngE) = strLine
                        arr_varHit(H_SDID, lngE) = arr_varProc(P_DID, lngX)
                        arr_varHit(H_SCID, lngE) = arr_varProc(P_VID, lngX)
                        arr_varHit(H_SPID, lngE) = arr_varProc(P_PID, lngX)
                      End If

                      lngLines = .CountOfLines  ' ** The Find resets this to the found line!
                      lngLine = lngLine + 1&
                      lngColumn = 1&
                      blnDoc = False
                      If lngLine > lngLines Then
                        blnFound = False
                      Else
                        blnFound = .Find(arr_varProc(P_PNAM, lngX), lngLine, lngColumn, lngLines, -1, True, True, False)
                        lngLines = .CountOfLines  ' ** The Find resets this to the found line!
                      End If
                      If blnFound = False Then
                        Exit Do
                      End If

                    Loop  ' ** blnFound.
                  End If  ' ** blnFound.

                  lngLines = .CountOfLines  ' ** The Find resets this to the found line!
                  lngLine = (lngDecLines + 1&)
                  lngColumn = 1&
                End With  ' ** cod.
                Set cod = Nothing

              End With  ' ** vbc.
              Set vbc = Nothing

            End If  ' ** blnFound.

          Next  ' ** lngY.

        Next  ' ** lngX.

      End With  ' ** vbp.
      Set vbp = Nothing

      ' ******************************************************
      ' ** Array: arr_varHit()
      ' **
      ' **   Field  Element  Name                 Constant
      ' **   =====  =======  ===================  ==========
      ' **     1       0     dbs_id_det           H_DID
      ' **     2       1     vbcom_id_det         H_VID
      ' **     3       2     vbcomproc_id_det     H_PID
      ' **     4       3     vbprocdet_linenum    H_LIN
      ' **     5       4     vbprocdet_raw        H_TXT
      ' **     6       5     dbs_id               H_SDID
      ' **     7       6     vbcom_id             H_SCID
      ' **     8       7     vbcomproc_id         H_SPID
      ' **
      ' ******************************************************

      If lngHits > 0& Then

        Set rst = .OpenRecordset("tblVBComponent_Procedure_Detail", dbOpenDynaset, dbConsistent)
        With rst
          For lngX = 0& To (lngHits - 1&)
            .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_SDID, lngX)) & " And " & "[vbcom_id] = " & CStr(arr_varHit(H_SCID, lngX)) & " And " & _
              "[vbcomproc_id] = " & CStr(arr_varHit(H_SPID, lngX)) & " And " & _
              "[dbs_id_det] = " & CStr(arr_varHit(H_DID, lngX)) & " And " & "[vbcom_id_det] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
              "[vbcomproc_id_det] = " & CStr(arr_varHit(H_PID, lngX)) & " And " & _
              "[vbprocdet_linenum] = " & CStr(arr_varHit(H_LIN, lngX))
            Select Case .NoMatch
            Case True
              .AddNew
              ![dbs_id] = arr_varHit(H_SDID, lngX)
              ![vbcom_id] = arr_varHit(H_SCID, lngX)
              ![vbcomproc_id] = arr_varHit(H_SPID, lngX)
              ' ** ![vbprocdet_id] : AutoNumber.
              ![dbs_id_det] = arr_varHit(H_DID, lngX)
              ![vbcom_id_det] = arr_varHit(H_VID, lngX)
              ![vbcomproc_id_det] = arr_varHit(H_PID, lngX)
              ![vbprocdet_linenum] = arr_varHit(H_LIN, lngX)
            Case False
              .Edit
            End Select
            strTmp01 = arr_varHit(H_TXT, lngX)
            lngPos2 = InStr(strTmp01, "' ** ")
            If lngPos2 > 0 Then
              strTmp01 = Trim$(Left$(strTmp01, (lngPos2 - 1&)))
            End If
            lngPos2 = InStr(strTmp01, " ")
            If IsNumeric(Left$(strTmp01, (lngPos2 - 1))) = True Then
              strTmp01 = Trim$(Mid$(strTmp01, lngPos2))
            End If
            lngPos2 = InStr(strTmp01, "=")
            If lngPos2 > 0& Then
              ![vbprocdet_assign] = Trim$(Left$(strTmp01, (lngPos2 - 1&)))
              strTmp01 = Trim$(Mid$(strTmp01, (lngPos2 + 1&)))
              If InStr(strTmp01, "(") > 0 And Right$(strTmp01, 1) = ")" Then
                blnNoParens = False
              Else
                blnNoParens = True
              End If
            Else
              blnNoParens = True
              ![vbprocdet_assign] = Null
            End If
            For lngZ = 0& To (lngProcs - 1&)
              If arr_varProc(P_PID, lngZ) = arr_varHit(H_SPID, lngX) Then
                Select Case arr_varProc(P_PCNT, lngZ)
                Case 1&
                  Select Case blnNoParens
                  Case True
                    lngPos2 = InStr(strTmp01, " ")
                    If lngPos2 > 0& Then
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      ![vbprocdet_param1] = strTmp02
                    Else
                      ![vbprocdet_param1] = Null
                    End If
                  Case False
                    lngPos2 = InStr(strTmp01, "(")
                    If Mid$(strTmp01, (lngPos2 + 1&), 1&) <> ")" Then  ' ** I'm not going to check whether these are Optional or not.
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                      ![vbprocdet_param1] = strTmp02
                    End If
                  End Select  ' ** blnNoParens.
                  ![vbprocdet_param2] = Null
                  ![vbprocdet_param3] = Null
                  ![vbprocdet_param4] = Null
                Case 2&
                  Select Case blnNoParens
                  Case True
                    lngPos2 = InStr(strTmp01, " ")
                    If lngPos2 > 0& Then
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      lngPos2 = InStr(strTmp02, ",")
                      If lngPos2 > 0& Then
                        ![vbprocdet_param1] = Left$(strTmp02, (lngPos2 - 1&))
                        ![vbprocdet_param2] = Mid$(strTmp02, (lngPos2 + 1&))
                      Else
                        ![vbprocdet_param1] = strTmp02
                        ![vbprocdet_param2] = Null
                      End If
                    Else
                      ![vbprocdet_param1] = Null
                      ![vbprocdet_param2] = Null
                    End If
                    ![vbprocdet_param3] = Null
                    ![vbprocdet_param4] = Null
                  Case False
                    lngPos2 = InStr(strTmp01, "(")
                    If Mid$(strTmp01, (lngPos2 + 1&), 1&) <> ")" Then
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      lngPos2 = InStr(strTmp02, ",")
                      If lngPos2 > 0& Then
                        ![vbprocdet_param1] = Left$(strTmp02, (lngPos2 - 1&))
                        strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                        If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                        ![vbprocdet_param2] = strTmp02
                        ![vbprocdet_param3] = Null
                        ![vbprocdet_param4] = Null
                      Else
                        If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                        ![vbprocdet_param1] = strTmp02
                        ![vbprocdet_param2] = Null
                        ![vbprocdet_param3] = Null
                        ![vbprocdet_param4] = Null
                      End If
                    End If
                  End Select  ' ** blnNoParens.
                Case 3&
                  Select Case blnNoParens
                  Case True
                    lngPos2 = InStr(strTmp01, " ")
                    If lngPos2 > 0& Then
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      lngPos2 = InStr(strTmp02, ",")
                      If lngPos2 > 0& Then
                        ![vbprocdet_param1] = Left$(strTmp02, (lngPos2 - 1&))
                        strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                        lngPos2 = InStr(strTmp02, ",")
                        If lngPos2 > 0& Then
                          ![vbprocdet_param2] = Left$(strTmp02, (lngPos2 - 1&))
                          ![vbprocdet_param3] = Mid$(strTmp02, (lngPos2 + 1&))
                        Else
                          ![vbprocdet_param2] = strTmp02
                          ![vbprocdet_param3] = Null
                        End If
                      Else
                        ![vbprocdet_param1] = strTmp02
                        ![vbprocdet_param2] = Null
                        ![vbprocdet_param3] = Null
                      End If
                    Else
                      ![vbprocdet_param1] = Null
                      ![vbprocdet_param2] = Null
                      ![vbprocdet_param3] = Null
                    End If
                    ![vbprocdet_param4] = Null
                  Case False
                    lngPos2 = InStr(strTmp01, "(")
                    If Mid$(strTmp01, (lngPos2 + 1&), 1&) <> ")" Then
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      lngPos2 = InStr(strTmp02, ",")
                      If lngPos2 > 0& Then
                        ![vbprocdet_param1] = Left$(strTmp02, (lngPos2 - 1&))
                        strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                        lngPos2 = InStr(strTmp02, ",")
                        If lngPos2 > 0& Then
                          ![vbprocdet_param2] = Left$(strTmp02, (lngPos2 - 1&))
                          strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                          If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                          ![vbprocdet_param3] = strTmp02
                          ![vbprocdet_param4] = Null
                        Else
                          ![vbprocdet_param2] = strTmp02
                          ![vbprocdet_param3] = Null
                          ![vbprocdet_param4] = Null
                        End If
                      Else
                        If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                        ![vbprocdet_param1] = strTmp02
                        ![vbprocdet_param2] = Null
                        ![vbprocdet_param3] = Null
                        ![vbprocdet_param4] = Null
                      End If
                    End If
                  End Select  ' ** blnNoParens.
                Case 4&
                  Select Case blnNoParens
                  Case True
                    lngPos2 = InStr(strTmp01, " ")
                    If lngPos2 > 0& Then
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      lngPos2 = InStr(strTmp02, ",")
                      If lngPos2 > 0& Then
                        ![vbprocdet_param1] = Left$(strTmp02, (lngPos2 - 1&))
                        strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                        lngPos2 = InStr(strTmp02, ",")
                        If lngPos2 > 0& Then
                          ![vbprocdet_param2] = Left$(strTmp02, (lngPos2 - 1&))
                          strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                          lngPos2 = InStr(strTmp02, ",")
                          If lngPos2 > 0& Then
                            ![vbprocdet_param3] = Left$(strTmp02, (lngPos2 - 1&))
                            ![vbprocdet_param4] = Mid$(strTmp02, (lngPos2 + 1&))
                          Else
                            ![vbprocdet_param3] = strTmp02
                            ![vbprocdet_param4] = Null
                          End If
                        Else
                          ![vbprocdet_param2] = strTmp02
                          ![vbprocdet_param3] = Null
                          ![vbprocdet_param4] = Null
                        End If
                      Else
                        ![vbprocdet_param1] = strTmp02
                        ![vbprocdet_param2] = Null
                        ![vbprocdet_param3] = Null
                        ![vbprocdet_param4] = Null
                      End If
                    Else
                      ![vbprocdet_param1] = Null
                      ![vbprocdet_param2] = Null
                      ![vbprocdet_param3] = Null
                      ![vbprocdet_param4] = Null
                    End If
                  Case False
                    lngPos2 = InStr(strTmp01, "(")
                    If Mid$(strTmp01, (lngPos2 + 1&), 1&) <> ")" Then
                      strTmp02 = Mid$(strTmp01, (lngPos2 + 1&))
                      lngPos2 = InStr(strTmp02, ",")
                      If lngPos2 > 0& Then
                        ![vbprocdet_param1] = Left$(strTmp02, (lngPos2 - 1&))
                        strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                        lngPos2 = InStr(strTmp02, ",")
                        If lngPos2 > 0& Then
                          ![vbprocdet_param2] = Left$(strTmp02, (lngPos2 - 1&))
                          strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                          lngPos2 = InStr(strTmp02, ",")
                          If lngPos2 > 0& Then
                            ![vbprocdet_param3] = Left$(strTmp02, (lngPos2 - 1&))
                            strTmp02 = Mid$(strTmp02, (lngPos2 + 1&))
                            If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                            ![vbprocdet_param4] = strTmp02
                          Else
                            If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                            ![vbprocdet_param3] = strTmp02
                            ![vbprocdet_param4] = Null
                          End If
                        Else
                          If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                          ![vbprocdet_param2] = strTmp02
                          ![vbprocdet_param3] = Null
                          ![vbprocdet_param4] = Null
                        End If
                      Else
                        If Right$(strTmp02, 1) = ")" Then strTmp02 = Left$(strTmp02, (Len(strTmp02) - 1))
                        ![vbprocdet_param1] = strTmp02
                        ![vbprocdet_param2] = Null
                        ![vbprocdet_param3] = Null
                        ![vbprocdet_param4] = Null
                      End If
                    End If
                  End Select  ' ** blnNoParens.
                End Select  ' ** P_PCNT.
                Exit For
              End If  ' ** P_PID, H_SPID.
            Next  ' ** lngZ.
            Select Case blnNoParens
            Case True
              lngPos2 = InStr(strTmp01, " ")
              If lngPos2 > 0 Then
                strTmp01 = Left$(strTmp01, (lngPos2 - 1&)) & "()"
              Else
                strTmp01 = strTmp01 & "()"
              End If
            Case False
              lngPos2 = InStr(strTmp01, "(")
              strTmp01 = Left$(strTmp01, lngPos2) & Right$(strTmp01, 1)
            End Select
            ![vbprocdet_proc] = strTmp01
            ![vbprocdet_raw] = arr_varHit(H_TXT, lngX)
            ![vbprocdet_datemodified] = Now()
            .Update
          Next  ' ** lngX.
          .Close
        End With  ' ** rst.
        Set rst = Nothing

        For lngX = 0& To (lngHits - 1&)
          lngHitXs = lngHitXs + 1&
          ReDim Preserve arr_varHitX(H_ELEMS, lngHitXs - 1&)
          For lngY = 0& To H_ELEMS
            arr_varHitX(lngY, lngX) = arr_varHit(lngY, lngX)
          Next
        Next

      End If  ' ** lngHits.

    Next  ' ** lngW.

    Debug.Print "'lngHitXs = " & CStr(lngHitXs)

    lngDels = 0&
    ReDim arr_varDel(0)

    Set rst = .OpenRecordset("tblVBComponent_Procedure_Detail", dbOpenDynaset, dbConsistent)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          blnFound = False
          For lngY = 0& To (lngHitXs - 1&)
            If arr_varHitX(H_DID, lngY) = ![dbs_id] And arr_varHitX(H_SCID, lngY) = ![vbcom_id] And _
                arr_varHitX(H_SPID, lngY) = ![vbcomproc_id] And arr_varHitX(H_VID, lngY) = ![vbcom_id_det] And _
                arr_varHitX(H_PID, lngY) = ![vbcomproc_id_det] And arr_varHitX(H_LIN, lngY) = ![vbprocdet_linenum] Then
              blnFound = True
              Exit For
            End If
          Next  ' ** lngY.
          If blnFound = False Then
            lngDels = lngDels + 1&
            ReDim Preserve arr_varDel(lngDels - 1&)
            arr_varDel(lngDels - 1&) = ![vbprocdet_id]
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next  ' ** lngX.
      .Close
    End With

    If lngDels > 0& Then
Debug.Print "'DELS: " & CStr(lngDels)
'lngDels = 0&
Stop
      For lngX = 0& To (lngDels - 1&)
        ' ** Delete tblVBComponent_Procedure_Detail, by specified [pdetid].
        Set qdf = .QueryDefs("zz_qry_VBComponent_WinDialog_01")
        With qdf.Parameters
          ![pdetid] = arr_varDel(lngX)
        End With
        qdf.Execute
      Next
    End If

    .Close
  End With  ' ** dbs.

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_WinDialog_Doc = blnRetVal

End Function

Public Function VBA_PublicVar_Doc() As Boolean
' ** Document all module level variables, constants, declares, types, enums,
' ** deftypes, and #Const to tblVBComponent_Declaration.
' ** Called by:
' **

  Const THIS_PROC As String = "VBA_PublicVar_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngFams As Long, arr_varFam As Variant
  Dim lngUniques As Long, arr_varUnique As Variant
  Dim lngAPIs As Long, arr_varAPI As Variant
  Dim lngLines As Long, lngDecLines As Long
  Dim lngThisDbsID As Long
  Dim strModName As String, strLine As String, strScope As String, strFamily As String
  Dim blnLineContinue As Boolean, blnDefType As Boolean, blnCompiler As Boolean, blnFound As Boolean, blnAdd As Boolean
  Dim blnBlockType As Boolean, blnBlockEnum As Boolean, lngBlockTypes As Long, lngBlockEnums As Long, lngDefTypes As Long
  Dim lngBlockStart As Long, strBlockName As String, strBlockType As String, strBlockMembers As String, strBlockScope As String
  Dim strVarName As String, strVarType As String, blnVarLoop As Boolean, blnArray As Boolean
  Dim strDeclareName As String, strDeclareType As String, lngDeclareID As Long, lngAPIID As Long
  Dim strConstName As String, strConstType As String, lngConstID As Long
  Dim blnNoScope As Boolean, blnChange As Boolean, blnOptCompare As Boolean, blnOptExplicit As Boolean
  Dim lngNoScopes As Long, arr_varNoScope() As Variant
  Dim lngChanges As Long, arr_varChange() As Variant
  Dim lngModCnt As Long, lngModTot As Long, lngVBComID As Long, lngDataTypeVb As Long
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Scope declarations.
  Const PPRIV As String = "Private"
  Const PPUB  As String = "Public"
  Const PSTAT As String = "Static"
  Const PDIM  As String = "Dim"
  Const PGLOB As String = "Global"  ' ** If any are found, change to Public.
  Const POPT  As String = "Option"
  Const PCOMP As String = "#Const"  ' ** Compiler declaration.
  Const PCIF  As String = "#If"     ' ** Compiler If/Then/Else
  Const PCELS As String = "#Else"   ' ** Compiler If/Then/Else
  Const PCEND As String = "#End"    ' ** Compiler If/Then/Else

  ' ** 2nd words, that should have scope.
  Const P2CONST As String = "Const"
  Const P2TYPE  As String = "Type"
  Const P2ENUM  As String = "Enum"
  Const P2DEC   As String = "Declare"

  ' ** Block closers.
  Const PENDT As String = "End Type"
  Const PENDE As String = "End Enum"

  ' ** Array: arr_varNoScope().
  Const NS_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const NS_MOD As Integer = 0
  Const NS_LIN As Integer = 1
  Const NS_TXT As Integer = 2

  ' ** Array: arr_varChange().
  Const CH_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const CH_MOD As Integer = 0
  Const CH_LIN As Integer = 1
  Const CH_TXT As Integer = 2

  ' ** Array: arr_varFam().
  Const F_FID As Integer = 0
  Const F_FAM As Integer = 1
  Const F_PFX As Integer = 2
  Const F_LEN As Integer = 3

  ' ** Array: arr_varUnique().
  Const U_FID As Integer = 0
  Const U_FAM As Integer = 1
  Const U_CON As Integer = 2

  ' ** Array: arr_varAPI().
  Const A_DID    As Integer = 0
  Const A_COMID  As Integer = 1
  Const A_APIID  As Integer = 2
  Const A_APINAM As Integer = 3

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    ' ** tblVBComponent_Declaration_Family, with pfxlen.
    Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_10")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngFams = .RecordCount
      .MoveFirst
      arr_varFam = .GetRows(lngFams)
      ' ****************************************************
      ' ** Array: arr_varFam()
      ' **
      ' **   Field  Element  Name               Constant
      ' **   =====  =======  =================  ==========
      ' **     1       0     vbdecfam_id        F_FID
      ' **     1       0     vbdecfam_name      F_FAM
      ' **     1       0     vbdecfam_prefix    F_PFX
      ' **     1       0     pfxlen             F_LEN
      ' **
      ' ****************************************************
      .Close
    End With

    ' ** Union of zz_qry_VBComponent_Declaration_09a (tblControlType, with with zz_qry_VBComponent_Declaration_08
    ' ** (tblVBComponent_Declaration_Family, just unique families), for 'AcControlType'; Cartesian),
    ' ** zz_qry_VBComponent_Declaration_09b (tblObjectType, with zz_qry_VBComponent_Declaration_08
    ' ** (tblVBComponent_Declaration_Family, just unique families), for 'AcObjectType'; Cartesian),
    ' ** zz_qry_VBComponent_Declaration_09c (tblIndexOrderType, with zz_qry_VBComponent_Declaration_08
    ' ** (tblVBComponent_Declaration_Family, just unique families), for 'DbIndexOrder'; Cartesian),
    ' ** zz_qry_VBComponent_Declaration_09d (tblDataTypeVb, with zz_qry_VBComponent_Declaration_08
    ' ** (tblVBComponent_Declaration_Family, just unique families), for 'VbDataType'; Cartesian),
    ' ** zz_qry_VBComponent_Declaration_09e (tblSystemColor_Base, with zz_qry_VBComponent_Declaration_08
    ' ** (tblVBComponent_Declaration_Family, just unique families),for 'VbSystemColors'; Cartesian),
    ' ** zz_qry_VBComponent_Declaration_09f (tblWindowsAccessRights, with zz_qry_VBComponent_Declaration_08
    ' ** (tblVBComponent_Declaration_Family, just unique families), for 'WinRights'; Cartesian).
    Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_09g")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngUniques = .RecordCount
      .MoveFirst
      arr_varUnique = .GetRows(lngUniques)
      ' ******************************************************
      ' ** Array: arr_varUnique()
      ' **
      ' **   Field  Element  Name                 Constant
      ' **   =====  =======  ===================  ==========
      ' **     1       0     vbdecfam_id          U_FID
      ' **     2       1     vbdecfam_name        U_FAM
      ' **     3       2     vbdecfam_constant    U_CON
      ' **
      ' ******************************************************
      .Close
    End With

    ' ** tblVBComponent_API, just needed fields, by specified CurrentAppName().
    Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_07")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngAPIs = .RecordCount
      .MoveFirst
      arr_varAPI = .GetRows(lngAPIs)
      ' **************************************************
      ' ** Array: arr_varAPI()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ==========
      ' **     1       0     dbs_id           A_DID
      ' **     2       1     vbcom_id         A_COMID
      ' **     3       2     vbcomapi_id      A_APIID
      ' **     4       3     vbcomapi_name    A_APINAM
      ' **
      ' **************************************************
      .Close
    End With

  End With  ' ** dbs.

  ' ** Update tblVBComponent_Declaration, for vbdec_notused = True, by specified CurrentAppName().
  Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Declaration_05")
  qdf.Execute dbFailOnError

  Set rst = dbs.OpenRecordset("tblVBComponent_Declaration", dbOpenDynaset, dbConsistent)

  lngNoScopes = 0&
  ReDim arr_varNoScope(NS_ELEMS, 0)

  lngChanges = 0&
  ReDim arr_varChange(CH_ELEMS, 0)

  lngBlockTypes = 0&: lngBlockEnums = 0&: lngDefTypes = 0&

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngModTot = .VBComponents.Count
    lngModCnt = 0&
    For Each vbc In .VBComponents  ' ** One-Based.
      lngModCnt = lngModCnt + 1&
      blnOptCompare = False: blnOptExplicit = False
      With vbc
        strModName = .Name
        varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & strModName & "'")
        Select Case IsNull(varTmp00)
        Case True
          lngVBComID = 0&
          Debug.Print "'MODULE NOT IN DATABASE! " & strModName
        Case False
          lngVBComID = varTmp00
        End Select
        Set cod = .CodeModule
        With cod
          lngLines = .CountOfLines
          lngDecLines = .CountOfDeclarationLines
          blnLineContinue = False: blnBlockType = False: blnBlockEnum = False
          strBlockName = vbNullString: strBlockType = vbNullString: strBlockMembers = vbNullString: strBlockScope = vbNullString
          lngBlockStart = 0&
          strDeclareName = vbNullString: strDeclareType = vbNullString: lngDeclareID = 0&
          strConstName = vbNullString: strConstType = vbNullString: lngConstID = 0&
          For lngX = 1& To lngDecLines
            strLine = Trim$(.Lines(lngX, 1))
            strTmp01 = vbNullString: strTmp02 = vbNullString
            strVarName = vbNullString: strVarType = vbNullString
            blnNoScope = False: blnChange = False: blnDefType = False: blnCompiler = False
            lngDataTypeVb = vbUndeclared
            If strLine <> vbNullString Then
              If Left$(strLine, 1) <> "'" Then
                If blnLineContinue = False Then
                  If Left$(strLine, Len(PCIF)) = PCIF Or Left$(strLine, Len(PCELS)) = PCELS Or Left$(strLine, Len(PCEND)) = PCEND Then
                    ' ** Compiler Directive logic, not constant or variable.
                  Else

                    strScope = vbNullString
                    intPos1 = InStr(strLine, " ")
                    If intPos1 > 0 Then
                      strTmp01 = Left$(strLine, (intPos1 - 1))  ' ** 1st word.
                      Select Case strTmp01
                      Case PPUB
                        ' ** Doc these below.
                        strScope = strTmp01
                      Case PPRIV
                        ' ** Doc these below.
                        strScope = strTmp01
                      Case PSTAT
                        ' ** Shouldn't be in a module's declaration section.
                        blnChange = True
                      Case PDIM
                        ' ** Should be changed. (It is allowed in Class Modules.)
                        blnChange = True
                      Case PGLOB
                        ' ** Change to Public.
                        blnChange = True
                      Case P2CONST
                        ' ** Unscoped!
                        blnNoScope = True
                      Case P2TYPE
                        ' ** Unscoped!
                        blnNoScope = True
                      Case P2ENUM
                        ' ** Unscoped!
                        blnNoScope = True
                      Case P2DEC
                        ' ** Unscoped!
                        blnNoScope = True
                      Case POPT
                        If strLine = "Option Compare Database" Then
                          blnOptCompare = True
                        ElseIf strLine = "Option Compare Text" Then  ' ** Just 1 usage?
                          blnOptCompare = True
                        ElseIf strLine = "Option Explicit" Then
                          blnOptExplicit = True
                        Else
                          ' ** I don't think we use any others!
                          Debug.Print "'WHAT'S THIS? 1 " & strModName & "  " & strLine
                        End If
                      Case PCOMP
                        ' ** Compiler directive.
                        ' **   #Const Directive
                        ' **   #If...Then...#Else Directive, including #If, #Else, #ElseIf, #End If
                        ' ** Doc these below.
                        blnCompiler = True
                      Case Else
                        If strTmp01 = "DefStr" Or strTmp01 = "DefLng" Or _
                            strTmp01 = "DefVar" Or strTmp01 = "DefBool" Then
                          ' ** DefType Statements.
                          ' ** Doc these below.
                          blnDefType = True
                          lngDefTypes = lngDefTypes + 1&
                        ElseIf strLine = PENDT Then
                          ' ** Doc these below.
                        ElseIf strLine = PENDE Then
                          ' ** Doc these below.
                        Else
                          ' ** Variables within a Type or Enum block.
                          intPos1 = InStr(strLine, "'")
                          If intPos1 > 0 Then strLine = Trim$(Left$(strLine, (intPos1 - 1)))
                          If blnBlockType = True Then
                            ' ** Doc these when finished.
                            strLine = Rem_Spaces(strLine)  ' ** Module Function: modStringFuncs.
                            If strBlockMembers <> vbNullString Then strBlockMembers = strBlockMembers & vbCrLf
                            strBlockMembers = strBlockMembers & strLine
                          ElseIf blnBlockEnum = True Then
                            ' ** Doc these when finished.
                            strLine = Rem_Spaces(strLine)  ' ** Module Function: modStringFuncs.
                            If strBlockMembers <> vbNullString Then strBlockMembers = strBlockMembers & vbCrLf
                            strBlockMembers = strBlockMembers & strLine
                          Else
                            ' ** What else?
                            Debug.Print "'WHAT'S THIS? 2 " & strModName & "  " & strLine
                            Stop
                          End If
                        End If  ' ** strTmp01.
                      End Select
                    Else
                      ' ** What types of one-word declarations could there be?
                      Debug.Print "'WHAT'S THIS? 3 " & strModName & "  " & strLine
                    End If  ' ** Space.
                    If blnNoScope = True Then
                      lngNoScopes = lngNoScopes + 1&
                      lngE = lngNoScopes - 1&
                      ReDim Preserve arr_varNoScope(NS_ELEMS, lngE)
                      arr_varNoScope(NS_MOD, lngE) = strModName
                      arr_varNoScope(NS_LIN, lngE) = lngX
                      arr_varNoScope(NS_TXT, lngE) = strLine
                    ElseIf blnChange = True Then
                      lngChanges = lngChanges + 1&
                      lngE = lngChanges - 1&
                      ReDim Preserve arr_varChange(CH_ELEMS, lngE)
                      arr_varChange(CH_MOD, lngE) = strModName
                      arr_varChange(CH_LIN, lngE) = lngX
                      arr_varChange(CH_TXT, lngE) = strLine
                    ElseIf (blnOptCompare = True And lngX = 1) Or (blnOptExplicit = True And lngX = 2) Then
                      ' ** Let it pass unmolested.
                    ElseIf blnDefType = True Then
                      ' ** Doc these here.
                      strScope = "Private"
                      strTmp03 = Trim$(Mid$(strLine, (Len(strTmp02) + 1)))
                      strTmp03 = Trim$(Mid$(strTmp03, (Len(strTmp01) + 1)))  ' ** Strip the DefType.
                      With rst
                        blnAdd = False
                        If .BOF = True And .EOF = True Then
                          blnAdd = True
                        Else
                          .MoveFirst
                          .FindFirst "[vbcom_id] = " & CStr(lngVBComID) & " And [vbdec_name] = '" & strTmp01 & "'"
                          If .NoMatch = True Then
                            blnAdd = True
                          Else
                            .Edit
                          End If
                        End If
                        If blnAdd = True Then
                          .AddNew  ' ** DefType Statement.
                          ' ** ![vbdec_id] : AutoNumber.
                          ![dbs_id] = lngThisDbsID
                          ![vbcom_id] = lngVBComID
                          ![vbdec_module] = strModName
                          ![vbdec_name] = strTmp01
                        End If
                        ![vbdec_linenum1] = lngX
                        ![vbdec_linenum2] = Null
                        If strScope <> vbNullString Then
                          ![scopetype_type] = strScope
                        Else
                          If IsNull(![scopetype_type]) = False Then
                            ![scopetype_type] = Null
                          End If
                        End If
                        ![dectype_type] = "DefType"
                        ![vbdec_family] = Null
                        ![vbdec_value] = Null
                        ![vbdec_parameter] = Null
                        ![vbdec_member] = strTmp03
                        lngDataTypeVb = vbUndeclared
                        strTmp03 = "{untyped}"
                        Select Case strTmp01
                        Case "DefStr"
                          lngDataTypeVb = vbString
                          strTmp03 = "String"
                        Case "DefLng"
                          lngDataTypeVb = vbLong
                          strTmp03 = "Long"
                        Case "DefVar"
                          lngDataTypeVb = vbVariant
                          strTmp03 = "Variant"
                        Case "DefBool"
                          lngDataTypeVb = vbBoolean
                          strTmp03 = "Boolean"
                        End Select
                        ![datatype_vb_type] = lngDataTypeVb
                        ![vbdec_vbtype] = strTmp03
                        ![vbdec_userdefined] = Null
                        ![vbdec_notused] = False    ' ** Later.
                        ![vbdec_usecnt] = 0&        ' ** Later.
                        ![vbdec_datemodified] = Now()
                        .Update
                      End With
                    ElseIf strLine = PENDT Then
                      ' ** Doc these here.
                      With rst
                        blnAdd = False
                        If .BOF = True And .EOF = True Then
                          blnAdd = True
                        Else
                          .MoveFirst
                          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                            "[vbdec_name] = '" & strBlockName & "'"
                          If .NoMatch = True Then
                            blnAdd = True
                          Else
                            .Edit
                          End If
                        End If
                        If blnAdd = True Then
                          .AddNew  ' ** User-Defined Type Block.
                          ' ** ![vbdec_id] : AutoNumber.
                          ![dbs_id] = lngThisDbsID
                          ![vbcom_id] = lngVBComID
                          ![vbdec_module] = strModName
                          ![vbdec_name] = strBlockName
                        End If
                        ![vbdec_linenum1] = lngBlockStart
                        ![vbdec_linenum2] = lngX
                        If strBlockScope <> vbNullString Then
                          ![scopetype_type] = strBlockScope
                        Else
                          If IsNull(![scopetype_type]) = False Then
                            ![scopetype_type] = Null
                          End If
                        End If
                        ![dectype_type] = "Type"
                        ![vbdec_family] = Null
                        ![vbdec_value] = Null
                        ![vbdec_parameter] = Null
                        If strBlockMembers <> vbNullString Then
                          ![vbdec_member] = strBlockMembers
                        Else
                          If IsNull(![vbdec_member]) = False Then
                            ![vbdec_member] = Null
                          End If
                        End If
                        lngDataTypeVb = vbUserDefinedType
                        ![datatype_vb_type] = lngDataTypeVb
                        ![vbdec_vbtype] = "{user-defined}"
                        ![vbdec_userdefined] = strBlockName
                        ![vbdec_notused] = False    ' ** Later.
                        ![vbdec_usecnt] = 0&        ' ** Later.
                        ![vbdec_datemodified] = Now()
                        If Left(strModName, 5) = "Form_" Or Left(strModName, 7) = "Report_" Then
                          ![comtype_type] = vbext_ct_Document
                        ElseIf Left(strModName, 3) = "mod" Or Left(strModName, 6) = "zz_mod" Then
                          ![comtype_type] = vbext_ct_StdModule
                        ElseIf Left(strModName, 3) = "cls" Then
                          ![comtype_type] = vbext_ct_ClassModule
                        End If
                        ![compdiropt_type] = 0  ' ** {none}.
                        .Update
                      End With
                      lngBlockStart = 0&
                      strBlockName = vbNullString: strBlockType = vbNullString: strBlockMembers = vbNullString: strBlockScope = vbNullString
                      blnBlockType = False
                    ElseIf strLine = PENDE Then
                      ' ** Doc these here.
                      With rst
                        blnAdd = False
                        If .BOF = True And .EOF = True Then
                          blnAdd = True
                        Else
                          .MoveFirst
                          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                            "[vbdec_name] = '" & strBlockName & "'"
                          If .NoMatch = True Then
                            blnAdd = True
                          Else
                            .Edit
                          End If
                        End If
                        If blnAdd = True Then
                          .AddNew  ' ** User-Defined Enum Block.
                          ' ** ![vbdec_id] : AutoNumber.
                          ![dbs_id] = lngThisDbsID
                          ![vbcom_id] = lngVBComID
                          ![vbdec_module] = strModName
                          ![vbdec_name] = strBlockName
                        End If
                        ![vbdec_linenum1] = lngBlockStart
                        ![vbdec_linenum2] = lngX
                        If strBlockScope <> vbNullString Then
                          ![scopetype_type] = strBlockScope
                        Else
                          If IsNull(![scopetype_type]) = False Then
                            ![scopetype_type] = Null
                          End If
                        End If
                        ![dectype_type] = "Enum"
                        ![vbdec_family] = Null
                        ![vbdec_value] = Null
                        ![vbdec_parameter] = Null
                        If strBlockMembers <> vbNullString Then
                          ![vbdec_member] = strBlockMembers
                        Else
                          If IsNull(![vbdec_member]) = False Then
                            ![vbdec_member] = Null
                          End If
                        End If
                        lngDataTypeVb = vbUserDefinedType
                        ![datatype_vb_type] = lngDataTypeVb
                        ![vbdec_vbtype] = "{user-defined}"
                        ![vbdec_userdefined] = strBlockName
                        ![vbdec_notused] = False    ' ** Later.
                        ![vbdec_usecnt] = 0&        ' ** Later.
                        ![vbdec_datemodified] = Now()
                        .Update
                      End With
                      lngBlockStart = 0&
                      strBlockName = vbNullString: strBlockType = vbNullString: strBlockMembers = vbNullString: strBlockScope = vbNullString
                      blnBlockEnum = False
                    ElseIf blnCompiler = True Then
                      strTmp03 = Trim$(Mid$(strLine, intPos1))
                      intPos1 = InStr(strTmp03, " ")
                      If intPos1 > 0 Then
                        strTmp02 = Left$(strTmp03, (intPos1 - 1))  ' ** 2nd word.
                        strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                        If Left$(strTmp03, 1) = "=" Then
                          strTmp03 = Trim$(Mid$(strTmp03, 2))
                          intPos1 = InStr(strTmp03, "'")
                          If intPos1 > 0 Then strTmp03 = Trim$(Left$(strTmp03, (intPos1 - 1)))
                          With rst
                            blnAdd = False
                            If .BOF = True And .EOF = True Then
                              blnAdd = True
                            Else
                              .MoveFirst
                              .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                                "[vbdec_name] = '" & strTmp02 & "'"
                              If .NoMatch = True Then
                                blnAdd = True
                              Else
                                .Edit
                              End If
                            End If
                            If blnAdd = True Then
                              .AddNew  ' ** Compiler Directive.
                              ' ** ![vbdec_id] : AutoNumber.
                              ![dbs_id] = lngThisDbsID
                              ![vbcom_id] = lngVBComID
                              ![vbdec_module] = strModName
                              ![vbdec_name] = strTmp02
                            End If
                            ![vbdec_linenum1] = lngX
                            ![vbdec_linenum2] = Null
                            ![scopetype_type] = "Private"
                            ![dectype_type] = "Compiler"
                            ![vbdec_family] = Null
                            If strTmp03 <> vbNullString Then
                              ![vbdec_value] = strTmp03
                            Else
                              If IsNull(![vbdec_value]) = False Then
                                ![vbdec_value] = Null
                              End If
                            End If
                            ![vbdec_parameter] = Null
                            ![vbdec_member] = Null
                            ![datatype_vb_type] = vbUndeclared
                            ![vbdec_vbtype] = "{untyped}"
                            ![vbdec_userdefined] = Null
                            ![vbdec_notused] = False    ' ** Later.
                            ![vbdec_usecnt] = 0&        ' ** Later.
                            ![vbdec_datemodified] = Now()
                            If Left(strModName, 5) = "Form_" Or Left(strModName, 7) = "Report_" Then
                              ![comtype_type] = vbext_ct_Document
                            ElseIf Left(strModName, 3) = "mod" Or Left(strModName, 6) = "zz_mod" Then
                              ![comtype_type] = vbext_ct_StdModule
                            ElseIf Left(strModName, 3) = "cls" Then
                              ![comtype_type] = vbext_ct_ClassModule
                            End If
                            ![compdiropt_type] = 0  ' ** {none}.
                            .Update
                          End With
                        Else
                          Debug.Print "'WHAT GIVES? " & strModName & "  " & strLine
                        End If
                      Else
                        Debug.Print "'I DON'T UNDERSTAND! " & strModName & "  " & strLine
                      End If
                    ElseIf blnBlockType = True Or blnBlockEnum = True Then
                      ' ** Already dealt with.
                    Else

                      strTmp03 = Trim$(Mid$(strLine, intPos1))
                      intPos1 = InStr(strTmp03, " ")
                      If intPos1 > 0 Then
                        strTmp02 = Left$(strTmp03, (intPos1 - 1))  ' ** 2nd word.
                        Select Case strTmp02
                        Case P2CONST
                          ' ** Doc these here.
                          If Right$(strLine, 1) = "_" Then
                            blnLineContinue = True
                          Else
                            blnLineContinue = False
                          End If
                          strTmp03 = Trim$(Mid$(strTmp03, intPos1))  ' ** 3rd word to the end of line.
                          intPos1 = InStr(strTmp03, " ")
                          If intPos1 > 0 Then
                            strConstName = Left$(strTmp03, (intPos1 - 1))
                            strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                            If Left$(strTmp03, 3) = "As " Then
                              intPos1 = InStr(strTmp03, " ")
                              strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                              intPos1 = InStr(strTmp03, "=")
                              If intPos1 > 0 Then
                                strConstType = Trim$(Left$(strTmp03, (intPos1 - 1)))
                                strTmp03 = Trim$(Mid$(strTmp03, (intPos1 + 1)))
                              Else
                                strConstType = strTmp03
                                strTmp03 = "{undefined}"
                              End If
                            Else
                              intPos1 = InStr(strTmp03, "=")
                              If intPos1 > 0 Then
                                strTmp03 = Trim$(Mid$(strTmp03, (intPos1 + 1)))
                              Else
                                strTmp03 = "{undefined}"
                              End If
                              strConstType = "{untyped}"
                            End If
                            ' ** Remove any remarks from the definition.
                            If Left$(strTmp03, 1) = Chr(34) Then  ' ** Quotes.
                              'FIGURE OUT HOW TO DEAL WITH CONCATENATED STRINGS!!
                              If InStr(strTmp03, "'") > 0 Then
                                If InStr(strTmp03, "&") = 0 Then
                                  ' ** No possible concatenation on the line (because I never use '+' for strings!).
                                  For lngY = 2& To Len(strTmp03)
                                    If Mid$(strTmp03, lngY, 1) = Chr(34) Then
                                      strTmp03 = Left$(strTmp03, lngY)
                                      Exit For
                                    End If
                                  Next
                                Else
                                  ' ** Both an apostrophe and an ampersand on the line.
                                  Debug.Print "'WHAT'S THIS LOOK LIKE? " & strModName & "  " & strLine
                                End If
                              Else
                                ' ** No possible remark on the line.
                              End If
                            Else
                              intPos1 = InStr(strTmp03, "'")
                              If intPos1 > 0 Then strTmp03 = Trim$(Left$(strTmp03, (intPos1 - 1)))
                            End If
                          Else
                            ' ** Not Typed.
                            strConstName = strTmp03
                            strConstType = "{untyped}"
                            strTmp03 = "{undefined}"
                          End If
                          If Right$(strConstName, 2) = "()" Then
                            strConstName = Left$(strConstName, (Len(strConstName) - 2))
                          End If
                          ' ** Now see if we can identify its family.
                          blnFound = False: strFamily = vbNullString
                          For lngY = 0& To (lngFams - 1&)
                            If arr_varFam(F_LEN, lngY) <> 0& Then
                              If Left$(strConstName, arr_varFam(F_LEN, lngY)) = arr_varFam(F_PFX, lngY) Then
                                If Compare_StringA_StringB(Left$(strConstName, arr_varFam(F_LEN, lngY)), "=", _
                                    arr_varFam(F_PFX, lngY)) = True Then  ' ** Module Function: modStringFuncs.
                                  ' ** Only if it matches letter-for-letter, including case.
                                  blnFound = True
                                  strFamily = arr_varFam(F_FAM, lngY)
                                End If
                                Exit For
                              End If
                            Else
                              For lngZ = 0& To (lngUniques - 1&)
                                If arr_varUnique(U_CON, lngZ) = strConstName Then
                                  blnFound = True
                                  strFamily = arr_varUnique(U_FAM, lngZ)
                                  Exit For
                                End If
                              Next
                            End If
                          Next
                          With rst
                            blnAdd = False
                            If .BOF = True And .EOF = True Then
                              blnAdd = True
                            Else
                              .MoveFirst
                              .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                                "[vbdec_name] = '" & strConstName & "'"
                              If .NoMatch = True Then
                                blnAdd = True
                              Else
                                .Edit
                              End If
                            End If
                            If blnAdd = True Then
                              .AddNew  ' ** Constant.
                              ' ** ![vbdec_id] : AutoNumber.
                              ![dbs_id] = lngThisDbsID
                              ![vbcom_id] = lngVBComID
                              ![vbdec_module] = strModName
                              ![vbdec_name] = strConstName
                            End If
                            ![vbdec_linenum1] = lngX
                            ![vbdec_linenum2] = Null
                            If strScope <> vbNullString Then
                              ![scopetype_type] = strScope
                            Else
                              If IsNull(![scopetype_type]) = False Then
                                ![scopetype_type] = Null
                              End If
                            End If
                            ![dectype_type] = "Constant"
                            If strFamily <> vbNullString Then
                              ![vbdec_family] = strFamily
                            Else
                              If IsNull(![vbdec_family]) = False Then
                                ![vbdec_family] = Null
                              End If
                            End If
                            ![vbdec_value] = strTmp03
                            ![vbdec_parameter] = Null
                            ![vbdec_member] = Null
                            lngDataTypeVb = GetDataType(strConstType)  ' ** Function: Below.
                            ![datatype_vb_type] = lngDataTypeVb
                            If lngDataTypeVb = vbUserDefinedType Then
                              ![vbdec_vbtype] = "{user-defined}"
                              ![vbdec_userdefined] = strConstType
                            Else
                              ![vbdec_vbtype] = strConstType
                              ![vbdec_userdefined] = Null
                            End If
                            ![vbdec_notused] = False    ' ** Later.
                            ![vbdec_usecnt] = 0&        ' ** Later.
                            ![vbdec_datemodified] = Now()
                            If Left(strModName, 5) = "Form_" Or Left(strModName, 7) = "Report_" Then
                              ![comtype_type] = vbext_ct_Document
                            ElseIf Left(strModName, 3) = "mod" Or Left(strModName, 6) = "zz_mod" Then
                              ![comtype_type] = vbext_ct_StdModule
                            ElseIf Left(strModName, 3) = "cls" Then
                              ![comtype_type] = vbext_ct_ClassModule
                            End If
                            ![compdiropt_type] = 0  ' ** {none}.
                            .Update
                            .Bookmark = .LastModified
                            lngConstID = ![vbdec_id]
                          End With
                          Select Case blnLineContinue
                          Case True
                            ' ** Leave the Constant variables available.
                          Case False
                            ' ** All done.
                            strConstName = vbNullString
                            strConstType = vbNullString
                            lngConstID = 0&
                          End Select
                        Case P2TYPE
                          ' ** Doc these when finished.
                          blnBlockType = True
                          lngBlockTypes = lngBlockTypes + 1&
                          lngBlockStart = lngX
                          strBlockMembers = vbNullString
                          strBlockScope = strScope
                          strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                          intPos1 = InStr(strTmp03, "'")
                          If intPos1 > 0 Then strTmp03 = Trim$(Left$(strTmp03, (intPos1 - 1)))
                          intPos1 = InStr(strTmp03, " ")
                          If intPos1 > 0 Then
                            strBlockName = Left$(strTmp03, (intPos1 - 1))
                            strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                            If Left$(strTmp03, 3) = "As " Then
                              intPos1 = InStr(strTmp03, " ")
                              strBlockType = Trim$(Mid$(strTmp03, intPos1))
                              intPos1 = InStr(strBlockType, "'")
                              If intPos1 > 0 Then  ' ** There's a remark.
                                strBlockType = Left$(strBlockType, (intPos1 - 1))
                              End If
                            Else
                              ' ** Else What?
                              strBlockType = "{WHAT?}"
                            End If
                          Else
                            ' ** Not Typed.
                            strBlockName = strTmp03
                            strBlockType = "{untyped}"
                          End If
                        Case P2ENUM
                          ' ** Doc these when finished.
                          blnBlockEnum = True
                          lngBlockEnums = lngBlockEnums + 1&
                          lngBlockStart = lngX
                          strBlockMembers = vbNullString
                          strBlockScope = strScope
                          strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                          intPos1 = InStr(strTmp03, "'")
                          If intPos1 > 0 Then strTmp03 = Trim$(Left$(strTmp03, (intPos1 - 1)))
                          intPos1 = InStr(strTmp03, " ")
                          If intPos1 > 0 Then
                            strBlockName = Left$(strTmp03, (intPos1 - 1))
                            strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                            If Left$(strTmp03, 3) = "As " Then
                              intPos1 = InStr(strTmp03, " ")
                              strBlockType = Trim$(Mid$(strTmp03, intPos1))
                              intPos1 = InStr(strBlockType, "'")
                              If intPos1 > 0 Then  ' ** There's a remark.
                                strBlockType = Left$(strBlockType, (intPos1 - 1))
                              End If
                            Else
                              ' ** Else What?
                              strBlockType = "{WHAT?}"
                            End If
                          Else
                            ' ** Not Typed.
                            strBlockName = strTmp03
                            strBlockType = "{untyped}"
                          End If
                        Case P2DEC
                          ' ** Doc these here.
                          If Right$(strLine, 1) = "_" Then
                            blnLineContinue = True
                          Else
                            blnLineContinue = False
                          End If
                          strTmp03 = Trim$(Mid$(strTmp03, intPos1))  ' ** 3rd word to end of line.
                          intPos1 = InStr(strTmp03, " ")
                          If intPos1 > 0 Then
                            strDeclareType = Trim$(Left$(strTmp03, (intPos1 - 1)))
                            strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                            intPos1 = InStr(strTmp03, " ")
                            If intPos1 > 0 Then
                              strDeclareName = Trim$(Left$(strTmp03, (intPos1 - 1)))
                              strTmp03 = Trim$(Mid$(strTmp03, intPos1))
                            Else
                              ' ** I give up!
                              Debug.Print "'WHAT!?? " & strLine
                              Stop
                            End If
                          Else
                            ' ** That's crazy!
                            Debug.Print "'WHAT??! " & strLine
                            Stop
                          End If
                          lngAPIID = 0&
                          For lngY = 0& To (lngAPIs - 1&)
                            If arr_varAPI(A_APINAM, lngY) = strDeclareName Then
                              lngAPIID = arr_varAPI(A_APIID, lngY)
                              Exit For
                            End If
                          Next
                          intPos1 = InStr(strTmp03, "'")
                          If intPos1 > 0 Then strTmp03 = Trim$(Left$(strTmp03, (intPos1 - 1)))  ' ** Remark.
                          With rst
                            blnAdd = False
                            If .BOF = True And .EOF = True Then
                              blnAdd = True
                            Else
                              .MoveFirst
                              .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                                "[vbdec_name] = '" & strDeclareName & "'"
                              If .NoMatch = True Then
                                blnAdd = True
                              Else
                                .Edit
                              End If
                            End If
                            If blnAdd = True Then
                              .AddNew  ' ** API Declaration.
                              ![dbs_id] = lngThisDbsID
                              ![vbcom_id] = lngVBComID
                              ![vbdec_module] = strModName
                              ![vbdec_name] = strDeclareName
                            End If
                            If lngAPIID > 0& Then
                              ![vbcomapi_id] = lngAPIID
                            Else
                              If IsNull(![vbcomapi_id]) = False Then
                                ![vbcomapi_id] = Null
                              End If
                            End If
                            ![vbdec_linenum1] = lngX
                            ![vbdec_linenum2] = Null
                            If strScope <> vbNullString Then
                              ![scopetype_type] = strScope
                            Else
                              If IsNull(![scopetype_type]) = False Then
                                ![scopetype_type] = Null
                              End If
                            End If
                            ![dectype_type] = strDeclareType  ' ** Function or Sub.
                            ![vbdec_family] = Null
                            ![vbdec_value] = Null
                            ![vbdec_parameter] = strTmp03
                            ![vbdec_member] = Null
                            Select Case blnLineContinue
                            Case False
                              For lngY = Len(strTmp03) To 1& Step -1&
                                If Mid$(strTmp03, lngY, 1) = " " Then
                                  strTmp03 = Trim$(Mid$(strTmp03, lngY))
                                  Exit For
                                ElseIf Mid$(strTmp03, lngY, 1) = ")" Then
                                  strTmp03 = "{untyped}"
                                  Exit For
                                End If
                              Next
                              lngDataTypeVb = GetDataType(strTmp03)  ' ** Function: Below.
                              ![datatype_vb_type] = lngDataTypeVb
                              If lngDataTypeVb = vbUserDefinedType Then
                                ![vbdec_vbtype] = Null
                                ![vbdec_userdefined] = strTmp03
                              Else
                                ![vbdec_vbtype] = strTmp03
                                ![vbdec_userdefined] = Null
                              End If
                            Case True
                              ![datatype_vb_type] = vbUndeclared
                              ![vbdec_vbtype] = "{untyped}"
                              ![vbdec_userdefined] = Null
                            End Select
                            ![vbdec_notused] = False    ' ** Later.
                            ![vbdec_usecnt] = 0&        ' ** Later.
                            ![vbdec_datemodified] = Now()
                            If Left(strModName, 5) = "Form_" Or Left(strModName, 7) = "Report_" Then
                              ![comtype_type] = vbext_ct_Document
                            ElseIf Left(strModName, 3) = "mod" Or Left(strModName, 6) = "zz_mod" Then
                              ![comtype_type] = vbext_ct_StdModule
                            ElseIf Left(strModName, 3) = "cls" Then
                              ![comtype_type] = vbext_ct_ClassModule
                            End If
                            ![compdiropt_type] = 0  ' ** {none}.
                            .Update
                            .Bookmark = .LastModified
                            lngDeclareID = ![vbdec_id]
                          End With
                          Select Case blnLineContinue
                          Case True
                            ' ** Leave the Declare variables available.
                          Case False
                            ' ** All done.
                            strDeclareName = vbNullString
                            strDeclareType = vbNullString
                            lngDeclareID = 0&
                          End Select
                        Case Else
                          ' ** Variable name, etc.
                          If blnBlockType = False And blnBlockEnum = False And lngX > 2& Then
                            If InStr(strLine, " As ") = 0 Then
                              ' ** Untyped?
                              Debug.Print "'UNTYPED? " & strModName & "  " & strLine
                            Else
                              ' ** strTmp02 is 1st variable name.
                              ' ** strTmp03 is whole line, including 1st variable name.
                              intPos1 = InStr(strTmp03, "'")
                              If intPos1 > 0 Then strTmp03 = Trim$(Left$(strTmp03, (intPos1 - 1)))
                              blnVarLoop = True: blnArray = False
                              Do While blnVarLoop = True
                                strVarName = strTmp02
                                intPos1 = InStr(strTmp03, ",")
                                If InStr(strTmp03, "(") > 0 Then
                                  If intPos1 > 0 Then
                                    If InStr(strTmp03, "(") < intPos1 Then
                                      intPos1 = InStr(InStr(strTmp03, ")"), strTmp03, ",")
                                      blnArray = True
                                    End If
                                  Else
                                    blnArray = True
                                  End If
                                End If
                                If intPos1 > 0 Then
                                  strVarType = Left$(strTmp03, (intPos1 - 1))
                                  strVarType = Trim$(Mid$(strVarType, (Len(strTmp02) + 1)))
                                  If Left$(strVarType, 3) = "As " Then
                                    strVarType = Trim$(Mid$(strVarType, 3))
                                    strTmp03 = Trim$(Mid$(strTmp03, (intPos1 + 1)))
                                    intPos1 = InStr(strTmp03, " ")
                                    If intPos1 > 0 Then
                                      strTmp02 = Trim$(Left$(strTmp03, (intPos1 - 1)))  ' ** Ready for the next loop.
                                    Else
                                      ' ** Last one, and not typed?
                                      strTmp02 = strTmp03
                                    End If
                                  Else
                                    Debug.Print "'WHAT'S THIS? 4 " & strModName & "  " & strLine
                                    strVarType = "{untyped}"
                                  End If
                                  intPos1 = InStr(strVarType, "'")
                                  If intPos1 > 0 Then strVarType = Trim$(Left$(strVarType, (intPos1 - 1)))  ' ** Remark.
                                Else
                                  ' ** Last (or only) variable on the line.
                                  blnVarLoop = False
                                  If strTmp02 = strTmp03 Then
                                    strVarType = "{untyped}"
                                  Else
                                    If blnArray = True Then
                                      intPos1 = InStr(strTmp03, " As ")
                                      If intPos1 > 0 Then
                                        strVarType = Trim$(Mid$(strTmp03, (intPos1 + 3)))
                                        strVarName = Trim$(Left$(strTmp03, intPos1))
                                      Else
                                        strVarName = strTmp03
                                        strVarType = "{untyped}"
                                      End If
                                    Else
                                      strTmp03 = Trim$(Mid$(strTmp03, (Len(strTmp02) + 1)))
                                      If Left$(strTmp03, 3) = "As " Then
                                        strVarType = Trim$(Mid$(strTmp03, 3))
                                        intPos1 = InStr(strVarType, "'")
                                        If intPos1 > 0 Then strVarType = Trim$(Left$(strVarType, (intPos1 - 1)))  ' ** Remark.
                                      Else
                                        Debug.Print "'WHAT'S THIS? 5 " & strModName & "  " & strLine
                                        strVarType = "{untyped}"
                                      End If
                                    End If
                                  End If
                                End If
                                If Right$(strVarName, 2) = "()" Then
                                  strVarName = Left$(strVarName, (Len(strVarName) - 2))
                                End If
                                With rst
                                  blnAdd = False
                                  If .BOF = True And .EOF = True Then
                                    blnAdd = True
                                  Else
                                    .MoveFirst
                                    .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                                      "[vbdec_name] = '" & strVarName & "'"
                                    If .NoMatch = True Then
                                      blnAdd = True
                                    Else
                                      .Edit
                                    End If
                                  End If
                                  If blnAdd = True Then
                                    .AddNew  ' ** Variable.
                                    ![dbs_id] = lngThisDbsID
                                    ![vbcom_id] = lngVBComID
                                    ![vbdec_module] = strModName
                                    ![vbdec_name] = strVarName
                                  End If
                                  ![vbdec_linenum1] = lngX
                                  ![vbdec_linenum2] = Null
                                  If strScope <> vbNullString Then
                                    ![scopetype_type] = strScope
                                  Else
                                    If IsNull(![scopetype_type]) = False Then
                                      ![scopetype_type] = Null
                                    End If
                                  End If
                                  ![dectype_type] = "Variable"
                                  'If Right$(strVarName, 2) = "()" Then strVarName = Left$(strVarName, (Len(strVarName) - 2))
                                  ![vbdec_family] = Null
                                  ![vbdec_value] = Null
                                  ![vbdec_parameter] = Null
                                  ![vbdec_member] = Null
                                  lngDataTypeVb = GetDataType(strVarType)  ' ** Function: Below.
                                  If lngDataTypeVb = vbUserDefinedType Then
                                    intPos1 = InStr(strVarType, ".")
                                    If intPos1 > 0 Then
                                      varTmp00 = Left$(strVarType, (intPos1 - 1))
                                      Select Case varTmp00
                                      Case "DAO", "VBA", "Access", "Office", "Scripting"
                                        ' ** Examples: DAO.Database, VBA.Collection, Access.Control, Office.CommandBar, Scripting.Folder.
                                        lngDataTypeVb = vbObject
                                        ![vbdec_vbtype] = Mid$(strVarType, (intPos1 + 1))
                                        ![vbdec_member] = varTmp00
                                        ![vbdec_userdefined] = Null
                                      Case Else
                                        ![vbdec_vbtype] = Null
                                        ![vbdec_userdefined] = strVarType
                                      End Select
                                      varTmp00 = Null
                                    Else
                                      If strVarType = "VbMsgBoxResult" Then
                                        lngDataTypeVb = vbInteger
                                        ![vbdec_vbtype] = strVarType
                                        ![vbdec_userdefined] = Null
                                      Else
                                        ![vbdec_vbtype] = Null
                                        ![vbdec_userdefined] = strVarType
                                      End If
                                    End If
                                  Else
                                    ![vbdec_vbtype] = strVarType
                                    ![vbdec_userdefined] = Null
                                  End If
                                  ![datatype_vb_type] = lngDataTypeVb
                                  ![vbdec_notused] = False    ' ** Later.
                                  ![vbdec_usecnt] = 0&        ' ** Later.
                                  ![vbdec_datemodified] = Now()
                                  If Left(strModName, 5) = "Form_" Or Left(strModName, 7) = "Report_" Then
                                    ![comtype_type] = vbext_ct_Document
                                  ElseIf Left(strModName, 3) = "mod" Or Left(strModName, 6) = "zz_mod" Then
                                    ![comtype_type] = vbext_ct_StdModule
                                  ElseIf Left(strModName, 3) = "cls" Then
                                    ![comtype_type] = vbext_ct_ClassModule
                                  End If
                                  ![compdiropt_type] = 0  ' ** {none}.
                                  .Update
                                End With
                              Loop  ' ** blnVarLoop.
                            End If
                          End If
                        End Select
                      Else
                        ' ** What?
                        Debug.Print "'WHAT? " & strModName & "  LINE: " & CStr(lngX) & "  " & strLine
                        'Stop
                      End If  ' ** Space.
                    End If  ' ** blnNoScope/blnChange.
                  End If  ' ** #If, #Else, #End If
                Else  ' ** Line continuation.
                  If lngDeclareID > 0& Then
                    If Right$(strLine, 1) = "_" Then
                      blnLineContinue = True
                    Else
                      blnLineContinue = False
                    End If
                    lngDataTypeVb = vbUndeclared
                    strTmp03 = "{untyped}"
                    intPos1 = InStr(strLine, "'")
                    If intPos1 > 0 Then strLine = Trim$(Left$(strLine, (intPos1 - 1)))
                    If blnLineContinue = False Then
                      For lngY = Len(strLine) To 1& Step -1&
                        If Mid$(strLine, lngY, 1) = " " Then
                          strTmp03 = Trim$(Mid$(strLine, lngY))
                          Exit For
                        ElseIf Mid$(strLine, lngY, 1) = ")" Then
                          strTmp03 = "{untyped}"
                          Exit For
                        End If
                      Next
                      lngDataTypeVb = GetDataType(strTmp03)  ' ** Function: Below.
                    End If
                    With rst
                      .MoveFirst
                      .FindFirst "[vbdec_id] = " & CStr(lngDeclareID)
                      If .NoMatch = False Then
                        .Edit
                        If IsNull(![vbdec_parameter]) = True Then
                          ![vbdec_parameter] = strLine
                        Else
                          ![vbdec_parameter] = ![vbdec_parameter] & vbCrLf & strLine
                        End If
                        ![vbdec_linenum2] = lngX
                        ![datatype_vb_type] = lngDataTypeVb
                        If lngDataTypeVb = vbUserDefinedType Then
                          ![vbdec_vbtype] = Null
                          ![vbdec_userdefined] = strTmp03
                        Else
                          ![vbdec_vbtype] = strTmp03
                          ![vbdec_userdefined] = Null
                        End If
                        ![vbdec_datemodified] = Now()
                        .Update
                      Else
                        Debug.Print "'NOT FOUND! " & strDeclareName & "  ID: " & CStr(lngDeclareID)
                      End If
                    End With
                    Select Case blnLineContinue
                    Case True
                      ' ** Leave the Declare variables available.
                    Case False
                      ' ** All done.
                      strDeclareName = vbNullString
                      strDeclareType = vbNullString
                      lngDeclareID = 0&
                    End Select
                  ElseIf lngConstID > 0& Then
                    With rst
                      .MoveFirst
                      .FindFirst "[vbdec_id] = " & CStr(lngConstID)
                      If .NoMatch = False Then
                        .Edit
                        If IsNull(![vbdec_value]) = True Then
                          ![vbdec_value] = strLine
                        Else
                          ![vbdec_value] = ![vbdec_value] & vbCrLf & strLine
                        End If
                        ![vbdec_linenum2] = lngX
                        ![vbdec_datemodified] = Now()
                        .Update
                      Else
                        Debug.Print "'NOT FOUND! " & strConstName & "  ID: " & CStr(lngConstID)
                      End If
                    End With
                    If Right$(strLine, 1) = "_" Then
                      blnLineContinue = True
                    Else
                      blnLineContinue = False
                    End If
                    Select Case blnLineContinue
                    Case True
                      ' ** Leave the Constant variables available.
                    Case False
                      ' ** All done.
                      strConstName = vbNullString
                      strConstType = vbNullString
                      lngConstID = 0&
                    End Select
                  Else
                    Debug.Print "'WHAT'S THIS? 6 " & strModName & "  " & strLine
                  End If
                End If  ' ** blnLineContinue.
              End If  ' ** Remark.
              If Right$(strLine, 1) = "_" Then
                blnLineContinue = True
              Else
                blnLineContinue = False
              End If
            End If  ' ** vbNullString.
          Next     ' ** lngX.
        End With  ' ** cod.
        If blnOptCompare = False Or blnOptExplicit = False Then
          Debug.Print "'MISSING OPTION! " & strModName
        End If
      End With  ' ** vbc.
    Next      ' ** vbc.
  End With  ' ** vbp.

  rst.Close

  ' ** Delete zz_qry_VBComponent_Declaration_01a (tblVBComponent_Declaration,
  ' ** for vbdec_notused = True, by specified CurrentAppName()).
  Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Declaration_01b")
  qdf.Execute

  ' ** Update zz_qry_VBComponent_Declaration_08a (tblVBComponent_Declaration,
  ' ** just names ending '()', with vbdec_name_new).
  'Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Declaration_08b")
  'qdf.Execute dbFailOnError

  dbs.Close

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  'Debug.Print "'TYPES: " & CStr(lngBlockTypes) & "  ENUMS: " & CStr(lngBlockEnums) & "  DEFTYPES: " & CStr(lngDefTypes)

  If lngNoScopes > 0& Then
    For lngX = 0& To (lngNoScopes - 1&)
      Debug.Print "'NO SCOPE! " & arr_varNoScope(NS_MOD, lngX) & "  LINE: " & CStr(arr_varNoScope(NS_LIN, lngX)) & _
        "  " & arr_varNoScope(NS_TXT, lngX)
    Next
  Else
    'Debug.Print "'NO MISSING SCOPES!"
  End If

  If lngChanges > 0& Then
    For lngX = 0& To (lngChanges - 1&)
      Debug.Print "'CHANGE! " & arr_varChange(CH_MOD, lngX) & "  LINE: " & CStr(arr_varChange(CH_LIN, lngX)) & _
        "  " & arr_varChange(CH_TXT, lngX)
    Next
  Else
    'Debug.Print "'NO DECLARATION CHANGES!"
  End If

  'Debug.Print "'SURVEYED " & CStr(lngModCnt) & " OF " & CStr(lngModTot) & " MODULES!"
'TYPES: 58  ENUMS: 14  DEFTYPES: 4
'NO MISSING SCOPES!
'NO DECLARATION CHANGES!
'SURVEYED 447 OF 447 MODULES!

  Debug.Print "'DONE! " & THIS_PROC & "()"

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_PublicVar_Doc = blnRetVal

End Function

Public Function VBA_PublicUsage_Doc() As Boolean
' ** Document usage of all Public variables and constants to tblVBComponent_Declaration_Detail.
' ** Called by:
' **

  Const THIS_PROC As String = "VBA_PublicUsage_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngComps As Long, arr_varComp As Variant
  Dim lngVars As Long, arr_varVar As Variant
  Dim lngHits As Long, arr_varHit() As Variant, lngHitXs As Long, arr_varHitX As Variant
  Dim lngProcs As Long, arr_varProc As Variant
  Dim strModName As String, strLine As String, lngVBComID As Long
  Dim lngLines As Long, lngDecLines As Long
  Dim lngLine As Long, lngColumn As Long, lngLastLine As Long
  Dim lngCharCnt As Long, lngHitsTot As Long
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean, blnDoc As Boolean, blnAdd As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer, intLen As Integer
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varVar().
  Const V_DECID As Integer = 0
  Const V_DID   As Integer = 1
  Const V_DNAM  As Integer = 2
  Const V_VID   As Integer = 3
  Const V_MOD   As Integer = 4
  Const V_LIN   As Integer = 5
  Const V_SCP   As Integer = 6
  Const V_TYP   As Integer = 7
  Const V_NAM   As Integer = 8
  Const V_NOT   As Integer = 9
  Const V_CNT   As Integer = 10
  Const V_ARR   As Integer = 11

  ' ** Array: arr_varHit().
  Const H_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const H_DID As Integer = 0
  Const H_VID As Integer = 1
  Const H_LIN As Integer = 2
  Const H_PID As Integer = 3

  ' ** Array: arr_varProc().
  Const P_DID  As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_VID  As Integer = 2
  Const P_VNAM As Integer = 3
  Const P_PID  As Integer = 4
  Const P_PNAM As Integer = 5
  Const P_BEG  As Integer = 6
  Const P_END  As Integer = 7

  ' ** Array: arr_varComp().
  Const C_DID  As Integer = 0
  Const C_DNAM As Integer = 1
  Const C_VID  As Integer = 2
  Const C_VNAM As Integer = 3
  Const C_LINS As Integer = 4

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  For lngW = 1& To 2&

    Set dbs = CurrentDb
    With dbs

      ' ** tblVBComponent, just needed fields, by specified CurrentAppName().
      Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_02")
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngComps = .RecordCount
        .MoveFirst
        arr_varComp = .GetRows(lngComps)
        ' ************************************************
        ' ** Array: arr_varComp()
        ' **
        ' **   Field  Element  Name           Constant
        ' **   =====  =======  =============  ==========
        ' **     1       0     dbs_id         C_DID
        ' **     2       1     dbs_name       C_DNAM
        ' **     3       2     vbcom_id       C_VID
        ' **     4       3     vbcom_name     C_VNAM
        ' **     5       4     vbcom_lines    C_LINS
        ' **
        ' ************************************************
        .Close
      End With

      Select Case lngW
      Case 1&
        ' ** tblVBComponent_Declaration, just Public Variables, by specified CurrentAppName().
        Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_03a")
        strTmp01 = "VARIABLES"
      Case 2&
        ' ** tblVBComponent_Declaration, just Public Constants, by specified CurrentAppName().
        Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_03b")
        strTmp01 = "CONSTANTS"
      End Select
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngVars = .RecordCount
        .MoveFirst
        arr_varVar = .GetRows(lngVars)
        ' ***************************************************
        ' ** Array: arr_varVar()
        ' **
        ' **   Field  Element  Name              Constant
        ' **   =====  =======  ================  ==========
        ' **     1       0     vbdec_id          V_DECID
        ' **     2       1     dbs_id            V_DID
        ' **     3       2     dbs_name          V_DNAM
        ' **     4       3     vbcom_id          V_VID
        ' **     5       4     vbdec_module      V_MOD
        ' **     6       5     vbdec_linenum1    V_LIN
        ' **     7       6     scopetype_type    V_SCP
        ' **     8       7     dectype_type      V_TYP
        ' **     9       8     vbdec_name        V_NAM
        ' **    10       9     vbdec_notused     V_NOT
        ' **    11      10     vbdec_usecnt      V_CNT
        ' **    12      11     vbdec_arr         V_ARR
        ' **
        ' ***************************************************
        .Close
      End With  ' ** rst.

      .Close
    End With  ' ** dbs.

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'TRACING " & CStr(lngVars) & " " & strTmp01 & ":"
    DoEvents

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      strModName = vbNullString
      Debug.Print "'|";  ' ** 0
      DoEvents
      lngHitsTot = 0&
      For lngX = 0& To (lngVars - 1&)
        lngHits = 0&
        ReDim arr_varHit(H_ELEMS, 0)
        lngLastLine = 0&
        For Each vbc In .VBComponents
          lngVBComID = 0&
          With vbc

            strModName = .Name
            For lngY = 0& To (lngComps - 1&)
              If arr_varComp(C_VNAM, lngY) = strModName Then
                lngVBComID = arr_varComp(C_VID, lngY)
                Exit For
              End If
            Next  ' ** lngY.

            Set cod = .CodeModule
            With cod
              lngLines = .CountOfLines
              lngDecLines = .CountOfDeclarationLines
              lngLine = 1&: lngColumn = 1&
              blnFound = .Find(arr_varVar(V_NAM, lngX), lngLine, lngColumn, lngLines, -1, True, False, False)
              ' ** object.Find(target, startline, startcol, endline, endcol [, wholeword] [, matchcase] [, patternsearch]) As Boolean
              lngLines = .CountOfLines  ' ** The Find resets this to the found line!
              If blnFound = True Then
                lngY = 0&: lngZ = 0&
                Do While blnFound = True

                  ' ** Check to make sure it's really a match, with no underscores to the left or right.
                  strLine = .Lines(lngLine, 1)
                  blnDoc = False

                  intLen = Len(strLine)
                  intPos1 = InStr(strLine, arr_varVar(V_NAM, lngX))
                  lngCharCnt = 0&
                  If intPos1 > 0 Then
                    If Left$(Trim$(strLine), 1) <> "'" Then

                      Do While intPos1 > 0

                        blnDoc = True

                        ' ** Check for a remark at the end of the line.
                        intPos2 = InStr(strLine, "'")

                        ' ** Check for quotes, since the variable, or an apostrophe, could be inside them.
                        intPos3 = InStr(strLine, Chr(34))
                        If intPos3 > 0 Then
                          ' ** There should be an even number of quotes, open and close.
                          lngCharCnt = CharCnt(strLine, Chr(34))  ' ** Module Function: modStringFuncs.
                          If intPos2 > 0 Then
                            ' ** There's both an apostrophe and quotes.
                            If intPos1 < intPos2 And intPos1 < intPos3 Then
                              ' ** The variable is before both the quotes and the remark, so OK, continue.
                            Else
                              If intPos2 < intPos1 And intPos1 < intPos3 Then
                                ' ** The variable appears to be within a remark.
                                blnDoc = False
                              Else
                                ' ** We'll have to see how far I should take this.
                                'blnDoc = False
                              End If
                            End If
                          Else
                            ' ** There are quotes, but no remark.
                            If intPos1 < intPos3 Then
                              ' ** OK, continue.
                            Else
                              If lngCharCnt = 2& And InStr((intPos3 + 1), strLine, Chr(34)) < intPos1 Then
                                ' ** The variable is to the right of the 2nd and final quote, so OK, continue.
                              Else
                                ' ** We'll have to see how far I should take this.
                                'blnDoc = False
                              End If
                            End If
                          End If  ' ** intPos2.
                        Else
                          ' ** There are no quotes on the line.
                          If intPos2 > 0 Then
                            ' ** There's a remark on the line.
                            If intPos2 < intPos1 Then
                              ' ** The variable appears to be within a remark.
                              blnDoc = False
                            Else
                              ' ** OK, continue.
                            End If
                          Else
                            ' ** OK, continue.
                          End If  ' ** intPos2.
                        End If  ' ** intPos3.

                        If blnDoc = True Then
                          If intPos1 = 1 Then
                            If intLen > Len(arr_varVar(V_NAM, lngX)) Then
                              strTmp01 = Mid$(strLine, (intPos1 + Len(arr_varVar(V_NAM, lngX))), 1)  ' ** Character after variable.
                              Select Case strTmp01
                              Case Chr(32)  ' ** Space.
                                ' ** OK.
                              Case "_"
                                ' ** It's a part of a longer variable.
                                blnDoc = False
                              Case Else
                                ' ** I think anything else is OK.
                              End Select
                              If blnDoc = True Then
                                lngHits = lngHits + 1&
                                lngE = lngHits - 1&
                                ReDim Preserve arr_varHit(H_ELEMS, lngE)
                                ' *****************************************************
                                ' ** Array: arr_varHit()
                                ' **
                                ' **   Field  Element  Name                Constant
                                ' **   =====  =======  ==================  ==========
                                ' **     1       0     dbs_id              H_DID
                                ' **     2       1     vbcom_id            H_VID
                                ' **     3       2     vbdecdet_linenum    H_LIN
                                ' **     4       3     vbcomproc_id        H_PID
                                ' **
                                ' *****************************************************
                                arr_varHit(H_DID, lngE) = lngThisDbsID
                                arr_varHit(H_VID, lngE) = lngVBComID
                                arr_varHit(H_LIN, lngE) = lngLine
                                arr_varHit(H_PID, lngE) = Null
                              End If
                            Else
                              ' ** All by itself? That's odd.
                              lngHits = lngHits + 1&
                              lngE = lngHits - 1&
                              ReDim Preserve arr_varHit(H_ELEMS, lngE)
                              arr_varHit(H_DID, lngE) = lngThisDbsID
                              arr_varHit(H_VID, lngE) = lngVBComID
                              arr_varHit(H_LIN, lngE) = lngLine
                              arr_varHit(H_PID, lngE) = Null
                            End If
                          Else
                            strTmp01 = Mid$(strLine, (intPos1 - 1), 1)  ' ** Character before variable.
                            Select Case strTmp01
                            Case Chr(32)  ' ** Space.
                              ' ** OK.
                            Case "_"
                              ' ** It's a part of a longer variable.
                              blnDoc = False
                            Case Else
                              ' ** I think anything else is OK.
                            End Select
                            If blnDoc = True Then
                              strTmp01 = Mid$(strLine, (intPos1 + Len(arr_varVar(V_NAM, lngX))), 1)  ' ** Character after variable.
                              Select Case strTmp01
                              Case Chr(32)  ' ** Space.
                                ' ** OK.
                              Case "_"
                                ' ** It's a part of a longer variable.
                                blnDoc = False
                              Case Else
                                ' ** I think anything else is OK.
                              End Select
                              If blnDoc = True Then
                                lngHits = lngHits + 1&
                                lngE = lngHits - 1&
                                ReDim Preserve arr_varHit(H_ELEMS, lngE)
                                arr_varHit(H_DID, lngE) = lngThisDbsID
                                arr_varHit(H_VID, lngE) = lngVBComID
                                arr_varHit(H_LIN, lngE) = lngLine
                                arr_varHit(H_PID, lngE) = Null
                              End If
                            End If  ' ** blnDoc.
                          End If  ' ** intPos1.
                        End If  ' ** blnDoc.

                        If blnDoc = True Then
                          ' ** We only need one instance of the variable being used per line.
                          intPos1 = 0
                          Exit Do
                        Else
                          ' ** Check for the variable being used more than once on the line.
                          intPos1 = InStr((intPos1 + 1), strLine, arr_varVar(V_NAM, lngX))
                          If intPos1 = 0 Then
                            Exit Do
                          End If
                        End If

                      Loop   ' ** intpos1.

                    End If  ' ** Remark.
                  End If  ' ** intPos1.

                  lngLines = .CountOfLines  ' ** The Find resets this to the found line!
                  lngLine = lngLine + 1&

If lngLine = lngY Or lngLine = lngZ Then
  If strModName = "modGlobConst" Then
    Exit Do
  Else
Stop
    Exit Do
  End If
Else
  If lngY < lngZ Then
    lngY = lngLine
  Else
    lngZ = lngLine
  End If
End If
                  If lngLine > lngLines Then
                    blnFound = False
                  Else
                    blnFound = .Find(arr_varVar(V_NAM, lngX), lngLine, lngColumn, lngLines, -1, True, False, False)
                    lngLines = .CountOfLines  ' ** The Find resets this to the found line!
                  End If
                  If blnFound = False Then
                    Exit Do
                  End If

                Loop   ' ** blnFound.
              End If  ' ** blnFound.
            End With  ' ** cod.

lngZ = 0&: lngE = 0&
For lngY = 0& To (lngHits - 1&)
  If arr_varHit(H_LIN, lngY) = lngZ Then
    For lngZ = (lngY + 1&) To (lngHits - 1&)
      If arr_varHit(H_LIN, lngZ) = arr_varHit(H_LIN, lngY) And arr_varHit(H_VID, lngZ) = arr_varHit(H_VID, lngY) Then
        lngE = lngY + 1&
Stop
        Exit For
      End If
    Next
    If lngE <> 0& Then
      lngHits = lngE
      Exit For
    End If
  Else
    lngZ = arr_varHit(H_LIN, lngY)
  End If
Next

          End With  ' ** vbc.
        Next      ' ** vbc.


        If lngHits > 0& Then
          lngHitsTot = lngHitsTot + lngHits
          arr_varVar(V_CNT, lngX) = lngHits
          arr_varVar(V_ARR, lngX) = arr_varHit
          arr_varVar(V_NOT, lngX) = CBool(False)
        Else
          'arr_varVar(V_NOT, lngX) = CBool(True)
          arr_varVar(V_CNT, lngX) = CLng(0)
        End If

        If (lngX + 1&) Mod 100& = 0& Then
          Debug.Print "|  " & CStr(lngX + 1&)  ' ** 100
          Debug.Print "'|";  ' ** 100
        ElseIf (lngX + 1&) Mod 10& = 0& Then
          Debug.Print "|";
        Else
          Debug.Print ".";
        End If
        DoEvents

      Next      ' ** lngX.
    End With  ' ** vbp.

    Debug.Print
    Debug.Print "'HITS: " & CStr(lngHitsTot) & " FOR " & CStr(lngVars) & " VARS!"
    DoEvents

    ' ***************************************************
    ' ** Array: arr_varVar()
    ' **
    ' **   Field  Element  Name              Constant
    ' **   =====  =======  ================  ==========
    ' **     1       0     vbdec_id          V_DECID
    ' **     2       1     dbs_id            V_DID
    ' **     3       2     dbs_name          V_DNAM
    ' **     4       3     vbcom_id          V_VID
    ' **     5       4     vbdec_module      V_MOD
    ' **     6       5     vbdec_linenum1    V_LIN
    ' **     7       6     scopetype_type    V_SCP
    ' **     8       7     dectype_type      V_TYP
    ' **     9       8     vbdec_name        V_NAM
    ' **    10       9     vbdec_notused     V_NOT
    ' **    11      10     vbdec_usecnt      V_CNT
    ' **    12      11     vbdec_arr         V_ARR
    ' **
    ' ***************************************************

    Set dbs = CurrentDb
    With dbs

      Set rst = .OpenRecordset("tblVBComponent_Declaration", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngVars - 1&)
          .FindFirst "[vbdec_id] = " & CStr(arr_varVar(V_DECID, lngX)) & " And [dbs_id] = " & CStr(arr_varVar(V_DID, lngX))
          If .NoMatch = False Then
            .Edit
            If arr_varVar(V_NOT, lngX) = False Then
              ![vbdec_notused] = arr_varVar(V_NOT, lngX)
            End If
            ![vbdec_usecnt] = arr_varVar(V_CNT, lngX)
            ![vbdec_datemodified] = Now()
            .Update
          Else
            Stop
          End If
        Next
        .Close
      End With  ' ** rst.

      lngHits = 0&
      ReDim arr_varHit(H_ELEMS, 0)

      ' *****************************************************
      ' ** Array: arr_varHitX()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     dbs_id              H_DID
      ' **     2       1     vbcom_id            H_VID
      ' **     3       2     vbdecdet_linenum    H_LIN
      ' **     4       3     vbcomproc_id        H_PID
      ' **
      ' *****************************************************

      lngProcs = 0&
      arr_varProc = Empty
      strModName = vbNullString: blnDoc = False
      For lngX = 0& To (lngVars - 1&)
        If arr_varVar(V_CNT, lngX) > 0& Then
          arr_varHitX = arr_varVar(V_ARR, lngX)
          lngHitXs = (UBound(arr_varHitX, 2) + 1&)
          For lngY = 0& To (lngHitXs - 1&)

            strTmp01 = vbNullString: blnDoc = True
            For lngZ = 0& To (lngComps - 1&)
              If arr_varComp(C_VID, lngZ) = arr_varHitX(H_VID, lngY) Then
                strTmp01 = arr_varComp(C_VNAM, lngZ)
                Exit For
              End If
            Next
            If strTmp01 = vbNullString Then
              Stop
            End If

            If strTmp01 <> strModName Then
              strModName = strTmp01
              lngProcs = 0&
              arr_varProc = Empty
              ' ** tblVBComponent_Procedure, sorted, by specified [vbcomid].
              Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_04")
              With qdf.Parameters
                ![vbcomid] = arr_varHitX(H_VID, lngY)
              End With
              Set rst = qdf.OpenRecordset
              With rst
                If .BOF = True And .EOF = True Then
                  ' ** Shouldn't happen!
                  blnDoc = False
                  Stop
                Else
                  .MoveLast
                  lngProcs = .RecordCount
                  .MoveFirst
                  arr_varProc = .GetRows(lngProcs)
                  ' *******************************************************
                  ' ** Array: arr_varProc()
                  ' **
                  ' **   Field  Element  Name                  Constant
                  ' **   =====  =======  ====================  ==========
                  ' **     1       0     dbs_id                P_DID
                  ' **     2       1     dbs_name              P_DNAM
                  ' **     3       2     vbcom_id              P_VID
                  ' **     4       3     vbcom_name            P_VNAM
                  ' **     5       4     vbcomproc_id          P_PID
                  ' **     6       5     vbcomproc_name        P_PNAM
                  ' **     7       6     vbcomproc_line_beg    P_BEG
                  ' **     8       7     vbcomproc_line_end    P_END
                  ' **
                  ' *******************************************************
                End If
                .Close
              End With
              Set rst = Nothing
              Set qdf = Nothing
            End If

            If blnDoc = True Then
              For lngZ = 0& To (lngProcs - 1&)
                If (arr_varHitX(H_LIN, lngY) >= arr_varProc(P_BEG, lngZ)) And (arr_varHitX(H_LIN, lngY) <= arr_varProc(P_END, lngZ)) Then
                  arr_varHitX(H_PID, lngY) = arr_varProc(P_PID, lngZ)
                  Exit For
                End If
              Next  ' ** lngZ.
              If IsNull(arr_varHitX(H_PID, lngY)) = True Then
                ' ** Maybe there are changes in the code.
                arr_varHitX(H_PID, lngY) = arr_varProc(P_PID, 0)  ' ** Default to the Declaration section.
              Else
                If arr_varHitX(H_PID, lngY) = 0& Then
                  ' ** Maybe there are changes in the code.
                  arr_varHitX(H_PID, lngY) = arr_varProc(P_PID, 0)  ' ** Default to the Declaration section.
                End If
              End If
            End If  ' ** blnDoc.
            If IsNull(arr_varHitX(H_PID, lngY)) = True Then
              Stop
            Else
              If arr_varHitX(H_PID, lngY) = 0& Then
                Stop
              End If
            End If

          Next   ' ** lngY.
        End If  ' ** V_CNT.
        arr_varVar(V_ARR, lngX) = arr_varHitX
        lngHitXs = 0&
        arr_varHitX = Empty
      Next  ' ** lngX

      ' ** Update tblVBComponent_Declaration_Detail, for vbdecdet_notused = True, by specified CurrentAppName().
      Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_06")
      qdf.Execute

      Set rst = .OpenRecordset("tblVBComponent_Declaration_Detail", dbOpenDynaset, dbConsistent)
      With rst
        Debug.Print "'|";
        For lngX = 0& To (lngVars - 1&)
          If arr_varVar(V_CNT, lngX) > 0& Then
            arr_varHitX = arr_varVar(V_ARR, lngX)
            lngHitXs = (UBound(arr_varHitX, 2) + 1&)
            For lngY = 0& To (lngHitXs - 1&)
              blnAdd = False
              If .BOF = True And .EOF = True Then
                blnAdd = True
              Else
                .FindFirst "[dbs_id] = " & CStr(arr_varVar(V_DID, lngX)) & " And " & _
                  "[vbdec_id] = " & CStr(arr_varVar(V_DECID, lngX)) & " And " & _
                  "[vbcom_id] = " & CStr(arr_varHitX(H_VID, lngY)) & " And " & _
                  "[vbcomproc_id] = " & CStr(arr_varHitX(H_PID, lngY)) & " And " & _
                  "[vbdecdet_linenum] = " & CStr(arr_varHitX(H_LIN, lngY))
                Select Case .NoMatch
                Case True
                  blnAdd = True
                Case False
                  .Edit
                End Select
              End If
              If blnAdd = True Then
                .AddNew
                ![dbs_id] = arr_varVar(V_DID, lngX)
                ![vbdec_id] = arr_varVar(V_DECID, lngX)
                ![vbcom_id] = arr_varHitX(H_VID, lngY)
                ' ** ![vbdecdet_id] : AutoNumber.
                ![vbcomproc_id] = arr_varHitX(H_PID, lngY)
                ![vbdecdet_linenum] = arr_varHitX(H_LIN, lngY)
              End If
              ![vbdecdet_notused] = False
              ![vbdecdet_datemodified] = Now()
              ![dectype_type] = arr_varVar(V_TYP, lngX)
              ![compdiropt_type] = 0  ' ** {none}.
              .Update
            Next   ' ** lngY.
          End If  ' ** V_CNT.
          If (lngX + 1&) Mod 100 = 0 Then
            Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngVars)
            Debug.Print "'|";
          ElseIf (lngX + 1&) Mod 10 = 0 Then
            Debug.Print "|";
          Else
            Debug.Print ".";
          End If
          DoEvents
        Next  ' ** lngX
        .Close
      End With

      Debug.Print
      DoEvents

      ' ** Delete tblVBComponent_Declaration_Detail, for vbdecdet_notused = True.
      'Set qdf = .QueryDefs("zz_qry_VBComponent_Declaration_17")
      'qdf.Execute

      .Close
    End With  ' ** dbs.

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Beep

  Next  ' ** lngW.

  Debug.Print "'DONE! " & THIS_PROC & "()"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_PublicUsage_Doc = blnRetVal

End Function

Private Function GetDataType(strDataType As String) As Long
' ** Called by:
' **   VBA_PublicVar_Doc(), Above.

  Const THIS_PROC As String = "GetDataType"

  Dim lngRetVal As Long

  lngRetVal = vbUndeclared

  If strDataType <> vbNullString Then
    Select Case strDataType
    Case "Integer"
      lngRetVal = vbInteger
    Case "Long"
      lngRetVal = vbLong
    Case "Single"
      lngRetVal = vbSingle
    Case "Double"
      lngRetVal = vbDouble
    Case "Currency"
      lngRetVal = vbCurrency
    Case "Date"
      lngRetVal = vbDate
    Case "String"
      lngRetVal = vbString
    Case "Object"
      lngRetVal = vbObject
    Case "Boolean"
      lngRetVal = vbBoolean
    Case "Variant"
      lngRetVal = vbVariant
    Case "Byte"
      lngRetVal = vbByte
    Case "{untyped}"
      lngRetVal = vbUndeclared
    Case Else
      lngRetVal = vbUserDefinedType
    End Select
  End If

  GetDataType = lngRetVal

End Function

Public Function VBA_MsgBox_Parse(varInput As Variant, strInfo As String) As Variant
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

Public Function VBA_EventProcPrefix_Doc() As Boolean
' ** Currently not called or used.
'WAIT! The EventProcPrefix is just the word that begins the procedure
'event, which is then followed by the underscore and and event name.
'IT'S THE CONTROL'S NAME, which may be different if, for example,
'there's a space in the Control name, i.e.,
'Control name is "account name type", so
'EventProcPrefix is "account_name_type".
'NO NEED TO DOCUMENT!!
'It can, however, be used to match a control with one
'of its procedures, in case they're somewhat different.
'Perhaps it should be added to tblForm_Control or tblReport_Control.

  Const THIS_PROC As String = "VBA_EventProcPrefix_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim frm As Access.Form, ctl As Access.Control
  Dim lngEvennts As Long, arr_varEvennt() As Variant
  Dim lngRecs As Long
  Dim lngX As Long, lngE As Long

  Const E_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const E_NAM As Integer = 0
  Const E_PRX As Integer = 1
  Const E_FND As Integer = 2

  blnRetValx = True

  lngEvennts = 0&
  ReDim arr_varEvennt(E_ELEMS, 0)

  Set dbs = CurrentDb
  Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Event_01")
  Set rst = qdf.OpenRecordset
  With rst
    .MoveLast
    lngRecs = .RecordCount
    .MoveFirst
    For lngX = 1& To lngRecs
      lngEvennts = lngEvennts + 1&
      lngE = lngEvennts - 1&
      ReDim Preserve arr_varEvennt(E_ELEMS, lngE)
      arr_varEvennt(E_NAM, lngE) = ![vbcom_event_name]
      arr_varEvennt(E_PRX, lngE) = vbNullString
      arr_varEvennt(E_FND, lngE) = CBool(False)
      If lngX < lngRecs Then .MoveNext
    Next
    .Close
  End With
  dbs.Close

  Set frm = Forms(0)
  With frm
    For Each ctl In .Controls
      With ctl
        For lngX = 0& To (lngEvennts - 1&)

        Next
      End With
    Next
  End With

'EventProcPrefix:
'BoundObjectFrame
'Chart
'CheckBox
'ComboBox
'CommandButton
'FormSection
'Image
'Label
'Line
'ListBox
'OptionButton
'OptionGroup
'Page
'PageBreak
'Rectangle
'ReportSection
'SubForm
'Tab
'TextBox
'ToggleButton
'UnboundObjectFrame

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_EventProcPrefix_Doc = blnRetValx

End Function

Private Function VBA_ChkDocQrys(Optional varSkip As Variant) As Boolean

  Const THIS_PROC As String = "VBA_ChkDocQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strPath1 As String, strPath2 As String, strFile1 As String, strFile2 As String, strPathFile1 As String, strPathFile2 As String
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

    strPath2 = CurrentAppPath & LNK_SEP & "Backups"  ' ** Module Function: modFileUtilities.
    strFile2 = "TrstXAdm - Copy (6)_keep.mdb"
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

    Debug.Print "'MOD DOC QRYS: " & CStr(lngQrys)
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
  '    Set rst = .OpenRecordset("tblQuery_Documentation", dbOpenDynaset, dbAppendOnly)
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
  '        ![qrydoc_datemodified] = Now()
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
              Debug.Print "'QRY NOT FOUND!  " & arr_varQry(Q_QNAM, lngX)
              DoEvents
            Else
On Error GoTo 0
              arr_varQry(Q_IMP, lngX) = CBool(True)
            End If
          Else
On Error GoTo 0
            arr_varQry(Q_IMP, lngX) = CBool(True)
          End If
        End If
      Next
    Else
      Debug.Print "'ALL MOD DOC QRYS PRESENT!"
    End If

    Beep

    Debug.Print "'DONE!"
    DoEvents

  End If  ' ** blnSkip.

  Set dbs = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_ChkDocQrys = blnRetValx

End Function
