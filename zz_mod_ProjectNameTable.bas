Attribute VB_Name = "zz_mod_ProjectNameTable"
Option Compare Database
Option Explicit

'VGC 08/18/2013: CHANGES!

' ** ===================================
' ** PROJECT-NAME TABLE:
' **  32,768 ENTRIES MAX!
' **   CONSTANT NAMES
' **   VARIABLE NAMES
' **   TYPE DEFINITION NAMES
' **   MODULE NAMES
' **   DLL-PROCEDURE DECLARATION NAMES
' ** DLL DECLARE TABLE:
' **  ROUGHLY 1,5000 PER MODULE
' ** IMPORT TABLE:
' **  CROSS-MODULE REFERENCES
' **   ROUGHLY 2,000 REFS PER MODULE
' ** ===================================

'tblVBComponent                     : All modules: standard, class, form, report.
'  Unique:  493                     :   zz_qry_VBComponent_ProjectName_20a
'tblVBComponent_Declaration         : All module-level declarations.
'  Unique: 3207                     :   zz_qry_VBComponent_ProjectName_30b
'tblVBComponent_Declaration_Local   : All local variables and constants within Functions and Subs.
'  Unique: 4782                     :   zz_qry_VBComponent_ProjectName_04
'tblVBComponent_Declaration_Type    : All members of user-defined types.
'  Unique:  394                     :   zz_qry_VBComponent_ProjectName_36
'tblVBComponent_Procedure_Parameter : All procedure parameters.
'  Unique:  709                     :   zz_qry_VBComponent_ProjectName_41
'tblVBComponent_API                 : All API declarations.
'  Unique:  177                     :   zz_qry_VBComponent_ProjectName_51a
'tblVBComponent_API_Parameter       : All API parameters.
'  Unique:  254                     :   zz_qry_VBComponent_ProjectName_58
'==============
'Total:   10016

'tblDatabase_Table                  : All tables in or linked to Trust Accountant.
'  Unique:  367                     :   zz_qry_VBComponent_ProjectName_60
'tblQuery                           : All queries, including dev zz's.
'  Unique: 9715                     :   zz_qry_VBComponent_ProjectName_61
'tblForm                            : All forms.
'  Unique:  217                     :   zz_qry_VBComponent_ProjectName_62
'tblReport                          : All reports.
'  Unique:  177                     :   zz_qry_VBComponent_ProjectName_63
'tblMacro                           : All macros.
'  Unique:  187                     :   zz_qry_VBComponent_ProjectName_64
'===================
'Total:   10663
'===================
'Total:   20679

'tblVBComponent_CodeLine:           : All code line numbers.
' Unique:  8295                     :   zz_qry_VBComponent_ProjectName_23
'tblVBComponent_LineName            : All named code lines.
' Unique:     4                     :   zz_qry_VBComponent_ProjectName_24e
'tblVBComponent_Procedure           : All functions and procedures.
' Unique:  3982                     :   zz_qry_VBComponent_ProjectName_28
'===================
'Total:   12281
'===================
'Total:   32960
'
'32768 - 32960 = -192!  SHOULD HAVE HIT OUT-OF-MEMORY 192 NAMES AGO!
'(UNLESS ONE OF THOSE ABOVE SHOULDN'T BE INCLUDED!)
'(AND I AM RIGHT ON THE CUSP RIGHT NOW!)

' ** Code Line Numbers:
' **   Numbers with unique or few uses:
' **   Uses:    1     2     3     4     5
' **   Cnt:    572 + 809 + 347 + 778 + 423
' **          ===== ===== ===== ===== =====
' **   Total:  2929  SLOTS FREED BY SPLITING MODULES!
' ** 1:  modVersionConvertFuncs  Highest now 67020
' ** 2:  modVersionConvertFuncs
' **     frmRpt_CourtReports_NY  Highest still 75480
' ** 3:  modVersionConvertFuncs
' **     frmJournal_Columns_Sub
' **     frmStatementParameters
' ** 4:  frmAccountProfile_Sub
' **     frmJournal_Columns_Sub
' **     frmRpt_Checks
' **     frmStatementParameters
' ** 5:  frmAccountProfile_Sub
' **     frmCheckReconcile
' **     frmJournal_Sub4_Sold
' **     frmRpt_Checks
' **     modVersionConvertFuncs
'REPRESENTATIVE!

'NEEDED:
'A THOROUGH CHECK FOR UNUSED VARIABLES!

' ** VBA_PublicVar_Doc()
' **   Documents all module level variables, constants, declares, types,
' **   enums, deftypes, and #Const to tblVBComponent_Declaration.
' ** INCLUDES BOTH PUBLIC AND PRIVATE.

' ** VBA_PublicUsage_Doc()
' **   Document usage of all Public variables and constants to tblVBComponent_Declaration_Detail.
' ** NOT USEFUL HERE.

' ** VBA_WinDialog_Doc()
' **   Document usage of browse dialog Functions and Subs to tblVBComponent_Procedure_Detail.
' ** NOT USEFUL BY ITSELF.

' ** Libraries:
' **   gdi32.dll:
' **     Graphics Device Interface (GDI) functions for device output, such as those for drawing and font management.
' **   Kernel32.dll:
' **     Low-level operating system functions for memory management and resource handling.
' **   User32.dll:
' **     Windows management functions for message handling, timers, menus, and communications.

' ** An Alias is the name within the DLL. The declaration name is the one to be used in the code.

' ** As Any:
' ** In Visual Basic 6.0, when you declare a reference to an external procedure with the Declare statement,
' ** you can specify As Any for the data type of any of the parameters and for the return type.
' ** The As Any keyword disables type checking and allow any data type to be passed in or returned.
' ** VB7 does not support the Any keyword. In a Declare statement, you must specifically declare the data
' ** type of every parameter and of the return. This improves type safety. You can overload your procedure
' ** declaration to accommodate various return types.

'Primary, secondary, tertiary,quaternary, quinary, senary, septenary, octonary, nonary, denary.
'Words also exist for `twelfth order' (duodenary) and `twentieth order' (vigenary).

Private Const BOF As String = "xx"

Private Const THIS_NAME As String = "zz_mod_ProjectNameTable"
' **

Public Function VBA_LocalVar_Doc() As Boolean
' ** Document all local variables and constants to tblVBComponent_Declaration_Local.

  Const THIS_PROC As String = "VBA_LocalVar_Doc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim strModName As String, strLine As String, strProcName As String
  Dim strVarName As String, strVarScope As String, strVarDeclare As String, strVarType As String, strConstValue As String
  Dim blnIsArray As Boolean, strArray As String
  Dim lngModLines As Long, lngModDecLines As Long
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngVars As Long, arr_varVar() As Variant
  Dim lngDataTypes As Long, arr_varDataType As Variant
  Dim lngTotVars As Long, lngHighVars As Long, strHighVars As String, lngLoopCnt As Long, lngLastLoopCnt As Long
  Dim lngThisDbsID As Long
  Dim blnLineContNext As Boolean, blnFound As Boolean, blnSkip As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long, lngF As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varProc.
  Const P_DID  As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_CID  As Integer = 2
  Const P_CNAM As Integer = 3
  Const P_CTYP As Integer = 4
  Const P_PID  As Integer = 5
  Const P_PNAM As Integer = 6
  Const P_BEG  As Integer = 7
  Const P_END  As Integer = 8

  ' ** Array: arr_varVar().
  Const V_ELEMS As Integer = 15  ' ** Array's first-element UBound().
  Const V_DID   As Integer = 0
  Const V_DNAM  As Integer = 1
  Const V_CID   As Integer = 2
  Const V_CNAM  As Integer = 3
  Const V_CTYP  As Integer = 4
  Const V_PID   As Integer = 5
  Const V_PNAM  As Integer = 6
  Const V_VNAM  As Integer = 7
  Const V_SCOP  As Integer = 8
  Const V_DECL  As Integer = 9
  Const V_DTYP  As Integer = 10
  Const V_VTYP  As Integer = 11
  Const V_ISARR As Integer = 12
  Const V_ARR   As Integer = 13
  Const V_CVAL  As Integer = 14
  Const V_LIN   As Integer = 15

  ' ** Array: arr_varDataType().
  Const D_ID    As Integer = 0
  Const D_TYPE  As Integer = 1
  Const D_CONST As Integer = 2
  Const D_NAME  As Integer = 3
  Const D_DATE  As Integer = 4

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs
    ' ** Empty tblVBComponent_Declaration_Local.
    Set qdf = .QueryDefs("zz_qry_VBComponent_ProjectName_01")
    qdf.Execute
    DoEvents
  End With

  Set vbp = Application.VBE.ActiveVBProject
  With vbp

    lngTotVars = 0&: lngHighVars = 0&: strHighVars = vbNullString
    lngLoopCnt = 0&: lngLastLoopCnt = 0&

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

    For Each vbc In .VBComponents
      With vbc

        strModName = .Name

        blnSkip = False

        lngVars = 0&
        ReDim arr_varVar(V_ELEMS, 0)

        ' ** tblVBComponent_Procedure, by specified [dbid], [comnam].
        Set qdf = dbs.QueryDefs("zz_qry_VBComponent_ProjectName_02")
        With qdf.Parameters
          ![dbid] = lngThisDbsID
          ![comnam] = strModName
        End With
        Set rst = qdf.OpenRecordset
        With rst
          If .BOF = True And .EOF = True Then
            blnSkip = True
            Debug.Print "'SKIPPED: " & strModName
          Else
            .MoveLast
            lngProcs = .RecordCount
            .MoveFirst
            arr_varProc = .GetRows(lngProcs)
            ' *******************************************************
            ' ** Array: arr_varProc()
            ' **
            ' **   FIELD  ELEMENT  NAME                  CONSTANT
            ' **   =====  =======  ====================  ==========
            ' **     1       0     dbs_id                P_DID
            ' **     2       1     dbs_name              P_DNAM
            ' **     3       2     vbcom_id              P_CID
            ' **     4       3     vbcom_name            P_CNAM
            ' **     5       4     comtype_type          P_CTYP
            ' **     6       5     vbcomproc_id          P_PID
            ' **     7       6     vbcomproc_name        P_PNAM
            ' **     8       7     vbcomproc_line_beg    P_BEG
            ' **     9       8     vbcomproc_line_end    P_END
            ' **
            ' *******************************************************
          End If
          .Close
        End With  ' ** rst.
        Set rst = Nothing
        Set qdf = Nothing

        Set rst = dbs.OpenRecordset("tblDataTypeVb", dbOpenDynaset, dbReadOnly)
        With rst
          .MoveLast
          lngDataTypes = .RecordCount
          .MoveFirst
          arr_varDataType = .GetRows(lngDataTypes)
          ' *************************************************************
          ' ** Array: arr_varDataType()
          ' **
          ' **   FIELD  ELEMENT  NAME                        CONSTANT
          ' **   =====  =======  ==========================  ==========
          ' **     1       0     datatype_vb_id              D_ID
          ' **     2       1     datatype_vb_type            D_TYPE
          ' **     3       2     datatype_vb_constant        D_CONST
          ' **     4       3     datatype_vb_name            D_NAME
          ' **     5       4     datatype_vb_datemodified    D_DATE
          ' **
          ' *************************************************************
          .Close
        End With
        Set rst = Nothing

        If blnSkip = False Then

          Set cod = .CodeModule
          With cod
            lngModLines = .CountOfLines
            lngModDecLines = .CountOfDeclarationLines
            For lngY = lngModDecLines To lngModLines
              strLine = Trim$(.Lines(lngY, 1))
              If strLine <> vbNullString Then
                strProcName = .ProcOfLine(lngY, vbext_pk_Proc)
                strVarScope = vbNullString: strVarDeclare = vbNullString
                If Left$(strLine, 1) <> "'" And Left$(strLine, 1) <> Chr(34) And _
                    Left$(strLine, 3) <> "#If" And Left$(strLine, 4) <> "#End" Then
                  If Left$(strLine, 4) = "Dim " Or Left$(strLine, 6) = "Const " Or Left$(strLine, 7) = "Static " Then
                    intPos1 = InStr(strLine, " ")
                    If intPos1 > 0 Then
                      If Trim$(Left$(strLine, intPos1)) = "Static" Then
                        strVarScope = Trim$(Left$(strLine, intPos1))
                      Else
                        strVarScope = "Local"
                      End If
                      If Trim$(Left$(strLine, intPos1)) = "Const" Then
                        strVarDeclare = "Constant"
                      Else
                        strVarDeclare = "Variable"
                      End If
                      strTmp02 = Trim$(Mid$(strLine, intPos1))
                      strConstValue = vbNullString
                      intPos2 = InStr(strTmp02, "' **")
                      If intPos2 > 0 Then strTmp02 = Trim$(Left$(strTmp02, (intPos2 - 1)))
                      intPos1 = InStr(strTmp02, ",")
                      If intPos1 > 0 And strVarDeclare <> "Constant" Then
                        Do While intPos1 > 0
                          blnIsArray = False: strArray = vbNullString
                          strTmp03 = Left$(strTmp02, (intPos1 - 1))        ' ** First variable.
                          strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))  ' ** Next variables.
                          intPos2 = InStr(strTmp03, " ")
                          If intPos2 > 0 Then
                            strVarName = Trim$(Left$(strTmp03, intPos2))
                            strVarType = Trim$(Mid$(strTmp03, intPos2))
                            If Left$(strVarType, 3) = "As " Then
                              strVarType = Trim$(Mid$(strVarType, 3))
                              If InStr(strVarName, "(") > 0 Then
                                blnIsArray = True
                                strArray = Trim$(Mid$(strVarName, InStr(strVarName, "(")))
                                strVarName = Trim$(Left$(strVarName, (InStr(strVarName, "(") - 1)))  ' ** Remove array disignation.
                              Else
                                If Left$(strVarName, 4) = "arr_" Then
                                  blnIsArray = True
                                  strArray = "{assigned}"
                                End If
                              End If
                              If InStr(strVarType, "'") > 0 Then _
                                strVarType = Trim$(Left$(strVarType, (InStr(strVarType, "'") - 1)))  ' ** Remove any remarks.
                              blnFound = False: lngF = -1&
                              For lngZ = 0& To (lngProcs - 1&)
                                If lngY >= arr_varProc(P_BEG, lngZ) And lngY <= arr_varProc(P_END, lngZ) Then
                                  blnFound = True
                                  lngF = lngZ
                                  If strProcName <> arr_varProc(P_PNAM, lngZ) Then strProcName = arr_varProc(P_PNAM, lngZ)
                                  Exit For
                                End If
                              Next  ' ** lngZ.
                              Select Case blnFound
                              Case True
                                lngVars = lngVars + 1&
                                lngE = lngVars - 1&
                                ReDim Preserve arr_varVar(V_ELEMS, lngE)
                                arr_varVar(V_DID, lngE) = arr_varProc(P_DID, lngF)
                                arr_varVar(V_DNAM, lngE) = arr_varProc(P_DNAM, lngF)
                                arr_varVar(V_CID, lngE) = arr_varProc(P_CID, lngF)
                                arr_varVar(V_CNAM, lngE) = strModName
                                arr_varVar(V_CTYP, lngE) = arr_varProc(P_CTYP, lngF)
                                arr_varVar(V_PID, lngE) = arr_varProc(P_PID, lngF)
                                arr_varVar(V_PNAM, lngE) = strProcName
                                arr_varVar(V_VNAM, lngE) = strVarName
                                arr_varVar(V_SCOP, lngE) = strVarScope
                                arr_varVar(V_DECL, lngE) = strVarDeclare
                                arr_varVar(V_DTYP, lngE) = CLng(0)
                                arr_varVar(V_VTYP, lngE) = strVarType
                                arr_varVar(V_ISARR, lngE) = blnIsArray
                                arr_varVar(V_ARR, lngE) = strArray
                                arr_varVar(V_CVAL, lngE) = Null
                                arr_varVar(V_LIN, lngE) = lngY
                                arr_varVar(V_CVAL, lngE) = Null
                              Case False
                                Debug.Print "'NOT FOUND! 1  PROC: '" & strProcName & "'  LINE: " & CStr(lngY) & "  MOD: " & strModName
                                Stop
                              End Select
                            Else
                              ' ** The first three characters are not 'As '!
                              Debug.Print "'TYPE? 1  " & strVarType & "  MOD: " & strModName & "  LINE: " & CStr(lngY)
                              Stop
                            End If  ' ** 'As '.
                          Else
                            strVarName = strTmp03
                            Debug.Print "'NO TYPE! 1  " & strVarName & "  MOD: " & strModName & "  PROC: " & strProcName
                            Stop
                          End If  ' ** intPos2: space.
                          intPos1 = InStr(strTmp02, ",")
                        Loop  ' ** intPos1: comma.
                        ' ** Now the last one in the line.
                        intPos2 = InStr(strTmp02, " ")
                        If intPos2 > 0 Then
                          blnIsArray = False: strArray = vbNullString
                          strVarName = Trim$(Left$(strTmp02, intPos2))
                          strVarType = Trim$(Mid$(strTmp02, intPos2))
                          If Left$(strVarType, 3) = "As " Then
                            strVarType = Trim$(Mid$(strVarType, 3))
                            If InStr(strVarName, "(") > 0 Then
                              blnIsArray = True
                              strArray = Trim$(Mid$(strVarName, InStr(strVarName, "(")))
                              strVarName = Trim$(Left$(strVarName, (InStr(strVarName, "(") - 1)))  ' ** Remove array designation.
                            Else
                              If Left$(strVarName, 4) = "arr_" Then
                                blnIsArray = True
                                strArray = "{assigned}"
                              End If
                            End If
                            If InStr(strVarType, "'") > 0 Then _
                              strVarType = Trim$(Left$(strVarType, (InStr(strVarType, "'") - 1)))  ' ** Remove any remarks.
                            blnFound = False: lngF = -1&
                            For lngZ = 0& To (lngProcs - 1&)
                              If lngY >= arr_varProc(P_BEG, lngZ) And lngY <= arr_varProc(P_END, lngZ) Then
                                blnFound = True
                                lngF = lngZ
                                If strProcName <> arr_varProc(P_PNAM, lngZ) Then strProcName = arr_varProc(P_PNAM, lngZ)
                                Exit For
                              End If
                            Next  ' ** lngZ.
                            Select Case blnFound
                            Case True
                              lngVars = lngVars + 1&
                              lngE = lngVars - 1&
                              ReDim Preserve arr_varVar(V_ELEMS, lngE)
                              arr_varVar(V_DID, lngE) = arr_varProc(P_DID, lngF)
                              arr_varVar(V_DNAM, lngE) = arr_varProc(P_DNAM, lngF)
                              arr_varVar(V_CID, lngE) = arr_varProc(P_CID, lngF)
                              arr_varVar(V_CNAM, lngE) = strModName
                              arr_varVar(V_CTYP, lngE) = arr_varProc(P_CTYP, lngF)
                              arr_varVar(V_PID, lngE) = arr_varProc(P_PID, lngF)
                              arr_varVar(V_PNAM, lngE) = strProcName
                              arr_varVar(V_VNAM, lngE) = strVarName
                              arr_varVar(V_SCOP, lngE) = strVarScope
                              arr_varVar(V_DECL, lngE) = strVarDeclare
                              arr_varVar(V_DTYP, lngE) = CLng(0)
                              arr_varVar(V_VTYP, lngE) = strVarType
                              arr_varVar(V_ISARR, lngE) = blnIsArray
                              arr_varVar(V_ARR, lngE) = strArray
                              arr_varVar(V_CVAL, lngE) = Null
                              arr_varVar(V_LIN, lngE) = lngY
                            Case False
                              Debug.Print "'NOT FOUND! 2  PROC: '" & strProcName & "'  LINE: " & CStr(lngY) & "  MOD: " & strModName
                              Stop
                            End Select
                          Else
                            ' ** The first three characters are not 'As '!
                            Debug.Print "'TYPE? 2  " & strVarType & "  MOD: " & strModName & "  LINE: " & CStr(lngY)
                            Stop
                          End If
                        Else
                          strVarName = strTmp02
                          Debug.Print "'NO TYPE! 2  " & strVarName & "  MOD: " & strModName & "  PROC: " & strProcName
                          Stop
                        End If  ' ** intPos2: space.
                      Else
                        ' ** Single variable or constant on the line.
                        blnIsArray = False: strArray = vbNullString
                        If InStr(strTmp02, "(") > 0 And InStr(strTmp02, ")") > 0 Then
                          ' ** It's a dimensioned array.
                          blnIsArray = True
                          strArray = Mid$(strTmp02, InStr(strTmp02, "("), ((InStr(strTmp02, ")") - InStr(strTmp02, "(")) + 1))
                          strTmp02 = Left$(strTmp02, (InStr(strTmp02, "(") - 1)) & Mid$(strTmp02, (InStr(strTmp02, ")") + 1))
                        End If
                        intPos2 = InStr(strTmp02, " ")
                        If intPos2 > 0 Then
                          strVarName = Trim$(Left$(strTmp02, intPos2))
                          strVarType = Trim$(Mid$(strTmp02, intPos2))
                          If Left$(strVarType, 3) = "As " Then
                            strVarType = Trim$(Mid$(strVarType, 3))  ' ** Strip the 'As'.
                            If strVarDeclare = "Constant" Then
                              intPos3 = InStr(strVarType, "=")
                              If intPos3 > 0 Then
                                strConstValue = Trim$(Mid$(strVarType, (intPos3 + 1)))
                                strVarType = Trim$(Left$(strVarType, (intPos3 - 1)))
                              Else
                                ' ** No '=' sign after type.
                                Debug.Print "'WHAT? 5  " & strVarType & "  MOD: " & strModName & "  LINE: " & CStr(lngY)
                                Stop
                              End If
                            End If  ' ** Const.
                            If InStr(strVarName, "(") > 0 Then
                              blnIsArray = True
                              strArray = Trim$(Mid$(strVarName, InStr(strVarName, "(")))
                              strVarName = Trim$(Left$(strVarName, (InStr(strVarName, "(") - 1)))  ' ** Remove array designation.
                            Else
                              If Left$(strVarName, 4) = "arr_" Then
                                blnIsArray = True
                                strArray = "{assigned}"
                              End If
                            End If
                            If InStr(strVarType, "'") > 0 Then _
                              strVarType = Trim$(Left$(strVarType, (InStr(strVarType, "'") - 1)))  ' ** Remove any remarks.
                            blnFound = False: lngF = -1&
                            For lngZ = 0& To (lngProcs - 1&)
                              If lngY >= arr_varProc(P_BEG, lngZ) And lngY <= arr_varProc(P_END, lngZ) Then
                                blnFound = True
                                lngF = lngZ
                                If strProcName <> arr_varProc(P_PNAM, lngZ) Then strProcName = arr_varProc(P_PNAM, lngZ)
                                Exit For
                              End If
                            Next  ' ** lngZ.
                            Select Case blnFound
                            Case True
                              lngVars = lngVars + 1&
                              lngE = lngVars - 1&
                              ReDim Preserve arr_varVar(V_ELEMS, lngE)
                              arr_varVar(V_DID, lngE) = arr_varProc(P_DID, lngF)
                              arr_varVar(V_DNAM, lngE) = arr_varProc(P_DNAM, lngF)
                              arr_varVar(V_CID, lngE) = arr_varProc(P_CID, lngF)
                              arr_varVar(V_CNAM, lngE) = strModName
                              arr_varVar(V_CTYP, lngE) = arr_varProc(P_CTYP, lngF)
                              arr_varVar(V_PID, lngE) = arr_varProc(P_PID, lngF)
                              arr_varVar(V_PNAM, lngE) = strProcName
                              arr_varVar(V_VNAM, lngE) = strVarName
                              arr_varVar(V_SCOP, lngE) = strVarScope
                              arr_varVar(V_DECL, lngE) = strVarDeclare
                              arr_varVar(V_DTYP, lngE) = CLng(0)
                              arr_varVar(V_VTYP, lngE) = strVarType
                              arr_varVar(V_ISARR, lngE) = blnIsArray
                              arr_varVar(V_ARR, lngE) = strArray
                              If strVarDeclare = "Constant" Then
                                arr_varVar(V_CVAL, lngE) = strConstValue
                              Else
                                arr_varVar(V_CVAL, lngE) = Null
                              End If
                              arr_varVar(V_LIN, lngE) = lngY
                            Case False
                              Debug.Print "'NOT FOUND! 3  PROC: '" & strProcName & "'  LINE: " & CStr(lngY) & "  MOD: " & strModName
                              Stop
                            End Select  ' ** blnFound.
                          Else
                            ' ** The first three characters are not 'As '!
                            Debug.Print "'TYPE? 3  " & strVarType & "  MOD: " & strModName & "  LINE: " & CStr(lngY)
                            Stop
                          End If  ' ** 'As '.
                        Else
                          strVarName = strTmp02
                          Debug.Print "'NO TYPE! 3  " & strVarName & "  MOD: " & strModName & "  PROC: " & strProcName
                          Stop
                        End If  ' ** intPos2: space.
                      End If  ' ** intPos1: comma.
                    End If  ' ** Pos1: space.
                  End If  ' ** Dim, Const, Static.
                End If  ' ** Remark.
              End If  ' ** vbNullString.
              If lngLoopCnt > lngLastLoopCnt And (lngLoopCnt Mod 100&) = 0 Then
                lngLastLoopCnt = lngLoopCnt
                Stop
              End If
            Next  ' ** lngY.
          End With  ' ** cod.

          If lngVars > 0& Then

            For lngY = 0& To (lngVars - 1&)
              For lngZ = 0& To (lngDataTypes - 1&)
                strTmp01 = arr_varVar(V_VTYP, lngY)
                If Left$(strTmp01, 8) = "String *" Then  ' ** String * 255
                  strTmp01 = "String"
                ElseIf strTmp01 = "VbMsgBoxResult" Then
                  strTmp01 = "Integer"
                ElseIf InStr(strTmp01, ".") > 0 Then  ' ** DAO.Database
                  strTmp01 = "Object"
                Else
                  Select Case strTmp01
                  Case "VBProject", "VBComponent", "CodeModule", "Module", "CurrentProject", _
                      "Reference", "Collection", "New Collection", "Encrypt"
                    strTmp01 = "Object"
                  Case "clsDevice", "clsDevices", "clsVersionInfo"
                    strTmp01 = "Object"
                  Case "RelBlob", "RelWindow", "WinLink", "adhFileFlags"
                    strTmp01 = "UserDefinedType"
                  Case "Variant"
                    ' ** Leave as is.
                  Case Else
                    Select Case arr_varVar(V_DECL, lngY)
                    Case "Variable"
                      If IsUC(strTmp01, True, True) = True Then
                        ' ** All upper-case, and may include underscores and numerals.
                        strTmp01 = "UserDefinedType"
                      Else
                        ' ** Leave as Variant.
                      End If
                    Case "Constant"
                      ' ** Leave as is (unlikely).
                    End Select
                  End Select
                End If
                If arr_varDataType(D_CONST, lngZ) = ("vb" & strTmp01) Then
                  Select Case arr_varVar(V_ISARR, lngY)
                  Case True
                    ' ** vbArray = 8192, vbArrayByte = 8209 (vbByte = 17: 8192 + 17 = 8209).
                    arr_varVar(V_DTYP, lngY) = (arr_varDataType(D_TYPE, lngZ) + 8192&)
                  Case False
                    arr_varVar(V_DTYP, lngY) = arr_varDataType(D_TYPE, lngZ)
                  End Select
                  Exit For
                End If
              Next  ' ** lngZ.
              If arr_varVar(V_DTYP, lngY) = 0& Then
                arr_varVar(V_DTYP, lngY) = vbVariant
              End If
            Next  ' ** lngY

            Set rst = dbs.OpenRecordset("tblVBComponent_Declaration_Local", dbOpenDynaset, dbAppendOnly)
            With rst
              For lngY = 0& To (lngVars - 1&)
                If arr_varVar(V_VNAM, lngY) = "arr_varTmp01()" Then
                  arr_varVar(V_VNAM, lngY) = Left(arr_varVar(V_VNAM, lngY), (Len(arr_varVar(V_VNAM, lngY)) - 2))
                End If
                .AddNew
'dbs_id
                ![dbs_id] = arr_varVar(V_DID, lngY)
'vbcom_id
                ![vbcom_id] = arr_varVar(V_CID, lngY)
'vbcomproc_id
                ![vbcomproc_id] = arr_varVar(V_PID, lngY)
'vbdecloc_id
                ' ** ![vbdecloc_id] : AutoNumber.
'vbdecloc_module
                ![vbdecloc_module] = strModName
'comtype_type
                ![comtype_type] = arr_varVar(V_CTYP, lngY)
'dectype_type
                ![dectype_type] = arr_varVar(V_DECL, lngY)  ' ** Variable or Constant.
'scopetype_type
                ![scopetype_type] = arr_varVar(V_SCOP, lngY)  ' ** Local or Static.
                intPos1 = InStr(arr_varVar(V_VNAM, lngY), "'")
                If intPos1 > 0 Then arr_varVar(V_VNAM, lngY) = Trim(Left(arr_varVar(V_VNAM, lngY), (intPos1 - 1)))
'vbdecloc_name
                ![vbdecloc_name] = arr_varVar(V_VNAM, lngY)
'datatype_vb_type
                ![datatype_vb_type] = arr_varVar(V_DTYP, lngY)
                intPos1 = InStr(arr_varVar(V_VTYP, lngY), "'")
                If intPos1 > 0 Then arr_varVar(V_VTYP, lngY) = Trim(Left(arr_varVar(V_VTYP, lngY), (intPos1 - 1)))
'vbdecloc_vbtype
                ![vbdecloc_vbtype] = arr_varVar(V_VTYP, lngY)
'vbdecloc_isarray
                ![vbdecloc_isarray] = arr_varVar(V_ISARR, lngY)
'vbdecloc_array
                If arr_varVar(V_ARR, lngY) <> vbNullString Then
                  ![vbdecloc_array] = arr_varVar(V_ARR, lngY)
                Else
                  ![vbdecloc_array] = Null
                End If
'vbdecloc_value
                ![vbdecloc_value] = arr_varVar(V_CVAL, lngY)
'vbdecloc_linenum
                ![vbdecloc_linenum] = arr_varVar(V_LIN, lngY)
'vbdecloc_datemodified
                ![vbdecloc_datemodified] = Now()
                .Update
              Next  ' ** lngY.
              .Close
            End With
            Set rst = Nothing

            lngTotVars = lngTotVars + lngVars
            If lngVars > lngHighVars Then
              lngHighVars = lngVars
              strHighVars = strModName
            End If

          End If  ' ** lngVars.

        End If  ' ** blnSkip.

      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.

  dbs.Close

  Debug.Print "'TOTAL VARS:  " & CStr(lngTotVars)
  Debug.Print "'HIGHEST CNT: " & CStr(lngHighVars) & "  " & strHighVars
  Debug.Print "'DONE!  " & THIS_PROC & "()"

'SKIPPED: modGlobConst
'TOTAL VARS:  31417
'HIGHEST CNT: 820  zz_mod_FormDocFuncs
'DONE!  VBA_LocalVar_Doc()

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_LocalVar_Doc = blnRetVal

End Function

Public Function VBA_UDTypeName_Doc() As Boolean
' ** Document all user-defined type members to tblVBComponent_Declaration_Type.

  Const THIS_PROC As String = "VBA_UDTypeName_Doc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngTypes As Long, arr_varType As Variant
  Dim lngNames As Long, arr_varName() As Variant
  Dim lngDataTypes As Long, arr_varDataType As Variant
  Dim lngModLines As Long, lngModDecLines As Long
  Dim strModName As String, strLastModName As String, strLine As String
  Dim strVarName As String, strVarType As String, strVarValue As String
  Dim blnStart As Boolean, blnEnd As Boolean, blnSave As Boolean
  Dim blnIsArray As Boolean, strArray As String
  Dim lngTotNames As Long, lngHighNames As Long, strHighNames As String
  Dim lngThisDbsID As Long
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp01 As String, strTmp02 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varType().
  Const T_DID  As Integer = 0
  Const T_DNAM As Integer = 1
  Const T_CID  As Integer = 2
  Const T_CNAM As Integer = 3
  Const T_CTYP As Integer = 4
  Const T_TID  As Integer = 5
  Const T_TNAM As Integer = 6
  Const T_SCOP As Integer = 7
  Const T_DECL As Integer = 8
  Const T_TYP1 As Integer = 9
  Const T_TYP2 As Integer = 10
  Const T_UDT  As Integer = 11
  Const T_LIN1 As Integer = 12
  Const T_LIN2 As Integer = 13

  ' ** Array: arr_varName().
  Const N_ELEMS As Integer = 14  ' ** Array's first-element UBound().
  Const N_DID   As Integer = 0
  Const N_DNAM  As Integer = 1
  Const N_CID   As Integer = 2
  Const N_CNAM  As Integer = 3
  Const N_CTYP  As Integer = 4
  Const N_TID   As Integer = 5
  Const N_TNAM  As Integer = 6
  Const N_DECL  As Integer = 7
  Const N_VNAM  As Integer = 8
  Const N_VTYP  As Integer = 9
  Const N_DTYP  As Integer = 10
  Const N_ISARR As Integer = 11
  Const N_ARR   As Integer = 12
  Const N_CVAL  As Integer = 13
  Const N_LIN   As Integer = 14

  ' ** Array: arr_varDataType().
  Const D_ID    As Integer = 0
  Const D_TYPE  As Integer = 1
  Const D_CONST As Integer = 2
  Const D_NAME  As Integer = 3
  Const D_DATE  As Integer = 4

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** Empty tblVBComponent_Declaration_Type.
    Set qdf = .QueryDefs("zz_qry_VBComponent_ProjectName_31")
    qdf.Execute
    Set qdf = Nothing

    ' ** tblVBComponent_Declaration, just 'Enum', 'Type', by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_VBComponent_ProjectName_32")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngTypes = .RecordCount
      .MoveFirst
      arr_varType = .GetRows(lngTypes)
      ' ******************************************************
      ' ** Array: arr_varType()
      ' **
      ' **   FIELD  ELEMENT  NAME                 CONSTANT
      ' **   =====  =======  ===================  ==========
      ' **     1       0     dbs_id               T_DID
      ' **     2       1     dbs_name             T_DNAM
      ' **     3       2     vbcom_id             T_CID
      ' **     4       3     vbcom_name           T_CNAM
      ' **     5       4     comtype_type         T_CTYP
      ' **     6       5     vbdec_id             T_TID
      ' **     7       6     vbdec_name           T_TNAM
      ' **     8       7     scopetype_type       T_SCOP
      ' **     9       8     dectype_type         T_DECL
      ' **    10       9     datatype_vb_type     T_TYP1
      ' **    11      10     vbdec_vbtype         T_TYP2
      ' **    12      11     vbdec_userdefined    T_UDT
      ' **    13      12     vbdec_linenum1       T_LIN1
      ' **    14      13     vbdec_linenum2       T_LIN2
      ' **
      ' ******************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    Set rst = .OpenRecordset("tblDataTypeVb", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngDataTypes = .RecordCount
      .MoveFirst
      arr_varDataType = .GetRows(lngDataTypes)
      ' *************************************************************
      ' ** Array: arr_varDataType()
      ' **
      ' **   FIELD  ELEMENT  NAME                        CONSTANT
      ' **   =====  =======  ==========================  ==========
      ' **     1       0     datatype_vb_id              D_ID
      ' **     2       1     datatype_vb_type            D_TYPE
      ' **     3       2     datatype_vb_constant        D_CONST
      ' **     4       3     datatype_vb_name            D_NAME
      ' **     5       4     datatype_vb_datemodified    D_DATE
      ' **
      ' *************************************************************
      .Close
    End With
    Set rst = Nothing

    Set rst = .OpenRecordset("tblVBComponent_Declaration_Type", dbOpenDynaset, dbConsistent)

    Set vbp = Application.VBE.ActiveVBProject
    With vbp

      strLastModName = vbNullString
      lngTotNames = 0&: lngHighNames = 0&: strHighNames = vbNullString

      For lngX = 0& To (lngTypes - 1&)

        If arr_varType(T_CNAM, lngX) <> strLastModName Then
          Set vbc = Nothing
          Set vbc = .VBComponents(arr_varType(T_CNAM, lngX))
          strLastModName = arr_varType(T_CNAM, lngX)
          lngNames = 0&
          ReDim arr_varName(N_ELEMS, 0)
        End If

        With vbc
          strModName = .Name
          Set cod = .CodeModule
          With cod
            lngModLines = .CountOfLines
            lngModDecLines = .CountOfDeclarationLines
            blnStart = False: blnEnd = False
            For lngY = arr_varType(T_LIN1, lngX) To arr_varType(T_LIN2, lngX)
              strLine = Trim$(.Lines(lngY, 1))
              strVarName = vbNullString: strVarType = vbNullString: strVarValue = vbNullString
              blnIsArray = False: strArray = vbNullString
              If Left$(strLine, 1) <> "'" Then
                Select Case arr_varType(T_DECL, lngX)
                Case "Type"
                  Select Case blnStart
                  Case True
                    strTmp01 = strLine
                    intPos1 = InStr(strTmp01, "'")
                    If intPos1 > 0 Then strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))  ' ** Strip off any remarks.
                    If strTmp01 <> "End Type" Then
                      intPos1 = InStr(strTmp01, " As ")
                      If intPos1 > 0 Then
                        strTmp02 = Trim$(Mid$(strTmp01, (intPos1 + 3)))
                        strTmp01 = Trim(Left$(strTmp01, intPos1))
                        intPos2 = InStr(strTmp01, "(")
                        If intPos2 > 0 Then
                          'prgDayState(11)
                          blnIsArray = True
                          strArray = Mid$(strTmp01, intPos2)
                          strTmp01 = Left$(strTmp01, (intPos2 - 1))
                        End If
                        strVarName = strTmp01
                        strVarType = strTmp02
                        lngNames = lngNames + 1&
                        lngE = lngNames - 1&
                        ReDim Preserve arr_varName(N_ELEMS, lngE)
                        arr_varName(N_DID, lngE) = arr_varType(T_DID, lngX)
                        arr_varName(N_DNAM, lngE) = arr_varType(T_DNAM, lngX)
                        arr_varName(N_CID, lngE) = arr_varType(T_CID, lngX)
                        arr_varName(N_CNAM, lngE) = arr_varType(T_CNAM, lngX)
                        arr_varName(N_CTYP, lngE) = arr_varType(T_CTYP, lngX)
                        arr_varName(N_TID, lngE) = arr_varType(T_TID, lngX)
                        arr_varName(N_TNAM, lngE) = arr_varType(T_TNAM, lngX)
                        arr_varName(N_DECL, lngE) = arr_varType(T_DECL, lngX)
                        arr_varName(N_VNAM, lngE) = strVarName
                        arr_varName(N_VTYP, lngE) = strVarType
                        arr_varName(N_DTYP, lngE) = Null
                        If Left$(strVarType, 8) = "String *" Then
                          strVarType = "String"
                        End If
                        For lngZ = 0& To (lngDataTypes - 1&)
                          If arr_varDataType(D_CONST, lngZ) = ("vb" & strVarType) Then
                            Select Case blnIsArray
                            Case True
                              ' ** vbArray = 8192, vbArrayByte = 8209 (vbByte = 17: 8192 + 17 = 8209).
                              arr_varName(N_DTYP, lngE) = (arr_varDataType(D_TYPE, lngZ) + 8192&)
                            Case False
                              arr_varName(N_DTYP, lngE) = arr_varDataType(D_TYPE, lngZ)
                            End Select
                            Exit For
                          End If
                        Next  ' ** lngZ.
                        arr_varName(N_ISARR, lngE) = blnIsArray
                        Select Case blnIsArray
                        Case True
                          arr_varName(N_ARR, lngE) = strArray
                        Case False
                          arr_varName(N_ARR, lngE) = Null
                        End Select
                        arr_varName(N_CVAL, lngE) = Null
                        arr_varName(N_LIN, lngE) = lngY
                      Else
                        ' ** No type?
                        strVarType = "Variant"
                        Stop
                      End If
                    Else
                      blnEnd = True
                    End If
                  Case False
                    ' ** Public Type BROWSEINFO
                    If InStr(strLine, "Type ") > 0 And InStr(strLine, arr_varType(T_TNAM, lngX)) > 0 Then
                      blnStart = True
                    Else
                      Stop
                    End If
                  End Select
                Case "Enum"
                  ' ** Enumerations are Constants, and my examples are all untyped.
                  Select Case blnStart
                  Case True
                    strTmp01 = strLine
                    intPos1 = InStr(strTmp01, "'")
                    intPos2 = InStr(strTmp01, Chr(34))
                    If intPos1 > 0 Then
                      If intPos2 > 0 Then
                        ' ** See if the single quote is within quotation marks.
                        If intPos1 < intPos2 Then
                          ' ** Remark is before the quotes.
                          strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))  ' ** Strip off any remarks.
                        Else
                          If InStr((intPos2 + 1), strTmp01, Chr(34)) < intPos1 Then
                            ' ** Remark is after quotes.
                            strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))  ' ** Strip off any remarks.
                          Else
                            ' ** Single quote appears to be within the constant declaration.
                          End If
                        End If
                      Else
                        strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))  ' ** Strip off any remarks.
                      End If
                    End If
                    If strLine <> "End Enum" Then
                      intPos1 = InStr(strTmp01, " As ")
                      intPos2 = InStr(strTmp01, "(")
                      If intPos1 > 0 Then
                        strTmp02 = Trim$(Mid$(strTmp01, (intPos1 + 3)))
                        strTmp01 = Trim$(Left$(strTmp01, intPos1))
                        If intPos2 > 0 Then
                          blnIsArray = True
                          strArray = Mid$(strTmp01, intPos2)  ' ** Trimming would have only been to the right side.
                          strTmp01 = Left$(strTmp01, (intPos2 - 1))
                        End If
                        intPos1 = InStr(strTmp02, "=")
                        If intPos1 > 0 Then
                          strVarType = Trim$(Left$(strTmp02, (intPos1 - 1)))
                          strVarValue = Trim$(Mid$(strTmp02, (intPos1 + 1)))
                        Else
                          ' ** No value?
                          strVarType = strTmp02
                        End If
                        strVarName = strTmp01
                      Else
                        ' ** None of mine are typed.
                        strVarType = "Variant"
                        intPos1 = InStr(strTmp01, "=")
                        If intPos1 > 0 Then
                          strTmp02 = Trim$(Mid$(strTmp01, (intPos1 + 1)))
                          strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))
                          If intPos2 > 0 Then
                            blnIsArray = True
                            strArray = Mid$(strTmp01, intPos2)  ' ** Trimming would have only been to the right side.
                            strTmp01 = Left$(strTmp01, (intPos2 - 1))
                          End If
                          strVarName = strTmp01
                          strVarValue = strTmp02
                        Else
                          ' ** No type, no value?
                          strVarName = strTmp01
                          Stop
                        End If
                      End If
                      lngNames = lngNames + 1&
                      lngE = lngNames - 1&
                      ReDim Preserve arr_varName(N_ELEMS, lngE)
                      arr_varName(N_DID, lngE) = arr_varType(T_DID, lngX)
                      arr_varName(N_DNAM, lngE) = arr_varType(T_DNAM, lngX)
                      arr_varName(N_CID, lngE) = arr_varType(T_CID, lngX)
                      arr_varName(N_CNAM, lngE) = arr_varType(T_CNAM, lngX)
                      arr_varName(N_CTYP, lngE) = arr_varType(T_CTYP, lngX)
                      arr_varName(N_TID, lngE) = arr_varType(T_TID, lngX)
                      arr_varName(N_TNAM, lngE) = arr_varType(T_TNAM, lngX)
                      arr_varName(N_DECL, lngE) = arr_varType(T_DECL, lngX)
                      arr_varName(N_VNAM, lngE) = strVarName
                      If strVarType = vbNullString Then strVarType = "Variant"
                      arr_varName(N_VTYP, lngE) = strVarType
                      arr_varName(N_DTYP, lngE) = Null
                      If Left$(strVarType, 8) = "String *" Then
                        strVarType = "String"
                      End If
                      For lngZ = 0& To (lngDataTypes - 1&)
                        If arr_varDataType(D_CONST, lngZ) = ("vb" & strVarType) Then
                          Select Case blnIsArray
                          Case True
                            ' ** vbArray = 8192, vbArrayByte = 8209 (vbByte = 17: 8192 + 17 = 8209).
                            arr_varName(N_DTYP, lngE) = (arr_varDataType(D_TYPE, lngZ) + 8192&)
                          Case False
                            arr_varName(N_DTYP, lngE) = arr_varDataType(D_TYPE, lngZ)
                          End Select
                          Exit For
                        End If
                      Next  ' ** lngZ.
                      arr_varName(N_ISARR, lngE) = blnIsArray
                      Select Case blnIsArray
                      Case True
                        arr_varName(N_ARR, lngE) = strArray
                      Case False
                        arr_varName(N_ARR, lngE) = Null
                      End Select
                      If strVarValue = vbNullString Then
                        arr_varName(N_CVAL, lngE) = Null
                      Else
                        arr_varName(N_CVAL, lngE) = strVarValue
                      End If
                      arr_varName(N_LIN, lngE) = lngY
                    Else
                      blnEnd = True
                    End If
                  Case False
                    ' ** Public Enum RegRoot
                    If InStr(strLine, "Enum ") > 0 And InStr(strLine, arr_varType(T_TNAM, lngX)) > 0 Then
                      blnStart = True
                    Else
                      Stop
                    End If
                  End Select
                End Select
                If blnEnd = True Then
                  ' ** Though the count should be finished anyway.
                  Exit For
                End If
              End If  ' ** Remark.
            Next  ' ** lngY.
          End With  ' ** cod.
        End With  ' ** vbc.

        blnSave = False
        If lngX < (lngTypes - 1&) Then
          If strModName <> arr_varType(T_CNAM, lngX + 1&) Then
            blnSave = True
          End If
        Else
          ' ** Last Type.
          blnSave = True
        End If

        If blnSave = True Then
          For lngY = 0& To (lngNames - 1&)
            With rst
              .AddNew
              ![dbs_id] = arr_varName(N_DID, lngY)
              ![vbcom_id] = arr_varName(N_CID, lngY)
              ![vbdec_id] = arr_varName(N_TID, lngY)
              ' ** ![vbdectype_id] : AutoNumber.
              ![vbdec_module] = arr_varName(N_CNAM, lngY)
              ![comtype_type] = arr_varName(N_CTYP, lngY)
              ![vbdec_name] = arr_varName(N_TNAM, lngY)
              ![dectype_type] = arr_varName(N_DECL, lngY)
              ![vbdectype_name] = arr_varName(N_VNAM, lngY)
              Select Case IsNull(arr_varName(N_DTYP, lngY))
              Case True
                ![datatype_vb_type] = vbUserDefinedType
              Case False
                ![datatype_vb_type] = arr_varName(N_DTYP, lngY)
              End Select
              ![vbdectype_vbtype] = arr_varName(N_VTYP, lngY)
              ![vbdectype_isarray] = arr_varName(N_ISARR, lngY)
              Select Case arr_varName(N_ISARR, lngY)
              Case True
                Select Case IsNull(arr_varName(N_ARR, lngY))
                Case True
                  ![vbdectype_array] = Null
                Case False
                  If arr_varName(N_ARR, lngY) = vbNullString Then
                    ![vbdectype_array] = Null
                  Else
                    ![vbdectype_array] = arr_varName(N_ARR, lngY)
                  End If
                End Select
              Case False
                ![vbdectype_array] = Null
              End Select
              If arr_varName(N_DECL, lngY) = "Enum" Then
                Select Case IsNull(arr_varName(N_CVAL, lngY))
                Case True
                  ![vbdectype_value] = Null
                Case False
                  ![vbdectype_value] = arr_varName(N_CVAL, lngY)
                End Select
              Else
                ![vbdectype_value] = Null
              End If
              ![vbdectype_linenum] = arr_varName(N_LIN, lngY)
              ![vbdectype_datemodified] = Now()
              .Update
              lngTotNames = lngTotNames + 1&
            End With
          Next  ' ** lngY.
          If lngNames > lngHighNames Then
            lngHighNames = lngNames
            strHighNames = arr_varType(T_TNAM, lngX)
          End If
        End If  ' ** blnSave.

      Next  ' ** lngX.

      rst.Close

    End With  ' ** vbp.

    .Close
  End With  ' ** dbs.

  Debug.Print "'TOTAL NAMES:  " & CStr(lngTotNames)
  Debug.Print "'HIGHEST CNT: " & CStr(lngHighNames) & "  " & strHighNames
  Debug.Print "'DONE!  " & THIS_PROC & "()"

'TOTAL NAMES:  461
'HIGHEST CNT: 134  WININFO_TYPE
'DONE!  VBA_UDTypeName_Doc()

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_UDTypeName_Doc = blnRetVal

End Function

Public Function VBA_CodeLine_Doc() As Boolean
' ** Document all code line numbers to tblVBComponent_CodeLine.
' ** tblVBComponent_Procedure must be up-to-date!

  Const THIS_PROC As String = "VBA_CodeLine_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngX As Long, sngY As Single
  Dim blnRetVal As Boolean

  ' ** Array: arr_varProc().
  Const P_DID As Integer = 0
  Const P_CID As Integer = 1
  Const P_PID As Integer = 2
  Const P_BEG As Integer = 3
  Const P_END As Integer = 4

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    ' ** Empty tblVBComponent_CodeLine.
    Set qdf = .QueryDefs("zz_qry_VBComponent_ProjectName_29a")
    qdf.Execute
    Set qdf = Nothing

    ' ** zz_qry_VBComponent_ProjectName_21a (tblVBComponent_Procedure, with vbcomproc_code_cnt), just needed fields.
    Set qdf = .QueryDefs("zz_qry_VBComponent_ProjectName_22a")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngProcs = .RecordCount
      .MoveFirst
      arr_varProc = .GetRows(lngProcs)
      ' *******************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   FIELD  ELEMENT  NAME                  CONSTANT
      ' **   =====  =======  ====================  ==========
      ' **     1       0     dbs_id                P_DID
      ' **     2       1     vbcom_id              P_CID
      ' **     3       2     vbcomproc_id          P_PID
      ' **     4       3     vbcomproc_code_beg    P_BEG
      ' **     5       4     vbcomproc_code_end    P_END
      ' **
      ' *******************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    Set rst = .OpenRecordset("tblVBComponent_CodeLine", dbOpenDynaset, dbAppendOnly)
    With rst
      For lngX = 0& To (lngProcs - 1&)
        If IsNull(arr_varProc(P_BEG, lngX)) = False And IsNull(arr_varProc(P_END, lngX)) = False Then
          For sngY = CSng(arr_varProc(P_BEG, lngX)) To CSng(arr_varProc(P_END, lngX)) Step 10
            .AddNew
            ![dbs_id] = arr_varProc(P_DID, lngX)
            ![vbcom_id] = arr_varProc(P_CID, lngX)
            ![vbcomproc_id] = arr_varProc(P_PID, lngX)
            ' ** ![vbcode_id] : AutoNumber.
            ![vbcode_codeline] = CStr(sngY)
            ![vbcode_datemodified] = Now()
            .Update
          Next  ' ** sngY.
        End If  ' ** IsNull().
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.

    .Close
  End With  ' ** dbs.

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_CodeLine_Doc = blnRetVal

End Function

Public Function VBA_LineName_Doc() As Boolean
' ** Document all line names to tblVBComponent_LineName.

  Const THIS_PROC As String = "VBA_LineName_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngModLines As Long, lngModDecLines As Long
  Dim strModName As String, strProcName As String, strLine As String
  Dim lngNames As Long, arr_varName() As Variant
  Dim lngThisDbsID As Long
  Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varName().
  Const N_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const N_DID  As Integer = 0
  Const N_CID  As Integer = 1
  Const N_CNAM As Integer = 2
  Const N_PID  As Integer = 3
  Const N_PNAM As Integer = 4
  Const N_LNAM As Integer = 5
  Const N_LINE As Integer = 6

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    ' ** Empty tblVBComponent_LineName.
    Set qdf = .QueryDefs("zz_qry_VBComponent_ProjectName_29b")
    qdf.Execute
    Set qdf = Nothing

    Set rst = dbs.OpenRecordset("tblVBComponent_LineName", dbOpenDynaset, dbAppendOnly)

  End With

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    strModName = vbNullString
    For Each vbc In .VBComponents
      With vbc

        strModName = .Name
        strProcName = vbNullString

        lngNames = 0&
        ReDim arr_varName(N_ELEMS, 0)

        Set cod = .CodeModule
        With cod
          lngModLines = .CountOfLines
          lngModDecLines = .CountOfDeclarationLines
          For lngX = lngModDecLines To lngModLines
            strLine = .Lines(lngX, 1)  ' ** Don't trim!
            If strLine <> vbNullString Then
              If Left$(Trim$(strLine), 1) <> "'" Then
                If InStr(strLine, " ") = 0 Then
                  ' ** I don't think I've ever put a remark on one of these.
                  If IsNumeric(Left$(strLine, 1)) = False Then
                    If Right$(strLine, 1) = ":" Then
                      strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                      lngNames = lngNames + 1&
                      lngE = lngNames - 1&
                      ReDim Preserve arr_varName(N_ELEMS, lngE)
                      arr_varName(N_DID, lngE) = lngThisDbsID
                      arr_varName(N_CID, lngE) = CLng(0)
                      arr_varName(N_CNAM, lngE) = strModName
                      arr_varName(N_PID, lngE) = CLng(0)
                      arr_varName(N_PNAM, lngE) = strProcName
                      arr_varName(N_LNAM, lngE) = strLine
                      arr_varName(N_LINE, lngE) = lngX
                    End If
                  End If  ' ** IsNumeric().
                Else
                  ' ** But just in case...
                  If InStr(strLine, "'") > 0 Then
                    strTmp01 = Trim$(Left$(strLine, (InStr(strLine, "'") - 1)))
                    If InStr(strTmp01, " ") = 0 Then
                      If IsNumeric(Left$(strTmp01, 1)) = False Then
                        If Right$(strTmp01, 1) = ":" Then
                          strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                          lngNames = lngNames + 1&
                          lngE = lngNames - 1&
                          ReDim Preserve arr_varName(N_ELEMS, lngE)
                          arr_varName(N_DID, lngE) = lngThisDbsID
                          arr_varName(N_CID, lngE) = CLng(0)
                          arr_varName(N_CNAM, lngE) = strModName
                          arr_varName(N_PID, lngE) = CLng(0)
                          arr_varName(N_PNAM, lngE) = strProcName
                          arr_varName(N_LNAM, lngE) = strLine
                          arr_varName(N_LINE, lngE) = lngX
                        End If
                      End If
                    End If
                  End If
                End If
              End If  ' ** Remark.
            End If  ' ** vbNullString.
          Next  ' ** lngX
        End With  ' ** cod.

        If lngNames > 0& Then

          For lngX = 0& To (lngNames - 1&)
            varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
              "[vbcom_name] = '" & arr_varName(N_CNAM, lngX) & "'")
            If IsNull(varTmp00) = False Then
              arr_varName(N_CID, lngX) = CLng(varTmp00)
If Left$(strModName, 3) <> "cls" Then
              varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
                "[vbcom_id] = " & CStr(arr_varName(N_CID, lngX)) & " And " & _
                "[vbcomproc_name] = '" & arr_varName(N_PNAM, lngX) & "'")
              If IsNull(varTmp00) = False Then
                arr_varName(N_PID, lngX) = CLng(varTmp00)
              Else
                Stop
              End If
Else
  ' ** I'm having trouble getting the right vbcomproc_id!
  varTmp00 = DCount("*", "tblVBComponent_Procedure", "[vbcom_id] = " & CStr(arr_varName(N_CID, lngX)) & " And " & _
    "[vbcomproc_name] = '" & arr_varName(N_PNAM, lngX) & "'")
  If IsNull(varTmp00) = False Then
    If varTmp00 > 1 Then
      lngTmp02 = CLng(varTmp00)
      varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
        "[vbcom_id] = " & CStr(arr_varName(N_CID, lngX)) & " And " & _
        "[vbcomproc_name] = '" & arr_varName(N_PNAM, lngX) & "'")
      If IsNull(varTmp00) = False Then
        lngTmp03 = CLng(varTmp00)
        varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
          "[vbcom_id] = " & CStr(arr_varName(N_CID, lngX)) & " And " & _
          "[vbcomproc_name] = '" & arr_varName(N_PNAM, lngX) & "' And " & _
          "[vbcomproc_id] <> " & CStr(lngTmp03))
        If IsNull(varTmp00) = False Then
          lngTmp04 = CLng(varTmp00)
          If lngTmp02 = 2& Then
            ' ** lngTmp03 and lngTmp04 should be different procedures with the same name, a Get and a Let.
            varTmp00 = DLookup("[vbcomproc_line_beg]", "tblVBComponent_Procedure", "[vbcomproc_id] = " & CStr(lngTmp03))
            If arr_varName(N_LINE, lngX) > varTmp00 Then
              varTmp00 = DLookup("[vbcomproc_line_beg]", "tblVBComponent_Procedure", "[vbcomproc_id] = " & CStr(lngTmp04))
              If arr_varName(N_LINE, lngX) > varTmp00 Then
                arr_varName(N_PID, lngX) = CLng(lngTmp04)
              Else
                arr_varName(N_PID, lngX) = CLng(lngTmp03)
              End If
            Else
              varTmp00 = DLookup("[vbcomproc_line_beg]", "tblVBComponent_Procedure", "[vbcomproc_id] = " & CStr(lngTmp04))
              If arr_varName(N_LINE, lngX) > varTmp00 Then
                arr_varName(N_PID, lngX) = CLng(lngTmp04)
              Else
                Stop
              End If
            End If
          Else
            Stop
          End If
        Else
          Stop
        End If
      Else
        Stop
      End If
    ElseIf varTmp00 = 1 Then
      varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
        "[vbcom_id] = " & CStr(arr_varName(N_CID, lngX)) & " And " & _
        "[vbcomproc_name] = '" & arr_varName(N_PNAM, lngX) & "'")
      If IsNull(varTmp00) = False Then
        arr_varName(N_PID, lngX) = CLng(varTmp00)
      Else
        Stop
      End If
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
          Next  ' ** lngX.

          With rst
            For lngX = 0& To (lngNames - 1&)
              .AddNew
              ![dbs_id] = arr_varName(N_DID, lngX)
              ![vbcom_id] = arr_varName(N_CID, lngX)
              ![vbcomproc_id] = arr_varName(N_PID, lngX)
              ' ** ![vbline_id] : AutoNumber.
              ![vbline_name] = arr_varName(N_LNAM, lngX)
              ![vbline_linenum] = arr_varName(N_LINE, lngX)
              ![vbline_datemodified] = Now()
              .Update
            Next  ' ** lngX
          End With  ' ** rst.

        End If  ' ** lngNames.

      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.

  rst.Close
  dbs.Close

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_LineName_Doc = blnRetVal

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
  Const P_CID  As Integer = 1
  Const P_CNAM As Integer = 2
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
  Const A_CNAM As Integer = 0
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
'    Left(arr_varProc(P_CNAM, (lngProcs - 1&)), 2) <> "zz" Then
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
              ' **     2       1     vbcom_id                P_CID
              ' **     3       2     vbcom_name              P_CNAM
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
              arr_varProc(P_CID, lngE) = lngVBComID
              arr_varProc(P_CNAM, lngE) = strModName
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
                .FindFirst "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varProc(P_CID, lngX)) & " And " & _
                  "[vbcomproc_name] = '" & arr_varProc(P_PNAM, lngX) & "'"
              Else
                .FindFirst "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varProc(P_CID, lngX)) & " And " & _
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
              ![vbcom_id] = arr_varProc(P_CID, lngX)
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
              arr_varAdd1(A_CNAM, lngE) = strModName
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
                        "[vbcom_id] = " & CStr(arr_varProc(P_CID, lngX)) & " And " & _
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
                          "[vbcom_id] = " & CStr(arr_varProc(P_CID, lngX)) & " And " & _
                          "[vbcomproc_id] = " & CStr(arr_varProc(P_PID, lngX)))
                        If IsNull(varTmp00) = False Then
                          lngTmp02 = varTmp00
                          ' ** See if another already has the new order.
                          varTmp00 = DLookup("[vbcomparam_id]", "tblVBComponent_Procedure_Parameter", _
                            "[dbs_id] = " & CStr(arr_varProc(P_DID, lngX)) & " And " & _
                            "[vbcom_id] = " & CStr(arr_varProc(P_CID, lngX)) & " And " & _
                            "[vbcomproc_id] = " & CStr(arr_varProc(P_PID, lngX)) & " And " & _
                            "[vbcomparam_order] = " & CStr(arr_varParm(M_ORD, lngY)))
                          If IsNull(varTmp00) = True Then
                            ' ** We're in luck! (Or maybe it was moved aside earlier.)
                            ![vbcomparam_order] = arr_varParm(M_ORD, lngY)
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
                      ![vbcom_id] = arr_varProc(P_CID, lngX)
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
        DoEvents
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
      Debug.Print "'  NEW: " & arr_varAdd1(A_PNAM, lngX) & "  IN  " & arr_varAdd1(A_CNAM, lngX) & "  " & arr_varAdd1(A_PTYP, lngX) & _
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

Public Function VBA_ProcCodeLine_Check() As Boolean

  Const THIS_PROC As String = "VBA_ProcCodeLine_Check"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngMods As Long, arr_varMod As Variant
  Dim strModName As String, strLine As String
  Dim strFirstCodeLine As String, strLastCodeLine As String
  Dim lngModLines As Long, lngModDecLines As Long
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim lngEdits As Long, blnEdited As Boolean
  Dim intPos1 As Integer
  Dim lngX As Long, lngY As Long, lngZ As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varMod().
  Const M_DID  As Integer = 0
  Const M_CID  As Integer = 1
  Const M_CNAM As Integer = 2
  Const M_CNT  As Integer = 3

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    ' ** zz_qry_VBComponent_ProjectName_22a (zz_qry_VBComponent_ProjectName_21a
    ' ** (tblVBComponent_Procedure, with vbcomproc_code_cnt), just needed fields),
    ' ** grouped by vbcom_id, just those missing code line numbers, with cnt.
    Set qdf = .QueryDefs("zz_qry_VBComponent_ProjectName_22b")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngMods = .RecordCount
      .MoveFirst
      arr_varMod = .GetRows(lngMods)
      ' ***********************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   FIELD  ELEMENT  NAME          CONSTANT
      ' **   =====  =======  ============  ==========
      ' **     1       0     dbs_id        M_DID
      ' **     2       1     vbcom_id      M_CID
      ' **     3       2     vbcom_name    M_CNAM
      ' **     4       3     cnt           M_CNT
      ' **
      ' ***********************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    lngEdits = 0&

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For lngX = 0& To (lngMods - 1&)
        If VBA_HasCodeLines(arr_varMod(M_CID, lngX)) = True Then  ' ** Module Function: zz_mod_ModuleMiscFuncs.
          ' ** tblVBComponent_Procedure, by specified [vbcid].
          Set qdf = dbs.QueryDefs("zz_qry_VBComponent_ProjectName_22c")
          With qdf.Parameters
            ![vbcid] = arr_varMod(M_CID, lngX)
          End With
          Set rst = qdf.OpenRecordset
          rst.MoveLast
          lngRecs = rst.RecordCount
          rst.MoveFirst
          Set vbc = .VBComponents(arr_varMod(M_CNAM, lngX))
          With vbc
            strModName = .Name
            Set cod = .CodeModule
            With cod
              lngModLines = .CountOfLines
              lngModDecLines = .CountOfDeclarationLines
              For lngY = 1& To lngRecs
                strFirstCodeLine = vbNullString: strLastCodeLine = vbNullString
                blnEdited = False
                For lngZ = rst![vbcomproc_line_beg] To rst![vbcomproc_line_end]
                  strLine = Trim(.Lines(lngZ, 1))
                  If strLine <> vbNullString Then
                    If Left$(strLine, 1) <> "'" Then
                      intPos1 = InStr(strLine, " ")
                      If intPos1 > 0 Then
                        strLine = Trim$(Left$(strLine, intPos1))
                        If IsNumeric(strLine) = True Then
                          If strFirstCodeLine = vbNullString Then
                            strFirstCodeLine = strLine
                          End If
                          strLastCodeLine = strLine
                        End If  ' ** IsNumeric().
                      End If  ' ** intPos1.
                    End If  ' ** Remark.
                  End If  ' ** vbNullString.
                Next  ' ** lngZ
                If strFirstCodeLine <> vbNullString Then
                  With rst
                    .Edit
                    ![vbcomproc_code_beg] = strFirstCodeLine
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdited = True
                  End With
                End If
                If strLastCodeLine <> vbNullString Then
                  With rst
                    .Edit
                    ![vbcomproc_code_end] = strLastCodeLine
                    ![vbcomproc_datemodified] = Now()
                    .Update
                    blnEdited = True
                  End With
                End If
                If blnEdited = True Then
                  lngEdits = lngEdits + 1&
                End If
                If lngY < lngRecs Then rst.MoveNext
              Next  ' ** lngY.
            End With  ' ** cod.
          End With  ' ** vbc.
          Set vbc = Nothing
          rst.Close
          Set rst = Nothing
          Set qdf = Nothing
        End If  ' ** VBA_HasCodeLines().
      Next  ' ** lngX.
    End With  ' ** vbp.

    .Close
  End With  ' ** dbs.

  If lngEdits > 0& Then
    Debug.Print "'PROCS EDITED: " & CStr(lngEdits)
  Else
    Debug.Print "'NO CHANGES!"
  End If

  Debug.Print "'DONE!  " & THIS_PROC & "()"

'PROCS EDITED: 8367
'DONE!  VBA_ProcCodeLine_Check()

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_ProcCodeLine_Check = blnRetVal

End Function

Public Function DatasheetCol() As Boolean

  Const THIS_PROC As String = "VBA_ProcCodeLine_Check"

  Dim ctl As Control, prp As Property
  Dim lngCtls As Long
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  'With Screen.ActiveDatasheet
  '  lngX = 0&
  '  For Each prp In .Properties
  '    With prp
  '      lngX = lngX + 1&
  '      Debug.Print "'" & Left(CStr(lngX) & "." & "   ", 4) & " " & .Name
  '    End With
  '    If lngX = 100& Then
  '      Stop
  '    End If
  '  Next
  'End With

  'With Screen.ActiveControl
  '  lngX = 0&
  '  For Each prp In .Properties
  '    With prp
  '      lngX = lngX + 1&
  '      Debug.Print "'" & Left(CStr(lngX) & "." & "   ", 4) & " " & .Name
  '    End With
  '    If lngX = 100& Then
  '      Stop
  '    End If
  '  Next
  'End With

  With Screen.ActiveDatasheet
    lngCtls = .Controls.Count
    For lngX = 0& To (lngCtls - 1&)
      Set ctl = .Controls(lngX)
      ctl.Properties("ColumnOrder") = (lngX + 1&)
    Next
  End With

'QRY: qryAccountProfile_Transactions_01  FLDS: 18
'journalno
'JournalType
'JournalType_Order
'accountno
'shortname
'legalname
'assetno
'totdesc
'totdc
'transdate
'shareface
'pershare
'icash
'pcash
'cost
'assetdate
'jcomment
'posted
'DONE!  Tbl_Fld_List()

'1.   ColumnWidth
'2.   ColumnOrder
'3.   ColumnHidden
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
'14.  SmartTags

'Datasheet properties:
'1.   RecordSource
'2.   Filter
'3.   FilterOn
'4.   OrderBy
'5.   OrderByOn
'6.   AllowFilters
'7.   Caption
'8.   DefaultView
'9.   ViewsAllowed
'10.  AllowFormView
'11.  AllowDatasheetView
'12.  AllowPivotTableView
'13.  AllowPivotChartView
'14.  AllowEditing
'15.  DefaultEditing
'16.  AllowEdits
'17.  AllowDeletions
'18.  AllowAdditions
'19.  DataEntry
'20.  AllowUpdating
'21.  RecordsetType
'22.  RecordLocks
'23.  ScrollBars
'24.  RecordSelectors
'25.  NavigationButtons
'26.  DividingLines
'27.  AutoResize
'28.  AutoCenter
'29.  PopUp
'30.  Modal
'31.  BorderStyle
'32.  ControlBox
'33.  MinButton
'34.  MaxButton
'35.  MinMaxButtons
'36.  CloseButton
'37.  WhatsThisButton
'38.  Width
'39.  Picture
'40.  PictureType
'41.  PictureSizeMode
'42.  PictureAlignment
'43.  PictureTiling
'44.  Cycle
'45.  MenuBar
'46.  Toolbar
'47.  ShortcutMenu
'48.  ShortcutMenuBar
'49.  GridX
'50.  GridY
'51.  LayoutForPrint
'52.  FastLaserPrinting
'53.  HelpFile
'54.  HelpContextId
'55.  SubdatasheetName
'56.  SubdatasheetName
'57.  LinkChildFields
'58.  LinkMasterFields
'59.  SubdatasheetHeight
'60.  SubdatasheetExpanded
'61.  RowHeight
'62.  DatasheetFontName
'63.  DatasheetFontHeight
'64.  DatasheetFontWeight
'65.  DatasheetFontItalic
'66.  DatasheetFontUnderline
'67.  DatasheetGridlinesBehavior
'68.  DatasheetGridlinesColor
'69.  DatasheetCellsEffect
'70.  DatasheetForeColor
'71.  ShowGrid
'72.  DatasheetBackColor
'73.  DatasheetBorderLineStyle
'74.  HorizontalDatasheetGridlineStyle
'75.  VerticalDatasheetGridlineStyle
'76.  DatasheetColumnHeaderUnderlineStyle
'77.  Hwnd
'78.  Count
'79.  LogicalPageWidth
'80.  Visible
'81.  Painting
'82.  PrtMip
'83.  PrtDevMode
'84.  PrtDevNames
'85.  FrozenColumns
'86.  Bookmark
'87.  Name
'88.  PaletteSource
'89.  Tag
'90.  PaintPalette
'91.  OpenArgs
'92.  OnCurrent
'93.  BeforeInsert
'94.  AfterInsert
'95.  BeforeUpdate
'96.  AfterUpdate
'97.  OnDirty
'98.  OnUndo
'99.  OnDelete
'100. BeforeDelConfirm
'101. AfterDelConfirm
'102. OnOpen
'103. OnLoad
'104. OnResize
'105. OnUnload
'106. OnClose
'107. OnActivate
'108. OnDeactivate
'109. OnGotFocus
'110. OnLostFocus
'111. OnClick
'112. OnDblClick
'113. OnMouseDown
'114. OnMouseMove
'115. OnMouseUp
'116. OnMouseWheel
'117. OnKeyDown
'118. OnKeyUp
'119. OnKeyPress
'120. KeyPreview
'121. OnError
'122. OnFilter
'123. OnApplyFilter
'124. OnTimer
'125. TimerInterval
'126. BeforeScreenTip
'127. OnCmdEnabled
'128. OnCmdChecked
'129. OnCmdBeforeExecute
'130. OnCmdExecute
'131. OnDataChange
'132. OnDataSetChange
'133. OnPivotTableChange
'134. OnSelectionChange
'135. OnViewChange
'136. OnConnect
'137. OnDisconnect
'138. BeforeQuery
'139. OnQuery
'140. AfterLayout
'141. BeforeRender
'142. AfterRender
'143. AfterFinalRender
'144. Dirty
'145. WindowWidth
'146. WindowHeight
'147. CurrentView
'148. CurrentSectionTop
'149. CurrentSectionLeft
'150. SelLeft
'151. SelTop
'152. SelWidth
'153. SelHeight
'154. CurrentRecord
'155. PictureData
'156. InsideHeight
'157. InsideWidth
'158. PicturePalette
'159. HasModule
'160. Orientation
'161. AllowDesignChanges
'162. WindowTop
'163. WindowLeft
'164. Moveable
'165. FetchDefaults

  DatasheetCol = blnRetVal

End Function
