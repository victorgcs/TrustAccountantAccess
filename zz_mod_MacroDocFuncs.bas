Attribute VB_Name = "zz_mod_MacroDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_MacroDocFuncs"

'VGC 11/23/2016: CHANGES!

'DOES EXPORTING A DIAGRAM INCLUDE THE RELATIONSHIP DIAGRAM?
'NO! DIAGRAMS ARE ONLY AVAILABLE IN ACCESS PROJECTS, WHICH ARE FRONTENDS FOR SQL SERVER.

' ** Graphic filters and file formats Microsoft Access can use:
' ** You can insert many popular graphics file formats into a form, report,
' ** or data access page either directly or with the use of separate graphics
' ** filters by using the Picture command on the Insert menu. For forms and
' ** reports, you don't need a separate graphics filter installed to insert
' ** Enhanced Metafile (.emf), Windows bitmap (.bmp, .rle, .dib), Windows
' ** Metafile (.wmf), and icon (.ico) graphics. However, you do need a graphics
' ** filter installed to insert all other graphics file formats listed below.
' **
' ** Graphics file formats requiring separate graphics filters
' **   Computer Graphics Metafile (.cgm) Kodak Photo CD (.pcd)
' **   CorelDRAW (.cdr) Macintosh PICT (.pct)
' **   Encapsulated PostScript (.eps) PC Paintbrush (.pcx)
' **   FlashPix (.fpx) Portable Network Graphics (.png)
' **   Graphic Interchange Format (.gif) Tagged Image File Format (.tif)
' **   Hanako (.jsh, .jah, .jbh) WordPerfect Graphics (.wpg)
' **   JPEG File Interchange Format (.jpg) X-Bitmap (.xbm)
' **
' ** For pages, you don't need a separate graphics filter installed to insert
' ** Windows bitmap (.bmp, .rle, .dib) graphics. However, you do need a graphics
' ** filter installed to insert Graphic Interchange Format (.gif), JPEG File
' ** Interchange Format (.jpg), Portable Network Graphics (.png), and X-Bitmap
' ** (.xbm) graphics.
' **
' ** The separate graphics filters are not supplied with the stand-alone version
' ** of Microsoft Access 2000. To use the filters, you need to install Microsoft
' ** Office 2000 Professional Edition, Microsoft Office 2000 Premier Edition, or
' ** a stand-alone version of Microsoft Word 2000 along with your stand-alone
' ** version of Access 2000.
' **
' ** Note: For more information on each separate graphics filter, see your Word documentation.

Private blnRetValx As Boolean
' **

Public Function QuikMacDoc() As Boolean
  Const THIS_PROC As String = "QuikMacDoc"
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Macro_ChkDocQrys(False) = True Then  ' ** Function: Below.
      blnRetValx = Macro_Doc  ' ** Function: Below.
      blnRetValx = Macro_Row_Doc  ' ** Function: Below.
      DoEvents
      DoBeeps  ' ** Module Function: modWindowFunctions.
      Debug.Print "'FINISHED!"
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED Macro_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikMacDoc = blnRetValx
End Function

Public Function Macro_Doc() As Boolean
' ** YAY! FOUND IT!
' ** Document all Macros to tblMacro, tblMacro_Text.

  Const THIS_PROC As String = "Macro_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstMcr As DAO.Recordset, rstTxt As DAO.Recordset, rstAct As DAO.Recordset
  Dim cntrs As DAO.Containers, cntr As DAO.Container, docs As DAO.Documents, doc As DAO.Document
  Dim fso As Scripting.FileSystemObject, fsfl As Scripting.File, fstxt As Scripting.TextStream
  Dim prps As DAO.Properties, prp As DAO.Property
  Dim strPath As String, strPathFile As String, strLine As String
  Dim lngMacs As Long, arr_varMac() As Variant
  Dim blnAdd As Boolean, blnFound As Boolean
  Dim lngLineCnt As Long
  Dim lngThisDbsID As Long
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngRecs As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varMac().
  Const M_ELEMS As Integer = 11  ' ** Array's first-element UBound().
  Const M_DID    As Integer = 0
  Const M_DNAM   As Integer = 1
  Const M_MID    As Integer = 2
  Const M_MNAM   As Integer = 3
  Const M_DSC    As Integer = 4
  Const M_PATH   As Integer = 5
  Const M_FILE   As Integer = 6
  Const M_TLINS  As Integer = 7
  Const M_ROWS   As Integer = 8
  Const M_COLS   As Integer = 9
  Const M_CREATE As Integer = 10
  Const M_LASTUP As Integer = 11

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.

  lngMacs = 0&
  ReDim arr_varMac(M_ELEMS, 0)

  ' ** Gather a list of all existing Macros into the arr_varMac() array.
  Set dbs = CurrentDb
  With dbs
    Set cntrs = .Containers
    For Each cntr In cntrs
      With cntr
        If .Name = "Scripts" Then
          Set docs = .Documents
          For Each doc In docs
            With doc
              lngMacs = lngMacs + 1&
              lngE = lngMacs - 1&
              ReDim Preserve arr_varMac(M_ELEMS, lngE)
              ' *****************************************************
              ' ** Array: arr_varMac()
              ' **
              ' **   Field  Element  Name                Constant
              ' **   =====  =======  ==================  ==========
              ' **     1       0     dbs_id              M_DID
              ' **     2       1     dbs_name            M_DNAM
              ' **     3       2     mcr_id              M_MID
              ' **     4       3     mcr_name            M_MNAM
              ' **     5       4     mcr_description     M_DSC
              ' **     6       5     Path                M_PATH
              ' **     7       6     Filename            M_FILE
              ' **     8       7     mcr_textlines       M_TLINS
              ' **     9       8     mcr_rows            M_ROWS
              ' **    10       9     mcr_columnsshown    M_COLS
              ' **    11      10     mcr_datecreated     M_CREATE
              ' **    12      11     mcr_lastupdated     M_LASTUP
              ' **
              ' *****************************************************
              arr_varMac(M_DID, lngE) = lngThisDbsID
              arr_varMac(M_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
              arr_varMac(M_MID, lngE) = CLng(0)
              arr_varMac(M_MNAM, lngE) = .Name
              arr_varMac(M_DSC, lngE) = vbNullString
              arr_varMac(M_PATH, lngE) = strPath
              arr_varMac(M_FILE, lngE) = StringReplace(.Name, " ", "_") & ".txt"  ' ** Module Function: modStringFuncs.
              arr_varMac(M_TLINS, lngE) = CLng(0)
              arr_varMac(M_ROWS, lngE) = CLng(0)
              arr_varMac(M_COLS, lngE) = CLng(0)
              arr_varMac(M_CREATE, lngE) = .DateCreated
              arr_varMac(M_LASTUP, lngE) = .LastUpdated
              ' ** Check for macro Description.
              Set prps = .Properties
              For Each prp In prps
                With prp
                  If .Name = "Description" Then
On Error Resume Next
                    strLine = .Value
                    If ERR.Number = 0 Then
On Error GoTo 0
                      If Trim(strLine) <> vbNullString Then
                        arr_varMac(M_DSC, lngE) = strLine
                      End If
                    Else
On Error GoTo 0
                    End If
                    Exit For
                  End If
                End With
              Next
            End With
          Next doc
          Exit For
        End If
      End With
    Next
    .Close
  End With

  If lngMacs > 0& Then
    Set dbs = CurrentDb
    With dbs

      Set rstMcr = .OpenRecordset("tblMacro", dbOpenDynaset, dbConsistent)
      Set rstTxt = .OpenRecordset("tblMacro_Text", dbOpenDynaset, dbConsistent)
      Set fso = CreateObject("Scripting.FileSystemObject")

      ' ** Update tblMacro, then generate the macro export and put it into tblMacro_Text.
      For lngX = 0& To (lngMacs - 1&)

        blnAdd = False: lngLineCnt = 0&: strPathFile = vbNullString: strLine = vbNullString

        With rstMcr
          If .BOF = True And .EOF = True Then
            blnAdd = True
            .AddNew
            ![dbs_id] = arr_varMac(M_DID, lngX)
          Else
            .FindFirst "[dbs_id] = " & CStr(arr_varMac(M_DID, lngX)) & " And [mcr_name] = '" & arr_varMac(M_MNAM, lngX) & "'"
            If .NoMatch = True Then
              blnAdd = True
              .AddNew
              ![dbs_id] = arr_varMac(M_DID, lngX)
            Else
              .Edit
            End If
          End If
          If blnAdd = True Then
            ![mcr_name] = arr_varMac(M_MNAM, lngX)
          End If
          If arr_varMac(M_DSC, lngX) <> vbNullString Then
            ![mcr_description] = arr_varMac(M_DSC, lngX)
          End If
          ![mcr_textlines] = arr_varMac(M_TLINS, lngX)
          ![mcr_rows] = arr_varMac(M_ROWS, lngX)  ' ** Update later.
          ![mcr_columnsshown] = arr_varMac(M_COLS, lngX)   ' ** Update later.
          ![mcr_datecreated] = arr_varMac(M_CREATE, lngX)
          ![mcr_lastupdated] = arr_varMac(M_LASTUP, lngX)
          ![mcr_datemodified] = Now()
          .Update
          .Bookmark = .LastModified
          arr_varMac(M_MID, lngX) = ![mcr_id]
        End With

        ' ** Since Access doesn't include Macros (also called Scripts) in its Object Model,
        ' ** export the macro to a text file, then read it back into tblMacro_Text.
        strPathFile = arr_varMac(M_PATH, lngX) & LNK_SEP & arr_varMac(M_FILE, lngX)
        Application.SaveAsText acMacro, arr_varMac(M_MNAM, lngX), strPathFile
        DoEvents

        If fso.FileExists(strPathFile) = True Then

          ' ** Delete tblMacro_Text, by specified [macid].
          Set qdf = .QueryDefs("zz_qry_Macro_01a")
          With qdf.Parameters
            ![macid] = arr_varMac(M_MID, lngX)
          End With
          qdf.Execute

          ' ** Now read the generated file, one line at a time, into tblMacro_Text.
          Set fsfl = fso.GetFile(strPathFile)
          Set fstxt = fsfl.OpenAsTextStream(ForReading, TristateFalse)
          With fstxt
            lngLineCnt = 0&
            Do While .AtEndOfStream <> True
              strLine = .ReadLine
              If Trim$(strLine) <> vbNullString Then
                lngLineCnt = lngLineCnt + 1&
                With rstTxt
                  .AddNew
                  ![dbs_id] = arr_varMac(M_DID, lngX)
                  ![mcr_id] = arr_varMac(M_MID, lngX)
                  ![mcrtxt_linenum] = lngLineCnt
                  ![mcrtxt_text] = strLine
                  ![mcrtxt_datemodified] = Now()
                  .Update
                End With
              End If
            Loop
            arr_varMac(M_TLINS, lngX) = lngLineCnt
            .Close
          End With
          Kill strPathFile
          Set fstxt = Nothing
          Set fsfl = Nothing
          Set qdf = Nothing

          ' ** Update tblMacro, by specified [macid], [tlin].
          Set qdf = .QueryDefs("zz_qry_Macro_03")
          With qdf.Parameters
            ![macid] = arr_varMac(M_MID, lngX)
            ![tlin] = arr_varMac(M_TLINS, lngX)
          End With
          qdf.Execute

        Else
          Stop
        End If

      Next

      rstTxt.Close
      rstMcr.Close

      ' ** Update zz_qry_Macro_04b (zz_qry_Macro_04a (tblMacro, linked to tblMacro_Text,
      ' ** by specified CurrentAppName()), just ColumnsShown, with mcr_columnsshown_new).
      Set qdf = .QueryDefs("zz_qry_Macro_04c")
      qdf.Execute

      lngDels = 0&
      ReDim arr_varDel(0)

      ' *****************************************************
      ' ** Array: arr_varMac()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     dbs_id              M_DID
      ' **     2       1     dbs_name            M_DNAM
      ' **     3       2     mcr_id              M_MID
      ' **     4       3     mcr_name            M_MNAM
      ' **     5       4     mcr_description     M_DSC
      ' **     6       5     Path                M_PATH
      ' **     7       6     Filename            M_FILE
      ' **     8       7     mcr_textlines       M_TLINS
      ' **     9       8     mcr_rows            M_ROWS
      ' **    10       9     mcr_columnsshown    M_COLS
      ' **    11      10     mcr_datecreated     M_CREATE
      ' **    12      11     mcr_lastupdated     M_LASTUP
      ' **
      ' *****************************************************

      ' ** Check for macros in tblMacro that aren't present anymore.
      Set rstMcr = .OpenRecordset("tblMacro", dbOpenDynaset, dbConsistent)
      With rstMcr
        If .BOF = True And .EOF = True Then
          ' ** It's empty!
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          For lngX = 1& To lngRecs
            If ![dbs_id] = lngThisDbsID Then
              blnFound = False
              For lngY = 0& To (lngMacs - 1&)
                If arr_varMac(M_MNAM, lngY) = ![mcr_name] Then
                  blnFound = True
                  Exit For
                End If
              Next
              If blnFound = False Then
                lngDels = lngDels + 1&
                lngE = lngDels - 1&
                ReDim Preserve arr_varDel(lngE)
                arr_varDel(lngE) = ![mcr_id]
              End If
            End If
            If lngX < lngRecs Then .MoveNext
          Next
        End If
        .Close
      End With

      If lngDels > 0& Then
        For lngX = 0& To (lngDels - 1&)
          ' ** Delete tblMacro, by specified [mcrid].
          Set qdf = .QueryDefs("zz_qry_Macro_01c")
          With qdf.Parameters
            ![mcrid] = arr_varDel(lngX)
          End With
          qdf.Execute
        Next
        Debug.Print "'MACROS DELETED: " & CStr(lngDels)
      End If

      .Close
    End With
  End If

'Large lists:
'Code Page
'Toolbar Name
'Command

' ** AcView enumeration:
' **   0  acViewNormal      (Default)
' **   1  acViewDesign
' **   2  acViewPreview
' **   3  acViewPivotTable
' **   4  acViewPivotChart

' ** AcFormView enumeration:
' **   0  acNormal   (Default)
' **   1  acDesign
' **   2  acPreview
' **   3  acFormDS

' ** AcFormOpenDataMode enumeration:
' **  -1  acFormPropertySettings  (Default)
' **   0  acFormAdd
' **   1  acFormEdit
' **   2  acFormReadOnly

' ** AcWindowMode enumeration:
' **   0  acWindowNormal  (Default)
' **   1  acHidden
' **   2  acIcon
' **   3  acDialog

' ** TextStream IOMode enumeration:
' **   1  ForReading    Open a file for reading only. You can't write to this file.
' **   2  ForWriting    Open a file for writing. If a file with the same name exists,
' **                    its previous contents are overwritten.
' **   8  ForAppending  Open a file and write to the end of the file.
 
' ** TextStream Format enumeration:
' **  -2  TristateUseDefault  Opens the file using the system default.
' **  -2  TristateMixed
' **  -1  TristateTrue        Opens the file as Unicode.
' **   0  TristateFalse       Opens the file as ASCII.

' ** '\"' : The backslash means take the next character as literal.

' *********************************
' ** EXAMPLES OF EXPORTED MACROS:
' *********************************

' ** zz_mcr_Backend_Switch:
'Version = 196611
'ColumnsShown = 0
'Begin
'    Action = "RunCode"
'    Comment = "z_mod_Backend_Compare.Backend_LinkAsk()"
'    Argument = "Backend_LinkAsk()"
'End

' ** AutoKeys:
'Version = 196611
'ColumnsShown = 3
'Begin
'    MacroName = "^+U"
'    Condition ="GetUserName()=\"VictorC\""
'    Action = "RunCode"
'    Comment = "Ctrl+Shift+U: Puts 'superuser' into ActiveControl; z_mod_Misc_Dev_Funcs.RetSuper"
'        "()"
'    Argument = "RetSuper()"
'End
'Begin
'    MacroName = "^+D"
'    Condition ="GetUserName()=\"VictorC\""
'    Action = "RunCode"
'    Comment = "Ctrl+Shift+d: Shows the Form Design toolbar; modWindowFunctions.CmdBars_Design()"
'    Argument = "CmdBars_Design()"
'End

  ' **********************************************************************************************
  ' ** The internal Access objects and methods below give NO INFORMATION on what's in the Macro!
  ' **********************************************************************************************

  'Set prj = Application.CurrentProject
  'With prj
  '  For Each mcr_ao In .AllMacros
  '    With mcr_ao
  '      'Debug.Print "'MACRO: " & .Name
  '      Set prps_ao = .Properties
  '      For Each prp_ao In prps_ao
  '        With prp_ao
  '          'Debug.Print "'PRP: " & .Name  ' ** 0, none, zip, nada, no properties.
  '        End With
  '      Next
  '      Exit For
  '    End With
  '  Next
  'End With

  'Set dbs = CurrentDb
  'With dbs
  '  Set cntrs = .Containers
  '  For Each cntr In cntrs
  '    With cntr
  '      If .Name = "Scripts" Then
  '        Set docs = .Documents
  '        For Each doc In docs
  '          With doc
  '            Set prps = .Properties
  '            Debug.Print "'MACRO: " & .Name & "  PRPS: " & CStr(prps.Count)
  '            For Each prp In prps
  '              With prp
  '                Debug.Print "'PRP: " & .Name
  '              End With
  '            Next
  '          End With
  '          'Exit For
  '        Next
  '        Exit For
  '      End If
  '    End With
  '  Next
  '  .Close
  'End With

' ** Macro properties generally available:
'MACRO: zz_mcr_Reset_Options  PRPS: 9
'PRP: Name
'PRP: Owner
'PRP: UserName
'PRP: Permissions
'PRP: AllPermissions
'PRP: Container
'PRP: DateCreated
'PRP: LastUpdated
'PRP: Description

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set fstxt = Nothing
  Set fsfl = Nothing
  Set fso = Nothing
  Set prp = Nothing
  Set prps = Nothing
  Set doc = Nothing
  Set docs = Nothing
  Set cntr = Nothing
  Set cntrs = Nothing
  Set rstAct = Nothing
  Set rstTxt = Nothing
  Set rstMcr = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Macro_Doc = blnRetValx

End Function

Public Function Macro_Row_Doc() As Boolean
' ** Document the Macro contents into tblMacro_Row, tblMacro_Row_Argument.

  Const THIS_PROC As String = "Macro_Row_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstTxt As DAO.Recordset, rstRow As DAO.Recordset, rstArg As DAO.Recordset
  Dim lngMacs As Long, arr_varMac() As Variant
  Dim lngThisMacroID As Long, lngThisRowID As Long
  Dim lngTheseMacroLines As Long, lngThisLine As Long, lngThisRow As Long
  Dim blnBeginFound As Boolean, blnEndFound As Boolean, blnAdd As Boolean
  Dim strMacroName As String, strCondition As String, strComment As String
  Dim strAction As String, varMcrActID As Variant, varMcrActArgID As Variant, lngMcrActArgs As Long
  Dim lngArgs As Long, arr_varArg() As Variant, strArgument As String
  Dim strLastType As String
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim intPos1 As Integer
  Dim varTmp00 As Variant
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varMac().
  Const M_ELEMS As Integer = 3
  Const M_DID  As Integer = 0
  Const M_DNAM As Integer = 1
  Const M_MID  As Integer = 2
  Const M_ROWS As Integer = 3

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    Set rstRow = .OpenRecordset("tblMacro_Row", dbOpenDynaset, dbConsistent)
    Set rstArg = .OpenRecordset("tblMacro_Row_Argument", dbOpenDynaset, dbConsistent)

    ' ** tblMacro_Text, sorted, by specified CurrentAppName().
    Set qdf = .QueryDefs("zz_qry_Macro_02")
    Set rstTxt = qdf.OpenRecordset
    With rstTxt
      If .BOF = True And .EOF = True Then
        ' ** Then what are you doing here?
      Else

        lngMacs = 0&
        ReDim arr_varMac(M_ELEMS, 0)

        lngThisMacroID = 0&: lngTheseMacroLines = 0&: lngThisLine = 0&: lngThisRow = 0&
        blnBeginFound = False: blnEndFound = False

        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs

          If ![mcr_id] <> lngThisMacroID Then

            lngThisMacroID = ![mcr_id]
            lngTheseMacroLines = ![mcr_textlines]
            lngThisRowID = 0&: lngThisRow = 0&
            blnBeginFound = False: blnEndFound = False
            strMacroName = vbNullString: strCondition = vbNullString: strAction = vbNullString
            strComment = vbNullString: strArgument = vbNullString

            lngMacs = lngMacs + 1&
            lngE = lngMacs - 1&
            ReDim Preserve arr_varMac(M_ELEMS, lngE)
            arr_varMac(M_DID, lngE) = lngThisDbsID
            arr_varMac(M_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
            arr_varMac(M_MID, lngE) = lngThisMacroID
            arr_varMac(M_ROWS, lngE) = CLng(0)

          End If

          lngThisLine = ![mcrtxt_linenum]
          If IsNull(![mcrtxt_text]) = False Then
            If Left$(Trim$(![mcrtxt_text]), 9) = "Version =" Or Left$(Trim$(![mcrtxt_text]), 14) = "ColumnsShown =" Then
              ' ** Skip these.
            Else
              If blnBeginFound = False Or (blnBeginFound = True And blnEndFound = True) Then
                If Left$(Trim$(![mcrtxt_text]), 5) = "Begin" Then

                  ' ** Reset all the row variables.
                  blnBeginFound = True
                  lngThisRowID = 0&: varMcrActID = Null: lngMcrActArgs = -1&
                  strMacroName = vbNullString: strCondition = vbNullString: strAction = vbNullString
                  strComment = vbNullString: strArgument = vbNullString
                  lngArgs = 0&
                  ReDim arr_varArg(0)
                  lngThisRow = lngThisRow + 1&

                End If
              Else
                If blnEndFound = False Then
                  If Left$(Trim$(![mcrtxt_text]), 3) = "End" Then
                    blnEndFound = True

                    ' ** Once 'End' is encountered, save the info into tblMacro_Row.
                    blnAdd = False
                    With rstRow
                      If .BOF = True And .EOF = True Then
                        blnAdd = True
                      Else
                        .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [mcr_id] = " & CStr(lngThisMacroID) & " And " & _
                          "[mcrrow_order] = " & CStr(lngThisRow)
                        If .NoMatch = True Then
                          blnAdd = True
                        End If
                      End If
                    End With

                    ' ** Delete tblMacro_Row_Argument, by specified [mcrid], [argord].
                    If blnAdd = False Then
                      Set qdf = dbs.QueryDefs("zz_qry_Macro_01b")
                      With qdf.Parameters
                        ![mcrid] = lngThisMacroID
                        ![argord] = lngThisRow
                      End With
                      qdf.Execute
                    End If

                    arr_varMac(M_ROWS, (lngMacs - 1&)) = arr_varMac(M_ROWS, (lngMacs - 1&)) + 1&

                    With rstRow
                      If blnAdd = True Then
                        .AddNew
                        ![dbs_id] = lngThisDbsID
                      Else
                        .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [mcr_id] = " & CStr(lngThisMacroID) & " And " & _
                          "[mcrrow_order] = " & CStr(lngThisRow)
                        .Edit
                      End If
                      If blnAdd = True Then
                        ![mcr_id] = lngThisMacroID
                      End If
                      If strAction = vbNullString Or IsNull(varMcrActID) = True Then
                        strAction = "{empty}"
                        ' ** It should find '{empty}' OK.
                        varMcrActID = DLookup("[mact_id]", "tblMacroAction", "[mact_name] = '" & strAction & "'")
                      End If
                      ![mact_id] = CLng(varMcrActID)
                      If blnAdd = True Then
                        ![mcrrow_order] = lngThisRow
                      End If
                      If strMacroName <> vbNullString Then
                        ![mcrrow_macroname] = strMacroName
                      Else
                        If IsNull(![mcrrow_macroname]) = False Then
                          ![mcrrow_macroname] = Null
                        End If
                      End If
                      If strCondition <> vbNullString Then
                        ![mcrrow_condition] = strCondition
                      Else
                        If IsNull(![mcrrow_condition]) = False Then
                          ![mcrrow_condition] = Null
                        End If
                      End If
                      If strComment <> vbNullString Then
                        ![mcrrow_comment] = strComment
                      Else
                        If IsNull(![mcrrow_comment]) = False Then
                          ![mcrrow_comment] = Null
                        End If
                      End If
                      ![mcrrow_datemodified] = Now()
                      If blnAdd = False Then
                        lngThisRowID = ![mcrrow_id]
                      End If
                      .Update
                      If blnAdd = True Then
                        .Bookmark = .LastModified
                        lngThisRowID = ![mcrrow_id]
                      End If
                    End With

                    ' ** Save the row's arguments into tblMacro_Row_Argument.
                    If lngArgs > 0& Then
                      With rstArg
                        For lngY = 0& To (lngArgs - 1&)
                          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [mcr_id] = " & CStr(lngThisMacroID) & " And " & _
                            "[mcrrow_id] = " & CStr(lngThisRowID) & " And [mact_id] = " & CStr(varMcrActID) & " And " & _
                            "[mcrrowarg_order] = " & CStr(lngY + 1&)
                          If .NoMatch = True Then
                            .AddNew
                            ![dbs_id] = lngThisDbsID
                          Else
                            .Edit
                          End If
                          ![mcr_id] = lngThisMacroID
                          ![mcrrow_id] = lngThisRowID
                          ![mact_id] = varMcrActID
                          ' ** This arugment's order, lngY + 1&, will match mactarg_order for this varMcrActID.
                          varMcrActArgID = DLookup("[mactarg_id]", "tblMacroActionArgument", _
                            "[mact_id] = " & CStr(varMcrActID) & " And " & _
                            "[mactarg_order] = " & CStr(lngY + 1&))
                          If IsNull(varMcrActArgID) = False Then
                            ' ** This links the exported Arugment line to a specific
                            ' ** choice in the Macro's list of inputs, allowing us to
                            ' ** know to what the argument refers, and its data type.
                            ' ** These, in turn, can be cross-referenced with
                            ' ** numerous type and option tables.
                            ![mactarg_id] = CLng(varMcrActArgID)
                          Else
                            ' ** I believe I've got them all!
                            ' ** (Though it's possible newer Access versions might differ.)
                            Stop
                          End If
                          ![mcrrowarg_order] = lngY + 1&
                          ![mcrrowarg_argument] = IIf(arr_varArg(lngY) = vbNullString, "NullString", arr_varArg(lngY))
                          ![mcrrowarg_datemodified] = Now()
                          .Update
                        Next
                      End With
                    End If

                    blnBeginFound = False: blnEndFound = False

                  Else

                    ' *********************************************
                    ' ** Here's where we parse the exported text.
                    ' *********************************************

                    intPos1 = InStr(![mcrtxt_text], "=")
                    If Left$(Trim$(![mcrtxt_text]), 9) = "MacroName" Then
                      ' **************************
                      ' ** MacroName column.
                      ' **************************
                      strLastType = "MacroName"
                      strMacroName = Trim$(Mid$(![mcrtxt_text], (intPos1 + 1)))
                      ' ** Remove the quotes from the export.
                      strMacroName = StringReplace(strMacroName, Chr(34), vbNullString)  ' ** Module Function: modStringFuncs.
                    ElseIf Left$(Trim$(![mcrtxt_text]), 9) = "Condition" Then
                      ' **************************
                      ' ** Condition column.
                      ' **************************
                      strLastType = "Condition"
                      strCondition = Trim$(Mid$(![mcrtxt_text], (intPos1 + 1)))
                      ' ** Remove the quotes from the export, checking for literal quotes that may be present and must remain.
                      ' ** 1. Replace '\"' with '^'.
                      ' ** 2. Remove the remaining quotes.
                      ' ** 3. Replace '^' with '"'.
                      strCondition = StringReplace(StringReplace(StringReplace(strCondition, ("\" & Chr(34)), "^"), _
                        Chr(34), vbNullString), "^", Chr(34))  ' ** Module Function: modStringFuncs.
                    ElseIf Left$(Trim$(![mcrtxt_text]), 6) = "Action" Then
                      ' **************************
                      ' ** Action column.
                      ' **************************
                      strLastType = "Action"
                      strAction = Trim$(Mid$(![mcrtxt_text], (intPos1 + 1)))
                      ' ** Remove the quotes from the export, checking for literal quotes that may be present and must remain.
                      strAction = StringReplace(strAction, Chr(34), vbNullString)  ' ** Module Function: modStringFuncs.
                      varMcrActID = DLookup("[mact_id]", "tblMacroAction", "[mact_name] = '" & strAction & "'")
                      If IsNull(varMcrActID) = False Then
                        lngMcrActArgs = DLookup("[mact_arguments]", "tblMacroAction", "[mact_id] = " & CStr(varMcrActID))
                        If lngMcrActArgs > 0& Then
                          ReDim arr_varArg(lngMcrActArgs - 1&)
                        Else
                          ReDim arr_varArg(0)
                        End If
                      Else
                        Stop
                      End If
                    ElseIf Left$(Trim$(![mcrtxt_text]), 7) = "Comment" Then
                      ' **************************
                      ' ** Comment column.
                      ' **************************
                      strLastType = "Comment"
                      strComment = Trim$(Mid$(![mcrtxt_text], (intPos1 + 1)))
                      ' ** Remove the quotes from the export, checking for literal quotes that may be present and must remain.
                      strComment = StringReplace(StringReplace(StringReplace(strComment, ("\" & Chr(34)), "^"), _
                        Chr(34), vbNullString), "^", Chr(34))  ' ** Module Function: modStringFuncs.
                    ElseIf Left$(Trim$(![mcrtxt_text]), 8) = "Argument" Then
                      ' **************************
                      ' ** Argument section.
                      ' **************************
                      strLastType = "Argument"
                      If lngMcrActArgs <> -1& Then
                        strArgument = Trim$(Mid$(![mcrtxt_text], (intPos1 + 1)))
                      ' ** Remove the quotes from the export, checking for literal quotes that may be present and must remain.
                        strArgument = StringReplace(StringReplace(StringReplace(strArgument, ("\" & Chr(34)), "^"), _
                          Chr(34), vbNullString), "^", Chr(34))  ' ** Module Function: modStringFuncs.
                        lngArgs = lngArgs + 1&
                        If lngArgs <= lngMcrActArgs Then
                          arr_varArg(lngArgs - 1&) = strArgument
                        Else
                          Stop
                        End If
                      Else
                        Stop
                      End If
                      strArgument = vbNullString
                    Else
                      If strLastType = "Comment" And intPos1 = 0 Then
                        ' **************************
                        ' ** Comment continuation.
                        ' **************************
                        ' ** I have only one example of this, so I don't know whether it will need StringReplace().
                        strComment = strComment & Trim$(![mcrtxt_text])
                      End If
                    End If

                  End If
                End If
              End If
            End If
          End If

          If lngX < lngRecs Then .MoveNext
        Next

      End If
      .Close
    End With

    rstArg.Close
    rstRow.Close

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    ' ** Delete any old rows in tblMacro_Row beyond the current count.
    If lngMacs > 0& Then
      For lngX = 0& To (lngMacs - 1&)
        If arr_varMac(M_DID, lngX) = lngThisDbsID Then
          varTmp00 = DCount("[mcrrow_id]", "tblMacro_Row", "[mcr_id] = " & CStr(arr_varMac(M_MID, lngX)))
          If CLng(varTmp00) > arr_varMac(M_ROWS, lngX) Then
            Debug.Print "'DEL: " & CStr(CLng(varTmp00) - arr_varMac(M_ROWS, lngX))
Stop
            ' ** Delete tblMacro_Row, by specified [mcrid], [roword].
            Set qdf = .QueryDefs("zz_qry_Macro_01d")
            With qdf.Parameters
              ![mcrid] = arr_varMac(M_MID, lngX)
              ![roword] = arr_varMac(M_ROWS, lngX)
            End With
            qdf.Execute
          End If
        End If
      Next
    Else
      Stop
    End If

    ' ** Update tblMacro, with DLookups() to zz_qry_Macro_05 (tblMacro_Row, grouped
    ' ** by mcr_id, with cnt, by specified CurrentAppName()), for mcr_actionlines.
    Set qdf = .QueryDefs("zz_qry_Macro_06")
    qdf.Execute

    .Close
  End With

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rstArg = Nothing
  Set rstRow = Nothing
  Set rstTxt = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Macro_Row_Doc = blnRetValx

End Function

Private Function Macro_ChkDocQrys(Optional varSkip As Variant) As Boolean
' ** Called by:
' **   QuikRelDoc(), Above

  Const THIS_PROC As String = "Macro_ChkDocQrys"

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

    Debug.Print "'MACRO DOC QRYS: " & CStr(lngQrys)
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
      Debug.Print "'ALL MACRO DOC QRYS PRESENT!"
    End If

    Debug.Print "'DONE!"
    DoEvents

    Beep

  End If  ' ** blnSkip.

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Macro_ChkDocQrys = blnRetValx

End Function
