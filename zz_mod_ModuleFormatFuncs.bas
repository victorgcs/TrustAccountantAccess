Attribute VB_Name = "zz_mod_ModuleFormatFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_ModuleFormatFuncs"

'VGC 09/02/2017: CHANGES!

' ** ACH     - Automated Clearing House (re personal checks)
' ** ALTCGD  - Asset Long Term Capital Gain Distribution
' ** CUSIP   - Committee of Uniform Security Identification Procedures
' ** EB      - Employee Benefit
' ** FBO     - For Benefit Of
' ** GTD     - Good-Till-Date Securities
' ** GUID    - Globally Unique IDentifier
' ** LT      - Long Term
' ** LTCG    - Long Term Capital Gains
' ** MICR    - Magnetic Ink Character Recognition
' ** OPEB    - Other Post Employment Benefits
' ** PC OD's - Principal Cash Overdrafts.
' ** REMIC   - Real Estate Mortgage Investment Conduits (Investment Grade Mortgage Bond)
' ** RMA     - Realized Market Adjustment
' ** ST      - Short Term

' ** MasterAsset table:
' **   Masterasset_TYPE field:
' **     RA = Regular Asset
' **     IA = Interest Asset
' **       CUSIP: 999999999
' **       This asset comes with the program, there can be only one, and it
' **       can't be deleted, though it can be renamed to whatever you want.
' **       Used with frmAccruedIncome.
' **   Description field prefixes:
' **     'HA-' = Hidden Asset
' **       Doesn't show up in drop-downs, and can only be hidden if 0 shares.
' **       Can have as many as you like.
' **     'SW-' = Sweep Asset
' **       Used with account_SWEEP field in Account table, and frmSweeper.
' **       Only 'SW-' assets show up in drop-down.
' **       Can have as many as you like.

' ** Type Declaration Character enumeration:
' **   @  Currency
' **   #  Double
' **   %  Integer
' **   &  Long
' **   !  Single
' **   $  String

' ** Printer's Quotes:
' **   Chr(145)  ‘  : Single Open
' **   Chr(146)  ’  : Single Close
' **   Chr(147)  “  : Double Open
' **   Chr(148)  ”  : Double Close

' ** Chr(160) is Arial Hard-Space (hardspace, hard space)!
' ** VBA_GetCode() in zz_mod_ModuleMiscFuncs.

' ** Array: arr_varBlock().
Private lngBlocks As Long, arr_varBlock() As Variant
Private Const B_ELEMS As Integer = 12  ' ** Array's first-element UBound().
Private Const B_OPEN       As Integer = 0
Private Const B_OPEN_LEN   As Integer = 1
Private Const B_OPEN_CNT   As Integer = 2
Private Const B_ALIGN1     As Integer = 3
Private Const B_ALIGN1_LEN As Integer = 4
Private Const B_ALIGN1_CNT As Integer = 5
Private Const B_ALIGN2     As Integer = 6
Private Const B_ALIGN2_LEN As Integer = 7
Private Const B_ALIGN2_CNT As Integer = 8
Private Const B_CLOSE      As Integer = 9
Private Const B_CLOSE_LEN  As Integer = 10
Private Const B_CLOSE_CNT  As Integer = 11
Private Const B_SCOPE      As Integer = 12

Private blnRetValx As Boolean  ' ** Universal replacement.

' ** To use these, only 1 QUIK_MOD_NAME can be active at a time. Put the name into the
' ** appropriate one, then remark-out the other 2. QUIK_FRM_NAME can always remain un-remarked.

'Private Const QUIK_MOD_NAME As String = "modVersionConvertFuncs1"
'Private Const QUIK_MOD_NAME As String = "Report_" & "rptCheckList_Users"
Private Const QUIK_FRM_NAME As String = "frmVersion_Main"
Private Const QUIK_MOD_NAME As String = "Form_" & QUIK_FRM_NAME
' **

Public Function QuikSpaces() As Boolean
  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  VBA_Strip_Spaces QUIK_MOD_NAME  ' ** Function: Below.
End Function

Public Function QuikBlanks() As Boolean
  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  VBA_Strip_Blanks QUIK_MOD_NAME  ' ** Function: Below.
 End Function

Public Function QuikFormat() As Boolean
  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  VBA_Module_Format QUIK_MOD_NAME  ' ** Function: Below.
End Function

Public Function QuikErrHandler() As Boolean
  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  VBA_Err_Handler QUIK_MOD_NAME  ' ** Function: Below.
End Function

Public Function QuikThisProc() As Boolean
  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  VBA_This_Proc QUIK_MOD_NAME  ' ** Function: Below.
End Function

Public Function QuikChkCtls() As Boolean
  'Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  'VBA_Chk_Ctls QUIK_FRM_NAME  ' ** Function: Below.
End Function

Public Function QuikAll() As Boolean
' ** Apply Quik's to all modules?

  Const THIS_PROC As String = "QuikAll"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngMods As Long, arr_varMod As Variant
  Dim lngX As Long

  ' ** Array: arr_varMod().
  Const M_ID   As Integer = 0
  Const M_NAME As Integer = 1

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs

    Set qdf = .QueryDefs("zz_qry_VBComponent_LineNumErrs_03")
    Set rst = qdf.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        ' ** Well you shouldn't be here then!
      Else
        .MoveLast
        lngMods = .RecordCount
        .MoveFirst
        arr_varMod = .GetRows(lngMods)
        ' *****************************************
        ' ** Array: arr_varMod()
        ' **
        ' **   Element  Description    Constant
        ' **   =======  =============  ==========
        ' **      0     vbcom_id       M_ID
        ' **      1     vbcom_name     M_NAME
        ' **
        ' *****************************************
      End If
      .Close
    End With

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    If lngMods > 0& Then
      Debug.Print "'MODS: " & CStr(lngMods) & "  ";
      Set rst = .OpenRecordset("tblVBComponent_Compile", dbOpenTable, dbAppendOnly)
      For lngX = 0& To (lngMods - 1&)
        DoCmd.Hourglass True
        ' ** Strip extra spaces from empty lines.
        VBA_Strip_Spaces arr_varMod(M_NAME, lngX), rst, arr_varMod(M_ID, lngX)  ' ** Function: Below.
        DoCmd.Hourglass True
        ' ** Strip multiple blank lines down to 1.
        VBA_Strip_Blanks arr_varMod(M_NAME, lngX), rst, arr_varMod(M_ID, lngX)  ' ** Function: Below.
        DoCmd.Hourglass True
        ' ** Renumber code lines, standardize indents, check for anomalies.
        VBA_Module_Format arr_varMod(M_NAME, lngX), rst, arr_varMod(M_ID, lngX)   ' ** Function: Below.
        DoCmd.Hourglass True
        ' ** Check for proper error handler section; add when not found.
        VBA_Err_Handler arr_varMod(M_NAME, lngX), rst, arr_varMod(M_ID, lngX)  ' ** Function: Below.
        DoCmd.Hourglass True
        ' ** Check THIS_PROC constant; add where not found.
        VBA_This_Proc arr_varMod(M_NAME, lngX), rst, arr_varMod(M_ID, lngX)  ' ** Function: Below.
        Beep
        Debug.Print ".";
        DoEvents
      Next
      rst.Close
    End If

    .Close
  End With

  Debug.Print
  Debug.Print "'FINISHED"

  DoCmd.Hourglass False

  Beep

'MODS: 64  ................................................................
'FINISHED

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  QuikAll = blnRetValx

End Function

Private Function VBA_Strip_Spaces(Optional varModName As Variant, Optional varComp As Variant, Optional varComID As Variant, Optional varDbsID As Variant) As Boolean
' ** Remove spaces from blank lines.
' ** Though extra spaces are automatically stripped from lines
' ** with code or remarks, spaces on blank lines remain.
' ** Called by:
' **   QuikSpaces(), Above
' **   QuikAll(), Above

  Const THIS_PROC As String = "VBA_Strip_Spaces"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngLines As Long, intLen As Integer
  Dim strLine As String, strModName As String
  Dim blnAllSpaces As Boolean
  Dim lngX As Long, lngY As Long

  blnRetValx = True

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        If strModName = varModName Then
          If IsMissing(varComp) = True Then
            Debug.Print "'" & strModName & "ª"
          End If
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            For lngX = 1& To lngLines
              strLine = .Lines(lngX, 1)
              intLen = Len(strLine)
              If intLen > 0 Then
                blnAllSpaces = True
                For lngY = 1& To intLen
                  If Mid(strLine, lngY, 1) <> " " And Mid(strLine, lngY, 1) <> "'" Then
                    blnAllSpaces = False
                    Exit For
                  End If
                Next
                If blnAllSpaces = True Then
                  .ReplaceLine lngX, ""
                  'Debug.Print "'Spaces: " & CStr(lngX)
                End If
              End If
            Next
          End With
        End If
      End With
    Next
  End With

  If IsMissing(varComp) = True Then
    Debug.Print blnRetValx
    Beep
  Else
    With varComp
      .AddNew
      ![dbs_id] = varDbsID
      ![vbcom_id] = varComID
      ![vbcomcomp_response] = "'" & varModName & "ª"
      ![vbcomcomp_return] = blnRetValx
      ![vbcomcomp_datemodified] = Now()
      .Update
    End With
  End If

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_Strip_Spaces = blnRetValx

End Function

Private Function VBA_Strip_Blanks(Optional varModName As Variant, Optional varComp As Variant, Optional varComID As Variant, Optional varDbsID As Variant) As Boolean
' ** Remove extra blank lines in code, leaving only one where there were two.
' ** Called by:
' **   QuikBlanks(), Above
' **   QuikAll(), Above
'### LOTS OF UNIQUE VARIABLES!

  Const THIS_PROC As String = "VBA_Strip_Blanks"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngModsChecked As Long
  Dim lngMods As Long, strModName As String
  Dim lngBlanks As Long, arr_varBlank() As Variant
  Dim lngAdds As Long, arr_varAdd() As Variant
  Dim lngLines As Long, lngTotBlanks As Long
  Dim strTmp01 As String, blnTmp02 As Boolean, blnTmp03 As Boolean
  Dim blnTmp04 As Boolean, blnTmp05 As Boolean, blnTmp06 As Boolean  'blnRepeat As Boolean, blnEndSelect As Boolean, blnNowLookForAdds As Boolean, blnLinesAdded As Boolean, blnFinished As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varBlank().
  Const B_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const B_LIN1 As Integer = 0
  Const B_LIN2 As Integer = 1

  ' ** Array: arr_varAdd().
  Const A_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const A_LIN1 As Integer = 0
  Const A_LIN2 As Integer = 1

  blnRetValx = False

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngMods = .VBComponents.Count
    lngModsChecked = 0&
    For Each vbc In .VBComponents
      lngModsChecked = lngModsChecked + 1&
      With vbc
        strModName = .Name
        If strModName = varModName Then
          blnRetValx = True
          If IsMissing(varComp) = True Then
            Debug.Print "'" & strModName & "¹"
          End If
          Set cod = .CodeModule
          With cod
            blnTmp02 = True
            blnTmp03 = False
            lngTotBlanks = 0&
            blnTmp04 = False
            blnTmp05 = False
            blnTmp06 = False
            Do While blnTmp02 = True
              ' ** Look for double blank lines.
              lngBlanks = 0&
              ReDim arr_varBlank(B_ELEMS, 0)
              ' **********************************************
              ' ** Array: arr_varBlank()
              ' **
              ' **   Element  Description       Constant
              ' **   =======  ================  ============
              ' **      0     Line Checked      B_LIN1
              ' **      1     2nd Blank Line    B_LIN2
              ' **
              ' **********************************************
              ' ** Look for no line between procedure.
              lngAdds = 0&
              ReDim arr_varAdd(A_ELEMS, 0)
              ' **********************************************
              ' ** Array: arr_varAdd()
              ' **
              ' **   Element  Description       Constant
              ' **   =======  ================  ============
              ' **      0     Line Checked      A_LIN1
              ' **      1     Add Blank Line    A_LIN2
              ' **
              ' **********************************************
              lngLines = .CountOfLines
              For lngX = 1& To (lngLines - 1&)
                strTmp01 = .Lines(lngX, 1)
                If strTmp01 <> vbNullString Then
                  If lngX < lngLines Then
                    strTmp01 = .Lines((lngX + 1), 1)
                    If strTmp01 = vbNullString Then
                      If (lngX + 1) < lngLines Then
                        strTmp01 = .Lines((lngX + 2), 1)
                        If strTmp01 = vbNullString Then
                          ' ** Extra line.
                          'Debug.Print "'EXTRA BLANKS: " & lngX
                          lngBlanks = lngBlanks + 1&
                          lngE = lngBlanks - 1&
                          ReDim Preserve arr_varBlank(B_ELEMS, lngE)
                          arr_varBlank(B_LIN1, lngE) = lngX
                          arr_varBlank(B_LIN2, lngE) = (lngX + 2&)
                        End If
                      End If
                    End If
                  End If
                Else
                  If blnTmp03 = True Then
                    ' ** That annoying blank line in the standard error handler!
                    ' ** Extra line.
                    blnTmp03 = False
                    lngBlanks = lngBlanks + 1&
                    lngE = lngBlanks - 1&
                    ReDim Preserve arr_varBlank(B_ELEMS, lngE)
                    arr_varBlank(B_LIN1, lngE) = lngX
                    arr_varBlank(B_LIN2, lngE) = lngX
                  End If
                End If
                strTmp01 = .Lines(lngX, 1)
                If InStr(strTmp01, "zErrorHandler") > 0 Then
                  ' ** A reference to zErrorHandler(), may be in Select-Else error handler.
                  strTmp01 = Trim(.Lines((lngX + 1&), 1))
                  If strTmp01 = vbNullString Then
                    ' ** Next line is a blank; looks good.
                    strTmp01 = .Lines((lngX - 1&), 1)
                    If InStr(strTmp01, "Case Else") > 0 Then
                      blnTmp03 = True
                    End If
                  End If
                End If
                If blnTmp04 = True Then
                  strTmp01 = Trim(.Lines(lngX, 1))
                  If strTmp01 <> vbNullString Then
                    If Left(strTmp01, 4) = "End " Then
                      ' ** Should be an end-of-procedure line.
                      If (lngX + 1&) < lngLines Then
                        strTmp01 = Trim(.Lines((lngX + 1&), 1))
                        If strTmp01 <> vbNullString Then
                          ' ** No blank between procedures!
                          lngAdds = lngAdds + 1&
                          lngE = lngAdds - 1&
                          ReDim Preserve arr_varAdd(A_ELEMS, lngE)
                          arr_varAdd(A_LIN1, lngE) = lngX
                          arr_varAdd(A_LIN2, lngE) = (lngX + 1&)
                        End If
                      End If
                    End If
                  End If
                End If
              Next
              If lngBlanks > 0& Then
                lngY = 0&
                For lngX = (lngBlanks - 1&) To 0& Step -1&
                  ' ** .CodeModule.DeleteLines(StartLine, Count)  Method
                  .DeleteLines arr_varBlank(B_LIN2, lngX)
                  lngY = lngY + 1&
                  If lngY > lngBlanks Then
                    If IsMissing(varComp) = True Then
                      Debug.Print "'STILL HERE?"
                    Else
                      With varComp
                        .AddNew
                        ![dbs_id] = varDbsID
                        ![vbcom_id] = varComID
                        ![vbcomcomp_response] = "'STILL HERE?"
                        ![vbcomcomp_return] = blnRetValx
                        ![vbcomcomp_datemodified] = Now()
                        .Update
                      End With
                    End If
                    Exit For
                  End If
                Next
              End If
              If lngBlanks > 0& Then
                lngTotBlanks = lngTotBlanks + lngBlanks
                'Debug.Print "'DELETED LINES FOUND! " & lngBlanks & " DELETED: " & lngY
              ElseIf lngBlanks = 0& And blnTmp04 = False And blnTmp06 = False Then
                ' ** Repeat one more time.
                blnTmp04 = True
              ElseIf lngBlanks = 0& And lngAdds > 0& And blnTmp04 = True And blnTmp06 = False Then
                lngY = 0&
                For lngX = (lngAdds - 1&) To 0& Step -1&
                  blnTmp05 = True
                  .InsertLines arr_varAdd(A_LIN2, lngX), ""
                  lngY = lngY + 1&
                  If lngY > lngAdds Then
                    If IsMissing(varComp) = True Then
                      Debug.Print "'NOT ADDED?"
                    Else
                      With varComp
                        .AddNew
                        ![dbs_id] = varDbsID
                        ![vbcom_id] = varComID
                        ![vbcomcomp_response] = "'NOT ADDED?"
                        ![vbcomcomp_return] = blnRetValx
                        ![vbcomcomp_datemodified] = Now()
                        .Update
                      End With
                    End If
                    Exit For
                  End If
                Next
                blnTmp06 = True
                ' ** I believe it will repeat one more time.
              Else
                blnTmp02 = False
                If lngAdds > 0& Then
                  If IsMissing(varComp) = True Then
                    Debug.Print "'NO BLANK LINES FOUND! " & lngTotBlanks & " DELETED, " & lngAdds & " ADDED"
                  Else
                    With varComp
                      .AddNew
                      ![dbs_id] = varDbsID
                      ![vbcom_id] = varComID
                      ![vbcomcomp_response] = "'NO BLANK LINES FOUND! " & lngTotBlanks & " DELETED, " & lngAdds & " ADDED"
                      ![vbcomcomp_return] = blnRetValx
                      ![vbcomcomp_datemodified] = Now()
                      .Update
                    End With
                  End If
                End If
              End If
            Loop  ' ** blnTmp02.
          End With
        End If  ' ** Specified module.
      End With
    Next
  End With

  If IsMissing(varComp) = True Then
    Debug.Print blnRetValx
    Beep
  Else
    With varComp
      .AddNew
      ![dbs_id] = varDbsID
      ![vbcom_id] = varComID
      ![vbcomcomp_response] = "'" & varModName & "¹"
      ![vbcomcomp_return] = blnRetValx
      ![vbcomcomp_datemodified] = Now()
      .Update
    End With
  End If

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_Strip_Blanks = blnRetValx

End Function

Private Function VBA_Module_Format(Optional varModName As Variant, Optional varComp As Variant, Optional varComID As Variant, Optional varDbsID As Variant) As Boolean
' ** Format module code to my standard, using numbered lines.
' ** Called by:
' **   QuikFormat(), Above
' **   QuikAll(), Above
'### LOTS OF UNIQUE VARIABLES!

  Const THIS_PROC As String = "VBA_Module_Format"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngLines As Long, lngLen As Long
  Dim lngDecLines As Long
  Dim strLine As String, lngLineNum As Long
  Dim lngModsChecked As Long
  Dim strModName As String, strProcName As String
  Dim blnNewMod As Boolean, blnNewProc As Boolean
  Dim blnProcStartFound As Boolean, blnProcEndFound As Boolean
  Dim blnIsProcStart As Boolean, blnIsProcEnd As Boolean
  Dim blnIsProp As Boolean, strPropKind As String, strLastPropKind As String
  Dim lngProcStart As Long, lngProcEnd As Long, lngLastProcEnd As Long, blnLastOnErrorFound As Boolean
  Dim blnOpeningRemarks As Boolean, blnOpenReport As Boolean
  Dim lngInds As Long, arr_varInd() As Variant
  Dim blnNumbered As Boolean, blnIsTerm As Boolean
  Dim blnLineContOn As Boolean, blnIsLineCont As Boolean
  Dim blnOpen As Boolean, blnClose As Boolean
  Dim blnIsOpeningRemark As Boolean, blnIsRemarkedLineNum As Boolean
  Dim blnOnErrorFound As Boolean, blnIsOnError As Boolean, lngFirstOnError As Long
  Dim blnDimGroupFound As Boolean, blnInDimGroup As Boolean
  Dim blnIsDim As Boolean, lngDimStart As Long, lngDimEnd As Long
  Dim lngIndNum As Long, strTermName As String, lngTermElem As Long
  Dim intPos01 As Integer, intPos02 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, lngTmp04 As Long, blnTmp05 As Boolean
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  Const MOD_LINE_NUM_START  As Long = 100&  ' ** Start the module line numbers at 100.
  Const MOD_LINE_NUM_INDENT As Integer = 6  ' ** Except proc open, close, and labels, indent 6 characters.
  Const MOD_INDENT          As Integer = 2  ' ** Standard indent for each structure block.

  ' ** Array: arr_varInd().
  Const I_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const I_TYP       As Integer = 0
  Const I_NUM       As Integer = 1
  Const I_TERM_ELEM As Integer = 2

  blnRetValx = False

  lngBlocks = 0&
  ReDim arr_varBlock(B_ELEMS, 0)
  VBA_Block_Term_Load  ' ** Procedure: Below.
  ' *********************************************************
  ' ** Array: arr_varBlock()
  ' **
  ' **   Element  Description                Constant
  ' **   =======  =========================  ==============
  ' **      0     Opening Term               B_OPEN
  ' **      1     Opening Term Length        B_OPEN_LEN
  ' **      2     Opening Term Words         B_OPEN_CNT
  ' **      3     Mid-Block Term 1           B_ALIGN1
  ' **      4     Mid-Block Term 1 Length    B_ALIGN1_LEN
  ' **      5     Mid-Block Term 1 Words     B_ALIGN1_CNT
  ' **      6     Mid-Block Term 2           B_ALIGN2
  ' **      7     Mid-Block Term 2 Length    B_ALIGN2_LEN
  ' **      8     Mid-Block Term 2 Words     B_ALIGN2_CNT
  ' **      9     Closing Term               B_CLOSE
  ' **     10     Closing Term Length        B_CLOSE_LEN
  ' **     11     Closing Term Words         B_CLOSE_CNT
  ' **     12     Scope Possible YN          B_SCOPE
  ' **
  ' *********************************************************

  lngInds = 0&
  ReDim arr_varInd(I_ELEMS, 0)
  ' ********************************************************
  ' ** Array: arr_varInd()
  ' **
  ' **   Element  Description                Constant
  ' **   =======  =========================  =============
  ' **      0     Structure Term Type        I_TYP
  ' **      1     Indent Number              I_NUM
  ' **      2     arr_varBlock() Element     I_TERM_ELEM
  ' **
  ' ********************************************************

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngModsChecked = 0&
    For Each vbc In .VBComponents
      lngModsChecked = lngModsChecked + 1&
      With vbc
        blnNewMod = True
        strModName = .Name
        blnNewProc = True
        blnProcStartFound = False: blnProcEndFound = False
        blnIsProcStart = False: blnIsProcEnd = False
        blnIsProp = False: strPropKind = vbNullString: strLastPropKind = vbNullString
        lngProcStart = 0&: lngProcEnd = 0&
        blnNumbered = False
        blnOpeningRemarks = True
        blnDimGroupFound = False: blnInDimGroup = False: blnIsDim = False: lngDimStart = 0&: lngDimEnd = 0&
        strProcName = vbNullString
        strLine = vbNullString
        lngLineNum = -1&
        lngIndNum = 0&
        lngInds = 0&
        ReDim arr_varInd(I_ELEMS, 0)
        blnOpenReport = False
        blnOnErrorFound = False: lngFirstOnError = 0&
        lngLastProcEnd = 0&: blnLastOnErrorFound = False

        If strModName = varModName Then

          blnRetValx = True
          If IsMissing(varComp) = True Then
            Debug.Print "'" & strModName & "²"
          End If
          Set cod = .CodeModule
          With cod

            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines

            For lngX = (lngDecLines + 1&) To lngLines
              ' ** .CodeModule.ProcOfLine(Line As Long, ProcKind As vbext_ProcKind) As String
              ' **   Returns name of procedure that specified line is in.
              ' **   Doesn't care if type of procedure is incorrect.

              blnNumbered = False
              strLine = vbNullString

              If .ProcOfLine(lngX, vbext_pk_Proc) <> vbNullString Then

'DEAL WITH DECLARATION SECTION!

'CHECK FOR COMPILER Directives!
'#Const, #If, #ElseIf, #Else, #End If

                ' ** For each procedure.
                If strProcName = vbNullString Then
                  ' ** The Declaration section has no procedure name, and will have no numbers.
                  ' ** This statement should fire only once per module.
                  blnNewProc = True
                  blnProcStartFound = False: blnProcEndFound = False
                  blnIsProcStart = False: blnIsProcEnd = False
                  blnIsProp = False: strPropKind = vbNullString: strLastPropKind = vbNullString
                  lngProcStart = 0&: lngProcEnd = 0&
                  blnLineContOn = False: blnIsLineCont = False
                  strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                  blnOpeningRemarks = True
                  blnOnErrorFound = False: lngFirstOnError = 0&
                  blnDimGroupFound = False: blnInDimGroup = False: blnIsDim = False: lngDimStart = 0&: lngDimEnd = 0&
                  lngIndNum = 0&
                  lngInds = 0&
                  ReDim arr_varInd(I_ELEMS, 0)
                ElseIf .ProcOfLine(lngX, vbext_pk_Proc) <> strProcName Or _
                    (.ProcOfLine(lngX, vbext_pk_Proc) = strProcName And _
                    blnProcEndFound = True And IsProperty(lngX, lngLines, lngProcEnd, cod) <> strLastPropKind) Then
                  strTmp01 = IsProperty(lngX, lngLines, lngProcEnd, cod)
                  If .ProcOfLine(lngX, vbext_pk_Proc) = strProcName Then
                    blnIsProp = True
                    strPropKind = strTmp01
                    strLastPropKind = strPropKind
                  Else
                    If strTmp01 <> vbNullString Then
                      blnIsProp = True
                      strPropKind = strTmp01
                      strLastPropKind = strPropKind
                    Else
                      blnIsProp = False
                      strPropKind = vbNullString
                      strLastPropKind = vbNullString
                    End If
                  End If
                  ' ** We've moved into a new procedure.
                  ' ** Doesn't care if type of procedure is incorrect.
                  If blnProcStartFound = False Or blnProcEndFound = False Then
                    blnRetValx = False
                    If IsMissing(varComp) = True Then
                      Beep
                      Debug.Print "'DIDN'T OPEN OR DIDN'T CLOSE! : " & lngX
                    Else
                      With varComp
                        .AddNew
                        ![dbs_id] = varDbsID
                        ![vbcom_id] = varComID
                        ![vbcomcomp_response] = "'DIDN'T OPEN OR DIDN'T CLOSE! : " & lngX
                        ![vbcomcomp_return] = blnRetValx
                        ![vbcomcomp_datemodified] = Now()
                        .Update
                      End With
                    End If
                  End If
                  If blnOnErrorFound = False Then
                    If IsMissing(varComp) = True Then
                      Beep
                      Debug.Print "'NO ERROR HANDLER! : " & strProcName
                    Else
                      With varComp
                        .AddNew
                        ![dbs_id] = varDbsID
                        ![vbcom_id] = varComID
                        ![vbcomcomp_response] = "'NO ERROR HANDLER! : " & strProcName
                        ![vbcomcomp_return] = blnRetValx
                        ![vbcomcomp_datemodified] = Now()
                        .Update
                      End With
                    End If
                  End If
                  blnNewProc = True
                  blnProcStartFound = False: blnProcEndFound = False
                  blnIsProcStart = False: blnIsProcEnd = False
                  lngLastProcEnd = lngProcEnd: blnLastOnErrorFound = blnOnErrorFound
                  lngProcStart = 0&: lngProcEnd = 0&
                  blnLineContOn = False: blnIsLineCont = False
                  strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                  blnOpeningRemarks = True
                  blnOnErrorFound = False: lngFirstOnError = 0&
                  blnDimGroupFound = False: blnInDimGroup = False: blnIsDim = False: lngDimStart = 0&: lngDimEnd = 0&
                  lngIndNum = 0&
                  lngInds = 0&
                  ReDim arr_varInd(I_ELEMS, 0)
                End If

                strLine = Trim(.Lines(lngX, 1))  ' ** Strip off leading or trailing spaces.
                lngLen = Len(strLine)
                If blnOpeningRemarks = True And lngLen = 0& And blnProcStartFound = True Then blnOpeningRemarks = False

'If Left(strLine, 1) = "#" Then
'  Debug.Print "'LINE: " & CStr(lngX) & "  " & strLine
'LINE: 1480  #If IsDev Then
'REPLACE 5!
'LINE: 1482  #Else
'LINE: 1484  #End If
'LINE: 1726  #If IsDev Then
'REPLACE 5!
'LINE: 1728  #Else
'LINE: 1730  #End If
'LINE: 1779  #If IsDev Then
'REPLACE 5!
'LINE: 1781  #Else
'LINE: 1783  #End If
'End If
                If lngLen > 0& Then
                  intPos01 = InStr(strLine, " ")  ' ** Find first space.
                  strTmp01 = vbNullString
                  If blnOpeningRemarks = True And intPos01 = 0 Then
                    If Left(strLine, 1) <> "'" Then blnOpeningRemarks = False
                  End If
                  If intPos01 > 0 Then
                    ' ** Label lines may not have any spaces, and can't be numbered anyway.
                    If blnNewProc = True And blnProcStartFound = False Then
                      If Left(strLine, 1) <> "'" Then
                        ' ** Find the proc's declaration line, it should be the first non-remark
                        ' ** line after blnNewProc = True.
                        ' ** A procedure includes all lines between it and the previous proc,
                        ' ** including remarks and blank lines.
                        ' ** Put it right to the left margin; it won't be numbered.
                        blnProcStartFound = True
                        blnIsProcStart = True
                        lngProcStart = lngX
                        .ReplaceLine lngX, strLine
'If InStr(strLine, "#If") > 0 Then Debug.Print "'REPLACE 1!"
                      Else
                        If blnOpeningRemarks = True Then
                          .ReplaceLine lngX, strLine
'If InStr(strLine, "#If") > 0 Then Debug.Print "'REPLACE 2!"
                        End If
                      End If
                    End If
                    If Left(strLine, 4) = "End " Then
                      If Mid(strLine, 5, 3) = "Sub" Or Mid(strLine, 5, 8) = "Function" Or _
                         Mid(strLine, 5, 8) = "Property" Then
                        ' ** Also move the proc's close to the left margin.
                        blnProcEndFound = True
                        blnIsProcEnd = True
                        lngProcEnd = lngX
                        .ReplaceLine lngX, strLine
'If InStr(strLine, "#If") > 0 Then Debug.Print "'REPLACE 3!"
                      End If
                    End If
                    strTmp01 = Left(strLine, (intPos01 - 1))  ' ** First word of line.
                    ' ** See if it's a numbered line.
                    If IsNumeric(strTmp01) = True Then
                      ' ** Check its context.
                      blnTmp05 = True
                      strTmp03 = Right(Trim(.Lines((lngX - 1), 1)), 1)
                      If strTmp03 = "_" Then
                        If InStr(strTmp01, "MsgBox") > 0 Then
                          'Debug.Print "'MSGBOX LINE: " & CStr(lngLineNum)
                        End If
                        ' ** Line before was continued, so definitely not a line number.
                        blnTmp05 = False
                        blnNumbered = False
                        If Len(Trim(.Lines(lngX, 1))) <= 5 Then
                          If IsMissing(varComp) = True Then
                            Debug.Print "SHORT CONTINUATION: " & lngX
                          Else
                            With varComp
                              .AddNew
                              ![dbs_id] = varDbsID
                              ![vbcom_id] = varComID
                              ![vbcomcomp_response] = "SHORT CONTINUATION: " & lngX
                              ![vbcomcomp_return] = blnRetValx
                              ![vbcomcomp_datemodified] = Now()
                              .Update
                            End With
                          End If
                        End If
                      Else
                        intPos02 = InStr(.Lines(lngX, 1), strTmp01)
                        For lngY = intPos02 To (InStr((intPos02 + 1), .Lines(lngX, 1), " ") - 1)
                          strTmp03 = Mid(.Lines(lngX, 1), lngY, 1)
                          If Asc(strTmp03) < 48 Or Asc(strTmp03) > 57 Then
                            ' ** Non numeric character, so definitely not a line number.
                            blnTmp05 = False
                            Exit For
                          End If
                        Next
                        If blnTmp05 = False Then
                          blnNumbered = False
                        Else
                          blnNumbered = True
                        End If
                      End If
                    End If
                    If blnIsProcStart = False And blnIsProcEnd = False And blnNumbered = True Then
                      ' ** This is a numbered line.
                      If blnOpeningRemarks = True Then blnOpeningRemarks = False
                      If lngLineNum = -1& Then
                        ' ** The first numbered line of a module; should fire only once per module.
                        lngLineNum = MOD_LINE_NUM_START
                        blnNewMod = False
                        blnNewProc = False
                      Else
                        If blnNewProc = True Then
                          ' ** The first numbered line of a procedure;
                          ' ** start at the next hundred to leave some room.
                          blnNewProc = False
                          lngLineNum = lngLineNum + 100&
                          strTmp01 = CStr(lngLineNum)
                          strTmp01 = Left(strTmp01, (Len(strTmp01) - 2)) & "00"
                          lngLineNum = Val(strTmp01)
                        Else
                          lngLineNum = lngLineNum + 10
                        End If
                      End If
                      intPos02 = 0
                      For lngY = intPos01 To Len(strLine)
                        If Mid(strLine, lngY, 1) <> " " And _
                           Mid(strLine, lngY, 1) <> vbCr And Mid(strLine, lngY, 1) <> vbLf Then
                          ' ** Find the first non-space character.
                          intPos02 = lngY
                          Exit For
                        End If
                      Next
                      If lngLineNum < MOD_LINE_NUM_START Then
                        ' ** How could it be less than the module's starting number?!
                        If IsMissing(varComp) = True Then
                          Beep
                          Debug.Print "'STOP : LINE NUM OFF " & lngLineNum
                        Else
                          With varComp
                            .AddNew
                            ![dbs_id] = varDbsID
                            ![vbcom_id] = varComID
                            ![vbcomcomp_response] = "'STOP : LINE NUM OFF " & lngLineNum
                            ![vbcomcomp_return] = blnRetValx
                            ![vbcomcomp_datemodified] = Now()
                            .Update
                          End With
                        End If
                      End If
                      strTmp01 = Left((CStr(lngLineNum) & Space(8)), 8)
                      If intPos02 > 0 Then
                        ' ** Warn of numbered Remarks!
                        If Mid(strLine, intPos01, 1) = "'" Then
                          If IsMissing(varComp) = True Then
                            Debug.Print "'REM NUM: " & strModName & "  LINE: " & Left(CStr(lngLineNum) & Space(8), 8) & strLine
                          Else
                            With varComp
                              .AddNew
                              ![dbs_id] = varDbsID
                              ![vbcom_id] = varComID
                              ![vbcomcomp_response] = "'STOP : LINE NUM OFF " & lngLineNum
                              ![vbcomcomp_return] = blnRetValx
                              ![vbcomcomp_datemodified] = Now()
                              .Update
                            End With
                          End If
                        End If
                        If intPos02 < MOD_LINE_NUM_INDENT Then
                          ' ** Make sure numbered lines are indented at least 8 characters.
                          strTmp01 = strTmp01 & Mid(strLine, intPos02)
                        Else
                          ' ** Otherwise leave indentation as is.
                          If Mid(strLine, MOD_LINE_NUM_INDENT, 1) <> " " Then
                            strTmp01 = strTmp01 & Mid(strLine, MOD_LINE_NUM_INDENT)
                          Else
                            strTmp01 = strTmp01 & Mid(strLine, (MOD_LINE_NUM_INDENT + 1))
                          End If
                        End If
                      Else
                        ' ** No characters!? Shouldn't have been able to number it then!
                        If IsMissing(varComp) = True Then
                          Beep
                          Debug.Print "'STOP : NO CHARACTERS!"
                        Else
                          With varComp
                            .AddNew
                            ![dbs_id] = varDbsID
                            ![vbcom_id] = varComID
                            ![vbcomcomp_response] = "'STOP : LINE NUM OFF " & lngLineNum
                            ![vbcomcomp_return] = blnRetValx
                            ![vbcomcomp_datemodified] = Now()
                            .Update
                          End With
                        End If
                      End If
                      .ReplaceLine lngX, strTmp01
'If InStr(strLine, "#If") > 0 Then Debug.Print "'REPLACE 4!"
                    ElseIf blnOpeningRemarks = True Then
                      If Left(strLine, 1) <> "'" And blnIsProcStart = False Then blnOpeningRemarks = False
                    End If  ' ** strLine is numbered.
                  Else
                    If lngX > 1& Then
                      strTmp03 = Right(Trim(.Lines((lngX - 1&), 1)), 1)
                      If strTmp03 = "_" Then
                        If Len(Trim(.Lines(lngX, 1))) <= 5 Then
                          If IsMissing(varComp) = True Then
                            Debug.Print "SHORT CONTINUATION: " & lngX
                          Else
                            With varComp
                              .AddNew
                              ![dbs_id] = varDbsID
                              ![vbcom_id] = varComID
                              ![vbcomcomp_response] = "SHORT CONTINUATION: " & lngX
                              ![vbcomcomp_return] = blnRetValx
                              ![vbcomcomp_datemodified] = Now()
                              .Update
                            End With
                          End If
                        End If
                      End If
                    End If
                  End If    ' ** intPos01 > 0, strLine contains a space (already trimmed).

                  ' ******************************************************************************
                  ' ** Indentation:
                  ' ******************************************************************************

                  ' ** Character 9 should be Dim's and all first level terms.
                  If (blnIsProcStart = False And blnIsProcEnd = False) And _
                     blnProcStartFound = True And blnProcEndFound = False Then
                    ' ** Don't check procedure open and close lines!

                    strLine = Trim(.Lines(lngX, 1))  ' ** Strip off leading or trailing spaces.
                    strTmp01 = strLine
                    lngLen = Len(strLine)
                    blnIsTerm = False
                    blnOpen = False
                    blnClose = False
                    strTermName = vbNullString
                    lngTermElem = -1

                    If blnNumbered = True Then strTmp01 = Trim(Mid(strTmp01, 9))  ' ** A numbered line.

                    For lngY = 0& To (lngBlocks - 1&)

                      If blnIsTerm = False Then
                        ' ** Check if it's an opening structure term.
                        strTmp02 = arr_varBlock(B_OPEN, lngY)
                        lngTmp04 = arr_varBlock(B_OPEN_LEN, lngY)
                        VBA_Block_Chk strTmp01, strTmp02, lngTmp04, blnIsTerm, strTermName, _
                          lngLen, blnNumbered, lngX, lngY  ' ** Function: Below.
                        If blnIsTerm = False Then
                          ' ** Check if it's a mid-block term.
                          strTmp02 = arr_varBlock(B_ALIGN1, lngY)
                          lngTmp04 = arr_varBlock(B_ALIGN1_LEN, lngY)
                          If strTmp02 <> vbNullString Then
                            VBA_Block_Chk strTmp01, strTmp02, lngTmp04, blnIsTerm, strTermName, _
                              lngLen, blnNumbered, lngX, lngY  ' ** Function: Below.
                          End If
                          If blnIsTerm = False Then
                            ' ** Check if it's a mid-block term.
                            strTmp02 = arr_varBlock(B_ALIGN2, lngY)
                            lngTmp04 = arr_varBlock(B_ALIGN2_LEN, lngY)
                            If strTmp02 <> vbNullString Then
                              VBA_Block_Chk strTmp01, strTmp02, lngTmp04, blnIsTerm, strTermName, _
                                lngLen, blnNumbered, lngX, lngY  ' ** Function: Below.
                            End If
                            If blnIsTerm = False Then
                              ' ** Check if it's a closing structure term.
                              strTmp02 = arr_varBlock(B_CLOSE, lngY)
                              lngTmp04 = arr_varBlock(B_CLOSE_LEN, lngY)
                              VBA_Block_Chk strTmp01, strTmp02, lngTmp04, blnIsTerm, strTermName, _
                                lngLen, blnNumbered, lngX, lngY  ' ** Function: Below.
                              If blnIsTerm = False Then
                                If arr_varBlock(B_SCOPE, lngY) = True Then
                                  ' ** Could still be an opening structure term with
                                  ' ** a Scope keyword at the beginning of the line.
                                  strTmp02 = arr_varBlock(B_OPEN, lngY)
                                  lngTmp04 = arr_varBlock(B_OPEN_LEN, lngY)
                                  If InStr(strTmp01, (" " & strTmp02 & " ")) > 0 Then
                                    If Left(strTmp01, 7) = "Public " Or Left(strTmp01, 8) = "Private " Or _
                                       Left(strTmp01, 7) = "Friend " Or Left(strTmp01, 7) = "Static " Or _
                                       Left(strTmp01, 7) = "Global " Then
                                      blnIsTerm = True
                                      blnOpen = True
                                      strTermName = strTmp02
                                    End If
                                  End If
                                End If
                              Else
                                blnClose = True
                              End If  ' ** Scope.
                            End If    ' ** Close Structure.
                          End If      ' ** Mid-Block.
                        Else
                          ' ** Check for single-line If-Then-Else block.
                          If blnIsTerm = True And strTermName = "If" Then
                            intPos01 = InStr(strTmp01, " Then")
                            If intPos01 > 0& Then
                              strTmp03 = Trim(Mid(strTmp01, intPos01))
                              If Len(strTmp03) > 4 Then
                                strTmp03 = Trim(Mid(strTmp03, 5))
                                If Left(strTmp03, 1) <> "'" Then
                                  ' ** Not a remark, so it must be a single-line block.
                                  blnIsTerm = False
                                End If
                              End If
                              If blnIsTerm = True Then blnOpen = True
                            Else
                              ' ** Check for line continuation.
                              blnIsTerm = False
                              If Right(strTmp01, 1) = "_" Then
                                For lngZ = (lngX + 1&) To lngLines
                                  strTmp03 = .Lines(lngZ, 1)
                                  intPos01 = InStr(strTmp03, " Then")
                                  If intPos01 > 0 Or Left(strTmp03, 4) = "Then" Then
                                    blnIsTerm = True
                                    Exit For
                                  ElseIf Right(strTmp03, 1) = "_" Then
                                    ' ** Keep going and check the next line.
                                  Else
                                    ' ** No "Then", no continuation, no dice!
                                    Exit For
                                  End If
                                Next
                                If InStr(strTmp01, "MsgBox") > 0 Then
                                  'Debug.Print "'MSGBOX LINE: " & CStr(lngLineNum)
                                End If
                              End If
                              If blnIsTerm = True Then
                                blnOpen = True
                              Else
                                If IsMissing(varComp) = True Then
                                  Beep
                                  Debug.Print "'IF WITH NO THEN! " & lngX
                                Else
                                  With varComp
                                    .AddNew
                                    ![dbs_id] = varDbsID
                                    ![vbcom_id] = varComID
                                    ![vbcomcomp_response] = "'IF WITH NO THEN! " & lngX
                                    ![vbcomcomp_return] = blnRetValx
                                    ![vbcomcomp_datemodified] = Now()
                                    .Update
                                  End With
                                End If
                              End If
                            End If
                          Else
                            blnOpen = True
                          End If
                        End If        ' ** Mid-Block.
                      End If          ' ** Open Structure.

                      If blnIsTerm = True Then
                        ' ** There can be only 1 structure component per line.
                        lngTermElem = lngY
                        Exit For
                      End If

                    Next  ' ** For each structure type, lngY.

                    If blnIsTerm = True Then
                      If blnOpen = True Then
                        ' ** If it's an Open, increment lngIndNum and add the new structure to the stack.
                        lngIndNum = lngIndNum + 1&
                        lngInds = lngInds + 1&
                        lngE = lngInds - 1&
                        ReDim Preserve arr_varInd(I_ELEMS, lngE)
                        arr_varInd(I_TYP, lngE) = strTermName
                        arr_varInd(I_NUM, lngE) = lngIndNum
                        arr_varInd(I_TERM_ELEM, lngE) = lngTermElem
                        ' ** This line should be indented appropriate to the previous indent, and all lines
                        ' ** up to, but not including this structure's close, should be indented one notch.
                      ElseIf blnClose = True Then
                        ' ** If it's a Close, decrement lngIndNum!
                        lngE = UBound(arr_varInd, 2)
                        If arr_varBlock(B_CLOSE, arr_varInd(I_TERM_ELEM, lngE)) = strTermName Then
                          ' ** This should close the last one on the stack.
                          lngIndNum = lngIndNum - 1&
                          lngInds = lngInds - 1&
                          If lngInds = 0& Then
                            ' ** Empty out the array.
                            ReDim arr_varInd(I_ELEMS, 0)
                          Else
                            lngE = lngInds - 1&
                            ReDim Preserve arr_varInd(I_ELEMS, lngE)  ' ** Lop off the last stack entry.
                          End If
                        Else
                          If IsMissing(varComp) = True Then
                            Beep
                            Debug.Print "'STOP : BLOCK NAME WRONG! " & lngX
                          Else
                            With varComp
                              .AddNew
                              ![dbs_id] = varDbsID
                              ![vbcom_id] = varComID
                              ![vbcomcomp_response] = "'STOP : BLOCK NAME WRONG! " & lngX
                              ![vbcomcomp_return] = blnRetValx
                              ![vbcomcomp_datemodified] = Now()
                              .Update
                            End With
                          End If
                        End If
                      End If
                    End If

                    strTmp01 = vbNullString
                    strTmp02 = vbNullString
                    If blnNumbered = True Then
                      strTmp01 = Left(strLine, MOD_LINE_NUM_INDENT)
                      strTmp02 = Trim(Mid(strLine, (MOD_LINE_NUM_INDENT + 1)))
                    Else
                      If strLine = vbNullString Then
                        strTmp01 = vbNullString
                        strTmp02 = vbNullString
                      ElseIf Right(strLine, 1) = ":" And Left(strLine, 1) <> "'" Then
                        strTmp01 = vbNullString
                        strTmp02 = strLine
                      Else
                        ' ** Even though Labels with Remarks won't be caught above,
                        ' ** they'll be automatically pushed to the left margin anyway.
                        strTmp01 = Space(MOD_LINE_NUM_INDENT)
                        strTmp02 = strLine
                      End If
                    End If

                    ' ** Add appropriate margins.
                    blnIsOnError = False
                    blnIsOpeningRemark = False
                    blnIsRemarkedLineNum = False
                    strTmp03 = vbNullString
                    If blnIsTerm = False And Left(strTmp02, 8) = "On Error" Or Left(strTmp02, 9) = "'On Error" Then
                      ' ** On Error's go to MOD_INDENT.
                      blnIsOnError = True
                      ' ** Check for first On Error of procedure.
                      If lngFirstOnError = 0& Then
                        lngFirstOnError = lngX
                        blnOnErrorFound = True
                      End If
                      If blnNumbered = False And Left(Trim(strLine), 1) <> "'" Then
                        If IsMissing(varComp) = True Then
                          Beep
                          Debug.Print "'ON ERROR NOT NUMBERED! : " & strModName & " " & strProcName & " " & lngX
                        Else
                          With varComp
                            .AddNew
                            ![dbs_id] = varDbsID
                            ![vbcom_id] = varComID
                            ![vbcomcomp_response] = "'ON ERROR NOT NUMBERED! : " & strModName & " " & strProcName & " " & lngX
                            ![vbcomcomp_return] = blnRetValx
                            ![vbcomcomp_datemodified] = Now()
                            .Update
                          End With
                        End If
                      End If
                    Else
                      ' ** Deal with line continuations!
                      If strLine = vbNullString Then
                        blnLineContOn = False
                        blnIsLineCont = False
                      ElseIf Left(strLine, 1) = "'" Then
                        blnLineContOn = False
                        blnIsLineCont = False
                      Else
                        If blnLineContOn = True Then
                          ' ** Line continuation is in effect.
                          blnIsLineCont = True
                          If Right(strLine, 1) = "_" Then
                            ' ** Line continuation continues to be in effect.
                            blnLineContOn = True
                          Else
                            ' ** Line continuation ends.
                            blnLineContOn = False
                          End If
                        Else
                          ' ** Line continuation not in effect.
                          blnIsLineCont = False
                          If Right(strLine, 1) = "_" Then
                            blnLineContOn = True
                          Else
                            blnLineContOn = False
                          End If
                        End If
                      End If
                      ' ** Deal with remarks.
                      If blnIsTerm = False Or (blnIsTerm = True And blnClose = True) Then
                        If blnOpeningRemarks = True And Left(strLine, 1) = "'" Then
                          blnIsOpeningRemark = True
                        Else
                          If blnOpeningRemarks = True Then blnOpeningRemarks = False
                          If Left(strTmp02, 1) = "'" Then
                            strTmp03 = Trim(Mid(strTmp02, 2))
                            intPos01 = InStr(strTmp03, " ")
                            If intPos01 > 0 Then
                              strTmp03 = Trim(Left(strTmp03, intPos01))
                              If IsNumeric(strTmp03) = True Then
                                ' ** If it's a remark, but one with a line number, move it to the left margin.
                                blnIsRemarkedLineNum = True
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                    If strTmp02 <> vbNullString Then
                      If Left(strTmp02, 4) = "Dim " Then
                        ' ** This is a Dim line.
                        blnIsDim = True
                        If blnDimGroupFound = False Then
                          ' ** First Dim line.
                          blnDimGroupFound = True
                          blnInDimGroup = True
                          lngDimStart = lngX  ' ** First Dim line.
                          If Right(CStr(lngLineNum), 2) <> "00" Then
                            If IsMissing(varComp) = True Then
                              Debug.Print "'DIM NOT AT BEGINNING!¹ " & lngX
                            Else
                              With varComp
                                .AddNew
                                ![dbs_id] = varDbsID
                                ![vbcom_id] = varComID
                                ![vbcomcomp_response] = "'DIM NOT AT BEGINNING!¹ " & lngX
                                ![vbcomcomp_return] = blnRetValx
                                ![vbcomcomp_datemodified] = Now()
                                .Update
                              End With
                            End If
                          End If
                        Else
                          ' ** Dim's were found on previous lines
                          If blnInDimGroup = False Then
                            If IsMissing(varComp) = True Then
                              Beep
                              Debug.Print "'DIM NOT AT BEGINNING!² " & lngX
                            Else
                              With varComp
                                .AddNew
                                ![dbs_id] = varDbsID
                                ![vbcom_id] = varComID
                                ![vbcomcomp_response] = "'DIM NOT AT BEGINNING!² " & lngX
                                ![vbcomcomp_return] = blnRetValx
                                ![vbcomcomp_datemodified] = Now()
                                .Update
                              End With
                            End If
                          End If
                        End If
                      Else
                        ' ** Not a Dim line.
                        If Left(strTmp02, 1) <> "'" And Left(strTmp02, 1) <> "#" Then  ' ** Compiler Directives: #Const, #If, #Else, #ElseIf, #End
                          ' ** Not a remark or Compiler Directive either.
                          If blnDimGroupFound = True Then
                            If blnInDimGroup = True Then
                              If blnIsLineCont = False Then
                                ' ** Whatever else it is, it's not a Dim!
                                If lngDimEnd = 0& Then
                                  lngDimEnd = lngX
                                Else
                                  If IsMissing(varComp) = True Then
                                    Debug.Print "'DOUBLE DIM END! " & lngX
                                  Else
                                    With varComp
                                      .AddNew
                                      ![dbs_id] = varDbsID
                                      ![vbcom_id] = varComID
                                      ![vbcomcomp_response] = "'DOUBLE DIM END! " & lngX
                                      ![vbcomcomp_return] = blnRetValx
                                      ![vbcomcomp_datemodified] = Now()
                                      .Update
                                    End With
                                  End If
                                End If
                                blnInDimGroup = False
                                blnIsDim = False
                              End If
                            End If
                          End If
                        End If
                      End If
                    End If
                    If blnIsOnError = True Then
                      If lngX = lngFirstOnError Then
                        If Right(Trim(strTmp01), 2) <> "00" Then
                          If IsMissing(varComp) = True Then
                            Debug.Print "'ON ERROR NOT FIRST NUMBER OF PROCEDURE! " & lngX
                          Else
                            With varComp
                              .AddNew
                              ![dbs_id] = varDbsID
                              ![vbcom_id] = varComID
                              ![vbcomcomp_response] = "'ON ERROR NOT FIRST NUMBER OF PROCEDURE! " & lngX
                              ![vbcomcomp_return] = blnRetValx
                              ![vbcomcomp_datemodified] = Now()
                              .Update
                            End With
                          End If
                        End If
                      End If
                      strTmp01 = strTmp01 & strTmp02
                    ElseIf blnIsOpeningRemark = True Then
                      strTmp01 = strTmp02
                    ElseIf blnIsRemarkedLineNum = True Then
                      strTmp01 = "'" & Trim(Mid(strTmp02, 2))
                    Else
If InStr(strTmp01, "#If") > 0 Then
'Stop
End If
                      If blnIsTerm = True And blnClose = False Then
                        If blnIsLineCont = True Then
                          strTmp01 = strTmp01 & Space((lngIndNum + 1&) * MOD_INDENT) & strTmp02
                        Else
                          strTmp01 = strTmp01 & Space(lngIndNum * MOD_INDENT) & strTmp02
                        End If
                      Else
                        If blnIsLineCont = True Then
                          strTmp01 = strTmp01 & Space((lngIndNum + 2&) * MOD_INDENT) & strTmp02
                        Else
                          strTmp01 = strTmp01 & Space((lngIndNum + 1&) * MOD_INDENT) & strTmp02
                        End If
                      End If
                    End If

                    If strLine <> vbNullString Then
If Left(Trim(strLine), 4) = "#If " Or Left(Trim(strLine), 5) = "#Else" Or Left(Trim(strLine), 5) = "#End " Then
  strTmp01 = Space(6) & Trim(strTmp01)
  'Debug.Print "'REPLACE 5!"
End If
                      .ReplaceLine lngX, strTmp01
                    End If

                  End If  ' ** Line within a procedure.

                  If blnIsProcStart = True Then
                    blnIsProcStart = False
                    '.ReplaceLine lngX, strLine
                  ElseIf blnIsProcEnd = True Then
                    blnIsProcEnd = False
                    '.ReplaceLine lngX, strLine
                  End If

                Else
                  ' ** Blank line.
                  blnLineContOn = False: blnIsLineCont = False
                End If      ' ** lngLen > 0, not a blank line.

              End If        ' ** Non-Declaration line.
            Next            ' ** For each line of CodeModule: lngX.

            If blnOnErrorFound = False And blnLastOnErrorFound = False Then
              If IsMissing(varComp) = True Then
                Beep
                Debug.Print "'NO ERROR HANDLER! : " & strProcName
              Else
                With varComp
                  .AddNew
                  ![dbs_id] = varDbsID
                  ![vbcom_id] = varComID
                  ![vbcomcomp_response] = "'NO ERROR HANDLER! : " & strProcName
                  ![vbcomcomp_return] = blnRetValx
                  ![vbcomcomp_datemodified] = Now()
                  .Update
                End With
              End If
            End If

          End With          ' ** This CodeModule: cod.

          If strModName = varModName Then
            Exit For
          End If
        End If  ' ** Specified module.

        If blnOpenReport = True Then
          'Debug.Print "'OpenReport: " & strModName
        End If

      End With            ' ** This VBComponent: vbc.
    Next                  ' ** For each VBComponent: vbc.
  End With                ' ** This VBProject: vbp.

  If IsMissing(varComp) = True Then
    Debug.Print blnRetValx
    Beep
  Else
    With varComp
      .AddNew
      ![dbs_id] = varDbsID
      ![vbcom_id] = varComID
      ![vbcomcomp_response] = "'" & varModName & "²"
      ![vbcomcomp_return] = blnRetValx
      ![vbcomcomp_datemodified] = Now()
      .Update
    End With
  End If

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_Module_Format = blnRetValx

End Function

Private Function VBA_This_Proc(Optional varModName As Variant, Optional varComp As Variant, Optional varComID As Variant, Optional varDbsID As Variant) As Boolean
' ** Called by:
' **   QuikThisProc(), Above
' **   QuikAll(), Above
'### LOTS OF UNIQUE VARIABLES!

  Const THIS_PROC As String = "VBA_This_Proc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngMods As Long, arr_varMod() As Variant
  Dim lngProcs As Long, arr_varProc() As Variant
  Dim arr_varWord() As Variant
  Dim lngLines As Long, lngProcLines As Long, lngModsChecked As Long
  Dim strModName As String, strProcName As String, strLine As String
  Dim lngProcStart As Long, lngProcEnd As Long
  Dim strProcKind As String, lngProcKind As Long
  Dim blnHasErrHand As Boolean, lngErrHandLine As Long
  Dim blnRemarkFound As Boolean, intRemarkPos As Integer
  Dim blnHasThisProc As Boolean, strThisProcErr As String, lngThisProcsAdded As Long
  Dim blnProcStartFound As Boolean, blnHasScope As Boolean, blnNumbered As Boolean
  Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngElemM As Long, lngLineNum As Long
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  ' ** Array: arr_varMod().
  Const M_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const M_NAM      As Integer = 0
  Const M_TYP      As Integer = 1
  Const M_PROCS    As Integer = 2
  Const M_PROC_ARR As Integer = 3

  ' ** Array: arr_varProc().
  Const P_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const P_NAM       As Integer = 0
  Const P_KND       As Integer = 1
  Const P_KNDNAM    As Integer = 2
  Const P_START     As Integer = 3
  Const P_END       As Integer = 4
  Const P_ERRH      As Integer = 5
  Const P_ERRH_LINE As Integer = 6
  Const P_EXIT      As Integer = 7
  Const P_MOD_ELEM  As Integer = 8
  Const P_THISPROC  As Integer = 9
  Const P_THISPROCE As Integer = 10

  ' ** Array: arr_varWord().
  Const W_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const W_LIN As Integer = 0
  'Const W_1 As Integer = 1
  Const W_2 As Integer = 2
  'Const W_3 As Integer = 3
  'Const W_4 As Integer = 4
  'Const W_5 As Integer = 5
  'Const W_6 As Integer = 6

  Const lngWords As Long = 6&

  Const TPROC As String = "        Const THIS_PROC As String = "

  blnRetValx = False

  lngMods = 0&
  ReDim arr_varMod(M_ELEMS, 0)
  ' ****************************************************
  ' ** Array: arr_varMod()
  ' **
  ' **   Element  Description             Constant
  ' **   =======  ======================  ============
  ' **      0     Module Name             M_NAM
  ' **      1     Module Type             M_TYP
  ' **      2     Number Of Procedures    M_PROCS
  ' **      3     arr_varProc() Array     M_PROC_ARR
  ' **
  ' ****************************************************

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngModsChecked = 0&
    For Each vbc In .VBComponents
      With vbc

        lngModsChecked = lngModsChecked + 1&
        strModName = .Name

        If strModName = varModName Then

          blnRetValx = True
          lngThisProcsAdded = 0&

          lngProcs = 0&
          ReDim arr_varProc(P_ELEMS, 0)  ' ** ReDim after every module.
          ' *****************************************************
          ' ** Array: arr_varProc()
          ' **
          ' **   Element  Description             Constant
          ' **   =======  ======================  =============
          ' **      0     Procedure Name          P_NAM
          ' **      1     Procedure Kind          P_KND
          ' **      2     Procedure Kind Name     P_KNDNAM
          ' **      3     Start Line              P_START
          ' **      4     End Line                P_END
          ' **      5     Has Error Handler       P_ERRH
          ' **      6     Has Exit Label          P_EXIT
          ' **      7     arr_varMod() Element    P_MOD_ELEM
          ' **      8     THIS_PROC True/False    P_THISPROC
          ' **      9     THIS_PROC Error         P_THISPROCE
          ' **
          ' *****************************************************

          ' ** Only 1 record.
          ReDim arr_varWord(W_ELEMS, 0)  ' ** ReDim after every procedure.
          ' ********************************************************
          ' ** Array: arr_varWord()
          ' **
          ' **   Element  Description                   Constant
          ' **   =======  ============================  ==========
          ' **      0     Procedure Declaration Line    W_LIN
          ' **      1     1st Word                      W_1
          ' **      2     2nd Word                      W_2
          ' **      3     3rd Word                      W_3
          ' **      4     4th Word                      W_4
          ' **      5     5th Word                      W_5
          ' **      6     6th Word                      W_6
          ' **
          ' ********************************************************

          lngMods = lngMods + 1&
          lngE = lngMods - 1&
          ReDim Preserve arr_varMod(M_ELEMS, lngE)
          arr_varMod(M_NAM, lngE) = .Name
          arr_varMod(M_TYP, lngE) = .Type
          ' **   vbext_ComponentType enumeration:
          ' **       1  vbext_ct_StdModule        Standard Module
          ' **       2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
          ' **       3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm in the VBA Editor.
          ' **      11  vbext_ct_ActiveXDesigner
          ' **     100  vbext_ct_Document         Module behind Form, Report, or Excel Worksheet.
          arr_varMod(M_PROCS, lngE) = 0&
          strModName = .Name
          lngElemM = lngE

          Set cod = .CodeModule
          With cod

            lngLines = .CountOfLines
            strProcName = vbNullString
            lngProcLines = 0&
            lngProcStart = 0&: lngProcEnd = 0&
            blnProcStartFound = False
            strProcKind = vbNullString: lngProcKind = -1
            blnRemarkFound = False: intRemarkPos = 0
            blnHasScope = False: blnHasErrHand = False: lngErrHandLine = 0&
            blnHasThisProc = False: strThisProcErr = vbNullString

            For lngX = 1& To lngLines
              lngLineNum = lngX

              blnNumbered = False
              strLine = vbNullString

              ' ** Declaration lines don't have a procedure name,
              ' ** and I don't believe they can have labels anyway.
              If .ProcOfLine(lngX, vbext_pk_Proc) <> vbNullString Then
                ' ** Returns name of procedure that the specified line is in.
                ' ** Doesn't care if type of procedure is incorrect.

                If .ProcOfLine(lngX, vbext_pk_Proc) <> strProcName Then
                  ' ** A new procedure.
                  If blnProcStartFound = True Then
                    strTmp01 = vbNullString
                    For lngY = 1& To lngWords
                      strTmp01 = strTmp01 & " " & arr_varWord(lngY, 0)
                    Next
                    lngProcs = lngProcs + 1&
                    lngE = lngProcs - 1&
                    ReDim Preserve arr_varProc(P_ELEMS, lngE)
                    arr_varProc(P_NAM, lngE) = strProcName
                    arr_varProc(P_KND, lngE) = lngProcKind
                    arr_varProc(P_KNDNAM, lngE) = strProcKind
                    arr_varProc(P_START, lngE) = lngProcStart
                    arr_varProc(P_END, lngE) = lngProcEnd
                    arr_varProc(P_ERRH, lngE) = blnHasErrHand
                    arr_varProc(P_ERRH_LINE, lngE) = lngErrHandLine
                    arr_varProc(P_EXIT, lngE) = False          ' ** Not figured out yet!
                    arr_varProc(P_MOD_ELEM, lngE) = lngElemM
                    arr_varProc(P_THISPROC, lngE) = blnHasThisProc
                    arr_varProc(P_THISPROCE, lngE) = strThisProcErr
                    strTmp01 = vbNullString
                  End If  ' ** Save previous procedure's info.
                  ' ** Only 1 record.
                  ReDim arr_varWord(W_ELEMS, 0)
                  strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                  strProcKind = vbNullString: lngProcKind = -1
                  blnProcStartFound = False
                  blnHasScope = False: blnHasErrHand = False: lngErrHandLine = 0&
                  blnHasThisProc = False: strThisProcErr = vbNullString
                  lngProcStart = 0&: lngProcLines = 0&: lngProcEnd = 0&
                End If  ' ** New procedure.

                strLine = Trim(.Lines(lngX, 1))

                If strLine <> vbNullString Then

                  If blnProcStartFound = False Then

                    If Left(strLine, 1) <> "'" Then
                      ' ** It's not a remark.
                      intPos02 = InStr(strLine, " ")
                      If intPos02 > 0 Then
                        ' ** It's got a space, as a procedure declaration line will.

                        ' ** Collect the first 6 words in the line.
                        intPos01 = 1&
                        For lngY = 1& To lngWords
                          If arr_varWord(lngY, 0) = vbNullString Then
                            If lngY = 1& Then
                              arr_varWord(W_LIN, 0) = strLine
                            End If
                            arr_varWord(lngY, 0) = Trim(Mid(strLine, intPos01, (intPos02 - intPos01)))
                            If Left(arr_varWord(lngY, 0), 1) = "'" Then
                              ' ** Remark encountered.
                              blnRemarkFound = True
                              intRemarkPos = intPos01 + IIf(lngY = 1&, 0, 1)
                              arr_varWord(lngY, 0) = vbNullString
                              Exit For
                            Else
                              If InStr(arr_varWord(lngY, 0), "(") > 0 Then arr_varWord(lngY, 0) = _
                                Left(arr_varWord(lngY, 0), (InStr(arr_varWord(lngY, 0), "(") - 1))
                              intPos01 = intPos02
                            End If
                          End If
                          intPos02 = InStr((intPos02 + 1), strLine, " ")
                          If intPos02 = 0 Then
                            Exit For
                          End If
                        Next

                      End If  ' ** Has at least 1 space.
                    Else
                      blnRemarkFound = True
                    End If  ' ** Not a remark.

                    ' ** Now pick up last word.
                    If blnRemarkFound = False And intPos01 > 0 Then
                      For lngY = 1& To lngWords
                        If arr_varWord(lngY, 0) = vbNullString Then
                          arr_varWord(lngY, 0) = Trim(Mid(strLine, intPos01))
                          If InStr(arr_varWord(lngY, 0), "(") > 0 Then
                            arr_varWord(lngY, 0) = _
                              Left(arr_varWord(lngY, 0), (InStr(arr_varWord(lngY, 0), "(") - 1))
                          End If
                          Exit For
                        End If
                      Next
                    End If

                    If arr_varWord(W_2, 0) <> vbNullString Then
                      ' ** Procedures declaration will have at least 1 space.

                      blnHasScope = False

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
                            blnRetValx = False
                            If IsMissing(varComp) = True Then
                              Beep
                              Debug.Print "'UNKNOWN PROPERTY TYPE! : " & lngX & " " & strTmp01 & " " & strTmp02
                            Else
                              With varComp
                                .AddNew
                                ![dbs_id] = varDbsID
                                ![vbcom_id] = varComID
                                ![vbcomcomp_response] = "'UNKNOWN PROPERTY TYPE! : " & lngX & " " & strTmp01 & " " & strTmp02
                                ![vbcomcomp_return] = blnRetValx
                                ![vbcomcomp_datemodified] = Now()
                                .Update
                              End With
                            End If
                          End Select
                        Case "Let", "Set", "Get"
                          ' ** Usually 3rd Word.
                          If lngProcKind = -1& Then
                            If IsMissing(varComp) = True Then
                              Debug.Print "'WHY KIND NOT FOUND ON PREVIOUS LOOP? : " & lngY & " LINE: " & lngX & " '" & strTmp01 & "'"
                            Else
                              With varComp
                                .AddNew
                                ![dbs_id] = varDbsID
                                ![vbcom_id] = varComID
                                ![vbcomcomp_response] = "'WHY KIND NOT FOUND ON PREVIOUS LOOP? : " & lngY & " LINE: " & lngX & " '" & strTmp01 & "'"
                                ![vbcomcomp_return] = blnRetValx
                                ![vbcomcomp_datemodified] = Now()
                                .Update
                              End With
                            End If
                          End If
                        Case Else
                          If strTmp01 = strProcName Then
                            ' ** Usually 4th word.
                            If lngY < lngWords Then
                              For lngZ = (lngY + 1&) To lngWords
                                arr_varWord(lngZ, 0) = vbNullString
                              Next
                            End If
                            If lngProcKind <> -1& Then
                              blnProcStartFound = True
                              lngProcStart = lngX '.ProcStartLine(strProcName, lngProcKind)
                              lngProcLines = .ProcCountLines(strProcName, lngProcKind)
                              lngZ = (lngProcStart + lngProcLines)
                              strTmp02 = "End " & strProcKind
                              strTmp03 = Trim(.Lines(lngZ, 1))
                              intPos01 = InStr(strTmp03, strTmp02)
                              Do While intPos01 = 0
                                lngZ = lngZ - 1&
                                intPos01 = InStr(.Lines(lngZ, 1), strTmp02)
                                If lngZ = lngX Then
                                  lngZ = 0&
                                  Exit Do  ' ** Exit the loop!
                                End If
                              Loop
                              If lngZ > 0& Then
                                lngProcEnd = lngZ
                              Else
                                If IsMissing(varComp) = True Then
                                  Debug.Print "'STOP : P_END NOT FOUND! " & strProcName & " " & lngX
                                Else
                                  With varComp
                                    .AddNew
                                    ![dbs_id] = varDbsID
                                    ![vbcom_id] = varComID
                                    ![vbcomcomp_response] = "'STOP : P_END NOT FOUND! " & strProcName & " " & lngX
                                    ![vbcomcomp_return] = blnRetValx
                                    ![vbcomcomp_datemodified] = Now()
                                    .Update
                                  End With
                                End If
                              End If
                            End If
                            Exit For
                          Else
                            ' ** Unknown 1st word; not a procedure declaration line.
                            'Beep
                            blnRetValx = False
                            If IsMissing(varComp) = True Then
                              Debug.Print "'UNKNOWN FIRST WORD! : lngX = " & lngX & ", lngY = " & lngY & " " & strTmp01
'LOTS OF "UNKNOWN FIRST WORD!" MEANS A WHOLE REMARKED OUT PROCEDURE IS THROWING OFF THE PROCESS.
                            Else
                              With varComp
                                .AddNew
                                ![dbs_id] = varDbsID
                                ![vbcom_id] = varComID
                                ![vbcomcomp_response] = "'UNKNOWN FIRST WORD! : lngX = " & lngX & ", lngY = " & lngY & " " & strTmp01
                                ![vbcomcomp_return] = blnRetValx
                                ![vbcomcomp_datemodified] = Now()
                                .Update
                              End With
                            End If
                          End If
                          Exit For
                        End Select

                      Next  ' ** For each word: lngY.

                    End If  ' ** Possible declaration line: arr_varWord(W_2, 0).

                  Else
                    ' ** We're within a procedure.

                    If Left(strLine, 1) <> "'" Then
                      ' ** Not a remark.
                      intPos01 = InStr(strLine, " ")
                      If intPos01 > 0 Then
                        strTmp01 = Trim(Left(strLine, intPos01))
                        If IsNumeric(strTmp01) = True Then
                          ' ** Constant, like variable declarations, can't be numbered.
                          blnNumbered = True
                          strTmp01 = Trim(Mid(strLine, intPos01))
                          If Left(strTmp01, 14) = "On Error GoTo " And blnHasErrHand = False Then
                            blnHasErrHand = True
                            lngErrHandLine = lngLineNum
                          End If
                        Else
                          If strTmp01 = "Const" Then
                            ' ** It's a constant declaration.
                            intPos02 = InStr((intPos01 + 1&), strLine, " ")  ' ** Next after constant name (e.g., 'THIS_PROC').
                            If intPos02 > 0 Then
                              strTmp02 = Trim(Mid(strLine, (intPos01 + 1&), ((intPos02 - intPos01) - 1&)))
                              If strTmp02 = "THIS_PROC" Then
                                intPos03 = InStr((intPos02 + 1&), strLine, " ")  ' ** Next after 'As' or '='.
                                If intPos03 > 0 Then
                                  If Left(Trim(Mid(strLine, (intPos02 + 1&))), 9) = "As String" Then
                                    intPos03 = InStr((intPos03 + 1&), strLine, " ")  ' ** Next after 'As String' (var reused).
                                    If intPos03 > 0 Then
                                      If Left(Trim(Mid(strLine, (intPos03 + 1&))), 1) = "=" Then
                                        intPos03 = InStr(strLine, Chr(34))  ' ** First quote (var reused).
                                        If intPos03 > 0 Then
                                          strTmp03 = Trim(Mid(strLine, intPos03))
                                          If InStr(strTmp03, "'") > 0 Then  ' ** Strip any remarks.
                                            strTmp03 = Trim(Left(strTmp03, (InStr(strTmp03, "'") - 1)))
                                          End If
                                          If Left(strTmp03, 1) = Chr(34) And Right(strTmp03, 1) = Chr(34) Then
                                            strTmp03 = Mid(Left(strTmp03, (Len(strTmp03) - 1)), 2)
                                            If strTmp03 = strProcName Then
                                              ' ** All's well!
                                              blnHasThisProc = True
                                              'Debug.Print "'" & strProcName & "  '" & strTmp03 & "'"
                                            Else
                                              strThisProcErr = "'WRONG PROC: " & CStr(lngLineNum) & " '" & strProcName & "'  '" & strTmp03 & "'"
                                              'Debug.Print strThisProcErr
                                            End If
                                          Else
                                            ' ** Something's off.
                                            strThisProcErr = "'SOMETHING'S OFF: " & CStr(lngLineNum) & "  " & strLine
                                            'Debug.Print strThisProcErr
                                          End If
                                        Else
                                          ' ** There must be quotes.
                                          strThisProcErr = "'WHERE'S THE QUOTE? " & CStr(lngLineNum) & "  " & strLine
                                          'Debug.Print strThisProcErr
                                        End If
                                      Else
                                        ' ** There must be an equal sign.
                                        strThisProcErr = "'WHERE'S THE OPERATOR? " & CStr(lngLineNum) & "  " & strLine
                                        'Debug.Print strThisProcErr
                                      End If
                                    Else
                                      ' ** There must be more words!
                                      strThisProcErr = "'WHAT? 3: " & CStr(lngLineNum) & "  " & strLine
                                      'Debug.Print strThisProcErr
                                    End If
                                  Else
                                    ' ** Must be typed as String.
                                    strThisProcErr = "'WRONG TYPE: " & CStr(lngLineNum) & "  " & strLine
                                    'Debug.Print strThisProcErr
                                  End If
                                Else
                                  ' ** There must be more words!
                                  strThisProcErr = "'WHAT? 2: " & CStr(lngLineNum) & "  " & strLine
                                  'Debug.Print strThisProcErr
                                End If
                              Else
                                'DON'T CARE.
                              End If
                            Else
                              ' ** Constant declarations need typing and assignment!
                              strThisProcErr = "'WHAT? 1: " & CStr(lngLineNum) & "  " & strLine
                              'Debug.Print strThisProcErr
                            End If

                          End If
                        End If
                      Else
                        ' ** Only 1 word.
                        If IsNumeric(strLine) = True Then
                          ' ** Line number with no code?!
                          blnNumbered = True
                          If IsMissing(varComp) = True Then
                            Beep
                            Debug.Print "'LINE NUM WITH NO CODE?! : " & lngX
                          Else
                            With varComp
                              .AddNew
                              ![dbs_id] = varDbsID
                              ![vbcom_id] = varComID
                              ![vbcomcomp_response] = "'LINE NUM WITH NO CODE?! : " & lngX
                              ![vbcomcomp_return] = blnRetValx
                              ![vbcomcomp_datemodified] = Now()
                              .Update
                            End With
                          End If
                        Else
                          'DON'T CARE
                        End If
                      End If  ' ** Multi-word line or 1-word line.

                    End If  ' ** Not a remark.

                  End If    ' ** blnProcStartFound.

                End If    ' ** Not a blank line.

              End If    ' ** Is a procedure.

            Next      ' ** For each code line: lngX.

            ' ** Save final procedure's info.
            If blnProcStartFound = True Then
              strTmp01 = vbNullString
              For lngY = 1& To lngWords
                strTmp01 = strTmp01 & " " & arr_varWord(lngY, 0)
              Next
              lngProcs = lngProcs + 1&
              lngE = lngProcs - 1&
              ReDim Preserve arr_varProc(P_ELEMS, lngE)
              arr_varProc(P_NAM, lngE) = strProcName
              arr_varProc(P_KND, lngE) = lngProcKind
              arr_varProc(P_KNDNAM, lngE) = strProcKind
              arr_varProc(P_START, lngE) = lngProcStart
              arr_varProc(P_END, lngE) = lngProcEnd
              arr_varProc(P_ERRH, lngE) = blnHasErrHand
              arr_varProc(P_ERRH_LINE, lngE) = lngErrHandLine
              arr_varProc(P_EXIT, lngE) = CBool(False)   ' ** Not figured out yet!
              arr_varProc(P_MOD_ELEM, lngE) = lngElemM
              arr_varProc(P_THISPROC, lngE) = blnHasThisProc
              arr_varProc(P_THISPROCE, lngE) = strThisProcErr
            End If  ' ** blnProcStartFound.

            ' *************************************************************************************
            ' ** Now analyze all the data.
            ' *************************************************************************************

            ' **   modUtilities: Standard Module
            ' **     Function EH    CopyToTempTable()
            ' **       CopyToTempTable_Exit: EXIT LABEL
            ' **       CopyToTempTable_Err:
            arr_varMod(M_PROCS, lngElemM) = lngProcs
            If lngProcs > 0 Then
              strTmp01 = " PROCS: " & lngProcs
            Else
              strTmp01 = " NO PROCS!"
            End If

            If IsMissing(varComp) = True Then
              Debug.Print "'" & arr_varMod(M_NAM, lngElemM) & "~"
            End If

            If lngProcs > 0& Then

              blnRetValx = True

              ' ** Report problems.
              For lngX = 0& To (lngProcs - 1&)
                If arr_varProc(P_THISPROC, lngX) = False Then
                  If arr_varProc(P_THISPROCE, lngX) = vbNullString Then
                    If IsMissing(varComp) = True Then
                      Debug.Print "'NO THIS_PROC: " & arr_varProc(P_NAM, lngX)
                    Else
                      With varComp
                        .AddNew
                        ![dbs_id] = varDbsID
                        ![vbcom_id] = varComID
                        ![vbcomcomp_response] = "'NO THIS_PROC: " & arr_varProc(P_NAM, lngX)
                        ![vbcomcomp_return] = blnRetValx
                        ![vbcomcomp_datemodified] = Now()
                        .Update
                      End With
                    End If
                  Else
                    If IsMissing(varComp) = True Then
                      Debug.Print arr_varProc(P_THISPROCE, lngX)
                    Else
                      With varComp
                        .AddNew
                        ![dbs_id] = varDbsID
                        ![vbcom_id] = varComID
                        ![vbcomcomp_response] = arr_varProc(P_THISPROCE, lngX)
                        ![vbcomcomp_return] = blnRetValx
                        ![vbcomcomp_datemodified] = Now()
                        .Update
                      End With
                    End If
                  End If
                End If
              Next  ' ** For each procedure: lngX.

              ' *******************************************************
              ' ** Array: arr_varProc()
              ' **
              ' **   Element  Description             Constant
              ' **   =======  ======================  ===============
              ' **      0     Procedure Name          P_NAM
              ' **      1     Procedure Kind          P_KND
              ' **      2     Procedure Kind Name     P_KNDNAM
              ' **      3     Start Line              P_START
              ' **      4     End Line                P_END
              ' **      5     Has Error Handler       P_ERRH
              ' **      6     Has Exit Label          P_EXIT
              ' **      7     arr_varMod() Element    P_MOD_ELEM
              ' **      8     THIS_PROC True/False    P_THISPROC
              ' **      9     THIS_PROC Error         P_THISPROCE
              ' **
              ' *******************************************************

              ' ** Add THIS_PROC.
              For lngX = (lngProcs - 1&) To 0& Step -1&
                If arr_varProc(P_THISPROC, lngX) = False And arr_varProc(P_THISPROCE, lngX) = vbNullString Then
                  If arr_varProc(P_ERRH, lngX) = True Then
                    ' ** Put if after the error-handler declaration.
                    .InsertLines (arr_varProc(P_ERRH_LINE, lngX) + 1&), TPROC & Chr(34) & arr_varProc(P_NAM, lngX) & Chr(34)
                    .InsertLines (arr_varProc(P_ERRH_LINE, lngX) + 1&), ""
                    lngThisProcsAdded = lngThisProcsAdded + 1&
                    arr_varProc(P_THISPROC, lngX) = True
                  Else
                    For lngY = arr_varProc(P_START, lngX) To arr_varProc(P_END, lngX)
                      If Trim(.Lines(lngY, 1)) = vbNullString Then
                        ' ** Put it after the first blank line.
                        .InsertLines (lngY + 1&), ""
                        .InsertLines (lngY + 1&), TPROC & Chr(34) & arr_varProc(P_NAM, lngX) & Chr(34)
                        arr_varProc(P_THISPROC, lngX) = True
                        lngThisProcsAdded = lngThisProcsAdded + 1&
                        Exit For
                      Else
                        strTmp01 = Trim(.Lines(lngY, 1))
                        If Left(strTmp01, 1) <> "'" Then
                          ' ** Not a remark.
                          intPos01 = InStr(strTmp01, " ")
                          If intPos01 > 0 Then
                            If IsNumeric(Trim(Left(strTmp01, intPos01))) = True Then
                              ' ** Put it before the first numbered line (since there is no error-handler declaration).
                              .InsertLines (lngY + 1&), ""
                              .InsertLines (lngY + 1&), TPROC & Chr(34) & arr_varProc(P_NAM, lngX) & Chr(34)
                              arr_varProc(P_THISPROC, lngX) = True
                              lngThisProcsAdded = lngThisProcsAdded + 1&
                              Exit For
                            Else
                              If Left(strTmp01, 3) = "Dim" Or Left(strTmp01, 5) = "Const" Then
                                ' ** Put it before any other declarations.
                                .InsertLines (lngY + 1&), ""
                                .InsertLines (lngY + 1&), TPROC & Chr(34) & arr_varProc(P_NAM, lngX) & Chr(34)
                                arr_varProc(P_THISPROC, lngX) = True
                                lngThisProcsAdded = lngThisProcsAdded + 1&
                                Exit For
                              End If
                            End If
                          End If
                        End If
                      End If
                    Next
                  End If
                End If
              Next  ' ** For each procedure: lngX.

            End If  ' ** lngProcs > 0&.

          End With  ' ** This CodeModule: cod.

          If lngProcs > 0& Then
            arr_varMod(M_PROC_ARR, lngElemM) = arr_varProc
          End If

          If strModName = varModName Then
            Exit For
          End If
        End If  ' ** Specified module.

      End With  ' ** This VBComponent: vbc.
    Next      ' ** For each VBComponent: vbc.
  End With  ' ** This ActiveProject: vbp.

  If lngThisProcsAdded > 0& Then
    If IsMissing(varComp) = True Then
      Debug.Print "'THIS_PROC'S ADDED: " & CStr(lngThisProcsAdded)
    Else
      With varComp
        .AddNew
        ![dbs_id] = varDbsID
        ![vbcom_id] = varComID
        ![vbcomcomp_response] = "'THIS_PROC'S ADDED: " & CStr(lngThisProcsAdded)
        ![vbcomcomp_return] = blnRetValx
        ![vbcomcomp_datemodified] = Now()
        .Update
      End With
    End If
  End If

  If IsMissing(varComp) = True Then
    Debug.Print blnRetValx
    Beep
  Else
    With varComp
      .AddNew
      ![dbs_id] = varDbsID
      ![vbcom_id] = varComID
      ![vbcomcomp_response] = "'" & varModName & "~"
      ![vbcomcomp_return] = blnRetValx
      ![vbcomcomp_datemodified] = Now()
      .Update
    End With
  End If

  VBA_This_Proc = blnRetValx

End Function

Private Function VBA_Err_Handler(Optional varModName As Variant, Optional varComp As Variant, Optional varComID As Variant, Optional varDbsID As Variant) As Boolean
' ** Check module error handling.
' ** Called by:
' **   QuikErrHandler(), Above
' **   QuikAll(), Above
'### LOTS OF UNIQUE VARIABLES!

  Const THIS_PROC As String = "VBA_Err_Handler"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strLine As String
  Dim lngModsChecked As Long
  Dim strModName As String, lngLines As Long
  Dim lngMods As Long, arr_varMod() As Variant
  Dim strProcName As String, lngProcLines As Long
  Dim lngProcs As Long, arr_varProc() As Variant
  Dim lngProcStart As Long, lngProcEnd As Long
  Dim blnProcStartFound As Boolean
  Dim strProcKind As String, lngProcKind As Long
  Dim blnHasScope As Boolean, blnHasErrHand As Boolean
  Dim blnNumbered As Boolean
  Dim blnRemarkFound As Boolean, intRemarkPos As Integer
  Dim arr_varWord() As Variant
  Dim lngLabels As Long, arr_varLabel() As Variant
  Dim lngRefs As Long, arr_varRef() As Variant
  Dim lngExits As Long, arr_varExit() As Variant
  Dim lngExitLine As Long, lngResumeLine As Long, lngErrHandLine As Long, lngErrZLine As Long, lngFormErrLine As Long
  Dim strExitName As String, strErrHandName As String
  Dim lngLen As Long, blnFound As Boolean, blnCase As Boolean, intCommaCnt As Integer
  Dim intPos01 As Integer, intPos02 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, arr_varTmp04 As Variant
  Dim lngElemM As Long, lngElemP As Long, lngElemL As Long, lngLineNum As Long
  Dim lngV As Long, lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  ' ** Array: arr_varMod().
  Const M_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const M_NAM      As Integer = 0
  Const M_TYP      As Integer = 1
  Const M_PROCS    As Integer = 2
  Const M_P_ARR As Integer = 3

  ' ** Array: arr_varProc().
  Const P_ELEMS As Integer = 13  ' ** Array's first-element UBound().
  Const P_NAME     As Integer = 0
  Const P_KIND     As Integer = 1
  Const P_KINDNAME As Integer = 2
  Const P_START    As Integer = 3
  Const P_END      As Integer = 4
  Const P_ERRH     As Integer = 5
  Const P_EXIT     As Integer = 6
  Const P_MOD_ELEM As Integer = 7
  Const P_LBLS     As Integer = 8
  Const P_LBL_ARR  As Integer = 9
  Const P_REFS     As Integer = 10
  Const P_REF_ARR  As Integer = 11
  Const P_EXITS    As Integer = 12
  Const P_EXIT_ARR As Integer = 13

  ' ** Array: arr_varWord().
  Const W_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const W_LIN As Integer = 0
  'Const W_1   As Integer = 1
  Const W_2   As Integer = 2
  'Const W_3   As Integer = 3
  'Const W_4   As Integer = 4
  'Const W_5   As Integer = 5
  'Const W_6   As Integer = 6

  Const lngWords As Long = 6&

  ' ** Array: arr_varLabel().
  Const L_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const L_LINE    As Integer = 0
  Const L_LINENUM As Integer = 1
  Const L_NAME    As Integer = 2
  Const L_ERRH    As Integer = 3
  Const L_EXIT    As Integer = 4
  Const L_REF     As Integer = 5

  ' ** Array: arr_varRef().
  Const R_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const R_LINE     As Integer = 0
  Const R_LINENUM  As Integer = 1
  Const R_ERR      As Integer = 2
  Const R_GOTO     As Integer = 3
  Const R_ZERO     As Integer = 4
  Const R_RESUME   As Integer = 5
  Const R_NEXT     As Integer = 6
  Const R_LABEL    As Integer = 7
  Const R_NUMBRD   As Integer = 8
  Const R_LBL_ELEM As Integer = 9

  ' ** Array: arr_varExit().
  Const E_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const E_LINE    As Integer = 0
  Const E_LINENUM As Integer = 1
  Const E_STMNT   As Integer = 2

  ' ** Analysis steps.
  Const SCAN_END   As Long = 1&
  Const SCAN_LABEL As Long = 2&
  Const SCAN_ERROR As Long = 3&
  Const SCAN_START As Long = 4&
  Const SCAN_EXIT  As Long = 5&
  Const SCAN_FIX   As Long = 6&
  Const SCAN_LIST  As Long = 7&

  blnRetValx = False

  lngMods = 0&
  ReDim arr_varMod(M_ELEMS, 0)
  ' **************************************************
  ' ** Array: arr_varMod()
  ' **
  ' **   Element  Description             Constant
  ' **   =======  ======================  ==========
  ' **      0     Module Name             M_NAM
  ' **      1     Module Type             M_TYP
  ' **      2     Number Of Procedures    M_PROCS
  ' **      3     arr_varProc() Array     M_P_ARR
  ' **
  ' **************************************************

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngModsChecked = 0&
    For Each vbc In .VBComponents
      With vbc

        lngModsChecked = lngModsChecked + 1&
        strModName = .Name

        If strModName = varModName Then

          blnRetValx = True

          ' ** Scan each module multiple times, because blank lines are being added.
          For lngV = 1& To SCAN_LIST

            lngProcs = 0&
            ReDim arr_varProc(P_ELEMS, 0)  ' ** ReDim after every module.
            ' ****************************************************
            ' ** Array: arr_varProc()
            ' **
            ' **   Element  Description             Constant
            ' **   =======  ======================  ============
            ' **      0     Procedure Name          P_NAME
            ' **      1     Procedure Kind          P_KIND
            ' **      2     Procedure Kind Name     P_KINDNAME
            ' **      3     Start Line              P_START
            ' **      4     End Line                P_END
            ' **      5     Has Error Handler       P_ERRH
            ' **      6     Has Exit Label          P_EXIT
            ' **      7     arr_varMod() Element    P_MOD_ELEM
            ' **      8     Number of Labels        P_LBLS
            ' **      9     arr_varLabel() Array    P_LBL_ARR
            ' **     10     Number of Refs          P_REFS
            ' **     11     arr_varRef() Array      P_REF_ARR
            ' **     12     Number of Exits         P_EXITS
            ' **     13     arr_varExit() Array     P_EXIT_ARR
            ' **
            ' ****************************************************

            ' ** Only 1 record.
            ReDim arr_varWord(W_ELEMS, 0)  ' ** ReDim after every procedure.
            ' ********************************************************
            ' ** Array: arr_varWord()
            ' **
            ' **   Element  Description                   Constant
            ' **   =======  ============================  ==========
            ' **      0     Procedure Declaration Line    W_LIN
            ' **      1     1st Word                      W_1
            ' **      2     2nd Word                      W_2
            ' **      3     3rd Word                      W_3
            ' **      4     4th Word                      W_4
            ' **      5     5th Word                      W_5
            ' **      6     6th Word                      W_6
            ' **
            ' ********************************************************

            lngLabels = 0&
            ReDim arr_varLabel(L_ELEMS, 0)  ' ** ReDim after every procedure.
            ' **************************************************
            ' ** Array: arr_varLabel()
            ' **
            ' **   Element  Description            Constant
            ' **   =======  =====================  ===========
            ' **      0     Line Text              L_LINE
            ' **      1     Line Number            L_LINENUM
            ' **      2     Label Name             L_NAME
            ' **      3     Is An Error Handler    L_ERRH
            ' **      4     Is An Exit             L_EXIT
            ' **      5     Is Referenced          L_REF
            ' **
            ' **************************************************

            lngRefs = 0&
            ReDim arr_varRef(R_ELEMS, 0)  ' ** ReDim after every procedure.
            ' ******************************************************
            ' ** Array: arr_varRef()
            ' **
            ' **   Element  Description               Constant
            ' **   =======  ========================  ============
            ' **      0     Line Text                 R_LINE
            ' **      1     Line Number               R_LINENUM
            ' **      2     Is On Error               R_ERR
            ' **      3     Has GoTo                  R_GOTO
            ' **      4     Has Goto 0                R_ZERO
            ' **      5     Has Resume                R_RESUME
            ' **      6     Has Resume Next           R_NEXT
            ' **      7     Label Referenced          R_LABEL
            ' **      8     Line Is Numbered          R_NUMBRD
            ' **      9     arr_varLabel() Element    R_LBL_ELEM
            ' **
            ' ******************************************************

            lngExits = 0&
            ReDim arr_varExit(E_ELEMS, 0)  ' ** ReDim after every procedure.
            ' *********************************************
            ' ** Array: arr_varExit()
            ' **
            ' **   Element  Description       Constant
            ' **   =======  ================  ===========
            ' **      0     Line Text         E_LINE
            ' **      1     Line Number       E_LINENUM
            ' **      2     Exit Statement    E_STMNT
            ' **
            ' *********************************************

            'Debug.Print "'" & .Name
            lngMods = lngMods + 1&
            lngE = lngMods - 1&
            ReDim Preserve arr_varMod(M_ELEMS, lngE)
            arr_varMod(M_NAM, lngE) = .Name
            arr_varMod(M_TYP, lngE) = .Type
            ' **   vbext_ComponentType enumeration:
            ' **       1  vbext_ct_StdModule        Standard Module
            ' **       2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
            ' **       3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm in the VBA Editor.
            ' **      11  vbext_ct_ActiveXDesigner
            ' **     100  vbext_ct_Document         Module behind Form, Report, or Excel Worksheet.
            arr_varMod(M_PROCS, lngE) = 0&
            strModName = .Name
            lngElemM = lngE

            Set cod = .CodeModule
            With cod

              lngLines = .CountOfLines
              strProcName = vbNullString
              lngProcLines = 0&
              lngProcStart = 0&: lngProcEnd = 0&
              blnProcStartFound = False
              strProcKind = vbNullString: lngProcKind = -1
              blnRemarkFound = False: intRemarkPos = 0
              blnHasScope = False: blnHasErrHand = False

              For lngX = 1& To lngLines

                blnNumbered = False
                strLine = vbNullString

                ' ** Declaration lines don't have a procedure name,
                ' ** and I don't believe they can have labels anyway.
                If .ProcOfLine(lngX, vbext_pk_Proc) <> vbNullString Then
                  ' ** Returns name of procedure that the specified line is in.
                  ' ** Doesn't care if type of procedure is incorrect.

                  If .ProcOfLine(lngX, vbext_pk_Proc) <> strProcName Then
                    ' ** A new procedure.
                    If blnProcStartFound = True Then
                      strTmp01 = vbNullString
                      For lngY = 1& To lngWords
                        strTmp01 = strTmp01 & " " & arr_varWord(lngY, 0)
                      Next
                      lngProcs = lngProcs + 1&
                      lngE = lngProcs - 1&
                      ReDim Preserve arr_varProc(P_ELEMS, lngE)
                      arr_varProc(P_NAME, lngE) = strProcName
                      arr_varProc(P_KIND, lngE) = lngProcKind
                      arr_varProc(P_KINDNAME, lngE) = strProcKind
                      arr_varProc(P_START, lngE) = lngProcStart
                      arr_varProc(P_END, lngE) = lngProcEnd
                      arr_varProc(P_ERRH, lngE) = blnHasErrHand  ' ** Not figured out yet!
                      arr_varProc(P_EXIT, lngE) = False          ' ** Not figured out yet!
                      arr_varProc(P_MOD_ELEM, lngE) = lngElemM
                      strTmp01 = vbNullString
                      If lngLabels > 0& Then
                        ' ** Check for an Exit label.
                        For lngY = 0& To (lngLabels - 1&)
                          lngW = lngProcEnd
                          For lngZ = 0& To (lngLabels - 1&)
                            If lngZ <> lngY Then
                              ' ** See if any labels come after this one.
                              If arr_varLabel(L_LINENUM, lngZ) > arr_varLabel(L_LINENUM, lngY) Then
                                ' ** Use this for the range to check, instead of lngProcEnd.
                                lngW = arr_varLabel(L_LINENUM, lngZ)
                                Exit For
                              End If
                            End If
                          Next
                          For lngZ = (arr_varLabel(L_LINENUM, lngY) + 1&) To (lngW - 1&)
                            ' ** Check these lines for an Exit statement.
                            strTmp01 = Trim(.Lines(lngZ, 1))
                            If strTmp01 <> vbNullString Then
                              intPos01 = InStr(strTmp01, "'")
                              If intPos01 > 0 Then
                                ' ** Strip off any remarks.
                                If intPos01 = 1 Then
                                  strTmp01 = vbNullString
                                Else
                                  strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
                                End If
                              End If
                              If strTmp01 <> vbNullString Then
                                If InStr(strTmp01, ("Exit " & strProcKind)) > 0 Then
                                  ' ** Yes, this is an Exit label.
                                  arr_varLabel(L_EXIT, lngY) = True
                                  Exit For
                                End If
                              End If
                            End If
                          Next
                        Next
                        If lngRefs > 0& Then
                          ' ** Cross-check Labels.
                          For lngY = 0& To (lngRefs - 1&)
                            For lngZ = 0& To (lngLabels - 1&)
                              If arr_varLabel(L_NAME, lngZ) = arr_varRef(R_LABEL, lngY) Then
                                arr_varRef(R_LBL_ELEM, lngY) = lngZ
                                arr_varLabel(L_REF, lngZ) = True
                                If arr_varRef(R_ERR, lngY) = True Then
                                  arr_varLabel(L_ERRH, lngZ) = True
                                End If
                                Exit For
                              End If
                            Next
                          Next  ' ** For each reference: lngY.
                          ' ** Update arr_varProc().
                          arr_varProc(P_LBLS, lngE) = lngLabels
                          If lngLabels > 0& Then arr_varProc(P_LBL_ARR, lngE) = arr_varLabel
                          arr_varProc(P_REFS, lngE) = lngRefs
                          If lngRefs > 0& Then arr_varProc(P_REF_ARR, lngE) = arr_varRef
                          For lngY = 0& To (lngLabels - 1&)
                            If arr_varLabel(L_ERRH, lngY) = True Then
                              blnHasErrHand = True
                              arr_varProc(P_ERRH, lngE) = blnHasErrHand
                            End If
                            If arr_varLabel(L_EXIT, lngY) = True Then
                              arr_varProc(P_EXIT, lngE) = True
                            End If
                          Next  ' ** For each label: lngY.
                        End If
                      Else
                        If lngRefs > 0& Then
                          arr_varProc(P_REF_ARR, lngE) = arr_varRef
                        End If
                      End If
                    End If  ' ** Save previous procedure's info.
                    ' ** Only 1 record.
                    ReDim arr_varWord(W_ELEMS, 0)
                    lngLabels = 0&
                    ReDim arr_varLabel(L_ELEMS, 0)
                    lngRefs = 0&
                    ReDim arr_varRef(R_ELEMS, 0)
                    lngExits = 0&
                    ReDim arr_varExit(E_ELEMS, 0)
                    strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                    strProcKind = vbNullString: lngProcKind = -1
                    blnProcStartFound = False
                    blnHasScope = False: blnHasErrHand = False
                    lngProcStart = 0&: lngProcLines = 0&: lngProcEnd = 0&
                  End If  ' ** New procedure.

                  strLine = Trim(.Lines(lngX, 1))

                  If strLine <> vbNullString Then

                    If blnProcStartFound = False Then

                      If Left(strLine, 1) <> "'" Then
                        ' ** It's not a remark.
                        intPos02 = InStr(strLine, " ")
                        If intPos02 > 0 Then
                          ' ** It's got a space, as a procedure declaration line will.

                          ' ** Collect the first 6 words in the line.
                          intPos01 = 1&
                          For lngY = 1& To lngWords
                            If arr_varWord(lngY, 0) = vbNullString Then
                              If lngY = 1& Then
                                arr_varWord(W_LIN, 0) = strLine
                              End If
                              arr_varWord(lngY, 0) = Trim(Mid(strLine, intPos01, (intPos02 - intPos01)))
                              If Left(arr_varWord(lngY, 0), 1) = "'" Then
                                ' ** Remark encountered.
                                blnRemarkFound = True
                                intRemarkPos = intPos01 + IIf(lngY = 1&, 0, 1)
                                arr_varWord(lngY, 0) = vbNullString
                                Exit For
                              Else
                                If InStr(arr_varWord(lngY, 0), "(") > 0 Then arr_varWord(lngY, 0) = _
                                  Left(arr_varWord(lngY, 0), (InStr(arr_varWord(lngY, 0), "(") - 1))
                                intPos01 = intPos02
                              End If
                            End If
                            intPos02 = InStr((intPos02 + 1), strLine, " ")
                            If intPos02 = 0 Then
                              Exit For
                            End If
                          Next

                        End If  ' ** Has at least 1 space.
                      Else
                        blnRemarkFound = True
                      End If  ' ** Not a remark.

                      ' ** Now pick up last word.
                      If blnRemarkFound = False And intPos01 > 0 Then
                        For lngY = 1& To lngWords
                          If arr_varWord(lngY, 0) = vbNullString Then
                            arr_varWord(lngY, 0) = Trim(Mid(strLine, intPos01))
                            If InStr(arr_varWord(lngY, 0), "(") > 0 Then
                              arr_varWord(lngY, 0) = _
                                Left(arr_varWord(lngY, 0), (InStr(arr_varWord(lngY, 0), "(") - 1))
                            End If
                            Exit For
                          End If
                        Next
                      End If

                      If arr_varWord(W_2, 0) <> vbNullString Then
                        ' ** Procedures declaration will have at least 1 space.

                        blnHasScope = False

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
                              blnRetValx = False
                              If IsMissing(varComp) = True Then
                                Beep
                                Debug.Print "'UNKNOWN PROPERTY TYPE! : " & lngX & " " & strTmp01 & " " & strTmp02
                              Else
                                With varComp
                                  .AddNew
                                  ![dbs_id] = varDbsID
                                  ![vbcom_id] = varComID
                                  ![vbcomcomp_response] = "'UNKNOWN PROPERTY TYPE! : " & lngX & " " & strTmp01 & " " & strTmp02
                                  ![vbcomcomp_return] = blnRetValx
                                  ![vbcomcomp_datemodified] = Now()
                                  .Update
                                End With
                              End If
                            End Select
                          Case "Let", "Set", "Get"
                            ' ** Usually 3rd Word.
                            If lngProcKind = -1& Then
                              If IsMissing(varComp) = True Then
                                Debug.Print "'WHY KIND NOT FOUND ON PREVIOUS LOOP? : " & lngY & " LINE: " & lngX & " '" & strTmp01 & "'"
                              Else
                                With varComp
                                  .AddNew
                                  ![dbs_id] = varDbsID
                                  ![vbcom_id] = varComID
                                  ![vbcomcomp_response] = "'WHY KIND NOT FOUND ON PREVIOUS LOOP? : " & lngY & " LINE: " & lngX & " '" & strTmp01 & "'"
                                  ![vbcomcomp_return] = blnRetValx
                                  ![vbcomcomp_datemodified] = Now()
                                  .Update
                                End With
                              End If
                            End If
                          Case Else
                            If strTmp01 = strProcName Then
                              ' ** Usually 4th word.
                              If lngY < lngWords Then
                                For lngZ = (lngY + 1&) To lngWords
                                  arr_varWord(lngZ, 0) = vbNullString
                                Next
                              End If
                              If lngProcKind <> -1& Then
                                blnProcStartFound = True
                                lngProcStart = lngX '.ProcStartLine(strProcName, lngProcKind)
                                lngProcLines = .ProcCountLines(strProcName, lngProcKind)
                                lngZ = (lngProcStart + lngProcLines)
                                strTmp02 = "End " & strProcKind
                                strTmp03 = Trim(.Lines(lngZ, 1))
                                intPos01 = InStr(strTmp03, strTmp02)
                                Do While intPos01 = 0
                                  lngZ = lngZ - 1&
                                  intPos01 = InStr(.Lines(lngZ, 1), strTmp02)
                                  If lngZ = lngX Then
                                    lngZ = 0&
                                    Exit Do  ' ** Exit the loop!
                                  End If
                                Loop
                                If lngZ > 0& Then
                                  lngProcEnd = lngZ
                                Else
                                  If IsMissing(varComp) = True Then
                                    Debug.Print "'STOP : P_END NOT FOUND! " & strProcName & " " & lngX
                                  Else
                                    With varComp
                                      .AddNew
                                      ![dbs_id] = varDbsID
                                      ![vbcom_id] = varComID
                                      ![vbcomcomp_response] = "'STOP : P_END NOT FOUND! " & strProcName & " " & lngX
                                      ![vbcomcomp_return] = blnRetValx
                                      ![vbcomcomp_datemodified] = Now()
                                      .Update
                                    End With
                                  End If
                                End If
                              End If
                              Exit For
                            Else
                              ' ** Unknown 1st word; not a procedure declaration line.
                              blnRetValx = False
                              'Beep
                              If IsMissing(varComp) = True Then
                                Debug.Print "'UNKNOWN FIRST WORD! : lngX = " & lngX & ", lngY = " & lngY & " " & strTmp01
'LOTS OF "UNKNOWN FIRST WORD!" MEANS A COMPLETELY REMARKED OUT PROCEDURE IS THROWING OFF THE PROCESS.
                              Else
                                With varComp
                                  .AddNew
                                  ![dbs_id] = varDbsID
                                  ![vbcom_id] = varComID
                                  ![vbcomcomp_response] = "'UNKNOWN FIRST WORD! : lngX = " & lngX & ", lngY = " & lngY & " " & strTmp01
                                  ![vbcomcomp_return] = blnRetValx
                                  ![vbcomcomp_datemodified] = Now()
                                  .Update
                                End With
                              End If
                            End If
                            Exit For
                          End Select

                        Next  ' ** For each word: lngY.

                      End If  ' ** Possible declaration line: arr_varWord(W_2, 0).

                    Else
                      ' ** We're within a procedure.

                      If Left(strLine, 1) <> "'" Then
                        ' ** Not a remark.
                        intPos01 = InStr(strLine, " ")
                        If intPos01 > 0 Then
                          strTmp01 = Trim(Left(strLine, intPos01))
                          If IsNumeric(strTmp01) Then
                            blnNumbered = True
                            strTmp01 = Trim(Mid(strLine, intPos01))
                          End If
                          If Left(strTmp01, 8) = "On Error" Then
                            lngRefs = lngRefs + 1&
                            lngE = lngRefs - 1&
                            ReDim Preserve arr_varRef(R_ELEMS, lngE)
                            arr_varRef(R_LINE, lngE) = strLine
                            arr_varRef(R_LINENUM, lngE) = lngX
                            arr_varRef(R_ERR, lngE) = True
                            arr_varRef(R_GOTO, lngE) = False
                            arr_varRef(R_ZERO, lngE) = False
                            arr_varRef(R_RESUME, lngE) = False
                            arr_varRef(R_NEXT, lngE) = False
                            arr_varRef(R_LABEL, lngE) = vbNullString
                            arr_varRef(R_NUMBRD, lngE) = blnNumbered
                            arr_varRef(R_LBL_ELEM, lngE) = -1&
                            strTmp01 = Trim(Mid(strTmp01, 9))
                            If Left(strTmp01, 5) = "GoTo " Then
                              arr_varRef(R_GOTO, lngE) = True
                              strTmp01 = Trim(Mid(strTmp01, 5))
                              If Left(strTmp01, 1) = "0" Then
                                If Len(strTmp01) > 1 Then
                                  If Mid(strTmp01, 2, 1) = " " Then
                                    ' ** No label reference.
                                    arr_varRef(R_ZERO, lngE) = True
                                  Else
                                    ' ** A label beginning with "0"?
                                    intPos01 = InStr(strTmp01, " ")
                                    If intPos01 > 0 Then
                                      strTmp03 = Trim(Left(strTmp01, intPos01))
                                    Else
                                      strTmp03 = strTmp01
                                    End If
                                    If Right(strTmp03, 1) = ":" Then
                                      strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
                                    End If
                                    arr_varRef(R_LABEL, lngE) = strTmp03
                                  End If
                                Else
                                  ' ** No label reference.
                                  arr_varRef(R_ZERO, lngE) = True
                                End If
                              Else
                                intPos01 = InStr(strTmp01, " ")
                                If intPos01 > 0 Then
                                  strTmp03 = Trim(Left(strTmp01, intPos01))
                                Else
                                  strTmp03 = strTmp01
                                End If
                                If Right(strTmp03, 1) = ":" Then
                                  strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
                                End If
                                arr_varRef(R_LABEL, lngE) = strTmp03
                              End If
                            Else
                              ' ** Might be "Resume Next".
                              If Left(strTmp01, 7) = "Resume " Then
                                arr_varRef(R_RESUME, lngE) = True
                                strTmp01 = Trim(Mid(strTmp01, 7))
                                intPos01 = InStr(strTmp01, " ")
                                If intPos01 > 0 Then
                                  If Left(strTmp01, 5) = "Next " Then
                                    ' ** No label reference.
                                    arr_varRef(R_NEXT, lngE) = True
                                  Else
                                    strTmp03 = Trim(Left(strTmp01, intPos01))
                                    If Right(strTmp03, 1) = ":" Then
                                      strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
                                    End If
                                    arr_varRef(R_LABEL, lngE) = strTmp03
                                  End If
                                Else
                                  If strTmp01 = "Next" Then
                                    ' ** No label reference.
                                    arr_varRef(R_NEXT, lngE) = True
                                  Else
                                    strTmp03 = strTmp01
                                    If Right(strTmp03, 1) = ":" Then
                                      strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
                                    End If
                                    arr_varRef(R_LABEL, lngE) = strTmp03
                                  End If
                                End If
                              Else
                                ' ** Unknown word.
                                arr_varRef(R_LABEL, lngE) = strTmp01
                                If IsMissing(varComp) = True Then
                                  Beep
                                  Debug.Print "'UNKNOWN WORD! : " & strTmp01 & " " & strLine
                                Else
                                  With varComp
                                    .AddNew
                                    ![dbs_id] = varDbsID
                                    ![vbcom_id] = varComID
                                    ![vbcomcomp_response] = "'UNKNOWN WORD! : " & strTmp01 & " " & strLine
                                    ![vbcomcomp_return] = blnRetValx
                                    ![vbcomcomp_datemodified] = Now()
                                    .Update
                                  End With
                                End If
                              End If
                            End If
                          Else
                            ' ** Doesn't begin with "On Error"; number already removed.
                            intPos01 = InStr(strTmp01, "'")
                            If intPos01 > 0 Then
                              ' ** There's a remark on the line, so strip that off.
                              strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
                            End If
                            lngLen = Len(strTmp01)
                            intPos01 = InStr(strTmp01, " ")
                            If intPos01 > 0 Then
                              strTmp01 = Trim(Left(strTmp01, intPos01))
                              If Right(strTmp01, 1) = ":" Then 'And InStr(strTmp01, "MsgBox") = 0 Then
                                ' ** It's a label.
                                lngLabels = lngLabels + 1&
                                lngE = lngLabels - 1&
                                ReDim Preserve arr_varLabel(L_ELEMS, lngE)
                                arr_varLabel(L_LINE, lngE) = strLine
                                arr_varLabel(L_LINENUM, lngE) = lngX
                                arr_varLabel(L_NAME, lngE) = Left(strTmp01, (Len(strTmp01) - 1))
                                arr_varLabel(L_ERRH, lngE) = False
                                arr_varLabel(L_EXIT, lngE) = False
                                arr_varLabel(L_REF, lngE) = False
                              Else
                                ' ** Might still be a "Resume" or "GoTo" statement.
                                ' ** strTmp01 IS ONLY THE FIRST WORD!
                                intPos01 = InStr(strLine, " GoTo ")
                                If intPos01 > 0 Then
                                  lngRefs = lngRefs + 1&
                                  lngE = lngRefs - 1&
                                  ReDim Preserve arr_varRef(R_ELEMS, lngE)
                                  arr_varRef(R_LINE, lngE) = strLine
                                  arr_varRef(R_LINENUM, lngE) = lngX
                                  arr_varRef(R_ERR, lngE) = False
                                  arr_varRef(R_GOTO, lngE) = True
                                  arr_varRef(R_ZERO, lngE) = False
                                  arr_varRef(R_RESUME, lngE) = False
                                  arr_varRef(R_NEXT, lngE) = False
                                  arr_varRef(R_LABEL, lngE) = vbNullString
                                  arr_varRef(R_NUMBRD, lngE) = blnNumbered
                                  arr_varRef(R_LBL_ELEM, lngE) = -1&
                                  strTmp01 = Trim(Mid(strLine, (intPos01 + 5)))
                                  intPos01 = InStr(strTmp01, " ")
                                  If intPos01 > 0 Then
                                    strTmp03 = Trim(Left(strTmp01, intPos01))
                                  Else
                                    strTmp03 = strTmp01
                                  End If
                                  If Right(strTmp03, 1) = ":" Then
                                    strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
                                  End If
                                  arr_varRef(R_LABEL, lngE) = strTmp03
                                Else
                                  intPos01 = InStr(strLine, " Resume ")
                                  If intPos01 > 0 Then
                                    lngRefs = lngRefs + 1&
                                    lngE = lngRefs - 1&
                                    ReDim Preserve arr_varRef(R_ELEMS, lngE)
                                    arr_varRef(R_LINE, lngE) = strLine
                                    arr_varRef(R_LINENUM, lngE) = lngX
                                    arr_varRef(R_ERR, lngE) = False
                                    arr_varRef(R_GOTO, lngE) = False
                                    arr_varRef(R_ZERO, lngE) = False
                                    arr_varRef(R_RESUME, lngE) = True
                                    arr_varRef(R_NEXT, lngE) = False
                                    arr_varRef(R_LABEL, lngE) = vbNullString
                                    arr_varRef(R_NUMBRD, lngE) = blnNumbered
                                    arr_varRef(R_LBL_ELEM, lngE) = -1&
                                    strTmp01 = Trim(Mid(strLine, (intPos01 + 7)))
                                    intPos01 = InStr(strTmp01, " ")
                                    If intPos01 > 0 Then
                                      strTmp03 = Trim(Left(strTmp01, intPos01))
                                    Else
                                      strTmp03 = strTmp01
                                    End If
                                    If Right(strTmp03, 1) = ":" Then
                                      strTmp03 = Left(strTmp03, (Len(strTmp03) - 1))
                                    End If
                                    arr_varRef(R_LABEL, lngE) = strTmp03
                                  Else
                                    ' ** Something else; not interested.
                                  End If
                                End If
                              End If
                            Else
                              ' ** Only 1 word; number already removed.
                              If Right(strTmp01, 1) = ":" And InStr(strTmp01, Chr(34)) = 0 Then
                                ' ** It's a label though labels aren't numbered!
                                ' ** Make sure the colon is not within a quoted string;
                                ' ** it was identifying words inside a MsgBox as a label.
                                ' **   L_LINENUM: 417
                                ' **   L_LINE: "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()", _
                                ' **   L_NAME: "Module:
                                ' **   L_ERRH: False
                                ' **   L_EXIT: False
                                ' **   L_REF:  False
                                lngLabels = lngLabels + 1&
                                lngE = lngLabels - 1&
                                ReDim Preserve arr_varLabel(L_ELEMS, lngE)
                                arr_varLabel(L_LINE, lngE) = strLine
                                arr_varLabel(L_LINENUM, lngE) = lngX
                                arr_varLabel(L_NAME, lngE) = strTmp01
                                arr_varLabel(L_ERRH, lngE) = False
                                arr_varLabel(L_EXIT, lngE) = False
                                arr_varLabel(L_REF, lngE) = False
                              Else
                                ' ** Something else; not interested.
                              End If
                            End If
                          End If
                        Else
                          ' ** Only 1 word.
                          If Right(strLine, 1) = ":" Then
                            ' ** It's a label.
                            lngLabels = lngLabels + 1&
                            lngE = lngLabels - 1&
                            ReDim Preserve arr_varLabel(L_ELEMS, lngE)
                            arr_varLabel(L_LINE, lngE) = strLine
                            arr_varLabel(L_LINENUM, lngE) = lngX
                            arr_varLabel(L_NAME, lngE) = Left(strLine, (Len(strLine) - 1))
                            arr_varLabel(L_ERRH, lngE) = False
                            arr_varLabel(L_EXIT, lngE) = False
                            arr_varLabel(L_REF, lngE) = False
                          ElseIf IsNumeric(strLine) = True Then
                            ' ** Line number with no code?!
                            blnNumbered = True
                            If IsMissing(varComp) = True Then
                              Beep
                              Debug.Print "'LINE NUM WITH NO CODE?! : " & lngX
                            Else
                              With varComp
                                .AddNew
                                ![dbs_id] = varDbsID
                                ![vbcom_id] = varComID
                                ![vbcomcomp_response] = "'LINE NUM WITH NO CODE?! : " & lngX
                                ![vbcomcomp_return] = blnRetValx
                                ![vbcomcomp_datemodified] = Now()
                                .Update
                              End With
                            End If
                          End If
                        End If  ' ** Multi-word line or 1-word line.

                      End If  ' ** Not a remark.

                    End If    ' ** blnProcStartFound.

                  End If    ' ** Not a blank line.

                End If    ' ** Is a procedure.

              Next      ' ** For each code line: lngX.

              ' ** Save final procedure's info.
              If blnProcStartFound = True Then
                strTmp01 = vbNullString
                For lngY = 1& To lngWords
                  strTmp01 = strTmp01 & " " & arr_varWord(lngY, 0)
                Next
                lngProcs = lngProcs + 1&
                lngE = lngProcs - 1&
                ReDim Preserve arr_varProc(P_ELEMS, lngE)
                arr_varProc(P_NAME, lngE) = strProcName
                arr_varProc(P_KIND, lngE) = lngProcKind
                arr_varProc(P_KINDNAME, lngE) = strProcKind
                arr_varProc(P_START, lngE) = lngProcStart
                arr_varProc(P_END, lngE) = lngProcEnd
                arr_varProc(P_ERRH, lngE) = blnHasErrHand  ' ** Not figured out yet!
                arr_varProc(P_EXIT, lngE) = CBool(False)   ' ** Not figured out yet!
                arr_varProc(P_MOD_ELEM, lngE) = lngElemM
                'Debug.Print "'" & Trim(strTmp01) & "()" & " " & lngProcStart & " - " & lngProcEnd
                strTmp01 = vbNullString
                If lngLabels > 0& Then
                  ' ** Check for Exit label.
                  For lngY = 0& To (lngLabels - 1&)
                    lngW = lngProcEnd
                    For lngZ = 0& To (lngLabels - 1&)
                      If lngZ <> lngY Then
                        ' ** See if any labels come after this one.
                        If arr_varLabel(L_LINENUM, lngZ) > arr_varLabel(L_LINENUM, lngY) Then
                          ' ** Use this for the range to check, instead of lngProcEnd.
                          lngW = arr_varLabel(L_LINENUM, lngZ)
                          Exit For
                        End If
                      End If
                    Next  ' ** For each label: lngZ.
                    For lngZ = (arr_varLabel(L_LINENUM, lngY) + 1&) To (lngW - 1&)
                      ' ** Check these lines for an Exit statement.
                      strTmp01 = Trim(.Lines(lngZ, 1))
                      If strTmp01 <> vbNullString Then
                        intPos01 = InStr(strTmp01, "'")
                        If intPos01 > 0 Then
                          ' ** Strip off any remarks.
                          If intPos01 = 1 Then
                            strTmp01 = vbNullString
                          Else
                            strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
                          End If
                        End If
                        If strTmp01 <> vbNullString Then
                          If InStr(strTmp01, ("Exit " & strProcKind)) > 0 Then
                            ' ** Yes, this is an Exit label.
                            arr_varLabel(L_EXIT, lngY) = True
                            Exit For
                          End If
                        End If
                      End If
                    Next  ' ** For each line: lngZ.
                  Next  ' ** For each label: lngY.
                  If lngRefs > 0& Then
                    ' ** Cross-check Labels.
                    For lngY = 0& To (lngRefs - 1&)
                      For lngZ = 0& To (lngLabels - 1&)
                        If arr_varLabel(L_NAME, lngZ) = arr_varRef(R_LABEL, lngY) Then
                          arr_varRef(R_LBL_ELEM, lngY) = lngZ
                          arr_varLabel(L_REF, lngZ) = True
                          If arr_varRef(R_ERR, lngY) = True Then
                            arr_varLabel(L_ERRH, lngZ) = True
                          End If
                          Exit For
                        End If
                      Next
                    Next  ' ** For each reference: lngY.
                    ' ** Update arr_varProc().
                    arr_varProc(P_LBL_ARR, lngE) = arr_varLabel
                    arr_varProc(P_REF_ARR, lngE) = arr_varRef
                    For lngY = 0& To (lngLabels - 1&)
                      If arr_varLabel(L_ERRH, lngY) = True Then
                        blnHasErrHand = True
                        arr_varProc(P_ERRH, lngE) = blnHasErrHand
                      End If
                      If arr_varLabel(L_EXIT, lngY) = True Then
                        arr_varProc(P_EXIT, lngE) = True
                      End If
                    Next  ' ** For each label: lngY.
                  End If  ' ** lngRefs > 0&.
                Else
                  If lngRefs > 0& Then
                    arr_varProc(P_REF_ARR, lngE) = arr_varRef
                  End If
                End If  ' ** lngLabels > 0&.
                arr_varProc(P_LBLS, lngE) = lngLabels
                arr_varProc(P_LBL_ARR, lngE) = arr_varLabel
                arr_varProc(P_REFS, lngE) = lngRefs
                arr_varProc(P_REF_ARR, lngE) = arr_varRef
                arr_varProc(P_EXITS, lngE) = CLng(0&)
              End If  ' ** blnProcStartFound.

              ' *************************************************************************************
              ' ** Now analyze all the data.
              ' *************************************************************************************

              ' **   modUtilities: Standard Module
              ' **     Function EH    CopyToTempTable()
              ' **       CopyToTempTable_Exit: EXIT LABEL
              ' **       CopyToTempTable_Err:
              arr_varMod(M_PROCS, lngElemM) = lngProcs
              If lngProcs > 0 Then
                strTmp01 = " PROCS: " & lngProcs
              Else
                strTmp01 = " NO PROCS!"
              End If
              If lngV = 2& Then
                If IsMissing(varComp) = True Then
                  Debug.Print "'" & arr_varMod(M_NAM, lngElemM) & "³"
                    '& ": " & VBA_Component_Type (arr_varMod(M_TYP, lngElemM)) & strTmp01 ' ** Function: Below.
                End If
              End If

              If lngProcs > 0& Then

                ' ** Check for procedures with multiple Exit statements.
                ' ** A procedure should have only 1 entry point and 1 exit point!
                ' ** Tracing program steps can become very difficult when
                ' ** it suddenly exits at various points mid-procedure.
                For lngX = 0& To (lngProcs - 1&)
                  lngExits = 0&
                  ReDim arr_varExit(E_ELEMS, 0)  ' ** ReDim after every procedure.
                  For lngY = arr_varProc(P_START, lngX) To arr_varProc(P_END, lngX)
                    strTmp01 = Trim(.Lines(lngY, 1))
                    If strTmp01 <> vbNullString Then
                      ' ** Ignore blank lines.
                      If Left(strTmp01, 1) <> "'" Then
                        ' ** Ignore remark lines.
                        intPos01 = InStr(strTmp01, " ")
                        If intPos01 > 0 Then
                          ' ** Exit statements will have a space.
                          strTmp02 = "Exit " & arr_varProc(P_KINDNAME, lngX)
                          intPos01 = InStr(strTmp01, strTmp02)
                          If intPos01 > 0 Then
                            strTmp03 = Trim(Left(strTmp01, (intPos01 + (Len(strTmp02) - 1))))
                            intPos01 = InStr(strTmp03, "'")
                            If intPos01 = 0 Then
                              ' ** Make sure the Exit statement isn't in a remark.
                              lngExits = lngExits + 1&
                              lngE = lngExits - 1&
                              ReDim Preserve arr_varExit(E_ELEMS, lngE)
                              arr_varExit(E_LINE, lngE) = strTmp01
                              arr_varExit(E_LINENUM, lngE) = lngY
                              arr_varExit(E_STMNT, lngE) = strTmp02
                            End If
                          End If
                        End If
                      End If
                    End If
                  Next  ' ** For each line: lngY.
                  arr_varProc(P_EXITS, lngX) = lngExits
                  If lngExits > 0& Then
                    arr_varProc(P_EXIT_ARR, lngX) = arr_varExit
                  End If
                Next  ' ** For each procedure: lngX.

' ****************************************************  ' ******************************************************
' ** Array: arr_varMod()                                ' ** Array: arr_varLabel()
' **                                                    ' **
' **   Element  Description             Constant        ' **   Element  Description            Constant
' **   =======  ======================  ==========      ' **   =======  =====================  ===========
' **      0     Module Name             M_NAM           ' **      0     Line Text              L_LINE
' **      1     Module Type             M_TYP           ' **      1     Line Number            L_LINENUM
' **      2     Number Of Procedures    M_PROCS         ' **      2     Label Name             L_NAME
' **      3     arr_varProc() Array     M_P_ARR         ' **      3     Is An Error Handler    L_ERRH
' **                                                    ' **      4     Is An Exit             L_EXIT
' ****************************************************  ' **      5     Is Referenced          L_REF
' ** Array: arr_varProc()                               ' **
' **                                                    ' ******************************************************
' **   Element  Description             Constant        ' ** Array: arr_varRef()
' **   =======  ======================  ============    ' **
' **      0     Procedure Name          P_NAME          ' **   Element  Description               Constant
' **      1     Procedure Kind          P_KIND          ' **   =======  ========================  ============
' **      2     Procedure Kind Name     P_KINDNAME      ' **      0     Line Text                 R_LINE
' **      3     Start Line              P_START         ' **      1     Line Number               R_LINENUM
' **      4     End Line                P_END           ' **      2     Is On Error               R_ERR
' **      5     Has Error Handler       P_ERRH          ' **      3     Has GoTo                  R_GOTO
' **      6     Has Exit Label          P_EXIT          ' **      4     Has Goto 0                R_ZERO
' **      7     arr_varMod() Element    P_MOD_ELEM      ' **      5     Has Resume                R_RESUME
' **      8     Number of Labels        P_LBLS          ' **      6     Has Resume Next           R_NEXT
' **      9     arr_varLabel() Array    P_LBL_ARR       ' **      7     Label Referenced          R_LABEL
' **     10     Number of Refs          P_REFS          ' **      8     Line Is Numbered          R_NUMBRD
' **     11     arr_varRef() Array      P_REF_ARR       ' **      9     arr_varLabel() Element    R_LBL_ELEM
' **     12     Number of Exits         P_EXITS         ' **
' **     13     arr_varExit() Array     P_EXIT_ARR      ' ******************************************************
' **
' ****************************************************
' ** Array: arr_varExit()
' **
' **   Element  Description       Constant
' **   =======  ================  ===========
' **      0     Line Text         E_LINE
' **      1     Line Number       E_LINENUM
' **      2     Exit Statement    E_STMNT
' **
' ****************************************************

                Select Case lngV
                Case SCAN_END
                  ' ** Add blank line above procedure end line, if necessary.
                  For lngX = (lngProcs - 1&) To 0& Step -1&
                    If .Lines((arr_varProc(P_END, lngX) - 1&), 1) <> vbNullString Then
                      .InsertLines arr_varProc(P_END, lngX), ""
                      'All line numbers after this now need to be incremented by 1!
                    End If
                  Next  ' ** For each procedure: lngX.

                Case SCAN_LABEL
                  ' ** Add blank line above each label, if necessary.
                  For lngX = (lngProcs - 1&) To 0& Step -1&  ' ** Start at end of proc.
                    If IsEmpty(arr_varProc(P_LBL_ARR, lngX)) = False Then  ' ** This proc has labels.
                      For lngY = UBound(arr_varProc(P_LBL_ARR, lngX), 2) To 0& Step -1&  ' ** For each label.
                        If .Lines((arr_varProc(P_LBL_ARR, lngX)(L_LINENUM, lngY) - 1&), 1) <> vbNullString Then
                          .InsertLines arr_varProc(P_LBL_ARR, lngX)(L_LINENUM, lngY), ""
                          'All line numbers after this now need to be incremented by 1!
                        End If
                      Next
                    End If
                  Next  ' ** For each procedure: lngX.

                Case SCAN_ERROR
                  ' ** Add blank line after first On Error, if necessary.
                  For lngX = (lngProcs - 1&) To 0& Step -1&
                    If IsEmpty(arr_varProc(P_REF_ARR, lngX)) = False Then
                      For lngY = 0& To UBound(arr_varProc(P_REF_ARR, lngX), 2)
                        If lngY = 0& Then  ' ** Only do it for first On Error.
                          If arr_varProc(P_REF_ARR, lngX)(R_ERR, lngY) = True Then
                            If .Lines((arr_varProc(P_REF_ARR, lngX)(R_LINENUM, lngY) + 1&), 1) <> vbNullString Then
                              .InsertLines (arr_varProc(P_REF_ARR, lngX)(R_LINENUM, lngY) + 1&), ""
                              Exit For
                              'All line numbers after this now need to be incremented by 1!
                            End If
                          End If
                        End If
                      Next
                    End If
                  Next  ' ** For each procedure: lngX.

                Case SCAN_START
                  ' ** Add blank line between procedure start line and first On Error, if necessary.
                  For lngX = (lngProcs - 1&) To 0& Step -1&
                    If IsEmpty(arr_varProc(P_REF_ARR, lngX)) = False Then
                      For lngY = 0& To UBound(arr_varProc(P_REF_ARR, lngX), 2)
                        If arr_varProc(P_REF_ARR, lngX)(R_ERR, lngY) = True Then
                          If (arr_varProc(P_REF_ARR, lngX)(R_LINENUM, lngY) - 1&) = arr_varProc(P_START, lngX) Then
                            .InsertLines (arr_varProc(P_START, lngX) + 1&), ""
                          End If
                        End If
                      Next
                    End If
                  Next  ' ** For each procedure: lngX.

                Case SCAN_EXIT
                  ' ** Add Exit label and Resume statement, if necessary.
                  For lngX = (lngProcs - 1&) To 0& Step -1&
                    If arr_varProc(P_EXIT, lngX) = False Then
                      If IsEmpty(arr_varProc(P_EXIT_ARR, lngX)) = False Then
                        If arr_varProc(P_ERRH, lngX) = False Then
                          If IsMissing(varComp) = True Then
                            Beep
                            Debug.Print "'EXIT W/O ERROR HANDLER! : " & arr_varProc(P_NAME, lngX)
                          Else
                            With varComp
                              .AddNew
                              ![dbs_id] = varDbsID
                              ![vbcom_id] = varComID
                              ![vbcomcomp_response] = "'EXIT W/O ERROR HANDLER! : " & arr_varProc(P_NAME, lngX)
                              ![vbcomcomp_return] = blnRetValx
                              ![vbcomcomp_datemodified] = Now()
                              .Update
                            End With
                          End If
                        Else
                          lngE = UBound(arr_varProc(P_EXIT_ARR, lngX), 2)  ' ** Last Exit statement.
                          lngExitLine = arr_varProc(P_EXIT_ARR, lngX)(E_LINENUM, lngE)
                          lngResumeLine = arr_varProc(P_END, lngX)
                          ' ** Find out where to put the Resume statement.
                          strTmp02 = Trim(.Lines((lngResumeLine - 1&), 1))
                          If strTmp02 = vbNullString Then
                            lngResumeLine = lngResumeLine - 1&
                          End If
                          strTmp01 = Trim(.Lines((lngExitLine - 2&), 1))
                          strTmp02 = Trim(.Lines((lngExitLine - 1&), 1))
                          strTmp01 = "  ~" & strTmp01 & "~" & vbCrLf & "  ~" & strTmp02 & "~"
                          strTmp03 = vbNullString
                          If strTmp02 <> vbNullString Then
                            strTmp03 = "  ~~" & vbCrLf
                          End If
                          strTmp03 = strTmp03 & "  ~" & "EXITP:" & "~"
                          strTmp01 = strTmp01 & vbCrLf & strTmp03
                          strTmp02 = Trim(.Lines(lngExitLine, 1))
                          strTmp01 = strTmp01 & vbCrLf & "  ~" & strTmp02 & "~"
                          For lngY = lngExitLine + 1& To arr_varProc(P_END, lngX)
                            If lngY = lngResumeLine Then
                              strTmp03 = "95      Resume EXITP"
                              strTmp01 = strTmp01 & vbCrLf & "  ~" & strTmp03 & "~"
                            End If
                            strTmp02 = Trim(.Lines(lngY, 1))
                            strTmp01 = strTmp01 & vbCrLf & "  ~" & strTmp02 & "~"
                          Next
                          If MsgBox("Add Exit label and Resume statement here?" & vbCrLf & vbCrLf & _
                             strTmp01, vbYesNo, "Add EXITP: Label?") = vbYes Then
                            strTmp03 = "95      Resume EXITP"
                            .InsertLines lngResumeLine, strTmp03
                            strTmp02 = Trim(.Lines((lngExitLine - 1&), 1))
                            strTmp03 = vbNullString
                            If strTmp02 <> vbNullString Then
                              strTmp03 = "" & vbCrLf
                            End If
                            strTmp03 = strTmp03 & "EXITP:"
                            .InsertLines lngExitLine, strTmp03
                          End If
                        End If
                      End If
                    End If
                  Next  ' ** For each procedure: lngX.

                Case SCAN_FIX

                  For lngX = (lngProcs - 1&) To 0& Step -1&
                    lngElemP = lngX

                    strProcName = arr_varProc(P_NAME, lngElemP)
                    lngProcEnd = arr_varProc(P_END, lngElemP)
                    lngExitLine = 0&: lngErrHandLine = 0&
                    strExitName = vbNullString: strErrHandName = vbNullString

                    lngLabels = arr_varProc(P_LBLS, lngElemP)
                    If lngLabels > 0& Then
                      arr_varTmp04 = arr_varProc(P_LBL_ARR, lngElemP)
                      For lngY = 0& To UBound(arr_varTmp04, 2)
                        lngElemL = lngY
                        If arr_varTmp04(L_ERRH, lngElemL) = True Then
                          lngErrHandLine = arr_varTmp04(L_LINENUM, lngElemL)
                          strErrHandName = arr_varTmp04(L_NAME, lngElemL)
                        ElseIf arr_varTmp04(L_EXIT, lngElemL) = True Then
                          lngExitLine = arr_varTmp04(L_LINENUM, lngElemL)
                          strExitName = arr_varTmp04(L_NAME, lngElemL)
                        End If
                      Next  ' ** For each label: lngY.
                    End If

                    If lngErrHandLine > 0& Then

                      If strExitName <> vbNullString Then
                        ' ** Check GoTo vs. Resume.
                        For lngY = (lngErrHandLine + 1) To (lngProcEnd - 1&)
                          lngLineNum = lngY
                          strLine = .Lines(lngLineNum, 1)  ' ** Don't trim.
                          intPos01 = InStr(strLine, "GoTo ")
                          If intPos01 > 0 Then
                            ' ** If this references the Exit label, replace it with a Resume.
                            If Mid(strLine, (intPos01 + 5), Len(strExitName)) = strExitName Then
                              strTmp01 = Left(strLine, (intPos01 - 1)) & "Resume" & Mid(strLine, (intPos01 + 4))
                              .ReplaceLine lngLineNum, strTmp01
                            End If
                          End If
                        Next
                        ' ** Check for presence of GoTo or Resume.
                        blnFound = False
                        For lngY = (lngErrHandLine + 1) To (lngProcEnd - 1&)
                          lngLineNum = lngY
                          strLine = .Lines(lngLineNum, 1)  ' ** Don't trim.
                          intPos01 = InStr(strLine, "GoTo ")
                          If intPos01 > 0 Then blnFound = True
                          intPos01 = InStr(strLine, "Resume ")
                          If intPos01 > 0 Then blnFound = True
                        Next
                        If blnFound = False Then
                          strTmp01 = "10      Resume " & strExitName
                          If Trim(.Lines((lngProcEnd - 1&), 1)) = vbNullString Then
                            .InsertLines (lngProcEnd - 1&), strTmp01
                            lngProcEnd = lngProcEnd + 1&
                          Else
                            .InsertLines (lngProcEnd - 1&), ""
                            lngProcEnd = lngProcEnd + 1&
                            .InsertLines (lngProcEnd - 1&), strTmp01
                            lngProcEnd = lngProcEnd + 1&
                          End If
                        End If
                      End If  ' ** Has Exit label: strExitName.

                      ' ** Check for Err vs. Err.Number in Case framework.
                      blnCase = False
                      For lngY = (lngErrHandLine + 1) To (lngProcEnd - 1&)
                        lngLineNum = lngY
                        strLine = .Lines(lngLineNum, 1)  ' ** Don't trim.
                        intPos01 = InStr(strLine, "Select Case Err")
                        If intPos01 > 0 Then
                          blnCase = True
                          If intPos01 + 15 < Len(strLine) Then
                            strTmp01 = Mid(strLine, (intPos01 + 15))
                            If Left(strTmp01, 7) = ".Number" Then
                              ' ** As it should be.
                            Else
                              ' ** Well, what is it?
                              If IsMissing(varComp) = True Then
                                Beep
                                Debug.Print "'ODD CASE: '" & strLine & "'"
                              Else
                                With varComp
                                  .AddNew
                                  ![dbs_id] = varDbsID
                                  ![vbcom_id] = varComID
                                  ![vbcomcomp_response] = "'ODD CASE: '" & strLine & "'"
                                  ![vbcomcomp_return] = blnRetValx
                                  ![vbcomcomp_datemodified] = Now()
                                  .Update
                                End With
                              End If
                            End If
                          Else
                            ' ** Doesn't have ".Number".
                            strTmp01 = strLine & ".Number"
                            .ReplaceLine lngLineNum, strTmp01
                          End If
                        End If
                      Next

                      ' ** Check for zErrorHandler() (doesn't handle multiples!).
                      blnFound = False
                      For lngY = (lngErrHandLine + 1) To (lngProcEnd - 1&)
                        lngLineNum = lngY
                        strLine = .Lines(lngLineNum, 1)  ' ** Don't trim.
                        intPos01 = InStr(strLine, "zErrorHandler")
                        If intPos01 > 0 Then
                          blnFound = True
                          lngErrZLine = lngLineNum
                          Exit For
                        End If
                      Next

                      If blnFound = True Then

                        strLine = .Lines(lngErrZLine, 1)  ' ** Don't trim.

                        ' ** Check for all 4 parameters.
                        intCommaCnt = 0
                        intPos01 = InStr(strLine, ",")
                        Do While intPos01 > 0
                          intCommaCnt = intCommaCnt + 1
                          intPos01 = InStr((intPos01 + 1), strLine, ",")
                        Loop
                        strTmp01 = vbNullString
                        Select Case intCommaCnt
                        Case 1
                          ' ** Parameters 1 & 2 only.
                          intPos01 = InStr(strLine, ", ")
                          intPos02 = InStr(intPos01, strLine, "'")  ' ** In case there's a Remark.
                          strTmp01 = ", Err.Number, Erl"
                          If intPos02 > 0 Then
                            strTmp01 = Left(strLine, (intPos02 - 1)) & strTmp01 & "  " & Mid(strLine, intPos02)
                          Else
                            strTmp01 = strLine & strTmp01
                          End If
                        Case 2
                          ' ** Parameters 1, 2, & 3 only.
                          intPos01 = InStr((InStr(strLine, ", ") + 1), strLine, ", ")
                          intPos02 = InStr(intPos01, strLine, "'")  ' ** In case there's a Remark.
                          strTmp01 = ", Erl"
                          If intPos02 > 0 Then
                            strTmp01 = Left(strLine, (intPos02 - 1)) & strTmp01 & "  " & Mid(strLine, intPos02)
                          Else
                            strTmp01 = strLine & strTmp01
                          End If
                        Case 3
                          ' ** All 4 parameters.
                          intPos01 = InStr(strLine, ", ,")  ' ** No Err.Number.
                          If intPos01 > 0 Then
                            strTmp01 = Left(strLine, (intPos01 + 1)) & "Err.Number" & Mid(strLine, (intPos01 + 2))
                          End If
                          intPos01 = InStr(strLine, ", Err,")  ' ** No .Number.
                          If intPos01 > 0 Then
                            strTmp01 = Left(strLine, (intPos01 + 4)) & ".Number" & Mid(strLine, (intPos01 + 5))
                          End If
                        Case Else
                          ' ** Both 1st and 2nd parameters are required, so this won't happen.
                        End Select  ' ** Check parameters.

                        ' ** Check for correct procedure name.
                        intPos01 = InStr(strLine, "THIS_PROC")
                        If intPos01 > 0 And strTmp01 <> vbNullString Then
                          .ReplaceLine lngLineNum, strTmp01
                        Else
                          If strTmp01 <> vbNullString Then
                            intPos01 = InStr(strTmp01, ", ")
                            intPos02 = InStr((intPos01 + 1), strTmp01, ", ")
                            strTmp02 = Trim(Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1)))
                            If Left(strTmp02, 1) = Chr(34) And Right(strTmp02, 1) = Chr(34) Then
                              ' ** Presumably it's surrounded by quotes, though it could be a different variable.
                              intPos01 = InStr(strTmp01, strTmp02)       ' ** Open quotes.
                              intPos02 = (intPos01 + Len(strTmp02)) - 1  ' ** Close quotes.
                              strTmp02 = Mid(Left(strTmp02, (Len(strTmp02) - 1)), 2)
                              If strTmp02 <> strProcName Then
                                strTmp01 = Left(strTmp01, intPos01) & strProcName & Mid(strTmp01, intPos02)
                              End If
                            Else
                              If IsMissing(varComp) = True Then
                                Debug.Print "'² " & lngLineNum & " " & strTmp02
'Stop
                              Else
                                With varComp
                                  .AddNew
                                  ![dbs_id] = varDbsID
                                  ![vbcom_id] = varComID
                                  ![vbcomcomp_response] = "'² " & lngLineNum & " " & strTmp02
                                  ![vbcomcomp_return] = blnRetValx
                                  ![vbcomcomp_datemodified] = Now()
                                  .Update
                                End With
                              End If
                            End If
                          Else
                            strTmp01 = strLine
                            intPos01 = InStr(strTmp01, ", ")
                            intPos02 = InStr((intPos01 + 1), strTmp01, ", ")
                            strTmp02 = Trim(Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1)))
                            If Left(strTmp02, 1) = Chr(34) And Right(strTmp02, 1) = Chr(34) Then
                              ' ** Presumably it's surrounded by quotes, though it could be a different variable.
                              intPos01 = InStr(strTmp01, strTmp02)       ' ** Open quotes.
                              intPos02 = (intPos01 + Len(strTmp02)) - 1  ' ** Close quotes.
                              strTmp02 = Mid(Left(strTmp02, (Len(strTmp02) - 1)), 2)
                              If strTmp02 <> strProcName Then
                                strTmp01 = Left(strTmp01, intPos01) & strProcName & Mid(strTmp01, intPos02)
                              Else
                                strTmp01 = vbNullString
                              End If
                            Else
                              If InStr(strTmp02, "THIS_PROC") = 0 Then
                                If varModName = "Form_frmRpt_CourtReports_CA" Or _
                                    varModName = "Form_frmRpt_CourtReports_FL" Or _
                                    varModName = "Form_frmRpt_CourtReports_NS" Then
                                  ' ** I know about these.
                                  'UNUSUAL WORD IN zERRORHANDLER¹ 2212 strControlName
                                  'UNUSUAL WORD IN zERRORHANDLER¹ 1960 strControlName
                                  'UNUSUAL WORD IN zERRORHANDLER¹ 1939 strControlName
                                  'UNUSUAL WORD IN zERRORHANDLER¹ 1589 strControlName
                                  'UNUSUAL WORD IN zERRORHANDLER¹ 3628 strControlName
                                  If strTmp02 <> "strControlName" Then
                                    If IsMissing(varComp) = True Then
                                      Debug.Print "'UNUSUAL WORD IN zERRORHANDLER¹ " & lngLineNum & " " & strTmp02
                                    Else
                                      With varComp
                                        .AddNew
                                        ![dbs_id] = varDbsID
                                        ![vbcom_id] = varComID
                                        ![vbcomcomp_response] = "'UNUSUAL WORD IN zERRORHANDLER¹ " & lngLineNum & " " & strTmp02
                                        ![vbcomcomp_return] = blnRetValx
                                        ![vbcomcomp_datemodified] = Now()
                                        .Update
                                      End With
                                    End If
                                  End If
                                Else
                                  If IsMissing(varComp) = True Then
                                    Debug.Print "'UNUSUAL WORD IN zERRORHANDLER¹ " & lngLineNum & " " & strTmp02
                                  Else
                                    With varComp
                                      .AddNew
                                      ![dbs_id] = varDbsID
                                      ![vbcom_id] = varComID
                                      ![vbcomcomp_response] = "'UNUSUAL WORD IN zERRORHANDLER¹ " & lngLineNum & " " & strTmp02
                                      ![vbcomcomp_return] = blnRetValx
                                      ![vbcomcomp_datemodified] = Now()
                                      .Update
                                    End With
                                  End If
                                End If
                              End If
                            End If
                          End If
                          If strTmp01 <> vbNullString Then
                            .ReplaceLine lngLineNum, strTmp01
                          End If
                        End If  ' ** Check proc name.

                        ' ** Check for Select Case framework.
                        If blnCase = False Then
                          ' ** Only add if currently a simple, 2-line error handler.
                          ' **   ERRH:
                          ' **   160     zErrorHandler, THIS_NAME, THIS_PROC, Err.Number, Erl
                          ' **   170     Resume EXITP
                          ' **
                          ' **   End Sub
                          If lngErrZLine = lngErrHandLine + 1& And lngProcEnd = lngErrHandLine + 4& Then
                            If Trim(.Lines((lngProcEnd - 1&), 1)) = vbNullString And _
                               InStr(.Lines((lngErrHandLine + 2), 1), "Resume") > 0 Then
                            .InsertLines (lngProcEnd - 2&), "46      End Select"
                            lngProcEnd = lngProcEnd + 1&
                            .InsertLines lngErrZLine, "        Case Else"
                            lngProcEnd = lngProcEnd + 1&
                            .InsertLines (lngErrHandLine + 1&), "45      Select Case Err.Number"
                            lngProcEnd = lngProcEnd + 1&
                            End If
                          End If
                        End If  ' ** No Select Case framework.

                      End If  ' ** Has zErrorHandler(): lngErrZLine.

                      ' ** Check for Form_Error() reference.
                      For lngY = (lngErrHandLine + 1) To (lngProcEnd - 1&)
                        lngLineNum = lngY
                        strLine = .Lines(lngLineNum, 1)  ' ** Don't trim.
                        intPos01 = InStr(strLine, "Call Form_Error(ERR.Number, 1)")
                        If intPos01 > 0 Then
                          lngFormErrLine = lngLineNum
                          strTmp02 = "Above."
                          If Left(strProcName, 5) = "Form_" Then
                            If strProcName <> "Form_Unload" And strProcName <> "Form_Close" Then
                              strTmp02 = "Below."
                            End If
                          End If
                          strTmp01 = Left(strLine, (intPos01 - 1)) & _
                            "Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: " & strTmp02
                          .ReplaceLine lngFormErrLine, strTmp01
                        End If
                      Next

                    End If  ' ** Error Handler line: lngErrHandLine.

                  Next  ' ** For each procedure: lngX.

                Case SCAN_LIST

                  For lngX = 0& To (lngProcs - 1&)

                    ' ** List only procedures missing pieces.
                    strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
                    If arr_varProc(P_ERRH, lngX) = False Then
                      ' ** If the procedure has no error handler, and no exit statement,
                      ' ** but does have a resume next statement, then it's OK.
                      If IsEmpty(arr_varProc(P_EXIT_ARR, lngX)) = True Then
                        ' ** No Exit statement.
                        strTmp01 = "OK"
                      End If
                      If IsEmpty(arr_varProc(P_REF_ARR, lngX)) = False Then
                        lngRefs = arr_varProc(P_REFS, lngX)
                        If lngRefs > 0& Then
                          arr_varTmp04 = arr_varProc(P_REF_ARR, lngX)
                          For lngY = 0& To UBound(arr_varTmp04, 2)
                            strTmp03 = arr_varTmp04(R_NEXT, lngY)
                            If arr_varTmp04(R_NEXT, lngY) = True Then
                              ' ** Has Resume Next statement.
                              strTmp02 = "OK"
                              Exit For
                            End If
                          Next
                        End If
                      End If
                    End If
                    If strTmp01 = "OK" And strTmp02 = "OK" Then
                      strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
                    Else
                      strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
                      If arr_varProc(P_ERRH, lngX) = False Then
                        strTmp02 = " NO EH!"
                        strTmp01 = strTmp01 & strTmp02
                      End If
                      If arr_varProc(P_EXIT, lngX) = False Then
                        strTmp02 = " NO EXIT!"
                        strTmp01 = strTmp01 & strTmp02
                      End If
                    End If
                    If IsEmpty(arr_varProc(P_REF_ARR, lngX)) = False Then
                      For lngY = 0& To UBound(arr_varProc(P_REF_ARR, lngX), 2)
                        If arr_varProc(P_REF_ARR, lngX)(R_NUMBRD, lngY) = False Then
                          strTmp02 = arr_varProc(P_REF_ARR, lngX)(R_LINENUM, lngY) & " NOT NUMBERED!"
                          strTmp01 = strTmp01 & strTmp02
                        End If
                      Next
                    End If
                    If strTmp01 <> vbNullString Then
                      If IsMissing(varComp) = True Then
                        Debug.Print "'  " & arr_varProc(P_NAME, lngX) & "() " & arr_varProc(P_START, lngX) & strTmp01
                      Else
                        With varComp
                          .AddNew
                          ![dbs_id] = varDbsID
                          ![vbcom_id] = varComID
                          ![vbcomcomp_response] = "'  " & arr_varProc(P_NAME, lngX) & "() " & arr_varProc(P_START, lngX) & strTmp01
                          ![vbcomcomp_return] = blnRetValx
                          ![vbcomcomp_datemodified] = Now()
                          .Update
                        End With
                      End If
                    End If

                    ' ** List only procedures with multiple exit statements.
                    If IsEmpty(arr_varProc(P_EXIT_ARR, lngX)) = False Then
                      lngW = UBound(arr_varProc(P_EXIT_ARR, lngX), 2)
                      If lngW > 0& Then
                        If strTmp01 = vbNullString Then
                          'Debug.Print "'  " & arr_varProc(P_NAME, lngX) & _
                          '  "() EXITS: " & (lngW + 1&) & " " & arr_varProc(P_START, lngX)
                        Else
                          'Debug.Print "'  MULTI-EXITS: " & (lngW + 1&) & " " & arr_varProc(P_START, lngX)
                        End If
                      End If
                    End If

                    ' ** List details about every procedure.
                    'strTmp02 = "'  " & Left(arr_varProc(P_KINDNAME, lngX) & Space(8), 8)
                    'strTmp01 = strTmp02 & " "
                    'If arr_varProc(P_ERRH, lngX) = True Then
                    '  strTmp02 = "EH   "
                    'Else
                    '  strTmp02 = "NO EH"
                    'End If
                    'strTmp01 = strTmp01 & strTmp02 & " "
                    'strTmp02 = vbNullString
                    'If arr_varProc(P_EXIT, lngX) = False Then
                    '  strTmp02 = "NO EXIT"
                    '  strTmp01 = strTmp01 & strTmp02 & " "
                    'Else
                    '  strTmp01 = strTmp01 & " "
                    'End If
                    'strTmp01 = strTmp01 & arr_varProc(P_NAME, lngX) & "()"
                    'Debug.Print strTmp01
                    'For lngY = 0& To UBound(arr_varProc(P_LBL_ARR, lngX), 2)
                    '  strTmp01 = vbNullString
                    '  If arr_varProc(P_LBL_ARR, lngX)(L_EXIT, lngY) = True Then
                    '    strTmp01 = " EXIT LABEL"
                    '  End If
                    '  If arr_varProc(P_LBL_ARR, lngX)(L_REF, lngY) = False Then
                    '    strTmp01 = strTmp01 & " NOT REFERENCED"
                    '  End If
                    '  Debug.Print "'    " & arr_varProc(P_LBL_ARR, lngX)(L_NAME, lngY) & ":" & strTmp01
                    'Next

                  Next  ' ** For each procedure: lngX.

                End Select  ' ** For each module scan: lngV.

              End If  ' ** lngProcs > 0&.

            End With  ' ** This CodeModule: cod.

          Next  ' ** Scan each module multiple times: lngV.

          If lngProcs > 0& Then
            arr_varMod(M_P_ARR, lngElemM) = arr_varProc
          End If

          If strModName = varModName Then
            Exit For
          End If
        End If  ' ** Specified module.

      End With  ' ** This VBComponent: vbc.
    Next      ' ** For each VBComponent: vbc.
  End With  ' ** This ActiveProject: vbp.

  If IsMissing(varComp) = True Then
    Debug.Print blnRetValx
    Beep
  Else
    With varComp
      .AddNew
      ![dbs_id] = varDbsID
      ![vbcom_id] = varComID
      ![vbcomcomp_response] = "'" & varModName & "³"
      ![vbcomcomp_return] = blnRetValx
      ![vbcomcomp_datemodified] = Now()
      .Update
    End With
  End If

  ' ** .CodeModule.ProcStartLine(ProcName As String, ProcKind As vbext_ProcKind) As Long
  ' **   Returns line number at which specified procedure begins. (including preceding blank lines and comments)
  ' **   Will error unless the type of procedure is correct.
  ' ** .CodeModule.ProcBodyLine(ProcName As String, ProcKind As vbext_ProcKind) As Long
  ' **   Returns line number of first line in specified procedure. (Declare line)
  ' **   Will error unless the type of procedure is correct.
  ' ** .CodeModule.ProcCountLines(ProcName As String,ProcKind As vbext_ProcKind) As Long
  ' **   Returns number of lines in specified procedure.
  ' **   Will error unless the type of procedure is correct.
  ' ** .CodeModule.ProcOfLine(Line As Long, ProcKind As vbext_ProcKind) As String
  ' **   Returns name of procedure that specified line is in.
  ' **   Doesn't care if type of procedure is incorrect.
  ' ** .CodeModule.Lines(StartLine As Long, Count As Long) As String
  ' **   Returns specified line.
  ' ** .CodeModule.InsertLines(Line as Long, String As String) Method
  ' **   Inserts a line or lines of code at a specified location in a block of code.
  ' ** Modules(0).InsertText Method
  ' **   When you insert a string by using the InsertText method, Microsoft Access
  ' **   places the new text at the end of the module, after all other procedures.
  ' ** .CodeModule.ReplaceLine(Line As Long, String As String) Method
  ' **   The ReplaceLine method replaces a specified line in a standard module.

  Set vbp = Nothing
  Set vbc = Nothing
  Set cod = Nothing

  VBA_Err_Handler = blnRetValx

End Function

Private Function VBA_Chk_Ctls(Optional varFrmName As Variant) As Boolean
' ** Called by:
' **   QuikChkCtls(), Above
'### LOTS OF UNIQUE VARIABLES!

  Const THIS_PROC As String = "VBA_Chk_Ctls"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim frm As Access.Form, ctl As Access.Control
  Dim strCtlName As String
  Dim lngCtls As Long, arr_varCtl() As Variant
  Dim strModName As String
  Dim lngLines As Long, lngLineNum As Long
  Dim strLine As String, lngFixes As Long, lngRems As Long
  Dim blnLoop As Boolean, blnIgnore As Boolean, blnQuote As Boolean, blnEnd As Boolean
  Dim blnSpace1 As Boolean, blnSpace2 As Boolean, blnBracket1 As Boolean, blnBracket2 As Boolean
  Dim blnOp1 As Boolean, blnOp2 As Boolean, blnSep1 As Boolean, blnSep2 As Boolean
  Dim blnFound As Boolean
  Dim intPos01 As Integer, intPos02 As Integer
  Dim strTmp01 As String, strTmp02 As String
  Dim lngX As Long, lngY As Long, lngElemC As Long

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const C_NAM   As Integer = 0
  Const C_TYP   As Integer = 1
  Const C_SPACE As Integer = 2
  Const C_SLASH As Integer = 3

  blnRetValx = True

  lngCtls = 0&
  ReDim arr_varCtl(C_ELEMS, 0)

If Left(varFrmName, 5) = "Form_" Then

  DoCmd.OpenForm varFrmName, acDesign, , , , acHidden
  Set frm = Forms(varFrmName)
  With frm
    lngCtls = .Controls.Count
    ReDim arr_varCtl(C_ELEMS, (lngCtls - 1&))
    For lngX = 0& To (lngCtls - 1&)
      lngElemC = lngX
      Set ctl = .Controls(lngElemC)
      With ctl
        strCtlName = .Name
        arr_varCtl(C_NAM, lngElemC) = strCtlName
        arr_varCtl(C_TYP, lngElemC) = .ControlType
        If InStr(strCtlName, " ") > 0 Then
          arr_varCtl(C_SPACE, lngElemC) = CBool(True)
        Else
          arr_varCtl(C_SPACE, lngElemC) = CBool(False)
        End If
        If InStr(strCtlName, "/") > 0 Then
          arr_varCtl(C_SLASH, lngElemC) = CBool(True)
        Else
          arr_varCtl(C_SLASH, lngElemC) = CBool(False)
        End If
     End With
    Next
  End With
  DoCmd.Close acForm, varFrmName, acSaveNo

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    strModName = "Form_" & varFrmName
    Debug.Print "'MOD: " & strModName & " CTLS: " & CStr(lngCtls)
    Set vbc = .VBComponents(strModName)
    With vbc
      Set cod = .CodeModule
      With cod
        lngLines = .CountOfLines
        lngFixes = 0&: lngRems = 0&
        For lngX = 0& To (lngCtls - 1&)

          lngElemC = lngX
          strCtlName = arr_varCtl(C_NAM, lngElemC)

          For lngY = 1& To lngLines
            lngLineNum = lngY
            strLine = Trim(.Lines(lngLineNum, 1))
            If strLine <> vbNullString Then
              If Left(strLine, 1) <> "'" Then

                blnLoop = True

                intPos01 = InStr(strLine, strCtlName)
                If intPos01 > 0 Then
                  ' ** OK, now we figure out how to isolate it and
                  ' ** determine if it should be prefixed with "Me."

                  If VBA_IsSQL_Ctl(cod, lngLineNum, intPos01) = False Then  ' ** Function: Below.

                    ' ** Repeat till done.
                    Do While blnLoop = True

                      blnIgnore = False: blnQuote = False: blnEnd = False
                      blnSpace1 = False: blnSpace2 = False: blnBracket1 = False: blnBracket2 = False
                      blnOp1 = False: blnOp2 = False: blnSep1 = False: blnSep2 = False

                      ' ** See what's just before it.
                      If intPos01 > 1 Then
                        strTmp02 = Mid(strLine, (intPos01 - 1), 1)
                        If strTmp02 <> " " Then
                          If (Asc(strTmp02) >= 65 And Asc(strTmp02) <= 90) Or _
                             (Asc(strTmp02) >= 97 And Asc(strTmp02) <= 122) Or _
                             (Asc(strTmp02) >= 48 And Asc(strTmp02) <= 57) Then
                            ' ** UpperCase, LowerCase, Numeral.
                            ' ** Not the control, ignore this instance of it.
                            blnIgnore = True
                          Else
                            Select Case strTmp02
                            Case "_"
                              ' ** Connector, so not the control, ignore this instance of it.
                              blnIgnore = True
                            Case "'"
                              ' ** Remark, ignore it and move on to another line.
                              blnIgnore = True
                              blnLoop = False
                            Case Chr(34)  ' "
                              ' ** Most likely not for this instance, but keep track of it.
                              blnIgnore = True
                              blnQuote = True
                              If InStr(strLine, "THIS_PROC") > 0 Or InStr(strLine, "zErrorHandler") > 0 Then
                                ' ** These instances are OK as-is.
                              Else
                                ' ** These instances are probably OK as-is.
                                'Debug.Print "'¹ " & Mid(strLine, (intPos01 - 1))
                              End If
                            Case "(", ")", "[", "]", "+", "-", "/", "*", "&", "^", ","
                              ' ** An operator, so it looks good.
                              blnOp1 = True
                              If strTmp02 = "[" Then
                                blnBracket1 = True
                                If Mid(strLine, (intPos01 - 2), 1) = "." Or Mid(strLine, (intPos01 - 2), 1) = "!" Then
                                  blnSep1 = True
                                End If
                              End If
                            Case ".", "!"
                              blnSep1 = True
                            Case Else
                               ' ** Let's see what it found.
                               blnIgnore = True
                               Debug.Print "'² " & strTmp02
                            End Select  ' ** Prfix.
                          End If  ' ** Prefix check.
                        Else
                          blnSpace1 = True
                        End If  ' ** Prefix Space or not.
                      Else
                        blnSpace1 = True
                      End If  ' ** Beginning of line or not.

                      ' ** Now see what's after it.
                      If blnIgnore = False Then
                        strTmp01 = Mid(strLine, intPos01)
                        intPos02 = InStr(strTmp01, " ")
                        If arr_varCtl(C_SPACE, lngElemC) = True Then
                          intPos02 = InStr((Len(strCtlName) - 2), strTmp01, " ")
                        End If
                        If intPos02 > 0 Then
                          strTmp01 = Left(strTmp01, (intPos02 - 1))
                        End If
                        blnFound = False
                        If Len(strTmp01) = Len(strCtlName) Then
                          ' ** Yes, it's the control.
                          blnFound = True
                          blnEnd = True
                        Else
                          strTmp02 = Mid(strTmp01, (Len(strCtlName) + 1), 1)
                          If (Asc(strTmp02) >= 65 And Asc(strTmp02) <= 90) Or _
                             (Asc(strTmp02) >= 97 And Asc(strTmp02) <= 122) Or _
                             (Asc(strTmp02) >= 48 And Asc(strTmp02) <= 57) Then
                            ' ** UpperCase, LowerCase, Numeral.
                            ' ** Not the control, ignore this instance of it.
                            blnIgnore = True
                          Else
                            Select Case strTmp02
                            Case "_"
                              ' ** Connector, so not the control, ignore this instance of it.
                            Case "'"
                              ' ** Remark, odd place, but looks good.
                              blnFound = True
                            Case Chr(34)  ' "
                              ' ** Odd, but let's keep going.
                              blnFound = True
                            Case "(", ")", "[", "]", "+", "-", "/", "*", "&", "^", ","
                              ' ** An operator, so it looks good.
                              blnFound = True
                              blnOp2 = True
                              If strTmp02 = "]" Then blnBracket2 = True
                            Case ".", "!"
                              blnFound = True
                              blnSep2 = True
                            Case ";"
                              ' ** End of a SQL statement, ignore it.
                            Case Else
                               ' ** Let's see what it found.
                               Debug.Print "'³ " & CStr(lngLineNum) & " '" & strTmp02 & "'"
                            End Select  ' ** Suffix.
                          End If  ' ** Suffix check.
                        End If  ' ** Suffix check.

                        ' ** OK, now figure out what we've got.
                        If blnFound = True Then

                          If (blnSep1 = True And blnSep2 = True) Or (blnSep1 = True And blnOp2 = True) Or _
                             (blnSep1 = True And blnEnd = True) Then
                            If blnSep1 = True Then
                              If Mid(strLine, (intPos01 - 3), 3) = "Me." Then
                                ' ** These really look like they're good, so ignore this instance of it.
                                'Debug.Print "'OK '" & strCtlName & "' " & strLine
                              Else
                                ' ** Often within SQL, so ignore this instance of it.
                                'SELECT, INSERT INTO, DELETE
                                'Debug.Print "'OK1 '" & strCtlName & "' " & strLine
                              End If
                            Else
                              Debug.Print "'OK2 '" & strCtlName & "' " & strLine
                            End If
                          Else
                            If (blnOp1 = True And blnOp2 = True) Or (blnOp1 = True And blnSep2 = True) Or _
                               (blnSep1 = True And blnOp2 = True) Or _
                               (blnSpace1 = True And (blnOp2 = True Or blnSep2 = True Or blnEnd = True)) Or _
                               (blnOp1 = True And blnEnd = True) Then
                              ' ** Now fix the line.
                              If blnBracket1 = True Then
                                ' ** If it's a control with a space, brackets must be used,
                                ' ** so the "Me." must go before the opening bracket.
                                ' ** Also, brackets might indicate it's actually within a SQL string!
                                blnFound = False
                                If blnBracket1 = True And blnBracket2 = True Then
                                  If arr_varCtl(C_SPACE, lngElemC) = False And _
                                     arr_varCtl(C_SLASH, lngElemC) = False Then
                                    ' ** Brackets aren't really necessary.
                                    blnFound = True
                                  End If
                                End If
                                If blnFound = True Then
                                  intPos02 = InStr(intPos01, strLine, "]")
                                  strTmp01 = Left(strLine, (intPos02 - 1)) & Mid(strLine, (intPos02 + 1))
                                  strTmp01 = Left(strTmp01, (intPos01 - 2)) & "Me." & Mid(strTmp01, intPos01)
                                  intPos01 = intPos01 + 3
                                Else
                                  strTmp01 = Left(strLine, (intPos01 - 2)) & "Me." & Mid(strLine, (intPos01 - 1))
                                  intPos01 = intPos01 + 3
                                  blnFound = True  ' ** Just reset it where it was.
                                End If
                              Else
                                strTmp01 = Left(strLine, (intPos01 - 1)) & "Me." & Mid(strLine, intPos01)
                                intPos01 = intPos01 + 3
                              End If
                              .ReplaceLine lngLineNum, strTmp01
                              lngFixes = lngFixes + 1&
                              strLine = .Lines(lngLineNum, 1)
                              'Debug.Print "'FIX '" & strCtlName & "' " & strLine
                            Else
                              Debug.Print "'?  " & blnSpace1 & " " & blnEnd & " '" & strCtlName & "' " & strLine
                            End If
                          End If

                        End If  ' ** blnFound.
                      End If  ' ** blnIgnore.

                      intPos01 = InStr((intPos01 + 1), strLine, strCtlName)
                      If intPos01 = 0 Then blnLoop = False

                    ' ** Repeat till done.
                    Loop

                  End If  ' ** Not within SQL.

                End If  ' ** Control name found.

              Else
                If strLine = "'" Then
                  ' ** Whole line is just a Remark character, blank it out.
                  'Debug.Print "'REM¹ " & .Lines(lngLineNum, 1)
                  .ReplaceLine lngLineNum, ""
                  lngRems = lngRems + 1&
                Else
                  ' ** How can I remove those giant blank spaces of remarked-out code?
                  If Left(strLine, 3) = "'  " And Mid(strLine, 4) = " " Then
                    'Debug.Print "'² " & .Lines(lngLineNum, 1)
                    strTmp01 = "'" & Trim(Mid(strLine, 2))
                    .ReplaceLine lngLineNum, strTmp01
                    lngRems = lngRems + 1&
                  End If
                End If
              End If  ' ** Not a Remark.
            End If  ' ** Not a blank line.
          Next  ' ** For each Line: lngY

'If lngX > 30 Then Exit For

        Next  ' ** For each Control: lngX.
      End With  ' ** This CodeModule: cod.
    End With  ' ** This VBComponent: vbc.
  End With  ' ** This VBProject: vbp.

  Debug.Print "'FIXES: " & CStr(lngFixes) & IIf(lngRems = 0&, "", " REMS: " & CStr(lngRems))

Else
  blnRetValx = False
End If

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set vbp = Nothing
  Set vbc = Nothing
  Set cod = Nothing

  VBA_Chk_Ctls = blnRetValx

End Function

Private Function VBA_IsSQL_Ctl(cod As CodeModule, lngLineNum As Long, intPosCtl As Integer) As Boolean
' ** Called by:
' **   VBA_Chk_Ctls(), Above

  Const THIS_PROC As String = "VBA_IsSQL_Ctl"

  Dim strLine As String
  Dim lngQuotes As Long, arr_varQuote() As Variant
  Dim lngSingles As Long, arr_varSingle() As Variant
  Dim lngCnt As Long
  Dim intPos01 As Integer
  Dim lngX As Long, lngElemQ As Long

  ' ** Array: arr_varQuote().
  Const Q_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const Q_OP As Integer = 0
  Const Q_CL As Integer = 1

  Const S_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  'Const S_OP As Integer = 0
  'Const S_CL As Integer = 1

  blnRetValx = False

  lngQuotes = 0&
  lngCnt = 0&
  ReDim arr_varQuote(Q_ELEMS, 0)

  lngSingles = 0&
  ReDim arr_varSingle(S_ELEMS, 0)

  With cod
    strLine = Trim(.Lines(lngLineNum, 1))
    If strLine <> vbNullString And intPosCtl > 0 Then
      intPos01 = InStr(strLine, Chr(34))  ' ** "
      If intPos01 > 0 Then
        ' ** String lines can't be broken within opening and closing quotes.

        ' ** Find the quotes: 1 open, 2 close, 3 open, 4 close, etc.
        Do While intPos01 > 0
          If lngCnt Mod 2& = 0& Then
            lngCnt = lngCnt + 1&
            lngQuotes = lngQuotes + 1&
            lngElemQ = lngQuotes - 1&
            ReDim Preserve arr_varQuote(Q_ELEMS, lngElemQ)
            arr_varQuote(Q_OP, lngElemQ) = intPos01
            arr_varQuote(Q_CL, lngElemQ) = CInt(0)
          Else
            lngCnt = lngCnt + 1&
            lngElemQ = lngQuotes - 1&
            arr_varQuote(Q_CL, lngElemQ) = intPos01
          End If
          intPos01 = InStr((intPos01 + 1), strLine, Chr(34))
        Loop

        ' ** Check for Remarks.
        intPos01 = InStr(strLine, "'")
        If intPos01 > 0 Then
          Do While intPos01 > 0
            If lngSingles Mod 2& = 0& Then
              lngSingles = lngSingles + 1&
              lngElemQ = lngSingles - 1&
              ReDim Preserve arr_varSingle(Q_ELEMS, lngElemQ)
              arr_varSingle(Q_OP, lngElemQ) = intPos01
              arr_varSingle(Q_CL, lngElemQ) = CInt(0)
            Else
              lngElemQ = lngSingles - 1&
              arr_varSingle(Q_CL, lngElemQ) = intPos01
            End If
            intPos01 = InStr((intPos01 + 1), strLine, "'")
          Loop
        End If

        ' ** OK, where is the control relative to the various quotes.
        For lngX = 0& To (lngQuotes - 1&)
          lngElemQ = lngX
          If arr_varQuote(Q_OP, lngElemQ) < intPosCtl And arr_varQuote(Q_CL, lngElemQ) > intPosCtl Then
            ' ** It's within quotes.
            ' ** Under what circumstances would singles interfere with this assessment?
            ' **   1. Single directly after open and directly before close
            ' **      should mean it's referring to previous and next lines.
            ' **   2. I'm wracking my brain and can't come up with an example of
            ' **      a single-quote (single apostrophe) appearing within a SQL
            ' **      statement that wouldn't blow it up. It's treated as an
            ' **      interchangeable double-quote in the Access SQL. You can put
            ' **      singles inside doubles or doubles inside singles, but the
            ' **      outside ones must be paired. For example:
            ' **      a. "' & account.accountno, "
            ' **         As a line that might be found in VBA, the ampersand is, in
            ' **         this case, treated as a SQL concatenator and not a VBA one.
            ' **         The single-quote is unrelated to the field designation.
            ' **      b. "' & account.accountno, '"
            ' **         In this case, everything within the singles is considered
            ' **         a string, and so is not interpreted as SQL. It could be,
            ' **         for example, part of a text field assignment, like an
            ' **         ad hoc query table.
            ' **         BUT WAIT!
            ' **      c. "'" & "ABCD" & "' AS accountno, '" & "xyz" & " AS bleep,"
            ' **         Here, the two singles have nothing to do with the field.
            ' **      It's hard to think of all possibilities. Perhaps a simple,
            ' **      viable rule has to only cover the likeliest scenarios.
            ' ** So...
            ' ** A single one (odd): Unrelated. Field is within SQL-interpreted text.
            ' **   Therefore, don't ever put a form reference ("Me.") before it.
            ' ** Two singles (even): Maybe, Maybe Not. As long as it's within a pair
            ' **   of doubles, don't use a form reference.
            ' ** ...
            ' ** In effect, I'm saying: IGNORE SINGLES! ONLY COUNT DOUBLES!
            blnRetValx = True
            Exit For
          End If
        Next

      End If
    End If
  End With

  VBA_IsSQL_Ctl = blnRetValx

End Function

Private Function VBA_Block_Chk(ByRef strLine As String, ByRef strTerm As String, ByRef lngTermLen As Long, ByRef blnIsTerm As Boolean, ByRef strTermName As String, ByRef lngLineLen As Long, ByRef blnNumbered As Boolean, lngX As Long, lngY As Long) As Boolean
' ** Determines whether statement contains a block term.
' ** This function can change the passed variables.
' ** Called by:
' **   VBA_Module_Format(), Above

  Const THIS_PROC As String = "VBA_Block_Chk"

  Dim intLen As Integer

  blnRetValx = True

  intLen = Len(strLine)

  If Left(strLine, lngTermLen) = strTerm Then
    ' ** So far, it looks like a structure statement.
    If lngTermLen = intLen Then
      ' ** Definitely a structure statement.
      blnIsTerm = True
      strTermName = strTerm
    Else
      ' ** See if the next character is a space.
      If Mid(strLine, (lngTermLen + 1), 1) = " " Then
        ' ** OK, looks like this is a structure statement.
        blnIsTerm = True
        strTermName = strTerm
      End If
    End If
  End If

  VBA_Block_Chk = blnRetValx

End Function

Private Sub VBA_Block_Term_Load()
' ** Terms that initiate an block indent.
' ** Called by:
' **   VBA_Module_Format(), Above

  Const THIS_PROC As String = "VBA_Block_Term_Load"

  Dim lngX As Long, lngE As Long

  ' *********************************************************
  ' ** Array: arr_varBlock()
  ' **
  ' **   Element  Description                Constant
  ' **   =======  =========================  ==============
  ' **      0     Opening Term               B_OPEN
  ' **      1     Opening Term Length        B_OPEN_LEN
  ' **      2     Opening Term Words         B_OPEN_CNT
  ' **      3     Mid-Block Term 1           B_ALIGN1
  ' **      4     Mid-Block Term 1 Length    B_ALIGN1_LEN
  ' **      5     Mid-Block Term 1 Words     B_ALIGN1_CNT
  ' **      6     Mid-Block Term 2           B_ALIGN2
  ' **      7     Mid-Block Term 2 Length    B_ALIGN2_LEN
  ' **      8     Mid-Block Term 2 Words     B_ALIGN2_CNT
  ' **      9     Closing Term               B_CLOSE
  ' **     10     Closing Term Length        B_CLOSE_LEN
  ' **     11     Closing Term Words         B_CLOSE_CNT
  ' **     12     Scope Possible YN          B_SCOPE
  ' **
  ' *********************************************************

  ' ** Select Case.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "Select Case"
  arr_varBlock(B_ALIGN1, lngE) = "Case"
  arr_varBlock(B_ALIGN2, lngE) = vbNullString
  arr_varBlock(B_CLOSE, lngE) = "End Select"
  arr_varBlock(B_SCOPE, lngE) = False

  ' ** If-Then-Else.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "If"
  arr_varBlock(B_ALIGN1, lngE) = "ElseIf"
  arr_varBlock(B_ALIGN2, lngE) = "Else"
  arr_varBlock(B_CLOSE, lngE) = "End If"
  arr_varBlock(B_SCOPE, lngE) = False

  ' ** For-Next.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "For"
  arr_varBlock(B_ALIGN1, lngE) = vbNullString
  arr_varBlock(B_ALIGN2, lngE) = vbNullString
  arr_varBlock(B_CLOSE, lngE) = "Next"
  arr_varBlock(B_SCOPE, lngE) = False

  ' ** Do-Loop.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "Do"
  arr_varBlock(B_ALIGN1, lngE) = vbNullString
  arr_varBlock(B_ALIGN2, lngE) = vbNullString
  arr_varBlock(B_CLOSE, lngE) = "Loop"
  arr_varBlock(B_SCOPE, lngE) = False

  ' ** While-Wend.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "While"
  arr_varBlock(B_ALIGN1, lngE) = vbNullString
  arr_varBlock(B_ALIGN2, lngE) = vbNullString
  arr_varBlock(B_CLOSE, lngE) = "Wend"
  arr_varBlock(B_SCOPE, lngE) = False

  ' ** With.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "With"
  arr_varBlock(B_ALIGN1, lngE) = vbNullString
  arr_varBlock(B_ALIGN2, lngE) = vbNullString
  arr_varBlock(B_CLOSE, lngE) = "End With"
  arr_varBlock(B_SCOPE, lngE) = False

  ' ** Type.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "Type"
  arr_varBlock(B_ALIGN1, lngE) = vbNullString
  arr_varBlock(B_ALIGN2, lngE) = vbNullString
  arr_varBlock(B_CLOSE, lngE) = "End Type"
  arr_varBlock(B_SCOPE, lngE) = True

  ' ** Enum.
  lngBlocks = lngBlocks + 1&
  lngE = lngBlocks - 1&
  ReDim Preserve arr_varBlock(B_ELEMS, lngE)
  arr_varBlock(B_OPEN, lngE) = "Enum"
  arr_varBlock(B_ALIGN1, lngE) = vbNullString
  arr_varBlock(B_ALIGN2, lngE) = vbNullString
  arr_varBlock(B_CLOSE, lngE) = "End Enum"
  arr_varBlock(B_SCOPE, lngE) = True

  ' ** Add lengths and word counts.
  For lngX = 0& To (lngBlocks - 1&)
    arr_varBlock(B_OPEN_LEN, lngX) = Len(arr_varBlock(B_OPEN, lngX))
    If InStr(arr_varBlock(B_OPEN, lngX), " ") = 0 Then
      arr_varBlock(B_OPEN_CNT, lngX) = 1
    Else
      arr_varBlock(B_OPEN_CNT, lngX) = 2
    End If
    If arr_varBlock(B_ALIGN1, lngX) <> vbNullString Then
      arr_varBlock(B_ALIGN1_LEN, lngX) = Len(arr_varBlock(B_ALIGN1, lngX))
      If InStr(arr_varBlock(B_ALIGN1, lngX), " ") = 0 Then
        arr_varBlock(B_ALIGN1_CNT, lngX) = 1
      Else
        arr_varBlock(B_ALIGN1_CNT, lngX) = 2
      End If
    Else
      arr_varBlock(B_ALIGN1_LEN, lngX) = 0
      arr_varBlock(B_ALIGN1_CNT, lngX) = 0
    End If
    If arr_varBlock(B_ALIGN2, lngX) <> vbNullString Then
      arr_varBlock(B_ALIGN2_LEN, lngX) = Len(arr_varBlock(B_ALIGN2, lngX))
      If InStr(arr_varBlock(B_ALIGN2, lngX), " ") = 0 Then
        arr_varBlock(B_ALIGN2_CNT, lngX) = 1
      Else
        arr_varBlock(B_ALIGN2_CNT, lngX) = 2
      End If
    Else
      arr_varBlock(B_ALIGN2_LEN, lngX) = 0
      arr_varBlock(B_ALIGN2_CNT, lngX) = 0
    End If
    arr_varBlock(B_CLOSE_LEN, lngX) = Len(arr_varBlock(B_CLOSE, lngX))
    If InStr(arr_varBlock(B_CLOSE, lngX), " ") = 0 Then
      arr_varBlock(B_CLOSE_CNT, lngX) = 1
    Else
      arr_varBlock(B_CLOSE_CNT, lngX) = 2
    End If
  Next

End Sub

Private Function IsProperty(lngLineNum As Long, lngModLines As Long, lngLastProcEnd As Long, codx As CodeModule) As String
' ** Called by:
' **   VBA_Module_Format(), Above

  Const THIS_PROC As String = "IsProperty"

  Dim strLine As String, strProcName As String
  Dim strScopeWord As String, strProcWord As String, strPropWord As String
  Dim intPos01 As Integer
  Dim strTmp01 As String
  Dim lngX As Long
  Dim strRetVal As String

  strRetVal = vbNullString

  If lngLastProcEnd < lngLineNum Then
    With codx
      strProcName = .ProcOfLine(lngLineNum, vbext_pk_Proc)
      strScopeWord = vbNullString: strProcWord = vbNullString: strPropWord = vbNullString
      For lngX = lngLineNum To lngModLines
        ' ** Look for the declaration line.
        strLine = Trim(.Lines(lngX, 1))
        If strLine <> vbNullString Then
          If Left(strLine, 1) <> "'" Then
            intPos01 = InStr(strLine, " ")
            If intPos01 > 0 Then
              strTmp01 = Trim(Left(strLine, intPos01))  ' ** 1st word.
              Select Case strTmp01
              Case "Public", "Private", "Friend", "Static", "Global"
                ' ** Looks like a declaration line.
                strScopeWord = strTmp01
              End Select
              If strScopeWord = vbNullString Then
                Select Case strTmp01
                Case "Sub", "Function", "Property"
                  ' ** Looks like a declaration line.
                  strProcWord = strTmp01
                End Select
              End If
              If strScopeWord <> vbNullString Or strProcWord <> vbNullString Then
                ' ** Might be the declaration line.
                If strProcWord = vbNullString Then
                  strTmp01 = Trim(Mid(strLine, intPos01))
                  intPos01 = InStr(strTmp01, " ")
                  If intPos01 > 0 Then
                    strTmp01 = Trim(Left(strTmp01, intPos01))  ' ** 2nd word.
                    Select Case strTmp01
                    Case "Sub", "Function", "Property"
                      ' ** Looks like a declaration line.
                      strProcWord = strTmp01
                    End Select
                  End If
                End If
                If strProcWord = "Property" Then
                  intPos01 = InStr(strLine, (strProcWord & " "))
                  If intPos01 > 0 Then
                    strTmp01 = Trim(Mid(strLine, (intPos01 + 8)))
                    intPos01 = InStr(strTmp01, " ")
                    If intPos01 > 0 Then
                      strTmp01 = Trim(Left(strTmp01, intPos01))
                      Select Case strTmp01
                      Case "Let", "Set", "Get"
                        strRetVal = strTmp01
                      End Select
                      If strRetVal <> vbNullString Then Exit For
                    End If
                  End If
                End If
              Else
                ' ** Try the next line.
              End If
            End If
          End If
        End If
        If Left(strLine, 4) = "End " Then Exit For
        If .ProcOfLine(lngLineNum, vbext_pk_Proc) <> strProcName Then Exit For
      Next
    End With
  End If

  IsProperty = strRetVal

End Function

Public Function VBA_THIS_NAME(Optional varModName As Variant) As Boolean
' ** Not called.
'### LOTS OF UNIQUE VARIABLES!

  Const THIS_PROC As String = "VBA_This_Name"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngMods As Long, arr_varMod() As Variant
  Dim lngLines As Long, lngDecLines As Long, lngModsChecked As Long, lngLineNum As Long
  Dim strLine As String, strModName As String
  Dim blnOneOnly As Boolean, blnSkip As Boolean
  Dim blnHasName As Boolean, blnHasComp As Boolean, blnHasExp As Boolean
  Dim intPos01 As Integer, intPos02 As Integer
  Dim strTmp01 As String
  Dim lngX As Long, lngE As Long, lngElemM As Long

  ' ** Array: arr_varMod().
  Const M_ELEMS As Integer = 11  ' ** Array's first-element UBound().
  Const M_NAME    As Integer = 0
  Const M_TYPE    As Integer = 1
  Const M_HASNAME As Integer = 2
  Const M_N_LNUM  As Integer = 3
  Const M_N_LINE  As Integer = 4
  Const M_HASCOMP As Integer = 5
  Const M_C_LNUM  As Integer = 6
  Const M_C_LINE  As Integer = 7
  Const M_HASEXP  As Integer = 8
  Const M_E_LNUM  As Integer = 9
  Const M_E_LINE  As Integer = 10
  Const M_ERR     As Integer = 11

  Const LINE_COMP As String = "Option Compare Database"
  Const LINE_EXP  As String = "Option Explicit"

  blnRetValx = False

  If IsMissing(varModName) = False Then
    blnOneOnly = True
  Else
    blnOneOnly = False
  End If

  lngMods = 0&
  ReDim arr_varMod(M_ELEMS, 0)
  ' ***********************************************************
  ' ** Array: arr_varMod()
  ' **
  ' **   Element  Description                     Constant
  ' **   =======  ==============================  ===========
  ' **      0     Module Name                     M_NAME
  ' **      1     Module Type                     M_TYPE
  ' **      2     Has THIS_NAME (True/False)      M_HASNAME
  ' **      3     THIS_NAME Line Number           M_N_LNUM
  ' **      4     Line Text                       M_N_LINE
  ' **      5     Has Compare (True/False)        M_HASCOMP
  ' **      6     Option Compare Line Number      M_C_LNUM
  ' **      7     Line Text                       M_C_LINE
  ' **      8     Has Explicit (True/False)       M_HASEXP
  ' **      9     Option Explicit Line Number     M_E_LNUM
  ' **     10     Line Text                       M_E_LINE
  ' **     11     THIS_NAME error (True/False)    M_ERR
  ' **
  ' ***********************************************************

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngModsChecked = 0&
    For Each vbc In .VBComponents
      With vbc

        lngModsChecked = lngModsChecked + 1&
        strModName = .Name

        If blnOneOnly = True Then
          If strModName = varModName Then blnSkip = False Else blnSkip = True
        Else
          blnSkip = False
        End If

        If blnSkip = False Then

          lngMods = lngMods + 1&
          lngE = lngMods - 1&
          ReDim Preserve arr_varMod(M_ELEMS, lngE)
          arr_varMod(M_NAME, lngE) = .Name
          arr_varMod(M_TYPE, lngE) = .Type
          ' **   vbext_ComponentType enumeration:
          ' **       1  vbext_ct_StdModule        Standard Module
          ' **       2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
          ' **       3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm in the VBA Editor.
          ' **      11  vbext_ct_ActiveXDesigner
          ' **     100  vbext_ct_Document         Module behind Form, Report, or Excel Worksheet.
          arr_varMod(M_HASNAME, lngE) = CBool(False)
          arr_varMod(M_N_LNUM, lngE) = CLng(0)
          arr_varMod(M_N_LINE, lngE) = vbNullString
          arr_varMod(M_HASCOMP, lngE) = CBool(False)
          arr_varMod(M_C_LNUM, lngE) = CLng(0)
          arr_varMod(M_C_LINE, lngE) = vbNullString
          arr_varMod(M_HASEXP, lngE) = CBool(False)
          arr_varMod(M_E_LNUM, lngE) = CLng(0)
          arr_varMod(M_E_LINE, lngE) = vbNullString
          arr_varMod(M_ERR, lngE) = CBool(False)
          strModName = .Name
          lngElemM = lngE

          Set cod = .CodeModule
          With cod

            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            blnHasName = False: blnHasComp = False: blnHasExp = False

            ' ** Look for THIS_NAME.
            For lngX = 1& To lngDecLines
              lngLineNum = lngX
              strLine = Trim(.Lines(lngX, 1))
              If InStr(strLine, "Const THIS_NAME As String = ") > 0 Then
                blnHasName = True
                arr_varMod(M_HASNAME, lngElemM) = CBool(True)
                arr_varMod(M_N_LNUM, lngElemM) = lngX
                arr_varMod(M_N_LINE, lngElemM) = strLine
                Exit For
              End If
            Next  ' ** For each line: lngX, lngLineNum.

            If blnHasName = True Then
              ' ** Check to make sure it's correct.
              strLine = Trim(.Lines(arr_varMod(M_N_LNUM, lngElemM), 1))
              intPos01 = InStr(strLine, Chr(34))
              If intPos01 > 0 Then
                intPos02 = InStr((intPos01 + 1), strLine, Chr(34))
                If intPos02 > 0 Then
                  strTmp01 = Mid(strLine, (intPos01 + 1), ((intPos02 - intPos01) - 1))
                  Select Case arr_varMod(M_TYPE, lngElemM)
                  Case vbext_ct_StdModule
                    If strTmp01 <> strModName Then
                      Debug.Print "'WRONG THIS_NAME: " & strModName
                      arr_varMod(M_ERR, lngElemM) = CBool(True)
                    End If
                  Case vbext_ct_ClassModule
                    If strTmp01 <> strModName Then
                      Debug.Print "'WRONG THIS_NAME: " & strModName
                      arr_varMod(M_ERR, lngElemM) = CBool(True)
                    End If
                  Case vbext_ct_Document
                    If Left(strModName, 5) = "Form_" Then
                      If Mid(strModName, 6) <> strTmp01 Then
                        Debug.Print "'WRONG THIS_NAME: " & strModName & "  '" & strTmp01 & "'"
                        arr_varMod(M_ERR, lngElemM) = CBool(True)
                      End If
                    ElseIf Left(strModName, 7) = "Report_" Then
                      If Mid(strModName, 8) <> strTmp01 Then
                        Debug.Print "'WRONG THIS_NAME: " & strModName & "  '" & strTmp01 & "'"
                        arr_varMod(M_ERR, lngElemM) = CBool(True)
                      End If
                    Else
                      Debug.Print "'WHAT? " & strModName & "  '" & strTmp01 & "'"
                      arr_varMod(M_ERR, lngElemM) = CBool(True)
                    End If
                  End Select
                Else
                  Debug.Print "'BAD THIS_NAME! " & strModName & "  '" & strTmp01 & "'"
                  arr_varMod(M_ERR, lngElemM) = CBool(True)
                End If
              Else
                Debug.Print "'BAD THIS_NAME! " & strModName & "  '" & strTmp01 & "'"
                arr_varMod(M_ERR, lngElemM) = CBool(True)
              End If
            End If

            ' ** Look for Option Compare Database.
            For lngX = 1& To lngDecLines
              lngLineNum = lngX
              strLine = Trim(.Lines(lngX, 1))
              If InStr(strLine, LINE_COMP) > 0 Then
                blnHasComp = True
                arr_varMod(M_HASCOMP, lngElemM) = CBool(True)
                arr_varMod(M_C_LNUM, lngElemM) = lngX
                arr_varMod(M_C_LINE, lngElemM) = strLine
                Exit For
              End If
            Next  ' ** For each line: lngX, lngLineNum.
            If blnHasComp = False Then
              Debug.Print "'NO OPTION COMPARE! " & strModName
            End If

            ' ** Look for Option Explicit.
            For lngX = 1& To lngDecLines
              lngLineNum = lngX
              strLine = Trim(.Lines(lngX, 1))
              If InStr(strLine, LINE_EXP) > 0 Then
                blnHasExp = True
                arr_varMod(M_HASEXP, lngElemM) = CBool(True)
                arr_varMod(M_E_LNUM, lngElemM) = lngX
                arr_varMod(M_E_LINE, lngElemM) = strLine
                Exit For
              End If
            Next  ' ** For each line: lngX, lngLineNum.
            If blnHasComp = False Then
              Debug.Print "'NO OPTION EXPLICIT! " & strModName
            End If

            If blnHasName = False Then

              strTmp01 = vbNullString
              Select Case arr_varMod(M_TYPE, lngElemM)
              Case vbext_ct_StdModule
                strTmp01 = strModName
              Case vbext_ct_ClassModule
                strTmp01 = strModName
              Case vbext_ct_Document
                If Left(strModName, 5) = "Form_" Then
                  strTmp01 = Mid(strModName, 6)
                ElseIf Left(strModName, 7) = "Report_" Then
                  strTmp01 = Mid(strModName, 8)
                End If
              End Select
              If strTmp01 = vbNullString Then
                Debug.Print "'PROBLEM 1: " & strModName
              Else

                strTmp01 = "Private Const THIS_NAME As String = " & Chr(34) & strTmp01 & Chr(34)

                If blnHasComp = True And blnHasExp = True Then
                  If arr_varMod(M_C_LNUM, lngElemM) = 1& Then
                    ' ** 1st line is 'Option Compare Database'.
                    If arr_varMod(M_E_LNUM, lngElemM) = 2& Then
                      ' ** 2nd line is 'Option Explicit'.
                      If Trim(.Lines(3, 1)) = vbNullString Then
                        ' ** 3rd line is blank.
                        If Trim(.Lines(3, 1)) <> .Lines(3&, 1) Then
'                          .ReplaceLine 3&, ""
                        End If
'                        .InsertLines 4&, ""
'                        .InsertLines 4&, strTmp01
                      Else
                        ' ** 3rd line not blank.
'                        .InsertLines 3&, ""
'                        .InsertLines 3&, strTmp01
'                        .InsertLines 3&, ""
                      End If
                    Else
                      ' ** 2nd line not 'Option Explicit'
                      Debug.Print "'DO MANUALLY: " & strModName
                    End If
                  Else
                    ' ** 1st line not 'Option Compare Database'.
'                    .ReplaceLine arr_varMod(M_C_LNUM, lngElemM), ""
'                    .ReplaceLine arr_varMod(M_E_LNUM, lngElemM), ""
'                    .InsertLines 1&, LINE_COMP
'                    .InsertLines 1&, LINE_EXP
'                    .InsertLines 1&, ""
'                    .InsertLines 1&, strTmp01
'                    .InsertLines 1&, ""
                  End If
                Else
                  ' ** At least 1 'Option' is missing.
                  If blnHasComp = True Then
                    If arr_varMod(M_C_LNUM, lngElemM) = 1& Then
                      ' ** 1st line is 'Option Compare Database'.
                      If Trim(.Lines(2&, 1&)) = vbNullString Then
                        ' ** 2nd line is blank.
                        If Trim(.Lines(2&, 1&)) <> .Lines(2&, 1&) Then
'                          .ReplaceLine 2&, ""
                        End If
'                        .InsertLines 2&, strTmp01
'                        .InsertLines 2&, ""
'                        .InsertLines 2&, LINE_EXP
                      Else
                        ' ** 2nd line not blank.
'                        .InsertLines 2&, ""
'                        .InsertLines 2&, strTmp01
'                        .InsertLines 2&, ""
'                        .InsertLines 2&, LINE_EXP
                      End If
                    Else
                      ' ** 1st line not 'Option Compare Database'.
'                      .ReplaceLine arr_varMod(M_C_LNUM, lngElemM), ""
'                      .InsertLines 1&, LINE_COMP
'                      .InsertLines 1&, LINE_EXP
'                      .InsertLines 1&, ""
'                      .InsertLines 1&, strTmp01
'                      .InsertLines 1&, ""
                    End If
                  End If
                  If blnHasExp = True Then
                    If arr_varMod(M_E_LNUM, lngElemM) = 1& Then
                      ' ** 1st line is 'Option Explicit'.
                      If Trim(.Lines(2&, 1&)) = vbNullString Then
                        ' ** 2nd line is blank.
                        If Trim(.Lines(2&, 1&)) <> .Lines(2&, 1&) Then
'                          .ReplaceLine 2&, ""
                        End If
'                        .InsertLines 2&, strTmp01
'                        .InsertLines 2&, ""
'                        .InsertLines 1&, LINE_COMP
                      Else
                        ' ** 2nd line not blank.
'                        .InsertLines 2&, ""
'                        .InsertLines 2&, strTmp01
'                        .InsertLines 2&, ""
'                        .InsertLines 1&, LINE_COMP
                      End If
                    Else
                      ' ** 1st line not 'Option Explicit'.
'                      .ReplaceLine arr_varMod(M_E_LNUM, lngElemM), ""
'                      .InsertLines 1&, LINE_COMP
'                      .InsertLines 1&, LINE_EXP
'                      .InsertLines 1&, ""
'                      .InsertLines 1&, strTmp01
'                      .InsertLines 1&, ""
                    End If
                  End If
                End If  ' ** blnHasComp, blnHasExp.

              End If  ' ** Valid strTmp01.

            End If  ' ** blnHasName = False.

            If blnHasName = True And (blnHasComp = False Or blnHasExp = False) Then
              Debug.Print "'MISSING OPTION: " & strModName
            End If

            If blnHasName = False Then
              Debug.Print "'MISSING HAS_NAME: " & strModName
            End If

          End With  ' ** This Code Module: cod.

        End If  ' ** blnSkip.

      End With  ' ** This Component: vbc.
    Next  ' ** For each Component: vbc.
  End With  ' ** This Project.

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_THIS_NAME = blnRetValx

End Function

Public Function VBA_Component_Properties() As Boolean
' ** Not called.

  Const THIS_PROC As String = "VBA_Component_Properties"

  Dim vbp As VBProject, vbc As VBComponent, vbc2 As VBComponent, prp As Object
  Dim dbs As DAO.Database, rst As DAO.Recordset, frm As Access.Form, rpt As Access.Report, ctl As Access.Control, fld As DAO.Field
  Dim lngMods As Long, arr_varMod() As Variant
  Dim lngProps As Long, arr_varProp() As Variant
  Dim intType As Integer, intVarType As Integer, strName As String, strName2 As String
  Dim lngRecs As Long, lngAdded As Long
  Dim blnFound As Boolean, blnAdd As Boolean, blnEdit As Boolean
  Dim intPos01 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, arr_varTmp04 As Variant
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  ' ** Array: arr_varMod().
  Const M_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const M_VNAM    As Integer = 0
  Const M_TYP     As Integer = 1
  Const M_VNAM2   As Integer = 2
  Const C_OBJ_TYP As Integer = 3
  Const C_PRP_CNT As Integer = 4

  ' ** Array: arr_varProp().
  Const P_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const P_ID      As Integer = 0
  Const P_NAME    As Integer = 1
  Const P_TYPE    As Integer = 2
  Const P_FRM     As Integer = 3
  Const P_RPT     As Integer = 4
  Const P_COM     As Integer = 5
  Const P_OBJS    As Integer = 6  ' ** String of objecttype_type's.
  Const P_TYPES   As Integer = 7  ' ** String of datatype_vb_type's.
  Const P_FOUND   As Integer = 8
  Const P_OBJECTS As Integer = 9  ' ** Form/Report name.

  blnRetValx = True

  lngMods = 0&
  ReDim arr_varMod(M_ELEMS, 0)

  lngAdded = 0&

  Set vbp = Application.VBE.ActiveVBProject
  With vbp

    ' ** First, get basic property info on each VBComponent.
    For Each vbc In .VBComponents
      With vbc
        strName = .Name
        intType = .Type
        ' ***********************************************
        ' ** Array: arr_varMod()
        ' **
        ' **   Element  Name              Constant
        ' **   =======  ================  =============
        ' **      0     Component Name    M_VNAM
        ' **      1     Type              M_TYP
        ' **      2     Object Name       M_VNAM2
        ' **      3     Object Type       C_OBJ_TYP
        ' **      4     Property Count    C_PRP_CNT
        ' **
        ' ***********************************************
        lngMods = lngMods + 1&
        lngE = lngMods - 1&
        ReDim Preserve arr_varMod(M_ELEMS, lngE)
        arr_varMod(M_VNAM, lngE) = strName
        arr_varMod(M_TYP, lngE) = intType
        If intType = vbext_ct_Document Then
          intPos01 = InStr(strName, "_")
          strName2 = Mid(strName, (intPos01 + 1))
          arr_varMod(M_VNAM2, lngE) = strName2
          If Left(strName, 5) = "Form_" Then
            arr_varMod(C_OBJ_TYP, lngE) = acForm
          ElseIf Left(strName, 7) = "Report_" Then
            arr_varMod(C_OBJ_TYP, lngE) = acReport
          End If
        Else
          arr_varMod(M_VNAM2, lngE) = Null
          arr_varMod(C_OBJ_TYP, lngE) = acModule
          arr_varMod(C_PRP_CNT, lngE) = .Properties.Count
        End If
      End With  ' ** vbc.
    Next  ' ** For each vbc.

    ' ** Now, get the count of properties in the corresponding form/report object.
    For lngX = 0& To (lngMods - 1&)
      If arr_varMod(C_OBJ_TYP, lngX) = acForm Then
        strName = arr_varMod(M_VNAM, lngX)
        strName2 = arr_varMod(M_VNAM2, lngX)
        DoCmd.OpenForm strName2, acDesign, , , , acHidden
        DoEvents
        Set vbc2 = Application.VBE.VBProjects("Trust").VBComponents(strName)
        arr_varMod(C_PRP_CNT, lngE) = vbc2.CodeModule.Parent.Properties.Count
        DoCmd.Close acForm, strName2, acSaveNo
        Set vbc2 = Nothing
      ElseIf arr_varMod(C_OBJ_TYP, lngX) = acReport Then
        strName = arr_varMod(M_VNAM, lngX)
        strName2 = arr_varMod(M_VNAM2, lngX)
        DoCmd.OpenReport strName2, acViewDesign
        DoEvents
        Set vbc2 = Application.VBE.VBProjects("Trust").VBComponents(strName)
        arr_varMod(C_PRP_CNT, lngE) = vbc2.Properties.Count
        DoCmd.Close acReport, strName2, acSaveNo
        Set vbc2 = Nothing
      End If
    Next

    lngProps = 0&
    ReDim arr_varProp(P_ELEMS, 0)

    Set dbs = CurrentDb
    With dbs

      ' ** Next get a list of properties already documented.
      Set rst = .OpenRecordset("tblVBComponent_Property", dbOpenDynaset, dbReadOnly)
      With rst
        If .BOF = True And .EOF = True Then
          ' ** Table not loaded yet.
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
          ' *******************************************************
          ' ** Array: arr_varProp()
          ' **
          ' **   Field  Element  Name               Constant
          ' **   =====  =======  =================  =============
          ' **     1       0     vbcom_prop_id      P_ID
          ' **     2       1     vbcom_prop_name    P_NAME
          ' **     3       2     datatype_vb_type      P_TYPE
          ' **     4       3     vbcom_frm          P_FRM
          ' **     5       4     vbcom_rpt          P_RPT
          ' **     6       5     vbcom_com          P_COM
          ' **     7       6     vbcom_objs         P_OBJS
          ' **     8       7     comtype_types      P_TYPES
          ' **     9       8     Found Yes/No       P_FOUND
          ' **    10       9     vbcom_objects      P_OBJECTS
          ' **
          ' *******************************************************
          For lngX = 1& To (lngRecs - 1&)
            lngProps = lngProps + 1&
            lngE = lngProps - 1&
            ReDim Preserve arr_varProp(P_ELEMS, lngE)
            arr_varProp(P_ID, lngE) = ![vbcom_prop_id]
            arr_varProp(P_NAME, lngE) = ![vbcom_prop_name]
            arr_varProp(P_TYPE, lngE) = ![datatype_vb_type]
            arr_varProp(P_FRM, lngE) = ![vbcom_frm]
            arr_varProp(P_RPT, lngE) = ![vbcom_rpt]
            arr_varProp(P_COM, lngE) = ![vbcom_com]
            arr_varProp(P_OBJS, lngE) = ![vbcom_objs]
            arr_varProp(P_TYPES, lngE) = ![comtype_types]
            arr_varProp(P_FOUND, lngE) = CBool(False)
          Next
        End If
        .Close
      End With

      .Close
    End With

    ' ** Now collect a comprehensive list of properties.
    For lngX = 0& To (lngMods - 1&)
      Set vbc = .VBComponents(arr_varMod(M_VNAM, lngX))
      With vbc

        Select Case arr_varMod(C_OBJ_TYP, lngX)
        Case acForm
          ' ** Form module.
          strName2 = arr_varMod(M_VNAM2, lngX)
          DoCmd.OpenForm strName2, acDesign, , , , acHidden
          Set frm = Forms(strName2)
          For Each prp In .Properties
            With prp
              intVarType = varType(vbc.Properties(.Name))
              blnFound = False
              ' ** First, look in the Form's Properties Collection.
              For lngY = 0& To (lngProps - 1&)
                If arr_varProp(P_NAME, lngY) = .Name Then
                  blnFound = True
                  If InStr(arr_varProp(P_OBJS, lngY), CStr(acForm)) = 0 Then
                    arr_varProp(P_OBJS, lngY) = arr_varProp(P_OBJS, lngY) & CStr(acForm) & ";"
                  End If
                  If InStr(arr_varProp(P_TYPES, lngY), CStr(intVarType)) = 0 Then
                    arr_varProp(P_TYPES, lngY) = arr_varProp(P_TYPES, lngY) & CStr(intVarType) & ";"
                  End If
                  Exit For
                End If
              Next
              If blnFound = False Then
                ' ** If not found in the Form's Properties Collection, see if it's in a Control's Properties Collection.
                For Each ctl In frm.Controls
                  With ctl
                    ' ** Check for spaces and/or periods.
                    strTmp02 = .Name
                    intPos01 = InStr(strTmp02, " ")
                    Do While intPos01 > 0
                      strTmp02 = Left(strTmp02, (intPos01 - 1)) & "_" & Mid(strTmp02, (intPos01 + 1))
                      intPos01 = InStr(strTmp02, " ")
                    Loop
                    intPos01 = InStr(strTmp02, ".")
                    Do While intPos01 > 0
                      strTmp02 = Left(strTmp02, (intPos01 - 1)) & "_" & Mid(strTmp02, (intPos01 + 1))
                      intPos01 = InStr(strTmp02, ".")
                    Loop
                    strTmp03 = prp.Name
                    intPos01 = InStr(strTmp03, " ")
                    Do While intPos01 > 0
                      strTmp03 = Left(strTmp03, (intPos01 - 1)) & "_" & Mid(strTmp03, (intPos01 + 1))
                      intPos01 = InStr(strTmp03, " ")
                    Loop
                    intPos01 = InStr(strTmp03, ".")
                    Do While intPos01 > 0
                      strTmp03 = Left(strTmp03, (intPos01 - 1)) & "_" & Mid(strTmp03, (intPos01 + 1))
                      intPos01 = InStr(strTmp03, ".")
                    Loop
                    If strTmp02 = strTmp03 Then
                      blnFound = True
                      Exit For
                    End If
                  End With
                Next
                If blnFound = False Then
                  ' ** If not found in a Control's Properties Collection, see if it's in RecordSource's Properties Collection.
                  If intVarType = vbObject Then
                    If frm.RecordSource <> vbNullString Then
                      Set dbs = CurrentDb
                      With dbs
On Error Resume Next
                        Set rst = .OpenRecordset(frm.RecordSource)
                        If ERR.Number = 0 Then
On Error GoTo 0
                          With rst
                            For Each fld In .Fields
                              With fld
                                ' ** Check for spaces and/or periods.
                                strTmp02 = .Name
                                intPos01 = InStr(strTmp02, " ")
                                Do While intPos01 > 0
                                  strTmp02 = Left(strTmp02, (intPos01 - 1)) & "_" & Mid(strTmp02, (intPos01 + 1))
                                  intPos01 = InStr(strTmp02, " ")
                                Loop
                                intPos01 = InStr(strTmp02, ".")
                                Do While intPos01 > 0
                                  strTmp02 = Left(strTmp02, (intPos01 - 1)) & "_" & Mid(strTmp02, (intPos01 + 1))
                                  intPos01 = InStr(strTmp02, ".")
                                Loop
                                strTmp03 = prp.Name
                                intPos01 = InStr(strTmp03, " ")
                                Do While intPos01 > 0
                                  strTmp03 = Left(strTmp03, (intPos01 - 1)) & "_" & Mid(strTmp03, (intPos01 + 1))
                                  intPos01 = InStr(strTmp03, " ")
                                Loop
                                intPos01 = InStr(strTmp03, ".")
                                Do While intPos01 > 0
                                  strTmp03 = Left(strTmp03, (intPos01 - 1)) & "_" & Mid(strTmp03, (intPos01 + 1))
                                  intPos01 = InStr(strTmp03, ".")
                                Loop
                                If strTmp02 = strTmp03 Then
                                  blnFound = True
                                  Exit For
                                End If
                              End With
                            Next
                            .Close
                          End With
                        Else
                          ' ** The RecordSource isn't valid right now (perhaps a temporary table, or something like that).
On Error GoTo 0
                        End If
                        .Close
                      End With  ' ** rst.
                    End If  ' ** Form has a RecordSource.
                  End If  ' ** The property we're checking is an object.
                End If  ' ** blnFound.
                If blnFound = False Then
                  ' ** If still not found, check it against a known list of objects.
                  If VBA_IsControl(.Name) = False Then
                    ' ** Put whatever's left into the arr_varProp() array
                    ' ** as a potential new Property for the list.
                    lngProps = lngProps + 1&
                    lngE = lngProps - 1&
                    ReDim Preserve arr_varProp(P_ELEMS, lngE)
                    arr_varProp(P_ID, lngE) = CLng(0)
                    arr_varProp(P_NAME, lngE) = .Name
                    arr_varProp(P_TYPE, lngE) = intVarType
                    arr_varProp(P_FRM, lngE) = CBool(False)
                    arr_varProp(P_RPT, lngE) = CBool(False)
                    arr_varProp(P_COM, lngE) = CBool(False)
                    arr_varProp(P_OBJS, lngE) = CStr(acForm) & ";"
                    arr_varProp(P_TYPES, lngE) = CStr(intVarType) & ";"
                    arr_varProp(P_FOUND, lngE) = CBool(False)
                  End If
                End If
              End If
            End With  ' ** prp.
          Next  ' ** For each prp.
          Set frm = Nothing
          DoCmd.Close acForm, strName2, acSaveNo

        Case acReport
          ' ** Report module.
          strName2 = arr_varMod(M_VNAM2, lngX)
          DoCmd.OpenReport strName2, acViewDesign
          Set rpt = Reports(strName2)
          For Each prp In .Properties
            With prp
              intVarType = varType(vbc.Properties(.Name))
              blnFound = False
              ' ** First, look in the Report's Properties Collection.
              For lngY = 0& To (lngProps - 1&)
                If arr_varProp(P_NAME, lngY) = .Name Then
                  blnFound = True
                  If InStr(arr_varProp(P_OBJS, lngY), CStr(acReport)) = 0 Then
                    arr_varProp(P_OBJS, lngY) = arr_varProp(P_OBJS, lngY) & CStr(acReport) & ";"
                  End If
                  If InStr(arr_varProp(P_TYPES, lngY), CStr(intVarType)) = 0 Then
                    arr_varProp(P_TYPES, lngY) = arr_varProp(P_TYPES, lngY) & CStr(intVarType) & ";"
                  End If
                  Exit For
                End If
              Next
              If blnFound = False Then
                ' ** If not found in the Report's Properties Collection, see if it's in a Control's Properties Collection.
                For Each ctl In rpt.Controls
                  With ctl
                    If .Name = prp.Name Then
                      blnFound = True
                      Exit For
                    End If
                  End With
                Next
                If blnFound = False Then
                  ' ** If not found in a Control's Properties Collection, see if it's in RecordSource's Properties Collection.
                  If intVarType = vbObject Then
                    If rpt.RecordSource <> vbNullString Then
                      Set dbs = CurrentDb
                      With dbs
On Error Resume Next
                        Set rst = .OpenRecordset(rpt.RecordSource)
                        If ERR.Number = 0 Then
On Error GoTo 0
                          With rst
                            For Each fld In .Fields
                              With fld
                                strTmp01 = .Name
                                intPos01 = InStr(strTmp01, ".")
                                Do While intPos01 > 0
                                  strTmp01 = Left(strTmp01, (intPos01 - 1)) & "_" & Mid(strTmp01, (intPos01 + 1))
                                  intPos01 = InStr(strTmp01, ".")
                                Loop
                                If strTmp01 = prp.Name Then
                                  blnFound = True
                                  Exit For
                                End If
                              End With
                            Next
                            .Close
                          End With
                        Else
                          ' ** The RecordSource isn't valid right now (perhaps a temporary table, or something like that).
On Error GoTo 0
                        End If
                        .Close
                      End With
                    End If
                  End If
                End If
                If blnFound = False Then
                  ' ** If still not found, check it against a known list of objects.
                  If VBA_IsControl(.Name) = False Then
                    ' ** Put whatever's left into the arr_varProp() array
                    ' ** as a potential new Property for the list.
                    lngProps = lngProps + 1&
                    lngE = lngProps - 1&
                    ReDim Preserve arr_varProp(P_ELEMS, lngE)
                    arr_varProp(P_ID, lngE) = CLng(0)
                    arr_varProp(P_NAME, lngE) = .Name
                    arr_varProp(P_TYPE, lngE) = intVarType
                    arr_varProp(P_FRM, lngE) = CBool(False)
                    arr_varProp(P_RPT, lngE) = CBool(False)
                    arr_varProp(P_COM, lngE) = CBool(False)
                    arr_varProp(P_OBJS, lngE) = CStr(acReport) & ";"
                    arr_varProp(P_TYPES, lngE) = CStr(intVarType) & ";"
                    arr_varProp(P_FOUND, lngE) = CBool(False)
                  End If
                End If
              End If
            End With  ' ** prp.
          Next  ' ** For each prp.
          Set rpt = Nothing
          DoCmd.Close acReport, strName2, acSaveNo

        Case acModule
          ' ** Standard module.
          For Each prp In .Properties
            With prp
              intVarType = varType(vbc.Properties(.Name))
              blnFound = False
              ' ** First, look in the Module's Properties Collection.
              For lngY = 0& To (lngProps - 1&)
                If arr_varProp(P_NAME, lngY) = .Name Then
                  blnFound = True
                  If InStr(arr_varProp(P_OBJS, lngY), CStr(acModule)) = 0 Then
                    arr_varProp(P_OBJS, lngY) = arr_varProp(P_OBJS, lngY) & CStr(acModule) & ";"
                  End If
                  If InStr(arr_varProp(P_TYPES, lngY), CStr(intVarType)) = 0 Then
                    arr_varProp(P_TYPES, lngY) = arr_varProp(P_TYPES, lngY) & CStr(intVarType) & ";"
                  End If
                  Exit For
                End If
              Next
              If blnFound = False Then
                ' ** Put whatever's left into the arr_varProp() array
                ' ** as a potential new Property for the list.
                lngProps = lngProps + 1&
                lngE = lngProps - 1&
                ReDim Preserve arr_varProp(P_ELEMS, lngE)
                arr_varProp(P_ID, lngE) = CLng(0)
                arr_varProp(P_NAME, lngE) = .Name
                arr_varProp(P_TYPE, lngE) = intVarType
                arr_varProp(P_FRM, lngE) = CBool(False)
                arr_varProp(P_RPT, lngE) = CBool(False)
                arr_varProp(P_COM, lngE) = CBool(False)
                arr_varProp(P_OBJS, lngE) = CStr(acModule) & ";"
                arr_varProp(P_TYPES, lngE) = CStr(intVarType) & ";"
                arr_varProp(P_FOUND, lngE) = CBool(False)
                arr_varProp(P_OBJECTS, lngE) = vbNullString
              End If
            End With  ' ** prp.
          Next  ' ** For each prp.

        End Select

      End With ' ** This component: vbc.
    Next  ' ** For each component in arr_varMod(): lngX.

  End With  ' ** vbp.

  ' ** Binary Sort arr_varProp() array by property name.
  For lngX = UBound(arr_varProp, 2) To 1 Step -1
    For lngY = 0 To (lngX - 1)
      If arr_varProp(P_NAME, lngY) > arr_varProp(P_NAME, (lngY + 1)) Then
        For lngZ = 0& To P_ELEMS
          varTmp00 = arr_varProp(lngZ, lngY)
          arr_varProp(lngZ, lngY) = arr_varProp(lngZ, (lngY + 1))
          arr_varProp(lngZ, (lngY + 1)) = varTmp00
        Next
      End If
    Next
  Next

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** Finally, add or update tblVBComponent_Property.
    Set rst = .OpenRecordset("tblVBComponent_Property", dbOpenDynaset, dbConsistent)
    With rst
      .MoveFirst
      For lngX = 0& To (lngProps - 1&)

        strTmp02 = vbNullString
        blnAdd = False: blnEdit = False
        .FindFirst "[vbcom_prop_name] = '" & CStr(arr_varProp(P_NAME, lngX)) & "'"
        If .NoMatch = True Then
          blnAdd = True
        Else
          ' ** Set blnEdit to True only if a property is actually updated.
        End If

        Select Case arr_varProp(P_NAME, lngX)
        Case "assetdate", "Cost", "DateEnd", "DateStart", "description_ma", "due_ma", "GainLoss", _
            "Holding Period", "ICash", "journaltype", "legalname", "PCash", "PurchaseDate", _
            "qryCourtReport-A2.taxcode", "rate_ma", "shareface", "shorttermdate"
          blnAdd = False
        Case Else
          blnFound = True
          If blnAdd = True Then
Debug.Print "'ADD THIS PROP? " & arr_varProp(P_NAME, lngX)

            lngAdded = lngAdded + 1&
'NEW: 53
'ADD THIS PROP? AfterBeginTransaction
'ADD THIS PROP? AfterCommitTransaction
'ADD THIS PROP? AfterFinalRender
'ADD THIS PROP? AfterLayout
'ADD THIS PROP? AfterRender
'ADD THIS PROP? AllowDatasheetView
'ADD THIS PROP? AllowFormView
'ADD THIS PROP? AllowPivotChartView
'ADD THIS PROP? AllowPivotTableView
'ADD THIS PROP? BatchUpdates
'ADD THIS PROP? BeforeBeginTransaction
'ADD THIS PROP? BeforeCommitTransaction
'ADD THIS PROP? BeforeQuery
'ADD THIS PROP? BeforeRender
'ADD THIS PROP? BeforeScreenTip
'ADD THIS PROP? BeginBatchEdit
'ADD THIS PROP? ChartSpace
'ADD THIS PROP? CommandBeforeExecute
'ADD THIS PROP? CommandChecked
'ADD THIS PROP? CommandEnabled
'ADD THIS PROP? CommandExecute
'ADD THIS PROP? CommitOnClose
'ADD THIS PROP? CommitOnNavigation
'ADD THIS PROP? DataChange
'ADD THIS PROP? DataSetChange
'ADD THIS PROP? FetchDefaults
'ADD THIS PROP? GroupFooter0
'ADD THIS PROP? GroupFooter1
'ADD THIS PROP? GroupFooter2
'ADD THIS PROP? GroupFooter3
'ADD THIS PROP? GroupHeader0
'ADD THIS PROP? GroupHeader1
'ADD THIS PROP? GroupHeader2
'ADD THIS PROP? GroupHeader3
'ADD THIS PROP? MouseWheel
'ADD THIS PROP? Moveable
'ADD THIS PROP? OnConnect
'ADD THIS PROP? OnDisconnect
'ADD THIS PROP? OnRecordExit
'ADD THIS PROP? OnUndo
'ADD THIS PROP? PivotTable
'ADD THIS PROP? PivotTableChange
'ADD THIS PROP? Printer
'ADD THIS PROP? Query
'ADD THIS PROP? RecordSourceQualifier
'ADD THIS PROP? RollbackTransaction
'ADD THIS PROP? SelectionChange
'ADD THIS PROP? Shape
'ADD THIS PROP? UndoBatchEdit
'ADD THIS PROP? UseDefaultPrinter
'ADD THIS PROP? ViewChange
'ADD THIS PROP? WindowLeft
'ADD THIS PROP? WindowTop
'PROPS: 543
            If blnAdd = True Then
              .AddNew
              ![vbcom_prop_name] = arr_varProp(P_NAME, lngX)
              ![datatype_vb_type] = arr_varProp(P_TYPE, lngX)
              ![vbcom_frm] = CBool(False)
              ![vbcom_rpt] = CBool(False)
              ![vbcom_com] = CBool(False)
              ![vbcom_prop_datemodified] = Now()
              .Update
            Else
              blnFound = False
            End If
          End If

          If blnFound = True Then
            If InStr(arr_varProp(P_OBJS, lngX), CStr(acForm)) > 0 Then
              If blnAdd = True Then
                .Edit
                ![vbcom_frm] = CBool(True)
                ![vbcom_prop_datemodified] = Now()
                .Update
              ElseIf ![vbcom_frm] <> True Then
                blnEdit = True
                .Edit
                ![vbcom_frm] = CBool(True)
                ![vbcom_prop_datemodified] = Now()
                .Update
              End If
              strTmp02 = "FRM"
            End If
            If InStr(arr_varProp(P_OBJS, lngX), CStr(acReport)) > 0 Then
              If blnAdd = True Then
                .Edit
                ![vbcom_rpt] = CBool(True)
                ![vbcom_prop_datemodified] = Now()
                .Update
              ElseIf ![vbcom_rpt] <> True Then
                blnEdit = True
                .Edit
                ![vbcom_rpt] = CBool(True)
                ![vbcom_prop_datemodified] = Now()
                .Update
              End If
              strTmp02 = "RPT"
            End If
            If InStr(arr_varProp(P_OBJS, lngX), CStr(acModule)) > 0 Then
              If blnAdd = True Then
                .Edit
                ![vbcom_com] = CBool(True)
                ![vbcom_prop_datemodified] = Now()
                .Update
              ElseIf ![vbcom_com] <> True Then
                blnEdit = True
                .Edit
                ![vbcom_com] = CBool(True)
                ![vbcom_prop_datemodified] = Now()
                .Update
              End If
              strTmp02 = "COM"
            End If
            If arr_varProp(P_OBJS, lngX) <> vbNullString Then
              If blnAdd = True Then
                .Edit
                ![vbcom_objs] = arr_varProp(P_OBJS, lngX)
                ![vbcom_prop_datemodified] = Now()
                .Update
              Else
                If IsNull(![vbcom_objs]) = True Then
                  blnEdit = True
                  .Edit
                  ![vbcom_objs] = arr_varProp(P_OBJS, lngX)
                  ![vbcom_prop_datemodified] = Now()
                  .Update
                Else
                  If ![vbcom_objs] <> arr_varProp(P_OBJS, lngX) Then
                    blnEdit = True
                    .Edit
                    ![vbcom_objs] = arr_varProp(P_OBJS, lngX)
                    ![vbcom_prop_datemodified] = Now()
                    .Update
                  End If
                End If
              End If
            End If
            If arr_varProp(P_TYPES, lngX) <> vbNullString Then
              If blnAdd = True Then
                .Edit
                ![comtype_types] = arr_varProp(P_TYPES, lngX)
                ![vbcom_prop_datemodified] = Now()
                .Update
              Else
                If IsNull(![comtype_types]) = True Then
                  blnEdit = True
                  .Edit
                  ![comtype_types] = arr_varProp(P_TYPES, lngX)
                  ![vbcom_prop_datemodified] = Now()
                  .Update
                Else
                  If ![comtype_types] <> arr_varProp(P_TYPES, lngX) Then
                    blnEdit = True
                    .Edit
                    ![comtype_types] = arr_varProp(P_TYPES, lngX)
                    ![vbcom_prop_datemodified] = Now()
                    .Update
                  End If
                End If
              End If
            End If
            If arr_varProp(P_OBJECTS, lngX) <> vbNullString Then
              If blnAdd = True Then
                .Edit
                ![vbcom_objects] = strTmp02 & " " & arr_varProp(P_OBJECTS, lngX)
                ![vbcom_prop_datemodified] = Now()
                .Update
              Else
                If IsNull(![vbcom_objects]) = True Then
                  blnEdit = True
                  .Edit
                  ![vbcom_objects] = strTmp02 & " " & arr_varProp(P_OBJECTS, lngX)
                  ![vbcom_prop_datemodified] = Now()
                  .Update
                Else
                  If ![vbcom_objects] <> arr_varProp(P_OBJECTS, lngX) Then
                    blnEdit = True
                    .Edit
                    ![vbcom_objects] = strTmp02 & " " & arr_varProp(P_OBJECTS, lngX)
                    ![vbcom_prop_datemodified] = Now()
                    .Update
                  End If
                End If
              End If
            End If
            arr_varTmp04 = Split(arr_varProp(P_TYPES, lngX), ",")
            If UBound(arr_varTmp04) > 0 Then
              Debug.Print "'" & arr_varProp(P_NAME, lngX) & " " & UBound(arr_varTmp04)
            End If
          End If  ' ** blnFound.

        End Select

      Next
      .Close
    End With  ' ** rst.

    .Close
  End With  ' ** dbs.

  Debug.Print "ADDS: " & CStr(lngAdded)
  Debug.Print "'PROPS: " & CStr(lngProps)

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rpt = Nothing
  Set fld = Nothing
  Set rst = Nothing
  Set dbs = Nothing
  Set prp = Nothing
  Set vbc = Nothing
  Set vbc2 = Nothing
  Set vbp = Nothing

  VBA_Component_Properties = blnRetValx

End Function

Private Function VBA_IsControl(varName As Variant) As Boolean
' ** Check these property names against control names.
' ** True will put the form/report name into vbcom_objects field.
' ** Called by:
' **   VBA_Component_Properties(), Above.

  Const THIS_PROC As String = "VBA_IsControl"

  blnRetValx = False

  Select Case varName
  Case "Sum_Of_TotalCost", "TotalCost_Grand_Total_Sum", "tmpDate", "tmpRate"
    ' ** RPT rptAssetList, FRM frmAssets_Add.
    blnRetValx = True
  Case "assettype_description", "description_masterasset", "description_masterasset_lbl", _
    "Revcode_SortOrderHeader", "TotalCost_Label", "TotalShareface_Label"
    ' ** Court reports.
    blnRetValx = True
  Case "accountno", "ActiveAssets.accountno", "ActiveAssets.assetno", "assettype", "assetno", _
    "assettype", "assettype_description_Label", "CompanyAddress1", "CompanyAddress2", "CompanyCity"
    blnRetValx = True
  Case "CompanyName", "CompanyPhone", "CompanyState", "CompanyZip", "Current_Label", _
    "currentDate", "DescCusip", "description", "assettype_description", "Difference_Label"
    blnRetValx = True
  Case "Dividend", "due", "End_Date", "EndDate", "Interest", _
    "IsHid", "ledger.assetno", "ledger_HIDDEN", "Location_ID", "marketvalue"
    blnRetValx = True
  Case "marketvaluecurrent", "masterasset.assetno", "masterasset_description", "masterasset_description_Label", _
    "masterasset_TYPE", "Model_Label", "multiplier", "NegativeIcash", "NegativePcash", "PositiveIcash"
    blnRetValx = True
  Case "PositivePcash", "rate", "Schedule_ID", "Schedule_ID_lbl", "Schedule_Name", _
    "sequence_no", "sortOrder", "SumCost_dp", "SumCost_ws", "SumICash_dp"
    blnRetValx = True
  Case "SumICash_ws", "SumPCash_dp", "SumPCash_ws", "taxcode", "transdate", _
    "yield"
    blnRetValx = True
  End Select

  VBA_IsControl = blnRetValx

End Function

Private Function VBA_Component_Type(varType As Variant) As String
' ** Currently not called or used.

  Const THIS_PROC As String = "VBA_Component_Type"

  Dim strRetVal As String

  strRetVal = vbNullString

  ' ** vbext_ComponentType enumeration:
  ' **     1  vbext_ct_StdModule        Standard Module
  ' **     2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
  ' **     3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm
  ' **                                  in the VBA Editor is called a designer.
  ' **    11  vbext_ct_ActiveXDesigner
  ' **   100  vbext_ct_Document         Module behind Form or Report.
  If IsNull(varType) = False Then
    Select Case varType
    Case vbext_ct_StdModule
      strRetVal = "Standard Module"
    Case vbext_ct_ClassModule
      strRetVal = "Class Module"
    Case vbext_ct_MSForm
      strRetVal = "User Form"
    'Case vbext_ct_ResFile             ' ** Not supported in Access 2000.
    '  strRetVal = "Res File"
    'Case vbext_ct_RelatedDocument
    '  strRetVal = "Related Document"  ' ** Not supported in Access 2000.
    Case vbext_ct_ActiveXDesigner
      strRetVal = "ActiveX"
    Case vbext_ct_Document
      strRetVal = "Form/Report Module"
    Case Else
      strRetVal = "{unknown}"
    End Select
  End If

'The Type property settings for the Window object are described in the following table:
'Constant Value Description
'vbext_wt_CodeWindow 0 Code window
'vbext_wt_Designer 1 Designer
'vbext_wt_Browser 2 Object Browser
'vbext_wt_Immediate 5 Immediate window
'vbext_wt_ProjectWindow 6 Project window
'vbext_wt_PropertyWindow 7 Properties window
'vbext_wt_Find 8 Find dialog box
'vbext_wt_FindReplace 9 Search and Replace dialog box
'vbext_wt_LinkedWindowFrame 11 Linked window frame
'vbext_wt_MainWindow 12 Main window
'vbext_wt_Watch 3 Watch window
'vbext_wt_Locals 4 Locals window
'vbext_wt_Toolbox 10 Toolbox
'vbext_wt_ToolWindow 15 Tool window

'The Type property settings for the Reference object are described in the following table:
'Constant Value Description
'vbext_rk_TypeLib 0 Type library
'vbext_rk_Project 1 Project

  VBA_Component_Type = strRetVal

End Function

Public Function VBA_Chk_Code() As Boolean
' ** Check validity of field names in a specified module procedure.
' ** Currently not called or used.

  Const THIS_PROC As String = "VBA_Chk_Code"

  Dim vbp As VBProject, vbc As VBComponent
  Dim dbs As DAO.Database, tdf As DAO.TableDef, qdf As DAO.QueryDef, fld As DAO.Field, obj As Object
  Dim strFindMod As String
  Dim strFindProc1 As String, strFindProc2 As String
  Dim lngFinds As Long, arr_varFind() As Variant
  Dim lngLines As Long
  Dim lngProcStart As Long, lngProcEnd As Long
  Dim lngFindStart As Long, lngFindEnd As Long
  Dim lngNestWithCnt As Long
  Dim strLine As String
  Dim strProcName As String, strTblName As String
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim lngFlds As Long, arr_varFld() As Variant
  Dim lngTblElem As Long
  Dim blnAllOK As Boolean
  Dim lngPos01 As Long, lngPos02 As Long
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  ' ** Array: arr_varFind().
  Const FND_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const FND_FIND1 As Integer = 0
  Const FND_FIND2 As Integer = 1
  Const FND_OBJ   As Integer = 2
  Const FND_NEST  As Integer = 3
  Const FND_START As Integer = 4
  Const FND_END   As Integer = 5

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const T_NAME     As Integer = 0
  Const T_TYPE     As Integer = 1
  Const T_OK       As Integer = 2
  Const T_FND_ELEM As Integer = 3

  ' ** Array: arr_varFld().
  Const F_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const F_TBL      As Integer = 0
  Const F_FLD      As Integer = 1
  Const F_OK       As Integer = 2
  Const F_TBL_ELEM As Integer = 3

  blnRetValx = True

  ' ** Module to check.
  strFindMod = "modCourtReportsCA"

  ' ** Procedure to check.
  strFindProc1 = "Public Function CABuildCourtReportData("
  strFindProc2 = "End Function"

  lngFinds = 0&
  ReDim arr_varFind(FND_ELEMS, 0)
  ' ***********************************************
  ' ** Array: arr_varFind()
  ' **
  ' **   Element  Description          Constant
  ' **   =======  ===================  ==========
  ' **      0     Open Search Word     FND_FIND1
  ' **      1     Close Search Word    FND_FIND2
  ' **      2     Nesting Term         FND_NEST
  ' **      3     Object Name          FND_OBJ
  ' **      4     Line Start           FND_START
  ' **      5     Line End             FND_END
  ' **
  ' ***********************************************

  ' ** rsDataOut.Open "tmpCourtReportData"
  lngFinds = lngFinds + 1&
  lngE = lngFinds - 1&
  ReDim Preserve arr_varFind(FND_ELEMS, lngE)
  arr_varFind(FND_FIND1, lngE) = "With rsDataOut"
  arr_varFind(FND_FIND2, lngE) = "End With"
  arr_varFind(FND_OBJ, lngE) = "rsDataOut"
  arr_varFind(FND_NEST, lngE) = "With "  ' ** Keep track of nesting.
  arr_varFind(FND_START, lngE) = 0&
  arr_varFind(FND_END, lngE) = 0&
'I already know rsDataOut never uses a specified assignment, like "rsDataOut.", "rsDataOut(", or "rsDataOut!",
'only "With rsDataOut" followed by multiple " .Fields(".

  ' ** rsDataIn.Open "qryCourtReport - Summary-1-CA"
'59 - 345
  lngFinds = lngFinds + 1&
  lngE = lngFinds - 1&
  ReDim Preserve arr_varFind(FND_ELEMS, lngE)
  arr_varFind(FND_FIND1, lngE) = "rsDataIn.Open"
  arr_varFind(FND_FIND2, lngE) = "rsDataIn.Open"  ' ** Search till next assignment.
  arr_varFind(FND_OBJ, lngE) = "rsDataIn"
  arr_varFind(FND_NEST, lngE) = vbNullString  ' ** No nesting allowed.
  arr_varFind(FND_START, lngE) = 0&
  arr_varFind(FND_END, lngE) = 0&
'I already know rsDataIn never uses "With rsDataIn", only specified assignments.
'Between the assignments, there will be a "With rsDataOut" until "End With",
'so left of "=" is rsDataOut and right of "=" is rsDataIn (which is always specified).
'Collect all fields that are specified for rsDataIn.
'All those without assignment are rsDataOut.

  ' ** rsDataIn.Open "qryAssetList-CA"
  lngFinds = lngFinds + 1&
  lngE = lngFinds - 1&
  ReDim Preserve arr_varFind(FND_ELEMS, lngE)
  arr_varFind(FND_FIND1, lngE) = "rsDataIn.Open"
  arr_varFind(FND_FIND2, lngE) = "rsDataIn.Open"  ' ** Search till next assignment.
  arr_varFind(FND_OBJ, lngE) = "rsDataIn"
  arr_varFind(FND_NEST, lngE) = vbNullString  ' ** No nesting allowed.
  arr_varFind(FND_START, lngE) = 0&
  arr_varFind(FND_END, lngE) = 0&

  ' ** rsDataIn.Open "account"
  lngFinds = lngFinds + 1&
  lngE = lngFinds - 1&
  ReDim Preserve arr_varFind(FND_ELEMS, lngE)
  arr_varFind(FND_FIND1, lngE) = "rsDataIn.Open"
  arr_varFind(FND_FIND2, lngE) = "End Function"
  arr_varFind(FND_OBJ, lngE) = "rsDataIn"
  arr_varFind(FND_NEST, lngE) = vbNullString  ' ** No nesting allowed.
  arr_varFind(FND_START, lngE) = 0&
  arr_varFind(FND_END, lngE) = 0&

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)
  ' **********************************************
  ' ** Array: arr_varTbl()
  ' **
  ' **   Element  Description       Constant
  ' **   =======  ================  ============
  ' **      0     Table Name        T_NAME
  ' **      1     Table or Query    T_TYPE
  ' **      2     Table is OK       T_OK
  ' **      3     arr_varFnd()      T_FND_ELEM
  ' **
  ' **********************************************

  lngFlds = 0&
  ReDim arr_varFld(F_ELEMS, 0)
  ' ********************************************
  ' ** Array: arr_varFld()
  ' **
  ' **   Element  Description     Constant
  ' **   =======  ==============  ============
  ' **      0     Table Name      F_TBL
  ' **      1     Field Name      F_FLD
  ' **      2     Field is OK     F_OK
  ' **      3     arr_varTbl()    F_TBL_ELEM
  ' **
  ' ********************************************

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    Set vbc = .VBComponents(strFindMod)
    With vbc
      With .CodeModule

        For lngX = 0& To (lngFinds - 1&)

          strTblName = vbNullString
          lngTblElem = 0&
          lngFindStart = 0&: lngFindEnd = 0&
          lngProcStart = 0&: lngProcEnd = 0&

          ' ** Search for procedure start.
          lngLines = .CountOfLines
          lngProcStart = 0&: lngProcEnd = 0&
          strProcName = vbNullString
          For lngY = 1& To lngLines
            If .Lines(lngY, 1) <> vbNullString Then
              strLine = Trim(.Lines(lngY, 1))
              If Left(strLine, Len(strFindProc1)) = strFindProc1 Then
                lngProcStart = lngY
                strProcName = .ProcOfLine(lngY, vbext_pk_Proc)
                Exit For
              End If
            End If
          Next  ' ** lngY, each line.
          If lngProcStart > 0& Then
            ' ** Search for procedure end.
            For lngY = lngProcStart To lngLines
              If .Lines(lngY, 1) <> vbNullString Then
                strLine = Trim(.Lines(lngY, 1))
                If Left(strLine, Len(strFindProc2)) = strFindProc2 Then
                  If .ProcOfLine(lngY, vbext_pk_Proc) = strProcName Then
                    lngProcEnd = lngY
                  End If
                  Exit For
                End If
              End If
            Next  ' ** lngY, each line.
            If lngProcEnd > 0& Then
              ' ** Search for block start.
              lngFindStart = 0&: lngFindEnd = 0&
              If lngX = 0& Then
                lngZ = lngProcStart
              Else
                lngZ = arr_varFind(FND_END, (lngX - 1&))
              End If
              For lngY = (lngZ + 1&) To lngProcEnd
                If .Lines(lngY, 1) <> vbNullString Then
                  strLine = Trim(.Lines(lngY, 1))
                  lngPos01 = InStr(strLine, " "): lngPos02 = 0&
                  If lngPos01 > 0 Then
                    ' ** There will be a space in the line I'm looking for.
                    varTmp00 = Trim(Left(strLine, lngPos01))
                    If IsNumeric(varTmp00) = True Then
                      ' ** Numbered line.
                      strLine = Trim(Mid(strLine, lngPos01))
                    End If
                    If Left(strLine, 1) <> "'" Then
                      ' ** Ignore remarks.
                      If Left(strLine, Len(arr_varFind(FND_FIND1, lngX))) = arr_varFind(FND_FIND1, lngX) Then
                        lngFindStart = lngY
                        arr_varFind(FND_START, lngX) = lngFindStart
                      End If
                    End If
                  End If
                End If
                If lngFindStart > 0& Then
                  Exit For
                End If
              Next  ' ** lngY, each line.
              If blnRetValx = True Then
                ' ** Search for block end.
                If lngFindStart > 0& Then
                  lngNestWithCnt = 0&
                  For lngY = (lngFindStart + 1&) To lngProcEnd
                    If .Lines(lngY, 1) <> vbNullString Then
                      strLine = Trim(.Lines(lngY, 1))
                      lngPos01 = InStr(strLine, " "): lngPos02 = 0&
                      If lngPos01 > 0 Then
                        ' ** There will be a space in the line I'm looking for.
                        varTmp00 = Trim(Left(strLine, lngPos01))
                        If IsNumeric(varTmp00) = True Then
                          ' ** Numbered line.
                          strLine = Trim(Mid(strLine, lngPos01))
                        End If
                        If Left(strLine, 1) <> "'" Then
                          ' ** Ignore remarks.
                          If arr_varFind(FND_NEST, lngX) <> vbNullString Then
                            If Left(strLine, Len(arr_varFind(FND_NEST, lngX))) = arr_varFind(FND_NEST, lngX) Then
                              lngNestWithCnt = lngNestWithCnt + 1&
                            End If
                          End If
                          If Left(strLine, Len(arr_varFind(FND_FIND2, lngX))) = arr_varFind(FND_FIND2, lngX) Then
                            If lngNestWithCnt = 0& Then
                              If arr_varFind(FND_FIND2, lngX) = "End With" Then
                                lngFindEnd = lngY
                              Else
                                lngFindEnd = lngY - 1&
                              End If
                              arr_varFind(FND_END, lngX) = lngFindEnd
                            Else
                              lngNestWithCnt = lngNestWithCnt - 1&
                            End If
                          End If
                        End If
                      End If
                    End If
                    If lngFindEnd > 0& Then
                      Exit For
                    End If
                  Next  ' ** lngY, each line.
                  If lngFindEnd = 0& Then
                    Beep
                    blnRetValx = False
                    Debug.Print "'END FIND NOT FOUND! " & lngX & " '" & arr_varFind(FND_FIND1, lngX) & _
                      "', '" & arr_varFind(FND_FIND2, lngX) & "' " & lngFindStart & " - ?"
                  End If
                Else
                  Beep
                  blnRetValx = False
                  Debug.Print "'WITH NOT FOUND!"
                End If  ' ** lngFindStart > 0&
              End If  ' ** blnRetValx
            Else
              Beep
              blnRetValx = False
              Debug.Print "'PROC END NOT FOUND!"
            End If  ' ** lngProcEnd > 0&
          Else
            Beep
            blnRetValx = False
            Debug.Print "'PROC NOT FOUND!"
          End If  ' ** lngProcStart > 0&

          If blnRetValx = True Then

            ' ** Get the table name.
            strTblName = vbNullString
            For lngY = lngFindStart To lngProcStart Step -1&
              If .Lines(lngY, 1) <> vbNullString Then
                strLine = Trim(.Lines(lngY, 1))
                lngPos01 = InStr(strLine, " "): lngPos02 = 0&
                If lngPos01 > 0 Then
                  ' ** There will be a space in the line I'm looking for.
                  varTmp00 = Trim(Left(strLine, lngPos01))
                  If IsNumeric(varTmp00) = True Then
                    ' ** Numbered line.
                    strLine = Trim(Mid(strLine, lngPos01))
                  End If
                  If Left(strLine, 1) <> "'" Then
                    ' ** Ignore remarks.
                    strTmp01 = arr_varFind(FND_OBJ, lngX) & ".Open "
                    If Left(strLine, Len(strTmp01)) = strTmp01 Then
                      strTmp01 = Trim(Mid(strLine, Len(strTmp01)))
                      lngPos01 = InStr(strTmp01, Chr(34))
                      If lngPos01 > 0& Then
                        strTmp01 = Mid(strTmp01, (lngPos01 + 1&))
                        lngPos01 = InStr(strTmp01, Chr(34))
                        If lngPos01 > 0& Then
                          strTmp01 = Left(strTmp01, (lngPos01 - 1&))
                          strTblName = strTmp01
                        End If
                      End If
                      Exit For
                    End If
                  End If
                End If
              End If
            Next  ' ** lngY, each line backwards

            If strTblName <> vbNullString Then
 
              ' ** Find out whether it's a table or a query.
              lngTbls = lngTbls + 1&
              lngE = lngTbls - 1&
              ReDim Preserve arr_varTbl(T_ELEMS, lngE)
              arr_varTbl(T_NAME, lngE) = strTblName
              arr_varTbl(T_TYPE, lngE) = vbNullString
              arr_varTbl(T_OK, lngE) = True
              arr_varTbl(T_FND_ELEM, lngE) = lngX
              Set dbs = CurrentDb
              With dbs
                ' ** See if it's a table.
                For Each tdf In .TableDefs
                  With tdf
                    If .Name = strTblName Then
                      arr_varTbl(T_TYPE, lngE) = "Table"
                      Exit For
                    End If
                  End With
                Next
                If arr_varTbl(T_TYPE, lngE) = vbNullString Then
                  ' ** See if it's a query.
                  For Each qdf In .QueryDefs
                    With qdf
                      If .Name = strTblName Then
                        arr_varTbl(T_TYPE, lngE) = "Query"
                        Exit For
                      End If
                    End With
                  Next
                End If
                If arr_varTbl(T_TYPE, lngE) <> vbNullString Then
                  lngTblElem = lngE
                Else
                  Beep
                  blnRetValx = False
                  Debug.Print "'TABLE NOT FOUND! " & strTblName
                End If
                .Close
              End With

              If blnRetValx = True Then

                ' ** Now pick up the field names.
                For lngY = lngFindStart To lngFindEnd
                  If .Lines(lngY, 1) <> vbNullString Then
                    strLine = Trim(.Lines(lngY, 1))
                    lngPos01 = InStr(strLine, " "): lngPos02 = 0&
                    If lngPos01 > 0 Then
                      varTmp00 = Trim(Left(strLine, lngPos01))
                      If IsNumeric(varTmp00) = True Then
                        ' ** Numbered line.
                        strLine = Trim(Mid(strLine, lngPos01))
                      End If
                      If Left(strLine, 1) <> "'" Then
                        ' ** Ignore remarks.
                        '210         .Fields("ReportNumber") = counter * 10
                        lngPos01 = InStr(strLine, Chr(34))
                        If lngPos01 > 0& Then

                          If arr_varFind(FND_NEST, lngX) <> vbNullString Then
                            ' ** The first Find, for rsDataOut.

                            strTmp01 = Mid(strLine, (lngPos01 + 1&))
                            lngPos01 = InStr(strTmp01, Chr(34))
                            If lngPos01 > 0& Then
                              strTmp01 = Left(strTmp01, (lngPos01 - 1&))
                              lngFlds = lngFlds + 1&
                              lngE = lngFlds - 1&
                              ReDim Preserve arr_varFld(F_ELEMS, lngE)
                              arr_varFld(F_TBL, lngE) = arr_varTbl(T_NAME, lngTblElem)
                              arr_varFld(F_FLD, lngE) = strTmp01
                              arr_varFld(F_OK, lngE) = False
                              arr_varFld(F_TBL_ELEM, lngE) = lngTblElem
                            Else
                              Beep
                              blnRetValx = False
                              Debug.Print "'QUOTES OFF! " & lngY & " " & strLine
                            End If

                          Else

                            ' ** Check for non-specified assignments.
                            If Left(strLine, 8) = ".Fields(" Then
                              lngPos01 = 1&
                            Else
                              lngPos01 = InStr(strLine, " .Fields(")
                            End If
                            Do While lngPos01 > 0&
                              strTmp02 = Mid(strLine, (lngPos01 + 1&))
                              lngPos02 = InStr(strTmp02, Chr(34))
                              If lngPos02 > 0& Then
                                strTmp02 = Mid(strTmp02, (lngPos02 + 1&))
                                lngPos02 = InStr(strTmp02, Chr(34))
                                If lngPos02 > 0& Then
                                  strTmp02 = Left(strTmp02, (lngPos02 - 1&))
                                  lngFlds = lngFlds + 1&
                                  lngE = lngFlds - 1&
                                  ReDim Preserve arr_varFld(F_ELEMS, lngE)
'THIS REALLY SHOULDN'T BE HARD-CODED!
                                  arr_varFld(F_TBL, lngE) = arr_varTbl(T_NAME, 0)
                                  arr_varFld(F_FLD, lngE) = strTmp02
                                  arr_varFld(F_OK, lngE) = False
                                  arr_varFld(F_TBL_ELEM, lngE) = 0&
                                Else
                                  Beep
                                  blnRetValx = False
                                  Debug.Print "'QUOTES OFF! " & lngY & " " & strLine
                                End If
                              Else
                                Beep
                                blnRetValx = False
                                Debug.Print "'QUOTE NOT FOUND! NON-SPEC " & lngY & " " & strLine
                              End If
                              lngPos01 = InStr((lngPos01 + 1&), strTmp02, " .Fields(")
                              If blnRetValx = False Then
                                Exit Do
                              End If
                            Loop  ' ** Non-Specified assignment.

                            ' ** Check for specified assignments.
                            lngPos01 = InStr(strLine, (arr_varFind(FND_OBJ, lngX) & ".Fields("))
                            Do While lngPos01 > 0&
                              strTmp02 = Mid(strLine, (lngPos01 + 1&))
                              lngPos02 = InStr(strTmp02, Chr(34))
                              If lngPos02 > 0& Then
                                strTmp02 = Mid(strTmp02, (lngPos02 + 1&))
                                lngPos02 = InStr(strTmp02, Chr(34))
                                If lngPos02 > 0& Then
                                  strTmp02 = Left(strTmp02, (lngPos02 - 1&))
                                  lngFlds = lngFlds + 1&
                                  lngE = lngFlds - 1&
                                  ReDim Preserve arr_varFld(F_ELEMS, lngE)
                                  arr_varFld(F_TBL, lngE) = arr_varTbl(T_NAME, lngTblElem)
                                  arr_varFld(F_FLD, lngE) = strTmp02
                                  arr_varFld(F_OK, lngE) = False
                                  arr_varFld(F_TBL_ELEM, lngE) = lngTblElem
                                Else
                                  Beep
                                  blnRetValx = False
                                  Debug.Print "'QUOTES OFF! " & lngY & " " & strLine
                                End If
                              Else
                                Beep
                                blnRetValx = False
                                Debug.Print "'QUOTE NOT FOUND! SPEC " & lngY & " " & strLine
                              End If
                              lngPos01 = InStr((lngPos01 + 1&), strTmp02, (arr_varFind(FND_OBJ, lngX) & ".Fields("))
                              If blnRetValx = False Then
                                Exit Do
                              End If
                            Loop  ' ** Specified assignment.

                          End If  ' ** Nesting allowed or not.

                        End If  ' ** Has quotes.
                      End If  ' ** Not a remark.
                    End If  ' ** Has a space.
                  End If  ' ** <> vbNullString
                  If blnRetValx = False Then
                    Exit For
                  End If
                Next  ' ** lngY, each line.

              End If  ' ** blnRetValx

            End If  ' ** strTblName

          End If  ' ** blnRetValx

          If blnRetValx = False Then
            Exit For
          End If

        Next  ' ** lngX, each Find

        If blnRetValx = True Then
          ' ** We've now got a list of fields to check against the tables.
          Set dbs = CurrentDb
          With dbs
            ' **********************************************
            ' ** Array: arr_varFld()
            ' **
            ' **   Element  Description     Constant
            ' **   =======  ==============  ==============
            ' **      0     Table Name      F_TBL
            ' **      1     Field Name      F_FLD
            ' **      2     Field is OK     F_OK
            ' **      3     arr_varTbl()    F_TBL_ELEM
            ' **
            ' **********************************************
            blnAllOK = True
            For lngY = 0& To (lngFlds - 1&)
              Select Case arr_varTbl(T_TYPE, arr_varFld(F_TBL_ELEM, lngY))
              Case "Table"
                Set obj = .TableDefs(arr_varTbl(T_NAME, arr_varFld(F_TBL_ELEM, lngY)))
              Case "Query"
                Set obj = .QueryDefs(arr_varTbl(T_NAME, arr_varFld(F_TBL_ELEM, lngY)))
              End Select
              With obj
                For Each fld In .Fields
                  With fld
                    If .Name = arr_varFld(F_FLD, lngY) Then
                      arr_varFld(F_OK, lngY) = True
                      Exit For
                    End If
                  End With
                Next
                If arr_varFld(F_OK, lngY) = False Then
                  blnAllOK = False
                End If
              End With
            Next
            If blnAllOK = True Then
              For lngX = 0& To (lngTbls - 1&)
                lngE = 0&
                For lngY = 0& To (lngFlds - 1&)
                  If arr_varFld(F_TBL_ELEM, lngY) = lngX Then
                    lngE = lngE + 1&
                  End If
                Next
                Debug.Print "'FIELDS: " & Left(CStr(lngE) & "   ", 3) & ", " & _
                  "START: " & Right("   " & CStr(arr_varFind(FND_START, arr_varTbl(T_FND_ELEM, lngX))), 3) & _
                  ", END: " & Right("   " & CStr(arr_varFind(FND_END, arr_varTbl(T_FND_ELEM, lngX))), 3) & _
                  ", TABLE: " & UCase$(arr_varTbl(T_TYPE, lngX)) & ": '" & arr_varTbl(T_NAME, lngX)
              Next
            Else
              Debug.Print "'FIELD NOT FOUND!"
              For lngX = 0& To (lngFlds - 1&)
                If arr_varFld(F_OK, lngX) = False Then
                  Debug.Print "'" & UCase(arr_varTbl(T_TYPE, arr_varFld(F_TBL_ELEM, lngX))) & ": " & _
                    arr_varTbl(T_NAME, arr_varFld(F_TBL_ELEM, lngX)) & " ; FIELD: " & arr_varFld(F_FLD, lngX)
                End If
              Next
            End If
            .Close
          End With
        End If

      End With    ' ** CodeModule
    End With      ' ** vbc
  End With        ' ** vbp

  Debug.Print blnRetValx

  Beep

  Set vbc = Nothing
  Set vbp = Nothing
  Set fld = Nothing
  Set qdf = Nothing
  Set tdf = Nothing
  Set obj = Nothing
  Set dbs = Nothing

  VBA_Chk_Code = blnRetValx

End Function

Public Function VBA_ThisChk() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_ThisChk"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngLines As Long, lngDecLines As Long
  Dim strModName As String, strLine As String, strProcName As String, strLastProc As String
  Dim lngThisNames As Long, lngThisProcs As Long, lngErrs As Long
  Dim blnFound1 As Boolean, blnFound2 As Boolean
  Dim intPos01 As Integer
  Dim strTmp01 As String, strTmp02 As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisNames = 0&: lngThisProcs = 0&: lngErrs = 0&
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents  ' ** One-Based.
      With vbc
        strModName = .Name
        strProcName = vbNullString: strLastProc = vbNullString
        blnFound1 = False: blnFound2 = False
        strTmp01 = vbNullString
        Set cod = .CodeModule
        With cod
          lngLines = .CountOfLines
          lngDecLines = .CountOfDeclarationLines
          For lngX = 1& To lngLines
            strLine = Trim(.Lines(lngX, 1))
            If Left(strLine, 1) <> "'" Then
              If strLine <> vbNullString Then
                If InStr(strLine, "' ** THIS_NAME") = 0 Then
                  intPos01 = InStr(strLine, "THIS_NAME As String = ")  ' ** THIS_NAME As String = "zz_mod_ModuleFormatFuncs"
                  If intPos01 > 0 Then
                    lngThisNames = lngThisNames + 1&
                    blnFound1 = True
                    strTmp01 = Mid(strLine, intPos01)
                    intPos01 = InStr(strTmp01, Chr(34))
                    If intPos01 > 0 Then
                      strTmp01 = Mid(strTmp01, (intPos01 + 1))
                      intPos01 = InStr(strTmp01, Chr(34))
                      If intPos01 > 1 Then
                        strTmp01 = Left(strTmp01, (intPos01 - 1))
                        strTmp02 = strModName
                        If Left(strTmp02, 5) = "Form_" Then strTmp02 = Mid(strTmp02, 6)
                        If Left(strTmp02, 7) = "Report_" Then strTmp02 = Mid(strTmp02, 8)
                        If strTmp01 <> strTmp02 Then
                          lngErrs = lngErrs + 1&
                          Debug.Print "'NOT SAME!  MOD: '" & strModName & "'  THIS_NAME: '" & strTmp01 & "'"
                        End If
                      Else
                        If .ProcOfLine(lngX, vbext_pk_Proc) <> "VBA_ThisChk" And _
                            .ProcOfLine(lngX, vbext_pk_Proc) <> "VBA_This_Name" Then
                          lngErrs = lngErrs + 1&
                          Debug.Print "'WHAT?  MOD: '" & strModName & "'  LINE: " & strLine
                          'Stop
                        End If
                      End If
                    End If
                  ElseIf lngX > lngDecLines Then
                    strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                    If strProcName <> "VBA_ThisChk" Then
                      If strLastProc = vbNullString And strProcName = vbNullString Then
                        ' ** Evidently we're not out of the declaration section!
                      Else
                        If strLastProc = vbNullString And strProcName <> vbNullString Then
                          ' ** First procedure.
                          strLastProc = strProcName
                          blnFound2 = False
                        ElseIf strProcName <> strLastProc Then
                          ' ** New procedure:
                          If blnFound2 = False Then
                            If Left(strLastProc, 4) <> "Quik" Then
                              lngErrs = lngErrs + 1&
                              Debug.Print "'THIS_PROC NOT FOUND!  '" & strModName & "'  PROC: '" & strLastProc & "'"
                            End If
                            strLastProc = strProcName
                          Else
                            strLastProc = strProcName
                            blnFound2 = False
                          End If
                        End If
                        intPos01 = InStr(strLine, "THIS_PROC As String = ")
                        If intPos01 > 0 Then
                          lngThisProcs = lngThisProcs + 1&
                          blnFound2 = True
                          strTmp01 = Mid(strLine, intPos01)
                          intPos01 = InStr(strTmp01, Chr(34))
                          If intPos01 > 0 Then
                            strTmp01 = Mid(strTmp01, (intPos01 + 1))
                            intPos01 = InStr(strTmp01, Chr(34))
                            If intPos01 > 1 Then
                              strTmp01 = Left(strTmp01, (intPos01 - 1))
                              If strTmp01 <> strProcName Then
                                If Right(strTmp01, 4) <> " Let" And Right(strTmp01, 4) <> " Get" Then
                                  lngErrs = lngErrs + 1&
                                  Debug.Print "'PROC NOT SAME!  MOD: '" & strModName & "'  PROC: '" & strProcName & "'  /  '" & strTmp01 & "'"
                                End If
                              End If
                            Else
                              If Left(strLine, 11) <> "Const TPROC" And _
                                  InStr(strLine, Chr(34) & "Const THIS_NAME") = 0 And _
                                  InStr(strLine, Chr(34) & "Private Const THIS_NAME") = 0 Then
                                If .ProcOfLine(lngX, vbext_pk_Proc) <> "VBA_ThisChk" Then
                                  lngErrs = lngErrs + 1&
                                  Debug.Print "'WHAT?  MOD: '" & strModName & "'  LINE: " & strLine
                                  'Stop
                                End If
                              End If
                            End If
                          Else
                            Stop
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
              End If  ' ** vbNullString.
            End If  ' ** Remark.
          Next  ' ** lngX.
        End With  ' ** cod.
        If blnFound1 = False Then
          Debug.Print "'THIS_NAME NOT FOUND!  " & strModName
        End If
        If blnFound2 = False Then
          If strModName <> "modGlobConst" Then
            Debug.Print "'THIS_PROC NOT FOUND!  '" & strModName & "'  PROC: '" & strLastProc & "'"
          End If
        End If
      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.

  If lngErrs = 0& Then
    Debug.Print "'NO ERRS!"
  Else
    Debug.Print "'ERRS: " & CStr(lngErrs)
  End If
  Debug.Print "'THIS_NAME'S: " & CStr(lngThisNames)
  Debug.Print "'THIS_PROC'S: " & CStr(lngThisProcs)
  Debug.Print "'DONE!"

  Beep

'NO ERRS!
'THIS_NAME'S: 509
'THIS_PROC'S: 9974
'DONE!

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  VBA_ThisChk = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
  End Select
  Resume EXITP

End Function

Public Function VBA_KeyDownChk() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_KeyDownChk"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngDims As Long, arr_varDim() As Variant
  Dim lngVars As Long, arr_varVar As Variant
  Dim strModName As String, strLastModName As String, strLine As String, strProcName As String
  Dim lngLines As Long, lngDecLines As Long, intMode As Integer, lngLoopCnt As Long
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean
  Dim lngTmp01 As Long
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varDim().
  Const D_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const D_DID  As Integer = 0
  Const D_VID  As Integer = 1
  Const D_VNAM As Integer = 2
  Const D_PID  As Integer = 3
  Const D_PNAM As Integer = 4
  Const D_LIN  As Integer = 5
  Const D_RAW  As Integer = 6

  ' ** Array: arr_varVar().
  'Const V_AID  As Integer = 0
  'Const V_DID  As Integer = 1
  'Const V_VID  As Integer = 2
  Const V_VNAM As Integer = 3
  'Const V_PID  As Integer = 4
  Const V_PNAM As Integer = 5
  Const V_VAR1 As Integer = 6
  Const V_USE1 As Integer = 7
  Const V_VAR2 As Integer = 8
  Const V_USE2 As Integer = 9
  Const V_VAR3 As Integer = 10
  Const V_USE3 As Integer = 11
  Const V_BEG  As Integer = 12
  Const V_END  As Integer = 13

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  intMode = 2

  Select Case intMode
  Case 1
    ' ** Collect the data.

    lngDims = 0&
    ReDim arr_varDim(D_ELEMS, 0)

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For Each vbc In .VBComponents  ' ** One-Based.
        With vbc
          strModName = .Name
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngX = lngDecLines To lngLines
              strLine = Trim(.Lines(lngX, 1))
              If Left(strLine, 1) <> "'" Then
                If strLine <> vbNullString Then
                  strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                  If strProcName <> "VBA_KeyDown" Then
                    If Right(strProcName, 8) = "_KeyDown" Then
                      If Left(strLine, 4) = "Dim " Then
                        If strLine = "Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer" Then
                          ' ** Normal.
                        ElseIf strLine = "Dim intRetVal As Integer" Then
                          ' ** Normal.
                        Else
                          lngDims = lngDims + 1&
                          lngE = lngDims - 1&
                          ReDim Preserve arr_varDim(D_ELEMS, lngE)
                          arr_varDim(D_DID, lngE) = lngThisDbsID
                          arr_varDim(D_VID, lngE) = CLng(0)
                          arr_varDim(D_VNAM, lngE) = strModName
                          arr_varDim(D_PID, lngE) = CLng(0)
                          arr_varDim(D_PNAM, lngE) = strProcName
                          arr_varDim(D_LIN, lngE) = lngX
                          arr_varDim(D_RAW, lngE) = strLine
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            Next
          End With
        End With
      Next
    End With

    If lngDims > 0& Then
      Set dbs = CurrentDb
      With dbs
        Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
        With rst
          .MoveFirst
          For lngX = 0& To (lngDims - 1&)
            .FindFirst "[dbs_id] = " & CStr(arr_varDim(D_DID, lngX)) & " And [vbcom_name] = '" & arr_varDim(D_VNAM, lngX) & "'"
            If .NoMatch = False Then
              arr_varDim(D_VID, lngX) = ![vbcom_id]
            Else
              Stop
            End If
          Next
          .Close
        End With
        Set rst = Nothing
        Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
        With rst
          .MoveFirst
          For lngX = 0& To (lngDims - 1&)
            .FindFirst "[dbs_id] = " & CStr(arr_varDim(D_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varDim(D_VID, lngX)) & " And " & _
              "[vbcomproc_name] = '" & arr_varDim(D_PNAM, lngX) & "'"
            If .NoMatch = False Then
              arr_varDim(D_PID, lngX) = ![vbcomproc_id]
            End If
          Next
          .Close
        End With
        Set rst = Nothing
        Set rst = .OpenRecordset("zz_tbl_VBComponent_Array", dbOpenDynaset, dbAppendOnly)
        With rst
          For lngX = 0& To (lngDims - 1&)
            .AddNew
            ![dbs_id] = arr_varDim(D_DID, lngX)
            ![vbcom_id] = arr_varDim(D_VID, lngX)
            ![vbcomproc_id] = arr_varDim(D_PID, lngX)
            ' ** ![vbarr_id] : AutoNumber.
            ![vbarr_module] = arr_varDim(D_VNAM, lngX)
            ![vbarr_procedure] = arr_varDim(D_PNAM, lngX)
            ![vbarr_linenum] = arr_varDim(D_LIN, lngX)
            ![vbarr_text] = arr_varDim(D_RAW, lngX)
            ![vbarr_datemodified] = Now()
            .Update
          Next
          .Close
        End With
        .Close
      End With
    End If

    Debug.Print "'DIMS: " & CStr(lngDims)
'DIMS: 91
'DONE!

  Case 2
    ' ** See if they're used.

    Set dbs = CurrentDb
    With dbs
      Set qdf = .QueryDefs("zz_qry_VBComponent_KeyDown_04")
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngVars = .RecordCount
        .MoveFirst
        arr_varVar = .GetRows(lngVars)
        ' *******************************************************
        ' ** Array: arr_varVar()
        ' **
        ' **   Field  Element  Name                  Constant
        ' **   =====  =======  ====================  ==========
        ' **     1       0     vbarr_id              V_AID
        ' **     2       1     dbs_id                V_DID
        ' **     3       2     vbcom_id              V_VID
        ' **     4       3     vbcom_name            V_VNAM
        ' **     5       4     vbcomproc_id          V_PID
        ' **     6       5     vbcomproc_name        V_PNAM
        ' **     7       6     var1                  V_VAR1
        ' **     8       7     vb_use1               V_USE1
        ' **     9       8     var2                  V_VAR2
        ' **    10       9     vb_use2               V_USE2
        ' **    11      10     var3                  V_VAR3
        ' **    12      11     vb_use3               V_USE3
        ' **    13      12     vbcomproc_line_beg    V_BEG
        ' **    14      13     vbcomproc_line_end    V_END
        ' **
        ' *******************************************************
        .Close
      End With
      Set rst = Nothing
      Set qdf = Nothing
      .Close
    End With
    Set dbs = Nothing

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      strLastModName = vbNullString
      For lngX = 0& To (lngVars - 1&)
        If arr_varVar(V_VNAM, lngX) <> strLastModName Then
          Set vbc = .VBComponents(arr_varVar(V_VNAM, lngX))
          lngLines = vbc.CodeModule.CountOfLines
          lngDecLines = vbc.CodeModule.CountOfDeclarationLines
        End If
        With vbc
          Set cod = .CodeModule
          With cod
            If IsNull(arr_varVar(V_VAR2, lngX)) = True Then
              lngLoopCnt = 1&
            ElseIf IsNull(arr_varVar(V_VAR3, lngX)) = True Then
              lngLoopCnt = 2&
            Else
              lngLoopCnt = 3&
            End If
            For lngY = 1& To lngLoopCnt
              blnFound = False
              For lngZ = arr_varVar(V_BEG, lngX) To arr_varVar(V_END, lngX)
                strLine = Trim(.Lines(lngZ, 1))
                If strLine <> vbNullString Then
                  If Left(strLine, 1) <> "'" Then
                    Select Case blnFound
                    Case True
                      Select Case lngY
                      Case 1&
                        If InStr(strLine, arr_varVar(V_VAR1, lngX)) > 0 Then
                          arr_varVar(V_USE1, lngX) = arr_varVar(V_USE1, lngX) + 1&
                        End If
                      Case 2&
                        If InStr(strLine, arr_varVar(V_VAR2, lngX)) > 0 Then
                          arr_varVar(V_USE2, lngX) = arr_varVar(V_USE2, lngX) + 1&
                        End If
                      Case 3&
                        If InStr(strLine, arr_varVar(V_VAR3, lngX)) > 0 Then
                          arr_varVar(V_USE3, lngX) = arr_varVar(V_USE3, lngX) + 1&
                        End If
                      End Select
                    Case False
                      If Left(strLine, 4) = "Dim " Then
                        Select Case lngY
                        Case 1&
                          If InStr(strLine, " " & arr_varVar(V_VAR1, lngX) & " ") > 0 Then
                            blnFound = True
                          End If
                        Case 2&
                          If InStr(strLine, " " & arr_varVar(V_VAR2, lngX) & " ") > 0 Then
                            blnFound = True
                          End If
                        Case 3&
                          If InStr(strLine, " " & arr_varVar(V_VAR3, lngX) & " ") > 0 Then
                            blnFound = True
                          End If
                        End Select
                      End If
                    End Select
                  End If
                End If
              Next
            Next
          End With
        End With
      Next
    End With

    lngTmp01 = 0&
    For lngX = 0& To (lngVars - 1&)
      If arr_varVar(V_USE1, lngX) = 0& Then
        lngTmp01 = lngTmp01 + 1&
        Debug.Print "'NOT USED!  MOD: " & arr_varVar(V_VNAM, lngX) & "  PROC: " & arr_varVar(V_PNAM, lngX) & "  VAR: " & arr_varVar(V_VAR1, lngX)
        DoEvents
      ElseIf arr_varVar(V_USE1, lngX) = 1& Then
        Select Case arr_varVar(V_VAR1, lngX)
        Case "ctl", "dbs", "qdf", "rst"
          lngTmp01 = lngTmp01 + 1&
          Debug.Print "'NOT USED!  MOD: " & arr_varVar(V_VNAM, lngX) & "  PROC: " & arr_varVar(V_PNAM, lngX) & "  VAR: " & arr_varVar(V_VAR1, lngX)
          DoEvents
        Case Else
          ' ** Nothing right now.
        End Select
      End If
      If IsNull(arr_varVar(V_VAR2, lngX)) = False Then
        If arr_varVar(V_USE2, lngX) = 0& Then
          lngTmp01 = lngTmp01 + 1&
          Debug.Print "'NOT USED!  MOD: " & arr_varVar(V_VNAM, lngX) & "  PROC: " & arr_varVar(V_PNAM, lngX) & "  VAR: " & arr_varVar(V_VAR2, lngX)
          DoEvents
        ElseIf arr_varVar(V_USE2, lngX) = 1& Then
          Select Case arr_varVar(V_VAR2, lngX)
          Case "ctl", "dbs", "qdf", "rst"
            lngTmp01 = lngTmp01 + 1&
            Debug.Print "'NOT USED!  MOD: " & arr_varVar(V_VNAM, lngX) & "  PROC: " & arr_varVar(V_PNAM, lngX) & "  VAR: " & arr_varVar(V_VAR2, lngX)
            DoEvents
          Case Else
            ' ** Nothing right now.
          End Select
        End If
      End If
      If IsNull(arr_varVar(V_VAR3, lngX)) = False Then
        If arr_varVar(V_USE3, lngX) = 0& Then
          lngTmp01 = lngTmp01 + 1&
          Debug.Print "'NOT USED!  MOD: " & arr_varVar(V_VNAM, lngX) & "  PROC: " & arr_varVar(V_PNAM, lngX) & "  VAR: " & arr_varVar(V_VAR3, lngX)
          DoEvents
        ElseIf arr_varVar(V_USE3, lngX) = 1& Then
          Select Case arr_varVar(V_VAR3, lngX)
          Case "ctl", "dbs", "qdf", "rst"
            lngTmp01 = lngTmp01 + 1&
            Debug.Print "'NOT USED!  MOD: " & arr_varVar(V_VNAM, lngX) & "  PROC: " & arr_varVar(V_PNAM, lngX) & "  VAR: " & arr_varVar(V_VAR3, lngX)
            DoEvents
          Case Else
            ' ** Nothing right now.
          End Select
        End If
      End If
    Next

    If lngTmp01 > 0& Then
      Debug.Print "'UNUSED VARS: " & CStr(lngTmp01)
    Else
      Debug.Print "'NONE FOUND!"
    End If
'NONE FOUND!
'DONE!

  End Select

  Debug.Print "'DONE!"

  Beep

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_KeyDownChk = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
  End Select
  Resume EXITP

End Function

Public Function VBA_SetFocusChk() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_SetFocusChk"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngLines As Long, lngDecLines As Long
  Dim strModName As String, strLine As String
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean
  Dim lngTmp01 As Long
  Dim lngX As Long
  Dim blnRetVal As Boolean

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngTmp01 = 0&
  Set vbp = Application.VBE.ActiveVBProject
  With vbp

    For Each vbc In .VBComponents  ' ** One-Based.
      With vbc
        strModName = .Name
        If Left(strModName, 5) = "Form_" And InStr(strModName, "_Sub") = 0 Then
          blnFound = False
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngX = lngDecLines To lngLines
              strLine = Trim(.Lines(lngX, 1))
              If strLine <> vbNullString Then
                If Left(strLine, 1) <> "'" Then
                  If strLine = "Private Sub Form_Timer()" Or strLine = "Public Sub Form_Timer()" Then
                    blnFound = True
                    Exit For
                  End If
                End If
              End If
            Next
          End With
          If blnFound = False Then
            lngTmp01 = lngTmp01 + 1&
            Debug.Print "'NO TIMER!  " & strModName
          End If
        End If
      End With
    Next
    If lngTmp01 > 0& Then
      Debug.Print "'TIMERS: " & CStr(lngTmp01)
    Else
      Debug.Print "'ALL HAVE TIMERS!"
    End If


  End With

  Debug.Print "'DONE!"
  Beep

'NO TIMER!  Form_frmCalendar
'NO TIMER!  Form_zz_frmStatus
'TIMERS: 2
'DONE!

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_SetFocusChk = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
  End Select
  Resume EXITP

End Function

Public Function VBA_RenumErrh() As Boolean
' ** Reunumber ERRH: with all the same numbers.

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_RenumErrh"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strModName As String, strProcName As String, strLine As String, strNewLine As String
  Dim lngLines As Long, lngDecLines As Long
  Dim strNum As String, strLastNum As String, strNewNum As String, strLastProc As String
  Dim strERRH_Start As String, strLastErrNum As String
  Dim blnInERRH As Boolean, blnIsFirst As Boolean
  Dim intPos01 As Integer, intLen As Integer
  Dim strTmp01 As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

On Error GoTo 0

  blnRetVal = True

  ' ** This will only replace existing numbers.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    Set vbc = .VBComponents("Form_frmRpt_IncomeExpense")
      '1.  "Form_frmAccountContacts"
      '2.  "Form_frmAccountExport"
      '3.  "Form_frmAccountProfile_Sub"
      '4.  "Form_frmCheckReconcile"
      '5.  "Form_frmJournal_Columns"
      '6.  "Form_frmJournal_Columns_Sub"
      '7.  "Form_frmJournal_Sub3_Purchase"
      '8.  "Form_frmJournal_Sub4_Sold"
      '9.  "Form_frmJournal_Sub5_Misc"
      '10. "Form_frmRpt_Checks"
      '11. "Form_frmRpt_CourtReports_CA"
      '12. "Form_frmRpt_CourtReports_FL"
      '13. "Form_frmRpt_CourtReports_NS"
      '14. "Form_frmRpt_CourtReports_NY"
      '15. "Form_frmRpt_IncomeExpense"
      '16. "Form_frmStatementParameters"
      '17. "Form_frmTransaction_Audit"
      '18. "Form_frmTransaction_Audit_Sub"
      '19. "Form_frmTransaction_Audit_Sub_Criteria"
      '20. "modCourtReportsNY2"
      '21. "modStatementParamFuncs1"
    With vbc
      strModName = .Name
      Debug.Print "'" & strModName & "^"
      Set cod = .CodeModule
      With cod
        lngLines = .CountOfLines
        lngDecLines = .CountOfDeclarationLines
        strLastNum = vbNullString: strLastProc = vbNullString: strERRH_Start = vbNullString: strLastErrNum = vbNullString
        blnInERRH = False: blnIsFirst = False
        For lngX = lngDecLines To lngLines
          strLine = .Lines(lngX, 1)  ' ** No Trim()!
          strNewNum = vbNullString: strProcName = vbNullString
          If strLine <> vbNullString Then
            If Left(strLine, 1) <> "'" Then
              If IsNumeric(Left(strLine, 1)) = True Then
                intPos01 = InStr(strLine, " ")
                If intPos01 > 0 Then
                  strNum = Trim(Left(strLine, intPos01))
                  If IsNumeric(strNum) = True Then
                    strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                    If strProcName <> vbNullString Then
                      If strLastProc = vbNullString Then
                        ' ** 1st procedure of module.
                        strNewNum = "100"
                        strLastProc = strProcName
                        blnIsFirst = True
                      ElseIf strProcName = strLastProc Then
                        ' ** Continuing within same procedure.
                        Select Case blnInERRH
                        Case True
                          Select Case blnIsFirst
                          Case True
                            ' ** These will be numbered normally, but save starting number.
                            If strERRH_Start = vbNullString Then
                              strNewNum = CStr(Val(strLastNum) + 10)
                              strERRH_Start = strNewNum
                            Else
                              strNewNum = CStr(Val(strLastErrNum) + 10)
                            End If
                          Case False
                            If strLastErrNum = vbNullString Then
                              strNewNum = strERRH_Start
                            Else
                              strNewNum = CStr(Val(strLastErrNum) + 10)
                            End If
                          End Select
                        Case False
                          strNewNum = CStr(Val(strLastNum) + 10)
                        End Select
                      Else
                        ' ** Next procedure.
                        blnInERRH = False: blnIsFirst = False
                        strLastErrNum = vbNullString
                        strLastProc = strProcName
                        If Val(strLastNum) Mod 100 = 0 Then
                          ' ** Ended on a hundreds.
                          strNewNum = CStr(Val(strLastNum) + 100)
                        Else
                          strTmp01 = strLastNum
                          strTmp01 = Left(strTmp01, (Len(strTmp01) - 2)) & "00"
                          strNewNum = CStr(Val(strTmp01) + 100)
                        End If
                      End If
                      If strNewNum <> vbNullString Then
                        If Len(strNum) > Len(strNewNum) Then
                          intLen = Len(strNum)
                        ElseIf Len(strNewNum) > Len(strNum) Then
                          intLen = Len(strNewNum)
                        Else
                          intLen = Len(strNewNum)
                        End If
                        strTmp01 = Mid(strLine, (intLen + 1))  ' ** Code line without number.
                        strNewLine = Left(strNewNum & String(intLen, " "), intLen)
                        strNewLine = strNewLine & strTmp01
                        .ReplaceLine lngX, strNewLine
                        Select Case blnInERRH
                        Case True
                          strLastErrNum = strNewNum
                        Case False
                          strLastNum = strNewNum
                        End Select
                      Else
                        Debug.Print "'NO NEW LINE NUM!  LINE: " & CStr(lngX)
                        Stop
                      End If
                    Else
                      Debug.Print "'NUMBERED LINE WITH NO PROC!  LINE: " & CStr(lngX)
                      Stop
                    End If
                  End If  ' ** IsNumeric().
                End If  ' ** intPos01.
              Else
                If Left(strLine, 5) = "ERRH:" Then
                  blnInERRH = True
                End If
              End If  ' ** IsNumeric().
            End If  ' ** Remark.
          End If  ' ** vbNullString.
        Next  ' ** lngX.
      End With  ' ** cod.
    End With  ' ** vbc.
  End With  ' ** vbp.

  Debug.Print blnRetVal

  Beep

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  VBA_RenumErrh = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_RenumErrh_Find() As Boolean

  Const THIS_PROC As String = "VBA_RenumErrh_Find"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim strModName As String, strLine As String, strFind As String
  Dim lngLines As Long, lngDecLines As Long
  Dim intPos01 As Integer
  Dim lngTmp01 As Long
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  strFind = "VBA_RenumErrh"

  lngTmp01 = 0&

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        Set cod = .CodeModule
        With cod
          lngLines = .CountOfLines
          lngDecLines = .CountOfDeclarationLines
          For lngX = 1& To lngDecLines
            strLine = .Lines(lngX, 1)
            If Trim(strLine) <> vbNullString Then
              If Left(strLine, 1) = "'" Then
                intPos01 = InStr(strLine, strFind)
                If intPos01 > 0 Then
                  lngTmp01 = lngTmp01 + 1&
                  Debug.Print "'" & strModName
                  DoEvents
                End If
              End If  ' ** Remark.
            End If  ' ** vbNullString
          Next  ' ** lngX
        End With  ' ** cod.
      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.

  Debug.Print "'HITS: " & CStr(lngTmp01)
  DoEvents

'Form_frmAccountExport
'Form_frmAccountProfile_Sub
'Form_frmCheckReconcile
'Form_frmJournal_Columns
'Form_frmJournal_Columns_Sub
'Form_frmJournal_Sub3_Purchase
'Form_frmJournal_Sub4_Sold
'Form_frmJournal_Sub5_Misc
'Form_frmRpt_Checks
'Form_frmRpt_CourtReports_CA
'Form_frmRpt_CourtReports_FL
'Form_frmRpt_CourtReports_NS
'Form_frmRpt_CourtReports_NY
'Form_frmRpt_IncomeExpense
'Form_frmStatementParameters
'Form_frmTransaction_Audit
'Form_frmTransaction_Audit_Sub
'Form_frmTransaction_Audit_Sub_Criteria
'modCourtReportsNY2
'modStatementParamFuncs1
'HITS: 20
'DONE!

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_RenumErrh_Find = blnRetVal

End Function

Public Function VBA_CodeLineNum_Doc() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_CodeLineNum_Doc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim strModName As String, strLine As String, strMax As String
  Dim lngLines As Long, lngDecLines As Long
  Dim lngNums As Long, arr_varNum() As Variant
  Dim lngVBComs As Long, arr_varVBCom() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim blnFound As Boolean, blnAdd As Boolean, blnAddAll As Boolean
  Dim intPos01 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, lngTmp01 As Long
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varNum().
  Const N_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const N_NUM  As Integer = 0
  Const N_SORT As Integer = 1
  Const N_CNT  As Integer = 2

  ' ** Array: arr_varVBCom().
  Const V_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const V_VID  As Integer = 0
  Const V_VNAM As Integer = 1
  Const V_LINS As Integer = 2
  Const V_MAX  As Integer = 3

  ' ** Array: arr_varDel().
  Const D_ELEMS As Integer = 0  ' ** Array's first-element UBound().
  Const D_VID As Integer = 0

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngNums = 0&
  ReDim arr_varNum(N_ELEMS, 0)

  lngVBComs = 0&
  ReDim arr_varVBCom(V_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        strMax = vbNullString
        If strModName <> "zz_mod_ModuleFormatFuncs" And strModName <> "modGlobConst" Then

          ' ** Note: These modules may have their error handlers all numbered the same:
          ' **   frmJournal_Columns_Sub
          ' **   frmStatementParameters
          ' **   modVersionConvertFuncs1

          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngX = lngDecLines To lngLines
              strLine = Trim(.Lines(lngX, 1))
              If strLine <> vbNullString Then
                If Left(strLine, 1) <> "." Then
                  intPos01 = InStr(strLine, " ")
                  If intPos01 > 0 Then
                    strTmp01 = Trim(Left(strLine, intPos01))
                    If IsNumeric(strTmp01) = True Then
                      If Right(strTmp01, 1) = "0" Then  ' ** All line numbers are by tens.
                        blnFound = False
                        If Val(strTmp01) > Val(strMax) Then
                          ' ** Since some modules have identical error handler line numbers,
                          ' ** make sure we're getting the biggest number in the module.
                          strMax = strTmp01
                        End If
                        For lngY = 0& To (lngNums - 1&)
                          If arr_varNum(N_NUM, lngY) = strTmp01 Then
                            blnFound = True
                            arr_varNum(N_CNT, lngY) = arr_varNum(N_CNT, lngY) + 1&
                            Exit For
                          End If
                        Next  ' ** lngY.
                        If blnFound = False Then
                          lngNums = lngNums + 1&
                          lngE = lngNums - 1&
                          ReDim Preserve arr_varNum(N_ELEMS, lngE)
                          arr_varNum(N_NUM, lngE) = strTmp01
                          arr_varNum(N_SORT, lngE) = Right(String(5, "0") & strTmp01, 5)  '76890
                          arr_varNum(N_CNT, lngE) = CLng(1)
                        End If
                      Else
                        Debug.Print "'" & strTmp01 & "  " & strModName & "  LINE: " & CStr(lngX)
                        DoEvents
                      End If
                    End If
                  End If
                End If  ' ** Remark.
              End If  ' ** vbNullString.
            Next  ' ** lngX.
          End With  ' ** cod.
          Set cod = Nothing

          lngVBComs = lngVBComs + 1&
          lngE = lngVBComs - 1&
          ReDim Preserve arr_varVBCom(V_ELEMS, lngE)
          arr_varVBCom(V_VID, lngE) = CLng(0)
          arr_varVBCom(V_VNAM, lngE) = strModName
          arr_varVBCom(V_LINS, lngE) = lngLines
          arr_varVBCom(V_MAX, lngE) = strMax

        End If  ' ** strModName.
      End With  ' ** vbc.
    Next  ' ** vbc.
    Set vbc = Nothing
  End With  ' ** vbp.
  Set vbp = Nothing

  Debug.Print "'NUMS: " & CStr(lngNums)
  DoEvents

  If lngNums > 0& Then

    ' ** Binary Sort arr_varNum() array by sort order.
    For lngX = UBound(arr_varNum, 2) To 1 Step -1
      For lngY = 0 To (lngX - 1)
        If arr_varNum(N_SORT, lngY) > arr_varNum(N_SORT, (lngY + 1)) Then
          For lngZ = 0& To N_ELEMS
            varTmp00 = arr_varNum(lngZ, lngY)
            arr_varNum(lngZ, lngY) = arr_varNum(lngZ, (lngY + 1))
            arr_varNum(lngZ, (lngY + 1)) = varTmp00
            varTmp00 = Empty
          Next  ' ** lngZ.
        End If
      Next  ' ** lngY.
    Next  ' ** lngX.

    Set dbs = CurrentDb
    With dbs

      ' ** Empty tblVBComponent_CodeNum.
      Set qdf = .QueryDefs("qryVBComponent_CodeNum_01")
      qdf.Execute
      Set qdf = Nothing

      Set rst1 = .OpenRecordset("tblVBComponent_CodeNum", dbOpenDynaset, dbAppendOnly)
      With rst1
        For lngX = 0& To (lngNums - 1&)
          .AddNew
          ![dbs_id] = lngThisDbsID
          ' ** ![vbcn_id] : AutoNumber.
          ![vbcn_code] = arr_varNum(N_NUM, lngX)
          ![vbcn_sort] = arr_varNum(N_SORT, lngX)
          ![vbcn_cnt] = arr_varNum(N_CNT, lngX)
          ![vbcn_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      Set rst1 = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      With rst1
        .MoveFirst
        For lngX = 0& To (lngVBComs - 1&)
          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & arr_varVBCom(V_VNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varVBCom(V_VID, lngX) = ![vbcom_id]
          Else
            Stop
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      blnAddAll = False: blnAdd = False
      Set rst1 = .OpenRecordset("tblVBComponent_CodeNum_Max", dbOpenDynaset, dbConsistent)
      With rst1
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        Else
          .MoveFirst
        End If
        For lngX = 0& To (lngVBComs - 1&)
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varVBCom(V_VID, lngX))
            If .NoMatch = True Then
              blnAdd = True
            End If
          End Select
          Select Case blnAdd
          Case True
            .AddNew
            ![dbs_id] = lngThisDbsID
            ![vbcom_id] = arr_varVBCom(V_VID, lngX)
            ' ** ![vbcnm_id] : AutoNumber.
            ![vbcom_name] = arr_varVBCom(V_VNAM, lngX)
            ![vbcom_lines] = arr_varVBCom(V_LINS, lngX)
            ![vbcnm_max] = arr_varVBCom(V_MAX, lngX)
            ![vbcnm_sort] = Right(String(5, "0") & arr_varVBCom(V_MAX, lngX), 5)
            ![vbcnm_datemodified] = Now()
            .Update
          Case False
            ' ** ![dbs_id]
            ' ** ![vbcom_id]
            ' ** ![vbcnm_id]
            ' ** ![vbcom_name]
            If arr_varVBCom(V_LINS, lngX) <> ![vbcom_lines] Then
              .Edit
              ![vbcom_lines] = arr_varVBCom(V_LINS, lngX)
              ![vbcnm_datemodified] = Now()
              .Update
            End If
            If arr_varVBCom(V_MAX, lngX) <> ![vbcnm_max] Then
              .Edit
              ![vbcnm_max] = arr_varVBCom(V_MAX, lngX)
              ![vbcnm_datemodified] = Now()
              .Update
            End If
            strTmp01 = Right(String(5, "0") & arr_varVBCom(V_MAX, lngX), 5)
            If strTmp01 <> ![vbcnm_sort] Then
              .Edit
              ![vbcnm_sort] = strTmp01
              ![vbcnm_datemodified] = Now()
              .Update
            End If
          End Select
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      ' ** Delete qryVBComponent_CodeNum_04 (qryVBComponent_CodeNum_03 (tblVBComponent_CodeNum_Max,
      ' ** grouped, with cnt > 1, Min(vbcnm_datemodified)), linked back to tblVBComponent_CodeNum_Max).
      Set qdf = .QueryDefs("qryVBComponent_CodeNum_05")
      qdf.Execute

      lngDels = 0&
      ReDim arr_varDel(D_ELEMS, 0)

      Set rst1 = .OpenRecordset("tblVBComponent_CodeNum_Max", dbOpenDynaset, dbReadOnly)
      Set rst2 = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      With rst1
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          rst2.FindFirst "[dbs_id] = " & CStr(![dbs_id]) & " And [vbcom_id] = " & CStr(![vbcom_id])
          If .NoMatch = True Then
            lngDels = lngDels + 1
            lngE = lngDels - 1&
            ReDim Preserve arr_varDel(D_ELEMS, lngE)
            arr_varDel(D_VID, lngE) = ![vbcom_id]
          End If
          If lngX < lngRecs Then .MoveNext
        Next  ' ** lngX
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing
      rst2.Close
      Set rst2 = Nothing

      Debug.Print "'DELS: " & CStr(lngDels)
      DoEvents

      If lngDels > 0& Then
        lngTmp01 = 0&
        For lngX = 0& To (lngDels - 1&)
          ' ** Delete tblVBComponent_CodeNum_Max, by specified [vbcomid].
          Set qdf = .QueryDefs("qryVBComponent_CodeNum_06")
          With qdf.Parameters
            ![vbcomid] = arr_varDel(D_VID, lngX)
          End With
          qdf.Execute
          Set qdf = Nothing
          lngTmp01 = lngTmp01 + 1&
        Next
        Debug.Print "'RECS DELETED: " & CStr(lngTmp01)
        DoEvents
      End If  ' ** lngDels.

      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

  End If  ' ** lngNums.

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_CodeLineNum_Doc = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_CodeLineNum_Add() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_CodeLineNum_Add"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strModName As String, strProcName As String, strLine As String, strLastLineNum As String
  Dim lngLines As Long, lngDecLines As Long, lngEdits As Long
  Dim blnFound As Boolean, blnLineContOn As Boolean, blnIsLineCont As Boolean, blnCase As Boolean
  Dim lngPos01 As Long, lngLen01 As Integer, lngLen02 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  strModName = "Form_frmMasterBalance"
  strProcName = "ShowAcctMast_Win"

  ' *******************************************
  ' ** THIS ASSUMES CODE IS ALREADY INDENTED!
  ' *******************************************

'CHECK EXISTING NUMBERS AGAINST NEW ONES!

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    Set vbc = .VBComponents(strModName)
    With vbc
      Set cod = .CodeModule
      With cod
        lngLines = .CountOfLines
        lngDecLines = .CountOfDeclarationLines
        strLastLineNum = "90"  ' ** For this, we'll always start at 100.
        blnLineContOn = False: blnIsLineCont = False: blnCase = False: blnFound = False
        lngEdits = 0&
        For lngX = lngDecLines To lngLines
          If .ProcOfLine(lngX, vbext_pk_Proc) = strProcName Then
            strLine = .Lines(lngX, 1)  ' ** Don't trim!
            strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
            If Trim(strLine) <> vbNullString Then
              If Left(Trim(strLine), 1) <> "'" Then
                If Left(Trim(strLine), 4) = "Dim " Or Left(Trim(strLine), 6) = "Const " Or Left(Trim(strLine), 7) = "Static " Then
                  ' ** These don't get numbered.
                Else
                  ' ** Code starts at position 9, indent 8 spaces.
                  ' ** On Error's start at position 7, indent 6 spaces.
                  lngLen01 = Len(strLine)  ' ** Don't lose indent!
                  lngPos01 = 0&: lngLen02 = 0&
                  Select Case blnLineContOn
                  Case True
                    ' ** Line continuation, so no number.
                    blnIsLineCont = True
                    If Right(strLine, 1) = "_" Then
                      ' ** More continuing lines, so blnLineContOn remains True.
                    Else
                      ' ** This is the last line of a continuation, but blnIsLineCont remains True for this line.
                      blnLineContOn = False
                    End If
                  Case False
                    If Right(strLine, 1) = "_" Then
                      ' ** First line having continuation, so it gets a number.
                      blnLineContOn = True
                    End If
                  End Select
                  Select Case blnCase
                  Case True
                    ' ** It should only get here until the first Case has been found.
                    If blnFound = False Then
                      If Left(Trim(strLine), 5) = "Case " Then
                        ' ** First Case statement, so no number, and both are True.
                        blnFound = True
                      End If
                    End If
                  Case False
                    If Left(Trim(strLine), 11) = "Select Case" Then
                      ' ** First line of Select Case block, so it gets a number.
                      blnCase = True
                      blnFound = False
                    End If
                  End Select  ' ** blnCase.
                  If blnIsLineCont = False Then
                    If blnCase = True And blnFound = True Then
                      ' ** This is the first Case statement, so no number.
                      blnCase = False
                      blnFound = False
                    Else
                      If Left(strLine, 1) = " " Then
                        For lngY = 1& To lngLen01
                          If Mid(strLine, lngY, 1) = Chr(32) Then
                            strTmp01 = strTmp01 & Mid(strLine, lngY, 1)
                          Else
                            ' ** Start of text.
                            lngPos01 = lngY  ' ** {may be used to check replacement}
                            strTmp02 = Mid(strLine, lngY)
                            Exit For
                          End If
                        Next  ' ** lngY.
                        If strTmp01 <> vbNullString And strTmp02 <> vbNullString Then
                          strTmp03 = CStr(Val(strLastLineNum) + 10)
                          lngLen02 = Len(strTmp03)
                          If Len(strTmp01) > lngLen02 Then
                            strTmp01 = strTmp03 & Mid(strTmp01, (lngLen02 + 1&))
                            .ReplaceLine lngX, (strTmp01 & strTmp02)
                            strLastLineNum = strTmp03
                            lngEdits = lngEdits + 1&
                          Else
                            Stop
                          End If
                        End If
                      ElseIf Left(strLine, 15) = "Public Function" Or Left(strLine, 16) = "Private Function" Or _
                          Left(strLine, 10) = "Public Sub" Or Left(strLine, 11) = "Private Sub" Then
                        ' ** First line of designated procedure.
                      ElseIf Left(strLine, 1) = "#" Then
                        ' ** Conditional Compiler constant or structure.
                        ' ** No line numbers.
                      ElseIf strLine = "EXITP:" Or strLine = "ERRH:" Then
                        ' ** No line numbers.
                      ElseIf strLine = "End Function" Or strLine = "End Sub" Then
                        ' ** End of procedure.
                        Exit For
                      Else
                        lngPos01 = InStr(strLine, " ")
                        If lngPos01 > 0 Then
                          strTmp01 = Trim(Left(strLine, lngPos01))
                          If IsNumeric(strTmp01) = True Then
                            ' ** Existing line number.
                          Else
                            Debug.Print "'" & strLine
                            DoEvents
                            Stop
                          End If
                        Else
                          Debug.Print "'" & strLine
                          DoEvents
                          Stop
                        End If
                      End If
                    End If  ' ** blnCase.
                  Else
                    ' ** Reset for next line.
                    blnIsLineCont = False
                  End If  ' ** blnIsLineCont.
                End If  ' ** Dim, Const, Static.
              End If  ' ** Remark.
            End If  ' ** vbNullString.
          End If  ' ** strProcName
        Next  ' ** lngX.

        Debug.Print "'LINES NUMBERED: " & CStr(lngEdits)
        DoEvents

      End With  ' ** cod.
    End With  ' ** vbc.
  End With  ' ** vbp.

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  VBA_CodeLineNum_Add = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_Var_LocalDoc() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_Var_LocalDoc"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim lngVars As Long, arr_varVar() As Variant
  Dim lngThisDbsID As Long, lngVBComID As Long, lngVBComProcID As Long, lngComType As Long
  Dim blnHasCompDir As Boolean, strCompDir1 As String, strCompDir2 As String
  Dim lngCompDirType1 As Long, lngCompDirOptType1 As Long, lngCompDirType2 As Long, lngCompDirOptType2 As Long
  Dim lngLines As Long, lngDecLines As Long, lngRecs As Long
  Dim strModName As String, strProcName As String, strLine As String, strDecType As String
  Dim strLastModName As String, strLastProcName As String, strProcType As String, strProcSubType As String
  Dim blnContinue As Boolean, blnFound As Boolean, blnAddAll As Boolean, blnAdd As Boolean
  Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intCnt As Integer, intMode As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varVar().
  Const V_ELEMS As Integer = 19  ' ** Array's first element UBound().
  Const V_DID   As Integer = 0
  Const V_VID   As Integer = 1
  Const V_VNAM  As Integer = 2
  Const V_PID   As Integer = 3
  Const V_PNAM  As Integer = 4
  Const V_VAR   As Integer = 5
  Const V_VTYP  As Integer = 6
  Const V_VTYP2 As Integer = 7
  Const V_SCOP  As Integer = 8
  Const V_SCOP2 As Integer = 9
  Const V_DTYP  As Integer = 10
  Const V_DTYP2 As Integer = 11
  Const V_VAL   As Integer = 12
  Const V_LINE  As Integer = 13
  Const V_PTYP  As Integer = 14
  Const V_PTYP2 As Integer = 15
  Const V_PSTP  As Integer = 16
  Const V_PSTP2 As Integer = 17
  Const V_CTYP  As Integer = 18
  Const V_CDIR  As Integer = 19

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents


  'intMode = 1  ' ** Local Constants.
  'intMode = 2  ' ** Local Variables.
  intMode = 3  ' ** Static Variables.

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngVars = 0&
  ReDim arr_varVar(V_ELEMS, 0)

  Select Case intMode
  Case 1
    Debug.Print "'LOCAL CONSTANTS..."
  Case 2
    Debug.Print "'LOCAL VARIABLES..."
  Case 3
    Debug.Print "'LOCAL STATIC..."
  End Select
  DoEvents

  Debug.Print "'|";
  DoEvents

  lngZ = 0&
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngRecs = .VBComponents.Count
    For Each vbc In .VBComponents

      lngZ = lngZ + 1&
      With vbc
        strModName = .Name
        If strModName <> "modGlobConst" Then
          Set cod = .CodeModule
          With cod

            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines

            blnHasCompDir = False
            lngCompDirType1 = 0&: lngCompDirOptType1 = 0&: lngCompDirType2 = 0&: lngCompDirOptType2 = 0&
            strCompDir1 = vbNullString: strCompDir2 = vbNullString
            strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString: strTmp05 = vbNullString
            ' ** See if there are any Compiler Directives in this module.
            ' ** Current maximum is 2 in one module.
            For lngX = 1& To lngDecLines
              strLine = Trim(.Lines(lngX, 1))
              If Left(strLine, 1) = "#" Then
                If Left(strLine, 6) = "#Const" Then
                  blnHasCompDir = True
                  strTmp01 = GetNthWord(strLine, 2)  ' ** Module Function: modStringFuncs.
                  Select Case strTmp01
                  Case "IsDev"
                    ' ** #Const IsDev = 0
                    If lngCompDirType1 = 0& Then
                      lngCompDirType1 = 1&
                    Else
                      lngCompDirType2 = 1&
                    End If
                  Case "IsDemo"
                    ' ** #Const IsDemo = 0
                    If lngCompDirType1 = 0& Then
                      lngCompDirType1 = 2&
                    Else
                      lngCompDirType2 = 2&
                    End If
                  Case "NoExcel"
                    ' ** #Const NoExcel = 0
                    If lngCompDirType1 = 0& Then
                      lngCompDirType1 = 3&
                    Else
                      lngCompDirType2 = 3&
                    End If
                  Case "HasRepost"
                    ' ** #Const HasRepost = 0
                    If lngCompDirType1 = 0& Then
                      lngCompDirType1 = 4#
                    Else
                      lngCompDirType2 = 4&
                    End If
                  End Select  ' ** strTmp01.

                End If
              End If
            Next  ' ** lngX.
            strTmp01 = vbNullString

            For lngX = (lngDecLines + 1&) To lngLines
              strLine = Trim(.Lines(lngX, 1))
              If strLine <> vbNullString Then
                If Left(strLine, 1) <> "'" Then

                  strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString: strTmp05 = vbNullString

                  ' ** Interpret Compiler Directives.
                  ' ** I DID NOT EXPECT THIS MUCH CODE!
                  If blnHasCompDir = True Then
                    If Left(strLine, 1) = "#" Then
                      If Left(strLine, 3) = "#If" Then
                        ' ** Start of Compiler Directive section, e.g., '#If NoExcel Then'.
                        ' ** I'm assuming we've not used 'Not' with these.
                        strTmp03 = GetNthWord(strLine, 2)  ' ** Module Function: mod StringFuncs.
                        Select Case strTmp03
                        Case "IsDev"
                          If lngCompDirType1 = 1& Then
                            strCompDir1 = "OPEN 1.1 " & CStr(lngX)
                            lngCompDirOptType1 = 1&
                          ElseIf lngCompDirType2 = 1& Then
                            strCompDir2 = "OPEN 1.2 " & CStr(lngX)
                            lngCompDirOptType2 = 1&
                          End If
                        Case "IsDemo"
                          If lngCompDirType1 = 2& Then
                            strCompDir1 = "OPEN 2.1 " & CStr(lngX)
                            lngCompDirOptType1 = 3&
                          ElseIf lngCompDirType2 = 2& Then
                            strCompDir2 = "OPEN 2.2 " & CStr(lngX)
                            lngCompDirOptType2 = 3&
                          End If
                        Case "NoExcel"
                          If lngCompDirType1 = 3& Then
                            strCompDir1 = "OPEN 3.1 " & CStr(lngX)
                            lngCompDirOptType1 = 5&
                          ElseIf lngCompDirType2 = 3& Then
                            strCompDir2 = "OPEN 3.2 " & CStr(lngX)
                            lngCompDirOptType2 = 5&
                          End If
                        Case "HasRepost"
                          If lngCompDirType1 = 4& Then
                            strCompDir1 = "OPEN 4.1 " & CStr(lngX)
                            lngCompDirOptType1 = 7&
                          ElseIf lngCompDirType2 = 4& Then
                            strCompDir2 = "OPEN 4.2 " & CStr(lngX)
                            lngCompDirOptType2 = 7&
                          End If
                        End Select
                      ElseIf Left(strLine, 5) = "#Else" Then
                        ' ** Flip the option to False.
                        If strCompDir1 <> vbNullString And strCompDir2 <> vbNullString Then
                          If Left(strCompDir1, 4) = "OPEN" And Left(strCompDir2, 4) = "OPEN" Then
                            ' ** Both have opened, so figure out which is the most recent.
                            strTmp01 = GetLastWord(strCompDir1)  ' ** Module Function: modStringFuncs.
                            strTmp02 = GetLastWord(strCompDir2)  ' ** Module Function: modStringFuncs.
                            If Val(strTmp01) < Val(strTmp02) Then
                              ' ** 2 is flipping.
                              strCompDir2 = "FLIP" & Mid(strCompDir2, 5)
                              Select Case Mid(strCompDir2, 6, 1)
                              Case "1"
                                lngCompDirOptType2 = 2&
                              Case "2"
                                lngCompDirOptType2 = 4&
                              Case "3"
                                lngCompDirOptType2 = 6&
                              Case "4"
                                lngCompDirOptType2 = 8&
                              End Select
                              strCompDir2 = Left(strCompDir2, (Len(strCompDir2) - Len(strTmp02)))  ' ** Strip opening line number.
                              strCompDir2 = strCompDir2 & CStr(lngX)  ' ** Add flipping line number.  (Would you please add the flipping line number!)
                            Else
                              ' ** 1 is flipping.
                              strCompDir1 = "FLIP" & Mid(strCompDir1, 5)
                              Select Case Mid(strCompDir1, 6, 1)
                              Case "1"
                                lngCompDirOptType1 = 2&
                              Case "2"
                                lngCompDirOptType1 = 4&
                              Case "3"
                                lngCompDirOptType1 = 6&
                              Case "4"
                                lngCompDirOptType1 = 8&
                              End Select
                              strCompDir1 = Left(strCompDir1, (Len(strCompDir1) - Len(strTmp01)))  ' ** Strip opening line number.
                              strCompDir1 = strCompDir1 & CStr(lngX)  ' ** Add flipping line number.
                            End If
                          ElseIf (Left(strCompDir1, 4) = "OPEN" And Left(strCompDir2, 4) = "FLIP") Then
                            ' ** 1 is flipping.
                            strCompDir1 = "FLIP" & Mid(strCompDir1, 5)
                            Select Case Mid(strCompDir1, 6, 1)
                            Case "1"
                              lngCompDirOptType1 = 2&
                            Case "2"
                              lngCompDirOptType1 = 4&
                            Case "3"
                              lngCompDirOptType1 = 6&
                            Case "4"
                              lngCompDirOptType1 = 8&
                            End Select
                            strCompDir1 = Left(strCompDir1, (Len(strCompDir1) - Len(strTmp01)))  ' ** Strip opening line number.
                            strCompDir1 = strCompDir1 & CStr(lngX)  ' ** Add flipping line number.
                          ElseIf (Left(strCompDir1, 4) = "FLIP" And Left(strCompDir2, 4) = "OPEN") Then
                            ' ** 2 is flipping.
                            strCompDir2 = "FLIP" & Mid(strCompDir2, 5)
                            Select Case Mid(strCompDir2, 6, 1)
                            Case "1"
                              lngCompDirOptType2 = 2&
                            Case "2"
                              lngCompDirOptType2 = 4&
                            Case "3"
                              lngCompDirOptType2 = 6&
                            Case "4"
                              lngCompDirOptType2 = 8&
                            End Select
                            strCompDir2 = Left(strCompDir2, (Len(strCompDir2) - Len(strTmp02)))  ' ** Strip opening line number.
                            strCompDir2 = strCompDir2 & CStr(lngX)  ' ** Add flipping line number.
                          Else
                            ' ** If they've both already flipped, this shouldn't be here!
                            Stop
                          End If
                        Else
                          ' ** Only 1 directive in play.
                          If strCompDir1 <> vbNullString Then
                            If Left(strCompDir1, 4) = "OPEN" Then
                              strTmp01 = GetLastWord(strCompDir1)  ' ** Module Function: modStringFuncs.
                              strCompDir1 = "FLIP" & Mid(strCompDir1, 5)
                              Select Case Mid(strCompDir1, 6, 1)
                              Case "1"
                                lngCompDirOptType1 = 2&
                              Case "2"
                                lngCompDirOptType1 = 4&
                              Case "3"
                                lngCompDirOptType1 = 6&
                              Case "4"
                                lngCompDirOptType1 = 8&
                              End Select
                              strCompDir1 = Left(strCompDir1, (Len(strCompDir1) - Len(strTmp01)))  ' ** Strip opening line number.
                              strCompDir1 = strCompDir1 & CStr(lngX)  ' ** Add flipping line number.
                            Else
                              Stop
                            End If
                          ElseIf strCompDir2 <> vbNullString Then
                            If Left(strCompDir2, 4) = "OPEN" Then
                              strTmp02 = GetLastWord(strCompDir2)  ' ** Module Function: modStringFuncs.
                              strCompDir2 = "FLIP" & Mid(strCompDir2, 5)
                              Select Case Mid(strCompDir2, 6, 1)
                              Case "1"
                                lngCompDirOptType1 = 2&
                              Case "2"
                                lngCompDirOptType1 = 4&
                              Case "3"
                                lngCompDirOptType1 = 6&
                              Case "4"
                                lngCompDirOptType1 = 8&
                              End Select
                              strCompDir2 = Left(strCompDir2, (Len(strCompDir2) - Len(strTmp02)))  ' ** Strip opening line number.
                              strCompDir2 = strCompDir2 & CStr(lngX)  ' ** Add flipping line number.
                            Else
                              Stop
                            End If
                          End If
                        End If
                      ElseIf Left(strLine, 7) = "#ElseIf" Then
                        ' ** None of these in use at the moment.
                      ElseIf Left(strLine, 5) = "#End If" Then
                        ' ** Figure out who's closing.
                        If strCompDir1 <> vbNullString And strCompDir2 <> vbNullString Then
                          ' ** Both are still open.
                          strTmp01 = GetLastWord(strCompDir1)  ' ** Module Function: modStringFuncs.
                          strTmp02 = GetLastWord(strCompDir2)  ' ** Module Function: modStringFuncs.
                          ' ** It doesn't matter if they're OPEN or FLIP.
                          ' ** A block can't open outside then close inside another block.
                          If Val(strTmp01) < Val(strTmp02) Then
                            ' ** 2 is closing.
                            lngCompDirOptType2 = 0&
                            strCompDir2 = vbNullString
                          Else
                            ' ** 1 is closing.
                            lngCompDirOptType1 = 0&
                            strCompDir1 = vbNullString
                          End If
                        ElseIf strCompDir2 <> vbNullString Then
                          ' ** Only 2 is open.
                          lngCompDirOptType2 = 0&
                          strCompDir2 = vbNullString
                        ElseIf strCompDir1 <> vbNullString Then
                          ' ** Only 1 is open.
                          lngCompDirOptType1 = 0&
                          strCompDir1 = vbNullString
                        End If
                      End If
                    End If
                    strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
                  End If  ' ** blnHasCompDir.

                  intPos01 = InStr(strLine, " ")
                  If intPos01 > 0 Then  ' ** Not looking for anything that would be alone on the line.
                    strTmp01 = Trim(Left(strLine, intPos01))
                    If IsNumeric(strTmp01) = False Then  ' ** Dimension statements are never numbered.

                      blnContinue = False: strDecType = vbNullString
                      Select Case intMode
                      Case 1
                        If strTmp01 = "Const" Then
                          blnContinue = True
                          strDecType = "Constant"
                        End If
                      Case 2
                        If strTmp01 = "Dim" Then
                          blnContinue = True
                          strDecType = "Variable"
                        End If
                      Case 3
                        If strTmp01 = "Static" Then
                          blnContinue = True
                          strDecType = "Variable"
                        End If
                      End Select
                      If blnContinue = True Then

                        strTmp01 = Trim(Mid(strLine, intPos01))  ' ** Line without 'Dim' word.
                        strTmp02 = vbNullString: strTmp03 = vbNullString
                        intPos01 = InStr(strTmp01, "'")
                        If intPos01 > 0 Then
                          ' ** This would be a remark, unless within the quotes of a constant assignment.
                          intPos02 = InStr(strTmp01, Chr(34))
                          If intPos02 > 0 And intPos02 < intPos01 Then
                            'Stop
                            If InStr(strLine, "MY_CRLF") > 0 Then
                              strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
                            Else
                              'Stop
                              strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
                            End If
                          Else
                            ' ** Just get rid of it.
                            strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))  ' ** If this is something bizarre, it'll error eventually!
                          End If
                        End If

                        If intMode = 1 Then
                          ' ** Constant, which are always 1 per line.
                          intCnt = 1
                        Else
                          intCnt = (CharCnt(strTmp01, ",") + 1)  ' ** Module Function: modStringFuncs.
                          ' ** Unless it's an array with sizing numbers, of which there are very few.
                        End If

                        strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                        strProcType = vbNullString: strProcSubType = vbNullString
                        For lngY = (lngX - 1&) To lngDecLines Step -1&
                          strTmp04 = .Lines(lngY, 1)
                          If Left(strTmp04, 7) = "Private" Or Left(strTmp04, 6) = "Public" Then
                            ' ** All the procedures should be scoped.
                            If InStr(strTmp04, " Sub ") > 0 Then
                              strProcType = "Sub"
                              Exit For
                            ElseIf InStr(strTmp04, " Function ") > 0 Then
                              strProcType = "Function"
                              Exit For
                            ElseIf InStr(strTmp04, " Property ") > 0 Then
                              strProcType = "Property"
                              If InStr(strTmp04, " Get ") > 0 Then
                                strProcSubType = "Get"
                                Exit For
                              ElseIf InStr(strTmp04, " Let ") > 0 Then
                                strProcSubType = "Let"
                                Exit For
                              ElseIf InStr(strTmp04, " Set ") > 0 Then
                                strProcSubType = "Set"
                                Exit For
                              End If
                            End If
                          End If
                        Next  ' ** lngY.
                        If strProcType = vbNullString Then
                          Stop
                        End If

                        If intCnt = 1 Then
                          ' ** Only one item on the line.
                          intPos02 = InStr(strTmp01, " as ")
                          If intPos02 > 0 Then
                            strTmp02 = Trim(Mid(strTmp01, (intPos02 + 3)))  ' ** This may be its type, alone.
                            strTmp01 = Trim(Left(strTmp01, intPos02))
                            intPos03 = InStr(strTmp02, " ")
                            If intPos03 > 0 Then
                              If intMode = 1 Then
                                strTmp03 = Trim(Mid(strTmp02, intPos03))   ' ** This should be its assignment, with equal sign.
                                strTmp02 = Trim(Left(strTmp02, intPos03))  ' ** This should be its type.
                                If Left(strTmp03, 1) = "=" Then
                                  strTmp03 = Trim(Mid(strTmp03, 2))  ' ** Remove equal sign.
                                  If strTmp02 = "String" Then
                                    intPos03 = CharPos(strTmp03, 2, Chr(34))  ' ** Module Function: modStringFuncs.
                                    If intPos03 > 0 Then
                                      intPos03 = InStr(intPos03, strTmp03, "'")
                                      If intPos03 > 0 Then
                                        ' ** This really should be a remark.
                                        strTmp03 = Trim(Left(strTmp03, (intPos03 - 1)))
                                      End If
                                    Else
                                      Stop
                                    End If
                                  Else
                                    intPos03 = InStr(strTmp03, "'")
                                    If intPos03 > 0 Then
                                      ' ** This really should be a remark.
                                      strTmp03 = Trim(Left(strTmp03, (intPos03 - 1)))
                                    End If
                                  End If
                                Else
                                  Stop
                                End If
                              Else
                                ' ** Since this should be the only item on the line, we can throw out anything after it.
                                ' ** But maybe it's got the 'New' object word.
                                If Left(strTmp02, 4) = "New " Then
                                  intPos03 = CharPos(strTmp02, 2, " ")  ' ** Module Function: modStringFuncs.
                                  If intPos03 = 0 Then
                                    ' ** Keep the whole thing.
                                  Else
                                    strTmp02 = Trim(Left(strTmp02, intPos03))  ' ** This should be the type, with the 'New' keyword.
                                  End If
                                Else
                                  strTmp02 = Trim(Left(strTmp02, intPos03))  ' ** This should now be its type, alone.
                                End If
                              End If
                            End If  ' ** intPos03.
                            If strTmp02 = "String * 255" Then strTmp02 = "String"
                            If Left(strTmp02, 4) = "New " And Len(strTmp02) > 4 Then strTmp02 = Mid(strTmp02, 5)
                            lngVars = lngVars + 1&
                            lngE = lngVars - 1&
                            ReDim Preserve arr_varVar(V_ELEMS, lngE)
                            arr_varVar(V_DID, lngE) = lngThisDbsID
                            arr_varVar(V_VID, lngE) = Null
                            arr_varVar(V_VNAM, lngE) = strModName
                            arr_varVar(V_PID, lngE) = Null
                            arr_varVar(V_PNAM, lngE) = strProcName
                            arr_varVar(V_VAR, lngE) = strTmp01
                            arr_varVar(V_VTYP, lngE) = strTmp02
                            arr_varVar(V_VTYP2, lngE) = Null
                            Select Case intMode
                            Case 1, 2
                              arr_varVar(V_SCOP, lngE) = "Local"
                            Case 3
                              arr_varVar(V_SCOP, lngE) = "Static"
                            End Select
                            arr_varVar(V_DTYP, lngE) = strDecType
                            arr_varVar(V_VAL, lngE) = strTmp03
                            arr_varVar(V_LINE, lngE) = lngX
                            arr_varVar(V_PTYP, lngE) = strProcType
                            arr_varVar(V_PTYP2, lngE) = Null
                            If strProcSubType <> vbNullString Then
                              arr_varVar(V_PSTP, lngE) = strProcSubType
                            Else
                              arr_varVar(V_PSTP, lngE) = Null
                            End If
                            arr_varVar(V_PSTP2, lngE) = Null
                            arr_varVar(V_CTYP, lngE) = Null
                            Select Case blnHasCompDir
                            Case True
                              ' ** OK, what are we inside of?
                              If lngCompDirOptType1 = 0& And lngCompDirOptType2 = 0& Then
                                ' ** Nothing is open.
                                arr_varVar(V_CDIR, lngE) = 0&
                              ElseIf lngCompDirOptType1 > 0& And lngCompDirOptType2 = 0& Then
                                ' ** Only 1 is open.
                                arr_varVar(V_CDIR, lngE) = lngCompDirOptType1
                              ElseIf lngCompDirOptType1 = 0& And lngCompDirOptType2 > 0& Then
                                ' ** Only 2 is open.
                                arr_varVar(V_CDIR, lngE) = lngCompDirOptType2
                              Else
                                ' ** Both are open.
                                strTmp04 = GetLastWord(strCompDir1)  ' ** Module Function: modStringFuncs.
                                strTmp05 = GetLastWord(strCompDir2)  ' ** Module Function: modStringFuncs.
                                If Val(strTmp04) < Val(strTmp05) Then
                                  ' ** 2 has control.
                                  arr_varVar(V_CDIR, lngE) = lngCompDirOptType2
                                Else
                                  ' ** 1 has control.
                                  arr_varVar(V_CDIR, lngE) = lngCompDirOptType1
                                End If
                              End If
                            Case False
                              arr_varVar(V_CDIR, lngE) = 0&
                            End Select
                          Else
                            ' ** Not typed?
                            Debug.Print "'" & strLine
                            Stop
                          End If
                        Else
                          ' ** Multiple items on the line.
                          intPos01 = InStr(strTmp01, ",")
                          blnFound = True
                          Do While blnFound = True
                            blnFound = False
                            strTmp02 = Left(strTmp01, (intPos01 - 1))  ' ** Just this item.
                            strTmp01 = Trim(Mid(strTmp01, (intPos01 + 1)))  ' ** The rest of the line.
                            intPos02 = InStr(strTmp02, " as ")
                            If intPos02 > 0 Then
                              strTmp03 = Trim(Mid(strTmp02, (intPos02 + 3)))  ' ** This should just be the type.
                              strTmp02 = Trim(Left(strTmp02, intPos02))       ' ** And this should be the defined item.
                              strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                              If strTmp03 = "String * 255" Then strTmp03 = "String"
                              If Left(strTmp03, 4) = "New " And Len(strTmp03) > 4 Then strTmp03 = Mid(strTmp03, 5)
                              lngVars = lngVars + 1&
                              lngE = lngVars - 1&
                              ReDim Preserve arr_varVar(V_ELEMS, lngE)
                              arr_varVar(V_DID, lngE) = lngThisDbsID
                              arr_varVar(V_VID, lngE) = Null
                              arr_varVar(V_VNAM, lngE) = strModName
                              arr_varVar(V_PID, lngE) = Null
                              arr_varVar(V_PNAM, lngE) = strProcName
                              arr_varVar(V_VAR, lngE) = strTmp02
                              arr_varVar(V_VTYP, lngE) = strTmp03
                              arr_varVar(V_VTYP2, lngE) = Null
                              Select Case intMode
                              Case 1, 2
                                arr_varVar(V_SCOP, lngE) = "Local"
                              Case 3
                                arr_varVar(V_SCOP, lngE) = "Static"
                              End Select
                              arr_varVar(V_DTYP, lngE) = strDecType
                              arr_varVar(V_VAL, lngE) = Null
                              arr_varVar(V_LINE, lngE) = lngX
                              arr_varVar(V_PTYP, lngE) = strProcType
                              arr_varVar(V_PTYP2, lngE) = Null
                              If strProcSubType <> vbNullString Then
                                arr_varVar(V_PSTP, lngE) = strProcSubType
                              Else
                                arr_varVar(V_PSTP, lngE) = Null
                              End If
                              arr_varVar(V_PSTP2, lngE) = Null
                              arr_varVar(V_CTYP, lngE) = Null
                              Select Case blnHasCompDir
                              Case True
                                ' ** OK, what are we inside of?
                                If lngCompDirOptType1 = 0& And lngCompDirOptType2 = 0& Then
                                  ' ** Nothing is open.
                                  arr_varVar(V_CDIR, lngE) = 0&
                                ElseIf lngCompDirOptType1 > 0& And lngCompDirOptType2 = 0& Then
                                  ' ** Only 1 is open.
                                  arr_varVar(V_CDIR, lngE) = lngCompDirOptType1
                                ElseIf lngCompDirOptType1 = 0& And lngCompDirOptType2 > 0& Then
                                  ' ** Only 2 is open.
                                  arr_varVar(V_CDIR, lngE) = lngCompDirOptType2
                                Else
                                  ' ** Both are open.
                                  strTmp04 = GetLastWord(strCompDir1)  ' ** Module Function: modStringFuncs.
                                  strTmp05 = GetLastWord(strCompDir2)  ' ** Module Function: modStringFuncs.
                                  If Val(strTmp04) < Val(strTmp05) Then
                                    ' ** 2 has control.
                                    arr_varVar(V_CDIR, lngE) = lngCompDirOptType2
                                  Else
                                    ' ** 1 has control.
                                    arr_varVar(V_CDIR, lngE) = lngCompDirOptType1
                                  End If
                                End If
                              Case False
                                arr_varVar(V_CDIR, lngE) = 0&
                              End Select
                            Else
                              ' ** Not typed?
                              Debug.Print "'" & strLine
                              Stop
                              Exit Do
                            End If
                            intPos01 = InStr(strTmp01, ",")
                            If intPos01 > 0 Then
                              blnFound = True
                            Else
                              ' ** The last definition on the line.
                              strTmp02 = strTmp01
                              intPos02 = InStr(strTmp02, " as ")
                              If intPos02 > 0 Then
                                strTmp03 = Trim(Mid(strTmp02, (intPos02 + 3)))  ' ** This should just be the type.
                                strTmp02 = Trim(Left(strTmp02, intPos02))       ' ** And this should be the defined item.
                                strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                                If strTmp03 = "String * 255" Then strTmp03 = "String"
                                If Left(strTmp03, 4) = "New " And Len(strTmp03) > 4 Then strTmp03 = Mid(strTmp03, 5)
                                lngVars = lngVars + 1&
                                lngE = lngVars - 1&
                                ReDim Preserve arr_varVar(V_ELEMS, lngE)
                                arr_varVar(V_DID, lngE) = lngThisDbsID
                                arr_varVar(V_VID, lngE) = Null
                                arr_varVar(V_VNAM, lngE) = strModName
                                arr_varVar(V_PID, lngE) = Null
                                arr_varVar(V_PNAM, lngE) = strProcName
                                arr_varVar(V_VAR, lngE) = strTmp02
                                arr_varVar(V_VTYP, lngE) = strTmp03
                                arr_varVar(V_VTYP2, lngE) = Null
                                Select Case intMode
                                Case 1, 2
                                  arr_varVar(V_SCOP, lngE) = "Local"
                                Case 3
                                  arr_varVar(V_SCOP, lngE) = "Static"
                                End Select
                                arr_varVar(V_DTYP, lngE) = strDecType
                                arr_varVar(V_VAL, lngE) = Null
                                arr_varVar(V_LINE, lngE) = lngX
                                arr_varVar(V_PTYP, lngE) = strProcType
                                arr_varVar(V_PTYP2, lngE) = Null
                                If strProcSubType <> vbNullString Then
                                  arr_varVar(V_PSTP, lngE) = strProcSubType
                                Else
                                  arr_varVar(V_PSTP, lngE) = Null
                                End If
                                arr_varVar(V_PSTP2, lngE) = Null
                                arr_varVar(V_CTYP, lngE) = Null
                                Select Case blnHasCompDir
                                Case True
                                  ' ** OK, what are we inside of?
                                  If lngCompDirOptType1 = 0& And lngCompDirOptType2 = 0& Then
                                    ' ** Nothing is open.
                                    arr_varVar(V_CDIR, lngE) = 0&
                                  ElseIf lngCompDirOptType1 > 0& And lngCompDirOptType2 = 0& Then
                                    ' ** Only 1 is open.
                                    arr_varVar(V_CDIR, lngE) = lngCompDirOptType1
                                  ElseIf lngCompDirOptType1 = 0& And lngCompDirOptType2 > 0& Then
                                    ' ** Only 2 is open.
                                    arr_varVar(V_CDIR, lngE) = lngCompDirOptType2
                                  Else
                                    ' ** Both are open.
                                    strTmp04 = GetLastWord(strCompDir1)  ' ** Module Function: modStringFuncs.
                                    strTmp05 = GetLastWord(strCompDir2)  ' ** Module Function: modStringFuncs.
                                    If Val(strTmp04) < Val(strTmp05) Then
                                      ' ** 2 has control.
                                      arr_varVar(V_CDIR, lngE) = lngCompDirOptType2
                                    Else
                                      ' ** 1 has control.
                                      arr_varVar(V_CDIR, lngE) = lngCompDirOptType1
                                    End If
                                  End If
                                Case False
                                  arr_varVar(V_CDIR, lngE) = 0&
                                End Select
                                Exit Do
                              Else
                                ' ** Not typed?
                                Debug.Print "'" & strLine
                                Stop
                                Exit Do
                              End If
                            End If
                          Loop  ' ** blnFound.
                         End If  ' ** intCnt.

                      End If  ' ** blnContinue.
                    End If  ' ** IsNumeric().
                  End If  ' ** intPos01.
                End If  ' ** Remark.
              End If  ' ** vbNullString.
            Next  ' ** lngX.

          End With  ' ** cod.
        End If  ' ** modGlobConst.
      End With  ' ** vbc.

      If lngZ Mod 100 = 0& Then
        Debug.Print "|  " & CStr(lngZ) & " of " & CStr(lngRecs)
        Debug.Print "'|";
      ElseIf lngZ Mod 10 = 0& Then
        Debug.Print "|";
      Else
        Debug.Print ".";
      End If
      DoEvents

    Next  ' ** vbc.
    Set cod = Nothing
  End With  ' ** vbp.
  Set vbc = Nothing
  Set vbp = Nothing
  Debug.Print
  DoEvents

  Select Case intMode
  Case 1
    Debug.Print "'CONSTANTS FOUND: " & CStr(lngVars)
  Case 2
    Debug.Print "'VARIABLES FOUND: " & CStr(lngVars)
  Case 3
    Debug.Print "'STATIC FOUND: " & CStr(lngVars)
  End Select  ' ** intMode.
  DoEvents

  'strLastModName = vbNullString: strLastProcName = vbNullString
  'For lngX = 0& To (lngVars - 1&)
  '  If arr_varVar(V_VNAM, lngX) <> strLastModName Then
  '    Debug.Print "'" & arr_varVar(V_VNAM, lngX) & ":"
  '    DoEvents
  '    strLastModName = arr_varVar(V_VNAM, lngX)
  '  End If
  '  If arr_varVar(V_PNAM, lngX) <> strLastProcName Then
  '    Debug.Print "'  " & arr_varVar(V_PNAM, lngX) & ":"
  '    DoEvents
  '    strLastProcName = arr_varVar(V_PNAM, lngX)
  '  End If
  '  Debug.Print "'    " & arr_varVar(V_VAR, lngX) & "    " & arr_varVar(V_VTYP, lngX) & "    " & arr_varVar(V_VAL, lngX)
  '  DoEvents
  'Next  ' ** lngX.

  If lngVars > 0& Then

    Set dbs = CurrentDb
    With dbs

      Set rst1 = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      Set rst2 = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)

      Debug.Print "'CHECKING ID'S..."
      DoEvents

      Debug.Print "'|";
      DoEvents
 
      strLastModName = vbNullString: strLastProcName = vbNullString
      lngVBComID = 0&: lngVBComProcID = 0&: lngComType = 0&
      For lngX = 0& To (lngVars - 1&)
        If arr_varVar(V_VNAM, lngX) <> strLastModName Then
          With rst1
            .MoveFirst
            If ![dbs_id] = arr_varVar(V_DID, lngX) And ![vbcom_name] = arr_varVar(V_VNAM, lngX) Then
              lngVBComID = ![vbcom_id]
              lngComType = ![comtype_type]
            Else
              .FindFirst "[dbs_id] = " & CStr(arr_varVar(V_DID, lngX)) & " And [vbcom_name] = '" & arr_varVar(V_VNAM, lngX) & "'"
              If .NoMatch = False Then
                lngVBComID = ![vbcom_id]
                lngComType = ![comtype_type]
              Else
                Stop
              End If
            End If
          End With  ' ** rst1.
        End If
        arr_varVar(V_VID, lngX) = lngVBComID
        arr_varVar(V_CTYP, lngX) = lngComType
        If arr_varVar(V_PNAM, lngX) <> strLastProcName Then
          With rst2
            .MoveFirst
            If ![dbs_id] = arr_varVar(V_DID, lngX) And ![vbcom_id] = lngVBComProcID And ![vbcomproc_name] = arr_varVar(V_PNAM, lngX) Then
              lngVBComProcID = ![vbcomproc_id]
            Else
              .FindFirst "[dbs_id] = " & CStr(arr_varVar(V_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varVar(V_VID, lngX)) & " And " & _
                "[vbcomproc_name] = '" & arr_varVar(V_PNAM, lngX) & "'"
              If .NoMatch = False Then
                lngVBComProcID = ![vbcomproc_id]
              Else
                If arr_varVar(V_PNAM, lngX) = THIS_PROC Then
                  lngVBComProcID = 0&
                Else
                  'Stop
                End If
              End If
            End If
          End With  ' ** rst2.
        End If
        arr_varVar(V_PID, lngX) = lngVBComProcID
        If Left(arr_varVar(V_VAR, lngX), 4) = "arr_" Then
          ' ** It's an array. Further typing will have to come later.
          arr_varVar(V_VTYP2, lngX) = vbArray
        Else
          Select Case arr_varVar(V_VTYP, lngX)
          Case "Integer"
            arr_varVar(V_VTYP2, lngX) = vbInteger
          Case "Long"
            arr_varVar(V_VTYP2, lngX) = vbLong
          Case "Single"
            arr_varVar(V_VTYP2, lngX) = vbSingle
          Case "Double"
            arr_varVar(V_VTYP2, lngX) = vbDouble
          Case "Currency"
            arr_varVar(V_VTYP2, lngX) = vbCurrency
          Case "Date"
            arr_varVar(V_VTYP2, lngX) = vbDate
          Case "String"
            arr_varVar(V_VTYP2, lngX) = vbString
          Case "Object"
            arr_varVar(V_VTYP2, lngX) = vbObject
          Case "Boolean"
            arr_varVar(V_VTYP2, lngX) = vbBoolean
          Case "Variant"
            arr_varVar(V_VTYP2, lngX) = vbVariant
          Case Else
            ' ** All other various objects.
            intPos01 = InStr(arr_varVar(V_VTYP, lngX), ".")
            If intPos01 > 0 Then
              arr_varVar(V_VTYP2, lngX) = vbObject
            Else
              ' ** For now, give it this.
              arr_varVar(V_VTYP2, lngX) = vbUserDefinedType
            End If
          End Select
        End If  ' ** V_VTYP.
        Select Case arr_varVar(V_SCOP, lngX)
        Case "Public"
          arr_varVar(V_SCOP2, lngX) = 1&
        Case "Private"
          arr_varVar(V_SCOP2, lngX) = 2&
        Case "Static"
          arr_varVar(V_SCOP2, lngX) = 4&
        Case "Local"
          arr_varVar(V_SCOP2, lngX) = 6&
        Case Else
          Stop
        End Select  ' ** V_SCOP.
        Select Case arr_varVar(V_DTYP, lngX)
        Case "Compiler"
          arr_varVar(V_DTYP2, lngX) = 1&
        Case "Constant"
          arr_varVar(V_DTYP2, lngX) = 2&
        Case "DefType"
          arr_varVar(V_DTYP2, lngX) = 3&
        Case "Enum"
          arr_varVar(V_DTYP2, lngX) = 4&
        Case "Function"
          arr_varVar(V_DTYP2, lngX) = 5&
        Case "Sub"
          arr_varVar(V_DTYP2, lngX) = 6&
        Case "Type"
          arr_varVar(V_DTYP2, lngX) = 7&
        Case "Variable"
          arr_varVar(V_DTYP2, lngX) = 8&
        Case Else
          Stop
        End Select  ' ** V_DTYP.
        Select Case arr_varVar(V_PTYP, lngX)
        Case "Declaration"
          arr_varVar(V_PTYP2, lngX) = 1&
        Case "Function"
          arr_varVar(V_PTYP2, lngX) = 2&
        Case "Sub"
          arr_varVar(V_PTYP2, lngX) = 3&
        Case "Property"
          arr_varVar(V_PTYP2, lngX) = 4&
        End Select  ' ** V_PTYP.
        If IsNull(arr_varVar(V_PSTP, lngX)) = True Then
          arr_varVar(V_PSTP, lngX) = "-"
          arr_varVar(V_PSTP2, lngX) = 4&
        Else
          Select Case arr_varVar(V_PSTP, lngX)
          Case "Get"
            arr_varVar(V_PSTP2, lngX) = 1&
          Case "Let"
            arr_varVar(V_PSTP2, lngX) = 2&
          Case "Set"
            arr_varVar(V_PSTP2, lngX) = 3&
          End Select  ' ** V_PSTP.
        End If
        If (lngX + 1&) Mod 1000 = 0& Then
          Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngVars)
          Debug.Print "'|";
        ElseIf (lngX + 1&) Mod 100 = 0& Then
          Debug.Print "|";
        ElseIf (lngX + 1&) Mod 10 = 0& Then
          Debug.Print ".";
        End If
        DoEvents
      Next  ' ** lngX.

      rst1.Close
      rst2.Close
      Set rst1 = Nothing
      Set rst2 = Nothing
      Debug.Print
      DoEvents

      Debug.Print "'WRITING..."
      DoEvents

      Debug.Print "'|";
      DoEvents

      blnAddAll = False: blnAdd = False
      Set rst1 = .OpenRecordset("tblVBComponent_Declaration_Local", dbOpenDynaset, dbConsistent)
      With rst1
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        Else
          Select Case intMode
          Case 1  ' ** Local Constant.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_02", "[scopetype_type] = 'Local' And [dectype_type] = 'Constant'")
          Case 2  ' ** Local Variable.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_02", "[scopetype_type] = 'Local' And [dectype_type] = 'Variable'")
          Case 3  ' ** Static Variable.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_02", "[scopetype_type] = 'Static'")
          End Select
          If varTmp00 = 0 Then
            blnAddAll = True
          End If
        End If
        For lngX = 0& To (lngVars - 1&)
          blnAdd = False
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            '.MoveFirst
            '.FindFirst "[dbs_id] = " & CStr(arr_varVar(V_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varVar(V_VID, lngX)) & " And " & _
            '  "[vbdec_name] = '" & arr_varVar(V_VAR, lngX) & "'"
            'Select Case .NoMatch
            'Case True
            '  blnAdd = True
            'Case False
            '
            'End Select
          End Select  ' ** blnAddAll.
          If blnAdd = True Then
            .AddNew
            ![dbs_id] = arr_varVar(V_DID, lngX)
            ![vbcom_id] = arr_varVar(V_VID, lngX)
            ![vbcomproc_id] = arr_varVar(V_PID, lngX)
            ' ** ![vbdecloc_id] : AutoNumber.
            ![vbdecloc_module] = arr_varVar(V_VNAM, lngX)
            ![vbdecloc_procedure] = arr_varVar(V_PNAM, lngX)
            ![vbdecloc_name] = arr_varVar(V_VAR, lngX)
            ![comtype_type] = arr_varVar(V_CTYP, lngX)
            ![scopetype_type] = arr_varVar(V_SCOP, lngX)
            ![dectype_type] = arr_varVar(V_DTYP, lngX)
            ![datatype_vb_type] = arr_varVar(V_VTYP2, lngX)
            ![vbdecloc_vbtype] = arr_varVar(V_VTYP, lngX)
            ![proctype_type] = arr_varVar(V_PTYP, lngX)
            ![procsubtype_type] = arr_varVar(V_PSTP, lngX)
            ![compdiropt_type] = arr_varVar(V_CDIR, lngX)
            ![vbdecloc_isarray] = False
            ![vbdecloc_array] = Null
            ![vbdecloc_value] = NullIfNullStr(arr_varVar(V_VAL, lngX))  ' ** Module Function: modStringFuncs.
            ![vbdecloc_linenum] = arr_varVar(V_LINE, lngX)
            ![vbdecloc_notused] = False
            ![vbdecloc_usecnt] = 0&
            ![vbdecloc_datemodified] = Now()
On Error Resume Next
            .Update
            If ERR.Number <> 0 Then
              .CancelUpdate
            End If
On Error GoTo 0
          End If  ' ** blnAdd.
          If (lngX + 1&) Mod 1000 = 0& Then
            Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngVars)
            Debug.Print "'|";
          ElseIf (lngX + 1&) Mod 100 = 0& Then
            Debug.Print "|";
          ElseIf (lngX + 1&) Mod 10 = 0& Then
            Debug.Print ".";
          End If
          DoEvents
        Next  ' ** lngX.
        Debug.Print
        DoEvents

        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

  End If  ' ** lngVars.

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_Var_LocalDoc = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_Var_Usage1() As Boolean
' ** Document module-level constant and variable usage.

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_Var_Usage1"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim lngLines As Long, lngDecLines As Long
  Dim strModName As String, strProcName As String, strLine As String, strLastModName As String, strLastProcName As String
  Dim lngMods As Long, arr_varMod As Variant
  Dim lngVars As Long, arr_varVar As Variant
  Dim lngHits As Long, arr_varHit() As Variant
  Dim lngThisDbsID As Long, lngVBComID As Long, lngVBComProcID As Long, lngTotalRecs As Long
  Dim strFind As String
  Dim blnContinue As Boolean, blnFound As Boolean, blnAddAll As Boolean, blnAdd As Boolean
  Dim intPos01 As Integer, intLen As Integer, intMode As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
  Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varMod().
  'Const M_DID  As Integer = 0
  Const M_VID  As Integer = 1
  Const M_VNAM As Integer = 2
  'Const M_CNT  As Integer = 3

  ' ** Array: arr_varVar().
  Const V_DID  As Integer = 0
  'Const V_VID  As Integer = 1
  Const V_XID  As Integer = 2
  'Const V_VNAM As Integer = 3
  Const V_XNAM As Integer = 4
  'Const V_SCOP As Integer = 5
  Const V_TYP  As Integer = 6
  'Const V_CNT  As Integer = 7

  ' ** Array: arr_varHit().
  Const H_ELEMS As Integer = 10  ' ** Array's first element UBound().
  Const H_DID  As Integer = 0
  Const H_VID  As Integer = 1
  Const H_VNAM As Integer = 2
  Const H_PID  As Integer = 3
  Const H_PNAM As Integer = 4
  Const H_XID  As Integer = 5
  Const H_VAR  As Integer = 6
  Const H_XTYP As Integer = 7
  Const H_LINE As Integer = 8
  Const H_CODE As Integer = 9
  Const H_RAW  As Integer = 10

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  'intMode = 1  ' ** Public Constant.
  'intMode = 2  ' ** Public Variable.
  'intMode = 3  ' ** Public Miscellaneous.
  'intMode = 4  ' ** Private Constant.
  'intMode = 5  ' ** Private Variable.
  intMode = 6  ' ** Private Miscellaneous.

  'HOW DO I DELETE DEAD ONES?
  'FOR NOW, DELETE ALL AND REGENERATE!

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    Select Case intMode
    Case 1  ' ** Public Constant.
      ' ** tblVBComponent, linked to zz_qry_VBComponent_Var_10_01_03
      ' ** (zz_qry_VBComponent_Var_10_01_02 (zz_qry_VBComponent_Var_10_01_01
      ' ** (tblVBComponent_Declaration, just Public Constants), grouped by vbcom_id,
      ' ** with cnt_con), grouped and summed, with cnt_con), all Trust.mdb modules, with cnt_con.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_10_01_04")
    Case 2  ' ** Public Variable.
      ' ** tblVBComponent, linked to zz_qry_VBComponent_Var_11_01_03
      ' ** (zz_qry_VBComponent_Var_11_01_02 (zz_qry_VBComponent_Var_11_01_01
      ' ** (tblVBComponent_Declaration, just Public Variables), grouped by vbcom_id,
      ' ** with cnt_var), grouped and summed, with cnt_var), all Trust.mdb modules, with cnt_con.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_11_01_04")
    Case 3  ' ** Public Miscellaneous.
      ' ** tblVBComponent, linked to zz_qry_VBComponent_Var_17_01_03 (zz_qry_VBComponent_Var_17_01_02
      ' ** (zz_qry_VBComponent_Var_17_01_01 (Union of zz_qry_VBComponent_Var_12_01_01 (tblVBComponent_Declaration,
      ' ** just Public Functions), zz_qry_VBComponent_Var_13_01_01 (tblVBComponent_Declaration,
      ' ** just Public Subs), zz_qry_VBComponent_Var_14_01_01 (tblVBComponent_Declaration, just Public Enums),
      ' ** zz_qry_VBComponent_Var_15_01_01 (tblVBComponent_Declaration, just Public Types), Public miscellaneous),
      ' ** grouped by vbcom_id, with cnt_var), grouped and summed, with cnt_var), all Trust.mdb modules, with cnt_var.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_17_01_04")
    Case 4  ' ** Private Constant.
      ' ** zz_qry_VBComponent_Var_20_01_01 (tblVBComponent_Declaration,
      ' ** just Private Constants), grouped by vbcom_id, with cnt_con.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_20_01_02")
    Case 5  ' ** Private Variable.
      ' ** zz_qry_VBComponent_Var_21_01_01 (tblVBComponent_Declaration,
      ' ** just Private Variables), grouped by vbcom_id, with cnt_var.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_21_01_02")
    Case 6  ' ** Private Miscellaneous.
      ' ** zz_qry_VBComponent_Var_27_01_01 (Union of zz_qry_VBComponent_Var_22_01_01
      ' ** (tblVBComponent_Declaration, just Private Functions), zz_qry_VBComponent_Var_23_01_01
      ' ** (tblVBComponent_Declaration, just Private Subs), zz_qry_VBComponent_Var_24_01_01
      ' ** (tblVBComponent_Declaration, just Private Enums), zz_qry_VBComponent_Var_25_01_01
      ' ** (tblVBComponent_Declaration, just Private Types), zz_qry_VBComponent_Var_26_01_01
      ' ** (tblVBComponent_Declaration, just Private Compiler Directives),
      ' ** Private miscellaneous), grouped by vbcom_id, with cnt_var.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_27_01_02")
    End Select
    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngMods = .RecordCount
      .MoveFirst
      arr_varMod = .GetRows(lngMods)
      ' *************************************************
      ' ** Array: arr_varMod()
      ' **
      ' **   Field  Element  Name            Constant
      ' **   =====  =======  ==============  ==========
      ' **     1       0     dbs_id          M_DID
      ' **     2       1     vbcom_id        M_VID
      ' **     3       2     vbdec_module    M_VNAM
      ' **     4       3     cnt_var         M_CNT
      ' **
      ' *************************************************
      .Close
    End With  ' ** rst1.
    Set rst1 = Nothing
    Set qdf = Nothing

    lngHits = 0&
    ReDim arr_varHit(H_ELEMS, 0)

    Select Case intMode
    Case 1  ' ** Public Constant.
      ' ** zz_qry_VBComponent_Var_10_01_02 (zz_qry_VBComponent_Var_10_01_01
      ' ** (tblVBComponent_Declaration, just Public Constants), grouped
      ' ** by vbcom_id, with cnt_con), grouped and summed, with cnt_con.
      lngTotalRecs = DLookup("[cnt_con]", "zz_qry_VBComponent_Var_10_01_03")
      lngTotalRecs = (lngTotalRecs * lngMods)
      Debug.Print "'TRACING " & CStr(lngTotalRecs) & " CONSTANTS!"
    Case 2  ' ** Public Variable.
      ' ** zz_qry_VBComponent_Var_11_01_02 (zz_qry_VBComponent_Var_11_01_01
      ' ** (tblVBComponent_Declaration, just Public Variables), grouped by
      ' ** vbcom_id, with cnt_var), grouped and summed, with cnt_var.
      lngTotalRecs = DLookup("[cnt_var]", "zz_qry_VBComponent_Var_11_01_03")
      lngTotalRecs = (lngTotalRecs * lngMods)
      Debug.Print "'TRACING " & CStr(lngTotalRecs) & " VARIABLES!"
    Case 3  ' ** Public Miscellaneous.
      ' ** zz_qry_VBComponent_Var_17_01_02 (zz_qry_VBComponent_Var_17_01_01 (Union of
      ' ** zz_qry_VBComponent_Var_12_01_01 (tblVBComponent_Declaration, just Public Functions),
      ' ** zz_qry_VBComponent_Var_13_01_01 (tblVBComponent_Declaration, just Public Subs),
      ' ** zz_qry_VBComponent_Var_14_01_01 (tblVBComponent_Declaration, just Public Enums),
      ' ** zz_qry_VBComponent_Var_15_01_01 (tblVBComponent_Declaration, just Public Types),
      ' ** Public miscellaneous), grouped by vbcom_id, with cnt_var), grouped and summed, with cnt_var.
      lngTotalRecs = DLookup("[cnt_var]", "zz_qry_VBComponent_Var_17_01_03")
      lngTotalRecs = (lngTotalRecs * lngMods)
      Debug.Print "'TRACING " & CStr(lngTotalRecs) & " MISCELLANEOUS!"
    Case 4  ' ** Private Constant.
      ' ** zz_qry_VBComponent_Var_20_01_02 (zz_qry_VBComponent_Var_20_01_01
      ' ** (tblVBComponent_Declaration, just Private Constants), grouped by
      ' ** vbcom_id, with cnt_con), grouped and summed, with cnt_con.
      lngTotalRecs = DLookup("[cnt_con]", "zz_qry_VBComponent_Var_20_01_03")
      Debug.Print "'TRACING " & CStr(lngTotalRecs) & " CONSTANTS!"
    Case 5  ' ** Private Variable.
      ' ** zz_qry_VBComponent_Var_21_01_02 (zz_qry_VBComponent_Var_21_01_01
      ' ** (tblVBComponent_Declaration, just Private Variables), grouped by
      ' ** vbcom_id, with cnt_var), grouped and summed, with cnt_var.
      lngTotalRecs = DLookup("[cnt_var]", "zz_qry_VBComponent_Var_21_01_03")
      Debug.Print "'TRACING " & CStr(lngTotalRecs) & " VARIABLES!"
    Case 6  ' ** Private Miscellaneous.
      ' ** zz_qry_VBComponent_Var_27_01_02 (zz_qry_VBComponent_Var_27_01_01 (Union of
      ' ** zz_qry_VBComponent_Var_22_01_01 (tblVBComponent_Declaration, just Private Functions),
      ' ** zz_qry_VBComponent_Var_23_01_01 (tblVBComponent_Declaration, just Private Subs),
      ' ** zz_qry_VBComponent_Var_24_01_01 (tblVBComponent_Declaration, just Private Enums),
      ' ** zz_qry_VBComponent_Var_25_01_01 (tblVBComponent_Declaration, just Private Types),
      ' ** zz_qry_VBComponent_Var_26_01_01 (tblVBComponent_Declaration, just Private Compiler Directives),
      ' ** Private miscellaneous), grouped by vbcom_id, with cnt_var), grouped and summed, with cnt_var.
      lngTotalRecs = DLookup("[cnt_var]", "zz_qry_VBComponent_Var_27_01_03")
      Debug.Print "'TRACING " & CStr(lngTotalRecs) & " MISCELLANEOUS!"
    End Select
    lngZ = 0&
    DoEvents

    Debug.Print "'|";
    DoEvents

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For lngW = 0& To (lngMods - 1&)

        Select Case intMode
        Case 1  ' ** Public Constant.
          ' ** zz_qry_VBComponent_Var_10_01_01 (tblVBComponent_Declaration,
          ' ** just Public Constants), by unused specified [comid].
          Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_10_01_05")
        Case 2  ' ** Public Variable.
          ' ** zz_qry_VBComponent_Var_11_01_01 (tblVBComponent_Declaration,
          ' ** just Public Variables), by unused specified [comid].
          Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_11_01_05")
        Case 3  ' ** Public Miscellaneous.
          ' ** zz_qry_VBComponent_Var_17_01_01 (xx), by unused specified [comid].
          Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_17_01_05")
        Case 4  ' ** Private Constant.
          ' ** zz_qry_VBComponent_Var_20_01_01 (tblVBComponent_Declaration,
          ' ** just Private Constants), by specified [comid].
          Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_20_01_04")
        Case 5  ' ** Private Variable.
          ' ** zz_qry_VBComponent_Var_21_01_01 (tblVBComponent_Declaration,
          ' ** just Private Variables), by specified [comid].
          Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_21_01_04")
        Case 6  ' ** Private Miscellaneous.
          ' ** zz_qry_VBComponent_Var_27_01_01 (xx), by specified [comid].
          Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_27_01_04")
        End Select
        With qdf.Parameters
          ![comid] = arr_varMod(M_VID, lngW)
        End With
        Set rst1 = qdf.OpenRecordset
        With rst1
          .MoveLast
          lngVars = .RecordCount
          .MoveFirst
          arr_varVar = .GetRows(lngVars)
          ' ***************************************************
          ' ** Array: arr_varVar()
          ' **
          ' **   Field  Element  Name              Constant
          ' **   =====  =======  ================  ==========
          ' **     1       0     dbs_id            V_DID
          ' **     2       1     vbcom_id          V_VID
          ' **     3       2     vbdec_id          V_XID
          ' **     4       3     vbdec_module      V_VNAM
          ' **     5       4     vbdec_name        V_XNAM
          ' **     6       5     scopetype_type    V_SCOP
          ' **     7       6     dectype_type      V_TYP
          ' **     8       7     vbdec_usecnt      V_CNT
          ' **
          ' ***************************************************
          .Close
        End With
        Set rst1 = Nothing
        Set qdf = Nothing

        Set vbc = .VBComponents(arr_varMod(M_VNAM, lngW))
        With vbc
          strModName = .Name
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngX = 0& To (lngVars - 1&)
              lngZ = lngZ + 1&
              strFind = arr_varVar(V_XNAM, lngX)
              For lngY = 1& To lngLines
                strLine = Trim(.Lines(lngY, 1))
                If strLine <> vbNullString Then
                  If Left(strLine, 1) <> "'" Then
                    intPos01 = InStr(strLine, strFind)
                    If intPos01 > 0 Then
                      ' ** Check characters fore and aft.
                      intLen = Len(strFind)
                      blnContinue = True
                      If intPos01 > 1 Then
                        strTmp01 = Mid(strLine, (intPos01 - 1), 1)
                        Select Case strTmp01
                        Case " ", "(", ",", ":", "-", "="
                          ' ** These are OK.
                        Case ".", "[", "_", "'", Chr(34)
                          ' ** Not OK.
                          blnContinue = False
                        Case Else
                          If (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Or _
                              (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or _
                              (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Then  ' ** Numbers and letters.
                            ' ** Nope
                            blnContinue = False
                          Else
                            ' ** Let's see what's there.
                            Debug.Print "'  '" & strTmp01 & "'"
                            Stop
                          End If
                        End Select
                      Else
                        ' ** Variable at beginning of line.
                      End If
                      If blnContinue = True Then
                        If (intPos01 + intLen) > Len(strLine) Then
                          ' ** At end of line, so OK.
                        Else
                          strTmp01 = Mid(strLine, (intPos01 + intLen), 1)
                          Select Case strTmp01
                          Case " ", "(", ")", ",", ":", "&", "!", "="
                            ' ** These are OK.
                          Case "]", "_", Chr(34)
                            ' ** Not OK.
                            blnContinue = False
                          Case "."
                            ' ** This is OK for objects, but I don't think it is for plain variables.
                            strTmp02 = Left(strFind, 3)
                            Select Case strTmp02
                            Case "int", "lng", "sng", "dbl", "cur", "bln", "str", "dat", "var", "arr"
                              ' ** I don't think so.
                              blnContinue = False
                            Case Else
                              ' ** Good to go.
                            End Select
                          Case Else
                            If (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Or _
                                (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or _
                                (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Then  ' ** Numbers and letters.
                              ' ** Nope
                              blnContinue = False
                            Else
                              ' ** Let's see what's there.
                              Debug.Print "'  '" & strTmp01 & "'"
                              Stop
                            End If
                          End Select
                        End If
                      End If  ' ** blnContinue.
                      If blnContinue = True Then
                        strTmp02 = vbNullString
                        strProcName = .ProcOfLine(lngY, vbext_pk_Proc)
                        If strProcName = vbNullString Then strProcName = "Declaration"
                        intPos01 = InStr(strLine, " ")
                        If intPos01 > 0 Then
                          strTmp02 = Trim(Left(strLine, intPos01))
                          If IsNumeric(strTmp02) = False Then
                            strTmp02 = vbNullString
                          End If
                        End If
                        lngHits = lngHits + 1&
                        lngE = lngHits - 1&
                        ReDim Preserve arr_varHit(H_ELEMS, lngE)
                        ' ******************************************************
                        ' ** Array: arr_varHit()
                        ' **
                        ' **   Field  Element  Name                 Constant
                        ' **   =====  =======  ===================  ==========
                        ' **     1       0     dbs_id               H_DID
                        ' **     2       1     vbcom_id             H_VID
                        ' **     3       2     vbcom_name           H_VNAM
                        ' **     4       3     vbcomproc_id         H_PID
                        ' **     5       4     vbcomproc_name       H_PNAM
                        ' **     6       5     vbdec_id             H_XID
                        ' **     7       6     vbdec_name           H_VAR
                        ' **     8       7     vbsearch_linenum     H_LINE
                        ' **     9       8     vbsearch_codeline    H_CODE
                        ' **    10       9     vbsearch_text        H_RAW
                        ' **
                        ' ******************************************************
                        arr_varHit(H_DID, lngE) = arr_varVar(V_DID, lngX)
                        arr_varHit(H_VID, lngE) = Null
                        arr_varHit(H_VNAM, lngE) = strModName
                        arr_varHit(H_PID, lngE) = Null
                        arr_varHit(H_PNAM, lngE) = strProcName
                        arr_varHit(H_XID, lngE) = arr_varVar(V_XID, lngX)
                        arr_varHit(H_VAR, lngE) = strFind
                        arr_varHit(H_XTYP, lngE) = arr_varVar(V_TYP, lngX)
                        arr_varHit(H_LINE, lngE) = lngY
                        If strTmp02 <> vbNullString Then
                          arr_varHit(H_CODE, lngE) = CLng(strTmp02)
                        Else
                          arr_varHit(H_CODE, lngE) = Null
                        End If
                        arr_varHit(H_RAW, lngE) = strLine
                      End If  ' ** blnContinue.
                    End If  ' ** intPos01.
                  End If  ' ** Remark.
                End If  ' ** vbNullString.
              Next  ' ** lngY.
              If lngZ Mod 10000& = 0& Then
                Debug.Print "|  " & CStr(lngZ) & " of " & CStr(lngTotalRecs)
                Debug.Print "'|";
              ElseIf lngZ Mod 1000& = 0& Then
                Debug.Print "|";
              ElseIf lngZ Mod 100& = 0& Then
                Debug.Print ".";
              End If
              DoEvents
            Next  ' ** lngX.
          End With  ' ** cod.
        End With  ' ** vbc.

      Next  ' ** lngW
    End With  ' ** vbp.
    Debug.Print
    DoEvents

    Debug.Print "'HITS: " & CStr(lngHits)
    DoEvents

    If lngHits > 0& Then

      Set rst1 = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      Set rst2 = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)

      Debug.Print "'CHECKING ID'S..."
      DoEvents

      Debug.Print "'|";
      DoEvents

      lngVBComID = 0&: lngVBComProcID = 0&
      strLastModName = vbNullString: strLastProcName = vbNullString
      For lngX = 0& To (lngHits - 1&)
        If arr_varHit(H_VNAM, lngX) <> strLastModName Then
          strLastProcName = vbNullString
          With rst1
            .MoveFirst
            If ![dbs_id] = arr_varHit(H_DID, lngX) And ![vbcom_name] = arr_varHit(H_VNAM, lngX) Then
              lngVBComID = ![vbcom_id]
            Else
              .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_name] = '" & arr_varHit(H_VNAM, lngX) & "'"
              If .NoMatch = False Then
                lngVBComID = ![vbcom_id]
              Else
                Stop
              End If
            End If
            strLastModName = ![vbcom_name]
          End With  ' ** rst1.
        End If  ' ** strLastModName.
        arr_varHit(H_VID, lngX) = lngVBComID
        If arr_varHit(H_PNAM, lngX) <> strLastProcName Then
          With rst2
            .MoveFirst
            If ![dbs_id] = arr_varHit(H_DID, lngX) And ![vbcom_id] = lngVBComID And ![vbcomproc_name] = arr_varHit(H_PNAM, lngX) Then
              lngVBComProcID = ![vbcomproc_id]
            Else
              .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
                "[vbcomproc_name] = '" & arr_varHit(H_PNAM, lngX) & "'"
              If .NoMatch = False Then
                lngVBComProcID = ![vbcomproc_id]
              Else
                'Stop
              End If
            End If
            strLastProcName = ![vbcomproc_name]
          End With  ' ** rst2.
        End If  ' ** strLastProcName.
        arr_varHit(H_PID, lngX) = lngVBComProcID
        If (lngX + 1&) Mod 1000& = 0& Then
          Debug.Print "|  " & CStr((lngX + 1&)) & " of " & CStr(lngHits)
          Debug.Print "'|";
        ElseIf (lngX + 1&) Mod 100& = 0& Then
          Debug.Print "|";
        ElseIf (lngX + 1&) Mod 10& = 0& Then
          Debug.Print ".";
        End If
        DoEvents
      Next  ' ** lngX.
      Debug.Print
      DoEvents

      rst1.Close
      rst2.Close
      Set rst1 = Nothing
      Set rst2 = Nothing

      Debug.Print "'WRITING..."
      DoEvents

      Debug.Print "'|";
      DoEvents

      blnAddAll = False: blnAdd = False
      Set rst1 = .OpenRecordset("tblVBComponent_Declaration_Detail", dbOpenDynaset, dbConsistent)
      With rst1
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        Else
          Select Case intMode
          Case 1  ' ** Public Constant.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_01", "[scopetype_type] = 'Public' And [dectype_type] = 'Constant'")
          Case 2  ' ** Public Variable.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_01", "[scopetype_type] = 'Public' And [dectype_type] = 'Variable'")
          Case 3  ' ** Public Miscellaneous.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_01", "[scopetype_type] = 'Public' And [dectype_type] " & _
              "Not In ('Constant','Variable')")
          Case 4  ' ** Private Constant.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_01", "[scopetype_type] = 'Private' And [dectype_type] = 'Constant'")
          Case 5  ' ** Private Variable.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_01", "[scopetype_type] = 'Private' And [dectype_type] = 'Variable'")
          Case 6  ' ** Private Miscellaneous.
            varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_01", "[scopetype_type] = 'Private' And [dectype_type] " & _
              "Not In ('Constant','Variable')")
          End Select
          If varTmp00 = 0 Then
            blnAddAll = True
          End If
        End If
  'intMode = 1  ' ** Public Constant.
  'intMode = 2  ' ** Public Variable.
  'intMode = 3  ' ** Public Miscellaneous.
  'intMode = 4  ' ** Private Constant.
  'intMode = 5  ' ** Private Variable.
  'intMode = 6  ' ** Private Miscellaneous.
        For lngX = 0& To (lngHits - 1&)
          blnFound = False
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            .MoveFirst
            If ![dbs_id] = arr_varHit(H_DID, lngX) And ![vbcom_id] = arr_varHit(H_VID, lngX) And ![vbcomproc_id] = arr_varHit(H_PID, lngX) And _
            ![vbdecdet_linenum] = arr_varHit(H_LINE, lngX) Then
              blnFound = True
            Else
              .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
                "[vbcomproc_id] = " & CStr(arr_varHit(H_PID, lngX)) & " And " & "[vbdecdet_linenum] = " & CStr(arr_varHit(H_LINE, lngX))
              Select Case .NoMatch
              Case True
                blnAdd = True
              Case False
                blnFound = True
              End Select
            End If
          End Select  ' ** blnAddAll.
          If blnFound = True Then
            'If ![vbsearch_codeline] <> arr_varHit(H_CODE, lngX) Then
            '  .Edit
            '  ![vbsearch_codeline] = arr_varHit(H_CODE, lngX)
            '  ![vbsearch_datemodified] = Now()
            '  .Update
            'End If
            'If ![vbsearch_text] <> arr_varHit(H_RAW, lngX) Then
            '  .Edit
            '  ![vbsearch_text] = arr_varHit(H_RAW, lngX)
            '  ![vbsearch_datemodified] = Now()
            '  .Update
            'End If
          End If
          If blnAdd = True Then
            .AddNew
            ![vbdec_id] = arr_varHit(H_XID, lngX)
            ' ** ![vbdecdet_id] : AutoNumber.
            ![dbs_id] = arr_varHit(H_DID, lngX)
            ![vbcom_id] = arr_varHit(H_VID, lngX)
            ![vbcomproc_id] = arr_varHit(H_PID, lngX)
            ![dectype_type] = arr_varHit(H_XTYP, lngX)
            ![compdiropt_type] = 0&
            ![vbdecdet_linenum] = arr_varHit(H_LINE, lngX)
            ![vbdecdet_notused] = False
            ![vbdecdet_datemodified] = Now()
On Error Resume Next
            .Update
            If ERR.Number <> 0 Then
              .CancelUpdate
            End If
On Error GoTo 0
          End If  ' ** blnAdd.
          If (lngX + 1&) Mod 10000& = 0& Then
            Debug.Print "|  " & CStr((lngX + 1&)) & " of " & CStr(lngHits)
            Debug.Print "'|";
          ElseIf (lngX + 1&) Mod 1000& = 0& Then
            Debug.Print "|";
          ElseIf (lngX + 1&) Mod 100& = 0& Then
            Debug.Print ".";
          End If
          DoEvents
        Next  ' ** lngX.
        Debug.Print
        DoEvents

        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

    End If  ' ** lngHits.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_Var_Usage1 = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_Var_Usage2() As Boolean
' ** Document procedure-level constant and variable usage.

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_Var_Usage2"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngMods As Long, arr_varMod As Variant
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngItems As Long, arr_varItem As Variant
  Dim lngHits As Long, arr_varHit() As Variant
  Dim lngDels As Long
  Dim strModName As String, strProcName As String, strLine As String, strFind As String
  Dim lngThisDbsID As Long, lngVBComID As Long, lngVBComProcID As Long
  Dim lngTotMods As Long, lngTotProcs As Long, lngTotItems As Long
  Dim lngLines As Long, lngStartLine As Long, lngEndLine As Long, lngCnt As Long
  Dim blnContinue As Boolean, blnAddAll As Boolean, blnAdd As Boolean
  Dim intMode As Integer, intPos01 As Integer, intPos02 As Integer, intLen As Integer
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngV As Long, lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varMod().
  'Const M_DID  As Integer = 0
  Const M_VID  As Integer = 1
  Const M_VNAM As Integer = 2
  'Const M_PCNT As Integer = 3
  'Const M_CCNT As Integer = 4

  ' ** Array: arr_varProc().
  'Const P_DID  As Integer = 0
  'Const P_VID  As Integer = 1
  'Const P_VNAM As Integer = 2
  Const P_PID  As Integer = 3
  Const P_PNAM As Integer = 4
  Const P_PTYP As Integer = 5
  Const P_STYP As Integer = 6
  'Const P_CCNT As Integer = 7

  ' ** Array: arr_varItem().
  Const I_DID  As Integer = 0
  Const I_VID  As Integer = 1
  'Const I_VNAM As Integer = 2
  Const I_PID  As Integer = 3
  Const I_PNAM As Integer = 4
  Const I_LID  As Integer = 5
  Const I_LNAM As Integer = 6
  'Const I_LINE As Integer = 7

  ' ** Array: arr_varHit().
  Const H_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const H_DID  As Integer = 0
  Const H_VID  As Integer = 1
  Const H_PID  As Integer = 2
  Const H_PNAM As Integer = 3
  Const H_LID  As Integer = 4
  Const H_LNAM As Integer = 5
  Const H_LINE As Integer = 6

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  'intMode = 1  ' ** Local Constants.
  'intMode = 2  ' ** Local Variables.
  intMode = 3  ' ** Static Variables.

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    Select Case intMode
    Case 1
      ' ** zz_qry_VBComponent_Var_01_01_03 (zz_qry_VBComponent_Var_01_01_02
      ' ** (zz_qry_VBComponent_Var_01_01_01 (tblVBComponent_Declaration_Local, just
      ' ** Local Constants), grouped by vbcom_id, vbcomproc_id, with cnt_con), grouped and
      ' ** summed, by vbcom_id, with cnt_proc, cnt_con), grouped and summed,
      ' ** by dbs_id, with cnt_mod, cnt_proc, cnt_con.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_01_01_04")
    Case 2
      ' ** zz_qry_VBComponent_Var_02_01_03 (zz_qry_VBComponent_Var_02_01_02
      ' ** (zz_qry_VBComponent_Var_02_01_01 (tblVBComponent_Declaration_Local,
      ' ** just Local Variables), grouped by vbcom_id, vbcomproc_id, with cnt_var),
      ' ** grouped and summed, by vbcom_id, with cnt_proc, cnt_var), grouped and summed,
      ' ** by dbs_id, with cnt_mod, cnt_proc, cnt_var.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_02_01_04")
    Case 3
      ' ** zz_qry_VBComponent_Var_03_01_03 (zz_qry_VBComponent_Var_03_01_02
      ' ** (zz_qry_VBComponent_Var_03_01_01 (tblVBComponent_Declaration_Local,
      ' ** just Static Variables), grouped by vbcom_id, vbcomproc_id, with cnt_var),
      ' ** grouped and summed, by vbcom_id, with cnt_proc, cnt_var), grouped and
      ' ** summed, by dbs_id, with cnt_mod, cnt_proc, cnt_var.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_03_01_04")
    End Select
    Set rst = qdf.OpenRecordset
    With rst
      .MoveFirst
      lngTotMods = ![cnt_mod]
      lngTotProcs = ![cnt_proc]
      Select Case intMode
      Case 1
        lngTotItems = ![cnt_con]
      Case 2, 3
        lngTotItems = ![cnt_var]
      End Select
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Set qdf = Nothing

    Select Case intMode
    Case 1
      Debug.Print "'TRACING " & CStr(lngTotItems) & " CONSTANTS, IN " & CStr(lngTotProcs) & " PROCS FROM " & CStr(lngTotMods) & " MODS!"
    Case 2
      Debug.Print "'TRACING " & CStr(lngTotItems) & " VARIABLES, IN " & CStr(lngTotProcs) & " PROCS FROM " & CStr(lngTotMods) & " MODS!"
    Case 3
      Debug.Print "'TRACING " & CStr(lngTotItems) & " STATICS, IN " & CStr(lngTotProcs) & " PROCS FROM " & CStr(lngTotMods) & " MODS!"
    End Select
    DoEvents

    Select Case intMode
    Case 1
      ' ** zz_qry_VBComponent_Var_01_01_02 (zz_qry_VBComponent_Var_01_01_01
      ' ** (tblVBComponent_Declaration_Local, just Local Constants), grouped by vbcom_id,
      ' ** vbcomproc_id, with cnt_con), grouped and summed, by vbcom_id, with cnt_proc, cnt_con.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_01_01_03")
    Case 2
      ' ** zz_qry_VBComponent_Var_02_01_02 (zz_qry_VBComponent_Var_02_01_01
      ' ** (tblVBComponent_Declaration_Local, just Local Variables), grouped by vbcom_id,
      ' ** vbcomproc_id, with cnt_var), grouped and summed, by vbcom_id, with cnt_proc, cnt_var.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_02_01_03")
    Case 3
      ' ** zz_qry_VBComponent_Var_03_01_02 (zz_qry_VBComponent_Var_03_01_01
      ' ** (tblVBComponent_Declaration_Local, just Static Variables), grouped by vbcom_id,
      ' ** vbcomproc_id, with cnt_var), grouped and summed, by vbcom_id, with cnt_proc, cnt_var
      Set qdf = .QueryDefs("zz_qry_VBComponent_Var_03_01_03")
    End Select
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngMods = .RecordCount
      .MoveFirst
      arr_varMod = .GetRows(lngMods)
      ' *****************************************************
      ' ** Array: arr_varMod()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     dbs_id              M_DID
      ' **     2       1     vbcom_id            M_VID
      ' **     3       2     vbdecloc_module     M_VNAM
      ' **     4       3     cnt_proc            M_PCNT
      ' **     5       4     cnt_con, cnt_var    M_CCNT
      ' **
      ' *****************************************************
      .Close
    End With  ' ** rst
    Set rst = Nothing
    Set qdf = Nothing

    If lngMods > 0& Then

      lngHits = 0&
      ReDim arr_varHit(H_ELEMS, 0)

      Set vbp = Application.VBE.ActiveVBProject
      With vbp

        Debug.Print "'|";
        DoEvents

        lngV = 0&: lngDels = 0&
        ' ** For each documented module.
        For lngW = 0& To (lngMods - 1&)
          Set vbc = .VBComponents(arr_varMod(M_VNAM, lngW))
          With vbc

            strModName = .Name
            lngVBComID = arr_varMod(M_VID, lngW)

            Select Case intMode
            Case 1
              ' ** zz_qry_VBComponent_Var_01_01_02 (zz_qry_VBComponent_Var_01_01_01
              ' ** (tblVBComponent_Declaration_Local, just Local Constants), grouped by vbcom_id,
              ' ** vbcomproc_id, with cnt_con), list of 1 module's procedures, with cnt_con, by specified [comid].
              Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_01_02_02")
            Case 2
              ' ** zz_qry_VBComponent_Var_02_01_02 (zz_qry_VBComponent_Var_02_01_01
              ' ** (tblVBComponent_Declaration_Local, just Local Variables), grouped by vbcom_id,
              ' ** vbcomproc_id, with cnt_var), list of 1 module's procedures, with cnt_var, by specified [comid].
              Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_02_02_02")
            Case 3
              ' ** zz_qry_VBComponent_Var_03_01_02 (zz_qry_VBComponent_Var_03_01_01
              ' ** (tblVBComponent_Declaration_Local, just Static Variables), grouped by vbcom_id,
              ' ** vbcomproc_id, with cnt_var), list of 1 module's procedures, with cnt_var, by specified [comid].
              Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_03_02_02")
            End Select
            With qdf.Parameters
              ![comid] = lngVBComID
            End With
            Set rst = qdf.OpenRecordset
            With rst
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
              ' **     2       1     vbcom_id              P_VID
              ' **     3       2     vbdecloc_module       P_VNAM
              ' **     4       3     vbcomproc_id          P_PID
              ' **     5       4     vbdecloc_procedure    P_PNAM
              ' **     6       5     proctype_type         P_PTYP
              ' **     7       6     procsubtype_type      P_STYP
              ' **     8       7     cnt_con, cnt_var      P_CCNT
              ' **
              ' *******************************************************
              .Close
            End With  ' ** rst.
            Set rst = Nothing
            Set qdf = Nothing

            If lngProcs > 0& Then

              Set cod = .CodeModule
              With cod
                ' ** For each documented procedure.
                For lngX = 0& To (lngProcs - 1&)

                  strProcName = arr_varProc(P_PNAM, lngX)
                  lngVBComProcID = arr_varProc(P_PID, lngX)

                  Select Case intMode
                  Case 1
                    ' ** zz_qry_VBComponent_Var_01_01_01 (tblVBComponent_Declaration_Local,
                    ' ** just Local Constants), list of 1 procedure's constants, by specified [procid].
                    Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_01_02_03")
                  Case 2
                    ' ** zz_qry_VBComponent_Var_02_01_01 (tblVBComponent_Declaration_Local,
                    ' ** just Local Variables), list of 1 procedure's constants, by specified [procid].
                    Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_02_02_03")
                  Case 3
                    ' ** zz_qry_VBComponent_Var_03_01_01 (tblVBComponent_Declaration_Local,
                    ' ** just Static Variables), list of 1 procedure's constants, by specified [procid].
                    Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_03_02_03")
                  End Select
                  With qdf.Parameters
                    ![procid] = lngVBComProcID
                  End With
                  Set rst = qdf.OpenRecordset
                  With rst
                    .MoveLast
                    lngItems = .RecordCount
                    .MoveFirst
                    arr_varItem = .GetRows(lngItems)
                    ' *******************************************************
                    ' ** Array: arr_varItem()
                    ' **
                    ' **   Field  Element  Name                  Constant
                    ' **   =====  =======  ====================  ==========
                    ' **     1       0     dbs_id                I_DID
                    ' **     2       1     vbcom_id              I_VID
                    ' **     3       2     vbdecloc_module       I_VNAM
                    ' **     4       3     vbcomproc_id          I_PID
                    ' **     5       4     vbdecloc_procedure    I_PNAM
                    ' **     6       5     vbdecloc_id           I_LID
                    ' **     7       6     vbdecloc_name         I_LNAM
                    ' **     8       7     vbdecloc_linenum      I_LINE
                    ' **
                    ' *******************************************************
                    .Close
                  End With  ' ** rst.
                  Set rst = Nothing
                  Set qdf = Nothing

                  Select Case arr_varProc(P_PTYP, lngX)
                  Case "Sub", "Function"
                    lngStartLine = .ProcBodyLine(strProcName, vbext_pk_Proc)
                    lngLines = .ProcCountLines(strProcName, vbext_pk_Proc)
                  Case "Property"
                    Select Case arr_varProc(P_STYP, lngX)
                    Case "Get"
                      lngStartLine = .ProcBodyLine(strProcName, vbext_pk_Get)
                      lngLines = .ProcCountLines(strProcName, vbext_pk_Get)
                    Case "Let"
                      lngStartLine = .ProcBodyLine(strProcName, vbext_pk_Let)
                      lngLines = .ProcCountLines(strProcName, vbext_pk_Let)
                    Case "Set"
                      lngStartLine = .ProcBodyLine(strProcName, vbext_pk_Set)
                      lngLines = .ProcCountLines(strProcName, vbext_pk_Set)
                    End Select
                  End Select
                  lngEndLine = ((lngStartLine - 1) + (lngLines - 1))

                  ' ** For each documented Constant or Variable.
                  For lngY = 0& To (lngItems - 1&)
                    ' ** We don't care about multiple hits on a line.
                    strFind = arr_varItem(I_LNAM, lngY)
                    intLen = Len(strFind)
                    lngCnt = 0&
                    For lngZ = lngStartLine To lngEndLine
                      strLine = Trim(.Lines(lngZ, 1))
                      If strLine <> vbNullString Then
                        If Left(strLine, 1) <> "'" Then
                          intPos01 = InStr(strLine, strFind)
                          If intPos01 > 0 Then
                            intPos02 = InStr(strLine, "' **")
                            If intPos02 > 0 And intPos02 < intPos01 Then
                              ' ** it's definitely in a Remark.
                            Else
                              If intPos01 > 1 Then
                                ' ** Check the first character before the hit.
                                strTmp01 = Mid(strLine, (intPos01 - 1), 1)
                                blnContinue = True
                                Select Case strTmp01
                                Case " ", "(", ",", ":", "-", "#", "="
                                  ' ** These are OK.
                                Case "[", "_", "'", Chr(34)
                                  ' ** Not OK.
                                  blnContinue = False
                                Case "."
                                  ' ** Iffy. User-Defined Type's will use periods.
                                  ' ** So let it go through
                                Case Else
                                  If (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Or _
                                      (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or _
                                      (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Then  ' ** Numbers and letters.
                                    ' ** Nope
                                    blnContinue = False
                                  Else
                                    ' ** Let's see what's there.
                                    Debug.Print "'  '" & strTmp01 & "'"
                                    Stop
                                  End If
                                End Select
                                If blnContinue = True Then
                                  If (intPos01 + intLen) > Len(strLine) Then
                                    ' ** Last word on the line.
                                  Else
                                    ' ** Now check the first character after the hit.
                                    strTmp01 = Mid(strLine, (intPos01 + intLen), 1)
                                    blnContinue = True
                                    Select Case strTmp01
                                    Case " ", "(", ")", ",", ":", "&", "!"
                                      ' ** These are OK.
                                    Case "]", "_", Chr(34)
                                      ' ** Not OK.
                                      blnContinue = False
                                    Case "."
                                      ' ** Iffy. User-Defined Type's will use periods.
                                      ' ** So let it go through
                                    Case Else
                                      If (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Or _
                                          (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or _
                                          (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Then  ' ** Numbers and letters.
                                        ' ** Nope
                                        blnContinue = False
                                      Else
                                        ' ** Let's see what's there.
                                        Debug.Print "'  '" & strTmp01 & "'"
                                        Stop
                                      End If
                                    End Select
                                  End If
                                End If  ' ** blnContinue.
                                If blnContinue = True Then
                                  lngHits = lngHits + 1&
                                  lngE = lngHits - 1&
                                  ReDim Preserve arr_varHit(H_ELEMS, lngE)
                                  arr_varHit(H_DID, lngE) = arr_varItem(I_DID, lngY)
                                  arr_varHit(H_VID, lngE) = arr_varItem(I_VID, lngY)
                                  arr_varHit(H_PID, lngE) = arr_varItem(I_PID, lngY)
                                  arr_varHit(H_PNAM, lngE) = arr_varItem(I_PNAM, lngY)
                                  arr_varHit(H_LID, lngE) = arr_varItem(I_LID, lngY)
                                  arr_varHit(H_LNAM, lngE) = arr_varItem(I_LNAM, lngY)
                                  arr_varHit(H_LINE, lngE) = lngZ
                                  lngCnt = lngCnt + 1&
                                End If  ' ** blnContinue.
                              Else
                                ' ** First word on line?
                                If Len(strLine) > Len(strFind) Then
                                  ' ** Check the first character after the hit.
                                  strTmp01 = Mid(strLine, (intLen + 1), 1)
                                  blnContinue = True
                                  Select Case strTmp01
                                  Case " ", "(", ")", ",", ":", "&", "!"
                                    ' ** These are OK.
                                  Case "]", "_", Chr(34)
                                    ' ** Not OK.
                                    blnContinue = False
                                  Case "."
                                    ' ** Iffy. User-Defined Type's will use periods.
                                    ' ** So let it go through
                                  Case Else
                                    If (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Or _
                                        (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or _
                                        (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Then  ' ** Numbers and letters.
                                      ' ** Nope
                                      blnContinue = False
                                    Else
                                      ' ** Let's see what's there.
                                      Debug.Print "'  '" & strTmp01 & "'"
                                      Stop
                                    End If
                                  End Select
                                  If blnContinue = True Then
                                    lngHits = lngHits + 1&
                                    lngE = lngHits - 1&
                                    ReDim Preserve arr_varHit(H_ELEMS, lngE)
                                    arr_varHit(H_DID, lngE) = arr_varItem(I_DID, lngY)
                                    arr_varHit(H_VID, lngE) = arr_varItem(I_VID, lngY)
                                    arr_varHit(H_PID, lngE) = arr_varItem(I_PID, lngY)
                                    arr_varHit(H_PNAM, lngE) = arr_varItem(I_PNAM, lngY)
                                    arr_varHit(H_LID, lngE) = arr_varItem(I_LID, lngY)
                                    arr_varHit(H_LNAM, lngE) = arr_varItem(I_LNAM, lngY)
                                    arr_varHit(H_LINE, lngE) = lngZ
                                    lngCnt = lngCnt + 1&
                                  End If  ' ** blnContinue.
                                Else
                                  ' ** Only word on line? Unlikely!
                                  lngHits = lngHits + 1&
                                  lngE = lngHits - 1&
                                  ReDim Preserve arr_varHit(H_ELEMS, lngE)
                                  arr_varHit(H_DID, lngE) = arr_varItem(I_DID, lngY)
                                  arr_varHit(H_VID, lngE) = arr_varItem(I_VID, lngY)
                                  arr_varHit(H_PID, lngE) = arr_varItem(I_PID, lngY)
                                  arr_varHit(H_PNAM, lngE) = arr_varItem(I_PNAM, lngY)
                                  arr_varHit(H_LID, lngE) = arr_varItem(I_LID, lngY)
                                  arr_varHit(H_LNAM, lngE) = arr_varItem(I_LNAM, lngY)
                                  arr_varHit(H_LINE, lngE) = lngZ
                                  lngCnt = lngCnt + 1&
                                End If
                              End If
                            End If
                          End If  ' ** intPos01.

                        End If  ' ** Remark.
                      End If  ' ** vbNullString.
                    Next  ' ** lngZ, lngLines.
                    If lngCnt = 0& Then
                      ' ** It should at least find its declaration.
                      'Stop
                      lngDels = lngDels + 1&
                    End If
                    lngV = lngV + 1&
                    If lngV Mod 1000& = 0& Then
                      Debug.Print "|  " & CStr(lngV) & " of " & CStr(lngTotItems)
                      Debug.Print "'|";
                    ElseIf lngV Mod 100& = 0& Then
                      Debug.Print "|";
                    ElseIf lngV Mod 10& = 0& Then
                      Debug.Print ".";
                    End If
                    DoEvents

                  Next  ' ** lngY, lngItems.

                Next  ' ** lngX, lngProcs.
              End With  ' ** cod.

            End If  ' ** lngProcs.

          End With  ' ** VBC.
          Set cod = Nothing

        Next  ' ** lngW, lngMods.
        Set vbc = Nothing
        Debug.Print
        DoEvents

      End With  ' ** vbp.
      Set vbp = Nothing

      Debug.Print "'HITS: " & CStr(lngHits)
      DoEvents

      If lngItems > 0& Then

        Debug.Print "'WRITING..."
        DoEvents

        Debug.Print "'|";
        DoEvents

        blnAddAll = False: blnAdd = False
        Set rst = .OpenRecordset("tblVBComponent_Declaration_Local_Detail", dbOpenDynaset, dbConsistent)
        With rst
          If .BOF = True And .EOF = True Then
            blnAddAll = True
          Else
            Select Case intMode
            Case 1
              varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_02", "[scopetype_type] = 'Local' And [dectype_type] = 'Constant'")
            Case 2
              varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_02", "[scopetype_type] = 'Local' And [dectype_type] = 'Variable'")
            Case 3
              varTmp00 = DCount("*", "zz_qry_VBComponent_Var_06_02", "[scopetype_type] = 'Static' ")
            End Select
            If varTmp00 = 0 Then
              blnAddAll = True
            End If
          End If
          For lngX = 0& To (lngHits - 1&)
            Select Case blnAddAll
            Case True
              blnAdd = True
            Case False
              .MoveFirst
              .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
                 "[vbcomproc_id] = " & CStr(arr_varHit(H_PID, lngX)) & " And [vbdecloc_id] = " & CStr(arr_varHit(H_LID, lngX)) & " And " & _
                 "[vbdeclocdet_linenum] = " & CStr(arr_varHit(H_LINE, lngX))
              If .NoMatch = True Then
                blnAdd = True
              End If
            End Select
            If blnAdd = True Then
              .AddNew
              ![vbdecloc_id] = arr_varHit(H_LID, lngX)
              ' ** ![vbdeclocdet_id] : AutoNumber.
              ![dbs_id] = arr_varHit(H_DID, lngX)
              ![vbcom_id] = arr_varHit(H_VID, lngX)
              ![vbcomproc_id] = arr_varHit(H_PID, lngX)
              Select Case intMode
              Case 1
                ![dectype_type] = "Constant"
              Case 2, 3
                ![dectype_type] = "Variable"
              End Select
              ![compdiropt_type] = 0&
              ![vbdeclocdet_linenum] = arr_varHit(H_LINE, lngX)
              ![vbdeclocdet_notused] = False
              ![vbdeclocdet_datemodified] = Now()
              .Update
            End If
            If (lngX + 1&) Mod 10000& = 0 Then
              Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngHits)
              Debug.Print "'|";
            ElseIf (lngX + 1&) Mod 1000& = 0 Then
              Debug.Print "|";
            ElseIf (lngX + 1&) Mod 100& = 0 Then
              Debug.Print ".";
            End If
            DoEvents
          Next  ' ** lngX.
          Debug.Print

          .Close
        End With  ' ** rst.
        Set rst = Nothing

        If lngDels > 0& Then
          Debug.Print "'ITEMS NOT FOUND: " & CStr(lngDels)
          DoEvents
        End If

      End If  ' ** lngItems.

    End If  ' ** lngMods.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_Var_Usage2 = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_Var_ChkProp() As Boolean
' ** Make sure all Get's, Let's, and Set's are separately listed.

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_Var_ChkProp"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngMods As Long, arr_varMod As Variant
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngLines As Long, lngDecLines As Long
  Dim lngVBComID As Long, lngNotFounds As Long
  Dim strModName As String, strLine As String
  Dim blnFound As Boolean
  Dim varTmp00 As Variant, varTmp01 As Variant, varTmp02 As Variant, varTmp03 As Variant
  Dim lngW As Long, lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varMod().
  'Const M_DID  As Integer = 0
  'Const M_DNAM As Integer = 1
  Const M_VID  As Integer = 2
  Const M_VNAM As Integer = 3
  'Const M_CTYP As Integer = 4
  'Const M_PCNT As Integer = 5

  ' ** Array: arr_varProc().
  'Const P_DID  As Integer = 0
  'Const P_DNAM As Integer = 1
  'Const P_VID  As Integer = 2
  'Const P_VNAM As Integer = 3
  'Const P_CTYP As Integer = 4
  'Const P_PID  As Integer = 5
  Const P_PNAM As Integer = 6
  Const P_PTYP As Integer = 7
  Const P_STYP As Integer = 8
  Const P_BEG  As Integer = 9
  'Const P_END  As Integer = 10
  Const P_FND  As Integer = 11
  Const P_CHG  As Integer = 12

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** zz_qry_VBComponent_25_01 (tblVBComponent_Procedure,
    ' ** just Property's), grouped by vbcom_id, with cnt_proc.
    Set qdf = .QueryDefs("zz_qry_VBComponent_Var_25_02")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngMods = .RecordCount
      .MoveFirst
      arr_varMod = .GetRows(lngMods)
      ' *************************************************
      ' ** Array: arr_varMod()
      ' **
      ' **   Field  Element  Name            Constant
      ' **   =====  =======  ==============  ==========
      ' **     1       0     dbs_id          M_DID
      ' **     2       1     dbs_name        M_DNAM
      ' **     3       2     vbcom_id        M_VID
      ' **     4       3     vbcom_name      M_VNAM
      ' **     5       4     comtype_type    M_CTYP
      ' **     6       5     cnt_proc        M_PCNT
      ' **
      ' *************************************************
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Set qdf = Nothing

    Debug.Print "'MODS: " & CStr(lngMods)
    DoEvents

    If lngMods > 0& Then

      lngNotFounds = 0&
      Set vbp = Application.VBE.ActiveVBProject
      With vbp
        For lngW = 0& To (lngMods - 1&)
          lngVBComID = arr_varMod(M_VID, lngW)
          Set vbc = .VBComponents(arr_varMod(M_VNAM, lngW))
          With vbc
            strModName = .Name
            ' ** zz_qry_VBComponent_Var_25_01 (tblVBComponent_Procedure, just
            ' ** Property's), just 1 module's procedures, by specified [comid].
            Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_25_03")
            With qdf.Parameters
              ![comid] = lngVBComID
            End With
            Set rst = qdf.OpenRecordset
            With rst
              .MoveLast
              lngProcs = .RecordCount
              .MoveFirst
              arr_varProc = .GetRows(lngProcs)
              ' ******************************************************
              ' ** Array: arr_varProc()
              ' **
              ' **   Field  Element  Name                 Constant
              ' **   =====  =======  ===================  ==========
              ' **     1       0     dbs_id               P_DID
              ' **     2       1     dbs_name             P_DNAM
              ' **     3       2     vbcom_id             P_VID
              ' **     4       3     vbcom_name           P_VNAM
              ' **     5       4     comtype_type         P_CTYP
              ' **     6       5     vbcomproc_id         P_PID
              ' **     7       6     vbcomproc_name       P_PNAM
              ' **     8       7     proctype_type        P_PTYP
              ' **     9       8     procsubtype_type     P_STYP
              ' **    10       9     vbcomproc_line_beg   P_BEG
              ' **    11      10     vbcomproc_line_end   P_END
              ' **    12      11     proc_fnd             P_FND
              ' **    13      12     proc_chg             P_CHG
              ' **
              ' ******************************************************
              .Close
            End With  ' ** rst.
            Set rst = Nothing
            Set qdf = Nothing

            Set cod = .CodeModule
            With cod

              lngLines = .CountOfLines
              lngDecLines = .CountOfDeclarationLines

              For lngX = lngDecLines To lngLines
                strLine = Trim(.Lines(lngX, 1))
                If strLine <> vbNullString Then
                  If Left(strLine, 1) <> "'" Then
                    varTmp00 = GetNthWord(strLine, 1)  ' ** Module Function: modStringFuncs.
                    If IsNull(varTmp00) = False Then
                      If varTmp00 = "Public" Or varTmp00 = "Private" Then
                        varTmp01 = GetNthWord(strLine, 2)  ' ** Module Function: modStringFuncs.
                        If IsNull(varTmp01) = False Then
                          If varTmp01 = "Property" Then
                            ' ** Will it be Get, Let, or Set?
                            varTmp02 = GetNthWord(strLine, 3)  ' ** Module Function: modStringFuncs.
                            varTmp03 = GetNthWord(strLine, 4)  ' ** Module Function: modStringFuncs.
                            If InStr(varTmp03, "(") > 0 Then varTmp03 = Left(varTmp03, (InStr(varTmp03, "(") - 1))
                            blnFound = False
                            For lngY = 0& To (lngProcs - 1&)
                              If arr_varProc(P_PNAM, lngY) = varTmp03 And arr_varProc(P_PTYP, lngY) = varTmp01 Then
                                ' ** It's a Property with the same name.
                                If arr_varProc(P_STYP, lngY) = varTmp02 Then
                                  ' ** And the same subtype.
                                  blnFound = True
                                  arr_varProc(P_FND, lngY) = CBool(True)
                                  If arr_varProc(P_BEG, lngY) <> lngY Then
                                    arr_varProc(P_BEG, lngY) = lngY
                                    arr_varProc(P_CHG, lngY) = CBool(True)
                                  End If
                                  Exit For
                                End If
                              End If
                            Next  ' ** lngY
                            If blnFound = False Then
                              Debug.Print "'NOT FOUND!  " & varTmp00 & " " & varTmp01 & " " & varTmp02 & "  " & varTmp03 & "()"
                              DoEvents
                              lngNotFounds = lngNotFounds + 1&
                            End If  ' ** blnFound.
                          End If  ' ** Property.
                        End If  ' ** IsNull().
                      End If  ' ** Public, Private.
                    End If  ' ** IsNull().
                  End If  ' ** Remark.
                End If  ' ** vbNullString.
              Next  ' ** lngX.

            End With  ' ** cod.
            Set cod = Nothing
          End With  ' ** vbc.
        Next  ' ** lngW.
        Set vbc = Nothing
      End With  ' ** vbp.
      Set vbp = Nothing

      If lngNotFounds > 0& Then
        Debug.Print "'NOT FOUND: " & CStr(lngNotFounds)
        DoEvents
      Else
        Debug.Print "'ALL TYPES FOUND AND LISTED!"
        DoEvents
      End If

    End If  ' ** lngMods

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

'MODS: 4
'ALL TYPES FOUND AND LISTED!
'DONE!
  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_Var_ChkProp = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_Var_Delete() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_Var_Delete"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngItems As Long, arr_varItem As Variant
  Dim strLastModName As String, strLine As String, strFind As String
  Dim lngLine As Long, lngDels As Long, lngEdits As Long, lngDecLines As Long
  Dim blnContinue As Boolean
  Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngY As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varItem().
  Const I_LID  As Integer = 0
  Const I_VNAM As Integer = 1
  Const I_PNAM As Integer = 2
  Const I_LINE As Integer = 3
  Const I_LNAM As Integer = 4
  Const I_DTYP As Integer = 5
  Const I_MARK As Integer = 6

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** tblMark_AutoNum, linked to tblVBComponent_Declaration_Local_Detail, just needed fields.
    Set qdf = .QueryDefs("zz_qry_VBComponent_Var_29_05")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngItems = .RecordCount
      .MoveFirst
      arr_varItem = .GetRows(lngItems)
      ' ********************************************************
      ' ** Array: arr_varItem()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     vbdecloc_id            I_LID
      ' **     2       1     vbdecloc_module        I_VNAM
      ' **     3       2     vbdecloc_procedure     I_PNAM
      ' **     4       3     vbdeclocdet_linenum    I_LINE
      ' **     5       4     vbdecloc_name          I_LNAM
      ' **     6       5     dectype_type           I_DTYP
      ' **     7       6     mark                   I_MARK
      ' **
      ' ********************************************************
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Set qdf = Nothing

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      strLastModName = vbNullString
      lngDels = 0&: lngEdits = 0&
      For lngX = 0& To (lngItems - 1&)
        strFind = arr_varItem(I_LNAM, lngX)
        If arr_varItem(I_VNAM, lngX) <> strLastModName Then
          If strLastModName <> vbNullString Then
            With vbc.CodeModule
              lngDecLines = .CountOfDeclarationLines
              For lngY = 1& To lngDecLines
                strLine = .Lines(lngY, 1)
                If Trim(strLine) <> vbNullString Then
                  If Left(strLine, 4) = "'VGC" And Right(strLine, 8) = "CHANGES!" Then
                    strTmp01 = "'VGC 11/23/2016: CHANGES!"
                    .ReplaceLine lngY, strTmp01
                    Exit For
                  End If
                End If
              Next  ' ** lngY.
            End With
          End If
          Set cod = Nothing
          Set vbc = Nothing
'If strLastModName = "modTransactionAuditFuncs1" Then
'Debug.Print "'STOP 1!"
'DoEvents
'Stop
'Exit For
'End If
          Set vbc = .VBComponents(arr_varItem(I_VNAM, lngX))
          strLastModName = arr_varItem(I_VNAM, lngX)
        End If
        With vbc
          Set cod = .CodeModule
          With cod
            blnContinue = True
            lngLine = arr_varItem(I_LINE, lngX)
            strLine = .Lines(lngLine, 1)
            intPos01 = InStr(strLine, strFind)
            If intPos01 = 0 Then
              lngLine = (arr_varItem(I_LINE, lngX) - 1&)
              strLine = .Lines(lngLine, 1)
              intPos01 = InStr(strLine, strFind)
              If intPos01 = 0 Then
                lngLine = (arr_varItem(I_LINE, lngX) + 1&)
                strLine = .Lines(lngLine, 1)
                intPos01 = InStr(strLine, strFind)
                If intPos01 = 0 Then
'Debug.Print "'STOP 1!"
                  blnContinue = False
                  Debug.Print "'NOT FOUND! " & strFind & "  " & strLastModName & "  " & arr_varItem(I_PNAM, lngX)
                  DoEvents
                  'Stop
                End If
              End If
            End If
            If blnContinue = True Then
              strTmp01 = Left(strLine, (intPos01 - 1))  ' ** Everything up to, but not including, our target.
              strTmp02 = Mid(strLine, intPos01)         ' ** Target to end of line.
              strTmp03 = vbNullString
              intPos02 = InStr(strTmp02, " As ")
              If intPos02 > 0 Then
                If arr_varItem(I_DTYP, lngX) = "Constant" Then
                  intPos03 = InStr(strTmp02, "=")
                  If intPos03 > 0 Then
                    ' ** Since all Constants are on their own line, this is enough.
                    'NOT YET PROGRAMMED!
Debug.Print "'STOP 2!"
DoEvents
                    Stop
                  Else
Debug.Print "'STOP 3!"
DoEvents
                    Stop
                  End If
                Else
                  ' ** Variable.
                  intPos03 = InStr(intPos02, strTmp02, ",")  ' ** Looking for comma after ' As '.
                  strTmp03 = vbNullString
                  If intPos03 > 0 Then
                    ' ** Comma found, so may be more variables after ours.
                    intPos04 = InStr(strTmp02, "'")
                    If intPos04 > 0 Then
                      If intPos04 < intPos03 Then
                        ' ** The comma is within a remark.
                        strTmp03 = Mid(strTmp02, intPos04)               ' ** Remark only, after target variable.
                        strTmp02 = Trim(Left(strTmp02, (intPos04 - 1)))  ' ** This should be our variable alone.
                      Else
                        strTmp03 = Mid(strTmp02, intPos03)               ' ** From comma to end of line.
                        strTmp02 = Left(strTmp02, (intPos03 - 1))        ' ** This should be our variable alone.
                      End If
                    Else
                      ' ** No remarks on the line.
                      strTmp03 = Mid(strTmp02, intPos03)                 ' ** From comma to end of the line.
                      strTmp02 = Left(strTmp02, (intPos03 - 1))          ' ** This should be our variable alone.
                    End If
                  Else
                    ' ** No comma found after our variable.
                    intPos04 = InStr(strTmp02, "'")
                    If intPos04 > 0 Then
                      strTmp03 = Mid(strTmp02, intPos04)                 ' ** Remark only, after target variable.
                      strTmp02 = Trim(Left(strTmp02, (intPos04 - 1)))    ' ** This should be our variable alone.
                    Else
                      ' ** strTmp02 should be our target variable, alone.
                    End If
                  End If
                  If Trim(strTmp01) = "Dim" Or Trim(strTmp01) = "Static" Then
                    ' ** Target was first on the line.
                    If strTmp03 = vbNullString Then
                      ' ** Target only thing on the line; trash it!
                      strTmp01 = "{DELETE}"                    ' ** Delete line.
                      strTmp02 = vbNullString: strTmp03 = vbNullString
                    Else
                      If Left(strTmp03, 1) = "," Then
                        ' ** Another declaration, but now it's first.
                        strTmp03 = Trim(Mid(strTmp03, 2))      ' ** Delete comma.
                        strTmp01 = strTmp01 & strTmp03         ' ** Replace line.
                      Else
                        ' ** Must be a remark, but there's no longer any variable declaration.
                        strTmp01 = strTmp03                    ' ** Replace line.
                      End If
                    End If
                  Else
                    If strTmp03 = vbNullString Then
                      ' ** Target was last on line, and no remark.
                      If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
                      ' ** Replace line.
                    Else
                      If Left(strTmp03, 1) = "," Then
                        ' ** Another declaration.
                        If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
                        strTmp01 = strTmp01 & strTmp03         ' ** Replace line.
                      Else
                        ' ** Must be a remark.
                        If Right(strTmp01, 1) = "," Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
                        strTmp01 = strTmp01 & "  " & strTmp03  ' ** Replace line.
                      End If
                    End If
                  End If
                End If  ' ** Constant or Variable.
                ' ** Fixes we shouldn't need!
                If strTmp01 <> "{DELETE}" And (Trim(strTmp01) = "Dim" Or Trim(strTmp01) = "Static") Then
                  strTmp01 = "{DELETE}"
                ElseIf Right(Trim(strTmp01), 1) = "," Then
                  strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
                  If Right(Trim(strTmp01), 1) = "," Then
                    strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
                    If Right(Trim(strTmp01), 1) = "," Then
                      strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
                    End If
                  End If
                ElseIf InStr(strTmp01, ", ,") > 0 Then
                  strTmp01 = StringReplace(strTmp01, ", ,", ",") ' ** Module Function: modStringFuncs.
                End If
                ' ** Now update.
                If strTmp01 = "{DELETE}" Then
                  .DeleteLines lngLine, 1
                  arr_varItem(I_MARK, lngX) = CBool(True)
                  lngDels = lngDels + 1&
                  ' ** Update tblMark_AutoNum, by specified [locid].
                  Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_29_06")
                  With qdf.Parameters
                    ![locid] = arr_varItem(I_LID, lngX)
                  End With
                  qdf.Execute
                  Set qdf = Nothing
                Else
                  ' ** We should have a good replacement line.
                  If Trim(strTmp01) = vbNullString Then
Debug.Print "'STOP 4!"
DoEvents
                    Stop
                  ElseIf InStr(strTmp01, strFind) > 0 Then
Debug.Print "'STOP 5!"
DoEvents
                    Stop
                  ElseIf Right(strTmp01, 1) = "," Then
Debug.Print "'STOP 6!"
DoEvents
                    Stop
                  Else
                    .ReplaceLine lngLine, strTmp01
                    arr_varItem(I_MARK, lngX) = CBool(True)
                    lngEdits = lngEdits + 1&
                    ' ** Update tblMark_AutoNum, by specified [locid].
                    Set qdf = dbs.QueryDefs("zz_qry_VBComponent_Var_29_06")
                    With qdf.Parameters
                      ![locid] = arr_varItem(I_LID, lngX)
                    End With
                    qdf.Execute
                    Set qdf = Nothing
                  End If
                End If
              Else
Debug.Print "'STOP 7!"
DoEvents
                Stop
              End If
            End If  ' ** blnContinue.
          End With  ' ** cod.
        End With  ' ** vbc.
      Next  ' ** lngX.
      Set cod = Nothing
      Set vbc = Nothing
    End With  ' ** vbp.
    Set vbp = Nothing

    .Close
  End With  ' ** dbs
  Set dbs = Nothing

  Debug.Print "'DELETES: " & CStr(lngDels)
  Debug.Print "'EDITS  : " & CStr(lngEdits)
  DoEvents

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_Var_Delete = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_CompilerDirective_Doc() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_CompilerDirective_Doc"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset
  Dim lngItems As Long, arr_varItem() As Variant
  Dim strModName As String, strProcName As String, strLine As String, strFind As String
  Dim lngLines As Long, lngDecLines As Long
  Dim lngThisDbsID As Long, lngVBComID As Long, lngVBComProcID As Long, lngCompDirType As Long
  Dim blnAddAll As Boolean, blnAdd As Boolean
  Dim intPos01 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varItem().
  Const I_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const I_DID  As Integer = 0   'dbs_id
  Const I_VID  As Integer = 1   'vbcom_id
  Const I_VNAM As Integer = 2   'vbcom_name
  Const I_PID  As Integer = 3   'vbcomproc_id
  Const I_PNAM As Integer = 4   'vbcomproc_name
  Const I_CON  As Integer = 5   'compdir_constant
  Const I_TYP  As Integer = 6   'compdir_type
  Const I_VAL  As Integer = 7   'vbcomdir_value
  Const I_DSC  As Integer = 8   'vbcomdir_description
  Const I_LIN  As Integer = 9   'vbcomdir_linenum
  Const I_RAW  As Integer = 10  'vbcomdir_raw

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  strFind = "#Const"

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngItems = 0&
  ReDim arr_varItem(I_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp

    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        Set cod = .CodeModule
        With cod

          lngLines = .CountOfLines
          lngDecLines = .CountOfDeclarationLines
          strProcName = "Declaration"

          For lngX = 1& To lngDecLines
            strLine = Trim(.Lines(lngX, 1))
            strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
            If strLine <> vbNullString Then
              If Left(strLine, 1) <> "'" Then
                intPos01 = InStr(strLine, strFind)
                If intPos01 > 0 Then  ' ** Should always be 1.
                  ' ** #Const IsDev = 0  ' ** 0 = release, -1 = development.
                  strTmp01 = GetNthWord(strLine, 2)  ' ** Module Function: modStringFuncs.
                  strTmp02 = GetNthWord(strLine, 4)  ' ** Module Function: modStringFuncs.
                  intPos01 = InStr(strLine, "'")
                  If intPos01 > 0 Then
                    strTmp03 = Mid(strLine, intPos01)
                  End If
                  lngItems = lngItems + 1&
                  lngE = lngItems - 1&
                  ReDim Preserve arr_varItem(I_ELEMS, lngE)
                  ' *********************************************************
                  ' ** Array: arr_varItem()
                  ' **
                  ' **   Field  Element  Name                    Constant
                  ' **   =====  =======  ======================  ==========
                  ' **     1       0     dbs_id                  I_DID
                  ' **     2       1     vbcom_id                I_VID
                  ' **     3       2     vbcom_name              I_VNAM
                  ' **     4       3     vbcomproc_id            I_PID
                  ' **     5       4     vbcomproc_name          I_PNAM
                  ' **     6       5     compdir_constant        I_CON
                  ' **     7       6     compdir_type            I_TYP
                  ' **     8       7     vbcomdir_value          I_VAL
                  ' **     9       8     vbcomdir_description    I_DSC
                  ' **    10       9     vbcomdir_linenum        I_LIN
                  ' **    11      10     vbcomdir_raw            I_RAW
                  ' **
                  ' *********************************************************
                  arr_varItem(I_DID, lngE) = lngThisDbsID
                  arr_varItem(I_VID, lngE) = Null
                  arr_varItem(I_VNAM, lngE) = strModName
                  arr_varItem(I_PID, lngE) = Null
                  arr_varItem(I_PNAM, lngE) = strProcName
                  arr_varItem(I_CON, lngE) = strTmp01
                  arr_varItem(I_TYP, lngE) = Null
                  arr_varItem(I_VAL, lngE) = strTmp02
                  If strTmp03 <> vbNullString Then
                    arr_varItem(I_DSC, lngE) = strTmp03
                  Else
                    arr_varItem(I_DSC, lngE) = Null
                  End If
                  arr_varItem(I_LIN, lngE) = lngX
                  arr_varItem(I_RAW, lngE) = strLine
                End If
              End If  ' ** Remark.
            End If  ' ** vbNullString.
          Next  ' ** lngX.

        End With  ' ** cod
      End With  ' ** vbc.
    Next  ' ** vbc.
    Set cod = Nothing
    Set vbc = Nothing

  End With  ' ** vbp
  Set vbp = Nothing

  Debug.Print "'DIRECTIVES: " & CStr(lngItems)
  DoEvents

  If lngItems > 0& Then

    Set dbs = CurrentDb
    With dbs

      Set rst1 = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      Set rst2 = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
      Set rst3 = .OpenRecordset("tblCompilerDirective", dbOpenDynaset, dbReadOnly)
      For lngX = 0& To (lngItems - 1&)
        lngVBComID = 0&: lngVBComProcID = 0&: lngCompDirType = 0&
        With rst1
          .MoveFirst
          .FindFirst "[dbs_id] = " & CStr(arr_varItem(I_DID, lngX)) & " And [vbcom_name] = '" & arr_varItem(I_VNAM, lngX) & "'"
          If .NoMatch = False Then
            lngVBComID = ![vbcom_id]
          Else
            Stop
          End If
        End With  ' ** rst1.
        arr_varItem(I_VID, lngX) = lngVBComID
        With rst2
          .MoveFirst
          .FindFirst "[dbs_id] = " & CStr(arr_varItem(I_DID, lngX)) & " And [vbcom_id] = " & CStr(lngVBComID) & " And " & _
            "[vbcomproc_name] = '" & arr_varItem(I_PNAM, lngX) & "'"
          If .NoMatch = False Then
            lngVBComProcID = ![vbcomproc_id]
          Else
            Stop
          End If
        End With  ' ** rst2.
        arr_varItem(I_PID, lngX) = lngVBComProcID
        With rst3
          .MoveFirst
          .FindFirst "[compdir_constant] = '" & arr_varItem(I_CON, lngX) & "'"
          If .NoMatch = False Then
            lngCompDirType = ![compdir_type]
          Else
            Stop
          End If
        End With  ' ** rst3.
        arr_varItem(I_TYP, lngX) = lngCompDirType
      Next  ' ** lngX.
      rst1.Close
      rst2.Close
      rst3.Close
      Set rst1 = Nothing
      Set rst2 = Nothing
      Set rst3 = Nothing

      Set rst1 = .OpenRecordset("tblVBComponent_Directive", dbOpenDynaset, dbConsistent)
      With rst1
        blnAddAll = False: blnAdd = False
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        End If
        For lngX = 0& To (lngItems - 1&)
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            .FindFirst "[dbs_id] = " & CStr(arr_varItem(I_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varItem(I_VID, lngX)) & " And " & _
              "[vbcomproc_id] = " & CStr(arr_varItem(I_PID, lngX))
            If .NoMatch = False Then
              If ![compdir_constant] <> arr_varItem(I_CON, lngX) Then
                .Edit
                ![compdir_constant] = arr_varItem(I_CON, lngX)
                ![vbcomdir_datemodified] = Now()
                .Update
              End If
              If ![compdir_type] <> arr_varItem(I_TYP, lngX) Then
                .Edit
                ![compdir_type] = arr_varItem(I_TYP, lngX)
                ![vbcomdir_datemodified] = Now()
                .Update
              End If
              If ![vbcomdir_value] <> arr_varItem(I_VAL, lngX) Then
                .Edit
                ![vbcomdir_value] = arr_varItem(I_VAL, lngX)
                ![vbcomdir_datemodified] = Now()
                .Update
              End If
              If ![vbcomdir_description] <> arr_varItem(I_DSC, lngX) Then
                .Edit
                ![vbcomdir_description] = arr_varItem(I_DSC, lngX)
                ![vbcomdir_datemodified] = Now()
                .Update
              End If
              If ![vbcomdir_linenum] <> arr_varItem(I_LIN, lngX) Then
                .Edit
                ![vbcomdir_linenum] = arr_varItem(I_LIN, lngX)
                ![vbcomdir_datemodified] = Now()
                .Update
              End If
              If ![vbcomdir_raw] <> arr_varItem(I_RAW, lngX) Then
                .Edit
                ![vbcomdir_raw] = arr_varItem(I_RAW, lngX)
                ![vbcomdir_datemodified] = Now()
                .Update
              End If
            Else
              blnAdd = True
            End If
          End Select  ' ** blnAddAll.
          If blnAdd = True Then
            .AddNew
            ![dbs_id] = arr_varItem(I_DID, lngX)
            ![vbcom_id] = arr_varItem(I_VID, lngX)
            ![vbcomproc_id] = arr_varItem(I_PID, lngX)
            ' ** ![vbcomdir_id] : AutoNumber.
            ![vbcomdir_module] = arr_varItem(I_VNAM, lngX)
            ![vbcomdir_procedure] = arr_varItem(I_PNAM, lngX)
            ![compdir_constant] = arr_varItem(I_CON, lngX)
            ![compdir_type] = arr_varItem(I_TYP, lngX)
            ![vbcomdir_value] = arr_varItem(I_VAL, lngX)
            ![vbcomdir_description] = arr_varItem(I_DSC, lngX)
            ![vbcomdir_linenum] = arr_varItem(I_LIN, lngX)
            ![vbcomdir_raw] = arr_varItem(I_RAW, lngX)
            ![vbcomdir_datemodified] = Now()
            .Update
          End If  ' ** blnAdd.
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

  End If  ' ** lngItems.

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_CompilerDirective_Doc = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_CompilerDirective_Usage() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_CompilerDirective_Usage"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngDirs As Long, arr_varDir As Variant
  Dim lngItems As Long, arr_varItem() As Variant
  Dim strModName As String, strProcName As String, strLine As String
  Dim strCompDir1 As String, strCompDir2 As String, strCompDirStat1 As String, strCompDirStat2 As String
  Dim lngLines As Long, lngDecLines As Long, lngLine As Long
  Dim lngCompDirType1 As Long, lngCompDirType2 As Long, lngCompDirOpt1 As Long, lngCompDirOpt2 As Long
  Dim intMode As Integer
  Dim blnAddAll As Boolean, blnAdd As Boolean, blnIsNeg As Boolean
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varDir().
  Const D_DID  As Integer = 0
  Const D_VID  As Integer = 1
  Const D_VNAM As Integer = 2
  'Const D_PID  As Integer = 3
  'Const D_PNAM As Integer = 4
  Const D_CID1 As Integer = 5
  'Const D_CON1 As Integer = 6
  'Const D_TYP1 As Integer = 7
  'Const D_LIN1 As Integer = 8
  Const D_CID2 As Integer = 9
  'Const D_CON2 As Integer = 10
  'Const D_TYP2 As Integer = 11
  'Const D_LIN2 As Integer = 12

  ' ** Array: arr_varItem().
  Const I_ELEMS As Integer = 12  ' ** Array's first-element UBound().
  Const I_DID  As Integer = 0
  Const I_VID  As Integer = 1
  Const I_VNAM As Integer = 2
  Const I_PID  As Integer = 3
  Const I_PNAM As Integer = 4
  Const I_CID  As Integer = 5
  Const I_STMT As Integer = 6
  Const I_TYP  As Integer = 7
  Const I_CON  As Integer = 8
  Const I_OPT  As Integer = 9
  Const I_LIN  As Integer = 10
  Const I_NOT  As Integer = 11
  Const I_RAW  As Integer = 12

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** tblVBComponent_Directive, just needed fields.
    Set qdf = .QueryDefs("zz_qry_VBComponent_Var_04_08")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngDirs = .RecordCount
      .MoveFirst
      arr_varDir = .GetRows(lngDirs)
      ' *******************************************************
      ' ** Array: arr_varDir()
      ' **
      ' **   Field  Element  Name                  Constant
      ' **   =====  =======  ====================  ==========
      ' **     1       0     dbs_id                D_DID
      ' **     2       1     vbcom_id              D_VID
      ' **     3       2     vbcomdir_module       D_VNAM
      ' **     4       3     vbcomproc_id          D_PID
      ' **     5       4     vbcomdir_procedure    D_PNAM
      ' **     6       5     vbcomdir_id1          D_CID1
      ' **     7       6     compdir_constant1     D_CON1
      ' **     8       7     compdir_type1         D_TYP1
      ' **     9       8     vbcomdir_linenum1     D_LIN1
      ' **    10       9     vbcomdir_id2          D_CID2
      ' **    11      10     compdir_constant2     D_CON2
      ' **    12      11     compdir_type2         D_TYP2
      ' **    13      12     vbcomdir_linenum2     D_LIN2
      ' **
      ' *******************************************************
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Set qdf = Nothing

    Debug.Print "'Declarations: " & CStr(lngDirs)
    DoEvents

    lngItems = 0&
    ReDim arr_varItem(I_ELEMS, 0)

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For lngX = 0& To (lngDirs - 1&)
        Set vbc = .VBComponents(arr_varDir(D_VNAM, lngX))
        With vbc
          strCompDir1 = vbNullString: strCompDir2 = vbNullString: strCompDirStat1 = vbNullString: strCompDirStat2 = vbNullString
          lngCompDirOpt1 = 0&: lngCompDirOpt2 = 0&: intMode = 0
          strModName = .Name
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngY = 1& To lngLines
              strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString
              blnIsNeg = False
              strLine = Trim(.Lines(lngY, 1))
              If strLine <> vbNullString Then
                If Left(strLine, 1) <> "'" Then
                  If Left(strLine, 1) = "#" Then  ' ** Compiler Directive indicator.
                    ' ** #Const
                    ' ** #If
                    ' ** #ElseIf
                    ' ** #Else
                    ' ** #End If

                    strProcName = .ProcOfLine(lngY, vbext_pk_Proc)
                    If strProcName = vbNullString Then strProcName = "Declaration"
                    strTmp01 = GetNthWord(strLine, 1)  ' ** Module Function: modStringFuncs.
                    Select Case strTmp01
                    Case "#Const"
                      ' ** #Const IsDev = 0
                      strTmp02 = GetNthWord(strLine, 2)  ' ** Module Function: modStringFuncs.
                      If strCompDir1 = vbNullString Then
                        strCompDir1 = strTmp02
                        Select Case strCompDir1
                        Case "IsDev"
                          lngCompDirType1 = 1&
                          lngCompDirOpt1 = 1&
                        Case "IsDemo"
                          lngCompDirType1 = 2&
                          lngCompDirOpt1 = 5&
                        Case "NoExcel"
                          lngCompDirType1 = 3&
                          lngCompDirOpt1 = 9&
                        Case "HasRepost"
                          lngCompDirType1 = 4&
                          lngCompDirOpt1 = 13&
                        End Select
                        intMode = 1
                      Else
                        strCompDir2 = strTmp02
                        Select Case strCompDir2
                        Case "IsDev"
                          lngCompDirType2 = 1&
                          lngCompDirOpt2 = 1&
                        Case "IsDemo"
                          lngCompDirType2 = 2&
                          lngCompDirOpt2 = 5&
                        Case "NoExcel"
                          lngCompDirType2 = 3&
                          lngCompDirOpt2 = 9&
                        Case "HasRepost"
                          lngCompDirType2 = 4&
                          lngCompDirOpt2 = 13&
                        End Select
                        intMode = 2
                      End If
                      lngLine = lngY
                    Case "#If"
                      ' ** #If NoExcel Then
                      strTmp02 = GetNthWord(strLine, 2)  ' ** Module Function: modStringFuncs.
                      If strTmp02 = "Not" Then
                        blnIsNeg = True
                        strTmp02 = GetNthWord(strLine, 3)  ' ** Module Function: modStringFuncs.
                        Select Case strTmp02
                        Case "IsDev", "IsDemo", "NoExcel", "HasRepost"
                          lngLine = lngY
                          If strCompDir1 = strTmp02 Then
                            strCompDirStat1 = "OPEN " & strCompDir1 & " Not " & CStr(lngLine)
                            Select Case lngCompDirType1
                            Case 1&
                              lngCompDirOpt1 = 3&
                            Case 2&
                              lngCompDirOpt1 = 7&
                            Case 3&
                              lngCompDirOpt1 = 11&
                            Case 4&
                              lngCompDirOpt1 = 15&
                            End Select
                            intMode = 1
                          ElseIf strCompDir2 = strTmp02 Then
                            strCompDirStat2 = "OPEN " & strCompDir2 & " Not " & CStr(lngLine)
                            Select Case lngCompDirType2
                            Case 1&
                              lngCompDirOpt2 = 3&
                            Case 2&
                              lngCompDirOpt2 = 7&
                            Case 3&
                              lngCompDirOpt2 = 11&
                            Case 4&
                              lngCompDirOpt2 = 15&
                            End Select
                            intMode = 2
                          Else
                            Stop
                          End If
                        Case Else
                          Stop
                        End Select
                      Else
                        Select Case strTmp02
                        Case "IsDev", "IsDemo", "NoExcel", "HasRepost"
                          lngLine = lngY
                          If strCompDir1 = strTmp02 Then
                            strCompDirStat1 = "OPEN " & strCompDir1 & " " & CStr(lngLine)
                            Select Case lngCompDirType1
                            Case 1&
                              lngCompDirOpt1 = 2&
                            Case 2&
                              lngCompDirOpt1 = 6&
                            Case 3&
                              lngCompDirOpt1 = 10&
                            Case 4&
                              lngCompDirOpt1 = 14&
                            End Select
                            intMode = 1
                          ElseIf strCompDir2 = strTmp02 Then
                            strCompDirStat2 = "OPEN " & strCompDir2 & " " & CStr(lngLine)
                            Select Case lngCompDirType2
                            Case 1&
                              lngCompDirOpt2 = 2&
                            Case 2&
                              lngCompDirOpt2 = 6&
                            Case 3&
                              lngCompDirOpt2 = 10&
                            Case 4&
                              lngCompDirOpt2 = 14&
                            End Select
                            intMode = 2
                          Else
                            Stop
                          End If
                        Case Else
                          Stop
                        End Select
                      End If  ' ** Not.
                    Case "#ElseIf"
                      ' ** Currently not being used.
                    Case "#Else"
                      lngLine = lngY
                      If strCompDirStat1 <> vbNullString And strCompDirStat2 <> vbNullString Then
                        ' ** 2 compiler directives are in use.
                        strTmp03 = GetLastWord(strCompDirStat1)  ' ** Module Function: modStringFuncs.
                        strTmp04 = GetLastWord(strCompDirStat2)  ' ** Module Function: modStringFuncs.
                        If Val(strTmp03) < Val(strTmp04) Then
                          ' ** We're within 2.
                          strTmp02 = GetNthWord(strCompDirStat2, 2)  ' ** Module Function: modStringFuncs.
                          strCompDirStat2 = "FLIP" & Mid(strCompDirStat2, 5)
                          strCompDirStat2 = Left(strCompDirStat2, (Len(strCompDirStat2) - Len(strTmp04)))  ' ** Strip previous line number.
                          strCompDirStat2 = strCompDirStat2 & CStr(lngLine)
                          If InStr(strCompDirStat2, " Not ") > 0 Then
                            blnIsNeg = True
                            Select Case lngCompDirType2
                            Case 1&
                              lngCompDirOpt2 = 2&
                            Case 2&
                              lngCompDirOpt2 = 6&
                            Case 3&
                              lngCompDirOpt2 = 10&
                            Case 4&
                              lngCompDirOpt2 = 14&
                            End Select
                          Else
                            Select Case lngCompDirType2
                            Case 1&
                              lngCompDirOpt2 = 3&
                            Case 2&
                              lngCompDirOpt2 = 7&
                            Case 3&
                              lngCompDirOpt2 = 11&
                            Case 4&
                              lngCompDirOpt2 = 15&
                            End Select
                          End If
                          intMode = 2
                        Else
                          ' ** We're within 1.
                          strTmp02 = GetNthWord(strCompDirStat1, 2)  ' ** Module Function: modStringFuncs.
                          strCompDirStat1 = "FLIP" & Mid(strCompDirStat1, 5)
                          strCompDirStat1 = Left(strCompDirStat1, (Len(strCompDirStat1) - Len(strTmp03)))  ' ** Strip previous line number.
                          strCompDirStat1 = strCompDirStat1 & CStr(lngLine)
                          If InStr(strCompDirStat1, " Not ") > 0 Then
                            blnIsNeg = True
                            Select Case lngCompDirType1
                            Case 1&
                              lngCompDirOpt1 = 2&
                            Case 2&
                              lngCompDirOpt1 = 6&
                            Case 3&
                              lngCompDirOpt1 = 10&
                            Case 4&
                              lngCompDirOpt1 = 14&
                            End Select
                          Else
                            Select Case lngCompDirType1
                            Case 1&
                              lngCompDirOpt1 = 3&
                            Case 2&
                              lngCompDirOpt1 = 7&
                            Case 3&
                              lngCompDirOpt1 = 11&
                            Case 4&
                              lngCompDirOpt1 = 15&
                            End Select
                          End If
                          intMode = 1
                        End If
                      ElseIf strCompDirStat1 <> vbNullString Then
                        ' ** Only 1st compiler directive to worry about.
                        strTmp03 = GetLastWord(strCompDirStat1)  ' ** Module Function: modStringFuncs.
                        strTmp02 = GetNthWord(strCompDirStat1, 2)  ' ** Module Function: modStringFuncs.
                        strCompDirStat1 = "FLIP" & Mid(strCompDirStat1, 5)
                        strCompDirStat1 = Left(strCompDirStat1, (Len(strCompDirStat1) - Len(strTmp03)))  ' ** Strip previous line number.
                        strCompDirStat1 = strCompDirStat1 & CStr(lngLine)
                        If InStr(strCompDirStat1, " Not ") > 0 Then
                          blnIsNeg = True
                          Select Case lngCompDirType1
                          Case 1&
                            lngCompDirOpt1 = 2&
                          Case 2&
                            lngCompDirOpt1 = 6&
                          Case 3&
                            lngCompDirOpt1 = 10&
                          Case 4&
                            lngCompDirOpt1 = 14&
                          End Select
                        Else
                          Select Case lngCompDirType1
                          Case 1&
                            lngCompDirOpt1 = 3&
                          Case 2&
                            lngCompDirOpt1 = 7&
                          Case 3&
                            lngCompDirOpt1 = 11&
                          Case 4&
                            lngCompDirOpt1 = 15&
                          End Select
                        End If
                        intMode = 1
                      ElseIf strCompDirStat2 <> vbNullString Then
                        ' ** Only 2nd compiler directive to worry about.
                        strTmp04 = GetLastWord(strCompDirStat2)  ' ** Module Function: modStringFuncs.
                        strTmp02 = GetNthWord(strCompDirStat2, 2)  ' ** Module Function: modStringFuncs.
                        strCompDirStat2 = "FLIP" & Mid(strCompDirStat2, 5)
                        strCompDirStat2 = Left(strCompDirStat2, (Len(strCompDirStat2) - Len(strTmp04)))  ' ** Strip previous line number.
                        strCompDirStat2 = strCompDirStat2 & CStr(lngLine)
                        If InStr(strCompDirStat2, " Not ") > 0 Then
                          blnIsNeg = True
                          Select Case lngCompDirType2
                          Case 1&
                            lngCompDirOpt2 = 2&
                          Case 2&
                            lngCompDirOpt2 = 6&
                          Case 3&
                            lngCompDirOpt2 = 10&
                          Case 4&
                            lngCompDirOpt2 = 14&
                          End Select
                        Else
                          Select Case lngCompDirType2
                          Case 1&
                            lngCompDirOpt2 = 3&
                          Case 2&
                            lngCompDirOpt2 = 7&
                          Case 3&
                            lngCompDirOpt2 = 11&
                          Case 4&
                            lngCompDirOpt2 = 15&
                          End Select
                        End If
                        intMode = 2
                      Else
                        Stop
                      End If
                    Case "#End"
                      lngLine = lngY
                      strTmp01 = strTmp01 & " If"
                      If strCompDirStat1 <> vbNullString And strCompDirStat2 <> vbNullString Then
                        ' ** 2 compiler directives are in use.
                        strTmp03 = GetLastWord(strCompDirStat1)  ' ** Module Function: modStringFuncs.
                        strTmp04 = GetLastWord(strCompDirStat2)  ' ** Module Function: modStringFuncs.
                        If Val(strTmp03) < Val(strTmp04) Then
                          ' ** We're in 2.
                          strTmp02 = GetNthWord(strCompDirStat2, 2)  ' ** Module Function: modStringFuncs.
                          strCompDirStat2 = vbNullString
                          Select Case lngCompDirType2
                          Case 1&
                            lngCompDirOpt2 = 4&
                          Case 2&
                            lngCompDirOpt2 = 8&
                          Case 3&
                            lngCompDirOpt2 = 12&
                          Case 4&
                            lngCompDirOpt2 = 16&
                          End Select
                          intMode = 2
                        Else
                          ' ** We're in 1.
                          strTmp02 = GetNthWord(strCompDirStat1, 2)  ' ** Module Function: modStringFuncs.
                          strCompDirStat1 = vbNullString
                          Select Case lngCompDirType1
                          Case 1&
                            lngCompDirOpt1 = 4&
                          Case 2&
                            lngCompDirOpt1 = 8&
                          Case 3&
                            lngCompDirOpt1 = 12&
                          Case 4&
                            lngCompDirOpt1 = 16&
                          End Select
                          intMode = 1
                        End If
                      ElseIf strCompDirStat1 <> vbNullString Then
                        strTmp02 = GetNthWord(strCompDirStat1, 2)  ' ** Module Function: modStringFuncs.
                        strCompDirStat1 = vbNullString
                        Select Case lngCompDirType1
                        Case 1&
                          lngCompDirOpt1 = 4&
                        Case 2&
                          lngCompDirOpt1 = 8&
                        Case 3&
                          lngCompDirOpt1 = 12&
                        Case 4&
                          lngCompDirOpt1 = 16&
                        End Select
                        intMode = 1
                      ElseIf strCompDirStat2 <> vbNullString Then
                        strTmp02 = GetNthWord(strCompDirStat2, 2)  ' ** Module Function: modStringFuncs.
                        strCompDirStat2 = vbNullString
                        Select Case lngCompDirType2
                        Case 1&
                          lngCompDirOpt2 = 4&
                        Case 2&
                          lngCompDirOpt2 = 8&
                        Case 3&
                          lngCompDirOpt2 = 12&
                        Case 4&
                          lngCompDirOpt2 = 16&
                        End Select
                        intMode = 2
                      Else
                        Stop
                      End If
                    Case Else
                      Stop
                    End Select

                    lngItems = lngItems + 1&
                    lngE = lngItems - 1&
                    ReDim Preserve arr_varItem(I_ELEMS, lngE)
                    arr_varItem(I_DID, lngE) = arr_varDir(D_DID, lngX)
                    arr_varItem(I_VID, lngE) = arr_varDir(D_VID, lngX)
                    arr_varItem(I_VNAM, lngE) = arr_varDir(D_VNAM, lngX)
                    varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", "[dbs_id] = " & CStr(arr_varItem(I_DID, lngE)) & " And " & _
                      "[vbcom_id] = " & CStr(arr_varItem(I_VID, lngE)) & " And [vbcomproc_name] = '" & strProcName & "'")
                    arr_varItem(I_PID, lngE) = varTmp00
                    arr_varItem(I_PNAM, lngE) = strProcName
                    arr_varItem(I_STMT, lngE) = strTmp01  ' ** Compiler Directive word or words.
                    arr_varItem(I_CON, lngE) = strTmp02  ' ** Compiler Directive constant.
                    arr_varItem(I_LIN, lngE) = lngLine
                    Select Case intMode
                    Case 1
                      arr_varItem(I_CID, lngE) = arr_varDir(D_CID1, lngX)
                      arr_varItem(I_TYP, lngE) = lngCompDirType1
                      arr_varItem(I_OPT, lngE) = lngCompDirOpt1
                    Case 2
                      arr_varItem(I_CID, lngE) = arr_varDir(D_CID2, lngX)
                      arr_varItem(I_TYP, lngE) = lngCompDirType2
                      arr_varItem(I_OPT, lngE) = lngCompDirOpt2
                    End Select
                    arr_varItem(I_NOT, lngE) = blnIsNeg
                    arr_varItem(I_RAW, lngE) = strLine

                  End If
                End If  ' ** Remark.
              End If  ' ** vbNullString.
            Next  ' ** lngY.

          End With  ' ** cod.
        End With  ' ** vbc.

      Next  ' ** lngX.
      Set cod = Nothing
      Set vbc = Nothing

    End With  ' ** vbp.
    Set vbp = Nothing

    Debug.Print "'HITS: " & CStr(lngItems)
    DoEvents

    If lngItems > 0& Then

      Set rst = .OpenRecordset("tblVBComponent_Directive_Detail", dbOpenDynaset, dbConsistent)
      blnAddAll = False: blnAdd = False
      With rst
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        End If
        For lngX = 0& To (lngItems - 1&)
          blnAdd = False
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            Stop  'NOT PROGRAMMED YET!
          End Select
          If blnAdd = True Then
            .AddNew
            ![dbs_id] = arr_varItem(I_DID, lngX)
            ![vbcom_id] = arr_varItem(I_VID, lngX)
            ![vbcomproc_id] = arr_varItem(I_PID, lngX)
            ![vbcomdir_id] = arr_varItem(I_CID, lngX)
            ' ** ![vbcomdirdet_id] : AutoNumber.
            ![compdir_type] = arr_varItem(I_TYP, lngX)
            ![vbcomdirdet_statement] = arr_varItem(I_STMT, lngX)
            ![compdir_constant] = arr_varItem(I_CON, lngX)
            ![vbcomdirdet_linenum] = arr_varItem(I_LIN, lngX)
            ![compdiropt_type] = arr_varItem(I_OPT, lngX)
            ![vbcomdirdet_not] = arr_varItem(I_NOT, lngX)
            ![vbcomdirdet_raw] = arr_varItem(I_RAW, lngX)
            ![vbcomdirdet_datemodified] = Now()
            .Update
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

    End If  ' ** lngItems.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_CompilerDirective_Usage = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function DimHiLbls() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "DimHiLbls"

  Dim frm As Access.Form, ctl As Access.Control
  Dim blnRetVal As Boolean

On Error GoTo 0

  blnRetVal = True

  Set frm = Forms(0)
  With frm
    For Each ctl In .FormHeader.Controls
      With ctl
        If Right(.Name, 7) = "_dim_hi" Then
          Debug.Print "'" & .Name
        End If
      End With
    Next
  End With

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set ctl = Nothing
  Set frm = Nothing
  DimHiLbls = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_Shortcut_Doc() As Boolean
' ** Document shortcut key remarks found at the top of every form module.

On Error GoTo ERRH

  Const THIS_PROC As String = "VBA_Shortcut_Doc"

  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim strModName As String, lngModType As Long, strLine As String, strLastType As String
  Dim strDesc As String, strKey As String, strCtl As String
  Dim strLastModName As String, lngVBComID As Long, strLastFrmName As String, lngFrmID As Long
  Dim lngLines As Long, lngDecLines As Long
  Dim lngKeys As Long, arr_varKey() As Variant
  Dim lngTypes As Long, arr_varType() As Variant
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean, blnFound2 As Boolean
  Dim intPos01 As Integer, intPos02 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, lngTmp03 As Long
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varKey().
  Const K_ELEMS As Integer = 16  ' ** Array's first-element UBound().
  Const K_DID   As Integer = 0
  Const K_VID   As Integer = 1
  Const K_VNAM  As Integer = 2
  Const K_FID1  As Integer = 3
  Const K_FNAM1 As Integer = 4
  Const K_KEY   As Integer = 5
  Const K_TYP   As Integer = 6
  Const K_TYPN  As Integer = 7
  Const K_CID   As Integer = 8
  Const K_CNAM  As Integer = 9
  Const K_DSC   As Integer = 10
  Const K_THIS  As Integer = 11
  Const K_FID2  As Integer = 12
  Const K_FNAM2 As Integer = 13
  Const K_MULTI As Integer = 14
  Const K_ISDEC As Integer = 15
  Const K_RAW   As Integer = 16

  ' ** Array: arr_varType().
  Const T_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const T_TYP As Integer = 0
  Const T_NAM As Integer = 1
  Const T_RAW As Integer = 2

On Error GoTo 0

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngKeys = 0&
  ReDim arr_varKey(K_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        lngModType = .Type
        ' **   vbext_ComponentType enumeration:
        ' **       1  vbext_ct_StdModule        Standard Module
        ' **       2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
        ' **       3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm in the VBA Editor.
        ' **      11  vbext_ct_ActiveXDesigner
        ' **     100  vbext_ct_Document         Module behind Form, Report, or Excel Worksheet.
        If Left(.Name, 5) = "Form_" Then
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            strLastType = vbNullString
            blnFound = False: blnFound2 = False
            For lngX = 1& To lngDecLines
              strLine = Trim(.Lines(lngX, 1))
              If strLine <> vbNullString Then
                If Left(strLine, 1) = "'" Then
                  ' ** Shortcut Alt keys responsive from this form:
                  If strLine = "' **   Name: ctlCurrentRecord  ControlSource: {unbound}" Or _
                      strLine = "' **   Choose RA:          (cmdChoose}" Or _
                      strLine = "' **   Choose RA:          (cmdChoose on frmAccountProfile_Add_Sub}" Or _
                      strLine = "' **   Country:            {country_name1}" Or _
                      strLine = "' **   Country:            {country_name1 on frmCurrency_Sub}" Or _
                      strLine = "' **   Enter Down:         {opgEnterKey_optDown}" Or _
                      strLine = "' **   Enter Down:         {opgEnterKey_optDown on frmJournal_Columns}" Or _
                      strLine = "' **     {Nothing}" Then
                    ' ** Skip these.
                  Else
                    If Left(strLine, 13) = "' ** Shortcut" Or InStr(strLine, "responsive from this form") > 0 Then
                      blnFound = True
                      blnFound2 = False
                      intPos01 = CharPos(strLine, 3, " ")
                      intPos02 = CharPos(strLine, 4, " ")
                      strTmp01 = Trim(Left(strLine, intPos02))
                      strTmp01 = Trim(Mid(strTmp01, intPos01))
                      If strTmp01 <> strLastType Then
                        strLastType = strTmp01
                      Else
                        If InStr(strModName, "frmAccountContacts") > 0 Then
                          ' ** These are OK.
                        Else
                          Stop
                        End If
                      End If
                    End If
                    If blnFound = True Then
                      ' **   Date:             D {transdate}
                      blnFound2 = False
                      If Left(strLine, 7) = "' **   " And Right(strLine, 1) = "}" Then
                        blnFound2 = True
                      ElseIf InStr(strLine, ":") > 0 And InStr(strLine, "{") > 0 And InStr(strLine, "}") > 0 Then
                        ' ** I'm trying to cover if the spacing isn't exactly right,
                        ' ** or there's something else at the end of the line.
                        blnFound2 = True
                      End If
                      If blnFound2 = True Then
                        intPos01 = InStr(strLine, ":")
                        If intPos01 > 0 Then
                          strTmp01 = Left(strLine, intPos01)
                          intPos02 = CharPos(strTmp01, 2, " ")
                          strDesc = Trim(Mid(strTmp01, intPos02))
                          strTmp01 = Trim(Mid(strLine, (intPos01 + 1)))
                          intPos01 = InStr(strTmp01, " ")
                          If intPos01 > 0 Then
                            strTmp02 = Trim(Left(strTmp01, intPos01))
                            If Len(strTmp02) = 1 Or ((Len(strTmp02) = 2 Or Len(strTmp02) = 3) And Left(strTmp02, 1) = "F") Then
                              strCtl = Trim(Mid(strTmp01, intPos01))
                              strKey = strTmp02
                              lngKeys = lngKeys + 1&
                              lngE = lngKeys - 1&
                              ReDim Preserve arr_varKey(K_ELEMS, lngE)
                              ' *****************************************************
                              ' ** Array: arr_varKey()
                              ' **
                              ' **   Field  Element  Name                Constant
                              ' **   =====  =======  ==================  ==========
                              ' **     1       0     dbs_id              K_DID
                              ' **     2       1     vbcom_id            K_VID
                              ' **     3       2     vbcom_name          K_VNAM
                              ' **     4       3     frm_id1             K_FID1
                              ' **     5       4     frm_name1           K_FNAM1
                              ' **     6       5     fs_key              K_KEY
                              ' **     7       6     keydowntype_type    K_TYP
                              ' **     8       7     keydowntype_name    K_TYPN
                              ' **     9       8     ctl_id              K_CID
                              ' **    10       9     ctl_name            K_CNAM
                              ' **    11      10     key_description     K_DSC
                              ' **    12      11     this form (T/F)     K_THIS
                              ' **    13      12     frm_id2             K_FID2
                              ' **    14      13     frm_name2           K_FNAM2
                              ' **    15      14     Multiple Forms      K_MULTI
                              ' **    16      15     Is Declaration      K_ISDEC
                              ' **    17      16     vbsct_raw           K_RAW
                              ' **                   Is KeyDown
                              ' **
                              ' *****************************************************
                              arr_varKey(K_DID, lngE) = lngThisDbsID
                              arr_varKey(K_VID, lngE) = Null
                              arr_varKey(K_VNAM, lngE) = strModName
                              arr_varKey(K_FID1, lngE) = Null
                              arr_varKey(K_FNAM1, lngE) = Mid(strModName, (InStr(strModName, "_") + 1))
                              arr_varKey(K_KEY, lngE) = strKey
                              arr_varKey(K_TYP, lngE) = Null
                              arr_varKey(K_TYPN, lngE) = strLastType
                              arr_varKey(K_CID, lngE) = Null
                              arr_varKey(K_CNAM, lngE) = strCtl
                              arr_varKey(K_DSC, lngE) = strDesc
                              If InStr(strDesc, " on ") > 0 Then
                                arr_varKey(K_THIS, lngE) = CBool(False)
                              Else
                                arr_varKey(K_THIS, lngE) = CBool(True)
                              End If
                              arr_varKey(K_FID2, lngE) = Null
                              arr_varKey(K_FNAM2, lngE) = Null
                              arr_varKey(K_MULTI, lngE) = CBool(False)
                              arr_varKey(K_ISDEC, lngE) = CBool(True)
                              arr_varKey(K_RAW, lngE) = strLine
                            Else
                              Debug.Print "'" & strLine
                              DoEvents
                            End If
                          Else
                            Debug.Print "'" & strLine
                            DoEvents
                          End If
                        Else
                          Debug.Print "'" & strLine
                          DoEvents
                          '' **   Show JournalNo    J {chkShowJournalNo}
                          '' **   Show JournalNo    J {chkShowJournalNo on frmAccountHideTrans2_One}
                          '' **   Show JournalNo    J {chkShowJournalNo on frmAccountHideTrans2_One}
                          '' **   One               E {ckgDeleteDates_chkOne}
                          '' **   One               E {ckgDeleteDates_chkOne on frmAssetPricing_History}
                          '' **   Issue Date        A {Issue_Date}
                          '' **   Issue Date        A {Issue_Date on frmCheckPOSPay}
                          '' **   Issue Date        A {Issue_Date on frmCheckPOSPay}
                          '' **   Issue Date        A {Issue_Date on frmCheckPOSPay}
                          '' **   Description       P {pp_description}
                          '' **   All               A {opgFilter_optAll}
                          '' **   All               A {opgFilter_optAll on frmCurrency_Account}
                          '' **   Principal Cash    H {Amount}
                          '' **   Site Map          M {cmdSiteMap}
                          '' **   Group Rev/Exp     U {chkGroupBy_IncExpCode}
                        End If
                      End If  ' ** blnFound2.
                    End If  ' ** blnFound.
                  End If  ' ** Bad ones.
                End If  ' ** Remark.
              End If  ' ** vbNullString.
            Next  ' ** lngX.
          End With  ' ** cod.
        End If  ' ** Form_...
      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.
  'DO WE WANT TO DO THE Form_KeyDown SUB AS WELL?  JUST TO COMPARE?

  Debug.Print "'KEYS: " & CStr(lngKeys)
  DoEvents
'KEYS: 3138

  If lngKeys > 0& Then

    lngTypes = 0&
    ReDim arr_varType(T_ELEMS, 0)

    For lngX = 0& To (lngKeys - 1&)
      blnFound = False
      For lngY = 0& To (lngTypes - 1&)
        If arr_varType(T_RAW, lngY) = arr_varKey(K_TYPN, lngX) Then
          blnFound = True
          Exit For
        End If
      Next  ' ** lngY.
      If blnFound = False Then
        lngTypes = lngTypes + 1&
        lngE = lngTypes - 1&
        ReDim Preserve arr_varType(T_ELEMS, lngE)
        arr_varType(T_TYP, lngE) = Null
        arr_varType(T_NAM, lngE) = Null
        arr_varType(T_RAW, lngE) = arr_varKey(K_TYPN, lngX)
      End If
    Next  ' ** lngX.

    Debug.Print "'TYPES: " & CStr(lngTypes)
    DoEvents

    For lngX = 0& To (lngTypes - 1&)
      If arr_varType(T_RAW, lngX) = "F-Keys" Then
        varTmp00 = DLookup("[keydowntype_type]", "tblKeyDownType", "[keydowntype_name] = 'Plain'")
      Else
        varTmp00 = DLookup("[keydowntype_type]", "tblKeyDownType", "[keydowntype_name] = '" & arr_varType(T_RAW, lngX) & "'")
      End If
      If IsNull(varTmp00) = False Then
        arr_varType(T_TYP, lngX) = varTmp00
        arr_varType(T_NAM, lngX) = IIf(arr_varType(T_RAW, lngX) = "F-Keys", "Plain", arr_varType(T_RAW, lngX))
      Else
        Stop
      End If
      ' ** I've tried to get it in there right, but I'll just have to resort to a fix.
      If arr_varType(T_NAM, lngX) = "F-Keys" Then
        arr_varType(T_NAM, lngX) = "Plain"
        arr_varType(T_TYP, lngX) = 0&
      ElseIf arr_varType(T_RAW, lngX) = "F-Keys" Then
        arr_varType(T_NAM, lngX) = "Plain"
        arr_varType(T_TYP, lngX) = 0&
      End If
    Next  ' ** lngX.

    For lngX = 0& To (lngKeys - 1&)
      For lngY = 0& To (lngTypes - 1&)
        If arr_varType(T_NAM, lngY) = arr_varKey(K_TYPN, lngX) Then
          arr_varKey(K_TYP, lngX) = arr_varType(T_TYP, lngY)
          Exit For
        End If
      Next  ' ** lngY.
      ' ** I've tried to get it in there right, but I'll just have to resort to a fix.
      If arr_varKey(K_TYPN, lngX) = "F-Keys" Then
        arr_varKey(K_TYPN, lngX) = "Plain"
        arr_varKey(K_TYP, lngX) = 0&
      End If
    Next  ' ** lngX.

    lngTmp03 = 0&
    For lngX = 0& To (lngKeys - 1&)
      intPos01 = InStr(arr_varKey(K_CNAM, lngX), "}")
      If intPos01 <> Len(arr_varKey(K_CNAM, lngX)) Then
        arr_varKey(K_CNAM, lngX) = Left(arr_varKey(K_CNAM, lngX), intPos01)
      End If
      If InStr(arr_varKey(K_CNAM, lngX), "and on each subform") > 0 Then
        arr_varKey(K_CNAM, lngX) = Trim(Left(arr_varKey(K_CNAM, lngX), InStr(arr_varKey(K_CNAM, lngX), " ")))
      End If
      intPos01 = InStr(arr_varKey(K_CNAM, lngX), " on ")
      If intPos01 > 0 Then
        strTmp01 = Trim(Left(arr_varKey(K_CNAM, lngX), intPos01))
        strTmp02 = Trim(Mid(arr_varKey(K_CNAM, lngX), (intPos01 + 3)))
        If Right(strTmp02, 1) = "}" Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
        If Left(strTmp01, 1) = "{" Then strTmp01 = Mid(strTmp01, 2)
        intPos02 = InStr(strTmp02, ",")
        If intPos02 > 0 Then
          ' ** More than 1 form listed!
          strTmp02 = Left(strTmp02, (intPos02 - 1))
          lngTmp03 = lngTmp03 + 1&
          arr_varKey(K_MULTI, lngX) = CBool(True)
        End If
        arr_varKey(K_CNAM, lngX) = strTmp01
        arr_varKey(K_FNAM2, lngX) = strTmp02
        arr_varKey(K_THIS, lngX) = CBool(False)
      Else
        strTmp01 = arr_varKey(K_CNAM, lngX)
        If Right(strTmp01, 1) = "}" Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
        If Left(strTmp01, 1) = "{" Then strTmp01 = Mid(strTmp01, 2)
        arr_varKey(K_CNAM, lngX) = strTmp01
      End If
      If IsNull(arr_varKey(K_FNAM2, lngX)) = False Then
        If Left(arr_varKey(K_FNAM2, lngX), 3) <> "frm" Then
          arr_varKey(K_FNAM2, lngX) = Null
        End If
      End If
    Next  ' ** lngX.

    If lngTmp03 > 0& Then
      Debug.Print "'" & CStr(lngTmp03) & " ENTRIES HAVE MULTIPLE FORMS!"
      DoEvents
    End If

    Set dbs = CurrentDb
    With dbs

      lngVBComID = 0&
      strLastModName = vbNullString

      Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      With rst
        For lngX = 0& To (lngKeys - 1&)
          If arr_varKey(K_VNAM, lngX) <> strLastModName Then
            .MoveFirst
            strLastModName = arr_varKey(K_VNAM, lngX)
            .FindFirst "[dbs_id] = " & CStr(arr_varKey(K_DID, lngX)) & " And [vbcom_name] = '" & strLastModName & "'"
            If .NoMatch = False Then
              lngVBComID = ![vbcom_id]
              arr_varKey(K_VID, lngX) = lngVBComID
            Else
              Stop
            End If
          Else
            arr_varKey(K_VID, lngX) = lngVBComID
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      lngFrmID = 0&
      strLastFrmName = vbNullString

      Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
      With rst
        For lngX = 0& To (lngKeys - 1&)
          If arr_varKey(K_FNAM1, lngX) <> strLastFrmName Then
            .MoveFirst
            strLastFrmName = arr_varKey(K_FNAM1, lngX)
            .FindFirst "[dbs_id] = " & CStr(arr_varKey(K_DID, lngX)) & " And [frm_name] = '" & strLastFrmName & "'"
            If .NoMatch = False Then
              lngFrmID = ![frm_id]
              arr_varKey(K_FID1, lngX) = lngFrmID
            Else
              Stop
            End If
          Else
            arr_varKey(K_FID1, lngX) = lngFrmID
          End If
          If IsNull(arr_varKey(K_FNAM2, lngX)) = False Then
            .MoveFirst
            .FindFirst "[dbs_id] = " & CStr(arr_varKey(K_DID, lngX)) & " And [frm_name] = '" & arr_varKey(K_FNAM2, lngX) & "'"
            If .NoMatch = False Then
              arr_varKey(K_FID2, lngX) = ![frm_id]
            Else
              Stop
            End If
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst
      Set rst = Nothing

      Set rst = .OpenRecordset("tblForm_Control", dbOpenDynaset, dbReadOnly)
      With rst
        For lngX = 0& To (lngKeys - 1&)
          .MoveFirst
          If arr_varKey(K_THIS, lngX) = True Then
            .FindFirst "[dbs_id] = " & CStr(arr_varKey(K_DID, lngX)) & " And [frm_id] = " & CStr(arr_varKey(K_FID1, lngX)) & " And " & _
              "[ctl_name] = '" & arr_varKey(K_CNAM, lngX) & "'"
            If .NoMatch = False Then
              arr_varKey(K_CID, lngX) = ![ctl_id]
            Else
              ' ** There'll be lots.
            End If
          Else
            If IsNull(arr_varKey(K_FID2, lngX)) = False Then
            .FindFirst "[dbs_id] = " & CStr(arr_varKey(K_DID, lngX)) & " And [frm_id] = " & CStr(arr_varKey(K_FID2, lngX)) & " And " & _
              "[ctl_name] = '" & arr_varKey(K_CNAM, lngX) & "'"
            If .NoMatch = False Then
              arr_varKey(K_CID, lngX) = ![ctl_id]
            Else
              ' ** There'll be lots.
            End If

            End If
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst
      Set rst = Nothing

      Set rst = .OpenRecordset("zz_tbl_VBComponent_Shortcut", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngKeys - 1&)
          .AddNew
          ' ** ![vbcomsct_id] : AutoNumber.
          ![dbs_id] = arr_varKey(K_DID, lngX)
          ![vbcom_id] = arr_varKey(K_VID, lngX)
          ![vbcom_name] = arr_varKey(K_VNAM, lngX)
          ![frm_id1] = arr_varKey(K_FID1, lngX)
          ![frm_name1] = arr_varKey(K_FNAM1, lngX)
          ![frm_id2] = arr_varKey(K_FID2, lngX)
          ![frm_name2] = arr_varKey(K_FNAM2, lngX)
          ![vbcomsct_key] = arr_varKey(K_KEY, lngX)
          ![keydowntype_type] = arr_varKey(K_TYP, lngX)
          ![keydowntype_name] = arr_varKey(K_TYPN, lngX)
          ![ctl_id] = arr_varKey(K_CID, lngX)
          ![ctl_name] = arr_varKey(K_CNAM, lngX)
          ![vbcomsct_description] = arr_varKey(K_DSC, lngX)
          ![vbcomsct_thisform] = arr_varKey(K_THIS, lngX)
          ![vbcomsct_isdec] = arr_varKey(K_ISDEC, lngX)
          ![vbcomsct_iskd] = False
          ![vbcomsct_multifrm] = arr_varKey(K_MULTI, lngX)
          intPos01 = InStr(arr_varKey(K_CNAM, lngX), ",")
          If intPos01 > 0 Then
            ![vbcomsct_multictl] = True
          Else
            ![vbcomsct_multictl] = False
          End If
          ![vbcomsct_raw] = arr_varKey(K_RAW, lngX)
          ![vbcomsct_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With
      Set rst = Nothing

      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

'KEYS: 3277
'TYPES: 7
'9 ENTRIES HAVE MULTIPLE FORMS!
'DONE!

  End If  ' ** lngKeys.

'TYPES: 7
'Alt
'Ctrl-Shift
'Ctrl
'F-Keys
'Alt-Shift
'Ctrl-Alt
'Ctrl-Alt-Shift

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set dbs = Nothing
  VBA_Shortcut_Doc = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_GetLineNum() As Boolean

  Const THIS_PROC As String = "VBA_GetLineNum"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim lngKeys As Long, arr_varKey As Variant
  Dim lngSubs As Long, arr_varSub As Variant
  Dim strModName As String, strProcName As String, strLine As String, strKeyType As String
  Dim lngLines As Long, lngDecLines As Long
  Dim blnFound As Boolean, blnFound2 As Boolean
  Dim intLen As Integer
  Dim strTmp01 As String, lngTmp02 As Long
  Dim lngX As Long, lngY As Long, lngZ As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varKey().
  Const K_ID6   As Integer = 0
  Const K_FID   As Integer = 2
  Const K_FNAM  As Integer = 3
  Const K_FID2  As Integer = 4
  Const K_FNAM2 As Integer = 5
  Const K_KTYP  As Integer = 7
  Const K_KEY   As Integer = 8
  Const K_LIN   As Integer = 15
  Const K_CONST As Integer = 20

  ' ** Array: arr_varSub().
  Const S_DID    As Integer = 0
  Const S_FID    As Integer = 1
  Const S_FNAM   As Integer = 2
  Const S_SUBID  As Integer = 3
  Const S_SUBNAM As Integer = 4

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    Set qdf = .QueryDefs("zzz_qry_xForm_Shortcut_29_20")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngKeys = .RecordCount
      .MoveFirst
      arr_varKey = .GetRows(lngKeys)
      ' *****************************************************
      ' ** Array: arr_varKey()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     vbcsc06_id          K_ID6
      ' **     2       1     dbs_id
      ' **     3       2     frm_id              K_FID
      ' **     4       3     frm_name            K_FNAM
      ' **     5       4     frm_id2             K_FID2
      ' **     6       5     frm_name2           K_FNAM2
      ' **     7       6     fs_order
      ' **     8       7     keydowntype_type    K_KTYP
      ' **     9       8     fs_key              K_KEY
      ' **    10       9     ctl_id
      ' **    11      10     fs_control
      ' **    12      11     ctltype_type
      ' **    13      12     fs_caption
      ' **    14      13     fs_unattached
      ' **    15      14     fs_parent
      ' **    16      15     fs_linenum          K_LIN
      ' **    17      16     fs_shift
      ' **    18      17     fs_alt
      ' **    19      18     fs_ctrl
      ' **    20      19     fs_letter
      ' **    21      20     keycode_constant    K_CONST
      ' **    22      21     fs_datemodified
      ' **
      ' *****************************************************
      .Close
    End With  ' ** rst.
    Set rst = Nothing
    Set qdf = Nothing

    Debug.Print "'KEYS: " & CStr(lngKeys)
    DoEvents

    If lngKeys > 0& Then

      strModName = vbNullString: strProcName = vbNullString
      intLen = Len("Private Sub Form_KeyDown(")
      Set vbp = Application.VBE.ActiveVBProject
      With vbp
        For lngX = 0& To (lngKeys - 1&)

          lngSubs = 0&
          arr_varSub = Empty

          ' ** zz_tbl_VBComponent_Shortcut_02, forms and their subforms, by specified [fnam].
          Set qdf = dbs.QueryDefs("zzz_qry_xForm_Shortcut_29_19")
          With qdf.Parameters
            ![fnam] = arr_varKey(K_FNAM, lngX)
          End With
          Set rst = qdf.OpenRecordset
          With rst
            If .BOF = True And .EOF = True Then
              Debug.Print "'NO SUBS: " & arr_varKey(K_FNAM, lngX)
            Else
              .MoveLast
              lngSubs = .RecordCount
              .MoveFirst
              arr_varSub = .GetRows(lngSubs)
              ' *************************************************
              ' ** Array: arr_varSub()
              ' **
              ' **   Field  Element  Name            Constant
              ' **   =====  =======  ==============  ==========
              ' **     1       0     dbs_id          S_DID
              ' **     2       1     frm_id          S_FID
              ' **     3       2     frm_name        S_FNAM
              ' **     4       3     frm_id_sub      S_SUBID
              ' **     5       4     frm_name_sub    S_SUBNAM
              ' **
              ' *************************************************
            End If
            .Close
          End With  ' ** rst.
          Set rst = Nothing
          Set qdf = Nothing

          For lngY = 0& To (lngSubs - 1&)

            strTmp01 = "Form_" & arr_varSub(S_SUBNAM, lngY)
            Set vbc = .VBComponents(strTmp01)

            With vbc
              Set cod = .CodeModule
              With cod
                lngLines = .CountOfLines
                lngDecLines = .CountOfDeclarationLines
                strKeyType = DLookup("[keydowntype_name]", "tblKeyDownType", "[keydowntype_type] = " & CStr(arr_varKey(K_KTYP, lngX)))
                blnFound = False: blnFound2 = False
                For lngZ = lngDecLines To lngLines
                  strLine = Trim(.Lines(lngZ, 1))
                  If strLine <> vbNullString Then
                    If blnFound = False Then
                      ' ** Private Sub Form_KeyDown(
                      If Left(strLine, intLen) = "Private Sub Form_KeyDown(" Then
                        blnFound = True
                      End If
                    Else
                      If Left(strLine, 7) = "End Sub" Then
                        Exit For
                      Else
                        If blnFound2 = False Then
                          ' ** Plain keys.
                          ' ** Alt keys.
                          ' ** Ctrl keys.
                          ' ** Ctrl-Shift keys.
                          strTmp01 = "' ** " & strKeyType & " keys."
                          If strLine = strTmp01 Then
                            blnFound2 = True
                          End If
                        Else
                          strTmp01 = "Case " & arr_varKey(K_CONST, lngX)
                          If InStr(strLine, strTmp01) > 0 Then
                            arr_varKey(K_LIN, lngX) = lngZ
                            arr_varKey(K_FID2, lngX) = arr_varSub(S_SUBID, lngY)
                            arr_varKey(K_FNAM2, lngX) = arr_varSub(S_SUBNAM, lngY)
                            Exit For
                          ElseIf InStr(strLine, "Case ") > 0 And InStr(strLine, arr_varKey(K_CONST, lngX)) > 0 Then
                            arr_varKey(K_LIN, lngX) = lngZ
                            arr_varKey(K_FID2, lngX) = arr_varSub(S_SUBID, lngY)
                            arr_varKey(K_FNAM2, lngX) = arr_varSub(S_SUBNAM, lngY)
                            Exit For
                          End If
                        End If  ' ** blnFound2.
                      End If
                    End If  ' ** blnFound.
                  End If  ' ** vbNullString.
                Next  ' ** lngZ.
              End With  ' ** cod.
            End With  ' ** vbc.

          Next  ' ** lngY.
        Next  ' ** lngX.
        Set cod = Nothing
        Set vbc = Nothing
      End With  ' ** vbp.
      Set vbp = Nothing

      lngTmp02 = 0&
      For lngX = 0& To (lngKeys - 1&)
        If IsNull(arr_varKey(K_LIN, lngX)) = False Then
          lngTmp02 = lngTmp02 + 1&
        End If
      Next  ' ** lngX.

      Debug.Print "'LINE NUMS FOUND: " & CStr(lngTmp02)
      DoEvents

      Set rst = .OpenRecordset("zz_tbl_VBComponent_Shortcut_07", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngKeys - 1&)
          If IsNull(arr_varKey(K_LIN, lngX)) = False Then
            .FindFirst "[vbcsc06_id] = " & CStr(arr_varKey(K_ID6, lngX))
            If .NoMatch = False Then
              .Edit
              ![fs_linenum] = arr_varKey(K_LIN, lngX)
              ![frm_id2] = arr_varKey(K_FID2, lngX)
              ![frm_name2] = arr_varKey(K_FNAM2, lngX)
              ![vbcsc07_datemodified] = Now()
              .Update
            End If
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst
      Set rst = Nothing

    End If  ' ** lngKeys.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Beep

  Debug.Print "'DONE!"
  DoEvents

EXITP:
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  VBA_GetLineNum = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function VBA_FindQrys() As Boolean

  Const THIS_PROC As String = "VBA_FindQrys"

  Dim wrk As DAO.Workspace, dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngFiles As Long, arr_varFile() As Variant
  Dim strPath As String, strFile As String, strPathFile As String, strDesc As String, strSQL As String
  Dim blnFound As Boolean
  Dim lngW As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const Q_DNAM As Integer = 0
  Const Q_PATH As Integer = 1
  Const Q_QNAM As Integer = 2
  Const Q_TYP  As Integer = 3
  Const Q_DSC  As Integer = 4
  Const Q_SQL  As Integer = 5
  Const Q_FND  As Integer = 6

  ' ** Array: arr_varFile()
  Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const F_DNAM As Integer = 0
  Const F_PATH As Integer = 1
  Const F_FND  As Integer = 2

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  strPath = "C:\Program Files\Delta Data\Trust Accountant"  ' ** Start here.

  lngFiles = 0&
  ReDim arr_varFile(F_ELEMS, 0)

  lngFiles = lngFiles + 1&
  lngE = lngFiles - 1&
  ReDim Preserve arr_varFile(F_ELEMS, lngE)
  arr_varFile(F_DNAM, lngE) = "TrustAux.mdb"
  arr_varFile(F_PATH, lngE) = strPath
  arr_varFile(F_FND, lngE) = CBool(False)

  lngFiles = lngFiles + 1&
  lngE = lngFiles - 1&
  ReDim Preserve arr_varFile(F_ELEMS, lngE)
  arr_varFile(F_DNAM, lngE) = "TrstXAdm.mdb"
  arr_varFile(F_PATH, lngE) = strPath
  arr_varFile(F_FND, lngE) = CBool(False)

  lngFiles = lngFiles + 1&
  lngE = lngFiles - 1&
  ReDim Preserve arr_varFile(F_ELEMS, lngE)
  arr_varFile(F_DNAM, lngE) = "TrstXAdm - Copy (6).mdb"
  arr_varFile(F_PATH, lngE) = strPath
  arr_varFile(F_FND, lngE) = CBool(False)

  lngFiles = lngFiles + 1&
  lngE = lngFiles - 1&
  ReDim Preserve arr_varFile(F_ELEMS, lngE)
  arr_varFile(F_DNAM, lngE) = "Trust_wshortcutstuff.mdb"
  arr_varFile(F_PATH, lngE) = strPath
  arr_varFile(F_FND, lngE) = CBool(False)

  strPath = strPath & LNK_SEP & "Client_Frontends"
  strPathFile = strPath & LNK_SEP & "*.mdb"

  blnFound = True: strFile = vbNullString
  Do While blnFound = True
    blnFound = False
    If strFile = vbNullString Then
      strFile = Dir(strPathFile, vbNormal)
      If strFile <> vbNullString Then
        blnFound = True
        lngFiles = lngFiles + 1&
        lngE = lngFiles - 1&
        ReDim Preserve arr_varFile(F_ELEMS, lngE)
        arr_varFile(F_DNAM, lngE) = strFile
        arr_varFile(F_PATH, lngE) = Parse_Path(strPathFile)  ' ** Module Function: modFileUtilities.
        arr_varFile(F_FND, lngE) = CBool(False)
      End If
    Else
      strFile = Dir()
      If strFile <> vbNullString Then
        blnFound = True
        lngFiles = lngFiles + 1&
        lngE = lngFiles - 1&
        ReDim Preserve arr_varFile(F_ELEMS, lngE)
        arr_varFile(F_DNAM, lngE) = strFile
        arr_varFile(F_PATH, lngE) = Parse_Path(strPathFile)  ' ** Module Function: modFileUtilities.
        arr_varFile(F_FND, lngE) = CBool(False)
      End If
    End If
  Loop

  Debug.Print "'FILES: " & CStr(lngFiles)
  DoEvents

  Set wrk = CreateWorkspace("tmpWRK", "Superuser", TA_SEC, dbUseJet)
  With wrk

    lngQrys = 0&
    ReDim arr_varQry(Q_ELEMS, 0)

    For lngW = 1& To (lngFiles - 1&)

      strPathFile = arr_varFile(F_PATH, lngW) & LNK_SEP & arr_varFile(F_DNAM, lngW)

      blnFound = False
      Set dbs = .OpenDatabase(strPathFile, True, True)  ' ** {pathfile}, {exclusive}, {read-only}
      With dbs
        For Each qdf In .QueryDefs
          strDesc = vbNullString
          With qdf
            If Left(.Name, 23) = "zz_qry_VBComponent_Var_" Then
              blnFound = True
              lngQrys = lngQrys + 1&
              lngE = lngQrys - 1&
              ReDim Preserve arr_varQry(Q_ELEMS, lngE)
              arr_varQry(Q_DNAM, lngE) = arr_varFile(F_DNAM, lngW)
              arr_varQry(Q_PATH, lngE) = arr_varFile(F_PATH, lngW)
              arr_varQry(Q_QNAM, lngE) = qdf.Name
              arr_varQry(Q_TYP, lngE) = qdf.Type
On Error Resume Next
              strDesc = .Properties("Description")
On Error GoTo 0
              If strDesc <> vbNullString Then
                arr_varQry(Q_DSC, lngE) = strDesc
              Else
                arr_varQry(Q_DSC, lngE) = Null
              End If
              arr_varQry(Q_SQL, lngE) = .SQL
              arr_varQry(Q_FND, lngE) = CBool(False)
            End If
          End With  ' ** qdf.
        Next  ' ** qdf.
        .Close
      End With  ' ** dbs.
      Set dbs = Nothing

    Next  ' ** lngW.

    .Close
  End With  ' ** wrk.

  Debug.Print "'QRYS: " & CStr(lngQrys)
  DoEvents

  If lngQrys > 0& Then
    blnFound = True
    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("zz_tbl_VBComponent_Var01", dbOpenDynaset, dbConsistent)
      With rst
        For lngW = 0& To (lngQrys - 1&)
          .AddNew
          ' ** ![vart_id] : AutoNumber.
          ![qry_name] = arr_varQry(Q_QNAM, lngW)
          ![dbs_name] = arr_varQry(Q_DNAM, lngW)
          ![dbs_path] = arr_varQry(Q_PATH, lngW)
          ![qrytype_type] = arr_varQry(Q_TYP, lngW)
          ![qry_description] = arr_varQry(Q_DSC, lngW)
          ![qry_sql] = arr_varQry(Q_SQL, lngW)
          ![vart_found] = False
          ![vart_datemodified] = Now()
          .Update
        Next  ' ** lngW.
        .Close
      End With
      Set rst = Nothing
      .Close
    End With  ' ** dbs.
  End If

  Beep

  Debug.Print "'DONE!  " & blnFound
  DoEvents

'FILES: 43
'QRYS: 325
'DONE!  True
EXITP:
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set wrk = Nothing
  VBA_FindQrys = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function CtlProp_Set() As Boolean

  Const THIS_PROC As String = "CtlProp_Set"

  Dim frm As Access.Form, ctl As Access.Control
  Dim blnRetVal As Boolean

  blnRetVal = True

  Set frm = Forms(0)
  With frm
    For Each ctl In .Detail.Controls
      With ctl
        If Right(.Name, 9) = "_lbl_recs" Then
          .RightMargin = 15&
        End If
      End With
    Next
  End With
  Set ctl = Nothing
  Set frm = Nothing

  Beep

  CtlProp_Set = blnRetVal

End Function

Public Function VBA_Check_TPP() As Boolean

  Const THIS_PROC As String = "VBA_Check_TPP"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim lngLines As Long, lngDecLines As Long, lngFixed As Long
  Dim strModName As String, strProcName As String, strLine As String, strFind As String, strNewLine As String
  Dim lngHits As Long, arr_varHit() As Variant
  Dim lngThisDbsID As Long
  Dim blnAdd As Boolean, blnSkip As Boolean
  Dim intPos01 As Integer, intPos02 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, lngTmp05 As Long, lngTmp06 As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varHit().
  Const H_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const H_VID  As Integer = 0
  Const H_VNAM As Integer = 1
  Const H_PID  As Integer = 2
  Const H_PNAM As Integer = 3
  Const H_LIN  As Integer = 4
  Const H_COD  As Integer = 5

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  strFind = "lngTpp = GetTPP"
  strNewLine = "lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!"

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
  lngFixed = 0&

  lngHits = 0&
  ReDim arr_varHit(H_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        Set cod = .CodeModule
        With cod
          lngLines = .CountOfLines
          lngDecLines = .CountOfDeclarationLines
          For lngX = 1& To lngLines
            strProcName = vbNullString
            strLine = .Lines(lngX, 1)
            If Trim(strLine) <> vbNullString Then
              intPos01 = InStr(strLine, strFind)
              If intPos01 > 0 Then
                strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                If strProcName = vbNullString Then strProcName = "Declaration"
                If strProcName <> THIS_PROC Then
                  lngHits = lngHits + 1&
                  lngE = lngHits - 1&
                  ReDim Preserve arr_varHit(H_ELEMS, lngE)
                  arr_varHit(H_VID, lngE) = Null
                  arr_varHit(H_VNAM, lngE) = strModName
                  arr_varHit(H_PID, lngE) = Null
                  arr_varHit(H_PNAM, lngE) = strProcName
                  arr_varHit(H_LIN, lngE) = lngX
                  arr_varHit(H_COD, lngE) = Null
                  If intPos01 > 1 Then
                    If Mid(strLine, (intPos01 - 1), 1) = "'" Then
                      ' ** Already remarked out and replaced.
                      strTmp01 = .Lines((lngX + 1&), 1)  ' ** Next line.
                      If InStr(strTmp01, strNewLine) = 0 Then  ' ** Might be just missing my remark.
                        Beep
                        Debug.Print "'" & strTmp01
                        Stop
                      End If
                    Else
                      ' ** This one needs to be replaced.
                      strTmp01 = Left(strLine, (intPos01 - 1))
                      strTmp03 = Mid(strLine, intPos01)          ' ** Find to end of line.
                      intPos02 = InStr(strTmp01, " ")            ' ** Find the first space.
                      strTmp02 = Left(strTmp01, (intPos01 - 1))  ' ** Expecting this to be a line number.
                      strTmp01 = Mid(strTmp01, intPos01)         ' ** Everything between number and strFind: just spaces most likely.
                      If IsNumeric(strTmp02) = False Then
                        ' ** If it's not a number, then what?
                        Beep
                        Debug.Print "'" & strTmp02
                        Stop
                      Else
                        arr_varHit(H_COD, lngE) = strTmp02
                        strTmp04 = Space(Len(strTmp02))     ' ** Replace numbers with spaces.
                        strTmp04 = strTmp04 & strTmp01      ' ** Spaces, then more spaces.
                        strTmp03 = "'" & strTmp03           ' ** Remark out the GetTPP() call.
                        strTmp04 = strTmp04 & strTmp03      ' ** Add remarked code to line.
                        .ReplaceLine lngX, strTmp04         ' ** Replace it.
                        strTmp04 = strTmp02                 ' ** Line number.
                        strTmp04 = strTmp04 & strTmp01      ' ** Add spaces.
                        strTmp04 = strTmp04 & strNewLine    ' ** The completed replacement line.
                        .InsertLines (lngX + 1&), strTmp04  ' ** Add new line after old one.
                        lngFixed = lngFixed + 1&
                        'TOTAL MODULE LINES IS NOW CHANGED! SHOULD I UPDATE lngLines?
                      End If  ' ** IsNumeric().
                    End If
                  Else
                    ' ** Not expecting to find at beginning of line!
                    Beep
                    Debug.Print "'" & strLine
                    Stop
                  End If
                End If  ' ** THIS_PROC.
              End If  ' ** intPos01.
            End If  ' ** vbNullString.
          Next  ' ** lngX.
        End With  ' ** cod.
      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.

  Debug.Print "'TOT HITS: " & CStr(lngHits)
  DoEvents
  Debug.Print "'FIXED: " & CStr(lngFixed)
  DoEvents

  If lngHits > 0& Then
    Set dbs = CurrentDb
    With dbs

      Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngHits - 1&)
          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & arr_varHit(H_VNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varHit(H_VID, lngX) = ![vbcom_id]
          Else
            Stop
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngHits - 1&)
          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
            "[vbcomproc_name] = '" & arr_varHit(H_PNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varHit(H_PID, lngX) = ![vbcomproc_id]
          Else
            Stop
          End If
        Next  ' ** lngX.
        .Close
      End With
      Set rst = Nothing

      ' ** Look for more than 1 hit in the same procedure.
      lngTmp05 = 0&: lngTmp06 = 0&
      For lngX = 0& To (lngHits - 1&)
        If arr_varHit(H_VID, lngX) = lngTmp05 And arr_varHit(H_PID, lngX) = lngTmp06 Then
          Debug.Print "'2 IN 1: " & arr_varHit(H_VNAM, lngX) & "  " & arr_varHit(H_PNAM, lngX)
          DoEvents
        Else
          lngTmp05 = arr_varHit(H_VID, lngX)
          lngTmp06 = arr_varHit(H_PID, lngX)
        End If
      Next  ' ** lngX.
'TOT HITS: 395
'FIXED: 0
'2 IN 1: Form_frmMap_Misc_LTCL  Form_Timer
'2 IN 1: Form_frmMap_Misc_STCGL  Form_Timer
'DONE!

      blnSkip = True
      If blnSkip = False Then
        Set rst = .OpenRecordset("zz_tbl_VBComponent_TPP", dbOpenDynaset, dbConsistent)
        With rst
          .MoveFirst
          For lngX = 0& To (lngHits - 1&)
            blnAdd = False
            .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
              "[vbcomproc_id] = " & CStr(arr_varHit(H_PID, lngX))
            Select Case .NoMatch
            Case True
              blnAdd = True
            Case False
Stop
            End Select
            If blnAdd = True Then
              .AddNew
              ' ** ![vbcomtpp_id] : AutoNumber.
              ![dbs_id] = lngThisDbsID
              ![vbcom_id] = arr_varHit(H_VID, lngX)
              ![vbcom_name] = arr_varHit(H_VNAM, lngX)
              ![vbcomproc_id] = arr_varHit(H_PID, lngX)
              ![vbcomproc_name] = arr_varHit(H_PNAM, lngX)
              ![vbcomtpp_linenum] = arr_varHit(H_LIN, lngX)
              ![vbcomtpp_codenum] = arr_varHit(H_COD, lngX)
              ![vbcomtpp_datemodified] = Now()
              .Update
            End If
          Next  ' ** lngX.
          .Close
        End With
        Set rst = Nothing
      End If  ' ** blnSkip.

      .Close
    End With
  End If  ' ** lngHits.

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set dbs = Nothing

  VBA_Check_TPP = blnRetVal

End Function

Public Function VBA_CodeLineNum_Chk() As Boolean

  Const THIS_PROC As String = "VBA_CodeLineNum_Chk"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent
  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim strModName As String
  Dim lngRecs As Long
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  Set dbs = CurrentDb
  With dbs

    Set rst = .OpenRecordset("tblVBComponent_CodeNum_Max", dbOpenDynaset, dbConsistent)
    rst.MoveLast
    lngRecs = rst.RecordCount
    rst.MoveFirst

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For Each vbc In .VBComponents
        With vbc
          strModName = .Name

          With rst
            .FindFirst "[vbcom_name] = '" & strModName & "'"
            If .NoMatch = True Then
              Debug.Print "'NOT FOUND: " & strModName
            End If
          End With

        End With
      Next
    End With
    rst.Close

    .Close
  End With

  Beep

  Debug.Print "'DONE!"
  DoEvents

  VBA_CodeLineNum_Chk = blnRetVal

End Function

Public Function MastBalChk() As Boolean

  Const THIS_PROC As String = "MastBalChk"

  Dim frm As Access.Form, ctl As Access.Control
  Dim lngLeft As Long, lngWidth As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  Set frm = Forms(0).frmMasterBalance_Sub.Form
  With frm
    For Each ctl In .Controls
      With ctl
On Error Resume Next
        lngLeft = .Left
        lngWidth = .Width
        If ERR.Number = 0 Then
On Error GoTo 0
          If lngLeft > 10980 Then
            Debug.Print "'" & .Name
          ElseIf lngLeft + lngWidth > 10980 Then
            Debug.Print "'" & .Name
          End If
        Else
On Error GoTo 0
        End If
      End With
    Next
  End With

  Set ctl = Nothing
  Set frm = Nothing

  MastBalChk = blnRetVal

End Function
