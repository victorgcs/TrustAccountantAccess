Attribute VB_Name = "zz_mod_VarConstDoc"
Option Compare Database
Option Explicit

'VGC 04/19/2016: CHANGES!

Private Const THIS_NAME As String = "zz_mod_VarConstDoc"
' **

Public Function Var_Doc() As Boolean

  Const THIS_PROC As String = "Var_Doc"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim lngLines As Long, lngDecLines As Long
  Dim strLine As String, strModName As String, strProcName As String, strVarType As String
  Dim lngLVars As Long, arr_varLVar() As Variant
  Dim lngPVars As Long, arr_varPVar() As Variant
  Dim lngGVars As Long, arr_varGVar() As Variant
  Dim blnFound As Boolean, blnOptional As Boolean
  Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intCnt As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
  Dim lngX As Long, lngY As Long, intZ As Integer, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varLVar(), arr_varPVar(), arr_varGVar().
  Const V_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const V_VNAM As Integer = 0
  Const V_MOD  As Integer = 1
  Const V_PROC As Integer = 2
  Const V_SCOP As Integer = 3
  Const V_TYPE As Integer = 4
  Const V_CNT  As Integer = 5
  Const V_PARM As Integer = 6
  Const V_OPT  As Integer = 7

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngLVars = 0&
  ReDim arr_varLVar(V_ELEMS, 0)
  lngPVars = 0&
  ReDim arr_varPVar(V_ELEMS, 0)
  lngGVars = 0&
  ReDim arr_varGVar(V_ELEMS, 0)

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
            strLine = .Lines(lngX, 1)
            strLine = Trim(strLine)
            strTmp01 = vbNullString: strTmp02 = vbNullString: strTmp03 = vbNullString
            intPos01 = 0: intPos02 = 0: intCnt = 0
            If strLine <> vbNullString Then
              If Left(strLine, 1) <> "'" Then
                intPos01 = InStr(strLine, "'")
                If intPos01 > 0 Then strLine = Trim(Left(strLine, (intPos01 - 1)))  ' ** Remove any remarks at end of line.
                intPos01 = InStr(strLine, " ")
                intPos02 = InStr((intPos01 + 2), strLine, " ")
                If intPos01 > 0 And intPos02 > 0 Then
                  strTmp01 = Trim(Left(strLine, intPos02))
                  Select Case strTmp01
                  Case "Public Const", "Private Const", "Public Type", "Private Type", "Private Declare", "Public Declare", _
                      "Public Enum", "Private Enum"
                    ' ** Not looking for these.
                  Case "Public Sub", "Public Function", "Private Sub", "Private Function", "Public Property"
                    ' ** Check for parameters.
                    intPos01 = InStr(strLine, "(")
                    If intPos01 > 0 Then
                      strTmp02 = Mid(strLine, (intPos01 + 1))
                      intPos02 = InStr(strTmp01, " ")
                      strTmp01 = "Local"
                      intPos03 = InStr(strTmp02, "(")  ' ** A 2nd opening paren.
                      If intPos03 > 0 Then
                        ' ** Private Sub StringToByte(InString As String, ByteArray() As Byte)
                        ' ** Public Function GUID2ByteArray(ByVal strGUID As String) As Byte()
                        If InStr(strLine, "Function") > 0 And Right(strLine, 9) = "As Byte()" Then
                          intPos02 = InStr(strTmp02, ")")
                        Else
                          intPos03 = InStr(intPos03, strTmp02, ")")
                          If InStr(intPos03, strTmp02, "(") = 0 Then  ' ** No 3rd opening paren.
                            intPos02 = InStr((intPos03 + 1), strTmp02, ")")
                          Else
                            Debug.Print "'HOW MANY?  " & strLine
                            Stop
                          End If
                        End If
                      Else
                        intPos02 = InStr(strTmp02, ")")
                      End If
                      If (intPos02 > 0) Then
                        If (intPos02 > (intPos01 + 1)) Then
                          strTmp02 = Left(strTmp02, (intPos02 - 1))
                          strTmp02 = StringReplace(strTmp02, "ByRef ", vbNullString)  ' ** Module Function: modStringFuncs.
                          strTmp02 = StringReplace(strTmp02, "ByVal ", vbNullString)  ' ** Module Function: modStringFuncs.
                          intPos03 = InStr(strTmp02, "Optional")
                          If intPos03 > 0 Then
                            blnOptional = True
                            strTmp02 = StringReplace(strTmp02, "Optional ", vbNullString)  ' ** Module Function: modStringFuncs.
                          Else
                            blnOptional = False
                          End If
                          intCnt = CharCnt(strTmp02, ",") + 1  ' ** Module Function: modStringFuncs.
                          Select Case intCnt
                          Case 1
                            intPos01 = InStr(strTmp02, " As ")
                            If intPos01 > 0 Then
                              strTmp03 = Trim(Mid(strTmp02, (intPos01 + 3)))
                              strTmp02 = Trim(Left(strTmp02, intPos01))
                              strProcName = .ProcOfLine((lngX + 1&), vbext_pk_Proc)
                              If strProcName = vbNullString Then strProcName = "Declaration"
                              blnFound = False
                              For lngY = 0& To (lngLVars - 1&)
                                If arr_varLVar(V_VNAM, lngY) = strTmp02 And arr_varLVar(V_PARM, lngY) = True And _
                                    arr_varLVar(V_OPT, lngY) = blnOptional Then
                                  blnFound = True
                                  arr_varLVar(V_CNT, lngY) = arr_varLVar(V_CNT, lngY) + 1&
                                  If InStr(arr_varLVar(V_MOD, lngY), strModName & ";") = 0 Then
                                    arr_varLVar(V_MOD, lngY) = arr_varLVar(V_MOD, lngY) & strModName & ";"
                                  End If
                                  If InStr(arr_varLVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                    arr_varLVar(V_PROC, lngY) = arr_varLVar(V_PROC, lngY) & strProcName & ";"
                                  End If
                                  Exit For
                                End If
                              Next
                              If blnFound = False Then
                                lngLVars = lngLVars + 1&
                                lngE = (lngLVars - 1&)
                                ReDim Preserve arr_varLVar(V_ELEMS, lngE)
                                arr_varLVar(V_VNAM, lngE) = strTmp02
                                arr_varLVar(V_MOD, lngE) = strModName & ";"
                                arr_varLVar(V_PROC, lngE) = strProcName & ";"
                                arr_varLVar(V_SCOP, lngE) = strTmp01
                                intPos03 = InStr(strTmp03, "=")
                                If intPos03 > 0 Then
                                  ' ** Boolean = False
                                  strTmp03 = Trim(Left(strTmp03, (intPos03 - 1)))
                                End If
                                strVarType = strTmp03
                                arr_varLVar(V_TYPE, lngE) = strVarType
                                arr_varLVar(V_CNT, lngE) = CLng(1)
                                arr_varLVar(V_PARM, lngE) = CBool(True)
                                arr_varLVar(V_OPT, lngE) = blnOptional
                              End If  ' ** blnFound.
                            Else
                              ' ** Untyped variable!
                              Debug.Print "'UNTYPED VAR!  " & strTmp03
                              Stop
                            End If
                          Case Else
                            strTmp04 = strTmp02
                            For intZ = 1 To intCnt
                              strTmp02 = vbNullString: strTmp03 = vbNullString
                              If intZ < intCnt Then
                                intPos01 = InStr(strTmp04, ",")
                                strTmp03 = Left(strTmp04, (intPos01 - 1))  ' ** This declaration only.
                                strTmp04 = Trim(Mid(strTmp04, (intPos01 + 1)))  ' ** Remainder of line.
                              Else
                                ' ** Last declaration on line.
                                strTmp03 = strTmp04
                                strTmp04 = vbNullString
                              End If
                              intPos01 = InStr(strTmp03, " As ")
                              If intPos01 > 0 Then
                                strTmp02 = Trim(Left(strTmp03, intPos01))
                                strTmp03 = Trim(Mid(strTmp03, (intPos01 + 3)))
                                strProcName = .ProcOfLine((lngX + 1&), vbext_pk_Proc)
                                If strProcName = vbNullString Then strProcName = "Declaration"
                                blnFound = False
                                For lngY = 0& To (lngLVars - 1&)
                                  If arr_varLVar(V_VNAM, lngY) = strTmp02 And arr_varLVar(V_PARM, lngY) = True And _
                                      arr_varLVar(V_OPT, lngY) = blnOptional Then
                                    blnFound = True
                                    arr_varLVar(V_CNT, lngY) = arr_varLVar(V_CNT, lngY) + 1&
                                    If InStr(arr_varLVar(V_MOD, lngY), strModName & ";") = 0 Then
                                      arr_varLVar(V_MOD, lngY) = arr_varLVar(V_MOD, lngY) & strModName & ";"
                                    End If
                                    If InStr(arr_varLVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                      arr_varLVar(V_PROC, lngY) = arr_varLVar(V_PROC, lngY) & strProcName & ";"
                                    End If
                                    Exit For
                                  End If
                                Next
                                If blnFound = False Then
                                  lngLVars = lngLVars + 1&
                                  lngE = (lngLVars - 1&)
                                  ReDim Preserve arr_varLVar(V_ELEMS, lngE)
                                  arr_varLVar(V_VNAM, lngE) = strTmp02
                                  arr_varLVar(V_MOD, lngE) = strModName & ";"
                                  arr_varLVar(V_PROC, lngE) = strProcName & ";"
                                  arr_varLVar(V_SCOP, lngE) = strTmp01
                                  intPos03 = InStr(strTmp03, "=")
                                  If intPos03 > 0 Then
                                    ' ** Boolean = False
                                    strTmp03 = Trim(Left(strTmp03, (intPos03 - 1)))
                                  End If
                                  strVarType = strTmp03
                                  arr_varLVar(V_TYPE, lngE) = strVarType
                                  arr_varLVar(V_CNT, lngE) = CLng(1)
                                  arr_varLVar(V_PARM, lngE) = CBool(True)
                                  arr_varLVar(V_OPT, lngE) = blnOptional
                                End If  ' ** blnFound.
                              Else
                                ' ** Untyped variable!
                                Debug.Print "'UNTYPED VAR!  " & strTmp03
                                Stop
                              End If
                            Next  ' ** intZ
                          End Select
                        End If  ' ** No Params.
                      Else
                        ' ** 2-line parameters?
                        Debug.Print "'2-LINE PARAMS?"
                        Stop
                      End If  ' ** intPos02.
                    End If  ' ** intPos01.
                  Case Else
                    strTmp01 = Trim(Left(strLine, intPos01))  ' ** First word.
                    If strTmp01 = "Public" Or strTmp01 = "Private" Or strTmp01 = "Dim" Or strTmp01 = "Static" Then
                      strTmp02 = Trim(Mid(strLine, intPos01))
                      intPos02 = InStr(strTmp02, " ")
                      blnOptional = False
                      If intPos02 > 0 Then
                        strTmp03 = Trim(Mid(strTmp02, intPos02))
                        strTmp02 = Trim(Left(strTmp02, intPos02))
                        strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                        If strProcName = vbNullString Then strProcName = "Declaration"
                        intCnt = CharCnt(strLine, ",") + 1
                        If InStr(strTmp02, "(") > 0 And InStr(strTmp03, ")") > 0 And InStr(strTmp03, "As ") > InStr(strTmp03, ")") Then
                          ' ** arr_typMDS(1000 To 3000, 1 To 14) As MONTHDAYSTATE
                          strTmp02 = Trim(Mid(strLine, intPos01))
                          intPos02 = InStr(strTmp02, " As ")
                          strTmp03 = Trim(Mid(strTmp02, intPos02))
                          strTmp02 = Trim(Left(strTmp02, intPos02))
                          If InStr(strTmp03, ",") > 0 Then
                            Stop
                          Else
                            intCnt = 1
                          End If
                        End If
                        If intCnt = 1 Then
                          blnFound = False
                          Select Case strTmp01
                          Case "Public"  ' ** Public, Global, application level.
                            For lngY = 0& To (lngGVars - 1&)
                              If arr_varGVar(V_VNAM, lngY) = strTmp02 Then
                                blnFound = True
                                arr_varGVar(V_CNT, lngY) = arr_varGVar(V_CNT, lngY) + 1&
                                If InStr(arr_varGVar(V_MOD, lngY), strModName & ";") = 0 Then
                                  arr_varGVar(V_MOD, lngY) = arr_varGVar(V_MOD, lngY) & strModName & ";"
                                End If
                                If InStr(arr_varGVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                  arr_varGVar(V_PROC, lngY) = arr_varGVar(V_PROC, lngY) & strProcName & ";"
                                End If
                                Exit For
                              End If
                            Next
                          Case "Private"  ' ** Private, module level.
                            For lngY = 0& To (lngPVars - 1&)
                              If arr_varPVar(V_VNAM, lngY) = strTmp02 Then
                                blnFound = True
                                arr_varPVar(V_CNT, lngY) = arr_varPVar(V_CNT, lngY) + 1&
                                If InStr(arr_varPVar(V_MOD, lngY), strModName & ";") = 0 Then
                                  arr_varPVar(V_MOD, lngY) = arr_varPVar(V_MOD, lngY) & strModName & ";"
                                End If
                                If InStr(arr_varPVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                  arr_varPVar(V_PROC, lngY) = arr_varPVar(V_PROC, lngY) & strProcName & ";"
                                End If
                                Exit For
                              End If
                            Next
                          Case "Dim", "Static"  ' ** Local, procedure level.
                            For lngY = 0& To (lngLVars - 1&)
                              If arr_varLVar(V_VNAM, lngY) = strTmp02 And arr_varLVar(V_PARM, lngY) = False Then
                                blnFound = True
                                arr_varLVar(V_CNT, lngY) = arr_varLVar(V_CNT, lngY) + 1&
                                If InStr(arr_varLVar(V_MOD, lngY), strModName & ";") = 0 Then
                                  arr_varLVar(V_MOD, lngY) = arr_varLVar(V_MOD, lngY) & strModName & ";"
                                End If
                                If InStr(arr_varLVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                  arr_varLVar(V_PROC, lngY) = arr_varLVar(V_PROC, lngY) & strProcName & ";"
                                End If
                                Exit For
                              End If
                            Next
                          End Select  ' ** strTmp01.
                          If blnFound = False Then
                            Select Case strTmp01
                            Case "Public"
                              lngGVars = lngGVars + 1&
                              lngE = (lngGVars - 1&)
                              ReDim Preserve arr_varGVar(V_ELEMS, lngE)
                              arr_varGVar(V_VNAM, lngE) = strTmp02
                              arr_varGVar(V_MOD, lngE) = strModName & ";"
                              arr_varGVar(V_PROC, lngE) = strProcName & ";"
                              arr_varGVar(V_SCOP, lngE) = strTmp01
                              If Left(strTmp03, 3) = "As " Then
                                strVarType = Mid(strTmp03, 4)
                              Else
                                Stop
                              End If
                              arr_varGVar(V_TYPE, lngE) = strVarType
                              arr_varGVar(V_CNT, lngE) = CLng(1)
                              arr_varGVar(V_PARM, lngE) = CBool(False)
                              arr_varGVar(V_OPT, lngE) = blnOptional
                            Case "Private"
                              lngPVars = lngPVars + 1&
                              lngE = (lngPVars - 1&)
                              ReDim Preserve arr_varPVar(V_ELEMS, lngE)
                              arr_varPVar(V_VNAM, lngE) = strTmp02
                              arr_varPVar(V_MOD, lngE) = strModName & ";"
                              arr_varPVar(V_PROC, lngE) = strProcName & ";"
                              arr_varPVar(V_SCOP, lngE) = strTmp01
                              If Left(strTmp03, 3) = "As " Then
                                strVarType = Mid(strTmp03, 4)
                              Else
                                Stop
                              End If
                              arr_varPVar(V_TYPE, lngE) = strVarType
                              arr_varPVar(V_CNT, lngE) = CLng(1)
                              arr_varPVar(V_PARM, lngE) = CBool(False)
                              arr_varPVar(V_OPT, lngE) = blnOptional
                            Case "Dim", "Static"
                              lngLVars = lngLVars + 1&
                              lngE = (lngLVars - 1&)
                              ReDim Preserve arr_varLVar(V_ELEMS, lngE)
                              arr_varLVar(V_VNAM, lngE) = strTmp02
                              arr_varLVar(V_MOD, lngE) = strModName & ";"
                              arr_varLVar(V_PROC, lngE) = strProcName & ";"
                              arr_varLVar(V_SCOP, lngE) = strTmp01
                              If Left(strTmp03, 3) = "As " Then
                                strVarType = Mid(strTmp03, 4)
                              Else
                                Stop
                              End If
                              arr_varLVar(V_TYPE, lngE) = strVarType
                              arr_varLVar(V_CNT, lngE) = CLng(1)
                              arr_varLVar(V_PARM, lngE) = CBool(False)
                              arr_varLVar(V_OPT, lngE) = blnOptional
                            End Select  ' ** strTmp01.
                          End If  ' ** blnFound.
                        Else
                          ' ** Multiple declarations per line.
                          ' ** strTmp01: First word.
                          strTmp04 = Trim(Mid(strLine, intPos01))  ' ** Everything after first space.
                          For intZ = 1 To intCnt
                            strTmp02 = vbNullString: strTmp03 = vbNullString
                            If intZ < intCnt Then
                              intPos01 = InStr(strTmp04, ",")
                              strTmp03 = Left(strTmp04, (intPos01 - 1))  ' ** This declaration only.
                              strTmp04 = Trim(Mid(strTmp04, (intPos01 + 1)))  ' ** Remainder of line.
                            Else
                              ' ** Last declaration on line.
                              strTmp03 = strTmp04
                              strTmp04 = vbNullString
                            End If
                            intPos01 = InStr(strTmp03, " ")
                            If intPos01 > 0 Then
                              strTmp02 = Trim(Left(strTmp03, intPos01))  ' ** Variable name.
                              strTmp03 = Trim(Mid(strTmp03, intPos01))
                              If Left(strTmp03, 3) = "As " Then
                                strVarType = Mid(strTmp03, 4)  ' ** Variable type.
                                blnFound = False
                                Select Case strTmp01
                                Case "Public"  ' ** Public, Global, application level.
                                  For lngY = 0& To (lngGVars - 1&)
                                    If arr_varGVar(V_VNAM, lngY) = strTmp02 Then
                                      blnFound = True
                                      arr_varGVar(V_CNT, lngY) = arr_varGVar(V_CNT, lngY) + 1&
                                      If InStr(arr_varGVar(V_MOD, lngY), strModName & ";") = 0 Then
                                        arr_varGVar(V_MOD, lngY) = arr_varGVar(V_MOD, lngY) & strModName & ";"
                                      End If
                                      If InStr(arr_varGVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                        arr_varGVar(V_PROC, lngY) = arr_varGVar(V_PROC, lngY) & strProcName & ";"
                                      End If
                                      Exit For
                                    End If
                                  Next
                                Case "Private"  ' ** Private, module level.
                                  For lngY = 0& To (lngPVars - 1&)
                                    If arr_varPVar(V_VNAM, lngY) = strTmp02 Then
                                      blnFound = True
                                      arr_varPVar(V_CNT, lngY) = arr_varPVar(V_CNT, lngY) + 1&
                                      If InStr(arr_varPVar(V_MOD, lngY), strModName & ";") = 0 Then
                                        arr_varPVar(V_MOD, lngY) = arr_varPVar(V_MOD, lngY) & strModName & ";"
                                      End If
                                      If InStr(arr_varPVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                        arr_varPVar(V_PROC, lngY) = arr_varPVar(V_PROC, lngY) & strProcName & ";"
                                      End If
                                      Exit For
                                    End If
                                  Next
                                Case "Dim", "Static"  ' ** Local, procedure level.
                                  For lngY = 0& To (lngLVars - 1&)
                                    If arr_varLVar(V_VNAM, lngY) = strTmp02 And arr_varLVar(V_PARM, lngY) = False Then
                                      blnFound = True
                                      arr_varLVar(V_CNT, lngY) = arr_varLVar(V_CNT, lngY) + 1&
                                      If InStr(arr_varLVar(V_MOD, lngY), strModName & ";") = 0 Then
                                        arr_varLVar(V_MOD, lngY) = arr_varLVar(V_MOD, lngY) & strModName & ";"
                                      End If
                                      If InStr(arr_varLVar(V_PROC, lngY), strProcName & ";") = 0 Then
                                        arr_varLVar(V_PROC, lngY) = arr_varLVar(V_PROC, lngY) & strProcName & ";"
                                      End If
                                      Exit For
                                    End If
                                  Next
                                End Select  ' ** strTmp01.
                                If blnFound = False Then
                                  Select Case strTmp01
                                  Case "Public"
                                    lngGVars = lngGVars + 1&
                                    lngE = (lngGVars - 1&)
                                    ReDim Preserve arr_varGVar(V_ELEMS, lngE)
                                    arr_varGVar(V_VNAM, lngE) = strTmp02
                                    arr_varGVar(V_MOD, lngE) = strModName & ";"
                                    arr_varGVar(V_PROC, lngE) = strProcName & ";"
                                    arr_varGVar(V_SCOP, lngE) = strTmp01
                                    arr_varGVar(V_TYPE, lngE) = strVarType
                                    arr_varGVar(V_CNT, lngE) = CLng(1)
                                    arr_varGVar(V_PARM, lngE) = CBool(False)
                                    arr_varGVar(V_OPT, lngE) = blnOptional
                                  Case "Private"
                                    lngPVars = lngPVars + 1&
                                    lngE = (lngPVars - 1&)
                                    ReDim Preserve arr_varPVar(V_ELEMS, lngE)
                                    arr_varPVar(V_VNAM, lngE) = strTmp02
                                    arr_varPVar(V_MOD, lngE) = strModName & ";"
                                    arr_varPVar(V_PROC, lngE) = strProcName & ";"
                                    arr_varPVar(V_SCOP, lngE) = strTmp01
                                    arr_varPVar(V_TYPE, lngE) = strVarType
                                    arr_varPVar(V_CNT, lngE) = CLng(1)
                                    arr_varPVar(V_PARM, lngE) = CBool(False)
                                    arr_varPVar(V_OPT, lngE) = blnOptional
                                  Case "Dim", "Static"
                                    lngLVars = lngLVars + 1&
                                    lngE = (lngLVars - 1&)
                                    ReDim Preserve arr_varLVar(V_ELEMS, lngE)
                                    arr_varLVar(V_VNAM, lngE) = strTmp02
                                    arr_varLVar(V_MOD, lngE) = strModName & ";"
                                    arr_varLVar(V_PROC, lngE) = strProcName & ";"
                                    arr_varLVar(V_SCOP, lngE) = strTmp01
                                    arr_varLVar(V_TYPE, lngE) = strVarType
                                    arr_varLVar(V_CNT, lngE) = CLng(1)
                                    arr_varLVar(V_PARM, lngE) = CBool(False)
                                    arr_varLVar(V_OPT, lngE) = blnOptional
                                  End Select  ' ** strTmp01.
                                End If  ' ** blnFound.
                              Else
                                Debug.Print "'WHAT?  " & strTmp03
                                Stop
                              End If  ' ** As.
                            Else
                              ' ** Untyped variable!
                              Debug.Print "'UNTYPED VAR!  " & strTmp03
                              Stop
                            End If
                          Next  ' ** intZ.
                        End If  ' ** intCnt.
                      End If  ' ** intPos02.
                    End If  ' ** Public, Private, Dim.
                  End Select  ' ** Sub, Function.
                End If  ' ** intPos01, intPos02.
              End If  ' ** Remark.
            End If  ' ** vbNullString.
          Next  ' ** lngX.
        End With  ' ** cod.
        Set cod = Nothing
      End With  ' ** vbc.
      Set vbc = Nothing
    Next  ' * ** vbc.
  End With  ' ** vbp.
  Set vbp = Nothing

  Debug.Print "'PUBLIC VARS:  " & CStr(lngGVars)
  DoEvents
  Debug.Print "'PRIVATE VARS: " & CStr(lngPVars)
  DoEvents
  Debug.Print "'LOCAL VARS:   " & CStr(lngLVars)
  DoEvents

  If lngGVars > 0& Then
    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("zz_tbl_VBComponent_Variable", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngGVars - 1&)
          .AddNew
          ' ** ![vbcomvar_id] : AutoNumber.
          ![vbcomvar_name] = arr_varGVar(V_VNAM, lngX)
          ![vbcomvar_module] = arr_varGVar(V_MOD, lngX)
          ![vbcomvar_procedure] = arr_varGVar(V_PROC, lngX)
          ![vbcomvar_scope] = arr_varGVar(V_SCOP, lngX)
          If arr_varGVar(V_SCOP, lngX) = "Dim" Then
            ![scopetype_type] = "Local"
          Else
            ![scopetype_type] = arr_varGVar(V_SCOP, lngX)
          End If
          If arr_varGVar(V_TYPE, lngX) = "DAO.index" Then
            arr_varGVar(V_TYPE, lngX) = "DAO.Index"
          End If
          If Left(arr_varGVar(V_TYPE, lngX), 4) = "New " Then
            arr_varGVar(V_TYPE, lngX) = Trim(Mid(arr_varGVar(V_TYPE, lngX), 5))
          End If
          intPos03 = InStr(arr_varGVar(V_TYPE, lngX), " * ")
          If intPos03 > 0 Then
            ' ** String * 255
            arr_varGVar(V_TYPE, lngX) = Trim(Left(arr_varGVar(V_TYPE, lngX), intPos03))
          End If
          ![vbcomvar_type] = arr_varGVar(V_TYPE, lngX)
          intPos03 = InStr(arr_varGVar(V_TYPE, lngX), ".")
          If intPos03 > 0 Then
            ![vbcomvar_object] = True
          Else
            Select Case arr_varGVar(V_TYPE, lngX)
            Case "AccessObject", "CodeModule", "CommandBarButton", "CurrentProject", "Drive", "encrypt", "Module", _
                "Object", "VBComponent", "VBProject", "RELBLOB", "RELWINDOW", "clsDevice", "clsDevices", "clsMonthCal"
              ![vbcomvar_object] = True
            Case Else
              ![vbcomvar_object] = False
            End Select
          End If
          If intPos03 > 0 Then
            ![vbcomvar_explicit] = True
            ![vbcomvar_objclass] = Left(arr_varGVar(V_TYPE, lngX), (intPos03 - 1))
            ![vbcomvar_objtype] = Mid(arr_varGVar(V_TYPE, lngX), (intPos03 + 1))
          Else
            ![vbcomvar_explicit] = False
            ![vbcomvar_objclass] = Null
            ![vbcomvar_objtype] = Null
          End If
          If IsUC(arr_varGVar(V_TYPE, lngX), True, True) = True Or Left(arr_varGVar(V_TYPE, lngX), 3) = "cls" Then   ' ** Module Function: modStringFuncs.
            ![vbcomvar_userdefined] = True
          Else
            ![vbcomvar_userdefined] = False
          End If
          If Left(arr_varGVar(V_TYPE, lngX), 3) = "cls" Then
            ![vbcomvar_class] = True
          Else
            ![vbcomvar_class] = False
          End If
          ![vbcomvar_parameter] = arr_varGVar(V_PARM, lngX)
          ![vbcomvar_optional] = arr_varGVar(V_OPT, lngX)
          ![vbcomvar_count] = arr_varGVar(V_CNT, lngX)
          ![vbcomvar_datemodified] = Now()
          .Update
        Next  ' ** lngX.
      End With
      Set rst = Nothing
      .Close
    End With
    Set dbs = Nothing
  End If

  If lngPVars > 0& Then
    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("zz_tbl_VBComponent_Variable", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngPVars - 1&)
          .AddNew
          ' ** ![vbcomvar_id] : AutoNumber.
          ![vbcomvar_name] = arr_varPVar(V_VNAM, lngX)
          ![vbcomvar_module] = arr_varPVar(V_MOD, lngX)
          ![vbcomvar_procedure] = arr_varPVar(V_PROC, lngX)
          ![vbcomvar_scope] = arr_varPVar(V_SCOP, lngX)
          If arr_varPVar(V_SCOP, lngX) = "Dim" Then
            ![scopetype_type] = "Local"
          Else
            ![scopetype_type] = arr_varPVar(V_SCOP, lngX)
          End If
          If arr_varPVar(V_TYPE, lngX) = "DAO.index" Then
            arr_varPVar(V_TYPE, lngX) = "DAO.Index"
          End If
          If Left(arr_varPVar(V_TYPE, lngX), 4) = "New " Then
            arr_varPVar(V_TYPE, lngX) = Trim(Mid(arr_varPVar(V_TYPE, lngX), 5))
          End If
          intPos03 = InStr(arr_varPVar(V_TYPE, lngX), " * ")
          If intPos03 > 0 Then
            ' ** String * 255
            arr_varPVar(V_TYPE, lngX) = Trim(Left(arr_varPVar(V_TYPE, lngX), intPos03))
          End If
          ![vbcomvar_type] = arr_varPVar(V_TYPE, lngX)
          intPos03 = InStr(arr_varPVar(V_TYPE, lngX), ".")
          If intPos03 > 0 Then
            ![vbcomvar_object] = True
          Else
            Select Case arr_varPVar(V_TYPE, lngX)
            Case "AccessObject", "CodeModule", "CommandBarButton", "CurrentProject", "Drive", "encrypt", "Module", _
                "Object", "VBComponent", "VBProject", "RELBLOB", "RELWINDOW", "clsDevice", "clsDevices", "clsMonthCal"
              ![vbcomvar_object] = True
            Case Else
              ![vbcomvar_object] = False
            End Select
          End If
          If intPos03 > 0 Then
            ![vbcomvar_explicit] = True
            ![vbcomvar_objclass] = Left(arr_varPVar(V_TYPE, lngX), (intPos03 - 1))
            ![vbcomvar_objtype] = Mid(arr_varPVar(V_TYPE, lngX), (intPos03 + 1))
          Else
            ![vbcomvar_explicit] = False
            ![vbcomvar_objclass] = Null
            ![vbcomvar_objtype] = Null
          End If
          If IsUC(arr_varPVar(V_TYPE, lngX), True, True) = True Or Left(arr_varPVar(V_TYPE, lngX), 3) = "cls" Then  ' ** Module Function: modStringFuncs.
            ![vbcomvar_userdefined] = True
          Else
            ![vbcomvar_userdefined] = False
          End If
          If Left(arr_varPVar(V_TYPE, lngX), 3) = "cls" Then
            ![vbcomvar_class] = True
          Else
            ![vbcomvar_class] = False
          End If
          ![vbcomvar_parameter] = arr_varPVar(V_PARM, lngX)
          ![vbcomvar_optional] = arr_varPVar(V_OPT, lngX)
          ![vbcomvar_count] = arr_varPVar(V_CNT, lngX)
          ![vbcomvar_datemodified] = Now()
          .Update
        Next  ' ** lngX.
      End With
      Set rst = Nothing
      .Close
    End With
    Set dbs = Nothing
  End If

  If lngLVars > 0& Then
    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("zz_tbl_VBComponent_Variable", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngLVars - 1&)
          .AddNew
          ' ** ![vbcomvar_id] : AutoNumber.
          ![vbcomvar_name] = arr_varLVar(V_VNAM, lngX)
          ![vbcomvar_module] = arr_varLVar(V_MOD, lngX)
          ![vbcomvar_procedure] = arr_varLVar(V_PROC, lngX)
          ![vbcomvar_scope] = arr_varLVar(V_SCOP, lngX)
          If arr_varLVar(V_SCOP, lngX) = "Dim" Then
            ![scopetype_type] = "Local"
          Else
            ![scopetype_type] = arr_varLVar(V_SCOP, lngX)
          End If
          If arr_varLVar(V_TYPE, lngX) = "DAO.index" Then
            arr_varLVar(V_TYPE, lngX) = "DAO.Index"
          End If
          If Left(arr_varLVar(V_TYPE, lngX), 4) = "New " Then
            arr_varLVar(V_TYPE, lngX) = Trim(Mid(arr_varLVar(V_TYPE, lngX), 4))
          End If
          intPos03 = InStr(arr_varLVar(V_TYPE, lngX), " * ")
          If intPos03 > 0 Then
            ' ** String * 255
            arr_varLVar(V_TYPE, lngX) = Trim(Left(arr_varLVar(V_TYPE, lngX), intPos03))
          End If
          ![vbcomvar_type] = arr_varLVar(V_TYPE, lngX)
          intPos03 = InStr(arr_varLVar(V_TYPE, lngX), ".")
          If intPos03 > 0 Then
            ![vbcomvar_object] = True
          Else
            Select Case arr_varLVar(V_TYPE, lngX)
            Case "AccessObject", "CodeModule", "CommandBarButton", "CurrentProject", "Drive", "encrypt", "Module", _
                "Object", "VBComponent", "VBProject", "RELBLOB", "RELWINDOW", "clsDevice", "clsDevices", "clsMonthCal"
              ![vbcomvar_object] = True
            Case Else
              ![vbcomvar_object] = False
            End Select
          End If
          If intPos03 > 0 Then
            ![vbcomvar_explicit] = True
            ![vbcomvar_objclass] = Left(arr_varLVar(V_TYPE, lngX), (intPos03 - 1))
            ![vbcomvar_objtype] = Mid(arr_varLVar(V_TYPE, lngX), (intPos03 + 1))
          Else
            ![vbcomvar_explicit] = False
            ![vbcomvar_objclass] = Null
            ![vbcomvar_objtype] = Null
          End If
          If IsUC(arr_varLVar(V_TYPE, lngX), True, True) = True Or Left(arr_varLVar(V_TYPE, lngX), 3) = "cls" Then  ' ** Module Function: modStringFuncs.
            ![vbcomvar_userdefined] = True
          Else
            ![vbcomvar_userdefined] = False
          End If
          If Left(arr_varLVar(V_TYPE, lngX), 3) = "cls" Then
            ![vbcomvar_class] = True
          Else
            ![vbcomvar_class] = False
          End If
          ![vbcomvar_parameter] = arr_varLVar(V_PARM, lngX)
          ![vbcomvar_optional] = arr_varLVar(V_OPT, lngX)
          ![vbcomvar_count] = arr_varLVar(V_CNT, lngX)
          ![vbcomvar_datemodified] = Now()
          .Update
        Next  ' ** lngX.
      End With
      Set rst = Nothing
      .Close
    End With
    Set dbs = Nothing
  End If

  Debug.Print "'TOTAL VARS:   " & CStr(lngGVars + lngPVars + lngLVars)
  DoEvents

  Beep

'PUBLIC VARS:  223
'PRIVATE VARS: 1466
'LOCAL VARS:   3231
'TOTAL VARS:   4920
'DONE!

'PUBLIC VARS:  223
'PRIVATE VARS: 1466
'LOCAL VARS:   3218
'TOTAL VARS:   4907
'DONE!

'PUBLIC VARS:  223
'PRIVATE VARS: 1466
'LOCAL VARS:   3153
'TOTAL VARS:   4842
'DONE!

'PUBLIC VARS:  223
'PRIVATE VARS: 1466
'LOCAL VARS:   3000
'TOTAL VARS:   4689
'DONE!

'PUBLIC VARS:  223
'PRIVATE VARS: 1469
'LOCAL VARS:   3016
'TOTAL VARS:   4708
'DONE!
  Debug.Print "'DONE!"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set dbs = Nothing
  Var_Doc = blnRetVal

End Function

Public Function Const_Doc() As Boolean

  Const THIS_PROC As String = "Const_Doc"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim lngLines As Long, lngDecLines As Long
  Dim strLine As String, strModName As String, strProcName As String, strVarType As String
  Dim lngLCons As Long, arr_varLCon() As Variant
  Dim lngPCons As Long, arr_varPCon() As Variant
  Dim lngGCons As Long, arr_varGCon() As Variant
  Dim lngScopeType_Public As Long, lngScopeType_Private As Long, lngScopeType_Local As Long
  Dim blnFound As Boolean, blnSkip As Boolean
  Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer ', intCnt As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
  Dim lngX As Long, lngY As Long, intZ As Integer, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varLCon(), arr_varPCon(), arr_varGCon().
  Const C_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const C_CNAM As Integer = 0
  Const C_MOD  As Integer = 1
  Const C_PROC As Integer = 2
  Const C_SCOP As Integer = 3
  Const C_SID  As Integer = 4
  Const C_TYPE As Integer = 5
  Const C_CNT  As Integer = 6
  Const C_VAL  As Integer = 7

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngLCons = 0&
  ReDim arr_varLCon(C_ELEMS, 0)
  lngPCons = 0&
  ReDim arr_varPCon(C_ELEMS, 0)
  lngGCons = 0&
  ReDim arr_varGCon(C_ELEMS, 0)

  blnSkip = False
  If blnSkip = False Then
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
              strLine = .Lines(lngX, 1)
              strLine = Trim(strLine)
              strProcName = vbNullString
              If strLine <> vbNullString Then
                If Left(strLine, 1) <> "'" Then
                  intPos01 = InStr(strLine, "'")
                  If intPos01 > 0 Then strLine = Trim(Left(strLine, (intPos01 - 1)))  ' ** Remove any remarks at end of line.
                  intPos01 = InStr(strLine, " ")                  ' ** First space.
                  intPos02 = InStr((intPos01 + 2), strLine, " ")  ' ** Second space.
                  If intPos01 > 0 And intPos02 > 0 Then

                    strTmp01 = Trim(Left(strLine, intPos01))  ' ** First word.
                    strTmp02 = Trim(Left(strLine, intPos02))  ' ** First two words.

                    Select Case strTmp01
                    Case "Public", "Private", "Const"

                      blnFound = False
                      Select Case strTmp01
                      Case "Public", "Private"
                        If GetLastWord(strTmp02) = "Const" Then  ' ** Module Function: modStringFuncs.
                          blnFound = True
                        End If
                      Case "Const"
                        blnFound = True
                      End Select

                      If blnFound = True Then
                        blnFound = False
                        strProcName = .ProcOfLine((lngX + 1&), vbext_pk_Proc)
                        If strProcName = vbNullString Then strProcName = "Declaration"
                        Select Case strTmp01
                        Case "Public", "Private"
                          strTmp02 = Trim(Mid(strLine, intPos02))
                          intPos03 = InStr(strTmp02, " As ")
                          If intPos03 > 0 Then
                            strTmp03 = Trim(Mid(strTmp02, (intPos03 + 3)))
                            strTmp02 = Trim(Left(strTmp02, intPos03))
                            blnFound = True
                          Else
                            ' ** Untyped constant!
                            intPos03 = InStr(strTmp02, "=")
                            If intPos03 > 0 Then
                              strTmp03 = "{untyped} " & Mid(strTmp02, intPos03)
                              strTmp02 = Trim(Left(strTmp02, (intPos03 - 1)))
                              blnFound = True
                            Else
                              ' ** Untyped AND no value?
                              Debug.Print "'UNTYPED AND NO VAL CONST!  " & strLine
                              Stop
                            End If
                            'Debug.Print "'UNTYPED CONST!  " & strTmp02
                            'Stop
                          End If
                        Case "Const"
                          strTmp02 = Trim(Mid(strLine, intPos01))
                          intPos03 = InStr(strTmp02, " As ")
                          If intPos03 > 0 Then
                            strTmp03 = Trim(Mid(strTmp02, (intPos03 + 3)))
                            strTmp02 = Trim(Left(strTmp02, intPos03))
                            blnFound = True
                          Else
                            ' ** Untyped constant!
                            intPos03 = InStr(strTmp02, "=")
                            If intPos03 > 0 Then
                              strTmp03 = "{untyped} " & Mid(strTmp02, intPos03)
                              strTmp02 = Trim(Left(strTmp02, (intPos03 - 1)))
                              blnFound = True
                            Else
                              ' ** Untyped AND no value?
                              Debug.Print "'UNTYPED AND NO VAL CONST!  " & strLine
                              Stop
                            End If
                            'Debug.Print "'UNTYPED CONST!  " & strTmp02
                            'Stop
                          End If
                        End Select
                      End If  ' ** blnFound.

                      If blnFound = True Then
                        intPos01 = InStr(strTmp03, "=")
                        If intPos01 > 0 Then
                          strTmp04 = Trim(Mid(strTmp03, (intPos01 + 1)))
                          strTmp03 = Trim(Left(strTmp03, (intPos01 - 1)))
                          blnFound = False
                          Select Case strTmp01
                          Case "Public"
                            For lngY = 0& To (lngGCons - 1&)
                              If arr_varGCon(C_CNAM, lngY) = strTmp02 Then
                                blnFound = True
                                arr_varGCon(C_CNT, lngY) = arr_varGCon(C_CNT, lngY) + 1&
                                arr_varGCon(C_MOD, lngY) = arr_varGCon(C_MOD, lngY) & strModName & "~"
                                arr_varGCon(C_PROC, lngY) = arr_varGCon(C_PROC, lngY) & strProcName & "~"
                                arr_varGCon(C_VAL, lngY) = arr_varGCon(C_VAL, lngY) & strTmp04 & "~"
                                Exit For
                              End If
                            Next  ' ** lngY.
                            If blnFound = False Then
                              lngGCons = lngGCons + 1&
                              lngE = lngGCons - 1&
                              ReDim Preserve arr_varGCon(C_ELEMS, lngE)
                              arr_varGCon(C_CNAM, lngE) = strTmp02
                              arr_varGCon(C_MOD, lngE) = strModName & "~"
                              arr_varGCon(C_PROC, lngE) = strProcName & "~"
                              arr_varGCon(C_SCOP, lngE) = strTmp01
                              arr_varGCon(C_SID, lngE) = CLng(0)
                              arr_varGCon(C_TYPE, lngE) = strTmp03
                              arr_varGCon(C_CNT, lngE) = CLng(1)
                              arr_varGCon(C_VAL, lngE) = strTmp04 & "~"
                            End If  ' ** blnFound.
                          Case "Private"
                            For lngY = 0& To (lngPCons - 1&)
                              If arr_varPCon(C_CNAM, lngY) = strTmp02 Then
                                blnFound = True
                                arr_varPCon(C_CNT, lngY) = arr_varPCon(C_CNT, lngY) + 1&
                                arr_varPCon(C_MOD, lngY) = arr_varPCon(C_MOD, lngY) & strModName & "~"
                                arr_varPCon(C_PROC, lngY) = arr_varPCon(C_PROC, lngY) & strProcName & "~"
                                arr_varPCon(C_VAL, lngY) = arr_varPCon(C_VAL, lngY) & strTmp04 & "~"
                                Exit For
                              End If
                            Next  ' ** lngY.
                            If blnFound = False Then
                              lngPCons = lngPCons + 1&
                              lngE = lngPCons - 1&
                              ReDim Preserve arr_varPCon(C_ELEMS, lngE)
                              arr_varPCon(C_CNAM, lngE) = strTmp02
                              arr_varPCon(C_MOD, lngE) = strModName & "~"
                              arr_varPCon(C_PROC, lngE) = strProcName & "~"
                              arr_varPCon(C_SCOP, lngE) = strTmp01
                              arr_varPCon(C_SID, lngE) = CLng(0)
                              arr_varPCon(C_TYPE, lngE) = strTmp03
                              arr_varPCon(C_CNT, lngE) = CLng(1)
                              arr_varPCon(C_VAL, lngE) = strTmp04 & "~"
                            End If  ' ** blnFound.
                          Case "Const"
                            For lngY = 0& To (lngLCons - 1&)
                              If arr_varLCon(C_CNAM, lngY) = strTmp02 Then
                                blnFound = True
                                arr_varLCon(C_CNT, lngY) = arr_varLCon(C_CNT, lngY) + 1&
                                arr_varLCon(C_MOD, lngY) = arr_varLCon(C_MOD, lngY) & strModName & "~"
                                arr_varLCon(C_PROC, lngY) = arr_varLCon(C_PROC, lngY) & strProcName & "~"
                                arr_varLCon(C_VAL, lngY) = arr_varLCon(C_VAL, lngY) & strTmp04 & "~"
                                Exit For
                              End If
                            Next  ' ** lngY.
                            If blnFound = False Then
                              lngLCons = lngLCons + 1&
                              lngE = lngLCons - 1&
                              ReDim Preserve arr_varLCon(C_ELEMS, lngE)
                              arr_varLCon(C_CNAM, lngE) = strTmp02
                              arr_varLCon(C_MOD, lngE) = strModName & "~"
                              arr_varLCon(C_PROC, lngE) = strProcName & "~"
                              arr_varLCon(C_SCOP, lngE) = "Local"
                              arr_varLCon(C_SID, lngE) = CLng(0)
                              arr_varLCon(C_TYPE, lngE) = strTmp03
                              arr_varLCon(C_CNT, lngE) = CLng(1)
                              arr_varLCon(C_VAL, lngE) = strTmp04 & "~"
                            End If  ' ** blnFound.
                          End Select
                        Else
                          ' ** Shouldn't ever be no assigned value.
                          Debug.Print "'NO VAL!  " & strLine
                          Stop
                        End If
                      End If  ' ** blnFound.

                    Case Else
                      ' ** What else?
                      intPos01 = InStr(strLine, " Const ")
                      If intPos01 > 0 Then
                        intPos02 = InStr(strLine, Chr(34))  ' ** Quotes.
                        If intPos02 < intPos01 Then
                          ' ** Something within quotes.
                          ' ** 'strTmp01 = "Private Const THIS_NAME As String = " & Chr(34) & strTmp01 & Chr(34)'
                        Else
                          Debug.Print "'WHAT?  " & strLine
                          Stop
                        End If
                      End If
                    End Select  ' ** strTmp01.
                  End If  ' ** Space.
                End If  ' ** Remark.
              End If  ' ** vbNullString.
            Next  ' ** lngX.
          End With  ' ** cod.
          Set cod = Nothing
        End With  ' ** vbc.
      Next  ' ** vbc.
      Set vbc = Nothing
    End With  ' ** vbp.
    Set vbp = Nothing
  End If  ' ** blnSkip.

  Debug.Print "'TOTAL CONSTS:  " & CStr(lngGCons + lngPCons + lngLCons)
  DoEvents

  blnSkip = False
  If blnSkip = False Then
    If lngGCons > 0& Or lngPCons > 0& Or lngLCons > 0& Then

      Set dbs = CurrentDb
      With dbs

        Set rst = .OpenRecordset("tblScopeType", dbOpenDynaset, dbConsistent)
        With rst
          .MoveFirst
          .FindFirst "[scopetype_type] = 'Public'"
          If .NoMatch = False Then
            lngScopeType_Public = ![scopetype_id]
            .MoveFirst
            .FindFirst "[scopetype_type] = 'Private'"
            If .NoMatch = False Then
              lngScopeType_Private = ![scopetype_id]
              .MoveFirst
              .FindFirst "[scopetype_type] = 'Local'"
              If .NoMatch = False Then
                lngScopeType_Local = ![scopetype_id]
              Else
                Stop
              End If
            Else
              Stop
            End If
          Else
            Stop
          End If
          .Close
        End With
        Set rst = Nothing

        For lngX = 0& To (lngGCons - 1&)
          Select Case arr_varGCon(C_SCOP, lngX)
          Case "Public"
            arr_varGCon(C_SID, lngX) = lngScopeType_Public
          Case "Private"
            arr_varGCon(C_SID, lngX) = lngScopeType_Private
          Case "Local"
            arr_varGCon(C_SID, lngX) = lngScopeType_Local
          End Select
        Next  ' ** lngX.
        For lngX = 0& To (lngPCons - 1&)
          Select Case arr_varPCon(C_SCOP, lngX)
          Case "Public"
            arr_varPCon(C_SID, lngX) = lngScopeType_Public
          Case "Private"
            arr_varPCon(C_SID, lngX) = lngScopeType_Private
          Case "Local"
            arr_varPCon(C_SID, lngX) = lngScopeType_Local
          End Select
        Next  ' ** lngX.
        For lngX = 0& To (lngLCons - 1&)
          Select Case arr_varLCon(C_SCOP, lngX)
          Case "Public"
            arr_varLCon(C_SID, lngX) = lngScopeType_Public
          Case "Private"
            arr_varLCon(C_SID, lngX) = lngScopeType_Private
          Case "Local"
            arr_varLCon(C_SID, lngX) = lngScopeType_Local
          End Select
        Next  ' ** lngX.

        Set rst = .OpenRecordset("zz_tbl_VBComponent_Constant", dbOpenDynaset, dbAppendOnly)
        With rst

          If lngGCons > 0& Then
            For lngX = 0& To (lngGCons - 1&)
              .AddNew
              ' ** ![vbcomcon_id] : AutoNumber.
              ![vbcomcon_name] = arr_varGCon(C_CNAM, lngX)
              ![vbcomcon_module] = arr_varGCon(C_MOD, lngX)
              ![vbcomcon_procedure] = arr_varGCon(C_PROC, lngX)
              ![vbcomcon_scope] = arr_varGCon(C_SCOP, lngX)
              ![scopetype_id] = arr_varGCon(C_SID, lngX)
              ![vbcomcon_type] = arr_varGCon(C_TYPE, lngX)
              ![vbcomcon_value] = arr_varGCon(C_VAL, lngX)
              ![vbcomcon_userdefined] = False
              ![vbcomcon_count] = arr_varGCon(C_CNT, lngX)
              ![vbcomcon_datemodified] = Now()
              .Update
            Next  ' ** lngX.
          End If  ' ** lngGCons.

          If lngPCons > 0& Then
            For lngX = 0& To (lngPCons - 1&)
              .AddNew
              ' ** ![vbcomcon_id] : AutoNumber.
              ![vbcomcon_name] = arr_varPCon(C_CNAM, lngX)
              ![vbcomcon_module] = arr_varPCon(C_MOD, lngX)
              ![vbcomcon_procedure] = arr_varPCon(C_PROC, lngX)
              ![vbcomcon_scope] = arr_varPCon(C_SCOP, lngX)
              ![scopetype_id] = arr_varPCon(C_SID, lngX)
              ![vbcomcon_type] = arr_varPCon(C_TYPE, lngX)
              ![vbcomcon_value] = arr_varPCon(C_VAL, lngX)
              ![vbcomcon_userdefined] = False
              ![vbcomcon_count] = arr_varPCon(C_CNT, lngX)
              ![vbcomcon_datemodified] = Now()
              .Update
            Next  ' ** lngX.
          End If  ' ** lngPCons.

          If lngLCons > 0& Then
            For lngX = 0& To (lngLCons - 1&)
              .AddNew
              ' ** ![vbcomcon_id] : AutoNumber.
              ![vbcomcon_name] = arr_varLCon(C_CNAM, lngX)
              ![vbcomcon_module] = arr_varLCon(C_MOD, lngX)
              ![vbcomcon_procedure] = arr_varLCon(C_PROC, lngX)
              ![vbcomcon_scope] = arr_varLCon(C_SCOP, lngX)
              ![scopetype_id] = arr_varLCon(C_SID, lngX)
              ![vbcomcon_type] = arr_varLCon(C_TYPE, lngX)
              ![vbcomcon_value] = arr_varLCon(C_VAL, lngX)
              ![vbcomcon_userdefined] = False
              ![vbcomcon_count] = arr_varLCon(C_CNT, lngX)
              ![vbcomcon_datemodified] = Now()
              .Update
            Next  ' ** lngX.
          End If  ' ** lngLCons.

          .Close
        End With  ' ** rst.
        Set rst = Nothing
        .Close
      End With  ' ** dbs.
      Set dbs = Nothing

    End If  ' ** lngGCons, lngPCons, lngLCons.
  End If  ' ** blnSkip.

  Debug.Print "'PUBLIC CONST:  " & CLng(lngGCons)
  DoEvents
  Debug.Print "'PRIVATE CONST: " & CLng(lngPCons)
  DoEvents
  Debug.Print "'LOCAL CONST:   " & CLng(lngLCons)
  DoEvents

  'Vb.net will regard "&h"-notation hex constants in the range from 0x80000000-0xFFFFFFFF
  'as negative numbers unless the type is explicitly specified as UInt32, Int64, or UInt64.
  'In present versions of VB, one may force the number to be evaluated correctly by
  'using a suffix of "&" suffix (Int64), "L" (Int64), "UL" (UInt64), or "UI" (UInt32).

  'Try explicitly marking it as a Long with a trailing ampersand &:
  ' &HFFFF&
  '&HFFFF without the trailing ampersand is an Integer literal, and
  'Integer in VB6 is a signed 2-byte integer with a range of -32,768 to +32,767.

  'IIf([vbcomcon_numeric1]>=-32768 And [vbcomcon_numeric1] <= 32767,'vbInteger',
  'IIf([vbcomcon_numeric1]>=-2147483648 And [vbcomcon_numeric1]<=2147483647,'vbLong',
  'Neg: -3.402823E38 to -1.401298E-45 ,'vbSingle',
  'Pos: 1.401298E-45 to 3.402823E+38 ,'vbSingle',
  'Neg: -1.79769313486231E+308 to -4.94065645841247E-324 ,'vbDouble',
  'Pos: 4.94065645841247E-324 to 1.79769313486232E+308 ,'vbDouble',

  ' ** &H6 : vbInteger
  ' ** &HB& : vbLong
  ' ** &HFFFF : vbInteger
  ' ** &H800000 : vbLong
  ' ** &HEF000000 : vbLong


'TOTAL CONSTS:  2514
'PUBLIC CONST:  526
'PRIVATE CONST: 1087
'LOCAL CONST:   901
'DONE!
  Beep

  Debug.Print "'DONE!"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set dbs = Nothing
  Const_Doc = blnRetVal

End Function

Public Function ProcUsage_Doc() As Boolean

  Const THIS_PROC As String = "ProcUsage_Doc"

  Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngUses As Long, arr_varUse() As Variant
  Dim lngDels As Long, arr_varDel() As Variant  ' ** Can't think of another name.
  Dim lngLines As Long, lngDecLines As Long
  Dim strLine As String, strModName As String, strProcName As String, strVarType As String
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean
  Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intLen As Integer
  Dim strTmp01 As String, lngTmp02 As Long
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varProc().
  Const P_DID  As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_VID  As Integer = 2
  Const P_VNAM As Integer = 3
  Const P_CTYP As Integer = 4
  Const P_PID  As Integer = 5
  Const P_PNAM As Integer = 6
  Const P_STYP As Integer = 7
  Const P_PTYP As Integer = 8

  ' ** Array: arr_varUse().
  Const U_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const U_VID   As Integer = 0
  Const U_VNAM  As Integer = 1
  Const U_PID   As Integer = 2
  Const U_PNAM  As Integer = 3
  Const U_VIDU  As Integer = 4
  Const U_VNAMU As Integer = 5
  Const U_PIDU  As Integer = 6
  Const U_PNAMU As Integer = 7
  Const U_LIN   As Integer = 8
  Const U_CODE  As Integer = 9
  Const U_RAW   As Integer = 10

  ' ** Array: arr_varDel().
  Const D_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const D_FIND As Integer = 0
  Const D_VNAM As Integer = 1
  Const D_PNAM As Integer = 2
  Const D_LIN  As Integer = 3
  Const D_RAW  As Integer = 4

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** tblVBComponent_Procedure, just Public procedures in Standard Modules.
    Set qdf = .QueryDefs("zzz_qry_VBComponent_Proc_01_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngProcs = .RecordCount
      .MoveFirst
      arr_varProc = .GetRows(lngProcs)
      ' ***************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   Field  Element  Name              Constant
      ' **   =====  =======  ================  ==========
      ' **     1       0     dbs_id            P_DID
      ' **     2       1     dbs_name          P_DNAM
      ' **     3       2     vbcom_id          P_VID
      ' **     4       3     vbcom_name        P_VNAM
      ' **     5       4     comtype_type      P_CTYP
      ' **     6       5     vbcomproc_id      P_PID
      ' **     7       6     vbcomproc_name    P_PNAM
      ' **     8       7     scopetype_type    P_STYP
      ' **     9       8     proctype_type     P_PTYP
      ' **
      ' ***************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With  ' ** dbs
  Set dbs = Nothing

  Debug.Print "'PROCS: " & CStr(lngProcs)
  DoEvents

  If lngProcs > 0& Then

    lngUses = 0&
    ReDim arr_varUse(U_ELEMS, 0)

    lngDels = 0&
    ReDim arr_varDel(D_ELEMS, 0)

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For lngX = 0& To (lngProcs - 1&)
        For Each vbc In .VBComponents
          With vbc
            strModName = .Name
            If strModName <> arr_varProc(P_VNAM, lngX) And strModName <> "zz_mod_ModuleFormatFuncs" And strModName <> "zz_mod_VarConstDoc" Then
              Set cod = .CodeModule
              With cod
                lngLines = .CountOfLines
                lngDecLines = .CountOfDeclarationLines
                For lngY = 1& To lngLines
                  strLine = .Lines(lngY, 1)
                  strLine = Trim(strLine)
                  strProcName = vbNullString
                  blnFound = False
                  If strLine <> vbNullString Then
                    If Left(strLine, 1) <> "'" Then
                      intPos02 = InStr(strLine, "'")
                      intPos01 = InStr(strLine, arr_varProc(P_PNAM, lngX))
                      If intPos01 > 0 Then
                        If intPos02 = 0 Or intPos02 > intPos01 Then
                          intLen = Len(arr_varProc(P_PNAM, lngX))
                          strProcName = .ProcOfLine(lngY, vbext_pk_Proc)
                          If intPos01 > 1 Then
                            strTmp01 = Mid(strLine, (intPos01 - 1), 1)  ' ** Character before proc name.
                            lngTmp02 = Asc(strTmp01)
                            Select Case strTmp01
                            Case " ", ",", "(", ")"
                              ' ** OK.
                              blnFound = True
                            Case Chr(34)  ' ** Quotes.
                              ' ** No.
                            Case "_", "."
                              ' ** No.
                            Case ";", ":", "[", "]", "{", "}", "\", "|"
                              ' ** Unsure.
                              If (arr_varProc(P_PNAM, lngX) = "CoInfo" Or arr_varProc(P_PNAM, lngX) = "Proper") And _
                                  (Mid(strLine, (intPos01 - 2), 2) = "![" Or Mid(strLine, (intPos01 - 2), 2) = (Chr(34) & "[")) Then
                                ' ** No.
                              Else
                                lngDels = lngDels + 1&
                                lngE = lngDels - 1&
                                ReDim Preserve arr_varDel(D_ELEMS, lngE)
                                arr_varDel(D_FIND, lngE) = arr_varProc(P_PNAM, lngX)
                                arr_varDel(D_VNAM, lngE) = strModName
                                arr_varDel(D_PNAM, lngE) = strProcName
                                arr_varDel(D_LIN, lngE) = lngY
                                arr_varDel(D_RAW, lngE) = strLine
                                'Debug.Print "'" & strLine
                                'Stop
                              End If
                            Case Else
                              If lngTmp02 >= 65 And lngTmp02 <= 90 Then  ' ** A - Z.
                                ' ** Nope.
                              ElseIf lngTmp02 >= 97 And lngTmp02 <= 122 Then  ' ** a - z.
                                ' ** Nope.
                              ElseIf lngTmp02 >= 48 And lngTmp02 <= 57 Then  ' ** 0 - 9.
                                ' ** Nope.
                              ElseIf arr_varProc(P_PNAM, lngX) = "FormRef" And strTmp01 = "<" Then
                                ' ** No.
                              ElseIf arr_varProc(P_PNAM, lngX) = "Scr" And strTmp01 = "-" Then
                                ' ** No.
                              ElseIf arr_varProc(P_PNAM, lngX) = "Proper" And Mid(strLine, (intPos01 - 2), 2) = (Chr(34) & "&") Then
                                  ' ** No.
                              Else
                                ' ** Unsure.
                                lngDels = lngDels + 1&
                                lngE = lngDels - 1&
                                ReDim Preserve arr_varDel(D_ELEMS, lngE)
                                arr_varDel(D_FIND, lngE) = arr_varProc(P_PNAM, lngX)
                                arr_varDel(D_VNAM, lngE) = strModName
                                arr_varDel(D_PNAM, lngE) = strProcName
                                arr_varDel(D_LIN, lngE) = lngY
                                arr_varDel(D_RAW, lngE) = strLine
                                'Debug.Print "'" & strLine
                                'Stop
                              End If
                            End Select
                          Else
                            ' ** Shouldn't begin a line unless it's a line continuation.
                            If Right(.Lines(lngY - 1&, 1), 1) = "_" Then
                              blnFound = True
                            Else
                              lngDels = lngDels + 1&
                              lngE = lngDels - 1&
                              ReDim Preserve arr_varDel(D_ELEMS, lngE)
                              arr_varDel(D_FIND, lngE) = arr_varProc(P_PNAM, lngX)
                              arr_varDel(D_VNAM, lngE) = strModName
                              arr_varDel(D_PNAM, lngE) = strProcName
                              arr_varDel(D_LIN, lngE) = lngY
                              arr_varDel(D_RAW, lngE) = strLine
                              'Debug.Print "'" & strLine
                              'Stop
                            End If
                          End If
                          If blnFound = True Then  ' ** It passed first test.
                            blnFound = False
                            If intPos01 + intLen < Len(strLine) Then
                              strTmp01 = Mid(strLine, (intPos01 + intLen), 1)  ' ** Character after proc name.
                              lngTmp02 = Asc(strTmp01)
                              Select Case strTmp01
                              Case " ", ",", "(", ")"
                                ' ** OK.
                                blnFound = True
                              Case Chr(34)  ' ** Quotes.
                                ' ** No.
                              Case "_"
                                ' ** No.
                              Case ";", ":", "[", "]", "{", "}", "\", "|"
                                ' ** Unsure.
                                lngDels = lngDels + 1&
                                lngE = lngDels - 1&
                                ReDim Preserve arr_varDel(D_ELEMS, lngE)
                                arr_varDel(D_FIND, lngE) = arr_varProc(P_PNAM, lngX)
                                arr_varDel(D_VNAM, lngE) = strModName
                                arr_varDel(D_PNAM, lngE) = strProcName
                                arr_varDel(D_LIN, lngE) = lngY
                                arr_varDel(D_RAW, lngE) = strLine
                                'Debug.Print "'" & strLine
                                'Stop
                              Case Else
                                If lngTmp02 >= 65 And lngTmp02 <= 90 Then  ' ** A - Z.
                                  ' ** Nope.
                                ElseIf lngTmp02 >= 97 And lngTmp02 <= 122 Then  ' ** a - z.
                                  ' ** Nope.
                                ElseIf lngTmp02 >= 48 And lngTmp02 <= 57 Then  ' ** 0 - 9.
                                  ' ** Nope.
                                Else
                                  ' ** Unsure.
                                  lngDels = lngDels + 1&
                                  lngE = lngDels - 1&
                                  ReDim Preserve arr_varDel(D_ELEMS, lngE)
                                  arr_varDel(D_FIND, lngE) = arr_varProc(P_PNAM, lngX)
                                  arr_varDel(D_VNAM, lngE) = strModName
                                  arr_varDel(D_PNAM, lngE) = strProcName
                                  arr_varDel(D_LIN, lngE) = lngY
                                  arr_varDel(D_RAW, lngE) = strLine
                                  'Debug.Print "'" & strLine
                                  'Stop
                                End If
                              End Select
                            Else
                              ' ** Last word on the line.
                              blnFound = True
                            End If
                            If blnFound = True Then
                              If strProcName = vbNullString Then strProcName = "Declaration"
                              lngUses = lngUses + 1&
                              lngE = lngUses - 1&
                              ReDim Preserve arr_varUse(U_ELEMS, lngE)
                              arr_varUse(U_VID, lngE) = arr_varProc(P_VID, lngX)
                              arr_varUse(U_VNAM, lngE) = arr_varProc(P_VNAM, lngX)
                              arr_varUse(U_PID, lngE) = arr_varProc(P_PID, lngX)
                              arr_varUse(U_PNAM, lngE) = arr_varProc(P_PNAM, lngX)
                              arr_varUse(U_VIDU, lngE) = CLng(0)
                              arr_varUse(U_VNAMU, lngE) = strModName
                              arr_varUse(U_PIDU, lngE) = CLng(0)
                              arr_varUse(U_PNAMU, lngE) = strProcName
                              arr_varUse(U_LIN, lngE) = lngY
                              intPos03 = InStr(strLine, " ")
                              strTmp01 = Trim(Left(strLine, intPos03))
                              If IsNumeric(strTmp01) = True Then
                                arr_varUse(U_CODE, lngE) = strTmp01
                              Else
                                arr_varUse(U_CODE, lngE) = Null
                              End If
                              arr_varUse(U_RAW, lngE) = strLine
                            End If
                          End If
                        End If
                      End If  ' ** intPos01.
                    End If  ' ** Remark.
                  End If  ' ** vbNullString.
                Next  ' ** lngY.
              End With  ' ** cod.
              Set cod = Nothing
            End If
          End With  ' ** vbc.
          Set vbc = Nothing
        Next  ' ** vbc.
      Next  ' ** lngX
    End With
    Set vbp = Nothing

  End If  ' ** lngProcs.

  Debug.Print "'USES: " & CStr(lngUses)
  DoEvents

  Debug.Print "'UNSURES: " & CStr(lngDels)
  DoEvents

  If lngUses > 0& Then
    Set dbs = CurrentDb
    With dbs

      lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

      Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngUses - 1&)
          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & arr_varUse(U_VNAMU, lngX) & "'"
          If .NoMatch = False Then
            arr_varUse(U_VIDU, lngX) = ![vbcom_id]
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
      With rst
        .MoveFirst
        For lngX = 0& To (lngUses - 1&)
          If IsNull(arr_varUse(U_VIDU, lngX)) = False Then
            .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varUse(U_VIDU, lngX)) & " And " & _
              "[vbcomproc_name] = '" & arr_varUse(U_PNAMU, lngX) & "'"
            If .NoMatch = False Then
              arr_varUse(U_PIDU, lngX) = ![vbcomproc_id]
            End If
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      Set rst = .OpenRecordset("zz_tbl_VBComponent_Usage", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngUses - 1&)
          .AddNew
          ' ** ![vbcomuse_id] : AutoNumber.
          ![dbs_id] = lngThisDbsID
          ![vbcom_id] = arr_varUse(U_VID, lngX)
          ![vbcom_name] = arr_varUse(U_VNAM, lngX)
          ![vbcomproc_id] = arr_varUse(U_PID, lngX)
          ![vbcomproc_name] = arr_varUse(U_PNAM, lngX)
          ![vbcom_id_use] = arr_varUse(U_VIDU, lngX)
          ![vbcom_name_use] = arr_varUse(U_VNAMU, lngX)
          ![vbcomproc_id_use] = arr_varUse(U_PIDU, lngX)
          ![vbcomproc_name_use] = arr_varUse(U_PNAMU, lngX)
          ![vbcomuse_line] = arr_varUse(U_LIN, lngX)
          If IsNull(arr_varUse(U_CODE, lngX)) = False Then
            ![vbcomuse_code] = arr_varUse(U_CODE, lngX)
          Else
            ![vbcomuse_code] = Null
          End If
          ![vbcomuse_raw] = arr_varUse(U_RAW, lngX)
          ![vbcomuse_datemodified] = Now()
          .Update
        Next  ' ** lngX.
        .Close
      End With  ' ** rst.
      Set rst = Nothing

      If lngDels > 0& Then
        Set rst = .OpenRecordset("zz_tbl_VBComponent_Unknown", dbOpenDynaset, dbConsistent)
        With rst
          For lngX = 0& To (lngDels - 1&)
            .AddNew
            ' ** ![vbcomunk_id] : AutoNumber.
            ![dbs_id] = lngThisDbsID
            ![vbcomunk_find] = arr_varDel(D_FIND, lngX)
            ![vbcom_name] = arr_varDel(D_VNAM, lngX)
            ![vbcomproc_name] = arr_varDel(D_PNAM, lngX)
            ![vbcomunk_line] = arr_varDel(D_LIN, lngX)
            ![vbcomunk_raw] = arr_varDel(D_RAW, lngX)
            ![vbcomunk_datemodified] = Now()
            .Update
          Next  ' ** lngX.
          .Close
        End With
        Set rst = Nothing
      End If  ' ** lngDels.

      .Close
    End With  ' ** dbs
    Set dbs = Nothing
  End If  ' ** lngUses.

'ALL 'Proper' ARE NOT!

'PROCS: 643
'USES: 15295
'UNSURES: 0
'DONE!
  Beep

  Debug.Print "'DONE!"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  ProcUsage_Doc = blnRetVal

End Function
