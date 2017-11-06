Attribute VB_Name = "zz_mod_ModuleMiscFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_ModuleMiscFuncs"

'VGC 08/18/2013: CHANGES!

' ** THIS HAS VBA_GetCode()!

'Security Event Privileges?
'SepAcquireTokenLockExclusive
'SepAcquireTokenLockShared
'SepCaptureAcl
'SepCaptureSecurityQualityOfService
'SepCaptureSid
'SepCreateImpersonationTokenDacl
'SepCreateSystemProcessToken
'SepDuplicateToken
'SepInitDACLs
'SepInitializeTokenImplementation
'SepInitPrivileges
'SepInitSDs
'SepInitSecurityIDs
'SepPrivilegeCheck
'SepReleaseAcl
'SepReleaseSecurityQualityOfService
'SepReleaseSid
'SepReleaseTokenLock
'SepTokenObjectType

'NO FAMILY:
'glngWarnRecs
'glngWarnSize
'gstrNoPermission
'gstrRegKeyName
'gblnDev_Debug
'gblnNoErrHandle_Repost
'gintShareFaceDecimals
'gstrDevUserName
'RET_ERR

' ** FORMS WITH FILE-FOLDER GRAPHIC LOOK STANDARDS:
' **
' **   WHITE ON BLUE, 9 PT. Fixedsys FONT:
' frmAccountExport
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmBackupRestore_File
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmAssetPricing_Import
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmAccountProfile_ReviewFreq
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmAccountProfile_StatementFreq
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_AccountProfile
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_AccountReviews
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_ArchivedTransactions
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_AssetHistory
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_CapitalGainAndLoss
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_Holdings
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_IncomeExpense
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_Locations
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_StatementOfCondition
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_TaxIncomeDeductions
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_TaxLot
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_TransactionsByType
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmRpt_UnrealizedGainAndLoss
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' frmStatementParameters
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
'   7 IN, 8 ABOVE   (blue)         9 Pt. Label (Fixedsys)
' **
' **   CLR_VDKGRY ON BEIGE, 9 PT. Arial FONT:
' frmArchiveTransactions
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmBackupRestore
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmLicense
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmLicense_Edit
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmLinkData
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmReportList
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmRpt_Checks
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmRpt_Checks_MICR_Adjust
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmRpt_Checks_MICR_Set
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' frmXAdmin_SysInfo
'   8 IN, 10 ABOVE  (beige)        9 Pt. Label (Arial)
' **
' **   CLR_DKGRY2 ON BEIGE, 8 PT. Arial FONT:
' frmRpt_Checks
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
' frmRpt_Checks_MICR_Adjust
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
' frmSweeper
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
' frmVersion_Main
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (beige, sm)    8 Pt. Label (Arial)
' **
' **   CLR_DKGRY2 ON NEW BLUE, 8 PT. Arial FONT:
' frmRpt_CourtReports_CA_Input
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
' frmRpt_CourtReports_FL_Input
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
' frmRpt_CourtReports_NS_Input
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
'   8 IN, 8 ABOVE   (new blue, sm) 8 Pt. Label (Arial)
' **
' **   CLR_DKGRY2 ON BEIGE, 11 PT. Arial FONT:
' frmVersion_Input
'   10 IN, 11 ABOVE (beige)       11 Pt. Label (Arial)
'   10 IN, 11 ABOVE (beige)       11 Pt. Label (Arial)

' ** TaKeyDownType enumeration:
Public Const taKeyDown_Plain        As Long = 0&
Public Const taKeyDown_Ctrl         As Long = 1&
Public Const taKeyDown_Alt          As Long = 2&
Public Const taKeyDown_Shift        As Long = 3&
Public Const taKeyDown_CtrlAlt      As Long = 4&
Public Const taKeyDown_CtrlShift    As Long = 5&
Public Const taKeyDown_AltShift     As Long = 6&
Public Const taKeyDown_CtrlAltShift As Long = 7&
Public Const taKeyDown_Unknown      As Long = -1&

Private blnRetValx As Boolean
' **

Public Function VBA_PubProcParams() As Boolean
' ** Document parameters fed to a public procedure.
'NO! ALREADY DONE IN VBA_Component_Proc_Doc()
' TO tblVBComponent_Procedure_Parameter.

  Const THIS_PROC As String = "VBA_PubProcParams"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngThisDbsID As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  DoCmd.Hourglass True
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs



    .Close
  End With




  DoCmd.Hourglass False

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Beep

  VBA_PubProcParams = blnRetVal

End Function

Public Function VBA_PubProcUsage() As Boolean

  Const THIS_PROC As String = "VBA_PubProcUsage"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngLines As Long, lngDecLines As Long
  Dim lngHits As Long
  Dim strModName As String, strLine As String
  Dim lngThisDbsID As Long
  Dim blnAddAll As Boolean, blnAdd As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intLen As Integer
  Dim strTmp01 As String, lngTmp02 As Long
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varProc().
  Const P_DID As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_VID  As Integer = 2
  Const P_VNAM As Integer = 3
  Const P_PID  As Integer = 4
  Const P_PNAM As Integer = 5
  Const P_TYP  As Integer = 6
  Const P_SCOP As Integer = 7

  ' ** Array: arr_varHit()
  Const H_ELEMS As Integer = 8  ' ** Array's first-element UBound().
  Const H_DID   As Integer = 0
  Const H_VID   As Integer = 1
  Const H_PID   As Integer = 2
  Const H_HVID  As Integer = 3
  Const H_HVNAM As Integer = 4
  Const H_HPID  As Integer = 5
  Const H_HPNAM As Integer = 6
  Const H_LIN   As Integer = 7
  Const H_RAW   As Integer = 8

  blnRetVal = True

  DoCmd.Hourglass True
  DoEvents

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs
    ' ** zz_qry_VBComponent_Proc_18a (tblVBComponent_Procedure, just Public,
    ' ** standard and class modules), just non-zz_.. Public procs.
    Set qdf = .QueryDefs("zz_qry_VBComponent_Proc_18c")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngProcs = .RecordCount
      .MoveFirst
      arr_varProc = .GetRows(lngProcs)
      ' ****************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   Fields  Element  Name              Constant
      ' **   ======  =======  ================  ==========
      ' **      1       0     dbs_id            P_DID
      ' **      2       1     dbs_name          P_DNAM
      ' **      3       2     vbcom_id          P_VID
      ' **      4       3     vbcom_name        P_VNAM
      ' **      5       4     vbcomproc_id      P_PID
      ' **      6       5     vbcomproc_name    P_PNAM
      ' **      7       6     comtype_type      P_TYP
      ' **      8       7     scopetype_type    P_SCOP
      ' **
      ' ****************************************************
      .Close
    End With
    .Close
  End With  ' ** dbs.

  Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.
  DoEvents

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
          For lngX = 0 To (lngProcs - 1&)
            If strModName <> arr_varProc(P_VNAM, lngX) Then  ' ** Don't search its own module.
              If arr_varProc(P_TYP, lngX) = vbext_ct_StdModule Then
                For lngY = lngDecLines To lngLines
                  strLine = .Lines(lngY, 1)
                  strLine = Trim$(strLine)
                  If strLine <> vbNullString Then
                    If Left$(strLine, 1) <> "'" Then
                      intPos1 = InStr(strLine, arr_varProc(P_PNAM, lngX))
                      If intPos1 > 0 Then
                        intLen = Len(arr_varProc(P_PNAM, lngX))
                        If intPos1 = 1 Then
                          If Len(strLine) > intLen Then
                            strTmp01 = Mid$(strLine, (intLen + 1), 1)
                            Select Case strTmp01
                            Case " ", "(", ")", ","  ' ** Space, parens, comma.
                              ' ** Looks good.
                              lngHits = lngHits + 1&
                              lngE = lngHits - 1&
                              ReDim Preserve arr_varHit(H_ELEMS, lngE)
                              arr_varHit(H_DID, lngE) = arr_varProc(P_DID, lngX)
                              arr_varHit(H_VID, lngE) = arr_varProc(P_VID, lngX)
                              arr_varHit(H_PID, lngE) = arr_varProc(P_PID, lngX)
                              arr_varHit(H_HVID, lngE) = CLng(0)
                              arr_varHit(H_HVNAM, lngE) = strModName
                              arr_varHit(H_HPID, lngE) = CLng(0)
                              arr_varHit(H_HPNAM, lngE) = .ProcOfLine(lngY, vbext_pk_Proc)
                              arr_varHit(H_LIN, lngE) = lngY
                              arr_varHit(H_RAW, lngE) = strLine
                            Case "_", "[", "]", "/", "'", "¹", Chr(34)
                              ' ** Nope, something else. (A line-continuation requires a space.)
                            Case Else
                              If (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Or _
                                  (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Then
                                ' ** Nope, something else.
                              ElseIf strTmp01 = "-" And arr_varProc(P_PNAM, lngX) = "scr" Then
                                ' ** Nope.
                              Else
                                ' ** What else?
                                Debug.Print "'1  " & strTmp01 & "  '" & arr_varProc(P_PNAM, lngX) & "'   '" & strLine & "'"
                                DoEvents
'3  -  'Scr'   '5260            MsgBox "The On-Screen Notice cannot be a negative number.", vbInformation + vbOKOnly, "Invalid Entry"'
                              End If
                            End Select
                          Else
                            lngHits = lngHits + 1&
                            lngE = lngHits - 1&
                            ReDim Preserve arr_varHit(H_ELEMS, lngE)
                            arr_varHit(H_DID, lngE) = arr_varProc(P_DID, lngX)
                            arr_varHit(H_VID, lngE) = arr_varProc(P_VID, lngX)
                            arr_varHit(H_PID, lngE) = arr_varProc(P_PID, lngX)
                            arr_varHit(H_HVID, lngE) = CLng(0)
                            arr_varHit(H_HVNAM, lngE) = strModName
                            arr_varHit(H_HPID, lngE) = CLng(0)
                            arr_varHit(H_HPNAM, lngE) = .ProcOfLine(lngY, vbext_pk_Proc)
                            arr_varHit(H_LIN, lngE) = lngY
                            arr_varHit(H_RAW, lngE) = strLine
                          End If
                        Else
                          strTmp01 = Mid$(strLine, (intPos1 - 1), 1)
                          Select Case strTmp01
                          Case " ", "(", ")", ","  ' ** Space, parens, comma.
                            ' ** Looks good.
                            intPos2 = intPos1 + intLen
                            If intPos2 < Len(strLine) Then
                              strTmp01 = Mid$(strLine, intPos2, 1)
                              Select Case strTmp01
                              Case " ", "(", ")", ","  ' ** Space, parens, comma.
                                ' ** Looks good.
                                lngHits = lngHits + 1&
                                lngE = lngHits - 1&
                                ReDim Preserve arr_varHit(H_ELEMS, lngE)
                                arr_varHit(H_DID, lngE) = arr_varProc(P_DID, lngX)
                                arr_varHit(H_VID, lngE) = arr_varProc(P_VID, lngX)
                                arr_varHit(H_PID, lngE) = arr_varProc(P_PID, lngX)
                                arr_varHit(H_HVID, lngE) = CLng(0)
                                arr_varHit(H_HVNAM, lngE) = strModName
                                arr_varHit(H_HPID, lngE) = CLng(0)
                                arr_varHit(H_HPNAM, lngE) = .ProcOfLine(lngY, vbext_pk_Proc)
                                arr_varHit(H_LIN, lngE) = lngY
                                arr_varHit(H_RAW, lngE) = strLine
                              Case "_", "[", "]", "/", "'", "¹", Chr(34)
                                ' ** Nope, something else. (A line-continuation requires a space.)
                              Case Else
                                If (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Or _
                                    (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Then
                                  ' ** Nope, something else.
                                ElseIf strTmp01 = "-" And arr_varProc(P_PNAM, lngX) = "scr" Then
                                  ' ** Nope.
                                Else
                                  ' ** What else?
                                  Debug.Print "'2  " & strTmp01 & "  '" & arr_varProc(P_PNAM, lngX) & "'   '" & strLine & "'"
                                  DoEvents
                                End If
                              End Select
                            Else
                              lngHits = lngHits + 1&
                              lngE = lngHits - 1&
                              ReDim Preserve arr_varHit(H_ELEMS, lngE)
                              arr_varHit(H_DID, lngE) = arr_varProc(P_DID, lngX)
                              arr_varHit(H_VID, lngE) = arr_varProc(P_VID, lngX)
                              arr_varHit(H_PID, lngE) = arr_varProc(P_PID, lngX)
                              arr_varHit(H_HVID, lngE) = CLng(0)
                              arr_varHit(H_HVNAM, lngE) = strModName
                              arr_varHit(H_HPID, lngE) = CLng(0)
                              arr_varHit(H_HPNAM, lngE) = .ProcOfLine(lngY, vbext_pk_Proc)
                              arr_varHit(H_LIN, lngE) = lngY
                              arr_varHit(H_RAW, lngE) = strLine
                            End If
                          Case "_", "[", "]", "/", "'", "¹", Chr(34), "."
                            ' ** Nope, something else.
                          Case Else
                            If (Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90) Or (Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122) Or _
                                (Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57) Then
                              ' ** Nope, something else.
                            ElseIf strTmp01 = "-" And arr_varProc(P_PNAM, lngX) = "scr" Then
                              ' ** Nope.
                            Else
                              ' ** What else?
                              Debug.Print "'3  " & strTmp01 & "  '" & arr_varProc(P_PNAM, lngX) & "'   '" & strLine & "'"
                              DoEvents
                            End If
                          End Select
                        End If
                      End If
                    End If
                  End If
                Next  ' ** lngY.
              ElseIf arr_varProc(P_TYP, lngX) = vbext_ct_ClassModule Then
                ' ** Because these sometimes have names the same as properties, handle them separately.

              Else
                ' ** Shouldn't be others!
                Debug.Print "'4  '" & arr_varProc(P_PNAM, lngX) & "'"
              End If  ' ** P_TYP.
            End If  ' ** strModName.
          Next  ' ** lngX.
        End With  ' ** cod.
      End With  ' ** vbc.
    Next  ' ** vbc.
  End With  ' ** vbp.
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  Set dbs = CurrentDb
  With dbs

    Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
    With rst
      For lngX = 0& To (lngHits - 1&)
        If arr_varHit(H_HVID, lngX) = 0& Then
          .FindFirst "[dbs_id] = " & arr_varHit(H_DID, lngX) & " And " & _
            "[vbcom_name] = '" & arr_varHit(H_HVNAM, lngX) & "'"
          Select Case .NoMatch
          Case True
            Stop
          Case False
            lngTmp02 = ![vbcom_id]
            arr_varHit(H_HVID, lngX) = lngTmp02
            For lngY = (lngX + 1&) To (lngHits - 1&)
              If arr_varHit(H_HVNAM, lngY) = arr_varHit(H_HVNAM, lngX) Then
                arr_varHit(H_HVID, lngY) = lngTmp02
              End If
            Next  ' ** lngY.
          End Select
        End If
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    Set rst = Nothing

    Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
    With rst
      For lngX = 0& To (lngHits - 1&)
        If arr_varHit(H_HPID, lngX) = 0& Then
          .FindFirst "[dbs_id] = " & arr_varHit(H_DID, lngX) & " And " & _
            "[vbcom_id] = " & CStr(arr_varHit(H_HVID, lngX)) & "And " & _
            "[vbcomproc_name] = '" & arr_varHit(H_HPNAM, lngX) & "'"
          Select Case .NoMatch
          Case True
            Stop
          Case False
            lngTmp02 = ![vbcomproc_id]
            arr_varHit(H_HPID, lngX) = lngTmp02
            For lngY = (lngX + 1&) To (lngHits - 1&)
              If arr_varHit(H_HVID, lngY) = arr_varHit(H_HVID, lngX) And _
                  arr_varHit(H_HPNAM, lngY) = arr_varHit(H_HPNAM, lngX) Then
                arr_varHit(H_HPID, lngY) = lngTmp02
              End If
            Next  ' ** lngY.
          End Select
        End If
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    Set rst = Nothing

    blnAddAll = False: blnAdd = False
    Set rst = .OpenRecordset("tblVBComponent_Procedure_Detail2", dbOpenDynaset, dbConsistent)
    With rst
      If .BOF = True And .EOF = True Then
        blnAddAll = True
      End If
      For lngX = 0& To (lngHits - 1&)
        If blnAddAll = True Then
          blnAdd = True
        Else
          '.FindFirst
          'Select Case .NoMatch
          'Case True
            blnAdd = True
          'Case False

          'End Select
        End If
        Select Case blnAdd
        Case True
          .AddNew
          ![dbs_id] = lngThisDbsID
          ![vbcom_id] = arr_varHit(H_VID, lngX)
          ![vbcomproc_id] = arr_varHit(H_PID, lngX)
          ' ** vbprocdet_id : AutoNumber.
          ![dbs_id_det] = lngThisDbsID
          ![vbcom_id_det] = arr_varHit(H_HVID, lngX)
          ![vbcomproc_id_det] = arr_varHit(H_HPID, lngX)
          ![vbprocdet_linenum] = arr_varHit(H_LIN, lngX)
          ![vbprocdet_proc] = arr_varHit(H_HPNAM, lngX)
          ![vbprocdet_param1] = Null
          ![vbprocdet_param2] = Null
          ![vbprocdet_param3] = Null
          ![vbprocdet_param4] = Null
          ![vbprocdet_assign] = Null
          ![vbprocdet_raw] = arr_varHit(H_RAW, lngX)
          ![vbprocdet_datemodified] = Now()
          .Update
        Case False
          '.Edit

          '.Update
        End Select
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.

    .Close
  End With  ' ** dbs.

  DoCmd.Hourglass False

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Beep

  VBA_PubProcUsage = blnRetVal

End Function

Public Function VBA_KeyDoc() As Boolean
' ** Collects shortcut key remarks from modules.

  Const THIS_PROC As String = "VBA_KeyDoc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strModName As String, strLine As String, strKeyDownType As String
  Dim lngLines As Long, lngDecLines As Long
  Dim lngRems As Long, arr_varRem() As Variant
  Dim lngRemXs As Long, arr_varRemX() As Variant
  Dim lngThisDbsID As Long, lngMods As Long
  Dim blnAdd As Boolean, blnAddAll As Boolean, blnDelete As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varRem().
  Const R_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const R_VID  As Integer = 0
  Const R_VNAM As Integer = 1
  Const R_TYP  As Integer = 2
  Const R_TNAM As Integer = 3
  Const R_KEY  As Integer = 4
  Const R_FID  As Integer = 5
  Const R_FNAM As Integer = 6
  Const R_CID  As Integer = 7
  Const R_CNAM As Integer = 8
  Const R_REM  As Integer = 9

  blnRetVal = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngRems = 0&
  ReDim arr_varRem(R_ELEMS, 0)

  lngRemXs = 0&
  ReDim arr_varRemX(R_ELEMS, 0)

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    lngMods = .VBComponents.Count
    lngY = 0&
    Debug.Print "'MODS: " & CStr(lngMods)
    'Debug.Print "'|";
    DoEvents
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        strKeyDownType = vbNullString
        If Left$(strModName, 3) <> "zz_" And strModName <> "modRegistryFuncs" Then
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngX = 1& To lngDecLines
              strLine = .Lines(lngX, 1)
              If Trim$(strLine) <> vbNullString Then
                If Left$(Trim$(strLine), 1) = "'" Then
                  strTmp01 = Trim$(strLine)
                  intPos1 = InStr(strTmp01, "keys responsive")
                  If intPos1 > 0 Then
                    strKeyDownType = Trim$(Left$(strTmp01, (intPos1 - 1)))
                    intPos1 = InStr(strKeyDownType, "Shortcut")
                    If intPos1 > 0 Then
                      strKeyDownType = Trim$(Mid$(strKeyDownType, (intPos1 + Len("Shortcut"))))
                    End If
                  ElseIf Left$(strTmp01, 7) = "' **   " Then
                    intPos1 = InStr(strTmp01, ":")
                    intPos2 = InStr(strTmp01, "{")
                    intPos3 = InStr(strTmp01, "}")
                    If intPos1 > 0 And intPos2 > 0 And intPos3 > 0 Then
If strTmp01 <> "' **   Name: ctlCurrentRecord  ControlSource: {unbound}" Then
                      lngRems = lngRems + 1&
                      lngE = lngRems - 1&
                      ReDim Preserve arr_varRem(R_ELEMS, lngE)
                      arr_varRem(R_VID, lngE) = CLng(0)
                      arr_varRem(R_VNAM, lngE) = strModName
                      arr_varRem(R_TYP, lngE) = taKeyDown_Unknown
                      If strKeyDownType = "F-" Then strKeyDownType = "Plain"
                      arr_varRem(R_TNAM, lngE) = strKeyDownType
                      arr_varRem(R_KEY, lngE) = Null
                      arr_varRem(R_FID, lngE) = Null   'frm_id
                      arr_varRem(R_FNAM, lngE) = Null  'frm_name
                      arr_varRem(R_CID, lngE) = Null   'ctl_id
                      arr_varRem(R_CNAM, lngE) = Null  'ctl_name
                      arr_varRem(R_REM, lngE) = strTmp01
End If
                    Else
                      If intPos1 = 0 And intPos2 > 0 And intPos3 > 0 Then  ' ** Maybe I just missed the colon.
                        If strTmp01 <> "' **     {Nothing}" Then
                          Debug.Print "'" & strModName & "  " & strTmp01
                        End If
                      End If
                    End If
                  End If
                End If  ' ** Remark.
              End If  ' ** vbNullString.
            Next  ' ** lngX.
          End With  ' ** cod.
          Set cod = Nothing
        End If  ' ** zz_.
      End With  ' ** vbc.
      Set vbc = Nothing
      lngY = lngY + 1&
      If lngY Mod 100& = 0& Then
        'Debug.Print "|"
        'Debug.Print "'|";
      ElseIf lngY Mod 10& = 0& Then
        'Debug.Print "|";
      Else
        'Debug.Print ".";
      End If
      DoEvents
    Next  ' ** vbc.
    'Debug.Print
  End With  ' ** vbp.
  Set vbp = Nothing

  If lngRems > 0& Then

    Debug.Print "'REMS: " & CStr(lngRems)
    DoEvents

    Set dbs = CurrentDb
    With dbs

      Set rst1 = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
      Set rst2 = .OpenRecordset("tblKeyDownType", dbOpenDynaset, dbReadOnly)
      With rst1
        .MoveLast
        .MoveFirst
        For lngX = 0& To (lngRems - 1&)
          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & arr_varRem(R_VNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varRem(R_VID, lngX) = ![vbcom_id]
            If IsNull(![vbcom_name2]) = False Then
              arr_varRem(R_FNAM, lngX) = ![vbcom_name2]
            End If
          Else
Debug.Print "'NOT FOUND!  " & arr_varRem(R_VNAM, lngX)
DoEvents
Stop
          End If
          With rst2
            .FindFirst "[keydowntype_name] = '" & arr_varRem(R_TNAM, lngX) & "'"
            If .NoMatch = False Then
              arr_varRem(R_TYP, lngX) = ![keydowntype_type]
            End If
          End With
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      rst2.Close
      Set rst1 = Nothing
      Set rst2 = Nothing

      ' ** Get the frm_id.
      Set rst1 = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
      With rst1
        .MoveLast
        .MoveFirst
        For lngX = 0& To (lngRems - 1&)
          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_name] = '" & arr_varRem(R_FNAM, lngX) & "'"
          If .NoMatch = False Then
            arr_varRem(R_FID, lngE) = ![frm_id]
          Else
Debug.Print "'NOT FOUND!  " & arr_varRem(R_FNAM, lngX)
DoEvents
Stop
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      ' ** Get the shortcut key, and ctl_name.
      For lngX = 0& To (lngRems - 1&)
        intPos1 = InStr(arr_varRem(R_REM, lngX), ":")
        If intPos1 > 0 Then
          strTmp01 = Trim(Mid(arr_varRem(R_REM, lngX), (intPos1 + 1)))
          intPos1 = InStr(strTmp01, "{")
          If intPos1 > 1 Then
            strTmp02 = Trim$(Left$(strTmp01, (intPos1 - 1)))
            arr_varRem(R_KEY, lngX) = strTmp02
          ElseIf intPos1 = 1 Then
            strTmp02 = vbNullString
            arr_varRem(R_KEY, lngX) = Null
          Else
            Beep
            Stop
          End If
          strTmp01 = Mid$(strTmp01, intPos1)
          intPos1 = InStr(strTmp01, "}")
          If intPos1 < Len(strTmp01) Then strTmp01 = Left$(strTmp01, intPos1)
          strTmp01 = Left$(Mid$(strTmp01, 2), (Len(Mid$(strTmp01, 2)) - 1))
          arr_varRem(R_CNAM, lngX) = strTmp01
        End If  ' ** intPos1.
      Next  ' ** lngX.

      ' ** Look for multiple-form shortcuts.
      For lngX = 0& To (lngRems - 1&)
        intPos1 = InStr(arr_varRem(R_CNAM, lngX), ",")
        If intPos1 > 0 Then
          strTmp01 = Trim(Left(arr_varRem(R_CNAM, lngX), (intPos1 - 1)))
          strTmp02 = Trim(Mid(arr_varRem(R_CNAM, lngX), (intPos1 + 1)))
          strTmp03 = vbNullString
          ' ** REMEMBER, the commas separate form names, not control names!
          intPos2 = InStr(strTmp01, " on frm")
          If intPos2 > 0 Then
            strTmp03 = Trim$(Left$(strTmp01, intPos2))
          End If
          strTmp02 = strTmp03 & " on " & strTmp02
          arr_varRem(R_CNAM, lngX) = strTmp01
          intPos3 = 0
          Do While intPos1 > 0
            intPos1 = InStr(strTmp02, ",")
            intPos3 = intPos3 + 1
            lngRemXs = lngRemXs + 1&
            lngE = lngRemXs - 1&
            ReDim Preserve arr_varRemX(R_ELEMS, lngE)
            arr_varRemX(R_VID, lngE) = arr_varRem(R_VID, lngX)
            arr_varRemX(R_VNAM, lngE) = arr_varRem(R_VNAM, lngX)
            arr_varRemX(R_TYP, lngE) = arr_varRem(R_TYP, lngX)
            arr_varRemX(R_TNAM, lngE) = arr_varRem(R_TNAM, lngX)
            arr_varRemX(R_KEY, lngE) = arr_varRem(R_KEY, lngX)
            arr_varRemX(R_FID, lngE) = arr_varRem(R_FID, lngX)
            arr_varRemX(R_FNAM, lngE) = arr_varRem(R_FNAM, lngX)
            arr_varRemX(R_CID, lngE) = Null
            If intPos1 > 0 Then
              strTmp01 = Trim$(Left$(strTmp02, (intPos1 - 1)))
              strTmp02 = Trim$(Mid$(strTmp02, (intPos1 + 1)))
              strTmp03 = vbNullString
              ' ** REMEMBER, the commas separate form names, not control names!
              intPos2 = InStr(strTmp01, " on frm")
              If intPos2 > 0 Then
                strTmp03 = Trim$(Left$(strTmp01, intPos2))
              End If
              strTmp02 = strTmp03 & " on " & strTmp02
            End If
            arr_varRemX(R_CNAM, lngE) = strTmp02
            arr_varRemX(R_REM, lngE) = arr_varRem(R_REM, lngX)
            If intPos3 > 100 Then
              Beep
              Stop
              Exit For
            End If
          Loop
        End If
      Next  ' ** lngX.

      ' ** Add the multiple-form shortcuts to the main array.
      If lngRemXs > 0& Then
        For lngX = 0& To (lngRemXs - 1&)
          lngRems = lngRems + 1&
          lngE = lngRems - 1&
          ReDim Preserve arr_varRem(R_ELEMS, lngE)
          For lngY = 0& To R_ELEMS
            arr_varRem(lngY, lngE) = arr_varRemX(lngY, lngX)
          Next  ' ** lngY.
        Next  ' ** lngX
      End If  ' ** lngRemXs.

      lngRemXs = 0&
      ReDim arr_varRemX(R_ELEMS, 0)

      ' ** Reassign remotely-listed shortcuts.
      For lngX = 0& To (lngRems - 1&)
        intPos1 = InStr(arr_varRem(R_CNAM, lngX), " on frm")
        If intPos1 > 0 Then
          strTmp01 = Trim(Mid(arr_varRem(R_CNAM, lngX), intPos1))
          strTmp02 = Trim(Left(arr_varRem(R_CNAM, lngX), intPos1))
          If Left$(strTmp01, 3) = "on " Then strTmp01 = Mid$(strTmp01, 4)
          arr_varRem(R_CNAM, lngX) = strTmp02
          intPos1 = InStr(strTmp01, ",")
          If intPos1 > 0 Then
            strTmp03 = Trim$(Left$(strTmp01, (intPos1 - 1)))
            If arr_varRem(R_FNAM, lngX) <> strTmp03 Then
              arr_varRem(R_FNAM, lngX) = strTmp03
            End If
            varTmp00 = DLookup("[frm_id]", "tblForm", "[frm_name] = '" & strTmp03 & "'")
            If IsNull(varTmp00) = False Then
              arr_varRem(R_FID, lngX) = varTmp00
            Else
              Beep
              Stop
            End If
          Else
            If arr_varRem(R_FNAM, lngX) <> strTmp01 Then
              arr_varRem(R_FNAM, lngX) = strTmp01
            End If
            varTmp00 = DLookup("[frm_id]", "tblForm", "[frm_name] = '" & strTmp01 & "'")
            If IsNull(varTmp00) = False Then
              arr_varRem(R_FID, lngX) = varTmp00
            Else
              Beep
              Stop
            End If
          End If
        End If
      Next  ' ** lngX

      ' ** Check for any missing frm_id's.
      For lngX = 0& To (lngRems - 1&)
        If IsNull(arr_varRem(R_FID, lngX)) = True Then
          varTmp00 = DLookup("[frm_id]", "tblForm", "[frm_name] = '" & arr_varRem(R_FNAM, lngX) & "'")
          If IsNull(varTmp00) = False Then
            arr_varRem(R_FID, lngX) = varTmp00
          Else
            Beep
            Stop
          End If
        End If
      Next

      ' ** Get the ctl_id.
      Set rst1 = .OpenRecordset("tblForm_Control", dbOpenDynaset, dbReadOnly)
      With rst1
        .MoveLast
        .MoveFirst
        For lngX = 0& To (lngRems - 1&)
          If IsNull(arr_varRem(R_CNAM, lngX)) = False Then
            intPos1 = InStr(arr_varRem(R_CNAM, lngX), " on each subform")
            If intPos1 > 0 Then arr_varRem(R_CNAM, lngX) = Trim(Left(arr_varRem(R_CNAM, lngX), intPos1))
            If Left(Trim(arr_varRem(R_CNAM, lngX)), 3) = "on " Then arr_varRem(R_CNAM, lngX) = Mid(Trim(arr_varRem(R_CNAM, lngX)), 4)
            .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_id] = " & CStr(arr_varRem(R_FID, lngX)) & " And " & _
              "[ctl_name] = '" & arr_varRem(R_CNAM, lngX) & "'"
            If .NoMatch = False Then
              arr_varRem(R_CID, lngX) = ![ctl_id]
            Else
              ' ** There are many that aren't controls.
            End If
          End If
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      ' ** Update tblVBComponent_Shortcut, for vbcomsc_mark = True.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_50")
      qdf.Execute
      Set qdf = Nothing

      Set rst1 = .OpenRecordset("tblVBComponent_Shortcut", dbOpenDynaset, dbConsistent)
      With rst1
        blnAddAll = False
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        Else
          .MoveLast
          .MoveFirst
        End If
        For lngX = 0& To (lngRems - 1&)
          blnAdd = False
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            Select Case IsNull(arr_varRem(R_KEY, lngX))
            Case True
              .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varRem(R_VID, lngX)) & " And " & _
                "[keydowntype_type] = " & CStr(arr_varRem(R_TYP, lngX)) & " And [vbcomsc_key] Is Null And " & _
                "[frm_id] = " & CStr(arr_varRem(R_FID, lngX)) & " And [ctl_id] = " & CStr(arr_varRem(R_CID, lngX))
            Case False
              Select Case IsNull(arr_varRem(R_CID, lngX))
              Case True
                .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varRem(R_VID, lngX)) & " And " & _
                  "[keydowntype_type] = " & CStr(arr_varRem(R_TYP, lngX)) & " And [vbcomsc_key] = '" & arr_varRem(R_KEY, lngX) & "' And " & _
                  "[frm_id] = " & CStr(arr_varRem(R_FID, lngX)) & " And [ctl_id] Is Null"
              Case False
                .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varRem(R_VID, lngX)) & " And " & _
                  "[keydowntype_type] = " & CStr(arr_varRem(R_TYP, lngX)) & " And [vbcomsc_key] = '" & arr_varRem(R_KEY, lngX) & "' And " & _
                  "[frm_id] = " & CStr(arr_varRem(R_FID, lngX)) & " And [ctl_id] = " & CStr(arr_varRem(R_CID, lngX))
              End Select
            End Select
            blnAdd = .NoMatch
          End Select
          Select Case blnAdd
          Case True
            .AddNew
          Case False
            .Edit
          End Select
          If blnAdd = True Then
            ![dbs_id] = lngThisDbsID
            ![vbcom_id] = arr_varRem(R_VID, lngX)
            ' ** ![vbcomsc_id] : AutoNumber.
            ![keydowntype_type] = arr_varRem(R_TYP, lngX)
            ![vbcomsc_key] = arr_varRem(R_KEY, lngX)
            ![frm_id] = arr_varRem(R_FID, lngX)
          End If
          If Left(Trim(![vbcomsc_key]), 3) = "on " Then
            ![vbcomsc_key] = Mid(![vbcomsc_key], 4)
          End If
          ![vbcom_name] = arr_varRem(R_VNAM, lngX)
          ![frm_name] = arr_varRem(R_FNAM, lngX)
          If IsNull(arr_varRem(R_CID, lngX)) = False Then
            ![ctl_id] = arr_varRem(R_CID, lngX)
          Else
            If IsNull(![ctl_id]) = False Then
              ![ctl_id] = Null
            End If
          End If
          ![ctl_name] = arr_varRem(R_CNAM, lngX)
          If arr_varRem(R_TNAM, lngX) = vbNullString Then
            ![keydowntype_name] = "{unk}"
          Else
            ![keydowntype_name] = arr_varRem(R_TNAM, lngX)
          End If
          ![vbcomsc_remark] = arr_varRem(R_REM, lngX)
          ![vbcomsc_mark] = False  ' ** False means it's been found.
          ![vbcomsc_datemodified] = Now()
On Error Resume Next
          .Update
On Error GoTo 0
        Next  ' ** lngX.
        .Close
      End With  ' ** rst1.
      Set rst1 = Nothing

      varTmp00 = DCount("*", "tblVBComponent_Shortcut", "[vbcomsc_mark] = True")  ' ** Those not found above.
      If IsNull(varTmp00) = False Then
        If varTmp00 > 0 Then
          blnDelete = True
          Beep
          Debug.Print "'DELS: " & CStr(varTmp00)
Stop
          If blnDelete = True Then
            ' ** Delete tblVBComponent_Shortcut, for vbcomsc_mark = True.
            Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_51")
            qdf.Execute
          End If
        End If
      End If

      ' ** Update zz_qry_VBComponent_Shortcut_25x (tblVBComponent_Shortcut, with
      ' ** DLookups() to zz_qry_VBComponent_Shortcut_22 (zz_qry_VBComponent_Shortcut_21
      ' ** (tblVBComponent_Shortcut, linked to tblForm, just local shortcuts, with
      ' ** frm_id_new, frm_name_new), linked to tblForm_Control, not yet updated, with ctl_id_new)).
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_25y")
      qdf.Execute
      Set qdf = Nothing

      ' ** Update zz_qry_VBComponent_Shortcut_15c (zz_qry_VBComponent_Shortcut_15b
      ' ** (zz_qry_VBComponent_Shortcut_15a (zz_qry_VBComponent_Shortcut_06
      ' ** (tblVBComponent_Shortcut, with ctl_name_newx), just ctl_id = Null),
      ' ** without 'cmdSave', 'MoveRec'), just extra words in ctl_name, with ctl_name_new).
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15d")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15f_01_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15f_02_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15f_02_08")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15f_03_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15f_04_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15f_05_04")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15h_01_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15h_02_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15h_03_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15h_04_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15h_05_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15j_01_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15m_01_03")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_15m_02_02")
      qdf.Execute
      Set qdf = Nothing

      ' ** Empty zz_tbl_Form_Shortcut_tmp01.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_38mx")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_38m")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_39z")
      qdf.Execute
      Set qdf = Nothing

      ' ** Empty zz_tbl_Form_Shortcut_tmp02.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_39jx")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_39j")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_40l")
      qdf.Execute
      Set qdf = Nothing

      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_40n")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append zz_qry_VBComponent_Shortcut_43h_12 (xx) to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_13")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append zz_qry_VBComponent_Shortcut_43h_14 (xx) to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_15")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append zz_qry_VBComponent_Shortcut_43h_21 (xx) to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_22")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append zz_qry_VBComponent_Shortcut_43h_23 (xx) to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_24")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append .._43h_27 to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_28")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append .._43h_29 to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_30")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append .._43h_31 to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_32")
      qdf.Execute
      Set qdf = Nothing

      ' ** Append .._43h_33 to tblVBComponent_Shortcut.
      Set qdf = .QueryDefs("zz_qry_VBComponent_Shortcut_43h_34")
      qdf.Execute
      Set qdf = Nothing

      .Close
    End With  ' ** dbs.

  End If

  Beep
  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_KeyDoc = blnRetVal

End Function

Public Function VBA_KeyDown() As Boolean

  Const THIS_PROC As String = "VBA_KeyDown"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngTabs As Long, arr_varTab() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim strModName As String, strProcName As String, strLine As String
  Dim strKeyType As String, lngKeyType As Long
  Dim blnCtrl As Boolean, blnAlt As Boolean, blnShift As Boolean
  Dim lngLines As Long, lngDecLines As Long, lngRecs As Long
  Dim lngThisDbsID As Long, lngVBComID As Long, lngVBComProcID As Long
  Dim blnFound As Boolean, blnSkip As Boolean
  Dim intPos1 As Integer, intPos2 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varProc().
  Const P_DID  As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_VID  As Integer = 2
  Const P_VNAM As Integer = 3
  Const P_FID  As Integer = 4
  Const P_FNAM As Integer = 5
  Const P_PID  As Integer = 6
  Const P_PNAM As Integer = 7
  Const P_BEG  As Integer = 8
  Const P_END  As Integer = 9

  ' ** Array: arr_varTab().
  Const T_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const T_DID As Integer = 0
  Const T_VID As Integer = 1
  Const T_PID As Integer = 2
  Const T_TYP As Integer = 3
  Const T_LIN As Integer = 4
  Const T_TAB As Integer = 5
  Const T_RET As Integer = 6

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs
    ' ** zz_qry_Form_Shortcut_06 (tblVBComponent_Procedure, just .._KeyDown()
    ' ** procedures, by specified CurrentAppName()), linked to tblForm.
    Set qdf = .QueryDefs("zz_qry_Form_Shortcut_07")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngProcs = .RecordCount
      .MoveFirst
      arr_varProc = .GetRows(lngProcs)
      ' ********************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   Fields  Element  Name                  Constant
      ' **   ======  =======  ====================  ==========
      ' **      1       0     dbs_id                P_DID
      ' **      2       1     dbs_name              P_DNAM
      ' **      3       2     vbcom_id              P_VID
      ' **      4       3     vbcom_name            P_VNAM
      ' **      5       4     frm_id                P_FID
      ' **      6       5     frm_name              P_FNAM
      ' **      7       6     vbcomproc_id          P_PID
      ' **      8       7     vbcomproc_name        P_PNAM
      ' **      9       8     vbcomproc_line_beg    P_BEG
      ' **     10       9     vbcomproc_line_end    P_END
      ' **
      ' ********************************************************
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    .Close
  End With
  Set dbs = Nothing

  lngTabs = 0&
  ReDim arr_varTab(T_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        If Left$(strModName, 3) <> "zz_" Then
          Set cod = .CodeModule
          With cod
            lngLines = .CountOfLines
            lngDecLines = .CountOfDeclarationLines
            For lngX = lngDecLines To lngLines
              strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
              If strProcName = "JC_Key_Sub" Or strProcName = "SkipKey" Then
                ' ** We'll skip this Standard Module procedures for now.
              Else
                strLine = Trim$(.Lines(lngX, 1))
                If strLine <> vbNullString Then
                  If Left$(strLine, 1) <> "'" Then
                    intPos1 = InStr(strLine, "vbKeyTab")
                    intPos2 = InStr(strLine, "vbKeyReturn")
                    If intPos1 > 0 Or intPos2 > 0 Then
                      lngTabs = lngTabs + 1&
                      lngE = lngTabs - 1&
                      ReDim Preserve arr_varTab(T_ELEMS, lngE)
                      arr_varTab(T_DID, lngE) = lngThisDbsID
                      varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", _
                        "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & strModName & "'")
                      If IsNull(varTmp00) = False Then
                        lngVBComID = varTmp00
                        arr_varTab(T_VID, lngE) = lngVBComID
                      Else
                        Stop
                      End If
                      varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", _
                        "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(lngVBComID) & " And [vbcomproc_name] = '" & strProcName & "'")
                      If IsNull(varTmp00) = False Then
                        lngVBComProcID = varTmp00
                        arr_varTab(T_PID, lngE) = lngVBComProcID
                      Else
                        Stop
                      End If
                      arr_varTab(T_TYP, lngE) = taKeyDown_Unknown
                      arr_varTab(T_LIN, lngE) = lngX
                      If intPos1 > 0 Then
                        arr_varTab(T_TAB, lngE) = CBool(True)
                      Else
                        arr_varTab(T_TAB, lngE) = CBool(False)
                      End If
                      If intPos2 > 0 Then
                        arr_varTab(T_RET, lngE) = CBool(True)
                      Else
                        arr_varTab(T_RET, lngE) = CBool(False)
                      End If
                      lngTmp02 = 0&
                      For lngY = (lngX - 1&) To 1& Step -1&
                        blnCtrl = False: blnAlt = False: blnShift = False
                        strKeyType = vbNullString: lngKeyType = -1&
                        strLine = Trim$(.Lines(lngY, 1))
                        If strLine <> vbNullString Then
                          If Left$(strLine, 1) <> "'" Then
                            intPos1 = InStr(strLine, " ")
                            If intPos1 > 0 Then
                              If IsNumeric(Left$(strLine, (intPos1 - 1))) = True Then
                                strTmp01 = Trim$(Mid$(strLine, intPos1))  ' ** Strip the line number.
                                If Left$(strTmp01, 3) = "If " Then
                                  ' ** If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
                                  If Mid$(strTmp01, 4, 5) = "(Not " Or Mid$(strTmp01, 4, 11) = "intCtrlDown" Then
                                    If InStr(strTmp01, "(Not intCtrlDown)") > 0 Then
                                      blnCtrl = False
                                    Else
                                      blnCtrl = True
                                    End If
                                    If InStr(strTmp01, "(Not intAltDown)") > 0 Then
                                      blnAlt = False
                                    Else
                                      blnAlt = True
                                    End If
                                    If InStr(strTmp01, "(Not intShiftDown)") > 0 Then
                                      blnShift = False
                                    Else
                                      blnShift = True
                                    End If
                                    If blnCtrl = False And blnAlt = False And blnShift = False Then
                                      strKeyType = "Plain"
                                      lngKeyType = taKeyDown_Plain
                                    ElseIf blnCtrl = True And blnAlt = False And blnShift = False Then
                                      strKeyType = "Ctrl"
                                      lngKeyType = taKeyDown_Ctrl
                                    ElseIf blnCtrl = False And blnAlt = True And blnShift = False Then
                                      strKeyType = "Alt"
                                      lngKeyType = taKeyDown_Alt
                                    ElseIf blnCtrl = False And blnAlt = False And blnShift = True Then
                                      strKeyType = "Shift"
                                      lngKeyType = taKeyDown_Shift
                                    ElseIf blnCtrl = True And blnAlt = True And blnShift = False Then
                                      strKeyType = "Ctrl-Alt"
                                      lngKeyType = taKeyDown_CtrlAlt
                                    ElseIf blnCtrl = True And blnAlt = False And blnShift = True Then
                                      strKeyType = "Ctrl-Shift"
                                      lngKeyType = taKeyDown_CtrlShift
                                    ElseIf blnCtrl = False And blnAlt = True And blnShift = True Then
                                      strKeyType = "Alt-Shift"
                                      lngKeyType = taKeyDown_AltShift
                                    ElseIf blnCtrl = True And blnAlt = True And blnShift = True Then
                                      strKeyType = "Ctrl-Alt-Shift"
                                      lngKeyType = taKeyDown_CtrlAltShift
                                    End If
'Private Const taKeyDown_Plain        As Long = 0&
'Private Const taKeyDown_Ctrl         As Long = 1&
'Private Const taKeyDown_Alt          As Long = 2&
'Private Const taKeyDown_Shift        As Long = 3&
'Private Const taKeyDown_CtrlAlt      As Long = 4&
'Private Const taKeyDown_CtrlShift    As Long = 5&
'Private Const taKeyDown_AltShift     As Long = 6&
'Private Const taKeyDown_CtrlAltShift As Long = 7&
'Private Const taKeyDown_Unknown      As Long = -1&
                                    lngTmp02 = lngY
                                  End If
                                End If
                              End If
                            End If  ' ** intPos1.
                          End If  ' ** Remark
                        End If  ' ** vbNullString.
                        If strKeyType <> vbNullString Then
                          arr_varTab(T_TYP, lngE) = lngKeyType
                          Exit For
                        ElseIf .ProcOfLine(lngY - 1&, vbext_pk_Proc) <> strProcName Then
                          Exit For
                        End If
                      Next  ' ** lngY.
                      If lngTmp02 = 0& Then
                        ' ** Type wasn't found, most likely a Standard Module.
                        For lngY = (lngX - 1&) To 1& Step -1&
                          strLine = Trim$(.Lines(lngY, 1))
                          If strLine <> vbNullString Then
                            If Left$(strLine, 1) <> "'" Then
                              intPos1 = InStr(strLine, "intAux")
                              If intPos1 > 0& Then
                                ' ** This is JC_Key_Sub() in modJrnlCol_Keys.
                                intPos2 = InStr(strLine, "Select Case")
                                If intPos2 > 0 And intPos2 < intPos1 Then
                                  lngTmp02 = lngY
                                  Exit For
                                End If
                              End If  ' ** intPos1.
                            End If  ' ** Remark.
                          End If  ' ** vbNullString.
                          If .ProcOfLine(lngY - 1&, vbext_pk_Proc) <> strProcName Then
                            Exit For
                          End If
                        Next  ' ** lngY.
                        If lngTmp02 > 0& Then
                          strTmp01 = vbNullString
                          For lngY = (lngX - 1&) To lngTmp02 Step -1&
                            strLine = Trim$(.Lines(lngY, 1))
                            If strLine <> vbNullString Then
                              If Left$(strLine, 1) <> "'" Then
                                intPos1 = InStr(strLine, " Case ")
                                If intPos1 > 0& Then
                                  strTmp01 = Trim$(Mid$(strLine, (intPos1 + 5)))
                                  Exit For
                                End If
                              End If  ' ** Remark.
                            End If  ' ** vbNullString.
                          Next  ' ** lngY.
                          If strTmp01 <> vbNullString Then
                            intPos1 = InStr(strTmp01, "'")
                            If intPos1 > 0 Then strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))
                            Select Case Val(strTmp01)
                            Case 0
                              ' **   0 : Plain keys.
                              strKeyType = "Plain"
                              lngKeyType = taKeyDown_Plain
                            Case 1
                              ' **   1 : Shift keys.
                              strKeyType = "Shift"
                              lngKeyType = taKeyDown_Shift
                            Case 2
                              ' **   2 : Tab copies accountno.
                              strKeyType = "Plain"
                              lngKeyType = taKeyDown_Plain
                            Case 3
                            ' **   3 : Tax Lot screen.
                              strKeyType = "Plain"
                              lngKeyType = taKeyDown_Plain
                            Case 4
                            ' **   4 : CommitRec.
                              strKeyType = "Plain"
                              lngKeyType = taKeyDown_Plain
                            End Select
                          End If  ' ** strTmp01.
                        End If  ' ** lngTmp02.
                      End If  ' ** lngTmp02.
                      If strKeyType <> vbNullString Then
                        arr_varTab(T_TYP, lngE) = lngKeyType
                      Else
                        If strProcName = "SkipKey" Then
                          strKeyType = "Plain"
                          lngKeyType = taKeyDown_Plain
                          arr_varTab(T_TYP, lngE) = lngKeyType
                        End If
                      End If
                    End If  ' ** vbKeyTab, vbKeyReturn.
                  End If  ' ** Remark.
                End If  ' ** vbNullString.
              End If  ' ** JC_Key_Sub(), SkipKey().
            Next  ' ** lngX.
          End With  ' ** cod.
        End If
      End With  ' ** vbc.
    Next  ' ** vbc.

  End With  ' ** vbp.
  Set vbp = Nothing

  lngDels = 0&
  ReDim arr_varDel(0)

  Set dbs = CurrentDb
  With dbs
    Set rst = .OpenRecordset("zz_tbl_VBComponent_KeyDown", dbOpenDynaset, dbConsistent)
    With rst
      For lngX = 0& To (lngTabs - 1&)
        .FindFirst "[dbs_id] = " & CStr(arr_varTab(T_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varTab(T_VID, lngX)) & " And " & _
          "[vbcomproc_id] = " & CStr(arr_varTab(T_PID, lngX)) & " And [keydowntype_type] = " & arr_varTab(T_TYP, lngX) & " And " & _
          "[vbkeydown_tab] = " & CStr(arr_varTab(T_TAB, lngX)) & " And [vbkeydown_return] = " & CStr(arr_varTab(T_RET, lngX))
        Select Case .NoMatch
        Case True
          .AddNew
          ![dbs_id] = arr_varTab(T_DID, lngX)
          ![vbcom_id] = arr_varTab(T_VID, lngX)
          ![vbcomproc_id] = arr_varTab(T_PID, lngX)
          ' ** ![vbkeydown_id] : AutoNumber.
          ![keydowntype_type] = arr_varTab(T_TYP, lngX)
          ![vbkeydown_linenum] = arr_varTab(T_LIN, lngX)
          ![vbkeydown_tab] = arr_varTab(T_TAB, lngX)
          ![vbkeydown_return] = arr_varTab(T_RET, lngX)
          ![vbkeydown_datemodified] = Now()
          .Update
        Case False
          .Edit
          ![vbkeydown_linenum] = arr_varTab(T_LIN, lngX)
          ![vbkeydown_datemodified] = Now()
          .Update
        End Select  ' ** NoMatch.
      Next  ' ** lngX.
      .MoveFirst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        blnFound = False
        For lngY = 0& To (lngTabs - 1&)
          If arr_varTab(T_DID, lngY) = ![dbs_id] And arr_varTab(T_VID, lngY) = ![vbcom_id] And arr_varTab(T_PID, lngY) = ![vbcomproc_id] Then
            If arr_varTab(T_TYP, lngY) = ![keydowntype_type] Then
              If CBool(arr_varTab(T_TAB, lngY)) = ![vbkeydown_tab] And CBool(arr_varTab(T_RET, lngY)) = ![vbkeydown_return] Then
                blnFound = True
                Exit For
              End If
            End If
          End If
        Next  ' ** lngY.
        If blnFound = False Then
          lngDels = lngDels + 1&
          lngE = lngDels - 1&
          ReDim Preserve arr_varDel(lngE)
          arr_varDel(lngE) = ![vbkeydown_id]
        End If
        If lngX < lngRecs Then .MoveNext
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    If lngDels > 0& Then
Beep
Debug.Print "'DELS: " & CStr(lngDels)
Stop
      For lngX = 0& To (lngDels - 1&)
        ' ** Delete zz_tbl_VBComponent_KeyDown, by specified [kydwnid].
        Set qdf = .QueryDefs("zz_qry_Form_Shortcut_01c")
        With qdf.Parameters
          ![kydwnid] = arr_varDel(lngX)
        End With
        qdf.Execute
        Set qdf = Nothing
      Next
    End If
    .Close
  End With  ' ** dbs.

  Beep
  Debug.Print "'DONE! " & THIS_PROC & "()"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_KeyDown = blnRetValx

End Function

Public Function DAOH() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "DAOH"

  blnRetValx = VBA_DAOHelp  ' ** Function: Below.

  DAOH = blnRetValx

End Function

Public Function VBA_DAOHelp() As Boolean
' ** Called by:
' **   DAOH(), Above.

  Const THIS_PROC As String = "VBA_DAOHelp"

  Dim dblReturn As Double

  blnRetValx = True

  OpenDAO "C:\Documents and Settings\VictorC\Desktop\ADO-DAO\DAO360.CHM"  ' ** Module Function: modShellFuncs.

  VBA_DAOHelp = blnRetValx

End Function

Public Function VBA_Find_Code(Optional varFindStr1 As Variant, Optional varFindStr2 As Variant, Optional varChk As Variant) As Boolean
' ** If all are missing, only do simple search for strFind1, hard-coded below. blnCalled is False
' ** If varFindStr1 present, then look and Debug.Print as found. blnCalled is True
' ** If both find's present (varChk = False), use array and list at finish.
' ** If varChk True, and search found, then list multiple, specific lines of context.
' ** Currently not called.

  Const THIS_PROC As String = "VBA_Find_Code"

  Dim vbp As VBProject, vbc As VBComponent
  Dim blnCalled As Boolean, blnFind2 As Boolean, blnChk As Boolean, blnDelete As Boolean, blnCountOnly As Boolean
  Dim lngHits As Long, arr_varHit() As Variant, lngE As Long, lngCnts As Long, lngDisps As Long
  Dim strFind1 As String, strFind2 As String
  Dim blnSkipToNext As Boolean
  Dim lngParLine As Long, blnLineContinue As Boolean
  Dim strMdlName As String, strProcName As String, lngStartLine As Long, lngBodyLine As Long
  Dim strScope As String, strPersist As String, strType As String, strSubtype As String
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, strTmp05 As String, strTmp06 As String
  Dim lngTmp07 As Long, lngTmp08 As Long, lngTmp09 As Long
  Dim blnTmp10 As Boolean, blnTmp11 As Boolean
  Dim intPos1 As Integer, intPos2 As Integer
  Dim lngX As Long, lngY As Long, lngZ As Long

  Const HIT_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const HIT_WORD1 As Integer = 0
  Const HIT_WORD2 As Integer = 1
  Const HIT_MDL   As Integer = 2
  Const HIT_PROC  As Integer = 3
  Const HIT_PROCL As Integer = 4
  Const HIT_LNNUM As Integer = 5
  Const HIT_LINE  As Integer = 6
  Const HIT_PAR   As Integer = 7
  Const HIT_CONT  As Integer = 8
  Const HIT_CNT   As Integer = 9
  Const HIT_DISP  As Integer = 10

  blnRetValx = True

  ' ** How to refer to the code within an Access application:
  ' ** Application.VBE.VBProjects.Count                   : Only 1 for a normal MDB (1-based).
  ' ** Application.VBE.VBProjects(1).VBComponents.Count   : All code modules, both Standard and Form (1-based).
  ' ** Application.VBE.VBProjects(1).VBComponents(1).Type : Identifies which type it is, Standard, Class, or Form.
  ' **   vbext_ComponentType enumeration:
  ' **       1  vbext_ct_StdModule        Standard Module
  ' **       2  vbext_ct_ClassModule      Class Module for user-defined classes and objects.
  ' **       3  vbext_ct_MSForm           A UserForm. The visual component of a UserForm in the VBA Editor is called a designer.
  ' **      11  vbext_ct_ActiveXDesigner
  ' **     100  vbext_ct_Document         Module behind Form or Report.
  ' ** Application.VBE.VBProjects(1).VBComponents(1).CodeModule.CountOfDeclarationLines
  ' ** Application.VBE.VBProjects(1).VBComponents(1).CodeModule.CountOfLines (all lines, including Declaration section)
  ' ** Application.VBE.VBProjects(1).VBComponents(1).CodeModule.Lines(StartLine As Long, Count As Long) As String
  ' **   Returns string containing specified number of lines of code.
  ' ** Application.VBE.VBProjects(1).VBComponents(1).CodeModule.Name
  ' ** Modules(0).Type
  ' **   AcModuleType enumeration:
  ' **     0  acStandardModule  The specified module is a standard module.
  ' **     1  acClassModule     The specified module is a class module.
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
  ' ** .CodeModule.AddFromFile                    Method
  ' ** .CodeModule.AddFromString                  Method
  ' ** .CodeModule.CreateEventProc                Method
  ' ** .CodeModule.DeleteLines(StartLine, Count) Method
  ' **   Deletes a single line or a specified range of lines.
  ' ** .CodeModule.Find                           Method
  ' ** .CodeModule.InsertLines(Line as Long, String As String) Method
  ' **   Inserts a line or lines of code at a specified location in a block of code.
  ' ** Modules(0).InsertText Method
  ' **   When you insert a string by using the InsertText method, Microsoft Access
  ' **   places the new text at the end of the module, after all other procedures.
  ' ** .CodeModule.ReplaceLine(Line As Long, String As String) Method
  ' **   The ReplaceLine method replaces a specified line in a standard module.

  ' ** Application.VBE.CodePanes(1)
  ' **   .CodeModule
  ' **   .CodePaneView
  ' **     Vbext_CodepaneView enumeration:
  ' **       0  vbext_cv_ProcedureView   The specified code pane is in Procedure view.
  ' **       1  Vbext_cv_FullModuleView  The specified project is in Full Module view.
  ' **   .CountOfVisibleLines  Lines currently showing in the CodePane window.
  ' **   .Topline              Sets or returns the top line showing in the CodePane window.
  ' **   .Window
  ' **     .Caption
  ' **     .Height
  ' **     .Left
  ' **     .LinkedWindowFrame
  ' **     .LinkedWindows
  ' **     .SetFocus
  ' **     .Top
  ' **     .Type
  ' **     .Visible
  ' **     .Width
  ' **     .WindowState
  ' **        Returns or sets a numeric value specifying the visual state of the window. DOESN'T WORK!!!!
  ' **          Vbext_WindowState enumeration:
  ' **            0  vbext_ws_Normal    Normal (Default)
  ' **            1  vbext_ws_Minimize  Minimized (minimized to an icon)
  ' **            2  vbext_ws_Maximize  Maximized (enlarged to maximum size)

  ' ** Vbext_ProcKind enumeration:
  ' **   0  vbext_pk_Proc  Sub or Function
  ' **   1  vbext_pk_Let   Proper Let
  ' **   2  vbext_pk_Set   Proper Set
  ' **   3  vbext_pk_Get   Proper Get

  If IsMissing(varFindStr1) = True Then
    blnCalled = False
    blnDelete = False: blnCountOnly = False
    strFind1 = "If gstrReportCallingForm <> vbNullString Then"
    blnFind2 = False
    strFind2 = vbNullString
    blnChk = False
  Else
    strFind1 = varFindStr1
    blnCalled = True
    blnDelete = False: blnCountOnly = False
    If IsMissing(varFindStr2) = True Then
      blnFind2 = False
      strFind2 = vbNullString
    Else
      blnFind2 = True
      strFind2 = varFindStr2
    End If
    If IsMissing(varChk) = True Then
      blnChk = False
    Else
      blnChk = CBool(varChk)
    End If
  End If

  lngHits = 0&
  ReDim arr_varHit(HIT_ELEMS, 0)

  Set vbp = Application.VBE.VBProjects(1)
  With vbp

    If blnChk = False Then

      For Each vbc In .VBComponents
        With vbc
          strMdlName = .Name
          If Left$(strMdlName, 6) <> "zz_mod_" Then

            With .CodeModule

              strProcName = vbNullString: lngStartLine = 0&: lngBodyLine = 0&
              blnSkipToNext = False

              If .Find(strFind1, 1, 1, -1, -1) = False Then  ' ** -1 indicates last line or last column.
                ' ** .Find(target,startline,startcol,endline,endcol[,wholeword][,matchcase][,patternsearch]) As Boolean
                blnSkipToNext = True
              End If

              If blnSkipToNext = False Then
                ' ** Only search for location if it's found in the module.

                strTmp02 = vbNullString
                lngParLine = 0&
                blnLineContinue = False
                strProcName = "Declaration"

                ' ** Walk through each declaration line of the module.
                For lngX = 1& To .CountOfDeclarationLines
                  strTmp02 = Trim$(.Lines(lngX, 1))
                  Select Case blnFind2
                  Case False
                    If InStr(strTmp02, strFind1) > 0 Then
                      Debug.Print "'" & .Name & ": Declaration Line: " & CStr(lngX)
                    End If
                  Case True
                    lngCnts = 0&
                    If blnLineContinue = False Then lngParLine = 0&
                      ' ** This line is not a continuation of the previous one.
                    If Right$(strTmp02, 1) = " _" Then
                      ' ** This line continues on to the next one.
                      If lngParLine = 0& Then
                        ' ** This is the first line of a multi-line statement.
                        lngParLine = lngX
                        blnLineContinue = True
                      Else
                        ' ** This line is mid-statement: not the first, not the last.
                        blnLineContinue = True  ' ** Though this should already be true from the last loop!
                      End If
                    Else
                      ' ** This line doesn't continue.
                      blnLineContinue = False
                    End If
                    If InStr(strTmp02, strFind1) > 0 Then
                      If InStr(strTmp02, strFind2) > 0 Then
                        lngCnts = 3&
                      Else
                        lngCnts = 1&
                      End If
                    ElseIf InStr(strTmp02, strFind2) > 0 Then
                      lngCnts = 2&
                    End If
                    If lngCnts > 0& Then
                      lngHits = lngHits + 1&
                      lngE = lngHits - 1&
                      ReDim Preserve arr_varHit(HIT_ELEMS, lngE)
                      arr_varHit(HIT_WORD1, lngE) = strFind1
                      arr_varHit(HIT_WORD2, lngE) = strFind2
                      arr_varHit(HIT_MDL, lngE) = strMdlName
                      arr_varHit(HIT_PROC, lngE) = strProcName
                      arr_varHit(HIT_PROCL, lngE) = strProcName
                      arr_varHit(HIT_LNNUM, lngE) = lngX
                      arr_varHit(HIT_LINE, lngE) = strTmp02
                      arr_varHit(HIT_PAR, lngE) = lngParLine
                      arr_varHit(HIT_CONT, lngE) = blnLineContinue
                      arr_varHit(HIT_CNT, lngE) = lngCnts
                      arr_varHit(HIT_DISP, lngE) = False
                      lngE = 0&
                    End If
                    If blnLineContinue = False Then lngParLine = 0&
                  End Select
                Next

                strTmp02 = vbNullString: strTmp03 = vbNullString: strTmp04 = vbNullString
                intPos1 = 0: intPos2 = 0
                lngParLine = 0&
                blnLineContinue = False

                ' ** Walk through each non-declaration line of the module.
                For lngX = (.CountOfDeclarationLines + 1&) To .CountOfLines
                  ' ** Search entire module and list every occurrence of strFind1.

                  strProcName = .ProcOfLine(lngX, vbext_pk_Proc)

                  blnSkipToNext = False

                  strTmp02 = .Lines(lngX, 1)
                  Select Case blnFind2
                  Case False
                    If InStr(strTmp02, strFind1) = 0 Then blnSkipToNext = True
                  Case True
                    lngCnts = 0&
                    If blnLineContinue = False Then lngParLine = 0&
                      ' ** This line is not a continuation of the previous one.
                    If Right$(strTmp02, 1) = " _" Then
                      ' ** This line continues on to the next one.
                      If lngParLine = 0& Then
                        ' ** This is the first line of a multi-line statement.
                        lngParLine = lngX
                        blnLineContinue = True
                      Else
                        ' ** This line is mid-statement: not the first, not the last.
                        blnLineContinue = True  ' ** Though this should already be true from the last loop!
                      End If
                    Else
                      ' ** This line doesn't continue.
                      blnLineContinue = False
                    End If
                    If InStr(strTmp02, strFind1) > 0 Then
                      If InStr(strTmp02, strFind2) > 0 Then
                        lngCnts = 3&
                      Else
                        lngCnts = 1&
                      End If
                    ElseIf InStr(strTmp02, strFind2) > 0 Then
                      lngCnts = 2&
                    End If
                    If lngCnts > 0& Then
                      lngHits = lngHits + 1&
                      lngE = lngHits - 1&
                      ReDim Preserve arr_varHit(HIT_ELEMS, lngE)
                      arr_varHit(HIT_WORD1, lngE) = strFind1
                      arr_varHit(HIT_WORD2, lngE) = strFind2
                      arr_varHit(HIT_MDL, lngE) = strMdlName
                      arr_varHit(HIT_PROC, lngE) = strProcName
                      arr_varHit(HIT_PROCL, lngE) = vbNullString
                      arr_varHit(HIT_LNNUM, lngE) = lngX
                      arr_varHit(HIT_LINE, lngE) = strTmp02
                      arr_varHit(HIT_PAR, lngE) = lngParLine
                      arr_varHit(HIT_CONT, lngE) = blnLineContinue
                      arr_varHit(HIT_CNT, lngE) = lngCnts
                      arr_varHit(HIT_DISP, lngE) = False
                      lngE = 0&
                    Else
                      blnSkipToNext = True
                    End If
                  End Select

                  If blnSkipToNext = False Then

                    ' ** Now we have to know the correct ProcKind in order to get any
                    ' ** more information, and I don't want to use trial-and-error.

                    ' ** Move backwards to find the procedure's declaration line.
                    For lngY = lngX To 1& Step -1&

                      blnSkipToNext = False

                      strScope = vbNullString: strPersist = vbNullString: _
                        strType = vbNullString: strSubtype = vbNullString

                      strTmp03 = Trim$(.Lines(lngY, 1))
                      intPos1 = InStr(strTmp03, " ")  ' ** Declare line will have at least 1 space.
                      If intPos1 = 0 Then blnSkipToNext = True

                      If blnSkipToNext = False Then
                        ' ** Check for Scope specifier.
                        strTmp04 = Trim$(Left$(strTmp03, intPos1))
                        Select Case strTmp04
                        Case "Public"  ' ** Scope.
                          ' ** It's a Public something.
                          ' ** Procedure is accessible to all other procedures in all modules.
                          strScope = strTmp04
                        Case "Private"  ' ** Scope.
                          ' ** It's a Private something.
                          ' ** Procedure is accessible only to other procedures in the module where it is declared.
                          strScope = strTmp04
                        Case "Friend"  ' ** Scope.
                          ' ** It's a Friendly something.
                          ' ** Modifies the definition of a procedure in a form module or class module
                          ' ** to make the procedure callable from modules that are outside the class,
                          ' ** but part of the project within which the class is defined.
                          ' ** Friend procedures cannot be used in standard modules.
                          strScope = strTmp04
                        Case Else
                          ' ** No scope specified.
                        End Select
                        If strScope <> vbNullString Then
                          ' ** Check out the next word.
                          intPos2 = InStr(intPos1 + 1, strTmp03, " ")
                          If intPos2 > 0 Then
                            strTmp04 = Trim$(Mid$(strTmp03, intPos1, (intPos2 - intPos1)))
                            intPos1 = intPos2
                          Else
                            blnSkipToNext = True
                          End If
                        End If
                      End If

                      If blnSkipToNext = False Then
                        ' ** Check for Persistence specifier.
                        Select Case strTmp04
                        Case "Static"  ' ** Persistence.
                          ' ** Precedes the type whether or not a scope is specified.
                          ' ** Without Static, the value of local variables is not preserved between calls.
                          strPersist = strTmp04
                        Case Else
                          ' ** No persistence specified.
                        End Select
                        If strPersist <> vbNullString Then
                          ' ** Check out the next word.
                          intPos2 = InStr(intPos1 + 1, strTmp03, " ")
                          If intPos2 > 0 Then
                            strTmp04 = Trim$(Mid$(strTmp03, intPos1, (intPos2 - intPos1)))
                            intPos1 = intPos2
                          Else
                            blnSkipToNext = True
                          End If
                        End If
                      End If

                      If blnSkipToNext = False Then
                        ' ** Check for procedure Type specifier.
                        Select Case strTmp04
                        Case "Sub"  ' ** Type.
                          ' ** Procedure that can't return a value.
                          ' ** Without Scope, Public by default.
                          strType = strTmp04
                        Case "Function"  ' ** Type.
                          ' ** Procedure that may return a value.
                          ' ** Without Scope, Public by default.
                          strType = strTmp04
                        Case "Property"  ' ** Type.
                          ' ** Form or Class Property.
                          ' ** Without Scope, Public by default.
                          strType = strTmp04
                        Case Else
                          ' ** Not the first line of the procedure, keep looking.
                          blnSkipToNext = True
                        End Select
                        If strType <> vbNullString Then
                          Select Case strType
                          Case "Property"
                            ' ** Check out the next word.
                            intPos2 = InStr(intPos1 + 1, strTmp03, " ")
                            If intPos2 > 0 Then
                              strTmp04 = Trim$(Mid$(strTmp03, intPos1, (intPos2 - intPos1)))
                              intPos1 = intPos2
                            Else
                              blnSkipToNext = True
                            End If
                            If blnSkipToNext = False Then
                              Select Case strTmp04
                              Case "Let"  ' ** Property Subtype.
                                strSubtype = strTmp04
                              Case "Get"  ' ** Property Subtype.
                                strSubtype = strTmp04
                              Case "Set"  ' ** Property Subtype.
                                strSubtype = strTmp04
                              Case Else
                                ' ** Not a properly defined property, keep looking.
                                blnSkipToNext = True
                                strType = vbNullString
                              End Select
                            End If
                          End Select
                        Else
                          blnSkipToNext = True
                        End If
                      End If

                      If blnSkipToNext = False Then
                        ' ** If it hasn't skipped on this line, it means we're on the declaration line.
                        lngBodyLine = lngY  'lngStartLine
                        Exit For
                      Else
                        ' ** Declaration line not yet found, keep looking.
                      End If

                    Next  ' ** lngY: previous line in procedure.

                    If strType <> vbNullString Then
                      strTmp02 = vbNullString
                      If strScope <> vbNullString Then strTmp02 = strScope & " "
                      If strPersist <> vbNullString Then strTmp02 = strTmp02 & strPersist & " "
                      strTmp02 = strTmp02 & strType & " "
                      If strSubtype <> vbNullString Then strTmp02 = strTmp02 & strSubtype & " "
                      strTmp02 = strTmp02 & strProcName
                      If blnFind2 = False And blnDelete = False Then
                        Debug.Print "'" & .Name & ": " & strTmp02 & "() Line: " & CStr(lngX)
                      Else
                        lngHits = lngHits + 1&
                        lngE = lngHits - 1&
                        ReDim Preserve arr_varHit(HIT_ELEMS, lngE)
                        arr_varHit(HIT_WORD1, lngE) = strFind1
                        arr_varHit(HIT_WORD2, lngE) = strFind2
                        arr_varHit(HIT_MDL, lngE) = strMdlName
                        arr_varHit(HIT_PROC, lngE) = strProcName
                        arr_varHit(HIT_PROCL, lngE) = strTmp02
                        arr_varHit(HIT_LNNUM, lngE) = lngX
                        arr_varHit(HIT_LINE, lngE) = strTmp02
                        arr_varHit(HIT_PAR, lngE) = lngParLine
                        arr_varHit(HIT_CONT, lngE) = blnLineContinue
                        arr_varHit(HIT_CNT, lngE) = lngCnts
                        arr_varHit(HIT_DISP, lngE) = False
                        lngE = 0&
                      End If
                    End If

                  End If  ' ** blnSkipToNext

                  If blnLineContinue = False Then lngParLine = 0&

                Next  ' ** lngX: next line in .CodeModule

              End If  ' ** blnSkipToNext, Find successful.

            End With  ' ** .CodeModule

          End If
        End With ' ** vbc (.VBComponent)
      Next  ' ** Each vbc: next component in .VBComponents

    Else  ' ** blnChk = True.

      lngHits = 15&
      lngE = lngHits - 1&
      ReDim arr_varHit(HIT_ELEMS, lngE)

      arr_varHit(HIT_MDL, 0) = "Form_XXX"
      arr_varHit(HIT_PROC, 0) = "Delete_Click"
      arr_varHit(HIT_LNNUM, 0) = 180&

      strMdlName = vbNullString: strProcName = vbNullString

      For lngX = 0& To (lngHits - 1&)
        Set vbc = .VBComponents(arr_varHit(HIT_MDL, lngX))
        With vbc
          With vbc.CodeModule
            If arr_varHit(HIT_MDL, lngX) <> strMdlName Then
              strMdlName = arr_varHit(HIT_MDL, lngX)
              Debug.Print "'" & strMdlName
              strProcName = vbNullString
            End If
            If arr_varHit(HIT_PROC, lngX) <> strProcName Then
              strProcName = arr_varHit(HIT_PROC, lngX)
              Debug.Print "'  ." & arr_varHit(HIT_PROC, lngX) & "()"
            End If
            Debug.Print "'    " & arr_varHit(HIT_LNNUM, lngX) & ": " & Trim$(.Lines(arr_varHit(HIT_LNNUM, lngX), 1))
          End With
        End With
      Next

    End If  ' ** blnChk.

  End With  ' ** vbp (.VBProject)

  If blnFind2 = True And blnChk = False Then

    lngDisps = 0&

    ' ** If both words hit on the same line, absolutely display it!
    For lngX = 0& To (lngHits - 1&)
      If arr_varHit(HIT_CNT, lngX) = 3& Then
        arr_varHit(HIT_DISP, lngX) = True
        lngDisps = lngDisps + 1&
      End If
    Next

    ' ** Check if both words are at least within the same multi-line statement.
    For lngX = 0& To (lngHits - 1&)
      If arr_varHit(HIT_DISP, lngX) = False Then

        If arr_varHit(HIT_PAR, lngX) > 0& Then
          ' ** This hit is within a multi-line statement.

          For lngY = 0& To (lngHits - 1&)
            ' ** See if another hit has the same parent, and it's the other word.
            If arr_varHit(HIT_PAR, lngY) = arr_varHit(HIT_PAR, lngX) And _
               arr_varHit(HIT_LNNUM, lngY) <> arr_varHit(HIT_LNNUM, lngX) Then
              ' ** Don't compare it to itself!
              If (arr_varHit(HIT_CNT, lngY) = 1& And arr_varHit(HIT_CNT, lngX) = 2&) Or _
                 (arr_varHit(HIT_CNT, lngY) = 2& And arr_varHit(HIT_CNT, lngX) = 1&) Then
                ' ** Yes, both words are within the same multi-line statement.
                arr_varHit(HIT_DISP, lngX) = True
                lngDisps = lngDisps + 1&
                Exit For
              End If
            End If
          Next

          If arr_varHit(HIT_DISP, lngX) = False Then
            For lngY = 0& To (lngHits - 1&)
              ' ** See if another hit is the parent, and it's the other word.
              If arr_varHit(HIT_LNNUM, lngY) = arr_varHit(HIT_PAR, lngX) Then
                ' ** Yes, the parent also had a hit (arr_varHit(HIT_PAR, lngY) should be 0).
                If (arr_varHit(HIT_CNT, lngY) = 1& And arr_varHit(HIT_CNT, lngX) = 2&) Or _
                   (arr_varHit(HIT_CNT, lngY) = 2& And arr_varHit(HIT_CNT, lngX) = 1&) Then
                  ' ** Yes, both words are within the same multi-line statement.
                  arr_varHit(HIT_DISP, lngX) = True
                  lngDisps = lngDisps + 1&
                  Exit For
                End If
              End If
            Next
          End If

        Else
          ' ** This hit could be the parent of another hit that's also the other word.

          For lngY = 0& To (lngHits - 1&)
            ' ** See if another hit has this as its parent, and it's the other word.
            If arr_varHit(HIT_PAR, lngY) = arr_varHit(HIT_LNNUM, lngX) Then
              ' ** Yes, another hit lists this as its parent.
              If (arr_varHit(HIT_CNT, lngY) = 1& And arr_varHit(HIT_CNT, lngX) = 2&) Or _
                 (arr_varHit(HIT_CNT, lngY) = 2& And arr_varHit(HIT_CNT, lngX) = 1&) Then
                ' ** Yes, both words are within the same multi-line statement.
                arr_varHit(HIT_DISP, lngX) = True
                lngDisps = lngDisps + 1&
                Exit For
              End If
            End If
          Next

        End If

      End If
    Next

'For now, we'll just look within the same statement.
'We could broaden it to the same procedure or module.

    If lngHits > 0& Then
      Debug.Print "'" & strFind1 & ", " & strFind2 & " : " & CStr(lngDisps); " of " & CStr(lngHits)
      strMdlName = vbNullString: strProcName = vbNullString
      lngCnts = 0&
      For lngX = 0& To (lngHits - 1&)
        If arr_varHit(HIT_DISP, lngX) = True Then
          lngCnts = lngCnts + 1&
          If arr_varHit(HIT_MDL, lngX) <> strMdlName Then
            strMdlName = arr_varHit(HIT_MDL, lngX)
            Debug.Print "'" & strMdlName
            strProcName = vbNullString
          End If
          If arr_varHit(HIT_PROC, lngX) <> strProcName Then
            strProcName = arr_varHit(HIT_PROC, lngX)
            Debug.Print "'  ." & arr_varHit(HIT_PROC, lngX) & "()"
          End If
          Debug.Print "'    " & arr_varHit(HIT_LNNUM, lngX) & ": " & Trim$(arr_varHit(HIT_LINE, lngX))
        End If
      Next
    End If

    ' ******************************************************************************
    ' ** Array: arr_varHit()
    ' **
    ' **   Element  Description          Source             Type       Constant
    ' **   =======  ===================  =================  =========  ===========
    ' **      0     Search term 1        strFind1           String     HIT_WORD1
    ' **      1     Search term 2        strFind2           String     HIT_WORD2
    ' **      2     Module Name          strMdlName         String     HIT_MDL
    ' **      3     Procedure Name       strProcName        String     HIT_PROC
    ' **      4       With Types         strTmp02            String     HIT_PROCL
    ' **      5     Line Number          lngX               Long       HIT_LNNUM
    ' **      6     Text of Line         strTmp02            String     HIT_LINE
    ' **      7     Procedure Line       lngParLine         Long       HIT_PAR
    ' **      8     Line Continuation    blnLineContinue    Boolean    HIT_CONT
    ' **      9     Hits                 lngCnts            Long       HIT_CNT
    ' **     10     Display              False              Boolean    HIT_DISP
    ' **
    ' ******************************************************************************

  End If  ' ** blnFind2.

  If blnDelete = True And lngHits > 0& Then

      ' ** Binary Sort arr_varHit() array by Module and descending Line Number.
      For lngZ = 1& To 2&  ' ** Twice to make sure line numbers also correctly sorted.
        For lngX = UBound(arr_varHit, 2) To 1& Step -1&
          For lngY = 0& To (lngX - 1&)
            If arr_varHit(HIT_MDL, lngY) > arr_varHit(HIT_MDL, (lngY + 1&)) Or _
              (arr_varHit(HIT_MDL, lngY) = arr_varHit(HIT_MDL, (lngY + 1&)) And _
               arr_varHit(HIT_LNNUM, lngY) < arr_varHit(HIT_LNNUM, (lngY + 1&))) Then
              strTmp01 = arr_varHit(HIT_WORD1, lngY)
              strTmp02 = arr_varHit(HIT_WORD2, lngY)
              strTmp03 = arr_varHit(HIT_MDL, lngY)
              strTmp04 = arr_varHit(HIT_PROC, lngY)
              strTmp05 = arr_varHit(HIT_PROCL, lngY)
              lngTmp07 = arr_varHit(HIT_LNNUM, lngY)
              strTmp06 = arr_varHit(HIT_LINE, lngY)
              lngTmp08 = arr_varHit(HIT_PAR, lngY)
              blnTmp10 = arr_varHit(HIT_CONT, lngY)
              lngTmp09 = arr_varHit(HIT_CNT, lngY)
              blnTmp11 = arr_varHit(HIT_DISP, lngY)
              arr_varHit(HIT_WORD1, lngY) = arr_varHit(HIT_WORD1, (lngY + 1&))
              arr_varHit(HIT_WORD2, lngY) = arr_varHit(HIT_WORD2, (lngY + 1&))
              arr_varHit(HIT_MDL, lngY) = arr_varHit(HIT_MDL, (lngY + 1&))
              arr_varHit(HIT_PROC, lngY) = arr_varHit(HIT_PROC, (lngY + 1&))
              arr_varHit(HIT_PROCL, lngY) = arr_varHit(HIT_PROCL, (lngY + 1&))
              arr_varHit(HIT_LNNUM, lngY) = arr_varHit(HIT_LNNUM, (lngY + 1&))
              arr_varHit(HIT_LINE, lngY) = arr_varHit(HIT_LINE, (lngY + 1&))
              arr_varHit(HIT_PAR, lngY) = arr_varHit(HIT_PAR, (lngY + 1&))
              arr_varHit(HIT_CONT, lngY) = arr_varHit(HIT_CONT, (lngY + 1&))
              arr_varHit(HIT_CNT, lngY) = arr_varHit(HIT_CNT, (lngY + 1&))
              arr_varHit(HIT_DISP, lngY) = arr_varHit(HIT_DISP, (lngY + 1&))
              arr_varHit(HIT_WORD1, (lngY + 1&)) = strTmp01
              arr_varHit(HIT_WORD2, (lngY + 1&)) = strTmp02
              arr_varHit(HIT_MDL, (lngY + 1&)) = strTmp03
              arr_varHit(HIT_PROC, (lngY + 1&)) = strTmp04
              arr_varHit(HIT_PROCL, (lngY + 1&)) = strTmp05
              arr_varHit(HIT_LNNUM, (lngY + 1&)) = lngTmp07
              arr_varHit(HIT_LINE, (lngY + 1&)) = strTmp06
              arr_varHit(HIT_PAR, (lngY + 1&)) = lngTmp08
              arr_varHit(HIT_CONT, (lngY + 1&)) = blnTmp10
              arr_varHit(HIT_CNT, (lngY + 1&)) = lngTmp09
              arr_varHit(HIT_DISP, (lngY + 1&)) = blnTmp11
            End If
          Next
        Next
      Next

      If blnCountOnly = True Then
        Debug.Print "'HITS: " & CStr(lngHits)
      Else
        lngCnts = 0&
        Set vbp = Application.VBE.VBProjects(1)
        With vbp
          For lngX = 0& To (lngHits - 1&)
            lngCnts = lngCnts + 1&
            If arr_varHit(HIT_MDL, lngX) <> "z_mdl_Misc_Dev_Funcs" Then
              Set vbc = .VBComponents(arr_varHit(HIT_MDL, lngX))
              With vbc
                With .CodeModule
                  .DeleteLines arr_varHit(HIT_LNNUM, lngX), 1
                End With
              End With
            End If
          Next
        End With
        Debug.Print "'LINES DELETED: " & CStr(lngCnts)
      End If

  End If

  Beep

  Set vbc = Nothing
  Set vbp = Nothing

  VBA_Find_Code = blnRetValx

End Function

Public Function VBA_Err_Sort(varInput As Variant) As String
' ** Currently not called.

  Const THIS_PROC As String = "VBA_Err_Sort"

  Dim lngErrs As Long, arr_varErr() As Variant
  Dim blnFound As Boolean
  Dim lngTmp00 As Long, strTmp01 As String, strTmp02 As String
  Dim intPos1 As Integer
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim strRetVal As String

  Const ERR_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const ERR_NUM As Integer = 0
  Const ERR_CNT As Integer = 1

  strRetVal = vbNullString

  If IsNull(varInput) = False Then
    ' ** Case 3101, 8519, 2108, 2116, 3020, 2501, 2169, 7753, 3314, 2237, 2046

    lngErrs = 0&
    ReDim arr_varErr(ERR_ELEMS, 0)

    strTmp01 = Trim$(varInput)
    intPos1 = InStr(strTmp01, " ")
    If intPos1 > 0 Then
      strTmp02 = Trim$(Left$(strTmp01, intPos1))
      If IsNumeric(strTmp02) = True Then strTmp01 = Trim$(Mid$(strTmp01, intPos1))
      If Left$(strTmp01, 5) = "Case " Then strTmp01 = Trim$(Mid$(strTmp01, 6))
      intPos1 = InStr(strTmp01, ",")
      Do While intPos1 > 0
        lngErrs = lngErrs + 1&
        lngE = lngErrs - 1&
        ReDim Preserve arr_varErr(ERR_ELEMS, lngE)
        arr_varErr(ERR_NUM, lngE) = CLng(Val(Left$(strTmp01, (intPos1 - 1))))
        arr_varErr(ERR_CNT, lngE) = CInt(0)
        strTmp01 = Trim$(Mid$(strTmp01, (intPos1 + 1)))
        intPos1 = InStr(strTmp01, ",")
      Loop
      lngErrs = lngErrs + 1&
      lngE = lngErrs - 1&
      ReDim Preserve arr_varErr(ERR_ELEMS, lngE)
      arr_varErr(ERR_NUM, lngE) = CLng(Val(strTmp01))
      arr_varErr(ERR_CNT, lngE) = CInt(0)
      ' ** Binary Sort arr_varErr() array.
      For lngX = UBound(arr_varErr, 2) To 1& Step -1&
        For lngY = 0& To (lngX - 1)
          If arr_varErr(ERR_NUM, lngY) > arr_varErr(ERR_NUM, (lngY + 1)) Then
            lngTmp00 = arr_varErr(ERR_NUM, lngY)
            arr_varErr(ERR_NUM, lngY) = arr_varErr(ERR_NUM, lngY + 1)
            arr_varErr(ERR_NUM, (lngY + 1)) = lngTmp00
          End If
        Next
      Next
      strTmp01 = "Case "
      For lngX = 0& To (lngErrs - 1&)
        blnFound = False
        For lngY = 0& To (lngErrs - 1&)
          If arr_varErr(ERR_NUM, lngY) = arr_varErr(ERR_NUM, lngX) Then
            arr_varErr(ERR_CNT, lngY) = arr_varErr(ERR_CNT, lngY) + 1
            Exit For
          End If
        Next
        If arr_varErr(ERR_CNT, lngX) = 1 Then
          ' ** Check for dupes.
          strTmp01 = strTmp01 & CStr(arr_varErr(ERR_NUM, lngX)) & ", "
        End If
      Next
      strTmp01 = Trim$(strTmp01)
      If Right$(strTmp01, 1) = "," Then strTmp01 = Left$(strTmp01, (Len(strTmp01) - 1))
      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      strRetVal = strTmp01
    End If

  End If

  Beep

  VBA_Err_Sort = strRetVal

End Function

Public Function VBA_Err_Copy() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "VBA_Err_Copy"

  Dim vbp As VBProject
  Dim vbc1 As VBComponent, cod1 As CodeModule
  Dim vbc2 As VBComponent, cod2 As CodeModule
  Dim strModName1 As String, lngModLines1 As Long
  Dim strModName2 As String, lngModLines2 As Long
  Dim lngRefs As Long, arr_varRef() As Variant
  Dim strFind1 As String, strFind2 As String
  Dim strLine As String, strProcName As String
  Dim lngElseNum As Long
  Dim blnFound As Boolean
  Dim lngResponse As Long
  Dim intPos1 As Integer
  Dim strTmp01 As String
  Dim lngX As Long, lngY As Long, lngElemR As Long

  Const PFX As String = "Form_"

  Const RF_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const RF_LINE2  As Integer = 0
  Const RF_LNNUM2 As Integer = 1
  Const RF_POS2   As Integer = 2
  Const RF_PROC   As Integer = 3
  Const RF_FOUND2 As Integer = 4
  Const RF_LINE1  As Integer = 5
  Const RF_LNNUM1 As Integer = 6
  Const RF_FOUND1 As Integer = 7

  Const FRMX As String = "Statement Parameters"
  Const FIND1 As String = "zErrorHandler THIS_NAME"
  Const FIND2 As String = "Form_Error("
  Const FIND3 As String = "Form_Error "

  'VBA_FrmImport FRMX  ' ** Function: Below
  'Exit Function

' ** Also called in response to errors within other form procedures.
' ** Refer to this instead of zErrorHandler(), since it has the
' ** individual error responses accounted for.

        ' ** AcDataError enumeration:
        ' **   0  acDataErrContinue  Ignore the error and continue without displaying the default Microsoft Access
        ' **                         error message. A custom error message may be displayed in place of the default
        ' **                         error message.
        ' **   1  acDataErrDisplay   Display the default Microsoft Access error message. (Default)
        ' **   2  acDataErrAdded     Don't display the default Microsoft Access error message. The entry may be
        ' **                         added to the combo box list in the NotInList event procedure. After the entry
        ' **                         is added, Microsoft Access updates the list by requerying the combo box.
        ' **                         Microsoft Access then rechecks the string against the combo box list, and saves
        ' **                         the value in the NewData argument in the field the combo box is bound to. If
        ' **                         the string is not in the list, then Microsoft Access displays an error message.

            ' ** Do nothing.
            lngResponse = acDataErrContinue
            ' ** Errors dismissed:
            ' **     13  Type mismatch.
            ' **     91  Object variable or With block variable not set.
            ' **     94  Invalid use of Null.
            ' **   2001  You Canceled the previous operation.
            ' **   2046  The command or action '|' isn't available now.
            ' **   2105  You can't go to the specified record.
            ' **   2108  You must save the field before you execute the GoToControl action, the GoToControl method,
            ' **         or the SetFocus method.
            ' **   2110  Access can't move the focus to the control '|'.
            ' **   2113  The value you entered isn't valid for this field.
            ' **   2116  The value in the field or record violates the validation rule for the record or field.
            ' **   2135  This Property Is ReadOnly And can't be set.
            ' **   2169  You can 't save this record at this time.
            ' **   2202  You must install a printer before you design, print, or preview.
            ' **   2237  The text you entered isn't an item in the list.
            ' **   2279  The value you entered isn't appropriate for the input mask '|' specified for this field.
            ' **   2427  You entered an expression that has no value.
            ' **   2474  The expression you entered requires the control to be in the active window.
            ' **   2501  The '|' action was Canceled.
            ' **   3020  Update or CancelUpdate without AddNew or Edit.
            ' **   3022  The changes you requested to the table were not successful because they would create
            ' **         duplicate values in the index, primary key, or relationship.
            ' **   3058  Index or primary key cannot contain a Null value.
            ' **   3075  |1 in query expression '|2'.
            ' **   3101  The Microsoft Jet database engine cannot find a record in the table '|2'
            ' **         with key matching field(s) '|1'.
            ' **   3162  You tried to assign the Null value to a variable that is not a Variant data type.
            ' **   3163  The field is too small to accept the amount of data you attempted to add.
            ' **   3314  The field '|' cannot contain a Null value because the Required property for this field
            ' **         is set to True.
            ' **   3315  Field '|' cannot be a zero-length string.
            ' **   7753  The value in the control violates the validation rule for the control.
            ' **   8519  You are about to delete '|' record(s).
            ' **   8530  Relationships that specify cascading deletes are about to cause '|' record(s) in this
            ' **         table and in related tables to be deleted.
            ' **  10503  You are about to run a delete query that will modify data in your table.
            ' **  10508  You are about to delete | row(s) from the specified table.

  blnRetValx = True

  lngRefs = 0&
  ReDim arr_varRef(RF_ELEMS, 0)


  Set vbp = Application.VBE.ActiveVBProject
  With vbp

    ' ** New form module.
    Set vbc1 = .VBComponents(PFX & FRMX)
    strModName1 = vbc1.Name
    Set cod1 = vbc1.CodeModule
    lngModLines1 = cod1.CountOfLines

    ' ** Old form module.
    Set vbc2 = .VBComponents(PFX & FRMX & "1")
    strModName2 = vbc2.Name
    Set cod2 = vbc2.CodeModule
    lngModLines2 = cod2.CountOfLines

    ' ** Find references to Form_Error() in the old form module.
    With cod2
      For lngX = 1& To lngModLines2
        strLine = Trim$(.Lines(lngX, 1))
        If strLine <> vbNullString Then
          If Left$(strLine, 1) <> "'" Then
            intPos1 = InStr(strLine, FIND2)
            If intPos1 > 0 Then
              strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
              Select Case strProcName
              Case "cmbTaxCodes_BeforeUpdate", "Form_Error", "cmdUndo_Click"
                ' ** Skip 'em!
              Case Else
                lngRefs = lngRefs + 1&
                lngElemR = lngRefs - 1&
                ReDim Preserve arr_varRef(RF_ELEMS, lngElemR)
                arr_varRef(RF_LINE2, lngElemR) = strLine
                arr_varRef(RF_LNNUM2, lngElemR) = lngX
                arr_varRef(RF_POS2, lngElemR) = intPos1
                Select Case strProcName
                Case "cmdDividendClose_Click", "cmdInterestClose_Click", "cmdSaleClose_Click", _
                     "cmdPurchaseClose_Click", "cmdMiscClose_Click", "cmdAddEditAssetClose_Click"
                  strProcName = "cmdClose_Click"
                Case "opgJournal_Click"
                  strProcName = "opgOptions_Click"
                Case "dividendTransDate_Enter"
                  Select Case FRMX
                  Case "frmJournal_Sub1_Dividend"
                    strProcName = "dividendTransDate_Enter"
                  End Select
                Case "dividendTransDate_LostFocus"
                  Select Case FRMX
                  Case "frmJournal_Sub1_Dividend"
                    strProcName = "dividendTransDate_LostFocus"
                  End Select
                Case "dividendTransDate_Exit"
                  Select Case FRMX
                  Case "frmJournal_Sub1_Dividend"
                    strProcName = "dividendTransDate_Exit"
                  End Select
                Case "Text270_Change"
                  Select Case FRMX
                  Case "frmAccountProfile"
                    strProcName = "Notes_Change"
                  End Select
                Case "Text270_Exit"
                  Select Case FRMX
                  Case "frmAccountProfile"
                    strProcName = "Notes_Exit"
                  End Select
                Case "cmdPreviewReview_Click"
                  Select Case FRMX
                  Case "frmRpt_AccountReviews"
                    strProcName = "cmdPreview_Click"
                  End Select
                Case "cmdPrintReview_Click", "cmdPrintTaxLot_Click", "cmdPrintTransactions_Click"
                  Select Case FRMX
                  Case "frmRpt_AccountReviews"
                    strProcName = "cmdPrint_Click"
                  Case "frmRpt_TaxLot"
                    strProcName = "cmdPrint_Click"
                  Case "frmTransactions"
                    strProcName = "cmdPrint_Click"
                  End Select
                Case "cmbPrintTransactions_Click"
                  Select Case FRMX
                  Case "Statement Parameters"
                    strProcName = "cmdTransactionsPrint_Click"
                  End Select
                Case "cmdAssetListtoExcel_Click"
                  Select Case FRMX
                  Case "Statement Parameters"
                    strProcName = "cmdAssetListExcel_Click"
                  End Select
                Case "Form_Load"
                  Select Case FRMX
                  Case "frmJournal_Sub5_Misc", "frmJournal_Sub3_Purchase", "frmJournal_Sub4_Sold", "frmMap_Split"
                    strProcName = "Form_Open"
                  End Select
                Case "purchaseType_Click", "saleType_Click"
                  Select Case FRMX
                  Case "frmJournal_Sub3_Purchase"
                    strProcName = "purchaseJournalType_Click"
                  Case "frmJournal_Sub4_Sold"
                    strProcName = "saleJournalType_Click"
                  End Select
                Case "miscType_Enter", "purchaseType_Enter", "saleType_Enter"
                  Select Case FRMX
                  Case "frmJournal_Sub5_Misc"
                    strProcName = "miscJournalType_Enter"
                  Case "frmJournal_Sub3_Purchase"
                    strProcName = "purchaseJournalType_Enter"
                  Case "frmJournal_Sub4_Sold"
                    strProcName = "saleJournalType_Enter"
                  End Select
                Case "miscType_Exit", "purchaseType_Exit", "saleType_Exit"
                  Select Case FRMX
                  Case "frmJournal_Sub5_Misc"
                    strProcName = "miscJournalType_Exit"
                  Case "frmJournal_Sub3_Purchase"
                    strProcName = "purchaseJournalType_Exit"
                  Case "frmJournal_Sub4_Sold"
                    strProcName = "saleJournalType_Exit"
                  End Select
                Case "cmdExit_Click"
                  Select Case FRMX
                  Case "frmAssets_Add_Purchase"
                    strProcName = "cmdClose_Click"
                  End Select
                Case "cmbCancel_Click"
                  Select Case FRMX
                  Case "frmReinvest_Dividend", "frmReinvest_Interest"
                    strProcName = "cmdCancel_Click"
                  End Select
                End Select
                arr_varRef(RF_PROC, lngElemR) = strProcName
                arr_varRef(RF_FOUND2, lngElemR) = CBool(False)
                arr_varRef(RF_LINE1, lngElemR) = vbNullString
                arr_varRef(RF_LNNUM1, lngElemR) = 0&
                arr_varRef(RF_FOUND1, lngElemR) = CBool(False)
              End Select
            End If
          End If
        End If
      Next
    End With

    If lngRefs > 0& Then

      With cod1
        For lngX = 1& To lngModLines1
          strLine = Trim$(.Lines(lngX, 1))
          If strLine <> vbNullString Then
            If Left$(strLine, 1) <> "'" Then
              blnFound = False
              If InStr(strLine, FIND1) > 0 Then
                strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                For lngY = 0& To (lngRefs - 1&)
                  lngElemR = lngY
                  If arr_varRef(RF_PROC, lngY) = strProcName Then
                    blnFound = True
                    arr_varRef(RF_FOUND2, lngElemR) = CBool(True)
                    arr_varRef(RF_LINE1, lngElemR) = .Lines(lngX, 1)
                    arr_varRef(RF_LNNUM1, lngElemR) = lngX
                    Exit For
                  End If
                Next
              ElseIf InStr(strLine, FIND2) > 0 Or InStr(strLine, FIND3) > 0 Then
                strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
                For lngY = 0& To (lngRefs - 1&)
                  lngElemR = lngY
                  If arr_varRef(RF_PROC, lngY) = strProcName Then
                    blnFound = True
                    arr_varRef(RF_FOUND2, lngElemR) = CBool(True)
                    arr_varRef(RF_LINE1, lngElemR) = .Lines(lngX, 1)
                    arr_varRef(RF_LNNUM1, lngElemR) = lngX
                    'Exit For
                  End If
                Next
              End If
            End If
          End If
        Next
      End With

      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

      blnFound = True
      For lngX = 0& To (lngRefs - 1&)
        lngElemR = lngX
        If arr_varRef(RF_FOUND2, lngElemR) = False Then
          blnFound = False
          Debug.Print "'NOT FOUND!¹ " & arr_varRef(RF_PROC, lngElemR)
        End If
      Next
      If blnFound = True Then
        'Debug.Print "'ALL FOUND!"
      End If

      If blnFound = True Then
        With cod1

          For lngX = 0& To (lngRefs - 1&)
            lngElemR = lngX
            lngElseNum = (arr_varRef(RF_LNNUM1, lngElemR) - 1&)
            strLine = .Lines(lngElseNum, 1)
            intPos1 = InStr(strLine, "Case Else")
            If intPos1 > 0 Then
              arr_varRef(RF_FOUND1, lngElemR) = CBool(True)
              If InStr(strLine, "Form_Error") = 0 Then
                strTmp01 = strLine & "  'Call " & Mid$(arr_varRef(RF_LINE2, lngElemR), arr_varRef(RF_POS2, lngElemR))
                .ReplaceLine lngElseNum, strTmp01
              End If
            ElseIf InStr(arr_varRef(RF_LINE1, lngElemR), FIND2) > 0 Then
              arr_varRef(RF_FOUND1, lngElemR) = CBool(True)
            End If
          Next

          blnFound = True
          For lngX = 0& To (lngRefs - 1&)
            lngElemR = lngX
            If arr_varRef(RF_FOUND2, lngElemR) = False Then
              blnFound = False
              Debug.Print "'CASE ELSE NOT FOUND!² " & arr_varRef(RF_LNNUM1, lngElemR) & " " & arr_varRef(RF_PROC, lngElemR)
            End If
          Next
          If blnFound = True Then
            Debug.Print "'ALL FOUND!"
          End If

        End With
      End If

    Else
      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      Debug.Print "'NO Form_Error() REFS FOUND!"
    End If

  End With

  Beep

  Set cod1 = Nothing
  Set cod2 = Nothing
  Set vbc1 = Nothing
  Set vbc2 = Nothing
  Set vbp = Nothing

  VBA_Err_Copy = blnRetValx

End Function

Public Function VBA_FrmImport(Optional varFrmName As Variant) As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "VBA_FrmImport"

  blnRetValx = True

  If Len(TA_SEC) > Len(TA_SEC2) Then
    DoCmd.TransferDatabase acImport, "Microsoft Access", "C:\VictorGCS_Clients\TrustAccountant\NewDemo\Trust.mdb", _
      acForm, CStr(varFrmName), (CStr(varFrmName) & "1")
  Else
    DoCmd.TransferDatabase acImport, "Microsoft Access", "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Trust.mdb", _
      acForm, CStr(varFrmName), (CStr(varFrmName) & "1")
  End If

  VBA_FrmImport = blnRetValx

End Function

Public Function VBA_Find_Reports() As Boolean
' ** Currently not called.

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strFind As String, blnForms As Boolean
  Dim strModName As String
  Dim lngLines As Long
  Dim strLine As String
  Dim lngX As Long

  blnRetValx = True

  strFind = "OpenReport"
  blnForms = False

  If blnForms = False Then
    ' ** Walk through every module.
    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For Each vbc In .VBComponents
        With vbc
          strModName = .Name
          If Left$(strModName, 2) <> "z_" Then
            Set cod = .CodeModule
            With cod
              lngLines = .CountOfLines
              For lngX = 1& To lngLines
                strLine = Trim$(.Lines(lngX, 1))
                If strLine <> vbNullString Then
                  If Left$(strLine, 1) <> "'" Then
                    If InStr(strLine, strFind) > 0 Then
                      Debug.Print "'" & strModName
                      Exit For
                    End If
                  End If
                End If
              Next
            End With
          End If
        End With
      Next
    End With
  Else
    ' ** Just use Frm_Prop() in zz_mod_Misc_Dev_Funcs.
  End If

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_Find_Reports = blnRetValx

End Function

Public Function VBA_DupeFormFormat() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "VBA_DupeFormFormat"

  Dim frm1 As Access.Form, frm2 As Access.Form, ctl1 As Access.Control, ctl2 As Access.Control
  Dim strForm1 As String, strForm2 As String

  blnRetValx = True

  strForm1 = "frmReinvest_Interest"  ' ** New one.
  strForm2 = "frmReinvest_Dividend"  ' ** Standard one.

  Set frm1 = Forms(strForm1)
  Set frm2 = Forms(strForm2)

  With frm1
    For Each ctl1 In .Controls
      With ctl1
        .Top = frm2.Controls(.Name).Top
        .Left = frm2.Controls(.Name).Left
        .Width = frm2.Controls(.Name).Width
        .Height = frm2.Controls(.Name).Height
        .FontName = frm2.Controls(.Name).FontName
        .FontSize = frm2.Controls(.Name).FontSize
      End With
    Next
  End With

  Beep

  Set ctl1 = Nothing
  Set ctl2 = Nothing
  Set frm1 = Nothing
  Set frm2 = Nothing

  VBA_DupeFormFormat = blnRetValx

End Function

Public Function VBA_FixCode() As Boolean
' ** Currently not called.

  Const THIS_NAME As String = "VBA_FixCode"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strFind1 As String, strFind2 As String
  Dim strModName As String
  Dim lngLines As Long, lngCnt As Long
  Dim strLine As String, strNewLine As String
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim lngX As Long

  Const strMsg As String = "' ** To assure it's there and correct."
  Const strCmd As String = ".OpenReport"
  Const strVar1 As String = "strDocName ="
  Const strVar2 As String = "strReportName ="

  blnRetValx = True

  strFind1 = "Report_"
  strFind2 = ".Name"
  lngCnt = 0&

  ' ** Walk through every module.
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For Each vbc In .VBComponents
      With vbc
        strModName = .Name
        Set cod = .CodeModule
        With cod
          lngLines = .CountOfLines
          For lngX = 1& To lngLines
            strLine = Trim$(.Lines(lngX, 1))
            If strLine <> vbNullString Then
              If Left$(strLine, 1) <> "'" Then
                intPos1 = InStr(strLine, strFind1)
                intPos2 = InStr(strLine, strFind2)
                If intPos1 > 0 And intPos2 > 0 Then
                  If InStr(strLine, strCmd) > 0 Or InStr(strLine, strVar1) > 0 Then
                    lngCnt = lngCnt + 1&
                    If InStr(strLine, strMsg) = 0 Then
                      ' ** None!
                      'Debug.Print "'" & strModName & "  LN: " & CStr(lngX) & "  " & strLine
                    Else
                      If Trim$(Mid$(strLine, (intPos2 + Len(strFind2)))) = strMsg Then
                        ' ** Simple syntax.
                        strNewLine = Left$(strLine, (intPos1 - 1)) & Chr(34) & Mid$(strLine, (intPos1 + Len(strFind1)))
                        intPos2 = InStr(strNewLine, strFind2)
                        strNewLine = Left$(strNewLine, (intPos2 - 1)) & Chr(34)
                        'Debug.Print "'A: " & strNewLine
                        .ReplaceLine lngX, strNewLine
                      Else
                        ' ** Additional parameters.
                        strNewLine = Left$(strLine, (intPos1 - 1)) & Chr(34) & Mid$(strLine, (intPos1 + Len(strFind1)))
                        intPos2 = InStr(strNewLine, strFind2)
                        strNewLine = Left$(strNewLine, (intPos2 - 1)) & Chr(34) & Mid$(strNewLine, (intPos2 + Len(strFind2)))
                        intPos3 = InStr(strNewLine, strMsg)
                        strNewLine = Trim$(Left$(strNewLine, (intPos3 - 1)))
                        Debug.Print "'B: " & strNewLine
                        '.ReplaceLine lngX, strNewLine
                      End If
                    End If
                  Else
                    ' ** All OK.
                    'Debug.Print "'" & strModName & "  LN: " & CStr(lngX) & "  " & strLine
                  End If
                End If
              End If
            End If
          Next
        End With
      End With
    Next
  End With

  Debug.Print "'CNT: " & CStr(lngCnt)

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_FixCode = blnRetValx

End Function

Public Function VBA_Proc_ChkLineContinuation() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "VBA_Proc_ChkLineContinuation"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngProcs As Long, arr_varProc As Variant
  Dim strLine As String
  Dim lngX As Long

  Const P_CID     As Integer = 0
  Const P_CNAME   As Integer = 1
  Const P_ID      As Integer = 2
  Const P_NAME    As Integer = 3
  Const P_LINENUM As Integer = 4
  Const P_MULTI   As Integer = 5

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs
    Set qdf = .QueryDefs("zz_qry_VBComponent_Proc_10")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngProcs = .RecordCount
      .MoveFirst
      arr_varProc = .GetRows(lngProcs)
      ' *******************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   Field  Element  Name                 Constant
      ' **   =====  =======  ===================  ===========
      ' **     1       0     vbcom_id             P_CID
      ' **     2       1     vbcom_name           P_CNAME
      ' **     3       2     vbcomproc_id         P_ID
      ' **     4       3     vbcomproc_name       P_NAME
      ' **     5       4     vbcomproc_line_beg   P_LINENUM
      ' **     6       5     IsMulti              P_MULTI
      ' **
      ' *******************************************************
      .Close
    End With
    .Close
  End With

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For lngX = 0& To (lngProcs - 1&)
      Set vbc = .VBComponents(arr_varProc(P_CNAME, lngX))
      With vbc
        Set cod = .CodeModule
        With cod
          strLine = .Lines(arr_varProc(P_LINENUM, lngX), 1)
          If Right$(strLine, 1) = "_" Then
            arr_varProc(P_MULTI, lngX) = CBool(True)
            Debug.Print "'MOD: " & arr_varProc(P_CNAME, lngX) & "  PROC: " & arr_varProc(P_NAME, lngX)
          End If
        End With
      End With
    Next
  End With

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  VBA_Proc_ChkLineContinuation = blnRetValx

End Function

Public Function VBA_GetCode(varComID As Variant, varProcID As Variant, varLine As Variant, Optional varCodeNum As Variant, Optional varNext As Variant) As Variant
' ** Returns the line of VBA code specified by absolute line number.
' ** Optionally returns just code line number.
' ** Called by some queries.
'CodeLine_Get()

  Const THIS_PROC As String = "VBA_GetCode"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngComID As Long, lngProcID As Long, lngLine As Long
  Dim strModName As String, strProcName As String
  Dim lngLineBeg As Long, lngLineEnd As Long
  Dim lngLines As Long
  Dim blnCodeNum As Boolean, strCodeNum As String, blnNext As Boolean, strNext As String
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngX As Long
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varComID) = False And IsNull(varProcID) = False And IsNull(varLine) = False Then
    Select Case IsMissing(varCodeNum)
    Case True
      blnCodeNum = False
    Case False
      blnCodeNum = CBool(varCodeNum)
    End Select
    Select Case IsMissing(varNext)
    Case True
      blnNext = False
      strNext = vbNullString
    Case False
      If Trim(varNext) <> vbNullString Then
        blnNext = True
        strNext = varNext
      Else
        blnNext = False
        strNext = vbNullString
      End If
    End Select
    strCodeNum = vbNullString
    lngComID = CLng(varComID): lngProcID = CLng(varProcID): lngLine = CLng(varLine)
    varTmp00 = DLookup("[vbcom_name]", "tblVBComponent", "[vbcom_id] = " & CStr(lngComID))
    If IsNull(varTmp00) = False Then
      strModName = varTmp00
      ' ** Check that line is within specified procedure.
      varTmp00 = DLookup("[vbcomproc_line_beg]", "tblVBComponent_Procedure", "[vbcom_id] = " & CStr(lngComID) & " And " & _
        "[vbcomproc_id] = " & CStr(varProcID))
      If IsNull(varTmp00) = False Then
        lngLineBeg = varTmp00
        varTmp00 = DLookup("[vbcomproc_line_end]", "tblVBComponent_Procedure", "[vbcom_id] = " & CStr(lngComID) & " And " & _
          "[vbcomproc_id] = " & CStr(varProcID))
        If IsNull(varTmp00) = False Then
          lngLineEnd = varTmp00
          If lngLine >= lngLineBeg And lngLine <= lngLineEnd Then
            Set vbp = Application.VBE.ActiveVBProject
            With vbp
              Set vbc = .VBComponents(strModName)
              With vbc
                Set cod = .CodeModule
                With cod
                  lngLines = .CountOfLines
                  If lngLine <= lngLines Then
                    Select Case blnCodeNum
                    Case True
                      strTmp01 = Trim(.Lines(lngLine, 1))
                      If strTmp01 <> vbNullString Then
                        intPos1 = InStr(strTmp01, " ")
                        If intPos1 > 0 Then
                          strTmp01 = Trim$(Left$(strTmp01, intPos1))
                          If IsNumeric(strTmp01) = True Then
                            strCodeNum = strTmp01
                          End If
                        End If
                      End If
                      If strCodeNum = vbNullString Then
                        If blnNext = True Then
                          Select Case strNext
                          Case "Next"
                            For lngX = lngLine To lngLines
                              strTmp01 = Trim(.Lines(lngX, 1))
                              If strTmp01 <> vbNullString Then
                                intPos1 = InStr(strTmp01, " ")
                                If intPos1 > 0 Then
                                  strTmp01 = Trim$(Left$(strTmp01, intPos1))
                                  If IsNumeric(strTmp01) = True Then
                                    strCodeNum = strTmp01
                                    Exit For
                                  End If
                                End If
                              End If
                            Next
                          Case "Prev"
                            For lngX = lngLine To .CountOfDeclarationLines Step -1&
                              strTmp01 = Trim(.Lines(lngX, 1))
                              If strTmp01 <> vbNullString Then
                                intPos1 = InStr(strTmp01, " ")
                                If intPos1 > 0 Then
                                  strTmp01 = Trim$(Left$(strTmp01, intPos1))
                                  If IsNumeric(strTmp01) = True Then
                                    strCodeNum = strTmp01
                                    Exit For
                                  End If
                                End If
                              End If
                            Next
                          End Select
                        End If
                      End If
                      If strCodeNum <> vbNullString Then
                        varRetVal = strCodeNum
                      End If
                    Case False
                      varRetVal = .Lines(lngLine, 1)
                    End Select
                  End If
                End With  ' ** cod.
              End With  ' ** vbc.
            End With  ' ** vbp.
          End If  ' ** lngLineBeg -> lngLine <- lngLineEnd.
        End If  ' ** lngProcID, vbcomproc_line_end.
      End If  ' ** lngProcID, vbcomproc_line_beg.
    End If  ' ** lngComID.
  End If  ' ** IsNull().

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_GetCode = varRetVal

End Function

Public Function VBA_HasCodeLines(varComID As Variant) As Boolean

  Const THIS_PROC As String = "VBA_HasCodeLines"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strModName As String, strLine As String
  Dim lngLines As Long, lngDecLines As Long
  Dim lngLinesChecked As Long, lngHits As Long
  Dim lngThisDbsID As Long
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = False

  If IsNull(varComID) = False Then
    lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
    varTmp00 = DLookup("[dbs_id]", "tblVBComponent", "[vbcom_id] = " & CStr(varComID))
    If IsNull(varTmp00) = False Then  ' ** Make sure it exists.
      If varTmp00 = lngThisDbsID Then  ' ** Make sure it's in this database.
        varTmp00 = DLookup("[vbcom_name]", "tblVBComponent", "[vbcom_id] = " & CStr(varComID))
        If IsNull(varTmp00) = False Then
          If ModExists(varTmp00) = True Then  ' ** Function: Below.
            strModName = varTmp00
            Set vbp = Application.VBE.ActiveVBProject
            With vbp
              Set vbc = .VBComponents(strModName)
              With vbc
                Set cod = .CodeModule
                With cod
                  lngLines = .CountOfLines
                  lngDecLines = .CountOfDeclarationLines
                  lngLinesChecked = 0&: lngHits = 0&
                  For lngX = lngDecLines To lngLines
                    strLine = .Lines(lngX, 1)
                    strLine = Trim$(strLine)
                    If strLine <> vbNullString Then
                      If Left$(strLine, 1) <> "'" Then
                        lngLinesChecked = lngLinesChecked + 1&
                        intPos1 = InStr(strLine, " ")
                        If intPos1 > 0 Then
                          strTmp01 = Trim$(Left$(strLine, intPos1))
                          If IsNumeric(strTmp01) = True Then
                            lngHits = lngHits + 1&
                          End If
                        End If
                      End If
                    End If
                    If lngHits >= 5 Then
                      blnRetVal = True
                      Exit For
                    ElseIf lngHits = 0& And lngLinesChecked >= 100& Then
                      Exit For
                    End If
                  Next
                End With  ' ** cod.
              End With  ' ** vbc.
            End With  ' ** vbp.
          End If
        End If
      End If
    End If
  End If

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_HasCodeLines = blnRetVal

End Function

Public Function VBA_GetVar(varComID As Variant, varProcID As Variant, varLine As Variant, varVariable As Variant) As Variant
' ** Given a variable, find its assignment prior to the line where it's used.

  Const THIS_PROC As String = "VBA_GetVar"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngComID As Long, lngProcID As Long, lngLine As Long
  Dim strModName As String, strProcName As String
  Dim lngLineBeg As Long, lngLineEnd As Long
  Dim lngLines As Long
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngX As Long
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varComID) = False And IsNull(varProcID) = False And IsNull(varLine) = False And IsNull(varVariable) = False Then
    lngComID = CLng(varComID): lngProcID = CLng(varProcID): lngLine = CLng(varLine)
    varTmp00 = DLookup("[vbcom_name]", "tblVBComponent", "[vbcom_id] = " & CStr(lngComID))
    If IsNull(varTmp00) = False Then
      strModName = varTmp00
      ' ** Check that line is within specified procedure.
      varTmp00 = DLookup("[vbcomproc_line_beg]", "tblVBComponent_Procedure", "[vbcom_id] = " & CStr(lngComID) & " And " & _
        "[vbcomproc_id] = " & CStr(varProcID))
      If IsNull(varTmp00) = False Then
        lngLineBeg = varTmp00
        varTmp00 = DLookup("[vbcomproc_line_end]", "tblVBComponent_Procedure", "[vbcom_id] = " & CStr(lngComID) & " And " & _
          "[vbcomproc_id] = " & CStr(varProcID))
        If IsNull(varTmp00) = False Then
          lngLineEnd = varTmp00
          If lngLine >= lngLineBeg And lngLine <= lngLineEnd Then
            Set vbp = Application.VBE.ActiveVBProject
            With vbp
              Set vbc = .VBComponents(strModName)
              With vbc
                Set cod = .CodeModule
                With cod
                  lngLines = .CountOfLines
                  If lngLine <= lngLines Then
                    For lngX = (lngLine - 1&) To lngLineBeg Step -1&
                      strTmp01 = Trim(.Lines(lngX, 1))
                      If strTmp01 <> vbNullString Then
                        If Left$(strTmp01, 1) <> "'" Then
                          intPos1 = InStr(strTmp01, varVariable)
                          If intPos1 > 0 Then
                            strTmp01 = Mid$(strTmp01, intPos1)
                            intPos1 = InStr(strTmp01, Chr(34))
                            If intPos1 > 0 Then
                              strTmp01 = Mid$(strTmp01, (intPos1 + 1))
                              intPos1 = InStr(strTmp01, Chr(34))
                              If intPos1 > 0 Then
                                strTmp01 = Left$(strTmp01, (intPos1 - 1))
                                varRetVal = strTmp01
                                Exit For
                              End If
                            End If
                          End If
                        End If
                      End If
                    Next
                  End If
                End With  ' ** cod.
              End With  ' ** vbc.
            End With  ' ** vbp.
          End If  ' ** lngLineBeg -> lngLine <- lngLineEnd.
        End If  ' ** lngProcID, vbcomproc_line_end.
      End If  ' ** lngProcID, vbcomproc_line_beg.
    End If  ' ** lngComID.
  End If  ' ** IsNull().

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  VBA_GetVar = varRetVal

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
' **   zz_mod_QuerySQLDocFuncs (This):
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
    If TableExists("USysRibbons") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "USysRibbons", acTable, "zz_USysRibbons"
    End If
    If TableExists("LedgerArchive_Backup") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "LedgerArchive_Backup", acTable, "tblTemplate_LedgerArchive"
    End If
    If TableExists("zz_tbl_m_TBL_tmp01") = False Then  ' ** Module Function: modFileUtilities.
      DoCmd.CopyObject , "zz_tbl_m_TBL_tmp01", acTable, "tblTemplate_m_TBL"
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
      RePost_TmpDB_Link  ' ** Module Function: modRePostFuncs.
      RePost_TmpDB_Link_RT True  ' ** Module Function: modRePostFuncs.
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
    TableDelete "zz_tbl_m_TBL_tmp01"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp01"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp02"  ' ** Module Function: modFileUtilities.
    TableDelete "tblDatabase_Table_Link_tmp03"  ' ** Module Function: modFileUtilities.

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    If blnJrnlTmp = True Then
      RePost_TmpDB_Unlink  ' ** Module Function: modRePostFuncs.
      RePost_TmpDB_Link_RT False  ' ** Module Function: modRePostFuncs.
    End If

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

  End Select

  Set qdf0 = Nothing
  Set dbs0 = Nothing

  Qry_TmpTables = blnRetValx

End Function

Public Function Qry_UpdateRef(Optional varOldRef As Variant, Optional varNewRef As Variant, Optional varCase As Variant) As Boolean
' ** Change all references from one source to another.

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_UpdateRef"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strFind As String, strFind2 As String
        Dim lngQrys As Long
        Dim strSQL As String
        Dim blnCase As Boolean, blnContinue As Boolean, blnCalled As Boolean
        Dim lngPos1 As Long, lngLen As Long
        Dim intX As Integer
        Dim blnRetVal As Boolean

15310   blnRetVal = True

15320   If IsMissing(varOldRef) = True Then
15330     strFind = "tblVBComponent_Procedure_Detail2"
15340     strFind2 = "tblVBComponent_Procedure_Detail"
15350     blnCase = False  ' ** Match case.
15351     blnCalled = False
15360   Else
15370     strFind = CStr(varOldRef)
15380     strFind2 = CStr(varNewRef)
15390     If IsMissing(varCase) = True Then
15400       blnCase = False
15410     Else
15420       blnCase = CBool(varCase)
15430     End If
15431     blnCalled = True
15440   End If

15450   Set dbs = CurrentDb
15460   With dbs
15470     lngQrys = 0&
15480     For Each qdf In .QueryDefs
15490       With qdf
15500         If Left$(.Name, 1) <> "~" Then  ' ** Skip those pesky system queries.
15510           strSQL = .SQL
15520           If blnCase = False Then
15530             lngPos1 = InStr(strSQL, strFind)
15540           Else
15550             lngPos1 = InStr(strSQL, strFind)  'ERRORS WITH Type Mismatch! WHY? : , vbTextCompare)
15560           End If
15570           If lngPos1 > 0 Then
15580             lngQrys = lngQrys + 1&
15590             Do While lngPos1 > 0
15600               strSQL = Left$(strSQL, (lngPos1 - 1)) & strFind2 & Mid$(strSQL, (lngPos1 + Len(strFind)))
15610               lngPos1 = InStr((lngPos1 + 1), strSQL, strFind)
                    ' ** VbCompare enumeration.
                    ' **    0  vbBinaryCompare     Performs a binary comparison.
                    ' **    1  vbTextCompare       Performs a textual comparison.
                    ' **    2  vbDatabaseCompare   Microsoft Access only. Performs a comparison based on information in your database.
                    ' **    3  vbUseCompareOption  Performs a comparison using the setting of the Option Compare statement. (Stated value, -1, is wrong!)
15620               If blnCase = True And lngPos1 > 0 And strFind = strFind2 Then
15630                 lngLen = Len(strFind): blnContinue = False
15640                 For intX = 1 To lngLen
15650                   If Asc(Mid$(strFind, intX, 1)) <> Asc(Mid$(strSQL, ((lngPos1 + intX) - 1), 1)) Then
                          ' ** If they're not equal, then continue the loop
15660                     blnContinue = True
15670                     Exit For
15680                   End If
15690                 Next
15700                 If blnContinue = False Then
                        ' ** All characters were identical, so this isn't a match.
15710                   lngPos1 = 0
15720                 End If
15730               End If
15740             Loop
15750             .SQL = strSQL
15760           End If
15770         End If
15780       End With
15790     Next
15800     .Close
15810   End With

15820   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

15821   If blnCalled = False Then
15830     Debug.Print "'QRYS CHANGED: " & CStr(lngQrys)
15831   End If

15840   Beep

EXITP:
15850   Set qdf = Nothing
15860   Set dbs = Nothing
15870   Qry_UpdateRef = blnRetVal
15880   Exit Function

ERRH:
15890   blnRetVal = False
15900   Select Case ERR.Number
        Case Else
15910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15920   End Select
15930   Resume EXITP

End Function

Public Function Find_QryStr(Optional varRenFind1 As Variant) As Variant
' ** Find a specified string within a query's SQL.

  Const THIS_PROC As String = "Find_QryStr"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, prp As Object
  Dim lngRefs As Long, arr_varRef() As Variant
  Dim strFind1 As String, lngFinds As Long, lngFindSpaces As Long
  Dim intIteration As Integer
  Dim strSQL As String
  Dim blnFound As Boolean, blnExact As Boolean
  Dim blnCalled As Boolean, blnFormRef As Boolean, blnToTable As Boolean
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngE As Long
  Dim varRetVal As Variant

  Const RF_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const RF_QRY As Integer = 0
  Const RF_SQL As Integer = 1
  Const RF_TYP As Integer = 2
  Const RF_DSC As Integer = 3

  If IsMissing(varRenFind1) = True Then
    strFind1 = "VBA_GetVar"
    blnExact = False
    blnFormRef = False  ' ** When True, trims right side to just return FormRef() piece.
    blnCalled = False
    varRetVal = False
    blnToTable = False  ' ** When True, puts results into a table.
    intIteration = 10
  Else
    strFind1 = varRenFind1
    blnExact = False
    blnFormRef = False
    blnCalled = True
    blnToTable = False
  End If

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

  Set dbs = CurrentDb

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  lngRefs = 0&
  ReDim arr_varRef(RF_ELEMS, 0)
  arr_varRef(RF_QRY, 0) = "0"

  With dbs
    For Each qdf In .QueryDefs
      blnFound = False: lngFindSpaces = 0&
      With qdf
        strSQL = .SQL
        If Left$(.Name, 1) <> "~" Then
          intPos1 = InStr(strSQL, vbCrLf)
          Do While intPos1 > 0
            strSQL = Left$(strSQL, (intPos1 - 1)) & " " & Mid$(strSQL, (intPos1 + 2))
            intPos1 = InStr(strSQL, vbCrLf)
          Loop
          intPos1 = InStr(strSQL, "  ")
          Do While intPos1 > 0
            strSQL = Left$(strSQL, intPos1) & Mid$(strSQL, (intPos1 + 2))
            intPos1 = InStr(strSQL, "  ")
          Loop
          intPos1 = InStr(strSQL, strFind1)
          If intPos1 > 0 Then
            blnFound = True
            strTmp01 = Mid$(strSQL, intPos1)
            intPos2 = InStr(strFind1, " ")  ' ** See if the Find has a space.
            If intPos2 > 0 Then
              If InStr((intPos2 + 1), strFind1, " ") = 0 Then  ' ** Only 1 space in Find.
                lngFindSpaces = 1&
                intPos2 = InStr((intPos2 + 1), strSQL, " ")  ' ** 2nd space.
              Else  ' ** At least 2 spaces in Find.
                lngFindSpaces = 2&
                intPos2 = InStr((intPos2 + 1), strSQL, " ")  ' ** 2nd space.
                intPos2 = InStr((intPos2 + 1), strSQL, " ")  ' ** 3rd space.
              End If
            Else
              lngFindSpaces = 0&
              intPos2 = InStr((intPos1 + 1), strSQL, " ")
            End If
            If intPos2 > 0 Then
              Select Case lngFindSpaces
              Case 0&
                strTmp02 = Trim$(Left$(strTmp01, intPos2))
              Case 1&
                strTmp02 = Trim$(Left$(strTmp01, intPos2))
              Case 2&
                strTmp02 = Trim$(Left$(strTmp01, intPos2))
              End Select
            Else
              strTmp02 = Mid$(strSQL, intPos1)
            End If
            If blnFormRef = False Then
              intPos2 = InStr(strTmp02, "]")
              If intPos2 > 0 Then strTmp02 = Left$(strTmp02, (intPos2 - 1))
            End If
            intPos2 = InStr(strTmp02, "!")
            If intPos2 > 0 Then strTmp02 = Left$(strTmp02, (intPos2 - 1))
            intPos2 = InStr(strTmp02, ".")
            If intPos2 > 0 Then strTmp02 = Left$(strTmp02, (intPos2 - 1))
            intPos2 = InStr(strTmp02, ",")
            If intPos2 > 0 Then strTmp02 = Left$(strTmp02, (intPos2 - 1))
            intPos2 = InStr(strTmp02, "+")
            If intPos2 > 0 Then strTmp02 = Left$(strTmp02, (intPos2 - 1))
            intPos2 = InStr(strTmp02, "/")
            If intPos2 > 0 Then strTmp02 = Left$(strTmp02, (intPos2 - 1))
            If blnCalled = False Then
              Select Case blnFormRef
              Case True
                'intPos1 = InStr(strTmp02, "*")
                'If intPos1 > 0 Then
                '  strTmp02 = Left$(strTmp02, (intPos1 - 1))
                'End If
                'intPos1 = InStr(strTmp02, "))")
                'Do While intPos1 > 0
                '  strTmp02 = Left$(strTmp02, intPos1)
                '  intPos1 = InStr(strTmp02, "))")
                'Loop
                intPos1 = InStr(strTmp02, ")")
                If intPos1 > 0 Then
                  strTmp02 = Left$(strTmp02, intPos1)
                End If
                intPos1 = InStr(strTmp02, "(")
                If intPos1 > 0 Then
                  lngFinds = lngFinds + 1&
                  If blnToTable = False Then
                    Debug.Print "'QRY: " & Left$(qdf.Name & Space(34), 34) & " REF: " & strTmp02
                  End If
                  varRetVal = True
                End If
              Case False
                Select Case blnExact
                Case True
                  intPos1 = InStr(strTmp02, " ")
                  If intPos1 > 0 And InStr(strFind1, " ") = 0 Then
                    strTmp02 = Left$(strTmp02, (intPos1 - 1))
                  End If
                  If Len(strTmp02) > Len(strFind1) Then
                    strTmp03 = Mid$(strTmp02, (Len(strFind1) + 1), 1)
                    If ((Asc(strTmp03) >= 65) And (Asc(strTmp03) <= 90)) Or ((Asc(strTmp03) >= 97) And (Asc(strTmp03) <= 122)) Or _
                        ((Asc(strTmp03) >= 48) And (Asc(strTmp03) <= 57)) Or strTmp03 = "_" Then
                      blnFound = False
                    Else

                    End If
                  End If
                  If blnFound = True Then
                    lngFinds = lngFinds + 1&
                    If blnToTable = False Then
                      Debug.Print "'QRY: '" & qdf.Name & "' " & strTmp02 & "  TYP: " & Qry_Type(qdf.Type)
                    End If
                    varRetVal = True
                  End If
                Case False
                  lngFinds = lngFinds + 1&
                  If blnToTable = False Then
                    Debug.Print "'QRY: '" & qdf.Name & "' " & strTmp02 & "  TYP: " & Qry_Type(qdf.Type)
                  End If
                  varRetVal = True
                End Select
              End Select
            End If
          End If
          If blnFound = True Then
            lngRefs = lngRefs + 1&
            lngE = lngRefs - 1&
            ReDim Preserve arr_varRef(RF_ELEMS, lngE)
            arr_varRef(RF_QRY, lngE) = qdf.Name
            arr_varRef(RF_SQL, lngE) = qdf.SQL
            If blnToTable = True Then
              arr_varRef(RF_TYP, lngE) = qdf.Type
              arr_varRef(RF_DSC, lngE) = vbNullString
              For Each prp In qdf.Properties
                With prp
                  If .Name = "Description" Then
                    If IsNull(.Value) = False Then
                      If Trim(.Value) <> vbNullString Then
                        arr_varRef(RF_DSC, lngE) = .Value
                      End If
                    End If
                  End If
                End With
              Next
              Set prp = Nothing
            End If
          End If
        End If
      End With
    Next

    If blnToTable = True Then
      Set rst = .OpenRecordset("zz_tbl_RecurringItems_01", dbOpenDynaset, dbAppendOnly)
      With rst
        For lngX = 0& To (lngRefs - 1&)
          .AddNew
          ' ** ![qry_id] =
          ![qryx_name] = arr_varRef(RF_QRY, lngX)
          ![qrytype_type] = arr_varRef(RF_TYP, lngX)
          If arr_varRef(RF_DSC, lngX) <> vbNullString Then
            ![qryx_description] = arr_varRef(RF_DSC, lngX)
          End If
          ![qryx_run] = intIteration
          ![qryx_sql] = arr_varRef(RF_SQL, lngX)
          ![qryx_datemodified] = Now()
          .Update
        Next
        .Close
      End With
    End If

    .Close
  End With

  If blnCalled = True Then
    varRetVal = arr_varRef
  End If

  Set prp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  If blnCalled = False Then
    Beep
    If lngFinds > 0& Then
      Debug.Print "'DONE!  QRYS: " & CStr(lngFinds) & "  " & THIS_PROC & "()"
    Else
      Debug.Print "'NONE FOUND!  " & THIS_PROC & "()"
    End If
  End If

  Find_QryStr = varRetVal

End Function

Public Function ModExists(varInput As Variant) As Boolean

  Const THIS_PROC As String = "ModExists"

  Dim vbp As VBProject, vbc As VBComponent
  Dim blnRetVal As Boolean

  blnRetVal = False

  If IsNull(varInput) = False Then
    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For Each vbc In .VBComponents
        With vbc
          If .Name = varInput Then
            blnRetVal = True
            Exit For
          End If
        End With  ' ** vbc.
        Set vbc = Nothing
      Next  ' ** vbc.
    End With  ' ** vbp.
  End If

  Set vbc = Nothing
  Set vbp = Nothing

  ModExists = blnRetVal

End Function
