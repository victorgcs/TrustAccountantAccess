Attribute VB_Name = "zz_mod_ReportDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_ReportDocFuncs"

'VGC 03/26/2017: CHANGES!

' ** AcNewRowOrCol enumeration: (my own)  ' ** Not currently used!
Public Const acNewRowOrColNone   As Integer = 0
Public Const acNewRowOrColBefore As Integer = 1
Public Const acNewRowOrColAfter  As Integer = 2
Public Const acNewRowOrColBoth   As Integer = 3

Private blnRetValx As Boolean
' **

Public Function QuikRptDoc() As Boolean
  Const THIS_PROC As String = "QuikRptDoc"
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Rpt_ChkDocQrys(False) = True Then  ' ** Function: Below.
      blnRetValx = Rpt_Doc  ' ** Function: Below.
      blnRetValx = Rpt_Sec_Doc
      blnRetValx = Rpt_Ctl_Doc  ' ** Function: Below.
      'DoCmd.Close acForm, Forms(0).Name
      blnRetValx = Rpt_RecSrc_Doc  ' ** Function: Below.
      blnRetValx = Rpt_Subreport_Doc
      blnRetValx = Rpt_Grp_Doc
      DoEvents
      DoBeeps  ' ** Module Function: modWindowFunctions.
      Debug.Print "'FINISHED!"
'DON'T FORGET Rpt_ReportList_Doc(), BELOW!!!!!!!!!!!!!
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED Rpt_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikRptDoc = blnRetValx
End Function

Private Function Rpt_Doc() As Boolean
' ** Document all Reports in Trust Accountant to tblReport.
' ** Called by:
' **   QuikRptDoc(), Above

  Const THIS_PROC As String = "Rpt_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, rpt As Access.Report, ctl As Access.Control
  Dim prj As CurrentProject, rptao As AccessObject
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngProps As Long
  Dim strName As String, strSubs As String
  Dim lngSubs As Long
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim blnFound As Boolean, blnDelete As Boolean, blnAdd As Boolean
  Dim intPos1 As Integer
  Dim strTmp00 As String
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 15  ' ** Array's first-element UBound().
  Const R_DID     As Integer = 0
  Const R_DNAM    As Integer = 1
  Const R_RID     As Integer = 2
  Const R_OTYP    As Integer = 3
  Const R_RNAM    As Integer = 4
  Const R_FNAM    As Integer = 5
  Const R_CTLS    As Integer = 6
  Const R_CAP     As Integer = 7
  Const R_HAS_SUB As Integer = 8
  Const R_IS_SUB  As Integer = 9
  Const R_ACT     As Integer = 10
  Const R_HID     As Integer = 11
  Const R_DSC     As Integer = 12
  Const R_TAG     As Integer = 13
  Const R_PARSUB  As Integer = 14
  Const R_SUBS    As Integer = 15

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  If Reports.Count > 0 Then
    Do While Reports.Count > 0
      DoCmd.Close acReport, Reports(0).Name, acSaveNo
    Loop
  End If

  Set prj = Application.CurrentProject
  With prj
    lngRpts = .AllReports.Count
    lngX = -1&
    ReDim arr_varRpt(R_ELEMS, (lngRpts - 1&))
    For Each rptao In .AllReports
      strName = vbNullString
      With rptao
        lngX = lngX + 1&
        strName = .Name
        ' *********************************************************
        ' ** Array: arr_varRpt()
        ' **
        ' **   Field  Element  Name                   Constant
        ' **   =====  =======  =====================  ===========
        ' **     1       0     dbs_id                 R_DID
        ' **     2       1     dbs_name               R_DNAM
        ' **     3       2     rpt_id                 R_RID
        ' **     4       3     objtype_type           R_OTYP
        ' **     5       4     rpt_name               R_RNAM
        ' **     6       5     rpt_fullname           R_FNAM
        ' **     7       6     rpt_controls           R_CTLS
        ' **     8       7     rpt_caption/Caption    R_CAP
        ' **     9       8     rpt_hassub             R_HAS_SUB
        ' **    10       9     rpt_issub              R_IS_SUB
        ' **    11      10     rpt_active             R_ACT
        ' **    12      11     sec_hidden             R_HID
        ' **    13      12     rpt_description        R_DSC
        ' **    14      13     rpt_tag/Tag            R_TAG
        ' **    15      14     rpt_parent_sub         R_PARSUB
        ' **    16      15     rpt_subs               R_SUBS
        ' **
        ' *********************************************************
        arr_varRpt(R_DID, lngX) = lngThisDbsID
        arr_varRpt(R_DNAM, lngX) = CurrentAppName  ' ** Module Function: modFileUtilities.
        arr_varRpt(R_RID, lngX) = CLng(0)
        arr_varRpt(R_OTYP, lngX) = .Type
        arr_varRpt(R_RNAM, lngX) = strName
        arr_varRpt(R_FNAM, lngX) = .FullName  ' ** Not saved.
        arr_varRpt(R_CTLS, lngX) = CLng(0)
        arr_varRpt(R_CAP, lngX) = vbNullString
        arr_varRpt(R_HAS_SUB, lngX) = CBool(False)
        If InStr(strName, "_Sub") > 0 Then
          arr_varRpt(R_IS_SUB, lngX) = CBool(True)
        Else
          arr_varRpt(R_IS_SUB, lngX) = CBool(False)
        End If
        If Left$(strName, 3) = "zz_" Then
          arr_varRpt(R_ACT, lngX) = CBool(False)
        Else
          arr_varRpt(R_ACT, lngX) = CBool(True)
        End If
        arr_varRpt(R_HID, lngX) = CBool(False)
        arr_varRpt(R_DSC, lngX) = vbNullString
        lngProps = CurrentDb.Containers("Reports").Documents(strName).Properties.Count
        For lngY = 0& To (lngProps - 1&)
          If CurrentDb.Containers("Reports").Documents(strName).Properties(lngY).Name = "Description" Then
            arr_varRpt(R_DSC, lngX) = CurrentDb.Containers("Reports").Documents(strName).Properties(lngY).Value
            Exit For
          End If
        Next
        arr_varRpt(R_TAG, lngX) = vbNullString
        arr_varRpt(R_PARSUB, lngX) = vbNullString
        arr_varRpt(R_SUBS, lngX) = CLng(0)
      End With
    Next
  End With

  For lngX = 0& To (lngRpts - 1&)
    strName = arr_varRpt(R_RNAM, lngX)
    lngSubs = 0&
    DoCmd.OpenReport strName, acViewDesign, , , acHidden
    Set rpt = Reports(strName)
    With rpt
      arr_varRpt(R_CTLS, lngX) = .Controls.Count
      arr_varRpt(R_CAP, lngX) = vbNullString
      If IsNull(.Caption) = False Then
        If Trim(.Caption) <> vbNullString Then
          arr_varRpt(R_CAP, lngX) = .Caption
        End If
      End If
      If IsNull(.Tag) = False Then
        If Trim(.Tag) <> vbNullString Then
          arr_varRpt(R_TAG, lngX) = .Tag
        End If
      End If
          'If InStr(arr_varRpt(R_TAG, lngX), "Is Subreport") > 0 Then
      For Each ctl In .Controls
        With ctl
          Select Case .ControlType
          Case acSubform
            lngSubs = lngSubs + 1&
          Case Else
            ' ** Don't care.
          End Select
        End With
      Next
      If lngSubs > 0& Then
        arr_varRpt(R_HAS_SUB, lngX) = CBool(True)
      End If
      arr_varRpt(R_SUBS, lngX) = lngSubs
    End With
    DoCmd.Close acReport, strName, acSaveNo
    Set ctl = Nothing
    Set rpt = Nothing
  Next

  Set dbs = CurrentDb
  With dbs

    Set rst = .OpenRecordset("tblReport", dbOpenDynaset, dbConsistent)
    With rst
      For lngX = 0& To (lngRpts - 1&)
        blnAdd = False
        .FindFirst "[dbs_id] = " & arr_varRpt(R_DID, lngX) & " And [rpt_name] = '" & arr_varRpt(R_RNAM, lngX) & "'"
'rpt_id
        If .NoMatch = True Then
          blnAdd = True
          .AddNew
          ![dbs_id] = arr_varRpt(R_DID, lngX)
        Else
          arr_varRpt(R_RID, lngX) = ![rpt_id]
          .Edit
        End If
        If blnAdd = True Then
'rpt_name
          ![rpt_name] = arr_varRpt(R_RNAM, lngX)
'objtype_type
          ![objtype_type] = arr_varRpt(R_OTYP, lngX)
        End If
'rpt_controls
        ![rpt_controls] = arr_varRpt(R_CTLS, lngX)
'rpt_caption
        If arr_varRpt(R_CAP, lngX) <> vbNullString Then
          ![rpt_caption] = arr_varRpt(R_CAP, lngX)
        Else
          If IsNull(![rpt_caption]) = False Then
            ![rpt_caption] = Null
          End If
        End If
'rpt_hassub
        ![rpt_hassub] = arr_varRpt(R_HAS_SUB, lngX)
'rpt_issub
        ![rpt_issub] = arr_varRpt(R_IS_SUB, lngX)
'rpt_active
        ![rpt_active] = arr_varRpt(R_ACT, lngX)
'sec_hidden
        If blnAdd = True Then
          ![sec_hidden] = arr_varRpt(R_HID, lngX)
        Else
          ' ** Leave it as stands.
        End If
'rpt_description
        If IsNull(![rpt_description]) = True And arr_varRpt(R_DSC, lngX) = vbNullString Then
          ' ** Do nothing.
        Else
          If IsNull(![rpt_description]) = True And arr_varRpt(R_DSC, lngX) <> vbNullString Then
            ![rpt_description] = arr_varRpt(R_DSC, lngX)
          Else
            If IsNull(![rpt_description]) = False And arr_varRpt(R_DSC, lngX) = vbNullString Then
              ![rpt_description] = Null
            Else
              If ![rpt_description] <> arr_varRpt(R_DSC, lngX) Then
                ![rpt_description] = arr_varRpt(R_DSC, lngX)
              End If
            End If
          End If
        End If
'rpt_tag
        If arr_varRpt(R_TAG, lngX) <> vbNullString Then
          ![rpt_tag] = arr_varRpt(R_TAG, lngX)
        Else
          If IsNull(![rpt_tag]) = False Then
            ![rpt_tag] = Null
          End If
        End If
'rpt_parent_sub
        ![rpt_parent_sub] = Null  ' ** Set below.
'rpt_subs
        ![rpt_subs] = arr_varRpt(R_SUBS, lngX)
'rpt_datemodified
        ![rpt_datemodified] = Now()
        .Update
        If blnAdd = True Then
          .Bookmark = .LastModified
          arr_varRpt(R_RID, lngX) = ![rpt_id]
        End If
      Next
    End With  ' ** rst.

    For lngX = 0& To (lngRpts - 1&)

      strSubs = vbNullString
      strName = arr_varRpt(R_RNAM, lngX)
      DoCmd.OpenReport strName, acViewDesign, , , acHidden
      Set rpt = Reports(strName)
      If arr_varRpt(R_SUBS, lngX) > 0& Then
        With rpt
          For Each ctl In .Controls
            With ctl
              Select Case .ControlType
              Case acSubform
                If .SourceObject <> vbNullString Then
                  strTmp00 = .SourceObject
                  intPos1 = InStr(strTmp00, ".")
                  If intPos1 > 0 Then
                    strTmp00 = Mid$(strTmp00, (intPos1 + 1))
                  End If
                  strSubs = strSubs & strTmp00 & ";"
                Else
                  strSubs = strSubs & "{empty};"
                End If
              Case Else
                ' ** Don't care.
              End Select
            End With
          Next
        End With
      End If  ' ** R_SUBS.

      blnRetValx = Rpt_Specs_Doc(rpt, dbs, CLng(arr_varRpt(R_RID, lngX)), lngThisDbsID)  ' ** Function: Below.

      DoCmd.Close acReport, strName, acSaveNo
      Set ctl = Nothing
      Set rpt = Nothing

      If strSubs <> vbNullString Then
        With rst
          .FindFirst "[dbs_id] = " & CStr(arr_varRpt(R_DID, lngX)) & " And [rpt_id] = " & CStr(arr_varRpt(R_RID, lngX))
          If .NoMatch = False Then
            If Right$(strSubs, 1) = ";" Then strSubs = Left$(strSubs, (Len(strSubs) - 1))
            .Edit
            ![rpt_parent_sub] = strSubs
            ![rpt_datemodified] = Now()
            .Update
          Else
            Stop
          End If
          intPos1 = InStr(strSubs, ";")
          If intPos1 = 0 Then
            strName = strSubs
            strSubs = vbNullString
          Else
            strName = Left$(strSubs, (intPos1 - 1))
            strSubs = Mid$(strSubs, (intPos1 + 1))
          End If
          Do While strName <> vbNullString
            If strName <> "{empty}" Then
              .MoveFirst
              .FindFirst "[dbs_id] = " & CStr(arr_varRpt(R_DID, lngX)) & " And [rpt_name] = '" & strName & "'"
              If .NoMatch = False Then
                .Edit
                ![rpt_parent_sub] = arr_varRpt(R_RNAM, lngX)
                ![rpt_datemodified] = Now()
                .Update
              Else
                Stop
              End If
            End If
            If strSubs <> vbNullString Then
              intPos1 = InStr(strSubs, ";")
              If intPos1 = 0 Then
                strName = strSubs
                strSubs = vbNullString
              Else
                strName = Left$(strSubs, (intPos1 - 1))
                strSubs = Mid$(strSubs, (intPos1 + 1))
              End If
            Else
              Exit Do
            End If
          Loop  ' ** strName.
        End With  ' ** rst.
      End If

    Next  ' ** lngx.
    rst.Close
    Set rst = Nothing

    lngDels = 0&
    ReDim arr_varDel(0)

    Set rst = .OpenRecordset("tblReport", dbOpenDynaset, dbConsistent)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          blnFound = False
          For lngY = 0& To (lngRpts - 1&)
            If arr_varRpt(R_RNAM, lngY) = ![rpt_name] Then
              blnFound = True
              Exit For
            End If
          Next
          If blnFound = False Then
            lngDels = lngDels + 1&
            lngE = lngDels - 1&
            ReDim Preserve arr_varDel(lngE)
            arr_varDel(lngE) = ![rpt_id]
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With  ' ** rst.

    If lngDels > 0& Then
      For lngX = 0& To (lngDels - 1&)
        blnDelete = True
        strTmp00 = DLookup("[rpt_name]", "tblReport", "[rpt_id] = " & CStr(arr_varDel(lngX)))
        Debug.Print "'DEL RPT? " & strTmp00
Stop
        If blnDelete = True Then
          ' ** Delete tblReport, by specified [rpid].
          Set qdf = .QueryDefs("zz_qry_Report_01")
          With qdf.Parameters
            ![rpid] = arr_varDel(lngX)
          End With
          qdf.Execute
          Set qdf = Nothing
        End If
      Next
    End If

    .Close
  End With  ' ** dbs.

  ' ** AcObjectType enumeration:
  ' **   -1  acDefault
  ' **    0  acTable
  ' **    1  acQuery
  ' **    2  acForm
  ' **    3  acReport
  ' **    4  acMacro
  ' **    5  acModule
  ' **    6  acDataAccessPage
  ' **    7  acServerView
  ' **    8  acDiagram
  ' **    9  acStoredProcedure
  ' **   10  acFunction
  ' ** Same as standard ObjectType.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  Set ctl = Nothing
  Set rpt = Nothing
  Set rptao = Nothing
  Set prj = Nothing

  Rpt_Doc = blnRetValx

End Function

Private Function Rpt_Sec_Doc() As Boolean
' ** Document all report Sections to tblReport_Section.
' ** Called by:
' **   QuikRptDoc(), Above

  Const THIS_PROC As String = "Rpt_Sec_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, rpt As Access.Report, prp As Object
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim lngSecs As Long, arr_varSec() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim strSection As String
  Dim blnFound As Boolean, blnDelete As Boolean, blnLoop As Boolean
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim lngTmp00 As Long
  Dim lngX As Long, lngY As Long, lngElemS As Long, lngE As Long

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 15  ' ** Array's first-element UBound().
  Const R_DID     As Integer = 0
  Const R_DNAM    As Integer = 1
  Const R_RID     As Integer = 2
  Const R_OTYP    As Integer = 3
  Const R_RNAM    As Integer = 4
  Const R_CTLS    As Integer = 5
  Const R_CAP     As Integer = 6
  Const R_HAS_SUB As Integer = 7
  Const R_IS_SUB  As Integer = 8
  Const R_ACT     As Integer = 9
  Const R_HID     As Integer = 10
  Const R_DSC     As Integer = 11
  Const R_TAG     As Integer = 12
  Const R_PARSUB  As Integer = 13
  Const R_SUBS    As Integer = 14
  Const R_DAT     As Integer = 15

  ' ** Array: arr_varSec().
  Const S_ELEMS As Integer = 23  ' ** Array's first-element UBound().
  Const S_DID   As Integer = 0
  Const S_DNAM  As Integer = 1
  Const S_RID   As Integer = 2
  Const S_RNAM  As Integer = 3
  Const S_SID   As Integer = 4
  Const S_OTYP  As Integer = 5
  Const S_IDX   As Integer = 6   'sec_index
  Const S_SNAM  As Integer = 7
  Const S_BCLR  As Integer = 8
  Const S_GROW  As Integer = 9   'sec_cangrow
  Const S_SHRNK As Integer = 10  'sec_canshrink
  Const S_FORCE As Integer = 11  'sec_forcenewpage
  Const S_HGHT  As Integer = 12
  Const S_KEEP  As Integer = 13  'sec_keeptogether
  Const S_NEWRC As Integer = 14  'sec_newroworcol
  Const S_FRMT  As Integer = 15
  Const S_PRNT  As Integer = 16
  Const S_RTRT  As Integer = 17
  Const S_REP   As Integer = 18  'sec_repeatsection
  Const S_SPEF  As Integer = 19
  Const S_TAG   As Integer = 20  'sec_tag
  Const S_VIS   As Integer = 21
  Const S_GRLVL As Integer = 22
  Const S_DAT   As Integer = 23

  Const S_MAX As Long = 20&  ' ** Reports may have many more, AND THE COLLECTION CAN HAVE HOLES IN IT!

  Const DEL_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const DEL_DID  As Integer = 0
  Const DEL_DNAM As Integer = 1
  Const DEL_RID  As Integer = 2
  Const DEL_RNAM As Integer = 3
  Const DEL_SID  As Integer = 4
  Const DEL_SNAM As Integer = 5

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  If Reports.Count > 0 Then
    Do While Reports.Count > 0
      DoCmd.Close acReport, Reports(0).Name, acSaveNo
      DoEvents
    Loop
  End If

  Set dbs = CurrentDb
  With dbs

    lngRpts = 0&
    ReDim arr_varRpt(R_ELEMS, 0)

    ' ** Get a list of all reports.
    Set rst = .OpenRecordset("tblReport", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngRpts = lngRpts + 1&
          lngE = lngRpts - 1&
          ReDim Preserve arr_varRpt(R_ELEMS, lngE)
          ' ******************************************************
          ' ** Array: arr_varRpt()
          ' **
          ' **   Field  Element  Name                Constant
          ' **   =====  =======  ==================  ===========
          ' **     1       0     dbs_id              R_DID
          ' **     2       1     dbs_name            R_DNAM
          ' **     3       2     rpt_id              R_RID
          ' **     4       3     objtype_type        R_OTYP
          ' **     5       4     rpt_name            R_RNAM
          ' **     6       5     rpt_controls        R_CTLS
          ' **     7       6     rpt_caption         R_CAP
          ' **     8       7     rpt_hassub          R_HAS_SUB
          ' **     9       8     rpt_issub           R_IS_SUB
          ' **    10       9     rpt_active          R_ACT
          ' **    11      10     sec_hidden          R_HID
          ' **    12      11     rpt_description     R_DSC
          ' **    13      12     rpt_tag             R_TAG
          ' **    14      13     rpt_parent_sub      R_PARSUB
          ' **    15      14     rpt_subs            R_SUBS
          ' **    16      15     rpt_datemodified    R_DAT
          ' **
          ' ******************************************************
          arr_varRpt(R_DID, lngE) = ![dbs_id]
          arr_varRpt(R_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
          arr_varRpt(R_RID, lngE) = ![rpt_id]
          arr_varRpt(R_OTYP, lngE) = ![objtype_type]
          arr_varRpt(R_RNAM, lngE) = ![rpt_name]
          arr_varRpt(R_CTLS, lngE) = ![rpt_controls]
          arr_varRpt(R_CAP, lngE) = ![rpt_caption]
          arr_varRpt(R_HAS_SUB, lngE) = ![rpt_hassub]
          arr_varRpt(R_IS_SUB, lngE) = ![rpt_issub]
          arr_varRpt(R_ACT, lngE) = ![rpt_active]
          arr_varRpt(R_HID, lngE) = ![sec_hidden]
          arr_varRpt(R_DSC, lngE) = ![rpt_description]
          arr_varRpt(R_TAG, lngE) = ![rpt_tag]
          arr_varRpt(R_PARSUB, lngE) = ![rpt_parent_sub]
          arr_varRpt(R_SUBS, lngE) = ![rpt_subs]
          arr_varRpt(R_DAT, lngE) = ![rpt_datemodified]
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    lngSecs = 0&
    ReDim arr_varSec(S_ELEMS, 0)

    For lngX = 0& To (lngRpts - 1&)
      DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
      Reports(arr_varRpt(R_RNAM, lngX)).Visible = False
      Set rpt = Reports(arr_varRpt(R_RNAM, lngX))
      With rpt
        ' ** Report Sections.
        ' ** NOTE: THE REPORT SECTION COLLECTION CAN HAVE HOLES IN IT!!
        For lngY = 0& To (S_MAX - 1&)
          lngElemS = lngY
On Error Resume Next
          strSection = .Section(lngElemS).Name
          If ERR = 0 Then
On Error GoTo 0
            lngSecs = lngSecs + 1&
            lngE = lngSecs - 1&
            ReDim Preserve arr_varSec(S_ELEMS, lngE)
            ' *******************************************************
            ' ** Array: arr_varRpt()
            ' **
            ' **   Field  Element  Name                 Constant
            ' **   =====  =======  ===================  ===========
            ' **     1       0     dbs_id               S_DID
            ' **     2       1     dbs_name             S_DNAM
            ' **     3       2     rpt_id               S_RID
            ' **     4       3     rpt_name             S_RNAM
            ' **     5       4     sec_id               S_SID
            ' **     6       5     objtype_type         S_OTYP
            ' **     7       6     sec_index            S_IDX
            ' **     8       7     sec_name             S_SNAM
            ' **     9       8     sec_backcolor        S_BCLR
            ' **    10       9     sec_cangrow          S_GROW
            ' **    11      10     sec_canshrink        S_SHRNK
            ' **    12      11     sec_forcenewpage     S_FORCE
            ' **    13      12     sec_height           S_HGHT
            ' **    14      13     sec_keeptogether     S_KEEP
            ' **    15      14     sec_newroworcol      S_NEWRC
            ' **    16      15     sec_onformat         S_FRMT
            ' **    17      16     sec_onprint          S_PRNT
            ' **    18      17     sec_onretreat        S_RTRT
            ' **    19      18     sec_repeatsection    S_REP
            ' **    20      19     sec_specialeffect    S_SPEF
            ' **    21      20     sec_tag              S_TAG
            ' **    22      21     sec_visible          S_VIS
            ' **    23      22     sec_grouplevel       S_GRLVL
            ' **    24      23     sec_datemodified     S_DAT
            ' **
            ' *******************************************************
            arr_varSec(S_DID, lngE) = arr_varRpt(R_DID, lngX)
            arr_varSec(S_DNAM, lngE) = arr_varRpt(R_DNAM, lngX)
            arr_varSec(S_RID, lngE) = arr_varRpt(R_RID, lngX)
            arr_varSec(S_RNAM, lngE) = rpt.Name
            arr_varSec(S_SID, lngE) = CLng(0)
            arr_varSec(S_OTYP, lngE) = arr_varRpt(R_OTYP, lngX)
            arr_varSec(S_IDX, lngE) = lngElemS
            arr_varSec(S_SNAM, lngE) = strSection
            arr_varSec(S_GROW, lngE) = CBool(False)         ' ** Default.
            arr_varSec(S_SHRNK, lngE) = CBool(False)        ' ** Default.
            arr_varSec(S_FORCE, lngE) = acForceNewPageNone  ' ** Default.
            arr_varSec(S_KEEP, lngE) = CBool(False)         ' ** Default.
            arr_varSec(S_NEWRC, lngE) = acNewRowOrColNone   ' ** Default.
            arr_varSec(S_FRMT, lngE) = Null
            arr_varSec(S_PRNT, lngE) = Null
            arr_varSec(S_RTRT, lngE) = Null
            arr_varSec(S_REP, lngE) = CBool(False)          ' ** Default.
            arr_varSec(S_TAG, lngE) = Null
            arr_varSec(S_GRLVL, lngE) = CLng(-1)
            arr_varSec(S_DAT, lngE) = Null
            For Each prp In .Section(lngElemS).Properties
              With prp
                Select Case .Name
                Case "BackColor"
                  arr_varSec(S_BCLR, lngE) = .Value
                Case "CanGrow"
                  arr_varSec(S_GROW, lngE) = .Value
                Case "CanShrink"
                  arr_varSec(S_SHRNK, lngE) = .Value
                Case "ForceNewPage"
                  arr_varSec(S_FORCE, lngE) = .Value
                Case "Height"
                  arr_varSec(S_HGHT, lngE) = .Value
                Case "KeepTogether"
                  arr_varSec(S_KEEP, lngE) = .Value
                Case "NewRowOrCol"
                  arr_varSec(S_NEWRC, lngE) = .Value
                Case "OnFormat"
                  If IsNull(.Value) = False Then
                    If Trim(.Value) <> vbNullString Then
                      arr_varSec(S_FRMT, lngE) = .Value
                    End If
                  End If
                Case "OnPrint"
                  If IsNull(.Value) = False Then
                    If Trim(.Value) <> vbNullString Then
                      arr_varSec(S_PRNT, lngE) = .Value
                    End If
                  End If
                Case "OnRetreat"
                  If IsNull(.Value) = False Then
                    If Trim(.Value) <> vbNullString Then
                      arr_varSec(S_RTRT, lngE) = .Value
                    End If
                  End If
                Case "RepeatSection"
                  arr_varSec(S_REP, lngE) = .Value
                Case "SpecialEffect"
                  arr_varSec(S_SPEF, lngE) = .Value
                Case "Tag"
                  If IsNull(.Value) = False Then
                    If Trim(.Value) <> vbNullString Then
                      arr_varSec(S_TAG, lngE) = .Value
                    End If
                  End If
                Case "Visible"
                  arr_varSec(S_VIS, lngE) = .Value
                End Select
              End With
            Next
          Else
On Error GoTo 0
          End If
        Next
      End With
      DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX), acSaveNo
    Next

    Set rst = dbs.OpenRecordset("tblReport_Section", dbOpenDynaset, dbConsistent)
    With rst

      For lngX = 0& To (lngSecs - 1&)
        blnLoop = True: lngTmp00 = 0&
        Do While blnLoop = True
          blnLoop = False
          blnFound = False
          .FindFirst "[sec_name] = '" & arr_varSec(S_SNAM, lngX) & "' And [rpt_id] = " & CStr(arr_varSec(S_RID, lngX)) & " And " & _
            "[dbs_id] = " & CStr(arr_varSec(S_DID, lngX))
          If .NoMatch = True Then
            .FindFirst "[sec_index] = " & CStr(arr_varSec(S_IDX, lngX)) & " And [rpt_id] = " & CStr(arr_varSec(S_RID, lngX)) & " And " & _
              "[dbs_id] = " & CStr(arr_varSec(S_DID, lngX))
            If .NoMatch = True Then
              .AddNew
'dbs_id
              ![dbs_id] = lngThisDbsID
'rpt_id
              ![rpt_id] = arr_varSec(S_RID, lngX)
'sec_id
              ' ** ![sec_id] : AutoNumber.
'objtype_type
              ![objtype_type] = acReport
              ![sec_index] = arr_varSec(S_IDX, lngX)
              ![sec_name] = arr_varSec(S_SNAM, lngX)
              ![sec_backcolor] = arr_varSec(S_BCLR, lngX)
              ![sec_cangrow] = arr_varSec(S_GROW, lngX)
              ![sec_canshrink] = arr_varSec(S_SHRNK, lngX)
              ![sec_forcenewpage] = arr_varSec(S_FORCE, lngX)
              ![sec_height] = arr_varSec(S_HGHT, lngX)
              ![sec_keeptogether] = arr_varSec(S_KEEP, lngX)
              ![sec_newroworcol] = arr_varSec(S_NEWRC, lngX)
              If IsNull(arr_varSec(S_FRMT, lngE)) = False Then
                ![sec_onformat] = arr_varSec(S_FRMT, lngE)
              Else
                ![sec_onformat] = Null
              End If
              If IsNull(arr_varSec(S_PRNT, lngE)) = False Then
                ![sec_onprint] = arr_varSec(S_PRNT, lngE)
              Else
                ![sec_onprint] = Null
              End If
              If IsNull(arr_varSec(S_RTRT, lngE)) = False Then
                ![sec_onretreat] = arr_varSec(S_RTRT, lngE)
              Else
                ![sec_onretreat] = Null
              End If
              ![sec_repeatsection] = arr_varSec(S_REP, lngX)
              ![sec_specialeffect] = arr_varSec(S_SPEF, lngX)
              If IsNull(arr_varSec(S_TAG, lngE)) = False Then
                ![sec_tag] = arr_varSec(S_TAG, lngX)
              Else
                ![sec_tag] = Null
              End If
              ![sec_visible] = arr_varSec(S_VIS, lngX)
              '![sec_grouplevel] =
              ![sec_datemodified] = Now()
              .Update
            Else
              blnFound = True
            End If
          Else
            blnFound = True
          End If
          If blnFound = True Then
'sec_index
            If IsNull(![sec_index]) = True Then
              .Edit
              ![sec_index] = arr_varSec(S_IDX, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_index] <> arr_varSec(S_IDX, lngX) Then
                .Edit
                ![sec_index] = arr_varSec(S_IDX, lngX)
                ![sec_datemodified] = Now()
On Error Resume Next
                .Update
                If ERR.Number <> 0 Then
On Error GoTo 0
                  ' ** Because I renamed some sections, it's causing problems.
                  ' ** Delete the previous entry for this sec_index.
                  ' ** Unique Indexes:
                  ' **   rpt_id, sec_index
                  ' **   rpt_id, sec_name
                  ' ** DbUpdate enumeration:
                  ' **   1  dbUpdateRegular        Update: Pending changes are not cached and are written to disk immediately. (Default)
                  ' **                             Cancel: Cancels pending changes that aren’t cached. (Default)
                  ' **   2  dbUpdateCurrentRecord  Update: Only the current record's pending changes are written to disk.
                  ' **                             Cancel: N/A.
                  ' **   4  dbUpdateBatch          Update: All pending changes in the update cache are written to disk.
                  ' **                             Cancel: Cancels pending changes in the update cache.
                  .CancelUpdate dbUpdateRegular
                  ' ** Delete tblReport_Section, by specified [rptid], [secidx].
                  Set qdf = dbs.QueryDefs("zz_qry_Report_Section_01b")
                  With qdf.Parameters
                    ![rptid] = arr_varSec(S_RID, lngX)
                    ![secidx] = arr_varSec(S_IDX, lngX)
                  End With
                  qdf.Execute
                  ' ** Delete tblReport_Section, by specified [rptid], [secnam].
                  Set qdf = dbs.QueryDefs("zz_qry_Report_Section_01c")
                  With qdf.Parameters
                    ![rptid] = arr_varSec(S_RID, lngX)
                    ![secnam] = arr_varSec(S_SNAM, lngX)
                  End With
                  qdf.Execute
                  rst.Requery
                  blnLoop = True
                Else
On Error GoTo 0
                End If
              End If
            End If
'sec_name
            If ![sec_name] <> arr_varSec(S_SNAM, lngX) Then
              .Edit
              ![sec_name] = arr_varSec(S_SNAM, lngX)
              ![sec_datemodified] = Now()
              .Update
            End If
            If blnLoop = False Then
'sec_backcolor
              If ![sec_backcolor] <> arr_varSec(S_BCLR, lngX) Then
                .Edit
                ![sec_backcolor] = arr_varSec(S_BCLR, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
'sec_cangrow
              If ![sec_cangrow] <> arr_varSec(S_GROW, lngX) Then
                .Edit
                ![sec_cangrow] = arr_varSec(S_GROW, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
'sec_canshrink
              If ![sec_canshrink] <> arr_varSec(S_SHRNK, lngX) Then
                .Edit
                ![sec_canshrink] = arr_varSec(S_SHRNK, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
'sec_forcenewpage
              If IsNull(![sec_forcenewpage]) = True Then
                .Edit
                ![sec_forcenewpage] = arr_varSec(S_FORCE, lngX)
                ![sec_datemodified] = Now()
                .Update
              Else
                If ![sec_forcenewpage] <> arr_varSec(S_FORCE, lngX) Then
                  .Edit
                  ![sec_forcenewpage] = arr_varSec(S_FORCE, lngX)
                  ![sec_datemodified] = Now()
                  .Update
                End If
              End If
'sec_height
              If ![sec_height] <> arr_varSec(S_HGHT, lngX) Then
                .Edit
                ![sec_height] = arr_varSec(S_HGHT, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
'sec_keeptogether
              If ![sec_keeptogether] <> arr_varSec(S_KEEP, lngX) Then
                .Edit
                ![sec_keeptogether] = arr_varSec(S_KEEP, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
'sec_newroworcol
              If IsNull(![sec_newroworcol]) = True Then
                .Edit
                ![sec_newroworcol] = arr_varSec(S_NEWRC, lngX)
                ![sec_datemodified] = Now()
                .Update
              Else
                If ![sec_newroworcol] <> arr_varSec(S_NEWRC, lngX) Then
                  .Edit
                  ![sec_newroworcol] = arr_varSec(S_NEWRC, lngX)
                  ![sec_datemodified] = Now()
                  .Update
                End If
              End If
'sec_onformat
              If IsNull(arr_varSec(S_FRMT, lngE)) = False Then
                If IsNull(![sec_onformat]) = True Then
                  .Edit
                  ![sec_onformat] = arr_varSec(S_FRMT, lngE)
                  ![sec_datemodified] = Now()
                  .Update
                Else
                  If ![sec_onformat] <> arr_varSec(S_FRMT, lngE) Then
                    .Edit
                    ![sec_onformat] = arr_varSec(S_FRMT, lngE)
                    ![sec_datemodified] = Now()
                    .Update
                  End If
                End If
              Else
                If IsNull(![sec_onformat]) = False Then
                  .Edit
                  ![sec_onformat] = Null
                  ![sec_datemodified] = Now()
                  .Update
                End If
              End If
'sec_onprint
              If IsNull(arr_varSec(S_PRNT, lngE)) = False Then
                If IsNull(![sec_onprint]) = True Then
                  .Edit
                  ![sec_onprint] = arr_varSec(S_PRNT, lngE)
                  ![sec_datemodified] = Now()
                  .Update
                Else
                  If ![sec_onprint] <> arr_varSec(S_PRNT, lngE) Then
                    .Edit
                    ![sec_onprint] = arr_varSec(S_PRNT, lngE)
                    ![sec_datemodified] = Now()
                    .Update
                  End If
                End If
              Else
                If IsNull(![sec_onprint]) = False Then
                  .Edit
                  ![sec_onprint] = Null
                  ![sec_datemodified] = Now()
                  .Update
                End If
              End If
'sec_onretreat
              If IsNull(arr_varSec(S_RTRT, lngE)) = False Then
                If IsNull(![sec_onretreat]) = True Then
                  .Edit
                  ![sec_onretreat] = arr_varSec(S_RTRT, lngE)
                  ![sec_datemodified] = Now()
                  .Update
                Else
                  If ![sec_onretreat] <> arr_varSec(S_RTRT, lngE) Then
                    .Edit
                    ![sec_onretreat] = arr_varSec(S_RTRT, lngE)
                    ![sec_datemodified] = Now()
                    .Update
                  End If
                End If
              Else
                If IsNull(![sec_onretreat]) = False Then
                  .Edit
                  ![sec_onretreat] = Null
                  ![sec_datemodified] = Now()
                  .Update
                End If
              End If
'sec_repeatsection
              If ![sec_repeatsection] <> arr_varSec(S_REP, lngX) Then
                .Edit
                ![sec_repeatsection] = arr_varSec(S_REP, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
'sec_specialeffect
              If IsNull(![sec_specialeffect]) = True Then
                .Edit
                ![sec_specialeffect] = arr_varSec(S_SPEF, lngX)
                ![sec_datemodified] = Now()
                .Update
              Else
                If ![sec_specialeffect] <> arr_varSec(S_SPEF, lngX) Then
                  .Edit
                  ![sec_specialeffect] = arr_varSec(S_SPEF, lngX)
                  ![sec_datemodified] = Now()
                  .Update
                End If
              End If
'sec_tag
              If IsNull(arr_varSec(S_TAG, lngX)) = False Then
                If IsNull(![sec_tag]) = True Then
                  .Edit
                  ![sec_tag] = arr_varSec(S_TAG, lngX)
                  ![sec_datemodified] = Now()
                  .Update
                Else
                  If ![sec_tag] <> arr_varSec(S_TAG, lngX) Then
                    .Edit
                    ![sec_tag] = arr_varSec(S_TAG, lngX)
                    ![sec_datemodified] = Now()
                    .Update
                  End If
                End If
              Else
                If IsNull(![sec_tag]) = False Then
                  .Edit
                  ![sec_tag] = Null
                  ![sec_datemodified] = Now()
                  .Update
                End If
              End If
'sec_visible
              If ![sec_visible] <> arr_varSec(S_VIS, lngX) Then
                .Edit
                ![sec_visible] = arr_varSec(S_VIS, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
'sec_grouplevel
'sec_datemodified
            End If  ' ** blnLoop.
          End If
          lngTmp00 = lngTmp00 + 1&
          If lngTmp00 > 5& Then
            Stop
          End If
        Loop  ' ** blnLoop.
      Next

      lngDels = 0&
      ReDim arr_varDel(DEL_ELEMS, 0)

      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          blnFound = False
          For lngY = 0& To (lngSecs - 1&)
            If arr_varSec(S_DID, lngY) = ![dbs_id] And arr_varSec(S_RID, lngY) = ![rpt_id] And arr_varSec(S_SNAM, lngY) = ![sec_name] Then
              blnFound = True
              Exit For
            End If
          Next
          If blnFound = False Then
            lngDels = lngDels + 1&
            lngE = lngDels - 1&
            ReDim Preserve arr_varDel(DEL_ELEMS, lngE)
            arr_varDel(DEL_DID, lngE) = ![dbs_id]
            arr_varDel(DEL_DNAM, lngE) = DLookup("[dbs_name]", "tblDatabase", "[dbs_id] = " & CStr(![dbs_id]))
            arr_varDel(DEL_RID, lngE) = ![rpt_id]
            arr_varDel(DEL_RNAM, lngE) = DLookup("[rpt_name]", "tblReport", "[rpt_id] = " & CStr(![rpt_id]))
            arr_varDel(DEL_SID, lngE) = ![sec_id]
            arr_varDel(DEL_SNAM, lngE) = ![sec_name]
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next

      For lngX = 0& To (lngDels - 1&)
        If arr_varDel(DEL_DID, lngX) = lngThisDbsID Then
          blnDelete = True
          Debug.Print "'DEL SEC? " & arr_varDel(DEL_SNAM, lngX) & " on " & arr_varDel(DEL_RNAM, lngX)
Stop
          If blnDelete = True Then
            ' ** Delete tblReport_Section, by specified [secid].
            Set qdf = dbs.QueryDefs("zz_qry_Report_Section_01a")
            With qdf.Parameters
              ![secid] = arr_varDel(DEL_SID, lngX)
            End With
            qdf.Execute
          End If
        End If
      Next

      .Close
    End With

    .Close
  End With

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set prp = Nothing
  Set rpt = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_Sec_Doc = blnRetValx

End Function

Private Function Rpt_Ctl_Doc() As Boolean
' ** Document all report controls in Trust Accountant to tblReport_Control.
' ** Called by:
' **   QuikRptDoc(), Above

  Const THIS_PROC As String = "Rpt_Ctl_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset
  Dim rpt As Access.Report, ctl As Access.Control, prp As Object
  Dim cbof As Access.BoundObjectFrame, cchk As Access.CheckBox, cimg As Access.Image, clbl As Access.Label
  Dim clin As Access.Line, cpgb As Access.PageBreak, cbox As Access.Rectangle, csub As Access.SubForm
  Dim ctxt As Access.TextBox
  Dim lngRpts As Long, arr_varRpt As Variant
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim strParent As String
  Dim lngThisDbsID As Long, lngCtlID As Long, lngCtlSpecID As Long
  Dim blnFound As Boolean
  Dim lngLen As Long
  Dim arr_varTmp() As Variant
  Dim lngX As Long, lngY As Long, dblZ As Long

  ' ** Array: arr_varRpt().
  Const R_DID  As Integer = 0
  Const R_DNAM As Integer = 1
  Const R_RID  As Integer = 2
  Const R_RNAM As Integer = 3

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const C_DID  As Integer = 0
  Const C_DNAM As Integer = 1
  Const C_RID  As Integer = 2
  Const C_RNAM As Integer = 3
  Const C_CID  As Integer = 4
  Const C_CNAM As Integer = 5
  Const C_FND  As Integer = 6

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  DoCmd.Hourglass True
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** 1. Get a list of all the reports.

    ' ** tblReport, just needed fields, by specified CurrentAppName().
    Set qdf = .QueryDefs("zz_qry_Report_02")
    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRpts = .RecordCount
      .MoveFirst
      arr_varRpt = .GetRows(lngRpts)
      ' *********************************************
      ' ** Array: arr_varRpt()
      ' **
      ' **   Field  Element  Name        Constant
      ' **   =====  =======  ==========  ==========
      ' **     1       0     dbs_id      R_DID
      ' **     2       1     dbs_name    R_DNAM
      ' **     3       2     rpt_id      R_RID
      ' **     4       3     rpt_name    R_RNAM
      ' **
      ' *********************************************
      .Close
    End With

    ' ** 2. Get a list of all the controls.

    ' ** tblReport_Control, just needed fields, by specified CurrentAppName().
    Set qdf = .QueryDefs("zz_qry_Report_Control_02")
    Set rst1 = qdf.OpenRecordset
    With rst1
      If .BOF = True And .EOF = True Then
        ReDim arr_varTmp(C_ELEMS, 0)
        arr_varTmp(C_DID, 0) = lngThisDbsID
        arr_varTmp(C_DNAM, 0) = CurrentAppName  ' ** Module Function: modFileUtilities.
        arr_varTmp(C_RID, 0) = CLng(0)
        arr_varTmp(C_RNAM, 0) = vbNullString
        arr_varTmp(C_CID, 0) = CLng(0)
        arr_varTmp(C_CNAM, 0) = vbNullString
        arr_varTmp(C_FND, 0) = CBool(False)
        arr_varCtl = arr_varTmp
      Else
        .MoveLast
        lngCtls = .RecordCount
        .MoveFirst
        arr_varCtl = .GetRows(lngCtls)
        ' **********************************************
        ' ** Array: arr_varSub()
        ' **
        ' **   Field  Element  Name         Constant
        ' **   =====  =======  ===========  ==========
        ' **     1       0     dbs_id       C_DID
        ' **     2       1     dbs_name     C_DNAM
        ' **     3       2     rpt_id       C_RID
        ' **     4       3     rpt_name     C_RNAM
        ' **     5       4     ctl_id       C_CID
        ' **     6       5     ctl_name     C_CNAM
        ' **     7       6     ctl_found    C_FND
        ' **
        ' **********************************************
      End If
      .Close
    End With

    Set rst1 = .OpenRecordset("tblReport_Control", dbOpenDynaset, dbConsistent)
    Set rst2 = .OpenRecordset("tblReport_Control_Specification_A", dbOpenDynaset, dbConsistent)
    Set rst3 = .OpenRecordset("tblReport_Control_Specification_B", dbOpenDynaset, dbConsistent)

    ' ** 3. Now open each report and collect the control name into tblReport_Control.

    For lngX = 0& To (lngRpts - 1&)
      DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
      Reports(arr_varRpt(R_RNAM, lngX)).Visible = False
      Set rpt = Reports(arr_varRpt(R_RNAM, lngX))
      With rpt
        For Each ctl In .Controls
          With ctl
            Select Case ctl.ControlType
            Case acBoundObjectFrame
              Set cbof = rpt.Controls(ctl.Name)
            Case acCheckBox
              Set cchk = rpt.Controls(ctl.Name)
            Case acImage
              Set cimg = rpt.Controls(ctl.Name)
            Case acLabel
              Set clbl = rpt.Controls(ctl.Name)
            Case acLine
              Set clin = rpt.Controls(ctl.Name)
            Case acPageBreak
              Set cpgb = rpt.Controls(ctl.Name)
            Case acRectangle
              Set cbox = rpt.Controls(ctl.Name)
            Case acSubform
              Set csub = rpt.Controls(ctl.Name)
            Case acTextBox
              Set ctxt = rpt.Controls(ctl.Name)
            Case Else
              'ccbx, ccmd, clbx, copb, copg, cpag, ctab
            End Select
            blnFound = False
            If rst1.BOF = True And rst1.EOF = True Then
              ' ** Table is empty.
            Else
              rst1.FindFirst "[ctl_name] = '" & .Name & "' And [rpt_id] = " & CStr(arr_varRpt(R_RID, lngX))
              If rst1.NoMatch = False Then
                blnFound = True: lngCtlID = 0&: lngCtlSpecID = 0&
                For lngY = 0& To (lngCtls - 1&)
                  If arr_varCtl(C_CID, lngY) = rst1![ctl_id] Then
                    arr_varCtl(C_FND, lngY) = True
                    lngCtlID = rst1![ctl_id]
                    Exit For
                  End If
                Next
                rst2.FindFirst "[ctl_id] = " & CStr(lngCtlID)
                Select Case rst2.NoMatch
                Case True
                  Stop
                Case False
                  lngCtlSpecID = rst2![ctlspec_id]
                End Select
                rst3.FindFirst "[ctlspec_id] = " & CStr(lngCtlSpecID)
                If rst3.NoMatch = True Then
                  Stop
                End If
              End If
            End If
            strParent = vbNullString
On Error Resume Next
            strParent = ctl.Parent.Name
On Error GoTo 0
            If blnFound = False Then

                rst1.AddNew
'dbs_id
                rst1![dbs_id] = arr_varRpt(R_DID, lngX)
'rpt_id
                rst1![rpt_id] = arr_varRpt(R_RID, lngX)
'ctl_id
                ' ** rst1![ctl_id] : AutoNumber.
'objtype_type
                rst1![objtype_type] = acReport
'ctl_name
                rst1![ctl_name] = ctl.Name
'ctltype_type
                rst1![ctltype_type] = ctl.ControlType
                Select Case ctl.ControlType
                Case acLabel
'ctl_caption
                  If IsNull(clbl.Caption) = False Then
                    If Trim(clbl.Caption) <> vbNullString Then
                      rst1![ctl_caption] = clbl.Caption
                    End If
                  End If
                End Select
'ctl_section
                rst1![ctl_section] = ctl.Section
                Select Case ctl.ControlType
                Case acBoundObjectFrame
'ctl_source
                  If IsNull(cbof.ControlSource) = False Then
                    If Trim(cbof.ControlSource) <> vbNullString Then
                      rst1![ctl_source] = cbof.ControlSource
                    End If
                  End If
                Case acCheckBox
                  If IsNull(cchk.ControlSource) = False Then
                    If Trim(cchk.ControlSource) <> vbNullString Then
                      rst1![ctl_source] = cchk.ControlSource
                    End If
                  End If
                Case acTextBox
                  If IsNull(ctxt.ControlSource) = False Then
                    If Trim(ctxt.ControlSource) <> vbNullString Then
                      rst1![ctl_source] = ctxt.ControlSource
                    End If
                  End If
                End Select
                Select Case ctl.ControlType
                Case acImage
'ctl_sourceobject
                  rst1![ctl_sourceobject] = cimg.Picture
                Case acSubform
                  If IsNull(csub.SourceObject) = False Then
                    If Trim(csub.SourceObject) <> vbNullString Then
                      rst1![ctl_sourceobject] = csub.SourceObject
                    End If
                  End If
                End Select
'ctl_controls
On Error Resume Next
                rst1![ctl_controls] = ctl.Controls.Count
                If ERR.Number <> 0 Then
On Error GoTo 0
                  rst1![ctl_controls] = 0&
                Else
On Error GoTo 0
                End If
'ctl_id_parent
                'rst1![ctl_id_parent] = 0&
'ctl_parent
                If strParent <> vbNullString Then
                  rst1![ctl_parent] = strParent
                End If
                rst1![ctl_datemodified] = Now()
                rst1.Update
                rst1.Bookmark = rst1.LastModified
                lngCtlID = rst1![ctl_id]
                rst2.AddNew
'dbs_id
                rst2![dbs_id] = arr_varRpt(R_DID, lngX)
'rpt_id
                rst2![rpt_id] = arr_varRpt(R_RID, lngX)
'ctl_id
                rst2![ctl_id] = lngCtlID
'ctlspec_id
                ' ** rst2![ctlspec_id] : AutoNumber.
                Select Case ctl.ControlType
                Case acBoundObjectFrame
'ctlspec_backcolor
                  rst2![ctlspec_backcolor] = cbof.BackColor
'ctlspec_backstyle
                  rst2![ctlspec_backstyle] = cbof.BackStyle
'ctlspec_bordercolor
                  rst2![ctlspec_bordercolor] = cbof.BorderColor
'ctlspec_borderlinestyle (deprecated)
'ctlspec_borderstyle
                  rst2![ctlspec_borderstyle] = cbof.BorderStyle
'ctlspec_borderwidth
                  rst2![ctlspec_borderwidth] = cbof.BorderWidth
'ctlspec_height
                  rst2![ctlspec_height] = cbof.Height
                Case acCheckBox
                  rst2![ctlspec_bordercolor] = cchk.BorderColor
                  rst2![ctlspec_borderstyle] = cchk.BorderStyle
                  rst2![ctlspec_borderwidth] = cchk.BorderWidth
                  rst2![ctlspec_height] = cchk.Height
                Case acImage
                  rst2![ctlspec_backcolor] = cimg.BackColor
                  rst2![ctlspec_backstyle] = cimg.BackStyle
                  rst2![ctlspec_bordercolor] = cimg.BorderColor
                  rst2![ctlspec_borderstyle] = cimg.BorderStyle
                  rst2![ctlspec_borderwidth] = cimg.BorderWidth
                  rst2![ctlspec_height] = cimg.Height
                Case acLabel
                  rst2![ctlspec_backcolor] = clbl.BackColor
                  rst2![ctlspec_backstyle] = clbl.BackStyle
                  rst2![ctlspec_bordercolor] = clbl.BorderColor
                  rst2![ctlspec_borderstyle] = clbl.BorderStyle
                  rst2![ctlspec_borderwidth] = clbl.BorderWidth
                  rst2![ctlspec_forecolor] = clbl.ForeColor
                  rst2![ctlspec_height] = clbl.Height
                Case acLine
                  rst2![ctlspec_bordercolor] = clin.BorderColor
                  rst2![ctlspec_borderstyle] = clin.BorderStyle
                  rst2![ctlspec_borderwidth] = clin.BorderWidth
                  rst2![ctlspec_height] = clin.Height
                Case acPageBreak
                  ' ** No additional properties are saved at this time.
                Case acRectangle
                  rst2![ctlspec_backcolor] = cbox.BackColor
                  rst2![ctlspec_backstyle] = cbox.BackStyle
                  rst2![ctlspec_bordercolor] = cbox.BorderColor
                  rst2![ctlspec_borderstyle] = cbox.BorderStyle
                  rst2![ctlspec_borderwidth] = cbox.BorderWidth
                  rst2![ctlspec_height] = cbox.Height
                Case acSubform
                  rst2![ctlspec_bordercolor] = csub.BorderColor
                  rst2![ctlspec_borderstyle] = csub.BorderStyle
                  rst2![ctlspec_borderwidth] = csub.BorderWidth
                  rst2![ctlspec_height] = csub.Height
                Case acTextBox
                  rst2![ctlspec_backcolor] = ctxt.BackColor
                  rst2![ctlspec_backstyle] = ctxt.BackStyle
                  rst2![ctlspec_bordercolor] = ctxt.BorderColor
                  rst2![ctlspec_borderstyle] = ctxt.BorderStyle
                  rst2![ctlspec_borderwidth] = ctxt.BorderWidth
'ctlspec_forecolor
                  rst2![ctlspec_forecolor] = ctxt.ForeColor
'ctlspec_format
                  If IsNull(ctxt.Format) = False Then
                    If Trim(ctxt.Format) <> vbNullString Then
                      rst2![ctlspec_format] = ctxt.Format
                    End If
                  End If
                  rst2![ctlspec_height] = ctxt.Height
                Case Else
                  'ccbx, ccmd, clbx, copb, copg, cpag, ctab
                  Beep
                  Debug.Print "'RPT: " & rpt.Name & "  CTL: " & ctl.Name & "  TYP: " & CStr(ctl.ControlType)
                  '![ctl_id_parent] =
                  '![ctl_parent] =
                End Select

                For Each prp In ctl.Properties
                  Select Case prp.Name
                  Case "BorderLineStyle"
'ctlspec_borderlinestyle  {deprecated}
                  Case "BottomMargin"
'ctlspec_bottommargin
                    If IsNull(rst2![ctlspec_bottommargin]) = True Then
                      rst2![ctlspec_bottommargin] = prp.Value
                    Else
                      If prp.Value <> rst2![ctlspec_bottommargin] Then
                        rst2![ctlspec_bottommargin] = prp.Value
                      End If
                    End If
                  Case "CanGrow"
'ctlspec_cangrow
                    If prp.Value <> rst2![ctlspec_cangrow] Then
                      rst2![ctlspec_cangrow] = prp.Value
                    End If
                  Case "CanShrink"
'ctlspec_canshrink
                    If prp.Value <> rst2![ctlspec_canshrink] Then
                      rst2![ctlspec_canshrink] = prp.Value
                    End If
                  Case "Caption"
'ctlspec_caption
                    Select Case IsNull(prp.Value)
                    Case True
                      If IsNull(rst2![ctlspec_caption]) = False Then
                        rst2![ctlspec_caption] = Null
                      End If
                    Case False
                      If Trim(prp.Value) <> vbNullString Then
                        Select Case IsNull(rst2![ctlspec_caption])
                        Case True
                          rst2![ctlspec_caption] = prp.Value
                        Case False
                          If prp.Value <> rst2![ctlspec_caption] Then
                            rst2![ctlspec_caption] = prp.Value
                          End If
                        End Select
                      Else
                        If IsNull(rst2![ctlspec_caption]) = False Then
                          rst2![ctlspec_caption] = Null
                        End If
                      End If
                    End Select
                  Case "ControlSource"
'ctlspec_controlsource
                    Select Case IsNull(prp.Value)
                    Case True
                      If IsNull(rst2![ctlspec_controlsource]) = False Then
                        rst2![ctlspec_controlsource] = Null
                      End If
                    Case False
                      If Trim(prp.Value) <> vbNullString Then
                        Select Case IsNull(rst2![ctlspec_controlsource])
                        Case True
                          rst2![ctlspec_controlsource] = prp.Value
                        Case False
                          If prp.Value <> rst2![ctlspec_controlsource] Then
                            rst2![ctlspec_controlsource] = prp.Value
                          End If
                        End Select
                      Else
                        If IsNull(rst2![ctlspec_controlsource]) = False Then
                          rst2![ctlspec_controlsource] = Null
                        End If
                      End If
                    End Select
                  Case "ControlType"
'ctlspec_controltype
                    Select Case IsNull(rst2![ctlspec_controltype])
                    Case True
                      rst2![ctlspec_controltype] = prp.Value
                    Case False
                      If prp.Value <> rst2![ctlspec_controltype] Then
                        rst2![ctlspec_controltype] = prp.Value
                      End If
                    End Select
                  Case "DecimalPlaces"
'ctlspec_decimalplaces
                    Select Case IsNull(rst2![ctlspec_decimalplaces])
                    Case True
                      rst2![ctlspec_decimalplaces] = prp.Value
                    Case False
                      If prp.Value <> rst2![ctlspec_decimalplaces] Then
                        rst2![ctlspec_decimalplaces] = prp.Value
                      End If
                    End Select
                  Case "DefaultValue"
'ctlspec_defaultvalue
                    If IsNull(prp.Value) = False Then
                      If Trim(prp.Value) <> vbNullString Then
                        rst2![ctlspec_defaultvalue] = prp.Value
                      Else
                        If IsNull(rst2![ctlspec_defaultvalue]) = False Then
                          rst2![ctlspec_defaultvalue] = Null
                        End If
                      End If
                    Else
                      If IsNull(rst2![ctlspec_defaultvalue]) = False Then
                        rst2![ctlspec_defaultvalue] = Null
                      End If
                    End If
                  Case "EventProcPrefix"
'ctlspec_eventprocprefix
                    Select Case IsNull(rst2![ctlspec_eventprocprefix])
                    Case True
                      rst2![ctlspec_eventprocprefix] = prp.Value
                    Case False
                      If prp.Value <> rst2![ctlspec_eventprocprefix] Then
                        rst2![ctlspec_eventprocprefix] = prp.Value
                      End If
                    End Select
                  Case "FontBold"
'ctlspec_fontbold
                    If prp.Value <> rst2![ctlspec_fontbold] Then
                      rst2![ctlspec_fontbold] = prp.Value
                    End If
                  Case "FontItalic"
'ctlspec_fontitalic
                    If prp.Value <> rst2![ctlspec_fontitalic] Then
                      rst2![ctlspec_fontitalic] = prp.Value
                    End If
                  Case "FontName"
'ctlspec_fontname
                    If IsNull(rst2![ctlspec_fontname]) = True Then
                      rst2![ctlspec_fontname] = prp.Value
                    Else
                      If prp.Value <> rst2![ctlspec_fontname] Then
                        rst2![ctlspec_fontname] = prp.Value
                      End If
                    End If
                  Case "FontSize"
'ctlspec_fontsize
                    If IsNull(rst2![ctlspec_fontsize]) = True Then
                      rst2![ctlspec_fontsize] = prp.Value
                    Else
                      If prp.Value <> rst2![ctlspec_fontsize] Then
                        rst2![ctlspec_fontsize] = prp.Value
                      End If
                    End If
                  Case "FontUnderline"
'ctlspec_fontunderline
                    If prp.Value <> rst2![ctlspec_fontunderline] Then
                      rst2![ctlspec_fontunderline] = prp.Value
                    End If
                  Case "FontWeight"
'ctlspec_fontweight
                    If IsNull(rst2![ctlspec_fontweight]) = True Then
                      rst2![ctlspec_fontweight] = prp.Value
                    Else
                      If prp.Value <> rst2![ctlspec_fontweight] Then
                        rst2![ctlspec_fontweight] = prp.Value
                      End If
                    End If
                  Case "Height"
'ctlspec_height
                    If IsNull(rst2![ctlspec_height]) = True Then
                      rst2![ctlspec_height] = prp.Value
                    Else
                      If rst2![ctlspec_height] <> prp.Value Then
                        rst2![ctlspec_height] = prp.Value
                      End If
                    End If
                  Case "HideDuplicates"
'ctlspec_hideduplicates
                    If prp.Value <> rst2![ctlspec_hideduplicates] Then
                      rst2![ctlspec_hideduplicates] = prp.Value
                    End If
                  Case "HyperlinkAddress"
'ctlspec_hyperlinkaddress
                    Select Case IsNull(prp.Value)
                    Case True
                      If IsNull(rst2![ctlspec_hyperlinkaddress]) = False Then
                        rst2![ctlspec_hyperlinkaddress] = Null
                      End If
                    Case False
                      If Trim(prp.Value) <> vbNullString Then
                        Select Case IsNull(rst2![ctlspec_hyperlinkaddress])
                        Case True
                          rst2![ctlspec_hyperlinkaddress] = prp.Value
                        Case False
                          If prp.Value <> rst2![ctlspec_hyperlinkaddress] Then
                            rst2![ctlspec_hyperlinkaddress] = prp.Value
                          End If
                        End Select
                      Else
                        If IsNull(rst2![ctlspec_hyperlinkaddress]) = False Then
                          rst2![ctlspec_hyperlinkaddress] = Null
                        End If
                      End If
                    End Select
                  Case "HyperlinkSubaddress"
'ctlspec_hyperlinksubaddress
                    Select Case IsNull(prp.Value)
                    Case True
                      If IsNull(rst2![ctlspec_hyperlinksubaddress]) = False Then
                        rst2![ctlspec_hyperlinksubaddress] = Null
                      End If
                    Case False
                      If Trim(prp.Value) <> vbNullString Then
                        Select Case IsNull(rst2![ctlspec_hyperlinksubaddress])
                        Case True
                          rst2![ctlspec_hyperlinksubaddress] = prp.Value
                        Case False
                          If prp.Value <> rst2![ctlspec_hyperlinksubaddress] Then
                            rst2![ctlspec_hyperlinksubaddress] = prp.Value
                          End If
                        End Select
                      Else
                        If IsNull(rst2![ctlspec_hyperlinksubaddress]) = False Then
                          rst2![ctlspec_hyperlinksubaddress] = Null
                        End If
                      End If
                    End Select
                  Case "ImageWidth"
'ctlspec_imagewidth
                    If IsNull(rst2![ctlspec_imagewidth]) = True Then
                      rst2![ctlspec_imagewidth] = prp.Value
                    Else
                      If rst2![ctlspec_imagewidth] <> prp.Value Then
                        rst2![ctlspec_imagewidth] = prp.Value
                      End If
                    End If
                  Case "ImageHeight"
'ctlspec_imageheight
                    If IsNull(rst2![ctlspec_imageheight]) = True Then
                      rst2![ctlspec_imageheight] = prp.Value
                    Else
                      If rst2![ctlspec_imageheight] <> prp.Value Then
                        rst2![ctlspec_imageheight] = prp.Value
                      End If
                    End If
                  Case "InputMask"
'ctlspec_inputmask
                    Select Case IsNull(prp.Value)
                    Case True
                      If IsNull(rst2![ctlspec_inputmask]) = False Then
                        rst2![ctlspec_inputmask] = Null
                      End If
                    Case False
                      If Trim(prp.Value) <> vbNullString Then
                        Select Case IsNull(rst2![ctlspec_inputmask])
                        Case True
                          rst2![ctlspec_inputmask] = prp.Value
                        Case False
                          If prp.Value <> rst2![ctlspec_inputmask] Then
                            rst2![ctlspec_inputmask] = prp.Value
                          End If
                        End Select
                      Else
                        If IsNull(rst2![ctlspec_inputmask]) = False Then
                          rst2![ctlspec_inputmask] = Null
                        End If
                      End If
                    End Select
                  Case "IsHyperlink"
'ctlspec_ishyperlink
                    If prp.Value <> rst2![ctlspec_ishyperlink] Then
                      rst2![ctlspec_ishyperlink] = prp.Value
                    End If
                  End Select
                Next
                rst2.Update
                rst2.Bookmark = rst2.LastModified
                lngCtlSpecID = rst2![ctlspec_id]

                rst3.AddNew
'dbs_id
                rst3![dbs_id] = arr_varRpt(R_DID, lngX)
'rpt_id
                rst3![rpt_id] = arr_varRpt(R_RID, lngX)
'ctl_id
                rst3![ctl_id] = lngCtlID
'ctlspec_id
                rst3![ctlspec_id] = lngCtlSpecID

                Select Case ctl.ControlType
                Case acBoundObjectFrame
'ctlspec_left
                  rst3![ctlspec_left] = cbof.Left
'ctlspec_specialeffect
                  rst3![ctlspec_specialeffect] = cbof.SpecialEffect
'ctlspec_top
                  rst3![ctlspec_top] = cbof.Top
'ctlspec_visible
                  rst3![ctlspec_visible] = cbof.Visible
'ctlspec_width
                  rst3![ctlspec_width] = cbof.Width
                Case acCheckBox
                  rst3![ctlspec_left] = cchk.Left
                  rst3![ctlspec_specialeffect] = cchk.SpecialEffect
                  rst3![ctlspec_top] = cchk.Top
                  rst3![ctlspec_visible] = cchk.Visible
                  rst3![ctlspec_width] = cchk.Width
                Case acImage
                  rst3![ctlspec_left] = cimg.Left
                  rst3![ctlspec_top] = cimg.Top
                  rst3![ctlspec_visible] = cimg.Visible
                  rst3![ctlspec_width] = cimg.Width
                Case acLabel
                  rst3![ctlspec_left] = clbl.Left
                  rst3![ctlspec_specialeffect] = clbl.SpecialEffect
                  rst3![ctlspec_top] = clbl.Top
                  rst3![ctlspec_visible] = clbl.Visible
                  rst3![ctlspec_width] = clbl.Width
                Case acLine
                  rst3![ctlspec_left] = clin.Left
                  rst3![ctlspec_specialeffect] = clin.SpecialEffect
                  rst3![ctlspec_top] = clin.Top
                  rst3![ctlspec_visible] = clin.Visible
                  rst3![ctlspec_width] = clin.Width
                Case acPageBreak
                  rst3![ctlspec_top] = cpgb.Top
                Case acRectangle
                  rst3![ctlspec_left] = cbox.Left
                  rst3![ctlspec_specialeffect] = cbox.SpecialEffect
                  rst3![ctlspec_top] = cbox.Top
                  rst3![ctlspec_visible] = cbox.Visible
                  rst3![ctlspec_width] = cbox.Width
                Case acSubform
                  rst3![ctlspec_left] = csub.Left
                  rst3![ctlspec_specialeffect] = csub.SpecialEffect
                  rst3![ctlspec_top] = csub.Top
                  rst3![ctlspec_visible] = csub.Visible
                  rst3![ctlspec_width] = csub.Width
                Case acTextBox
                  rst3![ctlspec_left] = ctxt.Left
                  rst3![ctlspec_specialeffect] = ctxt.SpecialEffect
                  rst3![ctlspec_top] = ctxt.Top
                  rst3![ctlspec_visible] = ctxt.Visible
                  rst3![ctlspec_width] = ctxt.Width
                Case Else
                  'ccbx, ccmd, clbx, copb, copg, cpag, ctab
                  Beep
                  Debug.Print "'RPT: " & rpt.Name & "  CTL: " & ctl.Name & "  TYP: " & CStr(ctl.ControlType)
                  '![ctl_id_parent] =
                  '![ctl_parent] =
                End Select

                For Each prp In ctl.Properties
                  Select Case prp.Name
                  Case "KeyboardLanguage"
'ctlspec_keyboardlanguage
                    Select Case IsNull(rst3![ctlspec_keyboardlanguage])
                    Case True
                      rst3![ctlspec_keyboardlanguage] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_keyboardlanguage] Then
                        rst3![ctlspec_keyboardlanguage] = prp.Value
                      End If
                    End Select
                  Case "Left"
'ctlspec_left
                    If IsNull(rst3![ctlspec_left]) = True Then
                      rst3![ctlspec_left] = prp.Value
                    Else
                      If rst3![ctlspec_left] <> prp.Value Then
                        rst3![ctlspec_left] = prp.Value
                      End If
                    End If
                  Case "LeftMargin"
'ctlspec_leftmargin
                    If IsNull(rst3![ctlspec_leftmargin]) = True Then
                      rst3![ctlspec_leftmargin] = prp.Value
                    Else
                      If prp.Value <> rst3![ctlspec_leftmargin] Then
                        rst3![ctlspec_leftmargin] = prp.Value
                      End If
                    End If
                  Case "LineSlant"
'ctlspec_lineslant
                    Select Case IsNull(rst3![ctlspec_lineslant])
                    Case True
                      rst3![ctlspec_lineslant] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_lineslant] Then
                        rst3![ctlspec_lineslant] = prp.Value
                      End If
                    End Select
                  Case "LineSpacing"
'ctlspec_linespacing
                    If IsNull(rst3![ctlspec_linespacing]) = True Then
                      rst3![ctlspec_linespacing] = prp.Value
                    Else
                      If prp.Value <> rst3![ctlspec_linespacing] Then
                        rst3![ctlspec_linespacing] = prp.Value
                      End If
                    End If
                  Case "NumeralShapes"
'ctlspec_numeralshapes
                    Select Case IsNull(rst3![ctlspec_numeralshapes])
                    Case True
                      rst3![ctlspec_numeralshapes] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_numeralshapes] Then
                        rst3![ctlspec_numeralshapes] = prp.Value
                      End If
                    End Select
                  Case "Picture"
'ctlspec_picture
                    If IsNull(rst3![ctlspec_picture]) = True Then
                      rst3![ctlspec_picture] = prp.Value
                    Else
                      If rst3![ctlspec_picture] <> prp.Value Then
                        rst3![ctlspec_picture] = prp.Value
                      End If
                    End If
                  Case "PictureAlignment"
'ctlspec_picturealignment
                    If IsNull(rst3![ctlspec_picturealignment]) = True Then
                      rst3![ctlspec_picturealignment] = prp.Value
                    Else
                      If rst3![ctlspec_picturealignment] <> prp.Value Then
                        rst3![ctlspec_picturealignment] = prp.Value
                      End If
                    End If
                  Case "PictureType"
'ctlspec_picturetype
                    If IsNull(rst3![ctlspec_picturetype]) = True Then
                      rst3![ctlspec_picturetype] = prp.Value
                    Else
                      If rst3![ctlspec_picturetype] <> prp.Value Then
                        rst3![ctlspec_picturetype] = prp.Value
                      End If
                    End If
                  Case "ReadingOrder"
'ctlspec_readingorder
                    Select Case IsNull(rst3![ctlspec_readingorder])
                    Case True
                      rst3![ctlspec_readingorder] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_readingorder] Then
                        rst3![ctlspec_readingorder] = prp.Value
                      End If
                    End Select
                  Case "RightMargin"
'ctlspec_rightmargin
                    If IsNull(rst3![ctlspec_rightmargin]) = True Then
                      rst3![ctlspec_rightmargin] = prp.Value
                    Else
                      If prp.Value <> rst3![ctlspec_rightmargin] Then
                        rst3![ctlspec_rightmargin] = prp.Value
                      End If
                    End If
                  Case "RunningSum"
'ctlspec_runningsum
                    Select Case IsNull(rst3![ctlspec_runningsum])
                    Case True
                      rst3![ctlspec_runningsum] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_runningsum] Then
                        rst3![ctlspec_runningsum] = prp.Value
                      End If
                    End Select
                  Case "ScrollBarAlign"
'ctlspec_scrollbaralign
                    Select Case IsNull(rst3![ctlspec_scrollbaralign])
                    Case True
                      rst3![ctlspec_scrollbaralign] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_scrollbaralign] Then
                        rst3![ctlspec_scrollbaralign] = prp.Value
                      End If
                    End Select
                  Case "Section"
'ctlspec_section
                    Select Case IsNull(rst3![ctlspec_section])
                    Case True
                      rst3![ctlspec_section] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_section] Then
                        rst3![ctlspec_section] = prp.Value
                      End If
                    End Select
                  Case "ShortcutMenuBar"
'ctlspec_shortcutmenubar
                    Select Case IsNull(prp.Value)
                    Case True
                      If IsNull(rst3![ctlspec_shortcutmenubar]) = False Then
                        rst3![ctlspec_shortcutmenubar] = Null
                      End If
                    Case False
                      If Trim(prp.Value) <> vbNullString Then
                        Select Case IsNull(rst3![ctlspec_shortcutmenubar])
                        Case True
                          rst3![ctlspec_shortcutmenubar] = prp.Value
                        Case False
                          If prp.Value <> rst3![ctlspec_shortcutmenubar] Then
                            rst3![ctlspec_shortcutmenubar] = prp.Value
                          End If
                        End Select
                      Else
                        If IsNull(rst3![ctlspec_shortcutmenubar]) = False Then
                          rst3![ctlspec_shortcutmenubar] = Null
                        End If
                      End If
                    End Select
                  Case "SizeMode"
'ctlspec_sizemode
                    If IsNull(rst3![ctlspec_sizemode]) = True Then
                      rst3![ctlspec_sizemode] = prp.Value
                    Else
                      If rst3![ctlspec_sizemode] <> prp.Value Then
                        rst3![ctlspec_sizemode] = prp.Value
                      End If
                    End If
                  Case "SmartTags"
'ctlspec_smarttags
                    Select Case IsNull(prp.Value)
                    Case True
                      If IsNull(rst3![ctlspec_smarttags]) = False Then
                        rst3![ctlspec_smarttags] = Null
                      End If
                    Case False
                      If Trim(prp.Value) <> vbNullString Then
                        Select Case IsNull(rst3![ctlspec_smarttags])
                        Case True
                          rst3![ctlspec_smarttags] = prp.Value
                        Case False
                          If prp.Value <> rst3![ctlspec_smarttags] Then
                            rst3![ctlspec_smarttags] = prp.Value
                          End If
                        End Select
                      Else
                        If IsNull(rst3![ctlspec_smarttags]) = False Then
                          rst3![ctlspec_smarttags] = Null
                        End If
                      End If
                    End Select
                  Case "Tag"
'ctlspec_tag
                    If IsNull(prp.Value) = True Then
                      If IsNull(rst3![ctlspec_tag]) = False Then
                        rst3![ctlspec_tag] = Null
                      End If
                    Else
                      If Trim(prp.Value) = vbNullString Then
                        If IsNull(rst3![ctlspec_tag]) = False Then
                          rst3![ctlspec_tag] = Null
                        End If
                      Else
                        If IsNull(rst3![ctlspec_tag]) = True Then
                          rst3![ctlspec_tag] = prp.Value
                        Else
                          If rst3![ctlspec_tag] <> prp.Value Then
                            rst3![ctlspec_tag] = prp.Value
                          End If
                        End If
                      End If
                    End If
                  Case "TextAlign"
'ctlspec_textalign
                    If IsNull(rst3![ctlspec_textalign]) = True Then
                      rst3![ctlspec_textalign] = prp.Value
                    Else
                      If prp.Value <> rst3![ctlspec_textalign] Then
                        rst3![ctlspec_textalign] = prp.Value
                      End If
                    End If
                  Case "TextFontCharSet"
'ctlspec_textfontcharset
                    Select Case IsNull(rst3![ctlspec_textfontcharset])
                    Case True
                      rst3![ctlspec_textfontcharset] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_textfontcharset] Then
                        rst3![ctlspec_textfontcharset] = prp.Value
                      End If
                    End Select
                  Case "Top"
'ctlspec_top
                    If IsNull(rst3![ctlspec_top]) = True Then
                      rst3![ctlspec_top] = prp.Value
                    Else
                      If rst3![ctlspec_top] <> prp.Value Then
                        rst3![ctlspec_top] = prp.Value
                      End If
                    End If
                  Case "TopMargin"
'ctlspec_topmargin
                    If IsNull(rst3![ctlspec_topmargin]) = True Then
                      rst3![ctlspec_topmargin] = prp.Value
                    Else
                      If prp.Value <> rst3![ctlspec_topmargin] Then
                        rst3![ctlspec_topmargin] = prp.Value
                      End If
                    End If
                  Case "Vertical"
'ctlspec_vertical
                    If prp.Value <> rst3![ctlspec_vertical] Then
                      rst3![ctlspec_vertical] = prp.Value
                    End If
                  Case "Visible"
'ctlspec_visible
                    If rst3![ctlspec_visible] <> prp.Value Then
                      rst3![ctlspec_visible] = prp.Value
                    End If
                  Case "Width"
'ctlspec_width
                    If IsNull(rst3![ctlspec_width]) = True Then
                      rst3![ctlspec_width] = prp.Value
                    Else
                      If rst3![ctlspec_width] <> prp.Value Then
                        rst3![ctlspec_width] = prp.Value
                      End If
                    End If
                  End Select
                Next
'ctlspec_datemodified
                rst3![ctlspec_datemodified] = Now()
                rst3.Update
              'End With

            Else

              blnFound = False
              lngLen = Len(.Name)
              If Len(rst1![ctl_name]) = lngLen Then
                For lngY = 1& To lngLen
                  If Asc(Mid$(rst1![ctl_name], lngY, 1)) <> Asc(Mid$(.Name, lngY, 1)) Then
                    ' ** Capitalization change detected.
                    blnFound = True
                    Exit For
                  End If
                Next
              Else
                ' ** Name change detected.
                blnFound = True
              End If
              If blnFound = True Then
                ' ** The Control's name was changed.
                With rst1
                  .Edit
                  ![ctl_name] = ctl.Name
                  ![ctl_datemodified] = Now()
                  .Update
                End With
              End If

              With rst1
                If strParent <> vbNullString Then
                  If IsNull(![ctl_parent]) = True Then
                    .Edit
                    ![ctl_parent] = strParent
                    ![ctl_datemodified] = Now()
                    .Update
                  Else
                    If strParent <> ![ctl_parent] Then
                      .Edit
                      ![ctl_parent] = strParent
                      ![ctl_datemodified] = Now()
                      .Update
                    End If
                  End If
                Else
                  If IsNull(![ctl_parent]) = False Then
                    .Edit
                    ![ctl_parent] = strParent
                    ![ctl_datemodified] = Now()
                    .Update
                  End If
                End If
              End With
              If rst1![ctltype_type] <> ctl.ControlType Then
                ' ** The Control's type was changed.
                With rst1
                  .Edit
                  ![ctltype_type] = ctl.ControlType
                  ![ctl_datemodified] = Now()
                  .Update
                End With
              End If

              With rst1
                .Edit
                ![ctl_section] = ctl.Section
                ![ctl_datemodified] = Now()
                .Update
              End With

              With rst1
                .Edit
                Select Case ctl.ControlType
                Case acBoundObjectFrame
                  If IsNull(cbof.ControlSource) = False Then
                    If Trim(cbof.ControlSource) <> vbNullString Then
                      ![ctl_source] = cbof.ControlSource
                    End If
                  End If
                Case acCheckBox
                  If IsNull(cchk.ControlSource) = False Then
                    If Trim(cchk.ControlSource) <> vbNullString Then
                      ![ctl_source] = cchk.ControlSource
                    End If
                  End If
                Case acImage
                  ![ctl_sourceobject] = cimg.Picture
                Case acLabel
                  If IsNull(clbl.Caption) = False Then
                    If Trim(clbl.Caption) <> vbNullString Then
                      ![ctl_caption] = clbl.Caption
                    End If
                  End If
                Case acSubform
                  If IsNull(csub.SourceObject) = False Then
                    If Trim(csub.SourceObject) <> vbNullString Then
                      ![ctl_sourceobject] = csub.SourceObject
                    End If
                  End If
                Case acTextBox
                  If IsNull(ctxt.ControlSource) = False Then
                    If Trim(ctxt.ControlSource) <> vbNullString Then
                      ![ctl_source] = ctxt.ControlSource
                    End If
                  End If
                Case Else
                  'ccbx, ccmd, clbx, copb, copg, cpag, ctab
                  '![ctl_id_parent] =
                  '![ctl_parent] =
                End Select
                ![ctl_datemodified] = Now()
                .Update
              End With

              rst2.Edit

              Select Case ctl.ControlType
              Case acBoundObjectFrame
'ctlspec_backcolor
                rst2![ctlspec_backcolor] = cbof.BackColor
'ctlspec_backstyle
                rst2![ctlspec_backstyle] = cbof.BackStyle
'ctlspec_bordercolor
                rst2![ctlspec_bordercolor] = cbof.BorderColor
'ctlspec_borderlinestyle {deprecated}
'ctlspec_borderstyle
                rst2![ctlspec_borderstyle] = cbof.BorderStyle
'ctlspec_borderwidth
                rst2![ctlspec_borderwidth] = cbof.BorderWidth
'ctlspec_height
                rst2![ctlspec_height] = cbof.Height
              Case acCheckBox
                rst2![ctlspec_bordercolor] = cchk.BorderColor
                rst2![ctlspec_borderstyle] = cchk.BorderStyle
                rst2![ctlspec_borderwidth] = cchk.BorderWidth
                rst2![ctlspec_height] = cchk.Height
              Case acImage
                rst2![ctlspec_backcolor] = cimg.BackColor
                rst2![ctlspec_backstyle] = cimg.BackStyle
                rst2![ctlspec_bordercolor] = cimg.BorderColor
                rst2![ctlspec_borderstyle] = cimg.BorderStyle
                rst2![ctlspec_borderwidth] = cimg.BorderWidth
                rst2![ctlspec_height] = cimg.Height
              Case acLabel
                rst2![ctlspec_backcolor] = clbl.BackColor
                rst2![ctlspec_backstyle] = clbl.BackStyle
                rst2![ctlspec_bordercolor] = clbl.BorderColor
                rst2![ctlspec_borderstyle] = clbl.BorderStyle
                rst2![ctlspec_borderwidth] = clbl.BorderWidth
'ctlspec_forecolor
                rst2![ctlspec_forecolor] = clbl.ForeColor
                rst2![ctlspec_height] = clbl.Height
              Case acLine
                rst2![ctlspec_bordercolor] = clin.BorderColor
                rst2![ctlspec_borderstyle] = clin.BorderStyle
                rst2![ctlspec_borderwidth] = clin.BorderWidth
                rst2![ctlspec_height] = clin.Height
              Case acPageBreak
                ' ** No additional properties are saved at this time.
              Case acRectangle
                rst2![ctlspec_backcolor] = cbox.BackColor
                rst2![ctlspec_backstyle] = cbox.BackStyle
                rst2![ctlspec_bordercolor] = cbox.BorderColor
                rst2![ctlspec_borderstyle] = cbox.BorderStyle
                rst2![ctlspec_borderwidth] = cbox.BorderWidth
                rst2![ctlspec_height] = cbox.Height
              Case acSubform
                rst2![ctlspec_bordercolor] = csub.BorderColor
                rst2![ctlspec_borderstyle] = csub.BorderStyle
                rst2![ctlspec_borderwidth] = csub.BorderWidth
                rst2![ctlspec_height] = csub.Height
              Case acTextBox
                rst2![ctlspec_backcolor] = ctxt.BackColor
                rst2![ctlspec_backstyle] = ctxt.BackStyle
                rst2![ctlspec_bordercolor] = ctxt.BorderColor
                rst2![ctlspec_borderstyle] = ctxt.BorderStyle
                rst2![ctlspec_borderwidth] = ctxt.BorderWidth
                rst2![ctlspec_forecolor] = ctxt.ForeColor
'ctlspec_format
                If IsNull(ctxt.Format) = False Then
                  If Trim(ctxt.Format) <> vbNullString Then
                    rst2![ctlspec_format] = ctxt.Format
                  End If
                End If
                rst2![ctlspec_height] = ctxt.Height
              Case Else
                'ccbx, ccmd, clbx, copb, copg, cpag, ctab
                '![ctl_id_parent] =
                '![ctl_parent] =
              End Select

              For Each prp In ctl.Properties
                Select Case prp.Name
                Case "BorderLineStyle"
'ctlspec_borderlinestyle {deprecated}
                Case "BottomMargin"
'ctlspec_bottommargin
                  If IsNull(rst2![ctlspec_bottommargin]) = True Then
                    rst2![ctlspec_bottommargin] = prp.Value
                  Else
                    If prp.Value <> rst2![ctlspec_bottommargin] Then
                      rst2![ctlspec_bottommargin] = prp.Value
                    End If
                  End If
                Case "CanGrow"
'ctlspec_cangrow
                  If prp.Value <> rst2![ctlspec_cangrow] Then
                    rst2![ctlspec_cangrow] = prp.Value
                  End If
                Case "CanShrink"
'ctlspec_canshrink
                  If prp.Value <> rst2![ctlspec_canshrink] Then
                    rst2![ctlspec_canshrink] = prp.Value
                  End If
                Case "Caption"
'ctlspec_caption
                  Select Case IsNull(prp.Value)
                  Case True
                    If IsNull(rst2![ctlspec_caption]) = False Then
                      rst2![ctlspec_caption] = Null
                    End If
                  Case False
                    If Trim(prp.Value) <> vbNullString Then
                      Select Case IsNull(rst2![ctlspec_caption])
                      Case True
                        rst2![ctlspec_caption] = prp.Value
                      Case False
                        If prp.Value <> rst2![ctlspec_caption] Then
                          rst2![ctlspec_caption] = prp.Value
                        End If
                      End Select
                    Else
                      If IsNull(rst2![ctlspec_caption]) = False Then
                        rst2![ctlspec_caption] = Null
                      End If
                    End If
                  End Select
                Case "ControlSource"
'ctlspec_controlsource
                  Select Case IsNull(prp.Value)
                  Case True
                    If IsNull(rst2![ctlspec_controlsource]) = False Then
                      rst2![ctlspec_controlsource] = Null
                    End If
                  Case False
                    If Trim(prp.Value) <> vbNullString Then
                      Select Case IsNull(rst2![ctlspec_controlsource])
                      Case True
                        rst2![ctlspec_controlsource] = prp.Value
                      Case False
                        If prp.Value <> rst2![ctlspec_controlsource] Then
                          rst2![ctlspec_controlsource] = prp.Value
                        End If
                      End Select
                    Else
                      If IsNull(rst2![ctlspec_controlsource]) = False Then
                        rst2![ctlspec_controlsource] = Null
                      End If
                    End If
                  End Select
                Case "ControlType"
'ctlspec_controltype
                  Select Case IsNull(rst2![ctlspec_controltype])
                  Case True
                    rst2![ctlspec_controltype] = prp.Value
                  Case False
                    If prp.Value <> rst2![ctlspec_controltype] Then
                      rst2![ctlspec_controltype] = prp.Value
                    End If
                  End Select
                Case "DecimalPlaces"
'ctlspec_decimalplaces
                  Select Case IsNull(rst2![ctlspec_decimalplaces])
                  Case True
                    rst2![ctlspec_decimalplaces] = prp.Value
                  Case False
                    If prp.Value <> rst2![ctlspec_decimalplaces] Then
                      rst2![ctlspec_decimalplaces] = prp.Value
                    End If
                  End Select
                Case "DefaultValue"
'ctlspec_defaultvalue
                  If IsNull(prp.Value) = False Then
                    If Trim(prp.Value) <> vbNullString Then
                      rst2![ctlspec_defaultvalue] = prp.Value
                    Else
                      If IsNull(rst2![ctlspec_defaultvalue]) = False Then
                        rst2![ctlspec_defaultvalue] = Null
                      End If
                    End If
                  Else
                    If IsNull(rst2![ctlspec_defaultvalue]) = False Then
                      rst2![ctlspec_defaultvalue] = Null
                    End If
                  End If
                Case "EventProcPrefix"
'ctlspec_eventprocprefix
                  Select Case IsNull(rst2![ctlspec_eventprocprefix])
                  Case True
                    rst2![ctlspec_eventprocprefix] = prp.Value
                  Case False
                    If prp.Value <> rst2![ctlspec_eventprocprefix] Then
                      rst2![ctlspec_eventprocprefix] = prp.Value
                    End If
                  End Select
                Case "FontBold"
'ctlspec_fontbold
                  If prp.Value <> rst2![ctlspec_fontbold] Then
                    rst2![ctlspec_fontbold] = prp.Value
                  End If
                Case "FontItalic"
'ctlspec_fontitalic
                  If prp.Value <> rst2![ctlspec_fontitalic] Then
                    rst2![ctlspec_fontitalic] = prp.Value
                  End If
                Case "FontName"
'ctlspec_fontname
                  If IsNull(rst2![ctlspec_fontname]) = True Then
                    rst2![ctlspec_fontname] = prp.Value
                  Else
                    If prp.Value <> rst2![ctlspec_fontname] Then
                      rst2![ctlspec_fontname] = prp.Value
                    End If
                  End If
                Case "FontSize"
'ctlspec_fontsize
                  If IsNull(rst2![ctlspec_fontsize]) = True Then
                    rst2![ctlspec_fontsize] = prp.Value
                  Else
                    If prp.Value <> rst2![ctlspec_fontsize] Then
                      rst2![ctlspec_fontsize] = prp.Value
                    End If
                  End If
                Case "FontUnderline"
'ctlspec_fontunderline
                  If prp.Value <> rst2![ctlspec_fontunderline] Then
                    rst2![ctlspec_fontunderline] = prp.Value
                  End If
                Case "FontWeight"
'ctlspec_fontweight
                  If IsNull(rst2![ctlspec_fontweight]) = True Then
                    rst2![ctlspec_fontweight] = prp.Value
                  Else
                    If prp.Value <> rst2![ctlspec_fontweight] Then
                      rst2![ctlspec_fontweight] = prp.Value
                    End If
                  End If
                Case "Height"
'ctlspec_height
                  If IsNull(rst2![ctlspec_height]) = True Then
                    rst2![ctlspec_height] = prp.Value
                  Else
                    If rst2![ctlspec_height] <> prp.Value Then
                      rst2![ctlspec_height] = prp.Value
                    End If
                  End If
                Case "HideDuplicates"
'ctlspec_hideduplicates
                  If prp.Value <> rst2![ctlspec_hideduplicates] Then
                    rst2![ctlspec_hideduplicates] = prp.Value
                  End If
                Case "HyperlinkAddress"
'ctlspec_hyperlinkaddress
                  Select Case IsNull(prp.Value)
                  Case True
                    If IsNull(rst2![ctlspec_hyperlinkaddress]) = False Then
                      rst2![ctlspec_hyperlinkaddress] = Null
                    End If
                  Case False
                    If Trim(prp.Value) <> vbNullString Then
                      Select Case IsNull(rst2![ctlspec_hyperlinkaddress])
                      Case True
                        rst2![ctlspec_hyperlinkaddress] = prp.Value
                      Case False
                        If prp.Value <> rst2![ctlspec_hyperlinkaddress] Then
                          rst2![ctlspec_hyperlinkaddress] = prp.Value
                        End If
                      End Select
                    Else
                      If IsNull(rst2![ctlspec_hyperlinkaddress]) = False Then
                        rst2![ctlspec_hyperlinkaddress] = Null
                      End If
                    End If
                  End Select
                Case "HyperlinkSubaddress"
'ctlspec_hyperlinksubaddress
                  Select Case IsNull(prp.Value)
                  Case True
                    If IsNull(rst2![ctlspec_hyperlinksubaddress]) = False Then
                      rst2![ctlspec_hyperlinksubaddress] = Null
                    End If
                  Case False
                    If Trim(prp.Value) <> vbNullString Then
                      Select Case IsNull(rst2![ctlspec_hyperlinksubaddress])
                      Case True
                        rst2![ctlspec_hyperlinksubaddress] = prp.Value
                      Case False
                        If prp.Value <> rst2![ctlspec_hyperlinksubaddress] Then
                          rst2![ctlspec_hyperlinksubaddress] = prp.Value
                        End If
                      End Select
                    Else
                      If IsNull(rst2![ctlspec_hyperlinksubaddress]) = False Then
                        rst2![ctlspec_hyperlinksubaddress] = Null
                      End If
                    End If
                  End Select
                Case "ImageWidth"
'ctlspec_imagewidth
                  If IsNull(rst2![ctlspec_imagewidth]) = True Then
                    rst2![ctlspec_imagewidth] = prp.Value
                  Else
                    If rst2![ctlspec_imagewidth] <> prp.Value Then
                      rst2![ctlspec_imagewidth] = prp.Value
                    End If
                  End If
                Case "ImageHeight"
'ctlspec_imageheight
                  If IsNull(rst2![ctlspec_imageheight]) = True Then
                    rst2![ctlspec_imageheight] = prp.Value
                  Else
                    If rst2![ctlspec_imageheight] <> prp.Value Then
                      rst2![ctlspec_imageheight] = prp.Value
                    End If
                  End If
                Case "InputMask"
'ctlspec_inputmask
                  Select Case IsNull(prp.Value)
                  Case True
                    If IsNull(rst2![ctlspec_inputmask]) = False Then
                      rst2![ctlspec_inputmask] = Null
                    End If
                  Case False
                    If Trim(prp.Value) <> vbNullString Then
                      Select Case IsNull(rst2![ctlspec_inputmask])
                      Case True
                        rst2![ctlspec_inputmask] = prp.Value
                      Case False
                        If prp.Value <> rst2![ctlspec_inputmask] Then
                          rst2![ctlspec_inputmask] = prp.Value
                        End If
                      End Select
                    Else
                      If IsNull(rst2![ctlspec_inputmask]) = False Then
                        rst2![ctlspec_inputmask] = Null
                      End If
                    End If
                  End Select
                Case "IsHyperlink"
'ctlspec_ishyperlink
                  If prp.Value <> rst2![ctlspec_ishyperlink] Then
                    rst2![ctlspec_ishyperlink] = prp.Value
                  End If
                End Select
              Next

              rst2.Update
              rst3.Edit

              Select Case ctl.ControlType
              Case acBoundObjectFrame
'ctlspec_left
                rst3![ctlspec_left] = cbof.Left
'ctlspec_specialeffect
                rst3![ctlspec_specialeffect] = cbof.SpecialEffect
'ctlspec_top
                rst3![ctlspec_top] = cbof.Top
'ctlspec_visible
                rst3![ctlspec_visible] = cbof.Visible
'ctlspec_width
                rst3![ctlspec_width] = cbof.Width
              Case acCheckBox
                rst3![ctlspec_left] = cchk.Left
                rst3![ctlspec_specialeffect] = cchk.SpecialEffect
                rst3![ctlspec_top] = cchk.Top
                rst3![ctlspec_visible] = cchk.Visible
                rst3![ctlspec_width] = cchk.Width
              Case acImage
                rst3![ctlspec_left] = cimg.Left
                rst3![ctlspec_top] = cimg.Top
                rst3![ctlspec_visible] = cimg.Visible
                rst3![ctlspec_width] = cimg.Width
              Case acLabel
                rst3![ctlspec_left] = clbl.Left
                rst3![ctlspec_specialeffect] = clbl.SpecialEffect
                rst3![ctlspec_top] = clbl.Top
                rst3![ctlspec_visible] = clbl.Visible
                rst3![ctlspec_width] = clbl.Width
              Case acLine
                rst3![ctlspec_left] = clin.Left
                rst3![ctlspec_specialeffect] = clin.SpecialEffect
                rst3![ctlspec_top] = clin.Top
                rst3![ctlspec_visible] = clin.Visible
                rst3![ctlspec_width] = clin.Width
              Case acPageBreak
                rst3![ctlspec_top] = cpgb.Top
              Case acRectangle
                rst3![ctlspec_left] = cbox.Left
                rst3![ctlspec_specialeffect] = cbox.SpecialEffect
                rst3![ctlspec_top] = cbox.Top
                rst3![ctlspec_visible] = cbox.Visible
                rst3![ctlspec_width] = cbox.Width
              Case acSubform
                rst3![ctlspec_left] = csub.Left
                rst3![ctlspec_specialeffect] = csub.SpecialEffect
                rst3![ctlspec_top] = csub.Top
                rst3![ctlspec_visible] = csub.Visible
                rst3![ctlspec_width] = csub.Width
              Case acTextBox
                rst3![ctlspec_left] = ctxt.Left
                rst3![ctlspec_specialeffect] = ctxt.SpecialEffect
                rst3![ctlspec_top] = ctxt.Top
                rst3![ctlspec_visible] = ctxt.Visible
                rst3![ctlspec_width] = ctxt.Width
              Case Else
                'ccbx, ccmd, clbx, copb, copg, cpag, ctab
                '![ctl_id_parent] =
                '![ctl_parent] =
              End Select

              For Each prp In ctl.Properties
                Select Case prp.Name
                Case "KeyboardLanguage"
'ctlspec_keyboardlanguage
                  Select Case IsNull(rst3![ctlspec_keyboardlanguage])
                  Case True
                    rst3![ctlspec_keyboardlanguage] = prp.Value
                  Case False
                    If prp.Value <> rst3![ctlspec_keyboardlanguage] Then
                      rst3![ctlspec_keyboardlanguage] = prp.Value
                    End If
                  End Select
                Case "Left"
'ctlspec_left
                  If IsNull(rst3![ctlspec_left]) = True Then
                    rst3![ctlspec_left] = prp.Value
                  Else
                    If rst3![ctlspec_left] <> prp.Value Then
                      rst3![ctlspec_left] = prp.Value
                    End If
                  End If
                Case "LeftMargin"
'ctlspec_leftmargin
                  If IsNull(rst3![ctlspec_leftmargin]) = True Then
                    rst3![ctlspec_leftmargin] = prp.Value
                  Else
                    If prp.Value <> rst3![ctlspec_leftmargin] Then
                      rst3![ctlspec_leftmargin] = prp.Value
                    End If
                  End If
                Case "LineSlant"
'ctlspec_lineslant
                  Select Case IsNull(rst3![ctlspec_lineslant])
                  Case True
                    rst3![ctlspec_lineslant] = prp.Value
                  Case False
                    If prp.Value <> rst3![ctlspec_lineslant] Then
                      rst3![ctlspec_lineslant] = prp.Value
                    End If
                  End Select
                Case "LineSpacing"
'ctlspec_linespacing
                  If IsNull(rst3![ctlspec_linespacing]) = True Then
                    rst3![ctlspec_linespacing] = prp.Value
                  Else
                    If prp.Value <> rst3![ctlspec_linespacing] Then
                      rst3![ctlspec_linespacing] = prp.Value
                    End If
                  End If
                Case "NumeralShapes"
'ctlspec_numeralshapes
                  Select Case IsNull(rst3![ctlspec_numeralshapes])
                  Case True
                    rst3![ctlspec_numeralshapes] = prp.Value
                  Case False
                    If prp.Value <> rst3![ctlspec_numeralshapes] Then
                      rst3![ctlspec_numeralshapes] = prp.Value
                    End If
                  End Select
                Case "Picture"
'ctlspec_picture
                  If IsNull(rst3![ctlspec_picture]) = True Then
                    rst3![ctlspec_picture] = prp.Value
                  Else
                    If rst3![ctlspec_picture] <> prp.Value Then
                      rst3![ctlspec_picture] = prp.Value
                    End If
                  End If
                Case "PictureAlignment"
'ctlspec_picturealignment
                  If IsNull(rst3![ctlspec_picturealignment]) = True Then
                    rst3![ctlspec_picturealignment] = prp.Value
                  Else
                    If rst3![ctlspec_picturealignment] <> prp.Value Then
                      rst3![ctlspec_picturealignment] = prp.Value
                    End If
                  End If
                Case "PictureType"
'ctlspec_picturetype
                  If IsNull(rst3![ctlspec_picturetype]) = True Then
                    rst3![ctlspec_picturetype] = prp.Value
                  Else
                    If rst3![ctlspec_picturetype] <> prp.Value Then
                      rst3![ctlspec_picturetype] = prp.Value
                    End If
                  End If
                  Case "ReadingOrder"
'ctlspec_readingorder
                    Select Case IsNull(rst3![ctlspec_readingorder])
                    Case True
                      rst3![ctlspec_readingorder] = prp.Value
                    Case False
                      If prp.Value <> rst3![ctlspec_readingorder] Then
                        rst3![ctlspec_readingorder] = prp.Value
                      End If
                    End Select
                Case "RightMargin"
'ctlspec_rightmargin
                  If IsNull(rst3![ctlspec_rightmargin]) = True Then
                    rst3![ctlspec_rightmargin] = prp.Value
                  Else
                    If prp.Value <> rst3![ctlspec_rightmargin] Then
                      rst3![ctlspec_rightmargin] = prp.Value
                    End If
                  End If
                Case "RunningSum"
'ctlspec_runningsum
                  Select Case IsNull(rst3![ctlspec_runningsum])
                  Case True
                    rst3![ctlspec_runningsum] = prp.Value
                  Case False
                    If prp.Value <> rst3![ctlspec_runningsum] Then
                      rst3![ctlspec_runningsum] = prp.Value
                    End If
                  End Select
                Case "ScrollBarAlign"
'ctlspec_scrollbaralign
                  Select Case IsNull(rst3![ctlspec_scrollbaralign])
                  Case True
                    rst3![ctlspec_scrollbaralign] = prp.Value
                  Case False
                    If prp.Value <> rst3![ctlspec_scrollbaralign] Then
                      rst3![ctlspec_scrollbaralign] = prp.Value
                    End If
                  End Select
                Case "Section"
'ctlspec_section
                  Select Case IsNull(rst3![ctlspec_section])
                  Case True
                    rst3![ctlspec_section] = prp.Value
                  Case False
                    If prp.Value <> rst3![ctlspec_section] Then
                      rst3![ctlspec_section] = prp.Value
                    End If
                  End Select
                Case "ShortcutMenuBar"
'ctlspec_shortcutmenubar
                  Select Case IsNull(prp.Value)
                  Case True
                    If IsNull(rst3![ctlspec_shortcutmenubar]) = False Then
                      rst3![ctlspec_shortcutmenubar] = Null
                    End If
                  Case False
                    If Trim(prp.Value) <> vbNullString Then
                      Select Case IsNull(rst3![ctlspec_shortcutmenubar])
                      Case True
                        rst3![ctlspec_shortcutmenubar] = prp.Value
                      Case False
                        If prp.Value <> rst3![ctlspec_shortcutmenubar] Then
                          rst3![ctlspec_shortcutmenubar] = prp.Value
                        End If
                      End Select
                    Else
                      If IsNull(rst3![ctlspec_shortcutmenubar]) = False Then
                        rst3![ctlspec_shortcutmenubar] = Null
                      End If
                    End If
                  End Select
                Case "SizeMode"
'ctlspec_sizemode
                  If IsNull(rst3![ctlspec_sizemode]) = True Then
                    rst3![ctlspec_sizemode] = prp.Value
                  Else
                    If rst3![ctlspec_sizemode] <> prp.Value Then
                      rst3![ctlspec_sizemode] = prp.Value
                    End If
                  End If
                Case "SmartTags"
'ctlspec_smarttags
                  Select Case IsNull(prp.Value)
                  Case True
                    If IsNull(rst3![ctlspec_smarttags]) = False Then
                      rst3![ctlspec_smarttags] = Null
                    End If
                  Case False
                    If Trim(prp.Value) <> vbNullString Then
                      Select Case IsNull(rst3![ctlspec_smarttags])
                      Case True
                        rst3![ctlspec_smarttags] = prp.Value
                      Case False
                        If prp.Value <> rst3![ctlspec_smarttags] Then
                          rst3![ctlspec_smarttags] = prp.Value
                        End If
                      End Select
                    Else
                      If IsNull(rst3![ctlspec_smarttags]) = False Then
                        rst3![ctlspec_smarttags] = Null
                      End If
                    End If
                  End Select
                Case "Tag"
'ctlspec_tag
                  If IsNull(prp.Value) = True Then
                    If IsNull(rst3![ctlspec_tag]) = False Then
                      rst3![ctlspec_tag] = Null
                    End If
                  Else
                    If Trim(prp.Value) = vbNullString Then
                      If IsNull(rst3![ctlspec_tag]) = False Then
                        rst3![ctlspec_tag] = Null
                      End If
                    Else
                      If IsNull(rst3![ctlspec_tag]) = True Then
                        rst3![ctlspec_tag] = prp.Value
                      Else
                        If rst3![ctlspec_tag] <> prp.Value Then
                          rst3![ctlspec_tag] = prp.Value
                        End If
                      End If
                    End If
                  End If
                Case "TextAlign"
'ctlspec_textalign
                  If IsNull(rst3![ctlspec_textalign]) = True Then
                    rst3![ctlspec_textalign] = prp.Value
                  Else
                    If prp.Value <> rst3![ctlspec_textalign] Then
                      rst3![ctlspec_textalign] = prp.Value
                    End If
                  End If
                Case "TextFontCharSet"
'ctlspec_textfontcharset
                  Select Case IsNull(rst3![ctlspec_textfontcharset])
                  Case True
                    rst3![ctlspec_textfontcharset] = prp.Value
                  Case False
                    If prp.Value <> rst3![ctlspec_textfontcharset] Then
                      rst3![ctlspec_textfontcharset] = prp.Value
                    End If
                  End Select
                Case "Top"
'ctlspec_top
                  If IsNull(rst3![ctlspec_top]) = True Then
                    rst3![ctlspec_top] = prp.Value
                  Else
                    If rst3![ctlspec_top] <> prp.Value Then
                      rst3![ctlspec_top] = prp.Value
                    End If
                  End If
                Case "TopMargin"
'ctlspec_topmargin
                  If IsNull(rst3![ctlspec_topmargin]) = True Then
                    rst3![ctlspec_topmargin] = prp.Value
                  Else
                    If prp.Value <> rst3![ctlspec_topmargin] Then
                      rst3![ctlspec_topmargin] = prp.Value
                    End If
                  End If
                Case "Vertical"
'ctlspec_vertical
                  If prp.Value <> rst3![ctlspec_vertical] Then
                    rst3![ctlspec_vertical] = prp.Value
                  End If
                Case "Visible"
'ctlspec_visible
                  If rst3![ctlspec_visible] <> prp.Value Then
                    rst3![ctlspec_visible] = prp.Value
                  End If
                Case "Width"
'ctlspec_width
                  If IsNull(rst3![ctlspec_width]) = True Then
                    rst3![ctlspec_width] = prp.Value
                  Else
                    If rst3![ctlspec_width] <> prp.Value Then
                      rst3![ctlspec_width] = prp.Value
                    End If
                  End If
                End Select
              Next

              rst3![ctlspec_datemodified] = Now()
              rst3.Update

            End If
          End With  ' ** This Control: ctl.
          DoEvents
        Next  ' ** For each Control: ctl.
      End With  ' ** This Form: rpt
      DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX), acSaveNo
      DoEvents
    Next  ' ** For each Form: lngX.
    rst1.Close

'TBL: tblReport_Control_Specification_A  FLDS: 35
'dbs_id
'rpt_id
'ctl_id
'ctlspec_id
'ctlspec_backcolor
'ctlspec_backstyle
'ctlspec_bordercolor
'ctlspec_borderlinestyle
'ctlspec_borderstyle
'ctlspec_borderwidth
'ctlspec_bottommargin
'ctlspec_cangrow
'ctlspec_canshrink
'ctlspec_caption
'ctlspec_controlsource
'ctlspec_controltype
'ctlspec_decimalplaces
'ctlspec_defaultvalue
'ctlspec_eventprocprefix
'ctlspec_fontbold
'ctlspec_fontitalic
'ctlspec_fontname
'ctlspec_fontsize
'ctlspec_fontunderline
'ctlspec_fontweight
'ctlspec_forecolor
'ctlspec_format
'ctlspec_height
'ctlspec_hideduplicates
'ctlspec_hyperlinkaddress
'ctlspec_hyperlinksubaddress
'ctlspec_imagewidth
'ctlspec_imageheight
'ctlspec_inputmask
'ctlspec_ishyperlink
'DONE!  Tbl_Fld_List()
'TBL: tblReport_Control_Specification_B  FLDS: 31
'dbs_id
'rpt_id
'ctl_id
'ctlspec_id
'ctlspec_keyboardlanguage
'ctlspec_left
'ctlspec_leftmargin
'ctlspec_lineslant
'ctlspec_linespacing
'ctlspec_numeralshapes
'ctlspec_picture
'ctlspec_picturealignment
'ctlspec_picturetype
'ctlspec_readingorder
'ctlspec_rightmargin
'ctlspec_runningsum
'ctlspec_scrollbaralign
'ctlspec_section
'ctlspec_shortcutmenubar
'ctlspec_sizemode
'ctlspec_smarttags
'ctlspec_specialeffect
'ctlspec_tag
'ctlspec_textalign
'ctlspec_textfontcharset
'ctlspec_top
'ctlspec_topmargin
'ctlspec_vertical
'ctlspec_visible
'ctlspec_width
'ctlspec_datemodified
'DONE!  Tbl_Fld_List()

    ' ** 4. Check for deleted or changed controls.

    ' ***************************************************************
    'dblPB_ThisStep = 4#
    'dblPB_ThisWidth = 0#
    'For dblZ = 1# To (dblPB_ThisStep - 1#)
    '  ' ** Assemble the weighted widths up to, but not including, this width.
    '  dblPB_ThisWidth = (dblPB_ThisWidth + arr_dblPB_ThisIncr(dblZ))
    'Next
    'dblPB_StepSubs = 0#
    'dblPB_ThisIncrSub = 0#
    'dblPB_StepSubs2 = 0#
    'dblPB_ThisIncrSub2 = 0#
    'dblPB_ThisStepSub = 0#
    'dblPB_ThisStepSub2 = 0#
    'ctlPB.Width = dblPB_ThisWidth
    'ctlPBPct2.Width = dblPB_ThisWidth + 30&
    'prpStat1 = "Checking for changed or deleted controls . . ."
    'strPB_ThisPct = Format$((dblPB_ThisWidth / dblPB_Width), "##0%")
    'prpPct1 = strSp & strPB_ThisPct
    'prpPct2 = strSp & strPB_ThisPct
    'ctlS1Cnt.Visible = False: ctlS1Of.Visible = False: ctlS1Tot.Visible = False
    'ctlS2Cnt.Visible = False: ctlS2Of.Visible = False: ctlS2Tot.Visible = False
    'DoEvents
    ' ***************************************************************

    For lngX = 0& To (lngCtls - 1&)
      If arr_varCtl(C_FND, lngX) = False And arr_varCtl(C_DID, lngX) = lngThisDbsID Then
        For lngY = 0& To (lngRpts - 1&)
          If arr_varRpt(R_RID, lngY) = arr_varCtl(C_RID, lngX) Then
            ' ** Delete tblReport_Control, by specified [ctlid].
            Set qdf = .QueryDefs("zz_qry_Report_Control_01")
            With qdf.Parameters
              ![ctlid] = arr_varCtl(C_CID, lngX)
            End With
            qdf.Execute
            Set qdf = Nothing
            Exit For
          End If
        Next
      End If
    Next

    .Close
  End With

  DoCmd.Hourglass False

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  ' ***************************************************************
  'DoCmd.SelectObject acForm, strFrmStat, False
  'ctlPB.Width = dblPB_Width
  'ctlPBPct2.Width = dblPB_Width
  'prpStat1 = "Finished . . ."
  'prpStat2 = vbNullString
  'strPB_ThisPct = Format$(1#, "##0%")
  'prpPct1 = strSp & strPB_ThisPct
  'prpPct2 = strSp & strPB_ThisPct
  'DoEvents
  ' ***************************************************************

  Set clbl = Nothing
  Set cbox = Nothing
  Set clin = Nothing
  Set cchk = Nothing
  Set ctxt = Nothing
  Set csub = Nothing
  Set cpgb = Nothing
  Set cimg = Nothing
  Set dbs = Nothing
  Set qdf = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst2 = Nothing
  Set rpt = Nothing
  Set prp = Nothing
  Set ctl = Nothing
  'Set frmStat = Nothing
  'Set ctlPB = Nothing
  'Set ctlStat1 = Nothing
  'Set ctlStat2 = Nothing
  'Set ctlPBPct1 = Nothing
  'Set ctlPBPct2 = Nothing
  'Set prpStat1 = Nothing
  'Set prpStat2 = Nothing
  'Set prpPct1 = Nothing
  'Set prpPct2 = Nothing
  'Set ctlS1Cnt = Nothing
  'Set ctlS1Tot = Nothing
  'Set ctlS2Cnt = Nothing
  'Set ctlS2Tot = Nothing
  'Set prpS1Cnt = Nothing
  'Set prpS1Tot = Nothing
  'Set prpS2Cnt = Nothing
  'Set prpS2Tot = Nothing

  Rpt_Ctl_Doc = blnRetValx

End Function

Private Function Rpt_RecSrc_Doc() As Boolean
' ** Document all report Record Sources to tblReport_RecordSource.
' ** Called by:
' **   QuikRptDoc(), Above

  Const THIS_PROC As String = "Rpt_RecSrc_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, rpt As Report
  Dim lngRpts As Long, arr_varRpt As Variant
  Dim strRecSrc As String, lngQryTblTypeType As Long
  Dim lngThisDbsID As Long
  Dim blnIsTbl As Boolean, blnIsQry As Boolean, blnIsSQL As Boolean
  Dim lngLen As Long
  Dim lngX As Long

  ' ** Array: arr_varRpt().
  Const R_DID  As Integer = 0
  Const R_DNAM As Integer = 1
  Const R_RID  As Integer = 2
  Const R_RNAM As Integer = 3

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  'If IsLoaded("zz_frmStatus", acForm) = True Then  ' ** Module Function: modFileUtilities.
  '  DoCmd.Close acForm, "zz_frmStatus"
  'End If

  Set dbs = CurrentDb
  With dbs

    Set rst2 = .OpenRecordset("tblReport_RecordSource", dbOpenDynaset, dbConsistent)

    ' ** tblReport, just needed fields, by specified CurrentAppName().
    Set qdf = .QueryDefs("zz_qry_Report_02")
    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngRpts = .RecordCount
      .MoveFirst
      arr_varRpt = .GetRows(lngRpts)
      ' *********************************************
      ' ** Array: arr_varRpt()
      ' **
      ' **   Field  Element  Name        Constant
      ' **   =====  =======  ==========  ==========
      ' **     1       0     dbs_id      R_DID
      ' **     2       1     dbs_name    R_DNAM
      ' **     3       2     rpt_id      R_RID
      ' **     4       3     rpt_name    R_RNAM
      ' **
      ' *********************************************
      .Close
    End With  ' ** rst1.

    lngLen = 0&
    For lngX = 0& To (lngRpts - 1&)

      lngQryTblTypeType = 0&: strRecSrc = vbNullString

      DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acViewDesign, , , acHidden
      Reports(arr_varRpt(R_RNAM, lngX)).Visible = False
      Set rpt = Reports(arr_varRpt(R_RNAM, lngX))
      With rpt
        If IsNull(.RecordSource) = False Then
          If .RecordSource <> vbNullString Then
            strRecSrc = .RecordSource
          End If
        End If
      End With
      DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX), acSaveNo

      If strRecSrc <> vbNullString Then
        blnIsTbl = CBool(IIf(Left$(strRecSrc, 3) <> "Qry" And Left$(strRecSrc, 4) <> "_Qry" And _
          strRecSrc <> "GLCodeLkpMgr" And Left$(strRecSrc, 6) <> "SELECT", True, False))
        If blnIsTbl = True Then
          lngQryTblTypeType = acTable
        Else
          blnIsQry = CBool(IIf(Left$(strRecSrc, 3) = "Qry" Or Left$(strRecSrc, 4) = "_Qry" Or strRecSrc = "GLCodeLkpMgr", True, False))
          If blnIsQry = True Then
            lngQryTblTypeType = acQuery
          Else
            blnIsSQL = CBool(IIf(Left$(strRecSrc, 6) = "SELECT", True, False))
            If blnIsSQL = True Then
              lngQryTblTypeType = acSQL
            Else
              Stop
            End If
          End If
        End If
      Else
        lngQryTblTypeType = acSQL
      End If

      With rst2
        .FindFirst "[rpt_id] = " & CStr(arr_varRpt(R_RID, lngX))
        If .NoMatch = True Then
          .AddNew
          ![dbs_id] = arr_varRpt(R_DID, lngX)
          ![rpt_id] = arr_varRpt(R_RID, lngX)
          ![recsrc_recordsource] = strRecSrc
          ![qrytbltype_type] = lngQryTblTypeType
          ![recsrc_datemodified] = Now()
          .Update
        Else
          .Edit
          ![recsrc_recordsource] = strRecSrc
          ![qrytbltype_type] = lngQryTblTypeType
          ![recsrc_datemodified] = Now()
          .Update
        End If
      End With

      If Len(strRecSrc) > lngLen Then lngLen = Len(strRecSrc)
    Next

    rst2.Close

    .Close
  End With

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rpt = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_RecSrc_Doc = blnRetValx

End Function

Private Function Rpt_Subreport_Doc() As Boolean
' ** Document all report subforms/subreports to tblReport_Subform.
' ** Called by:
' **   QuikRptDoc(), Above

  Const THIS_PROC As String = "Rpt_Subreport_Doc"

  Dim dbs As DAO.Database, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim cntr As Container, doc As Document, rpt As Report, ctl As Control
  Dim lngSubs As Long, arr_varSub() As Variant
  Dim lngRptID As Long, lngCtlID As Long
  Dim lngThisDbsID As Long
  Dim strSource As String
  Dim lngX As Long, lngE As Long

  Const SUB_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const S_DID     As Integer = 0
  Const S_DNAM    As Integer = 1
  Const S_RID     As Integer = 2
  Const S_RNAM    As Integer = 3
  Const S_CID     As Integer = 4
  Const S_CNAM    As Integer = 5
  Const S_SDID    As Integer = 6
  Const S_SRID    As Integer = 7
  Const S_SNAM    As Integer = 8
  Const S_LNKMAST As Integer = 9
  Const S_LNKCHLD As Integer = 10

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    Set rst1 = .OpenRecordset("tblReport", dbOpenDynaset, dbConsistent)
    If rst1.BOF = True And rst1.EOF = True Then
      ' ** No records?! Run rpt_Doc(), below.
      Beep
      blnRetValx = False
    End If

    If blnRetValx = True Then

      Set rst2 = .OpenRecordset("tblReport_Subform", dbOpenDynaset, dbConsistent)

      Set cntr = .Containers("Reports")
      With cntr
        For Each doc In .Documents
          With doc

            lngSubs = 0&
            ReDim arr_varSub(SUB_ELEMS, 0)
            ' ************************************************************
            ' ** Array: arr_varSub()
            ' **
            ' **   Field  Element  Name                       Constant
            ' **   =====  =======  =========================  ==========
            ' **     1       0     dbs_id                     S_DID
            ' **     2       1     dbs_name                   S_DNAM
            ' **     1       0     rpt_id                     S_RID
            ' **     2       1     rpt_name                   S_RNAM
            ' **     3       2     ctl_id                     S_CID
            ' **     4       3     ctl_name                   S_CNAM
            ' **     5       4     dbs_id_sub                 S_SDID
            ' **     5       4     rpt_id_sub                 S_SRID
            ' **     6       5     rpt_name2                  S_SNAM
            ' **     7       6     rptsub_linkmasterfields    S_LNKMAST
            ' **     8       7     rptsub_linkchildfields     S_LNKCHLD
            ' **
            ' ************************************************************

            DoCmd.OpenReport .Name, acViewDesign, , , acHidden
            Reports(.Name).Visible = False
            Set rpt = Reports(.Name)
            With rpt
              For Each ctl In .Controls
                With ctl
                  Select Case .ControlType
                  Case acSubform
                    With rst1
                      .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [rpt_name] = '" & doc.Name & "'"
                      If .NoMatch = False Then
                        .Edit
                        ![rpt_hassub] = True
                        lngSubs = lngSubs + 1&
                        lngE = lngSubs - 1&
                        ReDim Preserve arr_varSub(SUB_ELEMS, lngE)
                        arr_varSub(S_DID, lngE) = ![dbs_id]
                        arr_varSub(S_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
                        arr_varSub(S_RID, lngE) = ![rpt_id]
                        arr_varSub(S_RNAM, lngE) = rpt.Name
                        lngCtlID = DLookup("[ctl_id]", "tblReport_Control", "[dbs_id] = " & CStr(![dbs_id]) & " And " & _
                          "[rpt_id] = " & CStr(![rpt_id]) & " And [ctl_name] = '" & ctl.Name & "'")
                        arr_varSub(S_CID, lngE) = lngCtlID
                        arr_varSub(S_CNAM, lngE) = ctl.Name
                        If ctl.SourceObject = vbNullString Then
                          Debug.Print "'EMPTY SUB: " & rpt.Name & "  SUB: " & ctl.Name
                          arr_varSub(S_SDID, lngE) = 0&
                          arr_varSub(S_SRID, lngE) = 0&
                          arr_varSub(S_SNAM, lngE) = "{null}"
                          arr_varSub(S_LNKMAST, lngE) = Null
                          arr_varSub(S_LNKCHLD, lngE) = Null
                        Else
                          strSource = ctl.SourceObject
                          If InStr(strSource, ".") > 0 Then strSource = Mid$(strSource, (InStr(strSource, ".") + 1))
                          lngRptID = DLookup("[rpt_id]", "tblReport", "[dbs_id] = " & CStr(![dbs_id]) & " And [rpt_name] = '" & strSource & "'")
                          arr_varSub(S_SDID, lngE) = ![dbs_id]
                          arr_varSub(S_SRID, lngE) = lngRptID
                          arr_varSub(S_SNAM, lngE) = strSource
                          If IsNull(ctl.LinkMasterFields) = False Then
                            If ctl.LinkMasterFields <> vbNullString Then
                              arr_varSub(S_LNKMAST, lngE) = ctl.LinkMasterFields
                              arr_varSub(S_LNKCHLD, lngE) = ctl.LinkChildFields
                            Else
                              arr_varSub(S_LNKMAST, lngE) = Null
                              arr_varSub(S_LNKCHLD, lngE) = Null
                            End If
                          Else
                            arr_varSub(S_LNKMAST, lngE) = Null
                            arr_varSub(S_LNKCHLD, lngE) = Null
                          End If
                        End If
                        .Update
                      Else
                        Stop
                      End If
                    End With
                  Case Else
                    ' ** Don't care right now.
                  End Select
                End With  ' ** This Control: ctl.
              Next  ' ** For each Control: ctl.
            End With  ' ** This Report: rpt.
            DoCmd.Close acReport, .Name, acSaveNo

            If lngSubs > 0& Then

              For lngX = 0& To (lngSubs - 1&)
                If arr_varSub(S_SNAM, lngX) <> "{null}" Then
                  With rst1
                    .FindFirst "[dbs_id] = " & arr_varSub(S_SDID, lngX) & " And [rpt_id] = " & arr_varSub(S_SRID, lngX)
                    If .NoMatch = False Then
                      .Edit
                      ![rpt_issub] = True
                      ![rpt_datemodified] = Now()
                      .Update
                    Else
                      Stop
                    End If
                  End With
                End If
              Next
              With rst1
                .FindFirst "[dbs_id] = " & CStr(arr_varSub(S_DID, 0)) & " And [rpt_id] = " & CStr(arr_varSub(S_RID, 0))
                If .NoMatch = False Then
                  .Edit
                  ![rpt_subs] = lngSubs
                  ![rpt_datemodified] = Now()
                  .Update
                Else
                  Stop
                End If
              End With

              With rst2
                For lngX = 0& To (lngSubs - 1&)
                  If arr_varSub(S_SRID, lngX) > 0& Then
                    .FindFirst "[dbs_id] = " & CStr(arr_varSub(S_DID, lngX)) & " And [rpt_id] = " & CStr(arr_varSub(S_RID, lngX)) & _
                      " And [ctl_id] = " & CStr(arr_varSub(S_CID, lngX)) & _
                      " And [dbs_id_sub] = " & CStr(arr_varSub(S_SDID, lngX)) & " And [rpt_id_sub] = " & CStr(arr_varSub(S_SRID, lngX))
                    If .NoMatch = False Then
                      If IsNull(arr_varSub(S_LNKMAST, lngX)) = False Then
                        If IsNull(![rptsub_linkmasterfields]) = False Then
                          If arr_varSub(S_LNKMAST, lngX) <> ![rptsub_linkmasterfields] Then
                            If InStr(![rptsub_linkmasterfields], "{or}") = 0 Then
                              .Edit
                              ![rptsub_linkmasterfields] = arr_varSub(S_LNKMAST, lngX)
                              ![rptsub_datemodified] = Now()
                              .Update
                            Else
                              Debug.Print "'OR SRC: '" & arr_varSub(S_RNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'  SRC: '" & _
                                arr_varSub(S_LNKMAST, lngX) & "'"
                            End If
                          End If
                        Else
                          .Edit
                          ![rptsub_linkmasterfields] = arr_varSub(S_LNKMAST, lngX)
                          ![rptsub_datemodified] = Now()
                          .Update
                        End If
                      Else
                        If IsNull(![rptsub_linkmasterfields]) = False Then
                          Debug.Print "'rpt NULL, TBL NOT: '" & arr_varSub(S_RNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'"
                        End If
                      End If
                      If IsNull(arr_varSub(S_LNKCHLD, lngX)) = False Then
                        If IsNull(![rptsub_linkchildfields]) = False Then
                          If arr_varSub(S_LNKCHLD, lngX) <> ![rptsub_linkchildfields] Then
                            If InStr(![rptsub_linkchildfields], "{or}") = 0 Then
                              .Edit
                              ![rptsub_linkchildfields] = arr_varSub(S_LNKCHLD, lngX)
                              ![rptsub_datemodified] = Now()
                              .Update
                            Else
                              Debug.Print "'OR SRC: '" & arr_varSub(S_RNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'  SRC: '" & _
                                arr_varSub(S_LNKCHLD, lngX) & "'"
                            End If
                          End If
                        Else
                          .Edit
                          ![rptsub_linkchildfields] = arr_varSub(S_LNKCHLD, lngX)
                          ![rptsub_datemodified] = Now()
                          .Update
                        End If
                      Else
                        If IsNull(![rptsub_linkchildfields]) = False Then
                          Debug.Print "'RPT NULL, TBL NOT: '" & arr_varSub(S_RNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'"
                        End If
                      End If
                    Else
                      .AddNew
                      ![dbs_id] = arr_varSub(S_DID, lngX)
                      ![rpt_id] = arr_varSub(S_RID, lngX)
                      ![ctl_id] = arr_varSub(S_CID, lngX)
                      ![dbs_id_sub] = arr_varSub(S_SDID, lngX)
                      ![rpt_id_sub] = arr_varSub(S_SRID, lngX)
                      If IsNull(arr_varSub(S_LNKMAST, lngX)) = False Then
                        ![rptsub_linkmasterfields] = arr_varSub(S_LNKMAST, lngX)
                        ![rptsub_linkchildfields] = arr_varSub(S_LNKCHLD, lngX)
                      End If
                      ![rptsub_datemodified] = Now()
                      .Update
                    End If
                  End If
                Next
              End With

            End If

          End With  ' ** This Document: doc.
        Next  ' ** For each Document: doc.
      End With  ' ** This Container: cntr.

      rst2.Close

    End If
    rst1.Close

    .Close
  End With  ' ** dbs.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set ctl = Nothing
  Set rpt = Nothing
  Set doc = Nothing
  Set cntr = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set dbs = Nothing

  Rpt_Subreport_Doc = blnRetValx

End Function

Private Function Rpt_Grp_Doc() As Boolean
' ** Called by:
' **   QuikRptDoc(), Above

  Const THIS_PROC As String = "Rpt_Grp_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim prj As CurrentProject, rptao As AccessObject, rpt As Access.Report
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim lngGroups As Long, arr_varGroup() As Variant, arr_varTmpGrp As Variant
  Dim strName As String, varControlSource As Variant
  Dim lngThisDbsID As Long
  Dim blnAdd As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const R_DID  As Integer = 0
  Const R_DNAM As Integer = 1
  Const R_RID  As Integer = 2
  Const R_RNAM As Integer = 3
  Const R_GRPS As Integer = 4
  Const R_ARR  As Integer = 5

  ' ** Array: arr_varGroup().
  Const G_ELEMS As Integer = 7  ' ** Array's first-element UBound().
  Const G_IDX  As Integer = 0
  Const G_SRC  As Integer = 1
  Const G_SORT As Integer = 2
  Const G_HEAD As Integer = 3
  Const G_FOOT As Integer = 4
  Const G_ON   As Integer = 5
  Const G_INT  As Integer = 6
  Const G_KEEP As Integer = 7

  Const R_MAX_GRPS As Long = 10  ' ** 10 groups, 0 - 9.

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set prj = Application.CurrentProject
  With prj
    lngRpts = .AllReports.Count
    lngX = -1&
    ReDim arr_varRpt(R_ELEMS, (lngRpts - 1&))
    For Each rptao In .AllReports
      strName = vbNullString
      With rptao
        lngX = lngX + 1&
        strName = .Name
        arr_varRpt(R_DID, lngX) = lngThisDbsID
        arr_varRpt(R_DNAM, lngX) = CurrentAppName  ' ** Module Function: modFileUtilities.
        arr_varRpt(R_RID, lngX) = CLng(0)
        arr_varRpt(R_RNAM, lngX) = strName
        arr_varRpt(R_GRPS, lngX) = CLng(0)
        arr_varRpt(R_ARR, lngX) = Empty
      End With
    Next
  End With

  For lngX = 0& To (lngRpts - 1&)
    strName = arr_varRpt(R_RNAM, lngX)
    DoCmd.OpenReport strName, acViewDesign, , , acHidden
    Set rpt = Reports(strName)
    With rpt
      lngGroups = 0&
      ReDim arr_varGroup(G_ELEMS, 0)
      For lngY = 0& To (R_MAX_GRPS - 1&)
On Error Resume Next
        varControlSource = .GroupLevel(lngY).ControlSource
        If ERR.Number = 0 Then
On Error GoTo 0
          lngGroups = lngGroups + 1&
          lngE = lngGroups - 1&
          ReDim Preserve arr_varGroup(G_ELEMS, lngE)
          arr_varGroup(G_IDX, lngE) = lngY
          arr_varGroup(G_SRC, lngE) = varControlSource
          arr_varGroup(G_SORT, lngE) = .GroupLevel(lngY).sortOrder
          arr_varGroup(G_HEAD, lngE) = .GroupLevel(lngY).GroupHeader
          arr_varGroup(G_FOOT, lngE) = .GroupLevel(lngY).GroupFooter
          arr_varGroup(G_ON, lngE) = .GroupLevel(lngY).GroupOn
          arr_varGroup(G_INT, lngE) = .GroupLevel(lngY).GroupInterval
          arr_varGroup(G_KEEP, lngE) = .GroupLevel(lngY).KeepTogether
        Else
On Error GoTo 0
          Exit For
        End If
      Next
      arr_varRpt(R_GRPS, lngX) = lngGroups
      If lngGroups > 0& Then
        arr_varRpt(R_ARR, lngX) = arr_varGroup
      End If
    End With
    DoCmd.Close acReport, strName, acSaveNo
    Set rpt = Nothing
  Next

  Set dbs = CurrentDb
  With dbs

    Set rst = .OpenRecordset("tblReport", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveFirst
      For lngX = 0& To (lngRpts - 1&)
        .FindFirst "[dbs_id] = " & CStr(arr_varRpt(R_DID, lngX)) & " And [rpt_name] = '" & arr_varRpt(R_RNAM, lngX) & "'"
        If .NoMatch = False Then
          arr_varRpt(R_RID, lngX) = ![rpt_id]
        Else
          Stop
        End If
      Next
      .Close
    End With

    Set rst = .OpenRecordset("tblReport_Group", dbOpenDynaset, dbConsistent)
    With rst
      For lngX = 0& To (lngRpts - 1&)
        lngGroups = arr_varRpt(R_GRPS, lngX)
        If lngGroups > 0& Then
          arr_varTmpGrp = arr_varRpt(R_ARR, lngX)
          For lngY = 0& To (lngGroups - 1&)
            blnAdd = False
            If .BOF = True And .EOF = True Then
              blnAdd = True
            Else
              .FindFirst "[dbs_id] = " & CStr(arr_varRpt(R_DID, lngX)) & " And [rpt_id] = " & CStr(arr_varRpt(R_RID, lngX)) & " And " & _
                "[rptgrp_grouplevel] = " & CStr(arr_varTmpGrp(G_IDX, lngY))
              If .NoMatch = False Then
                If ![rptgrp_controlsource] <> arr_varTmpGrp(G_SRC, lngY) Then
                  .Edit
                  ![rptgrp_controlsource] = arr_varTmpGrp(G_SRC, lngY)
                  ![rptgrp_datemodified] = Now()
                  .Update
                End If
                If ![rptgrp_sortorder] <> arr_varTmpGrp(G_SORT, lngY) Then
                  .Edit
                  ![rptgrp_sortorder] = arr_varTmpGrp(G_SORT, lngY)
                  ![rptgrp_datemodified] = Now()
                  .Update
                End If
                If ![rptgrp_groupheader] <> arr_varTmpGrp(G_HEAD, lngY) Then
                  .Edit
                  ![rptgrp_groupheader] = arr_varTmpGrp(G_HEAD, lngY)
                  ![rptgrp_datemodified] = Now()
                  .Update
                End If
                If ![rptgrp_groupfooter] <> arr_varTmpGrp(G_FOOT, lngY) Then
                  .Edit
                  ![rptgrp_groupfooter] = arr_varTmpGrp(G_FOOT, lngY)
                  ![rptgrp_datemodified] = Now()
                  .Update
                End If
                If ![rptgrp_groupon] <> arr_varTmpGrp(G_ON, lngY) Then
                  .Edit
                  ![rptgrp_groupon] = arr_varTmpGrp(G_ON, lngY)
                  ![rptgrp_datemodified] = Now()
                  .Update
                End If
                If ![rptgrp_groupinterval] <> arr_varTmpGrp(G_INT, lngY) Then
                  .Edit
                  ![rptgrp_groupinterval] = arr_varTmpGrp(G_INT, lngY)
                  ![rptgrp_datemodified] = Now()
                  .Update
                End If
                If ![rptgrp_keeptogether] <> arr_varTmpGrp(G_KEEP, lngY) Then
                  .Edit
                  ![rptgrp_keeptogether] = arr_varTmpGrp(G_KEEP, lngY)
                  ![rptgrp_datemodified] = Now()
                  .Update
                End If
              Else
                blnAdd = True
              End If
            End If
            If blnAdd = True Then
              .AddNew
              ![dbs_id] = arr_varRpt(R_DID, lngX)
              ![rpt_id] = arr_varRpt(R_RID, lngX)
              ![rptgrp_grouplevel] = arr_varTmpGrp(G_IDX, lngY)
              ![rptgrp_controlsource] = arr_varTmpGrp(G_SRC, lngY)
              ![datatype_db_type] = dbText  'Default till I work this out!
              ![rptgrp_sortorder] = arr_varTmpGrp(G_SORT, lngY)
              ![rptgrp_groupheader] = arr_varTmpGrp(G_HEAD, lngY)
              ![rptgrp_groupfooter] = arr_varTmpGrp(G_FOOT, lngY)
              ![rptgrp_groupon] = arr_varTmpGrp(G_ON, lngY)
              ![rptgrp_groupinterval] = arr_varTmpGrp(G_INT, lngY)
              ![rptgrp_keeptogether] = arr_varTmpGrp(G_KEEP, lngY)
              ![rptgrp_datemodified] = Now()
              .Update
            End If
          Next  ' ** lngY.
        End If  ' ** lngGroups.
      Next  ' ** lngX.
      .Close
    End With

    ' ** zz_qry_Report_Group_01 (tblReport_Section, just grouping sections, linked to
    ' ** tblSectionType, by specified CurrentAppName()), just sec_grouplevel discrepancies.
    Set qdf = .QueryDefs("zz_qry_Report_Group_02a")
    Set rst = qdf.OpenRecordset
    If rst.BOF = True And rst.EOF = True Then
      ' ** All's well.
      rst.Close
    Else
      rst.Close
      'Beep
      'Debug.Print "'GRP LVL DISCREPANCIES!"

      ' ** Update zz_qry_Report_Group_02a (zz_qry_Report_Group_01 (tblReport_Section, just
      ' ** grouping sections, linked to tblSectionType, by specified CurrentAppName()),
      ' ** just sec_grouplevel discrepancies), for sec_grouplevel.
      Set qdf = .QueryDefs("zz_qry_Report_Group_02b")
      qdf.Execute
    End If

    ' ** Update zz_qry_Report_Group_04b (tblReport_Group, with DLookups() to zz_qry_Report_Group_04a
    ' ** (tblReport_Group, linked to zz_qry_Report_Group_03 (tblSectionType, linked to tblSectionBaseType,
    ' ** just group-level sections), with rptgrp_section_header_new, by specified CurrentAppName())).
    Set qdf = .QueryDefs("zz_qry_Report_Group_04c")
    qdf.Execute dbFailOnError

    ' ** Update zz_qry_Report_Group_05b (tblReport_Group, with DLookups() to zz_qry_Report_Group_05a
    ' ** (tblReport_Group, linked to zz_qry_Report_Group_03 (tblSectionType, linked to tblSectionBaseType,
    ' ** just group-level sections), with rptgrp_section_footer_new, by specified CurrentAppName())).
    Set qdf = .QueryDefs("zz_qry_Report_Group_05c")
    qdf.Execute

    .Close
  End With

  ' ** Starting at Section 5, Group Levels increase, in
  ' ** pairs, from 1-10, with GroupLevel() Indexes of 0-9.

  ' ** AcSection enumeration:
  ' **   0  acDetail
  ' **   1  acHeader
  ' **   2  acFooter
  ' **   3  acPageHeader
  ' **   4  acPageFooter
  ' **   5  acGroupLevel1Header  ' ** Level 1 corresponds to GroupLevel(0), etc.
  ' **   6  acGroupLevel1Footer
  ' **   7  acGroupLevel2Header
  ' **   8  acGroupLevel2Footer

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set rpt = Nothing
  Set rptao = Nothing
  Set prj = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_Grp_Doc = blnRetValx

End Function

Private Function Rpt_Specs_Doc(rpt As Access.Report, dbs As DAO.Database, lngRptID As Long, lngDbsID As Long) As Boolean
' ** Called by:
' **   Rpt_Doc(), Above

  Const THIS_PROC As String = "Rpt_Specs_Doc"

  Dim rst As DAO.Recordset
  Dim blnRetVal As Boolean  ' ** Don't confuse with the parent's blnRetValX!

  Set rst = dbs.OpenRecordset("tblReport_Specification", dbOpenDynaset, dbConsistent)
  With rst
    .FindFirst "[dbs_id] = " & CStr(lngDbsID) & " And [rpt_id] = " & CStr(lngRptID)
'spec_id
    If .NoMatch = True Then
      .AddNew
'dbs_id
      ![dbs_id] = lngDbsID
'rpt_id
      ![rpt_id] = lngRptID
    Else
      .Edit
    End If
'AutoCenter
    ![AutoCenter] = rpt.AutoCenter
'AutoResize
    ![AutoResize] = rpt.AutoResize
'BorderStyle
    ![BorderStyle] = rpt.BorderStyle
'Caption
    If rpt.Caption <> vbNullString Then
      ![Caption] = rpt.Caption
    Else
      If IsNull(![Caption]) = False Then
        ![Caption] = Null
      End If
    End If
'DateGrouping
    ![DateGrouping] = rpt.DateGrouping
'FastLaserPrinting
    ![FastLaserPrinting] = rpt.FastLaserPrinting
'Filter
    If rpt.Filter <> vbNullString Then
      ![Filter] = rpt.Filter
    Else
      If IsNull(![Filter]) = False Then
        ![Filter] = Null
      End If
    End If
'GridX
    ![GridX] = rpt.GridX
'GridY
    ![GridY] = rpt.GridY
'GrpKeepTogether
    ![GrpKeepTogether] = rpt.GrpKeepTogether
'HasModule
    ![HasModule] = rpt.HasModule
'HelpContextId
    ![HelpContextId] = rpt.HelpContextId
'HelpFile
    If IsNull(rpt.HelpFile) = False Then
      If Trim(rpt.HelpFile) <> vbNullString Then
        ![HelpFile] = rpt.HelpFile
      Else
        If IsNull(![HelpFile]) = False Then
          ![HelpFile] = Null
        End If
      End If
    Else
      If IsNull(![HelpFile]) = False Then
        ![HelpFile] = Null
      End If
    End If
'LayoutForPrint
    ![LayoutForPrint] = rpt.LayoutForPrint
'LogicalPageWidth
    ![LogicalPageWidth] = rpt.LogicalPageWidth
'MarginBottom
    ![MarginBottom] = Rpt_Margins_Get(rpt, 2) ' ** Module Function: modReportFunctions.
'MarginLeft
    ![MarginLeft] = Rpt_Margins_Get(rpt, 3) ' ** Module Function: modReportFunctions.
'MarginRight
    ![MarginRight] = Rpt_Margins_Get(rpt, 4) ' ** Module Function: modReportFunctions.
'MarginTop
    ![MarginTop] = Rpt_Margins_Get(rpt, 1) ' ** Module Function: modReportFunctions.
'MaxButton
    ![MaxButton] = rpt.MaxButton
'MenuBar
    If rpt.MenuBar <> vbNullString Then
      ![MenuBar] = rpt.MenuBar
    Else
      If IsNull(![MenuBar]) = False Then
        ![MenuBar] = Null
      End If
    End If
'MinButton
    ![MinButton] = rpt.MinButton
'MinMaxButtons
    ![MinMaxButtons] = rpt.MinMaxButtons
'Modal
    ![Modal] = rpt.Modal
'Moveable
    ![Moveable] = rpt.Moveable
'OrderBy
    If rpt.OrderBy <> vbNullString Then
      ![OrderBy] = rpt.OrderBy
    Else
      If IsNull(![OrderBy]) = False Then
        ![OrderBy] = Null
      End If
    End If
'Orientation
    ![Orientation] = Rpt_Orient_Get(rpt)  ' ** Module Function: modReportFunctions.
'PageFooter
    ![PageFooter] = rpt.PageFooter
'PageHeader
    ![PageHeader] = rpt.PageHeader
'Painting
    ![Painting] = rpt.Painting
'PaletteSource
    If IsNull(rpt.PaletteSource) = False Then
      If Trim(rpt.PaletteSource) <> vbNullString Then
        ![PaletteSource] = rpt.PaletteSource
      Else
        If IsNull(![PaletteSource]) = False Then
          ![PaletteSource] = Null
        End If
      End If
    Else
      If IsNull(![PaletteSource]) = False Then
        ![PaletteSource] = Null
      End If
    End If
'Picture
    ![Picture] = rpt.Picture
'PictureAlignment
    ![PictureAlignment] = rpt.PictureAlignment
'PicturePages
    ![PicturePages] = rpt.PicturePages
'PictureSizeMode
    ![PictureSizeMode] = rpt.PictureSizeMode  ' ** Reports use PictureSizeMode, controls use SizeMode.
'PictureTiling
    ![PictureTiling] = rpt.PictureTiling
'PictureType
    ![PictureType] = rpt.PictureType
'PopUp
    ![PopUp] = rpt.PopUp
'RecordLocks
    ![RecordLocks] = rpt.RecordLocks
'ShortcutMenuBar
    If rpt.ShortcutMenuBar <> vbNullString Then
      ![ShortcutMenuBar] = rpt.ShortcutMenuBar
    Else
      If IsNull(![ShortcutMenuBar]) = False Then
        ![ShortcutMenuBar] = Null
      End If
    End If
'Tag
    If rpt.Tag <> vbNullString Then
      ![Tag] = rpt.Tag
    Else
      If IsNull(![Tag]) = False Then
        ![Tag] = Null
      End If
    End If
'Toolbar
    If rpt.Toolbar <> vbNullString Then
      ![Toolbar] = rpt.Toolbar
    Else
      If IsNull(![Toolbar]) = False Then
        ![Toolbar] = Null
      End If
    End If
'Width
    ![Width] = rpt.Width
'WindowHeight
    ![WindowHeight] = rpt.WindowHeight
'WindowLeft
    ![WindowLeft] = rpt.WindowLeft
'WindowTop
    ![WindowTop] = rpt.WindowTop
'WindowWidth
    ![WindowWidth] = rpt.WindowWidth
'spec_datemodified
    ![spec_datemodified] = Now()
    .Update
    .Close
  End With  ' ** rst.

'PROPS: 66
'Count  'SAME AS .Controls.Count
'Orientation  'THIS IS DIFFERENT THAN PORTRAIT/LANDSCAPE!
'PicturePalette  'I DON'T THINK THIS IS A STRING OR FILE NAME!
'PrtDevMode
'PrtDevNames
'PrtMip

  Set rst = Nothing

  Rpt_Specs_Doc = blnRetVal

End Function

Public Function Rpt_ReportList_Doc() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "Rpt_ReportList_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngFrms As Long, arr_varFrm As Variant
  Dim strFormName As String, strFormName_Mod As String, strCallFormName As String, strCallFormName_Mod As String
  Dim strProcName As String, strCtlName As String, strCap As String
  Dim lngFrmID As Long, lngCtlID As Long
  Dim lngThisDbsID As Long
  Dim lngModLines As Long, lngDecLines As Long, strLine As String
  Dim lngOpens As Long, arr_varOpen() As Variant
  Dim strModName As String
  Dim blnNoVia As Boolean, blnSkip As Boolean
  Dim lngRecs As Long
  Dim blnAddAll As Boolean, blnAdd As Boolean
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varOpen().
  Const O_ELEMS As Integer = 12  ' ** Array's first-element UBound().
  Const O_VID  As Integer = 0
  Const O_VNAM As Integer = 1
  Const O_PID  As Integer = 2
  Const O_PNAM As Integer = 3
  Const O_FID  As Integer = 4
  Const O_FNAM As Integer = 5
  Const O_LIN  As Integer = 6
  Const O_COD  As Integer = 7
  Const O_TXT1 As Integer = 8
  Const O_TXT2 As Integer = 9
  Const O_TXT3 As Integer = 10
  Const O_MULT As Integer = 11
  Const O_RAW  As Integer = 12

  ' ** Array: arr_varFrm().
  Const F_RXID4      As Integer = 0
  Const F_DID        As Integer = 1
  Const F_DNAM       As Integer = 2
  Const F_FID        As Integer = 3
  Const F_FNAM       As Integer = 4
  Const F_DID_MENU1  As Integer = 5
  Const F_FID_MENU1  As Integer = 6
  Const F_FNAM_MENU1 As Integer = 7
  Const F_CID_MENU1  As Integer = 8
  Const F_CNAM_MENU1 As Integer = 9
  Const F_CCAP_MENU1 As Integer = 10
  Const F_DID_MENU2  As Integer = 11
  Const F_FID_MENU2  As Integer = 12
  Const F_FNAM_MENU2 As Integer = 13
  Const F_CID_MENU2  As Integer = 14
  Const F_CNAM_MENU2 As Integer = 15
  Const F_CCAP_MENU2 As Integer = 16
  Const F_DID_MENU3  As Integer = 17
  Const F_FID_MENU3  As Integer = 18
  Const F_FNAM_MENU3 As Integer = 19
  Const F_CID_MENU3  As Integer = 20
  Const F_CNAM_MENU3 As Integer = 21
  Const F_CCAP_MENU3 As Integer = 12
  Const F_DID_VIA1   As Integer = 23
  Const F_FID_VIA1   As Integer = 24
  Const F_FNAM_VIA1  As Integer = 25
  Const F_CID_VIA1   As Integer = 26
  Const F_CNAM_VIA1  As Integer = 27
  Const F_CCAP_VIA1  As Integer = 28
  Const F_DID_VIA2   As Integer = 29
  Const F_FID_VIA2   As Integer = 30
  Const F_FNAM_VIA2  As Integer = 31
  Const F_CID_VIA2   As Integer = 32
  Const F_CNAM_VIA2  As Integer = 33
  Const F_CCAP_VIA2  As Integer = 34
  Const F_DAT        As Integer = 35

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  ' ** Rpt_ReportList_Doc() -- this -- empties:
  ' **   tblReport_List
  ' **   tblReport_List_Staging
  ' **   tblReport_VBComponent
  ' **   zz_tbl_Report_VBComponent_04
  ' **   zz_tbl_Report_VBComponent_07
  ' **   zz_tbl_Report_VBComponent_08
  ' **   zz_tbl_Report_VBComponent_09
  ' **   zz_tbl_Report_VBComponent_10
  ' **   zz_tbl_Report_VBComponent_11

  ' ** Rpt_ReportList_ProcXRef(), below, empties:
  ' **   zz_tbl_Report_VBComponent_01
  ' **   zz_tbl_Report_VBComponent_02
  ' **   zz_tbl_Report_VBComponent_03

  ' ** zz_tbl_Report_VBComponent_06 is used
  ' ** when populating tblReport_Group, and
  ' ** not connected with the Report List.
  ' **   zz_qry_Report_Control_51a

  DoCmd.Hourglass True
  DoEvents

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

blnSkip = False
If blnSkip = False Then
  Set dbs = CurrentDb
  With dbs
    ' ** Empty tblReport_List, by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01j")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty tblReport_List_Staging, by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01k")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty tblReport_VBComponent, by specified [dbid].
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01l")
    With qdf.Parameters
      ![dbid] = lngThisDbsID
    End With
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Report_VBComponent_07.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01e")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Report_VBComponent_08.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01f")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Report_VBComponent_09.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01g")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Report_VBComponent_10.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01h")
    qdf.Execute
    Set qdf = Nothing
    ' ** Empty zz_tbl_Report_VBComponent_11.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01i")
    qdf.Execute
    Set qdf = Nothing
    .Close
  End With  ' ** dbs.
  Set qdf = Nothing
  Set dbs = Nothing

  DoCmd.Hourglass True
  Debug.Print "'AutoNum: tblReport_List"
  DoEvents
  ChangeSeed_Ext "tblReport_List"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: tblReport_List_Staging"
  DoEvents
  ChangeSeed_Ext "tblReport_List_Staging"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: tblReport_VBComponent"
  DoEvents
  ChangeSeed_Ext "tblReport_VBComponent"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: zz_tbl_Report_VBComponent_07"
  DoEvents
  ChangeSeed_Ext "zz_tbl_Report_VBComponent_07"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: zz_tbl_Report_VBComponent_08"
  DoEvents
  ChangeSeed_Ext "zz_tbl_Report_VBComponent_08"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: zz_tbl_Report_VBComponent_09"
  DoEvents
  ChangeSeed_Ext "zz_tbl_Report_VBComponent_09"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: zz_tbl_Report_VBComponent_10"
  DoEvents
  ChangeSeed_Ext "zz_tbl_Report_VBComponent_10"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: zz_tbl_Report_VBComponent_11"
  DoEvents
  ChangeSeed_Ext "zz_tbl_Report_VBComponent_11"  ' ** Module Function: modAutonumberFieldFuncs.
  DoCmd.Hourglass True
  Debug.Print "'AutoNum: Done!"
  DoEvents
End If  ' ** blnSkip.

blnSkip = False
If blnSkip = False Then
  Rpt_ReportList_ProcXRef  ' ** Function: Below.
End If  ' ** blnSkip.

'1. Rpt_ReportList_ProcXRef():
'   Collects code data and appends to
'     zz_tbl_Report_VBComponent_01
'   Then uses that to append break-outs into
'     zz_tbl_Report_VBComponent_02
'   Finally distilling that into
'     zz_tbl_Report_VBComponent_03
'   Calls Rpt_ReportList_ProcVet() for checking.

'MUST BE UP-TO-DATE!!  tblObject_Image

  Set dbs = CurrentDb
  With dbs

    DoCmd.Hourglass True
    DoEvents

blnSkip = False
If blnSkip = False Then
    ' ** Delete zz_qry_Report_VBComponent_10b (tblReport_VBComponent, not in zz_qry_Report_VBComponent_10a
    ' ** (zz_tbl_Report_VBComponent_03, just good lines), that don't belong, by specified CurrentAppName()).
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_10c")
    qdf.Execute
    Set qdf = Nothing

    ' ** Append zz_qry_Report_VBComponent_10a (zz_tbl_Report_VBComponent_03, just good lines) to tblReport_VBComponent.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_10e")
    qdf.Execute
    Set qdf = Nothing

    ' ** Append zz_qry_Report_VBComponent_10d (zz_qry_Report_VBComponent_10a (zz_tbl_Report_VBComponent_03,
    ' ** just good lines), not in tblReport_VBComponent) to tblReport_VBComponent.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_10f")
    qdf.Execute
    Set qdf = Nothing

    ' ** Append zz_qry_Report_VBComponent_12a (tblReport_Subform, linked to tblReport_VBComponent,
    ' ** via parent's rpt_id, by specified CurrentAppName()) to tblReport_VBComponent.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_12b")
    qdf.Execute
    Set qdf = Nothing

    ' ** Empty zz_tbl_Report_VBComponent_04.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01d")
    qdf.Execute
    Set qdf = Nothing

    ' ** Append zz_qry_Report_VBComponent_15d (zz_qry_Report_VBComponent_15c (zz_qry_Report_VBComponent_13e
    ' ** (Union of zz_qry_Report_VBComponent_13c (zz_qry_Report_VBComponent_13a (tblReport, linked to
    ' ** tblReport_VBComponent, with form and procedure names, without subreports, by specified CurrentAppName()),
    ' ** linked to zz_qry_Report_VBComponent_13b (zz_qry_Report_VBComponent_13a (tblReport, linked to
    ' ** tblReport_VBComponent, with form and procedure names, without subreports, by specified CurrentAppName()),
    ' ** grouped by rpt_name, with cnt, proc_priority), just cnt = 1), zz_qry_Report_VBComponent_13d
    ' ** (zz_qry_Report_VBComponent_13a (tblReport, linked to tblReport_VBComponent, with form and procedure
    ' ** names, without subreports, by specified CurrentAppName()), linked to zz_qry_Report_VBComponent_13b
    ' ** (zz_qry_Report_VBComponent_13a (tblReport, linked to tblReport_VBComponent, with form and procedure
    ' ** names, without subreports, by specified CurrentAppName()), grouped by rpt_name, with cnt, proc_priority),
    ' ** just cnt > 1, grouped by rpt_name, vbcom_id, with First(vbcom_proc_name))), linked to
    ' ** zz_qry_Report_VBComponent_15b (zz_qry_Report_VBComponent_14c (zz_qry_Report_VBComponent_14b
    ' ** (zz_qry_Report_VBComponent_14a (tblObject_Image, just command buttons), linked to tblForm_Control,
    ' ** just Print buttons), just those with Printer picture), linked to zz_qry_Report_VBComponent_15a
    ' ** (tblVBComponent_Procedure, just form procedures), to get OnClick procedure), complete report list),
    ' ** grouped by frm_id) to zz_tbl_Report_VBComponent_04.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_15f")
    qdf.Execute
    Set qdf = Nothing

    ' ** Update zz_qry_Report_VBComponent_15h (zz_tbl_Report_VBComponent_04, just numeric ctl_name).
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_15i")
    qdf.Execute
    Set qdf = Nothing
    ' ** Delete zz_qry_Report_VBComponent_15j for subreports (tblReport_VBComponent, just reports, subreports).
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_15k")
    qdf.Execute
    Set qdf = Nothing
    ' ** Delete zz_qry_Report_VBComponent_15j, for reports (tblReport_VBComponent, just reports, subreports).
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_15l")
    qdf.Execute
    Set qdf = Nothing

End If  ' ** blnSkip.

blnSkip = False
If blnSkip = False Then

    DoCmd.Hourglass True
    DoEvents

'MORE NEEDS TO HAPPEN HERE?
'Trace form names back to the menu that called them.
    ' ** zz_tbl_Report_VBComponent_04, for editing.
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_16")
    Set rst = qdf.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        ' ** Shouldn't be here!
        lngFrms = 0&
        arr_varFrm = Empty
      Else
        .MoveLast
        lngFrms = .RecordCount
        .MoveFirst
        arr_varFrm = .GetRows(lngFrms)
        ' *************************************************************
        ' ** Array: arr_varFrm()
        ' **
        ' **   Field  Element  Name                    Constant
        ' **   =====  =======  ======================  ==============
        ' **     1       0     rptxref_id4             F_RXID4
        ' **     2       1     dbs_id                  F_DID
        ' **     3       2     dbs_name                F_DNAM
        ' **     4       3     frm_id                  F_FID
        ' **     5       4     frm_name                F_FNAM
        ' **     6       5     dbs_id_menu1            F_DID_MENU1
        ' **     7       6     frm_id_menu1            F_FID_MENU1
        ' **     8       7     frm_name_menu1          F_FNAM_MENU1
        ' **     9       8     ctl_id_menu1            F_CID_MENU1
        ' **    10       9     ctl_name_menu1          F_CNAM_MENU1
        ' **    11      10     ctl_caption_menu1       F_CCAP_MENU1
        ' **    12      11     dbs_id_menu2            F_DID_MENU2
        ' **    13      12     frm_id_menu2            F_FID_MENU2
        ' **    14      13     frm_name_menu2          F_FNAM_MENU2
        ' **    15      14     ctl_id_menu2            F_CID_MENU2
        ' **    16      15     ctl_name_menu2          F_CNAM_MENU2
        ' **    17      16     ctl_caption_menu2       F_CCAP_MENU2
        ' **    18      17     dbs_id_menu3            F_DID_MENU3
        ' **    19      18     frm_id_menu3            F_FID_MENU3
        ' **    20      19     frm_name_menu3          F_FNAM_MENU3
        ' **    21      20     ctl_id_menu3            F_CID_MENU3
        ' **    22      21     ctl_name_menu3          F_CNAM_MENU3
        ' **    23      22     ctl_caption_menu3       F_CCAP_MENU3
        ' **    24      23     dbs_id_via01            F_DID_VIA1
        ' **    25      24     frm_id_via01            F_FID_VIA1
        ' **    26      25     frm_name_via01          F_FNAM_VIA1
        ' **    27      26     ctl_id_via01            F_CID_VIA1
        ' **    28      27     ctl_name_via01          F_CNAM_VIA1
        ' **    29      28     ctl_caption_via01       F_CCAP_VIA1
        ' **    30      29     dbs_id_via02            F_DID_VIA2
        ' **    31      30     frm_id_via02            F_FID_VIA2
        ' **    32      31     frm_name_via02          F_FNAM_VIA2
        ' **    33      32     ctl_id_via02            F_CID_VIA2
        ' **    34      33     ctl_name_via02          F_CNAM_VIA2
        ' **    35      34     ctl_caption_via02       F_CCAP_VIA2
        ' **    36      35     rptxref_datemodified    F_DAT
        ' **
        ' *************************************************************
      End If
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

'AT THIS POINT, zz_tbl_Report_VBComponent_04 HAS EVERY FORM WHERE
'A GTR_REF WAS FOUND; 64 FORMS REPRESENTING 136 REPORT CALLS!
'IT NOW NEEDS THE MENU PATH TO THAT FORM!
'FIRST, LET'S FIND THE FORM THAT CALLS THE ONE WITH THE BUTTON!
'BUT, WE DON'T YET KNOW HOW DEEP THE THREAD IS!
'I'M GOING TO USE THE ARRAY IN A DIFFERENT WAY WHILE I FIGURE THIS OUT.

'I COULD ALSO LOOK FOR EVERY INSTANCE OF DoCmd.OpenForm,
'THEN MATCH THEM BACK!

    lngOpens = 0&
    ReDim arr_varOpen(O_ELEMS, 0)

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For Each vbc In .VBComponents
        With vbc
          strModName = .Name
          If Left(strModName, 5) = "Form_" Then  ' ** For now, just forms.
            Set cod = .CodeModule
            With cod
              lngModLines = .CountOfLines
              lngDecLines = .CountOfDeclarationLines
              For lngY = lngDecLines To lngModLines
                strLine = Trim$(.Lines(lngY, 1))
                If strLine <> vbNullString Then
                  If Left$(strLine, 1) <> "'" Then
                    intPos1 = InStr(strLine, "DoCmd.OpenForm")
                    If intPos1 > 0 Then
                      strProcName = .ProcOfLine(lngY, vbext_pk_Proc)
                      lngOpens = lngOpens + 1&
                      lngE = lngOpens - 1&
                      ReDim Preserve arr_varOpen(O_ELEMS, lngE)
                      ' ******************************************************
                      ' ** Array: arr_varOpen()
                      ' **
                      ' **   Field  Element  Name                 Constant
                      ' **   =====  =======  ===================  ==========
                      ' **     1       0     vbcom_id             O_VID
                      ' **     2       1     vbcom_name           O_VNAM
                      ' **     3       2     vbcomproc_id         O_PID
                      ' **     4       3     vbcomproc_name       O_PNAM
                      ' **     5       4     frm_id1              O_FID
                      ' **     6       5     frm_name1            O_FNAM
                      ' **     7       6     rptxref_line         O_LIN
                      ' **     8       7     rptxref_code         O_COD
                      ' **     9       8     rptxref_text1        O_TXT1
                      ' **    10       9     rptxref_text2        O_TXT2
                      ' **    11      10     rptxref_text3        O_TXT3
                      ' **    12      11     rptxref_multiline    O_MULT
                      ' **    13      12     rptxref_raw          O_RAW
                      ' **
                      ' ******************************************************
                      arr_varOpen(O_VID, lngE) = CLng(0)
                      arr_varOpen(O_VNAM, lngE) = strModName
                      arr_varOpen(O_PID, lngE) = CLng(0)
                      arr_varOpen(O_PNAM, lngE) = strProcName
                      arr_varOpen(O_FID, lngE) = CLng(0)
                      arr_varOpen(O_FNAM, lngE) = Mid$(strModName, 6)
                      arr_varOpen(O_LIN, lngE) = lngY
                      intPos1 = InStr(strLine, " ")
                      strTmp01 = Trim$(Left$(strLine, intPos1))
                      strTmp02 = vbNullString
                      If IsNumeric(strTmp01) = True Then
                        arr_varOpen(O_COD, lngE) = CLng(strTmp01)
                        strTmp02 = Trim$(Mid$(strLine, intPos1))
                      Else
                        arr_varOpen(O_COD, lngE) = Null
                      End If
                      If strTmp02 <> vbNullString Then
                        arr_varOpen(O_TXT1, lngE) = strTmp02
                        If Right$(strLine, 1) = "_" Then
                          arr_varOpen(O_MULT, lngE) = CBool(True)
                          arr_varOpen(O_TXT2, lngE) = Trim$(.Lines(lngY + 1&, 1))
                          If Right$(Trim$(.Lines(lngY + 1&, 1)), 1) = "_" Then
                            arr_varOpen(O_TXT3, lngE) = Trim$(.Lines(lngY + 2&, 1))
                          Else
                            arr_varOpen(O_TXT3, lngE) = Null
                          End If
                        Else
                          arr_varOpen(O_MULT, lngE) = CBool(False)
                          arr_varOpen(O_TXT2, lngE) = Null
                          arr_varOpen(O_TXT3, lngE) = Null
                        End If
                      Else
                        arr_varOpen(O_TXT1, lngE) = strLine
                        arr_varOpen(O_MULT, lngE) = CBool(False)
                        arr_varOpen(O_TXT2, lngE) = Null
                        arr_varOpen(O_TXT3, lngE) = Null
                      End If
                      arr_varOpen(O_RAW, lngE) = strLine
                    End If  ' ** intPos1.
                  End If  ' ** Remark.
                End If  ' ** vbNullString.
              Next  ' ** lngY.
            End With  ' ** cod.
            Set cod = Nothing
          End If  ' ** strModName.
        End With  ' ** vbc.
      Next  ' ** vbc.
      Set vbc = Nothing
    End With  ' ** vbp.
    Set vbp = Nothing

    ' ** Add vbcom_id.
    Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
    With rst
      For lngX = 0& To (lngOpens - 1&)
        .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & arr_varOpen(O_VNAM, lngX) & "'"
        If .NoMatch = False Then
          arr_varOpen(O_VID, lngX) = ![vbcom_id]
        Else
          Stop
        End If
      Next
      .Close
    End With
    Set rst = Nothing

    ' ** Add vbcomproc_id.
    Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
    With rst
      For lngX = 0& To (lngOpens - 1&)
        .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & arr_varOpen(O_VID, lngX) & " And " & _
          "[vbcomproc_name] = '" & arr_varOpen(O_PNAM, lngX) & "'"
        If .NoMatch = False Then
          arr_varOpen(O_PID, lngX) = ![vbcomproc_id]
        Else
          Stop
        End If
      Next
      .Close
    End With
    Set rst = Nothing

    ' ** Add frm_id.
    Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
    With rst
      For lngX = 0& To (lngOpens - 1&)
        .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_name] = '" & arr_varOpen(O_FNAM, lngX) & "'"
        If .NoMatch = False Then
          arr_varOpen(O_FID, lngX) = ![frm_id]
        Else
          Stop
        End If
      Next
      .Close
    End With
    Set rst = Nothing

    blnAddAll = False: lngRecs = 0&
    Set rst = .OpenRecordset("zz_tbl_Report_VBComponent_07", dbOpenDynaset, dbAppendOnly)  ' ** First use of zz_tbl_Report_VBComponent_07.
    With rst
      If .BOF = True And .EOF = True Then
        blnAddAll = True
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
      End If
      For lngX = 0& To (lngOpens - 1&)
        blnAdd = False
        Select Case blnAddAll
        Case True
          blnAdd = True
        Case False
          .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id = " & CStr(arr_varOpen(O_VID, lngX)) & " And " & _
            "[vbcomproc_id] = " & CStr(arr_varOpen(O_PID, lngX)) & " And [frm_id1] = " & CStr(arr_varOpen(O_FID, lngX)) & " And " & _
            "[[rptxref_line]] = " & CStr(arr_varOpen(O_LIN, lngX))
          blnAdd = .NoMatch
        End Select
        If blnAdd = True Then
          .AddNew
          ' ** ![rptxref_id7] : AutoNumber.
          ![dbs_id] = lngThisDbsID
          ![dbs_name] = CurrentAppName()
          ![vbcom_id] = arr_varOpen(O_VID, lngX)
          ![vbcom_name] = arr_varOpen(O_VNAM, lngX)
          ![vbcomproc_id] = arr_varOpen(O_PID, lngX)
          ![vbcomproc_name] = arr_varOpen(O_PNAM, lngX)
          ![frm_id] = arr_varOpen(O_FID, lngX)
          ![frm_name1] = arr_varOpen(O_FNAM, lngX)
          ![rptxref_line] = arr_varOpen(O_LIN, lngX)
        Else
          .Edit
        End If
        ![rptxref_code] = arr_varOpen(O_COD, lngX)
        ![rptxref_text1] = arr_varOpen(O_TXT1, lngX)
        ![rptxref_text2] = arr_varOpen(O_TXT2, lngX)
        ![rptxref_text3] = arr_varOpen(O_TXT3, lngX)
        ![rptxref_multiline] = arr_varOpen(O_MULT, lngX)
        ![rptxref_raw] = arr_varOpen(O_RAW, lngX)
        ![rptxref_datemodified] = Now()
        .Update
        .MoveFirst
      Next
      .Close
    End With
    Set rst = Nothing

'zz_tbl_Report_VBComponent_07 has all OpenForm commands.
'frm_id1, and frm_name1, is the form in which the command is invoked.
'frm_id2, and frm_name2, is the form it's opening, so that is
'what we want to match to the forms in zz_tbl_Report_VBComponent_04.

    DoCmd.Hourglass True
    DoEvents

    ' *************************************************************
    ' ** Array: arr_varFrm()
    ' **
    ' **   Field  Element  Name                    Constant
    ' **   =====  =======  ======================  ==============
    ' **     1       0     rptxref_id4             F_RXID4
    ' **     2       1     dbs_id                  F_DID
    ' **     3       2     dbs_name                F_DNAM
    ' **     4       3     frm_id                  F_FID
    ' **     5       4     frm_name                F_FNAM
    ' **     6       5     dbs_id_menu1            F_DID_MENU1
    ' **     7       6     frm_id_menu1            F_FID_MENU1
    ' **     8       7     frm_name_menu1          F_FNAM_MENU1
    ' **     9       8     ctl_id_menu1            F_CID_MENU1
    ' **    10       9     ctl_name_menu1          F_CNAM_MENU1
    ' **    11      10     ctl_caption_menu1       F_CCAP_MENU1
    ' **    12      11     dbs_id_menu2            F_DID_MENU2
    ' **    13      12     frm_id_menu2            F_FID_MENU2
    ' **    14      13     frm_name_menu2          F_FNAM_MENU2
    ' **    15      14     ctl_id_menu2            F_CID_MENU2
    ' **    16      15     ctl_name_menu2          F_CNAM_MENU2
    ' **    17      16     ctl_caption_menu2       F_CCAP_MENU2
    ' **    18      17     dbs_id_menu3            F_DID_MENU3
    ' **    19      18     frm_id_menu3            F_FID_MENU3
    ' **    20      19     frm_name_menu3          F_FNAM_MENU3
    ' **    21      20     ctl_id_menu3            F_CID_MENU3
    ' **    22      21     ctl_name_menu3          F_CNAM_MENU3
    ' **    23      22     ctl_caption_menu3       F_CCAP_MENU3
    ' **    24      23     dbs_id_via01            F_DID_VIA1
    ' **    25      24     frm_id_via01            F_FID_VIA1
    ' **    26      25     frm_name_via01          F_FNAM_VIA1
    ' **    27      26     ctl_id_via01            F_CID_VIA1
    ' **    28      27     ctl_name_via01          F_CNAM_VIA1
    ' **    29      28     ctl_caption_via01       F_CCAP_VIA1
    ' **    30      29     dbs_id_via02            F_DID_VIA2
    ' **    31      30     frm_id_via02            F_FID_VIA2
    ' **    32      31     frm_name_via02          F_FNAM_VIA2
    ' **    33      32     ctl_id_via02            F_CID_VIA2
    ' **    34      33     ctl_name_via02          F_CNAM_VIA2
    ' **    35      34     ctl_caption_via02       F_CCAP_VIA2
    ' **    36      35     rptxref_datemodified    F_DAT
    ' **
    ' *************************************************************

End If  ' ** blnSkip.

blnSkip = False
If blnSkip = False Then

    DoCmd.Hourglass True
    DoEvents

    ' ** Append zz_qry_Report_VBComponent_32i (zz_qry_Report_VBComponent_32h (zz_qry_Report_VBComponent_32g
    ' ** (Union of zz_qry_Report_VBComponent_32c (zz_qry_Report_VBComponent_32a (tblReport, linked to
    ' ** tblReport_VBComponent, with form and procedure names, without subreports), linked to
    ' ** zz_qry_Report_VBComponent_32b (zz_qry_Report_VBComponent_32a (tblReport, linked to tblReport_VBComponent,
    ' ** with form and procedure names, without subreports), grouped by rpt_name, with cnt, proc_priority), just
    ' ** cnt = 1), zz_qry_Report_VBComponent_32f (zz_qry_Report_VBComponent_32d (zz_qry_Report_VBComponent_32a
    ' ** (tblReport, linked to tblReport_VBComponent, with form and procedure names, without subreports), linked to
    ' ** zz_qry_Report_VBComponent_32b (zz_qry_Report_VBComponent_32a (tblReport, linked to tblReport_VBComponent,
    ' ** with form and procedure names, without subreports), grouped by rpt_name, with cnt, proc_priority), grouped
    ' ** by rpt_name, vbcom_id, with First(vbcom_proc_name)), linked to zz_qry_Report_VBComponent_32e
    ' ** (tblVBComponent_Procedure, just 'cmdPrintReport_Click' on frmJournal_Columns, with tblReport, just
    ' ** 'rptPostingJournal_Column'; Cartesian), with 'rptPostingJournal_Column' fixed)), linked to tblForm,
    ' ** zz_qry_Report_VBComponent_15b (zz_qry_Report_VBComponent_14c (zz_qry_Report_VBComponent_14b
    ' ** (zz_qry_Report_VBComponent_14a (tblObject_Image, just command buttons), linked to tblForm_Control,
    ' ** just Print buttons), just those with Printer picture), linked to zz_qry_Report_VBComponent_15a
    ' ** (tblVBComponent_Procedure, just form procedures), to get OnClick procedure), complete report list),
    ' ** grouped by frm_id) to tblReport_List.
    'Set qdf = .QueryDefs("zz_qry_Report_VBComponent_32k")
    'qdf.Execute
    'Set qdf = Nothing

End If  ' ** blnSkip.

    .Close
  End With  ' ** dbs.

  blnRetValx = Rpt_ReportList_RunQrys  ' ** Function: Below.

  DoCmd.Hourglass False

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_ReportList_Doc = blnRetValx

End Function

Public Function Rpt_ReportList_RunQrys() As Boolean

  Const THIS_PROC As String = "Rpt_ReportList_RunQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim blnRetVal As Boolean

  blnRetVal = True

'zz_qry_Report_VBComponent_73f   '36
'zz_qry_Report_VBComponent_73e   '21
'zz_qry_Report_VBComponent_73j   ' 7
'zz_qry_Report_VBComponent_73l   ' 8
'                               =====
'                                 72

  Set dbs = CurrentDb
  With dbs
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_80a")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_91c")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_91e")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_92b")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_93b")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_93d")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_94e")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_95a_07")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_95a_11")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_95a_17")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_95a_23")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_95b_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_95d_02")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96g_03")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96h_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96i_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96j_10")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96j_14")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96j_17")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96j_24")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96k_08")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96k_11")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96k_14")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_96l_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97a_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97b_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97c_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97c_08")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97c_11")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97c_15")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97d_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97e_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97f_02")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97f_07")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97f_09")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97f_14")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97f_18")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97g_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97g_09")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97h_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97h_09")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97i_04")
    qdf.Execute
    Set qdf = Nothing
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97i_07")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97j_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97k_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97l_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_97m_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_98b_10")  'First use of zz_tbl_Report_VBComponent_08.
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_98c_10")  'First use of zz_tbl_Report_VBComponent_09.
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99g_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99g_08")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99g_12")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99g_15")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99g_19")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99g_23")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99h_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99h_12")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99h_16")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99j")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_04")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_08")  'SKIPPED 2!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_12")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_16")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_20")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_23")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_27")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_30")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_34")  'SKIPPED THEM!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_37")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_40")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_44")  'SKIPPED THEM!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_47")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_50")  'SKIPPED THEM!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_54")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_57")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_61")  'SKIPPED THEM!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_64")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_68")  'SKIPPED THEM!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_71")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99l_75")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99o_08a")  ' ** First use of zz_tbl_Report_VBComponent_10.
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99o_08c")  ' ** First use of zz_tbl_Report_VBComponent_11.
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99r_15")  'SKIPPED 'EM ALL!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99r_27")  'SKIPPED THEM!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99t_17")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99t_28")  'SKIPPED 106!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99t_44")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99w_24")  'SKIPPED IT!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99w_25")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99w_26")  '0!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99x_03")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99y_03")  '0!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99y_05")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99y_07")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_03")  '0!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_05")  '0!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_07")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_15")  '0!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_19")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_21")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_33")  '0!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_35")  'SKIPPED THEM!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_37")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_45")  '0!
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_49")
    qdf.Execute
    Set qdf = Nothing
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_99z_54")
    qdf.Execute
    Set qdf = Nothing

'zz_qry_Report_VBComponent_75k
'zz_qry_Report_VBComponent_75m
'zz_qry_Report_VBComponent_76j

'qryReport_List_65y
'qryReport_List_65w
'qryReport_List_66x
'qryReport_List_66z

'qryReport_List_68y


    .Close
  End With

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_ReportList_RunQrys = blnRetVal

End Function

Private Function Rpt_ReportList_ProcXRef() As Boolean
' ** This collects raw code references and
' ** puts them into zz_tbl_Report_VBComponent_01
' ** Called by:
' **   Rpt_ReportList_Doc(), Above

  Const THIS_PROC As String = "Rpt_ReportList_ProcXRef"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim lngVBComID As Long, lngProcID As Long
  Dim lngModLines As Long
  Dim lngThisDbsID As Long
  Dim strLine As String, strModName As String, strProcName As String
  Dim strRptName As String
  Dim lngRecs As Long
  Dim blnSkip As Boolean
  Dim intPos1 As Integer, intLen1 As Integer, intLen2 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, strTmp02 As String, lngTmp03 As Long
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 20  ' ** Array's first-element UBound().
  Const R_DID   As Integer = 0
  Const R_DNAM  As Integer = 1
  Const R_FID   As Integer = 2
  Const R_FNAM  As Integer = 3
  Const R_CID   As Integer = 4
  Const R_CNAM  As Integer = 5
  Const R_VID   As Integer = 6
  Const R_VNAM  As Integer = 7
  Const R_PID   As Integer = 8
  Const R_PNAM  As Integer = 9
  Const R_LNBEG As Integer = 10
  Const R_LNEND As Integer = 11
  Const R_RID   As Integer = 12
  Const R_RNAM  As Integer = 13
  Const R_CODE1 As Integer = 14
  Const R_LIN1  As Integer = 15
  Const R_CLIN1 As Integer = 16
  Const R_CODE2 As Integer = 17
  Const R_LIN2  As Integer = 18
  Const R_CLIN2 As Integer = 19
  Const R_FND   As Integer = 20

  Const GTR_REF As String = "'##GTR_Ref:"

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngRpts = 0&
  ReDim arr_varRpt(R_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    Set qdf = .QueryDefs("zz_qry_Report_List_05")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngRpts = lngRpts + 1&
          lngE = lngRpts - 1&
          ReDim Preserve arr_varRpt(R_ELEMS, lngE)
          arr_varRpt(R_DID, lngE) = ![dbs_id]
          arr_varRpt(R_DNAM, lngE) = ![dbs_name]
          arr_varRpt(R_FID, lngE) = ";"
          arr_varRpt(R_FNAM, lngE) = vbNullString
          arr_varRpt(R_CID, lngE) = CLng(0)
          arr_varRpt(R_CNAM, lngE) = vbNullString
          arr_varRpt(R_VID, lngE) = ";"            '![vbcom_id]  This is the report's vbcom_id!
          arr_varRpt(R_VNAM, lngE) = vbNullString  '![vbcom_name]
          arr_varRpt(R_PID, lngE) = ";"
          arr_varRpt(R_PNAM, lngE) = vbNullString
          arr_varRpt(R_LNBEG, lngE) = CLng(0)
          arr_varRpt(R_LNEND, lngE) = CLng(0)
          arr_varRpt(R_RID, lngE) = ![rpt_id]
          arr_varRpt(R_RNAM, lngE) = ![rpt_name]
          arr_varRpt(R_CODE1, lngE) = vbNullString
          arr_varRpt(R_LIN1, lngE) = ";"
          arr_varRpt(R_CLIN1, lngE) = ";"
          arr_varRpt(R_CODE2, lngE) = vbNullString
          arr_varRpt(R_LIN2, lngE) = ";"
          arr_varRpt(R_CLIN2, lngE) = ";"
          arr_varRpt(R_FND, lngE) = CInt(0)
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With
    Set rst = Nothing
    Set qdf = Nothing
    ' ** Empty zz_tbl_Report_VBComponent_01 (cascades to .._02, .._03).
    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01a")
    qdf.Execute
    Set qdf = Nothing
  End With  ' ** dbs.

  If lngRpts > 0& Then

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For Each vbc In .VBComponents
        With vbc
          strModName = .Name
          If Left(strModName, 5) = "Form_" Then
            Set cod = .CodeModule
            With cod
              lngModLines = .CountOfLines
              For lngY = 1& To lngModLines
                strLine = Trim$(.Lines(lngY, 1))
                If strLine <> vbNullString Then
                  ' ** ##GTR_Ref: rptAccountBalance
                  If InStr(strLine, GTR_REF) > 0 Then
                    If .ProcOfLine(lngY, vbext_pk_Proc) <> vbNullString Then  ' ** Ignore if in Declaration section.
                      lngVBComID = 0&: lngProcID = 0&
                      intPos1 = InStr(strLine, ": ")
                      If intPos1 > 0 Then
                        strRptName = Trim$(Mid$(strLine, (intPos1 + 1)))
                        For lngX = 0& To (lngRpts - 1&)
                          If arr_varRpt(R_RNAM, lngX) = strRptName Then
                            varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", "[vbcom_name] = '" & strModName & "'")
                            If IsNull(varTmp00) = False Then
                              lngVBComID = CLng(varTmp00)
                              If InStr(arr_varRpt(R_VID, lngX), (";" & CStr(lngVBComID) & ";")) = 0 Then  ' ** When empty, it still has a ';'.
                                arr_varRpt(R_VID, lngX) = arr_varRpt(R_VID, lngX) & CStr(lngVBComID) & ";"
                              End If
                            End If
                            strProcName = .ProcOfLine(lngY, vbext_pk_Proc)
                            If lngVBComID > 0& Then
                              varTmp00 = DLookup("[vbcomproc_id]", "tblVBComponent_Procedure", "[vbcom_id] = " & CStr(lngVBComID) & " And " & _
                                "[vbcomproc_name] = '" & strProcName & "'")
                              If IsNull(varTmp00) = False Then
                                lngProcID = CLng(varTmp00)
                                If InStr(arr_varRpt(R_PID, lngX), (";" & CStr(lngProcID) & ";")) = 0 Then  ' ** When empty, it still has a ';'.
                                  arr_varRpt(R_PID, lngX) = arr_varRpt(R_PID, lngX) & CStr(lngProcID) & ";"
                                End If
                              End If
                            End If
                            ' ** The GTR Ref is a remarked line, so find the next code line.
                            lngTmp03 = lngY + 1&
                            strTmp02 = Trim$(.Lines(lngTmp03, 1))  ' ** Check the next line
                            If InStr(strTmp02, GTR_REF) > 0 Then
                              lngTmp03 = lngY + 2&
                              strTmp02 = Trim$(.Lines(lngTmp03, 1))  ' ** Check the next line
                              If InStr(strTmp02, GTR_REF) > 0 Then
                                lngTmp03 = lngY + 3&
                                strTmp02 = Trim$(.Lines(lngTmp03, 1))  ' ** Check the next line
                                If InStr(strTmp02, GTR_REF) > 0 Then
                                  lngTmp03 = lngY + 4&
                                  strTmp02 = Trim$(.Lines(lngTmp03, 1))  ' ** Check the next line
                                  If InStr(strTmp02, GTR_REF) > 0 Then
                                    lngTmp03 = lngY + 5&
                                    strTmp02 = Trim$(.Lines(lngTmp03, 1))  ' ** Check the next line
                                    If InStr(strTmp02, GTR_REF) > 0 Then
                                      lngTmp03 = lngY + 6&
                                      strTmp02 = Trim$(.Lines(lngTmp03, 1))  ' ** Check the next line
                                      If InStr(strTmp02, GTR_REF) > 0 Then
                                        lngTmp03 = 0&
                                        strTmp02 = vbNullString
                                        Stop
                                      End If
                                    End If
                                  End If
                                End If
                              End If
                            End If
                            ' ** Put the code line into the array, not the GTR Ref line.
                            If InStr(arr_varRpt(R_LIN1, lngX), (";" & CStr(lngVBComID) & "." & CStr(lngTmp03) & ";")) = 0 Then
                              arr_varRpt(R_LIN1, lngX) = arr_varRpt(R_LIN1, lngX) & (CStr(lngVBComID) & "." & CStr(lngTmp03) & ";")
                            End If
                            intPos1 = InStr(strTmp02, " ")
                            If intPos1 > 0 Then
                              strTmp01 = Trim$(Left$(strTmp02, intPos1))
                              If IsNumeric(strTmp01) = True Then
                                If InStr(arr_varRpt(R_CLIN1, lngX), (";" & CStr(lngVBComID) & "." & strTmp01 & ";")) = 0 Then
                                  arr_varRpt(R_CLIN1, lngX) = arr_varRpt(R_CLIN1, lngX) & (CStr(lngVBComID) & "." & _
                                    strTmp01 & "(" & CStr(lngY) & ")" & ";")
                                End If
                              End If
                            End If
                          End If  ' ** strRptName.
                        Next  ' ** lngx.
                      Else
                        Beep
                        Stop
                      End If  ' ** Colon.
                    End If  ' ** strProcName.
                  End If  ' ** GTR_REF.
                End If  ' ** vbNullString.
              Next  ' ** lngY.
            End With  ' ** cod.
          End If  ' ** strModName.
        End With  ' ** vbc.
      Next  ' ** vbc.
    End With  ' ** vbp.

    With dbs

      ' ** Dump the raw data into this temporary table.
      Set rst = .OpenRecordset("zz_tbl_Report_VBComponent_01", dbOpenDynaset)
      With rst
        For lngX = 0& To (lngRpts - 1&)
          .AddNew
          ![dbs_id] = arr_varRpt(R_DID, lngX)
          ![dbs_name] = arr_varRpt(R_DNAM, lngX)
          ![rpt_id] = arr_varRpt(R_RID, lngX)
          ![rpt_name] = arr_varRpt(R_RNAM, lngX)
          If arr_varRpt(R_VID, lngX) <> vbNullString And arr_varRpt(R_VID, lngX) <> ";" Then
            ![vbcom_id] = Left(arr_varRpt(R_VID, lngX), 255)
          End If
          If arr_varRpt(R_PID, lngX) <> vbNullString And arr_varRpt(R_PID, lngX) <> ";" Then
            ![vbcomproc_id] = Left(arr_varRpt(R_PID, lngX), 255)
          End If
          If arr_varRpt(R_LIN1, lngX) <> vbNullString And arr_varRpt(R_LIN1, lngX) <> ";" Then
            ![rptvbcom_line] = Left(arr_varRpt(R_LIN1, lngX), 255)
          End If
          If arr_varRpt(R_CLIN1, lngX) <> vbNullString And arr_varRpt(R_CLIN1, lngX) <> ";" Then
            ![rptvbcom_code] = arr_varRpt(R_CLIN1, lngX)
          End If
          ![rptxref_datemodified] = Now()
          .Update
        Next
        .Close
      End With

blnSkip = False
If blnSkip = False Then
      ' ** Append zz_qry_Report_VBComponent_05 (zz_qry_Report_VBComponent_04 (zz_qry_Report_VBComponent_03
      ' ** (zz_qry_Report_VBComponent_02 (zz_tbl_Report_VBComponent_01, with vbcom_id, vbcomproc_id, rptvbcom_line,
      ' ** rptvbcom_codeline broken out through Pos4), with fields broken out through Pos8), with fields broken out
      ' ** through Pos11), without rptvbcom_code, with last of break-outs) to zz_tbl_Report_VBComponent_02.
      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_06a")
      qdf.Execute

      ' ** Update zz_qry_Report_VBComponent_06c (zz_qry_Report_VBComponent_06b (zz_qry_Report_VBComponent_04
      ' ** (zz_qry_Report_VBComponent_03 (zz_qry_Report_VBComponent_02 (zz_tbl_Report_VBComponent_01, with vbcom_id,
      ' ** vbcomproc_id, rptvbcom_line, rptvbcom_codeline broken out through Pos4), with fields broken out through
      ' ** Pos8), with fields broken out through Pos11), just rptvbcom_code, with last of break-outs), linked to
      ' ** zz_tbl_Report_VBComponent_02, with rptvbcom_code .._new fields).
      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_06d")
      qdf.Execute

      ' ** Update zz_qry_Report_VBComponent_06e (zz_tbl_Report_VBComponent_02, with add'l .._new fields).
      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_06f")
      qdf.Execute

      ' ** Update zz_qry_Report_VBComponent_06g (zz_tbl_Report_VBComponent_02, with final .._new fields).
      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_06h")
      qdf.Execute

      ' ** Append zz_qry_Report_VBComponent_07j (zz_qry_Report_VBComponent_07h (zz_qry_Report_VBComponent_07g
      ' ** (zz_qry_Report_VBComponent_07b (Union of zz_qry_Report_VBComponent_07a_01 (zz_tbl_Report_VBComponent_02,
      ' ** linked to tblVBComponent_Procedure, by vbcomproc_id01) - zz_qry_Report_VBComponent_07a_07
      ' ** (zz_tbl_Report_VBComponent_02, linked to tblVBComponent_Procedure, by vbcomproc_id07)), linked to
      ' ** tblVBComponent_Procedure, with .._line_beg, .._line_end, .._code_beg, .._code_end), linked to
      ' ** zz_qry_Report_VBComponent_07d (Union of zz_qry_Report_VBComponent_07c_01 (zz_tbl_Report_VBComponent_02,
      ' ** for rptvbcom_line01) - zz_qry_Report_VBComponent_07c_12 (zz_tbl_Report_VBComponent_02, for rptvbcom_line12))),
      ' ** linked to zz_qry_Report_VBComponent_07i (zz_qry_Report_VBComponent_07g (zz_qry_Report_VBComponent_07b
      ' ** (Union of zz_qry_Report_VBComponent_07a_01 (zz_tbl_Report_VBComponent_02, linked to tblVBComponent_Procedure,
      ' ** by vbcomproc_id01) - zz_qry_Report_VBComponent_07a_07 (zz_tbl_Report_VBComponent_02, linked to
      ' ** tblVBComponent_Procedure, by vbcomproc_id07)), linked to tblVBComponent_Procedure, with .._line_beg, .._line_end,
      ' ** .._code_beg, .._code_end), linked to zz_qry_Report_VBComponent_07f (Union of zz_qry_Report_VBComponent_07e_01
      ' ** (zz_tbl_Report_VBComponent_02, for rptvbcom_codeline01) - zz_qry_Report_VBComponent_07e_12
      ' ** (zz_tbl_Report_VBComponent_02, for rptvbcom_codeline12)))) to zz_tbl_Report_VBComponent_03.
      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_07k")
      qdf.Execute
End If  ' ** blnSkip.

      .Close
    End With

blnSkip = False
If blnSkip = False Then
    Rpt_ReportList_ProcVet False  ' ** Function: Below.
End If  ' ** blnSkip.

  Else
    dbs.Close
  End If  ' ** lngRpts.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_ReportList_ProcXRef = blnRetValx

End Function

Private Function Rpt_ReportList_ProcVet(Optional varForDelete As Variant) As Boolean
' ** Called by:
' **   Rpt_ReportList_ProcXRef(), Above

  Const THIS_PROC As String = "Rpt_ReportList_ProcVet"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngRpts As Long, arr_varRpt As Variant
  Dim lngRptXs As Long, arr_varRptX As Variant
  Dim strModName As String, strLine As String
  Dim blnForDelete As Boolean, blnDelete As Boolean, blnSkip As Boolean
  Dim lngThisDbsID As Long
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp00 As String, lngTmp01 As Long
  Dim lngX As Long

  ' ** Array: arr_varRpt().
  Const R_RVID   As Integer = 0
  Const R_DID    As Integer = 1
  Const R_DNAM   As Integer = 2
  Const R_RID    As Integer = 3
  Const R_RNAM   As Integer = 4
  Const R_VID    As Integer = 5
  Const R_VNAM   As Integer = 6
  Const R_LIN    As Integer = 7
  Const R_COD    As Integer = 8
  Const R_FND    As Integer = 9
  Const R_DEL    As Integer = 10
  Const R_PARID  As Integer = 11
  Const R_PARNAM As Integer = 12

  ' ** Array: arr_varRptX().
  Const RX_RXID1 As Integer = 0
  Const RX_RXID2 As Integer = 1
  Const RX_RXID3 As Integer = 2
  Const RX_DID   As Integer = 3
  Const RX_DNAM  As Integer = 4
  Const RX_RID   As Integer = 5
  Const RX_RNAM  As Integer = 6
  Const RX_VID   As Integer = 7
  Const RX_VNAM  As Integer = 8
  Const RX_PID   As Integer = 9
  Const RX_LIN   As Integer = 10
  Const RX_COD   As Integer = 11
  Const RX_TXT   As Integer = 12
  Const RX_REM   As Integer = 13
  Const RX_CAS   As Integer = 14
  Const RX_NA    As Integer = 15
  Const RX_FND   As Integer = 16

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  If IsMissing(varForDelete) = True Then
    blnForDelete = True  ' ** True: Delete from tblReport_VBComponent; False: Check entries in zz_tbl_Report_VBComponent_03.
  Else
    blnForDelete = varForDelete
  End If

  Set dbs = CurrentDb
  With dbs

    Select Case blnForDelete
    Case True

      ' ** tblReport_VBComponent, with add'l fields, by specified CurrentAppName().
      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_08")
      Set rst = qdf.OpenRecordset
      With rst
        If .BOF = True And .EOF = True Then
          ' ** Skip it!
          lngRpts = 0&
        Else
          .MoveLast
          lngRpts = .RecordCount
          .MoveFirst
          arr_varRpt = .GetRows(lngRpts)
          ' ****************************************************
          ' ** Array: arr_varRpt()
          ' **
          ' **   Field  Element  Name               Constant
          ' **   =====  =======  =================  ==========
          ' **     1       0     rptvbcom_id       R_RVID
          ' **     2       1     dbs_id            R_DID
          ' **     3       2     dbs_name          R_DNAM
          ' **     4       3     rpt_id            R_RID
          ' **     5       4     rpt_name          R_RNAM
          ' **     6       5     vbcom_id          R_VID
          ' **     7       6     vbcom_name        R_VNAM
          ' **     8       7     rptvbcom_line     R_LIN
          ' **     9       8     rptvbcom_code     R_COD
          ' **    10       9     rpt_fnd           R_FND
          ' **    11      10     rpt_del           R_DEL
          ' **    12      11     rpt_id_parent     R_PARID
          ' **    13      12     rpt_name_parent   R_PARNAM
          ' **
          ' ****************************************************
        End If
        .Close
      End With

      If lngRpts > 0& Then

blnSkip = False
If blnSkip = False Then
        Set vbp = Application.VBE.ActiveVBProject
        With vbp
          strModName = vbNullString
          For lngX = 0& To (lngRpts - 1&)
            If arr_varRpt(R_VNAM, lngX) <> strModName Then
              strModName = arr_varRpt(R_VNAM, lngX)
              Set vbc = .VBComponents(strModName)
            End If
            With vbc
              Set cod = .CodeModule
              With cod
                If IsNull(arr_varRpt(R_LIN, lngX)) = False Then
                  If arr_varRpt(R_LIN, lngX) > 0& And arr_varRpt(R_LIN, lngX) < .CountOfLines Then
                    strLine = Trim$(.Lines(arr_varRpt(R_LIN, lngX), 1))
                    intPos1 = InStr(strLine, (Chr(34) & arr_varRpt(R_RNAM, lngX) & Chr(34)))
                    If intPos1 = 0 Then
                      If IsNull(arr_varRpt(R_PARNAM, lngX)) = False Then
                        intPos1 = InStr(strLine, (Chr(34) & arr_varRpt(R_PARNAM, lngX) & Chr(34)))
                        If intPos1 = 0 Then
                          arr_varRpt(R_FND, lngX) = CBool(False)
                          arr_varRpt(R_DEL, lngX) = CBool(True)
                        Else
                          arr_varRpt(R_FND, lngX) = CBool(True)
                        End If
                      Else
                        arr_varRpt(R_FND, lngX) = CBool(False)
                        arr_varRpt(R_DEL, lngX) = CBool(True)
                      End If
                    Else
                      arr_varRpt(R_FND, lngX) = CBool(True)
                    End If
                  Else
                    arr_varRpt(R_FND, lngX) = CBool(False)
                    arr_varRpt(R_DEL, lngX) = CBool(True)
                  End If
                Else
                  arr_varRpt(R_FND, lngX) = CBool(False)
                  arr_varRpt(R_DEL, lngX) = CBool(True)
                End If
              End With  ' ** cod.
            End With  ' ** vbc.
          Next  ' ** lngX.
        End With  ' ** vbp
End If  ' ** blnSkip.

        lngTmp01 = 0&
        For lngX = 0& To (lngRpts - 1&)
          If arr_varRpt(R_DEL, lngX) = True Then
            lngTmp01 = lngTmp01 + 1&
          End If
        Next

        Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

        If lngTmp01 > 0& Then
          blnDelete = True
          Beep
          Debug.Print "'DELS: " & CStr(lngTmp01) & "  PROCEED?"
          Stop
          If blnDelete = True Then
            For lngX = 0& To (lngRpts - 1&)
              If arr_varRpt(R_DEL, lngX) = True Then
                ' ** Delete tblReport_VBComponent, by specified [rvcid].
                Set qdf = .QueryDefs("zz_qry_Report_VBComponent_01m")
                With qdf.Parameters
                  ![rvcid] = arr_varRpt(R_RVID, lngX)
                End With
                qdf.Execute
              End If
            Next
            Debug.Print "'DONE!  " & THIS_PROC & "()"
          Else
            Debug.Print "'NONE DELETED!  " & THIS_PROC & "()"
          End If
        Else
          Beep
          Debug.Print "'NO DELS!  " & THIS_PROC & "()"
        End If

      End If  ' ** lngRpts.

    Case False

      ' ** zz_tbl_Report_VBComponent_03, with add'l fields.
      Set qdf = .QueryDefs("zz_qry_Report_VBComponent_09")
      Set rst = qdf.OpenRecordset
      With rst
        .MoveLast
        lngRptXs = .RecordCount
        .MoveFirst
        arr_varRptX = .GetRows(lngRptXs)
        ' ***************************************************
        ' ** Array: arr_varRptX()
        ' **
        ' **   Field  Element  Name              Constant
        ' **   =====  =======  ================  ==========
        ' **     1       0     rptxref_id1       RX_RXID1
        ' **     2       1     rptxref_id2       RX_RXID2
        ' **     3       2     rptxref_id3       RX_RXID3
        ' **     4       3     dbs_id            RX_DID
        ' **     5       4     dbs_name          RX_DNAM
        ' **     6       5     rpt_id            RX_RID
        ' **     7       6     rpt_name          RX_RNAM
        ' **     8       7     vbcom_id          RX_VID
        ' **     9       8     vbcom_name        RX_VNAM
        ' **    10       9     vbcomproc_id      RX_PID
        ' **    11      10     rptvbcom_line     RX_LIN
        ' **    12      11     rptvbcom_code     RX_COD
        ' **    13      12     rptxref_text      RX_TXT
        ' **    14      13     rptxref_remark    RX_REM
        ' **    15      14     rptxref_case      RX_CAS
        ' **    16      15     rptxref_na        RX_NA
        ' **    17      16     rptxref_found     RX_FND
        ' **
        ' ***************************************************
        .Close
      End With

      If lngRptXs > 0& Then

        Set vbp = Application.VBE.ActiveVBProject
        With vbp
          strModName = vbNullString
          For lngX = 0& To (lngRptXs - 1&)
            If arr_varRptX(RX_VNAM, lngX) <> strModName Then
              strModName = arr_varRptX(RX_VNAM, lngX)
              Set vbc = .VBComponents(strModName)
            End If
            With vbc
              Set cod = .CodeModule
              With cod
                strLine = Trim$(.Lines(arr_varRptX(RX_LIN, lngX), 1))
                arr_varRptX(RX_TXT, lngX) = strLine
                If Left$(strLine, 1) <> "'" Then
                  intPos1 = InStr(strLine, arr_varRptX(RX_RNAM, lngX))
                  'If intPos1 > 0 Then
                    intPos2 = InStr(strLine, "Case ")
                    'If intPos2 > 0 And intPos2 < intPos1 Then
                    '  arr_varRptX(RX_CAS, lngX) = CBool(True)
                    'Else
                      intPos2 = InStr(strLine, " ")
                      If intPos2 > 0 Then
                        strTmp00 = Trim$(Left$(strLine, intPos2))
                        If IsNumeric(strTmp00) = False Then
                          ' ** No code line number means it might be a line continuation.
                          strTmp00 = Trim$(.Lines(arr_varRptX(RX_LIN, lngX) - 1, 1))
                          If Right$(strTmp00, 1) = "_" Then
                            arr_varRptX(RX_TXT, lngX) = Left$(strTmp00, (Len(strTmp00) - 1)) & strLine
                            intPos2 = InStr(strTmp00, " ")
                            If intPos2 > 0 Then
                              strTmp00 = Trim$(Left$(strTmp00, intPos2))
                              If IsNull(arr_varRptX(RX_COD, lngX)) = True Then
                                arr_varRptX(RX_COD, lngX) = CLng(strTmp00)
                              End If
                            End If
                          End If
                        Else
                          arr_varRptX(RX_COD, lngX) = CLng(strTmp00)
                        End If
                      End If
                      intPos2 = InStr(strLine, "' ** ")
                      If intPos2 > 0 And intPos2 < intPos1 Then
                        arr_varRptX(RX_REM, lngX) = CBool(True)
                      Else
                        intPos2 = InStr(strLine, "IsLoaded(")
                        If intPos2 > 0 And intPos2 < intPos1 Then
                          arr_varRptX(RX_NA, lngX) = CBool(True)
                        End If
                        intPos2 = InStr(strLine, "gstrReportQuerySpec =")
                        If intPos2 > 0 And intPos2 < intPos1 Then
                          arr_varRptX(RX_NA, lngX) = CBool(True)
                        End If
                        intPos2 = InStr(strLine, " ")
                        If intPos2 > 0 Then
                          strTmp00 = Trim$(Left$(strLine, intPos2))
                          If IsNumeric(strTmp00) = True Then
                            strTmp00 = Trim$(Mid$(strLine, intPos2))
                            If Left$(strTmp00, 3) = "If " Then
                              arr_varRptX(RX_NA, lngX) = CBool(True)
                            End If
                          End If
                        End If
                        ' ** Anything else to check for?

                      End If
                    'End If
                  'Else
                  '  arr_varRptX(RX_FND, lngX) = CBool(False)
                  'End If
                Else
                  arr_varRptX(RX_REM, lngX) = CBool(True)
                End If
              End With  ' ** cod.
            End With  ' ** vbc.
          Next  ' ** lngX.
        End With  ' ** vbp

        With dbs
          Set rst = .OpenRecordset("zz_tbl_Report_VBComponent_03", dbOpenDynaset, dbConsistent)
          With rst
            For lngX = 0& To (lngRptXs - 1&)
              If arr_varRptX(RX_FND, lngX) = False Then
                Debug.Print "'NOT FOUND? " & arr_varRptX(RX_VNAM, lngX) & "  LIN: " & CStr(arr_varRptX(RX_LIN, lngX))
              Else
                .FindFirst "[rptxref_id3] = " & CStr(arr_varRptX(RX_RXID3, lngX))
                If .NoMatch = False Then
                  If ![rptxref_remark] <> arr_varRptX(RX_REM, lngX) Then
                    .Edit
                    ![rptxref_remark] = True
                    ![rptxref_datemodified] = Now()
                    .Update
                  ElseIf ![rptxref_case] <> arr_varRptX(RX_CAS, lngX) Then
                    .Edit
                    ![rptxref_case] = True
                    ![rptxref_datemodified] = Now()
                    .Update
                  End If
                  If IsNull(![rptvbcom_code]) = True Then
                    If IsNull(arr_varRptX(RX_COD, lngX)) = False Then
                      .Edit
                      ![rptvbcom_code] = arr_varRptX(RX_COD, lngX)
                      ![rptxref_datemodified] = Now()
                      .Update
                    End If
                  Else
                    If IsNull(arr_varRptX(RX_COD, lngX)) = False Then
                      If ![rptvbcom_code] <> arr_varRptX(RX_COD, lngX) Then
                        .Edit
                        ![rptvbcom_code] = arr_varRptX(RX_COD, lngX)
                        ![rptxref_datemodified] = Now()
                        .Update
                      End If
                    End If
                  End If
                  If ![rptxref_na] <> arr_varRptX(RX_NA, lngX) Then
                    .Edit
                    ![rptxref_na] = arr_varRptX(RX_NA, lngX)
                    ![rptxref_datemodified] = Now()
                    .Update
                  End If
                  If IsNull(![rptxref_text]) = True Then
                    .Edit
                    ![rptxref_text] = arr_varRptX(RX_TXT, lngX)
                    ![rptxref_datemodified] = Now()
                    .Update
                  Else
                    If ![rptxref_text] <> arr_varRptX(RX_TXT, lngX) Then
                      .Edit
                      ![rptxref_text] = arr_varRptX(RX_TXT, lngX)
                      ![rptxref_datemodified] = Now()
                      .Update
                    End If
                  End If
                Else
                  Stop
                End If
              End If
            Next  ' ** lngX.
            .Close
          End With  ' ** rst.
        End With  ' ** dbs.

      End If  ' ** lngRpts.

      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

      Debug.Print "'DONE!  " & THIS_PROC & "()"
      DoEvents

    End Select

    .Close
  End With

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Beep

  Rpt_ReportList_ProcVet = blnRetValx

End Function

Public Function Rpt_Ctl_Assignment() As Boolean
' ** Not called.

  Const THIS_PROC As String = "Rpt_Ctl_Assignment"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, rpt As Access.Report
  Dim lngX As Long, lngY As Long, lngE As Long

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs

    .Close
  End With

' ** sec_index vs GroupLevel:
' **  0 Detail
' **  1 ReportHeader
' **  2 ReportFooter
' **  3 PageHeader
' **  4 PageHeader
' ** Starting with 5, each Section Index will correspond
' ** with a Group Header and/or Group Footer for those
' ** Groups designated to have a Header or Footer.
' ** So, GroupLevel 0 would have Sections 5 and 6,
' ** GroupLevel 1 would have 7 and 8,
' ** GroupLevel 2 9 and 10, etc.
' ** If no Group has a Header or Footer,
' ** there'll be no Sections greater than 4.
' ** If some Groups have a Header/Footer and some don't,
' ** the Sections don't close up; they skip and still
' ** take their proper place in the Section order.

  Beep

  Set rpt = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_Ctl_Assignment = blnRetValx

End Function

Public Function Rpt_Sec_Prop() As Boolean
' ** Not called.

  Const THIS_PROC As String = "Rpt_Sec_Prop"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, rpt As Access.Report, Sec As Access.Section
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngRpts As Long, arr_varRpt As Variant
  Dim strSecName As String, strModName As String
  Dim lngLines As Long, lngDecLines As Long, strLine As String
  Dim lngRptCnt As Long, lngSecCnt As Long, lngLineCnt As Long
  Dim intPos1 As Integer, intPos2 As Integer
  Dim blnRpt As Boolean, blnEdit As Boolean
  Dim strTmp00 As String
  Dim lngX As Long, lngY As Long

  ' ** Array: arr_varRpt().
  Const R_ID  As Integer = 0
  Const R_NAM As Integer = 1
  Const R_CNT As Integer = 2

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_qry_Report_Section_40 (tblReport_Section, just PageHeader, PageFooter), grouped by rpt_id, with cnt.
    Set qdf = .QueryDefs("zz_qry_Report_Section_41")
    ' ** zz_qry_Report_Section_37 (tblReport_Section, just sec_index = 5, 6,
    ' ** just discrepancies), grouped by rpt_id, with cnt.
    'Set qdf = .QueryDefs("zz_qry_Report_Section_37a")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRpts = .RecordCount
      .MoveFirst
      arr_varRpt = .GetRows(lngRpts)
      ' *********************************************
      ' ** Array: arr_varRpt()
      ' **
      ' **   Field  Element  Name        Constant
      ' **   =====  =======  ==========  ==========
      ' **     1       0     rpt_id      R_ID
      ' **     2       1     rpt_name    R_NAM
      ' **     3       2     cnt         R_CNT
      ' **
      ' *********************************************
      .Close
    End With
    .Close
  End With

  'lngRptCnt = 0&: lngSecCnt = 0&
  'For lngX = 0& To (lngRpts - 1&)
  '  DoCmd.OpenReport arr_varRpt(R_NAM, lngX), acViewDesign, , , acHidden
  '  Set rpt = Reports(0)
  '  With rpt
  '    blnRpt = False
  '    For lngY = 3& To 4& '0& To 18&
  '      strSecName = vbNullString: blnEdit = False
'On Error Resume Next
  '      strSecName = .Section(lngY).Name
'On Error GoTo 0
  '      If strSecName <> vbNullString Then
  '        If strSecName = "PageHeader" Then
  '          .Section(lngY).Name = acPageHeader  ' ** Because there's a property called 'PageHeader'!
  '          blnEdit = True
  '          blnRpt = True
  '        ElseIf strSecName = "PageFooter" Then
  '          .Section(lngY).Name = "PageFooterSection"  ' ** Because there's a property called 'PageFooter'!
  '          blnEdit = True
  '          blnRpt = True
  '        End If
  '        If blnEdit = True Then
  '          lngSecCnt = lngSecCnt + 1&
  '        End If
  '        'If Left$(strSecName, 5) = "Group" Then
  '        '  If Right$(strSecName, 1) = "0" Then
  '        '    lngSecCnt = lngSecCnt + 1&
  '        '    .Section(lngY).Name = Left$(strSecName, (Len(strSecName) - 1)) & "1"
  '        '  End If
  '        'End If
  '      End If
  '    Next  ' ** lngY.
  '    If blnRpt = True Then
  '      lngRptCnt = lngRptCnt + 1&
  '    End If
  '  End With  ' ** rpt.
  '  DoCmd.Close acReport, Reports(0).Name, acSaveYes
  'Next  ' ** lngX.

  'Debug.Print "'RPTS: " & CStr(lngRptCnt) & "  SECS: " & CStr(lngSecCnt)
  'DoEvents

'Stop
  lngRptCnt = 0&: lngLineCnt = 0&
  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    strModName = vbNullString
    lngRptCnt = 0&: lngSecCnt = 0&
    For lngX = 0& To (lngRpts - 1&)
      strModName = "Report_" & arr_varRpt(R_NAM, lngX)
      blnRpt = False
      Set vbc = .VBComponents(strModName)
      With vbc
        Set cod = .CodeModule
        With cod
          lngLines = .CountOfLines
          For lngY = 1& To lngLines
            strLine = .Lines(lngY, 1)
            blnEdit = False: strTmp00 = vbNullString: intPos1 = 0: intPos2 = 0
            If Trim$(strLine) <> vbNullString Then
              strTmp00 = strLine
              intPos1 = InStr(strTmp00, "PageHeader")
              If intPos1 > 0 Then
                intPos2 = InStr(strTmp00, acPageHeader)
                If intPos2 > 0 Then
                  If intPos2 <> intPos1 Then
                    Debug.Print "'LOOK AT: " & strModName
                  End If
                Else
                  strTmp00 = Left$(strTmp00, (intPos1 + 9)) & "Section" & Mid$(strTmp00, (intPos1 + 10))
                  blnEdit = True
                  blnRpt = True
                End If
              End If
              intPos1 = InStr(strTmp00, "PageFooter")
              If intPos1 > 0 Then
                intPos2 = InStr(strTmp00, "PageFooterSection")
                If intPos2 > 0 Then
                  If intPos2 <> intPos1 Then
                    Debug.Print "'LOOK AT: " & strModName
                  End If
                Else
                  strTmp00 = Left$(strTmp00, (intPos1 + 9)) & "Section" & Mid$(strTmp00, (intPos1 + 10))
                  blnEdit = True
                  blnRpt = True
                End If
              End If
              If blnEdit = True Then
                lngLineCnt = lngLineCnt + 1&
                .ReplaceLine lngY, strTmp00
              End If
            End If
          Next  ' ** lngY.
          'lngDecLines = .CountOfDeclarationLines
          'For lngY = 1& To lngDecLines
          '  strLine = .Lines(lngY, 1)
          '  blnEdit = False: strTmp00 = vbNullString
          '  If Trim$(strLine) <> vbNullString Then
          '    strTmp00 = strLine
          '    If Left$(strTmp00, 4) = "'VGC" And Right$(strTmp00, 8) = "CHANGES!" Then
          '      strTmp00 = "'VGC 01/25/2012: CHANGES!"
          '      .ReplaceLine lngY, strTmp00
          '      lngLineCnt = lngLineCnt + 1&
          '      blnRpt = True
          '      Exit For
          '    End If
          '  End If
          'Next  ' ** lngY.
          'For lngY = 1& To lngLines
          '  strLine = .Lines(lngY, 1)
          '  blnEdit = False: strTmp00 = vbNullString: intPos1 = 0: intPos2 = 0
          '  If Trim$(strLine) <> vbNullString Then
          '    strTmp00 = strLine
          '    intPos1 = InStr(strTmp00, "GroupHeader0")
          '    If intPos1 > 0 Then
          '      strTmp00 = Left$(strTmp00, (intPos1 + 10)) & "1" & Mid$(strTmp00, (intPos1 + 12))
          '      blnEdit = True
          '      blnRpt = True
          '    End If
          '    intPos2 = InStr(strTmp00, "GroupFooter0")
          '    If intPos2 > 0 Then
          '      strTmp00 = Left$(strTmp00, (intPos2 + 10)) & "1" & Mid$(strTmp00, (intPos2 + 12))
          '      blnEdit = True
          '      blnRpt = True
          '    End If
          '    If blnEdit = True Then
          '      lngLineCnt = lngLineCnt + 1&
          '      .ReplaceLine lngY, strTmp00
          '    End If
          '  End If
          'Next  ' ** lngY.
        End With  ' ** cod.
      End With  ' ** vbc.
      If blnRpt = True Then
        lngRptCnt = lngRptCnt + 1&
      End If
    Next  ' ** lngX.
  End With  ' ** vbp.

  Debug.Print "'RPTS: " & CStr(lngRptCnt) & "  LINES: " & CStr(lngLineCnt)
  Debug.Print "'DONE! " & THIS_PROC & "()"
  DoEvents

'RPTS: 80  LINES: 495
'DONE! Rpt_Sec_Prop()

'RPTS: 42  SECS: 68
'RPTS: 2  LINES: 3
'DONE! Rpt_Sec_Prop()

'RPTS: 42  LINES: 42
'DONE! Rpt_Sec_Prop()

  Set vbp = Nothing
  Set vbc = Nothing
  Set cod = Nothing
  Set Sec = Nothing
  Set rpt = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Beep

  Rpt_Sec_Prop = blnRetValx

End Function

Public Function Rpt_Props() As Boolean

  Const THIS_PROC As String = "Rpt_Props"

  Dim rpt As Access.Report, prp As Object

  blnRetValx = True

  Set rpt = Reports(0)
  With rpt
    Debug.Print "'PROPS: " & CStr(.Properties.Count)
    For Each prp In .Properties
      With prp
        Debug.Print "'" & .Name
      End With
    Next
  End With

  Beep

  Set prp = Nothing
  Set rpt = Nothing

  Rpt_Props = blnRetValx

End Function

Public Function Rpt_RptList_FrmName() As Boolean

  Const THIS_PROC As String = "Rpt_RptList_FrmName"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strLine As String, strDocVar As String, strModName As String
  Dim lngLineBeg As Long, lngLineEnd As Long
  Dim lngThisDbsID As Long
  Dim lngRecs As Long, lngEdits As Long
  Dim blnEdit As Boolean
  Dim intPos1 As Integer, intPos2 As Integer, intPos3 As Integer
  Dim varTmp00 As Variant, strTmp01 As String
  Dim lngX As Long, lngY As Long

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    lngEdits = 0&

    Set qdf = .QueryDefs("zz_qry_Report_VBComponent_95c_01")
    Set rst = qdf.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        blnRetValx = False
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          strModName = ![vbcom_name]
          strDocVar = ![rptxref_doc]
          lngLineBeg = ![vbcomproc_line_beg]
          lngLineEnd = ![vbcomproc_line_end]
          blnEdit = False
          Set vbp = Application.VBE.ActiveVBProject
          With vbp
            Set vbc = .VBComponents(strModName)
            With vbc
              Set cod = .CodeModule
              With cod
                For lngY = lngLineBeg To lngLineEnd
                  strLine = Trim(.Lines(lngY, 1))
                  If strLine <> vbNullString Then
                    If Left$(strLine, 1) <> "'" Then
                      intPos1 = InStr(strLine, strDocVar & " = ")
                      If intPos1 > 0 Then
                        intPos2 = InStr(intPos1, strLine, Chr(34))
                        If intPos2 > 0 Then
                          strTmp01 = Mid$(strLine, intPos2)
                          intPos3 = InStr(2, strTmp01, Chr(34))
                          If intPos3 < Len(strTmp01) Then strTmp01 = Left$(strTmp01, intPos3)
                          strTmp01 = Rem_Quotes(strTmp01)  ' ** Module Function: modStringFuncs.
                          varTmp00 = DLookup("[frm_id]", "tblForm", "[dbs_id] = " & CStr(lngThisDbsID) & " And " & _
                            "[frm_name] = '" & strTmp01 & "'")
                          If IsNull(rst![frm_id2]) = True Then
                            blnEdit = True
                          Else
                            If IsNull(rst![frm_name2]) = True Then
                              blnEdit = True
                            Else
                              If rst![frm_id2] <> varTmp00 Or rst![frm_name2] <> strTmp01 Then
                                blnEdit = True
                              End If
                            End If
                          End If
                          If blnEdit = True Then
                            rst.Edit
                            rst![frm_id2] = varTmp00
                            rst![frm_name2] = strTmp01
                            rst![rptxref_datemodified] = Now()
                            rst.Update
                            lngEdits = lngEdits + 1&
                          End If
                          Exit For
                        Else
                          intPos1 = InStr(strLine, "= strCallingForm")
                          If intPos1 > 0 Then
                            rst.Edit
                            rst![rptxref_doc] = "strCallingForm"
                            rst![rptxref_datemodified] = Now()
                            rst.Update
                          End If
                        End If
                      End If
                    End If
                  End If
                Next
              End With
              Set cod = Nothing
            End With
            Set vbc = Nothing
          End With
          Set vbp = Nothing
          If lngX < lngRecs Then .MoveNext
        Next
      End If
      .Close
    End With
    .Close
  End With

  Debug.Print "'DONE!  EDITS: " & CStr(lngEdits)

  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_RptList_FrmName = blnRetValx

End Function

Public Function Rpt_QryImport() As Boolean

  Const THIS_PROC As String = "Rpt_QryImport"

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
    ' ** zz_qry_Report_10x (tblQuery, just 'zz_qry_Report_..' queries in Trust.mdb),
    ' ** without 'zz_qry_Report_VBComponent_..'.
    Set qdf = .QueryDefs("zz_qry_Report_11x")
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

  Rpt_QryImport = blnRetVal

End Function

Private Function Rpt_ChkDocQrys(Optional varSkip As Variant) As Boolean

  Const THIS_PROC As String = "Rpt_ChkDocQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form
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

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Select Case IsMissing(varSkip)
  Case True
    blnSkip = True
  Case False
    blnSkip = varSkip
  End Select

  strPath = gstrDir_Dev
  strFile = CurrentAppName  ' ** Module Function: modFileUtilities.
  strPathFile = strPath & LNK_SEP & strFile

'On Error Resume Next
'  DoCmd.OpenForm "zz_frmStatus", acDesign, , , , acHidden
'  If ERR.Number <> 0 Then
'On Error GoTo 0
'    DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acForm, "zz_frmStatus", "zz_frmStatus"
'    DoEvents
'    CurrentDb.Containers("Forms").Documents.Refresh
'    Beep
'  Else
'On Error GoTo 0
'    DoCmd.Close acForm, "zz_frmStatus"
'  End If

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

    Debug.Print "'RPT DOC QRYS: " & CStr(lngQrys)
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
      Debug.Print "'ALL RPT DOC QRYS PRESENT!"
    End If

    Debug.Print "'DONE!"
    DoEvents

    Beep

  End If  ' ** blnSkip.

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Rpt_ChkDocQrys = blnRetValx

End Function
