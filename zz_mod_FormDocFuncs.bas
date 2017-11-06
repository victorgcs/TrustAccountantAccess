Attribute VB_Name = "zz_mod_FormDocFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "zz_mod_FormDocFuncs"

'VGC 03/26/2017: CHANGES!

'REMEMBER Painting PROPERTY WHEN RESIZING!

' ** AcMultiSelect enumeration: (my own)  ' ** Not currently used!
Public Const acMultiSelectNone     As Integer = 0
Public Const acMultiSelectSimple   As Integer = 1
Public Const acMultiSelectExtended As Integer = 2

' ** AcControlType enumeration:  (my own)
Public Const acNone              As Long = 99&
Public Const acDatasheetColumn   As Long = 115&
Public Const acEmptyCell         As Long = 127&
Public Const acWebBrowser        As Long = 128&
Public Const acNavigationControl As Long = 129&
Public Const acNavigationButton  As Long = 130&

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

' ** Array: arr_varSub().
Private lngSubs As Long, arr_varSub() As Variant
Private Const SUB_ELEMS As Integer = 9
Private Const SUB_FSID   As Integer = 0
Private Const SUB_PARID  As Integer = 1
Private Const SUB_PARNAM As Integer = 2
Private Const SUB_PARTYP As Integer = 3
Private Const SUB_CID    As Integer = 4
Private Const SUB_CNAM   As Integer = 5
Private Const SUB_CTYP   As Integer = 6
Private Const SUB_SUBID  As Integer = 7
Private Const SUB_SUBNAM As Integer = 8
Private Const SUB_SUBTYP As Integer = 9

Private blnRetValx As Boolean
' **

Public Function QuikFrmDoc() As Boolean
  Const THIS_PROC As String = "QuikFrmDoc"
  If Parse_File(CurrentBackendPath) = gstrDir_DevEmpty Or _
      (CurrentAppPath = gstrDir_Def And DCount("*", "account") = 2) Then ' ** Module Functions: modFileUtilities.
    If Frm_ChkDocQrys(False) = True Then  ' ** Function: Below.
      blnRetValx = Frm_Doc  ' ** Function: Below.
      blnRetValx = Frm_Sec_Doc  ' ** Function: Below.
Stop
      blnRetValx = Frm_Ctl_Doc  ' ** Function: Below.
      blnRetValx = Frm_RecSrc_Doc  ' ** Function: Below.
      blnRetValx = Frm_Shortcut_Doc  ' ** Function: Below.
      blnRetValx = Frm_Shortcut_Detail_Doc  ' ** Function: Below.
      blnRetValx = Frm_Subform_Doc  ' ** Function: Below.
      blnRetValx = Frm_Ctl_RowSource_Doc  ' ** Function: Below.
      blnRetValx = Frm_Image_Doc  ' ** Function: Below.
Stop
      blnRetValx = Frm_Ctl_Specs_Doc  ' ** Function: Below.
      DoEvents
      DoBeeps  ' ** Module Function: modWindowFunctions.
      Debug.Print "'FINISHED!"
    Else
      blnRetValx = False
      Beep
      Debug.Print "'FAILED Frm_ChkDocQrys()!"
    End If
  Else
    blnRetValx = False
    Beep
    Debug.Print "'NOT LINKED TO EMPTY!"
  End If
  QuikFrmDoc = blnRetValx
End Function

Private Function Frm_Doc() As Boolean
' ** Document all forms in Trust Accountant to tblForm,
' ** and specifications to tblForm_Specification.
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ctr As DAO.Container, doc As DAO.Document
  Dim frm1 As Access.Form, frm2 As Access.Form, ctl As Access.Control, prp As Object
  Dim strForm As String, strSubs As String
  Dim lngFrms As Long, arr_varFrm() As Variant
  Dim lngSubs As Long, lngRecs As Long
  Dim intSaveMode As Integer
  Dim lngThisDbsID As Long
  Dim blnAdd As Boolean, blnFound As Boolean, blnDelete As Boolean
  Dim lngFrmXs As Long, arr_varFrmX As Variant, lngFrmYs As Long, arr_varFrmY As Variant, lngFrmZs As Long, arr_varFrmZ() As Variant
  Dim lngTmp01 As Long
  Dim lngX As Long, lngY As Long, lngE As Long, lngF As Long

  ' ** Array: arr_varFrm().
  Const F_ELEMS As Integer = 15  ' ** Array's first-element UBound().
  Const F_DID     As Integer = 0
  Const F_FID     As Integer = 1
  Const F_OTYP    As Integer = 2
  Const F_FNAM    As Integer = 3
  Const F_CTLS    As Integer = 4
  Const F_CAP     As Integer = 5
  Const F_HAS_SUB As Integer = 6
  Const F_IS_SUB  As Integer = 7
  Const F_ACT     As Integer = 8
  Const F_HID     As Integer = 9
  Const F_DSC     As Integer = 10
  Const F_TAG     As Integer = 11
  Const F_PARSUB  As Integer = 12
  Const F_SUBS    As Integer = 13
  Const F_DAT     As Integer = 14
  Const F_FND     As Integer = 15

  ' ** Array: arr_varFrmX().
  Const FX_DID As Integer = 0
  Const FX_FID As Integer = 1

  ' ** Array: arr_varFrmY().
  Const FY_DID  As Integer = 0
  Const FY_FID  As Integer = 1
  Const FY_FNAM As Integer = 2
  Const FY_SUBS As Integer = 3
  Const FY_FNDS As Integer = 4

  ' ** Array: arr_varFrmZ().
  Const FZ_ELEMS As Integer = 3  ' ** Array's first-element UBound().
  Const FZ_DID  As Integer = 0
  Const FZ_FID  As Integer = 1
  Const FZ_FNAM As Integer = 2
  Const FZ_FND  As Integer = 3

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngFrms = 0&
  ReDim arr_varFrm(F_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    ' ** Add all the forms to tblForm.
    Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbConsistent)
    Set ctr = .Containers("Forms")
    With ctr

      For Each doc In .Documents
        strForm = vbNullString: strSubs = vbNullString: blnAdd = False
        lngSubs = 0&
        intSaveMode = acSaveNo
        With doc
          strForm = .Name
          lngFrms = lngFrms + 1&
          lngE = lngFrms - 1&
          ReDim Preserve arr_varFrm(F_ELEMS, lngE)
          ' *******************************************************
          ' ** Array: arr_varFrm()
          ' **
          ' **   Field  Element  Name                Constant
          ' **   =====  =======  ==================  ============
          ' **     1       0     dbs_id              F_DID
          ' **     2       1     frm_id              F_FID
          ' **     3       2     objtype_type        F_OTYP
          ' **     4       3     frm_name            F_FNAM
          ' **     5       4     frm_controls        F_CTLS
          ' **     6       5     frm_caption         F_CAP
          ' **     7       6     frm_hassub          F_HAS_SUB
          ' **     8       7     frm_issub           F_IS_SUB
          ' **     9       8     frm_active          F_ACT
          ' **    10       9     sec_hidden          F_HID
          ' **    11      10     frm_description     F_DSC
          ' **    12      11     frm_tag             F_TAG
          ' **    13      12     frm_parent_sub      F_PARSUB
          ' **    14      13     frm_subs            F_SUBS
          ' **    15      14     frm_datemodified    F_DAT
          ' **    16      15     Found               F_FND
          ' **
          ' *******************************************************
          arr_varFrm(F_DID, lngE) = lngThisDbsID
          arr_varFrm(F_FID, lngE) = CLng(0)
          arr_varFrm(F_OTYP, lngE) = acForm
          arr_varFrm(F_FNAM, lngE) = strForm
          arr_varFrm(F_PARSUB, lngE) = Null
          arr_varFrm(F_DAT, lngE) = Now()
          arr_varFrm(F_FND, lngE) = CBool(False)
          arr_varFrm(F_SUBS, lngE) = lngSubs
          DoCmd.OpenForm strForm, acDesign, , , , acHidden
          Set frm1 = Forms(strForm)
          rst.FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_name] = '" & strForm & "'"
'frm_id
          If rst.NoMatch = True Then
            blnAdd = True
            rst.AddNew
'dbs_id
            rst("dbs_id") = lngThisDbsID
          Else
            arr_varFrm(F_FID, lngE) = rst("frm_id")
            rst.Edit
          End If
'frm_name
          rst("frm_name") = strForm
'objtype_type
          rst("objtype_type") = acForm
'frm_controls
          rst("frm_controls") = frm1.Controls.Count
'frm_subs
          lngSubs = 0&
          For Each ctl In frm1.Controls
            With ctl
              If .ControlType = acSubform Then
                lngSubs = lngSubs + 1&
              End If
            End With
          Next
If lngSubs > 10& Then
Stop
End If
          rst("frm_subs") = lngSubs
          arr_varFrm(F_SUBS, lngE) = lngSubs
          rst("frm_parent_sub") = Null  ' ** To be updated below.
'frm_hassub
          If lngSubs > 0& Then
            rst("frm_hassub") = True
            arr_varFrm(F_HAS_SUB, lngE) = CBool(True)
          Else
            rst("frm_hassub") = False
            arr_varFrm(F_HAS_SUB, lngE) = CBool(False)
          End If
'frm_caption
          If IsNull(frm1.Caption) = False Then
            If frm1.Caption <> vbNullString Then
              rst("frm_caption") = frm1.Caption
            Else
              If blnAdd = False Then rst("frm_caption") = Null
            End If
          Else
            If blnAdd = False Then rst("frm_caption") = Null
          End If
          If IsNull(frm1.Tag) = False Then
            If frm1.Tag <> vbNullString Then
'frm_tag
              rst("frm_tag") = frm1.Tag
              If InStr(frm1.Tag, "Is Subform") > 0 Then
'frm_issub
                rst("frm_issub") = True
                arr_varFrm(F_IS_SUB, lngE) = CBool(True)
              Else
                If Right$(frm1.Name, 4) = "_Sub" Then
                  rst("frm_issub") = True
                  arr_varFrm(F_IS_SUB, lngE) = CBool(True)
                Else
                  rst("frm_issub") = False
                  arr_varFrm(F_IS_SUB, lngE) = CBool(False)
                End If
              End If
              If InStr(frm1.Tag, "Not Used") > 0 Then
'frm_active
                rst("frm_active") = False
              Else
                rst("frm_active") = True
              End If
            Else
              If Left$(frm1.Name, 5) = "Menu_" Then
                ' ** Add it's tag!
                intSaveMode = acSaveYes
                frm1.Tag = "Menu" & IIf(frm1.Tag = vbNullString, vbNullString, ";" & frm1.Tag)
                If InStr(frm1.Tag, "Is Subform") = 0 Then
                  rst("frm_issub") = False
                  arr_varFrm(F_IS_SUB, lngE) = CBool(False)
                End If
              End If
              If arr_varFrm(F_IS_SUB, lngE) = True Or Right$(frm1.Name, 4) = "_Sub" Then
                ' ** Add it's tag!
                intSaveMode = acSaveYes
                frm1.Tag = "Is Subform" & IIf(frm1.Tag = vbNullString, vbNullString, ";" & frm1.Tag)
                rst("frm_issub") = True
                arr_varFrm(F_IS_SUB, lngE) = CBool(True)
              End If
              If arr_varFrm(F_HAS_SUB, lngE) = True Then
                ' ** Add it's tag!
                intSaveMode = acSaveYes
                frm1.Tag = "Has Subform" & IIf(frm1.Tag = vbNullString, vbNullString, ";" & frm1.Tag)
                If InStr(frm1.Tag, "Is Subform") = 0 Then
                  rst("frm_issub") = False
                  arr_varFrm(F_IS_SUB, lngE) = CBool(False)
                End If
              End If
              rst("frm_active") = True
            End If
          Else
            ' ** It doesn't seem to ever be Null, just vbNullString.
          End If
'frm_parent_sub
          If rst("frm_issub") = True Then
            ' ** Processing moved to end.
          End If
          If rst("frm_hassub") = True Then
            For Each ctl In frm1.Controls
              With ctl
                If .ControlType = acSubform Then
                  If strSubs = vbNullString Then
                    strSubs = .SourceObject
                  Else
                    strSubs = strSubs & ";" & .SourceObject
                  End If
                End If
              End With
            Next
            If strSubs <> vbNullString Then
              rst("frm_parent_sub") = strSubs  ' ** Subs of parents get listed here.
            End If
If strSubs = vbNullString Then
Stop
End If
          Else
            ' ** It was nulled out first, above.
          End If
          For Each prp In .Properties
            With prp
'frm_description
              If .Name = "Description" Then
                If doc.Properties(.Name) <> vbNullString Then
                  rst("frm_description") = doc.Properties(.Name)
                End If
                Exit For
              End If
            End With
          Next
'sec_hidden
          Select Case blnAdd
          Case True
            rst("sec_hidden") = False
          Case False
            ' ** Leave it as stands.
          End Select
'frm_datemodified
          rst("frm_datemodified") = Now()
          rst.Update
          rst.Bookmark = rst.LastModified
          If arr_varFrm(F_FID, lngE) = 0& Then
            arr_varFrm(F_FID, lngE) = rst("frm_id")
          End If
          Frm_Specs_Doc frm1, dbs, CLng(arr_varFrm(F_FID, lngE))  ' ** Function: Below.
          DoCmd.Close acForm, strForm, intSaveMode
        End With  ' ** doc.
        Set prp = Nothing
        Set ctl = Nothing
        Set frm1 = Nothing
      Next  ' ** doc.
      Set doc = Nothing
    End With  ' ** ctr.
    Set ctr = Nothing
    DoEvents
    rst.Close
    Set rst = Nothing
    .Close
  End With  ' ** dbs.
  Set dbs = Nothing
  DoEvents

  Set dbs = CurrentDb
  With dbs
    ' ** Empty zz_tbl_Form_Doc.
    Set qdf = .QueryDefs("zz_qry_Form_01b")
    qdf.Execute
    Set rst = .OpenRecordset("zz_tbl_Form_Doc", dbOpenDynaset, dbAppendOnly)
    With rst
      For lngX = 0& To (lngFrms - 1&)
        .AddNew
        For lngY = 0& To F_ELEMS
          .Fields(lngY + 1&) = arr_varFrm(lngY, lngX)  ' ** Skip frmdoc_id.
        Next
        .Update
      Next  ' ** lngX.
      .Close
    End With  ' ** rst.
    .Close
  End With  ' ** dbs.
  Set rst = Nothing
  Set dbs = Nothing

  lngY = 0&
  ReDim arr_varFrm(F_ELEMS, 0)
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** zz_tbl_Form_Doc, just frm_issub = True.
    Set qdf = .QueryDefs("zz_qry_Form_02")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngFrmXs = .RecordCount
      .MoveFirst
      arr_varFrmX = .GetRows(lngFrmXs)
      .Close
    End With
    Set rst = Nothing
    DoEvents

    ' ** zz_tbl_Form_Doc, just frm_hassub = True.
    Set qdf = .QueryDefs("zz_qry_Form_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngFrmYs = .RecordCount
      .MoveFirst
      arr_varFrmY = .GetRows(lngFrmYs)
      .Close
    End With
    Set rst = Nothing
    DoEvents
    lngTmp01 = 0&

    ' ** Find each subform.
    For lngX = 0& To (lngFrmXs - 1&)  ' 60& To (lngFrmXs - 1&) '0& To 59&
      lngTmp01 = lngX
      ReDim arr_varFrm(F_ELEMS, 0)
      Set rst = .OpenRecordset("zz_tbl_Form_Doc", dbOpenDynaset, dbReadOnly)
      With rst
        .FindFirst "[dbs_id] = " & CStr(arr_varFrmX(FX_DID, lngX)) & " And [frm_id] = " & CStr(arr_varFrmX(FX_FID, lngX))
        If .NoMatch = False Then
          For lngY = 0& To F_ELEMS
            arr_varFrm(lngY, 0) = .Fields(lngY + 1&)
          Next
        Else
          Stop
        End If
        .Close
      End With
      Set rst = Nothing
      DoEvents
      Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbConsistent)
      strForm = vbNullString
      With rst
        .MoveFirst
        .FindFirst "[dbs_id] = " & CStr(arr_varFrm(F_DID, 0)) & " And [frm_id] = " & CStr(arr_varFrm(F_FID, 0))
        If .NoMatch = False Then
          .Edit
          .Fields("frm_parent_sub") = Null  ' ** To start fresh.
          .Update
        Else
          Stop
        End If
        .Close
      End With  ' ** rst.
      Set rst = Nothing
      DoEvents
      ' ** Now find each form with subforms.
      For lngY = 0& To (lngFrmYs - 1&)
        If arr_varFrmY(FY_FID, lngY) <> arr_varFrm(F_FID, 0) Then  ' ** Don't open the child form.
          If arr_varFrmY(FY_SUBS, lngY) > 0& And arr_varFrmY(FY_FNDS, lngY) < arr_varFrmY(FY_SUBS, lngY) Then  ' ** frm_subs, frm_foundsubs.
            DoCmd.OpenForm arr_varFrmY(FY_FNAM, lngY), acDesign, , , acFormPropertySettings, acHidden
            Set frm2 = Forms(0) '(arr_varFrmY(FY_FNAM, lngY))
            With frm2
              For Each ctl In .Controls
                With ctl
                  If .ControlType = acSubform Then
                    If .SourceObject = arr_varFrm(F_FNAM, 0) Then
                      ' ** This is the parent of the lngX subform.
                      strForm = frm2.Name
                      arr_varFrm(F_PARSUB, 0) = strForm
                      arr_varFrmY(FY_FNDS, lngY) = arr_varFrmY(FY_FNDS, lngY) + 1&
                      Exit For
                    End If
                  End If
                End With  ' ** ctl.
              Next  ' ** ctl.
            End With  ' ** frm2.
            DoCmd.Close acForm, arr_varFrmY(FY_FNAM, lngY), acSaveNo
            Set ctl = Nothing
            Set frm2 = Nothing
            DoEvents
          End If  ' ** F_SUBS.
        End If  ' ** F_FNAM.
        If strForm <> vbNullString Then Exit For
      Next  ' ** lngY.
      If strForm = vbNullString Then
      Stop
      End If
      If IsNull(arr_varFrm(F_PARSUB, 0)) = False Then
        Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbConsistent)
        With rst
          .MoveFirst
          .FindFirst "[dbs_id] = " & CStr(arr_varFrm(F_DID, 0)) & " And [frm_id] = " & CStr(arr_varFrm(F_FID, 0))
          If .NoMatch = False Then
            .Edit
            .Fields("frm_parent_sub") = arr_varFrm(F_PARSUB, 0)  ' ** Put in the new list of parents.
            .Update
          Else
            Stop
          End If
        End With
      End If
      rst.Close
      Set rst = Nothing
      'Debug.Print "'lngTmp01: " & CStr(lngTmp01)
    Next  ' ** lngX.
    .Close
  End With
  Set dbs = Nothing
  lngY = 0&
  ReDim arr_varFrm(F_ELEMS, 0)

  lngFrmZs = 0&
  ReDim arr_varFrmZ(FZ_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    ' ** zz_tbl_Form_Doc, all forms, with frm_foundx.
    Set qdf = .QueryDefs("zz_qry_Form_04")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngTmp01 = .RecordCount
      .MoveFirst
      For lngX = 1& To lngTmp01
        lngFrmZs = lngFrmZs + 1&
        lngF = lngFrmZs - 1&
        ReDim Preserve arr_varFrmZ(FZ_ELEMS, lngF)
        arr_varFrmZ(FZ_DID, lngF) = .Fields("dbs_id")
        arr_varFrmZ(FZ_FID, lngF) = .Fields("frm_id")
        arr_varFrmZ(FZ_FNAM, lngF) = .Fields("frm_name")
        arr_varFrmZ(FZ_FND, lngF) = CBool(False)
        If lngX < lngTmp01 Then .MoveNext
      Next
      .Close
    End With
    Set rst = Nothing
    DoEvents

    Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbConsistent)
    ' ** Look for obsolete forms.
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If .Fields("dbs_id") = lngThisDbsID Then
          blnFound = False
          For lngY = 0& To (lngFrmZs - 1&)
            If .Fields("dbs_id") = arr_varFrmZ(FZ_DID, lngY) And .Fields("frm_name") = arr_varFrmZ(FZ_FNAM, lngY) Then
              blnFound = True
              arr_varFrmZ(FZ_FND, lngY) = CBool(True)
              Exit For
            End If
          Next
          If blnFound = False Then
            Debug.Print "'FRM NOT FOUND!  " & .Fields("frm_name")
Stop
            lngFrmZs = lngFrmZs + 1&
            lngF = lngFrmZs - 1&
            ReDim Preserve arr_varFrmZ(FZ_ELEMS, lngF)
            arr_varFrmZ(FZ_DID, lngF) = .Fields("dbs_id")
            arr_varFrmZ(FZ_FID, lngF) = .Fields("frm_id")
            arr_varFrmZ(FZ_FNAM, lngF) = .Fields("frm_name")
            arr_varFrmZ(FZ_FND, lngF) = CBool(False)
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With
    Set rst = Nothing
    .Close
  End With
  Set dbs = Nothing

  If lngFrmZs > 0& Then
    Set dbs = CurrentDb
    With dbs
      For lngX = 0& To (lngFrmZs - 1&)
        If arr_varFrmZ(FZ_FND, lngX) = False And arr_varFrmZ(FZ_DID, lngX) = lngThisDbsID Then
          blnDelete = True
          Debug.Print "'DEL FRM? " & arr_varFrmZ(FZ_FNAM, lngX)
Stop
          If blnDelete = True Then
            ' ** Delete tblForm, by specified [frmid].
            Set qdf = dbs.QueryDefs("zz_qry_Form_01a")
            qdf.Parameters("frmid") = arr_varFrmZ(FZ_FID, lngX)
            qdf.Execute dbFailOnError
          End If
        End If
      Next
      .Close
    End With
    Set dbs = Nothing
  End If

  ' ** Make sure no forms are left open.
  If Forms.Count > 0 Then
    Do While Forms.Count > 0
      DoCmd.Close acForm, Forms(0).Name
    Loop
  End If

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set ctl = Nothing
  Set prp = Nothing
  Set frm1 = Nothing
  Set frm2 = Nothing
  Set doc = Nothing
  Set ctr = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Doc = blnRetValx

End Function

Private Function Frm_Sec_Doc() As Boolean
' ** Document all form Sections to tblForm_Section.
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_Sec_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Form
  Dim lngFrms As Long, arr_varFrm() As Variant
  Dim lngSecs As Long, arr_varSec() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim strSection As String
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean, blnDelete As Boolean
  Dim lngRecs As Long
  Dim lngX As Long, lngY As Long, lngElemS As Long, lngE As Long

  ' ** Array: arr_varFrm().
  Const F_ELEMS As Integer = 14  ' ** Array's first-element UBound().
  Const F_DID     As Integer = 0
  Const F_FID     As Integer = 1
  Const F_OTYP    As Integer = 2
  Const F_FNAM    As Integer = 3
  Const F_CTLS    As Integer = 4
  Const F_CAP     As Integer = 5
  Const F_HAS_SUB As Integer = 6
  Const F_IS_SUB  As Integer = 7
  Const F_ACT     As Integer = 8
  Const F_HID     As Integer = 9
  Const F_DSC     As Integer = 10
  Const F_TAG     As Integer = 11
  Const F_PARSUB  As Integer = 12
  Const F_SUBS    As Integer = 13
  Const F_DAT     As Integer = 14

  ' ** Array: arr_varSec().
  Const S_ELEMS As Integer = 20  ' ** Array's first-element UBound().
  Const S_DID    As Integer = 0
  Const S_FID    As Integer = 1
  Const S_FNAM   As Integer = 2
  Const S_SNAM   As Integer = 3
  Const S_IDX    As Integer = 4
  Const S_BKCLR  As Integer = 5   'BackColor
  Const S_GROW   As Integer = 6   'CanGrow
  Const S_SHRINK As Integer = 7   'CanShrink
  Const S_DISP   As Integer = 8   'DisplayWhen
  Const S_NEWPG  As Integer = 9   'ForceNewPage
  Const S_HT     As Integer = 10  'Height
  Const S_KEEP   As Integer = 11  'KeepTogether
  Const S_NEWROW As Integer = 12  'NewRowOrCol
  Const S_CLK    As Integer = 13  'OnClick
  Const S_DBLCLK As Integer = 14  'OnDblClick
  Const S_MDWN   As Integer = 15  'OnMouseDown
  Const S_MMOV   As Integer = 16  'OnMouseMove
  Const S_MUP    As Integer = 17  'OnMouseUp
  Const S_SPCEF  As Integer = 18  'SpecialEffect
  Const S_TAG    As Integer = 19  'Tag
  Const S_VIS    As Integer = 20  'Visible

  ' ** Array: arr_varDel().
  Const D_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const D_DID  As Integer = 0
  Const D_FID  As Integer = 1
  Const D_FNAM As Integer = 2
  Const D_SID  As Integer = 3
  Const D_SNAM As Integer = 4

  Const SEC_MAX As Long = 15&  '5&  ' ** Reports may have many more.

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    lngFrms = 0&
    ReDim arr_varFrm(F_ELEMS, 0)

    ' ** Get a list of all forms.
    Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngFrms = lngFrms + 1&
          lngE = lngFrms - 1&
          ReDim Preserve arr_varFrm(F_ELEMS, lngE)
          ' ******************************************************
          ' ** Array: arr_varFrm()
          ' **
          ' **   Field  Element  Name                Constant
          ' **   =====  =======  ==================  ===========
          ' **     1       0     dbs_id              F_DID
          ' **     2       1     frm_id              F_FID
          ' **     3       2     objtype_type        F_OTYP
          ' **     4       3     frm_name            F_FNAM
          ' **     5       4     frm_controls        F_CTLS
          ' **     6       5     frm_caption         F_CAP
          ' **     7       6     frm_hassub          F_HAS_SUB
          ' **     8       7     frm_issub           F_IS_SUB
          ' **     9       8     frm_active          F_ACT
          ' **    10       9     sec_hidden          F_HID
          ' **    11      10     frm_description     F_DSC
          ' **    12      11     frm_tag             F_TAG
          ' **    13      12     frm_parent_sub      F_PARSUB
          ' **    14      13     frm_subs            F_SUBS
          ' **    15      14     frm_datemodified    F_DAT
          ' **
          ' ******************************************************
          arr_varFrm(F_DID, lngE) = ![dbs_id]
          arr_varFrm(F_FID, lngE) = ![frm_id]
          arr_varFrm(F_OTYP, lngE) = ![objtype_type]
          arr_varFrm(F_FNAM, lngE) = ![frm_name]
          arr_varFrm(F_CTLS, lngE) = ![frm_controls]
          arr_varFrm(F_CAP, lngE) = ![frm_caption]
          arr_varFrm(F_HAS_SUB, lngE) = ![frm_hassub]
          arr_varFrm(F_IS_SUB, lngE) = ![frm_issub]
          arr_varFrm(F_ACT, lngE) = ![frm_active]
          arr_varFrm(F_HID, lngE) = ![sec_hidden]
          arr_varFrm(F_DSC, lngE) = ![frm_description]
          arr_varFrm(F_TAG, lngE) = ![frm_tag]
          arr_varFrm(F_PARSUB, lngE) = ![frm_parent_sub]
          arr_varFrm(F_SUBS, lngE) = ![frm_subs]
          arr_varFrm(F_DAT, lngE) = ![frm_datemodified]
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    lngSecs = 0&
    ReDim arr_varSec(S_ELEMS, 0)

    For lngX = 0& To (lngFrms - 1&)
      DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
      Set frm = Forms(arr_varFrm(F_FNAM, lngX))
      With frm
        ' ** Form Sections.
        For lngY = 0& To (SEC_MAX - 1&)
          lngElemS = lngY
On Error Resume Next
          strSection = .Section(lngElemS).Name
          If ERR = 0 Then
On Error GoTo 0
            lngSecs = lngSecs + 1&
            lngE = lngSecs - 1&
            ReDim Preserve arr_varSec(S_ELEMS, lngE)
            arr_varSec(S_DID, lngE) = arr_varFrm(F_DID, lngX)
            arr_varSec(S_FID, lngE) = arr_varFrm(F_FID, lngX)
            arr_varSec(S_FNAM, lngE) = frm.Name
            arr_varSec(S_SNAM, lngE) = strSection
            arr_varSec(S_IDX, lngE) = lngY
            arr_varSec(S_BKCLR, lngE) = .Section(lngElemS).BackColor
            arr_varSec(S_GROW, lngE) = .Section(lngElemS).CanGrow
            arr_varSec(S_SHRINK, lngE) = .Section(lngElemS).CanShrink
            arr_varSec(S_DISP, lngE) = .Section(lngElemS).DisplayWhen
            arr_varSec(S_NEWPG, lngE) = .Section(lngElemS).ForceNewPage
            arr_varSec(S_HT, lngE) = .Section(lngElemS).Height
            arr_varSec(S_KEEP, lngE) = .Section(lngElemS).KeepTogether
            arr_varSec(S_NEWROW, lngE) = .Section(lngElemS).NewRowOrCol
            arr_varSec(S_CLK, lngE) = .Section(lngElemS).OnClick
            arr_varSec(S_DBLCLK, lngE) = .Section(lngElemS).OnDblClick
            arr_varSec(S_MDWN, lngE) = .Section(lngElemS).OnMouseDown
            arr_varSec(S_MMOV, lngE) = .Section(lngElemS).OnMouseMove
            arr_varSec(S_MUP, lngE) = .Section(lngElemS).OnMouseUp
            arr_varSec(S_SPCEF, lngE) = .Section(lngElemS).SpecialEffect
            arr_varSec(S_TAG, lngE) = .Section(lngElemS).Tag
            arr_varSec(S_VIS, lngE) = .Section(lngElemS).Visible
          Else
On Error GoTo 0
            Exit For
          End If
        Next
      End With
      DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX), acSaveNo
    Next

    For lngX = 0& To (lngSecs - 1&)
      If IsNull(arr_varSec(S_GROW, lngX)) = True Then
        arr_varSec(S_GROW, lngX) = vbNullString
      End If
      If IsNull(arr_varSec(S_SHRINK, lngX)) = True Then
        arr_varSec(S_SHRINK, lngX) = vbNullString
      End If
      If IsNull(arr_varSec(S_CLK, lngX)) = True Then
        arr_varSec(S_CLK, lngX) = vbNullString
      End If
      If IsNull(arr_varSec(S_DBLCLK, lngX)) = True Then
        arr_varSec(S_DBLCLK, lngX) = vbNullString
      End If
      If IsNull(arr_varSec(S_MDWN, lngX)) = True Then
        arr_varSec(S_MDWN, lngX) = vbNullString
      End If
      If IsNull(arr_varSec(S_MMOV, lngX)) = True Then
        arr_varSec(S_MMOV, lngX) = vbNullString
      End If
      If IsNull(arr_varSec(S_MUP, lngX)) = True Then
        arr_varSec(S_MUP, lngX) = vbNullString
      End If
      If IsNull(arr_varSec(S_TAG, lngX)) = True Then
        arr_varSec(S_TAG, lngX) = vbNullString
      End If
    Next

    Set rst = dbs.OpenRecordset("tblForm_Section", dbOpenDynaset, dbConsistent)
    With rst

      For lngX = 0& To (lngSecs - 1&)
        .FindFirst "[dbs_id] = " & CStr(arr_varSec(S_DID, lngX)) & " And [sec_name] = '" & arr_varSec(S_SNAM, lngX) & "' And " & _
          "[frm_id] = " & CStr(arr_varSec(S_FID, lngX))
        If .NoMatch = True Then
          .AddNew
          ![dbs_id] = arr_varSec(S_DID, lngX)
          ![frm_id] = arr_varSec(S_FID, lngX)
          ![objtype_type] = acForm
          ![sec_name] = arr_varSec(S_SNAM, lngX)
          ![sec_index] = arr_varSec(S_IDX, lngX)
          ![sec_backcolor] = arr_varSec(S_BKCLR, lngX)
          If arr_varSec(S_GROW, lngX) <> vbNullString Then
            ![sec_cangrow] = arr_varSec(S_GROW, lngX)
          End If
          If arr_varSec(S_SHRINK, lngX) <> vbNullString Then
            ![sec_canshrink] = arr_varSec(S_SHRINK, lngX)
          End If
          ![sec_displaywhen] = arr_varSec(S_DISP, lngX)
          ![sec_forcenewpage] = arr_varSec(S_NEWPG, lngX)
          ![sec_height] = arr_varSec(S_HT, lngX)
          ![sec_keeptogether] = arr_varSec(S_KEEP, lngX)
          ![sec_newroworcol] = arr_varSec(S_NEWROW, lngX)
          If arr_varSec(S_CLK, lngX) <> vbNullString Then
            ![sec_onclick] = arr_varSec(S_CLK, lngX)
          End If
          If arr_varSec(S_DBLCLK, lngX) <> vbNullString Then
            ![sec_ondblclick] = arr_varSec(S_DBLCLK, lngX)
          End If
          If arr_varSec(S_MDWN, lngX) <> vbNullString Then
            ![sec_onmousedown] = arr_varSec(S_MDWN, lngX)
          End If
          If arr_varSec(S_MMOV, lngX) <> vbNullString Then
            ![sec_onmousemove] = arr_varSec(S_MMOV, lngX)
          End If
          If arr_varSec(S_MUP, lngX) <> vbNullString Then
            ![sec_onmouseup] = arr_varSec(S_MUP, lngX)
          End If
          ![sec_specialeffect] = arr_varSec(S_SPCEF, lngX)
          If arr_varSec(S_TAG, lngX) <> vbNullString Then
            ![sec_tag] = arr_varSec(S_TAG, lngX)
          End If
          ![sec_visible] = arr_varSec(S_VIS, lngX)
          ![sec_datemodified] = Now()
          .Update
        Else
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
              .Update
            End If
          End If
          If ![sec_backcolor] <> arr_varSec(S_BKCLR, lngX) Then
            .Edit
            ![sec_backcolor] = arr_varSec(S_BKCLR, lngX)
            ![sec_datemodified] = Now()
            .Update
          End If
          If arr_varSec(S_GROW, lngX) <> vbNullString Then
            If IsNull(![sec_cangrow]) = True Then
              .Edit
              ![sec_cangrow] = arr_varSec(S_GROW, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_cangrow] <> arr_varSec(S_GROW, lngX) Then
                .Edit
                ![sec_cangrow] = arr_varSec(S_GROW, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![sec_cangrow]) = False Then
              .Edit
              ![sec_cangrow] = Null
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If arr_varSec(S_SHRINK, lngX) <> vbNullString Then
            If IsNull(![sec_canshrink]) = True Then
              .Edit
              ![sec_canshrink] = arr_varSec(S_SHRINK, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_canshrink] <> arr_varSec(S_SHRINK, lngX) Then
                .Edit
                ![sec_canshrink] = arr_varSec(S_SHRINK, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![sec_canshrink]) = False Then
              .Edit
              ![sec_canshrink] = Null
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If IsNull(![sec_displaywhen]) = True Then
            .Edit
            ![sec_displaywhen] = arr_varSec(S_DISP, lngX)
            ![sec_datemodified] = Now()
            .Update
          Else
            If ![sec_displaywhen] <> arr_varSec(S_DISP, lngX) Then
              .Edit
              ![sec_displaywhen] = arr_varSec(S_DISP, lngX)
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If IsNull(![sec_forcenewpage]) = True Then
            .Edit
            ![sec_forcenewpage] = arr_varSec(S_NEWPG, lngX)
            ![sec_datemodified] = Now()
            .Update
          Else
            If ![sec_forcenewpage] <> arr_varSec(S_NEWPG, lngX) Then
              .Edit
              ![sec_forcenewpage] = arr_varSec(S_NEWPG, lngX)
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If ![sec_height] <> arr_varSec(S_HT, lngX) Then
            .Edit
            ![sec_height] = arr_varSec(S_HT, lngX)
            ![sec_datemodified] = Now()
            .Update
          End If
          If ![sec_keeptogether] <> arr_varSec(S_KEEP, lngX) Then
            .Edit
            ![sec_keeptogether] = arr_varSec(S_KEEP, lngX)
            ![sec_datemodified] = Now()
            .Update
          End If
          If IsNull(![sec_newroworcol]) = True Then
            .Edit
            ![sec_newroworcol] = arr_varSec(S_NEWROW, lngX)
            ![sec_datemodified] = Now()
            .Update
          Else
            If ![sec_newroworcol] <> arr_varSec(S_NEWROW, lngX) Then
              .Edit
              ![sec_newroworcol] = arr_varSec(S_NEWROW, lngX)
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If arr_varSec(S_CLK, lngX) <> vbNullString Then
            If IsNull(![sec_onclick]) = True Then
              .Edit
              ![sec_onclick] = arr_varSec(S_CLK, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_onclick] <> arr_varSec(S_CLK, lngX) Then
                .Edit
                ![sec_onclick] = arr_varSec(S_CLK, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![sec_onclick]) = False Then
              .Edit
              ![sec_onclick] = Null
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If arr_varSec(S_DBLCLK, lngX) <> vbNullString Then
            If IsNull(![sec_ondblclick]) = True Then
              .Edit
              ![sec_ondblclick] = arr_varSec(S_DBLCLK, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_ondblclick] <> arr_varSec(S_DBLCLK, lngX) Then
                .Edit
                ![sec_ondblclick] = arr_varSec(S_DBLCLK, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![sec_ondblclick]) = False Then
              .Edit
              ![sec_ondblclick] = Null
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If arr_varSec(S_MDWN, lngX) <> vbNullString Then
            If IsNull(![sec_onmousedown]) = True Then
              .Edit
              ![sec_onmousedown] = arr_varSec(S_MDWN, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_onmousedown] <> arr_varSec(S_MDWN, lngX) Then
                .Edit
                ![sec_onmousedown] = arr_varSec(S_MDWN, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![sec_onmousedown]) = False Then
              .Edit
              ![sec_onmousedown] = Null
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If arr_varSec(S_MMOV, lngX) <> vbNullString Then
            If IsNull(![sec_onmousemove]) = True Then
              .Edit
              ![sec_onmousemove] = arr_varSec(S_MMOV, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_onmousemove] <> arr_varSec(S_MMOV, lngX) Then
                .Edit
                ![sec_onmousemove] = arr_varSec(S_MMOV, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![sec_onmousemove]) = False Then
              .Edit
              ![sec_onmousemove] = Null
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If arr_varSec(S_MUP, lngX) <> vbNullString Then
            If IsNull(![sec_onmouseup]) = True Then
              .Edit
              ![sec_onmouseup] = arr_varSec(S_MUP, lngX)
              ![sec_datemodified] = Now()
              .Update
            Else
              If ![sec_onmouseup] <> arr_varSec(S_MUP, lngX) Then
                .Edit
                ![sec_onmouseup] = arr_varSec(S_MUP, lngX)
                ![sec_datemodified] = Now()
                .Update
              End If
            End If
          Else
            If IsNull(![sec_onmouseup]) = False Then
              .Edit
              ![sec_onmouseup] = Null
              ![sec_datemodified] = Now()
              .Update
            End If
          End If
          If ![sec_specialeffect] <> arr_varSec(S_SPCEF, lngX) Then
            .Edit
            ![sec_specialeffect] = arr_varSec(S_SPCEF, lngX)
            ![sec_datemodified] = Now()
            .Update
          End If
          If arr_varSec(S_TAG, lngX) <> vbNullString Then
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
          If ![sec_visible] <> arr_varSec(S_VIS, lngX) Then
            .Edit
            ![sec_visible] = arr_varSec(S_VIS, lngX)
            ![sec_datemodified] = Now()
            .Update
          End If
        End If
      Next

      lngDels = 0&
      ReDim arr_varDel(D_ELEMS, 0)

      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          blnFound = False
          For lngY = 0& To (lngSecs - 1&)
            If arr_varSec(S_DID, lngY) = ![dbs_id] And arr_varSec(S_FID, lngY) = ![frm_id] And arr_varSec(S_SNAM, lngY) = ![sec_name] Then
              blnFound = True
              Exit For
            End If
          Next
          If blnFound = False Then
            lngDels = lngDels + 1&
            lngE = lngDels - 1&
            ReDim Preserve arr_varDel(D_ELEMS, lngE)
            arr_varDel(D_DID, lngE) = ![dbs_id]
            arr_varDel(D_FID, lngE) = ![frm_id]
            arr_varDel(D_FNAM, lngE) = DLookup("[frm_name]", "tblForm", "[frm_id] = " & CStr(![frm_id]))
            arr_varDel(D_SID, lngE) = ![sec_id]
            arr_varDel(D_SNAM, lngE) = ![sec_name]
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next

      For lngX = 0& To (lngDels - 1&)
        blnDelete = True
        Debug.Print "'DEL SEC? " & arr_varDel(D_SNAM, lngX) & " on " & arr_varDel(D_FNAM, lngX)
Stop
        If blnDelete = True Then
          ' ** Delete tblForm_Section, by specified [secid].
          Set qdf = dbs.QueryDefs("zz_qry_Form_Section_01")
          With qdf.Parameters
            ![secid] = arr_varDel(D_SID, lngX)
          End With
          qdf.Execute dbFailOnError
        End If
      Next

      .Close
    End With

    .Close
  End With

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Sec_Doc = blnRetValx

End Function

Private Function Frm_Ctl_Doc() As Boolean
' ** Document all form controls in Trust Accountant to tblForm_Control.
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_Ctl_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Form, ctl1 As Control, ctl2 As Control, prp As Object
  Dim lngFrms As Long, arr_varFrm() As Variant
  Dim lngCtls As Long, arr_varCtl() As Variant
  Dim blnUpdate As Boolean, blnDelete As Boolean
  Dim strFrm As String, strParent As String
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim intPos1 As Integer
  Dim varTmp00 As Variant
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varFrm().
  Const F_ELEMS As Integer = 14  ' ** Array's first-element UBound().
  Const F_DID     As Integer = 0
  Const F_FID     As Integer = 1
  Const F_OTYP    As Integer = 2
  Const F_FNAM    As Integer = 3
  Const F_CTLS    As Integer = 4
  Const F_CAP     As Integer = 5
  Const F_HAS_SUB As Integer = 6
  Const F_IS_SUB  As Integer = 7
  Const F_ACT     As Integer = 8
  Const F_HID     As Integer = 9
  Const F_DSC     As Integer = 10
  Const F_TAG     As Integer = 11
  Const F_PARSUB  As Integer = 12
  Const F_SUBS    As Integer = 13
  Const F_DAT     As Integer = 14

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 14  ' ** Array's first-element UBound().
  Const C_DID   As Integer = 0
  Const C_FID   As Integer = 1
  Const C_CID   As Integer = 2
  Const C_OTYP  As Integer = 3
  Const C_CNAM  As Integer = 4
  Const C_CTYP  As Integer = 5
  Const C_CAP   As Integer = 6
  Const C_SEC   As Integer = 7
  Const C_SRC   As Integer = 8
  Const C_OBJ   As Integer = 9
  Const C_CTLS  As Integer = 10
  Const C_PARID As Integer = 11
  Const C_PAR   As Integer = 12
  Const C_DAT   As Integer = 13
  Const C_FND   As Integer = 14

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    lngFrms = 0&
    ReDim arr_varFrm(F_ELEMS, 0)

    ' ** Get a list of all forms.
    Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngFrms = lngFrms + 1&
          lngE = lngFrms - 1&
          ReDim Preserve arr_varFrm(F_ELEMS, lngE)
          ' ******************************************************
          ' ** Array: arr_varFrm()
          ' **
          ' **   Field  Element  Name                Constant
          ' **   =====  =======  ==================  ===========
          ' **     1       0     dbs_id              F_DID
          ' **     2       1     frm_id              F_FID
          ' **     3       2     objtype_type        F_OTYP
          ' **     4       3     frm_name            F_FNAM
          ' **     5       4     frm_controls        F_CTLS
          ' **     6       5     frm_caption         F_CAP
          ' **     7       6     frm_hassub          F_HAS_SUB
          ' **     8       7     frm_issub           F_IS_SUB
          ' **     9       8     frm_active          F_ACT
          ' **    10       9     sec_hidden          F_HID
          ' **    11      10     frm_description     F_DSC
          ' **    12      11     frm_tag             F_TAG
          ' **    13      12     frm_parent_sub      F_PARSUB
          ' **    14      13     frm_subs            F_SUBS
          ' **    15      14     frm_datemodified    F_DAT
          ' **
          ' ******************************************************
          arr_varFrm(F_DID, lngE) = ![dbs_id]
          arr_varFrm(F_FID, lngE) = ![frm_id]
          arr_varFrm(F_OTYP, lngE) = ![objtype_type]
          arr_varFrm(F_FNAM, lngE) = ![frm_name]
          arr_varFrm(F_CTLS, lngE) = ![frm_controls]
          arr_varFrm(F_CAP, lngE) = ![frm_caption]
          arr_varFrm(F_HAS_SUB, lngE) = ![frm_hassub]
          arr_varFrm(F_IS_SUB, lngE) = ![frm_issub]
          arr_varFrm(F_ACT, lngE) = ![frm_active]
          arr_varFrm(F_HID, lngE) = ![sec_hidden]
          arr_varFrm(F_DSC, lngE) = ![frm_description]
          arr_varFrm(F_TAG, lngE) = ![frm_tag]
          arr_varFrm(F_PARSUB, lngE) = ![frm_parent_sub]
          arr_varFrm(F_SUBS, lngE) = ![frm_subs]
          arr_varFrm(F_DAT, lngE) = ![frm_datemodified]
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With
    Set rst = Nothing

    lngCtls = 0&
    ReDim arr_varCtl(C_ELEMS, 0)

    ' ** Get a list of all controls currently documented.
    Set rst = .OpenRecordset("tblForm_Control", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngCtls = lngCtls + 1&
          lngE = lngCtls - 1&
          ReDim Preserve arr_varCtl(C_ELEMS, lngE)
          ' *****************************************************
          ' ** Array: arr_varCtl()
          ' **
          ' **   Field  Element  Name                Constant
          ' **   =====  =======  ==================  ==========
          ' **     1       0     dbs_id              C_DID
          ' **     2       1     frm_id              C_FID
          ' **     3       2     ctl_id              C_CID
          ' **     4       3     objtype_type        C_OTYP
          ' **     5       4     ctl_name            C_CNAM
          ' **     6       5     ctltype_type        C_CTYP
          ' **     7       6     ctl_caption         C_CAP
          ' **     8       7     ctl_section         C_SEC
          ' **     9       8     ctl_source          C_SRC
          ' **    10       9     ctl_sourceobject    C_OBJ
          ' **    11      10     ctl_controls        C_CTLS
          ' **    12      11     ctl_id_parent       C_PARID
          ' **    13      12     ctl_parent          C_PAR
          ' **    14      13     ctl_datemodified    C_DAT
          ' **    15      14     Found               C_FND
          ' **
          ' *****************************************************
          arr_varCtl(C_DID, lngE) = ![dbs_id]
          arr_varCtl(C_FID, lngE) = ![frm_id]
          arr_varCtl(C_CID, lngE) = ![ctl_id]
          arr_varCtl(C_OTYP, lngE) = ![objtype_type]
          arr_varCtl(C_CNAM, lngE) = ![ctl_name]
          arr_varCtl(C_CTYP, lngE) = ![ctltype_type]
          arr_varCtl(C_CAP, lngE) = ![ctl_caption]
          arr_varCtl(C_SEC, lngE) = ![ctl_section]
          arr_varCtl(C_SRC, lngE) = ![ctl_source]
          arr_varCtl(C_OBJ, lngE) = ![ctl_sourceobject]
          arr_varCtl(C_CTLS, lngE) = ![ctl_controls]
          arr_varCtl(C_PARID, lngE) = ![ctl_id_parent]
          arr_varCtl(C_PAR, lngE) = ![ctl_parent]
          arr_varCtl(C_DAT, lngE) = ![ctl_datemodified]
          arr_varCtl(C_FND, lngE) = CBool(False)
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    Set rst = .OpenRecordset("tblForm_Control", dbOpenDynaset, dbConsistent)

  End With

  For lngX = 0& To (lngFrms - 1&)
    DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
    Set frm = Forms(arr_varFrm(F_FNAM, lngX))
    With frm
      For Each ctl1 In .Controls
        With ctl1
          varTmp00 = Empty
'ctl_id
          rst.FindFirst "[dbs_id] = " & CStr(arr_varFrm(F_DID, lngX)) & " And [ctl_name] = '" & .Name & "' And " & _
            "[frm_id] = " & CStr(arr_varFrm(F_FID, lngX))
          If rst.NoMatch = True Then
            blnUpdate = True
            rst.AddNew
'dbs_id
            rst![dbs_id] = arr_varFrm(F_DID, lngX)
'frm_id
            rst![frm_id] = arr_varFrm(F_FID, lngX)
'ctl_name
            rst![ctl_name] = .Name
'objtype_type
            rst![objtype_type] = acForm
          Else
            blnUpdate = False
            rst.Edit
            For lngY = 0& To (lngCtls - 1&)
              If arr_varCtl(C_DID, lngY) = rst![dbs_id] And arr_varCtl(C_FID, lngY) = rst![frm_id] And _
                  arr_varCtl(C_CNAM, lngY) = rst![ctl_name] Then
                arr_varCtl(C_FND, lngY) = CBool(True)
                Exit For
              End If
            Next
            If TxtCaseComp(rst![ctl_name], .Name) = False Then  ' ** Function: Below.
              rst![ctl_name] = .Name
            End If
          End If
On Error Resume Next
          strParent = .Parent.Name
On Error GoTo 0
'ctl_parent
          If strParent <> vbNullString Then
            If IsNull(rst![ctl_parent]) = True Then
              blnUpdate = True
              rst![ctl_parent] = strParent
            Else
              If rst![ctl_parent] <> strParent Then
                blnUpdate = True
                rst![ctl_parent] = strParent
              End If
            End If
'ctl_id_parent
            If strParent <> frm.Name Then
              varTmp00 = DLookup("[ctl_id]", "tblForm_Control", "[dbs_id] = " & CStr(arr_varFrm(F_DID, lngX)) & " And " & _
                "[frm_id] = " & CStr(arr_varFrm(F_FID, lngX)) & " And " & _
                "[ctl_name] = '" & strParent & "'")
              If IsNull(varTmp00) = False Then
                rst![ctl_id_parent] = CLng(varTmp00)
              End If
            End If
          Else
            If IsNull(rst![ctl_parent]) = False Then
              blnUpdate = True
              rst![ctl_parent] = Null
            End If
          End If
          varTmp00 = Empty
'ctltype_type
          If blnUpdate = True Then
            rst![ctltype_type] = .ControlType
          Else
            If rst![ctltype_type] <> .ControlType Then
              blnUpdate = True
              rst![ctltype_type] = .ControlType
            End If
          End If
'ctl_controls
On Error Resume Next
          varTmp00 = .Controls.Count
On Error GoTo 0
          If IsEmpty(varTmp00) = False Then
            If IsNull(varTmp00) = False Then
              If IsNumeric(varTmp00) = True Then
                rst![ctl_controls] = varTmp00
              Else
                rst![ctl_controls] = Null
              End If
            Else
              rst![ctl_controls] = Null
            End If
          Else
            rst![ctl_controls] = Null
          End If
          For Each prp In .Properties
            With prp
              Select Case .Name
              Case "Caption"
'ctl_caption
                If .Value <> vbNullString Then
                  rst![ctl_caption] = .Value
                Else
                  If IsNull(rst![ctl_caption]) = False Then
                    rst![ctl_caption] = Null
                  End If
                End If
              Case "ControlSource"
'ctl_source
                If .Value <> vbNullString Then
                  rst![ctl_source] = .Value
                Else
                  If IsNull(rst![ctl_source]) = False Then
                    rst![ctl_source] = Null
                  End If
                End If
              Case "SourceObject"
'ctl_sourceobject
                If .Value <> vbNullString Then
                  rst![ctl_sourceobject] = .Value
                Else
                  If IsNull(rst![ctl_sourceobject]) = False Then
                    rst![ctl_sourceobject] = Null
                  End If
                End If
              Case "Section"
'ctl_section
                rst![ctl_section] = .Value
              End Select
            End With
          Next
'ctl_datemodified
          rst![ctl_datemodified] = Now()
          rst.Update
        End With  ' ** This Control: ctl1.
      Next  ' ** For each Control: ctl1.
    End With  ' ** This Form: frm.
    DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX), acSaveNo
  Next  ' ** For each Form: frm.

  For lngX = 0& To (lngCtls - 1&)
    If arr_varCtl(C_FND, lngX) = False And arr_varCtl(C_DID, lngX) = lngThisDbsID Then
      strFrm = vbNullString
      For lngY = 0& To (lngFrms - 1&)
        If arr_varFrm(F_DID, lngY) = arr_varCtl(C_DID, lngX) And arr_varFrm(F_FID, lngY) = arr_varCtl(C_FID, lngX) Then
          strFrm = arr_varFrm(F_FNAM, lngY)
          Exit For
        End If
      Next
      blnDelete = True
        Debug.Print "'DEL CTL? '" & arr_varCtl(C_CNAM, lngX) & "' on " & strFrm
Stop
      If blnDelete = True Then
        ' ** Delete tblForm_Control, by specified [ctlid].
        Set qdf = dbs.QueryDefs("zz_qry_Form_Control_01")
        With qdf.Parameters
          ![ctlid] = arr_varCtl(C_CID, lngX)
        End With
        qdf.Execute dbFailOnError
      End If
    End If
  Next

  ' ***************************************************************************************
  ' ** ScrollBarAlign Property: You can use the ScrollBarAlign to specify or
  ' **                          determine the alignment of a vertical scroll bar.
  ' ** Setting: The ScrollBarAlign property uses the following settings.
  ' **
  ' **     VB  Setting  Description
  ' **     ==  =======  =================================================================
  ' **     0   System   Vertical scroll bar is placed on the left if the form or report
  ' **                  Orientation property is right to left; and on the right if the
  ' **                  form or report Orientation property is left to right.
  ' **     1   Right    Aligns vertical scroll bar on the right side of the control.
  ' **     2   Left     Aligns vertical scroll bar on the left side of the control.
  ' **
  ' ***************************************************************************************

  ' ** Update tblForm_Shortcut, for fs_caption, from tblForm_Control.
  Set qdf = dbs.QueryDefs("zz_qry_Form_Control_02")
  qdf.Execute

  dbs.Close

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set ctl1 = Nothing
  Set ctl2 = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Ctl_Doc = blnRetValx

End Function

Private Function Frm_RecSrc_Doc() As Boolean
' ** Document all form Record Sources to tblForm_RecordSource.
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_RecSrc_Doc"

  Dim dbs As DAO.Database, qdf2 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, rst3 As DAO.Recordset, frm As Form
  Dim lngFrms As Long, arr_varFrm() As Variant
  Dim strRecSrc As String, lngQryTblTypeType As Long
  Dim blnIsTbl As Boolean, blnIsQry As Boolean, blnIsSQL As Boolean
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim strFormRef As String
  Dim lngLen As Long
  Dim intPos1 As Integer
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varFrm().
  Const F_ELEMS As Integer = 14  ' ** Array's first-element UBound().
  Const F_DID     As Integer = 0
  Const F_FID     As Integer = 1
  Const F_OTYP    As Integer = 2
  Const F_FNAM    As Integer = 3
  Const F_CTLS    As Integer = 4
  Const F_CAP     As Integer = 5
  Const F_HAS_SUB As Integer = 6
  Const F_IS_SUB  As Integer = 7
  Const F_ACT     As Integer = 8
  Const F_HID     As Integer = 9
  Const F_DSC     As Integer = 10
  Const F_TAG     As Integer = 11
  Const F_PARSUB  As Integer = 12
  Const F_SUBS    As Integer = 13
  Const F_DAT     As Integer = 14

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    Set rst1 = .OpenRecordset("tblForm_RecordSource", dbOpenDynaset, dbConsistent)

    lngFrms = 0&
    ReDim arr_varFrm(F_ELEMS, 0)

    Set rst2 = .OpenRecordset("tblForm", dbOpenDynaset, dbConsistent)
    With rst2
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngFrms = lngFrms + 1&
          lngE = lngFrms - 1&
          ReDim Preserve arr_varFrm(F_ELEMS, lngE)
          ' ******************************************************
          ' ** Array: arr_varFrm()
          ' **
          ' **   Field  Element  Name                Constant
          ' **   =====  =======  ==================  ===========
          ' **     1       0     dbs_id              F_DID
          ' **     2       1     frm_id              F_FID
          ' **     3       2     objtype_type        F_OTYP
          ' **     4       3     frm_name            F_FNAM
          ' **     5       4     frm_controls        F_CTLS
          ' **     6       5     frm_caption         F_CAP
          ' **     7       6     frm_hassub          F_HAS_SUB
          ' **     8       7     frm_issub           F_IS_SUB
          ' **     9       8     frm_active          F_ACT
          ' **    10       9     sec_hidden          F_HID
          ' **    11      10     frm_description     F_DSC
          ' **    12      11     frm_tag             F_TAG
          ' **    13      12     frm_parent_sub      F_PARSUB
          ' **    14      13     frm_subs            F_SUBS
          ' **    15      14     frm_datemodified    F_DAT
          ' **
          ' ******************************************************
          arr_varFrm(F_DID, lngE) = ![dbs_id]
          arr_varFrm(F_FID, lngE) = ![frm_id]
          arr_varFrm(F_OTYP, lngE) = ![objtype_type]
          arr_varFrm(F_FNAM, lngE) = ![frm_name]
          arr_varFrm(F_CTLS, lngE) = ![frm_controls]
          arr_varFrm(F_CAP, lngE) = ![frm_caption]
          arr_varFrm(F_HAS_SUB, lngE) = ![frm_hassub]
          arr_varFrm(F_IS_SUB, lngE) = ![frm_issub]
          arr_varFrm(F_ACT, lngE) = ![frm_active]
          arr_varFrm(F_HID, lngE) = ![sec_hidden]
          arr_varFrm(F_DSC, lngE) = ![frm_description]
          arr_varFrm(F_TAG, lngE) = ![frm_tag]
          arr_varFrm(F_PARSUB, lngE) = ![frm_parent_sub]
          arr_varFrm(F_SUBS, lngE) = ![frm_subs]
          arr_varFrm(F_DAT, lngE) = ![frm_datemodified]
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    lngLen = 0&
    For lngX = 0& To (lngFrms - 1&)

      lngQryTblTypeType = 0&: strRecSrc = vbNullString

On Error Resume Next
      DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
      If ERR.Number <> 0 Then
        Select Case ERR.Number
        Case 2001  ' ** You Canceled the previous operation.
On Error GoTo 0
          Debug.Print "'FRMX: " & arr_varFrm(F_FNAM, lngX)
        Case Else
Stop
On Error GoTo 0
        End Select
      Else
On Error GoTo 0

        Set frm = Forms(arr_varFrm(F_FNAM, lngX))
        With frm
          If IsNull(.RecordSource) = False Then
            If .RecordSource <> vbNullString Then
              strRecSrc = .RecordSource
            End If
          End If
        End With
        DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX), acSaveNo
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
          lngQryTblTypeType = acNothing
        End If

        With rst1
          .FindFirst "[dbs_id] = " & CStr(arr_varFrm(F_DID, lngX)) & " And [frm_id] = " & CStr(arr_varFrm(F_FID, lngX))
          If .NoMatch = True Then
            .AddNew
            ![dbs_id] = arr_varFrm(F_DID, lngX)
            ![frm_id] = arr_varFrm(F_FID, lngX)
            ![recsrc_recordsource] = NullIfNullStr(strRecSrc)  ' ** Module Function: modStringFuncs.
            intPos1 = InStr(strRecSrc, "Forms")
            If intPos1 > 0 Then
              If intPos1 > 1 Then
                If Mid$(strRecSrc, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
              End If
              strFormRef = Mid$(strRecSrc, intPos1)
              If InStr(strFormRef, " ") > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
              strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
              ![recsrc_hasformref] = True
              ![recsrc_formref] = strFormRef
            Else
              intPos1 = InStr(strRecSrc, "Reports")
              If intPos1 > 0 Then
                If intPos1 > 1 Then
                  If Mid$(strRecSrc, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
                End If
                strFormRef = Mid$(strRecSrc, intPos1)
                If InStr(strFormRef, " ") > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
                strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
                ![recsrc_hasformref] = True
                ![recsrc_formref] = strFormRef
                ' ** Now Check tblFormRef? Unlikely, so haven't created queries or code to process.
              End If
            End If
            ![qrytbltype_type] = lngQryTblTypeType
            ![recsrc_datemodified] = Now()
            .Update
            .Bookmark = .LastModified
          Else
            .Edit
            ![recsrc_recordsource] = NullIfNullStr(strRecSrc)
            intPos1 = InStr(strRecSrc, "Forms")
            If Mid$(strRecSrc, (intPos1 + 5), 1) <> "]" And Mid$(strRecSrc, (intPos1 + 5), 1) <> "." And _
                Mid$(strRecSrc, (intPos1 + 5), 1) <> "!" Then
              ' ** The word 'forms' is part of its name, and not a form reference!
              intPos1 = InStr((intPos1 + 1), strRecSrc, "Forms")
            End If
            If intPos1 > 0 Then
              If intPos1 > 1 Then
                If Mid$(strRecSrc, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
              End If
              strFormRef = Mid$(strRecSrc, intPos1)
              If InStr(strFormRef, " ") > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
              strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
              ![recsrc_hasformref] = True
              ![recsrc_formref] = strFormRef
            Else
              intPos1 = InStr(strRecSrc, "Reports")
              If Mid$(strRecSrc, (intPos1 + 7), 1) <> "]" And Mid$(strRecSrc, (intPos1 + 7), 1) <> "." And _
                  Mid$(strRecSrc, (intPos1 + 7), 1) <> "!" Then
                ' ** The word 'forms' is part of its name, and not a report reference!
                intPos1 = InStr((intPos1 + 1), strRecSrc, "Reports")
              End If
              If intPos1 > 0 Then
                If intPos1 > 1 Then
                  If Mid$(strRecSrc, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
                End If
                strFormRef = Mid$(strRecSrc, intPos1)
                If InStr(strFormRef, " ") > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
                strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
                ![recsrc_hasformref] = True
                ![recsrc_formref] = strFormRef
                ' ** Now Check tblFormRef? Unlikely, so haven't created queries or code to process.
              Else
                ![recsrc_hasformref] = False
                ![recsrc_formref] = Null
              End If
            End If
            ![qrytbltype_type] = lngQryTblTypeType
            ![recsrc_datemodified] = Now()
            .Update
          End If
        End With

      End If

      If Len(strRecSrc) > lngLen Then lngLen = Len(strRecSrc)
    Next

    rst1.Close

    .Close
  End With

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set frm = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set rst3 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  Frm_RecSrc_Doc = blnRetValx

End Function

Private Function Frm_Shortcut_Doc() As Boolean
' ** Document all form Shortcuts to tblForm_Shortcut.
' ** This includes both standard Alt-keys, via ampersand
' ** in label, and explicitly defined Form_KeyDown() keys.
' ** WARNING: This function requires that both Frm_Ctl_Doc(), above,
' ** and VBA_Component_Proc_Doc(), in zz_mod_ModuleFormatFuncs,
' ** be run first, so that the procedure start lines, current captions,
' ** and code are accurate!
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_Shortcut_Doc"

  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset, frm As Form, ctl As Control, prp As Object
  Dim lngFrms As Long, arr_varFrm() As Variant
  Dim lngCtls As Long, arr_varCtl() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngProcs As Long, arr_varProc As Variant
  Dim lngTypes As Long, arr_varType() As Variant
  Dim lngKeys As Long, arr_varKey() As Variant
  Dim lngCtlCnt As Long, lngCtlID As Long, lngTypeType As Long
  Dim strForm As String, strCap As String, strCtl As String
  Dim strLine As String, strLine2 As String, strModName As String
  Dim lngLines As Long, lngLen As Long, lngProcEnd As Long
  Dim lngPlain As Long, lngShift As Long, lngAlt As Long, lngCtrl As Long
  Dim lngCtrlShift As Long, lngAltShift As Long, lngCtrlAlt As Long, lngCtrlAltShift As Long, lngUnknown As Long
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim blnFound As Boolean, blnDelete As Boolean, blnAdd As Boolean, blnEdit As Boolean, blnPlainNotFound As Boolean, blnSkip As Boolean
  Dim intPos1 As Integer, intPos2 As Integer
  Dim varTmp00 As Variant, strTmp01 As String, lngTmp02 As Long, arr_varTmp03 As Variant, arr_varTmp04 As Variant, arr_varTmp05 As Variant
  Dim lngV As Long, lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long

  ' ** Array: arr_varFrm().
  Const FM_ELEMS As Integer = 15  ' ** Array's first-element UBound().
  Const F_DID     As Integer = 0
  Const F_FID     As Integer = 1
  Const F_OTYP    As Integer = 2
  Const F_FNAM    As Integer = 3
  Const F_CTLS    As Integer = 4
  Const F_CAP     As Integer = 5
  Const F_HAS_SUB As Integer = 6
  Const F_IS_SUB  As Integer = 7
  Const F_ACT     As Integer = 8
  Const F_HID     As Integer = 9
  Const F_DSC     As Integer = 10
  Const F_TAG     As Integer = 11
  Const F_PARSUB  As Integer = 12
  Const F_SUBS    As Integer = 13
  Const F_DAT     As Integer = 14
  Const F_ARR     As Integer = 15

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const C_DID  As Integer = 0
  Const C_FID  As Integer = 1
  Const C_CID  As Integer = 2
  Const C_CNAM As Integer = 3
  Const C_ORD  As Integer = 4
  Const C_CTYP As Integer = 5
  Const C_CAP  As Integer = 6
  Const C_LTR  As Integer = 7
  Const C_UN   As Integer = 8
  Const C_PAR  As Integer = 9
  Const C_DEL  As Integer = 10

  ' ** Array: arr_varDel().
  Const DEL_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const DEL_DID  As Integer = 0
  Const DEL_FID  As Integer = 1
  Const DEL_FAM  As Integer = 2
  Const DEL_SID  As Integer = 3
  Const DEL_SNAM As Integer = 4

  ' ** Array: arr_varProc().
  Const P_DID  As Integer = 0
  Const P_DNAM As Integer = 1
  Const P_CID  As Integer = 2
  Const P_CNAM As Integer = 3
  Const P_FID  As Integer = 4
  Const P_FNAM As Integer = 5
  Const P_PID  As Integer = 6
  Const P_PNAM As Integer = 7
  Const P_LIN  As Integer = 8
  Const P_ARR  As Integer = 9

  ' ** Array: arr_varType().
  Const T_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const T_TYP_NAM As Integer = 0
  Const T_TYP_TYP As Integer = 1
  Const T_SEC_BEG As Integer = 2
  Const T_SEC_END As Integer = 3
  Const T_SEC_IND As Integer = 4
  Const T_KEYS    As Integer = 5
  Const T_ARR     As Integer = 6

  ' ** Array: arr_varKey().
  Const K_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const K_CONST   As Integer = 0
  Const K_LINENUM As Integer = 1
  Const K_KEYVAL  As Integer = 2

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    lngFrms = 0&
    ReDim arr_varFrm(FM_ELEMS, 0)

    Set rst1 = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
    With rst1
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID And Left$(![frm_name], 3) <> "zz_" Then
          lngFrms = lngFrms + 1&
          lngE = lngFrms - 1&
          ReDim Preserve arr_varFrm(FM_ELEMS, lngE)
          ' ********************************************************
          ' ** Array: arr_varFrm()
          ' **
          ' **   Field  Element  Name                  Constant
          ' **   =====  =======  ====================  ===========
          ' **     1       0     dbs_id                F_DID
          ' **     2       1     frm_id                F_FID
          ' **     3       2     objtype_type          F_OTYP
          ' **     4       3     frm_name              F_FNAM
          ' **     5       4     frm_controls          F_CTLS
          ' **     6       5     frm_caption           F_CAP
          ' **     7       6     frm_hassub            F_HAS_SUB
          ' **     8       7     frm_issub             F_IS_SUB
          ' **     9       8     frm_active            F_ACT
          ' **    10       9     sec_hidden            F_HID
          ' **    11      10     frm_description       F_DSC
          ' **    12      11     frm_tag               F_TAG
          ' **    13      12     frm_parent_sub        F_PARSUB
          ' **    14      13     frm_subs              F_SUBS
          ' **    15      14     frm_datemodified      F_DAT
          ' **    16      15     arr_varCtl() Array    F_ARR
          ' **
          ' ********************************************************
          arr_varFrm(F_DID, lngE) = ![dbs_id]
          arr_varFrm(F_FID, lngE) = ![frm_id]
          arr_varFrm(F_OTYP, lngE) = ![objtype_type]
          arr_varFrm(F_FNAM, lngE) = ![frm_name]
          arr_varFrm(F_CTLS, lngE) = ![frm_controls]
          arr_varFrm(F_CAP, lngE) = ![frm_caption]
          arr_varFrm(F_HAS_SUB, lngE) = ![frm_hassub]
          arr_varFrm(F_IS_SUB, lngE) = ![frm_issub]
          arr_varFrm(F_ACT, lngE) = ![frm_active]
          arr_varFrm(F_HID, lngE) = ![sec_hidden]
          arr_varFrm(F_DSC, lngE) = ![frm_description]
          arr_varFrm(F_TAG, lngE) = ![frm_tag]
          arr_varFrm(F_PARSUB, lngE) = ![frm_parent_sub]
          arr_varFrm(F_SUBS, lngE) = ![frm_subs]
          arr_varFrm(F_DAT, lngE) = ![frm_datemodified]
          arr_varFrm(F_ARR, lngE) = Empty
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    lngDels = 0&
    ReDim arr_varDel(DEL_ELEMS, 0)

blnSkip = False
If blnSkip = False Then

    Set rst1 = .OpenRecordset("tblForm_Shortcut", dbOpenDynaset, dbConsistent)

    For lngX = 0& To (lngFrms - 1&)

      strForm = arr_varFrm(F_FNAM, lngX)
      DoCmd.OpenForm strForm, acDesign, , , , acHidden
      Set frm = Forms(strForm)
      With frm

        lngCtls = 0&
        ReDim arr_varCtl(C_ELEMS, 0)

        ' ** Update tblForm with number of controls.
        lngCtlCnt = .Controls.Count
        Set qdf = dbs.QueryDefs("zz_qry_Form_Shortcut_02")
        With qdf.Parameters
          ![frmid] = arr_varFrm(F_FID, lngX)
          ![ctls] = lngCtlCnt
        End With
        qdf.Execute

        ' ** Now look for shortcut keys by label; these are always Alt-keys.
        For Each ctl In .Controls
          With ctl
            lngCtlID = 0&
            For Each prp In .Properties
              With prp
                If .Name = "Caption" Then
                  If IsNull(.Value) = False Then
                    strCap = .Value
                    intPos1 = InStr(strCap, "&")
                    Do While intPos1 > 0
                      If Mid$(strCap, (intPos1 + 1), 1) <> "&" Then  ' ** Double means real ampersand, not shortcut key.
                        lngCtls = lngCtls + 1&
                        lngE = lngCtls - 1&
                        ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                        ' **************************************************
                        ' ** Array: arr_varCtl()
                        ' **
                        ' **   Field  Element  Name             Constant
                        ' **   =====  =======  ===============  ==========
                        ' **     1       0     dbs_id           C_DID
                        ' **     2       1     frm_id           C_FID
                        ' **     3       2     ctl_id           C_CID
                        ' **     4       3     ctl_name         C_CNAM
                        ' **     5       4     fs_order         C_ORD
                        ' **     6       5     ctltype_type     C_CTYP
                        ' **     7       6     fs_caption       C_CAP
                        ' **     8       7     fs_letter        C_LTR
                        ' **     9       8     fs_unattached    C_UN
                        ' **    10       9     fs_parent        C_PAR
                        ' **    11      10     Delete           C_DEL
                        ' **
                        ' **************************************************
                        arr_varCtl(C_DID, lngE) = arr_varFrm(F_DID, lngX)
                        arr_varCtl(C_FID, lngE) = arr_varFrm(F_FID, lngX)
                        arr_varCtl(C_CNAM, lngE) = ctl.Name
                        lngCtlID = DLookup("[ctl_id]", "tblForm_Control", "[frm_id] = " & CStr(arr_varFrm(F_FID, lngX)) & _
                          " And [ctl_name] = '" & ctl.Name & "'")
                        arr_varCtl(C_CID, lngE) = lngCtlID
                        arr_varCtl(C_ORD, lngE) = CLng(0)
                        arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                        arr_varCtl(C_CAP, lngE) = strCap
                        arr_varCtl(C_LTR, lngE) = UCase$(Mid$(strCap, (intPos1 + 1), 1))
                        arr_varCtl(C_PAR, lngE) = Null
                        arr_varCtl(C_DEL, lngE) = CBool(False)
                        If IsNull(ctl.Parent) = False Then
                          If ctl.ControlType <> acCommandButton Then
                            If ctl.Parent.Name <> frm.Name Then
                              arr_varCtl(C_UN, lngE) = False
                              arr_varCtl(C_PAR, lngE) = ctl.Parent.Name
                            Else
                              arr_varCtl(C_UN, lngE) = True
                            End If
                          End If
                        End If
                      End If
                      intPos1 = InStr((intPos1 + 2), strCap, "&")
                    Loop
                  End If
                End If
              End With
            Next
          End With
        Next  ' ** For each Control: ctl.

        ' ** Find the parent for unattached labels.
        For lngY = 0& To (lngCtls - 1&)
          If arr_varCtl(C_UN, lngY) = True Then
            intPos1 = InStr(arr_varCtl(C_CNAM, lngY), "_lbl")
            If intPos1 > 0 Then
              strCtl = Left$(arr_varCtl(C_CNAM, lngY), (intPos1 - 1))
              For Each ctl In .Controls
                With ctl
                  If .Name = strCtl Then
                    arr_varCtl(C_PAR, lngY) = strCtl
                    Exit For
                  End If
                End With
              Next
            End If
          End If
        Next

      End With  ' ** This Form: frm.

      DoCmd.Close acForm, strForm, acSaveNo

      ' ** Binary Sort arr_varCtl() array.
      For lngY = UBound(arr_varCtl, 2) To 1& Step -1
        For lngZ = 0 To (lngY - 1)
          If arr_varCtl(C_LTR, lngZ) > arr_varCtl(C_LTR, (lngZ + 1)) Then
            For lngV = 0& To C_ELEMS
              varTmp00 = arr_varCtl(lngV, lngZ)
              arr_varCtl(lngV, lngZ) = arr_varCtl(lngV, (lngZ + 1))
              arr_varCtl(lngV, (lngZ + 1)) = varTmp00
              varTmp00 = Empty
            Next  ' ** lngV.
          End If
        Next  ' ** lngZ.
      Next  ' ** lngY.

      ' ** Check for dupes.
'THIS IS STILL WITHIN THE SAME FORM, EVEN THOUGH THE FORM IS NOW CLOSED.
      For lngY = 0& To (lngCtls - 1&)
        If arr_varCtl(C_ORD, lngY) = 0& And arr_varCtl(C_DEL, lngY) = False Then
          arr_varCtl(C_ORD, lngY) = 1&
          lngTmp02 = 1&
          For lngZ = 0& To (lngCtls - 1&)
            If lngZ <> lngY Then  ' ** Skip its own record.
              If arr_varCtl(C_LTR, lngZ) = arr_varCtl(C_LTR, lngY) And arr_varCtl(C_DEL, lngZ) = False Then
'SINCE THIS IS ONLY LOOKING AT LETTERS, DIFFERENT CONTROLS WITH THE SAME
'LETTER (SUCH AS _lbl_dim, _lbl_dim_hi) WILL BE MATCHED HERE, AS WELL AS
'CONTROLS ON DIFFERENT TABS OF A TAB CONTROL.
                lngTmp02 = lngTmp02 + 1&
                arr_varCtl(C_ORD, lngZ) = lngTmp02
                'Exit For
              End If
            End If
          Next  ' ** lngZ.
        End If
      Next  ' ** lngY.

      'CHECK FOR DUPED CTL_ID'S?
      For lngY = 0& To (lngCtls - 1&)
        For lngZ = (lngY + 1&) To (lngCtls - 1&)
          If arr_varCtl(C_DEL, lngZ) = False Then
            If arr_varCtl(C_CID, lngZ) = arr_varCtl(C_CID, lngY) Then
              arr_varCtl(C_DEL, lngZ) = CBool(True)
            End If
          End If
        Next
      Next

      lngTmp02 = 0&
      For lngY = 0& To (lngCtls - 1&)
        If arr_varCtl(C_DEL, lngY) = True Then lngTmp02 = lngTmp02 + 1&
      Next
      If lngTmp02 > 0& Then
        Debug.Print "'DUPES: " & CLng(lngTmp02)
        'Stop
      End If

      ' ** Add shortcut keys to tblForm_Shortcut.
      For lngY = 0& To (lngCtls - 1&)
        If arr_varCtl(C_DEL, lngY) = False Then
          blnAdd = False: blnEdit = False
          With rst1
            .MoveFirst
            .FindFirst "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And [frm_id] = " & CStr(arr_varCtl(C_FID, lngY)) & " And " & _
              "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngY)) & " And [fs_letter] = '" & arr_varCtl(C_LTR, lngY) & "' And " & _
              "[keydowntype_type] = " & CStr(taKeyDown_Alt)
              If .NoMatch = True Then
                .MoveFirst
                .FindFirst "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And [frm_id] = " & CStr(arr_varCtl(C_FID, lngY)) & " And " & _
                  "[fs_control] = '" & CStr(arr_varCtl(C_CNAM, lngY)) & "' And [fs_letter] = '" & arr_varCtl(C_LTR, lngY) & "' And " & _
                  "[keydowntype_type] = " & CStr(taKeyDown_Alt)
                If .NoMatch = True Then
                  'Stop
                End If
              End If
              varTmp00 = DLookup("[ctl_name]", "tblForm_Control", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And " & _
                "[frm_id] = " & CStr(arr_varCtl(C_FID, lngY)) & " And [ctl_id] = " & CStr(arr_varCtl(C_CID, lngY)))
              If IsNull(varTmp00) = True Then
                Debug.Print "'dbs_id = " & CStr(arr_varCtl(C_DID, lngY)) & "; frm_id = " & CStr(arr_varCtl(C_FID, lngY)) & " " & _
                  DLookup("[frm_name]", "tblForm", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And " & _
                  "[frm_id] = " & CStr(arr_varCtl(C_FID, lngY))) & "; ctl_id = " & CStr(arr_varCtl(C_CID, lngY)) & " " & _
                  DLookup("[ctl_name]", "tblForm_Control", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And " & _
                  "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngY))) & "; " & _
                  "keydowntype_type = " & CStr(taKeyDown_Alt) & " Alt;  fs_letter = " & arr_varCtl(C_LTR, lngY)
                Stop
              Else
                If ![frm_id] <> arr_varCtl(C_FID, lngY) Then
                  Debug.Print "'NO MATCH: " & .NoMatch
                  Debug.Print "'dbs_id = " & CStr(arr_varCtl(C_DID, lngY)) & "; frm_id = " & CStr(arr_varCtl(C_FID, lngY)) & " " & _
                  DLookup("[frm_name]", "tblForm", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And " & _
                  "[frm_id] = " & CStr(arr_varCtl(C_FID, lngY))) & "; ctl_id = " & CStr(arr_varCtl(C_CID, lngY)) & " " & _
                  DLookup("[ctl_name]", "tblForm_Control", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And " & _
                  "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngY))) & "; " & _
                  "keydowntype_type = " & CStr(taKeyDown_Alt) & " Alt;  fs_letter = " & arr_varCtl(C_LTR, lngY)
                DoEvents
              End If

              ' ** Unique Index:  NOT QUITE!
              ' **   frm_id
              ' **   fs_letter
              ' **   fs_order
              Select Case .NoMatch
              Case True
                blnAdd = True
                .AddNew
                ![dbs_id] = arr_varCtl(C_DID, lngY)
                ![frm_id] = arr_varCtl(C_FID, lngY)
                ![fs_order] = arr_varCtl(C_ORD, lngY)
                ![keydowntype_type] = taKeyDown_Alt
                ![fs_key] = arr_varCtl(C_LTR, lngY)
                ![ctl_id] = arr_varCtl(C_CID, lngY)
                ![fs_control] = arr_varCtl(C_CNAM, lngY)
                ![ctltype_type] = arr_varCtl(C_CTYP, lngY)
                ![fs_caption] = arr_varCtl(C_CAP, lngY)
                ![fs_unattached] = arr_varCtl(C_UN, lngY)
                If IsNull(arr_varCtl(C_PAR, lngY)) = False Then
                  ![fs_parent] = arr_varCtl(C_PAR, lngY)
                Else
                  ![fs_parent] = Null
                End If
                ![fs_linenum] = Null
                ![fs_shift] = False
                ![fs_alt] = True
                ![fs_ctrl] = False
                ![fs_letter] = arr_varCtl(C_LTR, lngY)
                ![keycode_constant] = Null
                ![fs_datemodified] = Now()
                blnFound = False
On Error Resume Next
                .Update
                If ERR.Number <> 0 Then
                  blnFound = True
On Error GoTo 0
                  lngTmp02 = 0&
                  Do While blnFound = True
                    lngTmp02 = lngTmp02 + 1&
                    arr_varCtl(C_ORD, lngY) = arr_varCtl(C_ORD, lngY) + 1&
                    blnFound = False
                    ![fs_order] = arr_varCtl(C_ORD, lngY)
On Error Resume Next
                    .Update
                    If ERR.Number <> 0 Then
                      blnFound = True
On Error GoTo 0
                      If lngTmp02 > 100& Then
                        Debug.Print "'dbs_id = " & CStr(arr_varCtl(C_DID, lngY)) & "; frm_id = " & CStr(arr_varCtl(C_FID, lngY)) & " " & _
                          DLookup("[frm_name]", "tblForm", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And " & _
                          "[frm_id] = " & CStr(arr_varCtl(C_FID, lngY))) & "; ctl_id = " & CStr(arr_varCtl(C_CID, lngY)) & " " & _
                          DLookup("[ctl_name]", "tblForm_Control", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngY)) & " And " & _
                          "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngY))) & "; " & _
                          "keydowntype_type = " & CStr(taKeyDown_Alt) & " Alt;  fs_letter = " & arr_varCtl(C_LTR, lngY)
                        Stop
                        Exit Do
                      End If
                    Else
On Error GoTo 0
                    End If
                  Loop
                Else
On Error GoTo 0
                End If
              Case False
                If ![frm_id] <> arr_varCtl(C_FID, lngY) Then
                  blnEdit = True
                  .Edit
                  ![frm_id] = arr_varCtl(C_FID, lngY)
                  ![fs_datemodified] = Now()
                  .Update
                End If
                If ![fs_letter] <> arr_varCtl(C_LTR, lngY) Then
                  blnEdit = True
                  .Edit
                  ![fs_key] = arr_varCtl(C_LTR, lngY)
                  ![fs_letter] = arr_varCtl(C_LTR, lngY)
                  ![fs_datemodified] = Now()
                  .Update
                End If
                If ![fs_order] <> arr_varCtl(C_ORD, lngY) Then
                  blnEdit = True
                  .Edit
                  ![fs_order] = arr_varCtl(C_ORD, lngY)
                  ![fs_datemodified] = Now()
On Error Resume Next
                  .Update
                  If ERR.Number <> 0 Then
On Error GoTo 0
                    ![fs_order] = ![fs_order] + 1&
On Error Resume Next
                    .Update
                    If ERR.Number <> 0 Then
On Error GoTo 0
                      ![fs_order] = ![fs_order] + 1&
                      .Update
                    Else
On Error GoTo 0
                    End If
                  Else
On Error GoTo 0
                  End If

                End If
                If arr_varCtl(C_CID, lngY) > 0& Then
                  lngCtlID = arr_varCtl(C_CID, lngY)
                  Select Case IsNull(![ctl_id])
                  Case True
                    blnEdit = True
                    .Edit
                    ![ctl_id] = lngCtlID
                    ![fs_datemodified] = Now()
                    .Update
                  Case False
                    If ![ctl_id] <> arr_varCtl(C_CID, lngY) Then
                      blnEdit = True
                      .Edit
                      ![ctl_id] = lngCtlID
                      ![fs_datemodified] = Now()
                      .Update
                    End If
                  End Select
                End If
                If ![fs_control] <> arr_varCtl(C_CNAM, lngY) Then
                  blnEdit = True
                  .Edit
                  ![fs_control] = arr_varCtl(C_CNAM, lngY)
                  ![fs_datemodified] = Now()
                  .Update
                End If
                If ![ctltype_type] <> arr_varCtl(C_CTYP, lngY) Then
                  blnEdit = True
                  .Edit
                  ![ctltype_type] = arr_varCtl(C_CTYP, lngY)
                  ![fs_datemodified] = Now()
                  .Update
                End If
                If ![fs_caption] <> arr_varCtl(C_CAP, lngY) Then
                  blnEdit = True
                  .Edit
                  ![fs_caption] = arr_varCtl(C_CAP, lngY)
                  ![fs_datemodified] = Now()
                  .Update
                End If
                If ![fs_shift] = True Or ![fs_alt] = False Or ![fs_ctrl] = True Then
                  blnEdit = True
                  .Edit
                  ![fs_shift] = False
                  ![fs_alt] = True
                  ![fs_ctrl] = False
                  ![fs_datemodified] = Now()
                  .Update
                End If
                If ![fs_unattached] <> arr_varCtl(C_UN, lngY) Then
                  blnEdit = True
                  .Edit
                  ![fs_unattached] = arr_varCtl(C_UN, lngY)
                  ![fs_datemodified] = Now()
                  .Update
                End If
                If (IsNull(arr_varCtl(C_PAR, lngY)) = True And IsNull(![fs_parent]) = False) Or _
                    (IsNull(arr_varCtl(C_PAR, lngY)) = False And IsNull(![fs_parent]) = True) Then
                  If IsNull(arr_varCtl(C_PAR, lngY)) = True Then
                    blnEdit = True
                    .Edit
                    ![fs_parent] = Null
                    ![fs_datemodified] = Now()
                    .Update
                  Else
                    blnEdit = True
                    .Edit
                    ![fs_parent] = arr_varCtl(C_PAR, lngY)
                    ![fs_datemodified] = Now()
                    .Update
                  End If
                Else
                  If (IsNull(arr_varCtl(C_PAR, lngY)) = False And IsNull(![fs_parent]) = False) Then
                    If ![fs_parent] <> arr_varCtl(C_PAR, lngY) Then
                      blnEdit = True
                      .Edit
                      ![fs_parent] = arr_varCtl(C_PAR, lngY)
                      ![fs_datemodified] = Now()
                      .Update
                    End If
                  End If
                End If
              End Select  ' ** NoMatch.
            End If
          End With  ' ** rst1.
        End If  ' ** C_DEL.
      Next  ' ** For each Control in arr_varCtl(): lngY, ctl.

      arr_varFrm(F_ARR, lngX) = arr_varCtl

    Next  ' ** For each Form: lngX, frm.

End If  ' ** blnSkip.

    ' ** Now check shortcuts found only in Form_KeyDown() events.
    ' ** zz_qry_Form_Shortcut_04 (tblVBComponent_Procedure, just Form_KeyDown()
    ' ** event handlers, by specified CurrentAppName()), linked to tblForm.
    Set qdf = .QueryDefs("zz_qry_Form_Shortcut_05")
    Set rst1 = qdf.OpenRecordset
    With rst1
      .MoveLast
      lngProcs = .RecordCount
      .MoveFirst
      arr_varProc = .GetRows(lngProcs)
      ' ********************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     dbs_id                 P_DID
      ' **     2       1     dbs_name               P_DNAM
      ' **     3       2     vbcom_id               P_CID
      ' **     4       3     vbcom_name             P_CNAM
      ' **     5       4     frm_id                 P_FID
      ' **     6       5     frm_name               P_FNAM
      ' **     7       6     vbcomproc_id           P_PID
      ' **     8       7     vbcomproc_name         P_PNAM
      ' **     9       8     vbcomproc_line_beg     P_LIN
      ' **    10       9     arr_varType() Array    P_ARR
      ' **
      ' ********************************************************
      .Close
    End With

    Set vbp = Application.VBE.ActiveVBProject
    With vbp
      For lngW = 0& To (lngProcs - 1&)
        If Left$(arr_varProc(P_FNAM, lngW), 3) <> "zz_" And Left$(arr_varProc(P_CNAM, lngW), 8) <> "Form_zz_" Then
          Set vbc = .VBComponents(arr_varProc(P_CNAM, lngW))
          With vbc
            strModName = .Name
            Set cod = .CodeModule
            With cod

'If strModName = "Form_frmJournal_Columns" Then

              lngProcEnd = 0&
              lngLines = .CountOfLines
              For lngX = arr_varProc(P_LIN, lngW) To lngLines
                If .ProcOfLine(lngX, vbext_pk_Proc) <> arr_varProc(P_PNAM, lngW) Then
                  lngProcEnd = lngX
                  Exit For
                End If
                If lngX = lngLines And lngProcEnd = 0& Then
                  ' ** If the procedure we're checking happens to be the last
                  ' ** one in the module, it'll never encounter a new procedure.
                  lngProcEnd = lngLines
                End If
              Next
              If (lngProcEnd - arr_varProc(P_LIN, lngW)) < 10& Then Stop

              ' ** TaKeyDown enumeration:
              ' **    0 taKeyDown_Plain
              ' **    1 taKeyDown_Ctrl
              ' **    2 taKeyDown_Alt
              ' **    3 taKeyDown_Shift
              ' **    4 taKeyDown_CtrlAlt
              ' **    5 taKeyDown_CtrlShift
              ' **    6 taKeyDown_AltShift
              ' **    7 taKeyDown_CtrlAltShift
              ' **   -1 taKeyDown_Unknown
              lngPlain = 0&: lngCtrl = 1&: lngAlt = 2&: lngShift = 3&
              lngCtrlAlt = 4&: lngCtrlShift = 5&: lngAltShift = 6&: lngCtrlAltShift = 7&: lngUnknown = 8&

              lngTypes = 9&
              ReDim arr_varType(T_ELEMS, (lngTypes - 1&))  ' ** One extra, empty array row.  WHAT DID I MEAN?
              ' *************************************************
              ' ** Array: arr_varType()
              ' **
              ' **   Element  Name                  Constant
              ' **   =======  ====================  ===========
              ' **      0     keydowntype_name      T_TYP_NAM
              ' **      1     keydowntype_type      T_TYP_TYP
              ' **      2     Beginning Line        T_SEC_BEG
              ' **      3     Ending Line           T_SEC_END
              ' **      4     Line Indent           T_SEC_IND
              ' **      5     Count of Keys         T_KEYS
              ' **      6     arr_varKey() Array    T_ARR
              ' **
              ' *************************************************
              arr_varType(T_TYP_NAM, lngPlain) = "Plain"
              arr_varType(T_TYP_TYP, lngPlain) = taKeyDown_Plain
              arr_varType(T_SEC_BEG, lngPlain) = 0&
              arr_varType(T_SEC_END, lngPlain) = 0&
              arr_varType(T_SEC_IND, lngPlain) = 0&
              arr_varType(T_KEYS, lngPlain) = 0&
              arr_varType(T_ARR, lngPlain) = Empty
              arr_varType(T_TYP_NAM, lngCtrl) = "Ctrl"
              arr_varType(T_TYP_TYP, lngCtrl) = taKeyDown_Ctrl
              arr_varType(T_SEC_BEG, lngCtrl) = 0&
              arr_varType(T_SEC_END, lngCtrl) = 0&
              arr_varType(T_SEC_IND, lngCtrl) = 0&
              arr_varType(T_KEYS, lngCtrl) = 0&
              arr_varType(T_ARR, lngCtrl) = Empty
              arr_varType(T_TYP_NAM, lngAlt) = "Alt"
              arr_varType(T_TYP_TYP, lngAlt) = taKeyDown_Alt
              arr_varType(T_SEC_BEG, lngAlt) = 0&
              arr_varType(T_SEC_END, lngAlt) = 0&
              arr_varType(T_SEC_IND, lngAlt) = 0&
              arr_varType(T_KEYS, lngAlt) = 0&
              arr_varType(T_ARR, lngAlt) = Empty
              arr_varType(T_TYP_NAM, lngShift) = "Shift"
              arr_varType(T_TYP_TYP, lngShift) = taKeyDown_Shift
              arr_varType(T_SEC_BEG, lngShift) = 0&
              arr_varType(T_SEC_END, lngShift) = 0&
              arr_varType(T_SEC_IND, lngShift) = 0&
              arr_varType(T_KEYS, lngShift) = 0&
              arr_varType(T_ARR, lngShift) = Empty
              arr_varType(T_TYP_NAM, lngCtrlAlt) = "Ctrl-Alt"
              arr_varType(T_TYP_TYP, lngCtrlAlt) = taKeyDown_CtrlAlt
              arr_varType(T_SEC_BEG, lngCtrlAlt) = 0&
              arr_varType(T_SEC_END, lngCtrlAlt) = 0&
              arr_varType(T_SEC_IND, lngCtrlAlt) = 0&
              arr_varType(T_KEYS, lngCtrlAlt) = 0&
              arr_varType(T_ARR, lngCtrlAlt) = Empty
              arr_varType(T_TYP_NAM, lngCtrlShift) = "Ctrl-Shift"
              arr_varType(T_TYP_TYP, lngCtrlShift) = taKeyDown_CtrlShift
              arr_varType(T_SEC_BEG, lngCtrlShift) = 0&
              arr_varType(T_SEC_END, lngCtrlShift) = 0&
              arr_varType(T_SEC_IND, lngCtrlShift) = 0&
              arr_varType(T_KEYS, lngCtrlShift) = 0&
              arr_varType(T_ARR, lngCtrlShift) = Empty
              arr_varType(T_TYP_NAM, lngAltShift) = "Alt-Shift"
              arr_varType(T_TYP_TYP, lngAltShift) = taKeyDown_AltShift
              arr_varType(T_SEC_BEG, lngAltShift) = 0&
              arr_varType(T_SEC_END, lngAltShift) = 0&
              arr_varType(T_SEC_IND, lngAltShift) = 0&
              arr_varType(T_KEYS, lngAltShift) = 0&
              arr_varType(T_ARR, lngAltShift) = Empty
              arr_varType(T_TYP_NAM, lngCtrlAltShift) = "Ctrl-Alt-Shift"
              arr_varType(T_TYP_TYP, lngCtrlAltShift) = taKeyDown_CtrlAltShift
              arr_varType(T_SEC_BEG, lngCtrlAltShift) = 0&
              arr_varType(T_SEC_END, lngCtrlAltShift) = 0&
              arr_varType(T_SEC_IND, lngCtrlAltShift) = 0&
              arr_varType(T_KEYS, lngCtrlAltShift) = 0&
              arr_varType(T_ARR, lngCtrlAltShift) = Empty
              arr_varType(T_TYP_NAM, lngUnknown) = "{unk}"
              arr_varType(T_TYP_TYP, lngUnknown) = taKeyDown_Unknown
              arr_varType(T_SEC_BEG, lngUnknown) = 0&
              arr_varType(T_SEC_END, lngUnknown) = 0&
              arr_varType(T_SEC_IND, lngUnknown) = 0&
              arr_varType(T_KEYS, lngUnknown) = 0&
              arr_varType(T_ARR, lngUnknown) = Empty

              ' ** Find the beginning of each section of shortcut types.
              For lngX = arr_varProc(P_LIN, lngW) To lngProcEnd
                strLine = .Lines(lngX, 1)  ' ** Don't trim; we'll need the exact spacing. (If it's my indenting convention.)
                If strLine <> vbNullString Then

                  ' ** Get the character position of the statement's beginning.
                  lngLen = Len(strLine)
                  blnFound = False: intPos1 = 0: strLine2 = vbNullString
                  For lngY = 1& To lngLen
                    If Mid$(strLine, lngY, 1) <> " " And blnFound = False Then  ' ** Ignore leading spaces.
                      ' ** Found the first Character.
                      blnFound = True
                      If Mid$(strLine, lngY, 1) = "'" Then
                        ' ** A remark line, so skip it.
                        Exit For
                      Else
                        intPos2 = InStr(lngY, strLine, " ")  ' ** Find the first space within a statement.
                        If intPos2 > 0 Then
                          If IsNumeric(Trim(Left(strLine, intPos2))) = True Then
                            ' ** Move intPos1 up to the first statement character.
                            intPos1 = 0
                            For lngZ = intPos2 To lngLen
                              If Mid$(strLine, lngZ, 1) <> " " Then
                                ' ** The first statement character.
                                intPos1 = lngZ
                                strLine2 = Trim$(Mid$(strLine, intPos1))
                                Exit For
                              End If
                            Next
                            If intPos1 > 0 Then
                              Exit For
                            Else: Stop
                            End If
                          Else
                            ' ** lngY is already the first statement character.
                            intPos1 = lngY
                            strLine2 = Trim$(Mid$(strLine, intPos1))
                            Exit For
                          End If
                        Else
                          ' ** One word, so not our 'If' statement.
                          Exit For
                        End If
                      End If
                    End If
                  Next

'If strModName = "Form_frmJournal_Columns" Then
'Stop
'End If
                  ' ** See if this is one of the droids we're looking for.
                  If intPos1 > 0 And strLine2 <> vbNullString Then
                    intPos2 = InStr(strLine2, " ")
                    If intPos2 > 0 Then
                      If Left$(strLine2, 3) = "If " Then
                        If Right$(strLine2, 5) <> " Then" Then
                          ' ** A remark?
                          intPos2 = InStr(strLine2, " Then")
                          If intPos2 > 0 Then
                            strLine2 = Trim$(Left$(strLine2, ((intPos2 + Len(" Then")) - 1)))
                          Else
                            If Right$(strLine2, 1) = "_" Then
                              ' ** A line continuation of something we're not looking for.
                            Else: Stop
                            End If
                          End If
                        End If
                        Select Case strLine2
                        Case "If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then"
                          ' ** Plain keys.
                          If arr_varType(T_SEC_BEG, lngPlain) = 0& Then
                            arr_varType(T_SEC_BEG, lngPlain) = lngX
                            arr_varType(T_SEC_IND, lngPlain) = intPos1
                          Else: Stop
                          End If
                        Case "If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then"
                          ' ** Ctrl keys.
                          If arr_varType(T_SEC_BEG, lngCtrl) = 0& Then
                            arr_varType(T_SEC_BEG, lngCtrl) = lngX
                            arr_varType(T_SEC_IND, lngCtrl) = intPos1
                          Else: Stop
                          End If
                        Case "If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then"
                          ' ** Alt keys.
                          If arr_varType(T_SEC_BEG, lngAlt) = 0& Then
                            arr_varType(T_SEC_BEG, lngAlt) = lngX
                            arr_varType(T_SEC_IND, lngAlt) = intPos1
                          Else: Stop
                          End If
                        Case "If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then"
                          ' ** Shift keys.
                          If arr_varType(T_SEC_BEG, lngShift) = 0& Then
                            arr_varType(T_SEC_BEG, lngShift) = lngX
                            arr_varType(T_SEC_IND, lngShift) = intPos1
                          Else: Stop
                          End If
                        Case "If intCtrlDown And intAltDown And (Not intShiftDown) Then"
                          ' ** Ctrl-Alt keys.
                          If arr_varType(T_SEC_BEG, lngCtrlAlt) = 0& Then
                            arr_varType(T_SEC_BEG, lngCtrlAlt) = lngX
                            arr_varType(T_SEC_IND, lngCtrlAlt) = intPos1
                          Else: Stop
                          End If
                        Case "If intCtrlDown And (Not intAltDown) And intShiftDown Then"
                          ' ** Ctrl-Shift keys.
                          If arr_varType(T_SEC_BEG, lngCtrlShift) = 0& Then
                            arr_varType(T_SEC_BEG, lngCtrlShift) = lngX
                            arr_varType(T_SEC_IND, lngCtrlShift) = intPos1
                          Else: Stop
                          End If
                        Case "If (Not intCtrlDown) And intAltDown And intShiftDown Then"
                          ' ** Alt-Shift keys.
                          If arr_varType(T_SEC_BEG, lngAltShift) = 0& Then
                            arr_varType(T_SEC_BEG, lngAltShift) = lngX
                            arr_varType(T_SEC_IND, lngAltShift) = intPos1
                          Else: Stop
                          End If
                        Case "If intCtrlDown And intAltDown And intShiftDown Then"
                          ' ** Ctrl-Alt-Shift keys.
                          If arr_varType(T_SEC_BEG, lngCtrlAltShift) = 0& Then
                            arr_varType(T_SEC_BEG, lngCtrlAltShift) = lngX
                            arr_varType(T_SEC_IND, lngCtrlAltShift) = intPos1
                          Else: Stop
                          End If
                        Case Else
                          ' ** Other possible markers?
                          Select Case strLine2
                          Case "If KeyCode = vbKeyEscape Then"
                            ' **   DoCmd.Close acForm, THIS_NAME
                            ' **   DoCmd.OpenForm "frmMenu_Report"
                            ' ** End If
                            If arr_varType(T_SEC_BEG, lngUnknown) = 0& Then
                              arr_varType(T_SEC_BEG, lngUnknown) = lngX
                              arr_varType(T_SEC_IND, lngUnknown) = intPos1
                            Else: Stop
                            End If
                          Case Else
                            'If arr_varType(T_SEC_BEG, lngUnknown) = 0& Then
                            '  arr_varType(T_SEC_BEG, lngUnknown) = lngX
                            '  arr_varType(T_SEC_IND, lngUnknown) = intPos1
                            'Else: Stop
                            'End If
                          End Select
                        End Select
                      End If
                    End If
                  End If

                End If
              Next

              ' ** Make sure at least SOMTHING was found!
              lngRecs = 0&
              For lngX = 0& To (lngTypes - 1&)
                If arr_varType(T_SEC_BEG, lngX) = 0& Then
                  lngRecs = lngRecs + 1&
                End If
              Next
              If lngRecs = lngTypes Then Stop  ' ** None found?

'For lngX = 0& To (lngTypes - 1&)
'  If arr_varType(T_SEC_BEG, lngX) <> 0& Then
'    Debug.Print "'SEC BEG " & CStr(lngX + 1&) & ", " & arr_varType(T_TYP_NAM, lngX) & ": " & CStr(arr_varType(T_SEC_BEG, lngX))
'  Else
'    Debug.Print "'NO SEC " & CStr(lngX + 1&) & ", " & arr_varType(T_TYP_NAM, lngX) & "!"
'  End If
'Next
'Stop
'SEC BEG 1, Plain: 379
'SEC BEG 2, Ctrl: 511
'SEC BEG 3, Alt: 413
'NO SEC 4, Shift!
'NO SEC 5, Ctrl-Alt!
'SEC BEG 6, Ctrl-Shift: 650
'SEC BEG 7, Alt-Shift: 599
'NO SEC 8, Ctrl-Alt-Shift!
'NO SEC 9, {unk}!
              ' ** Find the end of each section.
              For lngX = 0& To (lngTypes - 1&)
                If arr_varType(T_SEC_BEG, lngX) <> 0& Then

'Debug.Print "'BEG SEC " & CStr(lngX + 1&) & ", " & arr_varType(T_TYP_NAM, lngX) & ": " & CStr(arr_varType(T_SEC_BEG, lngX))
'DoEvents
                  ' ** See if there are any more sections after this one.
                  lngZ = 0&
                  For lngY = 0& To (lngTypes - 1&) '(lngX + 1&) To (lngTypes - 1&)
                    If arr_varType(T_SEC_BEG, lngY) > arr_varType(T_SEC_BEG, lngX) Then
                      If lngZ = 0& Then
                        lngZ = arr_varType(T_SEC_BEG, lngY)  ' ** lngZ is the line too far.
                      Else
                        ' ** Since the order of types is not necessarily the order
                        ' ** they appear in code, make sure to find the very next one.
                        If arr_varType(T_SEC_BEG, lngY) < lngZ Then
                          lngZ = arr_varType(T_SEC_BEG, lngY)
                        End If
                      End If
                      'Exit For
                    End If
                  Next
'Debug.Print "'  lngZ: " & CStr(lngZ)
'DoEvents
'Stop
                  If lngZ = 0& Then
                    ' ** This is the last (or only) section of shortcut keys.
                    lngZ = lngProcEnd
                  End If

                  ' ** Now look for the matching 'End If'.
                  For lngY = arr_varType(T_SEC_BEG, lngX) To (lngZ - 1&)
                    strLine = .Lines(lngY, 1)  ' ** Don't trim; we'll need the exact spacing. (If it's my indenting convention.)
                    If strLine <> vbNullString Then
                      strLine2 = Trim$(strLine)
                      intPos1 = InStr(strLine2, " ")
                      If intPos1 > 0 Then
                        If IsNumeric(Trim(Left(strLine2, intPos1))) Then
                          strLine2 = Trim$(Mid$(strLine2, intPos1))  ' ** Strip off the line number.
                        End If
                        If Left$(strLine2, 6) = "End If" Then
                          If InStr(strLine, "End If") = arr_varType(T_SEC_IND, lngX) Then
                            arr_varType(T_SEC_END, lngX) = lngY  ' ** Let it keep incrementing till we're out of the loop.
                          Else
                            ' ** Not our 'End If', so keep looking.
                          End If
                        End If
                      End If
                    End If
                  Next

'Debug.Print "'END SEC " & CStr(lngX + 1&) & ", " & arr_varType(T_TYP_NAM, lngX) & ": " & CStr(arr_varType(T_SEC_END, lngX))
'DoEvents
'Stop
                  If arr_varType(T_SEC_END, lngX) = 0& Then
                    ' ** Couldn't find the right 'End If'!
                    Stop
                  End If

                End If
              Next  ' ** Each of the 7 shortcut key types: lngX.

'If strModName = "Form_frmJournal_Columns" Then
'Stop
'End If
              ' ** Now look for the keys.
              For lngX = 0& To (lngTypes - 1&)
                lngKeys = 0&
                ReDim arr_varKey(K_ELEMS, 0)
                If arr_varType(T_SEC_BEG, lngX) <> 0& Then
                  For lngY = arr_varType(T_SEC_BEG, lngX) To arr_varType(T_SEC_END, lngX)
                    strLine = .Lines(lngY, 1)
                    intPos1 = InStr(strLine, " vbKey")  ' ** vbKeyX
                    intPos2 = InStr(strLine, "'")
                    strLine2 = vbNullString
                    If intPos1 > 0 And (intPos2 = 0 Or intPos2 > intPos1) Then
                      If intPos2 > 0 Then
                        strLine2 = Left$(strLine, (intPos2 - 1))
                        strLine2 = Trim$(Mid$(strLine, (intPos1 + 1)))
                      Else
                        strLine2 = Trim$(Mid$(strLine, (intPos1 + 1)))
                      End If
                      intPos2 = InStr(strLine2, ",")
                      If intPos2 > 0 Then
                        ' ** Multiple keys on one line.
                        Do While intPos2 > 0
                          lngKeys = lngKeys + 1&
                          lngE = lngKeys - 1&
                          ReDim Preserve arr_varKey(K_ELEMS, lngE)
                          ' **************************************************
                          ' ** Array: arr_varKey()
                          ' **
                          ' **   Element  Name                  Constant
                          ' **   =======  ====================  ============
                          ' **      0     KeyCode Constant      K_CONST
                          ' **      1     Line Number           K_LINENUM
                          ' **      2     Key Constant Value    K_KEYVAL
                          ' **
                          ' **************************************************
                          arr_varKey(K_CONST, lngE) = Trim$(Left$(strLine2, (intPos2 - 1)))
                          arr_varKey(K_LINENUM, lngE) = lngY  'SAVE THIS!
                          strLine2 = Trim$(Mid$(strLine2, (intPos2 + 1)))
                          intPos2 = InStr(strLine2, ",")
                          If intPos2 = 0 Then
                            lngKeys = lngKeys + 1&
                            lngE = lngKeys - 1&
                            ReDim Preserve arr_varKey(K_ELEMS, lngE)
                            arr_varKey(K_CONST, lngE) = strLine2
                            arr_varKey(K_LINENUM, lngE) = lngY  'SAVE THIS!
                            arr_varKey(K_KEYVAL, lngE) = -1&
                            Exit Do
                          End If
                        Loop
                      Else
                        lngKeys = lngKeys + 1&
                        lngE = lngKeys - 1&
                        ReDim Preserve arr_varKey(K_ELEMS, lngE)
                        If InStr(strLine2, " ") > 0 Then strLine2 = Trim$(Left$(strLine2, InStr(strLine2, " ")))
                        arr_varKey(K_CONST, lngE) = strLine2
                        arr_varKey(K_LINENUM, lngE) = lngY  'SAVE THIS!
                        arr_varKey(K_KEYVAL, lngE) = -1&
                      End If
                    End If
                  Next  ' ** For each line in this key section: lngY.
                  If lngKeys = 0& Then
                    Stop
                  Else
                    For lngY = 0& To (lngKeys - 1&)
                      If IsNull(DLookup("[keycode_value]", "tblKeyCode", "[keycode_constant] = '" & arr_varKey(K_CONST, lngY) & "'")) = False Then
                        lngZ = DLookup("[keycode_value]", "tblKeyCode", "[keycode_constant] = '" & arr_varKey(K_CONST, lngY) & "'")
                        arr_varKey(K_KEYVAL, lngY) = lngZ
                      Else
                        Stop
                      End If
                    Next
                    arr_varType(T_KEYS, lngX) = lngKeys
                    arr_varType(T_ARR, lngX) = arr_varKey
                  End If
                End If
              Next  ' ** For each shortcut key type: lngX.

              arr_varProc(P_ARR, lngW) = arr_varType

'End If
            End With  ' ** This CodeModule: cod.
            Set cod = Nothing
          End With  ' ** This VBComponent: vbc.
          Set vbc = Nothing

        End If
      Next  ' ** For each Form_KeyDown() event procedure: lngW.
    End With  ' ** ActiveVBProject: vbp.
    Set vbp = Nothing

    Set rst1 = .OpenRecordset("tblForm_Shortcut", dbOpenDynaset, dbConsistent)
    With rst1

      ' ********************************************************
      ' ** Array: arr_varProc()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ==========
      ' **     1       0     dbs_id                 P_DID
      ' **     2       1     dbs_name               P_DNAM
      ' **     3       2     vbcom_id               P_CID
      ' **     4       3     vbcom_name             P_CNAM
      ' **     5       4     frm_id                 P_FID
      ' **     6       5     frm_name               P_FNAM
      ' **     7       6     vbcomproc_id           P_PID
      ' **     8       7     vbcomproc_name         P_PNAM
      ' **     9       8     vbcomproc_line_beg     P_LIN
      ' **    10       9     arr_varType() Array    P_ARR
      ' **
      ' ********************************************************
      For lngX = 0& To (lngProcs - 1&)
        If Left$(arr_varProc(P_FNAM, lngX), 3) <> "zz_" And Left$(arr_varProc(P_CNAM, lngX), 8) <> "Form_zz_" Then
          arr_varTmp04 = arr_varProc(P_ARR, lngX)
          lngTypes = (UBound(arr_varTmp04, 2) + 1&)
          ' *****************************************************
          ' ** Array: arr_varType(), arr_varTmp04()
          ' **
          ' **   Element  Name                     Constant
          ' **   =======  =======================  ============
          ' **      0     keydowntype_name         T_TYP_NAM
          ' **      1     keydowntype_type         T_TYP_TYP
          ' **      2     Beginning Line Number    T_SEC_BEG
          ' **      3     Ending Line Number       T_SEC_END
          ' **      4     Line Indent              T_SEC_IND
          ' **      5     Count of Keys            T_KEYS
          ' **      6     arr_varKey() Array       T_ARR
          ' **
          ' *****************************************************
          For lngY = 0& To (lngTypes - 1&)
            If arr_varTmp04(T_KEYS, lngY) > 0& Then
              arr_varTmp05 = arr_varTmp04(T_ARR, lngY)
              lngKeys = (UBound(arr_varTmp05, 2) + 1&)
              lngTypeType = -1&
              ' **************************************************
              ' ** Array: arr_varKey(), arr_varTmp05()
              ' **
              ' **   Element  Name                  Constant
              ' **   =======  ====================  ============
              ' **      0     KeyCode Constant      K_CONST
              ' **      1     Line Number           K_LINENUM
              ' **      2     Key Constant Value    K_KEYVAL
              ' **
              ' **************************************************
              For lngZ = 0& To (lngKeys - 1&)
                ' ** HOW DOES fs_order COME INTO PLAY HERE?
                blnAdd = False: blnEdit = False
                lngTypeType = arr_varTmp04(T_TYP_TYP, lngY)
                Select Case arr_varTmp04(T_TYP_NAM, lngY)
                Case "Plain"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = False And [fs_alt] = False And [fs_ctrl] = False"
                Case "Ctrl"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = False And [fs_alt] = False And [fs_ctrl] = True"
                Case "Alt"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = False And [fs_alt] = True And [fs_ctrl] = False"
                Case "Shift"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = True And [fs_alt] = False And [fs_ctrl] = False"
                Case "Ctrl-Alt"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = False And [fs_alt] = True And [fs_ctrl] = True"
                Case "Ctrl-Shift"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = True And [fs_alt] = False And [fs_ctrl] = True"
                Case "Alt-Shift"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = True And [fs_alt] = True And [fs_ctrl] = False"
                Case "Ctrl-Alt-Shift"
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = true And [fs_alt] = True And [fs_ctrl] = True"
                Case "{unk}"  ' ** Currently only vbKeyEscape, Plain.
                  .FindFirst "[frm_id] = " & CStr(arr_varProc(P_FID, lngX)) & " And " & _
                    "[fs_letter] = '" & CStr(arr_varTmp05(K_KEYVAL, lngZ)) & "' And " & _
                    "[fs_shift] = False And [fs_alt] = False And [fs_ctrl] = False"
                End Select
                If .NoMatch = True Then
                  blnAdd = True
                  .AddNew
                Else
                  .Edit
                End If
                If blnAdd = True Then
                  ![dbs_id] = arr_varProc(P_DID, lngX)
                  ![frm_id] = arr_varProc(P_FID, lngX)
                  ![keydowntype_type] = lngTypeType
                  strTmp01 = Mid(arr_varTmp05(K_CONST, lngZ), 6)
                  If Len(strTmp01) = 1 Then
                    ![fs_key] = strTmp01
                  ElseIf Len(strTmp01) = 2 And Left$(strTmp01, 1) = "F" Then
                    ![fs_key] = strTmp01
                  Else
                    ' ** Non-printing characters for fs_key:
                    Select Case strTmp01
                    Case "Tab"  ' ** vbKeyTab
                      ![fs_key] = "{Tab}"
                    Case "Return"  ' ** vbKeyReturn
                      ![fs_key] = "{Enter}"
                    Case "Up"  ' ** vbKeyUp
                      ![fs_key] = "{Up}"
                    Case "Down"  ' ** vbKeyDown
                      ![fs_key] = "{Down}"
                    Case "Left"  ' ** vbKeyLeft
                      ![fs_key] = "{Left}"
                    Case "Right"  ' ** vbKeyRight
                      ![fs_key] = "{Right}"
                    Case "PageUp"  ' ** vbKeyPageUp
                      ![fs_key] = "{PgUp}"
                    Case "PageDown"  ' ** vbKeyPageDown
                      ![fs_key] = "{PgDn}"
                    Case "Home"  ' ** vbKeyHome
                      ![fs_key] = "{Home}"
                    Case "End"  ' ** vbKeyEnd
                      ![fs_key] = "{End}"
                    Case "Delete"  ' ** vbKeyDelete
                      ![fs_key] = "{Del}"
                    Case "Escape"  ' ** vbKeyEscape
                      ![fs_key] = "{Esc}"
                    End Select
                  End If
                  ![ctl_id] = Null                                 ' ** Not required (none for these).
                  ![fs_control] = "Form_KeyDown"  ' ** Required.
                  ![ctltype_type] = acNone        ' ** Required, 99 (my own).
                  ![fs_linenum] = arr_varTmp05(K_LINENUM, lngZ)
                  ![fs_letter] = CStr(arr_varTmp05(K_KEYVAL, lngZ))
                  ![keycode_constant] = arr_varTmp05(K_CONST, lngZ)
                End If
                varTmp00 = vbNullString
                ' ** Special cases.
                Select Case arr_varTmp05(K_CONST, lngZ)
                Case "vbKeyN"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then
                    varTmp00 = "Next Record - Ctrl+N": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyP"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" And _
                      (arr_varProc(P_FNAM, lngX) <> "frmJournal_Columns" And arr_varProc(P_FNAM, lngX) <> "frmJournal_Columns_Sub") Then
                    varTmp00 = "Previous Record - Ctrl+P": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" And _
                      (arr_varProc(P_FNAM, lngX) = "frmJournal_Columns" Or arr_varProc(P_FNAM, lngX) = "frmJournal_Columns_Sub") Then
                    varTmp00 = "Print - Ctrl+P": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyS"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then
                    varTmp00 = "Save Record - Ctrl+S": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyX"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Alt" Then
                    varTmp00 = "E&xit": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyF4"
                  varTmp00 = "Dropdown - F4": blnEdit = True
                  If IsNull(![fs_caption]) = False Then
                    If ![fs_caption] = varTmp00 Then
                      varTmp00 = vbNullString: blnEdit = False
                    End If
                  End If
                Case "vbKeyF5"
                  varTmp00 = "Recalc - F5": blnEdit = True
                  If IsNull(![fs_caption]) = False Then
                    If ![fs_caption] = varTmp00 Then
                      varTmp00 = vbNullString: blnEdit = False
                    End If
                  End If
                Case "vbKeyF7"
                  varTmp00 = "Date Picker - F7": blnEdit = True
                  If IsNull(![fs_caption]) = False Then
                    If ![fs_caption] = varTmp00 Then
                      varTmp00 = vbNullString: blnEdit = False
                    End If
                  End If
                Case "vbKeyF8"
                  varTmp00 = "Date Picker - F8": blnEdit = True
                  If IsNull(![fs_caption]) = False Then
                    If ![fs_caption] = varTmp00 Then
                      varTmp00 = vbNullString: blnEdit = False
                    End If
                  End If
                Case "vbKeyEscape"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Cancel Form - Esc": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyTab"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Next Field - Tab": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Shift" Then
                    varTmp00 = "Previous Field - Shift-Tab": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then
                    varTmp00 = "Exit Table Foreward - Ctrl-Tab": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl-Shift" Then
                    varTmp00 = "Exit Table Backward - Ctrl-Shift-Tab": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyReturn"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Next Field - Enter": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Shift" Then
                    varTmp00 = "Previous Field - Shift-Enter": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then
                    varTmp00 = "Exit Table Foreward - Ctrl-Enter": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl-Shift" Then
                    varTmp00 = "Exit Table Backward - Ctrl-Shift-Enter": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyUp"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Previous Record - Up Arrow": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyDown"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Next Record - Down Arrow": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyPageUp"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Previous Page - PageUp": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then
                    varTmp00 = "First Page - Ctrl-PageUp": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyPageDown"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Next Page - PageDown": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  ElseIf arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then
                    varTmp00 = "Last Page - Ctrl-PageDown": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case "vbKeyHome"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then

                  End If
                Case "vbKeyEnd"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then

                  End If
                Case "vbKeyLeft"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then

                  End If
                Case "vbKeyRight"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then

                  End If
                Case "vbKeyDelete"
                  If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                    varTmp00 = "Delete Record - Delete": blnEdit = True
                    If IsNull(![fs_caption]) = False Then
                      If ![fs_caption] = varTmp00 Then
                        varTmp00 = vbNullString: blnEdit = False
                      End If
                    End If
                  End If
                Case Else
                  ' ** Others?
                End Select
                If blnAdd = True Or blnEdit = True Then
                  If varTmp00 <> vbNullString Then
                    ![fs_caption] = varTmp00
                  ElseIf blnAdd = True Then  ' ** Since there are no other captions except those created above,
                    ![fs_caption] = Null     ' ** there should be no case where a caption goes away.
                  End If
                End If
' vbKeyEscape
' cmdClose, cmdCancel, etc., on that form.
' Close Form - Alt+X  CHECK ALL FORMS FOR THIS ControlTipText!
' vbKeyF7 or or vbKeyF8
' cmdCalendar on that form (or cmdCalendar1, 2, etc.)
' Date Picker - F7 or F8
' vbKeyPageUp
'
'
' vbKeyPageDown
'
'
' vbKeyHome
'
'
' vbKeyEnd
'
'
'CAN WE FIND THESE VIA THE CODE OR CODE REMARKS?
                Select Case arr_varTmp04(T_TYP_NAM, lngY)
                Case "Plain"
                  ' ** No checks.
                Case "Ctrl"
                  Select Case blnAdd
                  Case True
                    ![fs_ctrl] = True
                  Case False
                    ' ** Let stand.
                  End Select
                Case "Alt"
                  Select Case blnAdd
                  Case True
                    ![fs_alt] = True
                  Case False
                    ' ** Let stand.
                  End Select
                Case "Shift"
                  Select Case blnAdd
                  Case True
                    ![fs_shift] = True
                  Case False
                    ' ** Let stand.
                  End Select
                Case "Ctrl-Alt"
                  Select Case blnAdd
                  Case True
                    ![fs_ctrl] = True
                    ![fs_alt] = True
                  Case False
                    ' ** Let stand.
                  End Select
                Case "Ctrl-Shift"
                  Select Case blnAdd
                  Case True
                    ![fs_ctrl] = True
                    ![fs_shift] = True
                  Case False
                    ' ** Let stand.
                  End Select
                Case "Alt-Shift"
                  Select Case blnAdd
                  Case True
                    ![fs_alt] = True
                    ![fs_shift] = True
                  Case False
                    ' ** Let stand.
                  End Select
                Case "Ctrl-Alt-Shift"
                  Select Case blnAdd
                  Case True
                    ![fs_ctrl] = True
                    ![fs_alt] = True
                    ![fs_shift] = True
                  Case False
                    ' ** Let stand.
                  End Select
                Case "{unk}"
                  ' ** Currently only vbKeyEscape, Plain.
Stop
                End Select
                If blnAdd = True Then
                  ![fs_unattached] = CBool(True)
                  ![fs_parent] = Null
                End If
                If blnAdd = False Then
                  If IsNull(![fs_linenum]) = True Then
                    blnEdit = True
                    ![fs_linenum] = arr_varTmp05(K_LINENUM, lngZ)
                  Else
                    If ![fs_linenum] <> arr_varTmp05(K_LINENUM, lngZ) Then
                      blnEdit = True
                      ![fs_linenum] = arr_varTmp05(K_LINENUM, lngZ)
                    End If
                  End If
                End If
                If blnAdd = True Or blnEdit = True Then
                  ![fs_datemodified] = Now()
                End If
                If blnAdd = True Then
                  ![fs_order] = 1&
                End If
                If blnAdd = True Then
On Error Resume Next
                  .Update
                  If ERR.Number <> 0 Then
On Error GoTo 0
                    For lngW = 2& To 20&
                      ![fs_order] = lngW
On Error Resume Next
                      .Update
                      If ERR.Number <> 0 Then
                        ' ** Go round again.
On Error GoTo 0
                      Else
On Error GoTo 0
                        Exit For
                      End If
                    Next
                  Else
On Error GoTo 0
                  End If
                Else
                  If blnEdit = True Then
                    .Update
                  Else
                    .CancelUpdate
                  End If
                End If
              Next  ' ** For each Key in arr_varKey() (arr_varTmp05): lngZ.
            End If
          Next  ' ** For each Type in arr_varType() (arr_varTmp04): lngY.
        End If  ' ** zz_'s.
      Next  ' ** For each Procedure in arr_varProc(): lngX.
      .Close
    End With

    ' ** Check for shortcuts changed or deleted.
    If lngCtls > 0& Then
      For lngV = 0& To (lngFrms - 1&)
        ' ** tblForm_Shortcut, by specified [frmid].
        Set qdf = .QueryDefs("zz_qry_Form_Shortcut_03")
        With qdf.Parameters
          ![frmid] = arr_varFrm(F_FID, lngV)
        End With
        Set rst2 = qdf.OpenRecordset
        With rst2
          If .BOF = True And .EOF = True Then
            ' ** No shortcuts for this form.
          Else
            .MoveLast
            lngRecs = .RecordCount
            .MoveFirst
            arr_varTmp03 = arr_varFrm(F_ARR, lngV)
            lngCtls = (UBound(arr_varTmp03, 2) + 1&)
            For lngW = 1& To lngRecs
              ' ** First check control shortcuts.
              blnFound = False
              For lngX = 0& To (lngCtls - 1&)
                If arr_varTmp03(C_CID, lngX) = ![ctl_id] And arr_varTmp03(C_LTR, lngX) = ![fs_letter] Then
                  blnFound = True
                  Exit For
                End If
              Next
              If blnFound = False Then
                ' ** Then check Form_KeyDown shortcuts.
                For lngX = 0& To (lngProcs - 1&)
                  If Left$(arr_varProc(P_FNAM, lngX), 3) <> "zz_" And Left$(arr_varProc(P_CNAM, lngX), 8) <> "Form_zz_" Then
                    If arr_varProc(P_FID, lngX) = ![frm_id] Then
                      arr_varTmp04 = arr_varProc(P_ARR, lngX)
                      lngTypes = (UBound(arr_varTmp04, 2) + 1&)
                      ' *****************************************************
                      ' ** Array: arr_varType()
                      ' **
                      ' **   Element  Name                     Constant
                      ' **   =======  =======================  ============
                      ' **      0     keydowntype_name         T_TYP_NAM
                      ' **      1     keydowntype_type         T_TYP_TYP
                      ' **      2     Beginning Line Number    T_SEC_BEG
                      ' **      3     Ending Line Number       T_SEC_END
                      ' **      4     Line Indent              T_SEC_IND
                      ' **      5     Count of Keys            T_KEYS
                      ' **      6     arr_varKey() Array       T_ARR
                      ' **
                      ' *****************************************************
                      blnPlainNotFound = False
                      For lngY = 0& To (lngTypes - 1&)
                        ' ** For each record, search only the type matching it!
                        blnFound = False
                        If ![fs_shift] = False And ![fs_alt] = False And ![fs_ctrl] = False And blnPlainNotFound = False Then
                          ' ** Plain.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = False And ![fs_alt] = False And ![fs_ctrl] = True Then
                          ' ** Ctrl.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = False And ![fs_alt] = True And ![fs_ctrl] = False Then
                          ' ** Alt.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Alt" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = True And ![fs_alt] = False And ![fs_ctrl] = False Then
                          ' ** Shift.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Shift" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = False And ![fs_alt] = True And ![fs_ctrl] = True Then
                          ' ** Ctrl-Alt.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl-Alt" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = True And ![fs_alt] = False And ![fs_ctrl] = True Then
                          ' ** Ctrl-Shift.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl-Shift" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = True And ![fs_alt] = True And ![fs_ctrl] = False Then
                          ' ** Alt-Shift.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Alt-Shift" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = True And ![fs_alt] = True And ![fs_ctrl] = True Then
                          ' ** Ctrl-Alt-Shift.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Ctrl-Alt-Shift" Then
                            blnFound = True
                          End If
                        ElseIf ![fs_shift] = False And ![fs_alt] = False And ![fs_ctrl] = False And blnPlainNotFound = True Then
                          ' ** Unknown (currently only vbKeyEscape, Plain).
                          If arr_varTmp04(T_TYP_NAM, lngY) = "{unk}" Then
                            blnFound = True
                          End If
                        End If
                        If blnFound = True Then
                          blnFound = False
                          If arr_varTmp04(T_KEYS, lngY) > 0& Then
                            arr_varTmp05 = arr_varTmp04(T_ARR, lngY)
                            lngKeys = (UBound(arr_varTmp05, 2) + 1&)
                            ' **************************************************
                            ' ** Array: arr_varKey()
                            ' **
                            ' **   Element  Name                  Constant
                            ' **   =======  ====================  ============
                            ' **      0     KeyCode Constant      K_CONST
                            ' **      1     Line Number           K_LINENUM
                            ' **      2     Key Constant Value    K_KEYVAL
                            ' **
                            ' **************************************************
                            For lngZ = 0& To (lngKeys - 1&)
                              If ![fs_letter] = CStr(arr_varTmp05(K_KEYVAL, lngZ)) Then
                                blnFound = True
                                Exit For
                              End If
                            Next  ' ** arr_varKey() (arr_varTmp05): lngZ.
                          End If  ' ** lngKeys > 0.
                          If arr_varTmp04(T_TYP_NAM, lngY) = "Plain" And blnFound = False Then
                            ' ** If this is the 'Plain' loop, and it wasn't found, stick around for the 'Other' loop.
                            blnPlainNotFound = True
                          Else
                            If arr_varTmp04(T_TYP_NAM, lngY) = "{unk}" And blnPlainNotFound = True Then
                              blnPlainNotFound = False
                            End If
                            ' ** Since this type is within the matching-Type 'If' block, always
                            ' ** exit the arr_varType() loop, whether or not it's been found.
                            Exit For
                          End If
                        End If  ' ** Only stops at matching shortcut type.
                      Next  ' ** arr_varType() (arr_varTmp04): lngY.
                      ' ** Since each arr_varProc() only has Form_KeyDown() events, it represents
                      ' ** 1 single, separate form, and should always exit after it's done.
                      Exit For
                    End If  ' ** arr_varProc() matches frm_id.
                  End If  ' ** zz's.
                Next  ' ** arr_varProc(): lngX.
              End If  ' ** Wasn't found among Controls.
              If blnFound = False Then
                lngDels = lngDels + 1&
                lngE = lngDels - 1&
                ReDim Preserve arr_varDel(DEL_ELEMS, lngE)
                arr_varDel(DEL_FID, lngE) = ![frm_id]
                strForm = DLookup("[frm_name]", "tblForm", "[frm_id] = " & CStr(![frm_id]))
                arr_varDel(DEL_FAM, lngE) = strForm
                arr_varDel(DEL_SID, lngE) = ![fs_id]
                arr_varDel(DEL_SNAM, lngE) = ![fs_control] & "  '" & _
                  IIf(IsNumeric(![fs_letter]) = False, ![fs_letter], ![keycode_constant]) & "'" & _
                  IIf(![fs_shift] = False And ![fs_alt] = False And ![fs_ctrl] = False, "  PLAIN", _
                  IIf(![fs_shift] = False And ![fs_alt] = False And ![fs_ctrl] = True, "  CTRL", _
                  IIf(![fs_shift] = False And ![fs_alt] = True And ![fs_ctrl] = False, "  ALT", _
                  IIf(![fs_shift] = True And ![fs_alt] = False And ![fs_ctrl] = False, "  SHIFT", _
                  IIf(![fs_shift] = False And ![fs_alt] = True And ![fs_ctrl] = True, "  CTRL-ALT", _
                  IIf(![fs_shift] = True And ![fs_alt] = False And ![fs_ctrl] = True, "  CTRL-SHIFT", _
                  IIf(![fs_shift] = True And ![fs_alt] = True And ![fs_ctrl] = False, "  ALT-SHIFT", _
                  IIf(![fs_shift] = True And ![fs_alt] = True And ![fs_ctrl] = True, "  CTRL-ALT-SHIFT", _
                  "  UNKNOWN"))))))))
              End If
              If lngW < lngRecs Then .MoveNext
            Next  ' ** For each Shortcut in form_id: lngW.
          End If
        End With  ' ** All shortcuts by form_id: rst2.
      Next  ' ** For each Form: lngV.
    End If  ' ** lngCtrls > 0&.

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

    ' ** Delete obsolete shortcuts.
    If lngDels > 0& Then
      For lngX = 0& To (lngDels - 1&)
        blnDelete = True
        Debug.Print "'DEL SHORTCUT? " & arr_varDel(DEL_SNAM, lngX) & " on " & arr_varDel(DEL_FAM, lngX)
Stop
'blnDelete = False
        If blnDelete = True Then
          ' ** Delete tblForm_Shortcut, by specified [fsid].
          Set qdf = .QueryDefs("zz_qry_Form_Shortcut_01a")
          With qdf.Parameters
            ![fsid] = arr_varDel(DEL_SID, lngX)
          End With
          qdf.Execute dbFailOnError
        End If
      Next
    End If

'Select Case arr_varTmp04(T_TYP_NAM, lngY)
'Case "Plain"
  '
'Case "Ctrl"
  '
'Case "Alt"
  '
'Case "Shift"
  '
'Case "Ctrl-Alt"
  '
'Case "Ctrl-Shift"
  '
'Case "Alt-Shift"
  '
'Case "Ctrl-Alt-Shift"
  '
'Case "{unk}"
  '
'End Select

    .Close
  End With  ' ** dbs.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set prp = Nothing
  Set ctl = Nothing
  Set frm = Nothing
  Set rst1 = Nothing
  Set rst2 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Shortcut_Doc = blnRetValx

End Function

Private Function Frm_Shortcut_Detail_Doc() As Boolean
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_Shortcut_Detail_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim lngKeys As Long, arr_varKey As Variant
  Dim strModName As String, strLine As String
  Dim lngModLines As Long, lngCaseLine As Long, lngNextCaseLine As Long
  Dim blnMoveRec As Boolean, lngMoveRecLine As Long, blnCmdBtn As Boolean, lngCmdBtnLine As Long
  Dim blnChkBox As Boolean, lngChkBoxLine As Long, blnOptGrp As Boolean, lngOptGrpLine As Long
  Dim blnSubCall As Boolean, lngSubCallLine As Long, blnFuncCall As Boolean, lngFuncCallLine As Long
  Dim blnSetFoc As Boolean, lngSetFocLine As Long
  Dim intTrueTot As Integer
  Dim lngThisDbsID
  Dim blnAddAll As Boolean, blnAdd As Boolean
  Dim intPos1 As Integer, intPos2 As Integer
  Dim strTmp01 As String
  Dim lngX As Long, lngY As Long, lngZ As Long

  ' ** Array: arr_varKey().
  Const K_FSID As Integer = 0
  Const K_DID  As Integer = 1
  Const K_FID  As Integer = 2
  Const K_FNAM As Integer = 3
  Const K_KEY  As Integer = 4
  Const K_LTR  As Integer = 5
  Const K_CON  As Integer = 6
  Const K_CAP  As Integer = 7
  Const K_SHFT As Integer = 8
  Const K_ALT  As Integer = 9
  Const K_CTRL As Integer = 10
  Const K_LINE As Integer = 11
  Const K_SORT As Integer = 12
  Const K_TYPE As Integer = 13
  Const K_DATA As Integer = 14

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs
    ' ** tblForm_Shortcut, just Form_KeyDown() shortcuts.
    Set qdf = .QueryDefs("zz_qry_System_90_04")  '"zz_qry_Form_Shortcut_19")
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
      ' **     1       0     fs_id               K_FSID
      ' **     2       1     dbs_id              K_DID
      ' **     3       2     frm_id              K_FID
      ' **     4       3     frm_name            K_FNAM
      ' **     5       4     Key                 K_KEY
      ' **     6       5     fs_letter           K_LTR
      ' **     7       6     keycode_constant    K_CON
      ' **     8       7     fs_caption          K_CAP
      ' **     9       8     fs_shift            K_SHFT
      ' **    10       9     fs_alt              K_ALT
      ' **    11      10     fs_ctrl             K_CTRL
      ' **    12      11     fs_linenum          K_LINE
      ' **    13      12     sort_ltr            K_SORT
      ' **    14      13     action_type         K_TYPE
      ' **    15      14     action_data         K_DATA
      ' **
      ' *****************************************************
      .Close
    End With
    .Close
  End With
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    For lngX = 0& To (lngKeys - 1&)
      Set vbc = .VBComponents("Form_" & arr_varKey(K_FNAM, lngX))
      With vbc
        strModName = .Name
        Set cod = .CodeModule
        With cod
          lngModLines = .CountOfLines
          ' ** arr_varKey(K_LINE, lngX) is line number of Case statement.

          ' ** Find the next Case statement, or the End Select statement.
          lngCaseLine = arr_varKey(K_LINE, lngX)
          lngNextCaseLine = 0&
          strLine = .Lines(lngCaseLine, 1)
          intPos1 = InStr(strLine, "Case")  ' ** Indentation.
          If intPos1 > 0 Then
            For lngY = (lngCaseLine + 1&) To lngModLines
              strLine = .Lines(lngY, 1)
              If Trim$(strLine) <> vbNullString Then  ' ** Not a blank line.
                If Left$(Trim$(strLine), 1) <> "'" Then  ' ** Not a remark.
                  intPos2 = InStr(strLine, "Case")
                  If intPos2 > 0 Then
                    If intPos2 = intPos1 Then  ' ** Same indentation, so this is good.
                      lngNextCaseLine = lngY
                      Exit For
                    End If
                  End If
                  If intPos2 = 0 Then
                    intPos2 = InStr(strLine, "End Select")
                    If intPos2 > 0 Then
                      If intPos2 = intPos1 Then  ' ** Same indentation, so this is good.
                        lngNextCaseLine = lngY
                        Exit For
                      End If
                    End If
                  End If
                End If
              End If
            Next  ' ** lngY.
          Else
            Stop
          End If

          If lngCaseLine > 0& And lngNextCaseLine > 0& Then

            blnMoveRec = False: lngMoveRecLine = 0&: blnCmdBtn = False: lngCmdBtnLine = 0&
            blnChkBox = False: lngChkBoxLine = 0&: blnOptGrp = False: lngOptGrpLine = 0&
            blnSubCall = False: lngSubCallLine = 0&: blnFuncCall = False: lngFuncCallLine = 0&
            blnSetFoc = False: lngSetFocLine = 0&

            ' ** Now try to divine its purpose.
            For lngY = 1& To 7&
              For lngZ = (lngCaseLine + 1&) To (lngNextCaseLine - 1&)
                strLine = .Lines(lngZ, 1)
                intPos1 = 0: intPos2 = 0
                If Trim$(strLine) <> vbNullString Then  ' ** Not a blank line.
                  If Left$(Trim$(strLine), 1) <> "'" Then ' ** Not just a remark.
                    Select Case lngY
                    Case 1&
                      'If there's a MoveRec, save the constant.
                      intPos1 = InStr(strLine, "MoveRec")
                      If intPos1 > 0 Then
                        blnMoveRec = True
                        lngMoveRecLine = lngZ
                      End If
                    Case 2&
                      'If there's a SetFocus, followed by a .._Click, then it's a command button.
                      intPos1 = InStr(strLine, "_Click")
                      If intPos1 > 0 Then
                        blnCmdBtn = True
                        lngCmdBtnLine = lngZ
                      End If
                    Case 3&
                      'If there's a SetFocus, followed by an .._AfterUpdate, then it's a CheckBox or OptionGroup.
                      intPos1 = InStr(strLine, "_AfterUpdate")
                      If intPos1 > 0 Then
                        intPos2 = InStr(strLine, "chk")
                        If intPos2 > 0 Then
                          blnChkBox = True
                          lngChkBoxLine = lngZ
                        End If
                      End If
                    Case 4&
                      'If there's a SetFocus, followed by an .._AfterUpdate, then it's a CheckBox or OptionGroup.
                      intPos1 = InStr(strLine, "_AfterUpdate")
                      If intPos1 > 0 Then
                        intPos2 = InStr(strLine, "opg")
                        If intPos2 > 0 Then
                          blnOptGrp = True
                          lngOptGrpLine = lngZ
                        End If
                      End If
                    Case 5&
                      'If it's got a remark with 'Procedure:' or 'Function:', then it's that.
                      intPos1 = InStr(strLine, "' **")
                      If intPos1 > 0 Then
                        intPos2 = InStr(strLine, "Procedure:")
                        If intPos2 > 0 Then
                          blnSubCall = True
                          lngSubCallLine = lngZ
                        End If
                      End If
                    Case 6&
                      'If it's got a remark with 'Procedure:' or 'Function:', then it's that.
                      intPos1 = InStr(strLine, "' **")
                      If intPos1 > 0 Then
                        intPos2 = InStr(strLine, "Function:")
                        If intPos2 > 0 Then
                          blnFuncCall = True
                          lngFuncCallLine = lngZ
                        End If
                      End If
                    Case 7&
                      'If it's only got a SetFocus, then it's a field.
                      intPos1 = InStr(strLine, ".SetFocus")
                      If intPos1 > 0 Then
                        blnSetFoc = True
                        lngSetFocLine = lngZ  ' ** This will continually update till the final one in the Case.
                      End If
                    End Select
                  End If  ' ** Remark.
                End If  ' ** vbNullString.
              Next  ' ** Case Line: lngZ.
            Next  ' ** Action Type: lngY.

            intTrueTot = blnMoveRec + blnCmdBtn + blnChkBox + blnOptGrp + blnSubCall + blnFuncCall + blnSetFoc
            intTrueTot = Abs(intTrueTot)
            If intTrueTot > 0 Then
              If blnMoveRec = True Then
                strTmp01 = vbNullString
                arr_varKey(K_TYPE, lngX) = "MoveRec"
                strLine = .Lines(lngMoveRecLine, 1)
                intPos1 = InStr(strLine, "acCmdRecordsGoTo")
                If intPos1 > 0 Then
                  strTmp01 = Mid$(strLine, intPos1)
                  intPos2 = InStr(strTmp01, "' **")
                  If intPos2 > 0 Then strTmp01 = Trim$(Left$(strTmp01, (intPos2 - 1)))  ' ** Though this should be 2 True's.
                  arr_varKey(K_DATA, lngX) = strTmp01
                Else
                  ' ** Possibly a specific ID.
                  arr_varKey(K_DATA, lngX) = "Specified"
                End If  ' ** intPos1.
                If intTrueTot = 2 And blnSetFoc = True Then
                  ' ** Let it stand.
                Else
                  ' ** Deal with it later.
                End If
              End If  ' ** blnMoveRec.
              If blnCmdBtn = True Then
                strTmp01 = vbNullString
                If arr_varKey(K_TYPE, lngX) = vbNullString Then
                  arr_varKey(K_TYPE, lngX) = "CmdBtn"
                  strLine = .Lines(lngCmdBtnLine, 1)
                  intPos1 = InStr(strLine, "_Click")
                  If intPos1 > 0 Then
                    For lngY = intPos1 To 1& Step -1&
                      If Mid$(strLine, lngY, 1) = "." Or Mid$(strLine, lngY, 1) = " " Then
                        strTmp01 = Left$(strLine, (intPos1 - 1))
                        strTmp01 = Mid$(strTmp01, (lngY + 1&))
                        arr_varKey(K_DATA, lngX) = strTmp01
                        Exit For
                      End If
                    Next
                  End If  ' ** intPos1.
                  If intTrueTot = 2 And blnSetFoc = True Then
                    ' ** Let it stand.
                  Else
                    ' ** Deal with it later.
                  End If
                Else
                  ' ** CmdBtn and MoveRec?
'Stop
                End If
              End If  ' ** blnCmdBtn.
              If blnChkBox = True Then
                strTmp01 = vbNullString
                If arr_varKey(K_TYPE, lngX) = vbNullString Then
                  arr_varKey(K_TYPE, lngX) = "ChkBox"
                  strLine = .Lines(lngChkBoxLine, 1)
                  intPos1 = InStr(strLine, "_AfterUpdate")
                  If intPos1 > 0 Then
                    For lngY = intPos1 To 1& Step -1&
                      If Mid$(strLine, lngY, 1) = "." Or Mid$(strLine, lngY, 1) = " " Then
                        strTmp01 = Left$(strLine, (intPos1 - 1))
                        strTmp01 = Mid$(strTmp01, (lngY + 1&))
                        arr_varKey(K_DATA, lngX) = strTmp01
                        Exit For
                      End If
                    Next
                  End If  ' ** intPos1.
                Else
                  ' ** Some weird mix!
'Stop
                End If
              End If  ' ** blnChkBox.
              If blnOptGrp = True Then
                strTmp01 = vbNullString
                If arr_varKey(K_TYPE, lngX) = vbNullString Then
                  arr_varKey(K_TYPE, lngX) = "OptGrp"
                  strLine = .Lines(lngOptGrpLine, 1)
                  intPos1 = InStr(strLine, "_AfterUpdate")
                  If intPos1 > 0 Then
                    For lngY = intPos1 To 1& Step -1&
                      If Mid$(strLine, lngY, 1) = "." Or Mid$(strLine, lngY, 1) = " " Then
                        strTmp01 = Left$(strLine, (intPos1 - 1))
                        strTmp01 = Mid$(strTmp01, (lngY + 1&))
                        arr_varKey(K_DATA, lngX) = strTmp01
                        Exit For
                      End If
                    Next
'move up to find an '.OptionValue' line!
                    strTmp01 = vbNullString
                    For lngY = lngOptGrpLine To arr_varKey(K_LINE, lngX) Step -1&
                      strLine = .Lines(lngY, 1)
                      If Trim$(strLine) <> vbNullString Then
                        If Left$(Trim$(strLine), 1) <> "'" Then
                          intPos1 = InStr(strLine, ".OptionValue")
                          If intPos1 > 0 Then
                            strTmp01 = Left$(strLine, (intPos1 - 1))
                            For lngZ = Len(strTmp01) To 1 Step -1&
                              If Mid$(strTmp01, lngZ, 1) = "." Or Mid$(strTmp01, lngZ, 1) = " " Then
                                strTmp01 = Mid$(strTmp01, (lngZ + 1))
                                Exit For
                              End If
                            Next
                            If strTmp01 <> vbNullString Then
                              arr_varKey(K_DATA, lngX) = strTmp01
                            End If
                            Exit For
                          End If
                        End If
                      End If
                    Next
                  End If  ' ** intPos1.
                Else
                  ' ** Some weird mix!
'Stop
                End If
              End If  ' ** blnOptGrp.
              If blnSubCall = True Then
                strTmp01 = vbNullString
                strLine = .Lines(lngSubCallLine, 1)
                intPos1 = InStr(strLine, "' **")
                If intPos1 > 0 Then
                  strTmp01 = Trim$(Left$(strLine, (intPos1 - 1)))
                  For lngY = Len(strTmp01) To 1& Step -1&
                    If Mid$(strTmp01, lngY, 1) = "." Or Mid$(strTmp01, lngY, 1) = " " Then
                      strTmp01 = Mid$(strTmp01, (lngY + 1&))
                      Exit For
                    End If
                  Next
                End If ' ** intPos1.
                If strTmp01 <> vbNullString Then
                  If arr_varKey(K_TYPE, lngX) = vbNullString Then
                    arr_varKey(K_TYPE, lngX) = "SubCall"
                    arr_varKey(K_DATA, lngX) = strTmp01
                    ' ** Exceptions.
                    If InStr(strLine, "ChangeAcct") > 0 Then
                      arr_varKey(K_TYPE, lngX) = "MoveRec"
                    End If
                  Else
                    ' ** Figure out which takes precedence.
'Stop
                  End If
                End If  ' ** vbNullString.
              End If  ' ** blnSubCall.
              If blnFuncCall = True Then
                strTmp01 = vbNullString
                strLine = .Lines(lngFuncCallLine, 1)
                intPos1 = InStr(strLine, "' **")
                If intPos1 > 0 Then
                  strTmp01 = Trim$(Left$(strLine, (intPos1 - 1)))
                  For lngY = Len(strTmp01) To 1& Step -1&
                    If Mid$(strTmp01, lngY, 1) = "." Or Mid$(strTmp01, lngY, 1) = " " Then
                      strTmp01 = Mid$(strTmp01, (lngY + 1&))
                      Exit For
                    End If
                  Next
                End If ' ** intPos1.
                If strTmp01 <> vbNullString Then
                  If arr_varKey(K_TYPE, lngX) = vbNullString Then
                    arr_varKey(K_TYPE, lngX) = "FuncCall"
                    arr_varKey(K_DATA, lngX) = strTmp01
                    ' ** Exceptions.
                    If InStr(strLine, "frmMenu_Background") > 0 Then
                      arr_varKey(K_TYPE, lngX) = "CmdBtn"
                      arr_varKey(K_DATA, lngX) = "Close Background"
                    ElseIf strTmp01 = "SetOption_Dev" Then
                      If arr_varKey(K_CON, lngX) = "vbKeyX" And arr_varKey(K_SHFT, lngX) = True And arr_varKey(K_CTRL, lngX) = True Then
                        arr_varKey(K_TYPE, lngX) = "CmdBtn"
                        arr_varKey(K_DATA, lngX) = "Close Form"
                      End If
                    ElseIf arr_varKey(K_CON, lngX) = "vbKeyEscape" Then
                      arr_varKey(K_TYPE, lngX) = "CmdBtn"
                      arr_varKey(K_DATA, lngX) = "Cancel Form - Esc"
                    End If
                  Else
                    ' ** Figure out which takes precedence.
'Stop
                  End If
                End If  ' ** vbNullString.
              End If  ' ** blnFuncCall.
              If blnSetFoc = True Then
                strTmp01 = vbNullString
                strLine = .Lines(lngSetFocLine, 1)
                intPos1 = InStr(strLine, ".SetFocus")
                If intPos1 > 0 Then
                  strTmp01 = Left$(strLine, (intPos1 - 1))
                  For lngY = Len(strTmp01) To 1 Step -1&
                    If Mid$(strTmp01, lngY, 1) = "." Or Mid$(strTmp01, lngY, 1) = " " Then
                      strTmp01 = Mid$(strTmp01, (lngY + 1&))
                      Exit For
                    End If
                  Next
                  If intTrueTot = 1 Then
                    arr_varKey(K_TYPE, lngX) = "SetFoc"
                    arr_varKey(K_DATA, lngX) = strTmp01
                  Else
                    ' ** Exceptions.
                    If arr_varKey(K_TYPE, lngX) = "SubCall" Then
                      strLine = .Lines(lngSubCallLine, 1)
                      If InStr(strLine, "posted_AfterUpdate") > 0 Then
                        arr_varKey(K_TYPE, lngX) = "ChkBox"
                        arr_varKey(K_DATA, lngX) = "posted"
                      End If
                    End If
                    If arr_varKey(K_TYPE, lngX) = "FuncCall" Then
                      strLine = .Lines(lngFuncCallLine, 1)
                      If InStr(strLine, "fSetScrollBarPosVT") > 0 Or InStr(strLine, "fSetScrollBarPosHZ") > 0 Or _
                          InStr(strLine, "JC_Key_Par_Next") > 0 Then
                        arr_varKey(K_TYPE, lngX) = "SetFoc"
                        arr_varKey(K_DATA, lngX) = strTmp01
                      ElseIf InStr(strLine, "frmCheckReconcile") > 0 And InStr(strLine, "_Sub_") > 0 Then
                        arr_varKey(K_TYPE, lngX) = "SetFoc"
                        arr_varKey(K_DATA, lngX) = strTmp01
                      ElseIf arr_varKey(K_DATA, lngX) = "RecCnt" Then
                        arr_varKey(K_TYPE, lngX) = "SetFoc"
                        arr_varKey(K_DATA, lngX) = strTmp01
                      End If
                    End If
                    If arr_varKey(K_TYPE, lngX) = "MoveRec" Then
                      strLine = .Lines(lngMoveRecLine, 1)
                      If InStr(strLine, "acCmdRecordsGoTo") = 0 Then
                        arr_varKey(K_TYPE, lngX) = "SetFoc"
                        arr_varKey(K_DATA, lngX) = strTmp01
                      ElseIf InStr(strLine, "frmFeeSchedules_Detail_Sub") > 0 And _
                          (arr_varKey(K_CON, lngX) = "vbKeyTab" Or arr_varKey(K_CON, lngX) = "vbKeyReturn") Then
                        arr_varKey(K_TYPE, lngX) = "SetFoc"
                        arr_varKey(K_DATA, lngX) = "frmFeeSchedules_Detail_Sub"
                      End If
                    End If
                  End If
                End If  ' ** intPos1.
              End If  ' ** blnSetFoc.
            End If  ' ** intTrueTot.

          End If  ' ** lngCaseLine, lngNextCaseLine.

        End With  ' ** cod.
      End With  ' ** vbc
    Next  ' ** lngX.

  End With  ' ** vbp.

  Set dbs = CurrentDb
  With dbs
    Set rst = .OpenRecordset("tblForm_Shortcut_Detail", dbOpenDynaset, dbConsistent)
    With rst
      blnAddAll = False
      If .BOF = True And .EOF = True Then
        blnAddAll = True
      Else
        .MoveLast
        .MoveFirst
      End If
      For lngX = 0& To (lngKeys - 1&)
        If arr_varKey(K_TYPE, lngX) <> vbNullString Then
          blnAdd = False
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            .MoveFirst
            .FindFirst "[fs_id] = " & CStr(arr_varKey(K_FSID, lngX))
            If .NoMatch = True Then
              .MoveLast
              .FindFirst "[fs_id] = " & CStr(arr_varKey(K_FSID, lngX))
              Select Case .NoMatch
              Case True
                blnAdd = True
              Case False
                If ![dbs_id] <> lngThisDbsID Then
                  Stop
                Else
                  If ![frm_id] <> arr_varKey(K_FID, lngX) Then
                    .FindNext "[fs_id] = " & CStr(arr_varKey(K_FSID, lngX))
                    ' ** I don't really know if there'll ever be more than 1 record.
                    Select Case .NoMatch
                    Case True
                      blnAdd = True
                    Case False
                      ' ** Edit it.
                    End Select
                  Else
                    ' ** Edit it.
                  End If
                End If
              End Select
            End If
          End Select
          If blnAdd = True Then
            .AddNew
            ![dbs_id] = lngThisDbsID
            ![frm_id] = arr_varKey(K_FID, lngX)
            ![fs_id] = arr_varKey(K_FSID, lngX)
            ' ** ![fsd_id] : AutoNumber.
            ![fsd_action_type] = arr_varKey(K_TYPE, lngX)
            If arr_varKey(K_DATA, lngX) <> vbNullString Then
              ![fsd_action_data] = arr_varKey(K_DATA, lngX)
            End If
            ![fsd_datemodified] = Now()
            .Update
          Else
            If ![fsd_action_type] <> arr_varKey(K_TYPE, lngX) Then
              .Edit
              ![fsd_action_type] = arr_varKey(K_TYPE, lngX)
              ![fsd_datemodified] = Now()
              .Update
            End If
            If arr_varKey(K_DATA, lngX) <> vbNullString Then
              If ![fsd_action_data] <> arr_varKey(K_DATA, lngX) Then
                .Edit
                ![fsd_action_data] = arr_varKey(K_DATA, lngX)
                ![fsd_datemodified] = Now()
                .Update
              End If
            End If
          End If
        End If
      Next ' ** lngX.
      .Close
    End With  ' ** rst.
    .Close
  End With  ' ** dbs.

  Beep
  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Shortcut_Detail_Doc = blnRetValx

End Function

Private Function Frm_Subform_Doc() As Boolean
' ** Document all form subforms to tblForm_Subform.
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_Subform_Doc"

  Dim dbs As DAO.Database, cntr As Container, doc As Document, frm As Form, ctl As Control
  Dim qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset
  Dim lngSubs As Long, arr_varSub() As Variant
  Dim lngFrmID As Long, lngCtlID As Long, lngThisDbsID As Long
  Dim strSource As String
  Dim lngRecs As Long
  Dim blnSkip As Boolean
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varSub().
  Const S_ELEMS As Integer = 9  ' ** Array's first-element UBound().
  Const S_DID     As Integer = 0
  Const S_DNAM    As Integer = 1
  Const S_FID     As Integer = 2
  Const S_FNAM    As Integer = 3
  Const S_CID     As Integer = 4
  Const S_CNAM    As Integer = 5
  Const S_SRCID   As Integer = 6
  Const S_SRCNAM  As Integer = 7
  Const S_LNKMAST As Integer = 8
  Const S_LNKCHLD As Integer = 9

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    lngSubs = 0&
    ReDim arr_varSub(S_ELEMS, 0)

    Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbConsistent)
    If rst.BOF = True And rst.EOF = True Then
      ' ** No records?! Run Frm_Doc(), above.
      Beep
      blnRetValx = False
    End If

    If blnRetValx = True Then
      Set cntr = .Containers("Forms")
      With cntr
        For Each doc In .Documents
          lngFrmID = 0&: strSource = vbNullString: lngSubs = 0&
          With doc
            ' ************************************************************
            ' ** Array: arr_varSub()
            ' **
            ' **   Field  Element  Name                       Constant
            ' **   =====  =======  =========================  ==========
            ' **     1       0     dbs_id                     S_DID
            ' **     2       1     dbs_name                   S_DNAM
            ' **     3       2     frm_id                     S_FID
            ' **     4       3     frm_name                   S_FNAM
            ' **     5       4     ctl_id                     S_CID
            ' **     6       5     ctl_name                   S_CNAM
            ' **     7       6     frm_id_sub                 S_SRCID
            ' **     8       7     frm_name2                  S_SRCNAM
            ' **     9       8     frmsub_linkmasterfields    S_LNKMAST
            ' **    10       9     frmsub_linkchildfields     S_LNKCHLD
            ' **
            ' ************************************************************
            DoCmd.OpenForm .Name, acDesign, , , , acHidden
            Set frm = Forms(.Name)
            With frm
              For Each ctl In .Controls
                With ctl
                  Select Case .ControlType
                  Case acSubform
                    With rst
                      .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_name] = '" & doc.Name & "'"
                      If .NoMatch = False Then
                        .Edit
                        ![frm_hassub] = True
                        lngSubs = lngSubs + 1&
                        lngE = lngSubs - 1&
                        ReDim Preserve arr_varSub(S_ELEMS, lngE)
                        arr_varSub(S_DID, lngE) = ![dbs_id]
                        arr_varSub(S_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
                        arr_varSub(S_FID, lngE) = ![frm_id]
                        arr_varSub(S_FNAM, lngE) = frm.Name
                        lngCtlID = DLookup("[ctl_id]", "tblForm_Control", "[frm_id] = " & CStr(![frm_id]) & " And [ctl_name] = '" & ctl.Name & "'")
                        arr_varSub(S_CID, lngE) = lngCtlID
                        arr_varSub(S_CNAM, lngE) = ctl.Name
                        If ctl.SourceObject = vbNullString Then
                          Debug.Print "'EMPTY SUB: " & frm.Name & "  SUB: " & ctl.Name
                          arr_varSub(S_SRCID, lngE) = 0&
                          arr_varSub(S_SRCNAM, lngE) = "{null}"
                          arr_varSub(S_LNKMAST, lngE) = Null
                          arr_varSub(S_LNKCHLD, lngE) = Null
                        Else
                          strSource = ctl.SourceObject
                          lngFrmID = DLookup("[frm_id]", "tblForm", "[frm_name] = '" & strSource & "'")
                          arr_varSub(S_SRCID, lngE) = lngFrmID
                          arr_varSub(S_SRCNAM, lngE) = ctl.SourceObject
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
            End With  ' ** This Form: frm.
            DoCmd.Close acForm, .Name, acSaveNo
            If lngSubs > 0& Then
              For lngX = 0& To (lngSubs - 1&)
                If arr_varSub(S_SRCNAM, lngX) <> "{null}" Then
                  With rst
                    .FindFirst "[dbs_id] = " & CStr(arr_varSub(S_DID, lngX)) & " And [frm_id] = " & CStr(arr_varSub(S_SRCID, lngX))
                    If .NoMatch = False Then
                      .Edit
                      ![frm_issub] = True
                      ![frm_datemodified] = Now()
                      .Update
                    Else
                      Stop
                    End If
                  End With
                End If
              Next
              With rst
                .FindFirst "[frm_id] = " & CStr(arr_varSub(S_FID, 0))  ' ** All these subforms are on the same form.
                If .NoMatch = False Then
                  .Edit
                  ![frm_subs] = lngSubs
                  ![frm_datemodified] = Now()
                  .Update
                Else
                  Stop
                End If
              End With
            End If
          End With  ' ** This Document: doc.
        Next  ' ** For each Document: doc.
      End With  ' ** This Container: cntr.
    End If
    rst.Close

    If blnRetValx = True And lngSubs > 0& Then
      Set rst = .OpenRecordset("tblForm_Subform", dbOpenDynaset, dbConsistent)
      With rst
        For lngX = 0& To (lngSubs - 1&)
          If arr_varSub(S_SRCID, lngX) > 0& Then
            .FindFirst "[dbs_id] = " & CStr(arr_varSub(S_DID, lngX)) & " And [frm_id] = " & CStr(arr_varSub(S_FID, lngX)) & _
              " And [ctl_id] = " & CStr(arr_varSub(S_CID, lngX)) & _
              " And [frm_id_sub] = " & CStr(arr_varSub(S_SRCID, lngX))
            If .NoMatch = False Then
              If IsNull(arr_varSub(S_LNKMAST, lngX)) = False Then
                If IsNull(![frmsub_linkmasterfields]) = False Then
                  If arr_varSub(S_LNKMAST, lngX) <> ![frmsub_linkmasterfields] Then
                    If InStr(![frmsub_linkmasterfields], "{or}") = 0 Then
                      .Edit
                      ![frmsub_linkmasterfields] = arr_varSub(S_LNKMAST, lngX)
                      ![frmsub_datemodified] = Now()
                      .Update
                    Else
                      Debug.Print "'OR SRC: '" & arr_varSub(S_FNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'  SRC: '" & _
                        arr_varSub(S_LNKMAST, lngX) & "'"
                    End If
                  End If
                Else
                  .Edit
                  ![frmsub_linkmasterfields] = arr_varSub(S_LNKMAST, lngX)
                  ![frmsub_datemodified] = Now()
                  .Update
                End If
              Else
                If IsNull(![frmsub_linkmasterfields]) = False Then
                  Debug.Print "'FRM NULL, TBL NOT: '" & arr_varSub(S_FNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'"
                End If
              End If
              If IsNull(arr_varSub(S_LNKCHLD, lngX)) = False Then
                If IsNull(![frmsub_linkchildfields]) = False Then
                  If arr_varSub(S_LNKCHLD, lngX) <> ![frmsub_linkchildfields] Then
                    If InStr(![frmsub_linkchildfields], "{or}") = 0 Then
                      .Edit
                      ![frmsub_linkchildfields] = arr_varSub(S_LNKCHLD, lngX)
                      ![frmsub_datemodified] = Now()
                      .Update
                    Else
                      Debug.Print "'OR SRC: '" & arr_varSub(S_FNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'  SRC: '" & _
                        arr_varSub(S_LNKCHLD, lngX) & "'"
                    End If
                  End If
                Else
                  .Edit
                  ![frmsub_linkchildfields] = arr_varSub(S_LNKCHLD, lngX)
                  ![frmsub_datemodified] = Now()
                  .Update
                End If
              Else
                If IsNull(![frmsub_linkchildfields]) = False Then
                  Debug.Print "'FRM NULL, TBL NOT: '" & arr_varSub(S_FNAM, lngX) & "'  SUB: '" & arr_varSub(S_CNAM, lngX) & "'"
                End If
              End If
            Else
              .AddNew
              ![dbs_id] = arr_varSub(S_DID, lngX)
              ![frm_id] = arr_varSub(S_FID, lngX)
              ![ctl_id] = arr_varSub(S_CID, lngX)
              ![frm_id_sub] = arr_varSub(S_SRCID, lngX)
              If IsNull(arr_varSub(S_LNKMAST, lngX)) = False Then
                ![frmsub_linkmasterfields] = arr_varSub(S_LNKMAST, lngX)
                ![frmsub_linkchildfields] = arr_varSub(S_LNKCHLD, lngX)
              End If
              ![frmsub_datemodified] = Now()
              .Update
            End If
          End If
        Next
        .Close
      End With
    End If

blnSkip = True
If blnSkip = False Then
    ' ** tblForm, linked to tblForm_Subform, just frmCalendar, frmReinvest_Dividend, frmReinvest_Interest.
    Set qdf1 = .QueryDefs("zz_qry_Form_Subform_01")
    Set rst = qdf1.OpenRecordset
    With rst
      If .BOF = True And .EOF = True Then
        Stop
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        If lngRecs <> 3& Then
          Stop
        Else
          For lngX = 1& To 3&
            Select Case lngX
            Case 1&
              ' ** First record is frmCalendar.
              Select Case IsNull(![frm_tag])
              Case True
                Stop
              Case False
                If InStr(![frm_tag], "Is Subform") = 0 Then
                  Stop
                End If
              End Select
              Select Case IsNull(![frm_parent_sub])
              Case True
                .Edit
                ![frm_parent_sub] = "frmReinvest_Dividend;frmReinvest_Interest"
                ![frm_datemodified] = Now()
                .Update
              Case False
                If InStr(![frm_parent_sub], "frmReinvest_Dividend") = 0 Then
                  .Edit
                  ![frm_parent_sub] = ![frm_parent_sub] & ";frmReinvest_Dividend"
                  ![frm_datemodified] = Now()
                  .Update
                End If
                If InStr(![frm_parent_sub], "frmReinvest_Interest") = 0 Then
                  .Edit
                  ![frm_parent_sub] = ![frm_parent_sub] & ";frmReinvest_Interest"
                  ![frm_datemodified] = Now()
                  .Update
                End If
              End Select
            Case 2&
              ' ** Then comes frmReinvest_Dividend.
              Select Case IsNull(![frm_tag])
              Case True
                Stop
              Case False
                If InStr(![frm_tag], "Has Subform") = 0 Then
                  Stop
                End If
              End Select
              Select Case IsNull(![frm_parent_sub])
              Case True
                .Edit
                ![frm_parent_sub] = ![frm_parent_sub] & "frmCalendar;frmCalendar"  ' ** It uses it twice.
                ![frm_datemodified] = Now()
                .Update
              Case False
                If InStr(![frm_parent_sub], "frmCalendar") = 0 Then
                  .Edit
                  ![frm_parent_sub] = ![frm_parent_sub] & "frmCalendar;frmCalendar"
                  ![frm_datemodified] = Now()
                  .Update
                End If
              End Select
            Case 3&
              ' ** And finally frmReinvest_Interest.
              Select Case IsNull(![frm_tag])
              Case True
                Stop
              Case False
                If InStr(![frm_tag], "Has Subform") = 0 Then
                  Stop
                End If
              End Select
              Select Case IsNull(![frm_parent_sub])
              Case True
                .Edit
                ![frm_parent_sub] = ![frm_parent_sub] & "frmCalendar;frmCalendar"  ' ** It uses it twice.
                ![frm_datemodified] = Now()
                .Update
              Case False
                If InStr(![frm_parent_sub], "frmCalendar") = 0 Then
                  .Edit
                  ![frm_parent_sub] = ![frm_parent_sub] & "frmCalendar;frmCalendar"
                  ![frm_datemodified] = Now()
                  .Update
                End If
              End Select
            End Select
            If lngX < lngRecs Then .MoveNext
          Next  ' ** lngX.
          .MoveFirst
          .MoveNext
          ' ** This should be frmReinvest_Dividend
          If IsNull(![frmsub_id_par]) = True Then
            ' ** Append zz_qry_Form_Subform_03 (tblForm, linked to tblForm_Control, zz_qry_Form_Subform_02
            ' ** (tblForm, just frmCalendar), just frmReinvest_Dividend subforms) to tblForm_Subform.
            Set qdf2 = dbs.QueryDefs("zz_qry_Form_Subform_04")
            qdf2.Execute
            Set qdf2 = Nothing
          End If
          .MoveNext
          ' ** This should be frmReinvest_Interest.
          If IsNull(![frmsub_id_par]) = True Then
            ' ** Append zz_qry_Form_Subform_05 (tblForm, linked to tblForm_Control, zz_qry_Form_Subform_02
            ' ** (tblForm, just frmCalendar), just frmReinvest_Interest subforms) to tblForm_Subform.
            Set qdf2 = dbs.QueryDefs("zz_qry_Form_Subform_06")
            qdf2.Execute
            Set qdf2 = Nothing
          End If
        End If
      End If
      .Close
    End With
End If  ' ** blnSkip.
    Set rst = Nothing
    Set qdf1 = Nothing

    .Close
  End With  ' ** dbs.

  ' ** AcControlType enumeration:
  ' **   100 acLabel            Label control.
  ' **   101 acRectangle        Rectangle control.
  ' **   102 acLine             Line control.
  ' **   103 acImage            Image control.
  ' **   104 acCommandButton    Command Button control.
  ' **   105 acOptionButton     Option Button control.
  ' **   106 acCheckBox         Check Box control.
  ' **   107 acOptionGroup      Option Group control.
  ' **   108 acBoundObjectFrame Bound Object Frame control.
  ' **   109 acTextBox          Text Box control.
  ' **   110 acListBox          List Box control.
  ' **   111 acComboBox         Combo Box control.
  ' **   112 acSubform          Subform control.
  ' **   114 acObjectFrame      Unbound Object Frame control.
  ' **   118 acPageBreak        Page Break control.
  ' **   119 acCustomControl    ActiveX control.
  ' **   122 acToggleButton     Toggle Button control.
  ' **   123 acTabCtl           Tab control.
  ' **   124 acPage             Page control, Tab control page.
  ' **   126 acAttachment       Attachment control.  (Access 2007)

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set doc = Nothing
  Set cntr = Nothing
  Set rst = Nothing
  Set qdf1 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  Frm_Subform_Doc = blnRetValx

End Function

Private Function Frm_Ctl_RowSource_Doc() As Boolean
' ** Document all form combo box and list box Row Sources to tblForm_Control_RowSource.
' ** Called by:
' **   QuikFrmDoc(), Above
'NO DELETE ROUTINE FOR CONTROLS THAT USED TO BE COMBO/LIST, NOW AREN'T!!

  Const THIS_PROC As String = "Frm_Ctl_RowSource_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim frm As Form, ctl As Control, prp As Object
  Dim lngRowSrcs As Long, arr_varRowSrc As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim strForm As String, strLastForm As String
  Dim strFormRef As String
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim blnAdd As Boolean, blnFound As Boolean
  Dim intPos1 As Integer
  Dim varTmp00 As Variant
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varRowSrc().
  Const ROW_DID  As Integer = 0
  Const ROW_DNAM As Integer = 1
  Const ROW_FID  As Integer = 2
  Const ROW_FNAM As Integer = 3
  Const ROW_CID  As Integer = 4
  Const ROW_CNAM As Integer = 5
  Const ROW_CTYP As Integer = 6
  Const ROW_CNST As Integer = 7
  Const ROW_RID  As Integer = 8
  Const ROW_FND  As Integer = 9

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set dbs = CurrentDb
  With dbs

    ' ** Get a list of all the form combo boxes and list boxes.
    ' ** tblForm_Control, just acComboBox, acListBox, by specified CurrentAppName().
    Set qdf = .QueryDefs("zz_qry_Form_Control_RowSource_02")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRowSrcs = .RecordCount
      .MoveFirst
      arr_varRowSrc = .GetRows(lngRowSrcs)
      ' *****************************************************
      ' ** Array: arr_varRowSrc()
      ' **
      ' **   Field  Element  Name                Constant
      ' **   =====  =======  ==================  ==========
      ' **     1       0     dbs_id              ROW_DID
      ' **     2       1     dbs_name            ROW_DNAM
      ' **     3       2     frm_id              ROW_FID
      ' **     4       3     frm_name            ROW_FNAM
      ' **     5       4     ctl_id              ROW_CID
      ' **     6       5     ctl_name            ROW_CNAM
      ' **     7       6     ctltype_type        ROW_CTYP
      ' **     8       7     ctltype_constant    ROW_CNST
      ' **     9       8     rowsrc_id           ROW_RID
      ' **    10       9     rowsrc_del          ROW_FND
      ' **
      ' *****************************************************
      .Close
    End With

    Set rst = .OpenRecordset("tblForm_Control_RowSource", dbOpenDynaset, dbConsistent)

    ' ** Open each form and get the control's properties; query is in form order.
    strForm = vbNullString: strLastForm = vbNullString
    For lngX = 0& To (lngRowSrcs - 1&)
      strForm = arr_varRowSrc(ROW_FNAM, lngX)
      If strForm <> strLastForm Then
        If strLastForm <> vbNullString Then
          DoCmd.Close acForm, strLastForm, acSaveNo
        End If
        DoCmd.OpenForm strForm, acDesign, , , , acHidden
        Set frm = Forms(strForm)
        strLastForm = strForm
      End If
      With frm
        Set ctl = .Controls(arr_varRowSrc(ROW_CNAM, lngX))
        If ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
          With rst
            blnAdd = False
            If .BOF = True And .EOF = True Then
              blnAdd = True
            Else
'rowsrc_id
              .FindFirst "[dbs_id] = " & CStr(arr_varRowSrc(ROW_DID, lngX)) & " And [ctl_id] = " & CStr(arr_varRowSrc(ROW_CID, lngX))
              If .NoMatch = True Then
                blnAdd = True
              End If
            End If
            If blnAdd = True Then
              .AddNew
'dbs_id
              ![dbs_id] = arr_varRowSrc(ROW_DID, lngX)
'frm_id
              ![frm_id] = arr_varRowSrc(ROW_FID, lngX)
'ctl_id
              ![ctl_id] = arr_varRowSrc(ROW_CID, lngX)
              If ctl.RowSourceType <> vbNullString Then
'rowsrctype_type
                ![rowsrctype_type] = ctl.RowSourceType
                If ctl.RowSourceType = "Table/Query" Then
                  If ctl.RowSource = vbNullString Then
'qrytbltype_type
                    ![qrytbltype_type] = acNothing
'rowsrc_rowsource
                    ![rowsrc_rowsource] = "{empty}"
                  Else
                    If Left$(ctl.RowSource, 3) = "qry" Then
                      ![qrytbltype_type] = acQuery
                    Else
                      If InStr(ctl.RowSource, "SELECT") > 0 Then
                        ![qrytbltype_type] = acSQL
                      Else
                        ![qrytbltype_type] = acTable
                      End If
                    End If
                    ![rowsrc_rowsource] = ctl.RowSource
                  End If
                Else
                  ![qrytbltype_type] = acNothing
                  If ctl.RowSource = vbNullString Then
                    ![rowsrc_rowsource] = "{empty}"
                  Else
                    ![rowsrc_rowsource] = ctl.RowSource
                  End If
                End If
                If ![rowsrc_rowsource] <> "{empty}" Then
'rowsrc_boundcolumn
                  ![rowsrc_boundcolumn] = ctl.BoundColumn
'rowsrc_columncount
                  ![rowsrc_columncount] = ctl.ColumnCount
'rowsrc_columnheads
                  ![rowsrc_columnheads] = ctl.ColumnHeads
                  If ctl.ColumnWidths <> vbNullString Then
'rowsrc_columnwidths
                    ![rowsrc_columnwidths] = ctl.ColumnWidths
                  End If
                  If ctl.ControlType = acComboBox Then
'rowsrc_listwidth
                    ![rowsrc_listwidth] = ctl.ListWidth
                  End If
                  If ctl.ControlType = acComboBox Then
'rowsrc_limittolist
                    ![rowsrc_limittolist] = ctl.LimitToList
                  Else
                    ![rowsrc_limittolist] = False
                  End If
                  If ctl.ControlType = acComboBox Then
'rowsrc_listrows
                    ![rowsrc_listrows] = ctl.ListRows
                  End If
                  If ctl.ControlType = acComboBox Then
'rowsrc_autoexpand
                    ![rowsrc_autoexpand] = ctl.AutoExpand
                  Else
                    ![rowsrc_autoexpand] = False
                  End If
                  If ctl.ControlType = acListBox Then
'rowsrc_multiselect
                    ![rowsrc_multiselect] = ctl.MultiSelect
                  Else
                    ![rowsrc_multiselect] = Null
                  End If
                  If ![qrytbltype_type] = acSQL Then
                    If InStr(ctl.RowSource, "forms") > 0 Or InStr(ctl.RowSource, "reports") > 0 Then
'rowsrc_hasformref
                      ![rowsrc_hasformref] = True
                      intPos1 = InStr(ctl.RowSource, "forms")
                      If intPos1 > 0 Then
                        If intPos1 > 1 Then
                          If Mid$(ctl.RowSource, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
                        End If
                        strFormRef = Mid$(ctl.RowSource, intPos1)
                        intPos1 = InStr(strFormRef, " ")
                        If intPos1 > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
                        strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
'rowsrc_formref
                        ![rowsrc_formref] = strFormRef
                      End If
                      intPos1 = InStr(ctl.RowSource, "reports")
                      If intPos1 > 0 Then
                        If intPos1 > 1 Then
                          If Mid$(ctl.RowSource, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
                        End If
                        strFormRef = Mid$(ctl.RowSource, intPos1)
                        intPos1 = InStr(strFormRef, " ")
                        If intPos1 > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
                        strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
                        ![rowsrc_formref] = strFormRef
                      End If
                    End If
                  End If
                End If
              Else
                ![rowsrctype_type] = "Table/Query"
                ![qrytbltype_type] = acNothing
                ![rowsrc_rowsource] = "{empty}"
              End If
'rowsrc_datemodified
              ![rowsrc_datemodified] = Now()
              .Update
              .Bookmark = .LastModified
              arr_varRowSrc(ROW_RID, lngX) = ![rowsrc_id]
              arr_varRowSrc(ROW_FND, lngX) = CBool(True)
            Else  ' ** Update.
              arr_varRowSrc(ROW_RID, lngX) = ![rowsrc_id]
              arr_varRowSrc(ROW_FND, lngX) = CBool(True)
              If ctl.RowSourceType <> vbNullString Then
                If IsNull(![rowsrctype_type]) = True Then
                  .Edit
'rowsrctype_type
                  ![rowsrctype_type] = ctl.RowSourceType
                  ![rowsrc_datemodified] = Now()
                  .Update
                Else
                  If ![rowsrctype_type] <> ctl.RowSourceType Then
                    .Edit
                    ![rowsrctype_type] = ctl.RowSourceType
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
              Else
                If IsNull(![rowsrctype_type]) = True Then
                  .Edit
                  ![rowsrctype_type] = "Table/Query"
                  ![qrytbltype_type] = acNothing
                  ![rowsrc_datemodified] = Now()
                  .Update
                Else
                  If ![rowsrctype_type] <> "Table/Query" Then
                    .Edit
                    ![rowsrctype_type] = "Table/Query"
                    ![qrytbltype_type] = acNothing
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
              End If
              If ctl.RowSource <> vbNullString Then
                If IsNull(![rowsrc_rowsource]) = True Then
                  .Edit
'rowsrc_rowsource
                  ![rowsrc_rowsource] = ctl.RowSource
                  ![rowsrc_datemodified] = Now()
                  .Update
                Else
                  If ![rowsrc_rowsource] <> ctl.RowSource Then
                    .Edit
                    ![rowsrc_rowsource] = ctl.RowSource
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
              Else
                If IsNull(![rowsrc_rowsource]) = True Then
                  .Edit
                  ![rowsrc_rowsource] = "{empty}"
                  ![rowsrc_datemodified] = Now()
                  .Update
                Else
                  If ![rowsrc_rowsource] <> "{empty}" Then
                    .Edit
                    ![rowsrc_rowsource] = "{empty}"
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
              End If
              If ctl.RowSource <> vbNullString Then
                If Left$(ctl.RowSource, 3) = "qry" Then
                  If IsNull(![qrytbltype_type]) = False Then
                    If ![qrytbltype_type] <> acQuery Then
                      .Edit
'qrytbltype_type
                      ![qrytbltype_type] = acQuery
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                  Else
                    .Edit
                    ![qrytbltype_type] = acQuery
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                Else
                  If InStr(ctl.RowSource, "SELECT") > 0 Then
                    If IsNull(![qrytbltype_type]) = False Then
                      If ![qrytbltype_type] <> acSQL Then
                        .Edit
                        ![qrytbltype_type] = acSQL
                        ![rowsrc_datemodified] = Now()
                        .Update
                      End If
                    Else
                      .Edit
                      ![qrytbltype_type] = acSQL
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                  Else
                    If IsNull(![qrytbltype_type]) = False Then
                      If ![qrytbltype_type] <> acTable Then
                        .Edit
                        ![qrytbltype_type] = acTable
                        ![rowsrc_datemodified] = Now()
                        .Update
                      End If
                    Else
                      .Edit
                      ![qrytbltype_type] = acTable
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                  End If
                End If
                If IsNull(![rowsrc_rowsource]) = True Then
                  .Edit
                  ![rowsrc_rowsource] = ctl.RowSource
                  ![rowsrc_datemodified] = Now()
                  .Update
                Else
                  If ![rowsrc_rowsource] <> ctl.RowSource Then
                    .Edit
                    ![rowsrc_rowsource] = ctl.RowSource
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
              Else  ' ** ctl.Rowsource is empty.
                If IsNull(![qrytbltype_type]) = False Then
                  If ![qrytbltype_type] <> acNothing Then
                    .Edit
                    ![qrytbltype_type] = acNothing
                    ![rowsrc_rowsource] = "{empty}"
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                  If ![rowsrc_rowsource] <> "{empty}" Then
                    .Edit
                    ![rowsrc_rowsource] = "{empty}"
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                Else
                  .Edit
                  ![qrytbltype_type] = acNothing
                  ![rowsrc_rowsource] = "{empty}"
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
              End If
              If ctl.RowSource <> vbNullString Then
                If IsNull(![rowsrc_boundcolumn]) = False Then
                  If ![rowsrc_boundcolumn] <> ctl.BoundColumn Then
                    .Edit
'rowsrc_boundcolumn
                    ![rowsrc_boundcolumn] = ctl.BoundColumn
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                Else
                  .Edit
                  ![rowsrc_boundcolumn] = ctl.BoundColumn
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If IsNull(![rowsrc_columncount]) = False Then
                  If ![rowsrc_columncount] <> ctl.ColumnCount Then
                    .Edit
'rowsrc_columncount
                    ![rowsrc_columncount] = ctl.ColumnCount
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                Else
                  .Edit
                  ![rowsrc_columncount] = ctl.ColumnCount
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If IsNull(![rowsrc_columnwidths]) = False Then
                  If ![rowsrc_columnwidths] <> ctl.ColumnWidths Then
                    .Edit
'rowsrc_columnwidths
                    ![rowsrc_columnwidths] = ctl.ColumnWidths
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                Else
                  .Edit
                  If ctl.ColumnWidths <> vbNullString Then
                    ![rowsrc_columnwidths] = ctl.ColumnWidths
                  Else
                    ![rowsrc_columnwidths] = Null
                  End If
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If ctl.ControlType = acComboBox Then
                  If IsNull(![rowsrc_listwidth]) = False Then
                    If ![rowsrc_listwidth] <> ctl.ListWidth Then
                      .Edit
'rowsrc_listwidth
                      ![rowsrc_listwidth] = ctl.ListWidth
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                  Else
                    .Edit
                    ![rowsrc_listwidth] = ctl.ListWidth
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
                If ![rowsrc_columnheads] <> ctl.ColumnHeads Then
                  .Edit
'rowsrc_columnheads
                  ![rowsrc_columnheads] = ctl.ColumnHeads
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If ctl.ControlType = acComboBox Then
                  If ![rowsrc_limittolist] <> ctl.LimitToList Then
                    .Edit
'rowsrc_limittolist
                    ![rowsrc_limittolist] = varTmp00
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                Else
                  If ![rowsrc_limittolist] = True Then
                    .Edit
                    ![rowsrc_limittolist] = False
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
                If ctl.ControlType = acComboBox Then
                  If IsNull(![rowsrc_listrows]) = False Then
                    If ![rowsrc_listrows] <> ctl.ListRows Then
                      .Edit
'rowsrc_listrows
                      ![rowsrc_listrows] = ctl.ListRows
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                  Else
                    .Edit
                    ![rowsrc_listrows] = ctl.ListRows
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
                If ctl.ControlType = acComboBox Then
                  If ![rowsrc_autoexpand] <> ctl.AutoExpand Then
                    .Edit
'rowsrc_autoexpand
                    ![rowsrc_autoexpand] = ctl.AutoExpand
                    ![rowsrc_datemodified] = Now()
                    .Update
                  End If
                End If
                If ctl.ControlType = acListBox Then
                  If IsNull(![rowsrc_multiselect]) = True Then
                    .Edit
'rowsrc_multiselect
                    ![rowsrc_multiselect] = ctl.MultiSelect
                    ![rowsrc_datemodified] = Now()
                    .Update
                  Else
                    If ![rowsrc_multiselect] <> ctl.MultiSelect Then
                      .Edit
'rowsrc_multiselect
                      ![rowsrc_multiselect] = ctl.MultiSelect
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                  End If
                End If
                If ![qrytbltype_type] = acSQL Then
                  If InStr(ctl.RowSource, "forms") > 0 Or InStr(ctl.RowSource, "reports") > 0 Then
                    If ![rowsrc_hasformref] <> True Then
                      .Edit
'rowsrc_hasformref
                      ![rowsrc_hasformref] = True
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                    .Edit
                    intPos1 = InStr(ctl.RowSource, "forms")
                    If intPos1 > 0 Then
                      If intPos1 > 1 Then
                        If Mid$(ctl.RowSource, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
                      End If
                      strFormRef = Mid$(ctl.RowSource, intPos1)
                      intPos1 = InStr(strFormRef, " ")
                      If intPos1 > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
                      strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
'rowsrc_formref
                      ![rowsrc_formref] = strFormRef
                    End If
                    intPos1 = InStr(ctl.RowSource, "reports")
                    If intPos1 > 0 Then
                      If intPos1 > 1 Then
                        If Mid$(ctl.RowSource, (intPos1 - 1), 1) = "[" Then intPos1 = intPos1 - 1
                      End If
                      strFormRef = Mid$(ctl.RowSource, intPos1)
                      intPos1 = InStr(strFormRef, " ")
                      If intPos1 > 0 Then strFormRef = Trim$(Left$(strFormRef, (intPos1 - 1)))
                      strFormRef = FrmRef_Trim(strFormRef)  ' ** Function: Below.
                      ![rowsrc_formref] = strFormRef
                    End If
                    ![rowsrc_datemodified] = Now()
                    .Update
                  Else
                    If ![rowsrc_hasformref] <> False Then
                      .Edit
                      ![rowsrc_hasformref] = False
                      ![rowsrc_formref] = Null
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                    If IsNull(![rowsrc_formref]) = False Then
                      .Edit
                      ![rowsrc_formref] = Null
                      ![rowsrc_datemodified] = Now()
                      .Update
                    End If
                  End If
                Else
                  .Edit
                  ![rowsrc_hasformref] = False
                  ![rowsrc_formref] = Null
'rowsrc_datemodified
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
              Else  ' ** ctl.RowSource is empty.
                If IsNull(![rowsrc_boundcolumn]) = False Then
                  .Edit
                  ![rowsrc_boundcolumn] = Null
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If IsNull(![rowsrc_columncount]) = False Then
                  .Edit
                  ![rowsrc_columncount] = Null
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If IsNull(![rowsrc_columnwidths]) = False Then
                  .Edit
                  ![rowsrc_columnwidths] = Null
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If IsNull(![rowsrc_listwidth]) = False Then
                  .Edit
                  ![rowsrc_listwidth] = Null
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If ![rowsrc_hasformref] <> False Then
                  .Edit
                  ![rowsrc_hasformref] = False
                  ![rowsrc_formref] = Null
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If IsNull(![rowsrc_formref]) = False Then
                  .Edit
                  ![rowsrc_formref] = Null
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If ![rowsrc_columnheads] = True Then
                  .Edit
                  ![rowsrc_columnheads] = False
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If ![rowsrc_limittolist] = True Then
                  .Edit
                  ![rowsrc_limittolist] = False
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If ![rowsrc_autoexpand] = True Then
                  .Edit
                  ![rowsrc_autoexpand] = False
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
                If ![rowsrc_multiselect] <> acMultiSelectNone Then
                  .Edit
                  ![rowsrc_multiselect] = acMultiSelectNone
                  ![rowsrc_datemodified] = Now()
                  .Update
                End If
              End If
            End If
          End With
        Else
          'arr_varRowSrc(ROW_FND, lngX) = CBool(False)
        End If
      End With
    Next
    DoCmd.Close acForm, strLastForm, acSaveNo

    rst.Close

    ' ** Check for mis-marked controls.
    For lngX = 0& To (lngRowSrcs - 1&)
      If arr_varRowSrc(ROW_FND, lngX) = False Then
        If arr_varRowSrc(ROW_RID, lngX) > 0& Then
          ' ** Delete tblForm_Control_RowSource, by specified [rsrcid].
          Set qdf = .QueryDefs("zz_qry_Form_Control_RowSource_01a")
          With qdf.Parameters
            ![rsrcid] = arr_varRowSrc(ROW_RID, lngX)
          End With
        Else
          ' ** Delete tblForm_Control_RowSource, by specified [dbid], [fmid], [ctid].
          Set qdf = .QueryDefs("zz_qry_Form_Control_RowSource_01b")
          With qdf.Parameters
            ![dbid] = arr_varRowSrc(ROW_DID, lngX)
            ![fmid] = arr_varRowSrc(ROW_FID, lngX)
            ![ctid] = arr_varRowSrc(ROW_CID, lngX)
          End With
        End If
        qdf.Execute
      End If
    Next

    lngDels = 0&
    ReDim arr_varDel(0)

    ' ** Then check for controls no longer combo box or list box.
    Set rst = .OpenRecordset("tblForm_Control_RowSource", dbOpenDynaset, dbConsistent)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          blnFound = False
          For lngY = 0& To (lngRowSrcs - 1&)
            If arr_varRowSrc(ROW_DID, lngY) = ![dbs_id] And arr_varRowSrc(ROW_FID, lngY) = ![frm_id] And _
                arr_varRowSrc(ROW_CID, lngY) = ![ctl_id] Then
              blnFound = True
              Exit For
            End If
          Next
          If blnFound = False Then
            lngDels = lngDels + 1&
            lngE = lngDels - 1&
            ReDim Preserve arr_varDel(lngE)
            arr_varDel(lngE) = ![rowsrc_id]
          End If
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    If lngDels > 0& Then
      For lngX = 0& To (lngDels - 1&)
        ' ** Delete tblForm_Control_RowSource, by specified [rsrcid].
        Set qdf = .QueryDefs("zz_qry_Form_Control_RowSource_01a")
        With qdf.Parameters
          ![rsrcid] = arr_varDel(lngX)
        End With
        qdf.Execute
      Next
    End If

    .Close
  End With

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set prp = Nothing
  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Ctl_RowSource_Doc = blnRetValx

End Function

Private Function Frm_Image_Doc() As Boolean
' ** Called by:
' **   QuikFrmDoc(), Above

  Const THIS_PROC As String = "Frm_Image_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim frm As Access.Form, rpt As Access.Report, ctl As Access.Control, prp As Object
  Dim lngFrms As Long, arr_varFrm() As Variant
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim lngCtls As Long, arr_varCtl() As Variant
  Dim lngDels As Long, arr_varDel() As Variant
  Dim lngThisDbsID As Long, lngRecs As Long
  Dim blnSkip As Boolean, blnFound As Boolean
  Dim varTmp00 As Variant
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varFrm().
  Const F_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const F_DID  As Integer = 0
  Const F_DNAM As Integer = 1
  Const F_FID  As Integer = 2
  Const F_FNAM As Integer = 3
  Const F_OTYP As Integer = 4
  Const F_PIC  As Integer = 5
  Const F_PTYP As Integer = 6
  Const F_PSIZ As Integer = 7
  Const F_PALN As Integer = 8
  Const F_PTIL As Integer = 9
  Const F_FND  As Integer = 10

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 10  ' ** Array's first-element UBound().
  Const R_DID  As Integer = 0
  Const R_DNAM As Integer = 1
  Const R_RID  As Integer = 2
  Const R_RNAM As Integer = 3
  Const R_OTYP As Integer = 4
  Const R_PIC  As Integer = 5
  Const R_PTYP As Integer = 6
  Const R_PSIZ As Integer = 7
  Const R_PALN As Integer = 8
  Const R_PTIL As Integer = 9
  Const R_FND  As Integer = 10

  ' ** Array: arr_varCtl().
  Const C_ELEMS As Integer = 12  ' ** Array's first-element UBound().
  Const C_DID  As Integer = 0
  Const C_OID  As Integer = 1
  Const C_OTYP As Integer = 2
  Const C_ONAM As Integer = 3
  Const C_CID  As Integer = 4
  Const C_CNAM As Integer = 5
  Const C_CTYP As Integer = 6
  Const C_PIC  As Integer = 7
  Const C_PTYP As Integer = 8
  Const C_PSIZ As Integer = 9
  Const C_PALN As Integer = 10
  Const C_PTIL As Integer = 11
  Const C_FND  As Integer = 12

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngFrms = 0&
  ReDim arr_varFrm(F_ELEMS, 0)

  lngRpts = 0&
  ReDim arr_varRpt(R_ELEMS, 0)

  lngCtls = 0&
  ReDim arr_varCtl(C_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![dbs_id] = lngThisDbsID Then
          lngFrms = lngFrms + 1&
          lngE = lngFrms - 1&
          ReDim Preserve arr_varFrm(F_ELEMS, lngE)
          ' *********************************************************
          ' ** Array: arr_varFrm()
          ' **
          ' **   Field  Element  Name                    Constant
          ' **   =====  =======  ======================  ==========
          ' **     1       0     frm_id                  F_DID
          ' **     2       1     frm_name                F_DNAM
          ' **     3       2     frm_id                  F_FID
          ' **     4       3     frm_name                F_FNAM
          ' **     5       4     objtype_type            F_OTYP
          ' **     6       5     img_picture             F_PIC
          ' **     7       6     img_picturetype         F_PTYP
          ' **     8       7     img_picturesizemode     F_PSIZ
          ' **     9       8     img_picturealignment    F_PALN
          ' **    10       9     img_picturetiling       F_PTIL
          ' **    11      10     Found                   F_FND
          ' **
          ' *********************************************************
          arr_varFrm(F_DID, lngE) = ![dbs_id]
          arr_varFrm(F_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
          arr_varFrm(F_FID, lngE) = ![frm_id]
          arr_varFrm(F_FNAM, lngE) = ![frm_name]
          arr_varFrm(F_OTYP, lngE) = ![objtype_type]
          arr_varFrm(F_PIC, lngE) = vbNullString
          arr_varFrm(F_PTYP, lngE) = CLng(-1)
          arr_varFrm(F_PSIZ, lngE) = CLng(-1)
          arr_varFrm(F_PALN, lngE) = CLng(-1)
          arr_varFrm(F_PTIL, lngE) = CBool(False)
          arr_varFrm(F_FND, lngE) = CBool(False)
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With  ' ** rst.
    Set rst = Nothing

'PUT THESE INTO REPORTS DOC!
    blnSkip = True
    If blnSkip = False Then
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
            ' *********************************************************
            ' ** Array: arr_varRpt()
            ' **
            ' **   Field  Element  Name                    Constant
            ' **   =====  =======  ======================  ==========
            ' **     1       0     dbs_id                  R_DID
            ' **     2       1     dbs_name                R_DNAM
            ' **     3       2     rpt_id                  R_RID
            ' **     4       3     rpt_name                R_RNAM
            ' **     5       4     objtype_type            R_OTYP
            ' **     6       5     img_picture             R_PIC
            ' **     7       6     img_picturetype         R_PTYP
            ' **     8       7     img_picturesizemode     R_PSIZ
            ' **     9       8     img_picturealignment    R_PALN
            ' **    10       9     img_picturetiling       R_PTIL
            ' **    11      10     Found                   R_FND
            ' **
            ' *********************************************************
            arr_varRpt(R_DID, lngE) = ![dbs_id]
            arr_varRpt(R_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
            arr_varRpt(R_RID, lngE) = ![rpt_id]
            arr_varRpt(R_RNAM, lngE) = ![rpt_name]
            arr_varRpt(R_OTYP, lngE) = ![objtype_type]
            arr_varRpt(R_PIC, lngE) = vbNullString
            arr_varRpt(R_PTYP, lngE) = CLng(-1)
            arr_varRpt(R_PSIZ, lngE) = CLng(-1)
            arr_varRpt(R_PALN, lngE) = CLng(-1)
            arr_varRpt(R_PTIL, lngE) = CBool(False)
            arr_varRpt(R_FND, lngE) = CBool(False)
          End If
          If lngX < lngRecs Then .MoveNext
        Next
        .Close
      End With  ' ** rst.
      Set rst = Nothing
    End If  ' ** blnSkip.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  blnSkip = True
  If blnSkip = False Then
    For lngX = 0& To (lngFrms - 1&)
      DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
      Set frm = Forms(arr_varFrm(F_FNAM, lngX))
      With frm
        For Each prp In .Properties
          Select Case prp.Name
          Case "Picture"
            If IsNull(prp) = False Then
              If Trim(prp) <> vbNullString Then
                arr_varFrm(F_PIC, lngX) = prp
              End If
            End If
          Case "PictureType"
            If IsNull(prp) = False Then
              arr_varFrm(F_PTYP, lngX) = prp
            End If
          Case "PictureSizeMode", "SizeMode"  ' ** Forms use PictureSizeMode, controls use SizeMode.
            If IsNull(prp) = False Then
              arr_varFrm(F_PSIZ, lngX) = prp
            End If
          Case "PictureAlignment"
            If IsNull(prp) = False Then
              arr_varFrm(F_PALN, lngX) = prp
            End If
          Case "PictureTiling"
            arr_varFrm(F_PTIL, lngX) = prp
          End Select
          Set prp = Nothing
        Next
      End With
      DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX), acSaveNo
      Set frm = Nothing
    Next
  End If  ' ** blnSkip.

  blnSkip = False
  If blnSkip = False Then
    For lngX = 0& To (lngFrms - 1&)
      DoCmd.OpenForm arr_varFrm(F_FNAM, lngX), acDesign, , , , acHidden
      Set frm = Forms(arr_varFrm(F_FNAM, lngX))
      With frm
        For Each ctl In .Controls
          lngE = -1&
          For Each prp In ctl.Properties
            Select Case prp.Name
            Case "Picture"
              If IsNull(prp) = False Then
                If Trim(prp) <> vbNullString Then
                  If lngE = -1& Then
                    lngCtls = lngCtls + 1&
                    lngE = lngCtls - 1&
                    ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                    ' *********************************************************
                    ' ** Array: arr_varCtl()
                    ' **
                    ' **   Field  Element  Name                    Constant
                    ' **   =====  =======  ======================  ==========
                    ' **     1       0     dbs_id                  C_DID
                    ' **     2       1     frm_id                  C_OID
                    ' **     3       2     objtype_type            C_OTYP
                    ' **     4       3     frm_name                C_ONAM
                    ' **     5       4     ctl_id                  C_CID
                    ' **     6       5     ctl_name                C_CNAM
                    ' **     7       6     ctltype_type            C_CTYP
                    ' **     8       7     img_picture             C_PIC
                    ' **     9       8     img_picturetype         C_PTYP
                    ' **    10       9     img_picturesizemode     C_PSIZ
                    ' **    11      10     img_picturealignment    C_PALN
                    ' **    12      11     img_picturetiling       C_PTIL
                    ' **    13      12     Found                   C_FND
                    ' **
                    ' *********************************************************
                    arr_varCtl(C_DID, lngE) = arr_varFrm(F_DID, lngX)
                    arr_varCtl(C_OID, lngE) = arr_varFrm(F_FID, lngX)
                    arr_varCtl(C_OTYP, lngE) = acForm
                    arr_varCtl(C_ONAM, lngE) = arr_varFrm(F_FNAM, lngX)
                    arr_varCtl(C_CID, lngE) = CLng(0)
                    arr_varCtl(C_CNAM, lngE) = ctl.Name
                    arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                    arr_varCtl(C_PTYP, lngE) = CLng(-1)
                    arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                    arr_varCtl(C_PALN, lngE) = CLng(-1)
                    arr_varCtl(C_PTIL, lngE) = CBool(False)
                    arr_varCtl(C_FND, lngE) = CBool(False)
                  End If
                  arr_varCtl(C_PIC, lngE) = Nz(prp.Value, vbNullString)
                End If
              End If
            Case "PictureType"
              If IsNull(prp) = False Then
                If lngE = -1& Then
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = arr_varFrm(F_DID, lngX)
                  arr_varCtl(C_OID, lngE) = arr_varFrm(F_FID, lngX)
                  arr_varCtl(C_OTYP, lngE) = acForm
                  arr_varCtl(C_ONAM, lngE) = arr_varFrm(F_FNAM, lngX)
                  arr_varCtl(C_CID, lngE) = CLng(0)
                  arr_varCtl(C_CNAM, lngE) = ctl.Name
                  arr_varCtl(C_PIC, lngE) = vbNullString
                  arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                  arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                  arr_varCtl(C_PALN, lngE) = CLng(-1)
                  arr_varCtl(C_PTIL, lngE) = CBool(False)
                  arr_varCtl(C_FND, lngE) = CBool(False)
                End If
                arr_varCtl(C_PTYP, lngE) = prp
              End If
            Case "PictureSizeMode", "SizeMode"  ' ** Forms use PictureSizeMode, controls use SizeMode.
              If IsNull(prp) = False Then
                If lngE = -1& Then
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = arr_varFrm(F_DID, lngX)
                  arr_varCtl(C_OID, lngE) = arr_varFrm(F_FID, lngX)
                  arr_varCtl(C_OTYP, lngE) = acForm
                  arr_varCtl(C_ONAM, lngE) = arr_varFrm(F_FNAM, lngX)
                  arr_varCtl(C_CID, lngE) = CLng(0)
                  arr_varCtl(C_CNAM, lngE) = ctl.Name
                  arr_varCtl(C_PIC, lngE) = vbNullString
                  arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                  arr_varCtl(C_PTYP, lngE) = CLng(-1)
                  arr_varCtl(C_PALN, lngE) = CLng(-1)
                  arr_varCtl(C_PTIL, lngE) = CBool(False)
                  arr_varCtl(C_FND, lngE) = CBool(False)
                End If
                arr_varCtl(C_PSIZ, lngE) = prp
              End If
            Case "PictureAlignment"
              If IsNull(prp) = False Then
                If lngE = -1& Then
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = arr_varFrm(F_DID, lngX)
                  arr_varCtl(C_OID, lngE) = arr_varFrm(F_FID, lngX)
                  arr_varCtl(C_OTYP, lngE) = acForm
                  arr_varCtl(C_ONAM, lngE) = arr_varFrm(F_FNAM, lngX)
                  arr_varCtl(C_CID, lngE) = CLng(0)
                  arr_varCtl(C_CNAM, lngE) = ctl.Name
                  arr_varCtl(C_PIC, lngE) = vbNullString
                  arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                  arr_varCtl(C_PTYP, lngE) = CLng(-1)
                  arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                  arr_varCtl(C_PTIL, lngE) = CBool(False)
                  arr_varCtl(C_FND, lngE) = CBool(False)
                End If
                arr_varCtl(C_PALN, lngE) = prp
              End If
            Case "PictureTiling"
              If lngE = -1& Then
                lngCtls = lngCtls + 1&
                lngE = lngCtls - 1&
                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                arr_varCtl(C_DID, lngE) = arr_varFrm(F_DID, lngX)
                arr_varCtl(C_OID, lngE) = arr_varFrm(F_FID, lngX)
                arr_varCtl(C_OTYP, lngE) = acForm
                arr_varCtl(C_ONAM, lngE) = arr_varFrm(F_FNAM, lngX)
                arr_varCtl(C_CID, lngE) = CLng(0)
                arr_varCtl(C_CNAM, lngE) = ctl.Name
                arr_varCtl(C_PIC, lngE) = vbNullString
                arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                arr_varCtl(C_PTYP, lngE) = CLng(-1)
                arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                arr_varCtl(C_PALN, lngE) = CLng(-1)
                arr_varCtl(C_FND, lngE) = CBool(False)
              End If
              arr_varCtl(C_PTIL, lngE) = prp
            End Select
            Set prp = Nothing
          Next
          Set ctl = Nothing
        Next
      End With
      DoCmd.Close acForm, arr_varFrm(F_FNAM, lngX), acSaveNo
      Set frm = Nothing
    Next
  End If  ' ** blnSkip.

  blnSkip = True
  If blnSkip = False Then
    For lngX = 0& To (lngRpts - 1&)
      DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acDesign, , , acHidden
      Set rpt = Reports(arr_varRpt(R_RNAM, lngX))
      With rpt
        .Visible = False
        For Each prp In .Properties
          Select Case prp.Name
          Case "Picture"
            If IsNull(prp) = False Then
              If Trim(prp) <> vbNullString Then
                arr_varRpt(R_PIC, lngX) = prp
              End If
            End If
          Case "PictureType"
            If IsNull(prp) = False Then
              arr_varRpt(R_PTYP, lngX) = prp
            End If
          Case "PictureSizeMode", "SizeMode"  ' ** Forms use PictureSizeMode, controls use SizeMode.
            If IsNull(prp) = False Then
              arr_varRpt(R_PSIZ, lngX) = prp
            End If
          Case "PictureAlignment"
            If IsNull(prp) = False Then
              arr_varRpt(R_PALN, lngX) = prp
            End If
          Case "PictureTiling"
            arr_varRpt(R_PTIL, lngX) = prp
          End Select
          Set prp = Nothing
        Next
      End With
      DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX), acSaveNo
      Set rpt = Nothing
    Next
  End If  ' ** blnSkip.

  blnSkip = True
  If blnSkip = False Then
    For lngX = 0& To (lngRpts - 1&)
      DoCmd.OpenReport arr_varRpt(R_RNAM, lngX), acDesign, , , acHidden
      Set rpt = Reports(arr_varRpt(R_RNAM, lngX))
      With rpt
        rpt.Visible = False
        For Each ctl In .Controls
          lngE = -1&
          For Each prp In ctl.Properties
            Select Case prp.Name
            Case "Picture"
              If IsNull(prp) = False Then
                If Trim(prp) <> vbNullString Then
                  If lngE = -1& Then
                    lngCtls = lngCtls + 1&
                    lngE = lngCtls - 1&
                    ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                    ' *********************************************************
                    ' ** Array: arr_varCtl()
                    ' **
                    ' **   Field  Element  Name                    Constant
                    ' **   =====  =======  ======================  ==========
                    ' **     1       0     dbs_id                  C_DID
                    ' **     2       1     rpt_id                  C_OID
                    ' **     3       2     objtype_type            C_OTYP
                    ' **     4       3     rpt_name                C_ONAM
                    ' **     5       4     ctl_id                  C_CID
                    ' **     6       5     ctl_name                C_CNAM
                    ' **     7       6     ctltype_type            C_CTYP
                    ' **     8       7     img_picture             C_PIC
                    ' **     9       8     img_picturetype         C_PTYP
                    ' **    10       9     img_picturesizemode     C_PSIZ
                    ' **    11      10     img_picturealignment    C_PALN
                    ' **    12      11     img_picturetiling       C_PTIL
                    ' **    13      12     Found                   C_FND
                    ' **
                    ' *********************************************************
                    arr_varCtl(C_DID, lngE) = arr_varRpt(R_DID, lngX)
                    arr_varCtl(C_OID, lngE) = arr_varRpt(R_RID, lngX)
                    arr_varCtl(C_OTYP, lngE) = acReport
                    arr_varCtl(C_ONAM, lngE) = arr_varRpt(R_RNAM, lngX)
                    arr_varCtl(C_CID, lngE) = CLng(0)
                    arr_varCtl(C_CNAM, lngE) = ctl.Name
                    arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                    arr_varCtl(C_PTYP, lngE) = CLng(-1)
                    arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                    arr_varCtl(C_PALN, lngE) = CLng(-1)
                    arr_varCtl(C_PTIL, lngE) = CBool(False)
                    arr_varCtl(C_FND, lngE) = CBool(False)
                  End If
                  arr_varCtl(C_PIC, lngE) = Nz(prp.Value, vbNullString)
                End If
              End If
            Case "PictureType"
              If IsNull(prp) = False Then
                If lngE = -1& Then
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = arr_varRpt(R_DID, lngX)
                  arr_varCtl(C_OID, lngE) = arr_varRpt(R_RID, lngX)
                  arr_varCtl(C_OTYP, lngE) = acReport
                  arr_varCtl(C_ONAM, lngE) = arr_varRpt(R_RNAM, lngX)
                  arr_varCtl(C_CID, lngE) = CLng(0)
                  arr_varCtl(C_CNAM, lngE) = ctl.Name
                  arr_varCtl(C_PIC, lngE) = vbNullString
                  arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                  arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                  arr_varCtl(C_PALN, lngE) = CLng(-1)
                  arr_varCtl(C_PTIL, lngE) = CBool(False)
                  arr_varCtl(C_FND, lngE) = CBool(False)
                End If
                arr_varCtl(C_PTYP, lngE) = prp
              End If
            Case "PictureSizeMode", "SizeMode"  ' ** Forms use PictureSizeMode, controls use SizeMode.
              If IsNull(prp) = False Then
                If lngE = -1& Then
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = arr_varRpt(R_DID, lngX)
                  arr_varCtl(C_OID, lngE) = arr_varRpt(R_RID, lngX)
                  arr_varCtl(C_OTYP, lngE) = acReport
                  arr_varCtl(C_ONAM, lngE) = arr_varRpt(R_RNAM, lngX)
                  arr_varCtl(C_CID, lngE) = CLng(0)
                  arr_varCtl(C_CNAM, lngE) = ctl.Name
                  arr_varCtl(C_PIC, lngE) = vbNullString
                  arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                  arr_varCtl(C_PTYP, lngE) = CLng(-1)
                  arr_varCtl(C_PALN, lngE) = CLng(-1)
                  arr_varCtl(C_PTIL, lngE) = CBool(False)
                  arr_varCtl(C_FND, lngE) = CBool(False)
                End If
                arr_varCtl(C_PSIZ, lngE) = prp
              End If
            Case "PictureAlignment"
              If IsNull(prp) = False Then
                If lngE = -1& Then
                  lngCtls = lngCtls + 1&
                  lngE = lngCtls - 1&
                  ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                  arr_varCtl(C_DID, lngE) = arr_varRpt(R_DID, lngX)
                  arr_varCtl(C_OID, lngE) = arr_varRpt(R_RID, lngX)
                  arr_varCtl(C_OTYP, lngE) = acReport
                  arr_varCtl(C_ONAM, lngE) = arr_varRpt(R_RNAM, lngX)
                  arr_varCtl(C_CID, lngE) = CLng(0)
                  arr_varCtl(C_CNAM, lngE) = ctl.Name
                  arr_varCtl(C_PIC, lngE) = vbNullString
                  arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                  arr_varCtl(C_PTYP, lngE) = CLng(-1)
                  arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                  arr_varCtl(C_PTIL, lngE) = CBool(False)
                  arr_varCtl(C_FND, lngE) = CBool(False)
                End If
                arr_varCtl(C_PALN, lngE) = prp
              End If
            Case "PictureTiling"
              If lngE = -1& Then
                lngCtls = lngCtls + 1&
                lngE = lngCtls - 1&
                ReDim Preserve arr_varCtl(C_ELEMS, lngE)
                arr_varCtl(C_DID, lngE) = arr_varRpt(R_DID, lngX)
                arr_varCtl(C_OID, lngE) = arr_varRpt(R_RID, lngX)
                arr_varCtl(C_OTYP, lngE) = acReport
                arr_varCtl(C_ONAM, lngE) = arr_varRpt(R_RNAM, lngX)
                arr_varCtl(C_CID, lngE) = CLng(0)
                arr_varCtl(C_CNAM, lngE) = ctl.Name
                arr_varCtl(C_PIC, lngE) = vbNullString
                arr_varCtl(C_CTYP, lngE) = ctl.ControlType
                arr_varCtl(C_PTYP, lngE) = CLng(-1)
                arr_varCtl(C_PSIZ, lngE) = CLng(-1)
                arr_varCtl(C_PALN, lngE) = CLng(-1)
                arr_varCtl(C_FND, lngE) = CBool(False)
              End If
              arr_varCtl(C_PTIL, lngE) = prp
            End Select
            Set prp = Nothing
          Next
          Set ctl = Nothing
        Next
      End With
      DoCmd.Close acReport, arr_varRpt(R_RNAM, lngX), acSaveNo
      Set frm = Nothing
    Next
  End If  ' ** blnSkip.

  Set dbs = CurrentDb
  With dbs
    Set rst = .OpenRecordset("tblObject_Image", dbOpenDynaset)
    With rst
      blnSkip = True
      If blnSkip = False Then
        For lngX = 0& To (lngFrms - 1&)
          If arr_varFrm(F_PIC, lngX) <> "(None)" Then
            If .BOF = True And .EOF = True Then
              .AddNew
              ![dbs_id] = arr_varFrm(F_DID, lngX)
            Else
              .MoveFirst
              .FindFirst "[dbs_id] = " & CStr(arr_varFrm(F_DID, lngX)) & " And [obj_id] = " & CStr(arr_varFrm(F_FID, lngX)) & " And " & _
                "[Objtype_type] = " & CStr(acForm) & " And [ctltype_type] = " & CStr(acNone)
              Select Case .NoMatch
              Case True
                .AddNew
                ![dbs_id] = arr_varFrm(F_DID, lngX)
              Case False
                .Edit
              End Select
            End If
            ' ** ![img_id] = {AutoNumber}
            ![objtype_type] = acForm
            ![obj_id] = arr_varFrm(F_FID, lngX)
            ![ctltype_type] = acNone
            ![ctl_id] = Null
            If arr_varFrm(F_PIC, lngX) <> vbNullString Then
              ![img_picture] = arr_varFrm(F_PIC, lngX)
            End If
            If arr_varFrm(F_PTYP, lngX) > -1& Then
              ![img_picturetype] = arr_varFrm(F_PTYP, lngX)
            End If
            If arr_varFrm(F_PSIZ, lngX) > -1& Then
              ![img_picturesizemode] = arr_varFrm(F_PSIZ, lngX)  ' ** Forms use PictureSizeMode, controls use SizeMode.
            End If
            If arr_varFrm(F_PALN, lngX) > -1& Then
              ![img_picturealignment] = arr_varFrm(F_PALN, lngX)
            End If
            ![img_picturetiling] = arr_varFrm(F_PTIL, lngX)
            ![img_datemodified] = Now()
            .Update
          End If  ' ** (None).
        Next
      End If  ' ** blnSkip.
      blnSkip = False
      If blnSkip = False Then
        For lngX = 0& To (lngCtls - 1&)
          If arr_varCtl(C_CID, lngX) = 0& Then
            varTmp00 = DLookup("[ctl_id]", "tblForm_Control", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngX)) & " And " & _
              "[frm_id] = " & CStr(arr_varCtl(C_OID, lngX)) & " And [ctl_name] = '" & arr_varCtl(C_CNAM, lngX) & "'")
            If IsNull(varTmp00) = False Then
              arr_varCtl(C_CID, lngX) = CLng(varTmp00)
            End If
          End If
        Next
        For lngX = 0& To (lngCtls - 1&)
          If arr_varCtl(C_PIC, lngX) <> "(None)" Then
            If .BOF = True And .EOF = True Then
              .AddNew
              ![dbs_id] = arr_varCtl(C_DID, lngX)
            Else
              .MoveFirst
              .FindFirst "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngX)) & " And [obj_id] = " & CStr(arr_varCtl(C_OID, lngX)) & " And " & _
                "[Objtype_type] = " & CStr(acForm) & " And [ctl_id] = " & CStr(arr_varCtl(C_CID, lngX)) & " And " & _
                "[ctltype_type] = " & CStr(arr_varCtl(C_CTYP, lngX))
              If .NoMatch = True Then
                .AddNew
                ![dbs_id] = arr_varCtl(C_DID, lngX)
              Else
                .Edit
              End If
            End If
            ' ** ![img_id] = {AutoNumber}
            ![objtype_type] = acForm
            ![obj_id] = arr_varCtl(C_OID, lngX)
            ![ctltype_type] = arr_varCtl(C_CTYP, lngX)
            ![ctl_id] = arr_varCtl(C_CID, lngX)
            If arr_varCtl(C_PIC, lngX) <> vbNullString Then
              ![img_picture] = arr_varCtl(C_PIC, lngX)
            End If
            If arr_varCtl(C_PTYP, lngX) > -1& Then
              ![img_picturetype] = arr_varCtl(C_PTYP, lngX)
            End If
            If arr_varCtl(C_PSIZ, lngX) > -1& Then
              ![img_picturesizemode] = arr_varCtl(C_PSIZ, lngX)  ' ** Forms use PictureSizeMode, controls use SizeMode.
            End If
            If arr_varCtl(C_PALN, lngX) > -1& Then
              ![img_picturealignment] = arr_varCtl(C_PALN, lngX)
            End If
            ![img_picturetiling] = arr_varCtl(C_PTIL, lngX)
            ![img_datemodified] = Now()
            .Update
          End If  ' ** (None).
        Next
      End If  ' ** blnSkip.
      blnSkip = True
      If blnSkip = False Then
        For lngX = 0& To (lngRpts - 1&)
          If arr_varRpt(R_PIC, lngX) <> "(None)" Then
            If .BOF = True And .EOF = True Then
              .AddNew
              ![dbs_id] = arr_varRpt(R_DID, lngX)
            Else
              .MoveFirst
              .FindFirst "[dbs_id] = " & CStr(arr_varRpt(R_DID, lngX)) & " And [obj_id] = " & CStr(arr_varRpt(R_RID, lngX)) & " And " & _
                "[Objtype_type] = " & CStr(acReport) & " And [ctltype_type] = " & CStr(acNone)
              If .NoMatch = True Then
                .AddNew
                ![dbs_id] = arr_varRpt(R_DID, lngX)
              Else
                .Edit
              End If
            End If
            ' ** ![img_id] = {AutoNumber}
            ![objtype_type] = acReport
            ![obj_id] = arr_varRpt(R_RID, lngX)
            ![ctltype_type] = acNone
            ![ctl_id] = Null
            If arr_varRpt(R_PIC, lngX) <> vbNullString Then
              ![img_picture] = arr_varRpt(R_PIC, lngX)
            End If
            If arr_varRpt(R_PTYP, lngX) > -1& Then
              ![img_picturetype] = arr_varRpt(R_PTYP, lngX)
            End If
            If arr_varRpt(R_PSIZ, lngX) > -1& Then
              ![img_picturesizemode] = arr_varRpt(R_PSIZ, lngX)  ' ** Forms use PictureSizeMode, controls use SizeMode.
            End If
            If arr_varRpt(R_PALN, lngX) > -1& Then
              ![img_picturealignment] = arr_varRpt(R_PALN, lngX)
            End If
            ![img_picturetiling] = arr_varRpt(R_PTIL, lngX)
            ![img_datemodified] = Now()
            .Update
          End If  ' ** (None).
        Next
      End If  ' ** blnSkip.
      blnSkip = True
      If blnSkip = False Then
        For lngX = 0& To (lngCtls - 1&)
          If arr_varCtl(C_CID, lngX) = 0& Then
            varTmp00 = DLookup("[ctl_id]", "tblReport_Control", "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngX)) & " And " & _
              "[rpt_id] = " & CStr(arr_varCtl(C_OID, lngX)) & " And [ctl_name] = '" & arr_varCtl(C_CNAM, lngX) & "'")
            If IsNull(varTmp00) = False Then
              arr_varCtl(C_CID, lngX) = CLng(varTmp00)
            End If
          End If
        Next
        For lngX = 0& To (lngCtls - 1&)
          If arr_varCtl(C_PIC, lngX) <> "(None)" Then
            If .BOF = True And .EOF = True Then
              .AddNew
              ![dbs_id] = arr_varCtl(C_DID, lngX)
            Else
              .MoveFirst
              .FindFirst "[dbs_id] = " & CStr(arr_varCtl(C_DID, lngX)) & " And [obj_id] = " & CStr(arr_varCtl(C_OID, lngX)) & " And " & _
                "[Objtype_type] = " & CStr(acReport) & " And [ctl_id] = " & CStr(arr_varCtl(C_CID, lngX)) & " And " & _
                "[ctltype_type] = " & CStr(arr_varCtl(C_CTYP, lngX))
              If .NoMatch = True Then
                .AddNew
                ![dbs_id] = arr_varCtl(C_DID, lngX)
              Else
                .Edit
              End If
            End If
            ' ** ![img_id] = {AutoNumber}
            ![objtype_type] = acReport
            ![obj_id] = arr_varCtl(C_OID, lngX)
            ![ctltype_type] = arr_varCtl(C_CTYP, lngX)
            ![ctl_id] = arr_varCtl(C_CID, lngX)
            If arr_varCtl(C_PIC, lngX) <> vbNullString Then
              ![img_picture] = arr_varCtl(C_PIC, lngX)
            End If
            If arr_varCtl(C_PTYP, lngX) > -1& Then
              ![img_picturetype] = arr_varCtl(C_PTYP, lngX)
            End If
            If arr_varCtl(C_PSIZ, lngX) > -1& Then
              ![img_picturesizemode] = arr_varCtl(C_PSIZ, lngX)  ' ** Forms use PictureSizeMode, controls use SizeMode.
            End If
            If arr_varCtl(C_PALN, lngX) > -1& Then
              ![img_picturealignment] = arr_varCtl(C_PALN, lngX)
            End If
            ![img_picturetiling] = arr_varCtl(C_PTIL, lngX)
            ![img_datemodified] = Now()
            .Update
          End If  ' ** (None).
        Next
      End If  ' ** blnSkip.
      .Close
    End With  ' ** rst.

    lngDels = 0&
    ReDim arr_varDel(0)

    Set rst = .OpenRecordset("tblObject_Image", dbOpenDynaset)
    With rst
      If .BOF = True And .EOF = True Then
        ' ** That would be strange!
      Else
        .MoveLast
        lngRecs = .RecordCount
        .MoveFirst
        For lngX = 1& To lngRecs
          If ![dbs_id] = lngThisDbsID Then
            blnFound = False
            Select Case ![objtype_type]
            Case acForm
              If IsNull(![ctl_id]) = True Then
                For lngY = 0& To (lngFrms - 1&)
                  If arr_varFrm(F_FID, lngY) = ![obj_id] Then
                    If arr_varFrm(F_PIC, lngY) <> vbNullString Then
                      If arr_varFrm(F_PIC, lngY) <> "(None)" Then
                        blnFound = True
                        Exit For
                      End If
                    End If
                  End If
                Next
                If blnFound = False Then
                  lngDels = lngDels + 1&
                  lngE = lngDels - 1&
                  ReDim Preserve arr_varDel(lngE)
                  arr_varDel(lngE) = ![img_id]
                End If
              Else
                For lngY = 0& To (lngCtls - 1&)
                  If arr_varCtl(C_OTYP, lngY) = acForm And arr_varCtl(C_OID, lngY) = ![obj_id] And arr_varCtl(C_CID, lngY) = ![ctl_id] Then
                    If arr_varCtl(C_PIC, lngY) <> vbNullString Then
                      If arr_varCtl(C_PIC, lngY) <> "(None)" Then
                        blnFound = True
                        Exit For
                      End If
                    End If
                  End If
                Next
                If blnFound = False Then
                  lngDels = lngDels + 1&
                  lngE = lngDels - 1&
                  ReDim Preserve arr_varDel(lngE)
                  arr_varDel(lngE) = ![img_id]
                End If
              End If
            Case acReport
blnSkip = True
If blnSkip = False Then
              If IsNull(![ctl_id]) = True Then
                For lngY = 0& To (lngRpts - 1&)
                  If arr_varRpt(R_RID, lngY) = ![obj_id] Then
                    If arr_varFrm(R_PIC, lngY) <> vbNullString Then
                      If arr_varFrm(R_PIC, lngY) <> "(None)" Then
                        blnFound = True
                        Exit For
                      End If
                    End If
                  End If
                Next
                If blnFound = False Then
                  lngDels = lngDels + 1&
                  lngE = lngDels - 1&
                  ReDim Preserve arr_varDel(lngE)
                  arr_varDel(lngE) = ![img_id]
                End If
              Else
                For lngY = 0& To (lngCtls - 1&)
                  If arr_varCtl(C_OTYP, lngY) = acReport And arr_varCtl(C_OID, lngY) = ![obj_id] And arr_varCtl(C_CID, lngY) = ![ctl_id] Then
                    If arr_varCtl(C_PIC, lngY) <> vbNullString Then
                      If arr_varCtl(C_PIC, lngY) <> "(None)" Then
                        blnFound = True
                        Exit For
                      End If
                    End If
                  End If
                Next
                If blnFound = False Then
                  lngDels = lngDels + 1&
                  lngE = lngDels - 1&
                  ReDim Preserve arr_varDel(lngE)
                  arr_varDel(lngE) = ![img_id]
                End If
              End If
End If  ' ** blnSkip.
            End Select
          End If
          If lngX < lngRecs Then .MoveNext
        Next
      End If
      .Close
    End With

    If lngDels > 0& Then
      For lngX = 0& To (lngDels - 1&)
        ' ** Delete tblObject_Image, by specified [imgid].
        Set qdf = .QueryDefs("zz_qry_Object_Image_01")
        With qdf.Parameters
          ![imgid] = arr_varDel(lngX)
        End With
        qdf.Execute
      Next
    End If

    .Close
  End With ' ** dbs.

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  Beep

  Set prp = Nothing
  Set ctl = Nothing
  Set rpt = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  ' ** AcImageType enumeration: (my own)
  ' **   0  acImageEmbedded  Embedded  The picture is embedded in the object and becomes part
  ' **                                 of the database file. (Default)
  ' **   1  acImageLinked    Linked    The picture is linked to the object. Microsoft Access
  ' **                                 stores a pointer to the location of the picture on the disk.

  ' ** AcImageSize enumeration: (my own)
  ' **   0  acImageSizeClip     Clip     The picture is displayed in its actual size. If the picture is larger
  ' **                               than the form or report, then the picture is clipped. (Default)
  ' **   1  acImageSizeStretch  Stretch  The picture is stretched horizontally and vertically to fill the entire
  ' **                               form, even if its original ratio of height to width is distorted.
  ' **   3  acImageSizeZoom     Zoom     The picture is enlarged to the maximum extent possible while keeping
  ' **                               its original ratio of height to width.

  ' ** AcImageAlign enumeration: (my own)
  ' **   0  acImageAlignTopLeft      Top Left      The picture is displayed in the top-left corner of the image control,
  ' **                                             Form window, or page of a report.
  ' **   1  acImageAlignTopRight     Top Right     The picture is displayed in the top-right corner of the image control,
  ' **                                             Form window, or page of a report.
  ' **   2  acImageAlignCenter       Center        The picture is centered in the image control, Form window,
  ' **                                             or page of a report. (Default)
  ' **   3  acImageAlignBottomLeft   Bottom Left   The picture is displayed in the bottom-left corner of the image control,
  ' **                                             Form window, or page of a report.
  ' **   4  acImageAlignBottomRight  Bottom Right  The picture is displayed in the bottom-right corner of the image control,
  ' **                                             Form window, or page of a report.
  ' **   5  acImageAlignFormCenter   Form Center   (Forms only) The form's picture is centered horizontally in relation to the
  ' **                                             width of the form and vertically in relation to the height the entire form.

  Frm_Image_Doc = blnRetValx

End Function

Private Function Frm_Ctl_Specs_Doc() As Boolean
' ** Called by:
' **   QuikFrmDoc(), Above.

  Const THIS_PROC As String = "Frm_Ctl_Specs_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rstA As DAO.Recordset, rstB As DAO.Recordset
  Dim cntr As Container, frm As Access.Form, ctl As Access.Control, prp As Object
  Dim lngForms As Long, strForm As String, strThisForm As String
  Dim lngControls As Long, arr_varControl() As Variant
  Dim lngElemF As Long
  Dim lngThisDbsID As Long
  Dim blnFound As Boolean, blnAdd As Boolean, blnDocFrm As Boolean, blnSkip As Boolean
  Dim intPos1 As Integer
  Dim varTmp00 As Variant, lngTmp02 As Long, lngTmp03 As Long, lngTmp04 As Long
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean

'TBL: tblForm_Control_Specification_A  FLDS: 00 (array minus: ctlspec_id)
'TBL: tblForm_Control_Specification_B  FLDS: 00 (array minus: ctlspec_id, ctlspec_datemodified)

  ' ** Array: arr_varControl().
  Const C_ELEMS As Integer = 79  ' ** Array's first-element UBound().
  Const C_DID      As Integer = 0
  Const C_DNAM     As Integer = 1
  Const C_FID      As Integer = 2
  Const C_FNAM     As Integer = 3
  Const C_CID      As Integer = 4
  Const C_CNAM     As Integer = 5
  Const C_AUTOCOR  As Integer = 6   ' ** Beg Specs_A.
  Const C_AUTOACT  As Integer = 7
  Const C_AUTOREP  As Integer = 8
  Const C_BAKCOLOR As Integer = 9
  Const C_BAKSTYLE As Integer = 10
  Const C_BRDCOLOR As Integer = 11
  Const C_BRDSTYLE As Integer = 12
  Const C_BRDWDTH  As Integer = 13
  Const C_BOTMRGN  As Integer = 14
  Const C_CANCEL   As Integer = 15
  Const C_CAP      As Integer = 16
  Const C_CLASS    As Integer = 17
  Const C_CTLSRC   As Integer = 18
  Const C_CTIPTXT  As Integer = 19
  Const C_CTYP     As Integer = 20
  Const C_DEFAULT  As Integer = 21
  Const C_DEFVAL   As Integer = 22
  Const C_DISPTYP  As Integer = 23
  Const C_ENABLED  As Integer = 24
  Const C_FNTBOLD  As Integer = 25
  Const C_FNTITAL  As Integer = 26
  Const C_FNTNAM   As Integer = 27
  Const C_FNTSIZ   As Integer = 28
  Const C_FNTUNDER As Integer = 29
  Const C_FNTWT    As Integer = 30
  Const C_FORCOLOR As Integer = 31
  Const C_FORMAT   As Integer = 32
  Const C_HEIGHT   As Integer = 33
  Const C_INPMASK  As Integer = 34
  Const C_LEFT     As Integer = 35
  Const C_LFTMRGN  As Integer = 36
  Const C_LINESLNT As Integer = 37
  Const C_LINESPAC As Integer = 38
  Const C_LOCKED   As Integer = 39
  Const C_MULTIROW As Integer = 40
  Const C_OLECLASS As Integer = 41
  Const C_OLETYPA  As Integer = 42
  Const C_OPTVAL   As Integer = 43  ' ** End Specs_A.
  Const C_PAGES    As Integer = 44  ' ** Beg Specs_B.
  Const C_PGIDX    As Integer = 45
  Const C_PIC      As Integer = 46
  Const C_PICALGN  As Integer = 47
  Const C_PICTILE  As Integer = 48
  Const C_PICTYP   As Integer = 49
  Const C_RGTMRGN  As Integer = 50
  Const C_SBARALGN As Integer = 51
  Const C_SCRLBAR  As Integer = 52
  Const C_SEC      As Integer = 53
  Const C_SHMENBAR As Integer = 54
  Const C_SIZMOD   As Integer = 55
  Const C_SRCDOC   As Integer = 56
  Const C_SRCITM   As Integer = 57
  Const C_SRCOBJ   As Integer = 58
  Const C_SPECEFF  As Integer = 59
  Const C_STBARTXT As Integer = 60
  Const C_STYLE    As Integer = 61
  Const C_TBFIXHT  As Integer = 62
  Const C_TBFIXWD  As Integer = 63
  Const C_TABIDX   As Integer = 64
  Const C_TABSTOP  As Integer = 65
  Const C_TAG      As Integer = 66
  Const C_TXTALGN  As Integer = 67
  Const C_TOP      As Integer = 68
  Const C_TOPMRGN  As Integer = 69
  Const C_TRANSP   As Integer = 70
  Const C_TRIPLE   As Integer = 71
  Const C_UPDATOPT As Integer = 72
  Const C_VALRUL   As Integer = 73
  Const C_VALTXT   As Integer = 74
  Const C_VERB     As Integer = 75
  Const C_VISIBLE  As Integer = 76
  Const C_WIDTH    As Integer = 77
  Const C_GAP      As Integer = 78
  Const C_AB       As Integer = 79

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  strThisForm = vbNullString  ' ** If vbNullString, all forms.

  lngControls = 0&
  ReDim arr_varControl(C_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    Set cntr = .Containers![Forms]
    With cntr

      lngForms = .Documents.Count

      ' ** Load arr_varSub() subform array.
      blnRetVal = FrmRpt_SubLoad  ' ** Function: Below.

      Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
      DoEvents

      lngTmp02 = 0&: lngTmp03 = 0&: lngTmp04 = 0&

      For lngX = 0& To (lngForms - 1&)

        blnDocFrm = False
        lngElemF = lngX

        If strThisForm = vbNullString Then
          blnDocFrm = True
        ElseIf .Documents(lngElemF).Name = strThisForm Then
           blnDocFrm = True
        End If

        If blnDocFrm = True Then

          strForm = .Documents(lngElemF).Name
          DoCmd.OpenForm strForm, acDesign, , , , acHidden
          Set frm = Forms(strForm)

          With frm

            lngTmp03 = 0&

            For Each ctl In .Controls
              With ctl

                lngControls = lngControls + 1&
                lngE = lngControls - 1&
                ReDim Preserve arr_varControl(C_ELEMS, lngE)
                arr_varControl(C_DID, lngE) = lngThisDbsID
                arr_varControl(C_DNAM, lngE) = CurrentAppName  ' ** Module Function: modFileUtilities.
                arr_varControl(C_FNAM, lngE) = frm.Name
                varTmp00 = DLookup("[frm_id]", "tblForm", "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_name] = '" & frm.Name & "'")
                If IsNull(varTmp00) = False Then
                  arr_varControl(C_FID, lngE) = CLng(varTmp00)
                Else
                  Stop
                End If
                arr_varControl(C_CNAM, lngE) = .Name
                varTmp00 = DLookup("[ctl_id]", "tblForm_Control", "[dbs_id] = " & CStr(lngThisDbsID) & " AND " & _
                  "[frm_id] = " & CStr(varTmp00) & " AND [ctl_name] = '" & .Name & "'")
                If IsNull(varTmp00) = False Then
                  arr_varControl(C_CID, lngE) = CLng(varTmp00)
                Else
                  Stop
                End If

                ' ** Initialize some elements.
                arr_varControl(C_GAP, lngE) = CLng(0)

                For Each prp In .Properties
                  With prp
                    Select Case .Name
                    Case "AllowAutoCorrect"
                      arr_varControl(C_AUTOCOR, lngE) = .Value
'ctlspec_allowautocorrect
                    Case "AutoActivate"
                      arr_varControl(C_AUTOACT, lngE) = .Value
'ctlspec_autoactivate
                    Case "AutoRepeat"
                      arr_varControl(C_AUTOREP, lngE) = .Value
'ctlspec_autorepeat
                    Case "BackColor"
                      arr_varControl(C_BAKCOLOR, lngE) = .Value
'ctlspec_backcolor
                    Case "BackStyle"
                      arr_varControl(C_BAKSTYLE, lngE) = .Value
'ctlspec_backstyle
                    Case "BorderColor"
                      arr_varControl(C_BRDCOLOR, lngE) = .Value
'ctlspec_bordercolor
                    Case "BorderStyle"
                      arr_varControl(C_BRDSTYLE, lngE) = .Value
'ctlspec_borderstyle
                    Case "BorderWidth"
                      arr_varControl(C_BRDWDTH, lngE) = .Value
'ctlspec_borderwidth
                    Case "BottomMargin"
                      arr_varControl(C_BOTMRGN, lngE) = .Value
'ctlspec_bottommargin
                    Case "Cancel"
                      arr_varControl(C_CANCEL, lngE) = .Value
'ctlspec_cancel
                    Case "Caption"
                      arr_varControl(C_CAP, lngE) = .Value
'ctlspec_caption
                    Case "Class"
                      arr_varControl(C_CLASS, lngE) = .Value
'ctlspec_class
                    Case "ControlSource"
                      arr_varControl(C_CTLSRC, lngE) = .Value
'ctlspec_controlsource
                    Case "ControlTipText"
                      arr_varControl(C_CTIPTXT, lngE) = .Value
'ctlspec_controltiptext
                    Case "ControlType"
                      arr_varControl(C_CTYP, lngE) = .Value
'ctlspec_controltype
                    Case "Default"
                      arr_varControl(C_DEFAULT, lngE) = .Value
'ctlspec_default
                    Case "DefaultValue"
                      arr_varControl(C_DEFVAL, lngE) = .Value
'ctlspec_defaultvalue
                    Case "DisplayType"
                      arr_varControl(C_DISPTYP, lngE) = .Value
'ctlspec_displaytype
                    Case "Enabled"
                      arr_varControl(C_ENABLED, lngE) = .Value
'ctlspec_enabled
                    Case "FontBold"
                      arr_varControl(C_FNTBOLD, lngE) = .Value
'ctlspec_fontbold
                    Case "FontItalic"
                      arr_varControl(C_FNTITAL, lngE) = .Value
'ctlspec_fontitalic
                    Case "FontName"
                      arr_varControl(C_FNTNAM, lngE) = .Value
'ctlspec_fontname
                    Case "FontSize"
                      arr_varControl(C_FNTSIZ, lngE) = .Value
'ctlspec_fontsize
                    Case "FontUnderline"
                      arr_varControl(C_FNTUNDER, lngE) = .Value
'ctlspec_fontunderline
                    Case "FontWeight"
                      arr_varControl(C_FNTWT, lngE) = .Value
'ctlspec_fontweight
                    Case "ForeColor"
                      arr_varControl(C_FORCOLOR, lngE) = .Value
'ctlspec_forecolor
                    Case "Format"
                      arr_varControl(C_FORMAT, lngE) = .Value
'ctlspec_format
                    Case "Height"
                      arr_varControl(C_HEIGHT, lngE) = .Value
'ctlspec_height
                    Case "InputMask"
                      arr_varControl(C_INPMASK, lngE) = .Value
'ctlspec_inputmask
                    Case "Left"
                      arr_varControl(C_LEFT, lngE) = .Value
'ctlspec_left
                    Case "LeftMargin"
                      arr_varControl(C_LFTMRGN, lngE) = .Value
'ctlspec_leftmargin
                    Case "LineSlant"
                      arr_varControl(C_LINESLNT, lngE) = .Value
'ctlspec_lineslant
                    Case "LineSpacing"
                      arr_varControl(C_LINESPAC, lngE) = .Value
'ctlspec_linespacing
                    Case "Locked"
                      arr_varControl(C_LOCKED, lngE) = .Value
'ctlspec_locked
                    Case "MultiRow"
                      arr_varControl(C_MULTIROW, lngE) = .Value
'ctlspec_multirow
                    Case "OLEClass"
                      arr_varControl(C_OLECLASS, lngE) = .Value
'ctlspec_oleclass
                    Case "OLETypeAllowed"
                      arr_varControl(C_OLETYPA, lngE) = .Value
'ctlspec_oletypeallowed
                    Case "OptionValue"                           ' ** End Specs_A.
                      arr_varControl(C_OPTVAL, lngE) = .Value
'ctlspec_optionvalue
                    Case "PageIndex"                             ' ** Beg Specs_B.
                      arr_varControl(C_PGIDX, lngE) = .Value
'ctlspec_pageindex
                    Case "Pages"
                      arr_varControl(C_PAGES, lngE) = .Count
'ctlspec_pages
                    Case "Picture"
                      arr_varControl(C_PIC, lngE) = .Value
'ctlspec_picture
                    Case "PictureAlignment"
                      arr_varControl(C_PICALGN, lngE) = .Value
'ctlspec_picturealignment
                    Case "PictureTiling"
                      arr_varControl(C_PICTILE, lngE) = .Value
'ctlspec_picturetiling
                    Case "PictureType"
                      arr_varControl(C_PICTYP, lngE) = .Value
'ctlspec_picturetype
                    Case "RightMargin"
                      arr_varControl(C_RGTMRGN, lngE) = .Value
'ctlspec_rightmargin
                    Case "ScrollBarAlign"
                      arr_varControl(C_SBARALGN, lngE) = .Value
'ctlspec_scrollbaralign
                    Case "ScrollBars"
                      arr_varControl(C_SCRLBAR, lngE) = .Value
'ctlspec_scrollbars
                    Case "Section"
                      arr_varControl(C_SEC, lngE) = .Value
'ctlspec_section
                    Case "ShortcutMenuBar"
                      arr_varControl(C_SHMENBAR, lngE) = .Value
'ctlspec_shortcutmenubar
                    Case "SizeMode"  ' ** Forms use PictureSizeMode, controls use SizeMode.
                      arr_varControl(C_SIZMOD, lngE) = .Value
'ctlspec_sizemode
                    Case "SourceDoc"
                      arr_varControl(C_SRCDOC, lngE) = .Value
'ctlspec_sourcedoc
                    Case "SourceItem"
                      arr_varControl(C_SRCITM, lngE) = .Value
'ctlspec_sourceitem
                    Case "SourceObject"
                      arr_varControl(C_SRCOBJ, lngE) = .Value
'ctlspec_sourceobject
                    Case "SpecialEffect"
                      arr_varControl(C_SPECEFF, lngE) = .Value
'ctlspec_specialeffect
                    Case "StatusBarText"
                      arr_varControl(C_STBARTXT, lngE) = .Value
'ctlspec_statusbartext
                    Case "Style"
                      arr_varControl(C_STYLE, lngE) = .Value
'ctlspec_style
                    Case "TabFixedHeight"
                      arr_varControl(C_TBFIXHT, lngE) = .Value
'ctlspec_tabfixedheight
                    Case "TabFixedWidth"
                      arr_varControl(C_TBFIXWD, lngE) = .Value
'ctlspec_tabfixedwidth
                    Case "TabIndex"
                      arr_varControl(C_TABIDX, lngE) = .Value
'ctlspec_tabindex
                    Case "TabStop"
                      arr_varControl(C_TABSTOP, lngE) = .Value
'ctlspec_tabstop
                    Case "Tag"
                      arr_varControl(C_TAG, lngE) = .Value
'ctlspec_tag
                    Case "TextAlign"
                      arr_varControl(C_TXTALGN, lngE) = .Value
'ctlspec_textalign
                    Case "Top"
                      arr_varControl(C_TOP, lngE) = .Value
'ctlspec_top
                    Case "TopMargin"
                      arr_varControl(C_TOPMRGN, lngE) = .Value
'ctlspec_topmargin
                    Case "Transparent"
                      arr_varControl(C_TRANSP, lngE) = .Value
'ctlspec_transparent
                    Case "TripleState"
                      arr_varControl(C_TRIPLE, lngE) = .Value
'ctlspec_triplestate
                    Case "UpdateOptions"
                      arr_varControl(C_UPDATOPT, lngE) = .Value
'ctlspec_updateoptions
                    Case "ValidationRule"
                      arr_varControl(C_VALRUL, lngE) = .Value
'ctlspec_validationrule
                    Case "ValidationText"
                      arr_varControl(C_VALTXT, lngE) = .Value
'ctlspec_validationtext
                    Case "Verb"
                      arr_varControl(C_VERB, lngE) = .Value
'ctlspec_verb
                    Case "Visible"
                      arr_varControl(C_VISIBLE, lngE) = .Value
'ctlspec_visible
                    Case "Width"
                      arr_varControl(C_WIDTH, lngE) = .Value
'ctlspec_width
                    End Select
                  End With  ' ** This property: prp.
                Next  ' ** For each property: prp.

              End With  ' ** This control: ctl.

              If IsEmpty(arr_varControl(C_LEFT, lngE)) = False And IsEmpty(arr_varControl(C_WIDTH, lngE)) = False Then
                If IsNull(arr_varControl(C_LEFT, lngE)) = False And IsNull(arr_varControl(C_WIDTH, lngE)) = False Then
                  If arr_varControl(C_LEFT, lngE) > 0& And arr_varControl(C_WIDTH, lngE) > 0& Then
                    arr_varControl(C_GAP, lngE) = (frm.Width - (arr_varControl(C_LEFT, lngE) + arr_varControl(C_WIDTH, lngE)))
                  End If
                End If
              End If

            Next  ' ** For each control: ctl.
            Set ctl = Nothing

          End With  ' ** This form: frm.
          Set frm = Nothing

          DoCmd.Close acForm, strForm, acSaveNo

          If lngTmp03 > 0& Then
            lngTmp04 = lngTmp04 + 1&
          End If

        End If  ' ** blnDocFrm.

      Next  ' ** For each form: lngX, lngElemF.

    End With  ' ** Document container: cntr.
    Set cntr = Nothing

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  For lngX = 0& To (lngControls - 1&)
    For lngY = 0& To C_ELEMS
      If IsEmpty(arr_varControl(lngY, lngX)) = True Then
        arr_varControl(lngY, lngX) = Null
      End If
    Next
  Next

  If strThisForm = vbNullString Then
'Stop
    Set dbs = CurrentDb
    With dbs
      ' ** Delete tblForm_Control_Specification_A, by specified [dbid].
      Set qdf = .QueryDefs("zz_qry_Form_Control_Specification_01")
      With qdf.Parameters
        ![dbid] = lngThisDbsID
      End With
      qdf.Execute
      .Close
    End With
    Set qdf = Nothing
    Set dbs = Nothing
    ChangeSeed_Ext "tblForm_Control_Specification_A"  ' ** Module Function: modAutonumberFieldFuncs.
  End If

  Set dbs = CurrentDb
  Set rstA = dbs.OpenRecordset("tblForm_Control_Specification_A", dbOpenDynaset, dbConsistent)
  Set rstB = dbs.OpenRecordset("tblForm_Control_Specification_B", dbOpenDynaset, dbConsistent)
  For lngX = 0& To (lngControls - 1&)
    With rstA
      blnAdd = False
      If .BOF = True And .EOF = True Then
        blnAdd = True
      Else
        .FindFirst "[dbs_id] = " & CStr(arr_varControl(C_DID, lngX)) & " And [ctl_id] = " & CStr(arr_varControl(C_CID, lngX))
        If .NoMatch = True Then
          blnAdd = True
        Else
          rstB.FindFirst "[ctlspec_id] = " & ![ctlspec_id]
          If rstB.NoMatch = True Then
            Stop
          End If
        End If
      End If
      If blnAdd = True Then
        .AddNew
        ![dbs_id] = arr_varControl(C_DID, lngX)
      Else
        .Edit
      End If
      ![frm_id] = arr_varControl(C_FID, lngX)
      ![ctl_id] = arr_varControl(C_CID, lngX)
      If blnAdd = True Then
        ' ** tblForm_Control_Specification_A![ctlspec_id] : AutoNumber.
        ' ** tblForm_Control_Specification_B![ctlspec_id] : Long Integer
        With rstB
          .AddNew
'dbs_id
          ![dbs_id] = arr_varControl(C_DID, lngX)
'frm_id
          ![frm_id] = arr_varControl(C_FID, lngX)
'ctl_id
          ![ctl_id] = arr_varControl(C_CID, lngX)
        End With
      End If
'ctlspec_allowautocorrect
      If IsNull(arr_varControl(C_AUTOCOR, lngX)) = False Then
        ![ctlspec_allowautocorrect] = arr_varControl(C_AUTOCOR, lngX)
      End If
'ctlspec_autoactivate
      If IsNull(arr_varControl(C_AUTOACT, lngX)) = False Then
        ![ctlspec_autoactivate] = arr_varControl(C_AUTOACT, lngX)
      End If
'ctlspec_autorepeat
      If IsNull(arr_varControl(C_AUTOREP, lngX)) = False Then
        ![ctlspec_autorepeat] = arr_varControl(C_AUTOREP, lngX)
      End If
'ctlspec_backcolor
      If IsNull(arr_varControl(C_BAKCOLOR, lngX)) = False Then
        ![ctlspec_backcolor] = arr_varControl(C_BAKCOLOR, lngX)
      End If
'ctlspec_backstyle
      If IsNull(arr_varControl(C_BAKSTYLE, lngX)) = False Then
        ![ctlspec_backstyle] = arr_varControl(C_BAKSTYLE, lngX)
      End If
'ctlspec_bordercolor
      If IsNull(arr_varControl(C_BRDCOLOR, lngX)) = False Then
        ![ctlspec_bordercolor] = arr_varControl(C_BRDCOLOR, lngX)
      End If
'ctlspec_borderstyle
      If IsNull(arr_varControl(C_BRDSTYLE, lngX)) = False Then
        ![ctlspec_borderstyle] = arr_varControl(C_BRDSTYLE, lngX)
      End If
'ctlspec_borderwidth
      If IsNull(arr_varControl(C_BRDWDTH, lngX)) = False Then
        ![ctlspec_borderwidth] = arr_varControl(C_BRDWDTH, lngX)
      End If
'ctlspec_bottommargin
      If IsNull(arr_varControl(C_BOTMRGN, lngX)) = False Then
        ![ctlspec_bottommargin] = arr_varControl(C_BOTMRGN, lngX)
      End If
'ctlspec_cancel
      If IsNull(arr_varControl(C_CANCEL, lngX)) = False Then
        ![ctlspec_cancel] = arr_varControl(C_CANCEL, lngX)
      End If
'ctlspec_caption
      If IsNull(arr_varControl(C_CAP, lngX)) = False Then
        If Trim(arr_varControl(C_CAP, lngX)) <> vbNullString Then
          ![ctlspec_caption] = arr_varControl(C_CAP, lngX)
        End If
      End If
'ctlspec_controlsource
      If IsNull(arr_varControl(C_CTLSRC, lngX)) = False Then
        If Trim(arr_varControl(C_CTLSRC, lngX)) <> vbNullString Then
          ![ctlspec_controlsource] = arr_varControl(C_CTLSRC, lngX)
        End If
      End If
'ctlspec_controltiptext
      If IsNull(arr_varControl(C_CTIPTXT, lngX)) = False Then
        If Trim(arr_varControl(C_CTIPTXT, lngX)) <> vbNullString Then
          ![ctlspec_controltiptext] = arr_varControl(C_CTIPTXT, lngX)
        End If
      End If
'ctlspec_controltype
      If IsNull(arr_varControl(C_CTYP, lngX)) = False Then
        ![ctlspec_controltype] = arr_varControl(C_CTYP, lngX)
      End If
'ctlspec_default
      If IsNull(arr_varControl(C_DEFAULT, lngX)) = False Then
        ![ctlspec_default] = arr_varControl(C_DEFAULT, lngX)
      End If
'ctlspec_defaultvalue
      If IsNull(arr_varControl(C_DEFVAL, lngX)) = False Then
        If Trim(arr_varControl(C_DEFVAL, lngX)) <> vbNullString Then
          ![ctlspec_defaultvalue] = arr_varControl(C_DEFVAL, lngX)
        End If
      End If
'ctlspec_displaytype
      If IsNull(arr_varControl(C_DISPTYP, lngX)) = False Then
        ![ctlspec_displaytype] = arr_varControl(C_DISPTYP, lngX)
      End If
'ctlspec_enabled
      If IsNull(arr_varControl(C_ENABLED, lngX)) = False Then
        ![ctlspec_enabled] = arr_varControl(C_ENABLED, lngX)
      End If
'ctlspec_fontbold
      If IsNull(arr_varControl(C_FNTBOLD, lngX)) = False Then
        ![ctlspec_fontbold] = arr_varControl(C_FNTBOLD, lngX)
      End If
'ctlspec_fontitalic
      If IsNull(arr_varControl(C_FNTITAL, lngX)) = False Then
        ![ctlspec_fontitalic] = arr_varControl(C_FNTITAL, lngX)
      End If
'ctlspec_fontname
      If IsNull(arr_varControl(C_FNTNAM, lngX)) = False Then
        ![ctlspec_fontname] = arr_varControl(C_FNTNAM, lngX)
      End If
'ctlspec_fontsize
      If IsNull(arr_varControl(C_FNTSIZ, lngX)) = False Then
        ![ctlspec_fontsize] = arr_varControl(C_FNTSIZ, lngX)
      End If
'ctlspec_fontunderline
      If IsNull(arr_varControl(C_FNTUNDER, lngX)) = False Then
        ![ctlspec_fontunderline] = arr_varControl(C_FNTUNDER, lngX)
      End If
'ctlspec_fontweight
      If IsNull(arr_varControl(C_FNTWT, lngX)) = False Then
        ![ctlspec_fontweight] = arr_varControl(C_FNTWT, lngX)
      End If
'ctlspec_forecolor
      If IsNull(arr_varControl(C_FORCOLOR, lngX)) = False Then
        ![ctlspec_forecolor] = arr_varControl(C_FORCOLOR, lngX)
      End If
'ctlspec_format
      If IsNull(arr_varControl(C_FORMAT, lngX)) = False Then
        If Trim(arr_varControl(C_FORMAT, lngX)) <> vbNullString Then
          ![ctlspec_format] = arr_varControl(C_FORMAT, lngX)
        End If
      End If
'ctlspec_height
      If IsNull(arr_varControl(C_HEIGHT, lngX)) = False Then
        ![ctlspec_height] = arr_varControl(C_HEIGHT, lngX)
      End If
'ctlspec_inputmask
      If IsNull(arr_varControl(C_INPMASK, lngX)) = False Then
        If Trim(arr_varControl(C_INPMASK, lngX)) <> vbNullString Then
          ![ctlspec_inputmask] = arr_varControl(C_INPMASK, lngX)
        End If
      End If
'ctlspec_left
      If IsNull(arr_varControl(C_LEFT, lngX)) = False Then
        ![ctlspec_left] = arr_varControl(C_LEFT, lngX)
      End If
'ctlspec_leftmargin
      If IsNull(arr_varControl(C_LFTMRGN, lngX)) = False Then
        ![ctlspec_leftmargin] = arr_varControl(C_LFTMRGN, lngX)
      End If
'ctlspec_lineslant
      If IsNull(arr_varControl(C_LINESLNT, lngX)) = False Then
        ![ctlspec_lineslant] = arr_varControl(C_LINESLNT, lngX)
      End If
'ctlspec_linespacing
      If IsNull(arr_varControl(C_LINESPAC, lngX)) = False Then
        ![ctlspec_linespacing] = arr_varControl(C_LINESPAC, lngX)
      End If
'ctlspec_locked
      If IsNull(arr_varControl(C_LOCKED, lngX)) = False Then
        ![ctlspec_locked] = arr_varControl(C_LOCKED, lngX)
      End If
'ctlspec_multirow
      If IsNull(arr_varControl(C_MULTIROW, lngX)) = False Then
        ![ctlspec_multirow] = arr_varControl(C_MULTIROW, lngX)
      End If
'ctlspec_oletypeallowed
      If IsNull(arr_varControl(C_OLETYPA, lngX)) = False Then
        ![ctlspec_oletypeallowed] = arr_varControl(C_OLETYPA, lngX)
      End If
'ctlspec_optionvalue
      If IsNull(arr_varControl(C_OPTVAL, lngX)) = False Then
        ![ctlspec_optionvalue] = arr_varControl(C_OPTVAL, lngX)
      End If
'ctlspec_pageindex
      If IsNull(arr_varControl(C_PGIDX, lngX)) = False Then
        rstB![ctlspec_pageindex] = arr_varControl(C_PGIDX, lngX)
      End If
'ctlspec_pages
      If IsNull(arr_varControl(C_PAGES, lngX)) = False Then
        rstB![ctlspec_pages] = arr_varControl(C_PAGES, lngX)
      End If
'ctlspec_picture
      If IsNull(arr_varControl(C_PIC, lngX)) = False Then
        If Trim(arr_varControl(C_PIC, lngX)) <> vbNullString Then
          rstB![ctlspec_picture] = arr_varControl(C_PIC, lngX)
        End If
      End If
'ctlspec_picturealignment
      If IsNull(arr_varControl(C_PICALGN, lngX)) = False Then
        rstB![ctlspec_picturealignment] = arr_varControl(C_PICALGN, lngX)
      End If
'ctlspec_picturetiling
      If IsNull(arr_varControl(C_PICTILE, lngX)) = False Then
        rstB![ctlspec_picturetiling] = arr_varControl(C_PICTILE, lngX)
      End If
'ctlspec_picturetype
      If IsNull(arr_varControl(C_PICTYP, lngX)) = False Then
        rstB![ctlspec_picturetype] = arr_varControl(C_PICTYP, lngX)
      End If
'ctlspec_rightmargin
      If IsNull(arr_varControl(C_RGTMRGN, lngX)) = False Then
        rstB![ctlspec_rightmargin] = arr_varControl(C_RGTMRGN, lngX)
      End If
'ctlspec_scrollbaralign
      If IsNull(arr_varControl(C_SBARALGN, lngX)) = False Then
        rstB![ctlspec_scrollbaralign] = arr_varControl(C_SBARALGN, lngX)
      End If
'ctlspec_scrollbars
      If IsNull(arr_varControl(C_SCRLBAR, lngX)) = False Then
        rstB![ctlspec_scrollbars] = arr_varControl(C_SCRLBAR, lngX)
      End If
'ctlspec_section
      If IsNull(arr_varControl(C_SEC, lngX)) = False Then
        rstB![ctlspec_section] = arr_varControl(C_SEC, lngX)
      End If
'ctlspec_shortcutmenubar
      If IsNull(arr_varControl(C_SHMENBAR, lngX)) = False Then
        If Trim(arr_varControl(C_SHMENBAR, lngX)) <> vbNullString Then
          rstB![ctlspec_shortcutmenubar] = arr_varControl(C_SHMENBAR, lngX)
        End If
      End If
'ctlspec_sizemode
      If IsNull(arr_varControl(C_SIZMOD, lngX)) = False Then
        rstB![ctlspec_sizemode] = arr_varControl(C_SIZMOD, lngX)  ' ** Forms use PictureSizeMode, controls use SizeMode.
      End If
'ctlspec_sourcedoc
      If IsNull(arr_varControl(C_SRCDOC, lngX)) = False Then
        If Trim(arr_varControl(C_SRCDOC, lngX)) <> vbNullString Then
          rstB![ctlspec_sourcedoc] = arr_varControl(C_SRCDOC, lngX)
        End If
      End If
'ctlspec_sourceitem
      If IsNull(arr_varControl(C_SRCITM, lngX)) = False Then
        If Trim(arr_varControl(C_SRCITM, lngX)) <> vbNullString Then
          rstB![ctlspec_sourceitem] = arr_varControl(C_SRCITM, lngX)
        End If
      End If
'ctlspec_sourceobject
      If IsNull(arr_varControl(C_SRCOBJ, lngX)) = False Then
        If Trim(arr_varControl(C_SRCOBJ, lngX)) <> vbNullString Then
          rstB![ctlspec_sourceobject] = arr_varControl(C_SRCOBJ, lngX)
        End If
      End If
'ctlspec_specialeffect
      If IsNull(arr_varControl(C_SPECEFF, lngX)) = False Then
        rstB![ctlspec_specialeffect] = arr_varControl(C_SPECEFF, lngX)
      End If
'ctlspec_statusbartext
      If IsNull(arr_varControl(C_STBARTXT, lngX)) = False Then
        If Trim(arr_varControl(C_STBARTXT, lngX)) <> vbNullString Then
          rstB![ctlspec_statusbartext] = arr_varControl(C_STBARTXT, lngX)
        End If
      End If
'ctlspec_style
      If IsNull(arr_varControl(C_STYLE, lngX)) = False Then
        rstB![ctlspec_style] = arr_varControl(C_STYLE, lngX)
      End If
'ctlspec_tabfixedheight
      If IsNull(arr_varControl(C_TBFIXHT, lngX)) = False Then
        rstB![ctlspec_tabfixedheight] = arr_varControl(C_TBFIXHT, lngX)
      End If
'ctlspec_tabfixedwidth
      If IsNull(arr_varControl(C_TBFIXWD, lngX)) = False Then
        rstB![ctlspec_tabfixedwidth] = arr_varControl(C_TBFIXWD, lngX)
      End If
'ctlspec_tabindex
      If IsNull(arr_varControl(C_TABIDX, lngX)) = False Then
        rstB![ctlspec_tabindex] = arr_varControl(C_TABIDX, lngX)
      End If
'ctlspec_tabstop
      If IsNull(arr_varControl(C_TABSTOP, lngX)) = False Then
        rstB![ctlspec_tabstop] = arr_varControl(C_TABSTOP, lngX)
      End If
'ctlspec_tag
      If IsNull(arr_varControl(C_TAG, lngX)) = False Then
        If Trim(arr_varControl(C_TAG, lngX)) <> vbNullString Then
          rstB![ctlspec_tag] = arr_varControl(C_TAG, lngX)
        End If
      End If
'ctlspec_textalign
      If IsNull(arr_varControl(C_TXTALGN, lngX)) = False Then
        rstB![ctlspec_textalign] = arr_varControl(C_TXTALGN, lngX)
      End If
'ctlspec_top
      If IsNull(arr_varControl(C_TOP, lngX)) = False Then
        rstB![ctlspec_top] = arr_varControl(C_TOP, lngX)
      End If
'ctlspec_topmargin
      If IsNull(arr_varControl(C_TOPMRGN, lngX)) = False Then
        rstB![ctlspec_topmargin] = arr_varControl(C_TOPMRGN, lngX)
      End If
'ctlspec_transparent
      If IsNull(arr_varControl(C_TRANSP, lngX)) = False Then
        rstB![ctlspec_transparent] = arr_varControl(C_TRANSP, lngX)
      End If
'ctlspec_triplestate
      If IsNull(arr_varControl(C_TRIPLE, lngX)) = False Then
        rstB![ctlspec_triplestate] = arr_varControl(C_TRIPLE, lngX)
      End If
'ctlspec_updateoptions
      If IsNull(arr_varControl(C_UPDATOPT, lngX)) = False Then
        rstB![ctlspec_updateoptions] = arr_varControl(C_UPDATOPT, lngX)
      End If
'ctlspec_validationrule
      If IsNull(arr_varControl(C_VALRUL, lngX)) = False Then
        If Trim(arr_varControl(C_VALRUL, lngX)) <> vbNullString Then
          rstB![ctlspec_validationrule] = arr_varControl(C_VALRUL, lngX)
        End If
      End If
'ctlspec_validationtext
      If IsNull(arr_varControl(C_VALTXT, lngX)) = False Then
        If Trim(arr_varControl(C_VALTXT, lngX)) <> vbNullString Then
          rstB![ctlspec_validationtext] = arr_varControl(C_VALTXT, lngX)
        End If
      End If
'ctlspec_verb
      If IsNull(arr_varControl(C_VERB, lngX)) = False Then
        rstB![ctlspec_verb] = arr_varControl(C_VERB, lngX)
      End If
'ctlspec_visible
      If IsNull(arr_varControl(C_VISIBLE, lngX)) = False Then
        rstB![ctlspec_visible] = arr_varControl(C_VISIBLE, lngX)
      End If
'ctlspec_width
      If IsNull(arr_varControl(C_WIDTH, lngX)) = False Then
        rstB![ctlspec_width] = arr_varControl(C_WIDTH, lngX)
      End If
'ctlspec_gap
      If arr_varControl(C_GAP, lngX) <> 0& Then
        rstB![ctlspec_gap] = arr_varControl(C_GAP, lngX)
      End If
      rstB![ctlspec_datemodified] = Now()
      .Update
      If blnAdd = True Then
        .Bookmark = .LastModified
'ctlspec_id
        rstB![ctlspec_id] = ![ctlspec_id]
      End If
      rstB.Update
    End With  ' ** rstA.
  Next  ' ** lngX.
  rstB.Close
  Set rstB = Nothing
  rstA.Close
  Set rstA = Nothing

blnSkip = True
If blnSkip = False Then
  ' ** Update tblJournal_Field with new tab indices.
  With dbs
    ' ** Update zz_qry_Journal_Column_21 (tblJournal_Field, with DLookups() to zz_qry_Journal_Column_20
    ' ** (tblJournal_Field, linked to tblForm_Control_Specification_A/_B, just discrepancies, by specified
    ' ** CurrentAppName())), for (ctlspec_tabindex * 100).
    Set qdf = .QueryDefs("zz_qry_Journal_Column_22")
    qdf.Execute
    ' ** Update zz_qry_Journal_Column_21 (tblJournal_Field, with DLookups() to zz_qry_Journal_Column_20
    ' ** (tblJournal_Field, linked to tblForm_Control_Specification_A/_B, just discrepancies, by specified
    ' ** CurrentAppName())), for ctlspec_tabindex.
    Set qdf = .QueryDefs("zz_qry_Journal_Column_23")
    qdf.Execute
    .Close
  End With
End If  ' ** blnSkip.

  Debug.Print "'DONE!  " & THIS_PROC & "()"

  Beep

  Set prp = Nothing
  Set ctl = Nothing
  Set frm = Nothing
  Set cntr = Nothing
  Set rstA = Nothing
  Set rstB = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Ctl_Specs_Doc = blnRetVal

End Function

Private Function Frm_Specs_Doc(frm As Access.Form, dbs As DAO.Database, lngFrmID As Long) As Boolean
' ** Called by:
' **   Frm_Doc(), Above.

  Const THIS_PROC As String = "Frm_Specs_Doc"

  Dim rstA As DAO.Recordset, rstB As DAO.Recordset
  Dim lngThisDbsID As Long, lngSpecID As Long

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Set rstA = dbs.OpenRecordset("tblForm_Specification_A", dbOpenDynaset, dbConsistent)
  With rstA
    .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_id] = " & CStr(lngFrmID)
'spec_id
    If .NoMatch = True Then
      .AddNew
'dbs_id
     .Fields("dbs_id") = lngThisDbsID
'frm_id
     .Fields("frm_id") = lngFrmID
      .Update
      .Bookmark = .LastModified
'spec_id
      lngSpecID = .Fields("spec_id")
      .Edit
    Else
      lngSpecID = .Fields("spec_id")
      .Edit
    End If
'AllowAdditions
   .Fields("AllowAdditions") = frm.AllowAdditions
'AllowDeletions
   .Fields("AllowDeletions") = frm.AllowDeletions
'AllowDatasheetView
   .Fields("AllowDatasheetView") = frm.AllowDatasheetView
'AllowDesignChanges
   .Fields("AllowDesignChanges") = frm.AllowDesignChanges
'AllowEdits
   .Fields("AllowEdits") = frm.AllowEdits
'AllowFilters
   .Fields("AllowFilters") = frm.AllowFilters
'AllowFormView
   .Fields("AllowFormView") = frm.AllowFormView
'AllowPivotChartView
   .Fields("AllowPivotChartView") = frm.AllowPivotChartView
'AllowPivotTableView
   .Fields("AllowPivotTableView") = frm.AllowPivotTableView
'AutoCenter
   .Fields("AutoCenter") = frm.AutoCenter
'AutoResize
   .Fields("AutoResize") = frm.AutoResize
'BorderStyle
   .Fields("BorderStyle") = frm.BorderStyle
'Caption
    If frm.Caption <> vbNullString Then
     .Fields("Caption") = frm.Caption
    Else
      If IsNull(.Fields("Caption")) = False Then
       .Fields("Caption") = Null
      End If
    End If
'CloseButton
   .Fields("CloseButton") = frm.CloseButton
'ControlBox
   .Fields("ControlBox") = frm.ControlBox
'Cycle
   .Fields("Cycle") = frm.Cycle
'DataEntry
   .Fields("DataEntry") = frm.DataEntry
'DatasheetBackColor
   .Fields("DatasheetBackColor") = frm.DatasheetBackColor
'•DatasheetBorderLineStyle Property  DEPRECATED!
'DatasheetCellsEffect
   .Fields("DatasheetCellsEffect") = frm.DatasheetCellsEffect
'DatasheetColumnHeaderUnderlineStyle
   .Fields("DatasheetColumnHeaderUnderlineStyle") = frm.Properties("DatasheetColumnHeaderUnderlineStyle")
'DatasheetFontHeight
   .Fields("DatasheetFontHeight") = frm.DatasheetFontHeight
'DatasheetFontItalic
   .Fields("DatasheetFontItalic") = frm.DatasheetFontItalic
'DatasheetFontName
   .Fields("DatasheetFontName") = frm.DatasheetFontName
'DatasheetFontUnderline
   .Fields("DatasheetFontUnderline") = frm.DatasheetFontUnderline
'DatasheetFontWeight
   .Fields("DatasheetFontWeight") = frm.DatasheetFontWeight
'DatasheetForeColor
   .Fields("DatasheetForeColor") = frm.DatasheetForeColor
'DatasheetGridlinesBehavior
   .Fields("DatasheetGridlinesBehavior") = frm.DatasheetGridlinesBehavior
'DatasheetGridlinesColor
   .Fields("DatasheetGridlinesColor") = frm.DatasheetGridlinesColor
'DefaultView
   .Fields("DefaultView") = frm.DefaultView
'DividingLines
   .Fields("DividingLines") = frm.DividingLines


'FastLaserPrinting
   .Fields("FastLaserPrinting") = frm.FastLaserPrinting
'•FillColor Property  ?
'Filter
    If frm.Filter <> vbNullString Then
     .Fields("Filter") = frm.Filter
    Else
      If IsNull(.Fields("Filter")) = False Then
       .Fields("Filter") = Null
      End If
    End If
'FrozenColumns
   .Fields("FrozenColumns") = frm.FrozenColumns
'GridX
   .Fields("GridX") = frm.GridX
'GridY
   .Fields("GridY") = frm.GridY
'HasModule
   .Fields("HasModule") = frm.HasModule
'HelpContextId
   .Fields("HelpContextId") = frm.HelpContextId
'HelpFile
    If IsNull(frm.HelpFile) = False Then
      If frm.HelpFile <> vbNullString Then
       .Fields("HelpFile") = frm.HelpFile
      Else
       .Fields("HelpFile") = Null
      End If
    Else
     .Fields("HelpFile") = Null
    End If
'HorizontalDatasheetGridlineStyle
   .Fields("HorizontalDatasheetGridlineStyle") = frm.Properties("HorizontalDatasheetGridlineStyle")
'InsideHeight
   .Fields("InsideHeight") = frm.InsideHeight
'InsideWidth
   .Fields("InsideWidth") = frm.InsideWidth
    .Update
    .Close
  End With
  Set rstA = Nothing

  Set rstB = dbs.OpenRecordset("tblForm_Specification_B", dbOpenDynaset, dbConsistent)
  With rstB
    .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_id] = " & CStr(lngFrmID) & " And [spec_id] = " & CStr(lngSpecID)
    If .NoMatch = True Then
      .AddNew
     .Fields("dbs_id") = lngThisDbsID
     .Fields("frm_id") = lngFrmID
     .Fields("spec_id") = lngSpecID
    Else
      .Edit
    End If
'KeyPreview
   .Fields("KeyPreview") = frm.KeyPreview
'LayoutForPrint
   .Fields("LayoutForPrint") = frm.LayoutForPrint
'LogicalPageWidth
   .Fields("LogicalPageWidth") = frm.LogicalPageWidth
'MenuBar
    If frm.MenuBar <> vbNullString Then
     .Fields("MenuBar") = frm.MenuBar
    Else
      If IsNull(.Fields("MenuBar")) = False Then
       .Fields("MenuBar") = Null
      End If
    End If
'MinMaxButtons
   .Fields("MinMaxButtons") = frm.MinMaxButtons
'Modal
   .Fields("Modal") = frm.Modal
'NavigationButtons
   .Fields("NavigationButtons") = frm.NavigationButtons
'OrderBy
    If frm.OrderBy <> vbNullString Then
     .Fields("OrderBy") = frm.OrderBy
    Else
      If IsNull(.Fields("OrderBy")) = False Then
       .Fields("OrderBy") = Null
      End If
    End If
'PaletteSource
    If IsNull(frm.PaletteSource) = False Then
      If frm.PaletteSource <> vbNullString Then
       .Fields("PaletteSource") = frm.PaletteSource
      Else
       .Fields("PaletteSource") = Null
      End If
    Else
     .Fields("PaletteSource") = Null
    End If
'Picture
   .Fields("Picture") = frm.Picture
'PictureAlignment
   .Fields("PictureAlignment") = frm.PictureAlignment
'PictureSizeMode
   .Fields("PictureSizeMode") = frm.PictureSizeMode  ' ** Forms use PictureSizeMode, controls use SizeMode.
'PictureTiling
   .Fields("PictureTiling") = frm.PictureTiling
'PictureType
   .Fields("PictureType") = frm.PictureType
'PopUp
   .Fields("PopUp") = frm.PopUp
'RecordLocks
   .Fields("RecordLocks") = frm.RecordLocks
'RecordSelectors
   .Fields("RecordSelectors") = frm.RecordSelectors
'RecordsetType
   .Fields("RecordsetType") = frm.RecordsetType
'RowHeight
   .Fields("RowHeight") = frm.RowHeight
'ScrollBars
   .Fields("ScrollBars") = frm.ScrollBars
'ShortcutMenu
   .Fields("ShortcutMenu") = frm.ShortcutMenu
'ShortcutMenuBar
    If frm.ShortcutMenuBar <> vbNullString Then
     .Fields("ShortcutMenuBar") = frm.ShortcutMenuBar
    Else
      If IsNull(.Fields("ShortcutMenuBar")) = False Then
       .Fields("ShortcutMenuBar") = Null
      End If
    End If
'SubdatasheetExpanded
   .Fields("SubdatasheetExpanded") = frm.SubdatasheetExpanded
'SubdatasheetHeight
   .Fields("SubdatasheetHeight") = frm.SubdatasheetHeight
'Tag
    If frm.Tag <> vbNullString Then
     .Fields("Tag") = frm.Tag
    Else
      If IsNull(.Fields("Tag")) = False Then
       .Fields("Tag") = Null
      End If
    End If
'TimerInterval
   .Fields("TimerInterval") = frm.TimerInterval
'Toolbar
    If frm.Toolbar <> vbNullString Then
     .Fields("Toolbar") = frm.Toolbar
    Else
      If IsNull(.Fields("Toolbar")) = False Then
       .Fields("Toolbar") = Null
      End If
    End If
'VerticalDatasheetGridlineStyle
   .Fields("VerticalDatasheetGridlineStyle") = frm.Properties("VerticalDatasheetGridlineStyle")
'ViewsAllowed
   .Fields("ViewsAllowed") = frm.ViewsAllowed
'Visible
   .Fields("Visible") = frm.Visible
'WhatsThisButton
   .Fields("WhatsThisButton") = frm.WhatsThisButton
'Width
   .Fields("Width") = frm.Width
'WindowHeight
   .Fields("WindowHeight") = frm.WindowHeight
'WindowWidth
   .Fields("WindowWidth") = frm.WindowWidth
'spec_datemodified
   .Fields("spec_datemodified") = Now()
    .Update
    .Close
  End With
  Set rstB = Nothing

  Set rstA = Nothing
  Set rstB = Nothing

  Frm_Specs_Doc = blnRetValx

End Function

Private Function TxtCaseComp(varInput1 As Variant, varInput2 As Variant) As Boolean
' ** Return True if they're identical, letter-by-letter, case-by-case.
' ** Called by:
' **   Frm_Ctl_Doc(), Above

  Const THIS_PROC As String = "TxtCaseComp"

  Dim intLen As Integer
  Dim strTmp1 As String, strTmp2 As String
  Dim intX As Integer

  blnRetValx = True

  If IsNull(varInput1) = False And IsNull(varInput2) = False Then
    strTmp1 = varInput1
    strTmp2 = varInput2
    intLen = Len(strTmp1)
    If intLen <> Len(strTmp2) Then
      blnRetValx = False
    Else
      For intX = 1 To intLen
        If Asc(Mid$(strTmp1, intX, 1)) <> Asc(Mid$(strTmp2, intX, 1)) Then
          blnRetValx = False
          Exit For
        End If
      Next
    End If
  End If

  TxtCaseComp = blnRetValx

End Function

Private Function FrmRpt_SubLoad() As Boolean
' ** Populate array of all subforms and subreports.
' ** Called by:
' **   Frm_Ctl_Specs_Doc(), Above

  Const THIS_PROC As String = "FrmRpt_SubLoad"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngRecs As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  lngSubs = 0&
  ReDim arr_varSub(SUB_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    ' ** tblForm_Subform, with add'l fields.
    Set qdf = .QueryDefs("qryForm_Subform_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        ' ****************************************************************
        ' ** Array: arr_varSub()
        ' **
        ' **   Field  Element  Name                         Constant
        ' **   =====  =======  ===========================  ============
        ' **     1       0     frmsub_id/rptsub_id          SUB_FSID
        ' **     2       1     frm_id/rpt_id                SUB_PARID
        ' **     3       2     frm_name/rpt_name            SUB_PARNAM
        ' **     4       3     objtype_type                 SUB_PARTYP
        ' **     5       4     ctl_id                       SUB_CID
        ' **     6       5     ctl_name                     SUB_CNAM
        ' **     7       6     ctltype_type                 SUB_CTYP
        ' **     8       7     frm_id_sub/rpt_id_sub        SUB_SUBID
        ' **     9       8     frm_name_sub/rpt_name_sub    SUB_SUBNAM
        ' **    10       9     objtype_type_sub             SUB_SUBTYP
        ' **
        ' ****************************************************************
        lngSubs = lngSubs + 1&
        lngE = lngSubs - 1&
        ReDim Preserve arr_varSub(SUB_ELEMS, lngE)
        arr_varSub(SUB_FSID, lngE) = ![frmsub_id]
        arr_varSub(SUB_PARID, lngE) = ![frm_id]
        arr_varSub(SUB_PARNAM, lngE) = ![frm_name]
        arr_varSub(SUB_PARTYP, lngE) = ![objtype_type]
        arr_varSub(SUB_CID, lngE) = ![ctl_id]
        arr_varSub(SUB_CNAM, lngE) = ![ctl_name]
        arr_varSub(SUB_CTYP, lngE) = ![ctltype_type]
        arr_varSub(SUB_SUBID, lngE) = ![frm_id_sub]
        arr_varSub(SUB_SUBNAM, lngE) = ![frm_name_sub]
        arr_varSub(SUB_SUBTYP, lngE) = ![objtype_type_sub]
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    ' ** tblReport_Subform, with add'l fields.
    Set qdf = .QueryDefs("qryReport_Subform_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        lngSubs = lngSubs + 1&
        lngE = lngSubs - 1&
        ReDim Preserve arr_varSub(SUB_ELEMS, lngE)
        arr_varSub(SUB_FSID, lngE) = ![rptsub_id]
        arr_varSub(SUB_PARID, lngE) = ![rpt_id]
        arr_varSub(SUB_PARNAM, lngE) = ![rpt_name]
        arr_varSub(SUB_PARTYP, lngE) = ![objtype_type]
        arr_varSub(SUB_CID, lngE) = ![ctl_id]
        arr_varSub(SUB_CNAM, lngE) = ![ctl_name]
        arr_varSub(SUB_CTYP, lngE) = ![ctltype_type]
        arr_varSub(SUB_SUBID, lngE) = ![rpt_id_sub]
        arr_varSub(SUB_SUBNAM, lngE) = ![rpt_name_sub]
        arr_varSub(SUB_SUBTYP, lngE) = ![objtype_type_sub]
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    .Close
  End With

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  FrmRpt_SubLoad = blnRetVal

End Function

Private Function FrmRef_Trim(varInput As Variant) As Variant
' ** Trim off right-end stuff from Form references.
' ** Called by:
' **   Frm_RecSrc_Doc(), Above
' **   Frm_Ctl_RowSource_Doc(), Above.

  Const THIS_PROC As String = "FrmRef_Trim"

  Dim blnLoop As Boolean
  Dim intPos1 As Integer
  Dim strTmp0 As String
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varInput) = False Then
    varRetVal = Trim$(varInput)

    intPos1 = InStr(varRetVal, ",")
    If intPos1 > 0 Then
      varRetVal = Trim$(Left$(varRetVal, (intPos1 - 1)))
    End If
    intPos1 = InStr(varRetVal, ")")
    If intPos1 > 0 Then
      varRetVal = Trim$(Left$(varRetVal, (intPos1 - 1)))
    End If
    intPos1 = InStr(varRetVal, ";")
    If intPos1 > 0 Then
      varRetVal = Trim$(Left$(varRetVal, (intPos1 - 1)))
    End If
    intPos1 = InStr(varRetVal, "+")
    If intPos1 > 0 Then
      varRetVal = Trim$(Left$(varRetVal, (intPos1 - 1)))
    End If
    intPos1 = InStr(varRetVal, Chr(10))
    If intPos1 > 0 Then
      varRetVal = Trim$(Left$(varRetVal, (intPos1 - 1)))
    End If
    intPos1 = InStr(varRetVal, Chr(13))
    If intPos1 > 0 Then
      varRetVal = Trim$(Left$(varRetVal, (intPos1 - 1)))
    End If

    blnLoop = True
    Do While blnLoop = True
      strTmp0 = Right$(varRetVal, 1)
      Select Case strTmp0
      Case ";", ")", ","
        varRetVal = Left$(varRetVal, (Len(varRetVal) - 1))
      Case Else
        blnLoop = False
      End Select
    Loop

  End If

EXITP:
  FrmRef_Trim = varRetVal
  Exit Function

ERRH:
  varRetVal = Null
  Select Case ERR.Number
  Case Else
    MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
      "Module: " & THIS_NAME & vbCrLf & "Proc: " & THIS_PROC, vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
  End Select
  Resume EXITP

End Function

Public Function Frm_Ctl_Color_Doc() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "Frm_Ctl_Color_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngSysClrs As Long, arr_varSysClr() As Variant
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim lngSysClr_Back_Elem As Long, lngSysClr_Brdr_Elem As Long, lngSysClr_Fore_Elem As Long
  Dim lngRecs As Long
  Dim blnFound_Back As Boolean, blnFound_Brdr As Boolean, blnFound_Fore As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varSysClr().
  Const SC_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const SC_BASEID As Integer = 0
  Const SC_BASE   As Integer = 1
  Const SC_CONST  As Integer = 2
  Const SC_SCLRID As Integer = 3
  Const SC_LONG   As Integer = 4
  Const SC_TYPE   As Integer = 5

  ' ** Array: arr_varCtl().
  Const C_FID  As Integer = 0
  Const C_FNAM As Integer = 1
  Const C_CID  As Integer = 2
  Const C_CNAM As Integer = 3
  Const C_BACK As Integer = 4
  Const C_BRDR As Integer = 5
  Const C_FORE As Integer = 6

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs

    lngSysClrs = 0&
    ReDim arr_varSysClr(SC_ELEMS, 0)

    ' ** tblSystemColor, linked to tblSystemColor_Base, with sysclrtype_name.
    Set qdf = .QueryDefs("zz_qry_SystemColor_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![sysclrtype_name] = "Developer" Then
          lngSysClrs = lngSysClrs + 1&
          lngE = lngSysClrs - 1&
          ReDim Preserve arr_varSysClr(SC_ELEMS, lngE)
          ' *********************************************************
          ' ** Array: arr_varSysClr()
          ' **
          ' **   Field  Element  Name                   Constant
          ' **   =====  =======  =====================  ===========
          ' **     1       0     sysclrbase_id          SC_BASEID
          ' **     2       1     sysclrbase_system      SC_BASE
          ' **     3       2     sysclrbase_constant    SC_CONST
          ' **     4       3     sysclr_id              SC_SCLRID
          ' **     5       4     sysclr_long            SC_LONG
          ' **     6       5     sysclrtype_name        SC_TYPE
          ' **     7       6     sysclr_long_new
          ' **     8       7     Cx
          ' **
          ' *********************************************************
          arr_varSysClr(SC_BASEID, lngE) = ![sysclrbase_id]
          arr_varSysClr(SC_BASE, lngE) = ![sysclrbase_system]
          arr_varSysClr(SC_CONST, lngE) = ![sysclrbase_constant]
          arr_varSysClr(SC_SCLRID, lngE) = ![sysclr_id]
          arr_varSysClr(SC_LONG, lngE) = ![sysclr_long]
          arr_varSysClr(SC_TYPE, lngE) = ![sysclrtype_name]
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    ' ** tblForm_Control, with ctlspec_backcolor, ctlspec_bordercolor, cltspec_forecolor.
    Set qdf = .QueryDefs("zz_qry_Form_Control_03")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' *********************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                   Constant
      ' **   =====  =======  =====================  ===========
      ' **     1       0     frm_id                 C_FID
      ' **     2       1     frm_name               C_FNAM
      ' **     3       2     ctl_id                 C_CID
      ' **     4       3     ctl_name               C_CNAM
      ' **     5       4     ctlspec_backcolor      C_BACK
      ' **     6       5     ctlspec_bordercolor    C_BRDR
      ' **     7       6     ctlspec_forecolor      C_FORE
      ' **
      ' *********************************************************
      .Close
    End With

    Set rst = .OpenRecordset("tblSystemColor_Control", dbOpenDynaset, dbConsistent)
    With rst

      For lngX = 0& To (lngCtls - 1&)

        lngSysClr_Back_Elem = -1&: lngSysClr_Brdr_Elem = -1&: lngSysClr_Fore_Elem = -1&
        blnFound_Back = False: blnFound_Brdr = False: blnFound_Fore = False

        If IsNull(arr_varCtl(C_BACK, lngX)) = False Then
          For lngY = 0& To (lngSysClrs - 1&)
            If arr_varSysClr(SC_BASE, lngY) = arr_varCtl(C_BACK, lngX) Then
              blnFound_Back = True
              lngSysClr_Back_Elem = lngY
            ElseIf (arr_varSysClr(SC_LONG, lngY) = arr_varCtl(C_BACK, lngX)) And (arr_varSysClr(SC_CONST, lngY) <> "vbMenuBar") Then
              ' ** Because vbMenuBar and vbButtonFace are the same on my computer,
              ' ** it encounters vbMenuBar first, but it's really meant to be the other.
              blnFound_Back = True
              lngSysClr_Back_Elem = lngY
            End If
            If blnFound_Back = True Then Exit For
          Next
          If blnFound_Back = True Then
            blnFound_Back = False
            If .BOF = True And .EOF = True Then
              ' ** Just add it.
            Else
              .FindFirst "[objtype_type] = " & CStr(acForm) & " And [obj_id] = " & CStr(arr_varCtl(C_FID, lngX)) & " And " & _
                "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngX)) & " And [sysclrctl_backcolor] = True"
              If .NoMatch = False Then
                blnFound_Back = True
              End If
            End If
            Select Case blnFound_Back
            Case True
              If ![sysclrbase_id] <> arr_varSysClr(SC_BASEID, lngSysClr_Back_Elem) Then
                .Edit
                ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Back_Elem)
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              If ![sysclr_id] <> arr_varSysClr(SC_SCLRID, lngSysClr_Back_Elem) Then
                .Edit
                ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Back_Elem)
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              'If ![sysclrctl_backcolor] = True Then
              '  .Edit
              '  ![sysclrctl_backcolor] = False
              '  ![sysclrctl_datemodified] = Now()
              '  .Update
              'End If
              If ![sysclrctl_bordercolor] = True Then
                .Edit
                ![sysclrctl_bordercolor] = False
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              If ![sysclrctl_forecolor] = True Then
                .Edit
                ![sysclrctl_forecolor] = False
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
            Case False
              .AddNew
              ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Back_Elem)
              ' ** [sysclrctl_id]: AutoNumber.
              ![obj_id] = arr_varCtl(C_FID, lngX)
              ![ctl_id] = arr_varCtl(C_CID, lngX)
              ![objtype_type] = acForm
              ![sysclrctl_backcolor] = True
              ![sysclrctl_bordercolor] = False
              ![sysclrctl_forecolor] = False
              ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Back_Elem)
              ![sysclrctl_datemodified] = Now()
              .Update
            End Select
          End If
        End If  ' ** ctlspec_backcolor.

        If IsNull(arr_varCtl(C_BRDR, lngX)) = False Then
          For lngY = 0& To (lngSysClrs - 1&)
            If arr_varSysClr(SC_BASE, lngY) = arr_varCtl(C_BRDR, lngX) Then
              blnFound_Brdr = True
              lngSysClr_Brdr_Elem = lngY
            ElseIf (arr_varSysClr(SC_LONG, lngY) = arr_varCtl(C_BRDR, lngX)) And (arr_varSysClr(SC_CONST, lngY) <> "vbMenuBar") Then
              ' ** Because vbMenuBar and vbButtonFace are the same on my computer,
              ' ** it encounters vbMenuBar first, but it's really meant to be the other.
              blnFound_Brdr = True
              lngSysClr_Brdr_Elem = lngY
            End If
            If blnFound_Brdr = True Then Exit For
          Next
          If blnFound_Brdr = True Then
            blnFound_Brdr = False
            If .BOF = True And .EOF = True Then
              ' ** Just add it.
            Else
              .FindFirst "[objtype_type] = " & CStr(acForm) & " And [obj_id] = " & CStr(arr_varCtl(C_FID, lngX)) & " And " & _
                "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngX)) & " And [sysclrctl_bordercolor] = True"
              If .NoMatch = False Then
                blnFound_Brdr = True
              End If
            End If
            Select Case blnFound_Brdr
            Case True
              If ![sysclrbase_id] <> arr_varSysClr(SC_BASEID, lngSysClr_Brdr_Elem) Then
                .Edit
                ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Brdr_Elem)
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              If ![sysclr_id] <> arr_varSysClr(SC_SCLRID, lngSysClr_Brdr_Elem) Then
                .Edit
                ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Brdr_Elem)
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              If ![sysclrctl_backcolor] = True Then
                .Edit
                ![sysclrctl_backcolor] = False
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              'If ![sysclrctl_bordercolor] = True Then
              '  .Edit
              '  ![sysclrctl_bordercolor] = False
              '  ![sysclrctl_datemodified] = Now()
              '  .Update
              'End If
              If ![sysclrctl_forecolor] = True Then
                .Edit
                ![sysclrctl_forecolor] = False
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
            Case False
              .AddNew
              ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Brdr_Elem)
              ' ** [sysclrctl_id]: AutoNumber.
              ![obj_id] = arr_varCtl(C_FID, lngX)
              ![ctl_id] = arr_varCtl(C_CID, lngX)
              ![objtype_type] = acForm
              ![sysclrctl_backcolor] = False
              ![sysclrctl_bordercolor] = True
              ![sysclrctl_forecolor] = False
              ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Brdr_Elem)
              ![sysclrctl_datemodified] = Now()
              .Update
            End Select
          End If
        End If  ' ** ctlspec_bordercolor.

        If IsNull(arr_varCtl(C_FORE, lngX)) = False Then
          For lngY = 0& To (lngSysClrs - 1&)
            If arr_varSysClr(SC_BASE, lngY) = arr_varCtl(C_FORE, lngX) Then
              blnFound_Fore = True
              lngSysClr_Fore_Elem = lngY
            ElseIf (arr_varSysClr(SC_LONG, lngY) = arr_varCtl(C_FORE, lngX)) And (arr_varSysClr(SC_CONST, lngY) <> "vbMenuBar") Then
              ' ** Because vbMenuBar and vbButtonFace are the same on my computer,
              ' ** it encounters vbMenuBar first, but it's really meant to be the other.
              blnFound_Fore = True
              lngSysClr_Fore_Elem = lngY
            End If
            If blnFound_Fore = True Then Exit For
          Next
          If blnFound_Fore = True Then
            blnFound_Fore = False
            If .BOF = True And .EOF = True Then
              ' ** Just add it.
            Else
              .FindFirst "[objtype_type] = " & CStr(acForm) & " And [obj_id] = " & CStr(arr_varCtl(C_FID, lngX)) & " And " & _
                "[ctl_id] = " & CStr(arr_varCtl(C_CID, lngX)) & " And [sysclrctl_forecolor] = True"
              If .NoMatch = False Then
                blnFound_Fore = True
              End If
            End If
            Select Case blnFound_Fore
            Case True
              If ![sysclrbase_id] <> arr_varSysClr(SC_BASEID, lngSysClr_Fore_Elem) Then
                .Edit
                ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Fore_Elem)
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              If ![sysclr_id] <> arr_varSysClr(SC_SCLRID, lngSysClr_Fore_Elem) Then
                .Edit
                ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Fore_Elem)
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              If ![sysclrctl_backcolor] = True Then
                .Edit
                ![sysclrctl_backcolor] = False
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              If ![sysclrctl_bordercolor] = True Then
                .Edit
                ![sysclrctl_bordercolor] = False
                ![sysclrctl_datemodified] = Now()
                .Update
              End If
              'If ![sysclrctl_forecolor] = True Then
              '  .Edit
              '  ![sysclrctl_forecolor] = False
              '  ![sysclrctl_datemodified] = Now()
              '  .Update
              'End If
            Case False
              .AddNew
              ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Fore_Elem)
              ' ** [sysclrctl_id]: AutoNumber.
              ![obj_id] = arr_varCtl(C_FID, lngX)
              ![ctl_id] = arr_varCtl(C_CID, lngX)
              ![objtype_type] = acForm
              ![sysclrctl_backcolor] = False
              ![sysclrctl_bordercolor] = False
              ![sysclrctl_forecolor] = True
              ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Fore_Elem)
              ![sysclrctl_datemodified] = Now()
              .Update
            End Select
          End If
        End If  ' ** ctlspec_forecolor.

      Next

      .Close
    End With

    .Close
  End With

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Beep

  Frm_Ctl_Color_Doc = blnRetValx

End Function

Public Function Frm_Sec_Color_Doc() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "Frm_Sec_Color_Doc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
  Dim lngSysClrs As Long, arr_varSysClr() As Variant
  Dim lngSecs As Long, arr_varSec As Variant
  Dim lngSysClr_Elem As Long
  Dim lngRecs As Long
  Dim blnFound As Boolean
  Dim lngX As Long, lngY As Long, lngE As Long

  ' ** Array: arr_varSysClr().
  Const SC_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const SC_BASEID As Integer = 0
  Const SC_BASE   As Integer = 1
  Const SC_CONST  As Integer = 2
  Const SC_SCLRID As Integer = 3
  Const SC_LONG   As Integer = 4
  Const SC_TYPE   As Integer = 5

  ' ** Array: arr_varSec().
  Const S_FID  As Integer = 0
  Const S_FNAM As Integer = 1
  Const S_SID  As Integer = 2
  Const S_IDX  As Integer = 3
  Const S_SNAM As Integer = 4
  Const S_BACK As Integer = 5

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs

    lngSysClrs = 0&
    ReDim arr_varSysClr(SC_ELEMS, 0)

    ' ** tblSystemColor, linked to tblSystemColor_Base, with sysclrtype_name.
    Set qdf = .QueryDefs("zz_qry_SystemColor_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![sysclrtype_name] = "Developer" Then
          lngSysClrs = lngSysClrs + 1&
          lngE = lngSysClrs - 1&
          ReDim Preserve arr_varSysClr(SC_ELEMS, lngE)
          ' *********************************************************
          ' ** Array: arr_varSysClr()
          ' **
          ' **   Field  Element  Name                   Constant
          ' **   =====  =======  =====================  ===========
          ' **     1       0     sysclrbase_id          SC_BASEID
          ' **     2       1     sysclrbase_system      SC_BASE
          ' **     3       2     sysclrbase_constant    SC_CONST
          ' **     4       3     sysclr_id              SC_SCLRID
          ' **     5       4     sysclr_long            SC_LONG
          ' **     6       5     sysclrtype_name        SC_TYPE
          ' **     7       6     sysclr_long_new
          ' **     8       7     Cx
          ' **
          ' *********************************************************
          arr_varSysClr(SC_BASEID, lngE) = ![sysclrbase_id]
          arr_varSysClr(SC_BASE, lngE) = ![sysclrbase_system]
          arr_varSysClr(SC_CONST, lngE) = ![sysclrbase_constant]
          arr_varSysClr(SC_SCLRID, lngE) = ![sysclr_id]
          arr_varSysClr(SC_LONG, lngE) = ![sysclr_long]
          arr_varSysClr(SC_TYPE, lngE) = ![sysclrtype_name]
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With

    ' ** tblForm_Section, with sec_backcolor, sec_backcolor_new.
    Set qdf = .QueryDefs("zz_qry_Form_Section_01")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngSecs = .RecordCount
      .MoveFirst
      arr_varSec = .GetRows(lngSecs)
      ' ***************************************************
      ' ** Array: arr_varSec()
      ' **
      ' **   Field  Element  Name             Constant
      ' **   =====  =======  ===============  ===========
      ' **     1       0     frm_id           S_FID
      ' **     2       1     frm_name         S_FNAM
      ' **     3       2     sec_id           S_SID
      ' **     4       3     sec_index        S_IDX
      ' **     5       4     sec_name         S_SNAM
      ' **     6       5     sec_backcolor    S_BACK
      ' **
      ' ***************************************************
      .Close
    End With

    Set rst = .OpenRecordset("tblSystemColor_Section", dbOpenDynaset, dbConsistent)
    With rst

      For lngX = 0& To (lngSecs - 1&)
        lngSysClr_Elem = -1&
        blnFound = False
        For lngY = 0& To (lngSysClrs - 1&)
          If arr_varSysClr(SC_BASE, lngY) = arr_varSec(S_BACK, lngX) Then
            blnFound = True
            lngSysClr_Elem = lngY
          ElseIf (arr_varSysClr(SC_LONG, lngY) = arr_varSec(S_BACK, lngX)) And (arr_varSysClr(SC_CONST, lngY) <> "vbMenuBar") Then
            ' ** Because vbMenuBar and vbButtonFace are the same on my computer,
            ' ** it encounters vbMenuBar first, but it's really meant to be the other.
            blnFound = True
            lngSysClr_Elem = lngY
          End If
          If blnFound = True Then Exit For
        Next

        If blnFound = True Then
          blnFound = False
          If .BOF = True And .EOF = True Then
            ' ** Just add it.
          Else
            .FindFirst "[objtype_type] = " & CStr(acForm) & " And [obj_id] = " & CStr(arr_varSec(S_FID, lngX)) & " And " & _
              "[sec_id] = " & CStr(arr_varSec(S_SID, lngX))
            If .NoMatch = False Then
              blnFound = True
            End If
          End If
          Select Case blnFound
          Case True
            If ![sysclrbase_id] <> arr_varSysClr(SC_BASEID, lngSysClr_Elem) Then
              .Edit
              ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Elem)
              ![sysclrsec_datemodified] = Now()
              .Update
            End If
            If ![sysclr_id] <> arr_varSysClr(SC_SCLRID, lngSysClr_Elem) Then
              .Edit
              ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Elem)
              ![sysclrsec_datemodified] = Now()
              .Update
            End If
          Case False
            .AddNew
            ![sysclrbase_id] = arr_varSysClr(SC_BASEID, lngSysClr_Elem)
            ' ** [sysclrsec_id]: AutoNumber.
            ![obj_id] = arr_varSec(S_FID, lngX)
            ![sec_id] = arr_varSec(S_SID, lngX)
            ![objtype_type] = acForm
            ![sysclr_id] = arr_varSysClr(SC_SCLRID, lngSysClr_Elem)
            ![sysclrsec_datemodified] = Now()
            .Update
          End Select
        End If

      Next

      .Close
    End With

    .Close
  End With

  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Beep

  Frm_Sec_Color_Doc = blnRetValx

End Function

Public Function Frm_Prop() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "Frm_Prop"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.CommandButton
  Dim lngForms As Long, arr_varForm As Variant
  Dim lngCtls As Long, arr_varCtl As Variant
  Dim lngX As Long, lngY As Long
  Dim strTmp00 As String

  ' ** Array: arr_varForm().
  Const FM_NAM    As Integer = 1
  Const FM_CAP    As Integer = 3
  Const FM_IS_SUB As Integer = 5

  ' ** Array: arr_varCtl().
  Const C_FID   As Integer = 0
  Const C_FNAM  As Integer = 1
  Const C_CID1  As Integer = 2
  Const C_CNAM1 As Integer = 3
  Const C_TIP1  As Integer = 4
  Const C_CID2  As Integer = 5
  Const C_CNAM2 As Integer = 6
  Const C_TIP2  As Integer = 7

  blnRetValx = True

  Set dbs = CurrentDb
  With dbs
    ' ** zz_qry_Form_Shortcut_13 (zz_qry_Form_Shortcut_12 (xx), with ctl_name <> Null), just OKCancel w/o ControlTipText.
    Set qdf = .QueryDefs("zz_qry_Form_Shortcut_13o")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngCtls = .RecordCount
      .MoveFirst
      arr_varCtl = .GetRows(lngCtls)
      ' **************************************************************
      ' ** Array: arr_varCtl()
      ' **
      ' **   Field  Element  Name                        Constant
      ' **   =====  =======  ==========================  ===========
      ' **     1       0     frm_id                     C_FID
      ' **     2       1     frm_name                   C_FNAM
      ' **     3       2     ctl_id1                    C_CID1
      ' **     4       3     ctl_name1                  C_CNAM1
      ' **     5       4     ctlspec_controltiptextx1   C_TIP1
      ' **     6       5     ctl_id2                    C_CID2
      ' **     7       6     ctl_name2                  C_CNAM2
      ' **     8       7     ctlspec_controltiptextx2   C_TIP2
      ' **
      ' **************************************************************
      .Close
    End With
    .Close
  End With

  For lngX = 0& To (lngCtls - 1&)
    If IsNull(arr_varCtl(C_TIP1, lngX)) = True Or IsNull(arr_varCtl(C_TIP2, lngX)) = True Then
      DoCmd.OpenForm arr_varCtl(C_FNAM, lngX), acDesign, , , , acHidden
      Set frm = Forms(arr_varCtl(C_FNAM, lngX))
      With frm
        For lngY = 1& To 2&
          Select Case lngY
          Case 1&
            strTmp00 = arr_varCtl(C_CNAM1, lngX)
          Case 2&
            strTmp00 = arr_varCtl(C_CNAM2, lngX)
          End Select
          Set ctl = .Controls(strTmp00)
          If InStr(strTmp00, "OK") > 0 Then
            ctl.ControlTipText = "OK - Alt+O"
          Else
            ctl.ControlTipText = "Cancel - Alt+C"
          End If
        Next
      End With
      DoCmd.Close acForm, arr_varCtl(C_FNAM, lngX), acSaveYes
    End If
  Next

  'Set dbs = CurrentDb
  'With dbs
    ' ** Get a list of all forms.
    'Set rst = .OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
    'With rst
    '  .MoveLast
    '  lngForms = .RecordCount
    '  .MoveFirst
    '  arr_varform = .GetRows(lngForms)
      ' ******************************************************
      ' ** Array: arr_varForm()
      ' **
      ' **   Field  Element  Name (Default)      Constant
      ' **   =====  =======  ==================  ===========
      ' **     1       0     frm_id
      ' **     2       1     frm_name            FM_NAM
      ' **     3       2     frm_controls
      ' **     4       3     frm_caption         FM_CAP
      ' **     5       4     frm_hassub
      ' **     6       5     frm_issub           FM_IS_SUB
      ' **     7       6     frm_active
      ' **     8       7     frm_description
      ' **     9       8     frm_tag
      ' **    10       9     frm_parent_sub
      ' **    11      10     frm_datemodified
      ' **
      ' ******************************************************
    '  .Close
    'End With
  '  .Close
  'End With

  'For lngX = 0& To (lngForms - 1&)
  '  If IsNull(arr_varForm(FM_CAP, lngX)) = True And arr_varForm(FM_IS_SUB, lngX) = True Then
  '    DoCmd.OpenForm arr_varForm(FM_NAM, lngX), acDesign
  '    Set frm = Forms(arr_varForm(FM_NAM, lngX))
  '    frm.Caption = arr_varForm(FM_NAM, lngX)
  '    DoCmd.Close acForm, arr_varForm(FM_NAM, lngX), acSaveYes
  '  End If
  'Next

  Beep

  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_Prop = blnRetValx

End Function

Public Function Frm_ImageRefresh() As Boolean
' ** Currently not called.

  Const THIS_PROC As String = "Frm_ImageRefresh"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Image
  Dim lngPics As Long, arr_varPic() As Variant
  Dim strFormName As String
  Dim lngRecs As Long
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varPic().
  Const P_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const P_NAM As Integer = 0
  Const P_PIC As Integer = 1
  Const P_ERR As Integer = 2

  blnRetValx = True

  strFormName = "frmJournal_Columns"

  lngPics = 0&
  ReDim arr_varPic(P_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    ' ** zz_qry_Object_Image_03 (tblObject_Image, just form images), just standard path.
    Set qdf = .QueryDefs("zz_qry_Object_Image_05")
    Set rst = qdf.OpenRecordset
    With rst
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      For lngX = 1& To lngRecs
        If ![frm_name] = strFormName Then
          lngPics = lngPics + 1&
          lngE = lngPics - 1&
          ReDim Preserve arr_varPic(P_ELEMS, lngE)
          arr_varPic(P_NAM, lngE) = ![ctl_name]
          arr_varPic(P_PIC, lngE) = ![img_picture]
          arr_varPic(P_ERR, lngE) = CBool(False)
        End If
        If lngX < lngRecs Then .MoveNext
      Next
      .Close
    End With
    .Close
  End With

  If lngPics > 0& Then
    DoCmd.OpenForm strFormName, acDesign, , , , acHidden
    Set frm = Forms(strFormName)
    With frm
      For lngX = 0& To (lngPics - 1&)
        Set ctl = .Controls(arr_varPic(P_NAM, lngX))
        With ctl
          .Picture = vbNullString
          DoEvents
          .Picture = arr_varPic(P_PIC, lngX)
          DoEvents
        End With
      Next
    End With
    DoCmd.Close acForm, strFormName, acSaveYes
  End If

  Beep

  Set ctl = Nothing
  Set frm = Nothing
  Set rst = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Frm_ImageRefresh = blnRetValx

End Function

Public Function Frm_Ctl_Cnt(strFormName As String) As Variant
' ** Currently not called.

  Const THIS_PROC As String = "Frm_Ctl_Cnt"

  Dim frm As Access.Form, ctl As Access.Control
  Dim arr_varRetVal(1) As Variant

  arr_varRetVal(0) = CLng(0)  ' ** Count of controls.
  arr_varRetVal(1) = CLng(0)  ' ** Count of subforms.

  DoCmd.OpenForm strFormName, acDesign, , , , acHidden
  Set frm = Forms(strFormName)
  With frm
    arr_varRetVal(0) = .Controls.Count
    For Each ctl In .Controls
      With ctl
        If .ControlType = acSubform Then
          arr_varRetVal(1) = arr_varRetVal(1) + 1&
        End If
      End With
    Next
  End With
  DoCmd.Close acForm, strFormName, acSaveNo

  Set ctl = Nothing
  Set frm = Nothing

  Frm_Ctl_Cnt = arr_varRetVal

End Function

Public Function Frm_Pub_Color_Doc() As Boolean
' ** Document all Public color constants to tblForm_Color.

  Const THIS_PROC As String = "Frm_Pub_Color_Doc"

  Dim dbs As DAO.Database, rst As DAO.Recordset
  Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
  Dim strModName As String, strLine As String
  Dim lngModLines As Long
  Dim lngClrs As Long, arr_varClr() As Variant
  Dim lngThisDbsID As Long
  Dim blnAdd As Boolean, blnAddAll As Boolean, blnCounted As Boolean
  Dim lngRecs As Long, lngSysColorValue As Long, lngAdds As Long, lngEdits As Long
  Dim intPos1 As Integer
  Dim strTmp00 As String, strTmp01 As String, strTmp02 As String
  Dim lngX As Long, lngE As Long

  ' ** Array: arr_varClr().
  Const C_ELEMS As Integer = 6  ' ** Array's first-element UBound().
  Const C_NAM As Integer = 0
  Const C_VAL As Integer = 1
  Const C_SYS As Integer = 2
  Const C_RED As Integer = 3
  Const C_GRN As Integer = 4
  Const C_BLU As Integer = 5
  Const C_FND As Integer = 6

  blnRetValx = True

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  lngClrs = 0&
  ReDim arr_varClr(C_ELEMS, 0)

  Set vbp = Application.VBE.ActiveVBProject
  With vbp
    Set vbc = .VBComponents("modPublicVariables")
    With vbc
      strModName = .Name
      Set cod = .CodeModule
      With cod
        lngModLines = .CountOfLines
        For lngX = 1& To lngModLines
          strLine = Trim$(.Lines(lngX, 1))
          If strLine <> vbNullString Then  ' ** Not a blank line.
            If Left$(strLine, 1) <> "'" Then  ' ** Not a remark.
              strTmp00 = "Public Const "
              If Left$(strLine, Len(strTmp00)) = strTmp00 Then
                intPos1 = Len(strTmp00) + 1
                strTmp00 = Trim$(Mid$(strLine, intPos1))  ' ** Strip 'Public Const'.
                If Left$(strTmp00, Len("MY_CLR_")) = "MY_CLR_" Or _
                    Left$(strTmp00, Len("CLR_")) = "CLR_" Or _
                    Left$(strTmp00, Len("WIN_CLR_")) = "WIN_CLR_" Then
                  intPos1 = InStr(strTmp00, " ")
                  If intPos1 > 0 Then
                    strTmp01 = Trim$(Mid$(strTmp00, intPos1))  ' ** Everything to the right of constant.
                    strTmp00 = Trim$(Left$(strTmp00, intPos1))  ' ** Just the constant.
                    intPos1 = InStr(strTmp01, "=")
                    If intPos1 > 0 Then
                      strTmp01 = Trim$(Mid$(strTmp01, (intPos1 + 1)))  ' ** Value, plus any description.
                      strTmp02 = vbNullString
                      intPos1 = InStr(strTmp01, "'")
                      If intPos1 > 0 Then
                        strTmp02 = Mid$(strTmp01, intPos1)  ' ** Description.
                        strTmp01 = Trim$(Left$(strTmp01, (intPos1 - 1)))  ' ** Value.
                      End If
                      lngClrs = lngClrs + 1&
                      lngE = lngClrs - 1&
                      ReDim Preserve arr_varClr(C_ELEMS, lngE)
                      arr_varClr(C_NAM, lngE) = strTmp00
                      If Val(strTmp01) < 0 Then
                        arr_varClr(C_VAL, lngE) = Null
                        arr_varClr(C_SYS, lngE) = CLng(Val(strTmp01))
                      Else
                        arr_varClr(C_VAL, lngE) = CLng(Val(strTmp01))
                        arr_varClr(C_SYS, lngE) = Null
                      End If
                      arr_varClr(C_RED, lngE) = Null
                      arr_varClr(C_GRN, lngE) = Null
                      arr_varClr(C_BLU, lngE) = Null
                      arr_varClr(C_FND, lngE) = CBool(False)
                    End If
                  End If
                End If  ' ** Color constant.
              End If  ' ** Public Const.
            End If  ' ** Remark.
          End If  ' ** vbNullString.
        Next  ' ** lngX
      End With  ' ** cod.
    End With  ' ** vbc
  End With  ' ** vbp
  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing

  Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.
  DoEvents


  If lngClrs > 0& Then

    Debug.Print "'CLRS: " & CStr(lngClrs)
    DoEvents

    For lngX = 0& To (lngClrs - 1&)
      lngSysColorValue = -1&
      If IsNull(arr_varClr(C_SYS, lngX)) = False Then
        lngSysColorValue = DLookup("[sysclrbase_index]", "tblSystemColor_Base", "[sysclrbase_system] = " & CStr(arr_varClr(C_SYS, lngX)))
        arr_varClr(C_VAL, lngX) = SystemColor_Get(lngSysColorValue)  ' ** Module Function: modSystemColorFuncs.
      End If
      strTmp00 = RGB_Split(arr_varClr(C_VAL, lngX), False)  ' ** Module Function: modSystemColorFuncs.
      intPos1 = InStr(strTmp00, ",")
      If intPos1 > 0 Then
        strTmp01 = Mid$(strTmp00, (intPos1 + 1))
        strTmp00 = Left$(strTmp00, (intPos1 - 1))  ' ** Red.
        intPos1 = InStr(strTmp01, ",")
        If intPos1 > 0 Then
          strTmp02 = Mid$(strTmp01, (intPos1 + 1))  ' ** Blue.
          strTmp01 = Left$(strTmp01, (intPos1 - 1))  ' ** Green.
          arr_varClr(C_RED, lngX) = CLng(Val(strTmp00))
          arr_varClr(C_GRN, lngX) = CLng(Val(strTmp01))
          arr_varClr(C_BLU, lngX) = CLng(Val(strTmp02))
        End If
      End If
    Next  ' ** lngX.

    lngAdds = 0&: lngEdits = 0&
    Set dbs = CurrentDb
    With dbs
      Set rst = .OpenRecordset("tblForm_Color", dbOpenDynaset, dbConsistent)
      With rst
        blnAdd = False: blnAddAll = False
        If .BOF = True And .EOF = True Then
          blnAddAll = True
        Else
          .MoveLast
          lngRecs = .RecordCount
          .MoveFirst
        End If
        For lngX = 0& To (lngClrs - 1&)
          blnAdd = False: blnCounted = False
          lngSysColorValue = -1&
          Select Case blnAddAll
          Case True
            blnAdd = True
          Case False
            .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [frmclr_constant] = '" & arr_varClr(C_NAM, lngX) & "'"
            Select Case .NoMatch
            Case True
              blnAdd = True
            Case False
              arr_varClr(C_FND, lngX) = CBool(True)
            End Select
          End Select
          Select Case blnAdd
          Case True
            .AddNew
            ![dbs_id] = lngThisDbsID
            ' ** ![frmclr_id] : AutoNumber.
            ![frmclr_constant] = arr_varClr(C_NAM, lngX)
            Select Case IsNull(arr_varClr(C_SYS, lngX))
            Case True
              ![frmclr_long] = arr_varClr(C_VAL, lngX)
            Case False
              ![frmclr_long] = arr_varClr(C_SYS, lngX)
            End Select
            ![frmclr_red] = arr_varClr(C_RED, lngX)
            ![frmclr_green] = arr_varClr(C_GRN, lngX)
            ![frmclr_blue] = arr_varClr(C_BLU, lngX)
            ![frmclr_datemodified] = Now()
            .Update
            lngAdds = lngAdds + 1&
          Case False
            Select Case IsNull(arr_varClr(C_SYS, lngX))
            Case True
              If ![frmclr_long] <> arr_varClr(C_VAL, lngX) Then
                .Edit
                ![frmclr_long] = arr_varClr(C_VAL, lngX)
                ![frmclr_datemodified] = Now()
                .Update
                If blnCounted = False Then
                  lngEdits = lngEdits + 1&
                  blnCounted = True
                End If
              End If
            Case False
              If ![frmclr_long] <> arr_varClr(C_SYS, lngX) Then
                .Edit
                ![frmclr_long] = arr_varClr(C_SYS, lngX)
                ![frmclr_datemodified] = Now()
                .Update
                If blnCounted = False Then
                  lngEdits = lngEdits + 1&
                  blnCounted = True
                End If
              End If
            End Select
            If ![frmclr_red] <> arr_varClr(C_RED, lngX) Then
              .Edit
              ![frmclr_red] = arr_varClr(C_RED, lngX)
              ![frmclr_datemodified] = Now()
              .Update
              If blnCounted = False Then
                lngEdits = lngEdits + 1&
                blnCounted = True
              End If
            End If
            If ![frmclr_green] <> arr_varClr(C_GRN, lngX) Then
              .Edit
              ![frmclr_green] = arr_varClr(C_GRN, lngX)
              ![frmclr_datemodified] = Now()
              .Update
              If blnCounted = False Then
                lngEdits = lngEdits + 1&
                blnCounted = True
              End If
            End If
            If ![frmclr_blue] <> arr_varClr(C_BLU, lngX) Then
              .Edit
              ![frmclr_blue] = arr_varClr(C_BLU, lngX)
              ![frmclr_datemodified] = Now()
              .Update
              If blnCounted = False Then
                lngEdits = lngEdits + 1&
                blnCounted = True
              End If
            End If
          End Select
        Next
        .Close
      End With
      .Close
    End With

    Debug.Print "'ADDS:  " & CStr(lngAdds)
    Debug.Print "'EDITS: " & CStr(lngEdits)

  Else
    Debug.Print "'NONE FOUND!"
  End If

  Debug.Print "'DONE!  " & THIS_PROC & "()"
  DoEvents

  DoCmd.Hourglass False
  Beep

  Set cod = Nothing
  Set vbc = Nothing
  Set vbp = Nothing
  Set rst = Nothing
  Set dbs = Nothing

  Frm_Pub_Color_Doc = blnRetValx

End Function

Public Function Frm_QryImport() As Boolean

  Const THIS_PROC As String = "Frm_QryImport"

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
    ' ** tblQuery, just 'zz_qry_Form_..' queries in Trust.mdb.
    Set qdf = .QueryDefs("zz_qry_Form_10x")
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

  Frm_QryImport = blnRetVal

End Function

Private Function Frm_ChkDocQrys(Optional varSkip As Variant) As Boolean

  Const THIS_PROC As String = "Frm_ChkDocQrys"

  Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, qdf2 As DAO.QueryDef, rst As DAO.Recordset, prp As Object
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

  lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

  Select Case IsMissing(varSkip)
  Case True
    blnSkip = True
  Case False
    blnSkip = varSkip
  End Select

  If TableExists("zz_tbl_Form_Doc") = False Then  ' ** Module Function: modFileUtilities.
    Set dbs = CurrentDb
    With dbs
      ' ** Data-Definition: Create table zz_tbl_Form_Doc.
      Set qdf1 = .QueryDefs("zz_qry_System_58_01")
      qdf1.Execute
      Set qdf1 = Nothing
      ' **    Data-Definition: Create index [dbs_id], [frm_id] Unique on table zz_tbl_Form_Doc.
      Set qdf1 = .QueryDefs("zz_qry_System_58_02")
      qdf1.Execute
      Set qdf1 = Nothing
      ' **    Data-Definition: Create index [frm_id] Unique on table zz_tbl_Form_Doc.
      Set qdf1 = .QueryDefs("zz_qry_System_58_03")
      qdf1.Execute
      Set qdf1 = Nothing
      ' **    Data-Definition: Create index [objtype_type] on table zz_tbl_Form_Doc.
      Set qdf1 = .QueryDefs("zz_qry_System_58_04")
      qdf1.Execute
      Set qdf1 = Nothing
      ' **    Data-Definition: Create index [frmdoc_id] PrimaryKey on table zz_tbl_Form_Doc.
      Set qdf1 = .QueryDefs("zz_qry_System_58_05")
      qdf1.Execute
      Set qdf1 = Nothing
      .Close
    End With  ' ** dbs.
    Set dbs = Nothing
  End If  ' ** TableExists().

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
      Set qdf1 = .QueryDefs("qryQuery_Documentation_01")
      With qdf1.Parameters
        ![vbnam] = THIS_NAME
      End With
      Set rst = qdf1.OpenRecordset
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
      Set qdf1 = Nothing
      .Close
    End With  ' ** dbs.
    Set dbs = Nothing

    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
    DoEvents

    Debug.Print "'FRM DOC QRYS: " & CStr(lngQrys)
    DoEvents

  End If  ' ** blnSkip.


  'If blnSkip = False Then
  '  Set dbs = CurrentDb
  '  With dbs
  '    varTmp00 = DLookup("[vbcom_id]", "tblVBComponent", "[vbcom_name] = '" & THIS_NAME & "'")
  '    If IsNull(varTmp00) = True Then
  '      Stop
  '    End If
  '    Set rst = .OpenRecordset("zz_tbl_Query_Documentation", dbOpenDynaset, dbAppendOnly)
  '    For lngX = 0& To (lngQrys - 1&)
  '      'Set qdf1 = .QueryDefs(arr_varQry(Q_QNAM, lngX))
  '      With rst
  '        .AddNew
  '        ' ** ![qrydoct_id] : AutoNumber.
  '        ![dbs_id] = lngThisDbsID
  '        ![vbcom_id] = varTmp00
  '        '![qry_id] =
  '        ![vbcom_name] = THIS_NAME
  '        ![qry_name] = arr_varQry(Q_QNAM, lngX)
  '        '![qrytype_type] = qdf1.Type
  '        '![qry_description] = qdf1.Properties("Description")
  '        '![qry_sql] = qdf1.SQL
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
              Set dbs = CurrentDb
              With dbs
                ' ** tblQuery_Documentation, by specified [qnam].
                Set qdf1 = .QueryDefs("qryQuery_Documentation_02")
                With qdf1.Parameters
                  ![qnam] = arr_varQry(Q_QNAM, lngX)
                End With
                Set rst = qdf1.OpenRecordset
                With rst
                  If .BOF = True And .EOF = True Then
                    Debug.Print "'QRY NOT FOUND!  " & arr_varQry(Q_QNAM, lngX)
                    DoEvents
                  Else
                    Set qdf2 = dbs.CreateQueryDef(![qry_name], ![qry_sql])
                    Set prp = qdf2.CreateProperty("Description", dbText, ![qry_description])
On Error Resume Next
                    qdf2.Properties.Append prp
                    If ERR.Number <> 0 Then
On Error GoTo 0
                      qdf2.Properties("Description") = ![qry_description]
                    Else
On Error GoTo 0
                    End If
                  End If
                  .Close
                End With
                .QueryDefs.Refresh
                Set rst = Nothing
                Set prp = Nothing
                Set qdf1 = Nothing
                Set qdf2 = Nothing
              End With
              Set dbs = Nothing
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
      Debug.Print "'ALL FRM DOC QRYS PRESENT!"
    End If

    Debug.Print "'DONE!"
    DoEvents

    Beep

  End If  ' ** blnSkip.

  Set prp = Nothing
  Set rst = Nothing
  Set qdf1 = Nothing
  Set qdf2 = Nothing
  Set dbs = Nothing

  Frm_ChkDocQrys = blnRetValx

End Function

Public Function RGB_Split(varInput As Variant, blnHex As Boolean) As String
' ** Converts color number to Red-Green-Blue constituents,
' ** given either a long integer or hex code.
' ** Parameters:
' **   Long Integer, False: 16777215 returns '255,255,255'.
' **   Hex,          True : #FFFFFF  returns '255,255,255'.
' ** See also RGBRev(), below, Hex(), internal, and Hexx(), modStringFuncs.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "RGB_Split"

        Dim intLen As Integer
        Dim strTmp01 As String
        Dim strRetVal As String

1810    strRetVal = vbNullString

1820    If IsNull(varInput) = False Then
1830      strTmp01 = Trim(varInput)
1840      If Left(strTmp01, 1) = "#" Then strTmp01 = Mid(strTmp01, 2)
1850      intLen = Len(strTmp01)
1860      If intLen > 0 Then
            ' ** Office 2007 lists them in hex: 2-2-2, B-G-R.
1870        If blnHex = False Then strTmp01 = Hex(varInput)
1880        strTmp01 = Right("000000" & strTmp01, 6)
1890        intLen = 6
1900        If blnHex = True Then
1910          strTmp01 = Right(strTmp01, 2) & Mid(strTmp01, 3, 2) & Left(strTmp01, 2)
1920        End If
1930        strRetVal = HexX(Right(strTmp01, 2))  ' ** Module Function: modStringFuncs.
1940        If intLen = 2 Then
1950          strRetVal = strRetVal & ",0,0"
1960        ElseIf intLen = 4 Then
1970          strRetVal = strRetVal & "," & HexX(Left(strTmp01, 2))  ' ** Module Function: modStringFuncs.
1980          strRetVal = strRetVal & ",0"
1990        ElseIf intLen = 6 Then
2000          strRetVal = strRetVal & "," & HexX(Mid(strTmp01, 3, 2))  ' ** Module Function: modStringFuncs.
2010          strRetVal = strRetVal & "," & HexX(Left(strTmp01, 2))  ' ** Module Function: modStringFuncs.
2020        End If
2030      End If
2040    End If

EXITP:
2050    RGB_Split = strRetVal
2060    Exit Function

ERRH:
2070    Select Case ERR.Number
        Case 13  ' ** Type mismatch.
2080      strRetVal = "#TYPE_MISMATCH"
2090    Case Else
2100      strRetVal = vbNullString
2110      Beep
2120      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()", _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
2130    End Select
2140    Resume EXITP

End Function
