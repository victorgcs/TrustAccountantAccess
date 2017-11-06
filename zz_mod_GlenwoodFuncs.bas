Attribute VB_Name = "zz_mod_GlenwoodFuncs"
Option Compare Database
Option Explicit

'VGC 06/09/2015: CHANGES!

Private Const THIS_NAME As String = "zz_mod_GlenwoodFuncs"
' **

Public Function OddChar(varInput As Variant) As Variant

  Const THIS_PROC As String = "OddChar"

  Dim intPos1 As Integer, intLen As Integer
  Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long
  Dim intX As Integer
  Dim varRetVal As Variant

  varRetVal = Null

  If IsNull(varInput) = False Then
    If Trim(varInput) <> vbNullString Then
      strTmp01 = Trim(varInput)
      intLen = Len(strTmp01)
      For intX = 1 To intLen
        strTmp02 = Mid(strTmp01, intX, 1)
        lngTmp03 = Asc(strTmp02)
        If (lngTmp03 >= 65& And lngTmp03 <= 90&) Or (lngTmp03 >= 97& And lngTmp03 <= 122&) Then
          ' ** Alphabetic: A-Z, a-z.
        ElseIf (lngTmp03 >= 48& And lngTmp03 <= 57) Then
          ' ** Numeric: 0-9.
        Else
          Select Case lngTmp03
          Case 32&
            ' ** Space.
          Case 38&, 40&, 41&, 44&, 45&, 46&, 47&
            ' ** Ampersand, open-paren, close-paren, comma, hyphen, period, slash.
          Case 13&, 10&
            ' ** Carriage-Return, Line-Feed.
            varRetVal = "CRLF"
            Exit For
          Case Else
            Debug.Print "'" & lngTmp03
          End Select
        End If
      Next


    End If
  End If

  OddChar = varRetVal

End Function

Public Function Glen_Qry_Copy() As Boolean

  Const THIS_PROC As String = "Glen_Qry_Copy"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngSteps As Long, strThisStep As String
  Dim strQryName As String, strSQL As String, strDesc As String
  Dim strNewBase As String, strOldSet As String, strNewSet As String
  Dim lngQrysCreated As Long
  Dim intPos1 As Integer
  Dim strTmp01 As String, strTmp02 As String, lngTmp03 As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean, blnSkip As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 5  ' ** Array's first-element UBound().
  Const Q_QNAM1 As Integer = 0
  Const Q_SQL1  As Integer = 1
  Const Q_DSC1  As Integer = 2
  Const Q_QNAM2 As Integer = 3
  Const Q_SQL2  As Integer = 4
  Const Q_DSC2  As Integer = 5

  Const QRY_BASE As String = "zzz_qry_Glenwood_yAsset_Trans_18_59_"

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  strNewBase = "20"
  lngSteps = 9&
  strOldSet = Left(Right(QRY_BASE, 3), 2)
  strNewSet = "36"

  Set dbs = CurrentDb
  With dbs

    lngQrys = 0&
    ReDim arr_varQry(Q_ELEMS, 0)

    strThisStep = vbNullString
    For lngX = 1& To lngSteps

      lngSteps = lngSteps + 1&
      strThisStep = Left(Right(QRY_BASE, 3), 2)
      If lngX > 1& Then
        strThisStep = Right("00" & CStr(Val(strThisStep) + (lngX - 1)), 2)
      End If

      strTmp01 = Left(QRY_BASE, 33) & strThisStep & "_"
      For Each qdf In .QueryDefs
        With qdf
          If Left(.Name, Len(strTmp01)) = strTmp01 Then
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_QNAM1, lngE) = .Name
            arr_varQry(Q_SQL1, lngE) = .SQL
            arr_varQry(Q_DSC1, lngE) = .Properties("Description")
            arr_varQry(Q_QNAM2, lngE) = Null
            arr_varQry(Q_SQL2, lngE) = Null
            arr_varQry(Q_DSC2, lngE) = Null
          End If
        End With
      Next  ' ** qdf.
      Set qdf = Nothing
    Next  ' ** lngX.

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    'For lngX = 0& To (lngQrys - 1&)
    '  Debug.Print "'" & arr_varQry(Q_QNAM1, lngX)
    '  DoEvents
    'Next

    For lngX = 0& To (lngQrys - 1&)
      strQryName = arr_varQry(Q_QNAM1, lngX)
      strQryName = Left(strQryName, 30) & strNewBase & Mid(strQryName, 33)
      If strNewSet <> vbNullString Then
        strTmp01 = Mid(strQryName, 34, 2)
        If strTmp01 = strOldSet Then
          strQryName = Left(strQryName, 33) & strNewSet & Mid(strQryName, 36)
        Else
          lngTmp03 = Val(strTmp01) - Val(strOldSet)
          strTmp02 = Right("00" & CStr(Val(strNewSet) + lngTmp03), 2)
          strQryName = Left(strQryName, 33) & strTmp02 & Mid(strQryName, 36)
        End If
      End If
      arr_varQry(Q_QNAM2, lngX) = strQryName
      strSQL = arr_varQry(Q_SQL1, lngX)
      If strNewSet <> vbNullString Then
'zzz_qry_Glenwood_yAsset_Trans_18_03
'zzz_qry_Glenwood_yAsset_Trans_20_03
        strTmp01 = "Trans_18_03"
        strTmp02 = "Trans_20_03"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_58
'zzz_qry_Glenwood_yAsset_Trans_20_35
        strTmp01 = "Trans_18_58"
        strTmp02 = "Trans_20_35"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_59
'zzz_qry_Glenwood_yAsset_Trans_20_36
        strTmp01 = "Trans_18_59"
        strTmp02 = "Trans_20_36"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_60
'zzz_qry_Glenwood_yAsset_Trans_20_37
        strTmp01 = "Trans_18_60"
        strTmp02 = "Trans_20_37"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_61
'zzz_qry_Glenwood_yAsset_Trans_20_38
        strTmp01 = "Trans_18_61"
        strTmp02 = "Trans_20_38"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_62
'zzz_qry_Glenwood_yAsset_Trans_20_39
        strTmp01 = "Trans_18_62"
        strTmp02 = "Trans_20_39"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_63
'zzz_qry_Glenwood_yAsset_Trans_20_40
        strTmp01 = "Trans_18_63"
        strTmp02 = "Trans_20_40"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_64
'zzz_qry_Glenwood_yAsset_Trans_20_41
        strTmp01 = "Trans_18_64"
        strTmp02 = "Trans_20_41"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_65
'zzz_qry_Glenwood_yAsset_Trans_20_42
        strTmp01 = "Trans_18_65"
        strTmp02 = "Trans_20_42"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_66
'zzz_qry_Glenwood_yAsset_Trans_20_43
        strTmp01 = "Trans_18_66"
        strTmp02 = "Trans_20_43"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'zzz_qry_Glenwood_yAsset_Trans_18_67
'zzz_qry_Glenwood_yAsset_Trans_20_44
        strTmp01 = "Trans_18_67"
        strTmp02 = "Trans_20_44"
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
      Else
        strTmp01 = Left(arr_varQry(Q_QNAM1, lngX), 33)
        strTmp02 = Left(arr_varQry(Q_QNAM2, lngX), 33)
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
      End If
      arr_varQry(Q_SQL2, lngX) = strSQL
      strDesc = arr_varQry(Q_DSC1, lngX)
      ' ** .._18_10_01, linked back to .._18_03, with gtt03_id1, gtt03_id2; 106.
      If strNewSet <> vbNullString Then
'.._18_03
'.._20_03
        strTmp01 = ".._18_03"
        strTmp02 = ".._20_03"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_58
'.._20_35
        strTmp01 = ".._18_58"
        strTmp02 = ".._20_35"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_59
'.._20_36
        strTmp01 = ".._18_59"
        strTmp02 = ".._20_36"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_60
'.._20_37
        strTmp01 = ".._18_60"
        strTmp02 = ".._20_37"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_61
'.._20_38
        strTmp01 = ".._18_61"
        strTmp02 = ".._20_38"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_62
'.._20_39
        strTmp01 = ".._18_62"
        strTmp02 = ".._20_39"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_63
'.._20_40
        strTmp01 = ".._18_63"
        strTmp02 = ".._20_40"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_64
'.._20_41
        strTmp01 = ".._18_64"
        strTmp02 = ".._20_41"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_65
'.._20_42
        strTmp01 = ".._18_65"
        strTmp02 = ".._20_42"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_66
'.._20_43
        strTmp01 = ".._18_66"
        strTmp02 = ".._20_43"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
'.._18_67
'.._20_44
        strTmp01 = ".._18_67"
        strTmp02 = ".._20_44"
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
      Else
        strTmp01 = ".." & Right(strTmp01, 4)
        strTmp02 = ".." & Right(strTmp02, 4)
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Funcstion: modStringFuncs.
      End If
      intPos1 = InStr(strDesc, ";")
      If intPos1 > 0 Then strDesc = Left(strDesc, intPos1) & " "
      arr_varQry(Q_DSC2, lngX) = strDesc
    Next  ' ** lngX.

    'For lngX = 0& To (lngQrys - 1&)
    '  Debug.Print "'" & Trim(arr_varQry(Q_DSC2, lngX))
    '  Debug.Print "'" & arr_varQry(Q_QNAM2, lngX)
    '  DoEvents
    'Next

    lngQrysCreated = 0&

    blnSkip = False
    If blnSkip = False Then
      For lngX = 0& To (lngQrys - 1&)
        Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngX), arr_varQry(Q_SQL2, lngX))
        With qdf
          Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC2, lngX))
On Error Resume Next
          .Properties.Append prp
          If ERR.Number <> 0 Then
On Error GoTo 0
            .Properties("Description") = arr_varQry(Q_DSC2, lngX)
          Else
On Error GoTo 0
          End If
        End With
        Set qdf = Nothing
        lngQrysCreated = lngQrysCreated + 1&
      Next
    End If  ' ** blnSkip.

    Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
    DoEvents

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing


'QRYS: 29
'.._20_41_01, linked back to .._20_03, with gtt03_id1, gtt03_id7;
'zzz_qry_Glenwood_yAsset_Trans_20_28_01
'.._20_41_02, linked back to .._20_03, with gatt02_id1, gatt02_id7;
'zzz_qry_Glenwood_yAsset_Trans_20_28_02
'.._20_03, linked to .._20_42_01, with gtt03_id2, gtt03_id6;
'zzz_qry_Glenwood_yAsset_Trans_20_29_01
'.._20_03, linked to .._20_42_02, with gatt02_id2, gatt02_id6;
'zzz_qry_Glenwood_yAsset_Trans_20_29_02
'.._20_03, linked to .._20_43_01, with gtt03_id3, gtt03_id5;
'zzz_qry_Glenwood_yAsset_Trans_20_30_01
'.._20_03, linked to .._20_43_02, with gatt02_id3, gatt02_id5;
'zzz_qry_Glenwood_yAsset_Trans_20_30_02
'.._20_03, linked to .._20_44_01, with gtt03_id4;
'zzz_qry_Glenwood_yAsset_Trans_20_31_01
'.._20_03, linked to .._20_44_02, with gatt02_id4;
'zzz_qry_Glenwood_yAsset_Trans_20_31_02
'.._20_45_01, linked to .._20_45_02, just cnt = 7, by gatt02_id1, gtt03_id1;
'zzz_qry_Glenwood_yAsset_Trans_20_32_01
'.._20_45_01, linked to .._20_45_02, just cnt = 7, by gatt02_id2, gtt03_id2;
'zzz_qry_Glenwood_yAsset_Trans_20_32_02
'.._20_45_01, linked to .._20_45_02, just cnt = 7, by gatt02_id3, gtt03_id3;
'zzz_qry_Glenwood_yAsset_Trans_20_32_03
'.._20_45_01, linked to .._20_45_02, just cnt = 7, by gatt02_id4, gtt03_id4;
'zzz_qry_Glenwood_yAsset_Trans_20_32_04
'.._20_45_01, linked to .._20_45_02, just cnt = 7, by gatt02_id5, gtt03_id5;
'zzz_qry_Glenwood_yAsset_Trans_20_32_05
'.._20_45_01, linked to .._20_45_02, just cnt = 7, by gatt02_id6, gtt03_id6;
'zzz_qry_Glenwood_yAsset_Trans_20_32_06
'.._20_45_01, linked to .._20_45_02, just cnt = 7, by gatt02_id7, gtt03_id7;
'zzz_qry_Glenwood_yAsset_Trans_20_32_07
'.._20_46_01, linked back to .._20_03;
'zzz_qry_Glenwood_yAsset_Trans_20_33_01
'.._20_46_02, linked back to .._20_03;
'zzz_qry_Glenwood_yAsset_Trans_20_33_02
'.._20_46_03, linked back to .._20_03;
'zzz_qry_Glenwood_yAsset_Trans_20_33_03
'.._20_46_04, linked back to .._20_03;
'zzz_qry_Glenwood_yAsset_Trans_20_33_04
'.._20_46_05, linked back to .._20_03;
'zzz_qry_Glenwood_yAsset_Trans_20_33_05
'.._20_46_06, linked back to .._20_03;
'zzz_qry_Glenwood_yAsset_Trans_20_33_06
'.._20_46_07, linked back to .._20_03;
'zzz_qry_Glenwood_yAsset_Trans_20_33_07
'Append .._20_47_01 to zz_tbl_Glenwood_Asset_Transaction_03.
'zzz_qry_Glenwood_yAsset_Trans_20_34_01
'Append .._20_47_02 to zz_tbl_Glenwood_Asset_Transaction_03.
'zzz_qry_Glenwood_yAsset_Trans_20_34_02
'Append .._20_47_03 to zz_tbl_Glenwood_Asset_Transaction_03.
'zzz_qry_Glenwood_yAsset_Trans_20_34_03
'Append .._20_47_04 to zz_tbl_Glenwood_Asset_Transaction_03.
'zzz_qry_Glenwood_yAsset_Trans_20_34_04
'Append .._20_47_05 to zz_tbl_Glenwood_Asset_Transaction_03.
'zzz_qry_Glenwood_yAsset_Trans_20_34_05
'Append .._20_47_06 to zz_tbl_Glenwood_Asset_Transaction_03.
'zzz_qry_Glenwood_yAsset_Trans_20_34_06
'Append .._20_47_07 to zz_tbl_Glenwood_Asset_Transaction_03.
'zzz_qry_Glenwood_yAsset_Trans_20_34_07
'QRYS CREATED: 0
'DONE!

  Beep

  Debug.Print "'DONE!"
  DoEvents

  Set prp = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Glen_Qry_Copy = blnRetVal

End Function

Public Function Glen_UpdateQry() As Boolean

  Const THIS_PROC As String = "Glen_UpdateQry"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
  Dim strQryName1 As String, strQryName2 As String, strSQL As String, strDesc As String
  Dim lngQrys As Long, lngQrysUpdated As Long
  Dim strTmp01 As String, strTmp02 As String
  Dim lngX As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    lngQrys = 24&
    lngQrysUpdated = 0&
    For lngX = 2& To lngQrys
      'zzz_qry_Glenwood_yAsset_Trans_20_60_01
      'zzz_qry_Glenwood_yAsset_Trans_20_60_24
      strQryName1 = "zzz_qry_Glenwood_yAsset_Trans_20_60_01"
      strQryName2 = strQryName1
      strQryName2 = Left(strQryName2, (Len(strQryName2) - 2))
      strQryName2 = strQryName2 & Right("00" & CStr(lngX), 2)
      DoCmd.CopyObject , strQryName2, acQuery, strQryName1
      .QueryDefs.Refresh
      DoEvents
      Set qdf = .QueryDefs(strQryName2)
      With qdf
        strSQL = .SQL
        strTmp01 = "zzz_qry_Glenwood_yAsset_Trans_20_59_01"
        strTmp02 = "zzz_qry_Glenwood_yAsset_Trans_20_59_"
        strTmp02 = strTmp02 & Right("00" & CStr(lngX), 2)
        strSQL = StringReplace(strSQL, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
        .SQL = strSQL
        strDesc = .Properties("Description")
        strTmp01 = ".." & Right(strTmp01, 9)
        strTmp02 = ".." & Right(strTmp02, 9)
        strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
        .Properties("Description") = strDesc
      End With
      Set qdf = Nothing
      'strTmp01 = "zzz_qry_Glenwood_yAsset_Trans_20_58_02"
      'strTmp02 = "zzz_qry_Glenwood_yAsset_Trans_20_58_"
      'strTmp02 = strTmp02 & Right("00" & CStr(lngX), 2)
      'blnRetVal = Qry_UpdateRef_rel(strTmp01, strTmp02, False, strQryName)  ' ** Module Function: modQueryFunctions.
      'If blnRetVal = False Then
      '  Stop
      'End If
      'strTmp01 = ".." & Right(strTmp01, 9)
      'strTmp02 = ".." & Right(strTmp02, 9)
      'strDesc = .QueryDefs(strQryName).Properties("Description")
      'strDesc = StringReplace(strDesc, strTmp01, strTmp02)  ' ** Module Function: modStringFuncs.
      '.QueryDefs(strQryName).Properties("Description") = strDesc
      lngQrysUpdated = lngQrysUpdated + 1&
    Next

    .Close
  End With
  Set dbs = Nothing

  Debug.Print "'QRYS UPDATED: " & CStr(lngQrysUpdated)
  DoEvents

  Beep

  Debug.Print "'DONE!"
  DoEvents

  Set prp = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Glen_UpdateQry = blnRetVal

End Function

Public Function Rpt_Ctl_Lbls() As Boolean

  Const THIS_PROC As String = "Rpt_Ctl_Lbls"

  Dim rpt As Access.Report, ctl As Access.Control
  Dim strCtlName As String
  Dim blnRetVal As Boolean

  blnRetVal = True

  Set rpt = Reports(0)
  With rpt
    For Each ctl In .Controls
      With ctl
        If .ControlType = acLabel Then
          strCtlName = .Parent.Name
          If strCtlName <> rpt.Name Then
            .Name = strCtlName & "_lbl"
          End If
        End If
      End With
    Next
  End With

  Beep

  'Debug.Print "'DONE!"
  DoEvents

  Set ctl = Nothing
  Set rpt = Nothing

  Rpt_Ctl_Lbls = blnRetVal

End Function

Public Function Glen_Rpt_Copy() As Boolean

  Const THIS_PROC As String = "Glen_Rpt_Copy"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, ctr As DAO.Container, doc As DAO.Document, prp As Object, rpt As Access.Report
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim strQryNum As String, strSQL As String, strDesc As String
  Dim strQrySource As String, strRptSource As String, strLastRpt As String
  Dim lngQrysCreated As Long, lngRptsCreated As Long
  Dim blnSkip As Boolean
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, lngTmp04 As Long
  Dim lngX As Long, lngY As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const Q_QNAM1 As Integer = 0
  Const Q_QNAM2 As Integer = 1
  Const Q_SQL2  As Integer = 2
  Const Q_DSC2  As Integer = 3
  Const Q_SUB   As Integer = 4

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const R_RNAM1  As Integer = 0
  Const R_RNAM2  As Integer = 1
  Const R_RECSRC As Integer = 2

  Const QRY_BASE As String = "zzz_qry_Glenwood_yTrans_03_11_"

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  strQrySource = "zzz_qry_Glenwood_yTrans_03_10_04_01"
  strRptSource = "zz_rptGlenwood_Sold_01"

  Set dbs = CurrentDb
  With dbs

    lngQrys = 0&
    ReDim arr_varQry(Q_ELEMS, 0)

    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
          strQryNum = Right(.Name, 2)
          If Val(strQryNum) >= 3 And Val(strQryNum) <= 21 Then
            ' ** zzz_qry_Glenwood_yTrans_03_11_03 - zzz_qry_Glenwood_yTrans_03_11_21.
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_QNAM1, lngE) = .Name
            arr_varQry(Q_QNAM2, lngE) = .Name & "_01"
            arr_varQry(Q_SQL2, lngE) = Null
            arr_varQry(Q_DSC2, lngE) = Null
            arr_varQry(Q_SUB, lngE) = lngE + 1&
          End If
        End If
      End With  ' ** qdf.
    Next  ' ** qdf.
    Set qdf = Nothing

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    If lngQrys > 0& Then

      Set qdf = .QueryDefs(strQrySource)
      With qdf
        strSQL = .SQL
        strDesc = .Properties("Description")
      End With
      Set qdf = Nothing

      For lngX = 0& To (lngQrys - 1&)
        arr_varQry(Q_SQL2, lngX) = strSQL
        arr_varQry(Q_DSC2, lngX) = strDesc
      Next  ' ** lngX.

      strTmp01 = "Trans_03_10_04"
      strTmp02 = ".._03_10_04"
      For lngX = 0& To (lngQrys - 1&)
        strTmp03 = Right(arr_varQry(Q_QNAM1, lngX), Len(strTmp01))
        strSQL = arr_varQry(Q_SQL2, lngX)
        strSQL = StringReplace(strSQL, strTmp01, strTmp03)  ' ** Module Function: modStringFuncs.
        arr_varQry(Q_SQL2, lngX) = strSQL
        strTmp03 = ".." & Right(strTmp03, 9)
        strDesc = arr_varQry(Q_DSC2, lngX)
        strDesc = StringReplace(strDesc, strTmp02, strTmp03)  ' ** Module Function: modStringFuncs.
        arr_varQry(Q_DSC2, lngX) = strDesc
      Next  ' ** lngX.

      blnSkip = False
      If blnSkip = False Then
        lngQrysCreated = 0&
        For lngX = 0& To (lngQrys - 1&)
          Set qdf = .CreateQueryDef(arr_varQry(Q_QNAM2, lngX), arr_varQry(Q_SQL2, lngX))
          With qdf
            Set prp = .CreateProperty("Description", dbText, arr_varQry(Q_DSC2, lngX))
On Error Resume Next
            .Properties.Append prp
            If ERR.Number <> 0 Then
On Error GoTo 0
              .Properties("Description") = arr_varQry(Q_DSC2, lngX)
            Else
On Error GoTo 0
            End If
          End With  ' ** qdf.
          Set qdf = Nothing
          lngQrysCreated = lngQrysCreated + 1&
        Next  ' ** lngX.
      End If  ' ** blnSkip.

      Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
      DoEvents

    End If  ' ** lngQrys.
    
    lngRpts = 0&
    ReDim arr_varRpt(R_ELEMS, 0)

    If lngQrys > 0& Then

      Set ctr = .Containers("Reports")
      For Each doc In ctr.Documents
        With doc
          If Left(.Name, Len(strRptSource)) = strRptSource Then
            lngRpts = lngRpts + 1&
            lngE = lngRpts - 1&
            ReDim Preserve arr_varRpt(R_ELEMS, lngE)
            arr_varRpt(R_RNAM1, lngE) = .Name
            arr_varRpt(R_RNAM2, lngE) = Null
            arr_varRpt(R_RECSRC, lngE) = Null
          End If
        End With  ' ** doc
      Next  ' ** doc.
      Set doc = Nothing
      Set ctr = Nothing

    End If  ' ** lngQrys.

    Debug.Print "'RPTS: " & CStr(lngRpts)
    DoEvents

    If lngRpts > 0& Then

      strLastRpt = arr_varRpt(R_RNAM1, (lngRpts - 1&))

      lngRpts = lngRpts + 1&
      lngE = lngRpts - 1&
      ReDim Preserve arr_varRpt(R_ELEMS, lngE)
      strTmp01 = Right(strLastRpt, 2)
      strTmp01 = CStr(Val(strTmp01) + 1)
      strTmp02 = Left(strLastRpt, (Len(strLastRpt) - 2))
      strTmp02 = strTmp02 + strTmp01
      arr_varRpt(R_RNAM1, lngE) = strTmp02
      arr_varRpt(R_RNAM2, lngE) = Null
      arr_varRpt(R_RECSRC, lngE) = Null
      lngRpts = lngRpts + 1&
      lngE = lngRpts - 1&
      ReDim Preserve arr_varRpt(R_ELEMS, lngE)
      strTmp01 = Right(strLastRpt, 2)
      strTmp01 = CStr(Val(strTmp01) + 2)
      strTmp02 = Left(strLastRpt, (Len(strLastRpt) - 2))
      strTmp02 = strTmp02 + strTmp01
      arr_varRpt(R_RNAM1, lngE) = strTmp02
      arr_varRpt(R_RNAM2, lngE) = Null
      arr_varRpt(R_RECSRC, lngE) = Null

      For lngX = 0& To (lngRpts - 1&)
        strTmp01 = arr_varRpt(R_RNAM1, lngX)
        strTmp01 = StringReplace(strTmp01, "Sold", "Misc")  ' ** Module Function: modStringFuncs.
        arr_varRpt(R_RNAM2, lngX) = strTmp01
        If InStr(strTmp01, "Sub") > 0 Then
          strTmp02 = Right(strTmp01, 2)
          lngTmp04 = Val(strTmp02)
          For lngY = 0& To (lngQrys - 1&)
            If arr_varQry(Q_SUB, lngY) = lngTmp04 Then
              arr_varRpt(R_RECSRC, lngX) = arr_varQry(Q_QNAM2, lngY)
              Exit For
            End If
          Next  ' ** lngY.
        End If
      Next  ' ** lngX.

      For lngX = 0& To (lngRpts - 1&)
        If Val(Right(arr_varRpt(R_RNAM1, lngX), 2)) > Val(Right(strLastRpt, 2)) Then
          arr_varRpt(R_RNAM1, lngX) = strLastRpt
        End If
      Next  ' ** lngX.

      lngRptsCreated = 0&
      For lngX = 0& To (lngRpts - 1&)
        DoCmd.CopyObject , arr_varRpt(R_RNAM2, lngX), acReport, arr_varRpt(R_RNAM1, lngX)
        DoEvents
        lngRptsCreated = lngRptsCreated + 1&
      Next  ' ** lngX.

      .Containers("Reports").Documents.Refresh
      DoEvents

      Debug.Print "'RPTS CREATED: " & CStr(lngRptsCreated)
      DoEvents

    End If  ' ** lngRpts.

    If lngRptsCreated > 0& Then
      For lngX = 0& To (lngRpts - 1&)
        If InStr(arr_varRpt(R_RNAM2, lngX), "Sub") > 0 Then
          DoCmd.OpenReport arr_varRpt(R_RNAM2, lngX), acViewDesign, , , acHidden
          Set rpt = Reports(0)
          With rpt
            .RecordSource = arr_varRpt(R_RECSRC, lngX)
          End With
          Set rpt = Nothing
          DoCmd.Close acReport, Reports(0).Name, acSaveYes
        End If
      Next  ' ** lngX.
    End If  ' ** lngRptsCreated.

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set rpt = Nothing
  Set prp = Nothing
  Set doc = Nothing
  Set ctr = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Glen_Rpt_Copy = blnRetVal

End Function

Public Function Glen_Qry_Desc() As Boolean

  Const THIS_PROC As String = "Glen_Qry_Desc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rpt As Access.Report, ctl As Access.Label
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim strDesc As String
  Dim intPos1 As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 4  ' ** Array's first-element UBound().
  Const Q_QNAM1 As Integer = 0
  Const Q_QNAM2 As Integer = 1
  Const Q_DSC1  As Integer = 2
  Const Q_DSC2  As Integer = 3
  Const Q_CNAM  As Integer = 4
  
  Const QRY_BASE As String = "zzz_qry_Glenwood_yTrans_03_11"

  blnRetVal = True

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs
    For Each qdf In .QueryDefs
      With qdf
        ' ** zzz_qry_Glenwood_yTrans_03_11_03_01
        If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
          If Len(.Name) = 35 Then
            lngQrys = lngQrys + 1&
            lngE = lngQrys - 1&
            ReDim Preserve arr_varQry(Q_ELEMS, lngE)
            arr_varQry(Q_QNAM1, lngE) = Null
            arr_varQry(Q_QNAM2, lngE) = .Name
            arr_varQry(Q_DSC1, lngE) = Null
            arr_varQry(Q_DSC2, lngE) = Null
            arr_varQry(Q_CNAM, lngE) = Null
          End If
        End If
      End With
    Next
    Set qdf = Nothing

    Debug.Print "'QRYS: " & CStr(lngQrys)
    DoEvents

    If lngQrys > 0& Then

      For lngX = 0& To (lngQrys - 1&)
        strTmp01 = arr_varQry(Q_QNAM2, lngX)
        strTmp01 = Left(strTmp01, (Len(strTmp01) - 3))
        arr_varQry(Q_QNAM1, lngX) = strTmp01
      Next  ' ** lngX.

      For lngX = 0& To (lngQrys - 1&)
        Set qdf = .QueryDefs(arr_varQry(Q_QNAM1, lngX))
        With qdf
          arr_varQry(Q_DSC1, lngX) = .Properties("Description")
        End With
      Next  ' ** lngX.
      Set qdf = Nothing

      For lngX = 0& To (lngQrys - 1&)
        ' ** .._03_11_02, just gt_units < 0, cash = 0, cost > 0; 1.  WHAT?
        strTmp01 = arr_varQry(Q_DSC1, lngX)
        intPos1 = InStr(strTmp01, "just")
        strTmp02 = Mid(strTmp01, (intPos1 + 8))
        intPos1 = InStr(strTmp02, ";")
        strTmp03 = Mid(strTmp02, intPos1)
        strTmp02 = Left(strTmp02, (intPos1 - 1))
        intPos1 = InStr(strTmp03, ".")
        If intPos1 > 0 Then
          strTmp03 = Trim(Mid(strTmp03, (intPos1 + 1)))
        End If
        strTmp02 = FormatProperCase(strTmp02)  ' ** Module Function: modStringFuncs.
        strTmp02 = StringReplace(strTmp02, " And ", " axd ")  ' ** Module Function: modStringFuncs.
        strTmp02 = StringReplace(strTmp02, " axd ", " and ")  ' ** Module Function: modStringFuncs.
        Select Case (lngX + 1&)
        Case 6&, 16&
'6   "and > 0"  "and Cash > 0"
'16  "and > 0"  "and Cash > 0"
          strTmp02 = StringReplace(strTmp02, "and > 0", "and Cash > 0")  ' ** Module Function: modStringFuncs.
        Case 3&, 10&, 15&, 17&
'3   "and > 0"  "and Cost > 0"
'10  "and > 0"  "and Cost > 0"
'15  "and > 0"  "and Cost > 0"
'17  "and > 0"  "and Cost > 0"
          strTmp02 = StringReplace(strTmp02, "and > 0", "and Cost > 0")  ' ** Module Function: modStringFuncs.
        End Select
        arr_varQry(Q_DSC2, lngX) = strTmp02 & Space(8) & strTmp03
      Next  ' ** lngX.

      For lngX = 0& To (lngQrys - 1&)
        strTmp01 = "zz_rptGlenwood_Misc_01_Sub_01_lbl"
        intPos1 = InStr(strTmp01, "Sub")
        strTmp02 = Right("00" & CStr(lngX + 1&), 2)
        strTmp03 = Left(strTmp01, (intPos1 + 3)) & strTmp02 & Right(strTmp01, 4)
        arr_varQry(Q_CNAM, lngX) = strTmp03
        strTmp01 = arr_varQry(Q_DSC2, lngX)
        strTmp02 = CStr(lngX + 1&) & ".  "
        arr_varQry(Q_DSC2, lngX) = strTmp02 & strTmp01
      Next  ' ** lngX.

      'For lngX = 0& To (lngQrys - 1&)
      '  'Debug.Print "'" & arr_varQry(Q_DSC2, lngX)
      '  Debug.Print "'" & arr_varQry(Q_CNAM, lngX)
      '  DoEvents
      'Next  ' ** lngX.

    End If  ' ** lngQrys.

    .Close
  End With
  Set dbs = Nothing

  Set rpt = Reports(0)
  With rpt
    For lngX = 0& To (lngQrys - 1&)
      Set ctl = .Controls(arr_varQry(Q_CNAM, lngX))
      ctl.Caption = arr_varQry(Q_DSC2, lngX)
    Next  ' ** lngX.
  End With
  Set rpt = Nothing

'QRYS: 19
'Units < 0, Cash = 0, Cost < 0        WHAT?
'Units < 0, Cash = 0, Cost > 0        WHAT?
'Units < 0, Cash = 0, Cost < 0 and Cost > 0        WHAT?
'Units < 0, Cash > 0, Cost < 0        SOLD?
'Units < 0, Cash < 0, Cost < 0        WHAT?
'Units < 0, Cash < 0 and Cash > 0, Cost = 0        WHAT?
'Units > 0, Cash = 0, Cost = 0        DEPOSIT WITH NO COST?
'Units > 0, Cash = 0, Cost > 0        DEPOSIT?
'Units > 0, Cash = 0, Cost < 0        WHAT?
'Units > 0, Cash = 0, Cost < 0 and Cost > 0        WHAT?
'Units > 0, Cash > 0, Cost > 0        WHAT?
'Units > 0, Cash > 0, Cost < 0        WHAT?
'Units > 0, Cash < 0, Cost > 0        WHAT?
'Units = 0, Cash = 0, Cost > 0        WHAT?
'Units = 0, Cash = 0, Cost < 0 and Cost > 0        WHAT?
'Units = 0, Cash < 0 and Cash > 0, Cost < 0        WHAT?
'Units = 0, Cash < 0 and Cash > 0, Cost < 0 and Cost > 0        WHAT?
'Units = 0, Cash = 0, Cost < 0        COST ADJ.?
'Units = 0, Cash < 0, Cost > 0        PURCHASE BUT ADD UNITS?
'DONE!

  Debug.Print "'DONE!"
  DoEvents

  Beep

  Set ctl = Nothing
  Set rpt = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  Glen_Qry_Desc = blnRetVal

End Function

Public Function AutoNum_Holes() As Boolean

On Error GoTo ERRH

  Const THIS_PROC As String = "AutoNum_Holes"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
  Dim lngNums As Long, arr_varNum() As Variant
  Dim lngRecs As Long, lngLastAutoNum As Long
  Dim strTableName As String, strFieldName As String
  Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
  Dim blnRetVal As Boolean 'varRetVal As Variant

On Error GoTo 0

  blnRetVal = True
  'varRetVal = Empty

  strTableName = "zz_tbl_Glenwood_Transaction_04"
  strFieldName = "gtt04_id"

  lngNums = 0&
  ReDim arr_varNum(0)
  lngLastAutoNum = 0&

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    ' ** Empty tblMark.
    Set qdf = .QueryDefs("qryTmp_Table_Empty_079_tblMark_AutoNum2")
    qdf.Execute
    Set qdf = Nothing

    Set rst1 = .OpenRecordset(strTableName, dbOpenDynaset, dbConsistent)
    With rst1
      .MoveLast
      lngRecs = .RecordCount
      .MoveFirst
      .Sort = "[" & strFieldName & "]"
      Set rst2 = rst1.OpenRecordset
      With rst2
        For lngX = 1& To lngRecs
          If .Fields(strFieldName) <> lngLastAutoNum + 1& Then
            lngNums = lngNums + 1&
            lngE = lngNums - 1&
            ReDim Preserve arr_varNum(lngE)
            lngLastAutoNum = lngLastAutoNum + 1&
            arr_varNum(lngE) = lngLastAutoNum
            ' ** Add missing AutoNum's till we catch up.
            lngZ = lngLastAutoNum + 1&
            If lngZ < .Fields(strFieldName) Then
              For lngY = lngZ To (.Fields(strFieldName) - 1&)
                lngNums = lngNums + 1&
                lngE = lngNums - 1&
                ReDim Preserve arr_varNum(lngE)
                lngLastAutoNum = lngLastAutoNum + 1&
                arr_varNum(lngE) = lngLastAutoNum
              Next
              lngLastAutoNum = .Fields(strFieldName)
            Else
              ' ** Ready to continue.
              lngLastAutoNum = .Fields(strFieldName)
            End If
          Else
            lngLastAutoNum = .Fields(strFieldName)
          End If
          If lngX < lngRecs Then .MoveNext
        Next
        .Close
      End With
      Set rst2 = Nothing
      .Close
    End With
    Set rst1 = Nothing

    Debug.Print "'HOLES: " & CStr(lngNums)
    DoEvents

    If lngNums > 0& Then
      Set rst1 = .OpenRecordset("tblMark_AutoNum2", dbOpenDynaset, dbConsistent)
      With rst1
        For lngX = 0& To (lngNums - 1&)
          .AddNew
          ![unique_id] = arr_varNum(lngX)
          ![mark] = False
          '![value_lng] =
          '![value_dbl] =
          '![value_txt] =
          ' ** ![autonum_id] : AutoNumber.
          .Update
        Next
        .Close
      End With
    End If

    .Close
  End With

  'varRetVal = arr_varNum

  Debug.Print "'DONE!"

  Beep

EXITP:
  Set rst2 = Nothing
  Set rst1 = Nothing
  Set qdf = Nothing
  Set dbs = Nothing
  'AutoNum_Holes = varRetVal
  AutoNum_Holes = blnRetVal
  Exit Function

ERRH:
  blnRetVal = False
  Select Case ERR.Number
  Case Else
    zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
  End Select
  Resume EXITP

End Function

Public Function Qry_FindDesc() As Boolean

  Const THIS_PROC As String = "Qry_FindDesc"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim strFind As String, strDesc As String
  Dim lngQrys As Long
  Dim intPos1 As Integer
  Dim blnRetVal As Boolean

  Const QRY_BASE As String = "zzz_qry_Glenwood_yTrans_"

  blnRetVal = True

  strFind = "kill"

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngQrys = 0&
  Set dbs = CurrentDb
  With dbs
    For Each qdf In .QueryDefs
      strDesc = vbNullString
      With qdf
        If Left(.Name, Len(QRY_BASE)) = QRY_BASE Then
On Error Resume Next
          strDesc = .Properties("Description")
On Error GoTo 0
          intPos1 = InStr(strDesc, strFind)
          If intPos1 > 0 Then
            Debug.Print "'" & .Name & "  '" & Trim(strDesc) & "'"
            DoEvents
            lngQrys = lngQrys + 1&
          End If
        End If
      End With
    Next
    .Close
  End With

  Debug.Print "'QRYS FOUND: " & CStr(lngQrys)
  DoEvents

'zzz_qry_Glenwood_yTrans_03_05_05_03  '.._03_08_01, corresponding 'Cost Adj.', kill; 2.'
'zzz_qry_Glenwood_yTrans_03_05_05_04  '.._03_06_01, corresponding 'Sold', kill; 1.'
'zzz_qry_Glenwood_yTrans_03_06_05_02  '.._03_06_05, ZERO'S OUT, kill; 2.'
'zzz_qry_Glenwood_yTrans_03_06_10_03  '.._03_06_01, just gtt04_id = 31861, 31862, kill; 2.'
'zzz_qry_Glenwood_yTrans_03_06_19_03  '.._03_11_01, just matching 'Misc.' entries, kill; 6.'
'zzz_qry_Glenwood_yTrans_03_11_17_02  '.._03_11_17, ZERO'S OUT, kill; 135.'
'QRYS FOUND: 6
'DONE!
  Debug.Print "'DONE!"

  Beep

  Set qdf = Nothing
  Set dbs = Nothing

  Qry_FindDesc = blnRetVal

End Function

Public Function Glen_Tbl_Delete() As Boolean

  Const THIS_PROC As String = "Glen_Tbl_Delete"

  Dim dbs As DAO.Database, tdf As DAO.TableDef
  Dim lngTbls As Long, arr_varTbl() As Variant
  Dim blnDelete As Boolean, lngTblsDeleted As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varTbl().
  Const T_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const T_TNAM As Integer = 0
  Const T_DEL  As Integer = 1

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngTbls = 0&
  ReDim arr_varTbl(T_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    .TableDefs.Refresh
    DoEvents

    For Each tdf In .TableDefs
      With tdf
        If Left(.Name, 11) = "tblGlenwood" Or Left(.Name, 15) = "zz_tbl_Glenwood" Then
          lngTbls = lngTbls + 1&
          lngE = lngTbls - 1&
          ReDim Preserve arr_varTbl(T_ELEMS, lngE)
          arr_varTbl(T_TNAM, lngE) = .Name
          arr_varTbl(T_DEL, lngE) = CBool(False)
        End If
      End With
    Next
    Set tdf = Nothing

    Debug.Print "'TBLS: " & CStr(lngTbls)
    DoEvents

    If lngTbls > 0& Then

      lngTblsDeleted = 0&
      blnDelete = True

      Debug.Print "'DELETE THESE TABLES?"
      Stop

      If blnDelete = True Then
        For lngX = (lngTbls - 1&) To 0& Step -1&
          DoCmd.DeleteObject acTable, arr_varTbl(T_TNAM, lngX)
          arr_varTbl(T_DEL, lngX) = CBool(True)
          lngTblsDeleted = lngTblsDeleted + 1&
          DoEvents
        Next
        .TableDefs.Refresh
      End If

      Debug.Print "'TBLS DELETED: " & CStr(lngTblsDeleted)
      DoEvents

    Else
      Debug.Print "'NONE FOUND!"
      DoEvents
    End If

    .Close
  End With
  Set dbs = Nothing

  Debug.Print "'DONE!"

  Beep

  Set tdf = Nothing
  Set dbs = Nothing

  Glen_Tbl_Delete = blnRetVal

End Function

Public Function Glen_Rpt_Delete() As Boolean

  Const THIS_PROC As String = "Glen_Rpt_Delete"

  Dim dbs As DAO.Database, ctr As DAO.Container, doc As DAO.Document
  Dim lngRpts As Long, arr_varRpt() As Variant
  Dim blnDelete As Boolean, lngRptsDeleted As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varRpt().
  Const R_ELEMS As Integer = 1  ' ** Array's first-element UBound().
  Const R_RNAM As Integer = 0
  Const R_DEL  As Integer = 1

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  lngRpts = 0&
  ReDim arr_varRpt(R_ELEMS, 0)

  Set dbs = CurrentDb
  With dbs

    .Containers("Reports").Documents.Refresh
    DoEvents

    Set ctr = .Containers("Reports")
    With ctr
      For Each doc In .Documents
        With doc
          If Left(.Name, 14) = "zz_rptGlenwood" Then
            lngRpts = lngRpts + 1&
            lngE = lngRpts - 1&
            ReDim Preserve arr_varRpt(R_ELEMS, lngE)
            arr_varRpt(R_RNAM, lngE) = .Name
            arr_varRpt(R_DEL, lngE) = CBool(False)
          End If
        End With
      Next
      Set doc = Nothing
    End With
    Set ctr = Nothing

    Debug.Print "'RPTS: " & CStr(lngRpts)
    DoEvents

    If lngRpts > 0& Then

      lngRptsDeleted = 0&
      blnDelete = True

      Debug.Print "'DELETE THESE REPORTS?"
      Stop

      If blnDelete = True Then
        For lngX = (lngRpts - 1&) To 0& Step -1&
          DoCmd.DeleteObject acReport, arr_varRpt(R_RNAM, lngX)
          arr_varRpt(R_DEL, lngX) = CBool(True)
          lngRptsDeleted = lngRptsDeleted + 1&
          DoEvents
        Next
      End If

      Debug.Print "'RPTS DELETED: " & CStr(lngRptsDeleted)
      DoEvents

    Else
      Debug.Print "'NONE FOUND!"
      DoEvents
    End If

    .Close
  End With
  Set dbs = Nothing

  Debug.Print "'DONE!"

  Beep

  Set doc = Nothing
  Set ctr = Nothing
  Set dbs = Nothing

  Glen_Rpt_Delete = blnRetVal

End Function
