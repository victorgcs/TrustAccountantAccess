Attribute VB_Name = "zz_mod_FirstAlabamaFuncs"
Option Compare Database
Option Explicit

'VGC 08/03/2016: CHANGES!

Private Const THIS_NAME As String = "zz_mod_FirstAlabamaFuncs"
' **

Public Function FA_RegisterImport() As Boolean

  Const THIS_PROC As String = "FA_RegisterImport"

  Dim fso As Scripting.FileSystemObject, fsfds As Scripting.Folders, fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder
  Dim fsfls As Scripting.Files, fsfl As Scripting.File
  Dim dbs As DAO.Database, qdf As DAO.QueryDef
  Dim lngDirs As Long, arr_varDir() As Variant
  Dim lngFiles As Long, arr_varFile() As Variant
  Dim strPath1 As String, strPath2 As String, strFile As String, strPathFile As String
  Dim strTableName1 As String, strTableName2 As String
  Dim strSQL As String
  Dim lngTotFiles As Long, lngFilesImported As Long
  Dim blnSkip As Boolean
  Dim intPos01 As Integer, intCnt As Integer
  Dim strTmp01 As String, strTmp02 As String, arr_varTmp03 As Variant
  Dim lngW As Long, lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varDir().
  Const D_ELEMS As Integer = 2  ' ** Array's first-element UBound().
  Const D_DNAM  As Integer = 0
  Const D_FILS  As Integer = 1
  Const D_F_ARR As Integer = 2

  ' ** Array: arr_varFile().
  Const F_ELEMS As Integer = 1
  Const F_FNAM  As Integer = 0
  Const F_DSC   As Integer = 1

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  strTableName1 = "tblFirstAlabama_Register_tmp"
  strTableName2 = "tblFirstAlabama_Register_raw"

  strPath1 = "C:\VictorGCS_Clients\TrustAccountant\Clients\FirstBankAlabama"

  lngDirs = 0&
  ReDim arr_varDir(D_ELEMS, 0)

  Set fso = CreateObject("Scripting.FileSystemObject")
  With fso
    Set fsfd1 = .GetFolder(strPath1)
    With fsfd1
      Set fsfds = .SubFolders
      For Each fsfd2 In fsfds
        With fsfd2
          If InStr(.Name, "Trust") > 0 Or InStr(.Name, "Estate") > 0 Then
            lngDirs = lngDirs + 1&
            lngE = lngDirs - 1&
            ReDim Preserve arr_varDir(D_ELEMS, lngE)
            arr_varDir(D_DNAM, lngE) = .Name
          End If
        End With   ' ** fsfd2.
      Next  ' ** fsfd2.
      Set fsfd2 = Nothing
      Set fsfds = Nothing
    End With  ' ** fsfd1.
    Set fsfd1 = Nothing
  End With  ' ** fso.
  Set fso = Nothing

  Debug.Print "'DIRS: " & CStr(lngDirs)
  DoEvents

  If lngDirs > 0& Then

    lngTotFiles = 0&
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso
      For lngW = 0& To (lngDirs - 1&)
        lngFiles = 0&
        ReDim arr_varfiles(F_ELEMS, 0)
        strPath2 = strPath1 & LNK_SEP & arr_varDir(D_DNAM, lngW)
        Set fsfd1 = .GetFolder(strPath2)
        With fsfd1
          Set fsfls = .Files
          For Each fsfl In fsfls
            With fsfl
              If Left(.Name, 15) = "Register Report" Then
                lngFiles = lngFiles + 1&
                lngE = lngFiles - 1&
                ReDim Preserve arr_varFile(F_ELEMS, lngE)
                arr_varFile(F_FNAM, lngE) = .Name
                strTmp01 = .Name
                intCnt = CharCnt(strTmp01, ".")  ' ** Module Function: modStringFuncs.
                intPos01 = CharPos(strTmp01, CLng(intCnt), ".")  ' ** Module Function: modStringFuncs.
                strTmp01 = Left(strTmp01, (intPos01 - 1))
                arr_varFile(F_DSC, lngE) = strTmp01
                lngTotFiles = lngTotFiles + 1&
              End If
            End With  ' ** fsfl.
          Next  ' ** fsfl.
        End With  ' ** fsfd1.
        arr_varDir(D_FILS, lngW) = lngFiles
        arr_varDir(D_F_ARR, lngW) = arr_varFile
      Next  ' ** lngW.
      Set fsfl = Nothing
      Set fsfls = Nothing
      Set fsfd1 = Nothing
    End With  ' ** fso.
    Set fso = Nothing

    Debug.Print "'FILES: " & CStr(lngTotFiles)
    DoEvents

    If lngTotFiles > 0& Then

      For lngW = 0& To (lngDirs - 1&)
        lngFiles = arr_varDir(D_FILS, lngW)
        arr_varTmp03 = arr_varDir(D_F_ARR, lngW)
        For lngX = 0& To (lngFiles - 1&)
          Debug.Print "'" & arr_varTmp03(F_DSC, lngX)
          DoEvents
        Next  ' ** lngX.
      Next  ' ** lngW.

      blnSkip = True
      If blnSkip = False Then
        Set dbs = CurrentDb
        With dbs

          lngFilesImported = 0&
          For lngW = 0& To (lngDirs - 1&)
            strPath2 = strPath1 & LNK_SEP & arr_varDir(D_DNAM, lngW)
            lngFiles = arr_varDir(D_FILS, lngW)
            arr_varTmp03 = arr_varDir(D_F_ARR, lngW)
            For lngX = 0& To (lngFiles - 1&)
              strFile = arr_varTmp03(F_FNAM, lngX)
              If arr_varTmp03(F_DSC, lngX) = "Register Report Alcoa stock" Or _
                  arr_varTmp03(F_DSC, lngX) = "Register Report FNTC" Or _
                  arr_varTmp03(F_DSC, lngX) = "Register Report Home" Or _
                  arr_varTmp03(F_DSC, lngX) = "Register Report Household Furnishings" Then
                ' ** Already imported.
'                  arr_varTmp03(F_DSC, lngX) = "Register Report Cash Account" Or
              Else
                strPathFile = strPath2 & LNK_SEP & strFile
                If TableExists(strTableName1) = True Then
                  DoCmd.DeleteObject acTable, strTableName1
                  DoEvents
                  .TableDefs.Refresh
                End If
                DoEvents
                DoCmd.TransferText acImportDelim, "FirstAlabama_Register_Import_Specification", strTableName1, strPathFile
                DoEvents
                strSQL = vbNullString: strTmp01 = vbNullString: strTmp02 = vbNullString: intPos01 = 0
                Set qdf = .QueryDefs("zzz_qry_FirstAlabama_Register_01")
                With qdf
                  ' ** INSERT INTO tblFirstAlabama_Register_raw ( faregr_account, faregr_description1, Field01, Field02, Field03,
                  ' **   Field04, Field05, Field06, Field07, Field08, Field09, Field10, Field11, faregr_datemodified )
                  ' ** SELECT 'Barbara Sweat Estate' AS faregr_account, 'Register Report Household Furnishings' AS faregr_description1,
                  ' **   tblFirstAlabama_Register_tmp.Field01, tblFirstAlabama_Register_tmp.Field02, tblFirstAlabama_Register_tmp.Field03,
                  ' **   tblFirstAlabama_Register_tmp.Field04, tblFirstAlabama_Register_tmp.Field05, tblFirstAlabama_Register_tmp.Field06,
                  ' **   tblFirstAlabama_Register_tmp.Field07, tblFirstAlabama_Register_tmp.Field08, tblFirstAlabama_Register_tmp.Field09,
                  ' **   tblFirstAlabama_Register_tmp.Field10, tblFirstAlabama_Register_tmp.Field11, Now() AS faregr_datemodified
                  ' ** FROM tblFirstAlabama_Register_tmp
                  ' ** ORDER BY tblFirstAlabama_Register_tmp.ID;
                  strSQL = .SQL
                  intPos01 = InStr(strSQL, "SELECT '")
                  strTmp01 = Left(strSQL, (intPos01 + 7))
                  intPos01 = InStr(intPos01, strSQL, "' AS faregr_account")
                  strTmp02 = Mid(strSQL, intPos01)
                  strSQL = strTmp01 & arr_varDir(D_DNAM, lngW) & strTmp02
                  intPos01 = InStr(strSQL, "' AS faregr_account,")
                  intPos01 = InStr(intPos01, strSQL, ", '")
                  strTmp01 = Left(strSQL, (intPos01 + 2))
                  intPos01 = InStr(strSQL, "' AS faregr_description1")
                  strTmp02 = Mid(strSQL, intPos01)
                  strSQL = strTmp01 & arr_varTmp03(F_DSC, lngX) & strTmp02
                  .SQL = strSQL
                  .Execute
                End With  ' ** qdf.
                Set qdf = Nothing
                'lngFiles = 0&
                'arr_varTmp03 = Empty
                lngFilesImported = lngFilesImported + 1&
              End If
            Next  ' ** lngX.
          Next  ' ** lngW.

          Debug.Print "'FILES IMPORTED: " & CStr(lngFilesImported)
          DoEvents

          .Close
        End With  ' ** dbs.
        Set dbs = Nothing
      End If  ' ** blnSkip.

    End If  ' ** lngTotFiles.

  End If  ' ** lngDirs.

  Beep

'DIRS: 9
'FILES: 60
'1.  Register Report Alcoa stock
'2.  Register Report Cash Account
'3.  Register Report FNTC
'4.  Register Report Home
'5.  Register Report Household Furnishings
'6.  Register Report News Corp
'7.  Register Report Savings Account
'8.  Register Report St Joe Company
'9.  Register Report Alcatel Lucent
'10. Register Report AT&T
'11. Register Report Cash Account
'12. Register Report Money Market
'13. Register Report Nokia Stock
'14. Register Report Cash Account
'15. Register Report FNTC Stock
'16. Register Report Manulife Stock
'17. Register Report Savings Account
'18. Register Report Southern Co
'19. Register Report Car
'20. Register Report Cash Account
'21. Register Report CD # 27396
'22. Register Report CD # 27399
'23. Register Report CD # 29194
'24. Register Report CD # 29811
'25. Register Report CD # 29965
'26. Register Report CD # 33566
'27. Register Report CD # 33612
'28. Register Report CD # 33613
'29. Register Report CD # 33614
'30. Register Report CD # 33628
'31. Register Report FNTC Stock
'32. Register Report IRA 2247
'33. Register Report Manulife Stock 2
'34. Register Report Manulife Stock
'35. Register Report Money Market
'36. Register Report Savings Account
'37. Register Report Southern Co
'38. Register Report Cash Account
'39. Register Report CD # 41112
'40. Register Report CD # 41405
'41. Register Report FirstBanc of Alabama Inc stock
'42. Register Report Money Market
'43. Register Report Cash Account
'44. Register Report CD # 40932
'45. Register Report CD # 40933
'46. Register Report CD # 40934
'47. Register Report CD # 41759
'48. Register Report Savings
'49. Register Report - Cash Account
'50. Register Report - Money Market Account
'51. Register Report - Savings Account
'52. Register Report Cash Account
'53. Register Report CD # 40892
'54. Register Report Money Market Account
'55. Register Report Raymond James
'56. Register Report Savings Account
'57. Register Report Cash Account
'58. Register Report CD # 40941
'59. Register Report CD # 41290
'60. Register Report Savings Account
'DONE!
  Debug.Print "'DONE!"
  DoEvents

  Set fsfl = Nothing
  Set fsfls = Nothing
  Set fsfd1 = Nothing
  Set fsfd2 = Nothing
  Set fsfds = Nothing
  Set fso = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  FA_RegisterImport = blnRetVal

End Function

Public Function FA_RegisterQrys() As Boolean

  Const THIS_PROC As String = "FA_RegisterQrys"

  Dim dbs As DAO.Database, qdf As DAO.QueryDef, prp As Object
  Dim strQryName1 As String, strQryName2 As String, strSQL1 As String, strSQL2 As String, strDesc1 As String, strDesc2 As String
  Dim lngQrysCreated As Long
  Dim intLen As Integer
  Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  blnRetVal = True

  Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
  DoEvents

  Set dbs = CurrentDb
  With dbs

    strDesc1 = "    zz_tbl_FirstAlabama_07, not in zz_tbl_FirstAlabama_08, grouped, by .._account, .._catord1, .._catord2, with Min(facoar_id), facoa_transord = 11; "
    strQryName1 = "zzz_qry_FirstAlabama_Chart_27_11"
    strSQL1 = "SELECT zz_tbl_FirstAlabama_07.faacct_id, zz_tbl_FirstAlabama_07.facoar_account, zz_tbl_FirstAlabama_07.facoa_category1, " & _
      "zz_tbl_FirstAlabama_07.facoa_category2, zz_tbl_FirstAlabama_07.facoa_catord1, zz_tbl_FirstAlabama_07.facoa_catord2, " & _
      "CLng(11) AS facoa_transord, Min(zz_tbl_FirstAlabama_07.facoar_id) AS facoar_id" & vbCrLf
    strSQL1 = strSQL1 & "FROM zz_tbl_FirstAlabama_07 LEFT JOIN zz_tbl_FirstAlabama_08 ON " & _
      "zz_tbl_FirstAlabama_07.facoar_id = zz_tbl_FirstAlabama_08.facoar_id" & vbCrLf
    strSQL1 = strSQL1 & "WHERE (((zz_tbl_FirstAlabama_08.facoar_id) Is Null))" & vbCrLf
    strSQL1 = strSQL1 & "GROUP BY zz_tbl_FirstAlabama_07.faacct_id, zz_tbl_FirstAlabama_07.facoar_account, " & _
      "zz_tbl_FirstAlabama_07.facoa_category1, zz_tbl_FirstAlabama_07.facoa_category2, " & _
      "zz_tbl_FirstAlabama_07.facoa_catord1, zz_tbl_FirstAlabama_07.facoa_catord2;"

    strDesc2 = "                        Append .._27_11 to zz_tbl_FirstAlabama_08."
    strQryName2 = "zzz_qry_FirstAlabama_Chart_27_11a"
    strSQL2 = "INSERT INTO zz_tbl_FirstAlabama_08 ( faacct_id, facoar_account, facoa_category1, facoa_category2, " & _
      "facoa_catord1, facoa_catord2, facoa_transord, facoar_id, fat08_datemodified )" & vbCrLf
    strSQL2 = strSQL2 & "SELECT zzz_qry_FirstAlabama_Chart_27_11.faacct_id, zzz_qry_FirstAlabama_Chart_27_11.facoar_account, " & _
      "zzz_qry_FirstAlabama_Chart_27_11.facoa_category1, zzz_qry_FirstAlabama_Chart_27_11.facoa_category2, " & _
      "zzz_qry_FirstAlabama_Chart_27_11.facoa_catord1, zzz_qry_FirstAlabama_Chart_27_11.facoa_catord2, " & _
      "zzz_qry_FirstAlabama_Chart_27_11.facoa_transord, zzz_qry_FirstAlabama_Chart_27_11.facoar_id, Now() AS fat08_datemodified" & vbCrLf
    strSQL2 = strSQL2 & "FROM zzz_qry_FirstAlabama_Chart_27_11;"

    lngQrysCreated = 0&
    For lngX = 31& To 49&

      intLen = Len(strQryName1)
      strTmp01 = strQryName1
      strTmp01 = Left(strTmp01, (intLen - 2)) & CStr(lngX)
      strTmp02 = strSQL1
      strTmp02 = StringReplace(strTmp02, "CLng(11)", "CLng(" & CStr(lngX) & ")")  ' ** Module Function: modStringFuncs.
      strTmp03 = strDesc1
      strTmp03 = StringReplace(strTmp03, "facoa_transord = 11;", "facoa_transord = " & CStr(lngX) & ";")  ' ** Module Function: modStringFuncs.
      Set qdf = .CreateQueryDef(strTmp01, strTmp02)
      With qdf
        Set prp = .CreateProperty("Description", dbText, strTmp03)
On Error Resume Next
        .Properties.Append prp
        If ERR.Number <> 0 Then
On Error GoTo 0
           .Properties("Description") = strTmp03
        Else
On Error GoTo 0
        End If
        lngQrysCreated = lngQrysCreated + 1&
      End With
      Set qdf = Nothing

      intLen = Len(strQryName2)
      strTmp01 = strQryName2
      strTmp01 = Left(strTmp01, (intLen - 3)) & CStr(lngX) & "a"
      strTmp02 = strSQL2
      strTmp02 = StringReplace(strTmp02, "Chart_27_11", "Chart_27_" & CStr(lngX))  ' ** Module Function: modStringFuncs.
      strTmp03 = strDesc2
      strTmp03 = StringReplace(strTmp03, ".._27_11", ".._27_" & CStr(lngX))  ' ** Module Function: modStringFuncs.
      Set qdf = .CreateQueryDef(strTmp01, strTmp02)
      With qdf
        Set prp = .CreateProperty("Description", dbText, strTmp03)
On Error Resume Next
        .Properties.Append prp
        If ERR.Number <> 0 Then
On Error GoTo 0
           .Properties("Description") = strTmp03
        Else
On Error GoTo 0
        End If
        lngQrysCreated = lngQrysCreated + 1&
      End With
      Set qdf = Nothing

    Next

    .QueryDefs.Refresh

    .Close
  End With  ' ** dbs.
  Set dbs = Nothing

  Debug.Print "'QRYS CREATED: " & CStr(lngQrysCreated)
  DoEvents

  Beep

  Debug.Print "'DONE!"
  DoEvents

  Set prp = Nothing
  Set qdf = Nothing
  Set dbs = Nothing

  FA_RegisterQrys = blnRetVal

End Function
