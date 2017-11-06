Attribute VB_Name = "zz_mod_RecoveryFuncs"
Option Compare Database
Option Explicit

'VGC 04/11/2016: CHANGES!

Private Const THIS_NAME As String = "zz_mod_RecoveryFuncs"
' **

Public Function VBCom_Import() As Boolean

  Const THIS_PROC As String = "VBCom_Import"

  Dim dbsLnk As DAO.Database, qdf As DAO.QueryDef
  Dim strPath As String, strFile As String, strPathFile As String, strSysPathFile As String
  Dim lngQrys As Long, arr_varQry() As Variant
  Dim lngQrysImported As Long
  Dim lngX As Long, lngE As Long
  Dim blnRetVal As Boolean

  ' ** Array: arr_varQry().
  Const Q_ELEMS As Integer = 0  ' ** Array's first elemnt UBound().
  Const Q_QNAM As Integer = 0

  blnRetVal = True

  strPath = "C:\Program Files\Delta Data\Trust Accountant"
  strFile = "Trust_bak73.mdb"
  strPathFile = strPath & "\" & strFile
  strSysPathFile = strPath & "\Database\TrustSec.mdw"

  'DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acModule, _
  '  "modFileUtilities", "modFileUtilities"

  lngQrys = 0&
  ReDim arr_varQry(Q_ELEMS, 0)

  DBEngine.SystemDB = strSysPathFile
  Set dbsLnk = DBEngine.OpenDatabase(strPathFile, False, True)
  With dbsLnk
    Debug.Print "'TOT QRYS: " & CStr(.QueryDefs.Count)
    DoEvents
    For Each qdf In .QueryDefs
      With qdf
        If Left(.Name, 1) <> "~" Then  ' ** Skip those pesky system queries.
          If InStr(.Name, "MSys") > 0 Then
            Stop
          End If
          lngQrys = lngQrys + 1&
          lngE = lngQrys - 1&
          ReDim Preserve arr_varQry(Q_ELEMS, lngE)
          arr_varQry(Q_QNAM, lngE) = .Name
        End If
      End With  ' ** qdf.
    Next  ' ** qdf.
    .Close
  End With  ' ** dbs
  Set qdf = Nothing
  Set dbsLnk = Nothing

  Debug.Print "'QRYS: " & CStr(lngQrys)
  DoEvents

  CurrentDb.QueryDefs.Refresh

  If lngQrys > 0& Then

    For lngX = 0& To (lngQrys - 1&)
      If Left(arr_varQry(Q_QNAM, lngX), 3) <> "qry" Then
        If Left(arr_varQry(Q_QNAM, lngX), 2) <> "zz" Then
          Debug.Print "'QRY: " & arr_varQry(Q_QNAM, lngX)
          Stop
        End If
      End If
    Next  ' ** lngX.

    Debug.Print "'|";
    DoEvents

    lngQrysImported = 0&
    For lngX = 0& To (lngQrys - 1&)
      DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acQuery, _
        arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
      If (lngX + 1&) Mod 1000 = 0 Then
        Debug.Print "|  " & CStr(lngX + 1&) & " of " & CStr(lngQrys)
        Debug.Print "'|";
      ElseIf (lngX + 1&) Mod 100 = 0 Then
        Debug.Print "|";
      ElseIf (lngX + 1&) Mod 10 = 0 Then
        Debug.Print ".";
      End If
      DoEvents
      lngQrysImported = lngQrysImported + 1&
    Next  ' ** lngX.
    Debug.Print
    DoEvents

  End If  ' ** lngQrys.

  Debug.Print "'QRYS IMPORTED: " & CStr(lngQrysImported)
  DoEvents

  Beep

  Debug.Print "'DONE!"
  DoEvents

'TOT QRYS: 9433
'QRYS: 9433
'QRYS IMPORTED: 9433
  Set qdf = Nothing
  Set dbsLnk = Nothing

  VBCom_Import = blnRetVal

End Function
