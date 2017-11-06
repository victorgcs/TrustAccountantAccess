Attribute VB_Name = "modAppVersionFuncs"
Option Compare Database
Option Explicit

'VGC 03/23/2017: CHANGES!

' ** Database Properties:
' ** CurrentDb().Containers("Databases")
' **   .Documents.Count = 4
' **   .Documents(0).Name = AccessLayout
' **   .Documents(1).Name = MSysDb
' **   .Documents(2).Name = SummaryInfo
' **   .Documents(3).Name = UserDefined
' **
' **   .Documents("MSysDb")
' **     .Properties.Count = 25  ' ** Property order may vary.
' **     .Properties(10).Name = AppTitle
' **     .Properties(11).Name = StartUpShowDBWindow
' **     .Properties(12).Name = StartUpShowStatusBar
' **     .Properties(13).Name = AllowShortcutMenus
' **     .Properties(14).Name = AllowFullMenus
' **     .Properties(15).Name = AllowBuiltInToolbars
' **     .Properties(16).Name = AllowToolbarChanges
' **     .Properties(17).Name = AllowSpecialKeys
' **     .Properties(18).Name = AppIcon
' **     .Properties(20).Name = Auto Compact
' **     .Properties(21).Name = StartUpForm
' **   .Documents("SummaryInfo")
' **     .Properties.Count = 11  ' ** Property Order may vary.
' **     .Properties(8).Name = Title
' **     .Properties(9).Name = Author
' **     .Properties(10).Name = Company
' **     .Properties(11).Name = Manager
' **   .Documents("UserDefined")
' **     .Properties.Count = 11
' **     .Properties(9).Name = AppVersion
' **     .Properties(10).Name = AppDate

Private Const THIS_NAME As String = "modAppVersionFuncs"
' **

Public Function App_SetAll() As Boolean
 
100   On Error GoTo ERRH

        Const THIS_PROC As String = "App_SetAll"

        Dim blnRetVal As Boolean

110     blnRetVal = True

120     AppVersion_Add  ' ** Function: Below.
130     AppDate_Add  ' ** Function: Below.
140     AppTitle_Let  ' ** Function: Below.
150     AppAuthor_Let  ' ** Function: Below.
160     AppManager_Add  ' ** Function: Below.
170     AppCompany_Let  ' ** Function: Below.
180     AppIcon_Let  ' ** Function: Below.

190     Beep

EXITP:
200     App_SetAll = blnRetVal
210     Exit Function

ERRH:
220     blnRetVal = False
230     Select Case ERR.Number
        Case Else
240       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
250     End Select
260     Resume EXITP

End Function

Public Function AppVersion_Get() As String
' ** Returns both the version and date, as a single string.

300   On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_Get"

        Dim dbs As DAO.Database, doc As DAO.Document
        Dim datVer_Date As Date
        Dim strRetVal As String

310     strRetVal = vbNullString

320     Set dbs = CurrentDb
330     Set doc = dbs.Containers("Databases").Documents("UserDefined")
340     strRetVal = "v " & doc.Properties("AppVersion").Value
350     datVer_Date = doc.Properties("AppDate").Value
360     strRetVal = strRetVal & " " & Format(datVer_Date, "mm/dd/yy")
370     dbs.Close

EXITP:
380     Set doc = Nothing
390     Set dbs = Nothing
400     AppVersion_Get = strRetVal
410     Exit Function

ERRH:
420     strRetVal = vbNullString
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
450     End Select
460     Resume EXITP

End Function

Public Function AppVersion_Get2() As String
' ** Returns just the version alone.

500   On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_Get2"

        Dim dbs As DAO.Database, doc As DAO.Document
        Dim strRetVal As String

510     strRetVal = vbNullString

520     Set dbs = CurrentDb
530     Set doc = dbs.Containers("Databases").Documents("UserDefined")
540     If doc.Properties.Count = 8 Then
550       strRetVal = RET_ERR
560     Else
570       strRetVal = doc.Properties("AppVersion").Value
580     End If
590     dbs.Close

EXITP:
600     Set doc = Nothing
610     Set dbs = Nothing
620     AppVersion_Get2 = strRetVal
630     Exit Function

ERRH:
640     strRetVal = vbNullString
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
670     End Select
680     Resume EXITP

End Function

Public Function AppVersion_GetDta() As String

700   On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_GetDta"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim strPathFile As String
        Dim strRetVal As String

710     strRetVal = "#NOTFOUND"

720   On Error Resume Next
730     Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
740     If ERR.Number <> 0 Then
750   On Error GoTo ERRH
760   On Error Resume Next
770       Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
780       If ERR.Number <> 0 Then
790   On Error GoTo ERRH
800   On Error Resume Next
810         Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
820         If ERR.Number <> 0 Then
830   On Error GoTo ERRH
840   On Error Resume Next
850           Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
860           If ERR.Number <> 0 Then
870   On Error GoTo ERRH
880   On Error Resume Next
890             Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
900             If ERR.Number <> 0 Then
910   On Error GoTo ERRH
920   On Error Resume Next
930               Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
940               If ERR.Number <> 0 Then
950   On Error GoTo ERRH
960   On Error Resume Next
970                 Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
980   On Error GoTo ERRH
990               Else
1000  On Error GoTo ERRH
1010              End If
1020            Else
1030  On Error GoTo ERRH
1040            End If
1050          Else
1060  On Error GoTo ERRH
1070          End If
1080        Else
1090  On Error GoTo ERRH
1100        End If
1110      Else
1120  On Error GoTo ERRH
1130      End If
1140    Else
1150  On Error GoTo ERRH
1160    End If

1170    With wrk
1180      strPathFile = CurrentBackendPathFile("account")  ' ** Module Function: modFileUtilities.
1190      Set dbs = .OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
1200      For Each doc In dbs.Containers("Databases").Documents
1210        If doc.Name = "UserDefined" Then
1220          For Each prp In doc.Properties
1230            If prp.Name = "AppVersion" Then
1240              strRetVal = doc.Properties("AppVersion").Value
1250              Exit For
1260            End If
1270          Next
1280          Exit For
1290        End If
1300      Next
1310      .Close
1320    End With

EXITP:
1330    Set prp = Nothing
1340    Set doc = Nothing
1350    Set dbs = Nothing
1360    Set wrk = Nothing
1370    AppVersion_GetDta = strRetVal
1380    Exit Function

ERRH:
1390    strRetVal = vbNullString
1400    Select Case ERR.Number
        Case Else
1410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
1420    End Select
1430    Resume EXITP

End Function

Public Function AppVersion_GetArch() As String

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_GetArch"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim strPathFile As String
        Dim strRetVal As String

1510    strRetVal = "#NOTFOUND"

1520  On Error Resume Next
1530    Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
1540    If ERR.Number <> 0 Then
1550  On Error GoTo ERRH
1560  On Error Resume Next
1570      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
1580      If ERR.Number <> 0 Then
1590  On Error GoTo ERRH
1600  On Error Resume Next
1610        Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
1620        If ERR.Number <> 0 Then
1630  On Error GoTo ERRH
1640  On Error Resume Next
1650          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
1660          If ERR.Number <> 0 Then
1670  On Error GoTo ERRH
1680  On Error Resume Next
1690            Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
1700            If ERR.Number <> 0 Then
1710  On Error GoTo ERRH
1720  On Error Resume Next
1730              Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
1740              If ERR.Number <> 0 Then
1750  On Error GoTo ERRH
1760  On Error Resume Next
1770                Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
1780  On Error GoTo ERRH
1790              Else
1800  On Error GoTo ERRH
1810              End If
1820            Else
1830  On Error GoTo ERRH
1840            End If
1850          Else
1860  On Error GoTo ERRH
1870          End If
1880        Else
1890  On Error GoTo ERRH
1900        End If
1910      Else
1920  On Error GoTo ERRH
1930      End If
1940    Else
1950  On Error GoTo ERRH
1960    End If

1970    With wrk
1980      strPathFile = CurrentBackendPathFile("LedgerArchive")  ' ** Module Function: modFileUtilities.
1990      Set dbs = .OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
2000      For Each doc In dbs.Containers("Databases").Documents
2010        If doc.Name = "UserDefined" Then
2020          For Each prp In doc.Properties
2030            If prp.Name = "AppVersion" Then
2040              strRetVal = doc.Properties("AppVersion").Value
2050              Exit For
2060            End If
2070          Next
2080          Exit For
2090        End If
2100      Next
2110      .Close
2120    End With

EXITP:
2130    Set prp = Nothing
2140    Set doc = Nothing
2150    Set dbs = Nothing
2160    Set wrk = Nothing
2170    AppVersion_GetArch = strRetVal
2180    Exit Function

ERRH:
2190    strRetVal = vbNullString
2200    Select Case ERR.Number
        Case Else
2210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
2220    End Select
2230    Resume EXITP

End Function

Public Function AppVersion_GetAux() As String

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_GetAux"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim strPathFile As String
        Dim strRetVal As String

2310    strRetVal = "#NOTFOUND"

2320  On Error Resume Next
2330    Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
2340    If ERR.Number <> 0 Then
2350  On Error GoTo ERRH
2360  On Error Resume Next
2370      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
2380      If ERR.Number <> 0 Then
2390  On Error GoTo ERRH
2400  On Error Resume Next
2410        Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
2420        If ERR.Number <> 0 Then
2430  On Error GoTo ERRH
2440  On Error Resume Next
2450          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
2460          If ERR.Number <> 0 Then
2470  On Error GoTo ERRH
2480  On Error Resume Next
2490            Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
2500            If ERR.Number <> 0 Then
2510  On Error GoTo ERRH
2520  On Error Resume Next
2530              Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
2540              If ERR.Number <> 0 Then
2550  On Error GoTo ERRH
2560  On Error Resume Next
2570                Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
2580  On Error GoTo ERRH
2590              Else
2600  On Error GoTo ERRH
2610              End If
2620            Else
2630  On Error GoTo ERRH
2640            End If
2650          Else
2660  On Error GoTo ERRH
2670          End If
2680        Else
2690  On Error GoTo ERRH
2700        End If
2710      Else
2720  On Error GoTo ERRH
2730      End If
2740    Else
2750  On Error GoTo ERRH
2760    End If

2770    With wrk
2780      strPathFile = CurrentBackendPathFile("tblDatabase")  ' ** Module Function: modFileUtilities.
2790      Set dbs = .OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
2800      For Each doc In dbs.Containers("Databases").Documents
2810        If doc.Name = "UserDefined" Then
2820          For Each prp In doc.Properties
2830            If prp.Name = "AppVersion" Then
2840              strRetVal = doc.Properties("AppVersion").Value
2850              Exit For
2860            End If
2870          Next
2880          Exit For
2890        End If
2900      Next
2910      .Close
2920    End With

EXITP:
2930    Set prp = Nothing
2940    Set doc = Nothing
2950    Set dbs = Nothing
2960    Set wrk = Nothing
2970    AppVersion_GetAux = strRetVal
2980    Exit Function

ERRH:
2990    strRetVal = vbNullString
3000    Select Case ERR.Number
        Case Else
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
3020    End Select
3030    Resume EXITP

End Function

Public Function AppVersion_GetMVP() As String

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_GetMVP"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim strRetVal As String

3110    strRetVal = vbNullString

3120    Set dbs = CurrentDb
3130    With dbs
3140      For Each tdf In .TableDefs
3150        With tdf
3160          If .Name = "m_VP" Then
3170  On Error Resume Next
3180            Set rst = dbs.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
3190            If ERR = 0 Then
3200  On Error GoTo ERRH
3210              With rst
3220                .MoveFirst
3230                strRetVal = CStr(![vp_MAIN]) & "." & CStr(![vp_MINOR]) & "." & _
                      IIf(Len(CStr(Nz(![vp_REVISION], 0))) = 1, CStr(Nz(![vp_REVISION], 0)) & "0", CStr(Nz(![vp_REVISION], 0)))
3240                .Close
3250              End With
3260            Else
3270  On Error GoTo ERRH
3280            End If
3290            Exit For
3300          End If
3310        End With
3320      Next
3330      .Close
3340    End With

EXITP:
3350    Set tdf = Nothing
3360    Set rst = Nothing
3370    Set dbs = Nothing
3380    AppVersion_GetMVP = strRetVal
3390    Exit Function

ERRH:
3400    strRetVal = vbNullString
3410    Select Case ERR.Number
        Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
3430    End Select
3440    Resume EXITP

End Function

Public Function AppVersion_GetMVD() As String

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_GetMVD"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim strRetVal As String

3510    strRetVal = vbNullString

3520    Set dbs = CurrentDb
3530    With dbs
3540      For Each tdf In .TableDefs
3550        With tdf
3560          If .Name = "m_VD" Then
3570  On Error Resume Next
3580            Set rst = dbs.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
3590            If ERR = 0 Then
3600  On Error GoTo ERRH
3610              With rst
3620                .MoveFirst
3630                strRetVal = CStr(![vd_MAIN]) & "." & CStr(![vd_MINOR]) & "." & _
                      IIf(Len(CStr(Nz(![vd_REVISION], 0))) = 1, CStr(Nz(![vd_REVISION], 0)) & "0", CStr(Nz(![vd_REVISION], 0)))
3640                .Close
3650              End With
3660            Else
3670  On Error GoTo ERRH
3680            End If
3690            Exit For
3700          End If
3710        End With
3720      Next
3730      .Close
3740    End With

EXITP:
3750    Set tdf = Nothing
3760    Set rst = Nothing
3770    Set dbs = Nothing
3780    AppVersion_GetMVD = strRetVal
3790    Exit Function

ERRH:
3800    strRetVal = vbNullString
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
3830    End Select
3840    Resume EXITP

End Function

Public Function AppVersion_GetMVA() As String

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_GetMVA"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim strRetVal As String

3910    strRetVal = vbNullString

3920    Set dbs = CurrentDb
3930    With dbs
3940      For Each tdf In .TableDefs
3950        With tdf
3960          If .Name = "m_VA" Then
3970  On Error Resume Next
3980            Set rst = dbs.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
3990            If ERR = 0 Then
4000  On Error GoTo ERRH
4010              With rst
4020                .MoveFirst
4030                strRetVal = CStr(![va_MAIN]) & "." & CStr(![va_MINOR]) & "." & _
                      IIf(Len(CStr(Nz(![va_REVISION], 0))) = 1, CStr(Nz(![va_REVISION], 0)) & "0", CStr(Nz(![va_REVISION], 0)))
4040                .Close
4050              End With
4060            Else
4070  On Error GoTo ERRH
4080            End If
4090            Exit For
4100          End If
4110        End With
4120      Next
4130      .Close
4140    End With

EXITP:
4150    Set tdf = Nothing
4160    Set rst = Nothing
4170    Set dbs = Nothing
4180    AppVersion_GetMVA = strRetVal
4190    Exit Function

ERRH:
4200    strRetVal = vbNullString
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
4230    End Select
4240    Resume EXITP

End Function

Public Function AppVersion_GetMVX() As String

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_GetMVX"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim strRetVal As String

4310    strRetVal = vbNullString

4320    Set dbs = CurrentDb
4330    With dbs
4340      For Each tdf In .TableDefs
4350        With tdf
4360          If .Name = "m_VX" Then
4370  On Error Resume Next
4380            Set rst = dbs.OpenRecordset(.Name, dbOpenDynaset, dbReadOnly)
4390            If ERR = 0 Then
4400  On Error GoTo ERRH
4410              With rst
4420                .MoveFirst
4430                strRetVal = CStr(![vx_MAIN]) & "." & CStr(![vx_MINOR]) & "." & _
                      IIf(Len(CStr(Nz(![vx_REVISION], 0))) = 1, CStr(Nz(![vx_REVISION], 0)) & "0", CStr(Nz(![vx_REVISION], 0)))
4440                .Close
4450              End With
4460            Else
4470  On Error GoTo ERRH
4480            End If
4490            Exit For
4500          End If
4510        End With
4520      Next
4530      .Close
4540    End With

EXITP:
4550    Set tdf = Nothing
4560    Set rst = Nothing
4570    Set dbs = Nothing
4580    AppVersion_GetMVX = strRetVal
4590    Exit Function

ERRH:
4600    strRetVal = vbNullString
4610    Select Case ERR.Number
        Case Else
4620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
4630    End Select
4640    Resume EXITP

End Function

Public Function AppVersion_Let(strVer As String) As Boolean

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_Let"

        Dim dbs As DAO.Database, doc As DAO.Document
        Dim blnRetVal As Boolean

4710    blnRetVal = True

4720    Set dbs = CurrentDb
4730    Set doc = dbs.Containers("Databases").Documents("UserDefined")
4740    doc.Properties("AppVersion").Value = strVer
4750    doc.Properties("AppDate").Value = Now()
4760    dbs.Close

EXITP:
4770    Set doc = Nothing
4780    Set dbs = Nothing
4790    AppVersion_Let = blnRetVal
4800    Exit Function

ERRH:
4810    blnRetVal = False
4820    Select Case ERR.Number
        Case Else
4830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
4840    End Select
4850    Resume EXITP

End Function

Public Function AppVersion_Add() As Boolean

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "AppVersion_Add"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

4910    blnRetVal = True

4920    Set dbs = CurrentDb
4930    Set doc = dbs.Containers("Databases").Documents("UserDefined")
4940    With doc
4950      Set prp = .CreateProperty("AppVersion", dbText, "2.2.24", True)  ' ** True prevents easy Edit or Delete.
4960  On Error Resume Next
4970      .Properties.Append prp
4980      If ERR.Number <> 0 Then
4990        Select Case ERR.Number
            Case 3367  ' ** Cannot append.  An object with that name already exists in the collection.
5000  On Error GoTo ERRH
5010          .Properties("AppVersion").Value = "2.2.24"
5020        Case Else
5030          Debug.Print "'AppVersion_Add()"
5040          Debug.Print "'  Error: " & CStr(ERR.Number)
5050          Debug.Print "'  Desc: " & ERR.description
5060          Debug.Print "'  Line: " & Erl
5070  On Error GoTo ERRH
5080        End Select
5090      Else
5100  On Error GoTo ERRH
5110      End If
5120    End With
5130    dbs.Close

EXITP:
5140    Set prp = Nothing
5150    Set doc = Nothing
5160    Set dbs = Nothing
5170    AppVersion_Add = blnRetVal
5180    Exit Function

ERRH:
5190    blnRetVal = False
5200    Select Case ERR.Number
        Case Else
5210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
5220    End Select
5230    Resume EXITP

End Function

Public Function AppDate_Get(Optional varTime As Variant) As String

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "AppDate_Get"

        Dim dbs As DAO.Database, doc As DAO.Document
        Dim datVer_Date As Date
        Dim strRetVal As String

5310    strRetVal = vbNullString

5320    Set dbs = CurrentDb
5330    Set doc = dbs.Containers("Databases").Documents("UserDefined")
5340    datVer_Date = doc.Properties("AppDate").Value
5350    If IsMissing(varTime) = True Then
5360      strRetVal = Format(datVer_Date, "mm/dd/yy")
5370    Else
5380      If CBool(varTime) = True Then
5390        strRetVal = Format(datVer_Date, "mm/dd/yyyy hh:nn:ss AM/PM")
5400      Else
5410        strRetVal = Format(datVer_Date, "mm/dd/yy")
5420      End If
5430    End If
5440    dbs.Close

EXITP:
5450    Set doc = Nothing
5460    Set dbs = Nothing
5470    AppDate_Get = strRetVal
5480    Exit Function

ERRH:
5490    strRetVal = vbNullString
5500    Select Case ERR.Number
        Case Else
5510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
5520    End Select
5530    Resume EXITP

End Function

Public Function AppDate_Let(Optional varBeep As Variant) As Boolean

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "AppDate_Let"

        Dim dbs As DAO.Database, doc As DAO.Document
        Dim blnBeep As Boolean
        Dim blnRetVal As Boolean

5610    blnRetVal = True

5620    Select Case IsMissing(varBeep)
        Case True
5630      blnBeep = False
5640    Case False
5650      blnBeep = CBool(varBeep)
5660    End Select

5670    Set dbs = CurrentDb
5680    Set doc = dbs.Containers("Databases").Documents("UserDefined")
5690    doc.Properties("AppDate").Value = Now()
5700    dbs.Close

5710    If blnBeep = True Then
5720      Beep
5730    End If

EXITP:
5740    Set doc = Nothing
5750    Set dbs = Nothing
5760    AppDate_Let = blnRetVal
5770    Exit Function

ERRH:
5780    blnRetVal = False
5790    Select Case ERR.Number
        Case Else
5800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
5810    End Select
5820    Resume EXITP

End Function

Public Function AppDate_Add() As Boolean

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "AppDate_Add"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

5910    blnRetVal = True

5920    Set dbs = CurrentDb
5930    Set doc = dbs.Containers("Databases").Documents("UserDefined")
5940    With doc
5950      Set prp = .CreateProperty("AppDate", dbDate, Now(), True)  ' ** True prevents easy Edit or Delete.
5960  On Error Resume Next
5970      .Properties.Append prp
5980      If ERR.Number <> 0 Then
5990        Select Case ERR.Number
            Case 3367  ' ** Cannot append.  An object with that name already exists in the collection.
6000  On Error GoTo ERRH
6010          .Properties("AppDate").Value = Now()
6020        Case Else
6030          Debug.Print "'AppDate_Add()"
6040          Debug.Print "'  Error: " & CStr(ERR.Number)
6050          Debug.Print "'  Desc: " & ERR.description
6060          Debug.Print "'  Line: " & Erl
6070  On Error GoTo ERRH
6080        End Select
6090      Else
6100  On Error GoTo ERRH
6110      End If
6120    End With
6130    dbs.Close

EXITP:
6140    Set prp = Nothing
6150    Set doc = Nothing
6160    Set dbs = Nothing
6170    AppDate_Add = blnRetVal
6180    Exit Function

ERRH:
6190    blnRetVal = False
6200    Select Case ERR.Number
        Case Else
6210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
6220    End Select
6230    Resume EXITP

End Function

Public Function AppDate_GetDta() As String

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "AppDate_GetDta"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim strPathFile As String
        Dim strRetVal As String

6310    strRetVal = "#NOTFOUND"

6320  On Error Resume Next
6330    Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
6340    If ERR.Number <> 0 Then
6350  On Error GoTo ERRH
6360  On Error Resume Next
6370      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
6380      If ERR.Number <> 0 Then
6390  On Error GoTo ERRH
6400  On Error Resume Next
6410        Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
6420        If ERR.Number <> 0 Then
6430  On Error GoTo ERRH
6440  On Error Resume Next
6450          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
6460          If ERR.Number <> 0 Then
6470  On Error GoTo ERRH
6480  On Error Resume Next
6490            Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
6500            If ERR.Number <> 0 Then
6510  On Error GoTo ERRH
6520  On Error Resume Next
6530              Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
6540              If ERR.Number <> 0 Then
6550  On Error GoTo ERRH
6560  On Error Resume Next
6570                Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
6580  On Error GoTo ERRH
6590              Else
6600  On Error GoTo ERRH
6610              End If
6620            Else
6630  On Error GoTo ERRH
6640            End If
6650          Else
6660  On Error GoTo ERRH
6670          End If
6680        Else
6690  On Error GoTo ERRH
6700        End If
6710      Else
6720  On Error GoTo ERRH
6730      End If
6740    Else
6750  On Error GoTo ERRH
6760    End If

6770    With wrk
6780      strPathFile = CurrentBackendPathFile("account")  ' ** Module Function: modFileUtilities.
6790      Set dbs = .OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
6800      For Each doc In dbs.Containers("Databases").Documents
6810        If doc.Name = "UserDefined" Then
6820          For Each prp In doc.Properties
6830            If prp.Name = "AppDate" Then
6840              strRetVal = doc.Properties("AppDate").Value
6850              Exit For
6860            End If
6870          Next
6880          Exit For
6890        End If
6900      Next
6910      .Close
6920    End With

EXITP:
6930    Set prp = Nothing
6940    Set doc = Nothing
6950    Set dbs = Nothing
6960    Set wrk = Nothing
6970    AppDate_GetDta = strRetVal
6980    Exit Function

ERRH:
6990    strRetVal = vbNullString
7000    Select Case ERR.Number
        Case Else
7010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
7020    End Select
7030    Resume EXITP

End Function

Public Function AppDate_GetArch() As String

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "AppDate_GetArch"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim strPathFile As String
        Dim strRetVal As String

7110    strRetVal = "#NOTFOUND"

7120  On Error Resume Next
7130    Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
7140    If ERR.Number <> 0 Then
7150  On Error GoTo ERRH
7160  On Error Resume Next
7170      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
7180      If ERR.Number <> 0 Then
7190  On Error GoTo ERRH
7200  On Error Resume Next
7210        Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
7220        If ERR.Number <> 0 Then
7230  On Error GoTo ERRH
7240  On Error Resume Next
7250          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
7260          If ERR.Number <> 0 Then
7270  On Error GoTo ERRH
7280  On Error Resume Next
7290            Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
7300            If ERR.Number <> 0 Then
7310  On Error GoTo ERRH
7320  On Error Resume Next
7330              Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
7340              If ERR.Number <> 0 Then
7350  On Error GoTo ERRH
7360  On Error Resume Next
7370                Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
7380  On Error GoTo ERRH
7390              Else
7400  On Error GoTo ERRH
7410              End If
7420            Else
7430  On Error GoTo ERRH
7440            End If
7450          Else
7460  On Error GoTo ERRH
7470          End If
7480        Else
7490  On Error GoTo ERRH
7500        End If
7510      Else
7520  On Error GoTo ERRH
7530      End If
7540    Else
7550  On Error GoTo ERRH
7560    End If

7570    With wrk
7580      strPathFile = CurrentBackendPathFile("LedgerArchive")  ' ** Module Function: modFileUtilities.
7590      Set dbs = .OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
7600      For Each doc In dbs.Containers("Databases").Documents
7610        If doc.Name = "UserDefined" Then
7620          For Each prp In doc.Properties
7630            If prp.Name = "AppDate" Then
7640              strRetVal = doc.Properties("AppDate").Value
7650              Exit For
7660            End If
7670          Next
7680          Exit For
7690        End If
7700      Next
7710      .Close
7720    End With

EXITP:
7730    Set prp = Nothing
7740    Set doc = Nothing
7750    Set dbs = Nothing
7760    Set wrk = Nothing
7770    AppDate_GetArch = strRetVal
7780    Exit Function

ERRH:
7790    strRetVal = vbNullString
7800    Select Case ERR.Number
        Case Else
7810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
7820    End Select
7830    Resume EXITP

End Function

Public Function AppDate_GetAux() As String

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "AppDate_GetAux"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim strPathFile As String
        Dim strRetVal As String

7910    strRetVal = "#NOTFOUND"

7920  On Error Resume Next
7930    Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)  ' ** New.
7940    If ERR.Number <> 0 Then
7950  On Error GoTo ERRH
7960  On Error Resume Next
7970      Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC2, dbUseJet)  ' ** New Demo.
7980      If ERR.Number <> 0 Then
7990  On Error GoTo ERRH
8000  On Error Resume Next
8010        Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC5, dbUseJet)  ' ** Old.
8020        If ERR.Number <> 0 Then
8030  On Error GoTo ERRH
8040  On Error Resume Next
8050          Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC6, dbUseJet)  ' ** Old Demo.
8060          If ERR.Number <> 0 Then
8070  On Error GoTo ERRH
8080  On Error Resume Next
8090            Set wrk = CreateWorkspace("tmpDB", "TAAdmin", TA_SEC3, dbUseJet)  ' ** New Admin.
8100            If ERR.Number <> 0 Then
8110  On Error GoTo ERRH
8120  On Error Resume Next
8130              Set wrk = CreateWorkspace("tmpDB", "Admin", "TA_SEC7", dbUseJet)  ' ** Old Admin.
8140              If ERR.Number <> 0 Then
8150  On Error GoTo ERRH
8160  On Error Resume Next
8170                Set wrk = CreateWorkspace("tmpDB", "Admin", "", dbUseJet)  ' ** Generic.
8180  On Error GoTo ERRH
8190              Else
8200  On Error GoTo ERRH
8210              End If
8220            Else
8230  On Error GoTo ERRH
8240            End If
8250          Else
8260  On Error GoTo ERRH
8270          End If
8280        Else
8290  On Error GoTo ERRH
8300        End If
8310      Else
8320  On Error GoTo ERRH
8330      End If
8340    Else
8350  On Error GoTo ERRH
8360    End If

8370    With wrk
8380      strPathFile = CurrentBackendPathFile("tblDatabase")  ' ** Module Function: modFileUtilities.
8390      Set dbs = .OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
8400      For Each doc In dbs.Containers("Databases").Documents
8410        If doc.Name = "UserDefined" Then
8420          For Each prp In doc.Properties
8430            If prp.Name = "AppDate" Then
8440              strRetVal = doc.Properties("AppDate").Value
8450              Exit For
8460            End If
8470          Next
8480          Exit For
8490        End If
8500      Next
8510      .Close
8520    End With

EXITP:
8530    Set prp = Nothing
8540    Set doc = Nothing
8550    Set dbs = Nothing
8560    Set wrk = Nothing
8570    AppDate_GetAux = strRetVal
8580    Exit Function

ERRH:
8590    strRetVal = vbNullString
8600    Select Case ERR.Number
        Case Else
8610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
8620    End Select
8630    Resume EXITP

End Function

Public Function AppDate_Update() As Boolean

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "AppDate_Update"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, doc As DAO.Document
        Dim strUser As String, strPass As String
        Dim strPathFile_MDB As String, strPathFile_MDW As String
        Dim blnExtras As Boolean
        Dim lngTmp01 As Long
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

8710    blnRetVal = True

8720    strUser = "Superuser"
8730    strPass = TA_SEC

8740    For lngX = 1& To 1&

8750      strPathFile_MDB = gstrDir_Dev
8760      strPathFile_MDW = gstrDir_Dev
8770      blnExtras = False
8780      Select Case lngX
          Case 1&
            '"\EmptyDatabase"
8790        strPathFile_MDB = strPathFile_MDB & LNK_SEP & gstrDir_DevEmpty
8800        strPathFile_MDW = strPathFile_MDW & LNK_SEP & gstrDir_DevEmpty
8810      Case 2&
            '"\EmptyDatabase\bak_2_2_10"
8820        strPathFile_MDB = strPathFile_MDB & LNK_SEP & gstrDir_DevEmpty & LNK_SEP & "bak_2_2_10"
8830        strPathFile_MDW = strPathFile_MDW & LNK_SEP & gstrDir_DevEmpty
8840      Case 3&
            '"\DemoDatabase"
8850        strPathFile_MDB = strPathFile_MDB & LNK_SEP & gstrDir_DevDemo
8860        strPathFile_MDW = strPathFile_MDW & LNK_SEP & gstrDir_DevDemo
8870        blnExtras = True
8880      Case 4&
            '"\DemoDatabase\bak_2_2_10"
8890        strPathFile_MDB = strPathFile_MDB & LNK_SEP & gstrDir_DevDemo & LNK_SEP & "bak_2_2_10"
8900        strPathFile_MDW = strPathFile_MDW & LNK_SEP & gstrDir_DevDemo
8910      End Select

8920      strPathFile_MDW = strPathFile_MDW & LNK_SEP & "TrustSec.mdw"

8930      Select Case blnExtras
          Case True
8940        lngTmp01 = 6&
8950      Case False
8960        lngTmp01 = 2&
8970      End Select

8980      For lngY = 1& To lngTmp01
8990        Select Case lngY
            Case 1&
              'TrstArch.mdb
9000          strPathFile_MDB = strPathFile_MDB & LNK_SEP & "TrstArch.mdb"
9010        Case 2&
              'TrustDta.mdb
9020          strPathFile_MDB = Parse_Path(strPathFile_MDB)  ' ** Module Function: modFileUtilities.
9030          strPathFile_MDB = strPathFile_MDB & LNK_SEP & "TrustDta.mdb"
9040        Case 3&
              'TrstArch_Hinton.mdb
9050          strPathFile_MDB = Parse_Path(strPathFile_MDB)  ' ** Module Function: modFileUtilities.
9060          strPathFile_MDB = strPathFile_MDB & LNK_SEP & "TrstArch_Hinton.mdb"
9070        Case 4&
              'TrustDta_Hinton.mdb
9080          strPathFile_MDB = Parse_Path(strPathFile_MDB)  ' ** Module Function: modFileUtilities.
9090          strPathFile_MDB = strPathFile_MDB & LNK_SEP & "TrustDta_Hinton.mdb"
9100        Case 5&
              'TrstArch_WmBJohnson.mdb
9110          strPathFile_MDB = Parse_Path(strPathFile_MDB)  ' ** Module Function: modFileUtilities.
9120          strPathFile_MDB = strPathFile_MDB & LNK_SEP & "TrstArch_WmBJohnson.mdb"
9130        Case 6&
              'TrustDta_WmBJohnson.mdb
9140          strPathFile_MDB = Parse_Path(strPathFile_MDB)  ' ** Module Function: modFileUtilities.
9150          strPathFile_MDB = strPathFile_MDB & LNK_SEP & "TrustDta_WmBJohnson.mdb"
9160        End Select

9170        Set wrk = CreateWorkspace("Tmp", strUser, strPass, dbUseJet)
9180        With wrk
9190          Set dbs = .OpenDatabase(strPathFile_MDB, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
9200          With dbs
9210            Set doc = dbs.Containers("Databases").Documents("UserDefined")
9220            doc.Properties("AppDate").Value = Now()
9230            .Close
9240          End With
9250          Set dbs = Nothing
9260          .Close
9270        End With
9280        Set wrk = Nothing

            ' **   strTmp01 = DBEngine.SystemDB
            ' **   Application.CurrentProject.Connection.Properties("Jet OLEDB:System database")

9290      Next  ' ** lngY

9300    Next  ' ** lngX.

9310    Beep

EXITP:
9320    Set doc = Nothing
9330    Set dbs = Nothing
9340    Set wrk = Nothing
9350    AppDate_Update = blnRetVal
9360    Exit Function

ERRH:
9370    blnRetVal = False
9380    Select Case ERR.Number
        Case Else
9390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
9400    End Select
9410    Resume EXITP

End Function

Public Function AppTime_Add() As Boolean

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "AppTime_Add"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

9510    blnRetVal = True

9520    Set dbs = CurrentDb
9530    Set doc = dbs.Containers("Databases").Documents("UserDefined")
9540    With doc
9550      Set prp = .CreateProperty("AppTime", dbDate, Now(), True)  ' ** True prevents easy Edit or Delete.
9560      .Properties.Append prp
9570    End With
9580    dbs.Close

EXITP:
9590    Set prp = Nothing
9600    Set doc = Nothing
9610    Set dbs = Nothing
9620    AppTime_Add = blnRetVal
9630    Exit Function

ERRH:
9640    blnRetVal = False
9650    Select Case ERR.Number
        Case Else
9660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
9670    End Select
9680    Resume EXITP

End Function

Public Function AppTitle_Let() As Boolean

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "AppTitle_Let"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

9710    blnRetVal = True

9720    Set dbs = CurrentDb
9730    Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
9740    doc.Properties("Title").Value = "Trust Accountant™"  ' ** Database Properties.
9750  On Error Resume Next
9760    dbs.Properties("AppTitle") = "Trust Accountant™"    ' ** Startup Properties.
9770    If ERR <> 0 Then
9780  On Error GoTo ERRH
9790      With dbs
9800        Set prp = .CreateProperty("AppTitle", dbText, "Trust Accountant™")
9810        .Properties.Append prp
9820      End With
9830    Else
9840  On Error GoTo ERRH
9850    End If
9860    dbs.Close

EXITP:
9870    Set prp = Nothing
9880    Set doc = Nothing
9890    Set dbs = Nothing
9900    AppTitle_Let = blnRetVal
9910    Exit Function

ERRH:
9920    blnRetVal = False
9930    Select Case ERR.Number
        Case Else
9940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
9950    End Select
9960    Resume EXITP

End Function

Public Function AppAuthor_Let() As Boolean

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "AppAuthor_Let"

        Dim dbs As DAO.Database, doc As DAO.Document
        Dim blnRetVal As Boolean

10010   blnRetVal = True

10020   Set dbs = CurrentDb
10030   Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
10040   doc.Properties("Author").Value = "gb, mike, vgc, fc, et al"
10050   dbs.Close

EXITP:
10060   Set doc = Nothing
10070   Set dbs = Nothing
10080   AppAuthor_Let = blnRetVal
10090   Exit Function

ERRH:
10100   blnRetVal = False
10110   Select Case ERR.Number
        Case Else
10120     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
10130   End Select
10140   Resume EXITP

End Function

Public Function AppManager_Add() As Boolean

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "AppManager_Add"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

10210   blnRetVal = True

10220   Set dbs = CurrentDb
10230   Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
10240   With doc
10250     Set prp = .CreateProperty("Manager", dbText, "Rich McCabe")
10260     .Properties.Append prp
10270   End With
10280   dbs.Close

EXITP:
10290   Set prp = Nothing
10300   Set doc = Nothing
10310   Set dbs = Nothing
10320   AppManager_Add = blnRetVal
10330   Exit Function

ERRH:
10340   blnRetVal = False
10350   Select Case ERR.Number
        Case Else
10360     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
10370   End Select
10380   Resume EXITP

End Function

Public Function AppManager_Let() As Boolean

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "AppManager_Let"

        Dim dbs As DAO.Database, doc As DAO.Document
        Dim blnRetVal As Boolean

10410   blnRetVal = True

10420   Set dbs = CurrentDb
10430   Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
10440   doc.Properties("Manager").Value = "Rich McCabe"
10450   dbs.Close

EXITP:
10460   Set doc = Nothing
10470   Set dbs = Nothing
10480   AppManager_Let = blnRetVal
10490   Exit Function

ERRH:
10500   blnRetVal = False
10510   Select Case ERR.Number
        Case Else
10520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
10530   End Select
10540   Resume EXITP

End Function

Public Function AppCompany_Let() As Boolean

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "AppCompany_Let"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

10610   blnRetVal = True

10620   Set dbs = CurrentDb
10630   Set doc = dbs.Containers("Databases").Documents("MSysDb")
10640 On Error Resume Next
10650   doc.Properties("Company").Value = "Delta Data, Inc."
10660   If ERR <> 0 Then
10670 On Error GoTo ERRH
10680     With dbs
10690       Set prp = .CreateProperty("Company", dbText, "Delta Data, Inc.")
10700       .Properties.Append prp
10710     End With
10720   Else
10730 On Error GoTo ERRH
10740   End If
10750   dbs.Close

EXITP:
10760   Set prp = Nothing
10770   Set doc = Nothing
10780   Set dbs = Nothing
10790   AppCompany_Let = blnRetVal
10800   Exit Function

ERRH:
10810   blnRetVal = False
10820   Select Case ERR.Number
        Case Else
10830     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
10840   End Select
10850   Resume EXITP

End Function

Public Function AppIcon_Let() As Boolean

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "AppIcon_Let"

        Dim dbs As DAO.Database
        Dim strPath As String, strPathFile As String
        Dim intX As Integer
        Dim blnRetVal As Boolean

10910   blnRetVal = True

10920   strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
10930   strPathFile = strPath & LNK_SEP & gstrFile_Icon
10940   If Dir(strPathFile) <> vbNullString Then
10950     Set dbs = CurrentDb
10960 On Error Resume Next
10970     dbs.Properties("AppIcon") = strPathFile
10980     If ERR <> 0 Then
10990 On Error GoTo ERRH
11000       intX = AppProperty_Add("AppIcon", dbText, strPathFile)
11010     Else
11020 On Error GoTo ERRH
11030     End If
11040     Application.RefreshTitleBar
11050   End If

EXITP:
11060   Set dbs = Nothing
11070   AppIcon_Let = blnRetVal
11080   Exit Function

ERRH:
11090   blnRetVal = False
11100   Select Case ERR.Number
        Case Else
11110     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
11120   End Select
11130   Resume EXITP

End Function

Public Function AppProperty_Add(strName As String, varType As Variant, varValue As Variant) As Integer

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "AppProperty_Add"

        Dim dbs As DAO.Database, prp As Object
        Dim blnRetVal As Boolean

        Const conPropNotFoundError As Long = 3270&

11210   blnRetVal = True

11220   Set dbs = CurrentDb
11230   dbs.Properties(strName) = varValue
11240   dbs.Close

EXITP:
11250   Set prp = Nothing
11260   Set dbs = Nothing
11270   AppProperty_Add = blnRetVal
11280   Exit Function

ERRH:
11290   Select Case ERR.Number
        Case conPropNotFoundError
11300     Set prp = dbs.CreateProperty(strName, varType, varValue)
11310     dbs.Properties.Append prp
11320     Resume        ' ** Execution resumes with the statement that caused the error.
          'Resume Next  ' ** Esecution resumes with the statement immediately following the statement that caused the error.
11330   Case Else
11340     blnRetVal = False
11350     Resume EXITP  ' ** Execution resumes at line specified.
11360   End Select

End Function

Public Function AppSupplement_Add() As Boolean
' ** Only for rare situations!

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "AppSupplement_Add"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

11410   blnRetVal = True

11420   Set dbs = CurrentDb
11430   Set doc = dbs.Containers("Databases").Documents("UserDefined")
11440   With doc
11450     Set prp = .CreateProperty("AppSupplement", dbBoolean, True, True)  ' ** True prevents easy Edit or Delete.
11460     .Properties.Append prp
11470   End With
11480   dbs.Close

EXITP:
11490   Set prp = Nothing
11500   Set doc = Nothing
11510   Set dbs = Nothing
11520   AppSupplement_Add = blnRetVal
11530   Exit Function

ERRH:
11540   blnRetVal = False
11550   Select Case ERR.Number
        Case Else
11560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
11570   End Select
11580   Resume EXITP

End Function

Public Function AppSupplement_Get() As Boolean

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "AppSupplement_Get"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

11610   blnRetVal = False

11620   Set dbs = CurrentDb
11630   Set doc = dbs.Containers("Databases").Documents("UserDefined")
11640   With doc
11650     For Each prp In .Properties
11660       With prp
11670         If .Name = "AppSupplement" Then
11680           blnRetVal = .Value
11690           Exit For
11700         End If
11710       End With
11720     Next
11730   End With
11740   dbs.Close

EXITP:
11750   Set prp = Nothing
11760   Set doc = Nothing
11770   Set dbs = Nothing
11780   AppSupplement_Get = blnRetVal
11790   Exit Function

ERRH:
11800   blnRetVal = False
11810   Select Case ERR.Number
        Case Else
11820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
11830   End Select
11840   Resume EXITP

End Function

Public Function AppNoExcel_Add() As Boolean
' ** For versions without Excel reference.

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "AppNoExcel_Add"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

11910   blnRetVal = True

11920   Set dbs = CurrentDb
11930   Set doc = dbs.Containers("Databases").Documents("UserDefined")
11940   With doc
11950     Set prp = .CreateProperty("AppNoExcel", dbBoolean, True, True)  ' ** True prevents easy Edit or Delete.
11960     .Properties.Append prp
11970   End With
11980   dbs.Close

11990   Beep

EXITP:
12000   Set prp = Nothing
12010   Set doc = Nothing
12020   Set dbs = Nothing
12030   AppNoExcel_Add = blnRetVal
12040   Exit Function

ERRH:
12050   blnRetVal = False
12060   Select Case ERR.Number
        Case Else
12070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
12080   End Select
12090   Resume EXITP

End Function

Public Function AppNoExcel_Get() As Boolean

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "AppNoExcel_Get"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

12110   blnRetVal = False

12120   Set dbs = CurrentDb
12130   Set doc = dbs.Containers("Databases").Documents("UserDefined")
12140   With doc
12150     For Each prp In .Properties
12160       With prp
12170         If .Name = "AppNoExcel" Then
12180           blnRetVal = .Value
12190           Exit For
12200         End If
12210       End With
12220     Next
12230   End With
12240   dbs.Close

EXITP:
12250   Set prp = Nothing
12260   Set doc = Nothing
12270   Set dbs = Nothing
12280   AppNoExcel_Get = blnRetVal
12290   Exit Function

ERRH:
12300   blnRetVal = False
12310   Select Case ERR.Number
        Case Else
12320     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
12330   End Select
12340   Resume EXITP

End Function

Public Function AppSubject_Let() As Boolean

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "AppSubject_Let"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

12410   blnRetVal = True

12420   Set dbs = CurrentDb
12430   Set doc = dbs.Containers("Databases").Documents("MSysDb")
12440 On Error Resume Next
12450   doc.Properties("Subject").Value = "Trust Accounting"
12460   If ERR <> 0 Then
12470 On Error GoTo ERRH
12480     With dbs
12490       Set prp = .CreateProperty("Subject", dbText, "Trust Accounting")
12500       .Properties.Append prp
12510     End With
12520   Else
12530 On Error GoTo ERRH
12540   End If
12550   dbs.Close

EXITP:
12560   Set prp = Nothing
12570   Set doc = Nothing
12580   Set dbs = Nothing
12590   AppSubject_Let = blnRetVal
12600   Exit Function

ERRH:
12610   blnRetVal = False
12620   Select Case ERR.Number
        Case Else
12630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
12640   End Select
12650   Resume EXITP

End Function

Public Function AppCategory_Let() As Boolean

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "AppCategory_Let"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

12710   blnRetVal = True

12720   Set dbs = CurrentDb
12730   Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
12740 On Error Resume Next
12750   doc.Properties("Category").Value = "Finance"
12760   If ERR <> 0 Then
12770 On Error GoTo ERRH
12780     With dbs
12790       Set prp = .CreateProperty("Category", dbText, "Finance")
12800       .Properties.Append prp
12810     End With
12820   Else
12830 On Error GoTo ERRH
12840   End If
12850   dbs.Close

EXITP:
12860   Set prp = Nothing
12870   Set doc = Nothing
12880   Set dbs = Nothing
12890   AppCategory_Let = blnRetVal
12900   Exit Function

ERRH:
12910   blnRetVal = False
12920   Select Case ERR.Number
        Case Else
12930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
12940   End Select
12950   Resume EXITP

End Function

Public Function AppKeywords_Let() As Boolean

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "AppKeywords_Let"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim blnRetVal As Boolean

13010   blnRetVal = True

13020   Set dbs = CurrentDb
13030   Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
13040 On Error Resume Next
13050   doc.Properties("Keywords").Value = "Trust;Accounting"
13060   If ERR <> 0 Then
13070 On Error GoTo ERRH
13080     With dbs
13090       Set prp = .CreateProperty("Keywords", dbText, "Trust;Accounting")
13100       .Properties.Append prp
13110     End With
13120   Else
13130 On Error GoTo ERRH
13140   End If
13150   dbs.Close

EXITP:
13160   Set prp = Nothing
13170   Set doc = Nothing
13180   Set dbs = Nothing
13190   AppKeywords_Let = blnRetVal
13200   Exit Function

ERRH:
13210   blnRetVal = False
13220   Select Case ERR.Number
        Case Else
13230     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
13240   End Select
13250   Resume EXITP

End Function

Public Function AppComments_Let() As Boolean

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "AppComments_Let"

        Dim dbs As DAO.Database, doc As DAO.Document, prp As Object
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

13310   blnRetVal = True

13320   Set dbs = CurrentDb
13330   Set doc = dbs.Containers("Databases").Documents("SummaryInfo")
13340   strTmp01 = "Copyright © 1998-2016 Delta Data Inc." & vbCrLf & _
          "All rights reserved" & vbCrLf & _
          "Printed in the U.S.A."
13350 On Error Resume Next
13360   doc.Properties("Comments").Value = strTmp01
13370   If ERR <> 0 Then
13380 On Error GoTo ERRH
13390     With dbs
13400       Set prp = .CreateProperty("Comments", dbText, strTmp01)
13410       .Properties.Append prp
13420     End With
13430   Else
13440 On Error GoTo ERRH
13450   End If
13460   dbs.Close

'Copyright © 1998-2017 Delta Data Inc.
'All rights reserved
'Printed in the U.S.A.
'trust;accounting;investments;fiduciary;
'Finance
'Delta Data, Inc.
'Trust Accounting

EXITP:
13470   Set prp = Nothing
13480   Set doc = Nothing
13490   Set dbs = Nothing
13500   AppComments_Let = blnRetVal
13510   Exit Function

ERRH:
13520   blnRetVal = False
13530   Select Case ERR.Number
        Case Else
13540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler
13550   End Select
13560   Resume EXITP

End Function
