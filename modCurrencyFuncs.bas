Attribute VB_Name = "modCurrencyFuncs"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCurrencyFuncs"

'VGC 10/28/2017: CHANGES!

Private intAct As Integer, intFun As Integer, intMet As Integer, intBMU As Integer, intAlt As Integer, intUni As Integer
' **

Public Function CurrSource_Get(strItem As String) As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "CurrSource_Get"

        Dim blnRetVal As Boolean

110     Select Case strItem
        Case "Active1"
120       Select Case intAct
          Case 0
130         blnRetVal = True
140       Case 1
150         blnRetVal = True
160       Case 2
170         blnRetVal = False
180       End Select
190     Case "Active2"
200       Select Case intAct
          Case 0
210         blnRetVal = False
220       Case 1
230         blnRetVal = True
240       Case 2
250         blnRetVal = False
260       End Select
270     Case "Fund1"
280       Select Case intFun
          Case 0
290         blnRetVal = True
300       Case 1
310         blnRetVal = True
320       Case 2
330         blnRetVal = False
340       End Select
350     Case "Fund2"
360       Select Case intFun
          Case 0
370         blnRetVal = False
380       Case 1
390         blnRetVal = True
400       Case 2
410         blnRetVal = False
420       End Select
430     Case "Metal1"
440       Select Case intMet
          Case 0
450         blnRetVal = True
460       Case 1
470         blnRetVal = True
480       Case 2
490         blnRetVal = False
500       End Select
510     Case "Metal2"
520       Select Case intMet
          Case 0
530         blnRetVal = False
540       Case 1
550         blnRetVal = True
560       Case 2
570         blnRetVal = False
580       End Select
590     Case "BMU1"
600       Select Case intBMU
          Case 0
610         blnRetVal = True
620       Case 1
630         blnRetVal = True
640       Case 2
650         blnRetVal = False
660       End Select
670     Case "BMU2"
680       Select Case intBMU
          Case 0
690         blnRetVal = False
700       Case 1
710         blnRetVal = True
720       Case 2
730         blnRetVal = False
740       End Select
750     Case "Alt1"
760       Select Case intAlt
          Case 0
770         blnRetVal = True
780       Case 1
790         blnRetVal = True
800       Case 2
810         blnRetVal = False
820       End Select
830     Case "Alt2"
840       Select Case intAlt
          Case 0
850         blnRetVal = False
860       Case 1
870         blnRetVal = True
880       Case 2
890         blnRetVal = False
900       End Select
910     Case "Unit1"
920       Select Case intUni
          Case 0
930         blnRetVal = True
940       Case 1
950         blnRetVal = True
960       Case 2
970         blnRetVal = False
980       End Select
990     Case "Unit2"
1000      Select Case intUni
          Case 0
1010        blnRetVal = False
1020      Case 1
1030        blnRetVal = True
1040      Case 2
1050        blnRetVal = False
1060      End Select
1070    End Select

EXITP:
1080    CurrSource_Get = blnRetVal
1090    Exit Function

ERRH:
1100    blnRetVal = False
1110    Select Case ERR.Number
        Case Else
1120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1130    End Select
1140    Resume EXITP

End Function

Public Function CurrSource_Set(intActive As Integer, intFund As Integer, intMetal As Integer, intBond As Integer, intMisc As Integer, intUnit As Integer) As Boolean

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrSource_Set"

        Dim blnRetVal As Boolean

1210    blnRetVal = True

1220    intAct = intActive
1230    intFun = intFund
1240    intMet = intMetal
1250    intBMU = intBond
1260    intAlt = intMisc
1270    intUni = intUnit

EXITP:
1280    CurrSource_Set = blnRetVal
1290    Exit Function

ERRH:
1300    blnRetVal = False
1310    Select Case ERR.Number
        Case Else
1320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1330    End Select
1340    Resume EXITP

End Function

Public Function CoInfoSet() As Boolean
' ** Not called by anything.

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfoSet"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

1410    blnRetVal = True

1420    gstrAccountNo = "20"
1430    gstrAccountName = "Harry T. Abbot Conservatorship"
1440    varTmp00 = DLookup("[currentDate]", "masterasset")
1450    gdatStartDate = CDate(varTmp00)
1460    gdatEndDate = Date

1470    Set dbs = CurrentDb
1480    With dbs
1490      Set rst = .OpenRecordset("CompanyInformation", dbOpenDynaset, dbReadOnly)
1500      With rst
1510        .MoveFirst
1520        gstrCo_Name = ![CoInfo_Name]
1530        gstrCo_Address1 = ![CoInfo_Address1]
1540        gstrCo_Address2 = ![CoInfo_Address2]
1550        gstrCo_City = ![CoInfo_City]
1560        gstrCo_State = ![CoInfo_State]
1570        gstrCo_Zip = ![CoInfo_Zip]
1580        gstrCo_Phone = ![CoInfo_Phone]
1590        .Close
1600      End With
1610      Set rst = Nothing
1620      .Close
1630    End With
1640    Set dbs = Nothing

1650    Beep

EXITP:
1660    Set rst = Nothing
1670    Set dbs = Nothing
1680    CoInfoSet = blnRetVal
1690    Exit Function

ERRH:
1700    blnRetVal = False
1710    Select Case ERR.Number
        Case Else
1720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1730    End Select
1740    Resume EXITP

End Function

Public Function GetWords(varInput As Variant, varCnt As Variant, Optional varRight As Variant) As Variant
' ** Return everything to the left of the varCnt space.
' ** Or return everything to the varRight of the varCnt space.

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "GetWords"

        Dim blnRight As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim varRetVal As Variant

1810    varRetVal = Null

1820    Select Case IsMissing(varRight)
        Case True
1830      blnRight = False
1840    Case False
1850      blnRight = CBool(varRight)
1860    End Select

1870    If IsNull(varInput) = False And IsNull(varCnt) = False Then
1880      strTmp01 = Trim(varInput)
1890      If strTmp01 <> vbNullString Then
1900        Select Case blnRight
            Case True
1910          If varCnt <> 99 Then
1920            intPos01 = CharPos(strTmp01, CLng(varCnt), " ")
1930            varRetVal = Trim(Mid(strTmp01, intPos01))
1940          End If
1950        Case False
1960          Select Case varCnt
              Case 99
1970            varRetVal = strTmp01
1980          Case Else
1990            intPos01 = CharPos(strTmp01, CLng(varCnt), " ")
2000            varRetVal = Trim(Left(strTmp01, intPos01))
2010          End Select
2020        End Select
2030      End If
2040    End If

EXITP:
2050    GetWords = varRetVal
2060    Exit Function

ERRH:
2070    varRetVal = Null
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Function

Public Function PopMark() As Boolean

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "PopMark"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngMax As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

2210    blnRetVal = True

2220    lngMax = 300&

        ' ** Empty tblMark.
2230    TableEmpty "tblMark"  ' ** Module Function: modFileUtilities."

2240    Set dbs = CurrentDb
2250    With dbs
2260      Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbConsistent)
2270      With rst
2280        For lngX = 1& To lngMax
2290          .AddNew
2300          ![unique_id] = lngX
2310          ![mark] = False
2320          .Update
2330        Next
2340        .Close
2350      End With
2360      .Close
2370    End With

2380    Beep

EXITP:
2390    Set rst = Nothing
2400    Set dbs = Nothing
2410    PopMark = blnRetVal
2420    Exit Function

ERRH:
2430    blnRetVal = False
2440    Select Case ERR.Number
        Case Else
2450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2460    End Select
2470    Resume EXITP

End Function

Public Function FindTypo() As Boolean

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "FindTypo"

        Dim rpt As Access.Report, ctl As Access.Control
        Dim blnRetVal As Boolean

2510    blnRetVal = True

2520    Set rpt = Reports(0)
2530    With rpt
2540      For Each ctl In .Controls
2550        With ctl
2560          If .ControlType = acTextBox Then
2570            If InStr(.ControlSource, "PercentMarket3x") > 0 Then
2580              Debug.Print "'HERE!  " & .Name
2590            End If
2600          End If
2610        End With
2620      Next
2630    End With

2640    Beep

EXITP:
2650    Set ctl = Nothing
2660    Set rpt = Nothing
2670    FindTypo = blnRetVal
2680    Exit Function

ERRH:
2690    blnRetVal = False
2700    Select Case ERR.Number
        Case Else
2710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2720    End Select
2730    Resume EXITP

End Function

Public Function CurrRptQrys() As Boolean

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrRptQrys"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

2810    blnRetVal = True

2820    Set dbs = CurrentDb
2830    With dbs
2840      For Each qdf In .QueryDefs
2850        With qdf
2860          If Left(.Name, 23) = "qryRpt_CurrencyRates_16" Then
2870            strTmp01 = .Properties("Description")
2880            If Left(strTmp01, 2) = ".." Then strTmp01 = "qryRpt_CurrencyRates" & Mid(strTmp01, 3)
2890            Debug.Print "' ** " & strTmp01
2900            Debug.Print "'" & .Name
2910          End If
2920        End With
2930      Next
2940      .Close
2950    End With

2960    Beep

EXITP:
2970    Set qdf = Nothing
2980    Set dbs = Nothing
2990    CurrRptQrys = blnRetVal
3000    Exit Function

ERRH:
3010    blnRetVal = False
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Function

Public Function Curr_UpdateTables() As Boolean

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Curr_UpdateTables"

        Dim dbsLoc As DAO.Database, rstLoc As DAO.Recordset, rstLocRaw As DAO.Recordset, rstLnk As DAO.Recordset, rstLnkRaw As DAO.Recordset
        Dim strPath As String, strPathFile As String
        Dim lngDbs As Long, arr_varDb() As Variant
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim lngDiffs As Long, arr_varDiff() As Variant
        Dim lngRecs As Long, lngFlds As Long
        Dim blnAdd As Boolean, blnChanged As Boolean
        Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varDb().
        Const D_ELEMS As Integer = 7  ' ** Array's first-element UBound().
        Const D_PATH As Integer = 0
        Const D_FILE As Integer = 1

        ' ** Array: arr_varTbl().
        Const T_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const T_TNAM1 As Integer = 0
        Const T_FLDS  As Integer = 1
        Const T_RECS  As Integer = 2
        Const T_AUTO  As Integer = 3
        Const T_TNAM2 As Integer = 4
        Const T_ADD   As Integer = 5

        ' ** Array: arr_varDiff().
        Const DF_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const DF_FIL As Integer = 0
        Const DF_TBL As Integer = 1
        Const DF_REC As Integer = 2
        Const DF_FLD As Integer = 3

3110  On Error GoTo 0

3120    blnRetVal = True

3130    blnChanged = False  ' ** True: Report differences, but don't change; False: Update.

3140    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
3150    DoEvents

3160    strPath = "C:\Program Files\Delta Data\Trust Accountant\Database"

3170    lngDbs = 0&
3180    ReDim arr_varDb(D_ELEMS, 0)

        ' ** Let's assume we're currently linked to an empty.

        ' ***********************************************
        ' ** EMPTIES!
        ' ***********************************************
3190    lngDbs = lngDbs + 1&
3200    lngE = lngDbs - 1&
3210    ReDim Preserve arr_varDb(D_ELEMS, lngE)
3220    arr_varDb(D_PATH, lngE) = strPath
3230    arr_varDb(D_FILE, lngE) = "TrustDta.mdb"

3240    lngDbs = lngDbs + 1&
3250    lngE = lngDbs - 1&
3260    ReDim Preserve arr_varDb(D_ELEMS, lngE)
3270    arr_varDb(D_PATH, lngE) = strPath
3280    arr_varDb(D_FILE, lngE) = "TrustDta_bak.mdb"

3290    lngDbs = lngDbs + 1&
3300    lngE = lngDbs - 1&
3310    ReDim Preserve arr_varDb(D_ELEMS, lngE)
3320    arr_varDb(D_PATH, lngE) = strPath
3330    arr_varDb(D_FILE, lngE) = "TrustDta_empty.mdb"

        ' ***********************************************
        ' ** DEMOS, WmBJohnson!
        ' ***********************************************
3340    lngDbs = lngDbs + 1&
3350    lngE = lngDbs - 1&
3360    ReDim Preserve arr_varDb(D_ELEMS, lngE)
3370    arr_varDb(D_PATH, lngE) = strPath
3380    arr_varDb(D_FILE, lngE) = "TrustDta_bak_WmB.mdb"

3390    lngDbs = lngDbs + 1&
3400    lngE = lngDbs - 1&
3410    ReDim Preserve arr_varDb(D_ELEMS, lngE)
3420    arr_varDb(D_PATH, lngE) = strPath
3430    arr_varDb(D_FILE, lngE) = "TrustDta_demo.mdb"

3440    lngDbs = lngDbs + 1&
3450    lngE = lngDbs - 1&
3460    ReDim Preserve arr_varDb(D_ELEMS, lngE)
3470    arr_varDb(D_PATH, lngE) = strPath
3480    arr_varDb(D_FILE, lngE) = "TrustDta_WmBJohnson.mdb"

        ' ***********************************************
        ' ** GEORGETOWN TRUST!
        ' ***********************************************
3490    lngDbs = lngDbs + 1&
3500    lngE = lngDbs - 1&
3510    ReDim Preserve arr_varDb(D_ELEMS, lngE)
3520    arr_varDb(D_PATH, lngE) = strPath
3530    arr_varDb(D_FILE, lngE) = "TrustDta_bak_george2.mdb"

3540    Debug.Print "'DBS TO UPDATE: " & CStr(lngDbs)
3550    DoEvents

3560    lngTbls = 0&
3570    ReDim arr_varTbl(T_ELEMS, 0)

        ' **************************************
        ' ** COUNTRY!
        ' **************************************
3580    lngTbls = lngTbls + 1&
3590    lngE = lngTbls - 1&
3600    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
3610    arr_varTbl(T_TNAM1, lngE) = "tblCountry"
3620    arr_varTbl(T_FLDS, lngE) = CLng(0)
3630    arr_varTbl(T_RECS, lngE) = CLng(0)
3640    arr_varTbl(T_AUTO, lngE) = CLng(0)  ' ** Field 0 is AutoNumber.
3650    arr_varTbl(T_TNAM2, lngE) = "tblCountry1"
3660    arr_varTbl(T_ADD, lngE) = CBool(False)

        ' **************************************
        ' ** CURRENCY!
        ' **************************************
3670    lngTbls = lngTbls + 1&
3680    lngE = lngTbls - 1&
3690    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
3700    arr_varTbl(T_TNAM1, lngE) = "tblCurrency"
3710    arr_varTbl(T_FLDS, lngE) = CLng(0)
3720    arr_varTbl(T_RECS, lngE) = CLng(0)
3730    arr_varTbl(T_AUTO, lngE) = CLng(0)  ' ** Field 0 is AutoNumber.
3740    arr_varTbl(T_TNAM2, lngE) = "tblCurrency1"
3750    arr_varTbl(T_ADD, lngE) = CBool(False)

        ' **************************************
        ' ** CURRENCY_SYMBOL!
        ' **************************************
3760    lngTbls = lngTbls + 1&
3770    lngE = lngTbls - 1&
3780    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
3790    arr_varTbl(T_TNAM1, lngE) = "tblCurrency_Symbol"
3800    arr_varTbl(T_FLDS, lngE) = CLng(0)
3810    arr_varTbl(T_RECS, lngE) = CLng(0)
3820    arr_varTbl(T_AUTO, lngE) = CLng(1)  ' ** Field 1 is AutoNumber.
3830    arr_varTbl(T_TNAM2, lngE) = "tblCurrency_Symbol1"
3840    arr_varTbl(T_ADD, lngE) = CBool(False)

        ' **************************************
        ' ** COUNTRY_CURRENCY!
        ' **************************************
3850    lngTbls = lngTbls + 1&
3860    lngE = lngTbls - 1&
3870    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
3880    arr_varTbl(T_TNAM1, lngE) = "tblCountry_Currency"
3890    arr_varTbl(T_FLDS, lngE) = CLng(0)
3900    arr_varTbl(T_RECS, lngE) = CLng(0)
3910    arr_varTbl(T_AUTO, lngE) = CLng(2)  ' ** Field 2 is AutoNumber.
3920    arr_varTbl(T_TNAM2, lngE) = "tblCountry_Currency1"
3930    arr_varTbl(T_ADD, lngE) = CBool(False)

        ' **************************************
        ' ** COUNTRY_CURRENCY_PRIMARY!
        ' **************************************
3940    lngTbls = lngTbls + 1&
3950    lngE = lngTbls - 1&
3960    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
3970    arr_varTbl(T_TNAM1, lngE) = "tblCountry_Currency_Primary"
3980    arr_varTbl(T_FLDS, lngE) = CLng(0)
3990    arr_varTbl(T_RECS, lngE) = CLng(0)
4000    arr_varTbl(T_AUTO, lngE) = CLng(3)  ' ** Field 3 is AutoNumber.
4010    arr_varTbl(T_TNAM2, lngE) = "tblCountry_Currency_Primary1"
4020    arr_varTbl(T_ADD, lngE) = CBool(False)

        ' **************************************
        ' ** CURRENCY_COUNTRY_PRIMARY!
        ' **************************************
4030    lngTbls = lngTbls + 1&
4040    lngE = lngTbls - 1&
4050    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
4060    arr_varTbl(T_TNAM1, lngE) = "tblCurrency_Country_Primary"
4070    arr_varTbl(T_FLDS, lngE) = CLng(0)
4080    arr_varTbl(T_RECS, lngE) = CLng(0)
4090    arr_varTbl(T_AUTO, lngE) = CLng(3)  ' ** Field 3 is AutoNumber.
4100    arr_varTbl(T_TNAM2, lngE) = "tblCurrency_Country_Primary1"
4110    arr_varTbl(T_ADD, lngE) = CBool(False)

        ' **************************************
        ' ** CURRENCY_HISTORY!
        ' **************************************
4120    lngTbls = lngTbls + 1&
4130    lngE = lngTbls - 1&
4140    ReDim Preserve arr_varTbl(T_ELEMS, lngE)
4150    arr_varTbl(T_TNAM1, lngE) = "tblCurrency_History"
4160    arr_varTbl(T_FLDS, lngE) = CLng(0)
4170    arr_varTbl(T_RECS, lngE) = CLng(0)
4180    arr_varTbl(T_AUTO, lngE) = CLng(1)  ' ** Field 1 is AutoNumber.
4190    arr_varTbl(T_TNAM2, lngE) = "tblCurrency_History1"
4200    arr_varTbl(T_ADD, lngE) = CBool(False)

4210    Debug.Print "'TBLS TO UPDATE: " & CStr(lngTbls)
4220    DoEvents

        ' ** Delete any ..1's hanging around.
4230    For lngX = 0& To (lngTbls - 1&)
4240      If TableExists(CStr(arr_varTbl(T_TNAM2, lngX))) = True Then  ' ** Module Function: modFileUtilities.
4250        TableDelete CStr(arr_varTbl(T_TNAM2, lngX))  ' ** Module Function: modFileUtilities.
4260      End If
4270    Next
4280    CurrentDb.TableDefs.Refresh

4290    lngDiffs = 0&
4300    ReDim arr_varDiff(DF_ELEMS, 0)

4310    For lngW = 0& To (lngDbs - 1&)
          ' ** FOR EACH DATABASE...

4320      If lngW = 0& Then
            ' ** Base all others on this local one.
4330        Set dbsLoc = CurrentDb
4340        With dbsLoc
4350          For lngX = 0& To (lngTbls - 1&)
4360            lngFlds = 0&: lngRecs = 0&
4370            Set rstLoc = .OpenRecordset(arr_varTbl(T_TNAM1, lngX), dbOpenDynaset, dbReadOnly)
4380            With rstLoc
4390              lngFlds = .Fields.Count
4400              If .BOF = True And .EOF = True Then
4410                lngRecs = 0&
4420              Else
4430                .MoveLast
4440                lngRecs = .RecordCount
4450              End If
4460              .Close
4470            End With
4480            Set rstLoc = Nothing
4490            arr_varTbl(T_FLDS, lngX) = lngFlds
4500            arr_varTbl(T_RECS, lngX) = lngRecs
4510          Next  ' ** lngX.
4520          .Close
4530        End With  ' ** dbsLoc.
4540        Set dbsLoc = Nothing
4550      Else
4560        Set dbsLoc = CurrentDb
4570        With dbsLoc
4580          strPathFile = arr_varDb(D_PATH, lngW) & LNK_SEP & arr_varDb(D_FILE, lngW)
4590          For lngX = 0& To (lngTbls - 1&)
4600            DoCmd.TransferDatabase acLink, "Microsoft Access", strPathFile, acTable, arr_varTbl(T_TNAM1, lngX), arr_varTbl(T_TNAM2, lngX)
4610            DoEvents
4620          Next  ' ** lngX.
4630          .TableDefs.Refresh
4640          DoEvents
4650          For lngX = 0& To (lngTbls - 1&)
                ' ** FOR EACH TABLE...
4660            blnAdd = False
4670            Set rstLocRaw = .OpenRecordset(arr_varTbl(T_TNAM1, lngX), dbOpenDynaset, dbReadOnly)
4680            rstLocRaw.sort = "[" & rstLocRaw.Fields(arr_varTbl(T_AUTO, lngX)).Name & "]"
4690            Set rstLoc = rstLocRaw.OpenRecordset
4700            Set rstLnkRaw = .OpenRecordset(arr_varTbl(T_TNAM2, lngX), dbOpenDynaset, dbConsistent)
4710            rstLnkRaw.sort = "[" & rstLnkRaw.Fields(arr_varTbl(T_AUTO, lngX)).Name & "]"
4720            Set rstLnk = rstLnkRaw.OpenRecordset
4730            With rstLnk
4740              If .Fields.Count <> arr_varTbl(T_FLDS, lngX) Then
4750                Stop
4760              Else
4770                .MoveLast
4780                lngRecs = .RecordCount
4790                .MoveFirst
4800                rstLoc.MoveFirst
4810                For lngY = 1& To arr_varTbl(T_RECS, lngX)
                      ' ** FOR EACH RECORD...
4820                  If blnAdd = False Then
4830                    For lngZ = 0& To (arr_varTbl(T_FLDS, lngX) - 1&)
                          ' ** FOR EACH FIELD...
4840                      If rstLoc.Fields(lngZ).Name <> .Fields(lngZ).Name Then
4850                        Stop
4860                      Else
                            ' ** THIS REPLACES VALUES, BUT I'D LIKE THE OPTION OF JUST NOTING DIFFERENCES!
4870                        If IsNull(rstLoc.Fields(lngZ)) = True And IsNull(.Fields(lngZ)) = True Then
                              ' ** Both Null.
4880                        Else
4890                          If IsNull(rstLoc.Fields(lngZ)) = True And IsNull(.Fields(lngZ)) = False Then
4900                            Select Case blnChanged
                                Case True
4910                              If Right(.Fields(lngZ).Name, 12) <> "datemodified" And Right(.Fields(lngZ).Name, 8) <> "username" Then
4920                                lngDiffs = lngDiffs + 1&
4930                                lngE = lngDiffs - 1&
4940                                ReDim Preserve arr_varDiff(DF_ELEMS, lngE)
4950                                arr_varDiff(DF_FIL, lngE) = arr_varDb(D_FILE, lngW)
4960                                arr_varDiff(DF_TBL, lngE) = arr_varTbl(T_TNAM1, lngX)
4970                                arr_varDiff(DF_REC, lngE) = lngY
4980                                arr_varDiff(DF_FLD, lngE) = .Fields(lngZ).Name
4990                              End If
5000                            Case False
5010                              .Edit
5020                              .Fields(lngZ) = Null
5030                              .Update
5040                            End Select
5050                          ElseIf IsNull(rstLoc.Fields(lngZ)) = False And IsNull(.Fields(lngZ)) = True Then
5060                            Select Case blnChanged
                                Case True
5070                              If Right(.Fields(lngZ).Name, 12) <> "datemodified" And Right(.Fields(lngZ).Name, 8) <> "username" Then
5080                                lngDiffs = lngDiffs + 1&
5090                                lngE = lngDiffs - 1&
5100                                ReDim Preserve arr_varDiff(DF_ELEMS, lngE)
5110                                arr_varDiff(DF_FIL, lngE) = arr_varDb(D_FILE, lngW)
5120                                arr_varDiff(DF_TBL, lngE) = arr_varTbl(T_TNAM1, lngX)
5130                                arr_varDiff(DF_REC, lngE) = lngY
5140                                arr_varDiff(DF_FLD, lngE) = .Fields(lngZ).Name
5150                              End If
5160                            Case False
5170                              .Edit
5180                              .Fields(lngZ) = rstLoc.Fields(lngZ)
5190                              .Update
5200                            End Select
5210                          Else
                                ' ** Hopefully, this will never happen on an AutoNumber field.
5220                            If .Fields(lngZ) <> rstLoc.Fields(lngZ) Then
5230                              Select Case blnChanged
                                  Case True
5240                                If Right(.Fields(lngZ).Name, 12) <> "datemodified" And Right(.Fields(lngZ).Name, 8) <> "username" Then
5250                                  lngDiffs = lngDiffs + 1&
5260                                  lngE = lngDiffs - 1&
5270                                  ReDim Preserve arr_varDiff(DF_ELEMS, lngE)
5280                                  arr_varDiff(DF_FIL, lngE) = arr_varDb(D_FILE, lngW)
5290                                  arr_varDiff(DF_TBL, lngE) = arr_varTbl(T_TNAM1, lngX)
5300                                  arr_varDiff(DF_REC, lngE) = lngY
5310                                  arr_varDiff(DF_FLD, lngE) = .Fields(lngZ).Name
5320                                End If
5330                              Case False
5340                                .Edit
5350                                .Fields(lngZ) = rstLoc.Fields(lngZ)
5360                                .Update
5370                              End Select
5380                            End If
5390                          End If
5400                        End If
5410                      End If
5420                    Next  ' ** lngZ, Fields.
5430                  Else
5440                    Select Case blnChanged
                        Case True
5450                      lngDiffs = lngDiffs + 1&
5460                      lngE = lngDiffs - 1&
5470                      ReDim Preserve arr_varDiff(DF_ELEMS, lngE)
5480                      arr_varDiff(DF_FIL, lngE) = arr_varDb(D_FILE, lngW)
5490                      arr_varDiff(DF_TBL, lngE) = arr_varTbl(T_TNAM1, lngX)
5500                      arr_varDiff(DF_REC, lngE) = lngY
5510                      arr_varDiff(DF_FLD, lngE) = "{AddNew}"
5520                    Case False
5530                      .AddNew
5540                      For lngZ = 0& To (arr_varTbl(T_FLDS, lngX) - 1&)  ' ** This includes the AutoNumber field.
5550                        .Fields(lngZ) = rstLoc.Fields(lngZ)
5560                      Next
5570                      .Update
5580                    End Select
5590                  End If  ' ** blnAdd.
5600                  If lngY = lngRecs And lngY < arr_varTbl(T_RECS, lngX) Then
5610                    blnAdd = True
5620                  Else
5630                    If lngY < lngRecs Then .MoveNext
5640                  End If
5650                  If lngY < arr_varTbl(T_RECS, lngX) Then rstLoc.MoveNext
5660                Next  ' ** lngY, Records.
5670              End If  ' ** Fields.Count.
5680              .Close
5690            End With  ' ** rstLnk.
5700            Set rstLnk = Nothing
5710            rstLnkRaw.Close
5720            Set rstLnkRaw = Nothing
5730            rstLoc.Close
5740            Set rstLoc = Nothing
5750            rstLocRaw.Close
5760            Set rstLocRaw = Nothing
5770            If blnAdd = True Then
5780              arr_varTbl(T_ADD, lngX) = CBool(True)
5790            End If
5800          Next  ' ** lngX, Tables.
5810          .Close
5820        End With  ' ** dbsLoc.
5830        Set dbsLoc = Nothing
5840      End If  ' ** lngX.

5850      For lngX = 0& To (lngTbls - 1&)
5860        If arr_varTbl(T_ADD, lngX) = True Then
              'DOESN'T WORK WITH OTHER DBS!
              'ChangeSeed_Ext (arr_varTbl(T_TNAM2, lngX))  ' ** Module Function: modAutonumberFieldFuncs.
5870          Debug.Print "'  RECS ADDED: " & arr_varTbl(T_TNAM2, lngX)
5880          arr_varTbl(T_ADD, lngX) = CBool(False)
5890        End If
5900      Next  ' ** lngX.

5910      If lngW > 0& Then

5920        For lngX = 0& To (lngTbls - 1&)
5930          DoCmd.DeleteObject acTable, arr_varTbl(T_TNAM2, lngX)
5940          DoEvents
5950        Next  ' ** lngX.
5960        CurrentDb.TableDefs.Refresh

5970        Debug.Print "'" & arr_varDb(D_FILE, lngW)
5980        DoEvents

5990      End If

6000    Next  ' ** lngW, Databases.

6010    Debug.Print "'DIFFS: " & CStr(lngDiffs)
6020    DoEvents

6030    If lngDiffs > 0& Then
6040      For lngX = 0& To (lngDiffs - 1&)
6050        Debug.Print "'" & arr_varDiff(DF_FIL, lngX) & "  " & arr_varDiff(DF_TBL, lngX) & "  " & _
              "REC: " & CStr(arr_varDiff(DF_REC, lngX)) & "  " & arr_varDiff(DF_FLD, lngX)
6060        DoEvents
6070      Next
6080    End If

6090    Beep
6100    Debug.Print "'DONE!"

        'DBS TO UPDATE: 7
        'TBLS TO UPDATE: 7
        'TrustDta_bak.mdb
        'TrustDta_empty.mdb
        'TrustDta_bak_WmB.mdb
        'TrustDta_demo.mdb
        'TrustDta_WmBJohnson.mdb
        'TrustDta_bak_george2.mdb
        'DIFFS: 0
        'DONE!

        'DBS TO UPDATE: 7
        'TBLS TO UPDATE: 7
        'TrustDta_bak.mdb
        'TrustDta_empty.mdb
        'TrustDta_bak_WmB.mdb
        'TrustDta_demo.mdb
        'TrustDta_WmBJohnson.mdb
        'TrustDta_bak_george2.mdb
        'DIFFS: 6
        'TrustDta_bak_george2.mdb  tblCurrency_Symbol  REC: 60  currsym_name
        'TrustDta_bak_george2.mdb  tblCurrency_Symbol  REC: 81  currsym_name
        'TrustDta_bak_george2.mdb  tblCurrency_Symbol  REC: 98  currsym_name
        'TrustDta_bak_george2.mdb  tblCurrency_Symbol  REC: 154  currsym_name
        'TrustDta_bak_george2.mdb  tblCurrency_Symbol  REC: 171  currsym_name
        'TrustDta_bak_george2.mdb  tblCurrency_Symbol  REC: 188  curr_code
        'DONE!

EXITP:
6110    Set rstLoc = Nothing
6120    Set rstLocRaw = Nothing
6130    Set rstLnk = Nothing
6140    Set rstLnkRaw = Nothing
6150    Set dbsLoc = Nothing
6160    Curr_UpdateTables = blnRetVal
6170    Exit Function

ERRH:
6180    blnRetVal = False
6190    Select Case ERR.Number
        Case Else
6200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6210    End Select
6220    Resume EXITP

End Function

Public Function HasForEx_All() As Boolean
' ** Returns True/False whether any account has foreign currency.
' ** This sets gblnHasForEx in modStartupFuncs and on frmMenu_Title.

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "HasForEx_All"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

6310    blnRetVal = False

6320    Set dbs = CurrentDb
6330    With dbs
          ' ** qryCurrency_07_03 (Union of qryCurrency_07_01 (Ledger, grouped,
          ' ** just curr_id <> 150, with cnt), qryCurrency_07_02 (LedgerArchive,
          ' ** grouped, just curr_id <> 150, with cnt)), grouped and summed, with cnt.
6340      Set qdf = .QueryDefs("qryCurrency_07_04")
6350      Set rst = qdf.OpenRecordset
6360      With rst
6370        If .BOF = True And .EOF = True Then
              ' ** Same as Zero
6380        Else
6390          .MoveFirst
6400          If IsNull(![cnt]) = True Then
                ' ** Same as Zero
6410          Else
6420            If ![cnt] > 0 Then
6430              blnRetVal = True
6440            End If
6450          End If
6460        End If
6470        .Close
6480      End With
6490      Set rst = Nothing
6500      Set qdf = Nothing
6510      If blnRetVal = False Then
            ' ** MasterAsset, grouped, just curr_id <> 150, with cnt.
6520        Set qdf = .QueryDefs("qryCurrency_08")
6530        Set rst = qdf.OpenRecordset
6540        With rst
6550          If .BOF = True And .EOF = True Then
                ' ** Same as Zero.
6560          Else
6570            .MoveFirst
6580            If ![cnt] > 0 Then
6590              blnRetVal = True
6600            End If
6610          End If
6620          .Close
6630        End With
6640        Set rst = Nothing
6650        Set qdf = Nothing
6660      End If
6670      .Close
6680    End With

        'gblnHasForEx
        'gblnHasForEx_This

EXITP:
6690    Set rst = Nothing
6700    Set qdf = Nothing
6710    Set dbs = Nothing
6720    HasForEx_All = blnRetVal
6730    Exit Function

ERRH:
6740    blnRetVal = False
6750    Select Case ERR.Number
        Case Else
6760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6770    End Select
6780    Resume EXITP

End Function

Public Function HasForEx_Acct(varAccountNo As Variant, Optional varSource As Variant) As Boolean
' ** Returns True/False whether account has foreign currency.
' ** Source: L = Ledger, A = ActiveAssets.

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "HasForEx_Acct"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strSource As String
        Dim blnRetVal As Boolean

6810    blnRetVal = False

6820    If IsNull(varAccountNo) = False Then

6830      Select Case IsMissing(varSource)
          Case True
6840        strSource = "L"
6850      Case False
6860        Select Case IsNull(varSource)
            Case True
6870          strSource = "L"
6880        Case False
6890          strSource = varSource
6900        End Select
6910      End Select

6920      Set dbs = CurrentDb
6930      With dbs
            ' ** tblCurrency_Account, by specified [actno].
6940        Set qdf = .QueryDefs("qryCurrency_14")
6950        With qdf.Parameters
6960          ![actno] = varAccountNo
6970        End With
6980        Set rst = qdf.OpenRecordset
6990        With rst
7000          If .BOF = True And .EOF = True Then
                ' ** Shouldn't happen.
7010          Else
7020            .MoveFirst
7030            Select Case strSource
                Case "L"
7040              If ![curracct_jno] > 0& Then
7050                blnRetVal = True
7060              End If
7070            Case "A"
7080              If ![curracct_aa] > 0& Then
7090                blnRetVal = True
7100              End If
7110            End Select
7120          End If
7130          .Close
7140        End With
7150        Set rst = Nothing
7160        Set qdf = Nothing
7170        .Close
7180      End With

7190    End If

EXITP:
7200    Set rst = Nothing
7210    Set qdf = Nothing
7220    Set dbs = Nothing
7230    HasForEx_Acct = blnRetVal
7240    Exit Function

ERRH:
7250    blnRetVal = False
7260    Select Case ERR.Number
        Case Else
7270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7280    End Select
7290    Resume EXITP

End Function

Public Function HasForEx_Suppress(varAccountNo As Variant) As Boolean
' ** Returns True/False whether to suppress foreign currency fields.

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "HasForEx_Suppress"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

7310    blnRetVal = False

7320    If IsNull(varAccountNo) = False Then
7330      If Trim(varAccountNo) <> vbNullString Then
7340        Set dbs = CurrentDb
7350        With dbs
              ' ** tblCurrency_Account, by specified [actno].
7360          Set qdf = .QueryDefs("qryCurrency_16")
7370          With qdf.Parameters
7380            ![actno] = varAccountNo
7390          End With
7400          Set rst = qdf.OpenRecordset
7410          With rst
7420            If .BOF = True And .EOF = True Then
                  ' ** New account? Should've been added already.
7430            Else
7440              .MoveFirst
7450              blnRetVal = ![curracct_suppress]  ' ** True means Yes, suppress foreign currency columns.
7460            End If
7470            .Close
7480          End With
7490          .Close
7500        End With
7510      End If
7520    End If

EXITP:
7530    Set rst = Nothing
7540    Set qdf = Nothing
7550    Set dbs = Nothing
7560    HasForEx_Suppress = blnRetVal
7570    Exit Function

ERRH:
7580    blnRetVal = False
7590    Select Case ERR.Number
        Case Else
7600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7610    End Select
7620    Resume EXITP

End Function

Public Sub HasForEx_Load()
' ** This populates tblCurrency_Account, which has the
' ** foreign currency counts for each account, as well as
' ** curracct_suppress, indicating whether to automatically
' ** suppress foreign currency fields on reports for those
' ** accounts that have no foreign currency transactions.
' ** This mechanism only applies to users with at least
' ** one foreign currency asset, otherwise it's ignored.

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "HasForEx_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnSuppress As Boolean

7710    Set dbs = CurrentDb
7720    With dbs

          ' ** Get the suppression preferences for this user.
7730      blnSuppress = Pref_Suppress  ' ** Module Function: modPreferenceFuncs.
7740      DoEvents

          ' ** qryCurrency_11_03 (Account, linked to qryCurrency_11_02 (qryCurrency_11_01 (Union of
          ' ** qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01 (Ledger, grouped by
          ' ** accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02 (LedgerArchive,
          ' ** grouped by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed, by
          ' ** accountno, with cnt_jno), qryCurrency_10_01 (ActiveAssets, grouped by accountno,
          ' ** just curr_id <> 150, with cnt_aa)), grouped, with Max(cnt_jno), Max(cnt_aa)),
          ' ** with acct_sort), not in tblCurrency_Account.
7750      Set qdf = .QueryDefs("qryCurrency_13_01")
7760      Set rst = qdf.OpenRecordset
7770      If rst.BOF = True And rst.EOF = True Then
            ' ** No new accounts.
7780        rst.Close
7790        Set rst = Nothing
7800        Set qdf = Nothing
7810      Else
7820        rst.Close
7830        Set rst = Nothing
7840        Set qdf = Nothing
7850        DoEvents

            ' ** If a new account has foreign currency, it gets a False (don't suppress).
            ' ** If it doesn't, it gets a setting based on the preference, above.

            ' ** Any with foreign currency:
            ' ** Append qryCurrency_13_02 (qryCurrency_13_01 (qryCurrency_11_03 (Account, linked to qryCurrency_11_02
            ' ** (qryCurrency_11_01 (Union of qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01
            ' ** (Ledger, grouped by accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02 (LedgerArchive,
            ' ** grouped by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed, by accountno, with
            ' ** cnt_jno), qryCurrency_10_01 (ActiveAssets, grouped by accountno, just curr_id <> 150, with cnt_aa)),
            ' ** grouped,with Max(cnt_jno), Max(cnt_aa)), with acct_sort), not in tblCurrency_Account), just those
            ' ** with foreign currency, for curracct_suppress = False) to tblCurrency_Account.
7860        Set qdf = .QueryDefs("qryCurrency_13_04")
7870        qdf.Execute
7880        Set qdf = Nothing
7890        DoEvents

            ' ** All those without foreign currency.
            ' ** Append qryCurrency_13_03 (qryCurrency_13_01 (qryCurrency_11_03 (Account, linked to qryCurrency_11_02
            ' ** (qryCurrency_11_01 (Union of qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01 (Ledger,
            ' ** grouped by accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02 (LedgerArchive, grouped
            ' ** by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed, by accountno, with cnt_jno),
            ' ** qryCurrency_10_01 (ActiveAssets, grouped by accountno, just curr_id <> 150, with cnt_aa)), grouped,
            ' ** with Max(cnt_jno), Max(cnt_aa)), with acct_sort), not in tblCurrency_Account), just those with no
            ' ** foreign currency, by specified [supr]) to tblCurrency_Account.
7900        Set qdf = .QueryDefs("qryCurrency_13_05")
7910        With qdf.Parameters
7920          ![supr] = blnSuppress
7930        End With
7940        qdf.Execute
7950        Set qdf = Nothing
7960        DoEvents

7970      End If

          ' ** With an established preference, the only discrepancies we can check
          ' ** are those with foreign currency having a curracct_suppress = True.

          ' ** tblCurrency_Account, linked to qryCurrency_11_03 (Account, linked to qryCurrency_11_02
          ' ** (qryCurrency_11_01 (Union of qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01
          ' ** (Ledger, grouped by accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02 (LedgerArchive,
          ' ** grouped by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed, by accountno, with
          ' ** cnt_jno), qryCurrency_10_01 (ActiveAssets, grouped by accountno, just curr_id <> 150, with cnt_aa)),
          ' ** grouped, with Max(cnt_jno), Max(cnt_aa)), with acct_sort), just discrepancies.
7980      Set qdf = .QueryDefs("qryCurrency_11_04")
7990      Set rst = qdf.OpenRecordset
8000      If rst.BOF = True And rst.EOF = True Then
            ' ** No discrepancies.
8010        rst.Close
8020        Set rst = Nothing
8030        Set qdf = Nothing
8040      Else
8050        rst.Close
8060        Set rst = Nothing
8070        Set qdf = Nothing
8080        DoEvents

            ' ** Update qryCurrency_12_01 (tblCurrency_Account, with DLookups() to qryCurrency_11_04
            ' ** (tblCurrency_Account, linked to qryCurrency_11_03 (Account, linked to qryCurrency_11_02
            ' ** (qryCurrency_11_01 (Union of qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01
            ' ** (Ledger, grouped by accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02
            ' ** (LedgerArchive, grouped by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed,
            ' ** by accountno, with cnt_jno), qryCurrency_10_01 (ActiveAssets, grouped by accountno, just
            ' ** curr_id <> 150, with cnt_aa)), grouped, with Max(cnt_jno), Max(cnt_aa)), with acct_sort),
            ' ** just discrepancies)), for curracct_jno.
8090        Set qdf = .QueryDefs("qryCurrency_12_02")
8100        qdf.Execute dbFailOnError
8110        Set qdf = Nothing
8120        DoEvents
            ' ** Update qryCurrency_12_01 (tblCurrency_Account, with DLookups() to qryCurrency_11_04
            ' ** (tblCurrency_Account, linked to qryCurrency_11_03 (Account, linked to qryCurrency_11_02
            ' ** (qryCurrency_11_01 (Union of qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01
            ' ** (Ledger, grouped by accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02
            ' ** (LedgerArchive, grouped by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed,
            ' ** by accountno, with cnt_jno), qryCurrency_10_01 (ActiveAssets, grouped by accountno, just
            ' ** curr_id <> 150, with cnt_aa)), grouped, with Max(cnt_jno), Max(cnt_aa)), with acct_sort),
            ' ** just discrepancies)), for curracct_aa.
8130        Set qdf = .QueryDefs("qryCurrency_12_03")
8140        qdf.Execute dbFailOnError
8150        Set qdf = Nothing
8160        DoEvents
            ' ** Update qryCurrency_12_01 (tblCurrency_Account, with DLookups() to qryCurrency_11_04
            ' ** (tblCurrency_Account, linked to qryCurrency_11_03 (Account, linked to qryCurrency_11_02
            ' ** (qryCurrency_11_01 (Union of qryCurrency_09_04 (qryCurrency_09_03 (Union of qryCurrency_09_01
            ' ** (Ledger, grouped by accountno, just curr_id <> 150, with cnt_jno), qryCurrency_09_02
            ' ** (LedgerArchive, grouped by accountno, just curr_id <> 150, with cnt_jno)), grouped and summed,
            ' ** by accountno, with cnt_jno), qryCurrency_10_01 (ActiveAssets, grouped by accountno, just
            ' ** curr_id <> 150, with cnt_aa)), grouped, with Max(cnt_jno), Max(cnt_aa)), with acct_sort),
            ' ** just discrepancies)), for curracct_suppress.
8170        Set qdf = .QueryDefs("qryCurrency_12_04")
8180        qdf.Execute dbFailOnError
8190        Set qdf = Nothing
8200        DoEvents

8210      End If

8220      .Close
8230    End With
8240    DoEvents

EXITP:
8250    Set rst = Nothing
8260    Set qdf = Nothing
8270    Set dbs = Nothing
8280    Exit Sub

ERRH:
8290    Select Case ERR.Number
        Case Else
8300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8310    End Select
8320    Resume EXITP

End Sub

Public Function CurrSym_Get(varCurrID As Variant) As Variant

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrSym_Get"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCurrID As Long
        Dim lngX As Long
        Dim varRetVal As Variant

        Static lngSyms As Long
        Static arr_varSym As Variant

        ' ** Array: arr_varSym().
        Const S_CID As Integer = 0
        'Const S_COD As Integer = 1
        Const S_SYM As Integer = 2
        'Const S_LEN As Integer = 3

8410    varRetVal = Null

8420    If IsNull(varCurrID) = False Then
8430      If varCurrID > 0 Then

8440        lngCurrID = varCurrID

8450        If lngSyms = 0& Or IsEmpty(arr_varSym) = True Then
8460          Set dbs = CurrentDb
8470          With dbs
                ' ** tblCurrency_Symbol, with len_sym (Max: 5).
8480            Set qdf = .QueryDefs("qryCurrency_Symbol_03")
8490            Set rst = qdf.OpenRecordset
8500            With rst
8510              .MoveLast
8520              lngSyms = .RecordCount
8530              .MoveFirst
8540              arr_varSym = .GetRows(lngSyms)
                  ' ***************************************************
                  ' ** Array: arr_varSym()
                  ' **
                  ' **   Field  Element  Name              Constant
                  ' **   =====  =======  ================  ==========
                  ' **     1       0     curr_id           S_CID
                  ' **     2       1     curr_code         S_COD
                  ' **     3       2     currsym_symbol    S_SYM
                  ' **     4       3     len_sym           S_LEN
                  ' **
                  ' ***************************************************
8550              .Close
8560            End With
8570            Set rst = Nothing
8580            Set qdf = Nothing
8590            .Close
8600          End With
8610          Set dbs = Nothing
8620        End If

8630        For lngX = 0& To (lngSyms - 1&)
8640          If arr_varSym(S_CID, lngX) = lngCurrID Then
8650            varRetVal = arr_varSym(S_SYM, lngX)
8660            Exit For
8670          End If
8680        Next

8690      End If
8700    End If

EXITP:
8710    Set rst = Nothing
8720    Set qdf = Nothing
8730    Set dbs = Nothing
8740    CurrSym_Get = varRetVal
8750    Exit Function

ERRH:
8760    varRetVal = Null
8770    Select Case ERR.Number
        Case Else
8780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8790    End Select
8800    Resume EXITP

End Function

Public Function CurrSymFont_Load() As Boolean

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrSymFont_Load"

        Dim dbs As DAO.Database, qdf1 As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngRecs As Long
        Dim blnSkip As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

8910  On Error GoTo 0

8920    blnRetVal = True

8930    Set dbs = CurrentDb
8940    With dbs

8950      Set qdf1 = .QueryDefs("qryCurrency_Font_Symbol_02")
8960      Set rst1 = qdf1.OpenRecordset
8970      rst1.MoveLast
8980      lngRecs = rst1.RecordCount
8990      rst1.MoveFirst

9000      blnSkip = False
9010      If blnSkip = False Then
9020        Set rst2 = .OpenRecordset("tblCurrency_Symbol_Font2", dbOpenDynaset, dbConsistent)
9030        With rst2
9040          .AddNew
9050          For lngX = 1& To lngRecs
9060            .Fields("curr_name_" & Right("000" & CStr(lngX), 3)) = rst1![currsym_name]
9070            If lngX < lngRecs Then rst1.MoveNext
9080          Next
9090          ![currfont2_datemodified] = Now()
9100          .Update
9110          .Close
9120        End With
9130      End If  ' ** blnSkip.

9140      blnSkip = True
9150      If blnSkip = False Then
9160        Set rst2 = .OpenRecordset("tblCurrency_Symbol_Font1", dbOpenDynaset, dbConsistent)
9170        With rst2
9180          .AddNew
9190          For lngX = 1& To lngRecs
9200            .Fields("currsym_symbol_" & Right("000" & CStr(lngX), 3)) = rst1![currsym_symbol]
9210            If lngX < lngRecs Then rst1.MoveNext
9220          Next
9230          ![currfont1_datemodified] = Now()
9240          .Update
9250          .Close
9260        End With
9270      End If  ' ** blnSkip.

9280      rst1.Close

9290      .Close
9300    End With

9310    Debug.Print "'DONE!"
9320    DoEvents

9330    Beep

EXITP:
9340    Set rst1 = Nothing
9350    Set rst2 = Nothing
9360    Set qdf1 = Nothing
9370    Set dbs = Nothing
9380    CurrSymFont_Load = blnRetVal
9390    Exit Function

ERRH:
9400    blnRetVal = False
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Function

Public Function CurrSymFont_Ctls() As Boolean

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrSymFont_Ctls"

        Dim frm As Access.Form, ctl As Access.Control
        Dim lngTpp As Long, lngCtlNum As Long
        Dim blnSkip As Boolean
        Dim strTmp01 As String, strTmp02 As String
        Dim lngW As Long, lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        'Const C_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        'Const C_COL  As Integer = 0
        'Const C_LFT  As Integer = 1
        'Const C_LAST As Integer = 2

9510  On Error GoTo 0

9520    blnRetVal = True

        'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
9530    lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

9540    Set frm = Forms(0)
9550    With frm

9560      blnSkip = True
9570      If blnSkip = False Then
9580        .currsym_symbol_box_106.Name = "currsym_symbol_box_001"
9590        lngX = 1&
9600        For Each ctl In .Detail.Controls
9610          With ctl
9620            If .ControlType = acRectangle Then
9630              If Left(.Name, 19) = "currsym_symbol_box_" And Right(.Name, 1) = "x" Then
9640                lngX = lngX + 1&
9650                .Name = "currsym_symbol_box_" & Right("000" & CStr(lngX), 3)
9660              End If
9670            End If
9680          End With
9690        Next
9700      End If  ' ** blnSkip.

9710      blnSkip = False
9720      If blnSkip = False Then
9730        For Each ctl In .Detail.Controls
9740          With ctl
9750            If .ControlType = acTextBox Then
                  'If Left(.Name, 10) = "curr_name_" Then
                  '  lngCtlNum = CLng(Right(.Name, 3))
                  '  .ControlSource = .Name
9760              If Left(.Name, 15) = "currsym_symbol_" Then
9770                strTmp01 = .Name
9780                .InSelection = True
9790                DoCmd.RunCommand acCmdChangeToLabel
9800                frm.Controls(strTmp01).InSelection = False
9810              End If
9820            End If
9830          End With
9840        Next

9850      End If

9860      blnSkip = True
9870      If blnSkip = False Then
9880        lngCtlNum = 0&

9890        For lngW = 1& To 3&
9900          Select Case lngW
              Case 1&
9910            strTmp01 = "curr_name_"
9920          Case 2&
9930            strTmp01 = "currsym_symbol_box_"
9940          Case 3&
9950            strTmp01 = "currsym_symbol_"
9960          End Select
9970          For lngX = 1& To 210&
9980            Set ctl = .Controls(strTmp01 & Right("000" & CStr(lngX), 3))
9990            With ctl
10000             lngCtlNum = lngX
10010             If lngCtlNum = 1& Or lngCtlNum = 22& Or lngCtlNum = 43& Or lngCtlNum = 64& Or lngCtlNum = 85& Then
                    ' ** Don't move these.
10020             Else
10030               If lngCtlNum < 106& Then
10040                 strTmp02 = strTmp01 & Right("000" & CStr(lngCtlNum - 1&), 3)
10050                 .Top = frm.Controls(strTmp02).Top + (26& * lngTpp)
10060               Else
10070                 If lngCtlNum = 106& Or lngCtlNum = 127& Or lngCtlNum = 148& Or lngCtlNum = 169& Or lngCtlNum = 190& Then
10080                   strTmp02 = strTmp01 & "021"
10090                   .Top = (frm.Controls(strTmp02).Top + frm.Controls(strTmp02).Height) + (20& * lngTpp)
10100                 Else
10110                   strTmp02 = strTmp01 & Right("000" & CStr(lngCtlNum - 1&), 3)
10120                   .Top = frm.Controls(strTmp02).Top + (26& * lngTpp)
10130                 End If
10140               End If
10150             End If
10160           End With
10170         Next  ' ** lngX.
10180       Next  ' ** lngW.
10190       .pg1_2.Top = .curr_name_106.Top - (6& * lngTpp)
10200       .Detail_vline01.Top = .curr_name_106.Top - (11& * lngTpp)
10210       .Detail_vline02.Top = .Detail_vline01.Top
10220       .Detail_hline01.Top = .curr_name_106.Top - (11& * lngTpp)
10230       .Detail_hline02.Top = .Detail_hline01.Top + lngTpp
10240     End If  ' ** blnSkip.

10250   End With

10260   Debug.Print "'DONE!"
10270   DoEvents

10280   Beep

EXITP:
10290   Set ctl = Nothing
10300   Set frm = Nothing
10310   CurrSymFont_Ctls = blnRetVal
10320   Exit Function

ERRH:
10330   blnRetVal = False
10340   Select Case ERR.Number
        Case Else
10350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10360   End Select
10370   Resume EXITP

End Function

Public Function Qry_ChkParams() As Boolean

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "Qry_ChkParams"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngCnt As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

10410 On Error GoTo 0

10420   blnRetVal = True

10430   Set dbs = CurrentDb
10440   With dbs
10450     For Each qdf In .QueryDefs
10460       With qdf
10470         If Left(.Name, 4) <> "~TMP" Then
10480           Select Case .Name
                Case "qryBackupRestore_05"
                  ' ** Looking for 'LedgerArchive_backup'.
10490           Case "zz_qry_VBComponent_MsgBox_02", "zz_qry_VBComponent_MsgBox_03"
                  ' ** Looking for VBAMsgBox_Parse().
10500           Case "zz_qry_VBComponent_MsgBox_20", "zz_qry_VBComponent_MsgBox_21", "zz_qry_VBComponent_MsgBox_22"
                  ' ** Looking for VBA_MsgBox_TitleParens().
10510           Case "zz_qry_VBComponent_MsgBox_24", "zz_qry_VBComponent_MsgBox_25"
                  ' ** Looking for VBA_MsgBox_TitleNum().
10520           Case "zz_qry_VBComponent_MsgBox_29", "zz_qry_VBComponent_MsgBox_30"
                  ' ** Looking for VBA_MsgBox_CrLf().
10530           Case Else
10540             lngCnt = .Parameters.Count
10550             If lngCnt > 0 Then
10560               For lngX = 0& To (lngCnt - 1&)
10570                 If .Parameters(lngX) = "NoAssets" Then
10580                   Debug.Print "'" & qdf.Name & "  PARAM: " & .Parameters(lngX)
10590                   DoEvents
10600                 End If
10610               Next
10620             End If
10630           End Select
10640         End If
10650       End With
10660     Next
10670     .Close
10680   End With

        'QRY: 'qryStatementParameters_AssetList_70_01' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_02' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_03' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_04' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_05' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_08' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_09' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_13' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_14' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_21' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_22' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_23' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_24' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_25' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_28' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_29' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_29_03' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_33' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_34' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_36_01' NoAssets
        'QRY: 'qryStatementParameters_AssetList_70_36_02' NoAssets
        'QRY: 'qryStatementParameters_AssetList_74_65' NoAssets
        'DONE!
10690   Debug.Print "'DONE!"
10700   DoEvents

10710   Beep

EXITP:
10720   Set qdf = Nothing
10730   Set dbs = Nothing
10740   Qry_ChkParams = blnRetVal
10750   Exit Function

ERRH:
10760   blnRetVal = False
10770   Select Case ERR.Number
        Case Else
10780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10790   End Select
10800   Resume EXITP

End Function
