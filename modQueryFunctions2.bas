Attribute VB_Name = "modQueryFunctions2"
Option Compare Database
Option Explicit

'VGC 10/29/2017: CHANGES!

Private Const THIS_NAME As String = "modQueryFunctions2"
' **

Public Function CoInfo(Optional varGroupBy As Variant) As String
' ** Return a string containing the Company Information
' ** that can be plugged into any SQL statement.
' ** See also CoInfoGet(), below.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo"

        Dim blnGroupBy As Boolean
        Dim strRetVal As String

110     If gstrCo_Name = vbNullString Then
120       CoOptions_Read  ' ** Module Function: modStartupFuncs.
130     End If

140     If IsMissing(varGroupBy) = True Then
150       blnGroupBy = False
160     Else
170       blnGroupBy = CBool(varGroupBy)
180     End If

190     Select Case blnGroupBy
        Case False
200       strRetVal = " '" & gstrCo_Name & "' As CompanyName, '" & gstrCo_Address1 & "' As CompanyAddress1, '" & _
            gstrCo_Address2 & "' As CompanyAddress2, '" & gstrCo_City & "' As CompanyCity, '" & gstrCo_State & "' As CompanyState, '" & _
            gstrCo_Zip & "' As CompanyZip, '" & gstrCo_Country & "' As CompanyCountry, '" & _
            gstrCo_PostalCode & "' As CompanyPostalCode, '" & gstrCo_Phone & "' As CompanyPhone "
          ''MyCompany' As CompanyName, 'MyAddr1' As CompanyAddress1, 'MyAddr2' As CompanyAddress2, 'MyCity' As CompanyCity, 'MyState' As CompanyState, 'MyZip' As CompanyZip, 'MyPhone' As CompanyPhone, 'MyCountry' As CompanyCountry, 'MyPostalCode' As CompanyPostalCode
210     Case True
220       strRetVal = " '" & gstrCo_Name & "', '" & gstrCo_Address1 & "', '" & gstrCo_Address2 & "', '" & gstrCo_City & "', '" & gstrCo_State & "', " & _
            "'" & gstrCo_Zip & "', '" & gstrCo_Country & "', '" & gstrCo_PostalCode & "' '" & gstrCo_Phone & "' "
          ''MyCompany', 'MyAddr1', 'MyAddr2', 'MyCity', 'MyState', 'MyZip', 'MyPhone', 'MyCountry', 'MyPostalCode'
230     End Select

        'CoInfoGet('gstrCo_Name') AS CompanyName, CoInfoGet('gstrCo_Address1') AS CompanyAddress1, CoInfoGet('gstrCo_Address2') AS CompanyAddress2, CoInfoGet('gstrCo_City') AS CompanyCity, CoInfoGet('gstrCo_State') AS CompanyState, CoInfoGet('gstrCo_Zip') AS CompanyZip, CoInfoGet('gstrCo_Phone') AS CompanyPhone, CoInfoGet('gstrCo_Country') AS CompanyCountry, CoInfoGet('gstrCo_PostalCode') AS CompanyPostalCode

EXITP:
240     CoInfo = strRetVal
250     Exit Function

ERRH:
260     Select Case ERR.Number
        Case Else
270       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
280     End Select
290     Resume EXITP

End Function

Public Function CoInfoGet(varPubVar As Variant) As Variant
' ** This function was created so that report queries could get the
' ** company info without requiring the SQL be written in the code.
' ** See also CoInfo(), above.

300   On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfoGet"

        Dim varRetVal As Variant

310     varRetVal = Null

320     If gstrCo_Name = vbNullString Then
330       CoOptions_Read  ' ** Module Function: modStartupFuncs.
340     End If

350     If IsNull(varPubVar) = False Then
360       Select Case varPubVar
          Case "gstrCo_Name"
            'CompanyName: CoInfoGet('gstrCo_Name')
370         varRetVal = gstrCo_Name
380       Case "gstrCo_Address1"
            'CompanyAddress1: CoInfoGet('gstrCo_Address1')
390         varRetVal = gstrCo_Address1
400       Case "gstrCo_Address2"
            'CompanyAddress2: CoInfoGet('gstrCo_Address2')
410         varRetVal = gstrCo_Address2
420       Case "gstrCo_City"
            'CompanyCity: CoInfoGet('gstrCo_City')
430         varRetVal = gstrCo_City
440       Case "gstrCo_State"
            'CompanyState: CoInfoGet('gstrCo_State')
450         varRetVal = gstrCo_State
460       Case "gstrCo_Zip"
            'CompanyZip: CoInfoGet('gstrCo_Zip')
470         varRetVal = gstrCo_Zip
480       Case "gstrCo_Country"
            'CompanyCountry: CoInfoGet('gstrCo_Country')
490         varRetVal = gstrCo_Country
500       Case "gstrCo_PostalCode"
            'CompanyPostalCode: CoInfoGet('gstrCo_PostalCode')
510         varRetVal = gstrCo_PostalCode
520       Case "gstrCo_Phone"
            'CompanyPhone: CoInfoGet('gstrCo_Phone')
530         varRetVal = gstrCo_Phone
540       Case Else
550         varRetVal = vbNullString
560       End Select
570     End If

        'CoInfoGet('gstrCo_Name')
        'CoInfoGet('gstrCo_Address1')
        'CoInfoGet('gstrCo_Address2')
        'CoInfoGet('gstrCo_City')
        'CoInfoGet('gstrCo_State')
        'CoInfoGet('gstrCo_Zip')
        'CoInfoGet('gstrCo_Country')
        'CoInfoGet('gstrCo_PostalCode')
        'CoInfoGet('gstrCo_Phone')

        'CoInfoGet('gstrCo_Name') AS CompanyName, CoInfoGet('gstrCo_Address1') AS CompanyAddress1, CoInfoGet('gstrCo_Address2') AS CompanyAddress2, CoInfoGet('gstrCo_City') AS CompanyCity, CoInfoGet('gstrCo_State') AS CompanyState, CoInfoGet('gstrCo_Zip') AS CompanyZip, CoInfoGet('gstrCo_Phone') AS CompanyPhone, CoInfoGet('gstrCo_Country') AS CompanyCountry, CoInfoGet('gstrCo_PostalCode') AS CompanyPostalCode

EXITP:
580     CoInfoGet = varRetVal
590     Exit Function

ERRH:
600     varRetVal = Null
610     Select Case ERR.Number
        Case Else
620       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
630     End Select
640     Resume EXITP

End Function

Public Function CoInfoGet_Block() As String
' ** CoInfoBlock Label:
' **   120   Top
' **   60    Left, from right edge
' **   3060  Width
' **   1260  Height
'Reports(0).CoInfoBlock.Left = 120&
'Reports(0).CoInfoBlock.Top = 120&
'Reports(0).CoInfoBlock.Height = 1260&
'Reports(0).CoInfoBlock.Width = 3060&
'Reports(0).CoInfoBlock.Left = (Reports(0).Width - Reports(0).CoInfoBlock.Width) - 60&
'130       .CoInfoBlock.Caption = CoInfoGet_Block  ' ** Module Function: modQueryFunctions2.

700   On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfoGet_Block"

        Dim strRetVal As String

710     strRetVal = vbNullString

720     If gstrCo_Name = vbNullString Then
730       CoOptions_Read  ' ** Module Function: modStartupFuncs.
740     End If

750     If gstrCo_Country = vbNullString Then
760       strRetVal = IIf(gstrCo_Name <> vbNullString, gstrCo_Name & vbCrLf, vbNullString) & _
            IIf(gstrCo_Address1 <> vbNullString, gstrCo_Address1 & vbCrLf, vbNullString) & _
            IIf(gstrCo_Address2 <> vbNullString, gstrCo_Address2 & vbCrLf, vbNullString) & _
            IIf(gstrCo_City <> vbNullString, gstrCo_City & ", " & gstrCo_State & "  " & FormatZip9(gstrCo_Zip) & vbCrLf, vbNullString) & _
            IIf(gstrCo_Phone <> vbNullString, FormatPhoneNum(gstrCo_Phone) & vbCrLf, vbNullString)  ' ** Module Functions: modStringFuncs.
770     Else
780       strRetVal = IIf(gstrCo_Name <> vbNullString, gstrCo_Name & vbCrLf, vbNullString) & _
            IIf(gstrCo_Address1 <> vbNullString, gstrCo_Address1 & vbCrLf, vbNullString) & _
            IIf(gstrCo_Address2 <> vbNullString, gstrCo_Address2 & vbCrLf, vbNullString) & _
            IIf(gstrCo_City <> vbNullString, gstrCo_City & vbCrLf, vbNullString) & _
            IIf(gstrCo_Country <> vbNullString, Trim(gstrCo_Country & "  " & gstrCo_PostalCode) & vbCrLf, vbNullString) & _
            IIf(gstrCo_Phone <> vbNullString, FormatPhoneNum(gstrCo_Phone) & vbCrLf, vbNullString)  ' ** Module Functions: modStringFuncs.
790     End If

EXITP:
800     CoInfoGet_Block = strRetVal
810     Exit Function

ERRH:
820     strRetVal = RET_ERR
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Function

Public Function Obj_Type(varType As Variant) As Variant

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Obj_Type"

        Dim varRetVal As Variant

910     varRetVal = Null

920     If IsNull(varType) = False Then
930       If IsNumeric(varType) = True Then
940         Select Case varType
            Case acNothing
950           varRetVal = "acNothing"
960         Case acSQL
970           varRetVal = "acSQL"
980         Case acDefault
990           varRetVal = "acDefault"
1000        Case acTable
1010          varRetVal = "acTable"
1020        Case acQuery
1030          varRetVal = "acQuery"
1040        Case acForm
1050          varRetVal = "acForm"
1060        Case acReport
1070          varRetVal = "acReport"
1080        Case acMacro
1090          varRetVal = "acMacro"
1100        Case acModule
1110          varRetVal = "acModule"
1120        Case acDataAccessPage
1130          varRetVal = "acDataAccessPage"
1140        Case acServerView
1150          varRetVal = "acServerView"
1160        Case acDiagram
1170          varRetVal = "acDiagram"
1180        Case acStoredProcedure
1190          varRetVal = "acStoredProcedure"
1200        Case Else
1210          varRetVal = "{unknown}"
1220        End Select
1230      Else
1240        varRetVal = RET_ERR
1250      End If
1260    End If

EXITP:
1270    Obj_Type = varRetVal
1280    Exit Function

ERRH:
1290    varRetVal = RET_ERR
1300    Select Case ERR.Number
        Case Else
1310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1320    End Select
1330    Resume EXITP

End Function

Public Function AssetDescription(varRI As Variant, varJT As Variant, varAN As Variant, varAD As Variant, varSF As Variant, varDesc As Variant, varRate As Variant, varDue As Variant, varJC As Variant) As String
' ** varRI:   Reccuring Item
' ** varJT:   Journaltype
' ** varAN:   Asset Number
' ** varAD:   Assestdate
' ** varSF:   shareface
' ** varJC:   Jcomment
' ** varDesc: Description
' ** varRate: rate
' ** varDue:  due date

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "AssetDescription"

        Dim strRetVal As String

        ' ** Origianl code:
        'XX: IIf(IsNull([RecurringItem]),"",IIf([journaltype]="Received",[RecurringItem],_
        '    IIf([journaltype]="Paid",[RecurringItem],[RecurringItem]))) & IIf([statement].[assetno] Is Not Null,
        '    IIf([assetdate] Is Not Null,Format([assetdate],"mm/dd/yyyy") & " ") &
        '    IIf([statement].[shareface]-CLng([statement].[shareface])=0,Format([statement].[shareface],"#,##0"),
        '    Format([statement].[shareface],"#,##0.0000")) & " " & CStr([statement].[Description]) &
        '    IIf([statement].[rate]>0," " & Format([statement].[rate],"#,##0.000%")),"") &
        '    IIf([statement].[due] Is Not Null,"  Due " & Format([statement].[due],"mm/dd/yyyy")) & " " & [Jcomment]

1410    strRetVal = vbNullString

1420    If IsNull(varRI) Then
1430      strRetVal = vbNullString
1440    Else
1450      If varJT = "Received" Then
1460        strRetVal = varRI
1470      Else
1480        If varJT = "Paid" Then
1490          strRetVal = varRI
1500        Else
1510          strRetVal = varRI
1520        End If
1530      End If
1540    End If

        ' ** Trim it now to get rid of varRI that are all spaces.
1550    strRetVal = Trim(strRetVal)

1560    If Not IsNull(varAN) Then
1570      If Not IsNull(varAD) Then
1580        strRetVal = Format(varAD, "mm/dd/yyyy") & " "
1590      End If
1600      If (varSF - CLng(varSF)) = 0 Then
1610        strRetVal = strRetVal & Format(varSF, "#,##0")
1620      Else
1630        strRetVal = strRetVal & Format(varSF, "#,##0.0000")
1640      End If
1650    End If

1660    If IsNull(varDesc) Then
1670      strRetVal = strRetVal & " "
1680    Else
1690      strRetVal = strRetVal & " " & CStr(varDesc)
1700    End If

1710    If varRate > 0 Then
1720      strRetVal = strRetVal & " " & Format(varRate, "#,##0.000%")
1730    End If

1740    If Not IsNull(varDue) Then
1750      strRetVal = strRetVal & "  Due " & Format(varDue, "mm/dd/yyyy")
1760    End If

1770    strRetVal = strRetVal & " " & varJC

EXITP:
1780    AssetDescription = strRetVal
1790    Exit Function

ERRH:
1800    strRetVal = vbNullString
1810    Select Case ERR.Number
        Case Else
1820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1830    End Select
1840    Resume EXITP

End Function

Public Function AssetListDesc(varActNo As Variant, varAstNo As Variant, varMastAstDesc As Variant) As String
'IIf([account].[accountno]='INCOME O/U' Or [account].[accountno]='SUSPENSE',[account].[accountno],IIf(IsNull([ActiveAssets].[assetno])=True,'No Assets For This Account',[masterasset].[description])) AS MasterAssetDescription

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "AssetListDesc"

        Dim strRetVal As String

1910    strRetVal = vbNullString

1920    If IsNull(varActNo) = False Then
1930      If varActNo = "INCOME O/U" Or varActNo = "SUSPENSE" Then
1940        strRetVal = varActNo
1950      Else
1960        If IsNull(varAstNo) = True Then
1970          strRetVal = "No Assets For This Account"
1980        Else
1990          If IsNull(varMastAstDesc) = True Then
2000            strRetVal = "Unknown"
2010          Else
2020            strRetVal = varMastAstDesc
2030          End If
2040        End If
2050      End If
2060    End If

EXITP:
2070    AssetListDesc = strRetVal
2080    Exit Function

ERRH:
2090    Select Case ERR.Number
        Case Else
2100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2110    End Select
2120    Resume EXITP

End Function

Public Function JournalNo_Holes() As Variant

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalNo_Holes"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngJNos As Long, arr_varJNo() As Variant
        Dim lngRecs As Long, lngLastJno As Long
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim varRetVal As Variant

2210    varRetVal = Empty

2220    lngJNos = 0&
2230    ReDim arr_varJNo(0)
2240    lngLastJno = 0&

2250    Set dbs = CurrentDb
2260    With dbs
2270      Set qdf = .QueryDefs("zzz_qry_MasterTrust_11")
2280      Set rst1 = qdf.OpenRecordset
          'Set rst1 = .OpenRecordset("ledger", dbOpenDynaset, dbConsistent)
2290      With rst1
2300        .MoveLast
2310        lngRecs = .RecordCount
2320        .MoveFirst
2330        .sort = "[journalno]"
2340        Set rst2 = rst1.OpenRecordset
2350        With rst2
2360          For lngX = 1& To lngRecs
2370            If ![journalno] <> lngLastJno + 1& Then
2380              lngJNos = lngJNos + 1&
2390              lngE = lngJNos - 1&
2400              ReDim Preserve arr_varJNo(lngE)
2410              lngLastJno = lngLastJno + 1&
2420              arr_varJNo(lngE) = lngLastJno
                  ' ** Add missing journalno's till we catch up.
2430              lngZ = lngLastJno + 1&
2440              If lngZ < ![journalno] Then
2450                For lngY = lngZ To (![journalno] - 1&)
2460                  lngJNos = lngJNos + 1&
2470                  lngE = lngJNos - 1&
2480                  ReDim Preserve arr_varJNo(lngE)
2490                  lngLastJno = lngLastJno + 1&
2500                  arr_varJNo(lngE) = lngLastJno
2510                Next
2520                lngLastJno = ![journalno]
2530              Else
                    ' ** Ready to continue.
2540                lngLastJno = ![journalno]
2550              End If
2560            Else
2570              lngLastJno = ![journalno]
2580            End If
2590            If lngX < lngRecs Then .MoveNext
2600          Next
2610          .Close
2620        End With
2630      End With
2640      .Close
2650    End With

2660    varRetVal = arr_varJNo

EXITP:
2670    Set rst2 = Nothing
2680    Set rst1 = Nothing
2690    Set qdf = Nothing
2700    Set dbs = Nothing
2710    JournalNo_Holes = varRetVal
2720    Exit Function

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Function

Public Function JournalNo_Holes_Get() As Long
'DOESN'T WORK!

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalNo_Holes_Get"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim lngX As Long
        Dim lngRetVal As Long

2810    lngRetVal = 0&

2820    Set dbs = CurrentDb
2830    With dbs
2840      Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbConsistent)
2850      With rst
2860        If .BOF = True And .EOF = True Then
              ' ** Well then I shouldn't have been using this!
2870        Else
2880          .MoveLast
2890          lngRecs = .RecordCount
2900          .MoveFirst
2910          For lngX = 1& To lngRecs
2920            If ![mark] = False Then
2930              lngRetVal = ![unique_id]
2940              .Edit
2950              ![mark] = CBool(True)
2960              .Update
2970              Exit For
2980            End If
2990            If lngX < lngRecs Then .MoveNext
3000          Next
3010        End If
3020        .Close
3030      End With
3040      .Close
3050    End With

EXITP:
3060    Set rst = Nothing
3070    Set dbs = Nothing
3080    JournalNo_Holes_Get = lngRetVal
3090    Exit Function

ERRH:
3100    lngRetVal = 0&
3110    Select Case ERR.Number
        Case Else
3120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3130    End Select
3140    Resume EXITP

End Function

Public Function JournalNo_Holes_Save() As Boolean

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalNo_Holes_Save"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngJNos As Long, arr_varJNo As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

3210    blnRetVal = True

3220    arr_varJNo = JournalNo_Holes  ' ** Function: Above.
3230    lngJNos = (UBound(arr_varJNo) + 1&)

3240    Set dbs = CurrentDb
3250    With dbs
          ' ** Empty tblMark.
3260      Set qdf = .QueryDefs("qrySystemUpdate_11x_m_TBL")
3270      qdf.Execute
3280      Set rst = .OpenRecordset("tblMark", dbOpenDynaset, dbAppendOnly)
3290      With rst
3300        For lngX = 0& To (lngJNos - 1&)
3310          .AddNew
3320          ![unique_id] = arr_varJNo(lngX)
3330          ![mark] = CBool(False)
3340          .Update
3350        Next
3360        .Close
3370      End With
3380      .Close
3390    End With

3400    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.

3410    Beep

3420    Debug.Print "'JNO'S RECOVERED: " & CStr(lngJNos)

3430    Debug.Print "'DONE! " & THIS_PROC & "()"

EXITP:
3440    Set rst = Nothing
3450    Set qdf = Nothing
3460    Set dbs = Nothing
3470    JournalNo_Holes_Save = blnRetVal
3480    Exit Function

ERRH:
3490    blnRetVal = False
3500    Select Case ERR.Number
        Case Else
3510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3520    End Select
3530    Resume EXITP

End Function

Public Function JournalNo_Increment(lngCurJNo As Long, strQrySource As String, Optional varForce As Variant, Optional varUseHoles As Variant, Optional varSource As Variant) As Long
' ** Increments [journalno] from the highest existing.
' **   lngCurJNO    : [journalno] of the source Ledger entry, needed to uniquely identify each record being appended.
' **   strQrySource : Query or table containing the source entries, with the field '[jno_source]'.

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "JournalNo_Increment"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngCurLastJNo As Long, lngNewJNo As Long
        Dim blnForce As Boolean, blnUseHoles As Boolean
        Dim blnAltSource As Boolean, strAltSource As String
        Dim lngX As Long, lngE As Long
        Dim lngRetVal As Long

        Static lngJNos As Long, arr_varJNo() As Variant
        Static lngMissingJNos As Long, arr_varMissingJNo As Variant

        Const J_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const J_NO_ORIG As Integer = 0
        Const J_NO_NEW  As Integer = 1

3610    lngRetVal = 0&

3620    Select Case IsMissing(varUseHoles)
        Case True
3630      blnUseHoles = False
3640    Case False
3650      blnUseHoles = varUseHoles
3660    End Select

3670    If blnUseHoles = True And lngMissingJNos = 0& Then
3680      arr_varMissingJNo = JournalNo_Holes  ' ** Function: Below.
3690      lngMissingJNos = (UBound(arr_varMissingJNo) + 1&)
3700    End If

3710    If lngCurJNo > 0& Then

3720      Select Case IsMissing(varForce)
          Case True
3730        blnForce = False
3740      Case False
3750        Select Case IsNull(varForce)
            Case True
3760          blnForce = False
3770        Case False
3780          blnForce = varForce
3790        End Select
3800      End Select

3810      Select Case IsMissing(varSource)
          Case True
3820        blnAltSource = False
3830        strAltSource = vbNullString
3840      Case False
3850        Select Case IsNull(varSource)
            Case True
3860          blnAltSource = False
3870          strAltSource = vbNullString
3880        Case False
3890          If Trim(varSource) = vbNullString Then
3900            blnAltSource = False
3910            strAltSource = vbNullString
3920          Else
3930            blnAltSource = True
3940            strAltSource = varSource
3950          End If
3960        End Select
3970      End Select

3980      If lngJNos = 0& Or blnForce = True Then
3990        ReDim arr_varJNo(J_ELEMS, 0)
4000        Select Case blnAltSource
            Case True
4010          Select Case strAltSource
              Case "00", "ledger"
4020            lngCurLastJNo = DMax("[journalno]", "ledger")
4030          Case "01"
4040            lngCurLastJNo = DMax("[journalno]", "tmpXAdmin_Ledger_01")
4050          Case "02"
4060            lngCurLastJNo = DMax("[journalno]", "tmpXAdmin_Ledger_02")
4070          Case "03"
4080            lngCurLastJNo = DMax("[journalno]", "tmpXAdmin_Ledger_03")
4090          Case "04"
4100            lngCurLastJNo = DMax("[journalno]", "tmpXAdmin_Ledger_04")
4110          End Select
4120        Case False
4130          lngCurLastJNo = DMax("[journalno]", "ledger")
4140        End Select
4150        Select Case blnUseHoles
            Case True
4160          lngNewJNo = 0&
4170        Case False
4180          lngNewJNo = lngCurLastJNo
4190        End Select
4200        Set dbs = CurrentDb
4210        With dbs
4220          Set rst = .OpenRecordset(strQrySource, dbOpenDynaset, dbReadOnly)
4230          With rst
4240            .MoveLast
4250            lngJNos = .RecordCount
4260            .MoveFirst
4270            For lngX = 1& To lngJNos
4280              lngJNos = lngJNos + 1&
4290              lngE = lngJNos - 1&
4300              ReDim Preserve arr_varJNo(J_ELEMS, lngE)
4310              arr_varJNo(J_NO_ORIG, lngE) = ![jno_source]
4320              Select Case blnUseHoles
                  Case True
4330                lngNewJNo = arr_varMissingJNo(lngX - 1&)  ' ** Loop is one-based, array is zero-based.
4340              Case False
4350                lngNewJNo = lngNewJNo + 1&
4360              End Select
4370              arr_varJNo(J_NO_NEW, lngE) = lngNewJNo
4380              If lngX < lngJNos Then .MoveNext
4390            Next
4400            .Close
4410          End With
4420          .Close
4430        End With
4440      End If

4450      For lngX = 0& To (lngJNos - 1&)
4460        If arr_varJNo(J_NO_ORIG, lngX) = lngCurJNo Then
4470          lngRetVal = arr_varJNo(J_NO_NEW, lngX)
4480          Exit For
4490        End If
4500      Next

4510    End If

EXITP:
4520    Set rst = Nothing
4530    Set dbs = Nothing
4540    JournalNo_Increment = lngRetVal
4550    Exit Function

ERRH:
4560    Select Case ERR.Number
        Case Else
4570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4580    End Select
4590    Resume EXITP

End Function

Public Function LinkID_Increment(lngCurLnkNo As Long, strQrySource As String) As Long
' ** Increments [tbllnk_id] from the highest existing.
' **   lngCurLnkNO  : [mtbl_ID] of the source m_TBL entry, needed to uniquely identify each record being appended.
' **   strQrySource : Query or table containing the source entries, with the field '[linkid_source]'.  (Yes, spelled out!)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "LinkID_Increment"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngCurLastLnkNo As Long, lngNewLnkNo As Long
        Dim lngX As Long, lngE As Long
        Dim lngRetVal As Long

        Static lngLnkNos As Long, arr_varLnkNo() As Variant

        ' ** Array: arr_varLnkNo().
        Const L_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const L_NO_ORIG As Integer = 0
        Const L_NO_NEW  As Integer = 1

4610    lngRetVal = 0&

4620    If lngCurLnkNo > 0& Then

4630      If lngLnkNos = 0& Then
4640        ReDim arr_varLnkNo(L_ELEMS, 0)
4650        lngCurLastLnkNo = DMax("[tbllnk_id]", "tblDatabase_Table_Link")
4660        lngNewLnkNo = lngCurLastLnkNo
4670        Set dbs = CurrentDb
4680        With dbs
4690          Set rst = .OpenRecordset(strQrySource, dbOpenDynaset, dbReadOnly)
4700          With rst
4710            .MoveLast
4720            lngLnkNos = .RecordCount
4730            .MoveFirst
4740            For lngX = 1& To lngLnkNos
4750              lngLnkNos = lngLnkNos + 1&
4760              lngE = lngLnkNos - 1&
4770              ReDim Preserve arr_varLnkNo(L_ELEMS, lngE)
4780              arr_varLnkNo(L_NO_ORIG, lngE) = ![linkid_source]
4790              lngNewLnkNo = lngNewLnkNo + 1&
4800              arr_varLnkNo(L_NO_NEW, lngE) = lngNewLnkNo
4810              If lngX < lngLnkNos Then .MoveNext
4820            Next
4830            .Close
4840          End With
4850          .Close
4860        End With
4870      End If

4880      For lngX = 0& To (lngLnkNos - 1&)
4890        If arr_varLnkNo(L_NO_ORIG, lngX) = lngCurLnkNo Then
4900          lngRetVal = arr_varLnkNo(L_NO_NEW, lngX)
4910          Exit For
4920        End If
4930      Next

4940    End If

EXITP:
4950    Set rst = Nothing
4960    Set dbs = Nothing
4970    LinkID_Increment = lngRetVal
4980    Exit Function

ERRH:
4990    Select Case ERR.Number
        Case Else
5000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5010    End Select
5020    Resume EXITP

End Function

Public Function ChkNY() As Boolean

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkNY"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim strFind As String
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const Q_QNAM As Integer = 0

5110    blnRetVal = True

5120    strFind = "Proper("
        '"CaseNum"
        '"cmdDateStart"
        '"cmdDateEnd"
        '"cmdAccountno"
        '"cmdgstrCrtRpt_CashAssets_Beg"
        '"cmdIncomeAtBegin"
        '"cmdNewInput"
        '"cmdIncomeCash"
        '"cmdInvestedIncome"
        '"Proper("
5130    lngQrys = 0&
5140    ReDim arr_varQry(Q_ELEMS, 0)

5150    Set dbs = CurrentDb
5160    With dbs
5170      For Each qdf In .QueryDefs
5180        With qdf
5190          If Left(.Name, Len("qryCourtReport_NY_")) = "qryCourtReport_NY_" Then
5200            If InStr(.Name, "_X_") = 0 Then
5210              If InStr(.SQL, strFind) > 0 Then
5220                lngQrys = lngQrys + 1&
5230                lngE = lngQrys - 1&
5240                ReDim Preserve arr_varQry(Q_ELEMS, lngE)
5250                arr_varQry(Q_QNAM, lngE) = .Name
5260              End If
5270            End If
5280          End If
5290        End With
5300      Next
5310      .Close
5320    End With

        ' ** Binary Sort arr_varQry() array.
5330    For lngX = UBound(arr_varQry, 2) To 1& Step -1&
5340      For lngY = 0& To (lngX - 1&)
5350        If arr_varQry(Q_QNAM, lngY) > arr_varQry(Q_QNAM, (lngY + 1&)) Then
5360          For lngZ = 0& To Q_ELEMS
5370            varTmp00 = arr_varQry(lngZ, lngY)
5380            arr_varQry(lngZ, lngY) = arr_varQry(lngZ, (lngY + 1&))
5390            arr_varQry(lngZ, (lngY + 1&)) = varTmp00
5400            varTmp00 = Empty
5410          Next
5420        End If
5430      Next
5440    Next

5450    For lngX = 0& To (lngQrys - 1&)
5460      Debug.Print "'" & arr_varQry(Q_QNAM, lngX)
5470      If ((lngX + 1&) Mod 100) = 0 Then
5480        Stop
5490      End If
5500    Next

5510    Beep

EXITP:
5520    Set qdf = Nothing
5530    Set dbs = Nothing
5540    ChkNY = blnRetVal
5550    Exit Function

ERRH:
5560    blnRetVal = False
5570    Select Case ERR.Number
        Case Else
5580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5590    End Select
5600    Resume EXITP

End Function

Public Function CoInfo_Find() As Boolean

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "CoInfo_Find"

        Dim vbp As VBProject, vbc As VBComponent, cod As CodeModule
        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strModName As String, strProcName As String, strLine As String, strItem As String
        Dim lngLines As Long, lngDecLines As Long
        Dim lngHits As Long, arr_varHit() As Variant
        Dim lngThisDbsID As Long
        Dim strFind1 As String, strFind2 As String, strFind3 As String
        Dim blnFound As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intPos04 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varHit().
        Const H_ELEMS As Integer = 6  ' ** Array's first-element UBound().
        Const H_VID  As Integer = 0
        Const H_VNAM As Integer = 1
        Const H_PID  As Integer = 2
        Const H_PNAM As Integer = 3
        Const H_LIN  As Integer = 4
        Const H_TXT  As Integer = 5
        Const H_ITM  As Integer = 6

5710  On Error GoTo 0

5720    blnRetVal = True

5730    Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
5740    DoEvents

5750    strFind1 = "CoInfoGet_Block"
5760    strFind2 = "CoInfoGet"
5770    strFind3 = "CoInfo"

5780    lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

5790    lngHits = 0&
5800    ReDim arr_varHit(H_ELEMS, 0)

5810    Set vbp = Application.VBE.ActiveVBProject
5820    With vbp
5830      For Each vbc In .VBComponents
5840        With vbc
5850          strModName = .Name
5860          Set cod = .CodeModule
5870          With cod
5880            lngLines = .CountOfLines
5890            lngDecLines = .CountOfDeclarationLines
5900            For lngX = lngDecLines To lngLines
5910              blnFound = False: strItem = vbNullString: strTmp01 = vbNullString: strTmp02 = vbNullString
5920              strLine = Trim(.Lines(lngX, 1))
5930              If strLine <> vbNullString Then
5940                If Left(strLine, 1) <> "'" Then
5950                  intPos01 = InStr(strLine, strFind1)
5960                  intPos02 = InStr(strLine, strFind2)
5970                  intPos03 = InStr(strLine, strFind3)
5980                  If intPos01 > 0 Then
5990                    strItem = strFind1
6000                    intPos04 = InStr(strLine, "'")
6010                    If intPos04 > 0 Then
                          ' ** Not sure how far to take this.
6020                      If intPos04 < intPos01 Then
6030                        Debug.Print "'" & strLine
6040                        Stop
6050                      Else
6060                        blnFound = True
6070                      End If
6080                    Else
6090                      blnFound = True
6100                    End If
6110                  ElseIf intPos02 > 0 Then
                        ' ** Unsure what it might look like.
6120                    strItem = strFind2
6130                    blnFound = True
6140                  ElseIf intPos03 > 0 Then
6150                    strItem = strFind3
6160                    strTmp01 = Mid(strLine, (intPos03 - 1))  ' ** First character before.
6170                    Select Case strTmp01
                        Case " ", "("
                          ' ** These are OK.
6180                      blnFound = True
6190                    Case "_", ".", ")", "[", "]", "'", Chr(34)
                          ' ** These are not.
6200                    Case ":", ";"
                          ' ** Don't know.
6210                      Debug.Print "'" & strLine
6220                      Stop
6230                    Case Else
6240                      If Asc(strTmp01) >= 48 And Asc(strTmp01) <= 57 Then  ' ** 0 - 9.
                            ' ** Numbers not OK.
6250                      ElseIf Asc(strTmp01) >= 65 And Asc(strTmp01) <= 90 Then  ' ** A - Z.
                            ' ** Letters not OK.
6260                      ElseIf Asc(strTmp01) >= 97 And Asc(strTmp01) <= 122 Then  ' ** a - z.
                            ' ** Letters not OK.
6270                      Else
                            ' ** Don't know what else to look for.
6280                        blnFound = True
6290                      End If
6300                    End Select
6310                    If blnFound = True Then
6320                      blnFound = False
6330                      intLen = Len(strFind3)
6340                      If (intPos03 + intLen) <= Len(strLine) Then
6350                        strTmp02 = Mid(strLine, (intPos03 + intLen), 1)  ' ** First character after.
6360                        Select Case strTmp02
                            Case " ", ")"
                              ' ** These are OK.
6370                          blnFound = True
6380                        Case "_", ".", "(", "[", "]", "'", Chr(34)
                              ' ** These are not.
6390                        Case ":", ";"
                              ' ** Don't know.
6400                          Debug.Print "'" & strLine
6410                          Stop
6420                        Case Else
6430                          If Asc(strTmp02) >= 48 And Asc(strTmp02) <= 57 Then  ' ** 0 - 9.
                                ' ** Numbers not OK.
6440                          ElseIf Asc(strTmp02) >= 65 And Asc(strTmp02) <= 90 Then  ' ** A - Z.
                                ' ** Letters not OK.
6450                          ElseIf Asc(strTmp02) >= 97 And Asc(strTmp02) <= 122 Then  ' ** a - z.
                                ' ** Letters not OK.
6460                          Else
                                ' ** Don't know what else to look for.
6470                            blnFound = True
6480                          End If
6490                        End Select
6500                      End If
6510                    End If
6520                  End If
6530                  If blnFound = True Then
6540                    strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
6550                    If strProcName <> THIS_PROC And strProcName <> strFind1 And _
                            strProcName <> strFind2 And strProcName <> strFind3 Then
6560                      lngHits = lngHits + 1&
6570                      lngE = lngHits - 1&
6580                      ReDim Preserve arr_varHit(H_ELEMS, lngE)
6590                      arr_varHit(H_VID, lngE) = Null
6600                      arr_varHit(H_VNAM, lngE) = strModName
6610                      arr_varHit(H_PID, lngE) = Null
6620                      arr_varHit(H_PNAM, lngE) = strProcName
6630                      arr_varHit(H_LIN, lngE) = lngX
6640                      arr_varHit(H_TXT, lngE) = strLine
6650                      arr_varHit(H_ITM, lngE) = strItem
6660                    End If
6670                  End If  ' ** blnFound.
6680                End If  ' ** Remark.
6690              End If  ' ** vbNullString.
6700            Next  ' ** lngX.
6710          End With  ' ** cod.
6720          Set cod = Nothing
6730        End With  ' ** vbc.
6740        Set vbc = Nothing
6750      Next  ' ** vbc.
6760    End With  ' ** vbp.

6770    Debug.Print "'HITS: " & CStr(lngHits)
6780    DoEvents

6790    If lngHits > 0& Then

6800      Set dbs = CurrentDb
6810      With dbs

6820        Set rst = .OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
6830        With rst
6840          .MoveFirst
6850          For lngX = 0& To (lngHits - 1&)
6860            .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_name] = '" & arr_varHit(H_VNAM, lngX) & "'"
6870            If .NoMatch = False Then
6880              arr_varHit(H_VID, lngX) = ![vbcom_id]
6890            Else
6900              Stop
6910            End If
6920          Next
6930          .Close
6940        End With
6950        Set rst = Nothing

6960        Set rst = .OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
6970        With rst
6980          .MoveFirst
6990          For lngX = 0& To (lngHits - 1&)
7000            .FindFirst "[dbs_id] = " & CStr(lngThisDbsID) & " And [vbcom_id] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
                  "[vbcomproc_name] = '" & arr_varHit(H_PNAM, lngX) & "'"
7010            If .NoMatch = False Then
7020              arr_varHit(H_PID, lngX) = ![vbcomproc_id]
7030            Else
7040              Stop
7050            End If
7060          Next
7070          .Close
7080        End With
7090        Set rst = Nothing

7100        Set rst = .OpenRecordset("zz_tbl_VBComponent_CoInfo", dbOpenDynaset, dbConsistent)
7110        With rst
7120          For lngX = 0& To (lngHits - 1&)
7130            .AddNew
                ' ** ![vbcomci_id : AutoNumber.
7140            ![dbs_id] = lngThisDbsID
7150            ![vbcom_id] = arr_varHit(H_VID, lngX)
7160            ![vbcom_name] = arr_varHit(H_VNAM, lngX)
7170            ![vbcomproc_id] = arr_varHit(H_PID, lngX)
7180            ![vbcomproc_name] = arr_varHit(H_PNAM, lngX)
7190            ![vbcomci_linenum] = arr_varHit(H_LIN, lngX)
7200            ![vbcomci_search] = arr_varHit(H_ITM, lngX)
7210            ![vbcomci_text] = arr_varHit(H_TXT, lngX)
7220            ![vbcomci_datemodified] = Now()
7230            .Update
7240          Next
7250          .Close
7260        End With
7270        Set rst = Nothing

7280        .Close
7290      End With  ' ** dbs.
7300      Set dbs = Nothing

7310    End If  ' ** lngHits.

        'HITS: 200
        'DONE!

        'OTHER PLACES TO CHECK:
        '  QUERIES
        '  REPORT CONTROLS

7320    Beep

7330    Debug.Print "'DONE!"
7340    DoEvents

EXITP:
7350    Set cod = Nothing
7360    Set vbc = Nothing
7370    Set vbp = Nothing
7380    Set rst = Nothing
7390    Set dbs = Nothing
7400    CoInfo_Find = blnRetVal
7410    Exit Function

ERRH:
7420    blnRetVal = False
7430    Select Case ERR.Number
        Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Function

Public Function VBAQryFind() As Boolean
' ** Document all queries found in the VBA code to zz_tbl_VBComponent_Query.

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "VBAQryFind"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngHits As Long, arr_varHit() As Variant
        Dim lngLines As Long, lngDecLines As Long
        Dim strModName As String, strProcName As String, strLine As String
        Dim strFind As String, strCode As String
        Dim lngThisDbsID As Long
        Dim blnContinue As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, intTmp04 As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varHit().
        Const H_ELEMS As Integer = 14  ' ** Array's first-element UBound().
        Const H_DID  As Integer = 0
        Const H_FID  As Integer = 1
        Const H_FNAM As Integer = 2
        Const H_RID  As Integer = 3
        Const H_RNAM As Integer = 4
        Const H_VID  As Integer = 5
        Const H_VNAM As Integer = 6
        Const H_PID  As Integer = 7
        Const H_PNAM As Integer = 8
        Const H_QID  As Integer = 9
        Const H_QNAM As Integer = 10
        Const H_LIN  As Integer = 11
        Const H_COD  As Integer = 12
        Const H_TXT  As Integer = 13
        Const H_MRK  As Integer = 14

7510  On Error GoTo 0

7520    blnRetVal = True

7530    lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

7540    lngHits = 0&
7550    ReDim arr_varHit(H_ELEMS, 0)

7560    Set vbp = Application.VBE.ActiveVBProject
7570    With vbp
7580      For Each vbc In .VBComponents
7590        With vbc
7600          strModName = .Name
7610          Set cod = .CodeModule
7620          With cod
7630            lngLines = .CountOfLines
7640            lngDecLines = .CountOfDeclarationLines
7650            For lngX = 1& To lngLines
7660              strProcName = vbNullString: strCode = vbNullString
7670              strLine = Trim(.Lines(lngX, 1))
7680              If strLine <> vbNullString Then
7690                If Left(strLine, 1) <> "'" Then
7700                  strTmp01 = strLine
7710                  strTmp01 = StringReplace(strTmp01, "QRY_SYS", "QQ1_SYS")  ' ** Module Function: modStringFuncs.
7720                  strTmp01 = StringReplace(strTmp01, "QRY_DDL", "QQ1_DDL")  ' ** Module Function: modStringFuncs.
7730                  strTmp01 = StringReplace(strTmp01, "QRY_BASE", "QQ1_BASE")  ' ** Module Function: modStringFuncs.
7740                  strTmp01 = StringReplace(strTmp01, "QRY:", "QQ1:")  ' ** Module Function: modStringFuncs.
7750                  strTmp01 = StringReplace(strTmp01, "QRYS:", "QQ1S:")  ' ** Module Function: modStringFuncs.
7760                  strTmp01 = StringReplace(strTmp01, "'QRY CNT:", "'QQ1 CNT:")  ' ** Module Function: modStringFuncs.
7770                  strTmp01 = StringReplace(strTmp01, "'QRY NOT FOUND", "'QQ1 NOT FOUND")  ' ** Module Function: modStringFuncs.
7780                  strTmp01 = StringReplace(strTmp01, "QRY NOT FOUND", "QQ1 NOT FOUND")  ' ** Module Function: modStringFuncs.
7790                  strTmp01 = StringReplace(strTmp01, "'QRYS CREATED:", "'QQ1S CREATED:")  ' ** Module Function: modStringFuncs.
7800                  strTmp01 = StringReplace(strTmp01, "'QRYS DELETED:", "'QQ1S DELETED:")  ' ** Module Function: modStringFuncs.
7810                  strTmp01 = StringReplace(strTmp01, "'QRYS EDITED:", "'QQ1S EDITED:")  ' ** Module Function: modStringFuncs.
7820                  strTmp01 = StringReplace(strTmp01, "'QRYS RUN:", "'QQ1S RUN:")  ' ** Module Function: modStringFuncs.
7830                  strTmp01 = StringReplace(strTmp01, "'QRYS CHANGED:", "'QQ1S CHANGED:")  ' ** Module Function: modStringFuncs.
7840                  strTmp01 = StringReplace(strTmp01, "_QRY As", "_QQ1 As")  ' ** Module Function: modStringFuncs.
7850                  strTmp01 = StringReplace(strTmp01, "_QRY  As", "_QQ1  As")  ' ** Module Function: modStringFuncs.
7860                  strTmp01 = StringReplace(strTmp01, "_QRY   As", "_QQ1   As")  ' ** Module Function: modStringFuncs.
7870                  strTmp01 = StringReplace(strTmp01, "_QRY1  As", "_QQ11  As")  ' ** Module Function: modStringFuncs.
7880                  strTmp01 = StringReplace(strTmp01, "_QRY2  As", "_QQ12  As")  ' ** Module Function: modStringFuncs.
7890                  strTmp01 = StringReplace(strTmp01, "_QRY, lng", "_QQ1, lng")  ' ** Module Function: modStringFuncs.
7900                  strTmp01 = StringReplace(strTmp01, "_QRY1, lng", "_QQ11, lng")  ' ** Module Function: modStringFuncs.
7910                  strTmp01 = StringReplace(strTmp01, "_QRY2, lng", "_QQ12, lng")  ' ** Module Function: modStringFuncs.
7920                  strTmp01 = StringReplace(strTmp01, "_QRY, 0", "_QQ1, 0")  ' ** Module Function: modStringFuncs.
7930                  strTmp01 = StringReplace(strTmp01, "_QRY,", "_QQ1,")  ' ** Module Function: modStringFuncs.
7940                  strTmp01 = StringReplace(strTmp01, Chr(34) & "qry" & Chr(34), Chr(34) & "QQ2" & Chr(34))  ' ** Module Function: modStringFuncs.
7950                  strTmp01 = StringReplace(strTmp01, "'qry'", "'QQ2'")  ' ** Module Function: modStringFuncs.
7960                  strTmp01 = StringReplace(strTmp01, "qry" & Chr(34) & ")", "QQ2" & Chr(34) & ")")  ' ** Module Function: modStringFuncs.
7970                  strTmp01 = StringReplace(strTmp01, "Qry_CheckBox", "QQ3_CheckBox")  ' ** Module Function: modStringFuncs.
7980                  strTmp01 = StringReplace(strTmp01, "Qry_ChkParams", "QQ3_ChkParams")  ' ** Module Function: modStringFuncs.
7990                  strTmp01 = StringReplace(strTmp01, "Qry_Chk", "QQ3_Chk")  ' ** Module Function: modStringFuncs.
8000                  strTmp01 = StringReplace(strTmp01, "Qry_Copy", "QQ3_Copy")  ' ** Module Function: modStringFuncs.
8010                  strTmp01 = StringReplace(strTmp01, "Qry_CurrentAppName", "QQ3_CurrentAppName")  ' ** Module Function: modStringFuncs.
8020                  strTmp01 = StringReplace(strTmp01, "Qry_Del_rel", "QQ3_Del_rel")  ' ** Module Function: modStringFuncs.
8030                  strTmp01 = StringReplace(strTmp01, "Qry_Doc_Simple", "QQ3_Doc_Simple")  ' ** Module Function: modStringFuncs.
8040                  strTmp01 = StringReplace(strTmp01, "Qry_Doc", "QQ3_Doc")  ' ** Module Function: modStringFuncs.
8050                  strTmp01 = StringReplace(strTmp01, "Qry_Export_rel", "QQ3_Export_rel")  ' ** Module Function: modStringFuncs.
8060                  strTmp01 = StringReplace(strTmp01, "Qry_FindDesc_rel", "QQ3_FindDesc_rel")  ' ** Module Function: modStringFuncs.
8070                  strTmp01 = StringReplace(strTmp01, "Qry_FindInMod", "QQ3_FindInMod")  ' ** Module Function: modStringFuncs.
8080                  strTmp01 = StringReplace(strTmp01, "Qry_FindStr_rel", "QQ3_FindStr_rel")  ' ** Module Function: modStringFuncs.
8090                  strTmp01 = StringReplace(strTmp01, "Qry_FldList_rel", "QQ3_FldList_rel")  ' ** Module Function: modStringFuncs.
8100                  strTmp01 = StringReplace(strTmp01, "Qry_Import_rel", "QQ3_Import_rel")  ' ** Module Function: modStringFuncs.
8110                  strTmp01 = StringReplace(strTmp01, "Qry_List", "QQ3_List")  ' ** Module Function: modStringFuncs.
8120                  strTmp01 = StringReplace(strTmp01, "Qry_LoadDDL", "QQ3_LoadDDL")  ' ** Module Function: modStringFuncs.
8130                  strTmp01 = StringReplace(strTmp01, "Qry_PropDoc", "QQ3_PropDoc")  ' ** Module Function: modStringFuncs.
8140                  strTmp01 = StringReplace(strTmp01, "Qry_RemExpr_rel", "QQ3_RemExpr_rel")  ' ** Module Function: modStringFuncs.
8150                  strTmp01 = StringReplace(strTmp01, "Qry_Rename", "QQ3_Rename")  ' ** Module Function: modStringFuncs.
8160                  strTmp01 = StringReplace(strTmp01, "Qry_Transfer_rel", "QQ3_Transfer_rel")  ' ** Module Function: modStringFuncs.
8170                  strTmp01 = StringReplace(strTmp01, "Qry_UpdateDesc_rel", "QQ3_UpdateDesc_rel")  ' ** Module Function: modStringFuncs.
8180                  strTmp01 = StringReplace(strTmp01, "Qry_UpdateRef_rel", "QQ3_UpdateRef_rel")  ' ** Module Function: modStringFuncs.
8190                  strTmp01 = StringReplace(strTmp01, "Qry_XAdminGfx", "QQ3_XAdminGfx")  ' ** Module Function: modStringFuncs.
8200                  strTmp01 = StringReplace(strTmp01, "Qry_ZZ_Tbl", "QQ3_ZZ_Tbl")  ' ** Module Function: modStringFuncs.
8210                  strTmp01 = StringReplace(strTmp01, "_Qry() As", "_QQ3() As")  ' ** Module Function: modStringFuncs.
8220                  strTmp01 = StringReplace(strTmp01, "_Qrys() As", "_QQ3s() As")  ' ** Module Function: modStringFuncs.
8230                  strTmp01 = StringReplace(strTmp01, "_Qry =", "_QQ3 =")  ' ** Module Function: modStringFuncs.
8240                  strTmp01 = StringReplace(strTmp01, "_Qrys =", "_QQ3s =")  ' ** Module Function: modStringFuncs.
8250                  strTmp01 = StringReplace(strTmp01, "_Qry" & Chr(34), "_QQ3" & Chr(34))  ' ** Module Function: modStringFuncs.
8260                  strTmp01 = StringReplace(strTmp01, "_Qrys" & Chr(34), "_QQ3s" & Chr(34))  ' ** Module Function: modStringFuncs.
8270                  strTmp01 = StringReplace(strTmp01, "[qry_datemodified]", "[QQ4_datemodified]")  ' ** Module Function: modStringFuncs.
8280                  strTmp01 = StringReplace(strTmp01, "[qry_description]", "[QQ4_description]")  ' ** Module Function: modStringFuncs.
8290                  strTmp01 = StringReplace(strTmp01, "[qry_fldcnt]", "[QQ4_fldcnt]")  ' ** Module Function: modStringFuncs.
8300                  strTmp01 = StringReplace(strTmp01, "[qry_formref]", "[QQ4_formref]")  ' ** Module Function: modStringFuncs.
8310                  strTmp01 = StringReplace(strTmp01, "[qry_formrefcnt]", "[QQ4_formrefcnt]")  ' ** Module Function: modStringFuncs.
8320                  strTmp01 = StringReplace(strTmp01, "[qry_id]", "[QQ4_id]")  ' ** Module Function: modStringFuncs.
8330                  strTmp01 = StringReplace(strTmp01, "[qry_name]", "[QQ4_name]")  ' ** Module Function: modStringFuncs.
8340                  strTmp01 = StringReplace(strTmp01, "[qry_param]", "[QQ4_param]")  ' ** Module Function: modStringFuncs.
8350                  strTmp01 = StringReplace(strTmp01, "[qry_param_clause]", "[QQ4_param_clause]")  ' ** Module Function: modStringFuncs.
8360                  strTmp01 = StringReplace(strTmp01, "[qry_paramcnt]", "[QQ4_paramcnt]")  ' ** Module Function: modStringFuncs.
8370                  strTmp01 = StringReplace(strTmp01, "[qry_sql]", "[QQ4_sql]")  ' ** Module Function: modStringFuncs.
8380                  strTmp01 = StringReplace(strTmp01, "[qry_tblcnt]", "[QQ4_tblcnt]")  ' ** Module Function: modStringFuncs.
8390                  strTmp01 = StringReplace(strTmp01, "[qryprop_name]", "[QQ4prop_name]")  ' ** Module Function: modStringFuncs.
8400                  strTmp01 = StringReplace(strTmp01, "[qryprop_datemodified]", "[QQ4prop_datemodified]")  ' ** Module Function: modStringFuncs.
8410                  strTmp01 = StringReplace(strTmp01, "[qrytype_type]", "[QQ4type_type]")  ' ** Module Function: modStringFuncs.
8420                  intPos01 = InStr(strTmp01, "zzz_qry")                   ' ** intPos01 is position in strTmp01.
8430                  If intPos01 = 0 Then
8440                    intPos01 = InStr(strTmp01, "zz_qry")
8450                    If intPos01 = 0 Then
8460                      intPos01 = InStr(strTmp01, "qry")
8470                    End If
8480                  End If
8490                  If intPos01 > 0 Then
8500                    blnContinue = True
8510                    strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
8520                    If strProcName = vbNullString Then strProcName = "Declaration"
8530                    If strProcName <> THIS_PROC Then
8540                      intPos02 = InStr(strTmp01, " ")
8550                      If intPos02 > 1 Then
8560                        strCode = Trim(Left(strTmp01, intPos02))
8570                        If IsNumeric(strCode) = False Then
8580                          strCode = vbNullString
8590                        End If
8600                      End If
8610                      strTmp03 = vbNullString
8620                      If intPos01 > 1 Then
8630                        strTmp02 = Mid(strTmp01, (intPos01 - 1), 1)  ' ** Character immediately preceding 'qry'.
8640                        intTmp04 = Asc(strTmp02)
8650                        Select Case strTmp02
                            Case ".", ",", ";", ":", "/", "\", ")", "]"
                              ' ** I don't think so. (But not sure!)
8660                          blnContinue = False
8670                        Case " ", "(", "["
                              ' ** Maybe.
8680                          intPos02 = InStr(intPos01, strTmp01, ".")
8690                          intPos03 = InStr(intPos01, strTmp01, " ")
8700                          If intPos02 > 0 Or intPos03 > 0 Then
8710                            If (intPos02 > 0 And intPos03 = 0) Or (intPos02 > 0 And intPos03 > 0 And intPos02 < intPos03) Then
8720                              strTmp02 = Mid(Left(strTmp01, (intPos02 - 1)), intPos01)
8730                              If IsNumeric(Right(strTmp02, 1)) = True Then
                                    ' ** I think so: qryRpt_CurrencyRates_15.[Notes]
8740                                strTmp03 = strTmp02
8750                              ElseIf Asc(Right(strTmp02, 1)) = 98 Or Asc(Right(strTmp02, 1)) = 114 Then
                                    ' ** Let's go with it: qryJournal_Interest_02b.assettype , qryPrintChecks_01_07d_Bank_Account_Number.
8760                                strTmp03 = strTmp02
8770                              End If
8780                            ElseIf ((intPos02 > 0 And intPos03 > 0 And intPos03 < intPos02) Or (intPos03 > 0 And intPos02 = 0)) Then
8790                              strTmp02 = Trim(Mid(Left(strTmp01, intPos03), intPos01))
8800                              strTmp03 = strTmp02
8810                            End If
8820                          Else
8830                            strTmp02 = Trim(Mid(strTmp01, intPos01))
8840                            strTmp03 = strTmp02
8850                          End If
8860                        Case Chr(34)  ' ** Quotes.
                              ' ** Definitely.
8870                          intPos02 = InStr(intPos01, strTmp01, Chr(34))
8880                          If intPos02 > (intPos01 + 3) Then
8890                            strTmp03 = Mid(strTmp01, (intPos01 - 1), ((intPos02 - intPos01) + 2))
8900                          Else
8910                            Debug.Print "'1 " & Mid(strTmp01, (intPos01 - 1))
8920                            DoEvents 'Stop
8930                            blnContinue = False
8940                          End If
8950                        Case Else
8960                          If (intTmp04 >= 65 And intTmp04 <= 90) Or (intTmp04 >= 97 And intTmp04 <= 122) Or (intTmp04 >= 48 And intTmp04 <= 57) Then
                                ' ** No!  A-Z, a-z, 0-9
8970                            blnContinue = False
8980                          Else
8990                            Debug.Print "'2 " & Mid(strTmp01, (intPos01 - 1))
9000                            DoEvents 'Stop
9010                            blnContinue = False
9020                          End If
9030                        End Select
9040                      Else
                            ' ** Maybe.
9050                      End If
9060                      If blnContinue = True Then
9070                        If strTmp03 = "zz_qry's" Or strTmp03 = "QRYS" Or strTmp03 = "zz_qry_" Or strTmp03 = "qry_formref" Or strTmp03 = "qry_id" Or _
                                strTmp03 = "qry_name" Or strTmp03 = "qryparam_datemodified" Or strTmp03 = "qryparam_id" Or strTmp03 = "qryrecsrc_datemodified" Then
9080                          strTmp03 = vbNullString
9090                        End If
9100                        If InStr(strTmp03, "'") > 0 Then
                              'Debug.Print "'" & strLine
9110                          DoEvents
9120                          strTmp03 = vbNullString
9130                        End If
9140                        If strTmp03 = vbNullString Then
9150                          If InStr(strTmp01, "zz_QQ1_SYStem_nn") = 0 And InStr(strTmp01, "zz_qry's") = 0 Then
9160                            Debug.Print "'3 " & Trim(Mid(strTmp01, (intPos01 - 1)))
9170                            DoEvents 'Stop
9180                            blnContinue = False
9190                          End If
9200                        Else
9210                          strTmp01 = StringReplace(strTmp01, "QQ1", "QRY")  ' ** Module Function: modStringFuncs.
9220                          strTmp01 = StringReplace(strTmp01, "QQ2", "qry")  ' ** Module Function: modStringFuncs.
9230                          strTmp01 = StringReplace(strTmp01, "QQ3", "Qry")  ' ** Module Function: modStringFuncs.
9240                          strTmp01 = StringReplace(strTmp01, "QQ4", "qry")  ' ** Module Function: modStringFuncs.
9250                          lngHits = lngHits + 1&
9260                          lngE = lngHits - 1&
9270                          ReDim Preserve arr_varHit(H_ELEMS, lngE)
9280                          arr_varHit(H_DID, lngE) = lngThisDbsID
9290                          arr_varHit(H_FID, lngE) = Null
9300                          If Left(strModName, 5) = "Form_" Then
9310                            arr_varHit(H_FNAM, lngE) = Mid(strModName, 6)
9320                          Else
9330                            arr_varHit(H_FNAM, lngE) = Null
9340                          End If
9350                          arr_varHit(H_RID, lngE) = Null
9360                          If Left(strModName, 7) = "Report_" Then
9370                            arr_varHit(H_RNAM, lngE) = Mid(strModName, 8)
9380                          Else
9390                            arr_varHit(H_RNAM, lngE) = Null
9400                          End If
9410                          arr_varHit(H_VID, lngE) = Null
9420                          arr_varHit(H_VNAM, lngE) = strModName
9430                          arr_varHit(H_PID, lngE) = Null
9440                          arr_varHit(H_PNAM, lngE) = strProcName
9450                          arr_varHit(H_QID, lngE) = Null
9460                          arr_varHit(H_QNAM, lngE) = strTmp03
9470                          arr_varHit(H_LIN, lngE) = lngX
9480                          If strCode <> vbNullString Then
9490                            arr_varHit(H_COD, lngE) = CLng(strCode)
9500                          Else
9510                            arr_varHit(H_COD, lngE) = Null
9520                          End If
9530                          arr_varHit(H_TXT, lngE) = strTmp01
9540                          arr_varHit(H_MRK, lngE) = CBool(False)
9550                        End If
9560                      End If  ' ** blnContinue.
9570                    End If  ' ** THIS_PROC.
9580                  End If  ' ** intPos01.
9590                End If  ' ** Remark.
9600              End If  ' ** vbNullString.
9610            Next  ' ** lngX.
9620          End With  ' ** cod.
9630        End With  ' ** vbc.
9640      Next  ' ** vbc.
9650    End With  ' ** vbp.

9660    Debug.Print "'HITS: " & CStr(lngHits)
9670    DoEvents

9680    If lngHits > 0& Then

9690      Set dbs = CurrentDb
9700      Set rst = dbs.OpenRecordset("tblForm", dbOpenDynaset, dbReadOnly)
9710      With rst
9720        .MoveFirst
9730        For lngX = 0& To (lngHits - 1&)
9740          If IsNull(arr_varHit(H_FNAM, lngX)) = False Then
9750            .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [frm_name] = '" & arr_varHit(H_FNAM, lngX) & "'"
9760            If .NoMatch = False Then
9770              arr_varHit(H_FID, lngX) = ![frm_id]
9780            Else
9790              Stop
9800            End If
9810          End If
9820        Next  ' ** lngX.
9830        .Close
9840      End With  ' ** rst
9850      Set rst = Nothing

9860      Set rst = dbs.OpenRecordset("tblReport", dbOpenDynaset, dbReadOnly)
9870      With rst
9880        .MoveFirst
9890        For lngX = 0& To (lngHits - 1&)
9900          If IsNull(arr_varHit(H_RNAM, lngX)) = False Then
9910            .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [rpt_name] = '" & arr_varHit(H_RNAM, lngX) & "'"
9920            If .NoMatch = False Then
9930              arr_varHit(H_RID, lngX) = ![rpt_id]
9940            Else
9950              Stop
9960            End If
9970          End If
9980        Next  ' ** lngX.
9990        .Close
10000     End With  ' ** rst
10010     Set rst = Nothing

10020     Set rst = dbs.OpenRecordset("tblVBComponent", dbOpenDynaset, dbReadOnly)
10030     With rst
10040       .MoveFirst
10050       For lngX = 0& To (lngHits - 1&)
10060         .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_name] = '" & arr_varHit(H_VNAM, lngX) & "'"
10070         If .NoMatch = False Then
10080           arr_varHit(H_VID, lngX) = ![vbcom_id]
10090         Else
10100           Stop
10110         End If
10120       Next  ' ** lngX.
10130       .Close
10140     End With  ' ** rst
10150     Set rst = Nothing

10160     Set rst = dbs.OpenRecordset("tblVBComponent_Procedure", dbOpenDynaset, dbReadOnly)
10170     With rst
10180       .MoveFirst
10190       For lngX = 0& To (lngHits - 1&)
10200         .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [vbcom_id] = " & CStr(arr_varHit(H_VID, lngX)) & " And " & _
                "[vbcomproc_name] = '" & arr_varHit(H_PNAM, lngX) & "'"
10210         If .NoMatch = False Then
10220           arr_varHit(H_PID, lngX) = ![vbcomproc_id]
10230         Else
10240           Stop
10250         End If
10260       Next  ' ** lngX.
10270       .Close
10280     End With  ' ** rst
10290     Set rst = Nothing

10300     Set rst = dbs.OpenRecordset("tblQuery", dbOpenDynaset, dbReadOnly)
10310     With rst
10320       .MoveFirst
10330       For lngX = 0& To (lngHits - 1&)
10340         If Left(arr_varHit(H_QNAM, lngX), 1) = Chr(34) And Right(arr_varHit(H_QNAM, lngX), 1) = Chr(34) Then
10350           arr_varHit(H_QNAM, lngX) = Mid(Left(arr_varHit(H_QNAM, lngX), (Len(arr_varHit(H_QNAM, lngX)) - 1)), 2)
10360         End If
10370         .FindFirst "[dbs_id] = " & CStr(arr_varHit(H_DID, lngX)) & " And [qry_name] = '" & arr_varHit(H_QNAM, lngX) & "'"
10380         If .NoMatch = False Then
10390           arr_varHit(H_QID, lngX) = ![qry_id]
10400         Else
10410           arr_varHit(H_QID, lngX) = CLng(0)
10420         End If
10430       Next  ' ** lngX.
10440       .Close
10450     End With  ' ** rst
10460     Set rst = Nothing

10470     Set rst = dbs.OpenRecordset("zz_tbl_VBComponent_Query", dbOpenDynaset, dbConsistent)
10480     With rst
10490       For lngX = 0& To (lngHits - 1&)
10500         .AddNew
              ' ** ![vbqry_id] : AutoNumber.
10510         ![dbs_id] = arr_varHit(H_DID, lngX)
10520         ![frm_id] = arr_varHit(H_FID, lngX)
10530         ![rpt_id] = arr_varHit(H_RID, lngX)
10540         ![vbcom_id] = arr_varHit(H_VID, lngX)
10550         ![vbcomproc_id] = arr_varHit(H_PID, lngX)
10560         ![frm_name] = arr_varHit(H_FNAM, lngX)
10570         ![rpt_name] = arr_varHit(H_RNAM, lngX)
10580         ![vbcom_name] = arr_varHit(H_VNAM, lngX)
10590         ![vbcomproc_name] = arr_varHit(H_PNAM, lngX)
10600         ![qry_id] = arr_varHit(H_QID, lngX)
10610         ![qry_name] = arr_varHit(H_QNAM, lngX)
10620         ![vbqry_line] = arr_varHit(H_LIN, lngX)
10630         ![vbqry_code] = arr_varHit(H_COD, lngX)
10640         ![vbqry_text] = arr_varHit(H_TXT, lngX)
10650         ![vbqry_mark] = arr_varHit(H_MRK, lngX)
10660         ![vbqry_datemodified] = Now()
10670         .Update
10680       Next  ' ** lngX.
10690       .Close
10700     End With  ' ** rst.
10710     Set rst = Nothing
10720     dbs.Close
10730     Set dbs = Nothing

10740   End If  ' ** lngHits.

10750   Beep
10760   Debug.Print "'DONE!"

EXITP:
10770   Set cod = Nothing
10780   Set vbc = Nothing
10790   Set vbp = Nothing
10800   Set rst = Nothing
10810   Set dbs = Nothing
10820   VBAQryFind = blnRetVal
10830   Exit Function

ERRH:
10840   blnRetVal = False
10850   Select Case ERR.Number
        Case Else
10860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10870   End Select
10880   Resume EXITP

End Function

Public Function FindLostQry() As Boolean
' ** See also VBA_FindQrys().

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "FindLostQry"

        Dim fso As Scripting.FileSystemObject, fsfds As Scripting.Folders, fsfd1 As Scripting.Folder, fsfd2 As Scripting.Folder
        Dim fsfls As Scripting.FILES, fsfl As Scripting.File
        Dim wrk As DAO.Workspace, dbsLnk As DAO.Database, dbsLoc As DAO.Database
        Dim qdf As DAO.QueryDef, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim lngDirs As Long, arr_varDir() As Variant
        Dim lngFiles As Long, arr_varFile() As Variant
        Dim strPath As String, strFile As String, strPathFile As String, strPathFile_This As String
        Dim strFind As String, strDesc As String
        Dim blnHasSfx As Boolean, blnFound As Boolean, blnSubs As Boolean, blnAdd As Boolean, blnAddAll As Boolean
        Dim intLen As Integer
        Dim strTmp01 As String, lngTmp02 As Long
        Dim lngW As Long, lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const Q_QNAM As Integer = 0
        Const Q_TYP  As Integer = 1
        Const Q_DSC  As Integer = 2
        Const Q_SQL  As Integer = 3
        Const Q_PATH As Integer = 4
        Const Q_FILE As Integer = 5

        ' ** Array: arr_varDir().
        Const D_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const D_DNAM As Integer = 0
        Const D_PATH As Integer = 1

        ' ** Array: arr_varFile().
        Const F_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const F_FNAM As Integer = 0
        Const F_PATH As Integer = 1

10910 On Error GoTo 0

10920   blnRetVal = True

10930   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
10940   DoEvents

10950   strFind = "qryVersion_Input_Notice_01"
10960   intLen = Len(strFind)
10970   blnHasSfx = False

10980   strPathFile_This = CurrentAppPath & LNK_SEP & CurrentAppName  ' ** Module Functions: modFileUtilities.

10990   lngQrys = 0&
11000   ReDim arr_varQry(Q_ELEMS, 0)

11010   lngFiles = 0&
11020   ReDim arr_varFile(F_ELEMS, 0)

11030   lngTmp02 = 0&

11040   For lngW = 1& To 12&

11050     blnSubs = False
11060     Select Case lngW
          Case 1&
11070       strPath = "C:\Program Files\Delta Data\Trust Accountant"  ' ** THIS!
11080     Case 2&
11090       strPath = "C:\Program Files\Delta Data\Trust Accountant\Client_Frontends"
11100     Case 3&
11110       strPath = "C:\Program Files (x86)\Delta Data\Trust Accountant"
11120     Case 4&
11130       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewWorking"
            '  Trust.mdb, Trust_OldStuff.mdb, many others
            '  TrustAux.mdb
11140     Case 5&
11150       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Old_Releases\Ver2210_Bak"
            '  Many!
11160     Case 6&
11170       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Old_Releases\Ver2200_Bak"
            '  Many!
11180     Case 7&
11190       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Old_Releases"
            '  Many subfolders!  DON'T REDO Ver2210_Bak, Ver2200_Bak!
11200       blnSubs = True
11210     Case 8&
11220       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\PreviousVersionDBs"
            '  Many subfolders, mostly backends, but some front
11230       blnSubs = True
11240     Case 9&
11250       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\PreviousVersionDBs\AnotherSourceDisk"
11260       blnSubs = True
11270     Case 10&
11280       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewUpgrade"
            '  Many subfolders with earlier versions
11290       blnSubs = True
11300     Case 11&
11310       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewUpgrade\Old_Releases"
11320       blnSubs = True
11330     Case 12&
11340       strPath = "C:\VictorGCS_Clients\TrustAccountant\NewUpgrade\DemoDatabase"
11350       blnSubs = True
11360     End Select

11370     DBEngine.SystemDB = "C:\Program Files\Delta Data\Trust Accountant\Database\TrustSec.mdw"
11380     Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)

11390     lngDirs = 0&
11400     ReDim arr_varDir(D_ELEMS, 0)

11410     Set fso = CreateObject("Scripting.FileSystemObject")
11420     With fso

11430       Set fsfd1 = .GetFolder(strPath)
11440       Select Case blnSubs
            Case True
11450         Set fsfds = fsfd1.SubFolders
11460         For Each fsfd2 In fsfds
11470           lngDirs = lngDirs + 1&
11480           lngE = lngDirs - 1&
11490           ReDim Preserve arr_varDir(D_ELEMS, lngE)
11500           arr_varDir(D_DNAM, lngE) = fsfd2.Name
11510           arr_varDir(D_PATH, lngE) = fsfd2.Path
11520         Next
11530       Case False
11540         lngDirs = 1&
11550       End Select

11560       For lngX = 0& To (lngDirs - 1&)

11570         blnFound = True
11580         If lngDirs = 1& Then
11590           Set fsfd2 = .GetFolder(strPath)
11600         Else
11610           Set fsfd2 = fsfds(arr_varDir(D_DNAM, lngX))
11620           If lngW = 7& And (fsfd2.Name = "Ver2200_Bak" Or fsfd2.Name = "Ver2210_Bak") Then
11630             blnFound = False
11640           End If
11650         End If
11660         If blnFound = True Then

11670           Set fsfls = fsfd2.FILES
11680           For Each fsfl In fsfls
11690             With fsfl
11700               If Left(.Name, 5) = "Trust" Or .Name = "TAJrnTmp.mdb" Then
11710                 strPath = Parse_Path(.Path)  ' ** Module Functions: modFileUtilities.
11720                 strFile = .Name
11730                 If Parse_Ext(strFile) = "mdb" Then
11740                   strTmp01 = Mid(strFile, 6, 1)

11750                   Select Case strTmp01
                        Case ".", "_", " ", "A"

11760                     strPathFile = strPath & LNK_SEP & strFile
11770                     If strPathFile <> strPathFile_This Then

11780                       Set dbsLnk = wrk.OpenDatabase(strPathFile, False, True)  ' ** {pathfile}, {exclusive}, {read-only}
11790                       With dbsLnk
'11800                         For Each tdf In .TableDefs
'11810                           With tdf
11800                         For Each qdf In .QueryDefs
11810                           With qdf
11820                             If Left(.Name, intLen) = strFind Then
11830                               strDesc = vbNullString
11840                               Select Case blnHasSfx
                                    Case True
11850                                 blnFound = False
11860                                 If .Name = strFind Then
11870                                   blnFound = True
11880                                 Else
11890                                   For lngY = 1& To 4&
11900                                     Select Case lngY
                                          Case 1&
11910                                       If .Name = strFind & "ActiveAssets_04" Then '"b" Then
11920                                         blnFound = True
11930                                         Exit For
11940                                       End If
11950                                     Case 2&
11960                                       If .Name = strFind & "ActiveAssets_05" Then '"c" Then
11970                                         blnFound = True
11980                                         Exit For
11990                                       End If
12000                                     Case 3&
12010                                       If .Name = strFind & "Journal_04" Then  '"d" Then
12020                                         blnFound = True
12030                                         Exit For
12040                                       End If
12050                                     Case 4&
12060                                       If .Name = strFind & "Ledger_02" Then '"e" Then
12070                                         blnFound = True
12080                                         Exit For
12090                                       End If
12100                                     End Select
12110                                   Next
12120                                 End If
12130                                 If blnFound = True Then
12140                                   lngQrys = lngQrys + 1&
12150                                   lngE = lngQrys - 1&
12160                                   ReDim Preserve arr_varQry(Q_ELEMS, lngE)
12170                                   arr_varQry(Q_QNAM, lngE) = .Name
'12180                                   arr_varQry(Q_TYP, lngE) = .Fields.Count
12180                                   arr_varQry(Q_TYP, lngE) = .Type
12190 On Error Resume Next
12200                                   strDesc = .Properties("Description")
12210 On Error GoTo 0
12220                                   If strDesc <> vbNullString Then
12230                                     arr_varQry(Q_DSC, lngE) = strDesc
12240                                   Else
12250                                     arr_varQry(Q_DSC, lngE) = Null
12260                                   End If
'12270                                   arr_varQry(Q_SQL, lngE) = .Connect
12270                                   arr_varQry(Q_SQL, lngE) = .SQL
12280                                   arr_varQry(Q_PATH, lngE) = fsfl.Path
12290                                   arr_varQry(Q_FILE, lngE) = fsfl.Name
'12300                                   Debug.Print "'TBL: " & .Name & "  FILE: " & Parse_Path(fsfl.Path) & LNK_SEP & fsfl.Name  ' ** Module Function: modFileUtilities.
12300                                   Debug.Print "'QRY: " & .Name & "  FILE: " & Parse_Path(fsfl.Path) & LNK_SEP & fsfl.Name  ' ** Module Function: modFileUtilities.
12310                                   DoEvents
12320                                 End If
12330                               Case False
12340                                 If .Name = strFind Then
12350                                   lngQrys = lngQrys + 1&
12360                                   lngE = lngQrys - 1&
12370                                   ReDim Preserve arr_varQry(Q_ELEMS, lngE)
12380                                   arr_varQry(Q_QNAM, lngE) = .Name
'12390                                   arr_varQry(Q_TYP, lngE) = .Fields.Count
12390                                   arr_varQry(Q_TYP, lngE) = .Type
12400 On Error Resume Next
12410                                   strDesc = .Properties("Description")
12420 On Error GoTo 0
12430                                   If strDesc <> vbNullString Then
12440                                     arr_varQry(Q_DSC, lngE) = strDesc
12450                                   Else
12460                                     arr_varQry(Q_DSC, lngE) = Null
12470                                   End If
'12480                                   arr_varQry(Q_SQL, lngE) = .Connect
12480                                   arr_varQry(Q_SQL, lngE) = .SQL
12490                                   arr_varQry(Q_PATH, lngE) = fsfl.Path
12500                                   arr_varQry(Q_FILE, lngE) = fsfl.Name
'12510                                   Debug.Print "'TBL: " & .Name & "  FILE: " & Parse_Path(fsfl.Path) & LNK_SEP & fsfl.Name  ' ** Module Function: modFileUtilities.
12510                                   Debug.Print "'QRY: " & .Name & "  FILE: " & Parse_Path(fsfl.Path) & LNK_SEP & fsfl.Name  ' ** Module Function: modFileUtilities.
12520                                   DoEvents
12530                                 End If
12540                               End Select  ' ** blnHasSfx.
12550                             End If  ' ** strFind.
12560                           End With  ' ** qdf.
12570                         Next  ' ** qdf.
12580                       End With  ' ** dbsLnk.
12590                       Set qdf = Nothing
12600                       dbsLnk.Close
12610                       Set dbsLnk = Nothing

12620                       lngFiles = lngFiles + 1&
12630                       lngE = lngFiles - 1&
12640                       ReDim Preserve arr_varFile(F_ELEMS, lngE)
12650                       arr_varFile(F_FNAM, lngE) = strFile
12660                       arr_varFile(F_PATH, lngE) = strPath

12670                     End If  ' ** strPathFile_This.
12680                   Case Else
                          ' ** No TrustAux.mdb, TrustDta.mdb, etc.
12690                   End Select

12700                 End If  ' ** mdb.
12710               End If
12720             End With  ' ** fsfl.
12730             DoEvents
12740             If lngFiles Mod 10 = 0 Then
12750               If lngFiles <> lngTmp02 Then
12760                 Debug.Print "'FILES: " & CStr(lngFiles)
12770                 DoEvents
12780                 lngTmp02 = lngFiles
12790               End If
12800             End If
12810           Next  ' ** fsfl.
12820           Set fsfl = Nothing
12830           Set fsfls = Nothing

12840         End If  ' ** blnFound.
12850       Next  ' ** lngX.
12860       Set fsfds = Nothing
12870       Set fsfd1 = Nothing
12880       Set fsfd2 = Nothing

12890     End With  ' ** fso.
12900     Set fso = Nothing

12910     wrk.Close
12920     Set wrk = Nothing

12930   Next  ' ** lngW.

12940   Debug.Print "'FILES CHECKED: " & CStr(lngFiles)
12950   DoEvents

'12960   Debug.Print "'TBLS: " & CStr(lngQrys)
12960   Debug.Print "'QRYS: " & CStr(lngQrys)
12970   DoEvents

12980   If lngFiles > 0& Then

12990     Set dbsLoc = CurrentDb
13000     With dbsLoc
13010       Set rst = .OpenRecordset("zz_tbl_File", dbOpenDynaset, dbConsistent)
13020       With rst
13030         blnAddAll = False: blnAdd = False
13040         If .BOF = True And .EOF = True Then
13050           blnAddAll = True
13060         End If
13070         For lngX = 0& To (lngFiles - 1&)
13080           blnAdd = False
13090           Select Case blnAddAll
                Case True
13100             blnAdd = True
13110           Case False
13120             .FindFirst "[file_name] = '" & arr_varFile(F_FNAM, lngX) & "' And [file_path] = '" & arr_varFile(F_PATH, lngX) & "'"
13130             If .NoMatch = True Then
13140               blnAdd = True
13150             End If
13160           End Select
13170           If blnAdd = True Then
13180             .AddNew
                  ' ** ![file_id] : AutoNumber.
13190             ![file_name] = arr_varFile(F_FNAM, lngX)
13200             ![file_path] = arr_varFile(F_PATH, lngX)
13210             ![file_datemodified] = Now()
13220             .Update
                End If

13230         Next
13240         .Close
13250       End With
13260       .Close
13270     End With
13280     Set dbsLoc = Nothing

13290   End If  ' ** lngFiles.

13300   If lngQrys > 0& Then

13310   End If  ' ** lngQrys.

13320   Beep
13330   Debug.Print "'DONE!"
13340   DoEvents

EXITP:
13350   Set fsfl = Nothing
13360   Set fsfls = Nothing
13370   Set fsfd1 = Nothing
13380   Set fsfd2 = Nothing
13390   Set fsfds = Nothing
13400   Set fso = Nothing
13410   Set qdf = Nothing
13420   Set tdf = Nothing
13430   Set dbsLoc = Nothing
13440   Set dbsLnk = Nothing
13450   Set wrk = Nothing
13460   FindLostQry = blnRetVal
13470   Exit Function

ERRH:
13480   blnRetVal = False
13490   Select Case ERR.Number
        Case Else
13500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13510   End Select
13520   Resume EXITP

End Function

Public Function ImportChkReg() As Boolean

100   On Error GoTo ERRH

        Const THIS_PROC As String = "ImportChkReg"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim lngRecs As Long
        Dim strPath As String, strFile As String, strPathFile As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const Q_QNAM As Integer = 0

110   On Error GoTo 0

120     blnRetVal = True

130     strPath = "C:\VictorGCS_Clients\TrustAccountant\NewWorking"
140     strFile = "Trust.mdb"
150     strPathFile = strPath & LNK_SEP & strFile

160     lngQrys = 0&
170     ReDim arr_varQry(Q_ELEMS, 0)

180     Set dbs = CurrentDb
190     Set qdf = dbs.QueryDefs("zzz_qry_zVBComponent_Query_07_03")
200     Set rst = qdf.OpenRecordset
210     With rst
220       .MoveLast
230       lngRecs = .RecordCount
240       .MoveFirst
250       For lngX = 1& To lngRecs
260         lngQrys = lngQrys + 1&
270         lngE = lngQrys - 1&
280         ReDim Preserve arr_varQry(Q_ELEMS, lngE)
290         arr_varQry(Q_QNAM, lngE) = ![qry_name]
300         If lngX < lngRecs Then .MoveNext
310       Next
320       .Close
330     End With
340     Set rst = Nothing
350     Set qdf = Nothing

360     Debug.Print "'QRYS: " & CStr(lngQrys)
370     DoEvents

380     If lngQrys > 0& Then
390       For lngX = 0& To (lngQrys - 1)
400         DoCmd.TransferDatabase acImport, "Microsoft Access", strPathFile, acQuery, arr_varQry(Q_QNAM, lngX), arr_varQry(Q_QNAM, lngX)
410         DoEvents
420       Next
430       dbs.QueryDefs.Refresh
440     End If

450     dbs.Close

460     Beep
470     Debug.Print "'DONE!"
480     DoEvents

EXITP:
800     Set rst = Nothing
810     Set qdf = Nothing
820     Set dbs = Nothing
840     ImportChkReg = blnRetVal
850     Exit Function

ERRH:
860     blnRetVal = False
870     Select Case ERR.Number
        Case Else
880       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
890     End Select
900     Resume EXITP

End Function
