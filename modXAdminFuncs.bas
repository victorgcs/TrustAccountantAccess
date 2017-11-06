Attribute VB_Name = "modXAdminFuncs"
Option Compare Database
Option Explicit

'VGC 09/21/2017: CHANGES!

' ** VBA_GetCode() in zz_mod_ModuleMiscFuncs.

' ** The column FileType, in the table MSysIMEXSpecs, can have the following values:
' **   Name                Value    Jet
' **   ==================  =====  =======
' **   Windows (ANSI)        0    Pre 4.0
' **   DOS or OS/2 (PC-8)    1    Pre 4.0
' **
' ** In Jet 4.0 and later, it is the code page to use.
' ** You can specify any code page there, not just 0 or 1.
' ** See the table in 'Character Set Recognition'.
' **   DOS (ASCII) Code Pages
' **   ======================
' **   DOS 437: United States

'Attributes: dbAutoIncrField

'MSysIMEXSpecs
'MSysIMEXColumns

' ** FileFolder graphic effect:
' **   Is there consistency in the point size of the label?
' **   Is the label, for a given point size, placed consistently?
' **   Is the tab's distance from the left edge consistent?
' **   Is the tab's rise consistent?
' **   Are the label's margins consistent?
' **   Is the mask line (hline03) placed consistently?
' **   I'm sure that box2 is a mix of heights and widths!
' **   Are there overlaps of the border lines?
' **   Is the Z-order of the pieces consistent?

' ** 10 Pt. Tab:
'frmAccountExport

' ** 8 Pt. Tab:
' **   Rise: 10 px
' **   Left:  8 px

' ** DbIndexType enumeration:
' **   0  dbIndexTypeNone    No                   No index. (Default)
' **   1  dbIndexTypeSimple  Yes (Duplicates OK)  The index allows duplicates.
' **   2  dbIndexTypeUnique  Yes (No Duplicates)  The index doesn't allow duplicates.

'Private dblPB_CurTot As Double, dblPB_CurCnt As Double
'Private dblPB_Incr As Double, dblPB_Width As Double, dblPB_CurWidth As Double
'Private dblPB_Sec1 As Double, dblPB_Sec2 As Double, dblPB_Sec3 As Double, dblPB_Sec4 As Double
'Private dblPB_Sec5 As Double, dblPB_Sec6 As Double, dblPB_Sec7 As Double

Private Const THIS_NAME As String = "modXAdminFuncs"
' **

Public Function GetMenu() As Boolean
' ** Only called if MDE.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "GetMenu"

        Dim varCurForm As Variant, varCurReport As Variant, varCurControl As Variant, varLastControl As Variant
        Dim lngFrms As Long, lngRpts As Long, lngErrNum As Long
        Dim strDocName As String
        Dim blnDoMenu As Boolean, blnWasInvisible As Boolean
        Dim lngX As Long
        Dim blnRetVal As Boolean

110     blnRetVal = True

        ' ** In we're here because of an error.
        ' ** If there's a hidden error box, I don't know of a way to find it.
120     lngErrNum = ERR.Number

130     lngFrms = Forms.Count
140     lngRpts = Reports.Count
150     gblnGoToReport = False

160     varCurForm = Null: varLastControl = Null: varCurReport = Null
170     strDocName = vbNullString
180     blnDoMenu = False: blnWasInvisible = False

190   On Error Resume Next
        ' ** 2467 : The expression you entered refers to an object that is closed or doesn't exist.
200     varLastControl = Screen.PreviousControl.Name
        ' ** 2474 : The expression you entered requires the control to be in the active window.
210     varCurControl = Screen.ActiveControl.Name
        ' ** 2475 : You entered an expression that requires a form be the active window.
220     varCurForm = Screen.ActiveForm.Name
        ' ** 2476 : You entered an expression that requires a report to be the active window.
230     varCurReport = Screen.ActiveReport.Name
240   On Error GoTo ERRH

250     If lngFrms = 0& Then
260       blnDoMenu = True
270     ElseIf lngFrms = 1& Then
280       If Forms(0).Name = "frmMenu_Background" Then
290         blnDoMenu = True
300       Else
310         If IsNull(varCurForm) = True And IsNull(varCurControl) = True Then
320           blnDoMenu = True
330         Else
340           If Forms(0).Visible = False Then
350             Forms(0).Visible = True
360             blnWasInvisible = True
370           Else
380             gblnSetFocus = True
390             DoCmd.SelectObject acForm, Forms(0).Name, False
400             Forms(0).TimerInterval = 100&
410           End If
420         End If
430       End If
440     Else
450       For lngX = 0& To (lngFrms - 1&)
460         If Forms(lngX).Name = "frmMenu_Background" Then
470           blnDoMenu = True
480           Exit For
490         End If
500       Next
510       If blnDoMenu = False Then
520         For lngX = 0& To (lngFrms - 1&)
530           If Forms(lngX).Visible = False Then
540             Forms(lngX).Visible = True
550             blnWasInvisible = True
560           End If
570         Next
580         If blnWasInvisible = False Then
590           gblnSetFocus = True
600           DoCmd.SelectObject acForm, Forms(lngFrms - 1&).Name, False
610           Forms(lngFrms - 1&).TimerInterval = 100&
620         End If
630       End If
640     End If

650     If blnDoMenu = True Then

660       If lngRpts > 0& Then
            ' ** Close any reports still open.
670         Do While Reports.Count > 0
680           DoCmd.Close acReport, Reports(0).Name
690           DoEvents
700         Loop
710       End If

720       If lngFrms > 0& Then
            ' ** Close any remaining open forms.
730         Do While Forms.Count > 0
740           DoCmd.Close acForm, Forms(0).Name
750           DoEvents
760         Loop
770       End If

          ' ** Changed my mind. Everything gets frmMenu_Title!
780       gstrTrustDataLocation = vbNullString  ' ** This'll trigger the whole startup routine.
790       strDocName = "frmMenu_title"
800       DoCmd.OpenForm strDocName

810     End If

EXITP:
820     GetMenu = blnRetVal
830     Exit Function

ERRH:
840     Select Case ERR.Number
        Case Else
850       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
860     End Select
870     Resume EXITP

End Function

Public Function ConType_Get(varConnect As Variant) As Long

900   On Error GoTo ERRH

        Const THIS_PROC As String = "ConType_Get"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strConnect As String, strSpec As String
        Dim intPos01 As Integer
        Dim lngX As Long
        Dim lngRetVal As Long

        Static lngConTypes As Long
        Static arr_varConType() As Variant

        ' ** Array: arr_varConType().
        'Const C_CID   As Integer = 0
        Const C_TYP   As Integer = 1
        'Const C_CONST As Integer = 2
        Const C_SPEC  As Integer = 3
        'Const C_NAM   As Integer = 4
        'Const C_PARM  As Integer = 5

910     lngRetVal = dbCNothing

920     If IsNull(varConnect) = False Then
930       If Trim(varConnect) <> vbNullString Then

940         strConnect = Trim(varConnect)

950         If lngConTypes = 0& Or IsEmpty(arr_varConType) = True Then
960           Set dbs = CurrentDb
970           With dbs
980             Set qdf = .QueryDefs("qryConnectionType_02")
990             Set rst = qdf.OpenRecordset
1000            With rst
1010              If .BOF = True And .EOF = True Then
                    ' ** Shouldn't happen.
1020              Else
1030                .MoveLast
1040                lngConTypes = .RecordCount
1050                .MoveFirst
1060                arr_varConType = .GetRows(lngConTypes)
                    ' *******************************************************
                    ' ** Array: arr_varConType()
                    ' **
                    ' **   Field  Element  Name                  Constant
                    ' **   =====  =======  ====================  ==========
                    ' **     1       0     contype_id            C_CID
                    ' **     2       1     contype_type          C_TYP
                    ' **     3       2     contype_constant      C_CONST
                    ' **     4       3     contype_specifier     C_SPEC
                    ' **     5       4     contype_name          C_NAM
                    ' **     6       5     contype_parameters    C_PARM
                    ' **
                    ' *******************************************************
1070              End If
1080              .Close
1090            End With
1100            .Close
1110          End With

1120        End If

1130        intPos01 = InStr(strConnect, ";")
            ' ** ;DATABASE=C:\VictorGCS_Clients\TrustAccountant\Clients\Ohana\TrustDta.mdb
1140        If intPos01 = 1 Then
1150          For lngX = 0& To (lngConTypes - 1&)
1160            If arr_varConType(C_SPEC, lngX) = "[database];" Then
1170              lngRetVal = arr_varConType(C_TYP, lngX)
1180              Exit For
1190            End If
1200          Next
1210        ElseIf intPos01 > 1 Then
1220          strSpec = Left(strConnect, intPos01)
1230          For lngX = 0& To (lngConTypes - 1&)
1240            If arr_varConType(C_SPEC, lngX) = strSpec Then
1250              lngRetVal = arr_varConType(C_TYP, lngX)
1260              Exit For
1270            End If
1280          Next
1290          If lngRetVal = dbCNothing Then
1300            Beep
                'Debug.Print "'" & strConnect
1310          End If
1320        Else
              ' ** No idea!
1330        End If
1340      Else
1350        lngRetVal = dbCAccess
1360      End If
1370    End If

EXITP:
1380    Set rst = Nothing
1390    Set qdf = Nothing
1400    Set dbs = Nothing
1410    ConType_Get = lngRetVal
1420    Exit Function

ERRH:
1430    lngRetVal = dbCNothing
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Function

Public Function zz_IMEX_Create() As Boolean
' ** May not be necessary now that I know how to create Data-Definition Language queries.

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "zz_IMEX_Create"

        Dim intID As Integer
        Dim strSQL As String
        Dim arr_varCol As Variant
        Dim strSpec_Name As String
        Dim intX As Integer
        Dim blnRetVal As Boolean

1510    blnRetVal = True

        ' ** create export spec if needed.
1520    strSpec_Name = "xx"
1530    intID = 0  'SpecID(strSpec_Name)
1540    If intID = 0 Then

1550      strSQL = _
            "INSERT INTO MSysIMEXSpecs " & _
            "(DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint," & _
            "FieldSeparator,FileType,SpecName,SpecType,StartRow,TextDelim,TimeDelim)" & _
            "VALUES ('/',Yes,No,2,'.',';',1252,'" & strSpec_Name & "',1,0,'',':');"
1560      DoCmd.RunSQL strSQL

1570      intID = 0  'SpecID(spec_name)

1580      arr_varCol = Array( _
            "(0,10,'FLD',0,0," & intID & ",151,100)", _
            "(0,10,'KYS',0,0," & intID & ",51,100)", _
            "(0,10,'TBL',0,0," & intID & ",1,50)", _
            "(0,10,'VAL',0,0," & intID & ",251,32000)")

1590      For intX = 0 To UBound(arr_varCol)
1600        strSQL = _
              "INSERT INTO MsysIMEXColumns " & _
              "VALUES " & arr_varCol(intX) & ";"
1610        DoCmd.RunSQL strSQL
1620      Next

1630    End If

EXITP:
1640    zz_IMEX_Create = blnRetVal
1650    Exit Function

ERRH:
1660    blnRetVal = False
1670    Select Case ERR.Number
        Case Else
1680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1690    End Select
1700    Resume EXITP

End Function

Public Function GetAccessVersion(strFileName As String) As String
'#####################################
'NEEDS UPDATING THROUGH ACCESS 2010!
'#####################################

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "GetAccessVersion"

        Dim intFileNum As Integer
        Dim strTmp01 As String
        Dim lngX As Long
        Dim strRetVal As String

1810    strRetVal = vbNullString

1820    intFileNum = FreeFile()
1830    Open strFileName For Binary Access Read Shared As #intFileNum
1840    strTmp01 = Space(LOF(intFileNum))
1850  On Error Resume Next
1860    Get #intFileNum, 1, strTmp01
1870    If ERR.Number <> 0 Then
1880  On Error GoTo ERRH
1890      strRetVal = RET_ERR
1900    Else
1910  On Error GoTo ERRH
1920    End If
1930    Close #intFileNum

1940    If strRetVal = vbNullString Then
1950      lngX = Asc(Mid(strTmp01, 115, 1))
1960      Select Case CStr(Hex(lngX))
          Case "12"
1970        strRetVal = "A97 mdb/mde"
1980      Case "7f"
1990        strRetVal = "A2K mdb"
2000      Case "15"
2010        strRetVal = "A2K2/3 mdb"
2020      Case "69"
2030        strRetVal = "A2K7 accdb"
2040      Case "ff"
2050        strRetVal = "A2K7 adp"
2060      Case Else
2070        strRetVal = "Unknown"
2080      End Select
2090      strRetVal = Left(strRetVal & Space(11), 11) & " : " & CStr(Hex(lngX))
2100    End If

EXITP:
2110    GetAccessVersion = strRetVal
2120    Exit Function

ERRH:
2130    strRetVal = RET_ERR
2140    Select Case ERR.Number
        Case Else
2150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2160    End Select
2170    Resume EXITP

End Function

Public Function GetCreatedVersion(strFileName As String) As Integer
'#####################################
'NEEDS UPDATING THROUGH ACCESS 2010!
'#####################################

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetCreatedVersion"

        Dim dbs As DAO.Database, prp As DAO.Property
        Dim strDir As String
        Dim intProjVer As Integer
        Dim intF As Integer
        Dim strKeyWord As String
        Dim intRetVal As Integer

2210    intRetVal = 0

2220    strDir = Dir(strFileName)

2230    If Len(strDir) <> 0 Then

2240      Set dbs = DBEngine(0).OpenDatabase(strFileName)

2250      Select Case Left(dbs.Properties("AccessVersion"), 2)
          Case "02"
2260        intRetVal = 2
2270      Case "06"
2280        intRetVal = 7
2290      Case "07"
2300        intRetVal = 8
2310      Case "08"
2320        For Each prp In dbs.Properties
2330          If prp.Name = "ProjVer" Then
2340            intProjVer = prp.Value
2350            Exit For
2360          End If
2370        Next
2380        Select Case intProjVer
            Case 0
2390          intRetVal = 9
2400        Case 24
2410          intRetVal = 10
2420        Case 35
2430          intRetVal = 11
2440        Case Else
2450          intRetVal = -1
2460        End Select
2470      Case "09"
2480        For Each prp In dbs.Properties
2490          If prp.Name = "ProjVer" Then
2500            intProjVer = prp.Value
2510            Exit For
2520          End If
2530        Next
2540        Select Case intProjVer
            Case 0
2550          intRetVal = 10
2560        Case 24
2570          intRetVal = 10
2580        Case 35
2590          intRetVal = 11
2600        Case Else
2610          intRetVal = -1
2620        End Select
2630      Case Else
2640        intRetVal = 0
2650      End Select
2660      dbs.Close

2670    End If

EXITP:
2680    Set prp = Nothing
2690    Set dbs = Nothing
2700    GetCreatedVersion = intRetVal
2710    Exit Function

ERRH:
2720    If Not dbs Is Nothing Then
2730      dbs.Close
2740      Set dbs = Nothing
2750    End If
2760    Select Case ERR.Number
        Case 3045  ' ** Couldn't use '|'; file already in use.
2770      Beep
2780      MsgBox ERR.description, vbInformation + vbOKOnly, "File In Use"
2790      intRetVal = 0
2800    Case 3343  ' ** Unrecognized database format '|'.
2810      intF = FreeFile
2820      Open strFileName For Input Access Read As intF
2830      If LOF(intF) < 20 Then
2840        Beep
2850        MsgBox ERR.description, vbInformation + vbOKOnly, "Unrecognized File"
2860        intRetVal = 0
2870      Else
2880        Seek intF, 5
2890        strKeyWord = Input(15, intF)
2900        If strKeyWord = "Standard Jet db" Then
2910          intRetVal = -1
2920        Else
2930          Beep
2940          MsgBox ERR.description, vbInformation + vbOKOnly, "Unrecognized File"
2950          intRetVal = 0
2960        End If
2970      End If
2980      Close intF
2990    Case Else
3000      intRetVal = 0
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3020    End Select
3030    Resume EXITP

End Function

Public Function GetFileSize(varInput As Variant, strMode As String, Optional varUnit As Variant) As String
'KEEP!!
3100  On Error GoTo ERRH

        Const THIS_PROC As String = "GetFileSize"

        Dim strUnit As String
        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, dblTmp02 As Double
        Dim strRetVal As String

3110    strRetVal = vbNullString

3120    If IsNull(varInput) = False Then

3130      Select Case IsMissing(varUnit)
          Case True
3140        strUnit = "calc"
3150      Case False
3160        strUnit = varUnit
3170      End Select

3180      Select Case strMode
          Case "Value"
            ' ** A raw size value, in Bytes, to be converted to Bt, Kb, Mb, Gb.
3190        dblTmp02 = varInput
3200      Case "PathFile"
            ' ** A path and filename, whose size is to be returned.
3210        dblTmp02 = FileLen(varInput)
3220      End Select

3230      Select Case strUnit
          Case "bt"
3240        strRetVal = Format(dblTmp02, "#,##0.##")
3250        If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3260        strRetVal = strRetVal & " Bt"
3270      Case "kb"
3280        dblTmp02 = (dblTmp02 / 1024#)
3290        strRetVal = Format(dblTmp02, "#,##0.##")
3300        If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3310        strRetVal = strRetVal & " Kb"
3320      Case "mb"
3330        dblTmp02 = ((dblTmp02 / 1024#) / 1024#)
3340        strRetVal = Format(dblTmp02, "#,##0.##")
3350        If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3360        strRetVal = strRetVal & " Mb"
3370      Case "gb"
3380        dblTmp02 = (((dblTmp02 / 1024#) / 1024#) / 1024#)
3390        strRetVal = Format(dblTmp02, "#,##0.##")
3400        If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3410        strRetVal = strRetVal & " Gb"
3420      Case "calc"
3430        strTmp01 = CStr(dblTmp02)
3440        intPos01 = InStr(strTmp01, ".")
3450        If intPos01 = 0 Then intLen = Len(strTmp01) Else intLen = Len(Left(strTmp01, (intPos01 - 1)))
3460        If intLen <= 3 Then
              ' ** Return bytes.
3470          strRetVal = Format(dblTmp02, "#,##0.##")
3480          If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3490          strRetVal = strRetVal & " Bt"
3500        ElseIf intLen <= 6 Then
              ' ** Return kilobytes.
3510          dblTmp02 = (dblTmp02 / 1024#)
3520          strRetVal = Format(dblTmp02, "#,##0.##")
3530          If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3540          strRetVal = strRetVal & " Kb"
3550        ElseIf intLen <= 9 Then
              ' ** Return megabytes.
3560          dblTmp02 = ((dblTmp02 / 1024#) / 1024#)
3570          strRetVal = Format(dblTmp02, "#,##0.##")
3580          If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3590          strRetVal = strRetVal & " Mb"
3600        Else
              ' ** Return gigabytes.
3610          dblTmp02 = (((dblTmp02 / 1024#) / 1024#) / 1024#)
3620          strRetVal = Format(dblTmp02, "#,##0.##")
3630          If Right(Trim(strRetVal), 1) = "." Then strRetVal = Left(Trim(strRetVal), (Len(strRetVal) - 1))
3640          strRetVal = strRetVal & " Gb"
3650        End If
3660      End Select

3670    End If  ' ** IsNull().

EXITP:
3680    GetFileSize = strRetVal
3690    Exit Function

ERRH:
3700    strRetVal = RET_ERR
3710    Select Case ERR.Number
        Case Else
3720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3730    End Select
3740    Resume EXITP

End Function

Public Function Tbl_Del_Rel(Optional varTbl As Variant) As Boolean

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_Del_Rel"

        Dim dbs As DAO.Database, tdf As DAO.TableDef
        Dim strFind As String, strMsg As String
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim msgResponse As VbMsgBoxResult
        Dim intLen As Integer
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varTbl().
        Const T_ELEMS As Integer = 1  ' ** Array's first element UBound().
        Const T_TNAM As Integer = 0
        Const T_IDX  As Integer = 1

3810    blnRetVal = True

3820    If IsMissing(varTbl) = True Then
3830      strFind = InputBox("Enter Table Name Partial:", "Delete Table", vbNullString)
3840    Else
3850      strFind = varTbl
3860    End If

3870    If strFind <> vbNullString Then

3880      DoCmd.Hourglass True
3890      DoEvents

3900      Set dbs = CurrentDb
3910      With dbs

3920        lngX = 0&
3930        lngTbls = 0&
3940        ReDim arr_varTbl(T_ELEMS, 0)

3950        intLen = Len(strFind)

3960        For Each tdf In .TableDefs
3970          With tdf
3980            lngX = lngX + 1&
3990            If Left(.Name, intLen) = strFind Then
4000              lngTbls = lngTbls + 1&
4010              lngE = lngTbls - 1&
4020              ReDim Preserve arr_varTbl(T_ELEMS, lngE)
4030              arr_varTbl(T_TNAM, lngE) = .Name
4040              arr_varTbl(T_IDX, lngE) = lngX
4050            End If
4060          End With  ' ** tdf.
4070        Next  ' ** tdf.
4080        Set tdf = Nothing

4090        .Close
4100      End With  ' ** dbs.
4110      Set dbs = Nothing

4120      If lngTbls > 0& Then
4130        strMsg = CStr(lngTbls) & " table" & IIf(lngTbls = 1&, " was ", "s were ") & "found beginning with string:" & vbCrLf & vbCrLf & _
              "    " & strFind & vbCrLf & vbCrLf & IIf(lngTbls = 1&, "Delete it?", "Delete them?")
4140        DoCmd.Hourglass False
4150        msgResponse = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, "Tables Found")
4160        If msgResponse = vbYes Then
4170          DoCmd.Hourglass True
4180          DoEvents
4190          For lngX = 0& To (lngTbls - 1&)
4200            DoCmd.DeleteObject acTable, arr_varTbl(T_TNAM, lngX)
4210            DoEvents
4220          Next ' ** lngX.
4230          CurrentDb.TableDefs.Refresh
4240          DoEvents
4250          Beep
4260          DoCmd.Hourglass False
4270          MsgBox "Finished", vbExclamation + vbOKOnly, ("Finished" & Space(40))
4280        End If  ' ** msgResponse.
4290      Else
4300        Beep
4310        DoCmd.Hourglass False
4320        MsgBox "No tables were found beginning with string:" & vbCrLf & vbCrLf & _
              "    " & strFind, vbInformation + vbOKOnly, "Table Not Found"
4330      End If  ' ** lngTbls.

4340    End If  ' ** strFind.

EXITP:
4350    Set tdf = Nothing
4360    Set dbs = Nothing
4370    Tbl_Del_Rel = blnRetVal
4380    Exit Function

ERRH:
4390    DoCmd.Hourglass False
4400    blnRetVal = False
4410    Select Case ERR.Number
        Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Function

Public Function Tbl_Fld_Find(Optional varInput As Variant) As Variant

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_Fld_Find"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim strFind As String
        Dim lngThisDbsID As Long, strThisDbsName As String
        Dim lngTmpJrnlID As Long, blnTmpJrnl As Boolean
        Dim blnCalled As Boolean
        Dim varTmp00 As Variant, arr_varTmp01() As Variant
        Dim lngX As Long, lngE As Long
        Dim varRetVal As Variant

        ' ** Array: arr_varTbl().
        Const T_ELEMS As Integer = 5  ' ** Array's first-element UBound.
        Const T_DID  As Integer = 0
        Const T_DNAM As Integer = 1
        Const T_TID  As Integer = 2
        Const T_TNAM As Integer = 3
        Const T_FID  As Integer = 4
        Const T_FNAM As Integer = 5

4510    varRetVal = Null

4520    Select Case IsMissing(varInput)
        Case True
4530      strFind = "cusip"
4540      blnCalled = False
4550      blnTmpJrnl = False
4560    Case False
4570      Select Case IsNull(varInput)
          Case True
4580        strFind = vbNullString
4590      Case False
4600        If Trim(varInput) <> vbNullString Then
4610          strFind = Trim(varInput)
4620        Else
4630          strFind = vbNullString
4640        End If
4650      End Select
4660      blnCalled = True
4670      blnTmpJrnl = False
4680    End Select

4690    If strFind <> vbNullString Then

4700      lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
4710      strThisDbsName = CurrentAppName  ' ** Module Function: modFileUtilities.
4720      varTmp00 = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & gstrFile_RePostDataName & "'")
4730      lngTmpJrnlID = varTmp00

4740      lngTbls = 0&
4750      ReDim arr_varTbl(T_ELEMS, 0)

4760      Set dbs = CurrentDb
4770      With dbs
4780        Set rst = .OpenRecordset("tblDatabase_Table_Field", dbOpenDynaset, dbReadOnly)
4790        With rst
4800          .FindFirst "[fld_name] = '" & strFind & "'"
4810          Select Case .NoMatch
              Case True
4820            Select Case blnCalled
                Case True
4830              ReDim arr_varTmp01(0, 0)
4840              arr_varTmp01(0, 0) = Null
4850              varRetVal = arr_varTmp01
4860            Case False
4870              Debug.Print "'NONE FOUND!"
4880              DoEvents
4890            End Select
4900          Case False
4910            Do While .NoMatch = False
4920              If blnTmpJrnl = False And ![dbs_id] = lngTmpJrnlID Then
                    ' ** Skip these!
4930              Else
4940                lngTbls = lngTbls + 1&
4950                lngE = lngTbls - 1&
4960                ReDim Preserve arr_varTbl(T_ELEMS, lngE)
4970                arr_varTbl(T_DID, lngE) = ![dbs_id]
4980                arr_varTbl(T_DNAM, lngE) = Null
4990                arr_varTbl(T_TID, lngE) = ![tbl_id]
5000                arr_varTbl(T_TNAM, lngE) = Null
5010                arr_varTbl(T_FID, lngE) = ![fld_id]
5020                arr_varTbl(T_FNAM, lngE) = ![fld_name]
5030              End If  ' ** blnTmpJrnl.
5040              .FindNext "[fld_name] = '" & strFind & "'"
5050            Loop  ' ** NoMatch.
5060          End Select
5070        End With  ' ** rst.
5080        Set rst = Nothing
5090        .Close
5100      End With  ' ** dbs.
5110      Set dbs = Nothing

5120      If lngTbls > 0& Then

5130        Debug.Print "'HITS: " & CStr(lngTbls)
5140        DoEvents

5150        For lngX = 0& To (lngTbls - 1&)
5160          If arr_varTbl(T_DID, lngX) = lngThisDbsID Then
5170            arr_varTbl(T_DNAM, lngX) = strThisDbsName
5180          Else
5190            varTmp00 = DLookup("[dbs_name]", "tblDatabase", "[dbs_id] = " & CStr(arr_varTbl(T_DID, lngX)))
5200            arr_varTbl(T_DNAM, lngX) = varTmp00
5210          End If
5220          varTmp00 = DLookup("[tbl_name]", "tblDatabase_Table", "[dbs_id] = " & CStr(arr_varTbl(T_DID, lngX)) & " And " & _
                "[tbl_id] = " & CStr(arr_varTbl(T_TID, lngX)))
5230          arr_varTbl(T_TNAM, lngX) = varTmp00
5240        Next  ' ** lngX.

5250        Select Case blnCalled
            Case True
5260          varRetVal = arr_varTbl
5270        Case False
5280          For lngX = 0& To (lngTbls - 1&)
5290            Debug.Print "'FLD: " & arr_varTbl(T_FNAM, lngX) & "  TBL: " & arr_varTbl(T_TNAM, lngX) & "  DBS: " & arr_varTbl(T_DNAM, lngX)
5300            DoEvents
5310          Next  ' ** lngX.
5320          Beep
5330          Debug.Print "'DONE!"
5340          DoEvents
5350        End Select

5360      End If

5370    End If

EXITP:
5380    Set rst = Nothing
5390    Set dbs = Nothing
5400    Tbl_Fld_Find = varRetVal
5410    Exit Function

ERRH:
5420    varRetVal = RET_ERR
5430    Select Case ERR.Number
        Case Else
5440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5450    End Select
5460    Resume EXITP

End Function

Public Function Tbl_Fld_List(Optional arr_varRetArray As Variant, Optional varRetFind As Variant) As Boolean
'KEEP!!

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_Fld_List"

        Dim dbs As DAO.Database, obj As Object, fld As DAO.Field
        Dim strFind1 As String
        Dim blnTable As Boolean, blnAlpha As Boolean, blnRequired As Boolean
        Dim lngObjs As Long, arr_varObj() As Variant
        Dim blnRetArray As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varObj().
        Const O_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const O_NAM As Integer = 0
        Const O_TYP As Integer = 1
        Const O_REQ As Integer = 2

5510    blnRetVal = True

5520    Select Case IsMissing(arr_varRetArray)
        Case True
5530      blnRetArray = False
5540      strFind1 = "zz_tbl_VBComponent_Query"
5550      blnAlpha = False  ' ** True: Alphabetize; False: Field Order.
5560      blnRequired = False  ' ** True: Show Required status; False: Don't show required status.
5570    Case False
5580      blnRetArray = True
5590      strFind1 = varRetFind
5600      blnAlpha = False
5610      blnRequired = False
5620    End Select

5630    If Left(strFind1, 3) = "qry" Or Left(strFind1, 6) = "zz_qry" Or Left(strFind1, 7) = "zzz_qry" Then
5640      blnTable = False
5650    Else
5660      blnTable = True
5670    End If
5680    If blnAlpha = True Then
5690      lngObjs = 0&
5700      ReDim arr_varObj(O_ELEMS, 0)
5710    End If

5720    Set dbs = CurrentDb
5730    With dbs
5740      Select Case blnTable
          Case True
5750        Set obj = .TableDefs(strFind1)
5760      Case False
5770        Set obj = .QueryDefs(strFind1)
5780      End Select
5790      If blnAlpha = False And blnRetArray = False Then
5800        Debug.Print "'" & IIf(blnTable = True, "TBL: ", "QRY: ") & strFind1 & "  FLDS: " & CStr(obj.Fields.Count)
5810      End If
5820      With obj
5830        For Each fld In .Fields
5840          With fld
5850            Select Case blnRetArray
                Case True
5860              lngObjs = lngObjs + 1&
5870              lngE = lngObjs - 1&
5880              ReDim Preserve arr_varRetArray(O_ELEMS, lngE)
5890              arr_varRetArray(O_NAM, lngE) = .Name
5900              arr_varRetArray(O_TYP, lngE) = .Type
5910              Select Case blnRequired
                  Case True
5920                arr_varRetArray(O_REQ, lngE) = .Required  ' ** Meaningless for queries.
5930              Case False
5940                arr_varRetArray(O_REQ, lngE) = CBool(False)
5950              End Select
5960            Case False
5970              Select Case blnAlpha
                  Case True
5980                lngObjs = lngObjs + 1&
5990                lngE = lngObjs - 1&
6000                ReDim Preserve arr_varObj(O_ELEMS, lngE)
6010                arr_varObj(O_NAM, lngE) = .Name
6020                arr_varObj(O_TYP, lngE) = .Type
6030                arr_varObj(O_REQ, lngE) = .Required  ' ** Meaningless for queries.
6040              Case False
6050                Select Case blnRequired
                    Case True
6060                  Debug.Print "'" & .Name & "    " & .Required
6070                Case False
6080                  Debug.Print "'" & .Name
6090                End Select
6100              End Select
6110            End Select
6120          End With
6130        Next
6140      End With
6150      .Close
6160    End With

6170    If blnAlpha = True Then
          ' ** Binary sort arr_varObj) array, by name. (I know, it's not a binary sort!)
6180      For lngX = UBound(arr_varObj, 2) To 1 Step -1
6190        For lngY = 0 To (lngX - 1)
6200          If arr_varObj(O_NAM, lngY) > arr_varObj(O_NAM, (lngY + 1)) Then
6210            For lngZ = 0& To O_ELEMS
6220              varTmp00 = arr_varObj(lngZ, lngY)
6230              arr_varObj(lngZ, lngY) = arr_varObj(lngZ, (lngY + 1))
6240              arr_varObj(lngZ, (lngY + 1)) = varTmp00
6250              varTmp00 = Empty
6260            Next
6270          End If
6280        Next
6290      Next
6300      Debug.Print "'" & IIf(blnTable = True, "TBL: ", "QRY: ") & strFind1 & "  FLDS: " & CStr(lngObjs)
6310      For lngX = 0& To (lngObjs - 1&)
6320        Debug.Print "'" & arr_varObj(O_NAM, lngX)
6330      Next
6340    End If

6350    If blnRetArray = False Then
6360      Debug.Print "'DONE!  " & THIS_PROC & "()"
6370      Beep
6380    End If

EXITP:
6390    Set fld = Nothing
6400    Set obj = Nothing
6410    Set dbs = Nothing
6420    Tbl_Fld_List = blnRetVal
6430    Exit Function

ERRH:
6440    blnRetVal = False
6450    Select Case ERR.Number
        Case Else
6460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6470    End Select
6480    Resume EXITP

End Function

Public Function Tbl_Fld_Type(varType As Variant) As Variant

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_Fld_Type"

        Dim varRetVal As Variant

6510    varRetVal = Null

6520    If IsNull(varType) = False Then
6530      If IsNumeric(varType) = True Then
6540        Select Case varType
            Case dbBoolean
6550          varRetVal = "dbBoolean"
6560        Case dbByte
6570          varRetVal = "dbByte"
6580        Case dbInteger
6590          varRetVal = "dbInteger"
6600        Case dbLong
6610          varRetVal = "dbLong"
6620        Case dbCurrency
6630          varRetVal = "dbCurrency"
6640        Case dbSingle
6650          varRetVal = "dbSingle"
6660        Case dbDouble
6670          varRetVal = "dbDouble"
6680        Case dbDate
6690          varRetVal = "dbDate"
6700        Case dbBinary
6710          varRetVal = "dbBinary"
6720        Case dbText
6730          varRetVal = "dbText"
6740        Case dbLongBinary
6750          varRetVal = "dbLongBinary"
6760        Case dbMemo
6770          varRetVal = "dbMemo"
6780        Case dbGUID
6790          varRetVal = "dbGUID"
6800        Case dbBigInt
6810          varRetVal = "dbBigInt"
6820        Case dbVarBinary
6830          varRetVal = "dbVarBinary"
6840        Case dbChar
6850          varRetVal = "dbChar"
6860        Case dbNumeric
6870          varRetVal = "dbNumeric"
6880        Case dbDecimal
6890          varRetVal = "dbDecimal"
6900        Case dbFloat
6910          varRetVal = "dbFloat"
6920        Case dbTime
6930          varRetVal = "dbTime"
6940        Case dbTimeStamp
6950          varRetVal = "dbTimeStamp"
6960        Case dbUnknown
6970          varRetVal = "dbUnknown"
6980        End Select
6990      End If
7000    End If

EXITP:
7010    Tbl_Fld_Type = varRetVal
7020    Exit Function

ERRH:
7030    varRetVal = RET_ERR
7040    Select Case ERR.Number
        Case Else
7050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7060    End Select
7070    Resume EXITP

End Function

Public Function Rel_Missing(Optional varNonContig As Variant, Optional varRegen As Variant, Optional varNoDebug As Variant) As Boolean
' ** When the frontend is recreated, it sometimes copies the local database-spanning relationships (non-contiguous),
' ** and sometimes it doesn't. This will check, and regenerate them if missing.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "Rel_Missing"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, Rel As DAO.Relation
        Dim lngRels As Long, arr_varRel As Variant
        Dim lngXRels As Long, arr_varXRel() As Variant
        Dim lngNewRels As Long
        Dim blnNonContig As Boolean, blnRegen As Boolean, blnExtraMulti As Boolean, blnListAll As Boolean, blnListNonContig As Boolean
        Dim blnFound As Boolean, blnFound2 As Boolean, blnContinue As Boolean, blnHasMissing As Boolean, blnCalled As Boolean
        Dim lngErrs As Long, arr_varErr() As Variant
        Dim lngThisDbsID As Long
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varRel().
        'Const R_RID    As Integer = 0
        Const R_NAM    As Integer = 1
        Const R_DID1   As Integer = 2
        Const R_DBNAM1 As Integer = 3
        Const R_TNAM1  As Integer = 5
        Const R_DID2   As Integer = 6
        Const R_DBNAM2 As Integer = 7
        Const R_TNAM2  As Integer = 9
        Const R_FLDS   As Integer = 10
        Const R_ATTR   As Integer = 11
        Const R_FND    As Integer = 12
        Const R_FNAM1  As Integer = 14
        Const R_FNAM1F As Integer = 15
        Const R_DTYP1  As Integer = 17
        Const R_FNAM2  As Integer = 19
        Const R_FNAM2F As Integer = 20
        Const R_DTYP2  As Integer = 22
        Const R_FNAM3  As Integer = 24
        Const R_FNAM3F As Integer = 25
        Const R_DTYP3  As Integer = 27
        Const R_FNAM4  As Integer = 29
        Const R_FNAM4F As Integer = 30
        Const R_DTYP4  As Integer = 32
        Const R_FNAM5  As Integer = 34
        Const R_FNAM5F As Integer = 35
        Const R_DTYP5  As Integer = 37
        Const R_FNAM6  As Integer = 39
        Const R_FNAM6F As Integer = 40
        Const R_DTYP6  As Integer = 42
        Const R_FNAM7  As Integer = 44
        Const R_FNAM7F As Integer = 45
        Const R_DTYP7  As Integer = 47

        ' ** Array: arr_varXRel().
        Const XR_ELEMS As Integer = 6  ' ** Array's first-element UBound().
        Const XR_NAM        As Integer = 0
        Const XR_FND        As Integer = 1
        Const XR_TNAM1      As Integer = 2
        Const XR_TNAM2      As Integer = 3
        Const XR_FLDS       As Integer = 4
        Const XR_ATTR       As Integer = 5
        Const XR_R_ARR_ELEM As Integer = 6

        ' ** Array: arr_varErr().
        Const E_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const E_NUM  As Integer = 0
        Const E_DSC  As Integer = 1
        Const E_RNAM As Integer = 2

7110    blnRetVal = True
7120    blnHasMissing = False
7130    lngNewRels = 0&

7140    lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

7150    Select Case IsMissing(varNonContig)
        Case True
7160      blnNonContig = True      ' ** True: Look for missing non-contiguous relationships; False: Look for missing contiguous relationships
7170      blnRegen = True          ' ** True: Regenerate missing relationships.
7180      blnListAll = False       ' ** True: List all relationships.
7190      blnListNonContig = False ' ** True: List non-contiguouse relationships.
7200      blnExtraMulti = False    ' ** True: List extra relationships.
7210      blnCalled = False
7220    Case False
7230      blnNonContig = CBool(varNonContig)
7240      Select Case IsMissing(varRegen)
          Case True
7250        blnRegen = False
7260      Case False
7270        blnRegen = CBool(varRegen)
7280      End Select
7290      blnListAll = False
7300      blnListNonContig = False
7310      blnExtraMulti = False
7320      blnCalled = True
7330    End Select

7340    lngXRels = 0&
7350    ReDim arr_varXRel(XR_ELEMS, 0)

7360    Set dbs = CurrentDb
7370    With dbs

7380      Select Case blnNonContig
          Case True

            ' ** tblXAdmin_Relation (just non-contiguous relationships), linked to tblXAdmin_Relation_Field, _
            ' ** qryRelation_01 (tblDatabase_Table_Field, with add'l fields), with add'l fields.
7390        Set qdf = .QueryDefs("qryRelation_02")
7400        Set rst = qdf.OpenRecordset
7410        With rst
7420          If .BOF = True And .EOF = True Then
                ' ** You shouldn't even be here!
7430            lngRels = 0&
7440            If blnCalled = False Then
7450              Debug.Print "'NO NON-CONTIGUOUS RELATIONSHIPS FOUND!"
7460            End If
7470          Else
7480            .MoveLast
7490            lngRels = .RecordCount
7500            .MoveFirst
7510            arr_varRel = .GetRows(lngRels)
                ' ********************************************************
                ' ** Array: arr_varRel()
                ' **
                ' **   Field  Element  Name                   Constant
                ' **   =====  =======  =====================  ==========
                ' **     1       0     rel_id
                ' **     2       1     rel_name               R_NAM
                ' **     3       2     dbs_id1
                ' **     4       3     dbs_name1              R_DBNAM1
                ' **     5       4     tbl_id1
                ' **     6       5     tbl_name1              R_TNAM1
                ' **     7       6     dbs_id2
                ' **     8       7     dbs_name2              R_DBNAM2
                ' **     9       8     tbl_id2
                ' **    10       9     tbl_name2              R_TNAM2
                ' **    11      10     rel_fld_cnt            R_FLDS
                ' **    12      11     rel_attributes         R_ATTR
                ' **    13      12     rel_found              R_FND
                ' **    14      13     fld_id1
                ' **    15      14     relfld_name1           R_FNAM1
                ' **    16      15     relfld_foreignname1    R_FNAM1F
                ' **    17      16     relfld_order1
                ' **    18      17     datatype_db_type1      R_DTYP1
                ' **    19      18     fld_id2
                ' **    20      19     relfld_name2           R_FNAM2
                ' **    21      20     relfld_foreignname2    R_FNAM2F
                ' **    22      21     relfld_order2
                ' **    23      22     datatype_db_type2      R_DTYP2
                ' **    24      23     Enforce
                ' **    25      24     DontEnforce
                ' **    26      25     Unique
                ' **    27      26     CascadeUpdate
                ' **    28      27     CascadeDelete
                ' **    29      28     LeftJoin
                ' **    30      29     RightJoin
                ' **    31      30     Inherited
                ' **
                ' ********************************************************
7520          End If
7530          .Close
7540        End With  ' ** rst.

7550        If blnCalled = False Then
7560          Debug.Print "'NON-CONTIGUOUS:"
7570          Debug.Print "'==============="
7580        End If

7590      Case False

            ' ** qryRelation_04 (qryRelation_03 (tblRelation, with add'l fields, just contiguous relationships),
            ' ** linked to tblRelation_Field, qryRelation_01 (tblDatabase_Table_Field, with add'l fields),
            ' ** with fld_id1, fld_id2, fld_id3), with fld_id4, fld_id5, fld_id6, fld_id7.
7600        Set qdf = .QueryDefs("qryRelation_05")
7610        Set rst = qdf.OpenRecordset
7620        With rst
7630          If .BOF = True And .EOF = True Then
                ' ** You shouldn't even be here!
7640            lngRels = 0&
7650            Debug.Print "'NO RELATIONSHIPS FOUND!"
7660          Else
7670            .MoveLast
7680            lngRels = .RecordCount
7690            .MoveFirst
7700            arr_varRel = .GetRows(lngRels)
                ' ********************************************************
                ' ** Array: arr_varRel()
                ' **
                ' **   Field  Element  Name                   Constant
                ' **   =====  =======  =====================  ==========
                ' **     1       0     rel_id                 R_RID
                ' **     2       1     rel_name               R_NAM
                ' **     3       2     dbs_id1                R_DID1
                ' **     4       3     dbs_name1              R_DBNAM1
                ' **     5       4     tbl_id1
                ' **     6       5     tbl_name1              R_TNAM1
                ' **     7       6     dbs_id2                R_DID2
                ' **     8       7     dbs_name2              R_DBNAM2
                ' **     9       8     tbl_id2
                ' **    10       9     tbl_name2              R_TNAM2
                ' **    11      10     rel_fld_cnt            R_FLDS
                ' **    12      11     rel_attributes         R_ATTR
                ' **    13      12     rel_found              R_FND
                ' **    14      13     fld_id1
                ' **    15      14     relfld_name1           R_FNAM1
                ' **    16      15     relfld_foreignname1    R_FNAM1F
                ' **    17      16     relfld_order1
                ' **    18      17     datatype_db_type1      R_DTYP1
                ' **    19      18     fld_id2
                ' **    20      19     relfld_name2           R_FNAM2
                ' **    21      20     relfld_foreignname2    R_FNAM2F
                ' **    22      21     relfld_order2
                ' **    23      22     datatype_db_type2      R_DTYP2
                ' **    24      23     fld_id3
                ' **    25      24     relfld_name3           R_FNAM3
                ' **    26      25     relfld_foreignname3    R_FNAM3F
                ' **    27      26     relfld_order3
                ' **    28      27     datatype_db_type3      R_DTYP3
                ' **    29      28     fld_id4
                ' **    30      29     relfld_name4           R_FNAM4
                ' **    31      30     relfld_foreignname4    R_FNAM4F
                ' **    32      31     relfld_order4
                ' **    33      32     datatype_db_type4      R_DTYP4
                ' **    34      33     fld_id5
                ' **    35      34     relfld_name5           R_FNAM5
                ' **    36      35     relfld_foreignname5    R_FNAM5F
                ' **    37      36     relfld_order5
                ' **    38      37     datatype_db_type5      R_DTYP5
                ' **    39      38     fld_id6
                ' **    40      39     relfld_name6           R_FNAM6
                ' **    41      40     relfld_foreignname6    R_FNAM6F
                ' **    42      41     relfld_order6
                ' **    43      42     datatype_db_type6      R_DTYP6
                ' **    44      43     fld_id7
                ' **    45      44     relfld_name7           R_FNAM7
                ' **    46      45     relfld_foreignname7    R_FNAM7F
                ' **    47      46     relfld_order7
                ' **    48      47     datatype_db_type7      R_DTYP7
                ' **    49      48     Enforce
                ' **    50      49     DontEnforce
                ' **    51      50     Unique
                ' **    52      51     CascadeUpdate
                ' **    53      52     CascadeDelete
                ' **    54      53     LeftJoin
                ' **    55      54     RightJoin
                ' **    56      55     Inherited
                ' **
                ' ********************************************************
7710          End If
7720          .Close
7730        End With  ' ** rst.

7740        Debug.Print "'CONTIGUOUS:"
7750        Debug.Print "'==========="

7760      End Select  ' ** blnNonContig.

7770      Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.

7780      For lngX = 0& To (lngRels - 1&)

7790        blnFound = False
7800        strTmp02 = "'" & Left(CStr(lngX + 1&) & ".    ", 5)

            ' ** Maximum name length: 64 chars.
7810        For Each Rel In .Relations
7820          With Rel

7830            blnContinue = True
7840            If .Table = arr_varRel(R_TNAM1, lngX) And .ForeignTable = arr_varRel(R_TNAM2, lngX) Then
7850              If (arr_varRel(R_NAM, lngX) = "tblDatabase_TabletblRelation" And Right(.Name, 29) = "tblDatabase_TabletblRelation1") Or _
                      (arr_varRel(R_NAM, lngX) = "tblDatabase_TabletblRelation1" And Right(.Name, 28) = "tblDatabase_TabletblRelation") Then
7860                blnContinue = False
7870              ElseIf (Left(arr_varRel(R_NAM, lngX), 43) = "tblMsgBoxStyleTypetblVBComponent_MessageBox") Then
7880                Select Case arr_varRel(R_NAM, lngX)
                    Case "tblMsgBoxStyleTypetblVBComponent_MessageBox"
7890                  If Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox2" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox3" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox4" Then
7900                    blnContinue = False
7910                  End If
7920                Case "tblMsgBoxStyleTypetblVBComponent_MessageBox2"
7930                  If Right(.Name, 43) = "tblMsgBoxStyleTypetblVBComponent_MessageBox" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox3" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox4" Then
7940                    blnContinue = False
7950                  End If
7960                Case "tblMsgBoxStyleTypetblVBComponent_MessageBox3"
7970                  If Right(.Name, 43) = "tblMsgBoxStyleTypetblVBComponent_MessageBox" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox2" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox4" Then
7980                    blnContinue = False
7990                  End If
8000                Case "tblMsgBoxStyleTypetblVBComponent_MessageBox4"
8010                  If Right(.Name, 43) = "tblMsgBoxStyleTypetblVBComponent_MessageBox" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox2" Or _
                          Right(.Name, 44) = "tblMsgBoxStyleTypetblVBComponent_MessageBox3" Then
8020                    blnContinue = False
8030                  End If
8040                End Select
8050              Else
                    ' ** Nothing else right now.
8060              End If
8070            ElseIf arr_varRel(R_NAM, lngX) = "tblDataTypeDbtblPricing_AppraiseColumnDataType" Then
                  ' ** 72.  NOT FOUND! tblDataTypeDbtblPricing_AppraiseColumnDataType  DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseColumnDataType
8080              Select Case .Table
                  Case "tblDataTypeDb"
8090                If arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb" Then
                      ' ** It should be all right!
8100                ElseIf arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb1" Then
8110                  arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb"
8120                Else
8130                  blnContinue = False
8140                End If
8150              Case "tblDataTypeDb1"
8160                If arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb1" Then
                      ' ** It should be all right!
8170                ElseIf arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb" Then
8180                  arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb1"
8190                Else
8200                  blnContinue = False
8210                End If
8220              Case Else
8230                blnContinue = False
8240              End Select
8250              If blnContinue = True Then
8260                If .ForeignTable = arr_varRel(R_TNAM2, lngX) Then
                      ' ** Good to go!
8270                Else
8280                  blnContinue = False
8290                End If
8300              End If
8310            ElseIf arr_varRel(R_NAM, lngX) = "tblDataTypeDbtblPricing_AppraiseItemType" Then
                  ' ** 73.  NOT FOUND! tblDataTypeDbtblPricing_AppraiseItemType        DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseItemType
8320              Select Case .Table
                  Case "tblDataTypeDb"
8330                If arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb" Then
                      ' ** It should be all right!
8340                ElseIf arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb1" Then
8350                  arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb"
8360                Else
8370                  blnContinue = False
8380                End If
8390              Case "tblDataTypeDb1"
8400                If arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb1" Then
                      ' ** It should be all right!
8410                ElseIf arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb" Then
8420                  arr_varRel(R_TNAM1, lngX) = "tblDataTypeDb1"
8430                Else
8440                  blnContinue = False
8450                End If
8460              Case Else
8470                blnContinue = False
8480              End Select
8490              If blnContinue = True Then
8500                If .ForeignTable = arr_varRel(R_TNAM2, lngX) Then
                      ' ** Good to go!
8510                Else
8520                  blnContinue = False
8530                End If
8540              End If
8550            Else
8560              blnContinue = False
8570            End If  ' ** R_TNAM1, R_TNAM2.

8580            If blnContinue = True Then
8590              Select Case blnNonContig
                  Case True
8600                If Rel_Attr(arr_varRel(R_ATTR, lngX), "dbRelationDontEnforce") = True Then
8610                  blnFound = True
8620                  arr_varRel(R_FND, lngX) = CBool(True)
8630                  If blnListNonContig = True Then
8640                    Debug.Print "'1 REL: " & .Name & "  TBL1: " & .Table & "  TBL2: " & .ForeignTable
8650                  End If
8660                  Exit For
8670                End If
8680              Case False
8690                strTmp01 = vbNullString
8700                If .Attributes = arr_varRel(R_ATTR, lngX) Then
8710                  blnFound = True
8720                  arr_varRel(R_FND, lngX) = CBool(True)
8730                  strTmp01 = .Name
8740                  Select Case strTmp01
                      Case "tblDataTypeDbtblPricing_AppraiseColumnDataType", "tblDataTypeDb1tblPricing_AppraiseColumnDataType"
                        ' ** 72.  NOT FOUND! tblDataTypeDbtblPricing_AppraiseColumnDataType  DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseColumnDataType
8750                    Stop
8760                  Case "tblDataTypeDbtblPricing_AppraiseItemType", "tblDataTypeDb1tblPricing_AppraiseItemType"
                        ' ** 73.  NOT FOUND! tblDataTypeDbtblPricing_AppraiseItemType        DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseItemType
8770                    Stop
8780                  Case Else
                        ' ** Nothing else.
8790                  End Select
8800                  If strTmp01 = arr_varRel(R_NAM, lngX) Then
8810                    If blnExtraMulti = False And blnListAll = True Then
8820                      Debug.Print "'2 " & strTmp02 & "REL: " & strTmp01 & "  TBL1: " & .Table & "  TBL2: " & .ForeignTable
8830                    End If
8840                  Else
8850                    intPos01 = InStr(strTmp01, "].")
8860                    If intPos01 > 0 Then
8870                      strTmp01 = Mid(strTmp01, (intPos01 + 2))
8880                      If strTmp01 = arr_varRel(R_NAM, lngX) Then
8890                        If blnExtraMulti = False And blnListAll = True Then
8900                          Debug.Print "'3 " & strTmp02 & "REL: " & strTmp01 & "  TBL1: " & .Table & "  TBL2: " & .ForeignTable
8910                        End If
8920                      Else
8930                        If blnExtraMulti = False And blnListAll = True Then
8940                          Debug.Print "'4 " & strTmp02 & " REL NAME: " & strTmp01 & " VS. " & arr_varRel(R_NAM, lngX)
8950                        End If
8960                      End If
8970                    Else
8980                      If blnExtraMulti = False And blnListAll = True Then
8990                        Debug.Print "'5 " & strTmp02 & " REL NAME: " & strTmp01 & " VS. " & arr_varRel(R_NAM, lngX)
9000                      End If
9010                    End If
9020                  End If
9030                  blnFound2 = False
9040                  For lngY = 0& To (lngXRels - 1&)
9050                    If arr_varXRel(XR_NAM, lngY) = strTmp01 Then
9060                      blnFound2 = True
9070                      arr_varXRel(XR_FND, lngY) = CBool(True)
9080                      If arr_varXRel(XR_R_ARR_ELEM, lngY) = vbNullString Then
9090                        arr_varXRel(XR_R_ARR_ELEM, lngY) = CStr(lngX)
9100                      Else
9110                        arr_varXRel(XR_R_ARR_ELEM, lngY) = arr_varXRel(XR_R_ARR_ELEM, lngY) & ";" & CStr(lngX)
9120                      End If
9130                      Exit For
9140                    End If
9150                  Next  ' ** lngY
9160                  If blnFound2 = False Then
9170                    lngXRels = lngXRels + 1&
9180                    lngE = lngXRels - 1&
9190                    ReDim Preserve arr_varXRel(XR_ELEMS, lngE)
9200                    arr_varXRel(XR_NAM, lngE) = strTmp01
9210                    arr_varXRel(XR_FND, lngE) = CBool(True)
9220                    arr_varXRel(XR_TNAM1, lngE) = .Table
9230                    arr_varXRel(XR_TNAM2, lngE) = .ForeignTable
9240                    arr_varXRel(XR_FLDS, lngE) = .Fields.Count
9250                    arr_varXRel(XR_ATTR, lngE) = .Attributes
9260                    arr_varXRel(XR_R_ARR_ELEM, lngE) = CStr(lngX)
9270                  End If
9280                  Exit For
9290                Else
9300                  If Rel_Attr(.Attributes, "dbRelationInherited") = True And _
                          Rel_Attr(arr_varRel(R_ATTR, lngX), "dbRelationInherited") = False Then
9310                    blnFound = True
9320                    arr_varRel(R_FND, lngX) = CBool(True)
9330                    strTmp01 = .Name
9340                    If strTmp01 = arr_varRel(R_NAM, lngX) Then
9350                      If blnExtraMulti = False And blnListAll = True Then
9360                        Debug.Print "'6 " & strTmp02 & "REL: " & strTmp01 & "  TBL1: " & .Table & "  TBL2: " & .ForeignTable
9370                      End If
9380                    Else
9390                      intPos01 = InStr(strTmp01, "].")
9400                      If intPos01 > 0 Then
9410                        strTmp01 = Mid(strTmp01, (intPos01 + 2))
9420                        If strTmp01 = arr_varRel(R_NAM, lngX) Then
9430                          If blnExtraMulti = False And blnListAll = True Then
9440                            Debug.Print "'7 " & strTmp02 & "REL: " & strTmp01 & "  TBL1: " & .Table & "  TBL2: " & .ForeignTable
9450                          End If
9460                        Else
9470                          If blnExtraMulti = False And blnListAll = True Then
9480                            Debug.Print "'8 " & strTmp02 & " REL NAME: " & strTmp01 & " VS. " & arr_varRel(R_NAM, lngX)
9490                          End If
9500                        End If
9510                      Else
9520                        If blnExtraMulti = False And blnListAll = True Then
9530                          Debug.Print "'9 " & strTmp02 & " REL NAME: " & strTmp01 & " VS. " & arr_varRel(R_NAM, lngX)
9540                        End If
9550                      End If
9560                    End If
9570                    blnFound2 = False
9580                    For lngY = 0& To (lngXRels - 1&)
9590                      If arr_varXRel(XR_NAM, lngY) = strTmp01 Then
9600                        blnFound2 = True
9610                        arr_varXRel(XR_FND, lngY) = CBool(True)
9620                        If arr_varXRel(XR_R_ARR_ELEM, lngY) = vbNullString Then
9630                          arr_varXRel(XR_R_ARR_ELEM, lngY) = CStr(lngX)
9640                        Else
9650                          arr_varXRel(XR_R_ARR_ELEM, lngY) = arr_varXRel(XR_R_ARR_ELEM, lngY) & ";" & CStr(lngX)
9660                        End If
9670                        Exit For
9680                      End If
9690                    Next  ' ** lngY
9700                    If blnFound2 = False Then
9710                      lngXRels = lngXRels + 1&
9720                      lngE = lngXRels - 1&
9730                      ReDim Preserve arr_varXRel(XR_ELEMS, lngE)
9740                      arr_varXRel(XR_NAM, lngE) = strTmp01
9750                      arr_varXRel(XR_FND, lngE) = CBool(True)
9760                      arr_varXRel(XR_TNAM1, lngE) = .Table
9770                      arr_varXRel(XR_TNAM2, lngE) = .ForeignTable
9780                      arr_varXRel(XR_FLDS, lngE) = .Fields.Count
9790                      arr_varXRel(XR_ATTR, lngE) = .Attributes
9800                      arr_varXRel(XR_R_ARR_ELEM, lngE) = CStr(lngX)
9810                    End If
9820                    Exit For
9830                  End If  ' ** dbRelationInherited.
9840                End If  ' ** R_ATTR.
9850              End Select  ' ** blnNonContig.
9860            End If  ' ** blnContinue.

9870            strTmp01 = .Name
9880            intPos01 = InStr(strTmp01, "].")
9890            If intPos01 > 0 Then
9900              strTmp01 = Mid(strTmp01, (intPos01 + 2))
9910            End If

                ' ** See that every relationship gets into arr_varXRel(), whether matching or not.
9920            blnFound2 = False
9930            For lngY = 0& To (lngXRels - 1&)
9940              If arr_varXRel(XR_NAM, lngY) = strTmp01 Then
9950                blnFound2 = True
9960                Exit For
9970              End If
9980            Next  ' ** lngY
9990            If blnFound2 = False Then
10000             lngXRels = lngXRels + 1&
10010             lngE = lngXRels - 1&
10020             ReDim Preserve arr_varXRel(XR_ELEMS, lngE)
10030             arr_varXRel(XR_NAM, lngE) = strTmp01
10040             arr_varXRel(XR_FND, lngE) = CBool(False)
10050             arr_varXRel(XR_TNAM1, lngE) = .Table
10060             arr_varXRel(XR_TNAM2, lngE) = .ForeignTable
10070             arr_varXRel(XR_FLDS, lngE) = .Fields.Count
10080             arr_varXRel(XR_ATTR, lngE) = .Attributes
10090             arr_varXRel(XR_R_ARR_ELEM, lngE) = vbNullString
10100           End If

10110         End With  ' ** rel.
10120       Next  ' ** rel.

10130       If blnFound = False Then
10140         blnHasMissing = True
10150         If blnExtraMulti = False And blnCalled = False Then
10160           Debug.Print strTmp02 & "NOT FOUND 1! " & Left(arr_varRel(R_NAM, lngX) & Space(46), 46) & "  DB1: " & arr_varRel(R_DBNAM1, lngX) & "  TBL1: " & arr_varRel(R_TNAM1, lngX) & "  DB2: " & arr_varRel(R_DBNAM2, lngX) & "  TBL2: " & arr_varRel(R_TNAM2, lngX)
10170         End If
10180       End If
10190       Set Rel = Nothing

10200     Next  ' ** lngRels: lngX)

10210     Select Case blnNonContig
          Case True
10220       blnFound2 = False
10230       For lngX = 0& To (lngRels - 1)
10240         If arr_varRel(R_FND, lngX) = False Then
10250           blnFound2 = True
10260           Exit For
10270         End If
10280       Next
10290     Case False
10300       blnFound2 = False
10310       For lngX = 0& To (lngRels - 1)
10320         If arr_varRel(R_FND, lngX) = False Then
10330           blnFound2 = True
10340           blnHasMissing = True
10350           Debug.Print "'NOT FOUND 2! " & arr_varRel(R_NAM, lngX) & "  DB1: " & arr_varRel(R_DBNAM1, lngX) & "  TBL1: " & arr_varRel(R_TNAM1, lngX) & "  DB2: " & arr_varRel(R_DBNAM2, lngX) & "  TBL2: " & arr_varRel(R_TNAM2, lngX)
10360         End If
10370       Next
10380     End Select

10390     lngErrs = 0&
10400     ReDim arr_varErr(E_ELEMS, 0)

10410     Select Case blnFound2
          Case True
10420       Select Case blnRegen
            Case True
10430         For lngX = 0& To (lngRels - 1)
10440           blnContinue = True
10450           If arr_varRel(R_FND, lngX) = False Then
10460             If blnNonContig = False And arr_varRel(R_DID1, lngX) <> lngThisDbsID And arr_varRel(R_DID2, lngX) <> lngThisDbsID Then
                    ' ** If both dbs_id's are in other databases, don't try to regenerate them here.
10470               blnContinue = False
10480             Else
10490               Set Rel = .CreateRelation(arr_varRel(R_NAM, lngX), arr_varRel(R_TNAM1, lngX), arr_varRel(R_TNAM2, lngX), arr_varRel(R_ATTR, lngX))
10500               With Rel
10510                 For lngY = 1& To arr_varRel(R_FLDS, lngX)
10520                   Select Case lngY
                        Case 1&
10530                     .Fields.Append .CreateField(arr_varRel(R_FNAM1, lngX), arr_varRel(R_DTYP1, lngX))
10540                     .Fields(arr_varRel(R_FNAM1, lngX)).ForeignName = arr_varRel(R_FNAM1F, lngX)
10550                   Case 2&
10560                     .Fields.Append .CreateField(arr_varRel(R_FNAM2, lngX), arr_varRel(R_DTYP2, lngX))
10570                     .Fields(arr_varRel(R_FNAM2, lngX)).ForeignName = arr_varRel(R_FNAM2F, lngX)
10580                   Case 3&
10590                     .Fields.Append .CreateField(arr_varRel(R_FNAM3, lngX), arr_varRel(R_DTYP3, lngX))
10600                     .Fields(arr_varRel(R_FNAM3, lngX)).ForeignName = arr_varRel(R_FNAM3F, lngX)
10610                   Case 4&
10620                     .Fields.Append .CreateField(arr_varRel(R_FNAM4, lngX), arr_varRel(R_DTYP4, lngX))
10630                     .Fields(arr_varRel(R_FNAM4, lngX)).ForeignName = arr_varRel(R_FNAM4F, lngX)
10640                   Case 5&
10650                     .Fields.Append .CreateField(arr_varRel(R_FNAM5, lngX), arr_varRel(R_DTYP5, lngX))
10660                     .Fields(arr_varRel(R_FNAM5, lngX)).ForeignName = arr_varRel(R_FNAM5F, lngX)
10670                   Case 6&
10680                     .Fields.Append .CreateField(arr_varRel(R_FNAM6, lngX), arr_varRel(R_DTYP6, lngX))
10690                     .Fields(arr_varRel(R_FNAM6, lngX)).ForeignName = arr_varRel(R_FNAM6F, lngX)
10700                   Case 7&
10710                     .Fields.Append .CreateField(arr_varRel(R_FNAM7, lngX), arr_varRel(R_DTYP7, lngX))
10720                     .Fields(arr_varRel(R_FNAM7, lngX)).ForeignName = arr_varRel(R_FNAM7F, lngX)
10730                   End Select
10740                 Next  ' ** R_FLDS: lngY.
10750               End With
10760 On Error Resume Next
10770               .Relations.Append Rel
10780               If ERR.Number <> 0 Then
10790                 blnContinue = False
10800                 lngErrs = lngErrs + 1&
10810                 lngE = lngErrs - 1&
10820                 ReDim Preserve arr_varErr(E_ELEMS, lngE)
10830                 arr_varErr(E_NUM, lngE) = ERR.Number
10840                 arr_varErr(E_DSC, lngE) = ERR.description
10850                 arr_varErr(E_RNAM, lngE) = Rel.Name
10860 On Error GoTo ERRH
10870               Else
10880 On Error GoTo ERRH
10890               End If
10900             End If
10910             If blnContinue = True Then
10920               lngNewRels = lngNewRels + 1&
10930               If blnCalled = False Then
10940                 Debug.Print "'NEW REL: " & arr_varRel(R_NAM, lngX) & "  TBL1: " & arr_varRel(R_TNAM1, lngX) & "  TBL2: " & arr_varRel(R_TNAM2, lngX)
10950               End If
10960             End If  ' ** blnContinue.
10970           End If  ' ** R_FND.
10980         Next  ' ** lngRels: lngX.
10990       Case False
11000         Debug.Print "'MISSING NON-CONTIGUOUS RELATIONSHIPS DETECTED!"
11010       End Select
11020     Case False
11030       Select Case blnNonContig
            Case True
11040         Select Case blnRegen
              Case True
11050           Debug.Print "'NO REGENERATION NECESSARY - ALL NON-CONTIGUOUS RELATIONSHIPS PRESENT!"
11060         Case False
11070           Debug.Print "'ALL NON-CONTIGUOUS RELATIONSHIPS PRESENT!"
11080         End Select
11090       Case False
11100         Select Case blnRegen
              Case True
11110           Debug.Print "'NO REGENERATION NECESSARY - ALL CONTIGUOUS RELATIONSHIPS PRESENT!"
11120         Case False
11130           Debug.Print "'ALL CONTIGUOUS RELATIONSHIPS PRESENT!"
11140         End Select
11150       End Select
11160     End Select

11170     .Close
11180   End With  ' ** dbs.

        ' ** ALL NON-CONTIGUOUS RELATIONSHIPS PRESENT!
        ' ** DONE!  Rel_Missing()

11190   If blnExtraMulti = True Then
11200     For lngX = 0& To (lngXRels - 1&)
11210       If arr_varXRel(XR_FND, lngX) = False Then
11220         Debug.Print "'     EXTRA REL: " & arr_varXRel(XR_NAM, lngX) & "  TBL1: " & arr_varXRel(XR_TNAM1, lngX) & "  TBL2: " & arr_varXRel(XR_TNAM2, lngX)
11230       End If
11240     Next
11250     For lngX = 0& To (lngXRels - 1&)
11260       If InStr(arr_varXRel(XR_R_ARR_ELEM, lngX), ";") > 0 Then
11270         Debug.Print "'     MULTI REL: " & arr_varXRel(XR_NAM, lngX) & "  TBL1: " & arr_varXRel(XR_TNAM1, lngX) & "  TBL2: " & arr_varXRel(XR_TNAM2, lngX)
11280       End If
11290     Next
11300   End If

        'NON-CONTIGUOUS:
        '===============
'6.   NOT FOUND 1! tblReportzz_tbl_Report_VBComponent_01           DB1: TrustAux.mdb  TBL1: tblReport  DB2: Trust.mdb  TBL2: zz_tbl_Report_VBComponent_01
        'MISSING NON-CONTIGUOUS RELATIONSHIPS DETECTED!
        'DONE!  Rel_Missing()

11310   If lngErrs > 0& Then
11320     Debug.Print "'ERRORS!  " & CStr(lngErrs)
11330     DoEvents
11340     Stop
11350     For lngX = 0& To (lngErrs - 1&)
11360       Debug.Print "'ERROR! " & CStr(arr_varErr(E_NUM, lngX)) & "  " & arr_varErr(E_DSC, lngX)
11370       Debug.Print "'  REL: " & arr_varErr(E_RNAM, lngX)
11380       DoEvents
11390       If ((lngX + 1&) Mod 50) = 0 Then
11400         Stop
11410       End If
11420     Next
11430   End If

        'NON-CONTIGUOUS:
        '===============
'1.   NOT FOUND 1! journaltypetblJournalType                       DB1: TrustDta.mdb  TBL1: journaltype  DB2: TrustAux.mdb  TBL2: tblJournalType
'2.   NOT FOUND 1! m_REVCODE_TYPEtblTaxCode                        DB1: TrustDta.mdb  TBL1: m_REVCODE_TYPE  DB2: TrustAux.mdb  TBL2: tblTaxCode
'3.   NOT FOUND 1! TaxCodetblTaxCode                               DB1: TrustDta.mdb  TBL1: TaxCode  DB2: TrustAux.mdb  TBL2: tblTaxCode
'4.   NOT FOUND 1! TaxCode_TypetblTaxCodeType                      DB1: TrustDta.mdb  TBL1: TaxCode_Type  DB2: TrustAux.mdb  TBL2: tblTaxCodeType
'6.   NOT FOUND 1! tblReportzz_tbl_Report_VBComponent_01           DB1: TrustAux.mdb  TBL1: tblReport  DB2: Trust.mdb  TBL2: zz_tbl_Report_VBComponent_01
        'NEW REL: journaltypetblJournalType              TBL1: journaltype     TBL2: tblJournalType
        'NEW REL: m_REVCODE_TYPEtblTaxCode               TBL1: m_REVCODE_TYPE  TBL2: tblTaxCode
        'NEW REL: TaxCodetblTaxCode                      TBL1: TaxCode         TBL2: tblTaxCode
        'NEW REL: TaxCode_TypetblTaxCodeType             TBL1: TaxCode_Type    TBL2: tblTaxCodeType
        'NEW REL: tblReportzz_tbl_Report_VBComponent_01  TBL1: tblReport       TBL2: zz_tbl_Report_VBComponent_01
        'DONE!  Rel_Missing()

        'CONTIGUOUS:
        '===========
'3.   NOT FOUND 1! {0B892C0B-8958-44B6-B3F9-2DA86BE429F2}          DB1: Trust.mdb  TBL1: tblSecurity_Group  DB2: Trust.mdb  TBL2: tblSecurity_GroupUser
'5.   NOT FOUND 1! {15D683DF-679F-4AD4-8766-5F7A00A92F12}          DB1: Trust.mdb  TBL1: tblRelation_View  DB2: Trust.mdb  TBL2: tblRelation_View_Window
'15.  NOT FOUND 1! {49EC4C69-930F-4C60-BF46-CC4470CD685E}          DB1: Trust.mdb  TBL1: tblSecurity_User  DB2: Trust.mdb  TBL2: tblSecurity_GroupUser
'21.  NOT FOUND 1! {5FE9140F-0A18-4B47-8254-A0BF4CEA889D}          DB1: Trust.mdb  TBL1: tblTemplate_Form_Control  DB2: Trust.mdb  TBL2: tblTemplate_Form_Graphics
'26.  NOT FOUND 1! {7530ED5C-75A8-4388-8531-350E99B9E7CD}          DB1: Trust.mdb  TBL1: tblTemplate_Form_Graphics  DB2: Trust.mdb  TBL2: tblTemplate_Form_Graphics_PictureData
'29.  NOT FOUND 1! {78C00C14-A236-4BC5-BAAE-8BEC2383D67D}          DB1: Trust.mdb  TBL1: tblTemplate_Database_Table  DB2: Trust.mdb  TBL2: tblTemplate_Database_Table_Link
'35.  NOT FOUND 1! {99BFAE34-862A-48D7-B415-DCCA7EB1121C}          DB1: Trust.mdb  TBL1: tblTemplate_Database  DB2: Trust.mdb  TBL2: tblTemplate_Form
'43.  NOT FOUND 1! {AFEF5575-88BD-4509-8C30-5DCEC8897BDA}          DB1: Trust.mdb  TBL1: tblTemplate_Database  DB2: Trust.mdb  TBL2: tblTemplate_Database_Table
'56.  NOT FOUND 1! {FE6027F1-7C2F-4DBB-BDBA-3632C69ADB61}          DB1: Trust.mdb  TBL1: tblTemplate_Form  DB2: Trust.mdb  TBL2: tblTemplate_Form_Control

        'THESE ARE JUST NOT FOUND HERE! THEY DO EXIST IN THEIR HOME DATABASE!
'10.  NOT FOUND 1! {26B2BF09-7874-4EB9-A14B-C7EC272D5389}          DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseColumnDataType
'14.  NOT FOUND 1! {4963733D-4B97-4090-8D57-D5D4F3B5AE49}          DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseItemType
'33.  NOT FOUND 1! {8EF786A6-1686-4557-A3F1-B2BC52D80FB1}          DB1: TrustDta.mdb  TBL1: tblRelation_View  DB2: TrustDta.mdb  TBL2: tblRelation_View_Window

'149. NOT FOUND 1! tblDatabasetblRelation_View                     DB1: TrustAux.mdb  TBL1: tblDatabase  DB2: TrustAux.mdb  TBL2: tblRelation_View
'353. NOT FOUND 1! tblRelation_ViewtblRelation_View_Window         DB1: TrustAux.mdb  TBL1: tblRelation_View  DB2: TrustAux.mdb  TBL2: tblRelation_View_Window
'468. NOT FOUND 1! tblXAdmin_GraphicstblTreeView_Icon              DB1: TrustAux.mdb  TBL1: tblXAdmin_Graphics  DB2: TrustAux.mdb  TBL2: tblTreeView_Icon

        'NOT FOUND 2! {0B892C0B-8958-44B6-B3F9-2DA86BE429F2}  DB1: Trust.mdb  TBL1: tblSecurity_Group  DB2: Trust.mdb  TBL2: tblSecurity_GroupUser
        'NOT FOUND 2! {15D683DF-679F-4AD4-8766-5F7A00A92F12}  DB1: Trust.mdb  TBL1: tblRelation_View  DB2: Trust.mdb  TBL2: tblRelation_View_Window
        'NOT FOUND 2! {49EC4C69-930F-4C60-BF46-CC4470CD685E}  DB1: Trust.mdb  TBL1: tblSecurity_User  DB2: Trust.mdb  TBL2: tblSecurity_GroupUser
        'NOT FOUND 2! {5FE9140F-0A18-4B47-8254-A0BF4CEA889D}  DB1: Trust.mdb  TBL1: tblTemplate_Form_Control  DB2: Trust.mdb  TBL2: tblTemplate_Form_Graphics
        'NOT FOUND 2! {7530ED5C-75A8-4388-8531-350E99B9E7CD}  DB1: Trust.mdb  TBL1: tblTemplate_Form_Graphics  DB2: Trust.mdb  TBL2: tblTemplate_Form_Graphics_PictureData
        'NOT FOUND 2! {78C00C14-A236-4BC5-BAAE-8BEC2383D67D}  DB1: Trust.mdb  TBL1: tblTemplate_Database_Table  DB2: Trust.mdb  TBL2: tblTemplate_Database_Table_Link
        'NOT FOUND 2! {99BFAE34-862A-48D7-B415-DCCA7EB1121C}  DB1: Trust.mdb  TBL1: tblTemplate_Database  DB2: Trust.mdb  TBL2: tblTemplate_Form
        'NOT FOUND 2! {AFEF5575-88BD-4509-8C30-5DCEC8897BDA}  DB1: Trust.mdb  TBL1: tblTemplate_Database  DB2: Trust.mdb  TBL2: tblTemplate_Database_Table
        'NOT FOUND 2! {FE6027F1-7C2F-4DBB-BDBA-3632C69ADB61}  DB1: Trust.mdb  TBL1: tblTemplate_Form  DB2: Trust.mdb  TBL2: tblTemplate_Form_Control

        'NOT FOUND 2! {26B2BF09-7874-4EB9-A14B-C7EC272D5389}  DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseColumnDataType
        'NOT FOUND 2! {4963733D-4B97-4090-8D57-D5D4F3B5AE49}  DB1: TrustDta.mdb  TBL1: tblDataTypeDb  DB2: TrustDta.mdb  TBL2: tblPricing_AppraiseItemType
        'NOT FOUND 2! {8EF786A6-1686-4557-A3F1-B2BC52D80FB1}  DB1: TrustDta.mdb  TBL1: tblRelation_View  DB2: TrustDta.mdb  TBL2: tblRelation_View_Window

        'NOT FOUND 2! tblDatabasetblRelation_View             DB1: TrustAux.mdb  TBL1: tblDatabase  DB2: TrustAux.mdb  TBL2: tblRelation_View
        'NOT FOUND 2! tblRelation_ViewtblRelation_View_Window DB1: TrustAux.mdb  TBL1: tblRelation_View  DB2: TrustAux.mdb  TBL2: tblRelation_View_Window
        'NOT FOUND 2! tblXAdmin_GraphicstblTreeView_Icon      DB1: TrustAux.mdb  TBL1: tblXAdmin_Graphics  DB2: TrustAux.mdb  TBL2: tblTreeView_Icon

        'NEW REL: {0B892C0B-8958-44B6-B3F9-2DA86BE429F2}  TBL1: tblSecurity_Group  TBL2: tblSecurity_GroupUser
        'NEW REL: {15D683DF-679F-4AD4-8766-5F7A00A92F12}  TBL1: tblRelation_View  TBL2: tblRelation_View_Window
        'NEW REL: {49EC4C69-930F-4C60-BF46-CC4470CD685E}  TBL1: tblSecurity_User  TBL2: tblSecurity_GroupUser
        'NEW REL: {5FE9140F-0A18-4B47-8254-A0BF4CEA889D}  TBL1: tblTemplate_Form_Control  TBL2: tblTemplate_Form_Graphics
        'NEW REL: {7530ED5C-75A8-4388-8531-350E99B9E7CD}  TBL1: tblTemplate_Form_Graphics  TBL2: tblTemplate_Form_Graphics_PictureData
        'NEW REL: {78C00C14-A236-4BC5-BAAE-8BEC2383D67D}  TBL1: tblTemplate_Database_Table  TBL2: tblTemplate_Database_Table_Link
        'NEW REL: {99BFAE34-862A-48D7-B415-DCCA7EB1121C}  TBL1: tblTemplate_Database  TBL2: tblTemplate_Form
        'NEW REL: {AFEF5575-88BD-4509-8C30-5DCEC8897BDA}  TBL1: tblTemplate_Database  TBL2: tblTemplate_Database_Table
        'NEW REL: {FE6027F1-7C2F-4DBB-BDBA-3632C69ADB61}  TBL1: tblTemplate_Form  TBL2: tblTemplate_Form_Control
        'DONE!  Rel_Missing()

11440   Select Case blnCalled
        Case True
11450     If lngNewRels > 0& Then
11460       Debug.Print "'NEW RELS: " & CStr(lngNewRels) & "  " & THIS_PROC & "()"
11470     End If
11480   Case False
11490     Debug.Print "'DONE!  " & THIS_PROC & "()"
11500   End Select

11510   Beep

EXITP:
11520   Set Rel = Nothing
11530   Set rst = Nothing
11540   Set qdf = Nothing
11550   Set dbs = Nothing
11560   Rel_Missing = blnRetVal
11570   Exit Function

ERRH:
11580   blnRetVal = False
11590   Select Case ERR.Number
        Case Else
11600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11610   End Select
11620   Resume EXITP

End Function

Public Function Rel_Add() As Boolean
' ** Add a new relationship.

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "Rel_Add"

        Dim wrk As DAO.Workspace, dbs As DAO.Database, Rel As DAO.Relation
        Dim strTblName1 As String, strTblName2 As String, strFldName1 As String, strFldName2 As String
        Dim strPath As String, strFile As String, strPathFile As String
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

11710 On Error GoTo 0

11720   blnRetVal = True

11730   strTblName1 = "tblCompilerDirectiveOption"
11740   strTblName2 = "tblVBComponent_Declaration_Local_Detail"
11750   strFldName1 = "compdiropt_type"
11760   strFldName2 = "compdiropt_type"

11770   strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
11780   strFile = gstrFile_AuxDataName
11790   strPathFile = strPath & LNK_SEP & strFile

11800   Set wrk = CreateWorkspace("tmpDB", "Superuser", TA_SEC, dbUseJet)
11810   With wrk

11820     Set dbs = .OpenDatabase(strPathFile, False, False)  ' ** {pathfile}, {exclusive}, {read-only}
11830     With dbs

11840       strTmp01 = "tblCompilerDirectiveOptiontblVBComponentDeclarationLocalDetail"  ' ** 62 Chars.
11850       Set Rel = .CreateRelation(strTmp01, strTblName1, strTblName2, (dbRelationUpdateCascade + dbRelationDeleteCascade))
11860       With Rel
11870         .Fields.Append .CreateField(strFldName1, dbLong)
11880         .Fields(strFldName1).ForeignName = strFldName2
11890       End With
11900       .Relations.Append Rel

11910       .Close
11920     End With  ' ** dbs.
11930     Set dbs = Nothing

11940     .Close
11950   End With  ' ** wrk.
11960   Set wrk = Nothing

11970   Beep

11980   Debug.Print "'DONE!"

EXITP:
11990   Set Rel = Nothing
12000   Set dbs = Nothing
12010   Set wrk = Nothing
12020   Rel_Add = blnRetVal
12030   Exit Function

ERRH:
12040   Select Case ERR.Number
        Case Else
12050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12060   End Select
12070   Resume EXITP

End Function

Public Function Rel_Attr(varAttr As Variant, strType As String) As Boolean
' ** Return True/False as to whether specified attribute is present.
' ** REMEMBER: DontEnforce is True for "Don't Enforce Referential Integrity",
' **           so my showing "Enforce" means "Not DontEnforce".
' **           I want it to say "Yes, Enforce Referential Integrity".
'KEEP!!
12100 On Error GoTo ERRH

        Const THIS_PROC As String = "Rel_Attr"

        Dim blnRetVal As Boolean

12110   If IsNull(varAttr) = False Then
12120     Select Case strType
          Case "dbRelationEnforce"
            ' ** dbRelationEnforce       0        The relationship is enforced (referential integrity).
12130       blnRetVal = (varAttr = 0) Or ((varAttr And dbRelationDontEnforce) = 0)
12140     Case "dbRelationUnique"
            ' ** dbRelationUnique        1        The relationship is one-to-one.
12150       blnRetVal = (varAttr And dbRelationUnique)
12160     Case "dbRelationDontEnforce"
            ' ** dbRelationDontEnforce   2        The relationship isn't enforced (no referential integrity).
12170       blnRetVal = (varAttr And dbRelationDontEnforce)
12180     Case "dbRelationInherited"
            ' ** dbRelationInherited     4        The relationship exists in a non-current database that contains the two linked tables.
12190       blnRetVal = (varAttr And dbRelationInherited)
12200     Case "dbRelationUpdateCascade"
            ' ** dbRelationUpdateCascade 256      Updates will cascade.
12210       blnRetVal = (varAttr And dbRelationUpdateCascade)
12220     Case "dbRelationDeleteCascade"
            ' ** dbRelationDeleteCascade 4096     Deletions will cascade.
12230       blnRetVal = (varAttr And dbRelationDeleteCascade)
12240     Case "dbRelationLeft"
            ' ** dbRelationLeft          16777216 In Design view, display a LEFT JOIN as the default join type. Microsoft Access only.
12250       blnRetVal = (varAttr And dbRelationLeft)
12260     Case "dbRelationRight"
            ' ** dbRelationRight         33554432 In Design view, display a RIGHT JOIN as the default join type. Microsoft Access only.
12270       blnRetVal = (varAttr And dbRelationRight)
12280     Case Else
12290       Beep
12300       MsgBox "Oops!"
12310     End Select
12320   End If

EXITP:
12330   Rel_Attr = blnRetVal
12340   Exit Function

ERRH:
12350   blnRetVal = False
12360   Select Case ERR.Number
        Case Else
12370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12380   End Select
12390   Resume EXITP

End Function

Public Function Rel_LinkXAdmin() As Boolean
' ** Link Trust.mdb's tblXAdmin_Relation, tblXAdmin_Relation_Field.

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "Rel_LinkXAdmin"

        Dim strPath As String, strFile As String, strExt As String, strPathFile As String
        Dim strTableName As String, strTableNameNew As String
        Dim blnRetVal As Boolean

12410   blnRetVal = True

12420   strPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
12430   strExt = Parse_Ext(CurrentAppName)  ' ** Module Functions: modFileUtilities.
12440   strFile = gstrFile_AuxDataName 'gstrFile_App & "." & strExt
12450   strPathFile = strPath & LNK_SEP & strFile

12460   strTableName = "tblReport_Control_Specification_A" '"tblXAdmin_Relation"
12470   strTableNameNew = strTableName '& "1"
12480   DoCmd.TransferDatabase acLink, "Microsoft Access", strPathFile, acTable, strTableName, strTableNameNew

12490   strTableName = "tblReport_Control_Specification_B" '"tblXAdmin_Relation_Field"
12500   strTableNameNew = strTableName '& "1"
12510   DoCmd.TransferDatabase acLink, "Microsoft Access", strPathFile, acTable, strTableName, strTableNameNew

12520   CurrentDb.TableDefs.Refresh
12530   CurrentDb.TableDefs.Refresh

12540   Beep

EXITP:
12550   Rel_LinkXAdmin = blnRetVal
12560   Exit Function

ERRH:
12570   blnRetVal = False
12580   Select Case ERR.Number
        Case Else
12590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12600   End Select
12610   Resume EXITP

End Function

Public Function PathCheck(Optional varSilent As Variant) As Boolean
' ** Check all saved database paths.

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "PathCheck"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim strThisDbsName As String, strThisDbsExt As String, strThisDbsPath As String, lngThisDbsID As Long
        Dim strThisBackendPath As String
        Dim lngRecs As Long
        Dim blnFound As Boolean, blnSecurityLicense As Boolean, blnSilent As Boolean
        Dim lngDataCnt As Long, lngTempCnt As Long, lngSecCnt As Long
        Dim strTmp01 As String
        Dim lngX As Long, lngY As Long
        Dim blnRetVal As Boolean

12710   blnRetVal = True

12720   Select Case IsMissing(varSilent)
        Case True
12730     blnSilent = False
12740   Case False
12750     blnSilent = CBool(varSilent)
12760   End Select

12770   strThisDbsName = CurrentAppName  ' ** Module Function: modFileUtilities.
12780   strThisDbsExt = Parse_Ext(strThisDbsName)  ' ** Module Function: modFileUtilities.
12790   strThisDbsPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
12800   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
12810   strThisBackendPath = CurrentBackendPath  ' ** Module Function: modFileUtilities.
12820   lngDataCnt = 0&: lngTempCnt = 0&: lngSecCnt = 0&

12830   Set dbs = CurrentDb
12840   With dbs
12850     For lngX = 1& To 3&
12860       gblnBadLink = False: blnSecurityLicense = False
12870       Select Case lngX
            Case 1&
12880         blnRetVal = TableExists("tblDatabase", , , True)  ' ** Module Function: modFileUtilities.
12890       Case 2&
12900         blnRetVal = TableExists("tblTemplate_Database")  ' ** Module Function: modFileUtilities.
12910       Case 3&
12920         blnRetVal = TableExists("tblSecurityLicense")  ' ** Module Function: modFileUtilities.
12930         blnSecurityLicense = True
12940       End Select
12950       If blnRetVal = True And gblnBadLink = False Then
12960         Select Case lngX
              Case 1&
12970           Set rst = .OpenRecordset("tblDatabase", dbOpenDynaset, dbConsistent)
12980         Case 2&
12990           Set rst = .OpenRecordset("tblTemplate_Database", dbOpenDynaset, dbConsistent)
13000         Case 3&
13010           Set rst = .OpenRecordset("tblSecurityLicense", dbOpenDynaset, dbConsistent)
13020         End Select
13030         With rst
13040           .MoveLast
13050           lngRecs = .RecordCount
13060           .MoveFirst
13070           For lngY = 1& To lngRecs
13080             Select Case blnSecurityLicense
                  Case True
                    ' ** Only client, data, auxilliary.
13090               If ![seclic_clientpath_ta] <> strThisDbsPath Then
13100                 .Edit
13110                 ![seclic_clientpath_ta] = strThisDbsPath
13120                 ![seclic_datemodified] = Now()
13130                 .Update
13140                 lngSecCnt = lngSecCnt + 1&
13150               End If
13160               If ![seclic_datapath_ta] <> strThisBackendPath Then
13170                 .Edit
13180                 ![seclic_clientpath_ta] = strThisBackendPath
13190                 ![seclic_datemodified] = Now()
13200                 .Update
13210                 lngSecCnt = lngSecCnt + 1&
13220               End If
13230               If ![seclic_auxiliarypath] <> strThisDbsPath Then
13240                 .Edit
13250                 ![seclic_clientpath_ta] = strThisDbsPath
13260                 ![seclic_datemodified] = Now()
13270                 lngSecCnt = lngSecCnt + 1&
13280                 .Update
13290               End If
13300             Case False
13310               Select Case ![dbs_name]
                    Case "Trust.mdb", "Trust.mde"
                      ' ** Should be this path, this extension.
13320                 If Parse_Ext(![dbs_name]) <> strThisDbsExt Then  ' ** Module Function: modFileUtilities.
13330                   .Edit
13340                   ![dbs_name] = strThisDbsName
13350                   ![dbs_datemodified] = Now()
13360                   .Update
13370                   If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
13380                 End If
13390                 If ![dbs_path] <> strThisDbsPath Then
13400                   .Edit
13410                   ![dbs_path] = strThisDbsPath
13420                   ![dbs_datemodified] = Now()
13430                   .Update
13440                   If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
13450                 End If
13460               Case "TrustDta.mdb"
13470                 If ![dbs_path] <> strThisBackendPath Then
13480                   .Edit
13490                   ![dbs_path] = strThisBackendPath
13500                   ![dbs_datemodified] = Now()
13510                   .Update
13520                   If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
13530                 End If
13540               Case "TrstArch.mdb"
13550                 If ![dbs_path] <> strThisBackendPath Then
13560                   .Edit
13570                   ![dbs_path] = strThisBackendPath
13580                   ![dbs_datemodified] = Now()
13590                   .Update
13600                   If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
13610                 End If
13620               Case "TrustAux.mdb"
13630                 If ![dbs_path] <> strThisDbsPath Then
13640                   .Edit
13650                   ![dbs_path] = strThisDbsPath
13660                   ![dbs_datemodified] = Now()
13670                   .Update
13680                   If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
13690                 End If
13700               Case "TAJrnTmp.mdb"
                      ' ** Not yet implemented anywhere.
13710                 If ![dbs_path] <> strThisBackendPath Then
13720                   .Edit
13730                   ![dbs_path] = strThisBackendPath
13740                   ![dbs_datemodified] = Now()
13750                   .Update
13760                   If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
13770                 End If
13780               Case "TrstXAdm.mdb", "TrstXAdm.mde"
13790                 strTmp01 = "TrstXAdm." & strThisDbsExt
13800                 strTmp01 = strThisDbsPath & LNK_SEP & strTmp01
13810                 blnFound = FileExists(strTmp01)  ' ** Module Function: modFileUtilities.
13820                 If blnFound = False Then
13830                   strTmp01 = Left(strTmp01, (Len(strTmp01) - 1)) & IIf(Right(strTmp01, 1) = "b", "e", "b")
13840                   blnFound = FileExists(strTmp01)  ' ** Module Function: modFileUtilities.
13850                 End If
13860                 Select Case blnFound
                      Case True
13870                   If Right(![dbs_name], 1) <> Right(strTmp01, 1) Then
13880                     .Edit
13890                     ![dbs_name] = Parse_File(strTmp01)  ' ** Module Function: modFileUtilities.
13900                     ![dbs_datemodified] = Now()
13910                     .Update
13920                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
13930                   End If
13940                   If ![dbs_path] <> Parse_Path(strTmp01) Then  ' ** Module Function: modFileUtilities.
13950                     .Edit
13960                     ![dbs_path] = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
13970                     ![dbs_datemodified] = Now()
13980                     .Update
13990                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
14000                   End If
14010                 Case False
14020                   blnRetVal = False
14030                   Beep
14040                   Select Case blnSilent
                        Case True
14050                     Debug.Print "'Neither TrstXAdm.mdb, nor TrstXAdm.mde were found in this directory!"
14060                     Debug.Print "'  " & strThisDbsPath
14070                   Case False
14080                     MsgBox "Neither TrstXAdm.mdb, nor TrstXAdm.mde were found in this directory!" & vbCrLf & vbCrLf & _
                            strThisDbsPath, vbCritical + vbOKOnly, "Trust Administration Not Found"
14090                   End Select
14100                 End Select
14110               Case "TrustImport.mdb", "TrustImport.mde"
                      ' ** Likely only in development environments for now.
                      ' ** First, move up one directory.
14120                 strTmp01 = Parse_Path(strThisDbsPath)  ' ** Module Function: modFileUtilities.
                      ' ** Then look for its directory.
14130                 blnFound = DirExists(strTmp01 & LNK_SEP & "Trust Import")  ' ** Module Function: modFileUtilities.
14140                 If blnFound = False Then  ' ** Try it without the space.
14150                   blnFound = DirExists(strTmp01 & LNK_SEP & "TrustImport")  ' ** Module Function: modFileUtilities.
14160                   Select Case blnFound
                        Case True
14170                     strTmp01 = strTmp01 & LNK_SEP & "TrustImport"
14180                   Case False
                          ' ** Move up another directory.
14190                     strTmp01 = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
14200                     blnFound = DirExists(strTmp01 & LNK_SEP & "TrustImport")  ' ** Module Function: modFileUtilities.
14210                     If blnFound = True Then
14220                       strTmp01 = strTmp01 & LNK_SEP & "TrustImport"
14230                     End If
14240                   End Select
14250                 Else
14260                   strTmp01 = strTmp01 & LNK_SEP & "Trust Import"
14270                 End If
14280                 Select Case blnFound
                      Case True
14290                   strTmp01 = strTmp01 & LNK_SEP & "TrustImport." & strThisDbsExt
14300                   blnFound = FileExists(strTmp01)  ' ** Module Function: modFileUtilities.
14310                   If blnFound = False Then
14320                     strTmp01 = Left(strTmp01, (Len(strTmp01) - 1)) & IIf(Right(strTmp01, 1) = "b", "e", "b")
14330                     blnFound = FileExists(strTmp01)  ' ** Module Function: modFileUtilities.
14340                     If blnFound = False Then strTmp01 = Left(strTmp01, (Len(strTmp01) - 3)) & strThisDbsExt
14350                   End If
                        ' ** If the directory is there, but not the file, then UFO's took it.
14360                   If Right(![dbs_name], 1) <> Right(strTmp01, 1) Then  ' ** Module Function: modFileUtilities.
14370                     .Edit
14380                     ![dbs_name] = Parse_File(strTmp01)  ' ** Module Function: modFileUtilities.
14390                     ![dbs_datemodified] = Now()
14400                     .Update
14410                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
14420                   End If
14430                   If ![dbs_path] <> Parse_Path(strTmp01) Then  ' ** Module Function: modFileUtilities.
14440                     .Edit
14450                     ![dbs_path] = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
14460                     ![dbs_datemodified] = Now()
14470                     .Update
14480                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
14490                   End If
14500                 Case False
                        ' ** Just give it the hypothetical location.
14510                   strTmp01 = Parse_Path(strThisDbsPath)  ' ** Module Function: modFileUtilities.
14520                   If InStr(Parse_File(strThisDbsPath), " ") > 0 Then
                          ' ** Should be 'Trust Accountant'.
14530                     strTmp01 = strTmp01 & LNK_SEP & "Trust Import"
14540                   Else
                          ' ** Should be 'NewWorking'.
14550                     strTmp01 = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
                          ' ** Should be 'TrustAccountant'
14560                     strTmp01 = strTmp01 & LNK_SEP & "TrustImport"
14570                   End If
14580                   strTmp01 = strTmp01 & LNK_SEP & "TrustImport." & strThisDbsExt
14590                   If Right(![dbs_name], 1) <> Right(strTmp01, 1) Then  ' ** Module Function: modFileUtilities.
14600                     .Edit
14610                     ![dbs_name] = Parse_File(strTmp01)  ' ** Module Function: modFileUtilities.
14620                     ![dbs_datemodified] = Now()
14630                     .Update
14640                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
14650                   End If
14660                   If ![dbs_path] <> Parse_Path(strTmp01) Then  ' ** Module Function: modFileUtilities.
14670                     .Edit
14680                     ![dbs_path] = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
14690                     ![dbs_datemodified] = Now()
14700                     .Update
14710                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
14720                   End If
14730                 End Select
14740               Case "FileSpec.mdb"
                      ' ** First, move up one directory.
14750                 strTmp01 = Parse_Path(strThisDbsPath)  ' ** Module Function: modFileUtilities.
                      ' ** Then look for its directory.
14760                 blnFound = DirExists(strTmp01 & LNK_SEP & "Trust Import")  ' ** Module Function: modFileUtilities.
14770                 If blnFound = False Then  ' ** Try it without the space.
14780                   blnFound = DirExists(strTmp01 & LNK_SEP & "TrustImport")  ' ** Module Function: modFileUtilities.
14790                   Select Case blnFound
                        Case True
14800                     strTmp01 = strTmp01 & LNK_SEP & "TrustImport"
14810                   Case False
                          ' ** Move up another directory.
14820                     strTmp01 = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
14830                     blnFound = DirExists(strTmp01 & LNK_SEP & "TrustImport")  ' ** Module Function: modFileUtilities.
14840                     If blnFound = True Then
14850                       strTmp01 = strTmp01 & LNK_SEP & "TrustImport"
14860                     End If
14870                   End Select
14880                 Else
14890                   strTmp01 = strTmp01 & LNK_SEP & "Trust Import"
14900                 End If
14910                 Select Case blnFound
                      Case True
14920                   strTmp01 = strTmp01 & LNK_SEP & "FileSpec.mdb"
14930                   blnFound = FileExists(strTmp01)  ' ** Module Function: modFileUtilities.
14940                   If blnFound = False Then
                          ' ** Hmmm... Put the path in anyway.
14950                   End If
                        ' ** If the directory is there, but not the file, then UFO's took it.
14960                   If ![dbs_path] <> Parse_Path(strTmp01) Then  ' ** Module Function: modFileUtilities.
14970                     .Edit
14980                     ![dbs_path] = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
14990                     ![dbs_datemodified] = Now()
15000                     .Update
15010                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
15020                   End If
15030                 Case False
                        ' ** Just give it the hypothetical location.
15040                   strTmp01 = Parse_Path(strThisDbsPath)  ' ** Module Function: modFileUtilities.
15050                   If InStr(Parse_File(strThisDbsPath), " ") > 0 Then
                          ' ** Should be 'Trust Accountant'.
15060                     strTmp01 = strTmp01 & LNK_SEP & "Trust Import"
15070                   Else
                          ' ** Should be 'NewWorking'.
15080                     strTmp01 = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
                          ' ** Should be 'TrustAccountant'
15090                     strTmp01 = strTmp01 & LNK_SEP & "TrustImport"
15100                   End If
15110                   strTmp01 = strTmp01 & LNK_SEP & "FileSpec.mdb"
15120                   If ![dbs_path] <> Parse_Path(strTmp01) Then  ' ** Module Function: modFileUtilities.
15130                     .Edit
15140                     ![dbs_path] = Parse_Path(strTmp01)  ' ** Module Function: modFileUtilities.
15150                     ![dbs_datemodified] = Now()
15160                     .Update
15170                     If lngX = 1& Then lngDataCnt = lngDataCnt + 1& Else lngTempCnt = lngTempCnt + 1&
15180                   End If
15190                 End Select
15200               End Select
15210             End Select  ' ** blnSecurityLicense.
15220             If lngY < lngRecs Then .MoveNext
15230           Next  ' ** lngY.
15240           .Close
15250         End With  ' ** rst.
15260       End If  ' ** blnRetVal, gblnBadLink.
15270     Next  ' ** lngX.
15280     .Close
15290   End With  ' ** dbs.

15300   If (lngDataCnt + lngTempCnt + lngSecCnt) = 0& Then
          'DoBeeps 2, 100&  ' ** Module Function: modWindowFunctions.
          'MsgBox "No changes were needed!", vbInformation + vbOKOnly, "No Update Needed"
15310   Else
15320     Select Case blnSilent
          Case True
            'Debug.Print "'tblDatabase: " & CStr(lngDataCnt)
            'Debug.Print "'tblTemplate_Database: " & CStr(lngTempCnt)
            'Debug.Print "'tblSecurityLicense: " & CStr(lngSecCnt)
15330     Case False
15340       Beep
15350       MsgBox "Changes were made in the following tables:" & vbCrLf & _
              "  tblDatabase: " & CStr(lngDataCnt) & vbCrLf & _
              "  tblTemplate_Database: " & CStr(lngTempCnt) & vbCrLf & _
              "  tblSecurityLicense: " & CStr(lngSecCnt)
15360     End Select
15370   End If

15380   If gblnBadLink = True Then
15390     Beep
15400     Select Case blnSilent
          Case True
15410       Debug.Print "'BAD LINK FOUND!"
15420     Case False
15430       MsgBox "Bad link found!", vbCritical + vbOKOnly, "Bad Link Found"
15440     End Select
15450   End If

EXITP:
15460   Set rst = Nothing
15470   Set dbs = Nothing
15480   PathCheck = blnRetVal
15490   Exit Function

ERRH:
15500   blnRetVal = False
15510   Select Case ERR.Number
        Case Else
15520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15530   End Select
15540   Resume EXITP

End Function

Public Sub ChkCalStage()

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "ChkCalStage"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCalID As Long
        Dim lngRecs As Long

15610   Set dbs = CurrentDb
15620   With dbs
15630     Set rst = .OpenRecordset("tblCalendar_Staging", dbOpenDynaset, dbConsistent)
15640     If rst.BOF = True And rst.EOF = True Then
15650       rst.Close
            ' ** Append tblForm_Graphics, for frmCalendar, to tblCalendar_Staging.
15660       Set qdf = .QueryDefs("qryCalendar_03")
15670       qdf.Execute dbFailOnError
15680       Set qdf = Nothing
15690       DoEvents
15700       Set rst = .OpenRecordset("tblCalendar_Staging", dbOpenDynaset, dbConsistent)
15710     End If
15720     lngCalID = 0&
15730     rst.MoveLast
15740     lngRecs = rst.RecordCount
15750     If lngRecs > 1& Then
15760       rst.MoveFirst
15770       lngCalID = rst![cal_id]
15780     End If
15790     rst.Close
15800     Set rst = Nothing
15810     If lngCalID > 0& Then
            ' ** Delete tblCalendar_Staging, by specified [calid] (<>[calid]).
15820       Set qdf = .QueryDefs("qryCalendar_04")
15830       With qdf.Parameters
15840         ![calid] = lngCalID
15850       End With
15860       qdf.Execute
15870       Set qdf = Nothing
15880     End If
          ' ** Update tblCalendar_Staging, for unique_id, by specified [unqid].
15890     Set qdf = .QueryDefs("qryCalendar_02")
15900     With qdf.Parameters
15910       ![unqid] = 0&
15920     End With
15930     qdf.Execute
15940     Set qdf = Nothing
15950     .Close
15960   End With  ' ** dbs.

EXITP:
15970   Set rst = Nothing
15980   Set qdf = Nothing
15990   Set dbs = Nothing
16000   Exit Sub

ERRH:
16010   Select Case ERR.Number
        Case Else
16020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16030   End Select
16040   Resume EXITP

End Sub

Public Function JrnlPrefCheck() As Boolean

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlPrefCheck"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngThisDbsID As Long
        Dim lngPrefCtlID As Long, lngFrmID As Long, lngCtlID As Long
        Dim blnAdd As Boolean
        Dim blnRetVal As Boolean

16110   blnRetVal = True

16120   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
16130   blnAdd = False

16140   Set dbs = CurrentDb
16150   With dbs
          ' ** qrySystemStartup_13a (tblPreference_Control, just frmPostingDate, opgInput),
          ' ** linked to tblPreference_User, by specified [dbid], [usr].
16160     Set qdf = .QueryDefs("qrySystemStartup_13b")
16170     With qdf.Parameters
16180       ![dbid] = lngThisDbsID
16190       ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
16200     End With
16210     Set rst = qdf.OpenRecordset
16220     With rst
16230       If .BOF = True And .EOF = True Then
16240         blnAdd = True
16250       Else
16260         .MoveLast
16270         .MoveFirst
16280         Select Case IsNull(![prefuser_id])
              Case True
16290           blnAdd = True
16300         Case False
16310           If ![prefuser_integer] = 1 Or ![prefuser_integer] = 2 Then
                  ' ** Fine.
16320           Else
16330             .Edit
16340             ![prefuser_integer] = 1
16350             ![DateModified] = Now()
16360             .Update
16370           End If
16380         End Select
16390       End If
16400       .Close
16410     End With
16420     Set rst = Nothing
16430     Set qdf = Nothing
16440     If blnAdd = True Then
            ' ** tblPreference_Control, just frmPostingDate, opgInput.
16450       Set qdf = .QueryDefs("qrySystemStartup_13a")
16460       Set rst = qdf.OpenRecordset
16470       With rst
16480         .MoveFirst
16490         lngPrefCtlID = ![prefctl_id]
16500         lngFrmID = ![frm_id]
16510         lngCtlID = ![ctl_id]
16520         .Close
16530       End With
16540       Set rst = Nothing
16550       Set qdf = Nothing
16560       Set rst = .OpenRecordset("tblPreference_User", dbOpenDynaset, dbConsistent)
16570       With rst
16580         .AddNew
16590         ![dbs_id] = lngThisDbsID
16600         ![frm_id] = lngFrmID
16610         ![ctl_id] = lngCtlID
16620         ![prefctl_id] = lngPrefCtlID
              ' ** ![prefuser_id] : AutoNumber.
16630         ![prefuser_integer] = 1
16640         ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
16650         ![DateCreated] = Now()
16660         ![DateModified] = Now()
16670         .Update
16680         .Close
16690       End With
16700     End If
16710     .Close
16720   End With

EXITP:
16730   Set rst = Nothing
16740   Set qdf = Nothing
16750   Set dbs = Nothing
16760   JrnlPrefCheck = blnRetVal
16770   Exit Function

ERRH:
16780   blnRetVal = False
16790   Select Case ERR.Number
        Case Else
16800     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16810   End Select
16820   Resume EXITP

End Function

Public Function JrnlDateCheck() As Boolean

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlDateCheck"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim datTransDate As Date
        Dim lngRecs As Long
        Dim blnFound As Boolean
        Dim lngTmp01 As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

16910   blnRetVal = True

16920   gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.

16930   Set dbs = CurrentDb
16940   With dbs

          ' ** Journal, grouped, with cnt, by specified [usr].
16950     Set qdf = .QueryDefs("qrySystemStartup_14a")
16960     With qdf.Parameters
16970       ![usr] = gstrJournalUser
16980     End With
16990     Set rst = qdf.OpenRecordset
17000     With rst
17010       If .BOF = True And .EOF = True Then
17020         lngRecs = 0&
17030       Else
17040         .MoveFirst
17050         Select Case IsNull(![cnt])
              Case True
17060           lngRecs = 0&
17070         Case False
17080           lngRecs = ![cnt]
17090           If lngRecs > 0& Then
17100             datTransDate = ![transdate]
17110           End If
17120         End Select
17130       End If
17140       .Close
17150     End With
17160     Set rst = Nothing
17170     Set qdf = Nothing

17180     If lngRecs > 0& Then

17190       glngPostingDateID = 0&

17200       Set rst = .OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
17210       With rst
17220         If .BOF = True And .EOF = True Then
17230           .AddNew
                ' ** ![PostingDate_ID] : AutoNumber.
17240           ![Posting_Date] = datTransDate
17250           ![Username] = gstrJournalUser
17260           .Update
17270           .Bookmark = .LastModified
17280           glngPostingDateID = ![PostingDate_ID]
17290         Else
17300           .MoveLast
17310           lngRecs = .RecordCount
17320           .MoveFirst

17330           blnFound = False: lngTmp01 = 0&
17340           For lngX = 1& To lngRecs
17350             Select Case IsNull(![Username])
                  Case True
17360               lngTmp01 = ![PostingDate_ID]
17370             Case False
17380               If ![Username] = gstrJournalUser Then
17390                 blnFound = True
17400                 glngPostingDateID = ![PostingDate_ID]
17410                 Select Case IsNull(![Posting_Date])
                      Case True
17420                   .Edit
17430                   ![Posting_Date] = datTransDate
17440                   .Update
17450                 Case False
17460                   If ![Posting_Date] <> datTransDate Then
17470                     .Edit
17480                     ![Posting_Date] = datTransDate
17490                     .Update
17500                   End If
17510                 End Select
17520                 Exit For
17530               End If
17540             End Select
17550             If lngX < lngRecs Then .MoveNext
17560           Next  ' ** lngX.

17570           If blnFound = False Then
17580             If lngTmp01 > 0& Then
17590               .MoveFirst
17600               .FindFirst "[ID] = " & CStr(lngTmp01)
17610               If .NoMatch = False Then
17620                 glngPostingDateID = ![PostingDate_ID]
17630                 .Edit
17640                 ![Posting_Date] = datTransDate
17650                 ![Username] = gstrJournalUser
17660                 .Update
17670               End If
17680             Else
17690               .AddNew
                    ' ** ![PostingDate_ID] : AutoNumber.
17700               ![Posting_Date] = datTransDate
17710               ![Username] = gstrJournalUser
17720               .Update
17730               .Bookmark = .LastModified
17740               glngPostingDateID = ![PostingDate_ID]
17750             End If
17760           End If  ' ** blnFound.

17770         End If  ' ** BOF, EOF.
17780         .Close
17790       End With  ' ** rst.
17800       Set rst = Nothing

17810       If glngPostingDateID > 0& Then
17820         Set rst = .OpenRecordset("tblCalendar_Staging", dbOpenDynaset, dbReadOnly)
17830         If rst.BOF = True And rst.EOF = True Then
17840           rst.Close
17850           ChkCalStage  ' ** Procedure: Above.
17860           DoEvents
17870         Else
17880           rst.Close
17890         End If
17900         Set rst = Nothing
17910         Set rst = .OpenRecordset("tblCalendar_Staging", dbOpenDynaset, dbConsistent)
17920         With rst
17930           If .BOF = True And .EOF = True Then
                  ' ** I give up!
17940           Else
17950             .MoveFirst
17960             .Edit
17970             ![unique_id] = glngPostingDateID
17980             ![cal_datemodified] = Now()
17990             .Update
18000           End If
18010           .Close
18020         End With  ' ** rst.
18030       End If  ' ** glngPostingDateID.

18040     End If  ' ** lngRecs.

18050     .Close
18060   End With  ' ** dbs.

EXITP:
18070   Set rst = Nothing
18080   Set qdf = Nothing
18090   Set dbs = Nothing
18100   JrnlDateCheck = blnRetVal
18110   Exit Function

ERRH:
18120   Select Case ERR.Number
        Case Else
18130     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18140   End Select
18150   Resume EXITP

End Function

Public Function LocAssetMove_Chk() As Boolean
' ** Not called.

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "LocAssetMove_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strQryName As String, lngParams As Long
        Dim strLtr1 As String, strLtr2 As String
        Dim lngCnt As Long
        Dim strTmp01 As String
        Dim lngW As Long, lngX As Long, lngY As Long, lngZ As Long
        Dim blnRetVal As Boolean

        Const QRY_BASE As String = "qryLocation_Asset_"

18210 On Error GoTo 0

18220   blnRetVal = True

18230   Set dbs = CurrentDb
18240   With dbs
18250     For lngW = 15& To 19&
18260       Select Case lngW
            Case 15&
18270         lngCnt = 1&
18280       Case Else
18290         lngCnt = 2&
18300       End Select
18310       For lngX = 1& To 4&
18320         Select Case lngX
              Case 1&
18330           strLtr1 = "b"
18340         Case 2&
18350           strLtr1 = "d"
18360         Case 3&
18370           strLtr1 = "f"
18380         Case 4&
18390           strLtr1 = "h"
18400         End Select
18410         For lngY = 1& To lngCnt
18420           Select Case lngY
                Case 1&
18430             strLtr2 = "a"
18440           Case 2&
18450             strLtr2 = "b"
18460           End Select
18470           strQryName = QRY_BASE & CStr(lngW) & strLtr1 & strLtr2
18480           Set qdf = .QueryDefs(strQryName)
18490           With qdf
18500             lngParams = .Parameters.Count
18510             For lngZ = 0& To (lngParams - 1&)
18520               strTmp01 = .Parameters(lngZ).Name
18530               Select Case strTmp01
                    Case "[actno]"
18540                 Select Case strLtr1
                      Case "d", "h"
                        ' ** OK.
18550                 Case Else
18560                   Debug.Print "'QRY: " & .Name & "  PARM: " & strTmp01
18570                 End Select
18580               Case "[astno]"
18590                 Select Case strLtr1
                      Case "f", "h"
                        ' ** OK.
18600                 Case Else
18610                   Debug.Print "'QRY: " & .Name & "  PARM: " & strTmp01
18620                 End Select
18630               Case Else
18640                 Debug.Print "'QRY: " & .Name & "  PARM: " & strTmp01
18650               End Select
18660             Next
18670           End With  ' ** qdf.
18680         Next  ' ** lngY.
18690       Next  ' ** lngX.
18700     Next  ' ** lngW.
18710     .Close
18720   End With  ' ** dbs.

18730   Beep

18740   Debug.Print "'DONE!"
18750   DoEvents

        'QRY: qryLocation_Asset_16fa  PARM: [locold]
        'QRY: qryLocation_Asset_16fa  PARM: [locnew]
        'QRY: qryLocation_Asset_16fb  PARM: [locold]
        'QRY: qryLocation_Asset_16fb  PARM: [locnew]

        'QRY: qryLocation_Asset_17fa  PARM: [locold]
        'QRY: qryLocation_Asset_17fa  PARM: [locnew]
        'QRY: qryLocation_Asset_17fb  PARM: [locold]
        'QRY: qryLocation_Asset_17fb  PARM: [locnew]

        'QRY: qryLocation_Asset_18fa  PARM: [locold]
        'QRY: qryLocation_Asset_18fa  PARM: [locnew]
        'QRY: qryLocation_Asset_18fb  PARM: [locold]
        'QRY: qryLocation_Asset_18fb  PARM: [locnew]

        'qryLocation_Asset_15ba
        'qryLocation_Asset_16ba
        'qryLocation_Asset_16bb
        'qryLocation_Asset_17ba
        'qryLocation_Asset_17bb
        'qryLocation_Asset_18ba
        'qryLocation_Asset_18bb
        'qryLocation_Asset_19ba
        'qryLocation_Asset_19bb

        'qryLocation_Asset_15da
        'qryLocation_Asset_16da
        'qryLocation_Asset_16db
        'qryLocation_Asset_17da
        'qryLocation_Asset_17db
        'qryLocation_Asset_18da
        'qryLocation_Asset_18db
        'qryLocation_Asset_19da
        'qryLocation_Asset_19db

        'qryLocation_Asset_15fa
        'qryLocation_Asset_16fa
        'qryLocation_Asset_16fb
        'qryLocation_Asset_17fa
        'qryLocation_Asset_17fb
        'qryLocation_Asset_18fa
        'qryLocation_Asset_18fb
        'qryLocation_Asset_19fa
        'qryLocation_Asset_19fb

        'qryLocation_Asset_15ha
        'qryLocation_Asset_16ha
        'qryLocation_Asset_16hb
        'qryLocation_Asset_17ha
        'qryLocation_Asset_17hb
        'qryLocation_Asset_18ha
        'qryLocation_Asset_18hb
        'qryLocation_Asset_19ha
        'qryLocation_Asset_19hb

EXITP:
18760   Set qdf = Nothing
18770   Set dbs = Nothing
18780   LocAssetMove_Chk = blnRetVal
18790   Exit Function

ERRH:
18800   blnRetVal = False
18810   Select Case ERR.Number
        Case Else
18820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18830   End Select
18840   Resume EXITP

End Function

Public Function AutoNum_Holes() As Boolean
' ** AutoNumber Holes, AutoNumber_Holes, AutoNum Holes.

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "AutoNum_Holes"

        Dim dbs As DAO.Database, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strTblName As String, strFldName As String
        Dim lngHoles As Long, arr_varHole() As Variant
        Dim lngRecs As Long, lngLastNum As Long
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varHole().
        Const H_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const H_ID As Integer = 0

18910   blnRetVal = True

'18920   strTblName = "LedgerArchive"
'18920   strTblName = "ledger"
'18920   strTblName = "tblCapata_Asset"
'18920   strTblName = "tblMark_AutoNum"
18920   strTblName = "tblPreference_Control"
'18930   strFldName = "journalno"
'18930   strFldName = "capast_id"
18930   strFldName = "prefctl_id"

18940   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
18950   DoEvents

18960   Debug.Print "'AUTONUM HOLES: " & strTblName
18970   DoEvents

18980   lngHoles = 0&
18990   ReDim arr_varHole(H_ELEMS, 0)

19000   Set dbs = CurrentDb
19010   With dbs

19020     lngLastNum = 0& '137307   ' ** Use this when doing Ledger and there's LedgerArchive.

          ' ** Empty tblMark_AutoNum.
19030     TableEmpty "tblMark_AutoNum"  ' ** Module Function: modFileUtilities.
          ' ** Empty tblMark_AutoNum2.
'19030     TableEmpty "tblMark_AutoNum2"  ' ** Module Function: modFileUtilities.

19040     Set rst1 = .OpenRecordset(strTblName, dbOpenDynaset, dbReadOnly)
19050     rst1.sort = "[" & strFldName & "]"
19060     Set rst2 = rst1.OpenRecordset
19070     With rst2
19080       .MoveLast
19090       lngRecs = .RecordCount
19100       .MoveFirst
19110       For lngX = 1& To lngRecs
19120         If .Fields(strFldName).Value = (lngLastNum + 1&) Then
19130           lngLastNum = .Fields(strFldName).Value
19140         Else
19150           Do Until .Fields(strFldName).Value = (lngLastNum + 1&)
19160             lngHoles = lngHoles + 1&
19170             lngE = lngHoles - 1&
19180             ReDim Preserve arr_varHole(H_ELEMS, lngE)
19190             arr_varHole(H_ID, lngE) = (lngLastNum + 1&)
19200             lngLastNum = lngLastNum + 1&
19210           Loop
19220           lngLastNum = .Fields(strFldName).Value
19230         End If
19240         If lngX < lngRecs Then .MoveNext
19250       Next  ' ** lngX.
19260       .Close
19270     End With  ' ** rst2.
19280     rst1.Close
19290     Set rst2 = Nothing
19300     Set rst1 = Nothing

19310     Debug.Print "'HOLES: " & CStr(lngHoles)
19320     DoEvents

19330     If lngHoles > 0& Then
19340       Set rst1 = .OpenRecordset("tblMark_AutoNum", dbOpenDynaset, dbAppendOnly)
'19340       Set rst1 = .OpenRecordset("tblMark_AutoNum2", dbOpenDynaset, dbAppendOnly)
19350       With rst1
19360         For lngX = 0& To (lngHoles - 1&)
19370           .AddNew
19380           ![unique_id] = arr_varHole(H_ID, lngX)
19390           ![mark] = False
                ' ** ![value_lng
                ' ** ![value_dbl
                ' ** ![value_txt
                ' ** ![autonum_id] : AutoNumber.
19400           .Update
                'Debug.Print "'" & arr_varHole(H_ID, lngX)
19410         Next
19420         .Close
19430       End With  ' ** rst1.
19440       Set rst1 = Nothing
19450     End If  ' ** lngHoles.

19460     .Close
19470   End With  ' ** dbs.

19480   Beep

19490   Debug.Print "'DONE!"
19500   DoEvents

        'AUTONUM HOLES: tblLedgerHidden
        'HOLES: 8
        'DONE!

EXITP:
19510   Set rst1 = Nothing
19520   Set rst2 = Nothing
19530   Set dbs = Nothing
19540   AutoNum_Holes = blnRetVal
19550   Exit Function

ERRH:
19560   Select Case ERR.Number
        Case Else
19570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19580   End Select
19590   Resume EXITP

End Function

Public Function XAdmin_FrmGfx_Qry() As Boolean
' ** Fill out form RecordSource query with image control names.

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "XAdmin_FrmGfx_Qry"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, fld As DAO.Field
        Dim lngFlds As Long, arr_varFld() As Variant
        Dim strQryName As String, strSQL As String, strNum As String, strSource As String, strFrmName As String
        Dim lngFldNames As Long, lngLastAltNum As Long
        Dim blnExists As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varFld().
        Const F_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const F_FNAM As Integer = 0
        Const F_NUM  As Integer = 1
        Const F_FVAL As Integer = 2
        Const F_IMG  As Integer = 3
        Const F_SQL  As Integer = 4

19610 On Error GoTo 0

19620   blnRetVal = True

19630   strQryName = "qryRpt_CourtReports_NY_01"
19640   strFrmName = "frmRpt_CourtReports_NY"
19650   strSource = "tblForm_Graphics"
19660   blnExists = True
19670   lngLastAltNum = 0&  ' ** Actually, this alt number.
19680   lngFldNames = 7&

19690   Set dbs = CurrentDb
19700   With dbs

19710     If blnExists = False Then
19720       strSQL = "SELECT tblForm_Graphics.frmgfx_id, tblForm_Graphics.dbs_id, tblDatabase.dbs_name, " & _
              "tblForm_Graphics.frm_id, tblForm.frm_name, tblForm_Graphics.frmgfx_alt, tblForm_Graphics.frmgfx_cnt,"
19730       strTmp03 = "FROM (tblDatabase INNER JOIN tblForm ON tblDatabase.dbs_id = tblForm.dbs_id) INNER JOIN " & _
              "(tblForm_Control INNER JOIN tblForm_Graphics ON (tblForm_Control.ctl_id = tblForm_Graphics.ctl_id_01) AND " & _
              "(tblForm_Control.frm_id = tblForm_Graphics.frm_id) AND (tblForm_Control.dbs_id = tblForm_Graphics.dbs_id)) ON " & _
              "(tblForm.frm_id = tblForm_Control.frm_id) AND (tblForm.dbs_id = tblForm_Control.dbs_id)"
19740       strTmp04 = "WHERE (((tblDatabase.dbs_name) In ('Trust.mdb','Trust.mde')) AND " & _
              "((tblForm.frm_name)='" & strFrmName & "') AND ((tblForm_Graphics.frmgfx_alt)=" & CStr(lngLastAltNum) & "));"
19750       strTmp02 = vbNullString
19760       For lngX = 1& To lngFldNames
              ' ** ctl_name_01, xadgfx_image_01.
19770         strNum = Right("00" & CStr(lngX), 2)
19780         strTmp02 = strTmp02 & strSource & ".ctl_name_" & strNum & ", "
19790         strTmp02 = strTmp02 & strSource & ".xadgfx_image_" & strNum & ", "
19800       Next ' ** lngX.
19810       strTmp02 = Trim(strTmp02)
19820       If Right(strTmp02, 1) = "," Then strTmp02 = Left(strTmp02, (Len(strTmp02) - 1))
19830       strSQL = strSQL & strTmp02 & vbCrLf & strTmp03 & vbCrLf & strTmp04
19840       Set qdf = .CreateQueryDef(strQryName, strSQL)
19850       .QueryDefs.Refresh
19860       .QueryDefs.Refresh
19870       Set qdf = Nothing
19880     End If
19890     DoEvents

19900     lngFlds = 0&
19910     ReDim arr_varFld(F_ELEMS, 0)

19920     Set qdf = .QueryDefs(strQryName)
19930     Set rst = qdf.OpenRecordset
19940     With rst
19950       If .BOF = True And .EOF = True Then
              ' ** No record!
19960       Else
              ' ** 1 record.
19970         .MoveFirst
19980         For Each fld In .Fields
19990           With fld
                  ' ** ctl_name_01, xadgfx_image_01.
20000             If Left(.Name, 9) = "ctl_name_" Then
20010               strNum = Right(.Name, 2)
20020               lngFlds = lngFlds + 1&
20030               lngE = lngFlds - 1&
20040               ReDim Preserve arr_varFld(F_ELEMS, lngE)
20050               arr_varFld(F_FNAM, lngE) = .Name
20060               arr_varFld(F_NUM, lngE) = CLng(strNum)
20070               arr_varFld(F_FVAL, lngE) = .Value
20080               arr_varFld(F_IMG, lngE) = "xadgfx_image_" & strNum
                    ' ** tblForm_Graphics.xadgfx_image_01 AS journalno_tgl_off_raised_img
20090               arr_varFld(F_SQL, lngE) = strSource & "." & arr_varFld(F_IMG, lngE) & " " & _
                      "AS " & .Value
                    ' *****************************************************
                    ' ** Array: arr_varFld()
                    ' **
                    ' **   Field  Element  Name                Constant
                    ' **   =====  =======  ==================  ==========
                    ' **     1       0     Field Name          F_FNAM
                    ' **     2       1     Field Number        F_NUM
                    ' **     3       2     Field Value         F_FVAL
                    ' **     4       3     Image Field Name    F_IMG
                    ' **     5       4     SQL Code            F_SQL
                    ' **
                    ' *****************************************************
20100             End If
20110           End With
20120         Next  ' ** fld.
20130         Set fld = Nothing
20140       End If
20150       .Close
20160     End With  ' ** rst
20170     Set rst = Nothing

20180     Debug.Print "'IMG FLDS: " & CStr(lngFlds)
20190     DoEvents

20200     strSQL = qdf.SQL
20210     strTmp01 = strSQL
20220     lngFldNames = 0&
20230     For lngX = 0& To (lngFlds - 1&)
20240       strTmp02 = strSource & "." & arr_varFld(F_IMG, lngX)  ' ** Image field with qualifying source table.
20250       intPos01 = InStr(strTmp01, strTmp02)
20260       If intPos01 > 0 Then
20270         strTmp03 = Left(strTmp01, (intPos01 - 1))  ' ** All SQL up to, but not including, this image field.
20280         intPos03 = InStr(strTmp01, "FROM ")
20290         intPos02 = InStr(intPos01, strTmp01, ",")
20300         If intPos02 = 0 Then
                ' ** Last field.
20310           intPos02 = InStr(intPos01, strTmp01, " ")  ' ** Space following this image field.
20320           If intPos02 = 0 Then
20330             Stop  ' ** There should always be a space somewhere!
20340           Else
20350             If intPos02 > intPos03 Then  ' ** Next space is after 'FROM' clause.
20360               intPos02 = InStr(intPos01, strTmp01, vbCr)  ' ** Carriage Return following this image field.
20370               If intPos02 > 0 Then
20380                 If intPos02 < intPos03 Then
20390                   strTmp04 = Mid(strTmp01, intPos02)  ' ** From CR following this image field to end of SQL.
20400                   strTmp01 = strTmp03 & arr_varFld(F_SQL, lngX) & strTmp04  ' ** All SQL, with new control name.
20410                 Else
20420                   Stop  ' ** If not a comma and not a space, what else could follow it?
20430                 End If
20440               Else
20450                 Stop  ' ** If not a comma, and not a space, and not a CR, there's nothing else?
20460               End If
20470             Else
20480               strTmp04 = Mid(strTmp01, intPos02)  ' ** From space following this image field to end of SQL.
20490               strTmp01 = strTmp03 & arr_varFld(F_SQL, lngX) & strTmp04  ' ** All SQL, with new control name.
20500             End If
20510           End If
20520         Else
20530           If intPos02 < intPos03 Then
                  ' ** Next field
20540             strTmp04 = Mid(strTmp01, intPos02)  ' ** From comma following this image field to end of SQL.
20550             strTmp01 = strTmp03 & arr_varFld(F_SQL, lngX) & strTmp04  ' ** All SQL, with new control name.
20560           Else
                  ' ** Last field before 'FROM' clause (if there are commas further on in the SQL.
20570             intPos02 = InStr(intPos01, strTmp01, " ")  ' ** Space following this image field.
20580             If intPos02 > 0 Then
20590               If intPos02 < intPos03 Then
20600                 strTmp04 = Mid(strTmp01, intPos02)  ' ** From space following this image field to end of SQL.
20610                 strTmp01 = strTmp03 & arr_varFld(F_SQL, lngX) & strTmp04  ' ** All SQL, with new control name.
20620               Else
20630                 intPos02 = InStr(intPos01, strTmp01, vbCr)  ' ** Carriage Return following this image field.
20640                 If intPos02 > 0 Then
20650                   If intPos02 < intPos03 Then
20660                     strTmp04 = Mid(strTmp01, intPos02)  ' ** From CR following this image field to end of SQL.
20670                     strTmp01 = strTmp03 & arr_varFld(F_SQL, lngX) & strTmp04  ' ** All SQL, with new control name.
20680                   Else
20690                     Stop  ' ** If not a comma and not a space, what else could follow it?
20700                   End If
20710                 Else
20720                   Stop  ' ** If not a comma, and not a space, and not a CR, there's nothing else?
20730                 End If
20740               End If
20750             Else
20760               Stop  ' ** There should always be a space somewhere!
20770             End If
20780           End If
20790         End If
20800       Else
20810         Stop  ' ** Image field not found!
20820       End If
20830       lngFldNames = lngFldNames + 1&
20840     Next  ' ** lngX.

          ' ** New SQL should be assembled in strTmp01.
20850     qdf.SQL = strTmp01

20860     .Close
20870   End With  ' ** dbs.
20880   Set dbs = Nothing

20890   Debug.Print "'FLDS NAMED: " & CStr(lngFldNames)
20900   DoEvents

20910   Beep

20920   Debug.Print "'DONE!"
20930   DoEvents

        'DONE!
        'DONE!
        'IMG FLDS: 52
        'FLDS NAMED: 52
        'DONE!
        'IMG FLDS: 42
        'FLDS NAMED: 42
        'DONE!

EXITP:
20940   Set fld = Nothing
20950   Set rst = Nothing
20960   Set qdf = Nothing
20970   Set dbs = Nothing
20980   XAdmin_FrmGfx_Qry = blnRetVal
20990   Exit Function

ERRH:
21000   blnRetVal = False
21010   Select Case ERR.Number
        Case Else
21020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21030   End Select
21040   Resume EXITP

End Function

Public Function XAdmin_FrmGfx_Ctl() As Boolean

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "XAdmin_FrmGfx_Ctl"

        Dim frm As Access.Form, ctl As Access.Control
        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, fld As DAO.Field
        Dim lngCtls As Long, arr_varCtl() As Variant
        Dim blnSkip As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long, lngY As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        Const C_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const C_CNAM1 As Integer = 0
        Const C_CNAM2 As Integer = 1

21110 On Error GoTo 0

21120   blnRetVal = True

21130   lngCtls = 0&
21140   ReDim arr_varCtl(C_ELEMS, 0)

21150   Set frm = Forms(0)
21160   With frm

21170     For Each ctl In .Detail.Controls
21180       With ctl
21190         If Left(.Name, 4) = "Text" Then
                ' ** Text402.
21200           lngCtls = lngCtls + 1&
21210           lngE = lngCtls - 1&
21220           ReDim Preserve arr_varCtl(C_ELEMS, lngE)
21230           arr_varCtl(C_CNAM1, lngE) = .Name
21240           arr_varCtl(C_CNAM2, lngE) = Null
21250         End If
21260       End With
21270     Next  ' ** ctl.
21280     Set ctl = Nothing

21290     Debug.Print "'CTLS: " & CStr(lngCtls)
21300     DoEvents

21310     If lngCtls > 0& Then

21320       strTmp01 = "journalno"  ' ** Already named.
21330       Set dbs = CurrentDb
21340       With dbs
21350         Set qdf = .QueryDefs("qryTransaction_Audit_05_06")
21360         Set rst = qdf.OpenRecordset
21370         With rst
21380           .MoveFirst
21390           lngY = 0&
21400           For Each fld In .Fields
21410             With fld
21420               If Right(.Name, 4) = "_img" Or Right(.Name, 8) = "_img_dis" Then
21430                 If Left(.Name, Len(strTmp01)) <> strTmp01 Then
21440                   lngY = lngY + 1&
21450                   If IsNull(arr_varCtl(C_CNAM2, (lngY - 1&))) = True Then
21460                     arr_varCtl(C_CNAM2, (lngY - 1&)) = .Name
21470                   End If
21480                 End If
21490               End If
21500             End With  ' ** fld.
21510           Next  ' ** fld.
21520           .Close
21530         End With  ' ** rst.
21540         Set rst = Nothing
21550         Set qdf = Nothing
21560         .Close
21570       End With  ' ** dbs.
21580       Set dbs = Nothing

            'For lngX = 0& To (lngCtls - 1&)
            '  Debug.Print "'" & arr_varCtl(C_CNAM2, lngX)
            '  If (lngX + 1&) Mod 100 = 0 Then
            '    Stop
            '  End If
            'Next

21590       blnSkip = False
21600       If blnSkip = False Then
21610         lngY = 0&
21620         strTmp02 = vbNullString
21630         For lngX = 0& To (lngCtls - 1&)
21640           If IsNull(arr_varCtl(C_CNAM2, lngX)) = False Then
21650             strTmp03 = arr_varCtl(C_CNAM2, lngX)
21660             intPos01 = InStr(strTmp03, "_tgl")
21670             If intPos01 > 0 Then
21680               strTmp03 = Left(strTmp03, (intPos01 - 1))
21690               If strTmp03 <> strTmp02 Then
21700                 If lngY = 0& Or lngY = 9& Then
21710                   strTmp02 = strTmp03
21720                   lngY = 1&
21730                 Else
21740                   Stop
21750                 End If
21760               Else
21770                 lngY = lngY + 1&
21780               End If
21790             Else
21800               Stop
21810             End If
21820           Else
21830             Stop
21840           End If
21850         Next  ' ** lngX.
21860       End If  ' ** blnSkip.

21870       For lngX = 0& To (lngCtls - 1&)
21880         Set ctl = .Controls(arr_varCtl(C_CNAM1, lngX))
21890         With ctl
21900           .Name = arr_varCtl(C_CNAM2, lngX)
21910           .ControlSource = arr_varCtl(C_CNAM2, lngX)
21920         End With
21930       Next  ' ** lngX

21940     End If  ' ** lngCtls.

21950   End With  ' ** frm.

21960   Beep

21970   Debug.Print "'DONE!"
21980   DoEvents

EXITP:
21990   Set ctl = Nothing
22000   Set frm = Nothing
22010   Set fld = Nothing
22020   Set rst = Nothing
22030   Set qdf = Nothing
22040   Set dbs = Nothing
22050   XAdmin_FrmGfx_Ctl = blnRetVal
22060   Exit Function

ERRH:
22070   blnRetVal = False
22080   Select Case ERR.Number
        Case Else
22090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22100   End Select
22110   Resume EXITP

End Function

Public Function Frm_Ctl_Prop_Set() As Boolean

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "Frm_Ctl_Prop_Set"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, frm As Access.Form, ctl As Access.Control
        Dim lngCtls As Long, arr_varCtl As Variant
        Dim strFrmName As String
        Dim lngX As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        'Const C_DID  As Integer = 0
        'Const C_FID  As Integer = 1
        Const C_FNAM As Integer = 2
        'Const C_CID  As Integer = 3
        Const C_CNAM As Integer = 4
        'Const C_TYP  As Integer = 5
        'Const C_DEF  As Integer = 6
        'Const C_VIS  As Integer = 7
        'Const C_TAG  As Integer = 8
        'Const C_NEW  As Integer = 9
        'Const C_TOP  As Integer = 6
        'Const C_TX   As Integer = 7
        'Const C_LFT  As Integer = 8
        Const C_LFTN As Integer = 9
        'Const C_PIX  As Integer = 10
        'Const C_PIXN As Integer = 11
        Const C_LX   As Integer = 12
        'Const C_WDT  As Integer = 13
        'Const C_WX   As Integer = 14
        'Const C_HGT  As Integer = 15
        'Const C_HX   As Integer = 16

22210 On Error GoTo 0

22220   blnRetVal = True

22230   Set dbs = CurrentDb
22240   With dbs
          ' ** zzz_qry_zForm_Control_01_05 (zzz_qry_zForm_Control_01_01 (tblForm_Control, just Trust.mdb,
          ' ** 'acCheckBox'), not in zzz_qry_zForm_Control_01_02 (tblPreference_Control, just Trust.mdb,
          ' ** 'dbBoolean')), just Visible = False, with ctlspec_tag_new.
          'Set qdf = .QueryDefs("zzz_qry_zForm_Control_01_06_01")
          ' ** zzz_qry_zForm_Control_01_05 (zzz_qry_zForm_Control_01_01 (tblForm_Control, just Trust.mdb,
          ' ** 'acCheckBox'), not in zzz_qry_zForm_Control_01_02 (tblPreference_Control, just Trust.mdb,
          ' ** 'dbBoolean'), just Visible = True, with ctlspec_tag_new.
          'Set qdf = .QueryDefs("zzz_qry_zForm_Control_01_07_01")
          ' ** zzz_qry_zForm_Control_01_01 (tblForm_Control, just Trust.mdb, 'acCheckBox'), just bad DefaultValue.
          'Set qdf = .QueryDefs("zzz_qry_zForm_Control_01_08")
          ' ** zzz_qry_zForm_Control_02_01 (tblForm_Control, just Trust.mdb, 'acOptionGroup'), not in
          ' ** zzz_qry_zForm_Control_02_02 (tblPreference_Control, just Trust.mdb, 'opg..'), with ctlspec_tag_new.
          'Set qdf = .QueryDefs("zzz_qry_zForm_Control_02_05")

          ' ** zzz_qry_zForm_Control_04_09, just Top, Left, Width, Height discrepancies.
22250     Set qdf = .QueryDefs("zzz_qry_zForm_Control_04_10")

22260     Set rst = qdf.OpenRecordset
22270     With rst
22280       .MoveLast
22290       lngCtls = .RecordCount
22300       .MoveFirst
22310       arr_varCtl = .GetRows(lngCtls)
            ' *********************************************************
            ' ** Array: arr_varCtl()
            ' **
            ' **   Field  Element  Name                    Constant
            ' **   =====  =======  ======================  ==========
            ' **     1       0     dbs_id                  C_DID
            ' **     2       1     frm_id                  C_FID
            ' **     3       2     frm_name                C_FNAM
            ' **     4       3     ctl_id                  C_CID
            ' **     5       4     ctl_name                C_CNAM
            ' **     6       5     ctltype_type            C_TYP
            ' **     7       6     ctlspec_defaultvalue    C_DEF
            ' **     8       7     ctlspec_visible         C_VIS
            ' **     9       8     ctlspec_tag             C_TAG
            ' **    10       9     ctlspec_tag_new         C_NEW
            ' **
            ' *********************************************************
            ' *********************************************************
            ' ** Array: arr_varCtl()
            ' **
            ' **   Field  Element  Name                    Constant
            ' **   =====  =======  ======================  ==========
            ' **     1       0     dbs_id                  C_DID
            ' **     2       1     frm_id                  C_FID
            ' **     3       2     frm_name                C_FNAM
            ' **     4       3     ctl_id                  C_CID
            ' **     5       4     ctl_name                C_CNAM
            ' **     6       5     ctltype_type            C_TYP
            ' **     7       6     ctlspec_top             C_TOP
            ' **     8       7     Tx                      C_TX
            ' **     9       8     ctlspec_left            C_LFT
            ' **    10       9     ctlspec_left_new        C_LFTN
            ' **    11      10     Left_Pix                C_PIX
            ' **    12      11     Left_Pix_new            C_PIXN
            ' **    13      12     Lx                      C_LX
            ' **    14      13     ctlspec_width           C_WDT
            ' **    15      14     Wx                      C_WX
            ' **    16      15     ctlspec_height          C_HGT
            ' **    17      16     Hx                      C_HX
            ' **
            ' *********************************************************
22320       .Close
22330     End With
22340     Set rst = Nothing
22350     Set qdf = Nothing
22360     .Close
22370   End With
22380   Set dbs = Nothing

22390   Debug.Print "'CTLS: " & CStr(lngCtls)
22400   DoEvents

22410   If lngCtls > 0& Then
22420     strFrmName = vbNullString
22430     For lngX = 0& To (lngCtls - 1&)
22440       If arr_varCtl(C_FNAM, lngX) <> strFrmName Then
22450         If strFrmName <> vbNullString Then
22460           DoCmd.Close acForm, strFrmName, acSaveYes
22470         End If
22480         strFrmName = arr_varCtl(C_FNAM, lngX)
22490         DoCmd.OpenForm strFrmName, acDesign, , , , acHidden
22500         Set frm = Forms(strFrmName)
22510       End If
22520       With frm
22530         Set ctl = .Controls(arr_varCtl(C_CNAM, lngX))
22540         If arr_varCtl(C_LX, lngX) = "X" Then
22550           ctl.Left = arr_varCtl(C_LFTN, lngX)
22560         End If
              'ctl.Tag = arr_varCtl(C_NEW, lngX)
              'ctl.DefaultValue = "False"
22570       End With
22580     Next
22590     Set frm = Nothing
22600     Set ctl = Nothing
22610     DoCmd.Close acForm, strFrmName, acSaveYes
22620   End If

22630   Beep

22640   Debug.Print "'DONE!"
22650   DoEvents

EXITP:
22660   Set frm = Nothing
22670   Set ctl = Nothing
22680   Set rst = Nothing
22690   Set qdf = Nothing
22700   Set dbs = Nothing
22710   Frm_Ctl_Prop_Set = blnRetVal
22720   Exit Function

ERRH:
22730   blnRetVal = False
22740   Select Case ERR.Number
        Case Else
22750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22760   End Select
22770   Resume EXITP

End Function

Public Function Frm_Ctl_Prop_List() As Boolean

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "Frm_Ctl_Prop_List"

        Dim frmPar As Access.Form, frmSub As Access.Form, ctl As Access.Control
        Dim lngCtls As Long, arr_varCtl As Variant
        Dim strFrmName1 As String, strFrmName2 As String
        Dim blnLocked As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varCtl().
        Const C_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        Const C_CNAM As Integer = 0
        Const C_LOCK As Integer = 1
        Const C_VIS  As Integer = 2

22810 On Error GoTo 0

22820   blnRetVal = True

22830   strFrmName1 = "frmAccountProfile_Add"
22840   strFrmName2 = "frmAccountProfile_Add_Sub"

22850   lngCtls = 0&
22860   ReDim arr_varCtl(C_ELEMS, 0)

22870   Set frmPar = Forms(strFrmName1)
22880   With frmPar
22890     Set frmSub = .Controls(strFrmName2).Form
22900     With frmSub
22910       For Each ctl In .Controls
22920         blnLocked = False
22930         With ctl
22940 On Error Resume Next
22950           blnLocked = .Locked
22960           If ERR.Number = 0 Then
22970 On Error GoTo 0
22980             If blnLocked = True Then
22990               lngCtls = lngCtls + 1&
23000               lngE = lngCtls - 1&
23010               ReDim Preserve arr_varCtl(C_ELEMS, lngE)
23020               arr_varCtl(C_CNAM, lngE) = .Name
23030               arr_varCtl(C_LOCK, lngE) = .Locked
23040               arr_varCtl(C_VIS, lngE) = .Visible
23050             End If
23060           Else
23070 On Error GoTo 0
23080           End If
23090         End With  ' ** ctl.
23100       Next  ' ** ctl.
23110     End With  ' ** frmSub.
23120   End With  ' ** frmPar.

23130   Debug.Print "'CTLS: " & CStr(lngCtls)
23140   DoEvents

23150   If lngCtls > 0& Then

23160     Debug.Print "'VIS:"
23170     DoEvents
23180     For lngX = 0& To (lngCtls - 1&)
23190       If arr_varCtl(C_VIS, lngX) = True Then
23200         Debug.Print "'  " & arr_varCtl(C_CNAM, lngX)
23210         DoEvents
23220       End If
23230     Next  ' ** lngX.

23240     Debug.Print "'NOT VIS:"
23250     DoEvents
23260     For lngX = 0& To (lngCtls - 1&)
23270       If arr_varCtl(C_VIS, lngX) = False Then
23280         Debug.Print "'  " & arr_varCtl(C_CNAM, lngX)
23290         DoEvents
23300       End If
23310     Next  ' ** lngX.

23320   End If  ' ** lngCtls.

23330   Beep

23340   Debug.Print "'DONE!"
23350   DoEvents

        'CTLS: 29
        'VIS:
        '  account_id
        '  description
        '  accountno2
        '  shortname2
        '  FocusHolder

        'NOT VIS:
        '  related_accountno_lbl2
        '  related_accountno_lbl4
        '  taxlot
        '  feeFrequency
        '  alphasort
        '  reviewfreq
        '  statementfreq
        '  birthdate
        '  fdictype
        '  statetype
        '  assetno_display
        '  icash
        '  pcash
        '  cost
        '  predate
        '  preicash
        '  prepcash
        '  preasset
        '  account_SWEEP
        '  ActiveAssets
        '  numCopies
        '  S_PQuotes
        '  L_PQuotes
        '  Acct_State_Pref
        'DONE!

EXITP:
23360   Set frmPar = Nothing
23370   Set frmSub = Nothing
23380   Set ctl = Nothing
23390   Frm_Ctl_Prop_List = blnRetVal
23400   Exit Function

ERRH:
23410   blnRetVal = False
23420   Select Case ERR.Number
        Case Else
23430     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23440   End Select
23450   Resume EXITP

End Function

Public Function Frm_Ctl_Lbl_Set() As Boolean

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "Frm_Ctl_Lbl_Set"

        Dim frm As Access.Form, ctl1 As Access.Control, ctl2 As Access.Control
        Dim blnRetVal As Boolean

23510 On Error GoTo 0

23520   blnRetVal = True

23530   Set frm = Forms(0)
23540   With frm

23550     For Each ctl1 In .Detail.Controls
23560       With ctl1
23570         Select Case .ControlType
              Case acLabel
23580           If Left(.Name, 5) = "Label" Then
23590             .Name = .Parent.Name & "_lbl"
                  '  If Left(.Parent.Name, 3) = "ckg" Then
                  '    'Set ctl2 = .Controls(0)
                  '    'ctl2.Name = .Name & "_lbl"
                  '  End If
23600           End If
23610         Case acCheckBox
23620           If Left(.Name, 5) = "Label" Then
23630             .Name = .Parent.Name & "_lbl"
                  '.Height = .Height - 30&
                  '.StatusBarText = vbNullString
                  '.ControlTipText = vbNullString
23640           End If
23650         End Select
23660       End With  ' ** ctl.
23670     Next  ' ** ctl.

23680   End With

23690   Beep

23700   Debug.Print "'DONE!"
23710   DoEvents

EXITP:
23720   Set ctl1 = Nothing
23730   Set ctl2 = Nothing
23740   Set frm = Nothing
23750   Frm_Ctl_Lbl_Set = blnRetVal
23760   Exit Function

ERRH:
23770   Select Case ERR.Number
        Case Else
23780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23790   End Select
23800   Resume EXITP

End Function

Public Function Tbl_Link_Chk() As Boolean
' ** Check linked tables against m_TBL and tblDatabase_Table_Link.

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_Link_Chk"

        Dim dbs As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim lngTbls As Long, arr_varTbl() As Variant
        Dim lngDbsID As Long
        Dim lngThisDbsID As Long, strThisDbsName As String
        Dim strDbsName As String, strConnect As String
        Dim blnFound As Boolean
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varTbl().
        Const T_ELEMS As Integer = 5  ' ** Array's first-element UBound().
        Const T_DID_AO As Integer = 0
        Const T_DID    As Integer = 1
        Const T_DNAM   As Integer = 2
        Const T_TID    As Integer = 3
        Const T_TNAM   As Integer = 4
        Const T_FND    As Integer = 5

23910   blnRetVal = True

23920   strDbsName = "TrustAux.mdb"
23930   lngDbsID = DLookup("[dbs_id]", "tblDatabase", "[dbs_name] = '" & strDbsName & "'")

23940   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
23950   DoEvents

23960   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
23970   strThisDbsName = CurrentAppName  ' ** Module Function: modFileUtilities.

23980   lngTbls = 0&
23990   ReDim arr_varTbl(T_ELEMS, 0)

24000   Set dbs = CurrentDb
24010   With dbs

24020     For Each tdf In .TableDefs
24030       With tdf
24040         strConnect = .Connect
24050         If strConnect <> vbNullString Then
24060           If InStr(strConnect, strDbsName) > 0 Then
24070             lngTbls = lngTbls + 1&
24080             lngE = lngTbls - 1&
24090             ReDim Preserve arr_varTbl(T_ELEMS, lngE)
24100             arr_varTbl(T_DID_AO, lngE) = lngThisDbsID
24110             arr_varTbl(T_DID, lngE) = lngDbsID
24120             arr_varTbl(T_DNAM, lngE) = strDbsName
24130             arr_varTbl(T_TID, lngE) = Null
24140             arr_varTbl(T_TNAM, lngE) = .Name
24150             arr_varTbl(T_FND, lngE) = CBool(False)
24160           End If
24170         End If
24180       End With
24190     Next
24200     Set tdf = Nothing

24210     Debug.Print "'TBLS: " & CStr(lngTbls)
24220     DoEvents

24230     If lngTbls > 0& Then

24240       Set rst = .OpenRecordset("tblDatabase_Table", dbOpenDynaset, dbReadOnly)
24250       With rst
24260         .MoveFirst
24270         For lngX = 0& To (lngTbls - 1&)
24280           blnFound = False
24290           .FindFirst "[dbs_id] = " & CStr(arr_varTbl(T_DID, lngX)) & " And [tbl_name] = '" & arr_varTbl(T_TNAM, lngX) & "'"
24300           If .NoMatch = False Then
24310             arr_varTbl(T_TID, lngX) = ![tbl_id]
24320           Else
24330             Debug.Print "'TBL NOT FOUND!  " & arr_varTbl(T_TNAM, lngX)
24340             DoEvents
24350           End If
24360         Next  ' ** lngX.
24370         .Close
24380       End With  ' ** rst.
24390       Set rst = Nothing

24400     End If

24410     .Close
24420   End With  ' ** dbs.
24430   Set dbs = Nothing

24440   Beep

24450   Debug.Print "'DONE!"
24460   DoEvents

EXITP:
24470   Set rst = Nothing
24480   Set tdf = Nothing
24490   Set dbs = Nothing
24500   Tbl_Link_Chk = blnRetVal
24510   Exit Function

ERRH:
24520   Select Case ERR.Number
        Case Else
24530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24540   End Select
24550   Resume EXITP

End Function

Public Function Tbl_Link_Add() As Boolean
' ** No queries were abused in the making of this link.

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "Tbl_Link_Add"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strTableName As String, strConnect As String, strAutoNum As String
        Dim strPath As String, strFile As String, strPathFile As String
        Dim lngThisDbsID As Long, lngThatDbsID As Long
        Dim lngTblID As Long, lngTblLinkID As Long, lngMTblID As Long, lngFldID As Long
        Dim blnNewRecs As Boolean
        Dim varTmp00 As Variant
        Dim blnRetVal As Boolean

24610 On Error GoTo 0

24620   blnRetVal = True

24630   strTableName = "tblTreeViewRelationship"
24640   blnNewRecs = False
        '"tblTreeView_Icon"
        '"tblTreeViewRelationship"
        '"tblCheckVoid"
        '"tblCheckBank"
        '"tblVBComponent_Monitor"
        '"tblTransaction_Audit_Filter"
        '"tblVBComponent_Shortcut_New"
        '"tblVBComponent_Shortcut_Form"
        '"tblForm_Shortcut_Form_Detail"
        '"tblForm_Shortcut_Form"
        '"tblForm_Shortcut_Publish"

24650   DoCmd.Hourglass True
24660   DoEvents

24670   Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.

24680   lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.

24690   Set dbs = CurrentDb
24700   With dbs

24710     If TableExists(strTableName) = True Then  ' ** Module Function: modFileUtilities.

24720       Set rst = .OpenRecordset("tblDatabase_Table_Link", dbOpenDynaset, dbReadOnly)
24730       With rst
24740         .FindFirst "[dbs_id_asof] = " & CStr(lngThisDbsID) & " And [tbllnk_name] = '" & strTableName & "'"
24750         Select Case .NoMatch
              Case True
                ' ** Proceed.
24760         Case False
24770           blnRetVal = False
24780           Beep
24790           Debug.Print "'TBL ALREADY LISTED!  " & strTableName
24800           DoEvents
24810         End Select
24820         .Close
24830       End With
24840       Set rst = Nothing

24850       If blnRetVal = True Then
24860         Set rst = .OpenRecordset("m_TBL", dbOpenDynaset, dbReadOnly)
24870         With rst
24880           .FindFirst "[mtbl_NAME] = '" & strTableName & "'"
24890           Select Case .NoMatch
                Case True
                  ' ** Proceed.
24900           Case False
24910             blnRetVal = False
24920             Beep
24930             Debug.Print "'TBL ALREADY LISTED!  " & strTableName
24940             DoEvents
24950           End Select
24960           .Close
24970         End With
24980         Set rst = Nothing
24990       End If  ' ** blnRetVal.

25000       If blnRetVal = True Then
25010         strConnect = .TableDefs(strTableName).Connect
25020         strFile = Parse_File(strConnect)  ' ** Module Function: modFileUtilities.
25030         Set rst = .OpenRecordset("tblDatabase", dbOpenDynaset, dbReadOnly)
25040         With rst
25050           .FindFirst "[dbs_name] = '" & strFile & "'"
25060           Select Case .NoMatch
                Case True
25070             blnRetVal = False
25080             Beep
25090             Debug.Print "'FILE NOT FOUND IN tblDatabase!  " & strFile
25100             DoEvents
25110           Case False
25120             lngThatDbsID = ![dbs_id]
25130             strPath = ![dbs_path]
25140             If strFile = gstrFile_DataName Or strFile = gstrFile_ArchDataName Or strFile = gstrFile_AuxDataName Then
                    ' ** OK to link.
25150             Else
25160               blnRetVal = False
25170               Beep
25180               Debug.Print "'DB NOT SUPPORTED!  " & strFile
25190               DoEvents
25200             End If
25210           End Select
25220           .Close
25230         End With
25240         Set rst = Nothing
25250       End If  ' ** blnRetVal.

25260       If blnRetVal = True Then
25270         Set rst = .OpenRecordset("tblDatabase_Table", dbOpenDynaset, dbReadOnly)
25280         With rst
25290           .FindFirst "[dbs_id] = " & CStr(lngThatDbsID) & " And [tbl_name] = '" & strTableName & "'"
25300           Select Case .NoMatch
                Case True
25310             blnRetVal = False
25320             Beep
25330             Debug.Print "'TBL NOT FOUND IN tblDatabase_Table!  " & strTableName
25340             DoEvents
25350           Case False
25360             lngTblID = ![tbl_id]
25370           End Select
25380           .Close
25390         End With
25400         Set rst = Nothing
25410       End If  ' ** blnRetVal.

25420       If blnRetVal = True Then
25430         Set rst = .OpenRecordset("tblDatabase_AutoNumber", dbOpenDynaset, dbReadOnly)
25440         With rst
25450           .FindFirst "[dbs_id] = " & CStr(lngThatDbsID) & " And [tbl_id] = " & CStr(lngTblID)
25460           Select Case .NoMatch
                Case True
25470             Beep
25480             Debug.Print "'NO AUTONUM FIELD FOUND!"
25490           Case False
25500             lngFldID = ![fld_id]
25510             varTmp00 = DLookup("[fld_name]", "tblDatabase_Table_Field", "[dbs_id] = " & CStr(lngThatDbsID) & " And " & _
                    "[fld_id] = " & CStr(lngFldID))
25520             If IsNull(varTmp00) = False Then
25530               strAutoNum = varTmp00
25540             End If
25550           End Select
25560           .Close
25570         End With
25580         Set rst = Nothing
25590         varTmp00 = Empty
25600       End If  ' ** blnRetVal.

25610       If blnRetVal = True Then
25620         strPathFile = strPath & LNK_SEP & strFile  ' ** Evidently this isn't needed!
25630         Set rst = .OpenRecordset("tblDatabase_Table_Link", dbOpenDynaset, dbConsistent)
25640         With rst
25650           .AddNew
25660           ![dbs_id] = lngThatDbsID
25670           ![tbl_id] = lngTblID
                ' ** ![tbllnk_id] : AutoNumber.
25680           ![dbs_id_asof] = lngThisDbsID
25690           ![tbllnk_name] = strTableName
25700           ![tbllnk_sourcetablename] = strTableName
25710           ![contype_type] = dbCJet
25720           ![tbllnk_connect] = strConnect
25730           ![tbllnk_versionadded] = "v" & AppVersion_Get2  ' ** Module Function: modAppVersionFuncs.
25740           ![tbllnk_datemodified] = Now()
25750           .Update
25760           .Bookmark = .LastModified
25770           lngTblLinkID = ![tbllnk_id]
25780           .Close
25790         End With
25800         Set rst = Nothing
25810         varTmp00 = DMax("[mtbl_ORDER]", "m_TBL")
25820         Set rst = .OpenRecordset("m_TBL", dbOpenDynaset, dbConsistent)
25830         With rst
25840           .AddNew
                ' ** ![mtbl_ID] : AutoNumber.
25850           ![mtbl_NAME] = strTableName
25860           ![mtbl_AUTONUMBER] = strAutoNum
25870           ![mtbl_ORDER] = varTmp00 + 10
25880           ![mtbl_NEWRecs] = blnNewRecs
25890           ![mtbl_ACTIVE] = True
25900           Select Case strFile
                Case gstrFile_DataName
25910             ![mtbl_DTA] = True
25920           Case gstrFile_ArchDataName
25930             ![mtbl_ARCH] = True
25940           Case gstrFile_AuxDataName
25950             ![mtbl_AUX] = True
25960           End Select
25970           .Update
25980           .Bookmark = .LastModified
25990           lngMTblID = ![mtbl_ID]
26000           .Close
26010         End With
26020         Set rst = Nothing
26030       End If  ' ** blnRetVal.

26040       If blnRetVal = True Then
              ' ** We're going to assume these queries find the new table.
              ' ** Append zzz_qry_Database_Table_Link_10_01 (m_TBL, not in
              ' ** in tblTemplate_m_TBL, by mtbl_ID) to tblTemplate_m_TBL.
26050         Set qdf = .QueryDefs("zzz_qry_Database_Table_Link_10_05")
26060         qdf.Execute
26070         Set qdf = Nothing
              ' ** Append zzz_qry_Database_Table_Link_13_01 (zzz_qry_Database_Table_Link_11_01
              ' ** (tblDatabase_Table_Link, just dbs_id_asof = 1), not in
              ' ** zzz_qry_Database_Table_Link_11_02 (tblTemplate_Database_Table_Link, just
              ' ** dbs_id_asof = 1), by tbllnk_id) to tblTemplate_Database_Table_Link.
26080         Set qdf = .QueryDefs("zzz_qry_Database_Table_Link_13_07")
26090         qdf.Execute
26100         Set qdf = Nothing
26110       End If  ' ** blnRetVal.

26120       If blnRetVal = True Then
26130         .TableDefs.Refresh
26140         DoEvents
26150         Debug.Print "'SUCCESSFULLY ADDED TO LINK TABLES!  " & strTableName
26160         DoEvents
26170       End If  ' ** blnRetVal.

26180     Else
26190       Beep
26200       Debug.Print "'TBL NOT LINKED!  " & strTableName
26210       DoEvents
26220     End If

26230     .Close
26240   End With
26250   Set dbs = Nothing

26260   DoCmd.Hourglass False

26270   Beep

26280   Debug.Print "'DONE!"
26290   DoEvents

EXITP:
26300   Set rst = Nothing
26310   Set qdf = Nothing
26320   Set dbs = Nothing
26330   Tbl_Link_Add = blnRetVal
26340   Exit Function

ERRH:
26350   DoCmd.Hourglass False
26360   blnRetVal = False
26370   Select Case ERR.Number
        Case Else
26380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26390   End Select
26400   Resume EXITP

End Function

Public Function Frm_Ctl_Chk() As Boolean

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "Frm_Ctl_Chk"

        Dim frm As Access.Form, ctl As Access.Control
        Dim lngForm_Width As Long, lngLeft As Long, lngWidth As Long
        Dim blnRetVal As Boolean

26510 On Error GoTo 0

26520   blnRetVal = True

        ' ** Width of form in design view.
26530   lngForm_Width = 14850&

26540   Set frm = Forms(0)
26550   With frm
26560     For Each ctl In .Controls
26570       lngLeft = 0&: lngWidth = 0&
26580       With ctl
26590 On Error Resume Next
26600         lngLeft = .Left
26610 On Error GoTo 0
26620         If lngLeft > 0& Then
26630           lngWidth = .Width
26640           If lngLeft + lngWidth > lngForm_Width Then
26650             Debug.Print "'" & .Name
26660           End If
26670         End If
26680       End With
26690     Next
26700   End With

        'Forms(0).ProgBar_box.Width = 2880  '5760
        'Forms(0).ProgBar_box.Height = 120  '195

        'Forms(1).ProgBar_box.Width = 4080
        'Forms(1).ProgBar_box.Height = 195

        'Forms(2).ProgBar_box.Width = 2880
        'Forms(2).ProgBar_box.Height = 120

        'Forms(3).ProgBar_box.Width = 2880
        'Forms(3).ProgBar_box.Height = 120

        'Forms(4).ProgBar_box.Width = 2880
        'Forms(4).ProgBar_box.Height = 120

        'Forms(5).ProgBar_box.Width = 7215
        'Forms(5).ProgBar_box.Height = 195

26710   Beep

26720   Debug.Print "'DONE!"

EXITP:
26730   Frm_Ctl_Chk = blnRetVal
26740   Exit Function

ERRH:
26750   blnRetVal = False
26760   Select Case ERR.Number
        Case Else
26770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26780   End Select
26790   Resume EXITP

End Function

Public Function GetCodeLine(strModName As String, lngLineNum As Long) As Variant

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "GetCodeLine"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim varRetVal As Variant

26810   varRetVal = Null

26820   Set vbp = Application.VBE.ActiveVBProject
26830   With vbp
26840     Set vbc = .VBComponents(strModName)
26850     With vbc
26860       Set cod = .CodeModule
26870       With cod
26880         varRetVal = .Lines(lngLineNum, 1)
26890       End With
26900     End With
26910   End With  ' ** vbp.

EXITP:
26920   Set cod = Nothing
26930   Set vbc = Nothing
26940   Set vbp = Nothing
26950   GetCodeLine = varRetVal
26960   Exit Function

ERRH:
26970   varRetVal = Null
26980   Select Case ERR.Number
        Case Else
26990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27000   End Select
27010   Resume EXITP

End Function

Public Function VBA_THIS_NAME() As Boolean
' ** Check for 'THIS_NAME' use in modules.
' ** This is to assure that procedures moved to standard modules
' ** because of the Project Name Table don't identify their parent
' ** form by the module name, instead of the original source form.

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_THIS_NAME"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim strModName As String, strProcName As String, strLine As String
        Dim lngLines As Long, lngDecLines As Long
        Dim lngItems As Long, arr_varItem() As Variant
        Dim strFind As String
        Dim intPos01 As Integer, intPos02 As Integer, intLen As Integer
        Dim lngX As Long, lngE As Long
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

        ' ** Array: arr_varItem().
        Const I_ELEMS As Integer = 3  ' ** Array's first-element UBound().
        Const I_VNAM As Integer = 0
        Const I_PNAM As Integer = 1
        Const I_LIN  As Integer = 2
        Const I_TXT  As Integer = 3

27110 On Error GoTo 0

27120   blnRetVal = True

27130   Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.
27140   DoEvents

27150   strFind = "THIS_NAME"
27160   intLen = Len(strFind)

27170   lngItems = 0&
27180   ReDim arr_varItem(I_ELEMS, 0)

27190   Set vbp = Application.VBE.ActiveVBProject
27200   With vbp
27210     For Each vbc In .VBComponents
27220       With vbc
27230         If .Type = vbext_ct_StdModule Then
27240           strModName = .Name
27250           If Left(strModName, 2) <> "zz" And strModName <> "modBackendUpdate" Then
27260             Set cod = .CodeModule
27270             With cod
27280               lngLines = .CountOfLines
27290               lngDecLines = .CountOfDeclarationLines
27300               For lngX = lngDecLines To lngLines
27310                 strLine = Trim(.Lines(lngX, 1))
27320                 strProcName = vbNullString
27330                 If strLine <> vbNullString Then
27340                   If Left(strLine, 1) <> "'" Then
27350                     intPos01 = InStr(strLine, strFind)
27360                     If intPos01 > 0 Then
27370                       intPos02 = InStr(strLine, "zErrorHandler")
27380                       If intPos02 = 0 Then
27390                         intPos02 = InStr(strLine, "zErrorWriteRecord")
27400                         If intPos02 = 0 Then
27410                           intPos02 = InStr(strLine, "Module:")
27420                           If intPos02 = 0 Then
27430                             strTmp01 = Mid(strLine, intPos01, intLen)
27440                             If IsUC(strTmp01, True) = True Then  ' ** Module Function: modStringFuncs.
27450                               strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
27460                               If strProcName <> THIS_PROC Then
27470                                 lngItems = lngItems + 1&
27480                                 lngE = lngItems - 1&
27490                                 ReDim Preserve arr_varItem(I_ELEMS, lngE)
27500                                 arr_varItem(I_VNAM, lngE) = strModName
27510                                 arr_varItem(I_PNAM, lngE) = strProcName
27520                                 arr_varItem(I_LIN, lngE) = lngX
27530                                 arr_varItem(I_TXT, lngE) = strLine
27540                               End If  ' ** THIS_PROC.
27550                             End If  ' ** IsUC().
27560                           End If
27570                         End If  ' ** intPos02.
27580                       End If  ' ** intPos02.
27590                     End If  ' ** intPos01.
27600                   End If  ' ** Remark.
27610                 End If  ' ** vbNullString.
27620               Next  ' ** lngX.
27630             End With  ' ** cod.
27640           End If  ' ** zz_.
27650         End If  ' ** Standard Module.
27660       End With  ' ** vbc.
27670     Next  ' ** vbc.
27680   End With  ' ** vbp.

27690   Debug.Print "'HITS: " & CStr(lngItems)
27700   DoEvents

27710   If lngItems > 0& Then
27720     For lngX = 0& To (lngItems - 1&)
27730       Debug.Print "'MOD: " & arr_varItem(I_VNAM, lngX) & "  PROC: " & arr_varItem(I_PNAM, lngX) & "  LINE: " & CStr(arr_varItem(I_LIN, lngX))
27740       Debug.Print "'  " & arr_varItem(I_TXT, lngX)
27750     Next  ' ** lngX.
27760   End If

        'HITS: 8
        'X MOD: modCourtReportsNS  PROC: AssetList_Word_NS  LINE: 2056
        'X   14770       gstrFormQuerySpec = THIS_NAME                                                          'FIXED!
        'X MOD: modCourtReportsNS  PROC: AssetList_Excel_NS  LINE: 2247
        'X   16070       gstrFormQuerySpec = THIS_NAME                                                          'FIXED!
        'X MOD: modGoToReportFuncs  PROC: GTRImageChk  LINE: 455
        'X   2810          If strModName <> THIS_NAME Then                                                      'OK!
        'X MOD: modGoToReportFuncs  PROC: GTRImageChk  LINE: 475
        'X   3010          End If  ' ** THIS_NAME.                                                              'OK!
        'X MOD: modJrnlCol_Forms  PROC: JC_Frm_TaxLot_Form  LINE: 1706
        'X   11200       intRetVal = OpenLotInfoForm(False, THIS_NAME)  ' ** Module Function: modPurchaseSold.  'OK!
        'X MOD: modPurchaseSold  PROC: SaleLotInfo_PS  LINE: 1708
        'X   11630         lngRetVal = OpenLotInfoForm(True, THIS_NAME)  ' ** Function: Above.                  'OK!
        'X MOD: modVersionConvertFuncs1  PROC: Version_Status  LINE: 4693
        'X   26430       DoCmd.OpenForm FRM_CNV_STATUS, , , , , acDialog, THIS_NAME                             'OK!
        'X MOD: modVersionConvertFuncs1  PROC: Version_Status  LINE: 4702
        'X   26510       DoCmd.OpenForm FRM_CNV_STATUS, , , , , acWindowNormal, THIS_NAME                       'OK!
        'DONE!

27770   Beep

27780   Debug.Print "'DONE!"
27790   DoEvents

EXITP:
27800   Set cod = Nothing
27810   Set vbc = Nothing
27820   Set vbp = Nothing
27830   VBA_THIS_NAME = blnRetVal
27840   Exit Function

ERRH:
27850   Select Case ERR.Number
        Case Else
27860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27870   End Select
27880   Resume EXITP

End Function

Public Function VBA_AccountNoDropdown_Qrys() As Boolean
' ** This helps track down accountno sorting to include
' ** expanding length of the number, starting with the
' ** 2-character demo accounts, followed by those with
' ** an increasing number of characters.

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "VBA_AccountNoDropdown_Qrys"

        Dim vbp As VBIDE.VBProject, vbc As VBIDE.VBComponent, cod As VBIDE.CodeModule
        Dim strModName As String, strProcName As String, strLine As String
        Dim lngLines As Long, lngDecLines As Long
        Dim lngItems As Long, arr_varItem() As Variant
        Dim strFind As String
        Dim intPos01 As Integer, intLen As Integer
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long
        Dim varTmp00 As Variant, strTmp01 As String
        Dim blnRetVal As Boolean

        ' ** Array: arr_varItem().
        Const I_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const I_VNAM As Integer = 0
        Const I_PNAM As Integer = 1
        Const I_LIN  As Integer = 2
        Const I_QRY  As Integer = 3
        Const I_TXT  As Integer = 4

27910 On Error GoTo 0

27920   blnRetVal = True

27930   Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.
27940   DoEvents

27950   strFind = "strSortOrig As String = " '& Chr(34) & "[accountno]"  '"qryAccountNoDropDown_"
27960   intLen = Len(strFind)

27970   lngItems = 0&
27980   ReDim arr_varItem(I_ELEMS, 0)

27990   Set vbp = Application.VBE.ActiveVBProject
28000   With vbp
28010     For Each vbc In .VBComponents
28020       With vbc
28030         strModName = .Name
28040         If Left(strModName, 2) <> "zz" Then
28050           Set cod = .CodeModule
28060           With cod
28070             lngLines = .CountOfLines
28080             lngDecLines = .CountOfDeclarationLines
28090             For lngX = 1& To lngDecLines 'lngDecLines To lngLines
28100               strLine = Trim(.Lines(lngX, 1))
28110               strProcName = vbNullString
28120               If strLine <> vbNullString Then
28130                 If Left(strLine, 1) <> "'" Then
28140                   intPos01 = InStr(strLine, strFind)
28150                   If intPos01 > 0 Then
28160                     strTmp01 = Mid(strLine, intPos01)
                          'intPos01 = InStr(strTmp01, Chr(34))
                          'If intPos01 > 0 Then
                          '  strTmp01 = Left(strTmp01, (intPos01 - 1))
                          'End If
28170                     strProcName = .ProcOfLine(lngX, vbext_pk_Proc)
28180                     If strProcName = vbNullString Then strProcName = "Declaration"
28190                     If strProcName <> THIS_PROC Then
28200                       lngItems = lngItems + 1&
28210                       lngE = lngItems - 1&
28220                       ReDim Preserve arr_varItem(I_ELEMS, lngE)
28230                       arr_varItem(I_VNAM, lngE) = strModName
28240                       arr_varItem(I_PNAM, lngE) = strProcName
28250                       arr_varItem(I_LIN, lngE) = lngX
28260                       arr_varItem(I_QRY, lngE) = strTmp01
28270                       arr_varItem(I_TXT, lngE) = strLine
28280                     End If  ' ** THIS_PROC.
28290                   End If  ' ** intPos01.
28300                 End If  ' ** Remark.
28310               End If  ' ** vbNuillString.
28320             Next  ' ** lngX.
28330           End With  ' ** cod.
28340         End If  ' ** zz_.
28350       End With  ' ** vbc.
28360     Next  ' ** vbc.
28370   End With  ' ** vbp

28380   Debug.Print "'HITS: " & CStr(lngItems)
28390   DoEvents

28400   If lngItems > 0& Then

          ' ** Binary Sort arr_varItem() array by query name.
28410     For lngX = UBound(arr_varItem, 2) To 1 Step -1
28420       For lngY = 0 To (lngX - 1)
28430         If arr_varItem(I_QRY, lngY) > arr_varItem(I_QRY, (lngY + 1)) Then
28440           For lngZ = 0& To I_ELEMS
28450             varTmp00 = arr_varItem(lngZ, lngY)
28460             arr_varItem(lngZ, lngY) = arr_varItem(lngZ, (lngY + 1))
28470             arr_varItem(lngZ, (lngY + 1)) = varTmp00
28480             varTmp00 = Empty
28490           Next  ' ** lngZ.
28500         End If
28510       Next  ' ** lngY.
28520     Next  ' ** lngX.

28530     strTmp01 = vbNullString
28540     For lngX = 0& To (lngItems - 1&)
28550       If arr_varItem(I_QRY, lngX) <> strTmp01 Then
28560         strTmp01 = arr_varItem(I_QRY, lngX)
28570         Debug.Print "'" & strTmp01
28580       End If
28590     Next

28600   End If

        ' ** Forms using tmpAccount for their own puposes:
        ' **   frmAccountProfile_Add
        ' **   frmAdminOfficer
        ' ** Forms using tmpAccount for accountno sorting:
        ' **   frmAccountContacts
        ' **   frmAccountProfile
        ' **   frmCheckReconcile
        ' **   frmJournal_Columns
        ' **   frmLocation_Asset
        ' **   frmLocation_Asset_Sub
        ' **   frmMap_Div_Detail_Sub
        ' **   frmMap_Int_Detail_Sub
        ' **   frmMap_Misc_LTCL_Detail_Sub
        ' **   frmMap_Misc_STCGL_Detail_Sub
        ' **   frmMap_Rec
        ' **   frmMap_Rec_Detail_Sub
        ' **   frmMap_Reinvest_DivInt_Detail_Sub
        ' **   frmMap_Reinvest_Rec_Detail_Sub
        ' **   frmMap_Split
        ' **   frmMap_Split_Detail_Sub
        ' **   frmMenu_Account
        ' **   frmPortfolioModeling_Select
        ' **   frmRpt_ArchivedTransactions
        ' **   frmRpt_Checks
        ' **   frmRpt_Checks_Bank2_Sub
        ' **   frmStatementBalance
        ' **   frmTransaction_Audit_Sub
        ' **   frmTransaction_Audit_Sub_ds
        ' ** Forms using tmpXAdmin_Account_02 for account sorting:
        ' **   frmAccountContacts
        ' ** AccountNoDropDown queries don't involve an additional table.

        ' ** Reports using tmpAccount for accountno sorting:
        ' **   rptMap_Dividend
        ' **   rptMap_Interest
        ' **   rptMap_Misc_LTCL
        ' **   rptMap_Misc_STCG
        ' **   rptMap_Misc_STCL
        ' **   rptMap_Received
        ' **   rptMap_Reinvest_Div
        ' **   rptMap_Reinvest_Int
        ' **   rptMap_Reinvest_Rec
        ' **   rptMap_Split
        ' **   rptChecks_Blank
        ' **   rptChecks_Preprinted
        ' **   rptTransaction_Audit_01
        'OTHER cmdPrintReport'S THAT'LL BE LOOKING FOR ALPHASORT?

        'FOR ALL SUBFORMS USING ALPHASORT, CROSS-CHECK FOR REPORTS!
        'CHECK ALL SORT_NOW'S FOR ACCOUNTNO!
        'CHECK ALL REPORT ORDER-BY'S FOR ACCOUNTNO!
        'CHECK ALL REPORT GROUPING FOR ACCOUNTNO!

        '##################
        'HAVEN'T DONE LIST BOX ON FRM CHECKS YET!
        '##################

        'X rptMap_Dividend:
        'X   qryMapReport_01
        'X rptMap_Interest:
        'X   qryMapReport_02
        'X rptMap_Misc_LTCL:
        'X   qryMapReport_07
        'X rptMap_Misc_STCG:
        'X   qryMapReport_08
        'X rptMap_Misc_STCL:
        'X   qryMapReport_09
        'X rptMap_Received:
        'X   qryMapReport_03
        'X rptMap_Reinvest_Div:
        'X   qryMapReport_05
        'X rptMap_Reinvest_Int:
        'X   qryMapReport_05
        'X rptMap_Reinvest_Rec:
        'X   qryMapReport_06
        'X rptMap_Split:
        'X   qryMapReport_04

        'X qryMap_Div_02_04
        'X qryMap_Int_02_04
        'X qryMap_Misc_LTCL_Detail_Sub_01_02
        'X qryMap_Misc_STCGL_Detail_Sub_01_02
        'X qryMap_Rec_02_04
        'X qryMap_Received_02_10
        'X qryMap_Reinvest_02_10
        'X qryMap_Split_02

        'HITS: 38
        'qryAccountNoDropDown_03
        'qryAccountNoDropDown_04
        'DONE!

        ' ** All default subform sorts:
        ' **   strSortOrig As String = "[account_sort]"
        ' **   strSortOrig As String = "[alphasort], [assettype], [asset_description], [assetdate]"
        ' **   strSortOrig As String = "[alphasort], [balance_date] DESC"
        ' **   strSortOrig As String = "[accounttype]"
        ' **   strSortOrig As String = "[accounttypegroup_sequence], [accounttype]"
        ' **   strSortOrig As String = "[alphasort]"
        ' **   strSortOrig As String = "[alphasort], [Contact_Number]"
        ' **   strSortOrig As String = "[assetdate]"
        ' **   strSortOrig As String = "[assettype]"
        ' **   strSortOrig As String = "[assettype], [description_masterasset_sort], [cusip]"
        ' **   strSortOrig As String = "[assettype], [totdesc], [cusip]"
        ' **   strSortOrig As String = "[assettypegroup_sequence], [assettype]"
        ' **   strSortOrig As String = "[chkbank_name], [chkvoid_chknum]"
        ' **   strSortOrig As String = "[ChkMemo_Memo]"
        ' **   strSortOrig As String = "[chkvoid_chknum]"
        ' **   strSortOrig As String = "[country_name]"
        ' **   strSortOrig As String = "[country_name_sort]"
        ' **   strSortOrig As String = "[curr_code], [curr_date] DESC"
        ' **   strSortOrig As String = "[curr_date] DESC, [curr_code]"
        ' **   strSortOrig As String = "[curr_name]"
        ' **   strSortOrig As String = "[curr_name_sort]"
        ' **   strSortOrig As String = "[curr_word1], [country_name_sort]"
        ' **   strSortOrig As String = "[curracct_sort]"
        ' **   strSortOrig As String = "[description_masterasset]"
        ' **   strSortOrig As String = "[ErrLog_Date] DESC"
        ' **   strSortOrig As String = "[fsfd_section], [fsfd_order], [fsfd_level], [keydowntype_order], [fsp_keycode]"
        ' **   strSortOrig As String = "[invobj_id]"
        ' **   strSortOrig As String = "[Journal_ID]"
        ' **   strSortOrig As String = "[journalno]"
        ' **   strSortOrig As String = "[JournalType_Order]"
        ' **   strSortOrig As String = "[ledghid_grpnum], [ledghid_ord]"
        ' **   strSortOrig As String = "[Loc_Name], [Loc_Address1]"
        ' **   strSortOrig As String = "[officer]"
        ' **   strSortOrig As String = "[pp_id] DESC"
        ' **   strSortOrig As String = "[Recur_Name], [Recur_Address]"
        ' **   strSortOrig As String = "[revcode_TYPE], [revcode_SORTORDER]"
        ' **   strSortOrig As String = "[scheddets_order]"
        ' **   strSortOrig As String = "[Schedule_ID]"
        ' **   strSortOrig As String = "[SpecSort2], [JrnlCol_ID]"
        ' **   strSortOrig As String = "[state_military] DESC, [state_canada] DESC, [state_territory] DESC, [state_multi], [state_name]"
        ' **   strSortOrig As String = "[taxcode_type], [taxcode_order]"
        ' **   strSortOrig As String = "[tmpDate]"
        ' **   strSortOrig As String = "[transdate] DESC, [croutchk_id]"
        ' **   strSortOrig As String = "[transdate] DESC, [journaltype], [journalno] DESC"
        ' **   strSortOrig As String = "[transdate] DESC, [JournalType_Order], [journalno]"
        ' **   strSortOrig As String = "[transdate], [JournalType_Order], [totdesc]"
        ' **   strSortOrig As String = "[Username]"
        ' **   strSortOrig As String = "qryPrintChecks_01_01a_Account_Number"

28610   Beep

28620   Debug.Print "'DONE!"
28630   DoEvents

EXITP:
28640   Set cod = Nothing
28650   Set vbc = Nothing
28660   Set vbp = Nothing
28670   VBA_AccountNoDropdown_Qrys = blnRetVal
28680   Exit Function

ERRH:
28690   Select Case ERR.Number
        Case Else
28700     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28710   End Select
28720   Resume EXITP

End Function

Public Function TblLink_Chk() As Boolean
' ** Compare tables currently linked with m_TBL.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "TblLink_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, tdf As DAO.TableDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnFound As Boolean
        Dim blnRetVal As Boolean

110   On Error GoTo 0

120     blnRetVal = True

130     Win_Mod_Restore  ' ** Module Procedure: modWindowsFuncs.
140     DoEvents

150     Set dbs = CurrentDb
160     With dbs

170       Set rst = .OpenRecordset("m_TBL", dbOpenDynaset, dbReadOnly)

180       lngRecs = 0&
190       For Each tdf In .TableDefs
200         With tdf
210           If .Connect <> vbNullString Then
220             blnFound = False
230             With rst
240               .FindFirst "[mtbl_NAME] = '" & tdf.Name & "'"
250               If .NoMatch = False Then
260                 blnFound = True
270               End If
280             End With
290             If blnFound = False Then
300               Debug.Print "'LINK NOT LISTED!  " & .Name
310               DoEvents
320               lngRecs = lngRecs + 1&
330             End If
340           End If
350         End With
360       Next

370       .Close
380     End With

390     If lngRecs > 0& Then
400       Debug.Print "'TBLS NOT IN m_TBL: " & CStr(lngRecs)
410     Else
420       Debug.Print "'ALL LINKED TBLS LISTED!"
430     End If
440     DoEvents

'ALL LINKED TBLS LISTED!
'DONE!

'LINK NOT LISTED!  tblTreeView_Icon
'LINK NOT LISTED!  tblTreeViewRelationship
'TBLS NOT IN m_TBL: 2
'DONE!

450     Beep
460     Debug.Print "'DONE!"
470     DoEvents

EXITP:
480     Set tdf = Nothing
490     Set rst = Nothing
500     Set qdf = Nothing
510     Set dbs = Nothing
520     TblLink_Chk = blnRetVal
530     Exit Function

ERRH:
540     blnRetVal = False
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Function
