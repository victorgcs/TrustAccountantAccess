Attribute VB_Name = "modErrorHandler"
Option Compare Database
Option Explicit

'VGC 04/04/2016: CHANGES!

Private Const THIS_NAME As String = "modErrorHandler"
' **

Public Function zErrorHandler(strModuleName As String, strFunctionName As String, Optional varErrNum As Variant, Optional varLineNum As Variant, Optional varErrDesc As Variant) As Boolean
' ** If a reference is made to a Missing parameter, the error generated is:
' ** 'Error 448': Named argument not found.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "zErrorHandler"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, fld As DAO.Field
        Dim rstErrLog As DAO.Recordset
        Dim strErrMsg As String, strErrDesc As String
        Dim dblErrNum As Double, lngLineNum As Long
        Dim blnFound As Boolean
        Dim blnRetVal As Boolean

110     blnRetVal = True
120     strErrDesc = vbNullString

130     If IsMissing(varErrNum) = False Then
140       If IsNull(varErrNum) = False Then
150         If IsNumeric(varErrNum) = True Then
160           dblErrNum = varErrNum
170         Else
180           dblErrNum = ERR.Number
190         End If
200       Else
210         dblErrNum = ERR.Number
220       End If
230     Else
          ' ** If it wasn't sent, the error object may be still active.
240       dblErrNum = ERR.Number
250     End If

260     If IsMissing(varErrDesc) = False Then
270       strErrDesc = varErrDesc
280     Else
290       strErrDesc = vbNullString
300     End If

310     If IsMissing(varLineNum) = False Then
320       If varLineNum > 0& Then
330         lngLineNum = varLineNum
340       Else
350         lngLineNum = 0&
360       End If
370     Else
380       lngLineNum = 0&
390     End If

400     If dblErrNum <> 0 Then

410       Set dbs = CurrentDb
420       Select Case TableExists("tblErrorLog")  ' ** Module Function: modFileUtilities.
          Case True
430   On Error Resume Next
440         Set rstErrLog = dbs.OpenRecordset("tblErrorLog", dbOpenDynaset, dbConsistent)
450         If ERR.Number <> 0 Then
460   On Error GoTo ERRH
470           If TableExists("tblErrorLog_tmp") = False Then  ' ** Module Function: modFileUtilities.
480             DoCmd.CopyObject , "tblErrorLog_tmp", acTable, "tblTemplate_ErrorLog"
490             CurrentDb.TableDefs.Refresh
500             CurrentDb.TableDefs.Refresh
510           End If
520           Set rstErrLog = dbs.OpenRecordset("tblErrorLog_tmp", dbOpenDynaset, dbConsistent)
530         Else
540   On Error GoTo ERRH
550         End If
560       Case False
570         If TableExists("tblErrorLog_tmp") = False Then  ' ** Module Function: modFileUtilities.
580           DoCmd.CopyObject , "tblErrorLog_tmp", acTable, "tblTemplate_ErrorLog"
590           DoEvents
600           CurrentDb.TableDefs.Refresh
610           CurrentDb.TableDefs.Refresh
620           DoEvents
630         End If
640         Set rstErrLog = dbs.OpenRecordset("tblErrorLog_tmp", dbOpenDynaset, dbConsistent)
650       End Select

660       Select Case dblErrNum
          Case 2046  ' ** The command or action '|' isn't available now.
            ' ** Ignore.
            ' ** I'd like to see the 'Close action was canceled' so I can fix it, or at least deal with it there.
670       Case 2486  ' ** You can't carry out this action at the present time.
            ' ** Ignore.
            ' ** Always in Form_Unload() event. Not sure what it objects to.
680       Case 2585  ' ** This action can't be carried out while processing a form or report event.
            ' ** Ignore.
            ' ** I believe this is an Access 2007 error regarding our TA Print Ribbon.
690       Case Else

700         If strErrDesc = vbNullString Then
710           blnFound = False
720           If dblErrNum <> 0 Then
730             With dbs
                  ' ** tblAccessAndJetErrors, by specified [errnum].
740               Set qdf = .QueryDefs("qryErrLog_02")
750               With qdf.Parameters
760                 ![errnum] = dblErrNum
770               End With
780               Set rst = qdf.OpenRecordset()
790               With rst
800                 If rst.EOF = True And rst.BOF = True Then
                      ' ** Error number not found in our table.
810                 Else
820                   strErrDesc = rst![ErrorString]
830                   blnFound = True
840                 End If
850                 .Close
860               End With  ' ** rst.
870             End With  ' ** dbs.
880           End If
890         Else
900           blnFound = True
910         End If

920         If blnFound = False Then
              ' ** Check the other two available sources.
930   On Error Resume Next
940           strErrDesc = AccessError(dblErrNum)  ' ** Normally, always returns something, but if it's way out of whack, it may error.
950   On Error GoTo ERRH
960           If strErrDesc = vbNullString Or strErrDesc = "Application-defined or object-defined error" Or _
                  strErrDesc = "Unknown Error" Then
970   On Error Resume Next
980             strErrDesc = Error(dblErrNum)
990   On Error GoTo ERRH
1000            If strErrDesc = vbNullString Or strErrDesc = "Application-defined or object-defined error" Or _
                    strErrDesc = "Unknown Error" Then
                  ' ** If it wasn't found, the error object may be still active.
                  ' ** NO IT WON'T! IT'S LOST AT THE FIRST 'On Error' IT HITS AFTER THE ERROR.
1010              strErrDesc = ERR.description
1020              If strErrDesc = vbNullString Then
1030                strErrDesc = "Unknown Error"
1040              End If
1050            End If
1060          End If
1070        End If  ' ** blnFound.

1080        strErrDesc = zErrorMsgClean(strErrDesc)  ' ** Function: Below.
1090        blnRetVal = zErrorWriteRecord(dblErrNum, strErrDesc, strModuleName, strFunctionName, lngLineNum, rstErrLog)  ' ** Function: Below.

1100        strErrMsg = vbNullString
1110        strErrMsg = strErrMsg & "Error:" & vbTab & vbTab & CStr(dblErrNum) & vbCrLf
1120        strErrMsg = strErrMsg & "Description:" & vbTab & strErrDesc & vbCrLf & vbCrLf
1130        strErrMsg = strErrMsg & "Module:" & vbTab & vbTab & strModuleName & vbCrLf
1140        strErrMsg = strErrMsg & "Sub/Function:" & vbTab & strFunctionName & "()" & vbCrLf
1150        strErrMsg = strErrMsg & "Line:" & vbTab & vbTab & CStr(varLineNum)

1160        If dblErrNum = 2501 Then  ' ** The '|' action was Canceled.
              ' ** Skip the message, do record it.
1170        Else
1180          MsgBox strErrMsg, vbCritical + vbOKOnly, "Error"
1190        End If

1200      End Select

1210      rstErrLog.Close
1220      dbs.Close

1230    End If  ' ** dblErrNum.

EXITP:
1240    Set fld = Nothing
1250    Set rst = Nothing
1260    Set rstErrLog = Nothing
1270    Set qdf = Nothing
1280    Set dbs = Nothing
1290    Exit Function

ERRH:
1300    Select Case ERR.Number
        Case Else
1310      strErrMsg = vbNullString
1320      If strModuleName <> vbNullString And strFunctionName <> vbNullString Then
1330        If IsMissing(varErrNum) = False Then
1340          strErrMsg = vbCrLf & vbCrLf & "From: " & vbCrLf & "Error: " & CStr(varErrNum)
1350          If IsMissing(varErrDesc) = False Then
1360            strErrMsg = strErrMsg & vbCrLf & varErrDesc
1370          End If
1380          strErrMsg = strErrMsg & vbCrLf & "Module: " & strModuleName & vbCrLf & "Function: " & strFunctionName & "()"
1390          If IsMissing(varLineNum) = False Then
1400            strErrMsg = strErrMsg & vbCrLf & "Line: " & CStr(varLineNum)
1410          End If
1420        Else
1430          strErrMsg = vbCrLf & vbCrLf & "From: " & vbCrLf & _
                "Module: " & strModuleName & vbCrLf & "Function: " & strFunctionName & "()"
1440        End If
1450      End If
1460      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl & strErrMsg, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1470    End Select
1480    Resume EXITP

End Function

Public Function zErrorWriteRecord(dblErrNum As Double, strErrDesc As String, strModuleName As String, strFunctionName As String, lngLineNum As Long, rstErrLog As DAO.Recordset) As Boolean

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "zErrorWriteRecord"

        Dim fld As DAO.Field
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

1510    blnRetVal = True

1520  On Error Resume Next
1530    strTmp01 = rstErrLog.Fields(0).Name  ' ** Just to see if the Recordset's valid.
1540    If ERR.Number = 3343 Then  ' ** Unrecognized database format '...\TrustDta.mdb'.
1550      blnRetVal = False
1560  On Error GoTo ERRH
1570      Beep
1580      MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer can read your data file!" & _
            vbCrLf & vbCrLf & "Restore your data from a backup outside of Trust Accountant.", _
            vbCritical + vbOKOnly, "Trust Accountant Data Corrupted"
1590    ElseIf ERR.Number = 3024 Then  ' ** Could not find file '...\TrustDta.mdb'.
1600      blnRetVal = False
1610  On Error GoTo ERRH
1620      Beep
1630      MsgBox "A serious error has occurred, and Trust Accountant can" & vbCrLf & "no longer find your data file!" & _
            vbCrLf & vbCrLf & "Check to see that it's still there, and/or" & vbCrLf & "relink Trust Accountant to the correct location.", _
            vbCritical + vbOKOnly, "Trust Accountant Data Not Found"
1640    ElseIf ERR.Number <> 0 Then
1650      blnRetVal = False
1660      Beep
1670      MsgBox "A serious error has occurred." & _
            vbCrLf & vbCrLf & ("Error: " & CStr(ERR.Number)) & vbCrLf & "Description: " & ERR.description, _
            vbCritical + vbOKOnly, "Trust Accountant Data Error"
1680  On Error GoTo ERRH
1690    Else
1700  On Error GoTo ERRH
1710      With rstErrLog
1720        .AddNew
1730        ![ErrLog_Date] = Now()
1740        ![ErrLog_Form] = IIf(strModuleName = vbNullString, "Unknown", strModuleName)
1750        ![ErrLog_FuncSub] = IIf(strFunctionName = vbNullString, "Unknown", strFunctionName)
1760        For Each fld In .Fields
1770          If fld.Name = "ErrLog_LineNum" Then
1780            ![ErrLog_LineNum] = lngLineNum
1790            Exit For
1800          End If
1810        Next
1820        ![ErrLog_ErrNum] = dblErrNum
1830        If strErrDesc <> vbNullString Then
1840          ![ErrLog_Message] = strErrDesc
1850        Else
1860          ![ErrLog_Message] = "{no description}"
1870        End If
1880        .Update
1890      End With
1900    End If

EXITP:
1910    Set fld = Nothing
1920    zErrorWriteRecord = blnRetVal
1930    Exit Function

ERRH:
1940    blnRetVal = False
1950    Select Case ERR.Number
        Case Else
1960      MsgBox ("Error: " & CStr(ERR.Number)) & vbCrLf & ERR.description & vbCrLf & vbCrLf & _
            "Module: " & THIS_NAME & vbCrLf & "Function: " & THIS_PROC & "()" & vbCrLf & "Line: " & Erl, _
            vbCritical + vbOKOnly, ("Error: " & CStr(ERR.Number))
1970    End Select
1980    Resume EXITP

End Function

Public Function zErrorMsgClean(varMsg As Variant, Optional varRpt As Variant, Optional varErrNum As Variant) As Variant
' **  Clean up error message.

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "zErrorMsgClean"

        Dim blnRpt As Boolean
        Dim intPass As Integer
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, intTmp02 As Integer, lngTmp03 As Long
        Dim varRetVal As Variant

        Const THIS_APP As String = "Trust Accountant™"
        Const THIS_PROG As String = "Microsoft Access"
        Const MY_CRLF As String = "{{CRLF}}"  '{{CRLF}}

2010    varRetVal = Null

2020    If IsNull(varMsg) = False Then

2030      intPass = 0
2040      If IsMissing(varErrNum) = False Then
2050        blnRpt = False
2060        lngTmp03 = varErrNum
2070      Else
2080        If IsMissing(varRpt) = True Then
2090          blnRpt = False
2100        Else
2110          If varRpt = -300 Then
2120            intPass = 3
2130            varRpt = -1
2140            blnRpt = CBool(varRpt)
2150          Else
2160            blnRpt = CBool(varRpt)
2170          End If
2180        End If
2190      End If

2200      varRetVal = varMsg

2210      intPos01 = InStr(varRetVal, THIS_APP)
2220      If intPos01 > 0 Then
2230        If intPos01 = 1 Then
              ' ** At beginning of Err.Description.
2240          varRetVal = THIS_PROG & Mid(varRetVal, (Len(THIS_APP) + 1))
2250        Else
2260          If intPos01 + Len(THIS_APP) < Len(varRetVal) Then
2270            varRetVal = Left(varRetVal, (intPos01 - 1)) & THIS_PROG & Mid(varRetVal, (intPos01 + Len(THIS_APP)))
2280          Else
                ' ** At end of Err.Description.
2290            varRetVal = Left(varRetVal, (intPos01 - 1)) & THIS_PROG
2300          End If
2310        End If
2320      End If

2330      If intPass = 3 Then

2340        If Right(varRetVal, Len(vbCrLf & "1" & vbCrLf & "1")) = (vbCrLf & "1" & vbCrLf & "1") Then
2350          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "1" & vbCrLf & "1")))
2360        End If

2370        intPos01 = InStr(varRetVal, (vbCrLf & "2" & vbCrLf))
2380        If intPos01 > 0 Then
2390          varRetVal = Left(varRetVal, (intPos01 - 1))
2400        End If

2410        intPos01 = InStr(varRetVal, (vbCrLf & "5" & vbCrLf))
2420        If intPos01 > 0 Then
2430          varRetVal = Left(varRetVal, (intPos01 - 1))
2440        End If

2450        intPos01 = InStr(varRetVal, (vbCrLf & "13" & vbCrLf))
2460        If intPos01 > 0 Then
2470          varRetVal = Left(varRetVal, (intPos01 - 1))
2480        End If

2490        intPos01 = InStr(varRetVal, (vbCrLf & "17" & vbCrLf))
2500        If intPos01 > 0 Then
2510          varRetVal = Left(varRetVal, (intPos01 - 1))
2520        End If

2530        intPos01 = InStr(varRetVal, (vbCrLf & "19" & vbCrLf))
2540        If intPos01 > 0 Then
2550          varRetVal = Left(varRetVal, (intPos01 - 1))
2560        End If

2570        intPos01 = InStr(varRetVal, (vbCrLf & "20" & vbCrLf))
2580        If intPos01 > 0 Then
2590          varRetVal = Left(varRetVal, (intPos01 - 1))
2600        End If

2610        intPos01 = InStr(varRetVal, (vbCrLf & "21" & vbCrLf))
2620        If intPos01 > 0 Then
2630          varRetVal = Left(varRetVal, (intPos01 - 1))
2640        End If

2650        intPos01 = InStr(varRetVal, (vbCrLf & "22" & vbCrLf))
2660        If intPos01 > 0 Then
2670          varRetVal = Left(varRetVal, (intPos01 - 1))
2680        End If

2690        intPos01 = InStr(varRetVal, (vbCrLf & "23" & vbCrLf))
2700        If intPos01 > 0 Then
2710          varRetVal = Left(varRetVal, (intPos01 - 1))
2720        End If

2730        If Right(varRetVal, Len(vbCrLf & "1")) = (vbCrLf & "1") Then
2740          intPos01 = InStr(varRetVal, (vbCrLf & "1"))
2750          intPos02 = InStr((intPos01 + 1), varRetVal, (vbCrLf & "1"))
2760          If intPos01 > 0 And intPos02 > 0 And intPos01 <> intPos02 Then
2770            varRetVal = Left(varRetVal, (intPos01 - 1))
2780          Else
2790            varRetVal = Left(varRetVal, (intPos01 - 1))
2800          End If
2810        End If

2820        If Right(varRetVal, Len(vbCrLf & "4" & vbCrLf & "2")) = (vbCrLf & "4" & vbCrLf & "2") Then
2830          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "4" & vbCrLf & "2")))
2840        End If

2850        If Right(varRetVal, Len(vbCrLf & "3")) = (vbCrLf & "3") Then
2860          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "3")))
2870        End If

2880        If Right(varRetVal, Len(vbCrLf & "1")) = (vbCrLf & "1") Then
2890          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "1")))
2900        End If

2910        If (varRetVal = (vbCrLf & "1" & vbCrLf & "209029")) Or (varRetVal = "**********") Or _
                (varRetVal = "|") Then
2920          varRetVal = Null
2930        End If

2940        If Left(varRetVal, 3) = ("|" & vbCrLf) Then
2950          varRetVal = "| " & Mid(varRetVal, 4)
2960        End If

2970        intTmp02 = 0
2980        If Right(varRetVal, Len(vbCrLf)) = vbCrLf Then
2990          intTmp02 = -1
3000          Do While intTmp02 = -1
3010            varRetVal = Left(varRetVal, (Len(varRetVal) - 2))
3020            If Right(varRetVal, Len(vbCrLf)) <> vbCrLf Then
3030              intTmp02 = 0
3040            End If
3050          Loop
3060        End If

3070        If Right(varRetVal, Len(vbCrLf & "1")) = (vbCrLf & "1") Then
3080          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "1")))
3090        End If

3100        If Right(varRetVal, Len(vbCrLf & "1" & vbCrLf & "2")) = (vbCrLf & "1" & vbCrLf & "2") Then
3110          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "1" & vbCrLf & "2")))
3120        End If

3130        If Right(varRetVal, Len(vbCrLf & "16" & vbCrLf & "4")) = (vbCrLf & "16" & vbCrLf & "4") Then
3140          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "16" & vbCrLf & "4")))
3150        End If

3160        If Right(varRetVal, Len(vbCrLf & "6" & vbCrLf & "4")) = (vbCrLf & "6" & vbCrLf & "4") Then
3170          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "6" & vbCrLf & "4")))
3180        End If

3190        If Right(varRetVal, Len(vbCrLf & "3" & vbCrLf & "2")) = (vbCrLf & "3" & vbCrLf & "2") Then
3200          varRetVal = Left(varRetVal, (Len(varRetVal) - Len(vbCrLf & "3" & vbCrLf & "2")))
3210        End If

3220        If IsNull(varRetVal) = True Then
3230          varRetVal = "{unknown}"
3240        Else
3250          If Trim(varRetVal) = vbNullString Then
3260            varRetVal = "{unknown}"
3270          End If
3280        End If

3290      Else

3300        If InStr(varRetVal, MY_CRLF) > 0 Then

3310          intPass = 2

              ' ** Check for double carriage return line feeds.
3320          intPos01 = InStr(varRetVal, MY_CRLF & MY_CRLF)
3330          If intPos01 > 0 Then
                ' ** Replace 2 with 1.
3340            Do While intPos01 > 0
3350              varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + Len(MY_CRLF)))
3360              intPos01 = InStr(varRetVal, MY_CRLF & MY_CRLF)
3370            Loop
3380          End If

              ' ** Now replace the temporary token with a real carriage return line feed.
3390          intPos01 = InStr(varRetVal, MY_CRLF)
3400          Do While intPos01 > 0
3410            varRetVal = Left(varRetVal, (intPos01 - 1)) & vbCrLf & Mid(varRetVal, (intPos01 + Len(MY_CRLF)))
3420            intPos01 = InStr(varRetVal, MY_CRLF)
3430          Loop

3440        Else

3450          intPass = 1

3460          If blnRpt = False Then

3470            intPos01 = InStr(varRetVal, vbCrLf)
3480            If intPos01 > 0 Then
3490              Do While intPos01 > 0
3500                varRetVal = Left(varRetVal, (intPos01 - 1)) & MY_CRLF & Mid(varRetVal, (intPos01 + 2))
3510                intPos01 = InStr(varRetVal, vbCrLf)
3520              Loop
3530            End If

3540            intPos01 = InStr(varRetVal, vbCr)
3550            If intPos01 > 0 Then
3560              Do While intPos01 > 0
3570                varRetVal = Left(varRetVal, (intPos01 - 1)) & MY_CRLF & Mid(varRetVal, (intPos01 + 1))
3580                intPos01 = InStr(varRetVal, vbCr)
3590              Loop
3600            End If

3610            intPos01 = InStr(varRetVal, vbLf)
3620            If intPos01 > 0 Then
3630              Do While intPos01 > 0
3640                varRetVal = Left(varRetVal, (intPos01 - 1)) & MY_CRLF & Mid(varRetVal, (intPos01 + 1))
3650                intPos01 = InStr(varRetVal, vbLf)
3660              Loop
3670            End If

                ' ** Check for vbLf without vbCr.
                'intPos01 = InStr(varRetVal, vbLf)  ' ** Chr(10)
                'If intPos01 > 0 Then
                '  Do While intPos01 > 0
                '    If Mid(varRetVal, (intPos01 - 1), 1) <> vbCr Then
                '      varRetVal = Left(varRetVal, (intPos01 - 1)) & vbCr & Mid(varRetVal, intPos01)
                '      intPos01 = InStr((intPos01 + 2), varRetVal, vbLf)
                '    Else
                '      intPos01 = InStr((intPos01 + 1), varRetVal, vbLf)
                '    End If
                '  Loop
                'End If

                ' ** Check for vbCr without vbLf.
                'intPos01 = InStr(varRetVal, vbCr)  ' ** Chr(13)
                'If intPos01 > 0 Then
                '  Do While intPos01 > 0
                '    If Mid(varRetVal, (intPos01 + 1), 1) <> vbLf Then
                '      varRetVal = Left(varRetVal, intPos01) & vbLf & Mid(varRetVal, (intPos01 + 1))
                '      intPos01 = InStr((intPos01 + 2), varRetVal, vbLf)
                '    Else
                '      intPos01 = InStr((intPos01 + 1), varRetVal, vbLf)
                '    End If
                '  Loop
                'End If

3680          Else
3690            If InStr(varRetVal, " the following error:") = 0 Then
3700              intPos01 = InStr(varRetVal, vbLf)
3710              If intPos01 > 0 Then
3720                varRetVal = Left(varRetVal, (intPos01 - 1))
3730              End If
3740              intPos01 = InStr(varRetVal, vbCr)
3750              If intPos01 > 0 Then
3760                varRetVal = Left(varRetVal, (intPos01 - 1))
3770              End If
3780            Else
3790              intPos01 = InStr(varRetVal, Chr(13))  ' ** Tab.
3800              Do While intPos01 > 0
3810                varRetVal = Left(varRetVal, (intPos01 - 1)) & " " & Mid(varRetVal, (intPos01 + 1))
3820                intPos01 = InStr(varRetVal, Chr(13))
3830              Loop
3840              varRetVal = Trim(varRetVal)
3850              strTmp01 = " the following error:": intTmp02 = 0
3860              intPos01 = InStr(varRetVal, strTmp01)
3870              intPos02 = InStr((intPos01 + Len(strTmp01)), varRetVal, "*")
3880              Do While intPos02 > 0
3890                intTmp02 = intTmp02 + 1
3900                If intTmp02 = 1 Then
3910                  varRetVal = Trim(Left(varRetVal, (intPos01 + Len(strTmp01)))) & " " & CStr(intTmp02) & ". " & _
                        IIf(Asc(Left(Trim(Mid(varRetVal, (intPos02 + 1))), 1)) < 32, _
                        Mid(Trim(Mid(varRetVal, (intPos02 + 1))), 2), Trim(Mid(varRetVal, (intPos02 + 1))))
3920                Else
3930                  varRetVal = IIf(Asc(Right(Trim(Left(varRetVal, (intPos02 - 1))), 1)) < 32, _
                        Left(Trim(Left(varRetVal, (intPos02 - 1))), (Len(Trim(Left(varRetVal, (intPos02 - 1)))) - 1)), _
                        Trim(Left(varRetVal, (intPos02 - 1)))) & " " & CStr(intTmp02) & ". " & _
                        IIf(Asc(Left(Trim(Mid(varRetVal, (intPos02 + 1))), 1)) < 32, _
                        Mid(Trim(Mid(varRetVal, (intPos02 + 1))), 2), Trim(Mid(varRetVal, (intPos02 + 1))))
3940                End If
3950                intPos02 = InStr((intPos01 + Len(strTmp01)), varRetVal, "*")
3960              Loop
                  ' ** Reformatted:
                  ' **   "The expression |2 you entered as the event property setting produced the following error: " & _
                  ' **     "1. The expression may not result in the name of a macro, the name of a user-defined function, or [Event Procedure].
                  ' **     "2. There may have been an error evaluating the function, event, or macro.
                  ' ** Original:
                  ' **   The expression |2 you entered as the event property setting produced the following error: |1.@* The expression may not result in the name of a macro, the name of a user-defined function, or [Event Procedure].
                  ' **   * There may have been an error evaluating the function, event, or macro.@@1@@1
3970            End If
3980          End If

              'intPos01 = InStr(varRetVal, "~~")
              'Do While intPos01 > 0
              '  varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1))
              '  intPos01 = InStr(varRetVal, "~~")
              'Loop

              'intPos01 = InStr(varRetVal, "~")
              'Do While intPos01 > 0
              '  varRetVal = Left(varRetVal, (intPos01 - 1)) & vbCrLf & Mid(varRetVal, (intPos01 + 1))
              '  intPos01 = InStr(varRetVal, "~")
              'Loop

3990          intPos01 = InStr(varRetVal, "@")
4000          Do While intPos01 > 0
4010            varRetVal = Left(varRetVal, (intPos01 - 1)) & MY_CRLF & Mid(varRetVal, (intPos01 + 1))
4020            intPos01 = InStr(varRetVal, "@")
4030          Loop

4040          intPos01 = InStr(varRetVal, Chr(13))  ' ** Tab.
4050          Do While intPos01 > 0
4060            varRetVal = Left(varRetVal, (intPos01 - 1)) & " " & Mid(varRetVal, (intPos01 + 1))
4070            intPos01 = InStr(varRetVal, Chr(13))
4080          Loop
4090          varRetVal = Trim(varRetVal)

4100          If blnRpt = True Then
4110            intPos01 = InStr(varRetVal, " Change the data")
4120            If intPos01 > 0 Then varRetVal = Left(varRetVal, (intPos01 - 1))
4130            intPos01 = InStr(varRetVal, "For example")
4140            If intPos01 > 0 Then varRetVal = Left(varRetVal, (intPos01 - 1))
4150            intPos01 = InStr(varRetVal, "Select an item from the list")
4160            If intPos01 > 0 Then varRetVal = Left(varRetVal, (intPos01 - 1))
4170            intPos01 = InStr(varRetVal, "You may be at the end")
4180            If intPos01 > 0 Then varRetVal = Left(varRetVal, (intPos01 - 1))
4190            intPos01 = InStr(varRetVal, "Try one of the following")
4200            If intPos01 > 0 Then varRetVal = Left(varRetVal, (intPos01 - 1))
4210            intPos01 = InStr(varRetVal, "You used a method")
4220            If intPos01 > 0 Then varRetVal = Left(varRetVal, (intPos01 - 1))
4230            intPos01 = InStr(varRetVal, "*")
4240            If intPos01 > 0 Then varRetVal = Left(varRetVal, (intPos01 - 1))
4250            varRetVal = Trim(varRetVal)
4260          Else
                ' ** Now that it's clean, replace MY_CRLF with the real thing again.

                ' ** Check for double carriage return line feeds.
4270            intPos01 = InStr(varRetVal, MY_CRLF & MY_CRLF)
4280            If intPos01 > 0 Then
                  ' ** Replace 2 with 1.
4290              Do While intPos01 > 0
4300                varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + Len(MY_CRLF)))
4310                intPos01 = InStr(varRetVal, MY_CRLF & MY_CRLF)
4320              Loop
4330            End If

                ' ** Now replace the temporary token with a real carriage return line feed.
4340            intPos01 = InStr(varRetVal, MY_CRLF)
4350            Do While intPos01 > 0
4360              varRetVal = Left(varRetVal, (intPos01 - 1)) & vbCrLf & Mid(varRetVal, (intPos01 + Len(MY_CRLF)))
4370              intPos01 = InStr(varRetVal, MY_CRLF)
4380            Loop

4390          End If

4400          If IsNull(varRetVal) = False Then
4410            If Trim(varRetVal) <> vbNullString Then
4420              If Asc(Right(varRetVal, 1)) < 32 Then
4430                varRetVal = Trim(Left(varRetVal, (Len(varRetVal) - 1)))
4440              Else
4450                If varRetVal = "(unknown)" Then
4460                  varRetVal = "{unknown}"
4470                End If
4480              End If
4490            Else
4500              varRetVal = "{unknown}"
4510            End If
4520          Else
4530            varRetVal = "{unknown}"
4540          End If

4550        End If

4560      End If

4570    End If

EXITP:
4580    zErrorMsgClean = varRetVal
4590    Exit Function

ERRH:
4600    varRetVal = "Error Parsing Message"
4610    Select Case ERR.Number
        Case Else
4620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Procedure: Above.
4630    End Select
4640    Resume EXITP

End Function

Public Sub zErrorLogWriter(strInputData As String)
' ** This procedure logs messages to a text file.
' ** If you send a zero length string, the routine will clear out the existing file.
' ** Called by:
' **   modUtilities
' **     CompareTableStructure()

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "zErrorLogWriter"

        Dim intFileNum As Integer
        Dim strPath As String, strPathFile As String

4710    intFileNum = FreeFile(0)
4720    strPath = CurrentBackendPathFile("LedgerArchive")  ' ** Module Function: modFileUtilities.
4730    strPath = Parse_Path(strPath)  ' ** Module Function: modFileUtilities.
4740    strPathFile = strPath & LNK_SEP & gstrFile_ArchiveLog

4750    If strInputData = vbNullString Then
4760      Open strPathFile For Output As #intFileNum
4770    Else
4780      Open strPathFile For Append As #intFileNum
4790    End If

4800    Write #intFileNum, Now & "  " & strInputData
4810    Close #intFileNum

EXITP:
4820    Exit Sub

ERRH:
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Procedure: Above.
4850    End Select
4860    Resume EXITP

End Sub

Public Function zErrorNumberHoles() As Boolean

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "zErrorNumberHoles"

        Dim dbs As DAO.Database, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngRecs As Long
        Dim lngErrNum As Long
        Dim blnSkip As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

4910    blnRetVal = True

4920    Set dbs = CurrentDb
4930    With dbs

4940      Set rst2 = .OpenRecordset("tblAccessAndJetErrors_05", dbOpenDynaset, dbAppendOnly)
4950      With rst2
4960        For lngX = 31508& To 65535

4970          varTmp00 = AccessError(lngX)
4980          If IsNull(varTmp00) = False Then
4990            If Trim(varTmp00) <> vbNullString Then
5000              If varTmp00 <> "Application-defined or object-defined error" Then
5010                .AddNew
5020                ![error_number] = lngX
5030                ![Error_Function] = 1
5040                ![Error_Description] = varTmp00
5050                .Update
5060              End If
5070            End If
5080          End If

5090          varTmp00 = Error(lngX)
5100          If IsNull(varTmp00) = False Then
5110            If Trim(varTmp00) <> vbNullString Then
5120              If varTmp00 <> "Application-defined or object-defined error" Then
5130                .AddNew
5140                ![error_number] = lngX
5150                ![Error_Function] = 2
5160                ![Error_Description] = varTmp00
5170                .Update
5180              End If
5190            End If
5200          End If

5210        Next
5220        .Close
5230      End With

5240      blnSkip = True
5250      If blnSkip = False Then
5260        Set rst2 = .OpenRecordset("tblAccessAndJetErrors_05", dbOpenDynaset, dbAppendOnly)
5270        Set rst1 = .OpenRecordset("tblAccessAndJetErrors_04", dbOpenDynaset, dbReadOnly)
5280        With rst1
5290          .MoveLast
5300          lngRecs = .RecordCount
5310          .MoveFirst
5320          lngErrNum = 0&
5330          For lngX = 1& To lngRecs
5340            If ![chk] = True Then
5350              If lngErrNum = 0& Then
5360                lngErrNum = ![error_number]
5370              Else
5380                If ![error_number] <> (lngErrNum + 1&) Then
5390                  Do Until lngErrNum = ![error_number]
5400                    With rst2

5410                      varTmp00 = AccessError(lngErrNum + 1&)
5420                      If IsNull(varTmp00) = False Then
5430                        If Trim(varTmp00) <> vbNullString Then
5440                          If varTmp00 <> "Application-defined or object-defined error" Then
5450                            .AddNew
5460                            ![error_number] = lngErrNum + 1&
5470                            ![Error_Function] = 1
5480                            ![Error_Description] = varTmp00
5490                            .Update
5500                          End If
5510                        End If
5520                      End If

5530                      varTmp00 = Error(lngErrNum + 1&)
5540                      If IsNull(varTmp00) = False Then
5550                        If Trim(varTmp00) <> vbNullString Then
5560                          If varTmp00 <> "Application-defined or object-defined error" Then
5570                            .AddNew
5580                            ![error_number] = lngErrNum + 1&
5590                            ![Error_Function] = 2
5600                            ![Error_Description] = varTmp00
5610                            .Update
5620                          End If
5630                        End If
5640                      End If

5650                    End With
5660                    lngErrNum = lngErrNum + 1&
5670                  Loop
5680                Else
5690                  lngErrNum = ![error_number]
5700                End If
5710              End If
5720            End If
5730            If lngX < lngRecs Then .MoveNext
5740          Next
5750          .Close
5760        End With
5770        rst2.Close
5780      End If

5790      .Close
5800    End With

5810    Beep

EXITP:
5820    Set rst1 = Nothing
5830    Set rst2 = Nothing
5840    Set dbs = Nothing
5850    zErrorNumberHoles = blnRetVal
5860    Exit Function

ERRH:
5870    Select Case ERR.Number
        Case Else
5880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Procedure: Above.
5890    End Select
5900    Resume EXITP

End Function

Public Function zErrorDescription(varErrNum As Variant) As String

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "zErrorDescription"

        Dim varTmp00 As Variant
        Dim strRetVal As String

6010    strRetVal = vbNullString

6020    If IsNull(varErrNum) = False Then
6030      If varErrNum <> 0 Then
6040        varTmp00 = DLookup("[ErrorString]", "tblAccessAndJetErrors", "[ErrorCode] = " & CStr(varErrNum))
6050        If IsNull(varTmp00) = False Then
6060          If Trim(varTmp00) <> vbNullString Then
6070            strRetVal = Trim(varTmp00)
6080          Else
6090            Debug.Print "'ERR NOT FOUND! " & CStr(varErrNum)
6100          End If
6110        Else
6120          Debug.Print "'ERR NOT FOUND! " & CStr(varErrNum)
6130        End If
6140      End If
6150    End If

EXITP:
6160    zErrorDescription = strRetVal
6170    Exit Function

ERRH:
6180    strRetVal = RET_ERR
6190    Select Case ERR.Number
        Case Else
6200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Procedure: Above.
6210    End Select
6220    Resume EXITP

End Function
