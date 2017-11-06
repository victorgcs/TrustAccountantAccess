Attribute VB_Name = "modStringFuncs"
Option Compare Database
Option Explicit

'VGC 03/23/2017: CHANGES!

' ** Single-quote, double-quote, printer's quote, printers quote.
' **  '  Chr(39)
' **  "  Chr(34)
' **  ‘  Chr(145)
' **  ’  Chr(146)
' **  “  Chr(147)
' **  ”  Chr(148)

' ** Chr(160) is Arial Hard-Space (hardspace, hard space)!

' ** Terminal font solid block: 'Û'

' ** ¼ ½ ¾  » «

Private Const THIS_NAME As String = "modStringFuncs"
' **

Public Function GetDateLong(varInput As Variant) As Long
' ** Return just the integer portion of a date, i.e., just the days, without the time.

100   On Error GoTo ERRH

        Const THIS_PROC As String = "GetDateLong"

        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim lngRetVal As Long

110     lngRetVal = -1&

120     If IsNull(varInput) = False Then
130       If IsDate(varInput) = True Then
            ' ** Going directly to CLng() will cause it to round
            ' ** the  decimal, possibly giving us the wrong date.
140         If varType(varInput) <> vbDate Then
150           strTmp01 = CStr(CDbl(CDate(varInput)))
160         Else
170           strTmp01 = CStr(CDbl(varInput))
180         End If
190         intPos01 = InStr(strTmp01, ".")
200         If intPos01 = 0 Then
210           lngRetVal = CLng(strTmp01)
220         Else
230           lngRetVal = CLng(Left(strTmp01, (intPos01 - 1)))
240         End If
250       End If
260     End If

EXITP:
270     GetDateLong = lngRetVal
280     Exit Function

ERRH:
290     lngRetVal = -1&
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

End Function

Public Function StringReplace(strString As String, strExistingChar As String, strNewChar As String) As String
' ** Removes strExistingChar from strString and replaces it with strNewChar.

400   On Error GoTo ERRH

        Const THIS_PROC As String = "StringReplace"

        Dim intPos01 As Integer
        Dim intX As Integer
        Dim strRetVal As String

410     strRetVal = vbNullString
420     If Len(strExistingChar) = 1 Then
430       For intX = 1 To Len(strString)
440         If Mid(strString, intX, 1) = strExistingChar Then
450           strRetVal = strRetVal & strNewChar
460         Else
470           strRetVal = strRetVal & Mid(strString, intX, 1)
480         End If
490       Next
500     Else
510       strRetVal = strString
520       intPos01 = InStr(strRetVal, strExistingChar)
530       If intPos01 > 0 Then
540         Do While intPos01 > 0
550           If intPos01 = 1 Then
560             strRetVal = strNewChar & Mid(strRetVal, (Len(strExistingChar) + 1))
570           ElseIf intPos01 = ((Len(strRetVal) - Len(strExistingChar)) + 1) Then
580             strRetVal = Left(strRetVal, (intPos01 - 1)) & strNewChar
590           Else
600             strRetVal = Left(strRetVal, (intPos01 - 1)) & strNewChar & Mid(strRetVal, (intPos01 + Len(strExistingChar)))
610           End If
620           intPos01 = InStr(strRetVal, strExistingChar)
630         Loop
640       End If
650     End If

EXITP:
660     StringReplace = strRetVal
670     Exit Function

ERRH:
680     Select Case ERR.Number
        Case Else
690       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
700     End Select
710     Resume EXITP

End Function

Public Function StringExactMatch(varInput As Variant, varMatch As Variant) As Boolean
' ** Not case-sensitive.

800   On Error GoTo ERRH

        Const THIS_PROC As String = "StringExactMatch"

        Dim intPos01 As Integer, intPos02 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String
        Dim blnRetVal As Boolean

810     blnRetVal = False

820     If IsNull(varInput) = False And IsNull(varMatch) = False Then
830       If Len(Trim(varMatch)) <= Len(Trim(varInput)) Then
840         strTmp01 = Trim(varInput)
850         strTmp02 = Trim(varMatch)
860         intPos01 = InStr(strTmp01, strTmp02)
870         If intPos01 > 0 Then
880           Do While intPos01 > 0
890             intLen = Len(strTmp02)
900             intPos02 = intPos01 + intLen
910             strTmp03 = vbNullString: strTmp04 = vbNullString
920             If intPos01 = 1 Then
930               If intPos02 <= Len(strTmp01) Then
940                 strTmp04 = Mid(strTmp01, intPos02, 1)  ' ** First character after match string.
950                 Select Case strTmp04
                    Case " ", "(", ")", ",", vbCr, vbLf  ' ** Space, parens, comma.
960                   blnRetVal = True
970                 Case "+", "-", "*", "/", "=", "<", ">"
                      ' ** Maybe.
980                   blnRetVal = True
990                 Case "'", Chr(34)
                      ' ** Maybe not.
1000                Case "_", "[", "]", "¹", "."
                      ' ** Definitely not.
1010                Case Else
1020                  If (Asc(strTmp04) >= 65 And Asc(strTmp04) <= 90) Or (Asc(strTmp04) >= 97 And Asc(strTmp04) <= 122) Or _
                          (Asc(strTmp04) >= 48 And Asc(strTmp04) <= 57) Then
                        ' ** Nope, something else.
1030                  Else
1040                    Debug.Print "'1 " & strTmp04 & "  '" & strTmp02 & "'  '" & Rem_CRLF(strTmp01) & "'"
1050                  End If
1060                End Select
1070              Else
1080                blnRetVal = True
1090              End If
1100            Else
1110              strTmp03 = Mid(strTmp01, (intPos01 - 1), 1)  ' ** First character before match string.
1120              Select Case strTmp03
                  Case " ", "(", ")", ",", vbCr, vbLf  ' ** Space, parens, comma.
1130                blnRetVal = True
1140              Case "+", "-", "*", "/", "=", "<", ">"
                    ' ** Maybe.
1150                blnRetVal = True
1160              Case "'", Chr(34)
                    ' ** Maybe not.
1170              Case "_", "[", "]", "¹", "."
                    ' ** Definitely not.
1180              Case Else
1190                If (Asc(strTmp03) >= 65 And Asc(strTmp03) <= 90) Or (Asc(strTmp03) >= 97 And Asc(strTmp03) <= 122) Or _
                        (Asc(strTmp03) >= 48 And Asc(strTmp03) <= 57) Then
                      ' ** Nope, something else.
1200                Else
1210                  Debug.Print "'2 " & strTmp03 & "  '" & strTmp02 & "'  '" & Rem_CRLF(strTmp01) & "'"
1220                End If
1230              End Select
1240              If blnRetVal = True Then
1250                blnRetVal = False
1260                If intPos02 <= Len(strTmp01) Then
1270                  strTmp04 = Mid(strTmp01, intPos02, 1)  ' ** First character after match string.
1280                  Select Case strTmp04
                      Case " ", "(", ")", ",", vbCr, vbLf  ' ** Space, parens, comma.
1290                    blnRetVal = True
1300                  Case "+", "-", "*", "/", "=", "<", ">"
                        ' ** Maybe.
1310                    blnRetVal = True
1320                  Case "'", Chr(34)
                        ' ** Maybe not.
1330                  Case "_", "[", "]", "¹", "."
                        ' ** Definitely not.
1340                  Case Else
1350                    If (Asc(strTmp04) >= 65 And Asc(strTmp04) <= 90) Or (Asc(strTmp04) >= 97 And Asc(strTmp04) <= 122) Or _
                            (Asc(strTmp04) >= 48 And Asc(strTmp04) <= 57) Then
                          ' ** Nope, something else.
1360                    Else
1370                      Debug.Print "'3 " & strTmp04 & "  '" & strTmp02 & "'  '" & Rem_CRLF(strTmp01) & "'"
1380                    End If
1390                  End Select
1400                Else
1410                  blnRetVal = True
1420                End If
1430              End If
1440            End If
1450            If blnRetVal = True Then
1460              Exit Do
1470            Else
1480              intPos01 = InStr((intPos01 + 1), strTmp01, strTmp02)
1490              If intPos01 = 0 Then Exit Do
1500            End If
1510          Loop
1520        End If
1530      End If
1540    End If

EXITP:
1550    StringExactMatch = blnRetVal
1560    Exit Function

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Function

Public Function Compare_StringA_StringB(varA As Variant, strOperator As String, varB As Variant) As Boolean
' ** Compare 2 strings, ASCII-for-ASCII.
' ** For now, strings of different lengths do not get an
' ** ASCII-for-ASCII comparison, so this can't be used for sorting.

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Compare_StringA_StringB"

        Dim strTmpAA As String, strTmpBB As String
        Dim intLenA As Integer, intLenB As Integer
        Dim intX As Integer
        Dim blnRetVal As Boolean

1710    blnRetVal = False

1720    If IsNull(varA) = False And IsNull(varB) = False Then
1730      strTmpAA = varA
1740      strTmpBB = varB
1750      intLenA = Len(strTmpAA)
1760      intLenB = Len(strTmpBB)
1770      Select Case strOperator
          Case "="
1780        blnRetVal = True  ' ** The first mismatch will break it.
1790        If intLenA = intLenB Then
1800          For intX = 1 To intLenA
1810            If Asc(Mid(strTmpAA, intX, 1)) = Asc(Mid(strTmpBB, intX, 1)) Then
                  ' ** Continue the loop.
1820            Else
1830              blnRetVal = False  ' ** The first mismatch will break it.
1840              Exit For
1850            End If
1860          Next
1870        Else
1880          blnRetVal = False
1890        End If
1900      Case "<>"
1910        blnRetVal = False  ' ** The first mismatch will confirm it.
1920        If intLenA = intLenB Then
1930          For intX = 1 To intLenA
1940            If Asc(Mid(strTmpAA, intX, 1)) = Asc(Mid(strTmpBB, intX, 1)) Then
                  ' ** Continue the loop.
1950            Else
1960              blnRetVal = True  ' ** The first mismatch will confirm it.
1970              Exit For
1980            End If
1990          Next
2000        Else
2010          blnRetVal = True
2020        End If
2030      Case ">"
2040        blnRetVal = False
2050        If intLenA > intLenB Then
2060          blnRetVal = True
2070        ElseIf intLenA = intLenB Then
2080          For intX = 1 To intLenA
2090            If Asc(Mid(strTmpAA, intX, 1)) = Asc(Mid(strTmpBB, intX, 1)) Then
                  ' ** Continue the loop, if it gets to the end all matching, the False stands.
2100            Else
2110              If Asc(Mid(strTmpAA, intX, 1)) > Asc(Mid(strTmpBB, intX, 1)) Then
2120                blnRetVal = True  ' ** The first More-Than confirms it.
2130                Exit For
2140              Else
2150                blnRetVal = False  ' ** The first Less-Than breaks it.
2160                Exit For
2170              End If
2180            End If
2190          Next
2200        Else
              ' ** Less-Than can never be More-Than.
2210        End If
2220      Case "<"
2230        blnRetVal = False
2240        If intLenA < intLenB Then
2250          blnRetVal = True
2260        ElseIf intLenA = intLenB Then
2270          For intX = 1 To intLenA
2280            If Asc(Mid(strTmpAA, intX, 1)) = Asc(Mid(strTmpBB, intX, 1)) Then
                  ' ** Continue the loop, if it gets to the end all matching, the False stands.
2290            Else
2300              If Asc(Mid(strTmpAA, intX, 1)) < Asc(Mid(strTmpBB, intX, 1)) Then
2310                blnRetVal = True  ' ** The first Less-Than confirms it.
2320                Exit For
2330              Else
2340                blnRetVal = False  ' ** The first More-Than breaks it.
2350                Exit For
2360              End If
2370            End If
2380          Next
2390        Else
              ' ** More-Than can never be Less-Than.
2400        End If
2410      End Select
2420    End If

EXITP:
2430    Compare_StringA_StringB = blnRetVal
2440    Exit Function

ERRH:
2450    blnRetVal = False
2460    Select Case ERR.Number
        Case Else
2470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2480    End Select
2490    Resume EXITP

End Function

Public Function Compare_DateA_DateB(datA As Date, strOperator As String, datB As Date) As Boolean
' ** Compare the day only, without any minutes or seconds.
' ** The CLng() function ROUNDS!!
' ** Is datA '>' datB?

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Compare_DateA_DateB"

        Dim dblDatA As Double, dblDatB As Double
        Dim strDatA As String, strDatB As String
        Dim lngData As Long, lngDatB As Long
        Dim intPosA As Integer, intPosB As Integer
        Dim blnRetVal As Boolean

2510    blnRetVal = False  ' ** Default to, "No, A is not '>' B".

2520    dblDatA = CDbl(datA)
2530    dblDatB = CDbl(datB)
2540    strDatA = CStr(dblDatA)
2550    strDatB = CStr(dblDatB)
2560    intPosA = InStr(strDatA, ".")
2570    intPosB = InStr(strDatB, ".")
2580    If intPosA > 0 Then strDatA = Left(strDatA, (intPosA - 1))
2590    If intPosB > 0 Then strDatB = Left(strDatB, (intPosB - 1))
2600    lngData = Val(strDatA)
2610    lngDatB = Val(strDatB)
2620    Select Case strOperator
        Case ">"
2630      If lngData > lngDatB Then blnRetVal = True
2640    Case "<"
2650      If lngData < lngDatB Then blnRetVal = True
2660    End Select

EXITP:
2670    Compare_DateA_DateB = blnRetVal
2680    Exit Function

ERRH:
2690    Select Case ERR.Number
        Case Else
2700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2710    End Select
2720    Resume EXITP

End Function

Public Function HasTypePfx(varAcct As Variant, Optional varRetArr As Variant) As Variant

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "HasTypePfx"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRelAccts As Long, arr_varRelAcct() As Variant
        Dim blnRetArr As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim lngX As Long, lngY As Long
        Dim varRetVal As Variant

        Static lngTypes As Long, arr_varType As Variant

        ' ** Array: arr_varType().
        Const AT_TYP As Integer = 0
        'Const AT_LNG As Integer = 1

        ' ** Array: arr_varRelAcct().
        Const RA_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const RA_ACTO As Integer = 0
        Const RA_PFX  As Integer = 1
        Const RS_TYPO As Integer = 2
        Const RA_TYPN As Integer = 3
        Const RA_ACTN As Integer = 4

2810    If IsMissing(varRetArr) = True Then
2820      blnRetArr = False
2830      varRetVal = CBool(False)
2840    Else
2850      blnRetArr = CBool(varRetArr)
2860      If blnRetArr = True Then
2870        lngRelAccts = 0&                  ' ** If it came with varRetArr, but varAcct wasn't a related_account,
2880        ReDim arr_varRelAcct(RA_ELEMS, 0)  ' ** varRetVal will just be assigned a Boolean True, below.
2890      Else
2900        varRetVal = CBool(False)
2910      End If
2920    End If
2930    If IsNull(varAcct) = False Then
2940      strTmp01 = Trim(CStr(varAcct))
2950      If strTmp01 <> vbNullString Then

2960        If lngTypes = 0& Then
2970          Set dbs = CurrentDb
2980          With dbs
                ' ** AccountType, with accounttype_lng.
2990            Set qdf = .QueryDefs("qryCompanyInformation_02")
3000            Set rst = qdf.OpenRecordset
3010            With rst
3020              .MoveLast
3030              lngTypes = .RecordCount
3040              .MoveFirst
3050              arr_varType = .GetRows(lngTypes)
                  ' ****************************************************
                  ' ** Array: arr_varType()
                  ' **
                  ' **   Field  Element  Name               Constant
                  ' **   =====  =======  =================  ==========
                  ' **     1       0     accounttype        AT_TYP
                  ' **     2       1     accounttype_lng    AT_LNG
                  ' **
                  ' ****************************************************
3060              .Close
3070            End With
3080            .Close
3090          End With
3100        End If

3110        intPos01 = InStr(strTmp01, ",")
3120        intPos02 = InStr(strTmp01, ";")  ' ** It's supposed to be a comma!
3130        If intPos01 = 0 And intPos02 = 0 Then

3140          If blnRetArr = True Then
                ' ** It shouldn't have gotten here!
3150            If strTmp01 = "SUSPENSE" Or strTmp01 = "INCOME O/U" Then
3160              arr_varRelAcct(RA_ACTO, 0) = strTmp01
3170              arr_varRelAcct(RA_PFX, 0) = CBool(False)
3180              arr_varRelAcct(RA_ACTN, 0) = "{Remove}"
3190            Else
3200              arr_varRelAcct(RA_ACTO, 0) = RET_ERR
3210            End If
3220            varRetVal = arr_varRelAcct
3230          Else
3240            If Len(strTmp01) >= 5 Then
                  ' ** Minimum 2-char accountno to get prefix, so accounttype-accountno = 5.
3250              If Mid(strTmp01, 3, 1) = "-" Then
3260                strTmp02 = Left(strTmp01, 2)
3270                If IsNumeric(strTmp02) = True Then
3280                  For lngX = 0& To (lngTypes - 1&)
3290                    If arr_varType(AT_TYP, lngX) = strTmp02 Then
3300                      varRetVal = CBool(True)
3310                      Exit For
3320                    End If
3330                  Next
3340                End If
3350              End If
3360            End If
3370          End If

3380        Else
              ' ** It's the Related_Account field that's being checked.

              ' ***************************************************
              ' ** Array: arr_varRelAcct()
              ' **
              ' **   Field  Element  Name              Constant
              ' **   =====  =======  ================= ==========
              ' **     1       0     accountno_old     RA_ACTO
              ' **     2       1     TypePfx (T/F)     RA_PFX
              ' **     3       2     accounttype_old   RA_TYPO
              ' **     4       3     accounttype_new   RA_TYPN
              ' **     5       4     accountno_new     RA_ACTN
              ' **
              ' ***************************************************

3390          If intPos01 = 0 And intPos02 > 0 Then strTmp01 = StringReplace(strTmp01, ";", ",")
3400          If intPos01 > 0 And intPos02 > 0 Then strTmp01 = StringReplace(strTmp01, ";", ",")

3410          lngRelAccts = (CharCnt(strTmp01, ",") + 1&)  ' ** Function: Below.
3420          ReDim arr_varRelAcct(RA_ELEMS, lngRelAccts - 1&)
3430          strTmp03 = strTmp01
3440          intPos01 = InStr(strTmp03, ",")
3450          lngX = 0&
3460          Do While intPos01 > 0
3470            arr_varRelAcct(RA_ACTO, lngX) = Left(strTmp03, (intPos01 - 1))
3480            strTmp03 = Mid(strTmp03, (intPos01 + 1))
3490            lngX = lngX + 1&
3500            intPos01 = InStr(strTmp03, ",")
3510            If intPos01 = 0 Then
3520              arr_varRelAcct(RA_ACTO, lngX) = strTmp03
3530              Exit Do
3540            End If
3550          Loop
3560          For lngX = 0& To (lngRelAccts - 1&)
3570            arr_varRelAcct(RA_PFX, lngX) = CBool(False)
3580            arr_varRelAcct(RS_TYPO, lngX) = vbNullString
3590            arr_varRelAcct(RA_TYPN, lngX) = vbNullString
3600            arr_varRelAcct(RA_ACTN, lngX) = vbNullString
3610          Next

              ' ** arr_varRelAcct() now has each related account in a separate element.
3620          For lngX = 0& To (lngRelAccts - 1&)
3630            strTmp01 = arr_varRelAcct(RA_ACTO, lngX)
3640            If Len(strTmp01) >= 5 Then
                  ' ** Minimum 2-char accountno to get prefix, so accounttype-accountno = 5.
3650              If Mid(strTmp01, 3, 1) = "-" Then
3660                strTmp02 = Left(strTmp01, 2)
3670                If IsNumeric(strTmp02) = True Then
3680                  For lngY = 0& To (lngTypes - 1&)
3690                    If arr_varType(AT_TYP, lngY) = strTmp02 Then
3700                      arr_varRelAcct(RA_PFX, lngX) = True
3710                      arr_varRelAcct(RS_TYPO, lngX) = strTmp02
3720                      If blnRetArr = False Then
3730                        varRetVal = CBool(True)
3740                        Exit For
3750                      End If
3760                    End If
3770                  Next
3780                End If
3790              End If
3800            End If
3810            If blnRetArr = False Then
3820              If varRetVal = True Then
                    ' ** Just need 1 hit to return True.
3830                Exit For
3840              End If
3850            Else
                  ' ** Collect the whole set, while supplies last!
3860            End If
3870          Next

3880          If blnRetArr = True Then
3890            varRetVal = arr_varRelAcct
3900          End If
3910        End If
3920      Else
3930        Select Case blnRetArr
            Case True
3940          arr_varRelAcct(RA_ACTO, 0) = RET_ERR
3950        Case False
3960          varRetVal = CBool(False)
3970        End Select
3980      End If
3990    Else
4000      Select Case blnRetArr
          Case True
4010        arr_varRelAcct(RA_ACTO, 0) = RET_ERR
4020      Case False
4030        varRetVal = CBool(False)
4040      End Select
4050    End If

EXITP:
4060    Set rst = Nothing
4070    Set qdf = Nothing
4080    Set dbs = Nothing
4090    HasTypePfx = varRetVal
4100    Exit Function

ERRH:
4110    varRetVal = CBool(False)
4120    Select Case ERR.Number
        Case Else
4130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4140    End Select
4150    Resume EXITP

End Function

Public Function IsLC(varInput As Variant) As Boolean
' ** IsLowerCase
' ** Numerals:
' **   48 - 57
' ** A - Z:
' **   65 - 90
' ** a - z:
' **   97 - 122

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "IsLC"

        Dim intLen As Integer
        Dim strTmp01 As String
        Dim intX As Integer
        Dim blnRetVal As Boolean

4210    blnRetVal = False

4220    If IsNull(varInput) = False Then
4230      If Trim(varInput) <> vbNullString Then
4240        blnRetVal = True
4250        strTmp01 = Trim(varInput)
4260        intLen = Len(strTmp01)
4270        For intX = 1 To intLen
4280          If Asc(Mid(strTmp01, intX, 1)) < 97 Or Asc(Mid(strTmp01, intX, 1)) > 122 Then
4290            blnRetVal = False
4300            Exit For
4310          End If
4320        Next
4330      End If
4340    End If

EXITP:
4350    IsLC = blnRetVal
4360    Exit Function

ERRH:
4370    blnRetVal = False
4380    Select Case ERR.Number
        Case Else
4390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4400    End Select
4410    Resume EXITP

End Function

Public Function IsUC(varInput As Variant, Optional varWithUnder As Variant, Optional varWithNum As Variant) As Boolean
' ** IsUpperCase
' ** Numerals:
' **   48 - 57
' ** A - Z:
' **   65 - 90
' ** a - z:
' **   97 - 122

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "IsUC"

        Dim intLen As Integer
        Dim blnWithUnderline As Boolean, blnWithNumeral As Boolean
        Dim strTmp01 As String
        Dim intX As Integer
        Dim blnRetVal As Boolean

4510    blnRetVal = False

4520    Select Case IsMissing(varWithUnder)
        Case True
4530      blnWithUnderline = False
4540    Case False
4550      blnWithUnderline = CBool(varWithUnder)
4560    End Select

4570    Select Case IsMissing(varWithNum)
        Case True
4580      blnWithNumeral = False
4590    Case False
4600      blnWithNumeral = CBool(varWithNum)
4610    End Select

4620    If IsNull(varInput) = False Then
4630      If Trim(varInput) <> vbNullString Then
4640        blnRetVal = True
4650        strTmp01 = Trim(varInput)
4660        intLen = Len(strTmp01)
4670        For intX = 1 To intLen
4680          If Asc(Mid(strTmp01, intX, 1)) >= 97 And Asc(Mid(strTmp01, intX, 1)) <= 122 Then
                ' ** Only need 1 lower case hit to negate it.
4690            blnRetVal = False
4700            Exit For
4710          End If
              'If Asc(Mid(strTmp01, intX, 1)) < 65 Or Asc(Mid(strTmp01, intX, 1)) > 90 Then
              '  If blnWithUnderline = True And Asc(Mid(strTmp01, intX, 1)) = 95 Then
              '    ' ** OK.
              '  ElseIf blnWithNumeral = True And (Asc(Mid(strTmp01, intX, 1)) >= 48 And Asc(Mid(strTmp01, intX, 1)) <= 57) Then
              '    ' ** OK.
              '  Else
              '    blnRetVal = False
              '    Exit For
              '  End If
              'End If
4720        Next
4730      End If
4740    End If

EXITP:
4750    IsUC = blnRetVal
4760    Exit Function

ERRH:
4770    blnRetVal = False
4780    Select Case ERR.Number
        Case Else
4790      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4800    End Select
4810    Resume EXITP

End Function

Public Function FormatAccountNum(varAcct As Variant) As String

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormatAccountNum"

        Dim blnAllNum As Boolean
        Dim intLen As Integer, intNonNum As Integer
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim intX As Integer
        Dim strRetVal As String

4910    strRetVal = vbNullString

4920    If IsNull(varAcct) = False Then

4930      strTmp01 = Trim(CStr(varAcct))
4940      intLen = Len(strTmp01)
4950      blnAllNum = True
4960      intNonNum = 0

4970      For intX = 1 To intLen
4980        If Asc(Mid(strTmp01, intX, 1)) >= 48 And Asc(Mid(strTmp01, intX, 1)) <= 57 Then
              ' ** Numeric, continue.
4990        Else
5000          blnAllNum = False
5010          intNonNum = intNonNum + 1
5020          If intNonNum = 1 Then intPos01 = intX
5030        End If
5040      Next

5050      If blnAllNum = True Then
5060        strRetVal = Right(String(15, "0") & strTmp01, 15)
5070      Else

5080        If intNonNum = 1 Then
              ' ** This could be a mid-number dash, a prefix, or a suffix.
5090          If intPos01 = 1 Then
                ' ** Prefix.
5100            If intLen > 1 Then
5110              strRetVal = Left(strTmp01, 1) & Right(String(14, "0") & strTmp01, 14)
5120            Else
                  ' ** One-character Account Number.
5130              strRetVal = Right(String(15, "_") & strTmp01, 15)
5140            End If
5150          ElseIf intPos01 = intLen Then
                ' ** Suffix.
5160            strRetVal = Right(String(14, "0") & Left(strTmp01, (intLen - 1)), 14) & Right(strTmp01, 1)
5170          Else
                ' ** Mid-number non-numeric character, most likely a dash.
5180            strRetVal = Right(String(7, "0") & Left(strTmp01, (intPos01 - 1)), 7) & Mid(strTmp01, intPos01, 1) & _
                  Right(String(7, "0") & Mid(strTmp01, (intPos01 + 1)), 7)
5190          End If
5200        Else
              ' ** Could be all letters, or something else.
5210          If intNonNum = intLen Then
                ' ** All non-numeric characters.
5220            strRetVal = Left(strTmp01 & String(15, "_"), 15)
5230          Else
                ' ** Should I consider a multi-character prefix or suffix?
                ' ** If these are strictly prefixes and suffixes, we could determine
                ' ** the longest of each within the Account table, then format them
                ' ** all with standard-length prefixes, padded numbers in the middle,
                ' ** and standard-length suffixes.
5240            If intPos01 = 1 Then
                  ' ** Though it starts at the beginning, just where are the numerals located?
5250              strRetVal = Left(strTmp01 & String(15, "_"), 15)
5260            Else
                  ' ** Who knows? Just do the best you can.
5270              strRetVal = Right(String(15, "0") & Left(strTmp01, (intPos01 - 1)), _
                    (15 - Len(Mid(strTmp01, intPos01)))) & Mid(strTmp01, intPos01)
5280            End If
5290          End If
5300        End If

            ' ** Replace any mid-number spaces with underscores.
5310        intPos01 = InStr(strRetVal, " ")
5320        Do While intPos01 > 0
5330          strRetVal = Left(strRetVal, (intPos01 - 1)) & "_" & Mid(strRetVal, (intPos01 + 1))
5340          intPos01 = InStr(strRetVal, " ")
5350        Loop

5360      End If

5370    End If

EXITP:
5380    FormatAccountNum = strRetVal
5390    Exit Function

ERRH:
5400    strRetVal = vbNullString
5410    Select Case ERR.Number
        Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Function

Public Function FormatPhoneNum(varPhone As Variant, Optional varStyle As Variant) As String
' ** This reformats a solid string of phone number digits
' ** into various structured formats.

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "FormatPhoneNum"

        Dim strRetVal As String
        Dim blnDashNotParen As Boolean, blnDotNotDash As Boolean, blnSpaceNotDash As Boolean
        Dim blnRemove As Boolean, blnWithOne As Boolean
        Dim intPos01 As Integer

5510    blnDashNotParen = False: blnDotNotDash = False: blnSpaceNotDash = False  ' ** Dafaults to parens around area-code.
5520    blnRemove = False: blnWithOne = False

5530    strRetVal = Trim(Nz(varPhone, vbNullString))

5540    If strRetVal <> vbNullString Then

5550      Select Case IsMissing(varStyle)
          Case True
            ' ** Default.
5560      Case False
5570        Select Case IsNull(varStyle)
            Case True
              ' ** Default.
5580        Case False
5590          Select Case varStyle
              Case "Remove", "Rem"
5600            blnRemove = True
5610          Case "Paren", "(", ")"
                ' ** Default.
5620          Case "Dash", "-"
5630            blnDashNotParen = True
5640          Case "Dot", "."
5650            blnDotNotDash = True
5660          Case "Space", " "
5670            blnSpaceNotDash = True
5680          End Select
5690        End Select
5700      End Select

          ' ** First, remove any dashes, dots, or parens already present.
5710      intPos01 = InStr(strRetVal, "-")
5720      If intPos01 > 0 Then
5730        Do While intPos01 > 0
5740          strRetVal = Left(strRetVal, (intPos01 - 1)) & Mid(strRetVal, (intPos01 + 1))
5750          intPos01 = InStr(strRetVal, "-")
5760        Loop
5770      End If
5780      intPos01 = InStr(strRetVal, ".")
5790      If intPos01 > 0 Then
5800        Do While intPos01 > 0
5810          strRetVal = Left(strRetVal, (intPos01 - 1)) & Mid(strRetVal, (intPos01 + 1))
5820          intPos01 = InStr(strRetVal, ".")
5830        Loop
5840      End If
5850      intPos01 = InStr(strRetVal, "(")
5860      If intPos01 > 0 Then
5870        Do While intPos01 > 0
5880          strRetVal = Left(strRetVal, (intPos01 - 1)) & Mid(strRetVal, (intPos01 + 1))
5890          intPos01 = InStr(strRetVal, "(")
5900        Loop
5910      End If
5920      intPos01 = InStr(strRetVal, ")")
5930      If intPos01 > 0 Then
5940        Do While intPos01 > 0
5950          strRetVal = Left(strRetVal, (intPos01 - 1)) & Mid(strRetVal, (intPos01 + 1))
5960          intPos01 = InStr(strRetVal, ")")
5970        Loop
5980      End If
5990      intPos01 = InStr(strRetVal, " ")
6000      If intPos01 > 0 Then
6010        Do While intPos01 > 0
6020          strRetVal = Left(strRetVal, (intPos01 - 1)) & Mid(strRetVal, (intPos01 + 1))
6030          intPos01 = InStr(strRetVal, " ")
6040        Loop
6050      End If

6060      Select Case blnRemove
          Case True
            ' ** We're done.
6070      Case False

6080        If Left(strRetVal, 1) = "1" Then
6090          blnWithOne = True
6100          strRetVal = Mid(strRetVal, 2)
6110        End If

6120        Select Case Len(strRetVal)
            Case 10
6130          If blnDashNotParen = True Or (blnWithOne = True And blnDotNotDash = False And blnSpaceNotDash = False) Then
6140            strRetVal = Left(strRetVal, 3) & "-" & Mid(strRetVal, 4, 3) & "-" & Right(strRetVal, 4)
6150          ElseIf blnDotNotDash = True Then
6160            strRetVal = Left(strRetVal, 3) & "." & Mid(strRetVal, 4, 3) & "." & Right(strRetVal, 4)
6170          ElseIf blnSpaceNotDash = True Then
6180            strRetVal = Left(strRetVal, 3) & " " & Mid(strRetVal, 4, 3) & " " & Right(strRetVal, 4)
6190          Else
6200            strRetVal = "(" & Left(strRetVal, 3) & ") " & Mid(strRetVal, 4, 3) & "-" & Right(strRetVal, 4)
6210          End If
6220        Case 7
6230          If blnDotNotDash = True Then
6240            strRetVal = Left(strRetVal, 3) & "." & Right(strRetVal, 4)
6250          ElseIf blnSpaceNotDash = True Then
6260            strRetVal = Left(strRetVal, 3) & " " & Right(strRetVal, 4)
6270          Else
6280            strRetVal = Left(strRetVal, 3) & "-" & Right(strRetVal, 4)
6290          End If
6300        Case Else
              ' ** Leave it alone.
6310        End Select

6320        If blnWithOne = True Then
6330          If blnDotNotDash = True Then
6340            strRetVal = "1." & strRetVal
6350          ElseIf blnSpaceNotDash = True Then
6360            strRetVal = "1 " & strRetVal
6370          Else
6380            strRetVal = "1-" & strRetVal
6390          End If
6400        End If

6410      End Select

6420    End If

EXITP:
6430    FormatPhoneNum = strRetVal
6440    Exit Function

ERRH:
6450    strRetVal = RET_ERR
6460    Select Case ERR.Number
        Case Else
6470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6480    End Select
6490    Resume EXITP

End Function

Public Function FormatUpperCase(strString As String) As String

6500  On Error GoTo ERRH

        Dim lngCnt As Long
        Dim lngLen As Long
        Dim strBuild As String
        Dim strChar As String

        Const THIS_PROC As String = "FormatUpperCase"

6510    lngLen = Len(strString)
6520    If lngLen = 0 Then
6530      FormatUpperCase = strString
6540    Else
6550      For lngCnt = 1 To lngLen
6560        strChar = Mid(strString, lngCnt, 1)
6570        If Asc(strChar) > 96 And Asc(strChar) < 123 Then
6580          strBuild = strBuild & Chr(Asc(strChar) - 32)
6590        Else
6600          strBuild = strBuild & strChar
6610        End If
6620      Next lngCnt
6630    End If

6640    FormatUpperCase = strBuild

EXITP:
6650    Exit Function

ERRH:
6660    FormatUpperCase = strString
6670    Select Case ERR.Number
        Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Function

Public Function FormatProperCase(strInput As String) As String

6800  On Error GoTo ERRH

        Dim lngLen As Long
        Dim blnWord As Boolean
        Dim lngX As Long
        Dim strRetVal As String

        Const THIS_PROC As String = "FormatProperCase"

6810    strRetVal = Trim(strInput)
6820    lngLen = Len(strRetVal)
6830    If lngLen > 0& Then
6840      strRetVal = UCase$(Left(strRetVal, 1)) & Mid(strRetVal, 2)
6850      blnWord = False
6860      For lngX = 2 To lngLen
6870        If Mid(strRetVal, lngX, 1) = " " Then
6880          blnWord = True
6890        Else
6900          If blnWord = False Then
6910            If lngX < lngLen Then
6920              strRetVal = Left(strRetVal, (lngX - 1&)) & LCase$(Mid(strRetVal, lngX, 1)) & Mid(strRetVal, (lngX + 1&))
6930            Else
6940              strRetVal = Left(strRetVal, (lngX - 1&)) & LCase$(Mid(strRetVal, lngX, 1))
6950            End If
6960          Else
6970            blnWord = False
6980            If lngX < lngLen Then
6990              strRetVal = Left(strRetVal, (lngX - 1&)) & UCase$(Mid(strRetVal, lngX, 1)) & Mid(strRetVal, (lngX + 1&))
7000            Else
7010              strRetVal = Left(strRetVal, (lngX - 1&)) & UCase$(Mid(strRetVal, lngX, 1))
7020            End If
7030          End If
7040        End If
7050      Next
7060    End If

EXITP:
7070    FormatProperCase = strRetVal
7080    Exit Function

ERRH:
7090    strRetVal = strInput
7100    Select Case ERR.Number
        Case Else
7110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7120    End Select
7130    Resume EXITP

End Function

Public Function FormatZip(varZip As Variant) As Variant

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "FormatZip"

        Dim varRetVal As Variant

7210    varRetVal = Null
7220    If IsNull(varZip) = False Then
7230      If Trim(varZip) <> vbNullString Then
7240        varRetVal = Trim(varZip)
7250        If InStr(varRetVal, " ") > 0 Then
7260          varRetVal = Rem_Spaces(varRetVal, True)  ' ** Function: Below.
7270        End If
7280        If Len(varRetVal) = 9 Then
7290          varRetVal = Left(varRetVal, 5) & "-" & Right(varRetVal, 4)
7300        End If
7310      End If
7320    End If

EXITP:
7330    FormatZip = varRetVal
7340    Exit Function

ERRH:
7350    varRetVal = RET_ERR
7360    Select Case ERR.Number
        Case Else
7370      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7380    End Select
7390    Resume EXITP

End Function

Public Function FormatZip9(varZip As Variant) As Variant

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormatZip9"

        Dim varRetVal As Variant

7410    varRetVal = Null
7420    If IsNull(varZip) = False Then
7430      If Trim(varZip) <> vbNullString Then
7440        varRetVal = Trim(varZip)
7450        If InStr(varRetVal, " ") > 0 Then
7460          varRetVal = Rem_Spaces(varRetVal, True)  ' ** Function: Below.
7470        End If
7480        If Len(varRetVal) = 9 Then
7490          varRetVal = Left(varRetVal, 5) & "-" & Right(varRetVal, 4)
7500        End If
7510      End If
7520    End If

EXITP:
7530    FormatZip9 = varRetVal
7540    Exit Function

ERRH:
7550    varRetVal = RET_ERR
7560    Select Case ERR.Number
        Case Else
7570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7580    End Select
7590    Resume EXITP

End Function

Public Function FixAmps(varInput As Variant) As Variant
' ** Return a double-ampersand for a single so that "Bill & Ted" doesn't look like "Bill _Ted"
' ** Make sure that if it already has a double-ampersand, we don't add to it!

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "FixAmps"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

7610    varRetVal = Null

7620    If IsNull(varInput) = False Then
7630      If Trim(varInput) <> vbNullString Then
7640        varRetVal = Trim(varInput)
7650        intPos01 = InStr(varRetVal, "&")
7660        If intPos01 > 0 Then
7670          Do While intPos01 > 0
7680            If intPos01 = 1 Then
7690              If Left(varRetVal, 2) <> "&&" Then
7700                varRetVal = "&" & varRetVal
7710              End If
7720              intPos01 = intPos01 + 1  ' ** So it won't keep finding the same one.
7730            ElseIf intPos01 = Len(varRetVal) Then
7740              If Right(varRetVal, 2) <> "&&" Then
7750                varRetVal = varRetVal & "&"
7760              End If
7770              intPos01 = intPos01 + 1
7780            Else
7790              If Mid(varRetVal, intPos01, 2) <> "&&" Then
7800                varRetVal = Left(varRetVal, (intPos01 - 1)) & "&" & Mid(varRetVal, intPos01)
7810              End If
7820              intPos01 = intPos01 + 1
7830            End If
7840            If (intPos01 + 1) >= Len(varRetVal) Then
7850              Exit Do
7860            Else
7870              intPos01 = InStr((intPos01 + 1), varRetVal, "&")
7880              If intPos01 = 0 Then Exit Do
7890            End If
7900          Loop
7910        End If
7920      End If
7930    End If

EXITP:
7940    FixAmps = varRetVal
7950    Exit Function

ERRH:
7960    varRetVal = RET_ERR
7970    Select Case ERR.Number
        Case Else
7980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7990    End Select
8000    Resume EXITP

End Function

Public Function FixQuotes(varInput As Variant) As Variant
' ** Replaces regular quotes with printers' quotes.

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "FixQuotes"

        Dim blnFound As Boolean
        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, intTmp02 As Integer
        Dim intX As Integer
        Dim varRetVal As Variant

8110    varRetVal = Null

        ' ** Single-quote, double-quote, printer's quote, printers quote.
        ' **  '  Chr(39)
        ' **  "  Chr(34)
        ' **  ‘  Chr(145)
        ' **  ’  Chr(146)
        ' **  “  Chr(147)
        ' **  ”  Chr(148)

8120    If IsNull(varInput) = False Then
8130      If Trim(varInput) <> vbNullString Then
8140        blnFound = False
8150        strTmp01 = Trim(varInput)
8160        intLen = Len(strTmp01)
            ' ** Check for single quotes.
8170        intPos01 = InStr(strTmp01, Chr(39))
8180        Do While intPos01 > 0
8190          blnFound = True
8200          Select Case intPos01
              Case 1
8210            If intLen > 1 Then
8220              strTmp01 = Chr(146) & Mid(strTmp01, 2)
8230            Else
8240              strTmp01 = Chr(146)
8250            End If
8260          Case intLen
8270            strTmp01 = Left(strTmp01, (intLen - 1)) & Chr(146)
8280          Case Else
8290            strTmp01 = Left(strTmp01, (intPos01 - 1)) & Chr(146) & Mid(strTmp01, (intPos01 + 1))
8300          End Select
8310          intPos01 = InStr(strTmp01, Chr(39))
8320        Loop
            ' ** Check for double quotes.
8330        intPos01 = InStr(strTmp01, Chr(34))
8340        If intPos01 > 0 Then
8350          blnFound = True
8360          intTmp02 = 0
8370          For intX = 1 To intLen
8380            If Asc(Mid(strTmp01, intX, 1)) = 34 Then
8390              intTmp02 = intTmp02 + 1
8400            End If
8410          Next
8420          If intTmp02 = 1 Then
8430            Select Case intPos01
                Case 1
8440              If intLen > 1 Then
8450                strTmp01 = Chr(147) & Mid(strTmp01, 2)
8460              Else
8470                strTmp01 = Chr(147)
8480              End If
8490            Case intLen
8500              strTmp01 = Left(strTmp01, (intLen - 1)) & Chr(148)
8510            Case Else
8520              strTmp01 = Left(strTmp01, (intPos01 - 1)) & Chr(147) & Mid(strTmp01, (intPos01 + 1))
8530            End Select
8540          Else
8550            intTmp02 = 0
8560            Do While intPos01 > 0
8570              intTmp02 = intTmp02 + 1
8580              Select Case intPos01
                  Case 1
8590                If intLen > 1 Then
8600                  strTmp01 = Chr(IIf((intTmp02 Mod 2) <> 0, 147, 148)) & Mid(strTmp01, 2)  ' ** Odd, Left; Even, Right.
8610                Else
8620                  strTmp01 = Chr(IIf((intTmp02 Mod 2) <> 0, 147, 148))
8630                End If
8640              Case intLen
8650                strTmp01 = Left(strTmp01, (intLen - 1)) & Chr(IIf((intTmp02 Mod 2) <> 0, 147, 148))
8660              Case Else
8670                strTmp01 = Left(strTmp01, (intPos01 - 1)) & Chr(IIf((intTmp02 Mod 2) <> 0, 147, 148)) & Mid(strTmp01, (intPos01 + 1))
8680              End Select
8690              intPos01 = InStr(strTmp01, Chr(34))
8700            Loop
8710          End If
8720        End If
8730        If blnFound = True Then
8740          varRetVal = strTmp01
8750        Else
8760          varRetVal = varInput
8770        End If
8780      Else
8790        varRetVal = varInput
8800      End If
8810    End If

EXITP:
8820    FixQuotes = varRetVal
8830    Exit Function

ERRH:
8840    varRetVal = RET_ERR
8850    Select Case ERR.Number
        Case Else
8860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8870    End Select
8880    Resume EXITP

End Function

Public Function FixQuotes2(varInput As Variant) As Variant
' ** Replaces double printers quotes with 2 single printers quotes,
' ** so that FixedSys can at least show it.

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "FixQuotes2"

        Dim blnFound As Boolean
        Dim intLen As Integer
        Dim intPos01 As Integer, intPos02 As Integer
        Dim strTmp01 As String, intTmp02 As Integer, intTmp03 As Integer
        Dim intX As Integer
        Dim varRetVal As Variant

8910    varRetVal = Null

        ' ** Though FixedSys has single printer's quotes, it doesn't have double ones.
        ' ** Single-quote, double-quote, printer's quote, printers quote.
        ' **  '  Chr(39)
        ' **  "  Chr(34)
        ' **  ‘  Chr(145)
        ' **  ’  Chr(146)
        ' **  “  Chr(147)
        ' **  ”  Chr(148)

8920    If IsNull(varInput) = False Then
8930      If Trim(varInput) <> vbNullString Then
8940        blnFound = False
8950        strTmp01 = Trim(varInput)
8960        intLen = Len(strTmp01)
            ' ** Check for double printer's quotes.
8970        intPos01 = InStr(strTmp01, Chr(147))  ' ** Opening quote.
8980        If intPos01 > 0 Then
8990          blnFound = True
9000          intTmp02 = 0
9010          For intX = 1 To intLen
9020            If Asc(Mid(strTmp01, intX, 1)) = 147 Then
9030              intTmp02 = intTmp02 + 1
9040            End If
9050          Next
9060          If intTmp02 = 1 Then
9070            Select Case intPos01
                Case 1
9080              If intLen > 1 Then
9090                strTmp01 = Chr(145) & Chr(145) & Mid(strTmp01, 2)
9100              Else
9110                strTmp01 = Chr(145) & Chr(145)
9120              End If
9130            Case intLen
9140              strTmp01 = Left(strTmp01, (intLen - 1)) & Chr(145) & Chr(145)
9150            Case Else
9160              strTmp01 = Left(strTmp01, (intPos01 - 1)) & Chr(145) & Chr(145) & Mid(strTmp01, (intPos01 + 1))
9170            End Select
9180          ElseIf intTmp02 > 1 Then
9190            intTmp02 = 0
9200            Do While intPos01 > 0
9210              intTmp02 = intTmp02 + 1
9220              Select Case intPos01
                  Case 1
9230                If intLen > 1 Then
9240                  strTmp01 = Chr(145) & Chr(145) & Mid(strTmp01, 2)
9250                Else
9260                  strTmp01 = Chr(145) & Chr(145)
9270                End If
9280              Case intLen
9290                strTmp01 = Left(strTmp01, (intLen - 1)) & Chr(145) & Chr(145)
9300              Case Else
9310                strTmp01 = Left(strTmp01, (intPos01 - 1)) & Chr(145) & Chr(145) & Mid(strTmp01, (intPos01 + 1))
9320              End Select
9330              intLen = Len(strTmp01)
9340              intPos01 = InStr(strTmp01, Chr(147))
9350            Loop
9360          End If
9370        End If
9380        intPos02 = InStr(strTmp01, Chr(148))  ' ** Closing quote.
9390        If intPos02 > 0 Then
9400          intLen = Len(strTmp01)
9410          intTmp03 = 0
9420          For intX = 1 To intLen
9430            If Asc(Mid(strTmp01, intX, 1)) = 148 Then
9440              intTmp03 = intTmp03 + 1
9450            End If
9460          Next
9470          If intTmp03 = 1 Then
9480            Select Case intPos02
                Case 1
9490              If intLen > 1 Then
9500                strTmp01 = Chr(146) & Chr(146) & Mid(strTmp01, 2)
9510              Else
9520                strTmp01 = Chr(146) & Chr(146)
9530              End If
9540            Case intLen
9550              strTmp01 = Left(strTmp01, (intLen - 1)) & Chr(146) & Chr(146)
9560            Case Else
9570              strTmp01 = Left(strTmp01, (intPos02 - 1)) & Chr(146) & Chr(146) & Mid(strTmp01, (intPos02 + 1))
9580            End Select
9590          ElseIf intTmp03 > 1 Then
9600            intTmp03 = 0
9610            Do While intPos02 > 0
9620              intTmp03 = intTmp03 + 1
9630              Select Case intPos02
                  Case 1
9640                If intLen > 1 Then
9650                  strTmp01 = Chr(146) & Chr(146) & Mid(strTmp01, 2)
9660                Else
9670                  strTmp01 = Chr(146) & Chr(146)
9680                End If
9690              Case intLen
9700                strTmp01 = Left(strTmp01, (intLen - 1)) & Chr(146) & Chr(146)
9710              Case Else
9720                strTmp01 = Left(strTmp01, (intPos02 - 1)) & Chr(146) & Chr(146) & Mid(strTmp01, (intPos02 + 1))
9730              End Select
9740              intPos02 = InStr(strTmp01, Chr(148))
9750            Loop
9760          End If
9770        End If
9780        If blnFound = True Then
9790          varRetVal = strTmp01
9800        Else
9810          varRetVal = varInput
9820        End If
9830      Else
9840        varRetVal = varInput
9850      End If
9860    End If

        'S_PQuotes: CBool(IIf(InStr([shortname],Chr(147))>0 Or InStr([shortname],Chr(148))>0,True,False))
        'L_PQuotes: CBool(IIf(IsNull([legalname])=True,False,IIf(InStr([legalname],Chr(147))>0 Or InStr([legalname],Chr(148))>0,True,False)))

EXITP:
9870    FixQuotes2 = varRetVal
9880    Exit Function

ERRH:
9890    varRetVal = RET_ERR
9900    Select Case ERR.Number
        Case Else
9910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9920    End Select
9930    Resume EXITP

End Function

Public Function Rem_Amp(varInput As Variant) As Variant
' ** Remove ampersands from a string.

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Amp"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

10010   varRetVal = Null

10020   If IsNull(varInput) = False Then
10030     varRetVal = varInput
10040     intPos01 = InStr(varRetVal, "&")
10050     Do While intPos01 > 0
10060       If intPos01 = 1 Then
10070         varRetVal = Mid(varRetVal, 2)
10080       ElseIf intPos01 = Len(varRetVal) Then
10090         varRetVal = Left(varRetVal, (Len(varRetVal) - 1))
10100       Else
10110         varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1))
10120       End If
10130       intPos01 = InStr(varRetVal, "&")
10140     Loop
10150   End If

EXITP:
10160   Rem_Amp = varRetVal
10170   Exit Function

ERRH:
10180   Select Case ERR.Number
        Case Else
10190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10200   End Select
10210   Resume EXITP

End Function

Public Function Rem_Apost(varInput As Variant) As Variant
' ** Remove apostrophes from a string.

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Apost"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

10310   varRetVal = Null

10320   If IsNull(varInput) = False Then
10330     varRetVal = varInput
10340     intPos01 = InStr(varRetVal, "'")
10350     Do While intPos01 > 0
10360       If intPos01 = 1 Then
10370         varRetVal = Mid(varRetVal, 2)
10380       ElseIf intPos01 = Len(varRetVal) Then
10390         varRetVal = Left(varRetVal, (Len(varRetVal) - 1))
10400       Else
10410         varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1))
10420       End If
10430       intPos01 = InStr(varRetVal, "'")
10440     Loop
10450   End If

EXITP:
10460   Rem_Apost = varRetVal
10470   Exit Function

ERRH:
10480   Select Case ERR.Number
        Case Else
10490     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10500   End Select
10510   Resume EXITP

End Function

Public Function Rem_Braces(varInput As Variant) As Variant
' ** Remove braces from a text string.

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Braces"

        Dim intLen As Integer
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

10610   varRetVal = Null

10620   strTmp01 = vbNullString: strTmp02 = vbNullString

10630   If IsNull(varInput) = False Then
10640     strTmp01 = Trim(varInput)
10650     If strTmp01 <> vbNullString Then
10660       intPos01 = InStr(strTmp01, "{")
10670       If intPos01 > 0 Then
10680         intLen = Len(strTmp01)
10690         For intX = 1 To intLen
10700           If Mid(strTmp01, intX, 1) = "{" Then
                  ' ** Skip it!
10710           ElseIf Mid(strTmp01, intX, 1) = "}" Then
                  ' ** Skip it!
10720           Else
10730             strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
10740           End If
10750         Next
10760         strTmp01 = strTmp02
10770       End If
10780       varRetVal = strTmp01
10790     End If
10800   End If

EXITP:
10810   Rem_Braces = varRetVal
10820   Exit Function

ERRH:
10830   varRetVal = Null
10840   Select Case ERR.Number
        Case Else
10850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10860   End Select
10870   Resume EXITP

End Function

Public Function Rem_Brackets(varInput As Variant) As Variant
' ** Remove brackets from a SQL string.

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Brackets"

        Dim intLen As Integer
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

10910   varRetVal = Null

10920   strTmp01 = vbNullString: strTmp02 = vbNullString

10930   If IsNull(varInput) = False Then
10940     strTmp01 = Trim(varInput)
10950     If strTmp01 <> vbNullString Then
10960       intPos01 = InStr(strTmp01, "[")
10970       If intPos01 > 0 Then
10980         intLen = Len(strTmp01)
10990         For intX = 1 To intLen
11000           If Mid(strTmp01, intX, 1) = "[" Then
                  ' ** Skip it!
11010           ElseIf Mid(strTmp01, intX, 1) = "]" Then
                  ' ** Skip it!
11020           Else
11030             strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
11040           End If
11050         Next
11060         strTmp01 = strTmp02
11070       End If
11080       varRetVal = strTmp01
11090     End If
11100   End If

EXITP:
11110   Rem_Brackets = varRetVal
11120   Exit Function

ERRH:
11130   varRetVal = Null
11140   Select Case ERR.Number
        Case Else
11150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11160   End Select
11170   Resume EXITP

End Function

Public Function Rem_Colon(varInput As Variant) As Variant
' ** Remove Colons from a string.

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Colon"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

11210   varRetVal = Null

11220   If IsNull(varInput) = False Then
11230     varRetVal = varInput
11240     intPos01 = InStr(varRetVal, ":")
11250     Do While intPos01 > 0
11260       If intPos01 = 1 Then
11270         varRetVal = Mid(varRetVal, 2)
11280       ElseIf intPos01 = Len(varRetVal) Then
11290         varRetVal = Left(varRetVal, (Len(varRetVal) - 1))
11300       Else
11310         varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1))
11320       End If
11330       intPos01 = InStr(varRetVal, ":")
11340     Loop
11350   End If

EXITP:
11360   Rem_Colon = varRetVal
11370   Exit Function

ERRH:
11380   Select Case ERR.Number
        Case Else
11390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11400   End Select
11410   Resume EXITP

End Function

Public Function Rem_Comma(varInput As Variant) As Variant
' ** Remove commas from a string.

11500 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Comma"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

11510   varRetVal = Null

11520   If IsNull(varInput) = False Then
11530     varRetVal = varInput
11540     intPos01 = InStr(varRetVal, ",")
11550     Do While intPos01 > 0
11560       If intPos01 = 1 Then
11570         varRetVal = Mid(varRetVal, 2)
11580       ElseIf intPos01 = Len(varRetVal) Then
11590         varRetVal = Left(varRetVal, (Len(varRetVal) - 1))
11600       Else
11610         varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1))
11620       End If
11630       intPos01 = InStr(varRetVal, ",")
11640     Loop
11650   End If

EXITP:
11660   Rem_Comma = varRetVal
11670   Exit Function

ERRH:
11680   Select Case ERR.Number
        Case Else
11690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11700   End Select
11710   Resume EXITP

End Function

Public Function Rem_CRLF(varInput As Variant) As Variant
' ** Remove Carriage-Return and Line-Feed from a string.

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_CRLF"

        Dim intPos01 As Integer
        Dim blnSpaceAdded As Boolean
        Dim varRetVal As Variant

11810   varRetVal = Null

11820   If IsNull(varInput) = False Then
11830     blnSpaceAdded = False
11840     varRetVal = Trim(CStr(varInput))
11850     intPos01 = InStr(varInput, Chr(13))
11860     If intPos01 > 0 Then
11870       Do While intPos01 > 0
11880         varRetVal = Left(varRetVal, (intPos01 - 1)) & " " & Mid(varRetVal, (intPos01 + 1))
11890         intPos01 = InStr(varRetVal, Chr(13))
11900         blnSpaceAdded = True
11910       Loop
11920     End If
11930     intPos01 = InStr(varInput, Chr(10))
11940     If intPos01 > 0 Then
11950       Do While intPos01 > 0
11960         If blnSpaceAdded = False Then
11970           varRetVal = Left(varRetVal, (intPos01 - 1)) & " " & Mid(varRetVal, (intPos01 + 1))
11980         Else
11990           varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1))
12000         End If
12010         intPos01 = InStr(varRetVal, Chr(10))
12020       Loop
12030     End If
12040   End If

EXITP:
12050   Rem_CRLF = varRetVal
12060   Exit Function

ERRH:
12070   Select Case ERR.Number
        Case Else
12080     Beep
12090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12100   End Select
12110   Resume EXITP

End Function

Public Function Rem_Dash(varInput As Variant) As Variant
' ** Remove dashes from a string.

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Dash"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

12210   varRetVal = Null

12220   If IsNull(varInput) = False Then
12230     varRetVal = varInput
12240     intPos01 = InStr(varRetVal, "-")
12250     Do While intPos01 > 0
12260       If intPos01 = 1 Then
12270         varRetVal = Mid(varRetVal, 2)
12280       ElseIf intPos01 = Len(varRetVal) Then
12290         varRetVal = Left(varRetVal, (Len(varRetVal) - 1))
12300       Else
12310         varRetVal = Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1))
12320       End If
12330       intPos01 = InStr(varRetVal, "-")
12340     Loop
12350     varRetVal = Trim(Rem_Spaces(varRetVal))
12360   End If

EXITP:
12370   Rem_Dash = varRetVal
12380   Exit Function

ERRH:
12390   Select Case ERR.Number
        Case Else
12400     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12410   End Select
12420   Resume EXITP

End Function

Public Function Rem_Dollar(varInput As Variant, Optional varCurrID As Variant) As String
' ** Remove currency sign sign and commas from a string.
' ** Supports foreign currency symbols.

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Dollar"

        Dim lngCurrID As Long, strCurrSym As String
        Dim blnFound As Boolean
        Dim intLen01 As Integer, intLen02 As Integer
        Dim intX As Integer, intY As Integer
        Dim strRetVal As String

12510   strRetVal = vbNullString

12520   Select Case IsMissing(varCurrID)
        Case True
12530     lngCurrID = 150&
12540     strCurrSym = "$"
12550   Case False
12560     Select Case IsNull(varCurrID)
          Case True
12570       lngCurrID = 150&
12580       strCurrSym = "$"
12590     Case False
12600       lngCurrID = varCurrID
12610       strCurrSym = CurrSym_Get(lngCurrID)  ' ** Module Function: modCurrencyFuncs.
12620     End Select
12630   End Select

12640   If IsNull(varInput) = False Then
12650     intLen01 = Len(varInput)
12660     intLen02 = Len(strCurrSym)
12670     blnFound = False: intY = 0
12680     For intX = 1 To intLen01
12690       If blnFound = True And intY < intLen02 Then
12700         intY = intY + 1
12710       Else
12720         If Mid(varInput, intX, intLen02) <> strCurrSym Then  ' ** Strip dollar sign.
12730           If Mid(varInput, intX, 1) <> " " Then  ' ** Strip spaces.
12740             If Mid(varInput, intX, 1) = "(" Then
12750               strRetVal = strRetVal & "-"
12760             ElseIf Mid(varInput, intX, 1) <> ")" Then
12770               If Mid(varInput, intX, 1) <> "," Then
12780                 strRetVal = strRetVal & Mid(varInput, intX, 1)
12790               End If
12800             End If
12810           End If
12820         Else
12830           blnFound = True
12840           intY = 1
12850         End If
12860       End If
12870     Next
12880   End If

EXITP:
12890   Rem_Dollar = strRetVal
12900   Exit Function

ERRH:
12910   strRetVal = vbNullString
12920   Select Case ERR.Number
        Case Else
12930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12940   End Select
12950   Resume EXITP

End Function

Public Function Rem_ENotation(varInput As Variant) As Variant
' ** Lessen precision to remove E-Notation.

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_ENotation"

        Dim strVal As String, dblVal As Double, strMultiplier As String, intMultiplier As Integer
        Dim blnNeg As Boolean
        Dim intPos01 As Integer, intPos02 As Integer
        Dim varRetVal As Variant

13010   varRetVal = Null
13020   blnNeg = False

13030   If IsNull(varInput) = False Then
13040     strVal = CStr(varInput)
13050     intPos01 = InStr(strVal, "E")
13060     If intPos01 > 0 Then                                                        '-1.24296162419991E-02
13070       If Left(strVal, 1) = "-" Then
13080         blnNeg = True
13090         strVal = Mid(strVal, 2)                                               '1.24296162419991E-02
13100       End If
13110       intPos01 = InStr(strVal, "E")
13120       Do While intPos01 > 0
13130         strMultiplier = Mid(strVal, (intPos01 + 1))                            '-02
13140         If Left(strMultiplier, 1) = "-" Then
13150           strMultiplier = Mid(strMultiplier, 2)                               '02
13160           strVal = Left(strVal, (intPos01 - 1))                                '1.24296162419991
13170           intMultiplier = CInt(Val(strMultiplier))                             '2
13180           If intMultiplier > 12 Then intMultiplier = 12
13190           strVal = Left(strVal, (Len(strVal) - intMultiplier))                '1.242961624199
13200           strVal = String(intMultiplier, "0") & strVal                         '001.242961624199
13210           intPos02 = InStr(strVal, ".")                                         '4
13220           strVal = Left(strVal, (intPos02 - 1)) & Mid(strVal, (intPos02 + 1))  '001242961624199
13230           intPos02 = intPos02 - intMultiplier                                    '2
13240           strVal = Left(strVal, (intPos02 - 1)) & "." & Mid(strVal, intPos02)  '0.01242961624199
13250           If blnNeg = True Then
13260             strVal = "-" & strVal                                              '-0.01242961624199
13270           End If
13280           dblVal = CDbl(Val(strVal))
13290           strVal = CStr(dblVal)
13300           intPos01 = InStr(strVal, "E")
13310         Else
                ' ** Not expecting to handle large numbers here.
13320           Beep
13330           varRetVal = -1
13340           Exit Do
13350         End If
13360       Loop
13370       varRetVal = CDbl(Val(strVal))
13380     Else
13390       varRetVal = varInput
13400     End If
13410   End If

EXITP:
13420   Rem_ENotation = varRetVal
13430   Exit Function

ERRH:
13440   Select Case ERR.Number
        Case Else
13450     Beep
13460     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13470   End Select
13480   Resume EXITP

End Function

Public Function Rem_ENotation2(varInput As Variant) As String
' ** Move the decimal point left as far as needed, and
' ** return a string of whatever length is required.
' ** Precision remains unchanged.

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_ENotation2"

        Dim intMovesToMake As Integer
        Dim intMovesMade As Integer
        Dim blnIsNeg As Boolean
        Dim intPos01 As Integer
        Dim strRetVal As String

13510   strRetVal = vbNullString

13520   blnIsNeg = False

13530   If IsNull(varInput) = False Then
13540     If varInput < 0 Then blnIsNeg = True
13550     strRetVal = CStr(varInput)
13560     strRetVal = Trim(strRetVal)
13570     If strRetVal <> vbNullString Then
13580       If blnIsNeg = True Then
13590         intPos01 = InStr(strRetVal, "-")
13600         If intPos01 = 1 Then
13610           strRetVal = Mid(strRetVal, 2)
13620         Else
                'Debug.Print "'" & strRetVal
13630         End If
13640       End If
13650       intPos01 = InStr(strRetVal, "E-")
13660       If intPos01 > 0 Then
13670         intMovesToMake = Val(Mid(strRetVal, (intPos01 + 2)))  ' ** Number of times to move the decimal point to the left.
13680         intMovesMade = 0
13690         strRetVal = Left(strRetVal, (intPos01 - 1))
              ' ** Decimal point should always be the 2nd character in E-Notation.
13700         strRetVal = "." & Left(strRetVal, 1) & Mid(strRetVal, 3)
13710         intMovesMade = 1
13720         Do While intMovesMade < intMovesToMake
13730           strRetVal = ".0" & Mid(strRetVal, 2)
13740           intMovesMade = intMovesMade + 1
                '2.31481000001477E-05
                '0.231481000001477E-04
                '0.0231481000001477E-03
                '0.00231481000001477E-02
                '0.000231481000001477E-01
                '0.0000231481000001477
                '0.0000231481
                '0.0000231481000001477
13750         Loop
13760         strRetVal = "0" & strRetVal  ' ** Add a zero before the decimal place.
13770         If blnIsNeg = True Then strRetVal = "-" & strRetVal
13780       End If
13790     End If
13800   End If

EXITP:
13810   Rem_ENotation2 = strRetVal
13820   Exit Function

ERRH:
13830   strRetVal = vbNullString
13840   Select Case ERR.Number
        Case Else
13850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13860   End Select
13870   Resume EXITP

End Function

Public Function Rem_Parens(varInput As Variant) As Variant
' ** Remove parentheses from a string.

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Parens"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

13910   varRetVal = Null

13920   If IsNull(varInput) = False Then
13930     varRetVal = Trim(varInput)
13940     intPos01 = InStr(varRetVal, "(")
13950     If intPos01 > 0 Then
13960       Do While intPos01 > 0
13970         If intPos01 = 1 Then
13980           If Len(varRetVal) > 1 Then
13990             varRetVal = Trim(Mid(varRetVal, 2))
14000           Else
14010             varRetVal = Null
14020           End If
14030         Else
14040           If intPos01 = Len(varRetVal) Then
14050             varRetVal = Trim(Left(varRetVal, (Len(varRetVal) - 1)))
14060           Else
14070             varRetVal = Trim(Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1)))
14080           End If
14090         End If
14100         intPos01 = InStr(varRetVal, "(")
14110       Loop
14120     End If
14130     intPos01 = InStr(varRetVal, ")")
14140     If intPos01 > 0 Then
14150       Do While intPos01 > 0
14160         If intPos01 = 1 Then
14170           If Len(varRetVal) > 1 Then
14180             varRetVal = Trim(Mid(varRetVal, 2))
14190           Else
14200             varRetVal = Null
14210           End If
14220         Else
14230           If intPos01 = Len(varRetVal) Then
14240             varRetVal = Trim(Left(varRetVal, (Len(varRetVal) - 1)))
14250           Else
14260             varRetVal = Trim(Left(varRetVal, (intPos01 - 1)) & Mid(varRetVal, (intPos01 + 1)))
14270           End If
14280         End If
14290         intPos01 = InStr(varRetVal, ")")
14300       Loop
14310     End If
14320   End If

EXITP:
14330   Rem_Parens = varRetVal
14340   Exit Function

ERRH:
14350   varRetVal = RET_ERR
14360   Select Case ERR.Number
        Case Else
14370     Beep
14380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14390   End Select
14400   Resume EXITP

End Function

Public Function Rem_Period(varInput As Variant, Optional varReplace As Variant)
' ** Remove Period from a string.

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Period"

        Dim strReplace As String
        Dim intDots As Integer, arr_varDot() As Variant
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim varRetVal As Variant

        Const D_ELEMS As Integer = 2  ' ** Array's first-element UBound().
        'Const D_POS As Integer = 0
        'Const D_LFT As Integer = 1
        'Const D_RGT As Integer = 2

14510   varRetVal = Null

14520   intDots = 0
14530   ReDim arr_varDot(D_ELEMS, 0)

14540   If IsMissing(varReplace) = True Then
14550     strReplace = vbNullString
14560   Else
14570     If IsNull(varReplace) = False Then
14580       strReplace = varReplace
14590     Else
14600       strReplace = vbNullString
14610     End If
14620   End If

14630   If IsNull(varInput) = False Then
14640     strTmp01 = Trim(CStr(varInput))
14650     If strTmp01 <> vbNullString Then
14660       intPos01 = InStr(strTmp01, ".")
14670       If intPos01 > 0 Then
14680         strTmp02 = Replace(strTmp01, ".", strReplace, 1, -1, vbBinaryCompare)
              ' ** VbCompare enumeration.
              ' **    0  vbBinaryCompare     Performs a binary comparison.
              ' **    1  vbTextCompare       Performs a textual comparison.
              ' **    2  vbDatabaseCompare   Microsoft Access only. Performs a comparison based on information in your database.
              ' **    3  vbUseCompareOption  Performs a comparison using the setting of the Option Compare statement. (Stated value, -1, is wrong!)
14690         varRetVal = strTmp02
14700       Else
14710         varRetVal = varInput
14720       End If
14730     Else
14740       varRetVal = varInput
14750     End If
14760   End If

EXITP:
14770   Rem_Period = varRetVal
14780   Exit Function

ERRH:
14790   varRetVal = RET_ERR
14800   Select Case ERR.Number
        Case Else
14810     Beep
14820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14830   End Select
14840   Resume EXITP

End Function

Public Function Rem_Quotes(varInput As Variant) As Variant
' ** Remove standard double-quotes from beginning and end of string.

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Quotes"

        Dim strTmp01 As String
        Dim varRetVal As Variant

14910   varRetVal = Null

14920   If IsNull(varInput) = False Then
14930     strTmp01 = Trim(varInput)
14940     If strTmp01 <> vbNullString Then
14950       If Left(strTmp01, 1) = Chr(34) Then
14960         strTmp01 = Mid(strTmp01, 2)
14970       End If
14980       If Right(strTmp01, 1) = Chr(34) Then
14990         strTmp01 = Left(strTmp01, (Len(strTmp01) - 1))
15000       End If
15010       varRetVal = strTmp01
15020     End If
15030   End If

EXITP:
15040   Rem_Quotes = varRetVal
15050   Exit Function

ERRH:
15060   varRetVal = Null
15070   Select Case ERR.Number
        Case Else
15080     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15090   End Select
15100   Resume EXITP

End Function

Public Function Rem_Semi(varInput As Variant) As Variant

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Semi"

        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

15210   varRetVal = Null

15220   If IsNull(varInput) = False Then
15230     strTmp01 = Trim(varInput)
15240     If strTmp01 <> vbNullString Then
15250       intLen = Len(strTmp01)
15260       strTmp02 = vbNullString
15270       intPos01 = InStr(strTmp01, ";")
15280       If intPos01 > 0 Then
15290         For intX = 1 To intLen
15300           If Mid(strTmp01, intX, 1) <> ";" Then
15310             strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
15320           End If
15330         Next
15340       Else
15350         strTmp02 = strTmp01
15360       End If
15370       varRetVal = strTmp02
15380     End If
15390   End If

EXITP:
15400   Rem_Semi = varRetVal
15410   Exit Function

ERRH:
15420   varRetVal = Null
15430   Select Case ERR.Number
        Case Else
15440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15450   End Select
15460   Resume EXITP

End Function

Public Function Rem_Slash(varInput As Variant) As Variant

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Slash"

        Dim intPos01 As Integer, intPos02 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

15510   varRetVal = Null

15520   If IsNull(varInput) = False Then
15530     strTmp01 = Trim(varInput)
15540     If strTmp01 <> vbNullString Then
15550       intLen = Len(strTmp01)
15560       strTmp02 = vbNullString
15570       intPos01 = InStr(strTmp01, "/")
15580       intPos02 = InStr(strTmp01, "\")
15590       If intPos01 > 0 Or intPos02 > 0 Then
15600         For intX = 1 To intLen
15610           If Mid(strTmp01, intX, 1) <> "/" And Mid(strTmp01, intX, 1) <> "\" Then
15620             strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
15630           End If
15640         Next
15650       Else
15660         strTmp02 = strTmp01
15670       End If
15680       varRetVal = strTmp02
15690     End If
15700   End If

EXITP:
15710   Rem_Slash = varRetVal
15720   Exit Function

ERRH:
15730   varRetVal = Null
15740   Select Case ERR.Number
        Case Else
15750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15760   End Select
15770   Resume EXITP

End Function

Public Function Rem_Spaces(varInput As Variant, Optional varAll As Variant) As Variant
' ** Remove extra spaces from inside a string.

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Spaces"

        Dim intLen As Integer, intCnt As Integer
        Dim blnAll As Boolean
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

15810   varRetVal = Null

15820   Select Case IsMissing(varAll)
        Case True
15830     blnAll = False
15840   Case False
15850     blnAll = CBool(varAll)
15860   End Select

15870   If IsNull(varInput) = False Then
15880     strTmp01 = Trim(varInput)
15890     If strTmp01 <> vbNullString Then
15900       intLen = Len(strTmp01)
15910       intCnt = 0
15920       strTmp02 = vbNullString
15930       For intX = 1 To intLen
15940         If Mid(strTmp01, intX, 1) <> " " Then
15950           strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
15960           intCnt = 0
15970         Else
15980           Select Case blnAll
                Case True
                  ' ** Skip it!
15990           Case False
16000             If intCnt = 0 Then
16010               intCnt = intCnt + 1
16020               strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
16030             Else
                    ' ** Skip it!
16040               intCnt = intCnt + 1
16050             End If
16060           End Select
16070         End If
16080       Next
16090       strTmp01 = strTmp02
16100     End If
16110     varRetVal = strTmp01
16120   End If

EXITP:
16130   Rem_Spaces = varRetVal
16140   Exit Function

ERRH:
16150   varRetVal = RET_ERR
16160   Select Case ERR.Number
        Case Else
16170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16180   End Select
16190   Resume EXITP

End Function

Public Function Rem_Tab(varInput As Variant) As Variant
' ** Remove all Tab characters from a string.

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Tab"

        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

16210   varRetVal = Null

16220   If IsNull(varInput) = False Then
16230     strTmp01 = Trim(varInput)
16240     If strTmp01 <> vbNullString Then
16250       intLen = Len(strTmp01)
16260       strTmp02 = vbNullString
16270       intPos01 = InStr(strTmp01, Chr(9))
16280       If intPos01 > 0 Then
16290         For intX = 1 To intLen
16300           If Mid(strTmp01, intX, 1) <> Chr(9) Then
16310             strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
16320           End If
16330         Next
16340       Else
16350         strTmp02 = strTmp01
16360       End If
16370       varRetVal = strTmp02
16380     End If
16390   End If

EXITP:
16400   Rem_Tab = varRetVal
16410   Exit Function

ERRH:
16420   varRetVal = Null
16430   Select Case ERR.Number
        Case Else
16440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16450   End Select
16460   Resume EXITP

End Function

Public Function Rem_Under(varInput As Variant) As Variant
' ** Remove all underscores/underlines from a string.

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "Rem_Under"

        Dim intPos01 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

16510   varRetVal = Null

16520   If IsNull(varInput) = False Then
16530     strTmp01 = Trim(varInput)
16540     If strTmp01 <> vbNullString Then
16550       intLen = Len(strTmp01)
16560       strTmp02 = vbNullString
16570       intPos01 = InStr(strTmp01, "_")
16580       If intPos01 > 0 Then
16590         For intX = 1 To intLen
16600           If Mid(strTmp01, intX, 1) <> "_" Then
16610             strTmp02 = strTmp02 & Mid(strTmp01, intX, 1)
16620           End If
16630         Next
16640       Else
16650         strTmp02 = strTmp01
16660       End If
16670       varRetVal = strTmp02
16680     End If
16690   End If

EXITP:
16700   Rem_Under = varRetVal
16710   Exit Function

ERRH:
16720   varRetVal = Null
16730   Select Case ERR.Number
        Case Else
16740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16750   End Select
16760   Resume EXITP

End Function

Public Function SpaceToUnder(varInput As Variant) As Variant

16800 On Error GoTo ERRH

        Const THIS_PROC As String = "SpaceToUnder"

        Dim intPos01 As Integer
        Dim varRetVal As Variant

16810   varRetVal = Null

16820   If IsNull(varInput) = False Then
16830     varRetVal = varInput
16840     intPos01 = InStr(varRetVal, " ")
16850     If intPos01 > 0 Then
16860       Do While intPos01 > 0
16870         varRetVal = Left(varRetVal, intPos01 - 1) & "_" & Mid(varRetVal, intPos01 + 1)
16880         intPos01 = InStr(varRetVal, " ")
16890       Loop
16900     Else
16910       varRetVal = varInput
16920     End If
16930   End If

EXITP:
16940   SpaceToUnder = varRetVal
16950   Exit Function

ERRH:
16960   varRetVal = Null
16970   Select Case ERR.Number
        Case Else
16980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16990   End Select
17000   Resume EXITP

End Function

Public Function GlobalVarGet(varGlobalVar As Variant) As Variant

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "GlobalVarGet"

        Dim varRetVal As Variant

17110   varRetVal = Null

17120   If IsNull(varGlobalVar) = False Then
17130     Select Case varGlobalVar
            ' ** CoInfo variables.
          Case "gstrCo_Name"
17140       varRetVal = gstrCo_Name
17150     Case "gstrCo_Address1"
17160       varRetVal = gstrCo_Address1
17170     Case "gstrCo_Address2"
17180       varRetVal = gstrCo_Address2
17190     Case "gstrCo_City"
17200       varRetVal = gstrCo_City
17210     Case "gstrCo_State"
17220       varRetVal = gstrCo_State
17230     Case "gstrCo_Zip"
17240       varRetVal = gstrCo_Zip
17250     Case "gstrCo_Country"
17260       varRetVal = gstrCo_Country
17270     Case "gstrCo_PostalCode"
17280       varRetVal = gstrCo_PostalCode
17290     Case "gstrCo_Phone"
17300       varRetVal = gstrCo_Phone
            ' ** Trust Accountant Options.
17310     Case "gblnIncomeTaxCoding"
17320       varRetVal = gblnIncomeTaxCoding
17330     Case "gblnRevenueExpenseTracking"
17340       varRetVal = gblnRevenueExpenseTracking
17350     Case "gblnAccountNoWithType"
17360       varRetVal = gblnAccountNoWithType
17370     Case "gblnSeparateCheckingAccounts"
17380       varRetVal = gblnSeparateCheckingAccounts
17390     Case "gblnTabCopyAccount"
17400       varRetVal = gblnTabCopyAccount
17410     Case "gblnLinkRevTaxCodes"
17420       varRetVal = gblnLinkRevTaxCodes
            ' ** Court Report variables.
17430     Case "gcurCrtRpt_NY_ICash"
17440       varRetVal = gcurCrtRpt_NY_ICash
17450     Case "gcurCrtRpt_NY_InputNew"
17460       varRetVal = gcurCrtRpt_NY_InputNew
17470     Case "gstrCrtRpt_Ordinal"
17480       varRetVal = gstrCrtRpt_Ordinal
17490     Case "gstrCrtRpt_Period"
17500       varRetVal = gstrCrtRpt_Period
17510     Case "gstrCrtRpt_Version"
17520       varRetVal = gstrCrtRpt_Version
            ' ** Miscellaneous variables.
17530     Case "gstrAccountNo"
17540       varRetVal = gstrAccountNo
17550     Case "gstrAccountName"
17560       varRetVal = gstrAccountName
17570     Case "glngAssetNo"
17580       varRetVal = glngAssetNo
17590     Case "glngCurrID"
17600       varRetVal = glngCurrID
17610     Case "gdatEndDate"
17620       varRetVal = gdatEndDate
17630     Case "gblnGoToReport"
17640       varRetVal = gblnGoToReport
17650     Case "gblnIsLiability"
17660       varRetVal = gblnIsLiability
17670     Case "gstrJournalUser"
17680       If gstrJournalUser = vbNullString Then gstrJournalUser = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
17690       varRetVal = gstrJournalUser
17700     Case "gblnLocSuggest"
17710       varRetVal = gblnLocSuggest
17720     Case "gdatMarketDate"
17730       varRetVal = gdatMarketDate
17740     Case "glngMonthID"
17750       varRetVal = glngMonthID
17760     Case "gstrPurchaseType"
17770       varRetVal = gstrPurchaseType
17780     Case "gdatStartDate"
17790       varRetVal = gdatStartDate
17800     Case "glngTaxCode_Distribution"
17810       varRetVal = glngTaxCode_Distribution
17820     End Select
17830   End If

EXITP:
17840   GlobalVarGet = varRetVal
17850   Exit Function

ERRH:
17860   varRetVal = RET_ERR
17870   Select Case ERR.Number
        Case Else
17880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17890   End Select
17900   Resume EXITP

End Function

Public Function AcctNameChk() As Boolean
' ** Check for quotes within Account ShortName and LegalName.

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "AcctNameChk"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

18010   blnRetVal = True

18020   Set dbs = CurrentDb
18030   With dbs
18040     Set rst = .OpenRecordset("account", dbOpenDynaset, dbConsistent)
18050     With rst
18060       If .BOF = True And .EOF = True Then
              ' ** Shouldn't happen!
18070         blnRetVal = False
18080       Else
18090         .MoveLast
18100         lngRecs = .RecordCount
18110         .MoveFirst
18120         For lngX = 1& To lngRecs
18130           If IsNull(![shortname]) = False Then
18140             If Trim(![shortname]) <> vbNullString Then
18150               varTmp00 = FixQuotes(![shortname])  ' ** Function: Below.
18160               If IsNull(varTmp00) = False Then
18170                 If varTmp00 <> ![shortname] Then
18180                   .Edit
18190                   ![shortname] = varTmp00
18200                   .Update
18210                 End If
18220               End If
18230             End If
18240           End If
18250           If IsNull(![legalname]) = False Then
18260             If Trim(![legalname]) <> vbNullString Then
18270               varTmp00 = FixQuotes(![legalname])  ' ** Function: Below.
18280               If IsNull(varTmp00) = False Then
18290                 If varTmp00 <> ![legalname] Then
18300                   .Edit
18310                   ![legalname] = varTmp00
18320                   .Update
18330                 End If
18340               End If
18350             End If
18360           End If
18370           If lngX < lngRecs Then .MoveNext
18380         Next
18390       End If
18400       .Close
18410     End With
18420     .Close
18430   End With

EXITP:
18440   Set rst = Nothing
18450   Set dbs = Nothing
18460   AcctNameChk = blnRetVal
18470   Exit Function

ERRH:
18480   blnRetVal = False
18490   Select Case ERR.Number
        Case Else
18500     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18510   End Select
18520   Resume EXITP

End Function

Public Function SkipKey(ByVal KeyCode As Integer) As Boolean
' ** Returns:
' **   True  : Non-printable character.
' **   False : Letters, numbers, punctuation.
' ** Don't tempt alteration of KeyCode parameter.

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "SkipKey"

        Dim intRetVal As Integer
        Dim blnRetVal As Boolean

18610   blnRetVal = False  ' ** False: Don't skip this key.
18620   intRetVal = KeyCode

18630   Select Case intRetVal
        Case vbKeyBack, vbKeyCancel, vbKeyClear, vbKeyMenu, vbKeyPause, vbKeyExecute
18640     blnRetVal = True
18650   Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageUp, vbKeyPageDown
18660     blnRetVal = True
18670   Case vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8
18680     blnRetVal = True
18690   Case vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12, vbKeyF13, vbKeyF14, vbKeyF15, vbKeyF16
18700     blnRetVal = True
18710   Case vbKeyEscape, vbKeyInsert, vbKeyDelete, vbKeyLButton, vbKeyMButton, vbKeyRButton
18720     blnRetVal = True
18730   Case vbKeyNumlock, vbKeyAdd, vbKeySubtract, vbKeyMultiply, vbKeyDivide, vbKeyDecimal, vbKeySeparator
18740     blnRetVal = True
18750   Case vbKeyControl, vbKeyShift, vbKeySelect, vbKeySnapshot, vbKeyHelp, vbKeyPrint
18760     blnRetVal = True
18770   Case vbKeyTab, vbKeyReturn, 0   ' ** Zero isn't listed, but that's what an Esc throws!
18780     blnRetVal = True
18790   End Select

EXITP:
18800   SkipKey = blnRetVal
18810   Exit Function

ERRH:
18820   blnRetVal = True  ' ** True: Skip this key.
18830   Select Case ERR.Number
        Case Else
18840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18850   End Select
18860   Resume EXITP

End Function

Public Function GetLastWord(varInput As Variant, Optional varMayHaveSfx As Variant, Optional varHowMany As Variant) As Variant
' ** SEE ALSO: GetFirstWord_AP in modPricingFuncs.
' ** SEE ALSO: GetLastWord_AP in modPricingFuncs.

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "GetLastWord"

        Dim blnMayHaveSfx As Boolean
        Dim intHowMany As Integer
        Dim intPos01 As Integer, intLen As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

18910   varRetVal = Null

18920   If IsNull(varInput) = False Then
18930     Select Case IsMissing(varMayHaveSfx)
          Case True
18940       blnMayHaveSfx = False
18950     Case False
18960       blnMayHaveSfx = CBool(varMayHaveSfx)
18970     End Select
18980     Select Case IsMissing(varHowMany)
          Case True
18990       intHowMany = 1
19000     Case False
19010       intHowMany = CInt(varHowMany)
19020     End Select
19030     varRetVal = Trim(CStr(varInput))
19040     strTmp01 = vbNullString: strTmp02 = vbNullString
19050     If InStr(varRetVal, " ") > 0 Then
19060       If blnMayHaveSfx = True Then
19070         intPos01 = InStr(varRetVal, ",")
19080         If intPos01 > 0 Then
19090           strTmp01 = Mid(varRetVal, intPos01)
19100           varRetVal = Left(varRetVal, (intPos01 - 1))
19110         End If
19120         If Right(varRetVal, 3) = " Jr" Or Right(varRetVal, 3) = " Sr" Or _
                  Right(varRetVal, 3) = " II" Then
19130           strTmp02 = Right(varRetVal, 2)
19140           varRetVal = Trim(Left(varRetVal, (Len(varRetVal) - 2)))
19150         End If
19160         If Right(varRetVal, 4) = " Jr." Or Right(varRetVal, 4) = " Sr." Or _
                  Right(varRetVal, 4) = " Esq" Or Right(varRetVal, 4) = " III" Then
19170           strTmp02 = Right(varRetVal, 3)
19180           varRetVal = Trim(Left(varRetVal, (Len(varRetVal) - 3)))
19190         End If
19200         If Right(varRetVal, 5) = " Esq." Then
19210           strTmp02 = Right(varRetVal, 4)
19220           varRetVal = Trim(Left(varRetVal, (Len(varRetVal) - 4)))
19230         End If
19240       End If
19250       intLen = Len(varRetVal)
19260       intCnt = 0
19270       For intX = intLen To 1 Step -1
19280         If Mid(varRetVal, intX, 1) = " " Then
19290           If intHowMany = 1 Then
19300             varRetVal = Mid(varRetVal, (intX + 1))
19310             Exit For
19320           Else
19330             intCnt = intCnt + 1
19340             If intCnt = intHowMany Then
19350               varRetVal = Mid(varRetVal, (intX + 1))
19360               Exit For
19370             End If
19380           End If
19390         End If
19400       Next
19410       If strTmp01 <> vbNullString Then
19420         varRetVal = varRetVal & "~" & strTmp01
19430       ElseIf strTmp02 <> vbNullString Then
19440         varRetVal = varRetVal & "~" & strTmp02
19450       End If
19460     End If
19470   End If

EXITP:
19480   GetLastWord = varRetVal
19490   Exit Function

ERRH:
19500   varRetVal = RET_ERR
19510   Select Case ERR.Number
        Case Else
19520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19530   End Select
19540   Resume EXITP

End Function

Public Function GetNthWord(varInput As Variant, intWord As Integer) As Variant
' ** Get the Nth word from a string containing multiple words.

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "GetNthWord"

        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String
        Dim intX As Integer
        Dim varRetVal As Variant

19610   varRetVal = Null

19620   If IsNull(varInput) = False Then
19630     If Trim(varInput) <> vbNullString Then
19640       strTmp01 = Trim(varInput)
19650       intPos01 = InStr(strTmp01, " ")
19660       If intWord = 1 Then
19670         If intPos01 > 0 Then
19680           varRetVal = Left(strTmp01, (intPos01 - 1))
19690         Else
19700           varRetVal = strTmp01
19710         End If
19720       ElseIf intPos01 > 0 Then
              ' ** 1 space, 2nd word; 2 spaces, 3rd word, 3 spaces, 4th word, etc.
19730         strTmp02 = strTmp01
19740         For intX = 2 To intWord
19750           strTmp02 = Trim(Mid(strTmp02, intPos01))  ' ** strTmp02 now begins with intX's word.
19760           intPos01 = InStr(strTmp02, " ")
19770           If intPos01 = 0 And intX < intWord Then
                  ' ** Too few words.
19780             strTmp02 = vbNullString
19790             Exit For
19800           End If
19810         Next
19820         If strTmp02 <> vbNullString Then
19830           If intPos01 > 0 Then
19840             strTmp02 = Trim(Left(strTmp02, intPos01))
19850           End If
19860           varRetVal = strTmp02
19870         End If
19880       End If
19890     End If
19900   End If

EXITP:
19910   GetNthWord = varRetVal
19920   Exit Function

ERRH:
19930   varRetVal = Null
19940   Select Case ERR.Number
        Case Else
19950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19960   End Select
19970   Resume EXITP

End Function

Public Function GetNthLine(varInput As Variant, intLine As Integer) As Variant
' ** Get the Nth line from a string containing multiple lines.

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "GetNthLine"

        Dim intPos01 As Integer, intLines As Integer
        Dim strTmp01 As String
        Dim varRetVal As Variant

20010   varRetVal = varInput

20020   If IsNull(varInput) = False Then
20030     If Trim(varInput) <> vbNullString Then
20040       strTmp01 = Trim(varInput)
20050       intPos01 = InStr(strTmp01, vbCrLf)
20060       If intPos01 > 0 Then
20070         intLines = (CharCnt(strTmp01, vbCrLf, True) + 1)
20080         If intLine <= intLines Then
20090           If intLine = 1 Then
20100             varRetVal = Left(strTmp01, (intPos01 - 1))
20110           ElseIf intLine = intLines Then
20120             intPos01 = CharPos(strTmp01, (intLines - 1), vbCrLf)
20130             varRetVal = Mid(strTmp01, (intPos01 + 2))
20140           Else
20150             intPos01 = CharPos(strTmp01, (intLine - 1), vbCrLf)  ' ** vbCrLf before line.
20160             strTmp01 = Mid(strTmp01, (intPos01 + 2))
20170             strTmp01 = Left(strTmp01, (InStr(strTmp01, vbCrLf) - 1))
20180             varRetVal = Trim(strTmp01)
20190           End If
20200         Else
20210           varRetVal = Null
20220         End If
20230       End If
20240     End If
20250   End If

EXITP:
20260   GetNthLine = varRetVal
20270   Exit Function

ERRH:
20280   varRetVal = Null
20290   Select Case ERR.Number
        Case Else
20300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20310   End Select
20320   Resume EXITP

End Function

Public Function NullStrIfNull(varInput As Variant) As String
' ** If varInput is null or does not exist, return null string.

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "NullStrIfNull"

        Dim strRetVal As String

20410   strRetVal = vbNullString

20420   If IsNull(varInput) = False Then strRetVal = CStr(varInput)

EXITP:
20430   NullStrIfNull = strRetVal
20440   Exit Function

ERRH:
20450   Select Case ERR.Number
        Case 13  ' ** Type mismatch.
20460     strRetVal = vbNullString
20470   Case 2427  ' ** You entered an expression that has no value.
20480     strRetVal = vbNullString
20490   Case 3021  ' ** No current record.
20500     strRetVal = vbNullString
20510   Case Else
20520     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20530   End Select
20540   Resume EXITP

End Function

Public Function NullIfNullStr(varInput As Variant) As Variant
' ** If varInput is a null string or does not exist, return null.

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "NullIfNullStr"

        Dim varRetVal As Variant

20610   varRetVal = Null

20620   If IsNull(varInput) = False Then
20630     If varInput <> vbNullString Then
20640       varRetVal = varInput
20650     End If
20660   End If

EXITP:
20670   NullIfNullStr = varRetVal
20680   Exit Function

ERRH:
20690   Select Case ERR.Number
        Case 13  ' ** Type mismatch.
20700     varRetVal = Null
20710   Case 2427  ' ** You entered an expression that has no value.
20720     varRetVal = Null
20730   Case 3021  ' ** No current record.
20740     varRetVal = Null
20750   Case Else
20760     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20770   End Select
20780   Resume EXITP

End Function

Public Function ZeroIfNull(varInput As Variant, Optional strCalling As String) As Variant
' ** If varInput is null or does not exist, return zero.

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "ZeroIfNull"

        Dim varRetVal As Variant

20810   varRetVal = 0

20820   If IsNull(varInput) = False Then
20830     If varType(varInput) = vbString Then
20840       If varInput = " " Then varInput = "0"
20850     End If
20860     varRetVal = varInput
20870   End If

EXITP:
20880   ZeroIfNull = varRetVal
20890   Exit Function

ERRH:
20900   Select Case ERR.Number
        Case 13  ' ** Type mismatch.
20910     varRetVal = 0
20920   Case 2427  ' ** You entered an expression that has no value.
20930     varRetVal = 0
20940   Case 3021  ' ** No current record.
20950     varRetVal = 0
20960   Case Else
20970     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20980   End Select
20990   Resume EXITP

End Function

Public Function NullIfZero(varInput As Variant) As Variant
' ** If varInput is 0, return Null.

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "NullIfZero"

        Dim varRetVal As Variant

21010   varRetVal = Null

21020   If IsNull(varInput) = False Then
21030     If varType(varInput) = vbString Then
21040       If varInput <> "0" And varInput <> "0.00" And varInput <> "$0.00" Then
21050         varRetVal = varInput
21060       End If
21070     Else
21080       If varInput <> 0 Then
21090         varRetVal = varInput
21100       End If
21110     End If
21120   End If

EXITP:
21130   NullIfZero = varRetVal
21140   Exit Function

ERRH:
21150   Select Case ERR.Number
        Case 13  ' ** Type mismatch.
21160     varRetVal = Null
21170   Case 2427  ' ** You entered an expression that has no value.
21180     varRetVal = Null
21190   Case 3021  ' ** No current record.
21200     varRetVal = Null
21210   Case Else
21220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21230   End Select
21240   Resume EXITP

End Function

Public Function CharPos(varInput As Variant, lngInst As Long, Optional varChar As Variant) As Long
' ** Return the position of the specified instance of the specified character(s) in the supplied string.

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "CharPos"

        Dim strChar As String
        Dim lngLen01 As Long, lngLen02 As Long, lngCnt As Long
        Dim strTmp01 As String
        Dim lngX As Long
        Dim lngRetVal As Long

21310   lngRetVal = 0&
21320   strTmp01 = vbNullString: lngCnt = 0&

21330   If IsMissing(varChar) = True Then
21340     strChar = ","  ' ** Default to a comma.
21350   Else
21360     If IsNull(varChar) = False Then
21370       strChar = varChar
21380     Else
21390       strChar = ","  ' ** Default to a comma.
21400     End If
21410   End If

21420   If IsNull(varInput) = False Then
21430     strTmp01 = Trim(varInput)
21440     If strTmp01 <> vbNullString Then
21450       lngLen01 = Len(strTmp01)
21460       lngLen02 = Len(strChar)
21470       If lngLen02 = 1 Then
21480         For lngX = 1& To lngLen01
21490           If Mid(strTmp01, lngX, 1) = strChar Then
21500             lngCnt = lngCnt + 1&
21510             If lngCnt = lngInst Then
21520               lngRetVal = lngX
21530               Exit For
21540             End If
21550           End If
21560         Next
21570       Else
21580         For lngX = 1& To lngLen01
21590           If Mid(strTmp01, lngX, lngLen02) = strChar Then
21600             lngCnt = lngCnt + 1
21610             If lngCnt = lngInst Then
21620               lngRetVal = lngX
21630               Exit For
21640             End If
21650           End If
21660         Next
21670       End If
21680     End If
21690   End If

EXITP:
21700   CharPos = lngRetVal
21710   Exit Function

ERRH:
21720   lngRetVal = 0&
21730   Select Case ERR.Number
        Case Else
21740     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21750   End Select
21760   Resume EXITP

End Function

Public Function CharCnt(varInput As Variant, strChar As String, Optional varMultiChar As Variant) As Long
' ** Count the occurrences of the specified character or characters in the supplied string.

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "CharCnt"

        Dim blnMultiChar As Boolean
        Dim lngLen As Long
        Dim strTmp01 As String
        Dim lngX As Long
        Dim lngRetVal As Long

21810   lngRetVal = 0&
21820   strTmp01 = vbNullString

21830   Select Case IsMissing(varMultiChar)
        Case True
21840     blnMultiChar = False
21850   Case False
21860     blnMultiChar = CBool(varMultiChar)
21870   End Select
        'If blnMultiChar = True Then
21880   If IsNull(varInput) = False Then
21890     strTmp01 = Trim(varInput)
21900     If strTmp01 <> vbNullString Then
21910       lngLen = Len(strTmp01)
21920       For lngX = 1& To lngLen
21930         Select Case blnMultiChar
              Case True
21940           If Mid(strTmp01, lngX, Len(strChar)) = strChar Then
21950             lngRetVal = lngRetVal + 1&
21960           End If
21970         Case False
21980           If Mid(strTmp01, lngX, 1) = strChar Then
21990             lngRetVal = lngRetVal + 1&
22000           End If
22010         End Select
22020       Next
22030     End If
22040   End If
        'End If

EXITP:
22050   CharCnt = lngRetVal
22060   Exit Function

ERRH:
22070   lngRetVal = 0&
22080   Select Case ERR.Number
        Case Else
22090     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22100   End Select
22110   Resume EXITP

End Function

Public Function Parse_GfxImg(varInput As Variant, Optional varName As Variant, Optional varSep As Variant) As Variant
' ** Parameters:
' **   varInput: Field name behind control, user-defined from a query.
' **   varName:  Control Tag, with special alternate name.
' **   varSep:   Whether the name separates the graphic description.
' **
' ** Used by:
' **   qryJournal_Columns_15 : [ctl_source]
' **     Sort1: Parse_GfxImg([ctl_source],[ctlspec_tag],True)
' **   qryJournal_Columns_16 : [xadgfx_name]
' **     Sort1: Parse_GfxImg([xadgfx_name],'',False)

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "Parse_GfxImg"

        Dim strReplacementName As String, blnName As Boolean, blnSep As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer
        Dim strTmp01 As String
        Dim intX As Integer
        Dim varRetVal As Variant
        Dim R_NOFOC As String
        Dim R_SFOC  As String
        Dim R_FOC   As String
        Dim R_DIS   As String
        Dim R_DOTS  As String
        Dim S_NOFOC As String
        Dim S_FOC   As String
        Dim NOSHORT As String

        'THE CONTROL SEPARATES RAISED_FOCUS!
        'THE IMAGE DOESN'T SEPARATE RAISEDFOCUS!
22210   varRetVal = Null

22220   If IsNull(varInput) = False Then
22230     If Trim(varInput) <> vbNullString Then

22240       strReplacementName = vbNullString
22250       blnName = False
22260       If IsMissing(varName) = False Then
22270         If Trim(varName) <> vbNullString Then
22280           strReplacementName = varName
22290           blnName = True
22300         End If
22310       End If

22320       Select Case IsMissing(varSep)
            Case True
22330         blnSep = True
22340       Case False
22350         blnSep = CBool(varSep)
22360       End Select

22370       Select Case blnSep
            Case True
22380         R_NOFOC = "_raised"
22390         R_SFOC = "_raised_semifocus"
22400         R_FOC = "_raised_focus"
22410         R_DIS = "_dis"
22420         R_DOTS = "_dots"
22430         S_NOFOC = "_sunken"
22440         S_FOC = "_sunken_focus"
22450         NOSHORT = "_no_shortcut"
22460       Case False
22470         R_NOFOC = "_raised"
22480         R_SFOC = "_raisedsemifocus"
22490         R_FOC = "_raisedfocus"
22500         R_DIS = "disabled"
22510         R_DOTS = "dots"
22520         S_NOFOC = "_sunken"
22530         S_FOC = "_sunkenfocus"
22540         NOSHORT = "_no_shortcut"
22550       End Select

22560       strTmp01 = Trim(varInput)

            ' ** Raised Focus.
22570       intPos02 = InStr(strTmp01, R_FOC)  ' ** intPos02 is an underscore.
22580       intPos03 = 0
22590       If intPos02 > 0 Then
22600         intPos03 = InStr(strTmp01, R_DOTS)
22610         intPos01 = 0
              ' ** Find next-previous underscore.
22620         For intX = (intPos02 - 1) To 1 Step -1
22630           If Mid(strTmp01, intX, 1) = "_" Then
22640             intPos01 = intX  ' ** intPos01 is previous underscore.
22650             Exit For
22660           End If
22670         Next
22680         If intPos01 > 0 Then
22690           varRetVal = Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1))
22700           varRetVal = varRetVal & R_FOC
22710         Else
22720           varRetVal = Left(strTmp01, (intPos02 - 1))
22730           varRetVal = varRetVal & R_FOC
22740         End If
22750         If intPos03 > 0 Then varRetVal = varRetVal & R_DOTS
22760         If blnName = True Then
22770           If strReplacementName <> "No_Shortcut" And strReplacementName <> "Classic" And strReplacementName <> "Columnar" Then
22780             varRetVal = strReplacementName & R_FOC
22790             If intPos03 > 0 Then varRetVal = varRetVal & R_DOTS
22800           End If
22810         End If
22820       End If

            ' ** Raised Semi-Focus.
22830       intPos02 = InStr(strTmp01, R_SFOC)  ' ** intPos02 is an underscore.
22840       intPos03 = 0
22850       If intPos02 > 0 Then
22860         intPos03 = InStr(strTmp01, R_DOTS)
22870         intPos01 = 0
              ' ** Find next-previous underscore.
22880         For intX = (intPos02 - 1) To 1 Step -1
22890           If Mid(strTmp01, intX, 1) = "_" Then
22900             intPos01 = intX  ' ** intPos01 is previous underscore.
22910             Exit For
22920           End If
22930         Next
22940         If intPos01 > 0 Then
22950           varRetVal = Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1))
22960           varRetVal = varRetVal & R_SFOC
22970         Else
22980           varRetVal = Left(strTmp01, (intPos02 - 1))
22990           varRetVal = varRetVal & R_SFOC
23000         End If
23010         If intPos03 > 0 Then varRetVal = varRetVal & R_DOTS
23020         If blnName = True Then
23030           If strReplacementName <> "No_Shortcut" And strReplacementName <> "Classic" And strReplacementName <> "Columnar" Then
23040             varRetVal = strReplacementName & R_SFOC
23050             If intPos03 > 0 Then varRetVal = varRetVal & R_DOTS
23060           End If
23070         End If
23080       End If

            ' ** Sunken Focus.
23090       If IsNull(varRetVal) = True Then
23100         intPos02 = InStr(strTmp01, S_FOC)
              ' ** Dots don't matter.
23110         If intPos02 > 0 Then
23120           intPos01 = 0
                ' ** Find next-previous underscore.
23130           For intX = (intPos02 - 1) To 1 Step -1
23140             If Mid(strTmp01, intX, 1) = "_" Then
23150               intPos01 = intX  ' ** intPos01 is previous underscore.
23160               Exit For
23170             End If
23180           Next
23190           If intPos01 > 0 Then
23200             varRetVal = Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1))
23210             varRetVal = varRetVal & S_FOC
23220           Else
23230             varRetVal = Left(strTmp01, (intPos02 - 1))
23240             varRetVal = varRetVal & S_FOC
23250           End If
23260           If blnName = True Then
23270             If strReplacementName <> "No_Shortcut" And strReplacementName <> "Classic" And strReplacementName <> "Columnar" Then
23280               varRetVal = strReplacementName & S_FOC
23290             End If
23300           End If
23310           If InStr(varRetVal, "Scroll") > 0 Then
23320             If InStr(varInput, "dots") > 0 And InStr(varRetVal, "dots") = 0 Then
23330               varRetVal = varRetVal & "_dots"
23340             End If
23350           End If
23360         End If
23370       End If

            ' ** Raised No Focus.
23380       If IsNull(varRetVal) = True Then
23390         intPos02 = InStr(strTmp01, R_NOFOC)
23400         If intPos02 > 0 Then
23410           intPos01 = 0
                ' ** Find next-previous underscore.
23420           For intX = (intPos02 - 1) To 1 Step -1
23430             If Mid(strTmp01, intX, 1) = "_" Then
23440               intPos01 = intX  ' ** intPos01 is previous underscore.
23450               Exit For
23460             End If
23470           Next
23480           If intPos01 > 0 Then
23490             varRetVal = Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1))
23500             varRetVal = varRetVal & R_NOFOC
23510           Else
23520             varRetVal = Left(strTmp01, (intPos02 - 1))
23530             varRetVal = varRetVal & R_NOFOC
23540           End If
23550           If blnName = True Then
23560             If strReplacementName <> "No_Shortcut" And strReplacementName <> "Classic" And strReplacementName <> "Columnar" Then
23570               varRetVal = strReplacementName & R_NOFOC
23580             End If
23590           End If
23600         End If
23610       End If

            'THIS SHOULDN'T EXIST!
            ' ** Sunken No Focus.
23620       If IsNull(varRetVal) = True Then
23630         intPos02 = InStr(strTmp01, S_NOFOC)
23640         If intPos02 > 0 Then
23650           intPos01 = 0
                ' ** Find next-previous underscore.
23660           For intX = (intPos02 - 1) To 1 Step -1
23670             If Mid(strTmp01, intX, 1) = "_" Then
23680               intPos01 = intX  ' ** intPos01 is previous underscore.
23690               Exit For
23700             End If
23710           Next
23720           If intPos01 > 0 Then
23730             varRetVal = Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1))
23740             varRetVal = varRetVal & S_NOFOC
23750           Else
23760             varRetVal = Left(strTmp01, (intPos02 - 1))
23770             varRetVal = varRetVal & S_NOFOC
23780           End If
23790           If blnName = True Then
23800             If strReplacementName <> "No_Shortcut" And strReplacementName <> "Classic" And strReplacementName <> "Columnar" Then
23810               varRetVal = strReplacementName & S_NOFOC
23820             End If
23830           End If
23840         End If
23850       End If

            ' ** Raised Disabled.
23860       If IsNull(varRetVal) = True And Left(R_DIS, 1) = "_" Then
23870         intPos02 = InStr(strTmp01, R_DIS)
23880         If intPos02 > 0 Then
23890           intPos01 = 0
                ' ** Find next-previous underscore.
23900           For intX = (intPos02 - 1) To 1 Step -1
23910             If Mid(strTmp01, intX, 1) = "_" Then
23920               intPos01 = intX  ' ** intPos01 is previous underscore.
23930               Exit For
23940             End If
23950           Next
23960           If intPos01 > 0 Then
23970             varRetVal = Mid(strTmp01, (intPos01 + 1), ((intPos02 - intPos01) - 1))
23980             varRetVal = varRetVal & R_DIS
23990           Else
24000             varRetVal = Left(strTmp01, (intPos02 - 1))
24010             varRetVal = varRetVal & R_DIS
24020           End If
24030           If blnName = True Then
24040             If strReplacementName <> "No_Shortcut" And strReplacementName <> "Classic" And strReplacementName <> "Columnar" Then
24050               varRetVal = strReplacementName & R_DIS
24060             End If
24070           End If
24080         End If
24090       ElseIf IsNull(varRetVal) = True And Left(R_DIS, 1) <> "_" Then
24100         intPos02 = InStr(strTmp01, R_DIS)
24110         If intPos02 > 0 Then
24120           varRetVal = Left(strTmp01, (intPos02 - 1)) & "_dis"
24130           intPos01 = InStr(varRetVal, "__")
24140           If intPos01 > 0 Then
24150             varRetVal = Left(varRetVal, intPos01) & "raised" & Mid(varRetVal, (intPos01 + 1))
24160           End If
24170         End If
24180       ElseIf InStr(strTmp01, R_DIS) > 0 Then
24190         varRetVal = varRetVal & "_dis"
24200       End If

            ' ** No Shortcut.
24210       intPos02 = InStr(strTmp01, NOSHORT)
            'If intPos02 > 0 Or strReplacementName = "No_Shortcut" Then
            '  varRetVal = varRetVal & "_ns"
            'End If

24220       If IsNull(varRetVal) = False Then

24230         If Left(varRetVal, 3) = "cmd" Then
24240           varRetVal = Mid(varRetVal, 4)
24250         End If

24260         If Left(varRetVal, 4) = "lbl_" Then
24270           If InStr(varInput, "Scroll") > 0 Then
24280             varRetVal = "Scroll_" & Mid(varRetVal, 5)
24290           ElseIf InStr(varInput, "_Sub") > 0 Then
24300             varRetVal = "Transactions_" & Mid(varRetVal, 5)
24310           End If
24320         End If

              ' ** Transactions subform label.
24330         If Left(varRetVal, 13) = "Transactions_" Then
                'If InStr(strTmp01, "_Bge") > 0 Or InStr(strTmp01, "_des") > 0 Then
                '  varRetVal = varRetVal & "_des"
                'ElseIf InStr(strTmp01, "_Gry") > 0 Or InStr(strTmp01, "_std") > 0 Then
                '  varRetVal = varRetVal & "_std"
                'End If
24340         End If

              ' ** Raised Focus.
24350         intPos01 = InStr(varRetVal, "_raisedfocus")
24360         intPos03 = 0
24370         If intPos01 > 0 Then
24380           intPos03 = InStr(varRetVal, "dots")
24390           strTmp01 = Left(varRetVal, (intPos01 - 1)) & "_raised_focus"
24400           If intPos01 + Len("_raisedfocus") > Len(varRetVal) Then
                  ' ** At the end of the input.
24410           Else
24420             strTmp01 = strTmp01 & Mid(varRetVal, (intPos01 + Len("_raisedfocus")))
24430           End If
24440           varRetVal = strTmp01
24450           If intPos03 > 0 Then
24460             varRetVal = Left(varRetVal, intPos03) & "_" & Mid(varRetVal, (intPos03 + 1))
24470           End If
24480         End If

              ' ** Raised Semi-Focus.
24490         intPos01 = InStr(varRetVal, "_raisedsemifocus")
24500         intPos03 = 0
24510         If intPos01 > 0 Then
24520           intPos03 = InStr(varRetVal, "dots")
24530           strTmp01 = Left(varRetVal, (intPos01 - 1)) & "_raised_semifocus"
24540           If intPos01 + Len("_raisedfocus") > Len(varRetVal) Then
                  ' ** At the end of the input.
24550           Else
24560             strTmp01 = strTmp01 & Mid(varRetVal, (intPos01 + Len("_raisedsemifocus")))
24570           End If
24580           varRetVal = strTmp01
24590           If intPos03 > 0 Then
24600             varRetVal = Left(varRetVal, intPos03) & "_" & Mid(varRetVal, (intPos03 + 1))
24610           End If
24620         End If

              ' ** Sunken Focus.
24630         intPos01 = InStr(varRetVal, "_sunkenfocus")
              ' ** Dots don't matter.
24640         If intPos01 > 0 Then
24650           strTmp01 = Left(varRetVal, (intPos01 - 1)) & "_sunken_focus"
24660           If intPos01 + Len("_sunkenfocus") > Len(varRetVal) Then
                  ' ** At the end of the input.
24670           Else
24680             strTmp01 = strTmp01 & Mid(varRetVal, (intPos01 + Len("_sunkenfocus")))
24690           End If
24700           varRetVal = strTmp01
24710         End If
24720         If varInput = "ScrollRight_Sunken_Blu" Then
24730           varRetVal = "ScrollRight_sunken_focus_dots"
24740         ElseIf varInput = "ScrollLeft_Sunken_Blu" Then
24750           varRetVal = "ScrollLeft_sunken_focus_dots"
24760         End If

24770         If Left(varRetVal, 7) = "Switch_" And (strReplacementName = "Columnar" Or strReplacementName = "Classic") Then
24780           varRetVal = "Switch" & strReplacementName & Mid(varRetVal, 7)
24790         End If

24800       End If

24810     End If
24820   End If

EXITP:
24830   Parse_GfxImg = varRetVal
24840   Exit Function

ERRH:
24850   varRetVal = RET_ERR
24860   Select Case ERR.Number
        Case Else
24870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24880   End Select
24890   Resume EXITP

End Function

Public Function TwipsToInches(varInput As Variant) As String
' ** Convert Twips to Inches, at 3 decimal places.
' ** ADD OPTION FOR 4 DECIMAL PLACES!!

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "TwipsToInches"

        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String, intTmp03 As Integer, dblTmp04 As Double
        Dim strRetVal As String

24910   strRetVal = vbNullString

24920   If IsNull(varInput) = False Then
24930     dblTmp04 = varInput / 1440#
24940     strTmp01 = CStr(dblTmp04)
24950     intPos01 = InStr(strTmp01, ".")
24960     If intPos01 > 0 Then
24970       strTmp02 = Mid(strTmp01, (intPos01 + 1))  ' ** Just the decimal portion.
24980       If Len(strTmp02) > 3 Then
24990         If Val(Mid(strTmp02, 4, 1)) >= 5 Then
25000           intTmp03 = Val(Mid(strTmp02, 3, 1))
25010           If intTmp03 = 9 Then
25020             If Val(Mid(strTmp02, 2, 1)) = 9 Then
25030               If Val(Left(strTmp02, 1)) = 9 Then
25040                 strTmp02 = CStr(Val(Left(strTmp01, (intPos01 - 1))) + 1) & ".000"
25050                 strRetVal = strTmp02 & Chr(34)
25060               Else
25070                 strTmp02 = CStr(Val(Left(strTmp02, 1)) + 1) & "00"
25080                 strRetVal = Left(strTmp01, intPos01) & strTmp02 & Chr(34)
25090               End If
25100             Else
25110               strTmp02 = Left(strTmp01, 1) & CStr(Val(Mid(strTmp02, 2, 1)) + 1) & "0"
25120               strRetVal = Left(strTmp01, intPos01) & strTmp02 & Chr(34)
25130             End If
25140           Else
25150             intTmp03 = intTmp03 + 1
25160             strTmp02 = Left(strTmp02, 2) & CStr(intTmp03)
25170             strRetVal = Left(strTmp01, intPos01) & strTmp02 & Chr(34)
25180           End If
25190         Else
25200           strTmp02 = Left(strTmp02, 3)
25210           strRetVal = Left(strTmp01, intPos01) & strTmp02 & Chr(34)
25220         End If
25230       Else
25240         strTmp02 = Left(strTmp02 & "000", 3)
25250         strRetVal = Left(strTmp01, intPos01) & strTmp02 & Chr(34)
25260       End If
25270     Else
25280       strRetVal = strTmp01 & ".000" & Chr(34)
25290     End If
25300   End If

EXITP:
25310   TwipsToInches = strRetVal
25320   Exit Function

ERRH:
25330   strRetVal = RET_ERR
25340   Select Case ERR.Number
        Case Else
25350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25360   End Select
25370   Resume EXITP

End Function

Public Function GetMax(varVal01 As Variant, varVal02 As Variant, Optional varVal03 As Variant, Optional varVal04 As Variant, Optional varVal05 As Variant, Optional varVal06 As Variant, Optional varVal07 As Variant, Optional varVal08 As Variant, Optional varVal09 As Variant, Optional varVal10 As Variant) As Variant
' ** Return the maximum value in a series of supplied numbers.

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "GetMax"

        Dim lngVals As Long
        Dim lngX As Long, lngE As Long
        Dim varRetVal As Variant

        ' ** Array: arr_varVal().
        Const V_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const V_VAL As Integer = 0

25410   lngVals = 0&
25420   ReDim arr_varVal(V_ELEMS, 0)

25430   If IsNull(varVal01) = False Then
25440     If IsNumeric(varVal01) = True Then
25450       lngVals = lngVals + 1&
25460       lngE = lngVals - 1&
25470       ReDim Preserve arr_varVal(V_ELEMS, lngE)
25480       arr_varVal(V_VAL, lngE) = varVal01
25490     End If
25500   End If

25510   If IsNull(varVal02) = False Then
25520     If IsNumeric(varVal02) = True Then
25530       lngVals = lngVals + 1&
25540       lngE = lngVals - 1&
25550       ReDim Preserve arr_varVal(V_ELEMS, lngE)
25560       arr_varVal(V_VAL, lngE) = varVal02
25570     End If
25580   End If

25590   If IsMissing(varVal03) = False Then
25600     If IsNull(varVal03) = False Then
25610       If IsNumeric(varVal03) = True Then
25620         lngVals = lngVals + 1&
25630         lngE = lngVals - 1&
25640         ReDim Preserve arr_varVal(V_ELEMS, lngE)
25650         arr_varVal(V_VAL, lngE) = varVal03
25660       End If
25670     End If
25680   End If
25690   If IsMissing(varVal04) = False Then
25700     If IsNull(varVal04) = False Then
25710       If IsNumeric(varVal04) = True Then
25720         lngVals = lngVals + 1&
25730         lngE = lngVals - 1&
25740         ReDim Preserve arr_varVal(V_ELEMS, lngE)
25750         arr_varVal(V_VAL, lngE) = varVal04
25760       End If
25770     End If
25780   End If
25790   If IsMissing(varVal05) = False Then
25800     If IsNull(varVal05) = False Then
25810       If IsNumeric(varVal05) = True Then
25820         lngVals = lngVals + 1&
25830         lngE = lngVals - 1&
25840         ReDim Preserve arr_varVal(V_ELEMS, lngE)
25850         arr_varVal(V_VAL, lngE) = varVal05
25860       End If
25870     End If
25880   End If
25890   If IsMissing(varVal06) = False Then
25900     If IsNull(varVal06) = False Then
25910       If IsNumeric(varVal06) = True Then
25920         lngVals = lngVals + 1&
25930         lngE = lngVals - 1&
25940         ReDim Preserve arr_varVal(V_ELEMS, lngE)
25950         arr_varVal(V_VAL, lngE) = varVal06
25960       End If
25970     End If
25980   End If
25990   If IsMissing(varVal07) = False Then
26000     If IsNull(varVal07) = False Then
26010       If IsNumeric(varVal07) = True Then
26020         lngVals = lngVals + 1&
26030         lngE = lngVals - 1&
26040         ReDim Preserve arr_varVal(V_ELEMS, lngE)
26050         arr_varVal(V_VAL, lngE) = varVal07
26060       End If
26070     End If
26080   End If
26090   If IsMissing(varVal08) = False Then
26100     If IsNull(varVal08) = False Then
26110       If IsNumeric(varVal08) = True Then
26120         lngVals = lngVals + 1&
26130         lngE = lngVals - 1&
26140         ReDim Preserve arr_varVal(V_ELEMS, lngE)
26150         arr_varVal(V_VAL, lngE) = varVal08
26160       End If
26170     End If
26180   End If
26190   If IsMissing(varVal09) = False Then
26200     If IsNull(varVal09) = False Then
26210       If IsNumeric(varVal09) = True Then
26220         lngVals = lngVals + 1&
26230         lngE = lngVals - 1&
26240         ReDim Preserve arr_varVal(V_ELEMS, lngE)
26250         arr_varVal(V_VAL, lngE) = varVal09
26260       End If
26270     End If
26280   End If
26290   If IsMissing(varVal10) = False Then
26300     If IsNull(varVal10) = False Then
26310       If IsNumeric(varVal10) = True Then
26320         lngVals = lngVals + 1&
26330         lngE = lngVals - 1&
26340         ReDim Preserve arr_varVal(V_ELEMS, lngE)
26350         arr_varVal(V_VAL, lngE) = varVal10
26360       End If
26370     End If
26380   End If

26390   If lngVals > 0& Then
26400     varRetVal = arr_varVal(V_VAL, 0)
26410     For lngX = 1& To (lngVals - 1&)
26420       If arr_varVal(V_VAL, lngX) > varRetVal Then
26430         varRetVal = arr_varVal(V_VAL, lngX)
26440       End If
26450     Next
26460   Else
26470     varRetVal = 0
26480   End If

EXITP:
26490   GetMax = varRetVal
26500   Exit Function

ERRH:
26510   varRetVal = 0
26520   Select Case ERR.Number
        Case Else
26530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26540   End Select
26550   Resume EXITP

End Function

Public Function GetMin(varVal01 As Variant, varVal02 As Variant, Optional varVal03 As Variant, Optional varVal04 As Variant, Optional varVal05 As Variant, Optional varVal06 As Variant, Optional varVal07 As Variant, Optional varVal08 As Variant, Optional varVal09 As Variant, Optional varVal10 As Variant) As Variant
' ** Return the minimum value in a series of supplied numbers.

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "GetMin"

        Dim lngVals As Long
        Dim lngX As Long, lngE As Long
        Dim varRetVal As Variant

        ' ** Array: arr_varVal().
        Const V_ELEMS As Integer = 0  ' ** Array's first-element UBound().
        Const V_VAL As Integer = 0

26610   lngVals = 0&
26620   ReDim arr_varVal(V_ELEMS, 0)

26630   If IsNull(varVal01) = False Then
26640     If IsNumeric(varVal01) = True Then
26650       lngVals = lngVals + 1&
26660       lngE = lngVals - 1&
26670       ReDim Preserve arr_varVal(V_ELEMS, lngE)
26680       arr_varVal(V_VAL, lngE) = varVal01
26690     End If
26700   End If

26710   If IsNull(varVal02) = False Then
26720     If IsNumeric(varVal02) = True Then
26730       lngVals = lngVals + 1&
26740       lngE = lngVals - 1&
26750       ReDim Preserve arr_varVal(V_ELEMS, lngE)
26760       arr_varVal(V_VAL, lngE) = varVal02
26770     End If
26780   End If

26790   If IsMissing(varVal03) = False Then
26800     If IsNull(varVal03) = False Then
26810       If IsNumeric(varVal03) = True Then
26820         lngVals = lngVals + 1&
26830         lngE = lngVals - 1&
26840         ReDim Preserve arr_varVal(V_ELEMS, lngE)
26850         arr_varVal(V_VAL, lngE) = varVal03
26860       End If
26870     End If
26880   End If
26890   If IsMissing(varVal04) = False Then
26900     If IsNull(varVal04) = False Then
26910       If IsNumeric(varVal04) = True Then
26920         lngVals = lngVals + 1&
26930         lngE = lngVals - 1&
26940         ReDim Preserve arr_varVal(V_ELEMS, lngE)
26950         arr_varVal(V_VAL, lngE) = varVal04
26960       End If
26970     End If
26980   End If
26990   If IsMissing(varVal05) = False Then
27000     If IsNull(varVal05) = False Then
27010       If IsNumeric(varVal05) = True Then
27020         lngVals = lngVals + 1&
27030         lngE = lngVals - 1&
27040         ReDim Preserve arr_varVal(V_ELEMS, lngE)
27050         arr_varVal(V_VAL, lngE) = varVal05
27060       End If
27070     End If
27080   End If
27090   If IsMissing(varVal06) = False Then
27100     If IsNull(varVal06) = False Then
27110       If IsNumeric(varVal06) = True Then
27120         lngVals = lngVals + 1&
27130         lngE = lngVals - 1&
27140         ReDim Preserve arr_varVal(V_ELEMS, lngE)
27150         arr_varVal(V_VAL, lngE) = varVal06
27160       End If
27170     End If
27180   End If
27190   If IsMissing(varVal07) = False Then
27200     If IsNull(varVal07) = False Then
27210       If IsNumeric(varVal07) = True Then
27220         lngVals = lngVals + 1&
27230         lngE = lngVals - 1&
27240         ReDim Preserve arr_varVal(V_ELEMS, lngE)
27250         arr_varVal(V_VAL, lngE) = varVal07
27260       End If
27270     End If
27280   End If
27290   If IsMissing(varVal08) = False Then
27300     If IsNull(varVal08) = False Then
27310       If IsNumeric(varVal08) = True Then
27320         lngVals = lngVals + 1&
27330         lngE = lngVals - 1&
27340         ReDim Preserve arr_varVal(V_ELEMS, lngE)
27350         arr_varVal(V_VAL, lngE) = varVal08
27360       End If
27370     End If
27380   End If
27390   If IsMissing(varVal09) = False Then
27400     If IsNull(varVal09) = False Then
27410       If IsNumeric(varVal09) = True Then
27420         lngVals = lngVals + 1&
27430         lngE = lngVals - 1&
27440         ReDim Preserve arr_varVal(V_ELEMS, lngE)
27450         arr_varVal(V_VAL, lngE) = varVal09
27460       End If
27470     End If
27480   End If
27490   If IsMissing(varVal10) = False Then
27500     If IsNull(varVal10) = False Then
27510       If IsNumeric(varVal10) = True Then
27520         lngVals = lngVals + 1&
27530         lngE = lngVals - 1&
27540         ReDim Preserve arr_varVal(V_ELEMS, lngE)
27550         arr_varVal(V_VAL, lngE) = varVal10
27560       End If
27570     End If
27580   End If

27590   If lngVals > 0& Then
27600     varRetVal = arr_varVal(V_VAL, 0)
27610     For lngX = 1& To (lngVals - 1&)
27620       If arr_varVal(V_VAL, lngX) < varRetVal Then
27630         varRetVal = arr_varVal(V_VAL, lngX)
27640       End If
27650     Next
27660   Else
27670     varRetVal = 0
27680   End If

EXITP:
27690   GetMin = varRetVal
27700   Exit Function

ERRH:
27710   varRetVal = 0
27720   Select Case ERR.Number
        Case Else
27730     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27740   End Select
27750   Resume EXITP

End Function
