Attribute VB_Name = "modCodeUtilities"
Option Compare Database
Option Explicit

'VGC 03/23/2017: CHANGES!

' *******************************************************************************
' ** SRSEncrypt.dll was written and compiled by Duane Johnson, Minneapolis, MN
' *******************************************************************************

'? CalcRegCode("01/01/2010")
' 29886 2010  Begins repeating!
' 29902 2009
' 29900 2008
' 29898 2007
' 29896 2006
' 29894 2005
' 29892 2004
' 29890 2003
' 29888 2002
' 29886 2001
' 29884 2000

'? CalcRegCode("05/25/2010")
' 29906 2010  Begins repeating!
' 29922 2009
' 29920 2008
' 29918 2007
' 29916 2006
' 29914 2005
' 29912 2004
' 29910 2003
' 29908 2002
' 29906 2001
' 29904 2000

Private Const THIS_NAME As String = "modCodeUtilities"
' **

Public Function CalcRegCode(strDate As String) As Long
' ** Parameters:
' **   strDate = frmLicense.txtPricingExpires
' **   strDate = frmLicense_Edit.txtPricingExpires
' ** Return:
' **   Val(frmLicense.txtRegCode)      {must equal}  lngRetVal
' **   Val(frmLicense_Edit.txtRegCode)  {must equal}  lngRetVal
' ---------------------------------------------------------------------------
' -- To calculate the registration code
' -- 1) Add the numbers in the date string (like: 12/31/2001 => 1+2+3+1+2+0+0+1 => 10)
' -- 2) Double the total
' -- 3) Add 29876 to the total, simply to make it harder to decipher
' ---------------------------------------------------------------------------

100   On Error GoTo ERRH

        Const THIS_PROC As String = "CalcRegCode"

        Dim lngTotal As Long
        Dim strTmp01 As String
        Dim intX As Integer
        Dim lngRetVal As Long

110     lngRetVal = 0&
120     lngTotal = 0&
130     strTmp01 = strDate

        ' ** Add the numbers in the date string.
140     For intX = 1 To Len(strDate)
150       If Asc(Mid(strTmp01, intX, 1)) > Asc("/") And Asc(Mid(strTmp01, intX, 1)) < Asc(":") Then
160         lngTotal = lngTotal + Val(Mid(strTmp01, intX, 1))
170       End If
180     Next
        ' ** Double the total.
190     lngTotal = lngTotal * 2
        ' ** Add 29876 to the total, simply to make it harder to decipher.
200     lngRetVal = lngTotal + 29876

EXITP:
210     CalcRegCode = lngRetVal
220     Exit Function

ERRH:
230     Select Case ERR.Number
        Case Else
240       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
250     End Select
260     Resume EXITP

End Function

Public Function CalcLimitCode(strInput As String) As String
' ** Parameters:
' **   strInput = frmLicense.txtLimit
' ** Doesn't compare, just calculates.

300   On Error GoTo ERRH

        Const THIS_PROC As String = "CalcLimitCode"

        Dim strTime As String
        Dim strTmp01 As String
        Dim strRetVal As String

310     strRetVal = vbNullString

        ' ** Calc code.
320     If Len(strInput) > 6 Then
          'MsgBox "Invalid number", vbExclamation, "Error"
330       strRetVal = vbNullString
340     Else
350       strTmp01 = strPadZ(strInput, 6)
360       strTime = Format(time, "Hh:Nn:Ss")
370       strRetVal = Mid(strTime, 5, 1) & _
            Mid(strTime, 8, 1) & _
            Mid(strTmp01, 5, 1) & _
            Mid(strTmp01, 4, 1) & _
            Mid(strTime, 2, 1) & _
            Mid(strTmp01, 1, 1) & _
            Mid(strTime, 7, 1) & _
            Mid(strTmp01, 6, 1) & _
            Mid(strTmp01, 2, 1) & _
            Mid(strTime, 4, 1) & _
            Mid(strTmp01, 3, 1) & _
            Mid(strTmp01, 6, 1)
380     End If

EXITP:
390     CalcLimitCode = strRetVal
400     Exit Function

ERRH:
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

End Function

Public Function CalcLimitNumber(strInput As String) As String
' ** Parameters:
' **   strInput = frmLicense.txtLicenseCode
' **   strInput = frmLicense_Edit.txtLicenseCode
' ** Return:
' **   frmLicense.txtLimit      {must equal}  strRetVal
' **   frmLicense_Edit.txtLimit  {must equal}  strRetVal

500   On Error GoTo ERRH

        Const THIS_PROC As String = "CalcLimitNumber"

        Dim strTmp01 As String
        Dim strRetVal As String

510     strRetVal = vbNullString

520     If Len(strInput) <> 12 Then
          'MsgBox "Invalid code", vbExclamation, "Error"
530       strRetVal = vbNullString
540     Else
550       strTmp01 = Mid(strInput, 6, 1) & _
            Mid(strInput, 9, 1) & _
            Mid(strInput, 11, 1) & _
            Mid(strInput, 4, 1) & _
            Mid(strInput, 3, 1) & _
            Mid(strInput, 8, 1)
          ' ** Remove leading zeroes.
560       strRetVal = CStr(Val(strTmp01))
570     End If

EXITP:
580     CalcLimitNumber = strRetVal
590     Exit Function

ERRH:
600     Select Case ERR.Number
        Case Else
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End Select
630     Resume EXITP

End Function

Public Function CalcPriceCode(strDate As String) As Long
' ** Parameters:
' **   strDate = frmLicense.txtPricingExpires
' **   strDate = frmLicense_Edit.txtPricingExpires
' ** Return:
' **   Val(frmLicense.txtPricingCode)      {must equal}  lngRetVal
' **   Val(frmLicense_Edit.txtPricingCode)  {must equal}  lngRetVal
' ---------------------------------------------------------------------------
' -- To calculate the registration code
' -- 1) Add the numbers in the date string (like: 12/31/2001 => 1+2+3+1+2+0+0+1 => 10)
' -- 2) Double the total
' -- 3) Add 17531 to the total, simply to make it harder to decipher
' ---------------------------------------------------------------------------

700   On Error GoTo ERRH

        Const THIS_PROC As String = "CalcPriceCode"

        Dim lngTotal As Long
        Dim strTmp01 As String
        Dim intX As Integer
        Dim lngRetVal As Long

710     lngRetVal = 0&
720     lngTotal = 0&
730     strTmp01 = strDate

        ' ** Add the numbers in the date string.
740     For intX = 1 To Len(strDate)
750       If Asc(Mid(strTmp01, intX, 1)) > Asc("/") And Asc(Mid(strTmp01, intX, 1)) < Asc(":") Then
760         lngTotal = lngTotal + Val(Mid(strTmp01, intX, 1))
770       End If
780     Next
        ' ** Double the total.
790     lngTotal = lngTotal * 2
        ' ** Add 17531 to the total, simply to make it harder to decipher.
800     lngRetVal = lngTotal + 17531

EXITP:
810     CalcPriceCode = lngRetVal
820     Exit Function

ERRH:
830     Select Case ERR.Number
        Case Else
840       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
850     End Select
860     Resume EXITP

End Function

Public Function DecodeString(strValue As String) As String

900   On Error GoTo ERRH

        Const THIS_PROC As String = "DecodeString"

        Dim lox As SRSencrypt.encrypt  'As Object  'As encrypt  ' ** VGC 07/29/2008:
        Dim strRetVal As String

        'Set lox = CreateObject("encrypt")  ' ** VGC 07/29/2008: Switched back to this after Error 429, ActiveX component can't create object.
910     Set lox = New SRSencrypt.encrypt
920     strRetVal = lox.DeCode(strValue)

EXITP:
930     DecodeString = strRetVal
940     Exit Function

ERRH:
950     DoCmd.Hourglass False
960     Select Case ERR.Number
        Case Else
970       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
980     End Select
990     Resume EXITP

End Function

Public Function EncodeString(strValue As String) As String

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "EncodeString"

        Dim lox As SRSencrypt.encrypt  'As Object  'As encrypt  ' ** VGC 07/29/2008:
        Dim strReturn As String

        'Set lox = CreateObject("encrypt")  ' ** VGC 07/29/2008: Switched back to this after Error 429, ActiveX component can't create object.
1010    Set lox = New SRSencrypt.encrypt
1020    strReturn = lox.EnCode(strValue)
1030    EncodeString = strReturn

EXITP:
1040    Exit Function

ERRH:
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Function

Public Function ForcePause(dblTimDurSEC As Double) As Boolean
' ** Pause execution for a specified number of seconds.
' ** Since dblTimDurSEC is Double, can take decimal fractions.
' ** For example, 0.25 is equivalent to 250 milliseconds as specified for a TimerInterval.

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "ForcePause"

        Dim MMT As MMTIME
        Dim lngTimStart As Long   ' ** The start time.
        Dim lngTimFinish As Long  ' ** The finish time.
        Dim lngTimLast As Long    ' ** The current time.
        Dim blnRetVal As Boolean

1110    blnRetVal = True

1120    MMT.wType = TIME_MS

1130    lngTimStart = GetTimeUnits  ' ** Module Function: modWindowFunctions.
1140    lngTimLast = lngTimStart
1150    lngTimFinish = lngTimStart + (dblTimDurSEC * 1000)

1160    Do While lngTimLast < lngTimFinish
          ' ** Dum-dee-dum-dum...
1170      lngTimLast = GetTimeUnits  ' ** Module Function: modWindowFunctions.
1180    Loop

EXITP:
1190    ForcePause = blnRetVal
1200    Exit Function

ERRH:
1210    blnRetVal = False
1220    Select Case ERR.Number
        Case Else
1230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1240    End Select
1250    Resume EXITP

End Function

Public Function strPadZ(strText As String, intLen As Integer) As String

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "strPadZ"

1310    strText = Trim(strText)
1320    Do While Len(strText) < intLen
1330      strText = "0" & strText
1340    Loop

EXITP:
1350    strPadZ = strText
1360    Exit Function

ERRH:
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Function

Public Function Pass_Check(varInput As Variant, varUserName As Variant, Optional varShowMsg As Variant) As Boolean

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Pass_Check"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnHasOKLen As Boolean, blnHasNoSpace As Boolean, blnHasUC As Boolean, blnHasLC As Boolean, blnHasNum As Boolean
        Dim blnIsFound As Boolean, blnIsDiff As Boolean, blnShowMsg As Boolean
        Dim strMsg As String
        Dim intLen As Integer, intPos01 As Integer
        Dim lngRecs As Long
        Dim varTmp00 As Variant, varTmp01 As Variant, strTmp02 As String
        Dim intX As Integer, lngY As Long
        Dim blnRetVal As Boolean

1510    blnRetVal = True

1520    blnHasOKLen = False: blnHasNoSpace = False: blnHasUC = False: blnHasLC = False: blnHasNum = False
1530    blnIsFound = False: blnIsDiff = False

1540    If IsNull(varInput) = False And IsNull(varUserName) = False Then
1550      strTmp02 = Trim(varInput)
1560      If strTmp02 <> vbNullString Then
1570        intLen = Len(strTmp02)
1580        If intLen >= 5 And intLen <= 14 Then
              ' ** 5-14 chars.
1590          blnHasOKLen = True
1600          intPos01 = InStr(strTmp02, Chr(32))
1610          If intPos01 = 0 Then
                ' ** No space.
1620            blnHasNoSpace = True
1630            For intX = 1 To intLen
1640              If Asc(Mid(strTmp02, intX, 1)) >= 65 And Asc(Mid(strTmp02, intX, 1)) <= 90 Then
                    ' ** 1 upper-case.
1650                blnHasUC = True
1660              ElseIf Asc(Mid(strTmp02, intX, 1)) >= 97 And Asc(Mid(strTmp02, intX, 1)) <= 122 Then
                    ' ** 1 lower-case.
1670                blnHasLC = True
1680              ElseIf Asc(Mid(strTmp02, intX, 1)) >= 48 And Asc(Mid(strTmp02, intX, 1)) <= 57 Then
                    ' ** 1 numeral.
1690                blnHasNum = True
1700              End If
1710            Next
1720            If blnHasUC = True And blnHasLC = True And blnHasNum = True Then
                  ' ** Finally, check the password against the current password for this user.
1730              varTmp00 = Null: varTmp01 = Null
1740              Set dbs = CurrentDb
1750              With dbs
                    ' ** Get the user's GUID from the Users table.
1760                Set rst = .OpenRecordset("Users", dbOpenDynaset, dbReadOnly)
1770                With rst
1780                  If .BOF = True And .EOF = True Then
                        ' ** Shouldn't happen!
1790                  Else
1800                    .MoveLast
1810                    lngRecs = .RecordCount
1820                    .MoveFirst
1830                    For lngY = 1& To lngRecs
1840                      If ![Username] = varUserName Then
1850                        varTmp00 = FilterGUIDString(StringFromGUID(![s_GUID]))  ' ** Function: Below.
1860                        Exit For
1870                      End If
1880                      If lngY < lngRecs Then .MoveNext
1890                    Next
1900                  End If
1910                  .Close
1920                End With
1930                If IsNull(varTmp00) = False Then
                      ' ** Check this user's current password found in the _~xusr table.
1940                  Set rst = .OpenRecordset("_~xusr", dbOpenDynaset, dbReadOnly)
1950                  With rst
1960                    If .BOF = True And .EOF = True Then
                          ' ** No users yet, so definitely a new user.
1970                    Else
1980                      .MoveLast
1990                      lngRecs = .RecordCount
2000                      .MoveFirst
2010                      For lngY = 1& To lngRecs
2020                        If FilterGUIDString(StringFromGUID(![s_GUID])) = varTmp00 Then  ' ** Function: Below.
2030                          varTmp01 = ![xusr_extant]  ' ** Checks new password against current password.
2040                          Exit For
2050                        End If
2060                        If lngY < lngRecs Then .MoveNext
2070                      Next
2080                    End If
2090                    .Close
2100                  End With
2110                  If IsNull(varTmp01) = False Then
                        ' ** Existing user.
2120                    blnIsFound = True
2130                    If DecodeString(CStr(varTmp01)) <> strTmp02 Then  ' ** Function: Above.
                          ' ** Not the same as their last password.
                          ' ** PASSES ALL THE TESTS!
2140                      blnIsDiff = True
2150                    End If
2160                  Else
                        ' ** New user.
                        ' ** PASSES ALL THE TESTS!
2170                    blnIsDiff = True
2180                  End If
2190                Else
                      ' ** New user.
                      ' ** PASSES ALL THE TESTS!
2200                  blnIsDiff = True
2210                End If
2220                .Close
2230              End With
2240            End If
2250          End If
2260        End If
2270        If blnHasOKLen = False Or blnHasNoSpace = False Or blnHasUC = False Or _
                blnHasLC = False Or blnHasNum = False Or blnIsDiff = False Then
2280          DoCmd.Hourglass False
2290          If blnHasOKLen = False Then
2300            strMsg = "The password cannot be shorter than 5 characters, nor longer than 14 characters."
2310          ElseIf blnHasNoSpace = False Then
2320            strMsg = "The password cannot contain a space."
2330          ElseIf blnHasUC = False Then
2340            strMsg = "The password must have at least 1 upper-case (capital) letter."
2350          ElseIf blnHasLC = False Then
2360            strMsg = "The password must have at least 1 lower-case letter."
2370          ElseIf blnHasNum = False Then
2380            strMsg = "The password must have at least 1 number."
2390          ElseIf blnIsDiff = False Then
2400            strMsg = "The password must be different from the one it's replacing."
2410          End If
2420          blnRetVal = False
2430          blnShowMsg = True
2440          Select Case IsMissing(varShowMsg)
              Case True
                ' ** Show message.
2450          Case False
2460            blnShowMsg = CBool(varShowMsg)
2470          End Select
2480          If blnShowMsg = True Then
2490            MsgBox strMsg & vbCrLf & vbCrLf & _
                  "Password Requirements:" & vbCrLf & _
                  "   5 to 14 characters in length" & vbCrLf & _
                  "   Contain at least 1 upper-case (capital) letter" & vbCrLf & _
                  "   Contain at least 1 lower-case letter" & vbCrLf & _
                  "   Contain at least 1 numeral" & vbCrLf & _
                  "   Cannot contain a space" & _
                  "   Password must be changed every 12 months" & vbCrLf & _
                  "   New password cannot be same as one replaced", _
                  vbInformation + vbOKOnly, ("Invalid Entry" & Space(40))
2500          End If
2510        End If
2520      Else
2530        blnRetVal = False
2540      End If
2550    Else
2560      blnRetVal = False
2570    End If

EXITP:
2580    Set rst = Nothing
2590    Set dbs = Nothing
2600    Pass_Check = blnRetVal
2610    Exit Function

ERRH:
2620    blnRetVal = False
2630    Select Case ERR.Number
        Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Function

Public Function FilterGUIDString(ByRef strGUIDtmp As String) As String
' ** Remove the type name prefix/suffix from a GUID returned by the StringFromGUID() function.

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FilterGUIDString"

        Dim strRetVal As String

2710    If InStr(strGUIDtmp, "guid") > 0 Then
          ' ** {guid {279EB3CC-E559-41F5-93F9-9665A6852453}}
2720      strRetVal = Left(Mid(strGUIDtmp, InStr(2, strGUIDtmp, "{")), (Len(Mid(strGUIDtmp, InStr(2, strGUIDtmp, "{"))) - 1))
2730    Else
2740      strRetVal = strGUIDtmp
2750    End If

EXITP:
2760    FilterGUIDString = strRetVal
2770    Exit Function

ERRH:
2780    strRetVal = vbNullString
2790    Select Case ERR.Number
        Case 5  ' ** Invalid procedure call or argument.
          ' ** Not sure why this showed up after deleting user.
2800    Case Else
2810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2820    End Select
2830    Resume EXITP

End Function

Public Function FilterGUID(ByRef strGUIDtmp As String) As String
' ** Remove the braces and dashes from a GUID string.

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "FilterGUID"

        Dim strRetVal As String

2910    strRetVal = vbNullString

2920    If InStr(strGUIDtmp, "guid") > 0 Then
2930      strGUIDtmp = Left(Mid(strGUIDtmp, InStr(2, strGUIDtmp, "{")), (Len(Mid(strGUIDtmp, InStr(2, strGUIDtmp, "{"))) - 1))
2940    End If

2950    strGUIDtmp = Replace(strGUIDtmp, "-", "", 1, Len(strGUIDtmp))
2960    strGUIDtmp = Replace(strGUIDtmp, "}", "", 1, Len(strGUIDtmp))
2970    strGUIDtmp = Replace(strGUIDtmp, "{", "", 1, Len(strGUIDtmp))

2980    strRetVal = strGUIDtmp

EXITP:
2990    FilterGUID = strRetVal
3000    Exit Function

ERRH:
3010    strRetVal = vbNullString
3020    Select Case ERR.Number
        Case Else
3030      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3040    End Select
3050    Resume EXITP

End Function

Public Function GUID2ByteArray(ByVal strGUID As String) As Byte()
' ** Convert a GUID string to a Byte array.

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "GUID2ByteArray"

        Dim i As Integer
        Dim j As Integer
        Dim sPos As Integer
        Dim Offset As Integer
        Dim sGUID(0 To 2) As Byte
        Dim bytArray() As Byte

3110    ReDim bytArray(0 To 15) As Byte

3120    sGUID(0) = 7
3130    sGUID(1) = 11
3140    sGUID(2) = 15

3150    Offset = 0
3160    sPos = 0

        ' ** AABBCCDD-AABB-CCDD-XXXX-XXXXXXXXXXXX 'Microsoft Access view.
        ' ** DDCCBBAA-BBAA-DDCC-XXXX-XXXXXXXXXXXX 'SQLServer view.
        ' ** Need to loop through to build the GUID byte array in the Microsoft
        ' ** Access storage format since the first eight bytes are reversed.
3170    For i = 0 To UBound(sGUID)
3180      For j = sGUID(i) To (Offset + 1) Step -2
3190        bytArray(sPos) = "&H" & Mid(strGUID, j, 2)
3200        sPos = sPos + 1
3210      Next j
3220      Offset = sGUID(i)
3230    Next i

3240    For i = 17 To 31 Step 2
3250      bytArray(sPos) = "&H" & Mid(strGUID, i, 2)
3260      sPos = sPos + 1
3270    Next i

EXITP:
3280    GUID2ByteArray = bytArray()
3290    Exit Function

ERRH:
3300    Select Case ERR.Number
        Case Else
3310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3320    End Select
3330    Resume EXITP

End Function

Public Function PTMgr_License1() As Boolean

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "PTMgr_License1"

        Dim strSection As String, strSubSection As String, strValue As String
        Dim strPath As String, strFile As String, strPathFile As String
        Dim strFirm As String, strExpires As String, strLimit As String, strAppName As String, strDate As String
        'Dim strInput As String
        'Dim lngLines As Long
        'Dim lngX As Long
        Dim blnRetVal As Boolean, blnSkip As Boolean

3410  On Error GoTo 0

        ' ** [License]
        ' ** Firm=Ã™Á²ÂÁ·u™Í¿ËÁÏÒ
        ' ** Expires=©susw|xz{z†
        ' ** Limit=©su
        ' ** Program=É¢©¡¸Äº¿ÀÌ
        ' ** Exp=©suswwxz{{‚
        ' **
        ' ** test=1

3420    strPath = "C:\VictorGCS_Clients\PTManager"
3430    strFile = "CSI.lic"
3440    strPathFile = strPath & LNK_SEP & strFile

3450    strSection = "License"
3460    strFirm = "Scott R. Blaser, CPA"
3470    strExpires = Format(DateAdd("n", 1, Date), "mm/dd/yyyy")  ' ** One month from today.
3480    strLimit = "35"
3490    strAppName = "PTManager"
3500    strDate = "08/19/2016"  ' ** I don't know what this represents.

        ' ** SubSection 1: Firm.
3510    strSubSection = "Firm"
3520    strValue = EncodeString(strFirm)  ' ** Function: Above.
3530    blnRetVal = PTMgr_IniFile_Set(strSection, strSubSection, strValue, strPathFile)  ' ** Function: Below.
3540    If blnRetVal = True Then
          ' ** SubSection 2: Expires.
3550      strSubSection = "Expires"
3560      strValue = EncodeString(strExpires)  ' ** Function: Above.
3570      blnRetVal = PTMgr_IniFile_Set(strSection, strSubSection, strValue, strPathFile)  ' ** Function: Below.
3580      If blnRetVal = True Then
            ' ** SubSection 3: Limit.
3590        strSubSection = "Limit"
3600        strValue = EncodeString(strLimit)  ' ** Function: Above.
3610        blnRetVal = PTMgr_IniFile_Set(strSection, strSubSection, strValue, strPathFile)  ' ** Function: Below.
3620        If blnRetVal = True Then
              ' ** SubSection 4: Program.
3630          strSubSection = "Program"
3640          strValue = EncodeString(strAppName)  ' ** Function: Above.
3650          blnRetVal = PTMgr_IniFile_Set(strSection, strSubSection, strValue, strPathFile)  ' ** Function: Below.
3660          If blnRetVal = True Then
                ' ** SubSection 5: Exp.
3670            strSubSection = "Exp"
3680            strValue = EncodeString(strDate)  ' ** Function: Above.
3690            blnRetVal = PTMgr_IniFile_Set(strSection, strSubSection, strValue, strPathFile)  ' ** Function: Below.
3700            If blnRetVal = True Then
3710              blnSkip = True
3720              If blnSkip = False Then
                    ' ** SubSection 6: Test.
3730                strSubSection = "Test"
                    ' ** strValue = 1 indicates a demo version.
3740                strValue = vbNullString
                    'I WAS TRYING TO DELETE THE 'TEST' LINE!
                    'lngLines = 0&
                    'Open strPathFile For Input As #1
                    'Do While Not EOF(1)
                    '  lngLines = lngLines + 1&  ' ** Find out how many lines are in the file.
                    '  Line Input #1, strInput
                    'Loop
                    'Close #1
                    'lngX = 0&
                    'Open strPathFile For Input As #1
                    'Do While Not EOF(1)
                    '  lngX = lngX + 1&
                    '  Line Input #1, strInput
                    '  If Left(strInput, 4) = "Test" Then

                    '  End If
                    'Loop
                    'Close #1
3750                blnRetVal = PTMgr_IniFile_Del(strSection, strSubSection, strPathFile)  ' ** Function: Below.
3760              End If  ' ** blnSkip.
3770            End If  ' ** blnRetVal.
3780          End If  ' ** blnRetVal.
3790        End If  ' ** blnRetVal.
3800      End If  ' ** blnRetVal.
3810    End If  ' ** blnRetVal.

3820    If blnRetVal = False Then
3830      Beep
3840      Debug.Print "'ERROR WRITING TO PT-MANAGER LICENSE FILE!"
3850    End If

3860    Beep

3870    Debug.Print "'DONE!"

EXITP:
3880    PTMgr_License1 = blnRetVal
3890    Exit Function

ERRH:
3900    blnRetVal = False
3910    Select Case ERR.Number
        Case Else
3920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3930    End Select
3940    Resume EXITP

End Function

Public Function PTMgr_IniFile_Set(strSection As String, strSubSection As String, strValue As String, strFile As String) As Boolean
' ** Example:
' **   strSection = "License"
' **   strSubSection = "Firm"
' **   strValue = EncodeString(Me.txtLicensedTo)
' **   strFile = gstrTrustDataLocation & gstrFile_LIC
' ** See also: zz_mod_MDEPrepFuncs.xIniFile_Set()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "PTMgr_IniFile_Set"

        Dim lngRetVal As Long
        Dim blnRetVal As Boolean

4010    lngRetVal = WritePrivateProfileStringA(strSection, strSubSection, strValue, strFile)
4020    If lngRetVal = 0 Then
4030      blnRetVal = False
4040    Else
4050      blnRetVal = True
4060    End If

EXITP:
4070    PTMgr_IniFile_Set = blnRetVal
4080    Exit Function

ERRH:
4090    Select Case ERR.Number
        Case Else
4100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select
4120    Resume EXITP

End Function

Public Function PTMgr_IniFile_Del(strSection As String, strSubSection As String, strFile As String) As Boolean
' ** Example:
' **   strSection = "License"
' **   strSubSection = "Firm"
' **   strValue = EncodeString(Me.txtLicensedTo)
' **   strFile = gstrTrustDataLocation & gstrFile_LIC

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "PTMgr_IniFile_Del"

        Dim intLen As Long
        Dim blnRetVal As Boolean

4210    intLen = WritePrivateProfileStringA(strSection, strSubSection, "", strFile)
4220    blnRetVal = True

EXITP:
4230    PTMgr_IniFile_Del = blnRetVal
4240    Exit Function

ERRH:
4250    Select Case ERR.Number
        Case Else
4260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4270    End Select
4280    Resume EXITP

End Function
