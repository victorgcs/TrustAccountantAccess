Attribute VB_Name = "modUtilities"
Option Compare Database
Option Explicit

'VGC 09/02/2017: CHANGES!

Private Const THIS_NAME As String = "modUtilities"
' **

Public Sub ExpandCombo(Optional ctl As Variant)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "ExpandCombo"

110     DoEvents
120     If IsMissing(ctl) = True Then
130   On Error Resume Next
140       SendKeys "{F4}"
150   On Error GoTo ERRH
160     Else
170   On Error Resume Next
180       ctl.Dropdown
190       If ERR.Number <> 0 Then
            ' ** Error 438: Object doesn't support this property or method.
200   On Error GoTo ERRH
210         SendKeys "{F4}"
220       Else
230   On Error GoTo ERRH
            ' ** Just because it didn't error doesn't mean it worked!
240         If ctl.Name = "cmbAccountHelper" Then
250           SendKeys "{F4}"
260         End If
270       End If
280     End If

EXITP:
290     Exit Sub

ERRH:
300     Select Case ERR.Number
        Case 70  ' ** Permission denied.
          ' ** The F4 evidently has different meanings in different environments.
310     Case Else
320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
330     End Select
340     Resume EXITP

End Sub

Public Function IsNothing(varInput As Variant) As Boolean
' ** VBA does not have its own IsNothing() function.

400   On Error GoTo ERRH

        Const THIS_PROC As String = "IsNothing"

        Dim blnRetVal As Boolean

410     blnRetVal = True

420   On Error Resume Next
430     blnRetVal = (varInput Is Nothing)
440     ERR.Clear
450   On Error GoTo ERRH

EXITP:
460     IsNothing = blnRetVal
470     Exit Function

ERRH:
480     blnRetVal = False
490     Select Case ERR.Number
        Case Else
500       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
510     End Select
520     Resume EXITP

End Function

Public Function DateCheck_Post(varDate As Variant) As Boolean
' ** Verify that posting dates entered are within an acceptable range.

600   On Error GoTo ERRH

        Dim datMaxDate As Date
        Dim blnRetVal As Boolean

        Const THIS_PROC As String = "DateCheck_Post"

610     blnRetVal = True    ' ** Unless proven otherwise.

        ' ** Allow posting up to 1 month into the future.
620     datMaxDate = DateAdd("m", 1, Date)

630     If IsNull(varDate) = False Then
640       If Trim(varDate) <> vbNullString Then
650         If IsDate(varDate) = False Then
660           DoCmd.Hourglass False
670           MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
680           blnRetVal = False
690         Else
700           If CDate(varDate) > datMaxDate Then
710             DoCmd.Hourglass False
720             MsgBox "The date must not be later than " & CStr(datMaxDate) & ".", vbInformation + vbOKOnly, "Invalid Date"
730             blnRetVal = False
740           Else
750             If year(CDate(varDate)) < 1900 Then
760               DoCmd.Hourglass False
770               MsgBox "The date entered is out of range." & vbCrLf & _
                    "  Year: " & CStr(year(CDate(varDate))), vbInformation + vbOKOnly, "Invalid Date"
780               blnRetVal = False
790             Else
800               If year(CDate(varDate)) < 1990 Then
810                 DoCmd.Hourglass False
820                 If MsgBox("Are you sure you want to post entries this old?" & vbCrLf & _
                        "  Year: " & CStr(year(CDate(varDate))), vbQuestion + vbYesNo + vbDefaultButton2, "Verify Old Date") = vbNo Then
830                   blnRetVal = False
840                 End If
850               Else
860                 If CDate(varDate) > Date Then
870                   DoCmd.Hourglass False
880                   If MsgBox("Are you sure you want to post entries for future dates?", vbQuestion + vbYesNo + vbDefaultButton2, _
                          "Verify Future Posting") = vbNo Then
890                     blnRetVal = False
900                   End If
910                 End If
920               End If
930             End If
940           End If
950         End If
960       Else
970         blnRetVal = False
980       End If
990     Else
1000      blnRetVal = False
1010    End If

EXITP:
1020    DateCheck_Post = blnRetVal
1030    Exit Function

ERRH:
1040    blnRetVal = False
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Function

Public Function DateCheck_Trade(varDate As Variant) As Boolean
' ** Verify that trading dates entered are within an acceptable range.

1100  On Error GoTo ERRH

        Dim datMaxDate As Date, dblMaxDate As Double, strMaxDate As String, lngMaxDate As Long
        Dim strTmp01 As String, lngTmp02 As Long, dblTmp03 As Double, datTmp04 As Date
        Dim blnRetVal As Boolean

        Const THIS_PROC As String = "DateCheck_Trade"

1110    blnRetVal = True    ' ** Unless proven otherwise.

        ' ** Allow  trading only up to today.
1120    datMaxDate = Date    ' ** For this date checking, a timestamp isn't needed.
1130    dblMaxDate = CDbl(datMaxDate)
1140    strMaxDate = CStr(dblMaxDate)
1150    If InStr(strMaxDate, ".") > 0 Then  ' ** There shouldn't be, but who knows...
1160      strMaxDate = Left(strMaxDate, (InStr(strMaxDate, ".") - 1))
1170      lngMaxDate = CLng(strMaxDate)
1180    Else
1190      lngMaxDate = CLng(strMaxDate)
1200    End If

1210    If IsNull(varDate) = False Then  ' ** (VGC 08/19/2010)
1220      If Trim(varDate) <> vbNullString Then
1230        If Right(varDate, 1) = "_" Then
              ' ** Text passes with InputMask: '08/08/10__'.
1240          Do While Right(varDate, 1) = "_"
1250            varDate = Left(varDate, (Len(varDate) - 1))
1260          Loop
1270        End If
1280        If IsDate(varDate) = False Then
1290          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1300          blnRetVal = False
1310        Else
1320          datTmp04 = CDate(varDate)
1330          dblTmp03 = CDbl(datTmp04)
1340          strTmp01 = CStr(dblTmp03)
1350          If InStr(strTmp01, ".") > 0 Then
1360            strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
1370            lngTmp02 = CLng(strTmp01)
1380          Else
1390            lngTmp02 = CLng(strTmp01)
1400          End If
1410          If lngTmp02 > lngMaxDate Then
                ' ** The Long value is just the Day portion of a date.
1420            MsgBox "Future Trade Dates are not allowed.", vbInformation + vbOKOnly, "Invalid Date"
1430            blnRetVal = False
1440          Else
1450            If year(CDate(varDate)) < 1900 Then
1460              MsgBox "The date entered is out of range." & vbCrLf & _
                    "  Year: " & CStr(year(CDate(varDate))), vbInformation + vbOKOnly, "Invalid Date"
1470              blnRetVal = False
1480            Else
1490              If year(CDate(varDate)) < 1990 Then
1500                If MsgBox("Are you sure you want to enter Trade Dates this old?" & vbCrLf & _
                        "  Year: " & CStr(year(CDate(varDate))), vbQuestion + vbYesNo + vbDefaultButton2, "Verify Old Date") = vbNo Then
1510                  blnRetVal = False
1520                End If
1530              End If
1540            End If
1550          End If
1560        End If
1570      Else
1580        blnRetVal = False
1590      End If
1600    Else
1610      blnRetVal = False
1620    End If

EXITP:
1630    DateCheck_Trade = blnRetVal
1640    Exit Function

ERRH:
1650    blnRetVal = False
1660    Select Case ERR.Number
        Case Else
1670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1680    End Select
1690    Resume EXITP

End Function

Public Function dbl_Round(dblNum As Double, intDecimal As Integer) As Double
' ** Function to do rounding.
' ** As it is now, it must round to >= 1 place past the decmal point; ie. intDecimal must be 1 or more.

1700  On Error GoTo ERRH

        Dim strWork As String
        Dim strOrg As String
        Dim intLoop As Integer
        Dim intPastDecimal As Integer
        Dim blnOut As Boolean
        Dim dblCarry As Double
        Dim intLastDigit As Integer
        Dim dblRetVal As Double

        Const THIS_PROC As String = "dbl_Round"

1710    dblRetVal = dblNum

1720    strWork = ""
1730    strOrg = dblNum  ' ** Conversion to string.

1740    strOrg = Trim(strOrg) & IIf(InStr(strOrg, ".") = 0, ".", "") & "000000000000"  ' ** Pad it out for convenience.
1750    intPastDecimal = 0
1760    blnOut = False
1770    dblCarry = 0#

1780    If intDecimal < 1 Then
1790      MsgBox "Rounding problem; NO ROUNDING DONE!", vbCritical + vbOKOnly, "Error"
1800      strWork = strOrg
1810    Else

          ' ** First, get any part to the left of the decimal.
1820      intLoop = 1
1830      Do While (intLoop <= Len(strOrg)) And Mid(strOrg, intLoop, 1) <> "."
1840        strWork = strWork & Mid(strOrg, intLoop, 1)
1850        intLoop = intLoop + 1
1860      Loop

          ' ** Now, get the decimal, if any.
1870      If (intLoop <= Len(strOrg)) Then
1880        strWork = strWork & Mid(strOrg, intLoop, 1)
1890        intLoop = intLoop + 1
1900        intPastDecimal = 0

            ' ** Now, get numbers to the right of the decimal, if any.
1910        Do While (Not blnOut) And (intLoop <= Len(strOrg))
1920          If intPastDecimal = intDecimal Then  ' ** Just got past where we want to round to!
1930            If Int(Mid(strOrg, intLoop, 1)) >= 5 Then  ' ** Need to increment previous one.
1940              intLastDigit = CInt(Right(strWork, 1)) + 1
1950              If intLastDigit < 10 Then
1960                strWork = Left(strWork, intLoop - 2) & CStr(intLastDigit)
1970              Else
1980                strWork = Left(strWork, intLoop - 2) & CStr(intLastDigit - 10)
1990                dblCarry = 10 ^ ((intPastDecimal - 1) * -1)
2000              End If
2010            End If
2020            blnOut = True
2030          Else
2040            strWork = strWork & Mid(strOrg, intLoop, 1)
2050            intLoop = intLoop + 1
2060            intPastDecimal = intPastDecimal + 1
2070          End If
2080        Loop

2090      End If
2100    End If

2110    dblRetVal = CDbl(IIf(strWork = vbNullString, "0", strWork)) + dblCarry

EXITP:
2120    dbl_Round = dblRetVal
2130    Exit Function

ERRH:
2140    dblRetVal = dblNum
2150    Select Case ERR.Number
        Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

End Function

Public Function GetDollarString(dblInput As Double) As String

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "GetDollarString"

        Dim dblVal As Double
        Dim blnNeedThousand As Boolean, blnSkip As Boolean
        Dim intPos01 As Integer, intPos02 As Integer, intLen As Integer
        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String, strTmp04 As String, dblTmp05 As Double
        Dim intX As Integer
        Dim strRetVal As String

2210    strRetVal = vbNullString

2220    blnNeedThousand = False     ' ** Unless proven otherwise.

        ' ** NOTE: THIS DOES NOT WORK OVER 100 MILLION!
2230    If dblInput > 99999999.99 Then  ' ** TO THE NINES! (Hopefully, not at sixes and sevens.) (Nor behind the 8 ball.) (Though it could by in the pipe five-by-five.)  (But not 6 ways from Sunday.)
2240      strRetVal = "Error - Number is Too Large"
2250    Else

          ' ** Paid may come through as negative.
2260      dblVal = Abs(Round(dblInput, 2))  ' ** Make sure it only has 2 decimal places.

2270      If dblVal / 1000000 >= 1 Then
2280        strTmp01 = getTens(Int(dblVal / 1000000))
2290        strTmp01 = strTmp01 & "Million "
2300        dblVal = dblVal - Int(dblVal / 1000000) * 1000000
2310      End If

2320      If Int(dblVal / 1000) >= 1 Then
2330        If Int(dblVal / 100000) > 0 Or Len(strTmp01) > 0 Then
2340          strTmp01 = strTmp01 & getOnes(Int(dblVal / 100000)) & IIf(Int(dblVal / 100000) > 0, "Hundred ", "")
2350          blnNeedThousand = True
2360        End If
2370        dblVal = dblVal - Int(dblVal / 100000) * 100000
2380        If Int(dblVal / 1000) > 0 Or Len(strTmp01) > 0 Then
2390          strTmp01 = strTmp01 & getTens(Int(dblVal / 1000))
2400          If Int(dblVal / 1000) > 0 Then blnNeedThousand = True
2410        End If
2420        If blnNeedThousand Then
2430          strTmp01 = strTmp01 & "Thousand "
2440        End If
2450        dblVal = dblVal - Int(dblVal / 1000) * 1000
2460      End If

2470      If Int(dblVal / 100) > 0 Then
2480        strTmp01 = strTmp01 & getOnes(Int(dblVal / 100)) & "Hundred "
2490      End If

2500      dblVal = dblVal - (Int(dblVal / 100) * 100)
2510      If Int(dblVal) > 0 Or Len(strTmp01) > 0 Then
2520        strTmp01 = strTmp01 & getTens(Int(dblVal)) & "and "
2530      End If
2540      dblVal = CDbl(Val(Right(Format(dblVal, "###.00"), 2)))
2550      Select Case dblVal
          Case Is > 0
            'strTmp01 = strTmp01 & getTens(dblVal) & "One Hundredths"
2560        strTmp01 = strTmp01 & CStr(dblVal) & "/100"
2570      Case Else
            'strTmp01 = strTmp01 & "Zero One Hundredths"
2580        strTmp01 = strTmp01 & "00/100"
2590      End Select

2600      intLen = 109 - Len(strTmp01)
2610      For intX = 0 To intLen
2620        strTmp04 = "_" & strTmp04  ' ** Just the filler, no text.
2630      Next

2640      blnSkip = True
2650      If blnSkip = False Then
            ' ** Compare input vs. output.
2660        If InStr(CStr(dblInput), ".") > 0 Then
              ' ** and Fifty Seven One Hundredths
              ' ** and Seven One Hundredths
              ' ** and Zero One Hundredths
2670          dblTmp05 = Abs(dblInput)
2680          strTmp02 = CStr(dblTmp05)
2690          intPos01 = InStr(strTmp01, "One Hundredths")
2700          If intPos01 > 0 Then
2710            intPos02 = intPos01
2720            intPos01 = 0
2730            For intX = intPos02 To 1 Step -1
2740              If Mid(strTmp01, intX, 5) = " and " Then
2750                intPos01 = intX + 1
2760                Exit For
2770              End If
2780            Next
2790            If intPos01 > 0 Then
2800              intPos01 = (intPos01 + 4)
2810              strTmp03 = Trim(Mid(strTmp01, intPos01, (intPos02 - intPos01)))
                  ' ** Reverse to number.
2820            End If
2830          Else
                ' ** No 100ths! Shouldn't happen!
2840          End If
2850        Else
2860          dblTmp05 = Abs(dblInput)
2870          strTmp02 = CStr(dblTmp05)
2880          intPos01 = InStr(strTmp01, "One Hundredths")
2890          If intPos01 > 0 Then
                ' ** Should be Zero.
2900          End If
2910        End If
2920      End If  ' ** blnSkip.

2930      strRetVal = strTmp01 & " " & strTmp04

2940    End If

EXITP:
2950    GetDollarString = strRetVal
2960    Exit Function

ERRH:
2970    strRetVal = vbNullString
2980    Select Case ERR.Number
        Case Else
2990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3000    End Select
3010    Resume EXITP

End Function

Public Function getOnes(dblInput As Double) As String

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "getOnes"

        Dim strRetVal As String

3110    strRetVal = vbNullString

3120    Select Case dblInput
        Case 1
3130      strRetVal = "One "
3140    Case 2
3150      strRetVal = "Two "
3160    Case 3
3170      strRetVal = "Three "
3180    Case 4
3190      strRetVal = "Four "
3200    Case 5
3210      strRetVal = "Five "
3220    Case 6
3230      strRetVal = "Six "
3240    Case 7
3250      strRetVal = "Seven "
3260    Case 8
3270      strRetVal = "Eight "
3280    Case 9
3290      strRetVal = "Nine "
3300    End Select

EXITP:
3310    getOnes = strRetVal
3320    Exit Function

ERRH:
3330    strRetVal = vbNullString
3340    Select Case ERR.Number
        Case Else
3350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3360    End Select
3370    Resume EXITP

End Function

Public Function getTens(dblInput As Double) As String

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "getTens"

        Dim blnContinue As Boolean
        Dim strRetVal As String

3410    strRetVal = vbNullString

3420    blnContinue = True

3430    Select Case Int(dblInput / 10)
        Case 1
3440      Select Case dblInput
          Case 10
3450        strRetVal = "Ten "
3460      Case 11
3470        strRetVal = "Eleven "
3480      Case 12
3490        strRetVal = "Twelve "
3500      Case 13
3510        strRetVal = "Thirteen "
3520      Case 14
3530        strRetVal = "Fourteen "
3540      Case 15
3550        strRetVal = "Fifteen "
3560      Case 16
3570        strRetVal = "Sixteen "
3580      Case 17
3590        strRetVal = "Seventeen "
3600      Case 18
3610        strRetVal = "Eighteen "
3620      Case 19
3630        strRetVal = "Nineteen "
3640      End Select
3650      blnContinue = False
3660    Case 2
3670      strRetVal = "Twenty "
3680    Case 3
3690      strRetVal = "Thirty "
3700    Case 4
3710      strRetVal = "Forty "
3720    Case 5
3730      strRetVal = "Fifty "
3740    Case 6
3750      strRetVal = "Sixty "
3760    Case 7
3770      strRetVal = "Seventy "
3780    Case 8
3790      strRetVal = "Eighty "
3800    Case 9
3810      strRetVal = "Ninety "
3820    Case 0

3830    End Select

3840    If blnContinue = True Then
3850      If Int(dblInput - Int(dblInput / 10) * 10) <> 0 Then
3860        strRetVal = Trim(strRetVal) & "-"
3870      End If
3880      strRetVal = strRetVal & getOnes(Int(dblInput - Int(dblInput / 10) * 10))
3890    End If

EXITP:
3900    getTens = strRetVal
3910    Exit Function

ERRH:
3920    strRetVal = vbNullString
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Function

Public Function SetDateSpecificSQL(strAccountNo As String, strOption As String, strActiveFormName As String, Optional varStartDate As Variant, Optional varEndDate As Variant, Optional varIsArchive As Variant) As Integer
' ** Return codes:
' **    0  Success.
' **   -2  No data.
' **   -4  Date criteria not met.
' **   -9  Error.
' ** Called by:
' **   frmRpt_ArchivedTransactions
' **     {remarked out!} cmdPreview_Click()
' **     {remarked out!}   Me.cmbAccounts  'Statements'
' **     {remarked out!}   'All'           'Statements'
' **     cmdPrint_Click
' **       Me.cmbAccounts  'Statements'
' **       'All'           'Statements'
' **   frmRpt_CourtReports_CA
' **     blnBuildAssetListInfo
' **       Me.cmbAccounts  'Statements'
' **   frmRpt_CourtReports_FL
' **     blnBuildAssetListInfo
' **       Me.cmbAccounts  'Statements'
' **   frmRpt_CourtReports_NS
' **     blnBuildAssetListInfo
' **       Me.cmbAccounts  'Statements'
' **   frmRpt_CourtReports_NY
' **     blnBuildAssetListInfo
' **       Me.cmbAccounts  'Statements'
' **   frmStatementParameters
' **     blnBuildAssetListInfo
' **       Me.cmbAccounts  'Statements'
' **       'All'           'Statements'
' **     blnCommonTransactionCode
' **       Me.cmbAccounts  'StatementTransactions'
' **       Me.cmbAccounts  'Statements'
' **       'All'           'StatementTransactions'
' **       'All'           'Statements'
' **   frmRpt_TransactionsByType
' **     blnCommonTransactionCode
' **       'All'           'StatementTransactions'

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "SetDateSpecificSQL"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strSQL As String
        Dim strStartDate As String, strEndDate As String
        Dim blnNoAccountsUpdated As Boolean, blnIsArchive As Boolean
        Dim strMaxDateFld As String
        Dim strQry_BalDate As String, strQry_SumIncRMA As String, strQry_SumInc As String, strQry_Inc As String
        Dim strQry_SumDec As String, strQry_Dec As String
        Dim blnIsCourtRpt As Boolean
        Dim intRetVal_CheckFirstAcctBal As Integer
        Dim intRetVal As Integer

4010    intRetVal = 0
4020    blnNoAccountsUpdated = True: blnIsCourtRpt = False

4030    If IsMissing(varIsArchive) = True Then
4040      blnIsArchive = False
4050    Else
4060      blnIsArchive = CBool(varIsArchive)
4070    End If

4080    strMaxDateFld = "MaxOfbalance date"
4090    strQry_BalDate = "qryMaxBalDates"
4100    strQry_SumIncRMA = "qrySumIncreasesRMA"
4110    strQry_SumInc = "qrySumIncreases"
4120    strQry_Inc = "qryIncreases"
4130    strQry_SumDec = "qrySumDecreases"
4140    strQry_Dec = "qryDecreases"

4150    If strOption = "StatementTransactions" Then
4160      strEndDate = Forms(strActiveFormName)!TransDateEnd
4170      strStartDate = Forms(strActiveFormName)!TransDateStart
4180    Else
4190      If IsMissing(varStartDate) = True Then
4200        strEndDate = Forms(strActiveFormName)!DateEnd
4210  On Error Resume Next
4220        strStartDate = Forms(strActiveFormName)!DateStart
4230        If ERR <> 0 Then
4240          If strActiveFormName = "frmStatementParameters" Then
4250            If IsNull(DLookup("[MaxOfbalance date]", "qryMaxBalDates", "[accountno] = '" & strAccountNo & "'")) = True Then
4260              strStartDate = "1/1/1900"
4270            Else
4280              strStartDate = DLookup("[MaxOfbalance date]", "qryMaxBalDates", "[accountno] = '" & strAccountNo & "'")
4290            End If
4300          Else
4310            strStartDate = "1/1/1900"
4320          End If
4330        End If
4340  On Error GoTo ERRH
4350      Else
4360        blnIsCourtRpt = True
4370        strStartDate = Format(CDate(varStartDate), "mm/dd/yyyy")
4380        strEndDate = Format(CDate(varEndDate), "mm/dd/yyyy")
4390        strMaxDateFld = "balance date"
4400        strQry_BalDate = "qryCourtReport_05"
4410        strQry_SumIncRMA = "qryCourtReport_10"
4420        strQry_SumInc = "qryCourtReport_11"
4430        strQry_SumDec = "qryCourtReport_12"
4440        strQry_Inc = "qryCourtReport_13"
4450        strQry_Dec = "qryCourtReport_14"
4460      End If
4470    End If

4480    Set dbs = CurrentDb

4490    If strAccountNo = "All" Then
          ' ** Loop through each account and set the right Balance information.
          ' ** The period ending date might be valid for some of the accounts,
          ' ** but invalid for others.  So here we only change those that are
          ' ** invalid so that they have a valid balance data.
          ' ** However, the accounts that have no transactions prior to the
          ' ** chosen period ending date will be passed by.
4500      strSQL = "SELECT [accountno] FROM account;"
4510      Set rst = dbs.OpenRecordset(strSQL)
4520      Do Until rst.EOF
            ' ** At this point, gvarCrtRpt_FL_SpecData may be initialized to 1.
4530        intRetVal_CheckFirstAcctBal = CheckFirstAcctBal(dbs, rst![accountno], strOption, strEndDate) ' ** Function: Below.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -4  Date criteria not met.
            ' **   -9  Error.
4540        If intRetVal_CheckFirstAcctBal = 0 Then
              ' ** If it gets at least one good hit, the variable's turned off.
4550          blnNoAccountsUpdated = False
4560        End If
4570        rst.MoveNext
4580      Loop
4590      rst.Close
          ' ** For all accounts the return value only represents 1 of many, and the Boolean (next line), represents the entire function.
4600      If blnNoAccountsUpdated = True Then
4610        intRetVal = -4
            ' ** None of the accounts were updated due to the fact that
            ' ** none of the accounts had transactions that were prior
            ' ** to the selected Period Ending date.
4620      End If
4630    Else
          ' ** Specific account was chosen.
4640      If blnIsArchive = False Then
4650        intRetVal_CheckFirstAcctBal = CheckFirstAcctBal(dbs, strAccountNo, strOption, strEndDate)  ' ** Function: Below.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -4  Date criteria not met.
            ' **       The account that was selected was not updated due to the fact that it
            ' **       had no transactions that were prior to the selected Period Ending date.
            ' **   -9  Error.
            ' ** For 1 account, this return value represents the entire function.
4660        intRetVal = intRetVal_CheckFirstAcctBal
4670      End If
4680    End If

4690    If intRetVal = 0 Then

4700      If strOption = "StatementTransactions" Then
4710        strSQL = "SELECT Balance.accountno As accountno, Max(Balance.[balance date]) AS [" & strMaxDateFld & "] " & _
              "FROM Balance " & _
              "GROUP BY Balance.accountno;"
4720      Else
            ' ** ACTUALLY, THE COMBO BOX LOOP NOW SKIPS NEWER ACCOUNTS!
            'If strActiveFormName = "frmStatementParameters" Then
            '  If Forms(strActiveFormName)![chkStatements] = True Then
            '    ' ** qryStatementParameters_20x.
            '    strSQL = "SELECT Balance.accountno AS accountno, Max(Balance.[balance date]) AS [" & strMaxDateFld & "] "
            '    strSQL = strSQL & "FROM account INNER JOIN Balance ON account.accountno = Balance.accountno "
            '    strSQL = strSQL & "WHERE (((Balance.[balance date]) < #" & strEndDate & "#) And " & _
            '      "((account.predate) <= #" & strEndDate & "#)) "
            '    strSQL = strSQL & "GROUP BY Balance.accountno;"
            '  Else
            '    strSQL = "SELECT Balance.accountno As accountno, Max(Balance.[balance date]) AS [" & strMaxDateFld & "] " & _
            '      "FROM Balance " & _
            '      "WHERE (((Balance.[balance date]) < #" & strEndDate & "#)) " & _
            '      "GROUP BY Balance.accountno;"
            '  End If
            'Else
4730        strSQL = "SELECT Balance.accountno As accountno, Max(Balance.[balance date]) AS [" & strMaxDateFld & "] " & _
              "FROM Balance " & _
              "WHERE (((Balance.[balance date]) < #" & strEndDate & "#)) " & _
              "GROUP BY Balance.accountno;"
            'End If
4740      End If
4750      dbs.QueryDefs(strQry_BalDate).SQL = strSQL
          ' ** For Statements:
          ' **   strQry_BalDate = "qryMaxBalDates"
          ' **   strMaxDateFld = "MaxOfbalance date"
          ' **   strEndDate = Forms("frmStatementParameters").DateEnd
          ' **   Reports are called in this order:
          ' **     Asset List  'THIS COMES OVER AS STATEMENTS EVEN IF IT'S ONLY THE ASSET LIST REPORT ALONE!
          ' **     Transaction List
          ' **     Summary
          ' ** qryMaxBalDates
          ' ** qryCourtReport_05

4760      If strOption <> "StatementTransactions" Then
4770        If strOption = "Statements" Then
              ' ** Setting the qrySumIncreases SQL.
4780          If blnIsArchive = True And blnIsCourtRpt = True Then
4790            strSQL = dbs.QueryDefs("qryCourtReport_08_01_archive_04").SQL
4800          Else
4810            strSQL = "SELECT ledger.accountno, Sum(IIf(ledger.icash<0,0,ledger.icash)) AS SumPositiveIcash, " & _
                  "Sum(IIf(ledger.pcash<0,0,ledger.pcash)) AS SumPositivePcash, Sum(ledger.cost) AS SumPositiveCost, " & _
                  strQry_SumIncRMA & ".RMA " & _
                  "FROM (ledger INNER JOIN " & strQry_BalDate & " ON ledger.accountno = " & strQry_BalDate & ".accountno) " & _
                  "LEFT JOIN " & strQry_SumIncRMA & " ON ledger.accountno = " & strQry_SumIncRMA & ".accountno " & _
                  "WHERE (((ledger.journaltype)='Dividend' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Interest' " & _
                  "OR (ledger.journaltype)='Sold' OR (ledger.journaltype)='Withdrawn' OR (ledger.journaltype)='Received') " & _
                  "AND ((ledger.icash)>=0) AND " & _
                  "((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.pcash)>=0) "
4820            strSQL = strSQL & "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.pcash)>=0) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)>=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)<0)) " & _
                  "GROUP BY ledger.accountno, " & strQry_SumIncRMA & ".RMA;"
4830          End If
4840        Else
              ' ** Setting the qrySumIncreases SQL.
4850          If blnIsArchive = True And blnIsCourtRpt = True Then
4860            strSQL = dbs.QueryDefs("qryCourtReport_08_02_archive_04").SQL
4870          Else
4880            strSQL = "SELECT ledger.accountno, Sum(IIf(ledger.icash<0,0,ledger.icash)) AS SumPositiveIcash, " & _
                  "Sum(IIf(ledger.pcash<0,0,ledger.pcash)) AS SumPositivePcash, Sum(ledger.cost) AS SumPositiveCost, " & _
                  strQry_SumIncRMA & ".RMA " & _
                  "FROM ledger LEFT JOIN " & strQry_SumIncRMA & " ON ledger.accountno = " & strQry_SumIncRMA & ".accountno " & _
                  "WHERE (((ledger.journaltype)='Dividend' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Interest' " & _
                  "OR (ledger.journaltype)='Sold' OR (ledger.journaltype)='Withdrawn' OR (ledger.journaltype)='Received') " & _
                  "AND ((ledger.icash)>=0) AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.pcash)>=0) AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') "
4890            strSQL = strSQL & "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.pcash)>=0) AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') " & _
                  "AND ((ledger.icash)>=0) AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)<0)) " & _
                  "GROUP BY ledger.accountno, " & strQry_SumIncRMA & ".RMA;"
4900          End If
4910        End If
4920        dbs.QueryDefs(strQry_SumInc).SQL = strSQL
            ' ** qrySumIncreases
            ' ** qryCourtReport_11

            'SELECT ledger.accountno, Sum(IIf(ledger.icash<0,0,ledger.icash)) AS SumPositiveIcash,
            '  Sum(IIf(ledger.pcash<0,0,ledger.pcash)) AS SumPositivePcash, Sum(ledger.cost) AS SumPositiveCost, qrySumIncreasesRMA.RMA
            'FROM (ledger INNER JOIN qryMaxBalDates ON ledger.accountno = qryMaxBalDates.accountno)
            '  LEFT JOIN qrySumIncreasesRMA ON ledger.accountno = qrySumIncreasesRMA.accountno
            'WHERE (((ledger.journaltype)='Dividend' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Interest'
            '  OR (ledger.journaltype)='Sold' OR (ledger.journaltype)='Withdrawn' OR (ledger.journaltype)='Received')
            '  AND ((ledger.icash)>=0) AND ((ledger.transdate)>=CDate(Format([qryMaxBalDates].[MaxOfbalance date],'mm/dd/yyyy'))+1
            '  AND (ledger.transdate)<=##) AND ((ledger.pcash)>=0) AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.')
            '  AND ((ledger.transdate)>=CDate(Format([qryMaxBalDates].[MaxOfbalance date],'mm/dd/yyyy'))+1
            '  AND (ledger.transdate)<=#12/31/2007#) AND ((ledger.pcash)>=0) AND ((ledger.ledger_HIDDEN)=False))
            '  OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)>=0)
            '  AND ((ledger.transdate)>=CDate(Format([qryMaxBalDates].[MaxOfbalance date],'mm/dd/yyyy'))+1
            '  AND (ledger.transdate)<=#12/31/2007#) AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Cost Adj.')
            '  AND ((ledger.transdate)>=CDate(Format([qryMaxBalDates].[MaxOfbalance date],'mm/dd/yyyy'))+1
            '  AND (ledger.transdate)<=#12/31/2007#) AND ((ledger.ledger_HIDDEN)=False) AND ((ledger.cost)<0))
            'GROUP BY ledger.accountno, qrySumIncreasesRMA.RMA;

4930        If strOption = "Statements" Then
              ' ** Setting the qrySumDecreases SQL.
4940          If blnIsArchive = True And blnIsCourtRpt = True Then
4950            strSQL = dbs.QueryDefs("qryCourtReport_08_03_archive_04").SQL
4960          Else
4970            strSQL = "SELECT ledger.accountno, Sum(IIf(ledger.icash>0,0,ledger.icash)) AS SumNegativeIcash, " & _
                  "Sum(IIf(ledger.pcash>0,0,ledger.pcash)) AS SumNegativePcash, Sum(ledger.cost) AS SumNegativeCost " & _
                  "FROM ledger INNER JOIN " & strQry_BalDate & " ON ledger.accountno = " & strQry_BalDate & ".accountno " & _
                  "WHERE (((ledger.journaltype)='Purchase' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Paid') " & _
                  "AND ((ledger.icash)<=0) AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Deposit') AND ((ledger.icash)<=0) AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) "
4980            strSQL = strSQL & "AND ((ledger.cost)<>0)) OR (((ledger.journaltype)='Misc.') AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)>0)) " & _
                  "GROUP BY ledger.accountno;"
4990          End If
5000        Else
              ' ** Setting the qrySumDecreases SQL.
5010          If blnIsArchive = True And blnIsCourtRpt = True Then
5020            strSQL = dbs.QueryDefs("qryCourtReport_08_04_archive_04").SQL
5030          Else
5040            strSQL = "SELECT ledger.accountno, Sum(IIf([ledger].[icash]>0,0,[ledger].[icash])) AS SumNegativeIcash, " & _
                  "Sum(IIf([ledger].[pcash]>0,0,[ledger].[pcash])) AS SumNegativePcash, Sum(ledger.cost) AS SumNegativeCost " & _
                  "FROM ledger " & _
                  "WHERE (((ledger.journaltype)='Purchase' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Paid') " & _
                  "AND ((ledger.icash)<=0) AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Deposit') AND ((ledger.icash)<=0) " & _
                  "AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False) AND ((ledger.cost)<>0)) "
5050            strSQL = strSQL & "OR (((ledger.journaltype)='Misc.') AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)<=0) " & _
                  "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)>0)) " & _
                  "GROUP BY ledger.accountno;"
5060          End If
5070        End If
5080        dbs.QueryDefs(strQry_SumDec).SQL = strSQL
            ' ** qrySumDecreases
            ' ** qryCourtReport_12

5090        If strOption = "Statements" Then
              ' ** Setting the qryIncreases SQL.
5100          If blnIsArchive = True And blnIsCourtRpt = True Then
5110            strSQL = dbs.QueryDefs("qryCourtReport_08_05_archive_04").SQL
5120          Else
5130            strSQL = "SELECT ledger.accountno, ledger.journaltype, Sum(IIf([ledger].[icash]<0,0,[ledger].[icash])) AS PositiveIcash, " & _
                  "Sum(IIf([ledger].[pcash]<0,0,[ledger].[pcash])) AS PositivePcash, Sum(ledger.cost) AS PositiveCost  " & _
                  "FROM ledger INNER JOIN " & strQry_BalDate & " ON ledger.accountno = " & strQry_BalDate & ".accountno " & _
                  "WHERE (((ledger.journaltype)='Dividend' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Interest' " & _
                  "OR (ledger.journaltype)='Sold' OR (ledger.journaltype)='Withdrawn' OR (ledger.journaltype)='Received') " & _
                  "AND ((ledger.icash)>=0) AND ((ledger.pcash)>=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Misc.') AND ((ledger.pcash)>=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) "
5140            strSQL = strSQL & "OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)>=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)<0)) " & _
                  "GROUP BY ledger.accountno, ledger.journaltype " & _
                  "HAVING (((ledger.accountno) = [Reports]![rptAccountSummary]![accountno]));"
5150          End If
5160        Else
              ' ** Setting the qryIncreases SQL.
5170          If blnIsArchive = True And blnIsCourtRpt = True Then
5180            strSQL = dbs.QueryDefs("qryCourtReport_08_06_archive_04").SQL
5190          Else
5200            strSQL = "SELECT ledger.accountno, ledger.journaltype, " & _
                  "Sum(IIf([ledger].[icash]<0,0,[ledger].[icash])) AS PositiveIcash, " & _
                  "Sum(IIf([ledger].[pcash]<0,0,[ledger].[pcash])) AS PositivePcash, Sum(ledger.cost) AS PositiveCost " & _
                  "FROM ledger " & _
                  "WHERE (((ledger.journaltype)='Dividend' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Interest' " & _
                  "OR (ledger.journaltype)='Sold' OR (ledger.journaltype)='Withdrawn' OR (ledger.journaltype)='Received') " & _
                  "AND ((ledger.icash)>=0) AND ((ledger.pcash)>=0) " & _
                  "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') AND ((ledger.pcash)>=0) "
5210            strSQL = strSQL & "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)>=0) " & _
                  "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)<0)) " & _
                  "GROUP BY ledger.accountno, ledger.journaltype " & _
                  "HAVING (((ledger.accountno) = [Reports]![rptAccountSummary]![accountno]));"
5220          End If
5230        End If
5240        dbs.QueryDefs(strQry_Inc).SQL = strSQL
            ' ** qryIncreases
            ' ** qryCourtReport_13

5250        If strOption = "Statements" Then
              ' ** Setting the qryDecreases SQL.
5260          If blnIsArchive = True And blnIsCourtRpt = True Then
5270            strSQL = dbs.QueryDefs("qryCourtReport_08_07_archive_04").SQL
5280          Else
5290            strSQL = "SELECT ledger.accountno, ledger.journaltype, Sum(IIf([ledger].[icash]>0,0,[ledger].[icash])) AS NegativeIcash, " & _
                  "Sum(IIf([ledger].[pcash]>0,0,[ledger].[pcash])) AS NegativePcash, Sum(ledger.cost) AS NegativeCost " & _
                  "FROM ledger INNER JOIN " & strQry_BalDate & " ON ledger.accountno = " & strQry_BalDate & ".accountno " & _
                  "WHERE (((ledger.journaltype)='Purchase' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Paid') " & _
                  "AND ((ledger.icash)<=0) AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Deposit') AND ((ledger.icash)<=0) AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) "
5300            strSQL = strSQL & "AND ((ledger.cost)<>0)) OR (((ledger.journaltype)='Misc.') AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)>0)) " & _
                  "GROUP BY ledger.accountno, ledger.journaltype " & _
                  "HAVING (((ledger.accountno) = [Reports]![rptAccountSummary]![accountno]));"
5310          End If
5320        Else
              ' ** Setting the qryDecreases SQL.
5330          If blnIsArchive = True And blnIsCourtRpt = True Then
5340            strSQL = dbs.QueryDefs("qryCourtReport_08_08_archive_04").SQL
5350          Else
5360            strSQL = "SELECT ledger.accountno, ledger.journaltype, " & _
                  "Sum(IIf([ledger].[icash]>0,0,[ledger].[icash])) AS NegativeIcash, " & _
                  "Sum(IIf([ledger].[pcash]>0,0,[ledger].[pcash])) AS NegativePcash, Sum(ledger.cost) AS NegativeCost " & _
                  "FROM ledger " & _
                  "WHERE (((ledger.journaltype)='Purchase' OR (ledger.journaltype)='Liability' OR (ledger.journaltype)='Paid') " & _
                  "AND ((ledger.icash)<=0) AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False)) " & _
                  "OR (((ledger.journaltype)='Deposit') AND ((ledger.icash)<=0) AND ((ledger.pcash)<=0) " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)<>0)) OR (((ledger.journaltype)='Misc.') AND ((ledger.pcash)<=0) "
5370            strSQL = strSQL & "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Misc.') AND ((ledger.icash)<=0) " & _
                  "AND ((ledger.transdate) Between #" & strStartDate & "# " & _
                  "AND #" & strEndDate & "#) " & _
                  "AND ((ledger.ledger_HIDDEN)=False)) OR (((ledger.journaltype)='Cost Adj.') " & _
                  "AND ((ledger.transdate)>=CDate(Format([" & strQry_BalDate & "].[" & strMaxDateFld & "],'mm/dd/yyyy'))+1 " & _
                  "AND (ledger.transdate)<=#" & strEndDate & "#) AND ((ledger.ledger_HIDDEN)=False) " & _
                  "AND ((ledger.cost)>0)) " & _
                  "GROUP BY ledger.accountno, ledger.journaltype " & _
                  "HAVING (((ledger.accountno) = [Reports]![rptAccountSummary]![accountno]));"
5380          End If
5390        End If
5400        dbs.QueryDefs(strQry_Dec).SQL = strSQL
            ' ** qryDecreases
            ' ** qryCourtReport_14

            ' ** Setup the qryAssetList query.
5410        If strAccountNo <> "All" Then
              ' ** Specific Account.
              ' ** VGC 11/25/2009: Ћ='90',-1,1)Л to Ћ='90',1,1)Л.
              ' ** VGC 12/04/2010: Added legalname, currentDate.
5420          If blnIsArchive = True And blnIsCourtRpt = True Then
5430            strSQL = dbs.QueryDefs("qryCourtReport_08_09_archive_01").SQL
5440          Else
5450            strSQL = "SELECT ActiveAssets.assetno, " & _
                  "masterasset.description AS MasterAssetDescription, " & _
                  "masterasset.due, masterasset.rate, Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, " & _
                  "Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface])) * " & _
                  "IIf([assettype].[assettype] = '90',1,1) AS TotalShareface, account.accountno, account.shortname, " & _
                  "account.legalname, assettype.assettype, assettype_description, " & _
                  "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
                  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
                  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc, " & _
                  "account.icash, account.pcash, " & CoInfo & ", " & _
                  "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX, " & _
                  "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX, " & _
                  "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX, masterasset.currentDate "
5460            strSQL = strSQL & "FROM account LEFT JOIN ((masterasset RIGHT JOIN ActiveAssets " & _
                  "ON masterasset.assetno = ActiveAssets.assetno) LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) " & _
                  "ON account.accountno = ActiveAssets.accountno " & _
                  "GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, account.accountno, " & _
                  "account.shortname, account.legalname, assettype.assettype, assettype_description, " & _
                  "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
                  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
                  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), account.icash, " & _
                  "account.pcash, IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), " & _
                  "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]), " & _
                  "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]), account.accountno, masterasset.currentDate " & _
                  "HAVING (((account.accountno) = '" & strAccountNo & "'));"
5470            strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.
5480          End If
5490        Else
              ' ** All.
              ' ** VGC 11/25/2009: Ћ='90',-1,1)Л to Ћ='90',1,1)Л.
              ' ** VGC 12/04/2010: Added legalname.
5500          If blnIsArchive = True And blnIsCourtRpt = True Then
5510            strSQL = dbs.QueryDefs("qryCourtReport_08_10_archive_01").SQL
5520          Else
5530            strSQL = "SELECT ActiveAssets.assetno, " & _
                  "masterasset.description AS MasterAssetDescription, " & _
                  "masterasset.due, masterasset.rate, Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, " & _
                  "Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface])) * " & _
                  "IIf([assettype].[assettype] = '90',1,1) AS TotalShareface, account.accountno, account.shortname, " & _
                  "account.legalname, assettype.assettype, assettype_description, " & _
                  "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
                  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
                  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc, " & _
                  "account.icash, account.pcash, masterasset.currentDate, " & CoInfo & ", " & _
                  "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX, " & _
                  "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX, " & _
                  "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX "
5540            strSQL = strSQL & "FROM account LEFT JOIN ((masterasset RIGHT JOIN ActiveAssets " & _
                  "ON masterasset.assetno = ActiveAssets.assetno) LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) ON " & _
                  "account.accountno = ActiveAssets.accountno " & _
                  "GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, account.accountno, " & _
                  "account.shortname, account.legalname, assettype.assettype, assettype_description, " & _
                  "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
                  "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
                  "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), account.icash, " & _
                  "account.pcash, IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), " & _
                  "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]), " & _
                  "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]), account.accountno, masterasset.currentDate;"
5550            strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.
5560          End If
5570        End If
5580        dbs.QueryDefs("qryAssetList").SQL = strSQL
            ' ** qryCurrentTotalMarketValue comes after qryAssetList has been written.
            ' ** qryTransRangeTotals uses qryMaxBalDates!

5590      End If ' ** <> StatementTransactions.

5600      dbs.Close

5610    End If

EXITP:
5620    Set rst = Nothing
5630    Set qdf = Nothing
5640    Set dbs = Nothing
5650    SetDateSpecificSQL = intRetVal
5660    Exit Function

ERRH:
5670    intRetVal = -9
5680    Select Case ERR.Number
        Case Else
5690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5700    End Select
5710    Resume EXITP

End Function

Public Function strFieldTypeToText(intType As Integer) As String
' Returns field type as a string
' NOTE: this function does NOT do anything with autoincrement fields,
' as handling them is not needed at this point

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "strFieldTypeToText"

5810    Select Case intType
        Case dbBoolean
5820      strFieldTypeToText = "boolean"
5830    Case dbByte
5840      strFieldTypeToText = "byte"
5850    Case dbCurrency
5860      strFieldTypeToText = "currency"
5870    Case dbDate
5880      strFieldTypeToText = "date"
5890    Case dbDouble
5900      strFieldTypeToText = "double"
5910    Case dbInteger
5920      strFieldTypeToText = "integer"
5930    Case dbLong
5940      strFieldTypeToText = "long"
          '###        Case dbNumeric
          '###        strFieldTypeToText = "numeric"
5950    Case dbText
5960      strFieldTypeToText = "text"
          '###       Case dbTime
          '###           strFieldTypeToText = "time"
5970    End Select

EXITP:
5980    Exit Function

ERRH:
5990    Select Case ERR.Number
        Case Else
6000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6010    End Select
6020    Resume EXITP

End Function

Public Function strNextDelimitedResult(strString As String, strDelimiter As String) As String
' ** Returns the string from the start of strString through the
' ** end of the string or to strDelimiter, whichever comes first.

6100  On Error GoTo ERRH

        Dim intLoop As Integer
        Dim strResult As String
        Dim blnOut As Boolean

        Const THIS_PROC As String = "strNextDelimitedResult"

6110    blnOut = False
6120    strResult = ""
6130    intLoop = 1
6140    Do While Not blnOut
6150      If Mid(strString, intLoop, 1) = strDelimiter Then
6160        blnOut = True
6170      Else
6180        strResult = strResult & Mid(strString, intLoop, 1)
6190      End If
6200      intLoop = intLoop + 1
6210      If intLoop > Len(strString) Then
6220        blnOut = True
6230      End If
6240    Loop

EXITP:
6250    strNextDelimitedResult = strResult
6260    Exit Function

ERRH:
6270    Select Case ERR.Number
        Case Else
6280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6290    End Select
6300    Resume EXITP

End Function

Public Function SQLFormatStr(varQuantity As Variant, intType As Integer) As Variant

6400  On Error GoTo ERRH

        Dim strReturn As Variant

        Const THIS_PROC As String = "SQLFormatStr"

6410    strReturn = Null

6420    If Not IsNull(varQuantity) Then
6430      Select Case intType
          Case dbText
6440        If InStr(varQuantity, Chr(39)) > 0 Then  ' ** Apostrophe, single-quote.
6450          strReturn = Chr(34) & varQuantity & Chr(34)
6460        Else
6470          strReturn = "'" & varQuantity & "'"
6480        End If
6490      Case dbLong, dbSingle, dbInteger, dbCurrency, dbDouble
6500        strReturn = str(varQuantity)
6510      Case dbDate
6520        strReturn = "#" & Format(varQuantity, "mm/dd/yyyy hh:mm:ss") & "#"
6530      Case dbBoolean
6540        strReturn = IIf(varQuantity, "True", "False")
6550      End Select
6560    Else
6570      strReturn = "Null"
6580    End If

EXITP:
6590    SQLFormatStr = strReturn
6600    Exit Function

ERRH:
6610    strReturn = Null
6620    Select Case ERR.Number
        Case Else
6630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6640    End Select
6650    Resume EXITP

End Function

Public Function CheckFirstAcctBal(dbs As DAO.Database, strAccountNo As String, strOption As String, strEndDate As String) As Integer
' ** There was not a balance date prior to the requested date.
' ** So, one must be created. Lets update the Initial Balance record
' ** to reflect the ending of the previous month;
' ** making like the account had been created at that time.
' ** THIS DOES NOT UPDATE THE ACCOUNT TABLE WITH CURRENT DATA!
' ** SEE ALSO: modStatementParamFuncs1.ChkFirstBal().
' ** Called by:
' **   SetDateSpecificSQL(), above, once each for All, Specific.
' ** Return codes:
' **    0  Success.
' **   -2  No data.
' **   -4  Date criteria not met.
' **   -9  Error.

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "CheckFirstAcctBal"

        Dim qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strNewDate As String
        Dim strSQL As String
        Dim datStartDate_Local As Date, datEndDate_Local As Date
        Dim lngAssetNo As Long, strJournalType As String, lngJrnlNum As Long
        Dim blnIncludeArchive As Boolean
        Dim varTmp00 As Variant
        Dim intRetVal As Integer

6710    intRetVal = 0

        ' ** I think it should always include it.
6720    blnIncludeArchive = True

        'strSQL = "SELECT ledger.accountno, Min(ledger.transdate) AS MinOftransdate " & _
        '  "FROM ledger " & _
        '  "GROUP BY ledger.accountno " & _
        '  "HAVING (((ledger.accountno) = '" & strAccountNo & "'));"

6730    Select Case blnIncludeArchive
        Case True
          ' ** LedgerArchive, grouped by accountno, with Min(transdate), by specified [actno].
6740      Set qdf = dbs.QueryDefs("qryStatementParameters_33_02")
6750      With qdf.Parameters
6760        ![actno] = strAccountNo
6770      End With
6780      Set rst = qdf.OpenRecordset
6790      If rst.BOF = True And rst.EOF = True Then
6800        rst.Close
6810        Set rst = Nothing
6820        Set qdf = Nothing
            ' ** Ledger, grouped by accountno, with Min(transdate), by specified [actno].
6830        Set qdf = dbs.QueryDefs("qryStatementParameters_33_01")
6840        With qdf.Parameters
6850          ![actno] = strAccountNo
6860        End With
6870        Set rst = qdf.OpenRecordset
6880        If rst.BOF = True And rst.EOF = True Then
              ' ** The account has no transactions.
6890          intRetVal = -2
6900          rst.Close
6910          Set rst = Nothing
6920          Set qdf = Nothing
6930        Else
6940          rst.MoveFirst
6950          datStartDate_Local = CDate(rst![MinOftransdate])
6960          rst.Close
6970          Set rst = Nothing
6980          Set qdf = Nothing
6990        End If
7000      Else
7010        rst.MoveFirst
7020        datStartDate_Local = CDate(rst![MinOftransdate])
7030        rst.Close
7040        Set rst = Nothing
7050        Set qdf = Nothing
7060      End If
7070    Case False
          ' ** Ledger, grouped by accountno, with Min(transdate), by specified [actno].
7080      Set qdf = dbs.QueryDefs("qryStatementParameters_33_01")
7090      With qdf.Parameters
7100        ![actno] = strAccountNo
7110      End With
7120      Set rst = qdf.OpenRecordset
7130      If rst.BOF = True And rst.EOF = True Then
            ' ** The account has no transactions.
7140        intRetVal = -2
7150        rst.Close
7160        Set rst = Nothing
7170        Set qdf = Nothing
7180      Else
7190        rst.MoveFirst
7200        datStartDate_Local = CDate(rst![MinOftransdate])
7210        rst.Close
7220        Set rst = Nothing
7230        Set qdf = Nothing
7240      End If
7250    End Select

7260    If intRetVal = 0 Then

7270      If IsNull(gvarCrtRpt_FL_SpecData) = True Then
            ' ** All other situations proceed normally.
7280        If datStartDate_Local > CDate(strEndDate) Then
              ' ** No transcations prior to date submitted.
7290          intRetVal = -4
7300        End If
7310      Else
7320        If datStartDate_Local <= CDate(strEndDate) Then
              ' ** No need for intervention; proceed normally.
7330        Else
              ' ** If this is Florida, and there's nothing before the specified start date (as in a new installation),
              ' ** then this criteria spits back 'no data', cancelling the rest of the report procedure.
              ' ** It SHOULD just continue with a zero beginning balance! But it doesn't.
              ' ** California comes through twice in the same way, but has a special section to handle that:
              ' **   12200  If blnContinue = False And intRetVal_BuildAssetListInfo = -2 Then ...
              ' ** National Standards doesn't hit this at all.
              ' ** Example:
              ' **   rst![MinOftransdate] = 01/01/2009
              ' **   strEndDate = 12/31/2008 (because this is supposed to be a beginning balance)
              ' ** Logic:
              ' **   If CDate(rst![MinOftransdate]) >= Forms("frmRpt_CourtReports_FL").DateStart And _
              ' **       CDate(rst![MinOftransdate]) <= Forms("frmRpt_CourtReports_FL").DateEnd Then ...

              ' ** At this point, gvarCrtRpt_FL_SpecData is initialized to 1.

              ' ** Check the Ledger again, specifically for this section of code.
              ' ** Ledger, grouped by accountno, with Min(transdate), Min(journalno), cnt.
7340          Set qdf = dbs.QueryDefs("qryLedger_10")  '![qryCourtReport_FL_00_New_10_01]
7350          Set rst = qdf.OpenRecordset
7360          rst.MoveFirst
7370          rst.FindFirst "[accountno] = '" & strAccountNo & "'"
7380          If rst.NoMatch = False Then
7390            If rst![transdate] = datStartDate_Local Then
                  ' ** Nothing I really need to do with this info, other than just to confirm.
                  ' **   ![accountno]  Group By
                  ' **   ![transdate]  Min
                  ' **   ![journalno]  Min
                  ' **   ![cnt]        Count(journalno)
7400              rst.Close
7410            Else
                  ' ** The queries are essentially the same; shouldn't get here.
7420              intRetVal = -2
7430              rst.Close
7440            End If
7450          Else
                ' ** Shouldn't get here, because the 1st query, above, would've caught it.
7460            intRetVal = -2
7470            rst.Close
7480          End If

7490          If intRetVal <> -2 Then
7500            If IsLoaded("frmRpt_ArchivedTransactions", acForm) = True Then  ' ** Module Function: modFileUtilities.
7510              datStartDate_Local = CDate(Forms("frmRpt_ArchivedTransactions").DateStart)
7520              datEndDate_Local = CDate(Forms("frmRpt_ArchivedTransactions").DateEnd)
7530            ElseIf IsLoaded("frmRpt_CourtReports_CA", acForm) = True Then  ' ** Module Function: modFileUtilities.
7540              datStartDate_Local = CDate(Forms("frmRpt_CourtReports_CA").DateStart)
7550              datEndDate_Local = CDate(Forms("frmRpt_CourtReports_CA").DateEnd)
7560            ElseIf IsLoaded("frmCourReportMenu_FL", acForm) = True Then  ' ** Module Function: modFileUtilities.
7570              datStartDate_Local = CDate(Forms("frmRpt_CourtReports_FL").DateStart)
7580              datEndDate_Local = CDate(Forms("frmRpt_CourtReports_FL").DateEnd)
7590            ElseIf IsLoaded("frmRpt_CourtReports_NS", acForm) = True Then  ' ** Module Function: modFileUtilities.
7600              datStartDate_Local = CDate(Forms("frmRpt_CourtReports_NS").DateStart)
7610              datEndDate_Local = CDate(Forms("frmRpt_CourtReports_NS").DateEnd)
7620            ElseIf IsLoaded("frmRpt_CourtReports_NY", acForm) = True Then  ' ** Module Function: modFileUtilities.
7630              datStartDate_Local = CDate(Forms("frmRpt_CourtReports_NY").DateStart)
7640              datEndDate_Local = CDate(Forms("frmRpt_CourtReports_NY").DateEnd)
7650            ElseIf IsLoaded("frmRpt_TransactionsByType", acForm) = True Then  ' ** Module Function: modFileUtilities.
7660              datStartDate_Local = CDate(Forms("frmRpt_TransactionsByType").TransDateStart)
7670              datEndDate_Local = CDate(Forms("frmRpt_TransactionsByType").TransDateEnd)
7680            ElseIf IsLoaded("frmStatementParameters", acForm) = True Then  ' ** Module Function: modFileUtilities.
7690              datStartDate_Local = CDate(Forms("frmStatementParameters").TransDateStart)
7700              datEndDate_Local = CDate(Forms("frmStatementParameters").TransDateEnd)
7710            Else
7720              datStartDate_Local = CDate("01/01/1900")
7730              datEndDate_Local = CDate(Date)
7740            End If
                ' ** Ledger, grouped by accountno, assetno, with journaltype1, journaltype2, by specified [actno], [datbeg], [datend].
7750            Set qdf = dbs.QueryDefs("qryLedger_11")  '![qryCourtReport_FL_00_New_10_02]
7760            With qdf.Parameters
7770              ![actno] = strAccountNo
7780              ![datbeg] = datStartDate_Local
7790              ![datEnd] = datEndDate_Local
7800            End With
7810            Set rst = qdf.OpenRecordset
7820            lngAssetNo = 0&
7830            If rst.BOF = True And rst.EOF = True Then
                  ' ** Shouldn't get here!
7840              intRetVal = -2
7850              rst.Close
7860            Else
                  ' ** If there's an Asset, use it, otherwise whatever.
                  ' **   ![accountno]
                  ' **   ![assetno]
                  ' **   ![journaltype1]
                  ' **   ![journaltype2]
7870            End If
7880          End If

7890          If intRetVal <> -2 Then

7900            rst.MoveFirst
7910            rst.FindFirst "[assetno] <> 0"
7920            If rst.NoMatch = False Then
7930              lngAssetNo = rst![assetno]
7940            Else
7950              rst.MoveFirst
7960              lngAssetNo = 0&
7970            End If
7980            strJournalType = rst![journaltype1]
7990            rst.Close

                ' ** Ledger, grouped, by specified [actno], [datbeg], [datend], [astno], [jtyp].
8000            Set qdf = dbs.QueryDefs("qryLedger_12")  '![qryCourtReport_FL_00_New_10_03]
8010            With qdf.Parameters
8020              ![actno] = strAccountNo
8030              ![datbeg] = datStartDate_Local
8040              ![datEnd] = datEndDate_Local
8050              ![astno] = lngAssetNo
8060              ![jtyp] = strJournalType
8070            End With
8080            Set rst = qdf.OpenRecordset
8090            rst.MoveFirst
8100            lngJrnlNum = rst![journalno]
8110            rst.Close

                ' ** Append qryLedger_13 to Ledger, by specified [jno], [jnox], [datspec]
8120            Set qdf = dbs.QueryDefs("qryLedger_14")  '![qryCourtReport_FL_00_New_10_05]
8130            With qdf.Parameters
8140              ![jno] = lngJrnlNum
8150              gvarCrtRpt_FL_SpecData = 999999999
8160              ![jnox] = gvarCrtRpt_FL_SpecData  ' ** I think this should be safe! (Though make sure it's gone when this printing is over!)
8170              ![datspec] = (datStartDate_Local - 30)  ' **                                       {see ChkSpecLedgerEntry()}
8180            End With
8190            qdf.Execute

                ' ** Now rerun the query that got us here in the first place, and reset datStartDate_Local.
8200            strSQL = "SELECT ledger.accountno, Min(ledger.transdate) AS MinOftransdate " & _
                  "FROM ledger " & _
                  "GROUP BY ledger.accountno " & _
                  "HAVING (((ledger.accountno) = '" & strAccountNo & "'));"

8210            Set rst = dbs.OpenRecordset(strSQL)
8220            If rst.BOF = True And rst.EOF = True Then
                  ' ** The account has no transactions.
8230              intRetVal = -2
8240            Else
8250              intRetVal = 0
8260              datStartDate_Local = CDate(rst![MinOftransdate])
8270            End If

                ' ** DON'T FORGET TO DELETE THE DUMMY LEDGER RECORD (gvarCrtRpt_FL_SpecData) WHEN THIS PRINTING IS OVER!
                ' ** SEE ChkSpecLedgerEntry().
8280          End If
8290        End If
8300      End If

8310      If intRetVal = 0 Then

            ' ** There are transactions that are before the requested date.
            ' ** So, let's figure out what date we need to update the record with.
8320        Select Case Format(datStartDate_Local, "m")
            Case "1"
8330          strNewDate = "12/31/" & CStr(CInt(Format(datStartDate_Local, "yyyy")) - 1)
8340        Case "2"
8350          strNewDate = "01/31/" & Format(datStartDate_Local, "yyyy")
8360        Case "3"
8370          strNewDate = "02/" & Format(CDate("03/01/" & Format(datStartDate_Local, "yyyy")) - 1, "dd") & "/" & Format(datStartDate_Local, "yyyy")
8380        Case "4"
8390          strNewDate = "03/31/" & Format(datStartDate_Local, "yyyy")
8400        Case "5"
8410          strNewDate = "04/30/" & Format(datStartDate_Local, "yyyy")
8420        Case "6"
8430          strNewDate = "05/31/" & Format(datStartDate_Local, "yyyy")
8440        Case "7"
8450          strNewDate = "06/30/" & Format(datStartDate_Local, "yyyy")
8460        Case "8"
8470          strNewDate = "07/31/" & Format(datStartDate_Local, "yyyy")
8480        Case "9"
8490          strNewDate = "08/31/" & Format(datStartDate_Local, "yyyy")
8500        Case "10"
8510          strNewDate = "09/30/" & Format(datStartDate_Local, "yyyy")
8520        Case "11"
8530          strNewDate = "10/31/" & Format(datStartDate_Local, "yyyy")
8540        Case "12"
8550          strNewDate = "11/30/" & Format(datStartDate_Local, "yyyy")
8560        End Select

8570        strSQL = "UPDATE Balance SET Balance.[balance date] = #" & strNewDate & "# " & _
              "WHERE (((Balance.accountno) = '" & strAccountNo & "') AND ((Balance.icash)=0) " & _
              "AND ((Balance.pcash)=0) AND ((Balance.cost)=0) AND ((Balance.TotalMarketValue)=0) AND ((Balance.AccountValue)=0));"

8580        DoCmd.SetWarnings False
8590        dbs.Execute strSQL
8600        DoCmd.SetWarnings True

8610      Else
8620        If intRetVal = -4 And strOption = "Statements" Then
8630          If IsLoaded("frmRpt_CourtReports_NY", acForm) = True Then  ' ** Module Function: modFileUtilities.
                ' ** Put one in as of the end of the prior year.
                ' ********
                ' ** strEndDate IS FED FROM SetDateSpecificSQL(), ABOVE.
                ' ** strEndDate THERE IS FED FROM BuildAssetListInfo() IN frmRpt_CourtReports_NY.
                ' ** varStarDate THERE IS FED FROM, AMONG OTHERS, cmdPreview01_Click() AS (.DateStart - 1).
                ' ** SO WE'LL PUT A ZERO ENTRY IN FOR (.DateStart - 1)!
                ' ********
                ' ** Get the earliest one that's in there now.
8640            varTmp00 = DMin("[balance date]", "Balance", "[accountno] = '" & strAccountNo & "'")
8650            Select Case IsNull(varTmp00)
                Case True
                  ' ** Put one in.
8660              With dbs
8670                Set rst = .OpenRecordset("Balance", dbOpenDynaset, dbConsistent)
8680                With rst
8690                  .AddNew
8700                  ![accountno] = strAccountNo
8710                  ![balance date] = CDate(strEndDate)
8720                  ![ICash] = 0#  ' ** All these are Double.
8730                  ![PCash] = 0#
8740                  ![Cost] = 0#
8750                  ![TotalMarketValue] = 0#
8760                  ![AccountValue] = 0#
8770                  .Update
8780                  intRetVal = 0
8790                  .Close
8800                End With
8810                Set rst = Nothing
8820              End With
8830            Case False
8840              If varTmp00 <= CDate(strEndDate) Then
                    ' ** There's one there already?!
8850                intRetVal = 0
8860              Else
                    ' ** If the one found is within the report's range, and
                    ' ** all Zero's, then just move it back to the one we want.
8870                With dbs
8880                  Set rst = .OpenRecordset("Balance", dbOpenDynaset, dbConsistent)
8890                  With rst
8900                    .MoveLast
8910                    .MoveFirst
8920                    If varTmp00 <= CDate(Forms("frmRpt_CourtReports_NY").DateEnd) And varTmp00 > CDate(strEndDate) Then
8930                      .FindFirst "[accountno] = '" & strAccountNo & "' And [balance date] = #" & Format(varTmp00, "mm/dd/yyyy") & "#"
8940                      If .NoMatch = False Then
8950                        If ![ICash] = 0# And ![PCash] = 0# And ![Cost] = 0# And ![TotalMarketValue] = 0# And ![AccountValue] = 0# Then
8960                          .Edit
8970                          ![balance date] = CDate(strEndDate)
8980                          .Update
8990                          intRetVal = 0
9000                        Else
                              ' ** Who knows?
9010                          .AddNew
9020                          ![accountno] = strAccountNo
9030                          ![balance date] = CDate(strEndDate)
9040                          ![ICash] = 0#  ' ** All these are Double.
9050                          ![PCash] = 0#
9060                          ![Cost] = 0#
9070                          ![TotalMarketValue] = 0#
9080                          ![AccountValue] = 0#
9090                          .Update
9100                          intRetVal = 0
9110                        End If
9120                      Else
                            ' ** Shouldn't happen!
9130                        intRetVal = -9
9140                      End If
9150                    Else
                          ' ** After the period?
9160                      .AddNew
9170                      ![accountno] = strAccountNo
9180                      ![balance date] = CDate(strEndDate)
9190                      ![ICash] = 0#  ' ** All these are Double.
9200                      ![PCash] = 0#
9210                      ![Cost] = 0#
9220                      ![TotalMarketValue] = 0#
9230                      ![AccountValue] = 0#
9240                      .Update
9250                      intRetVal = 0
9260                    End If
9270                    .Close
9280                  End With
9290                  Set rst = Nothing
9300                End With
9310              End If
9320            End Select
9330          End If  ' ** IsLoaded().
9340        End If  ' ** intRetVal, strOption.
9350      End If  ' ** intRetVal.
9360    End If  ' ** intRetVal.

EXITP:
9370    Set rst = Nothing
9380    Set qdf = Nothing
9390    CheckFirstAcctBal = intRetVal
9400    Exit Function

ERRH:
9410    intRetVal = -9
9420    Select Case ERR.Number
        Case Else
9430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9440    End Select
9450    Resume EXITP

End Function

Public Function ChkSpecLedgerEntry() As Boolean
' ** This just makes sure a special dummy Ledger record used
' ** by the Florida Court Reports isn't still hanging around.
' ** I believe it's a worry, because SetDateSpecificSQL()
' ** and CheckFirstAcctBal(), above, both use it,
' ** and they're also used by lots of other reports!

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkSpecLedgerEntry"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

9510    blnRetVal = True
9520    If IsNull(gvarCrtRpt_FL_SpecData) = False Then
9530      If gvarCrtRpt_FL_SpecData <> 1 Then  ' ** Make sure it's not the just-initialized one.
9540        Set dbs = CurrentDb
9550        With dbs
              ' ** Delete Ledger, by specified [jnox] (the 999999999 one).
9560          Set qdf = .QueryDefs("qryLedger_15")  '![qryCourtReport_FL_00_New_10_06]
9570          With qdf.Parameters
9580            ![jnox] = gvarCrtRpt_FL_SpecData
9590          End With
9600          qdf.Execute
9610          .Close
9620        End With
9630      End If
9640      gvarCrtRpt_FL_SpecData = Null
9650    End If

EXITP:
9660    Set qdf = Nothing
9670    Set dbs = Nothing
9680    ChkSpecLedgerEntry = blnRetVal
9690    Exit Function

ERRH:
9700    blnRetVal = False
9710    Select Case ERR.Number
        Case Else
9720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9730    End Select
9740    Resume EXITP

End Function

Public Function ChkSpecLedgerEntryNull() As Boolean

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "ChkSpecLedgerEntryNull"

        Dim blnRetVal As Boolean

9810    blnRetVal = True

9820    gvarCrtRpt_FL_SpecData = Null

EXITP:
9830    ChkSpecLedgerEntryNull = blnRetVal
9840    Exit Function

ERRH:
9850    blnRetVal = False
9860    Select Case ERR.Number
        Case Else
9870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9880    End Select
9890    Resume EXITP

End Function

Public Function RecurXFer(varRecurItem As Variant) As Boolean

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "RecurXFer"

        Dim blnRetVal As Boolean

9910    blnRetVal = False

9920    If IsNull(varRecurItem) = False Then
9930      If Trim(varRecurItem) <> vbNullString Then
9940        If Trim(varRecurItem) = RECUR_I_TO_P Or Trim(varRecurItem) = RECUR_P_TO_I Then
9950          blnRetVal = True
9960        End If
9970      End If
9980    End If

EXITP:
9990    RecurXFer = blnRetVal
10000   Exit Function

ERRH:
10010   blnRetVal = False
10020   Select Case ERR.Number
        Case Else
10030     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10040   End Select
10050   Resume EXITP

End Function

Public Function GetLineNumber(strFormName As String, strKeyName As String) As Long
' ** Part of the Highlight Current Record
' ** procedure posted by James H Brooks on
' ** The Access Web: http://www.mvps.org/access/
' **
' ** The function "GetLineNumber" is modified from the Microsoft Knowledge Base
' ** (Q120913), the only difference here is that the following items have been hard
' ** coded: frm, strKeyName, varKeyValue. This was done to add a slight performance
' ** increase. Change strKeyName and varKeyValue to reflect the key in your table.
' **
' ** ctlBack.ControlSource: =IIf([SelTop]=[ctlCurrentLine],"лллллллллллл",Null)
' ** ctlCurrentLine.ControlSource: =GetLineNumber("frmJournal_Columns_Sub","journal_id")

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "GetLineNumber"

        Dim rst As DAO.Recordset, frm As Access.Form
        Dim varKeyValue As Variant
        Dim blnContinue As Boolean
        Dim lngRetVal As Long

10110   lngRetVal = 0&
10120   blnContinue = True

10130   If strFormName <> vbNullString And strKeyName <> vbNullString Then

10140     If InStr(strFormName, "_Sub") > 0 Then
10150       Select Case strFormName
            Case "frmStatementBalance_Sub"
10160         Set frm = Forms.frmStatementBalance.frmStatementBalance_Sub.Form
10170       Case "frmJournal_Columns_Sub"
              ' ** ctlCurrentLine:
              ' **   =GetLineNumber("frmJournal_Columns_Sub","JrnlCol_ID")
10180         Set frm = Forms.frmJournal_Columns.frmJournal_Columns_Sub.Form
10190       Case "Yyy_Sub"
              'Set frm = Forms.Xxx.Yyy_Sub.Form
10200       End Select
10210     Else
10220       Set frm = Forms(strFormName)
10230     End If

10240     varKeyValue = frm(strKeyName)

10250     If IsNull(varKeyValue) = False Then

10260       Set rst = frm.RecordsetClone

            ' ** Find the current record.
10270       Select Case rst.Fields(strKeyName).Type
            Case DB_INTEGER, DB_LONG, DB_CURRENCY, DB_SINGLE, DB_DOUBLE, DB_BYTE
              ' ** Find using numeric data type key value.
10280         If IsNumeric(varKeyValue) = True Then
10290           rst.FindFirst "[" & strKeyName & "] = " & varKeyValue
10300         End If
10310       Case DB_DATE
              ' ** Find using date data type key value.
10320         If IsDate(varKeyValue) = True Then
10330           rst.FindFirst "[" & strKeyName & "] = #" & varKeyValue & "#"
10340         End If
10350       Case DB_Text
              ' ** Find using text data type key value.
10360         rst.FindFirst "Rem_Apost([" & strKeyName & "]) = '" & Rem_Apost(varKeyValue) & "'"  ' ** Module Function: modStringFuncs.
10370       Case Else
10380         blnContinue = False
10390         MsgBox "Invalid key field data type!" & vbCrLf & vbCrLf & _
                "Module:" & vbTab & vbTab & "modUtilities" & vbCrLf & _
                "Sub/Function:" & vbTab & "GetLineNumber()" & vbCrLf & _
                "Line:" & vbTab & vbTab & "7970", vbInformation + vbOKOnly, "Data Type Error"
10400       End Select

            ' ** Loop backward, counting the lines.
10410       If blnContinue = True Then
10420         Do Until rst.BOF
10430           lngRetVal = lngRetVal + 1
10440           rst.MovePrevious
10450         Loop
10460       End If

10470       rst.Close

10480     Else
10490       blnContinue = False
10500     End If

10510   Else
10520     blnContinue = False
10530   End If

EXITP:
10540   Set frm = Nothing
10550   Set rst = Nothing
10560   GetLineNumber = lngRetVal
10570   Exit Function

ERRH:
10580   lngRetVal = 0&
10590   Select Case ERR.Number
        Case 438  ' ** Object doesn't support this property or method.
          ' ** Ignore, since it might be just the form closing.
10600   Case 3059  ' ** Operation Canceled by user (don't know why this sometimes triggers).
          ' ** Ignore.
10610   Case 3077  ' ** Syntax error (missing operator) in expression.
          ' ** Seems to happen when a key is pressed before the form settles down.
10620   Case Else
10630     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10640   End Select
10650   Resume EXITP

End Function

Public Function JrnlCol_Set() As Boolean

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlCol_Set"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim blnRetVal As Boolean

10710   blnRetVal = True

10720   Set dbs = CurrentDb
10730   With dbs

          ' ** Delete qryJournal_Columns_02c (tblJournal_Column, not in Journal, for Journal_ID <> NULL).
10740     Set qdf = dbs.QueryDefs("qryJournal_Columns_02d")  ' ** Entries since posted.
10750     qdf.Execute

          ' ** Delete tblJournalColumn, for Journal_ID <> Null, posted = True.
10760     Set qdf = dbs.QueryDefs("qryJournal_Columns_02a")  ' ** Journal entries not edited (to be replaced).
10770     qdf.Execute

          ' ** What remains:
          ' **   Journal_ID = Null: New, unfinished entries.
          ' **   Journal_ID <> Null, posted = False: Journal entries that have been edited.

          ' ** Append qryJournal_Columns_07 (qryJournal_Columns_06 (Journal, not
          ' ** in tblJournal_Column), with add'l fields) to tblJournal_Column.
10780     Set qdf = .QueryDefs("qryJournal_Columns_08")  ' ** Appended with posted = True.
10790     qdf.Execute

          ' ** Update qryJournal_Columns_27_02 (tblJournal_Column, linked to qryJournal_Columns_27_01
          ' ** (tblJournal_Memo, linked to Journal), JrnlMemo_Memo_new).
10800     Set qdf = .QueryDefs("qryJournal_Columns_27_03")
10810     qdf.Execute

10820     .Close
10830   End With  ' ** dbs.

EXITP:
10840   Set qdf = Nothing
10850   Set dbs = Nothing
10860   JrnlCol_Set = blnRetVal
10870   Exit Function

ERRH:
10880   blnRetVal = False
10890   Select Case ERR.Number
        Case Else
10900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10910   End Select
10920   Resume EXITP

End Function

Public Sub JrnlEmptyRec_Chk()

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlEmptyRec_Chk"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long

11010   If gblnAdmin = True Then
11020     lngRecs = 0&
11030     Set dbs = CurrentDb
11040     With dbs
            ' ** Journal, just completely empty records.
11050       Set qdf = .QueryDefs("qryJournal_04")
11060       Set rst = qdf.OpenRecordset
11070       With rst
11080         If .BOF = True And .EOF = True Then
                ' ** Everything's OK.
11090         Else
11100           .MoveLast
11110           lngRecs = .RecordCount
11120         End If
11130         .Close
11140       End With
11150       If lngRecs > 0& Then
11160         Beep
11170         MsgBox "There appears to be a completely empty record in the Journal," & vbCrLf & _
                "which can cause problems when creating new entries." & vbCrLf & vbCrLf & _
                "To assure the Journal is clean, post all pending transactions, then clear all transactions." & vbCrLf & vbCrLf & _
                "If the problem persists, contact Delta Data, Inc.", vbCritical + vbOKOnly, "Empty Journal Record"
11180       End If
11190       .Close
11200     End With
11210   End If

EXITP:
11220   Set rst = Nothing
11230   Set qdf = Nothing
11240   Set dbs = Nothing
11250   Exit Sub

ERRH:
11260   Select Case ERR.Number
        Case Else
11270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11280   End Select
11290   Resume EXITP

End Sub

Public Function SetOption_Run(Optional varOnOpen As Variant) As Boolean

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "SetOption_Run"

        Dim prp As Object
        Dim blnOnOpen As Boolean
        Dim blnRetVal As Boolean

11310   blnRetVal = True

11320   If IsMissing(varOnOpen) = True Then
11330     blnOnOpen = True  ' ** Set to Trust Accountant's run settings.
11340   Else
11350     blnOnOpen = varOnOpen
11360   End If

        ' ** Startup Properties:
        ' **   Text In Startup Dialog Box        Property Name
        ' **   ================================  ========================
        ' **   Application Title                 AppTitle
        ' **   Application Icon                  AppIcon
        ' **   Display Form/Page                 StartupForm
        ' **   Display Database Window           StartupShowDBWindow
        ' **   Display Status Bar                StartupShowStatusBar
        ' **   Menu Bar                          StartupMenuBar
        ' **   Shortcut Menu Bar                 StartupShortcutMenuBar
        ' **   Allow Full Menus                  AllowFullMenus
        ' **   Allow Default Shortcut Menus      AllowShortcutMenus
        ' **   Allow Built-In Toolbars           AllowBuiltInToolbars
        ' **   Allow Toolbar/Menu Changes        AllowToolbarChanges
        ' **   Allow Viewing Code After Error    AllowBreakIntoCode
        ' **   Use Access Special Keys           AllowSpecialKeys

11370   Select Case blnOnOpen
        Case True

          ' ** 1. Display Database Window.
          ' ** No need to save, because it only applies to this database.  ' ** Check box: dbBoolean.
11380 On Error Resume Next
11390     CurrentDb.Properties("StartupShowDBWindow") = False  ' ** Also covers Navigation Pane.
11400     If ERR.Number <> 0 Then
11410 On Error GoTo ERRH
11420       Set prp = CurrentDb.CreateProperty("StartupShowDBWindow", dbBoolean, False)
11430 On Error Resume Next
            ' ** Error: 3367  Cannot append. An object with that name already exists in the collection.
11440       CurrentDb.Properties.Append prp
11450 On Error GoTo ERRH
11460     Else
11470 On Error GoTo ERRH
11480     End If

          ' ** 2. Arrow Key Behavior.
          ' ** No need to save, because I don't think the user knows or cares.  ' ** Option Group: dbInteger.
11490     Application.SetOption "Arrow Key Behavior", 1  ' ** Next character.

          ' ** 3. AutoIndex On Import/Create.
          ' ** No need to save, because I don't think the user knows or cares.  ' ** Text Box: dbText.
11500     Application.SetOption "AutoIndex On Import/Create", vbNullString

          ' ** 4. Default Database Directory.
11510     gstrDefaultDatabaseDirectory = GetOption("Default Database Directory")  ' ** Text Box: dbText.
11520     Application.SetOption "Default Database Directory", CurrentAppPath  ' ** Module Function: modFileUtilities.

          ' ** 5. Default Open Mode for Databases.
11530     gintDefaultOpenMode = GetOption("Default Open Mode for Databases")  ' ** Option Group: dbInteger.
11540     Application.SetOption "Default Open Mode for Databases", 0  ' ** Set to shared.

          ' ** 6. Default Record Locking.
11550     gintDefaultRecordLocking = GetOption("Default Record Locking")  ' ** Option Group: dbInteger.
11560     Application.SetOption "Default Record Locking", 2  ' ** Set to record level locking.

          ' ** 7. Confirm Record Changes.
11570     gblnConfirmRecordChanges = GetOption("Confirm Record Changes")  ' ** Check box: dbBoolean.
11580     Application.SetOption "Confirm Record Changes", False  ' ** Stop Record confirmation messages.

          ' ** 8. Confirm Action Queries.
11590     gblnConfirmActionQueries = GetOption("Confirm Action Queries")  ' ** Check box: dbBoolean.
11600     Application.SetOption "Confirm Action Queries", False  ' ** Set to record level locking.

          ' ** 9. Confirm Document Deletions.
11610     gblnConfirmDocumentDeletions = GetOption("Confirm Document Deletions")  ' ** Check box: dbBoolean.
11620     Application.SetOption "Confirm Document Deletions", False  ' ** Stop deletion confirmation messages.

          ' ** 10. Compact On Close.
11630     gblnAutoCompact = GetOption("Auto Compact")  ' ** Check box: dbBoolean.
11640     Application.SetOption "Auto Compact", True  ' ** Compact MDE when closing.

          ' ** 11. Windows In Taskbar.
11650     gblnShowWindowsInTaskbar = GetOption("ShowWindowsInTaskbar")  ' ** Check box: dbBoolean.
11660     Application.SetOption "ShowWindowsInTaskbar", False  ' ** Too many taskbar items show up.

          ' ** 12. Show Hidden Objects.
11670     gblnShowHiddenObjects = GetOption("Show Hidden Objects")  ' ** Check box: dbBoolean.
11680     If CurrentUser <> "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
11690       Application.SetOption "Show Hidden Objects", False  ' ** Don't show hidden objects.
11700     Else
            ' ** So that when we log into a users computer, we see everything.
11710       Application.SetOption "Show Hidden Objects", True  ' ** Do show hidden objects.
11720     End If

          ' ** 13. Remove Help question.
11730     Application.CommandBars.DisableAskAQuestionDropdown = True

          ' ** 14. Access 2007 special handling.
          'SetOption_Access2007 True, THIS_PROC  ' ** Module Function: modXAccess_07_10_Funcs.

          ' ** 15. Vista/Win7 special handling.
          'If IsWinVista = True Then  ' ** Module Function: modOperSysInfoFuncs1.

          'End If

          ' ** 16. Perform Name AutoCorrect.
          'gblnPerformNameAutoCorrect = GetOption("Perform Name AutoCorrect")  ' ** Check box: dbBoolean.
          'Application.SetOption "Perform Name AutoCorrect", False  ' ** Don't AutoCorrect while running.

          ' ** 17. Track Name AutoCorrect Info.
          'gblnTrackNameAutoCorrectInfo = GetOption("Track Name AutoCorrect Info")  ' ** Check box: dbBoolean.
          'Application.SetOption "Track Name AutoCorrect Info", False  ' ** Don't AutoCorrect while running.

          ' ** 18. Log Name AutoCorrect Changes.
          'gblnLogNameAutoCorrectChanges = GetOption("Log Name AutoCorrect Changes")  ' ** Check box: dbBoolean.
          'Application.SetOption "Log Name AutoCorrect Changes", False  ' ** Don't AutoCorrect while running.

11740   Case False

          ' ** 1. Display Database Window.
          ' ** Not saved.

          ' ** 2. Arrow Key Behavior.
          ' ** Not saved.

          ' ** 3. AutoIndex On Import/Create.
          ' ** Not saved.

          ' ** 4. Default Database Directory.
11750     Application.SetOption "Default Database Directory", gstrDefaultDatabaseDirectory

          ' ** 5. Default Open Mode for Databases.
11760     Application.SetOption "Default Open Mode for Databases", gintDefaultOpenMode

          ' ** 6. Default Record Locking.
11770     Application.SetOption "Default Record Locking", gintDefaultRecordLocking

          ' ** 7. Confirm Record Changes.
11780     Application.SetOption "Confirm Record Changes", gblnConfirmRecordChanges

          ' ** 8. Confirm Action Queries.
11790     Application.SetOption "Confirm Action Queries", gblnConfirmActionQueries

          ' ** 9. Confirm Document Deletions.
11800     Application.SetOption "Confirm Document Deletions", gblnConfirmDocumentDeletions

          ' ** 10. Compact On Close.
11810     Application.SetOption "Auto Compact", gblnAutoCompact

          ' ** 11. Windows In Taskbar.
11820     Application.SetOption "ShowWindowsInTaskbar", gblnShowWindowsInTaskbar

          ' ** 13. Remove Help question.
          ' ** Leave it off.

          ' ** 12. Show Hidden Objects.
11830     Application.SetOption "Show Hidden Objects", gblnShowHiddenObjects

          ' ** 14. Access 2007 special handling.
11840     SetOption_Access2007 False, THIS_PROC  ' ** Module Function: modXAccess_07_10_Funcs.

          ' ** 15. Vista special handling.
          'If IsWinVista = True Then  ' ** Module Function: modOperSysInfoFuncs1.

          'End If

          ' ** 16. Perform Name AutoCorrect.
          'Application.SetOption "Perform Name AutoCorrect", gblnPerformNameAutoCorrect

          ' ** 17. Track Name AutoCorrect Info.
          'Application.SetOption "Track Name AutoCorrect Info", gblnTrackNameAutoCorrectInfo

          ' ** 18. Log Name AutoCorrect Changes.
          'Application.SetOption "Log Name AutoCorrect Changes", gblnLogNameAutoCorrectChanges

          'DOESN'T WORK!
          'Call AccessCloseButtonEnabled(False)  ' ** Module Procedure: modWindowFunctions.

11850   End Select

EXITP:
11860   Set prp = Nothing
11870   SetOption_Run = blnRetVal
11880   Exit Function

ERRH:
11890   blnRetVal = False
11900   Select Case ERR.Number
        Case Else
11910     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11920   End Select
11930   Resume EXITP

End Function

Public Function SetOption_Dev() As Boolean
' ** Developer's settings.
' ** Called by:
' **   mcrReset_Options

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "SetOption_Dev"

        Dim dbs As DAO.Database, prp As DAO.Property
        Dim blnFound As Boolean
        Dim blnRetVal As Boolean

12010   blnRetVal = True

12020   SetOption_Access2010 False  ' ** Module Function: modXAccess_07_10_Funcs.

        ' ** 1. Display Database Window.
        ' ** Evidently, this doesn't exist naturally until it's been manually set.
12030 On Error Resume Next
12040   CurrentDb.Properties("StartupShowDBWindow") = True  ' ** Also covers Navigation Pane.
12050 On Error GoTo ERRH

        'DoCmd.SelectObject acForm, , True
        ' ** 2. Arrow Key Behavior.
        ' ** 3. AutoIndex On Import/Create.
        ' ** 4. Default Database Directory.
        ' ** 5. Default Open Mode for Databases.
        'Application.SetOption "Default Open Mode for Databases", 0  ' ** gintDefaultOpenMode
        ' ** 6. Default Record Locking.
        'Application.SetOption "Default Record Locking", 2           ' ** gintDefaultRecordLocking
        ' ** 7. Confirm Record Changes.
12060   Application.SetOption "Confirm Record Changes", False        ' ** gblnConfirmRecordChanges
        ' ** 8. Confirm Action Queries.
12070   Application.SetOption "Confirm Action Queries", True         ' ** gblnConfirmActionQueries
        ' ** 9. Confirm Document Deletions.
12080   Application.SetOption "Confirm Document Deletions", True     ' ** gblnConfirmDocumentDeletions
        ' ** 10. Compact On Close.
12090   Application.SetOption "Auto Compact", False                  ' ** gblnAutoCompact
        ' ** 11. Windows In Taskbar.
        'Application.SetOption "ShowWindowsInTaskbar", False         ' ** gblnShowWindowsInTaskbar
        ' ** 12. Show Hidden Objects.
12100   Application.SetOption "Show Hidden Objects", True            ' ** gblnShowHiddenObjects
        ' ** 13. Remove Help question.
12110   Application.CommandBars.DisableAskAQuestionDropdown = True
        ' ** 14. Access 2007 special handling.
12120   SetOption_Access2007 False, THIS_PROC  ' ** Module Function: modXAccess_07_10_Funcs.
        ' ** 15. Vista special handling.
        ' ** 16. Perform Name AutoCorrect.
        'Application.SetOption "Perform Name AutoCorrect", True       ' ** gblnPerformNameAutoCorrect
        ' ** 17. Track Name AutoCorrect Info.
        'Application.SetOption "Track Name AutoCorrect Info", True    ' ** gblnTrackNameAutoCorrectInfo
        ' ** 18. Log Name AutoCorrect Changes.
        'Application.SetOption "Log Name AutoCorrect Changes", False  ' ** gblnLogNameAutoCorrectChanges

        ' ************************
        ' ** Startup Properties:
        ' ************************
12130   Set dbs = CurrentDb
12140   With dbs
          ' ** 1.  AppTitle.
          ' **       {Leave as-is}
          ' ** 2.  AppIcon.
          ' **       {Leave as-is}
          ' ** 3.  StartupForm.
          ' **       {Leave as-is}
          ' ** 4.  StartupShowDBWindow.
12150     blnFound = False
12160     For Each prp In .Properties
12170       With prp
12180         If .Name = "StartupShowDBWindow" Then
12190           blnFound = True
12200           .Value = True
12210           Exit For
12220         End If
12230       End With
12240     Next  ' ** prp.
12250     Set prp = Nothing
12260     If blnFound = False Then
12270       Set prp = .CreateProperty("StartupShowDBWindow", dbBoolean, True)
12280       .Properties.Append prp
12290       Set prp = Nothing
12300     End If
          ' ** 5.  StartupShowStatusBar.
12310     blnFound = False
12320     For Each prp In .Properties
12330       With prp
12340         If .Name = "StartupShowStatusBar" Then
12350           blnFound = True
12360           .Value = True
12370           Exit For
12380         End If
12390       End With
12400     Next  ' ** prp.
12410     Set prp = Nothing
12420     If blnFound = False Then
12430       Set prp = .CreateProperty("StartupShowStatusBar", dbBoolean, True)
12440       .Properties.Append prp
12450       Set prp = Nothing
12460     End If
          ' ** 6.  AllowFullMenus.
12470     blnFound = False
12480     For Each prp In .Properties
12490       With prp
12500         If .Name = "AllowFullMenus" Then
12510           blnFound = True
12520           .Value = True
12530           Exit For
12540         End If
12550       End With
12560     Next  ' ** prp.
12570     Set prp = Nothing
12580     If blnFound = False Then
12590       Set prp = .CreateProperty("AllowFullMenus", dbBoolean, True)
12600       .Properties.Append prp
12610       Set prp = Nothing
12620     End If
          ' ** 7.  AllowShortcutMenus.
12630     blnFound = False
12640     For Each prp In .Properties
12650       With prp
12660         If .Name = "AllowShortcutMenus" Then
12670           blnFound = True
12680           .Value = True
12690           Exit For
12700         End If
12710       End With
12720     Next  ' ** prp.
12730     Set prp = Nothing
12740     If blnFound = False Then
12750       Set prp = .CreateProperty("AllowShortcutMenus", dbBoolean, True)
12760       .Properties.Append prp
12770       Set prp = Nothing
12780     End If
          ' ** 8.  AllowBuiltInToolbars.
12790     blnFound = False
12800     For Each prp In .Properties
12810       With prp
12820         If .Name = "AllowBuiltInToolbars" Then
12830           blnFound = True
12840           .Value = True
12850           Exit For
12860         End If
12870       End With
12880     Next  ' ** prp.
12890     Set prp = Nothing
12900     If blnFound = False Then
12910       Set prp = .CreateProperty("AllowBuiltInToolbars", dbBoolean, True)
12920       .Properties.Append prp
12930       Set prp = Nothing
12940     End If
          ' ** 9.  AllowToolbarChanges.
12950     blnFound = False
12960     For Each prp In .Properties
12970       With prp
12980         If .Name = "AllowToolbarChanges" Then
12990           blnFound = True
13000           .Value = True
13010           Exit For
13020         End If
13030       End With
13040     Next  ' ** prp.
13050     Set prp = Nothing
13060     If blnFound = False Then
13070       Set prp = .CreateProperty("AllowToolbarChanges", dbBoolean, True)
13080       .Properties.Append prp
13090       Set prp = Nothing
13100     End If
          ' ** 10. AllowSpecialKeys.
13110     blnFound = False
13120     For Each prp In .Properties
13130       With prp
13140         If .Name = "AllowSpecialKeys" Then
13150           blnFound = True
13160           .Value = True
13170           Exit For
13180         End If
13190       End With
13200     Next  ' ** prp.
13210     Set prp = Nothing
13220     If blnFound = False Then
13230       Set prp = .CreateProperty("AllowSpecialKeys", dbBoolean, True)
13240       .Properties.Append prp
13250       Set prp = Nothing
13260     End If
          ' ** 11. AllowBreakIntoCode.
13270     blnFound = False
13280     For Each prp In .Properties
13290       With prp
13300         If .Name = "AllowBreakIntoCode" Then
13310           blnFound = True
13320           .Value = True
13330           Exit For
13340         End If
13350       End With
13360     Next  ' ** prp.
13370     Set prp = Nothing
13380     If blnFound = False Then
13390       Set prp = .CreateProperty("AllowBreakIntoCode", dbBoolean, True)
13400       .Properties.Append prp
13410       Set prp = Nothing
13420     End If
          ' ** 12. StartupShortcutMenuBar.
          ' **       {Not set}
          ' ** 13. StartupMenuBar.
          ' **       {Not set}
13430     .Close
13440   End With  ' ** dbs.
13450   Set prp = Nothing
13460   Set dbs = Nothing

13470   If IsAccess2007 = True Or IsAccess2010 = True Then  ' ** Module Functions: modXAccess_07_10_Funcs.
13480     DoCmd.ShowToolbar "Ribbon", acToolbarYes  ' ** Turn on the Ribbons.
13490     DoEvents
13500   End If

13510   If IsLoaded("frmMenu_Background", acForm) = True Then  ' ** Module Function: modFileUtilities.
13520     DoCmd.Close acForm, "frmMenu_Background"
13530     gblnDev_NoAppBackground = True
13540   End If

13550   CmdBars_Clipboard True  ' ** Module Function: modWindowFunctions.
13560   Scr  ' ** Module Function: modWindowFunctions.

        'DOESN'T WORK!
        'Call AccessCloseButtonEnabled(True)  ' ** Module Procedure: modWindowFunctions.

13570   OpenAllDatabases False  ' ** Module Procedure: modStartupFuncs.
13580   Scr  ' ** Module Function: modWindowFunctions.

13590   Beep

EXITP:
13600   Set prp = Nothing
13610   Set dbs = Nothing
13620   SetOption_Dev = blnRetVal
13630   Exit Function

ERRH:
13640   blnRetVal = False
13650   Select Case ERR.Number
        Case Else
13660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13670   End Select
13680   Resume EXITP

End Function

Public Function TodaysDate() As Date
' ** Make sure we absolutely, positively get today's date.

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "TodaysDate"

        Dim intPos01 As Integer
        Dim strTmp01 As String, dblTmp02 As Double
        Dim datRetVal As Date

13710   dblTmp02 = CDbl(Now())
13720   strTmp01 = CStr(dblTmp02)
13730   intPos01 = InStr(strTmp01, ".")
13740   If intPos01 > 0 Then strTmp01 = Left(strTmp01, (intPos01 - 1))
13750   dblTmp02 = CDbl(strTmp01)
13760   datRetVal = CDate(dblTmp02)

EXITP:
13770   TodaysDate = datRetVal
13780   Exit Function

ERRH:
13790   datRetVal = Date
13800   Select Case ERR.Number
        Case Else
13810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13820   End Select
13830   Resume EXITP

End Function

Public Function GetPostDate() As Date

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "GetPostDate"

        Dim dbs As DAO.Database
        Dim datRetVal As Date

13910   Set dbs = CurrentDb
13920   Set grstPostingDate = dbs.OpenRecordset("PostingDate", dbOpenDynaset, dbConsistent)
13930   With grstPostingDate
13940     .FindFirst "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
13950     Select Case .NoMatch
          Case True
            ' ** There should always be one by this time.
13960       .AddNew
13970       ![Posting_Date] = Date
13980       ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
13990       .Update
14000       datRetVal = Date  'datPostingDate = Date
14010     Case False
14020       Select Case IsNull(![Posting_Date])
            Case True
14030         datRetVal = Date  'datPostingDate = Date
14040         .Edit
14050         ![Posting_Date] = datRetVal  'datPostingDate
14060         .Update
14070       Case False
14080         datRetVal = ![Posting_Date]  'datPostingDate = ![Posting_Date]
14090       End Select
14100     End Select
14110     .Close
14120   End With
14130   Set grstPostingDate = Nothing
14140   dbs.Close

EXITP:
14150   Set dbs = Nothing
14160   GetPostDate = datRetVal
14170   Exit Function

ERRH:
14180   datRetVal = Date
14190   Select Case ERR.Number
        Case Else
14200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14210   End Select
14220   Resume EXITP

End Function
