Attribute VB_Name = "modCourtReportsNY1"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCourtReportsNY1"

'VGC 09/29/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

'QRY: 'qryRpt_CourtReports_NY_Input_IncomeBalance_01' IncomeBalance_img
'QRY: 'qryRpt_CourtReports_NY_Input_InvestedIncome_01' IncomeBalance'));
'DONE!
'QRY: 'qryRpt_CourtReports_NY_Input_IncomeBalance_01' IncomeBalance_hline03_img
'DONE!

'2-Line column headers:
'rptCourtRptNY_07
'3-Line column headers:
'rptCourtRptNY_02
'3-Line column headers and 2-Line report name headers:
'rptCourtRptNY_03
'2-Line report name headers:
'rptCourtRptNY_06
'Messy:
'rptCourtRptNY_01

' ** Array: arr_varCap().
'Private Const C_RID   As Integer = 0
Private Const C_RNAM  As Integer = 1
'Private Const C_CAP   As Integer = 2
Private Const C_CAPN  As Integer = 3

' ** Array: arr_varFile().
Private lngFiles As Long, arr_varFile() As Variant
Private Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
Private Const F_RNAM As Integer = 0
Private Const F_FILE As Integer = 1
Private Const F_PATH As Integer = 2

Private blnExcel As Boolean, blnAllCancel As Boolean, blnNoData As Boolean
Private strThisProc As String ', strCaseNum As String
' **

Public Function cmdDateStart() As Variant

100   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDateStart"

        Dim datRetVal As Date

110     datRetVal = gdatCrtRpt_NY_DateStart

EXITP:
120     cmdDateStart = datRetVal
130     Exit Function

ERRH:
140     Select Case ERR.Number
        Case Else
150       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
160     End Select
170     Resume EXITP

End Function

Public Function cmdDateEnd() As Variant

200   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDateEnd"

        Dim datRetVal As Date

210     datRetVal = gdatCrtRpt_NY_DateEnd

EXITP:
220     cmdDateEnd = datRetVal
230     Exit Function

ERRH:
240     Select Case ERR.Number
        Case Else
250       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
260     End Select
270     Resume EXITP

End Function

Public Function cmdAccountno() As String

300   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAccountno"

        Dim strRetVal As String

310     strRetVal = gstrCrtRpt_NY_AccountNo

EXITP:
320     cmdAccountno = strRetVal
330     Exit Function

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Function

Public Function cmdgstrCrtRpt_CashAssets_Beg() As String

400   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdgstrCrtRpt_CashAssets_Beg"

        Dim strRetVal As String

410     strRetVal = gstrCrtRpt_CashAssets_Beg

EXITP:
420     cmdgstrCrtRpt_CashAssets_Beg = strRetVal
430     Exit Function

ERRH:
440     Select Case ERR.Number
        Case Else
450       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
460     End Select
470     Resume EXITP

End Function

Public Function cmdIncomeAtBegin() As Currency

500   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdIncomeAtBegin"

        Dim curRetVal As Currency

510     curRetVal = gcurCrtRpt_NY_IncomeBeg

EXITP:
520     cmdIncomeAtBegin = curRetVal
530     Exit Function

ERRH:
540     Select Case ERR.Number
        Case Else
550       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
560     End Select
570     Resume EXITP

End Function

Public Function cmdNewInput() As Currency

600   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNewInput"

        Dim curRetVal As Currency

610     curRetVal = gcurCrtRpt_NY_InputNew

EXITP:
620     cmdNewInput = curRetVal
630     Exit Function

ERRH:
640     Select Case ERR.Number
        Case Else
650       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
660     End Select
670     Resume EXITP

End Function

Public Function cmdIncomeCash() As Currency

700   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdIncomeCash"

        Dim curRetVal As Currency

710     gcurCrtRpt_NY_ICash = Nz(DLookup("[icash]", "qryCourtReport_NY_00_B_01"), 0)
720     curRetVal = gcurCrtRpt_NY_ICash

EXITP:
730     cmdIncomeCash = curRetVal
740     Exit Function

ERRH:
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Function

Public Function cmdInvestedIncome() As Currency

800   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdInvestedIncome"

        Dim curRetVal As Currency

810     curRetVal = Nz(DLookup("tcost", "qryCourtReport_NY_InvestedIncome_b"), 0)

EXITP:
820     cmdInvestedIncome = curRetVal
830     Exit Function

ERRH:
840     Select Case ERR.Number
        Case Else
850       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
860     End Select
870     Resume EXITP

End Function

Public Function NYBuildCourtReportData(strReportNumber As String) As Integer
' ** Return Codes:
' **   0  Success.
' **  -1  Canceled.
' **  -9  Error.

900   On Error GoTo ERRH

        Const THIS_PROC As String = "NYBuildCourtReportData"

        Dim intRetVal As Integer

910     intRetVal = 0

EXITP:
920     NYBuildCourtReportData = intRetVal
930     Exit Function

ERRH:
940     intRetVal = -9
950     Select Case ERR.Number
        Case Else
960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
970     End Select
980     Resume EXITP

End Function

Public Function BuildSummary_NY() As Boolean

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "BuildSummary_NY"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnPennyAdded As Boolean, lngPennyAddedID As Long
        Dim lngRecs As Long
        Dim lngX As Long
        Dim blnRetVal As Boolean

1010    blnRetVal = True

1020    Set dbs = CurrentDb
1030    With dbs

          ' ** Empty tmpCourtReportData.
1040      Set qdf = .QueryDefs("qryCourtReport_02")
1050      qdf.Execute
1060      Set qdf = Nothing
1070      DoEvents

1080      Set rst = dbs.OpenRecordset("tblCourtReports_NY_Def", dbOpenDynaset, dbReadOnly)
1090      With rst
1100        If .BOF = True And .EOF = True Then
1110          blnRetVal = False
1120        Else
1130          .MoveLast
1140          lngRecs = .RecordCount
1150          .MoveFirst
1160          For lngX = 1& To lngRecs
1170            If IsNull(.Fields("QueryNameSummary")) = False And .Fields("Schedule") <> "F" Then
1180              Set qdf = dbs.QueryDefs(.Fields("QueryNameSummary"))
1190              qdf.Execute
1200              Set qdf = Nothing
1210              DoEvents
1220            End If
1230            If lngX < lngRecs Then .MoveNext
1240          Next
1250        End If
1260        .Close
1270      End With
1280      Set rst = Nothing
1290      DoEvents

          ' ** There can be up to 10 records in tmpCourtReportData.
          'varTmp00 = DCount("*", "tmpCourtReportData")

1300      If blnRetVal = True Then
            ' ** AT THIS POINT, 1 SET OF RECS IS IN tmpCourtReportData, AND 'A' IS ZERO. (When no previous balance.)

1310        blnPennyAdded = False: lngPennyAddedID = 0&

            ' ** Update tmpCourtReportData, for Amount.
1320        Set qdf = .QueryDefs("qryCourtReport_NY_07_A_08_02")
            ' ** nz(FormRef('gcurCrtRpt_NY_IncomeBeg'),0)+nz(FormRef('NewInput'),0)  AA-1
1330        qdf.Execute
1340        Set qdf = Nothing
1350        DoEvents

            ' ** IF 'A' IS ZERO, GIVE IT A PENNY!
            ' ** tmpCourtReportData, just Schedule A.
1360        Set qdf = .QueryDefs("qryCourtReport_NY_AdjustA_0")
1370        Set rst = qdf.OpenRecordset
1380        With rst
1390          .MoveLast
1400          .MoveFirst
1410          If ![amount] = 0 Then
1420            lngPennyAddedID = ![crtrpt_id]  ' ** NEW FIELD! I HOPE IT DOESN'T WREAK HAVOC WITH ALL THE OTHERS!
1430            .Edit
1440            ![amount] = 0.01
1450            .Update
1460            blnPennyAdded = True
1470          End If
1480          .Close
1490        End With
1500        Set rst = Nothing
1510        Set qdf = Nothing
1520        DoEvents

            ' ** Append tblCourtReportData_Dummies_NY to tmpCourtReportData.
1530        Set qdf = .QueryDefs("qryCourtReport_NY_AppendDummy")
1540        qdf.Execute
1550        Set qdf = Nothing
1560        DoEvents

            ' ** AT THIS POINT, 2 SETS OF RECS ARE IN tmpCourtReportData, AND BOTH 'A' ARE ZERO.

            ' ** PRIN AT BEG: 0      nz([Amount],0)  DLookup("[Amount]", "tmpCourtReportData", "[ReportSchedule] = 'A'")
            ' ** CASH AT BEG: 0      nz(FormRef('gstrCrtRpt_CashAssets_Beg'),0)
            ' ** PRIN REC:    16232  nz(DSum("totamount","qryCourtReport_NY_Received"),0)
            ' ** INC CUR:     0      nz(FormRef('gcurCrtRpt_NY_IncomeBeg'),0)
            'Debug.Print "'Amount                    = " & DLookup("[Amount]", "tmpCourtReportData", "[ReportSchedule] = 'A'")
            'Debug.Print "'gstrCrtRpt_CashAssets_Beg = " & FormRef("gstrCrtRpt_CashAssets_Beg")
            'Debug.Print "'TotAmount                 = " & DSum("totamount", "qryCourtReport_NY_Received")
            'Debug.Print "'gcurCrtRpt_NY_IncomeBeg       = " & FormRef("gcurCrtRpt_NY_IncomeBeg")

            ' ** Update tmpCourtReportData, subtract Invested Income Schedule A.
1570        Set qdf = .QueryDefs("qryCourtReport_NY_AdjustA")
            ' ** nz([Amount],0) -
            ' ** nz(FormRef('gstrCrtRpt_CashAssets_Beg'),0) +
            ' ** nz(DSum("totamount","qryCourtReport_NY_Received"),0) -
            ' ** nz(FormRef('gcurCrtRpt_NY_IncomeBeg'),0)
            ' ** WHERE [Amount] <> 0
1580        qdf.Execute
1590        Set qdf = Nothing
1600        DoEvents

            ' ** SINCE IT'S EXPECTED THAT THE FIRST 'A' WILL HAVE AN ON-HAND BALANCE,
            ' ** THE '<>0' IS SUPPOSED TO ONLY UPDATE THE 1ST 'A'!
            ' ** REMOVING '<>0' UPDATES BOTH, AND WITH IT, BOTH REMAIN ZERO!

            ' ** AT THIS POINT, 2 SETS OF RECS ARE IN tmpCourtReportData, AND BOTH 'A' ARE 1632!

            ' ** IF FIRST WAS GIVEN A PENNY, REMOVE IT.
1610        If blnPennyAdded = True Then
              ' ** Update tmpCourtReportData, subtract 0.01 when Schedule A was Zero.
1620          Set qdf = .QueryDefs("qryCourtReport_NY_AdjustA_0")
1630          Set rst = qdf.OpenRecordset
1640          With rst
1650            .MoveLast
1660            .MoveFirst
1670            .FindFirst "[crtrpt_id] = " & CStr(lngPennyAddedID)
1680            If .NoMatch = False Then
1690              .Edit
1700              ![amount] = (![amount] - 0.01)
1710              .Update
1720            End If
1730            .Close
1740          End With
1750          Set qdf = Nothing
1760          Set rst = Nothing
1770          DoEvents
1780        End If

1790      End If  ' ** blnRetVal.

1800      .Close
1810    End With

EXITP:
1820    Set rst = Nothing
1830    Set qdf = Nothing
1840    Set dbs = Nothing
1850    BuildSummary_NY = blnRetVal
1860    Exit Function

ERRH:
1870    blnRetVal = False
1880    Select Case ERR.Number
        Case Else
1890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1900    End Select
1910    Resume EXITP

End Function

Public Function Proper(varInput As Variant) As Variant

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Proper"

        Dim strTmp01 As String, strTmp02 As String, strTmp03 As String
        Dim intX As Integer
        Dim varRetVal As Variant

2010    varRetVal = Null

2020    If IsNull(varInput) = False Then
2030      strTmp01 = CStr(LCase(varInput))
2040      strTmp03 = " "
2050      For intX = 1 To Len(strTmp01)
2060        strTmp02 = Mid(strTmp01, intX, 1)
2070        If strTmp02 >= "a" And strTmp02 <= "z" And (strTmp03 < "a" Or strTmp03 > "z") Then
2080          Mid(strTmp01, intX, 1) = UCase$(strTmp02)
2090        End If
2100        strTmp03 = strTmp02
2110      Next
2120      varRetVal = strTmp01
2130    End If

EXITP:
2140    Proper = varRetVal
2150    Exit Function

ERRH:
2160    varRetVal = RET_ERR
2170    Select Case ERR.Number
        Case Else
2180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2190    End Select
2200    Resume EXITP

End Function

Public Function FirstDate_NY(frm As Access.Form) As Boolean
' ** See if dates entered are earlier than first transaction.
' ** If dates not yet entered, let it go through.

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FirstDate_NY"

        Dim datFirstDate As Date
        Dim strFirstDateMsg As String
        Dim blnRetVal As Boolean

2310    blnRetVal = True

2320    With frm

2330      strFirstDateMsg = "There is no data for these reports."

2340      .FirstDateMsg_Set strFirstDateMsg  ' ** Form Procedure: frmRpt_CourtReports_NY.

2350      Select Case IsNull(.cmbAccounts.Column(8))
          Case True
2360        blnRetVal = False
2370        datFirstDate = DateAdd("y", 1, Date)  ' ** Tomorrow.
2380      Case False
2390        datFirstDate = CDate(.cmbAccounts.Column(8))
2400      End Select

2410      .FirstDate_Set datFirstDate  ' ** Form Procedure: frmRpt_CourtReports_NY.

2420      If blnRetVal = True Then
2430        If IsNull(.DateStart) = False Then
2440          If CDate(.DateStart) < datFirstDate Then  ' ** Starting date is early.
2450            If IsNull(.DateEnd) = False Then
2460              If CDate(.DateEnd) < datFirstDate Then  ' ** Ending date is too early.
2470                blnRetVal = False
2480              End If
2490            End If
2500          End If
2510        End If
2520      End If  ' ** blnRetVal.

2530    End With

EXITP:
2540    FirstDate_NY = blnRetVal
2550    Exit Function

ERRH:
2560    blnRetVal = False
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

End Function

Public Sub SetUserReportPath_NY(frm As Access.Form)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "SetUserReportPath_NY"

        Dim blnEnable As Boolean

2710    With frm
2720      blnEnable = True
2730      Select Case IsNull(.UserReportPath)
          Case True
2740        blnEnable = False
2750      Case False
2760        If Trim(.UserReportPath) = vbNullString Then
2770          blnEnable = False
2780        End If
2790      End Select
2800      Select Case blnEnable
          Case True
2810        .UserReportPath.BorderColor = CLR_LTBLU2
2820        .UserReportPath.BackStyle = acBackStyleNormal
2830        .UserReportPath.Enabled = True  ' ** It remains locked.
2840        .UserReportPath_chk.Enabled = True
2850        .UserReportPath_chk.Locked = False
2860        .UserReportPath_chk_lbl1.Visible = True
2870        .UserReportPath_chk_lbl1_dim.Visible = False
2880        .UserReportPath_chk_lbl1_dim_hi.Visible = False
2890        .UserReportPath_chk_lbl2.Visible = True
2900        .UserReportPath_chk_lbl2_dim.Visible = False
2910        .UserReportPath_chk_lbl2_dim_hi.Visible = False
2920      Case False
2930        .UserReportPath = vbNullString
2940        .UserReportPath.BorderColor = WIN_CLR_DISR
2950        .UserReportPath.BackStyle = acBackStyleTransparent
2960        .UserReportPath.Enabled = False
2970        .UserReportPath_chk.Enabled = False
2980        .UserReportPath_chk.Locked = False
2990        .UserReportPath_chk_lbl1.Visible = False
3000        .UserReportPath_chk_lbl1_dim.Visible = True
3010        .UserReportPath_chk_lbl1_dim_hi.Visible = True
3020        .UserReportPath_chk_lbl2.Visible = False
3030        .UserReportPath_chk_lbl2_dim.Visible = True
3040        .UserReportPath_chk_lbl2_dim_hi.Visible = True
3050      End Select
3060    End With

EXITP:
3070    Exit Sub

ERRH:
3080    DoCmd.Hourglass False
3090    Select Case ERR.Number
        Case Else
3100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3110    End Select
3120    Resume EXITP

End Sub

Public Sub ShowRptNums(blnShow As Boolean, frm As Access.Form)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "ShowRptNums"

        Dim lngX As Long

3210    With frm
3220      If CurrentUser = "Superuser" Then  ' ** Internal Access Function: Trust Accountant login.
3230        For lngX = 0& To 12&
3240          .Controls("cmdPrint" & Right("00" & CStr(lngX), 2) & "_lbl").Visible = blnShow
3250          If lngX = 0& Then
3260            .Controls("cmdPrint" & Right("00" & CStr(lngX), 2) & "B_lbl").Visible = blnShow
3270          End If
3280        Next
3290      End If
3300    End With

EXITP:
3310    Exit Sub

ERRH:
3320    Select Case ERR.Number
        Case Else
3330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3340    End Select
3350    Resume EXITP

End Sub

Public Sub SendToFile_NY(frm As Access.Form, strReportNumber As String, blnRebuildTable As Boolean, strControlName As String, lngCaps As Long, arr_varCap As Variant, THAT_NAME As String, Optional varExcel As Variant)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "SendToFile_NY"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strQry As String, strMacro As String
        Dim strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim lngRecs As Long
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal_BuildCourtReportData As Integer
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long

3410    blnContinue = True
3420    blnUseSavedPath = False

3430    With frm

3440      DoCmd.Hourglass True
3450      DoEvents

3460      If IsMissing(varExcel) = True Then
3470        blnExcel = False
3480      Else
3490        blnExcel = varExcel
3500      End If

3510      If blnExcel = True Then
3520        If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
              ' ** It seems like it's not quite closed when it gets here,
              ' ** because if I stop the code and run the function again,
              ' ** it always comes up False.
3530          ForcePause 2  ' ** Module Function: modCodeUtilities.
3540          If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
3550            DoCmd.Hourglass False
3560            msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                  "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                  "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                  "You may close Excel before proceding, then click Retry." & vbCrLf & _
                  "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
                ' ** ... Otherwise Trust Accountant will do it for you.
3570            If msgResponse <> vbRetry Then
3580              blnAllCancel = True
3590              .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
3600              AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
3610              blnContinue = False
3620            End If
3630          End If
3640        End If
3650      End If  ' ** blnExcel.

3660      If blnContinue = True Then

3670        DoCmd.Hourglass True
3680        DoEvents

3690        intRetVal_BuildCourtReportData = 0

3700        If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

3710          ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

3720          gstrFormQuerySpec = THAT_NAME

              ' ** Set global variables for report headers.
3730          gstrAccountNo = .cmbAccounts.Column(0)
3740          gstrAccountName = .cmbAccounts.Column(3)
3750          gdatStartDate = .DateStart.Value
3760          gdatEndDate = .DateEnd.Value
3770          gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal and gstrCrtRpt_Version should be populated from the input window.

3780          Set dbs = CurrentDb
3790          With dbs
                ' ** tblReport, captions of Court Reports, by specified [CrtTyp].
3800            Set qdf = .QueryDefs("qryCourtReport_15")
3810            With qdf.Parameters
3820              ![CrtTyp] = "NY"
3830            End With
3840            Set rst = qdf.OpenRecordset
3850            With rst
3860              .MoveLast
3870              lngCaps = .RecordCount
3880              .MoveFirst
3890              arr_varCap = .GetRows(lngCaps)
                  ' ****************************************************
                  ' ** Array: arr_varCap()
                  ' **
                  ' **   Field  Element  Name               Constant
                  ' **   =====  =======  =================  ==========
                  ' **     1       0     rpt_id             C_RID
                  ' **     2       1     rpt_name           C_RNAM
                  ' **     3       2     rpt_caption        C_CAP
                  ' **     4       3     rpt_caption_new    C_CAPN
                  ' **
                  ' ****************************************************
3900              .Close
3910            End With
3920            .Close
3930          End With

              ' ** Build a new summary report table.
3940          intRetVal_BuildCourtReportData = NYBuildCourtReportData(strReportNumber)
3950          If intRetVal_BuildCourtReportData = 0 Then
3960            blnRebuildTable = False
3970            If blnExcel = True Then

3980              Set dbs = CurrentDb

                  ' ** Empty tmpCourtReportData2.
3990              Set qdf = dbs.QueryDefs("qryCourtReport_NY_01_00")
4000              qdf.Execute

4010              Select Case strReportNumber

                  Case "0", "0A"
                    ' ** 0 Summary of Account.

                    ' ** tmpCourtReportData
                    ' **   "date >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And date < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#"
                    ' **     " and accountno = '" & gstrAccountNo & "'"

                    ' ** Empty tmpCourtReportData7.
4020                Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_09")
4030                qdf.Execute

                    ' ** Append qryCourtReport_NS_00_08 to tmpCourtReportData7.
4040                Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_10")
4050                With qdf.Parameters
4060                  ![actno] = gstrAccountNo
4070                  ![datbeg] = gdatStartDate
4080                  ![datEnd] = gdatEndDate
4090                End With
4100                qdf.Execute

4110                Set rst = dbs.OpenRecordset("tmpCourtReportData7", dbOpenDynaset, dbConsistent)
4120                With rst
4130                  .MoveLast
4140                  lngRecs = .RecordCount
4150                  .MoveFirst
                      ' ** Set the initial sorting order.
4160                  For lngX = 1& To lngRecs
4170                    .Edit
4180                    ![sort] = CSng(lngX)
4190                    .Update
4200                    If lngX < lngRecs Then .MoveNext
4210                  Next
4220                  .Close
4230                End With

4240                Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_15")
4250                qdf.Execute

4260                Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_17")
4270                qdf.Execute

4280                Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_20")
4290                qdf.Execute

                    ' ** line 6  is Rpt 40
                    ' ** line 10 is Rpt 60

4300                Set rst = dbs.OpenRecordset("tmpCourtReportData7", dbOpenDynaset, dbConsistent)
4310                With rst
                      ' ** Now add blank lines.
4320                  .AddNew
4330                  ![sort] = 4.5!
4340                  ![Report Division] = 20!
4350                  ![Report Group] = 10.3!
4360                  ![Report Number] = 20.4!
4370                  .Update
4380                  .AddNew
4390                  ![sort] = 7.5!
4400                  ![Report Division] = 20!
4410                  ![Report Group] = 20.3!
4420                  ![Report Number] = 40.4!
4430                  .Update
4440                  .AddNew
4450                  ![sort] = 8.5!
4460                  ![Report Division] = 20.3!
4470                  ![Report Group] = 20.5!
4480                  ![Report Number] = 40.6!
4490                  .Update
4500                  .AddNew
4510                  ![sort] = 10.5!
4520                  ![Report Division] = 40!
4530                  ![Report Group] = 30!
4540                  ![Report Number] = 60.1!
4550                  .Update
4560                  .AddNew
4570                  ![sort] = 13.5!
4580                  ![Report Division] = 60!
4590                  ![Report Group] = 40.3!
4600                  ![Report Number] = 70.4!
4610                  .Update
4620                  .AddNew
4630                  ![sort] = 16.5!
4640                  ![Report Division] = 60!
4650                  ![Report Group] = 50.3!
4660                  ![Report Number] = 90.4!
4670                  .Update
4680                  .AddNew
4690                  ![sort] = 17.5!
4700                  ![Report Division] = 60.3!
4710                  ![Report Group] = 50.5!
4720                  ![Report Number] = 90.6!
4730                  .Update
4740                  .Close
4750                End With

4760                blnNoData = False
4770                Select Case strReportNumber
                    Case "0"
                      ' ** Summary: For export.
4780                  strQry = "qryCourtReport_NY_00_11"

4790                Case "0A"
                      'Disbursements of Principal, by revcode_DESC
                      'Div 20, Grp 20, Rpt 30
                      'Disbursements of Income, by revcode_DESC
                      'Div 60, Grp 50, Rpt 80
                      ' ** Append qryCourtReport_NS_00_A_04 (Principal/Income detail entries, grouped and summed) to tmpCourtReportData7.
4800                  Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_A_05")
4810                  With qdf.Parameters
4820                    ![actno] = gstrAccountNo
4830                    ![datbeg] = gdatStartDate
4840                    ![datEnd] = gdatEndDate
4850                  End With
4860                  qdf.Execute
                      ' ** Append Disbursements of Principal/Income total lines to tmpCourtReportData7.
4870                  Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_A_06")
4880                  qdf.Execute
                      ' ** Update tmpCourtReportData7, Disbursements of Principal/Income header line, [Acquisition Value] = Null.
4890                  Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_A_07")
4900                  qdf.Execute
                      ' ** Summary: For export; grouped by Inc/Exp Codes.
4910                  strQry = "qryCourtReport_NY_00_A_08"
4920                End Select

4930              Case "1"
                    ' ** 1 Receipts of Principal.

                    ' ** qryCourtReport-A2
                    ' **   "transdate >= #" & Format(gdatStartDate,"mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#"
                    ' **     " and accountno = '" & gstrAccountNo & "'"
                    ' **     " and ((journaltype = 'Received' and pcash > 0) "
                    ' **     " or (journaltype = 'Misc.' and pcash > 0) "
                    ' **     " or (journaltype = 'Cost Adj.' and cost > 0) "
                    ' **     " or (journaltype = 'Deposit' and not(jcomment like '*stock split*')))"

                    ' ** Append qryCourtReport_NY_01_03 to tmpCourtReportData2.
4940                Set qdf = dbs.QueryDefs("qryCourtReport_NY_01_04")
4950                With qdf.Parameters
4960                  ![actno] = gstrAccountNo
4970                  ![datbeg] = gdatStartDate
4980                  ![datEnd] = gdatEndDate
4990                End With
5000                qdf.Execute
5010                dbs.Close

                    ' ** Receipts of Principal: For export.
5020                strQry = "qryCourtReport_NY_01_12"
                    ' ** Union of qryCourtReport_NY_07 (qryCourtReport_NY_05 (tmpCourtReportData2,
                    ' ** just needed fields), with qryCourtReport_NY_06 (qryCourtReport_NY_05
                    ' ** (tmpCourtReportData2, just needed fields), Top 1 uniqueid) uniqueidx;
                    ' ** Cartesian), qryCourtReport_NY_08 (qryCourtReport_NY_05 (tmpCourtReportData2,
                    ' ** just needed fields), grouped and summed by accountno).
5030                varTmp00 = DCount("*", "qryCourtReport_NY_01_11a")
5040                If IsNull(varTmp00) = True Then
5050                  blnNoData = True
5060                  strQry = "qryCourtReport_NY_01_15"
5070                Else
5080                  If varTmp00 = 0 Then
5090                    blnNoData = True
5100                    strQry = "qryCourtReport_NY_01_15"
5110                  End If
5120                End If

5130              Case "2"
                    ' ** 2 Gains (Losses) on Sale or Other Dispositions.

                    ' ** qryCourtReport-B
                    ' **   "transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#"
                    ' **     " and accountno = '" & gstrAccountNo & "'"
                    ' **     " and (journaltype = 'Sold' and GainLoss <> 0)"

                    ' ** Append qryCourtReport_NS_02_02 to tmpCourtReportData2.
5140                Set qdf = dbs.QueryDefs("qryCourtReport_NY_02_03")
5150                With qdf.Parameters
5160                  ![actno] = gstrAccountNo
5170                  ![datbeg] = gdatStartDate
5180                  ![datEnd] = gdatEndDate
5190                End With
5200                qdf.Execute
5210                dbs.Close

                    ' ** Gains/Losses on Sale: For export.
5220                strQry = "qryCourtReport_NY_02_11"
5230                varTmp00 = DCount("*", "qryCourtReport_NY_02_10a")
5240                If IsNull(varTmp00) = True Then
5250                  blnNoData = True
5260                  strQry = "qryCourtReport_NY_02_14"
5270                Else
5280                  If varTmp00 = 0 Then
5290                    blnNoData = True
5300                    strQry = "qryCourtReport_NY_02_14"
5310                  End If
5320                End If

5330              Case "3", "3A"
                    ' ** 3 Disbursements of Principal.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# AND transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#)"
                    ' **     " AND accountno = '" & gstrAccountNo & "'"
                    ' **     " AND ((journaltype = 'Paid' AND pcash <> 0 and taxcode <> 11)"  '<> "Distribution"
                    ' **     " OR (journaltype = 'Misc.' AND pcash < 0 )"
                    ' **     " OR (journaltype = 'Cost Adj.' AND cost < 0) "
                    ' **     " OR (journaltype = 'Withdrawn' AND taxcode <> 11))"  '<> "Distribution"

                    ' ** Append qryCourtReport_NS_03_02 to tmpCourtReportData2.
5340                Set qdf = dbs.QueryDefs("qryCourtReport_NY_03_03")    '####  TAXCODE  ####
5350                With qdf.Parameters
5360                  ![actno] = gstrAccountNo
5370                  ![datbeg] = gdatStartDate
5380                  ![datEnd] = gdatEndDate
5390                End With
5400                qdf.Execute
5410                dbs.Close

5420                Select Case strReportNumber
                    Case "3"
                      ' ** Disbursements of Principal: For export.
5430                  strQry = "qryCourtReport_NY_03_15"
5440                  varTmp00 = DCount("*", "qryCourtReport_NY_03_14a")
5450                  If IsNull(varTmp00) = True Then
5460                    blnNoData = True
5470                    strQry = "qryCourtReport_NY_03_15c"
5480                  Else
5490                    If varTmp00 = 0 Then
5500                      blnNoData = True
5510                      strQry = "qryCourtReport_NY_03_15c"
5520                    End If
5530                  End If

5540                Case "3A"
                      ' ** Disbursements of Principal: For export, grouped by Inc/Exp Codes.
5550                  strQry = "qryCourtReport_NY_03_17"
5560                  varTmp00 = DCount("*", "qryCourtReport_NY_03_16a")
5570                  If IsNull(varTmp00) = True Then
5580                    blnNoData = True
5590                    strQry = "qryCourtReport_NY_03_20"
5600                  Else
5610                    If varTmp00 = 0 Then
5620                      blnNoData = True
5630                      strQry = "qryCourtReport_NY_03_20"
5640                    End If
5650                  End If

5660                End Select

5670              Case "4"
                    ' ** 4 Distributions of Principal to Beneficiaries.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#)"
                    ' **     " and accountno = '" & gstrAccountNo & "'"
                    ' **     " and ((journaltype = 'Paid' and pcash <> 0 and taxcode = 11)"  '= "Distribution"
                    ' **     " or (journaltype = 'Withdrawn' and taxcode = 11))"  '= "Distribution"

                    ' ** Append qryCourtReport_NS_04_02 to tmpCourtReportData2.
5680                Set qdf = dbs.QueryDefs("qryCourtReport_NY_04_03")
5690                With qdf.Parameters
5700                  ![actno] = gstrAccountNo
5710                  ![datbeg] = gdatStartDate
5720                  ![datEnd] = gdatEndDate
5730                End With
5740                qdf.Execute
5750                dbs.Close

                    ' ** Distributions of Principal: For export.
5760                strQry = "qryCourtReport_NY_04_11"
5770                varTmp00 = DCount("*", "qryCourtReport_NY_04_10a")
5780                If IsNull(varTmp00) = True Then
5790                  blnNoData = True
5800                  strQry = "qryCourtReport_NY_04_14"
5810                Else
5820                  If varTmp00 = 0 Then
5830                    blnNoData = True
5840                    strQry = "qryCourtReport_NY_04_14"
5850                  End If
5860                End If

5870              Case "5"
                    ' ** 5 Information For Investments Made.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "(journaltype = 'Purchase')"

                    ' ** Append qryCourtReport_NS_05_02 to tmpCourtReportData2.
5880                Set qdf = dbs.QueryDefs("qryCourtReport_NY_05_03")
5890                With qdf.Parameters
5900                  ![actno] = gstrAccountNo
5910                  ![datbeg] = gdatStartDate
5920                  ![datEnd] = gdatEndDate
5930                End With
5940                qdf.Execute
5950                dbs.Close

                    ' ** Info For Investers: For export.
5960                strQry = "qryCourtReport_NY_05_11"
5970                varTmp00 = DCount("*", "qryCourtReport_NY_05_10a")
5980                If IsNull(varTmp00) = True Then
5990                  blnNoData = True
6000                  strQry = "qryCourtReport_NY_05_14"
6010                Else
6020                  If varTmp00 = 0 Then
6030                    blnNoData = True
6040                    strQry = "qryCourtReport_NY_05_14"
6050                  End If
6060                End If

6070              Case "6"
                    ' ** 6 Change in Investment Holdings.

                    ' ** qryCourtReport-A3
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((journaltype = 'Sold' And GainLoss = 0) Or "
                    ' **     "(journaltype = 'Deposit' And jcomment Like '*stock split*') Or "
                    ' **     "(journaltype = 'Liability'))"

                    ' ** Append qryCourtReport_NS_06_03 to tmpCourtReportData2.
6080                Set qdf = dbs.QueryDefs("qryCourtReport_NY_06_04")
6090                With qdf.Parameters
6100                  ![actno] = gstrAccountNo
6110                  ![datbeg] = gdatStartDate
6120                  ![datEnd] = gdatEndDate
6130                End With
6140                qdf.Execute
6150                dbs.Close

                    ' ** Change in Investments: For export.
6160                strQry = "qryCourtReport_NY_06_12"
6170                varTmp00 = DCount("*", "qryCourtReport_NY_06_11a")
6180                If IsNull(varTmp00) = True Then
6190                  blnNoData = True
6200                  strQry = "qryCourtReport_NY_06_15"
6210                Else
6220                  If varTmp00 = 0 Then
6230                    blnNoData = True
6240                    strQry = "qryCourtReport_NY_06_15"
6250                  End If
6260                End If

6270              Case "7"
                    ' ** 7 Receipts of Income.

                    ' ** qryCourtReport - Receipts of Income 1
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((icash > 0) Or "
                    ' **     "(journaltype = 'Purchase' And icash < 0 And pcash = -cost))"

                    ' ** Append qryCourtReport_NS_07_02 to tmpCourtReportData2.
6280                Set qdf = dbs.QueryDefs("qryCourtReport_NY_07_03")
6290                With qdf.Parameters
6300                  ![actno] = gstrAccountNo
6310                  ![datbeg] = gdatStartDate
6320                  ![datEnd] = gdatEndDate
6330                End With
6340                qdf.Execute
6350                dbs.Close

                    ' ** Receipts of Income: For export.
6360                strQry = "qryCourtReport_NY_07_16"
6370                varTmp00 = DCount("*", "qryCourtReport_NY_07_15a")
6380                If IsNull(varTmp00) = True Then
6390                  blnNoData = True
6400                  strQry = "qryCourtReport_NY_07_19"
6410                Else
6420                  If varTmp00 = 0 Then
6430                    blnNoData = True
6440                    strQry = "qryCourtReport_NY_07_19"
6450                  End If
6460                End If

6470              Case "8", "8A"
                    ' ** 8 Disbursements of Income.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((journaltype = 'Paid' And icash <> 0 And taxcode <> 11) Or "  '<> "Distribution"
                    ' **     "(journaltype = 'Misc.' And icash < 0) Or "
                    ' **     "(journaltype = 'Liability' And icash < 0))"

                    ' ** Append qryCourtReport_NS_08_02 to tmpCourtReportData2.
6480                Set qdf = dbs.QueryDefs("qryCourtReport_NY_08_03")    '####  TAXCODE  ####
6490                With qdf.Parameters
6500                  ![actno] = gstrAccountNo
6510                  ![datbeg] = gdatStartDate
6520                  ![datEnd] = gdatEndDate
6530                End With
6540                qdf.Execute
6550                dbs.Close

6560                Select Case strReportNumber
                    Case "8"
                      ' ** Disbursements of Income: For export.
6570                  strQry = "qryCourtReport_NY_08_15"
6580                  varTmp00 = DCount("*", "qryCourtReport_NY_08_14a")
6590                  If IsNull(varTmp00) = True Then
6600                    blnNoData = True
6610                    strQry = "qryCourtReport_NY_08_15c"
6620                  Else
6630                    If varTmp00 = 0 Then
6640                      blnNoData = True
6650                      strQry = "qryCourtReport_NY_08_15c"
6660                    End If
6670                  End If

6680                Case "8A"
                      ' ** Disbursements of Income: For export; grouped by Inc/Exp Codes.
6690                  strQry = "qryCourtReport_NY_08_17"
6700                  varTmp00 = DCount("*", "qryCourtReport_NY_08_16a")
6710                  If IsNull(varTmp00) = True Then
6720                    blnNoData = True
6730                    strQry = "qryCourtReport_NY_08_20"
6740                  Else
6750                    If varTmp00 = 0 Then
6760                      blnNoData = True
6770                      strQry = "qryCourtReport_NY_08_20"
6780                    End If
6790                  End If

6800                End Select

6810              Case "9"
                    ' ** 9 Distributions of Income.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((journaltype = 'Paid' And icash <> 0 And taxcode = 11))"  '= "Distribution"

                    ' ** Append qryCourtReport_NS_09_02 to tmpCourtReportData2.
6820                Set qdf = dbs.QueryDefs("qryCourtReport_NY_09_03")    '####  TAXCODE  ####
6830                With qdf.Parameters
6840                  ![actno] = gstrAccountNo
6850                  ![datbeg] = gdatStartDate
6860                  ![datEnd] = gdatEndDate
6870                End With
6880                qdf.Execute
6890                dbs.Close

                    ' ** Distributions of Income: For export.
6900                strQry = "qryCourtReport_NY_09_11"
6910                varTmp00 = DCount("*", "qryCourtReport_NY_09_10a")
6920                If IsNull(varTmp00) = True Then
6930                  blnNoData = True
6940                  strQry = "qryCourtReport_NY_09_14"
6950                Else
6960                  If varTmp00 = 0 Then
6970                    blnNoData = True
6980                    strQry = "qryCourtReport_NY_09_14"
6990                  End If
7000                End If

7010              End Select

7020              strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
7030              strRptPath = .UserReportPath

7040              If Len(strReportNumber) = 1 Then
7050                strRptName = ("rptCourtRptNY_0" & strReportNumber)
7060              ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
7070                strRptName = ("rptCourtRptNY_" & strReportNumber)
7080              Else
7090                strRptName = ("rptCourtRptNY_0" & strReportNumber)
7100              End If

7110              strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
7120              If blnNoData = True Then
7130                strMacro = strMacro & "_nd"
7140              End If

7150              For lngX = 0& To (lngCaps - 1&)
7160                If arr_varCap(C_RNAM, lngX) = strRptName Then
7170                  strRptCap = arr_varCap(C_CAPN, lngX)
7180                  Exit For
7190                End If
7200              Next

7210              If IsNull(.UserReportPath) = False Then
7220                If .UserReportPath <> vbNullString Then
7230                  If .UserReportPath_chk = True Then
7240                    If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
7250                      blnUseSavedPath = True
7260                    End If
7270                  End If
7280                End If
7290              End If

7300              Select Case blnUseSavedPath
                  Case True
7310                strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
7320              Case False
7330                DoCmd.Hourglass False
7340                strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
7350              End Select

7360              If strRptPathFile <> vbNullString Then
7370                DoCmd.Hourglass True
7380                DoEvents
7390                Select Case blnExcel
                    Case True
7400                  blnAutoStart = .chkOpenExcel
7410                Case False
7420                  blnAutoStart = .chkOpenWord
7430                End Select
7440                If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
7450                If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
7460                  Kill strRptPathFile
7470                End If
                    ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                    ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
7480                DoCmd.RunMacro strMacro
                    ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                    ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
7490                If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                        FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
7500                  If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
7510                    Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
7520                  Else
7530                    Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
7540                  End If
7550                  DoEvents
7560                  If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
7570                    DoEvents
7580                    Select Case gblnPrintAll
                        Case True
7590                      lngFiles = lngFiles + 1&
7600                      lngE = lngFiles - 1&
7610                      ReDim Preserve arr_varFile(F_ELEMS, lngE)
7620                      arr_varFile(F_RNAM, lngE) = strRptName
7630                      arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
7640                      arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
                          'FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY.
7650                    Case False
7660                      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
7670                        EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
7680                      End If
7690                      DoEvents
7700                      If blnAutoStart = True Then
7710                        OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
7720                      End If
7730                    End Select
                        'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                        '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                        'End If
                        'DoEvents
                        'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
7740                  End If
7750                End If
7760                strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7770                If strRptPath <> .UserReportPath Then
7780                  .UserReportPath = strRptPath
7790                  SetUserReportPath_NY frm  ' ** Procedure: Above.
7800                End If
7810              Else
7820                blnContinue = False
7830              End If

7840            Else

7850              strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
7860              strRptPath = .UserReportPath

7870              If Len(strReportNumber) = 1 Then
7880                strRptName = ("rptCourtRptNY_0" & strReportNumber)
7890              ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
7900                strRptName = ("rptCourtRptNY_" & strReportNumber)
7910              Else
7920                strRptName = ("rptCourtRptNY_0" & strReportNumber)
7930              End If

7940              strRptCap = vbNullString
7950              For lngX = 0& To (lngCaps - 1&)
7960                If arr_varCap(C_RNAM, lngX) = strRptName Then
7970                  strRptCap = arr_varCap(C_CAPN, lngX)
7980                  Exit For
7990                End If
8000              Next

8010              If IsNull(.UserReportPath) = False Then
8020                If .UserReportPath <> vbNullString Then
8030                  If .UserReportPath_chk = True Then
8040                    If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
8050                      blnUseSavedPath = True
8060                    End If
8070                  End If
8080                End If
8090              End If

8100              Select Case blnUseSavedPath
                  Case True
8110                strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
8120              Case False
8130                DoCmd.Hourglass False
8140                strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
8150              End Select

8160              If strRptPathFile <> vbNullString Then
8170                DoCmd.Hourglass True
8180                DoEvents
8190                Select Case blnExcel
                    Case True
8200                  blnAutoStart = .chkOpenExcel
8210                Case False
8220                  blnAutoStart = .chkOpenWord
8230                End Select
8240                If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
8250                If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
8260                  Kill strRptPathFile
8270                End If
8280                Select Case gblnPrintAll
                    Case True
8290                  lngFiles = lngFiles + 1&
8300                  lngE = lngFiles - 1&
8310                  ReDim Preserve arr_varFile(F_ELEMS, lngE)
8320                  arr_varFile(F_RNAM, lngE) = strRptName
8330                  arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
8340                  arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
                      'FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY.
8350                  DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
8360                Case False
8370                  DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
8380                End Select
                    'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
8390                strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
8400                If strRptPath <> .UserReportPath Then
8410                  .UserReportPath = strRptPath
8420                  SetUserReportPath_NY frm  ' ** Procedure: Above.
8430                End If
8440              Else
8450                blnContinue = False
8460              End If

8470            End If
8480          Else
                ' ** Return Codes:
                ' **   0  Success.
                ' **  -1  Canceled.
                ' **  -9  Error.
8490            blnContinue = False
8500          End If  ' ** intRetVal_BuildCourtReportData.

8510        End If  ' ** Validate.
8520      End If ' ** blnContinue.
8530    End With

8540    DoCmd.Hourglass False

EXITP:
8550    Set rst = Nothing
8560    Set qdf = Nothing
8570    Set dbs = Nothing
8580    Exit Sub

ERRH:
8590    DoCmd.Hourglass False
8600    Select Case ERR.Number
        Case 70  ' ** Permission denied.
8610      Beep
8620      MsgBox "Trust Accountant is unable to save the file." & vbCrLf & vbCrLf & _
            "If the program to which you're exporting is open," & vbCrLf & _
            "please close it and try again.", vbInformation + vbOKOnly, "Failed To Save File"
8630    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
8640    Case Else
8650      zErrorHandler frm.Name, strControlName, ERR.Number, Erl
8660    End Select
8670    Resume EXITP

End Sub

Public Sub AssetList_Word_NY(THAT_NAME As String, frm As Access.Form)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "AssetList_Word_NY"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCaps As Long, arr_varCap As Variant
        Dim strRptType As String, strRptName As String, strRptCap As String, strThisProc As String
        Dim strRptPath As String, strRptPathFile As String
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim lngX As Long, lngE As Long

8710    blnContinue = True
8720    blnUseSavedPath = False

8730    With frm

8740      DoCmd.Hourglass True
8750      DoEvents

8760      blnExcel = False
8770      blnAllCancel = False
8780      .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
8790      AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
8800      blnAutoStart = .chkOpenWord
8810      strThisProc = "cmdWord00B_Click"

8820      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

8830        gstrFormQuerySpec = THAT_NAME

8840        strRptType = vbNullString
8850        intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRptType, strThisProc, frm)  ' ** Function: Below.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry, e.g., date.
            ' **   -9  Error.

8860        Select Case intRetVal_BuildAssetListInfo
            Case 0

8870          If intRetVal_BuildAssetListInfo = -2 Then
8880            Set dbs = CurrentDb
8890            With dbs
                  ' ** Empty tmpAssetList2.
8900              Set qdf = dbs.QueryDefs("qryCourtReport_03")
8910              qdf.Execute
8920              .Close
8930            End With
8940          ElseIf intRetVal_BuildAssetListInfo < 0 Then
8950            blnContinue = False
8960          End If

8970          If blnContinue = True Then

                ' ** strRptType should return either "_00B" or "_00BA".
8980            If strRptType <> vbNullString Then

8990              gdatStartDate = .DateStart
9000              gdatEndDate = .DateEnd
9010              gstrAccountNo = .cmbAccounts.Column(0)
9020              gstrAccountName = .cmbAccounts.Column(3)

9030              lngCaps = 0&
9040              arr_varCap = Empty

9050              Set dbs = CurrentDb
9060              With dbs
                    ' ** tblReport, captions of Court Reports, by specified [CrtTyp].
9070                Set qdf = .QueryDefs("qryCourtReport_15")
9080                With qdf.Parameters
9090                  ![CrtTyp] = "NY"
9100                End With
9110                Set rst = qdf.OpenRecordset
9120                With rst
9130                  .MoveLast
9140                  lngCaps = .RecordCount
9150                  .MoveFirst
9160                  arr_varCap = .GetRows(lngCaps)
                      ' ****************************************************
                      ' ** Array: arr_varCap()
                      ' **
                      ' **   Field  Element  Name               Constant
                      ' **   =====  =======  =================  ==========
                      ' **     1       0     rpt_id             C_RID
                      ' **     2       1     rpt_name           C_RNAM
                      ' **     3       2     rpt_caption        C_CAP
                      ' **     4       3     rpt_caption_new    C_CAPN
                      ' **
                      ' ****************************************************
9170                  .Close
9180                End With
9190                .Close
9200              End With

9210              strRptCap = vbNullString: strRptPathFile = vbNullString
9220              strRptPath = .UserReportPath
9230              strRptName = "rptCourtRptNY" & strRptType

9240              strRptCap = vbNullString
9250              For lngX = 0& To (lngCaps - 1&)
9260                If arr_varCap(C_RNAM, lngX) = strRptName Then
9270                  strRptCap = arr_varCap(C_CAPN, lngX)
9280                  Exit For
9290                End If
9300              Next

9310              If IsNull(.UserReportPath) = False Then
9320                If .UserReportPath <> vbNullString Then
9330                  If .UserReportPath_chk = True Then
9340                    If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
9350                      blnUseSavedPath = True
9360                    End If
9370                  End If
9380                End If
9390              End If

9400              Select Case blnUseSavedPath
                  Case True
9410                strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
9420              Case False
9430                DoCmd.Hourglass False
9440                strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
9450              End Select

9460              If strRptPathFile <> vbNullString Then
9470                DoCmd.Hourglass True
9480                DoEvents
9490                Select Case blnExcel
                    Case True
9500                  blnAutoStart = .chkOpenExcel
9510                Case False
9520                  blnAutoStart = .chkOpenWord
9530                End Select
9540                If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
9550                If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
9560                  Kill strRptPathFile
9570                End If
9580                Select Case gblnPrintAll
                    Case True
9590                  lngFiles = lngFiles + 1&
9600                  lngE = lngFiles - 1&
9610                  ReDim Preserve arr_varFile(F_ELEMS, lngE)
9620                  arr_varFile(F_RNAM, lngE) = strRptName
9630                  arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
9640                  arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
                      'FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY.
9650                  DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
9660                Case False
9670                  DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
9680                End Select
                    'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
9690                strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
9700                If strRptPath <> .UserReportPath Then
9710                  .UserReportPath = strRptPath
9720                  SetUserReportPath_NY frm  ' ** Procedure: Above.
9730                End If
9740              Else
9750                blnAllCancel = True
9760                .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
9770                AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
9780                blnContinue = False
9790              End If  ' ** strRptPathFile.

9800            End If  ' ** strRptType.

9810          Else
9820            blnAllCancel = True
9830            .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
9840            AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
9850            blnContinue = False
9860            DoCmd.Hourglass False
9870            MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
9880          End If  ' ** blnContinue.

9890        Case -2
9900          Beep
9910          MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
9920        Case -3, -9
              ' ** Message shown below.
9930        End Select  ' ** intRetVal_BuildAssetListInfo

9940      End If  ' ** Validate.

9950      DoCmd.Hourglass False

9960    End With

EXITP:
9970    Set rst = Nothing
9980    Set qdf = Nothing
9990    Set dbs = Nothing
10000   Exit Sub

ERRH:
10010   blnAllCancel = True
10020   frm.AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
10030   AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
10040   gblnPrintAll = False
10050   DoCmd.Hourglass False
10060   Select Case ERR.Number
        Case Else
10070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10080   End Select
10090   Resume EXITP

End Sub

Public Sub AssetList_Excel_NY(frm As Access.Form, THAT_NAME As String, Optional varFromExcel00 As Variant)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetList_Excel_NY"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCaps As Long, arr_varCap As Variant
        Dim strQry As String, strMacro As String
        Dim strRptType As String, strRptName As String, strRptCap As String, strThisProc As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strLastAssetType As String
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean, blnFromExcell00 As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long
        Dim lngX As Long, lngE As Long

10110   With frm

10120     DoCmd.Hourglass True
10130     DoEvents

10140     blnContinue = True
10150     blnUseSavedPath = False
10160     blnExcel = True
10170     strThisProc = "cmdExcel00B"

10180     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
10190       ForcePause 2  ' ** Module Function: modCodeUtilities.
10200       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
10210         DoCmd.Hourglass False
10220         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
10230         If msgResponse <> vbRetry Then
10240           blnAllCancel = True
10250           .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
10260           AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
10270           blnContinue = False
10280         End If
10290       End If
10300     End If

10310     If blnContinue = True Then

10320       DoCmd.Hourglass True
10330       DoEvents

10340       If frm.Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

10350         gstrFormQuerySpec = THAT_NAME

10360         blnNoData = False
10370         strRptType = vbNullString
10380         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRptType, strThisProc, frm)  ' ** Function: Below.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

10390         Select Case intRetVal_BuildAssetListInfo
              Case 0
                ' ** Continue.
10400         Case -2
10410           blnNoData = True
10420           Set dbs = CurrentDb
10430           With dbs
                  ' ** Empty tmpAssetList2.
10440             Set qdf = dbs.QueryDefs("qryCourtReport_03")
10450             qdf.Execute
10460             .Close
10470           End With
10480         Case -3, -9
10490           blnAllCancel = True
10500           .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
10510           AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
10520           blnContinue = False
10530         End Select  ' ** intRetVal_BuildAssetListInfo.

10540         If blnContinue = True Then

                ' ** strRptType should return either "_00B" or "_00BA".
                ' ** '_00B':  Non-current end date.
                ' ** '_00BA': Current end date.
10550           If strRptType <> vbNullString Then

10560             gstrAccountNo = .cmbAccounts.Column(0)
10570             gstrAccountName = .cmbAccounts.Column(3)
10580             gdatStartDate = .DateStart
10590             gdatEndDate = .DateEnd
10600             gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
                  ' ** gstrCrtRpt_Ordinal and gstrCrtRpt_Version should be populated from the input window.

10610             lngCaps = 0&
10620             arr_varCap = Empty

10630             Set dbs = CurrentDb
10640             With dbs
                    ' ** tblReport, captions of Court Reports, by specified [CrtTyp].
10650               Set qdf = .QueryDefs("qryCourtReport_15")
10660               With qdf.Parameters
10670                 ![CrtTyp] = "NY"
10680               End With
10690               Set rst = qdf.OpenRecordset
10700               With rst
10710                 .MoveLast
10720                 lngCaps = .RecordCount
10730                 .MoveFirst
10740                 arr_varCap = .GetRows(lngCaps)
                      ' ****************************************************
                      ' ** Array: arr_varCap()
                      ' **
                      ' **   Field  Element  Name               Constant
                      ' **   =====  =======  =================  ==========
                      ' **     1       0     rpt_id             C_RID
                      ' **     2       1     rpt_name           C_RNAM
                      ' **     3       2     rpt_caption        C_CAP
                      ' **     4       3     rpt_caption_new    C_CAPN
                      ' **
                      ' ****************************************************
10750                 .Close
10760               End With
10770               .Close
10780             End With

10790             If blnNoData = False Then
10800               Set dbs = CurrentDb
10810               With dbs

                      ' ** Empty tmpCourtReportData2.
10820                 Set qdf = .QueryDefs("qryCourtReport_NY_00_B_06")
10830                 qdf.Execute

10840                 Select Case strRptType
                      Case "_00B"
                        ' ** Append qryCourtReport_NS_00_B_07a to tmpCourtReportData2.
10850                   Set qdf = .QueryDefs("qryCourtReport_NY_00_B_08a")
10860                 Case "_00BA"
10870                   Set qdf = .QueryDefs("qryCourtReport_NY_00_B_08b")
10880                 End Select
10890                 qdf.Execute

                      ' ** tmpCourtReportData2, sorted.
10900                 Set qdf = .QueryDefs("qryCourtReport_NY_00_B_10")
10910                 Set rst = qdf.OpenRecordset
10920                 With rst
10930                   .MoveLast
10940                   lngRecs = .RecordCount
10950                   .MoveFirst
10960                   strLastAssetType = vbNullString
10970                   For lngX = 1& To lngRecs
10980                     If ![assettype] <> strLastAssetType Then
10990                       .Edit
11000                       ![sort2] = 1&
11010                       .Update
11020                       strLastAssetType = ![assettype]
11030                     Else
11040                       .Edit
11050                       ![sort2] = 2&
11060                       .Update
11070                     End If
11080                     If lngX < lngRecs Then .MoveNext
11090                   Next
11100                   .Close
11110                 End With

11120                 .Close
11130               End With  ' ** dbs.
11140             End If  ' ** blnNoData.

11150             strQry = "qryCourtReport_NY_00_B_24"
11160             Select Case strRptType
                  Case "_00B"
                    ' ** Non-current end date.
                    ' ** Property on Hand at Ending of Account Period.
11170               strQry = "qryCourtReport_NY_00_B_24a"
11180             Case "_00BA"
                    ' ** Current end date.
                    ' ** Property on Hand at Ending of Account Period.
11190               strQry = "qryCourtReport_NY_00_B_24a"
11200             End Select
11210             If blnNoData = True Then
11220               strQry = "qryCourtReport_NY_00_B_29"
11230             End If
                  ' ** OK, WHICH ONE OF THESE SUBS GETS CALLED?
                  ' ** WHEN CALLED ALONE, BY ITSELF, cmdExcel00B_Click IS USED.
                  ' ** WHEN CALLED BY THE SUMMARY, TO BE INCLUDED, THIS PROCEDURE IS USED.

                  ' ** Property on Hand at Beginning of Account Period.
                  'qryCourtReport_NY_00_B_24b

11240             strRptCap = vbNullString: strRptPathFile = vbNullString
11250             strRptPath = .UserReportPath
11260             strRptName = "rptCourtRptNY" & strRptType

11270             For lngX = 0& To (lngCaps - 1&)
11280               If arr_varCap(C_RNAM, lngX) = strRptName Then
11290                 strRptCap = arr_varCap(C_CAPN, lngX)
11300                 Exit For
11310               End If
11320             Next

11330             strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
11340             If strMacro = "mcrExcelExport_CR_NY_00B" Then
11350               strMacro = "mcrExcelExport_CR_NY_00Bi"
11360             End If
11370             If blnNoData = True Then
                    ' ** mcrExcelExport_CR_NY_00Bi_nd
11380               strMacro = strMacro & "_nd"
11390             End If

11400             If IsNull(.UserReportPath) = False Then
11410               If .UserReportPath <> vbNullString Then
11420                 If .UserReportPath_chk = True Then
11430                   If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
11440                     blnUseSavedPath = True
11450                   End If
11460                 End If
11470               End If
11480             End If

11490             Select Case blnUseSavedPath
                  Case True
11500               strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
11510             Case False
11520               DoCmd.Hourglass False
11530               strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
11540             End Select

11550             If strRptPathFile <> vbNullString Then
11560               DoCmd.Hourglass True
11570               DoEvents
11580               Select Case blnExcel
                    Case True
11590                 blnAutoStart = .chkOpenExcel
11600               Case False
11610                 blnAutoStart = .chkOpenWord
11620               End Select
11630               If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
11640               If gblnPrintAll = False And blnFromExcell00 = True And .chkOpenExcel = True Then
11650                 blnAutoStart = False
                      ' ** Borrowing this from the Journal.
11660                 gstrSaleAccountNumber = strRptPathFile
11670               End If
11680               If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
11690                 Kill strRptPathFile
11700               End If
                    ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                    ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
11710               DoCmd.RunMacro strMacro
                    ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                    ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
11720               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                        FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
11730                 If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
11740                   Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
11750                 Else
11760                   Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
11770                 End If
11780                 DoEvents
11790                 If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
11800                   DoEvents
11810                   Select Case gblnPrintAll
                        Case True
11820                     lngFiles = lngFiles + 1&
11830                     lngE = lngFiles - 1&
11840                     ReDim Preserve arr_varFile(F_ELEMS, lngE)
11850                     arr_varFile(F_RNAM, lngE) = strRptName
11860                     arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
11870                     arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
                          'FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY.
11880                   Case False
11890                     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
11900                       EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
11910                     End If
11920                     DoEvents
11930                     If blnAutoStart = True Then
11940                       OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
11950                     End If
11960                   End Select
                        'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                        '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                        'End If
                        'DoEvents
                        'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
11970                 End If
11980               End If
11990               strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
12000               If strRptPath <> .UserReportPath Then
12010                 .UserReportPath = strRptPath
12020                 SetUserReportPath_NY frm  ' ** Procedure: Above.
12030               End If
12040             Else
12050               blnAllCancel = True
12060               .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
12070               AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
12080               blnContinue = False
12090             End If  ' ** strRptPathFile.

12100           End If  ' ** strRptType.

12110         Else
12120           blnAllCancel = True
12130           .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
12140           AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
12150           blnContinue = False
12160           DoCmd.Hourglass False
12170           MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
12180         End If  ' ** blnContinue.

12190       End If  ' ** Validate.
12200     End If  ' ** blnContinue.
12210   End With

12220   DoCmd.Hourglass False

EXITP:
12230   Set rst = Nothing
12240   Set qdf = Nothing
12250   Set dbs = Nothing
12260   Exit Sub

ERRH:
12270   blnAllCancel = True
12280   frm.AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
12290   AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
12300   DoCmd.Hourglass False
12310   Select Case ERR.Number
        Case 70  ' ** Permission denied.
12320     Beep
12330     MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
12340   Case Else
12350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12360   End Select
12370   Resume EXITP

End Sub

Public Sub BtnsEnable_Set_NY(frm As Access.Form)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "BtnsEnable_Set_NY"

12410   With frm

          ' ** If revenue / expense tracking is checked on the options screen,
          ' ** let user select reports to be grouped by revenue/expense codes.
12420     Select Case gblnRevenueExpenseTracking
          Case True
12430       .chkGroupBy_IncExpCode.Enabled = True
12440       .chkGroupBy_IncExpCode_lbl3.Visible = False     ' ** Off note.
12450       .chkGroupBy_IncExpCode_lbl4_txt.Locked = True   ' ** Asterisk
12460       .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
12470       .cmdPreview00_lbl4_txt.Locked = True
12480       .cmdPreview00_lbl4_cmd.Enabled = True
12490       .cmdPreview04_lbl4_txt.Locked = True
12500       .cmdPreview04_lbl4_cmd.Enabled = True
12510       .cmdPreview09_lbl4_txt.Locked = True
12520       .cmdPreview09_lbl4_cmd.Enabled = True
12530       .cmdPreview10_lbl4_txt.Locked = True
12540       .cmdPreview10_lbl4_cmd.Enabled = True
12550     Case False
12560       .chkGroupBy_IncExpCode.Enabled = False
12570       .chkGroupBy_IncExpCode_lbl3.Visible = True      ' ** Off note.
12580       .chkGroupBy_IncExpCode_lbl4_txt.Locked = False  ' ** Asterisk
12590       .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
12600       .cmdPreview00_lbl4_txt.Locked = False  ' ** So they'll look disabled.
12610       .cmdPreview00_lbl4_cmd.Enabled = False
12620       .cmdPreview04_lbl4_txt.Locked = False
12630       .cmdPreview04_lbl4_cmd.Enabled = False
12640       .cmdPreview09_lbl4_txt.Locked = False
12650       .cmdPreview09_lbl4_cmd.Enabled = False
12660       .cmdPreview10_lbl4_txt.Locked = False
12670       .cmdPreview10_lbl4_cmd.Enabled = False
12680     End Select

12690     If glngTaxCode_Distribution = 0& Then
12700       glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
12710     End If

12720     If gdatStartDate > 0 And gdatEndDate > 0 Then
12730       .DateStart = gdatStartDate
12740       .DateEnd = gdatEndDate
12750     End If

      #If NoExcel Then
12760     .cmdExcel00.Enabled = False
12770     .cmdExcel00B.Enabled = False
12780     .cmdExcel01.Enabled = False
12790     .cmdExcel02.Enabled = False
12800     .cmdExcel03.Enabled = False
12810     .cmdExcel04.Enabled = False
12820     .cmdExcel05.Enabled = False
12830     .cmdExcel06.Enabled = False
12840     .cmdExcel07.Enabled = False
12850     .cmdExcel08.Enabled = False
12860     .cmdExcel09.Enabled = False
12870     .cmdExcel10.Enabled = False
12880     .cmdExcel11.Enabled = False
12890     .cmdExcel12.Enabled = False
12900     .cmdExcel13.Enabled = False
12910     .cmdExcelAll.Enabled = False
12920     .chkOpenExcel.Enabled = False
12930     .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
12940     .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
12950     .cmdExcel00.Enabled = True
12960     .cmdExcel00B.Enabled = True
12970     .cmdExcel01.Enabled = True
12980     .cmdExcel02.Enabled = True
12990     .cmdExcel03.Enabled = True
13000     .cmdExcel04.Enabled = True
13010     .cmdExcel05.Enabled = True
13020     .cmdExcel06.Enabled = True
13030     .cmdExcel07.Enabled = True
13040     .cmdExcel08.Enabled = True
13050     .cmdExcel09.Enabled = True
13060     .cmdExcel10.Enabled = True
13070     .cmdExcel11.Enabled = True
13080     .cmdExcel12.Enabled = True
13090     .cmdExcel13.Enabled = False
13100     .cmdExcelAll.Enabled = True
13110     .chkOpenExcel.Enabled = True
13120     .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
13130     .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

13140   End With

EXITP:
13150   Exit Sub

ERRH:
13160   Select Case ERR.Number
        Case Else
13170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13180   End Select
13190   Resume EXITP

End Sub

Public Function PreviewOrPrint_NY(strReportNumber As String, strControlName As String, intView As Integer, blnRebuildTable As Boolean, frm As Access.Form) As Boolean

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "PreviewOrPrint_NY"

        Dim strTmp01 As String
        Dim intRetVal_BuildCourtReportData As Integer
        Dim blnRetVal As Boolean

13210   blnRetVal = True

13220   With frm
13230     If frm.Validate = True Then  ' ** Function: Below.

13240       ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

            ' ** Set global variables for report headers.
13250       gdatStartDate = .DateStart.Value
13260       gdatEndDate = .DateEnd.Value
13270       gstrAccountNo = .cmbAccounts.Column(0)
13280       gstrAccountName = .cmbAccounts.Column(3)

13290       intRetVal_BuildCourtReportData = NYBuildCourtReportData(strReportNumber)  ' ** Module Function: modCourtReportsNS.
            ' ** Return Codes:
            ' **   0  Success.
            ' **  -1  Canceled.
            ' **  -9  Error.

13300       If intRetVal_BuildCourtReportData = 0 Then
13310         blnRebuildTable = False
13320         Select Case IsNumeric(Right(strReportNumber, 1))
              Case True
13330           strTmp01 = "rptCourtRptNY_" & Right("00" & strReportNumber, 2)
13340         Case False
13350           strTmp01 = "rptCourtRptNY_" & Right("000" & strReportNumber, 3)
13360         End Select
13370         If intView = acViewPreview Then gblnMessage = True Else gblnMessage = False
13380         DoCmd.OpenReport strTmp01, intView
13390         If intView = acViewPreview Then
13400           DoCmd.Maximize
13410           DoCmd.RunCommand acCmdFitToWindow
13420         End If
13430       Else
13440         blnRetVal = False
13450       End If

13460     Else
13470       blnRetVal = False
13480     End If  ' ** Validate.
13490   End With

13500   If intView = acNormal Then
13510     If Reports.Count > 0 Then
13520       DoCmd.Close acReport, Reports(0).Name
13530     End If
13540   End If

EXITP:
13550   PreviewOrPrint_NY = blnRetVal
13560   Exit Function

ERRH:
13570   blnRetVal = False
13580   DoCmd.Hourglass False
13590   Select Case ERR.Number
        Case 2202  ' ** You must set a default printer before you design, print, or preview.
13600     MsgBox "There does not appear to be a default printer defined." & vbCrLf & vbCrLf & _
            "Please define a default printer before trying to print this report.", vbCritical + vbOKOnly, "Printer Not Found"
13610   Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
13620   Case Else
13630     zErrorHandler THIS_NAME, THIS_PROC & ": " & strControlName, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13640   End Select
13650   Resume EXITP

End Function

Public Sub AssetList_PreviewPrint_NY(intMode As Integer, THAT_NAME As String, frm As Access.Form, Optional THAT_PROC As Variant)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetList_PreviewPrint_NY"

        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim strRpt As String, strThisProc As String

13710   With frm

13720     blnContinue = True

13730     DoCmd.Hourglass True
13740     DoEvents

13750     strThisProc = THAT_PROC
          ' ** NY_CmdPrev00_Click -> cmdPreview00_Click
          ' ** NY_CmdPrint00_Click -> cmdPrint00_Click
          ' ** cmdPreview00B_Click
          ' ** cmdPrint00B_Click

13760     If frm.Validate = True Then  ' ** Function: Below.

13770       strRpt = vbNullString
13780       intRetVal_BuildAssetListInfo = 0

13790       intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm)  ' ** Function: Below.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry, e.g., date.
            ' **   -9  Error.

13800       Select Case intRetVal_BuildAssetListInfo
            Case 0

              ' ** strRpt should return either "_00B" or "_00BA".
13810         If strRpt <> vbNullString Then

13820           gdatStartDate = .DateStart
13830           gdatEndDate = .DateEnd
13840           gstrAccountNo = .cmbAccounts
13850           gstrAccountName = .cmbAccounts.Column(3)

13860           If intMode = acViewPreview Then gblnMessage = True Else gblnMessage = False
13870           DoCmd.OpenReport "rptCourtRptNY" & strRpt, intMode
13880           If intMode = acViewPreview Then
13890             DoCmd.Maximize
13900             DoCmd.RunCommand acCmdFitToWindow
13910           End If

13920         Else
13930           blnContinue = False
13940           DoCmd.Hourglass False
13950           MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
13960         End If

13970       Case -2
13980         blnContinue = False
13990         If intMode = acViewPreview Then gblnMessage = True Else gblnMessage = False
14000         strRpt = "rptCourtRptNY_00B"
14010         DoCmd.OpenReport strRpt, intMode
14020         If intMode = acViewPreview Then
14030           DoCmd.Maximize
14040           DoCmd.RunCommand acCmdFitToWindow
14050         End If
14060       Case -3, -9
              ' ** Message shown below.
14070       End Select  ' ** intRetVal_BuildAssetListInfo

14080     End If  ' ** Validate.
14090   End With

14100   DoCmd.Hourglass False

EXITP:
14110   Exit Sub

ERRH:
14120   DoCmd.Hourglass False
14130   Forms(THAT_NAME).Visible = True
14140   Select Case ERR.Number
        Case Else
14150     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14160   End Select
14170   Resume EXITP

End Sub

Public Function BuildAssetListInfo_NY(varStartDate As Variant, varEndDate As Variant, strWhen As String, strRpt As String, strControlName As String, frm As Access.Form) As Integer
' ** Copied from frmRpt_CourtReports_CA.
' ** Return codes:
' **    0  Success.
' **   -2  No data.
' **   -3  Missing entry, e.g., date.
' **   -9  Error.

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "BuildAssetListInfo_NY"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim rstWork As DAO.Recordset
        Dim rstTempAssetList As DAO.Recordset
        Dim rstTmpAccountInfo As DAO.Recordset
        Dim rstAccount As DAO.Recordset
        Dim rstMasterAsset As DAO.Recordset
        Dim strSQL As String
        Dim strWorkSQL As String
        Dim blnContinue As Boolean
        Dim blnNoMAsset As Boolean  ' ** TRUE if no MasterAsset involved.
        Dim intRetVal_BuildAssetListInfo As Integer, intRetVal_SetDateSpecificSQL As Integer

14210   blnContinue = True
14220   intRetVal_BuildAssetListInfo = 0   ' ** Unless proven otherwise.

14230   With frm
14240     If .cmbAccounts.Visible = True Then
14250       .cmbAccounts.SetFocus
14260       If .cmbAccounts.text = vbNullString Then
14270         DoCmd.Hourglass False
14280         MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
14290         intRetVal_BuildAssetListInfo = -3  ' ** Missing entry.
14300       End If
            ' ** Return focus to the button that called this.
14310       If Right(strControlName, 6) = "_Click" Then
              ' ** cmdPrint00_Click.
14320         .Controls(Left(strControlName, (Len(strControlName) - 6))).SetFocus
14330         DoEvents
14340       End If
14350     End If

14360     If intRetVal_BuildAssetListInfo = 0 Then
14370       If IsNull(varEndDate) Then
14380         DoCmd.Hourglass False
14390         MsgBox "Must enter Period " & strWhen & " date to continue.", vbInformation + vbOKOnly, "Entry Required"
14400         .DateEnd.SetFocus
14410         intRetVal_BuildAssetListInfo = -3  ' ** Missing entry.
14420       End If
14430     End If

14440     If intRetVal_BuildAssetListInfo = 0 Then
14450       DoEvents

            ' ** This code will update the qryMaxBalDates query to give
            ' ** us the balance numbers from the previous statement.
14460       intRetVal_SetDateSpecificSQL = SetDateSpecificSQL(.cmbAccounts, "Statements", frm.Name, varStartDate, varEndDate)  ' ** Module Function: modUtilities.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -4  Date criteria not met.
            ' **   -9  Error.

14470       DoEvents
14480       If intRetVal_SetDateSpecificSQL <> 0 Then
14490         blnContinue = False
14500         intRetVal_BuildAssetListInfo = -2  ' ** No data.
14510       End If
14520     End If

14530     If intRetVal_BuildAssetListInfo = 0 And blnContinue = True Then

14540       Set dbs = CurrentDb

            ' ** Empty tmpAssetList2.
14550       If TableExists("tmpAssetList2") = True Then  ' ** Module Function: modUtilities.
14560         Set qdf = dbs.QueryDefs("qryCourtReport_03")
14570         qdf.Execute
14580       End If
            ' ** Empty tmpAccountInfo.
14590       If TableExists("tmpAccountInfo") = True Then  ' ** Module Function: modUtilities.
14600         Set qdf = dbs.QueryDefs("qryCourtReport_04")
14610         qdf.Execute
14620       End If

14630       DoEvents
            ' ** qryCourtReport_NS_07z.
            ' ** VGC 11/25/2009: «='90',-1,1)» to «='90',1,1)».
14640       strSQL = "SELECT ActiveAssets.assetno, masterasset.description AS MasterAssetDescription, masterasset.due, " & _
              "masterasset.rate, Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, " & _
              "Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface]))*" & _
              "IIf([assettype].[assettype]='90',1,1) AS TotalShareface, account.accountno, account.shortname, " & _
              "account.legalname, assettype.assettype, assettype_description, " & _
              "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
              "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
              "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc, "
14650       strSQL = strSQL & "account.icash, account.pcash, masterasset.currentDate, " & _
              "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX, " & _
              "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX, " & _
              "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX, " & CoInfo & " "
14660       strSQL = strSQL & "FROM (masterasset LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) " & _
              "RIGHT JOIN (account LEFT JOIN ActiveAssets ON account.accountno = ActiveAssets.accountno) " & _
              "ON masterasset.assetno = ActiveAssets.assetno " & _
              "GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, " & _
              "account.accountno, account.shortname, account.legalname, assettype.assettype, assettype_description, " & _
              "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & "
14670       strSQL = strSQL & "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
              "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), " & _
              "account.icash, account.pcash, masterasset.currentDate, " & _
              "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), " & _
              "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]), " & _
              "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) " & _
              "HAVING (((account.accountno)='" & .cmbAccounts & "'));"
14680       strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.

14690       Set rst = dbs.OpenRecordset(strSQL)  ' ** This is what's used for appending to the tmpAssetList2 table and rstTempAssetList.
14700       DoEvents

            ' ** Ledger records to roll back if after specified 'To' date. (VGC 08/24/2010: CHANGES!)
14710       Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_B_03")
14720       With qdf.Parameters
14730         ![actno] = frm.cmbAccounts
14740         ![datEnd] = CDate(varEndDate)
14750       End With
14760       Set rstWork = qdf.OpenRecordset
14770       If rstWork.BOF = True And rstWork.EOF = True Then
              ' ** In the absense of newer data to roll back, take the total cost as it's normally computed for report 0.
14780         blnContinue = False  ' ** Leave intRetVal_BuildAssetListInfo = 0 so that the procedure calling this continues.
14790         rstWork.Close
14800         If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
14810           dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
14820         End If
14830         rst.Close
14840         dbs.Close
14850         strRpt = "_00BA"
14860       End If
14870       DoEvents

14880       If blnContinue Then  ' ** There are transactions after the specified To date.

14890         Set rstAccount = dbs.OpenRecordset("SELECT account.* FROM account ORDER BY accountno", dbOpenSnapshot)
14900         Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_B_04")
14910         Set rstMasterAsset = qdf.OpenRecordset()
14920         DoEvents

14930         If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
14940           dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
14950           Set rstTempAssetList = dbs.OpenRecordset("SELECT tmpAssetList2.* FROM tmpAssetList2", dbOpenDynaset)
14960         Else
                ' ** ERROR.
14970           intRetVal_BuildAssetListInfo = -9  ' ** Error.
14980           DoCmd.Hourglass False
14990           MsgBox "Unable to create temporary table for reporting.", vbCritical + vbOKOnly, "Error"
15000           rstAccount.Close
15010           rstMasterAsset.Close
15020           rstWork.Close
15030           dbs.Close
15040         End If
15050         DoEvents

15060         If intRetVal_BuildAssetListInfo = 0 Then

                ' ** Create table to contain temporary account master records to track "global" account info.
15070           If MakeTempTable(rst, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
15080             DoEvents
15090             dbs.Execute "INSERT INTO tmpAccountInfo " & strSQL  ' ** Copy data from query.
                  ' ** Now, make records here distinct for each account.
15100             dbs.Execute "UPDATE tmpAccountInfo SET assetno = Null, MasterAssetDescription = Null, " & _
                    "due = Null, rate = Null, TotalCost = Null, TotalShareFace = Null,  " & _
                    "assettype = Null, assettype_description = Null, TotDesc = Null,  " & _
                    "currentDate = Null, MarketValueX = Null, MarketValueCurrentX = Null, " & _
                    "YieldX = Null  " & _
                    "WHERE True"
15110             Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT DISTINCTROW tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenSnapshot)
15120             DoEvents
                  ' ** Empty tmpAccountInfo.
15130             Set qdf = dbs.QueryDefs("qryCourtReport_04")
15140             qdf.Execute
15150             If rstTmpAccountInfo.RecordCount > 0 Then
15160               rstTmpAccountInfo.MoveLast
15170               rstTmpAccountInfo.MoveFirst
15180             End If
15190             If Not CopyToTempTable(rstTmpAccountInfo, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
                    ' ** ERROR.
15200               intRetVal_BuildAssetListInfo = -9  ' ** Error.
15210               DoCmd.Hourglass False
15220               MsgBox "Unable to copy data to temporary table.", vbCritical + vbOKOnly, "Error"
15230             Else
15240               rstTmpAccountInfo.Close  ' ** Close old snapshot, then open as dynaset.
15250               Set rstTmpAccountInfo = Nothing
15260               Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenDynaset)
15270             End If
15280             DoEvents
15290           Else
                  ' ** ERROR.
15300             intRetVal_BuildAssetListInfo = -9  ' ** Error.
15310             DoCmd.Hourglass False
15320             MsgBox "Unable to create temporary account table for reporting.", vbCritical + vbOKOnly, "Error"
15330             rstAccount.Close
15340             rstMasterAsset.Close
15350             rstWork.Close
15360             rstTempAssetList.Close
15370             dbs.Close
15380           End If

15390         End If  ' ** intRetVal_BuildAssetListInfo.

15400         If intRetVal_BuildAssetListInfo = 0 Then

15410           rstTmpAccountInfo.MoveLast
15420           rstTmpAccountInfo.MoveFirst

                ' ** Get total number of records in recordset of subsequent ledger entries.
15430           rstWork.MoveLast
15440           rstWork.MoveFirst

15450           DoEvents
15460           Do While Not rstWork.EOF  ' ** Move through each record, tracking changes.

15470             rstTempAssetList.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "' AND [assetno] = " & CStr(rstWork![assetno])
15480             If rstTempAssetList.NoMatch Then
                    ' ** We need a new record because this asset was not in the reporting query after subsequent changes.
15490               rstMasterAsset.FindFirst "[assetno] = " & CStr(rstWork![assetno])
15500               blnNoMAsset = rstMasterAsset.NoMatch
15510               If blnNoMAsset Then
15520                 Select Case Trim(rstWork![journaltype])
                      Case "Misc.", "Paid", "Received"
                        ' ** OK to continue.
15530                 Case Else
                        ' ** ERROR.
15540                   intRetVal_BuildAssetListInfo = -9  ' ** Error.
15550                   DoCmd.Hourglass False
15560                   MsgBox "Missing Master Asset record.", vbCritical + vbOKOnly, "Error"
15570                   rstAccount.Close
15580                   rstMasterAsset.Close
15590                   rstWork.Close
15600                   rstTempAssetList.Close
15610                   rstTmpAccountInfo.Close
15620                   dbs.Close
15630                 End Select
15640               Else
15650                 rstTempAssetList.AddNew
15660                 rstTempAssetList![accountno] = Trim(rstWork![accountno])
15670                 rstTempAssetList![assetno] = rstMasterAsset![assetno]
15680                 rstTempAssetList![MasterAssetDescription] = rstMasterAsset![description]
15690                 rstTempAssetList![assettype] = rstMasterAsset![assettype]
15700                 rstTempAssetList![assettype_description] = rstMasterAsset![assettype_description]
15710                 rstTempAssetList![due] = rstMasterAsset![due]
15720                 rstTempAssetList![rate] = rstMasterAsset![rate]
15730                 rstTempAssetList![totdesc] = CStr(rstMasterAsset![description]) & _
                        IIf(rstMasterAsset![rate] > 0, " " & Format(rstMasterAsset![rate], "0.000%"), "") & _
                        IIf(Not IsNull(rstMasterAsset![due]), "  Due " & Format(rstMasterAsset![due], "mm/dd/yyyy"), "")
15740                 rstTempAssetList![TotalCost] = 0
15750                 rstTempAssetList![TotalShareFace] = 0
15760                 rstTempAssetList![ICash] = 0
15770                 rstTempAssetList![PCash] = 0
                      ' ** LEAVE in editing mode.
15780               End If
15790             Else
15800               blnNoMAsset = False
15810               rstTempAssetList.Edit  ' ** Edit existing temp record.
15820             End If
15830             DoEvents

15840             If intRetVal_BuildAssetListInfo = 0 Then

15850               rstTmpAccountInfo.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"

15860               If rstTmpAccountInfo.NoMatch Then
                      ' ** We need a new account record.

15870                 rstAccount.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"
15880                 If rstAccount.NoMatch Then
                        ' ** ERROR.
15890                   intRetVal_BuildAssetListInfo = -9  ' ** Error.
15900                   DoCmd.Hourglass False
15910                   MsgBox "Missing Account master record.", vbCritical + vbOKOnly, "Error"
15920                   rstAccount.Close
15930                   rstMasterAsset.Close
15940                   rstWork.Close
15950                   rstTempAssetList.Close
15960                   rstTmpAccountInfo.Close
15970                   dbs.Close
15980                 Else
15990                   rstTmpAccountInfo.AddNew
                        ' ** Copy info from Account table.
16000                   rstTmpAccountInfo![accountno] = rstAccount![accountno]
16010                   rstTmpAccountInfo![shortname] = rstAccount![shortname]
16020                   rstTmpAccountInfo![legalname] = rstAccount![legalname]
16030                   rstTmpAccountInfo![ICash] = rstAccount![ICash]
16040                   rstTmpAccountInfo![PCash] = rstAccount![PCash]
                        ' ** LEAVE in editing mode.
16050                 End If
16060               Else
16070                 rstTmpAccountInfo.Edit  ' ** Edit existing temp record.
16080               End If

16090             End If  ' ** intRetVal_BuildAssetListInfo.
16100             DoEvents

16110             If intRetVal_BuildAssetListInfo = 0 Then

16120               If (Not blnNoMAsset) And (Trim(rstWork![journaltype]) <> "Received") Then
                      ' ** Add/Subtract info from ASSET record. NOTE: These are SUBTRACTIONS
                      ' ** because the query returns the changes since the date.
                      'VGC 08/20/2010: CHANGES!
16130                 rstTempAssetList![TotalShareFace] = _
                        rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, _
                        rstWork![shareface] * IIf(rstTempAssetList![assettype] = 90, 1, 1)))
                      'rstTempAssetList![TotalShareFace] = _
                      '  rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, rstWork![shareface] * _
                      '  IIf(rstTempAssetList![assettype] = 90, -1, 1)))
16140                 rstTempAssetList![TotalCost] = rstTempAssetList![TotalCost] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost])
                      ' ** Save ASSET temp table record.
16150                 rstTempAssetList.Update
16160               End If

                    ' ** Add/Subtract info from ACCOUNT record. NOTE: These are SUBTRACTIONS
                    ' ** because the query returns the changes since the date.
16170               rstTmpAccountInfo![ICash] = rstTmpAccountInfo![ICash] - rstWork![ICash]
16180               rstTmpAccountInfo![PCash] = rstTmpAccountInfo![PCash] - rstWork![PCash]
                    ' ** Save ACCOUNT temp table record.
16190               rstTmpAccountInfo.Update

16200               rstWork.MoveNext

16210             End If  ' ** intRetVal_BuildAssetListInfo.
16220             DoEvents

16230             If intRetVal_BuildAssetListInfo <> 0 Then Exit Do

16240           Loop  ' ** Move through each record, tracking changes.

16250         End If  ' ** intRetVal_BuildAssetListInfo.

16260         If intRetVal_BuildAssetListInfo = 0 Then

                ' ** Remove any asset records which now have a ZERO totalshareface AND totalcost.
                'strWorkSQL = "DELETE tmpAssetList2.*, tmpAssetList2.TotalShareface, tmpAssetList2.TotalCost " & _
                '  "FROM tmpAssetList2 " & _
                '  "WHERE ((Round((tmpAssetList2.TotalShareface),4)=0) AND (Round((tmpAssetList2.TotalCost),2)=0));"
                ' ** Delete tmpAssetList2, for zero entries.
                ' ** 07/21/2008: Added rounding.
16270           Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_B_05")
16280           qdf.Execute
16290           DoEvents

                ' ** Now, update tmpAssetList2 to reflect the needed account information which is
                ' ** global to the report.
16300           rstTmpAccountInfo.MoveFirst
16310           While Not rstTmpAccountInfo.EOF
16320             strWorkSQL = "UPDATE tmpAssetList2 SET " & _
                    "shortname = " & SQLFormatStr(rstTmpAccountInfo![shortname], dbText) & ", " & _
                    "legalname = " & SQLFormatStr(rstTmpAccountInfo![legalname], dbText) & ", " & _
                    "icash = " & SQLFormatStr(rstTmpAccountInfo![ICash], dbCurrency) & ", " & _
                    "pcash = " & SQLFormatStr(rstTmpAccountInfo![PCash], dbCurrency) & " " & _
                    "WHERE accountno = " & SQLFormatStr(rstTmpAccountInfo![accountno], dbText) & ";"
16330             dbs.Execute strWorkSQL
16340             rstTmpAccountInfo.MoveNext
16350           Wend
16360           DoEvents

                ' ** Close temp & working recordsets.
16370           rstAccount.Close
16380           rstMasterAsset.Close
16390           rstWork.Close
16400           rstTempAssetList.Close
16410           rstTmpAccountInfo.Close

                ' ** Finally, base report on this instead of on qryAssetList.
16420           Set qdf = dbs.QueryDefs("qryCourtReport_NY_00_B_01")
16430           Set rst = qdf.OpenRecordset()
16440           strRpt = "_00B"

16450         End If  ' ** intRetVal_BuildAssetListInfo.

16460       End If  ' ** blnContinue.
16470       DoEvents

16480       If intRetVal_BuildAssetListInfo = 0 And blnContinue = True Then
16490         If rst.BOF = True And rst.EOF = True Then
16500           If strWhen = "Beginning" Then
                  ' ** This is OK!
                  ' ** It just means there's no beginning data.
16510             rst.Close
                  ' ** Append qryCourtReport_NY_01_46 (qryCourtReport_NY_01_45 (qryCourtReport_NY_01_44
                  ' ** (qryCourtReport_07 (qryCourtReport_06 (Ledger, linked to Account, qryCourtReport_05
                  ' ** (Balance, grouped by accountno, with max Balance Date), with add'l fields), grouped
                  ' ** and summed); excluding Misc. [icash]+[pcash]=0), grouped and summed), linked to
                  ' ** Account, just 1 dummy record) to tmpAssetList2.
16520             Set qdf = dbs.QueryDefs("qryCourtReport_NY_01_47")
16530             qdf.Execute
16540           Else
16550             rst.Close
16560             DoCmd.Hourglass False
16570             MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
16580             intRetVal_BuildAssetListInfo = -2  ' ** No data.
16590           End If
16600         End If
16610         dbs.Close
16620       End If
16630       DoEvents

16640     End If  ' ** intRetVal_BuildAssetListInfo.
16650   End With

EXITP:
16660   Set rstWork = Nothing
16670   Set rstTempAssetList = Nothing
16680   Set rstTmpAccountInfo = Nothing
16690   Set rstAccount = Nothing
16700   Set rstMasterAsset = Nothing
16710   Set rst = Nothing
16720   Set qdf = Nothing
16730   Set dbs = Nothing
16740   BuildAssetListInfo_NY = intRetVal_BuildAssetListInfo
16750   Exit Function

ERRH:
16760   intRetVal_BuildAssetListInfo = -9  ' ** Error.
16770   DoCmd.Hourglass False
16780   Select Case ERR.Number
        Case Else
16790     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16800   End Select
16810   Resume EXITP

End Function

Public Sub WordAll_NY(frm As Access.Form)

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "WordAll_NY"

        Dim strDocName As String
        Dim blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String
        Dim lngX As Long

        ' ** Access: 14272506  Very Light Red
        ' ** Access: 12295153  Medium Red
        ' ** Word:   16770233  Very Light Blue
        ' ** Word:   16434048  Medium Blue
        ' ** Excel:  14677736  Very Light Green
        ' ** Excel:  5952646   Medium Green

16910   With frm
16920     If .Validate = True Then  ' ** Function: Below.

16930       DoCmd.Hourglass True
16940       DoEvents

16950       .cmdWordAll_box01.Visible = True
16960       .cmdWordAll_box02.Visible = True
16970       If .chkAssetList = True Then
16980         .cmdWordAll_box03.Visible = True
16990       End If
17000       .cmdWordAll_box04.Visible = True
17010       DoEvents

17020       blnExcel = False
17030       blnAllCancel = False
17040       .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
17050       AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
17060       blnAutoStart = .chkOpenWord

17070       Beep
17080       DoCmd.Hourglass False
17090       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Word" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

17100       If msgResponse = vbOK Then

17110         DoCmd.Hourglass True
17120         DoEvents

17130         blnAllCancel = False
17140         .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
17150         AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
17160         gblnPrintAll = True
17170         blnAutoStart = False  ' ** They'll open only after all have been exported.
17180         strThisProc = "cmdWordAll_Click"

17190         lngFiles = 0&
17200         ReDim arr_varFile(F_ELEMS, 0)
17210         FileArrayInit  ' ** Module Procedure: modCourtReportsNY2.

              ' ** Get the Summary inputs first.

17220         If blnAllCancel = False Then
                ' ** Summary Statement.
17230           .cmdWord00.SetFocus
17240           .cmdWord00_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17250           DoEvents
17260         End If
17270         If blnAllCancel = False Then
                ' ** Principal Received.
17280           .cmdWord01.SetFocus
17290           .cmdWord01_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17300           DoEvents
17310         End If
17320         If blnAllCancel = False Then
                ' ** Increases On Sales, Liquidation or Distribution.
17330           .cmdWord02.SetFocus
17340           .cmdWord02_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17350           DoEvents
17360         End If
17370         If blnAllCancel = False Then
                ' ** Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability.
17380           .cmdWord03.SetFocus
17390           .cmdWord03_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17400           DoEvents
17410         End If
17420         If blnAllCancel = False Then
                ' ** Administration Expenses Chargeable to Principal.
17430           .cmdWord04.SetFocus
17440           .cmdWord04_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17450           DoEvents
17460         End If
17470         If blnAllCancel = False Then
                ' ** Distributions of Principal.
17480           .cmdWord05.SetFocus
17490           .cmdWord05_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17500           DoEvents
17510         End If
17520         If blnAllCancel = False Then
                ' ** New Investments, Exchanges and Stock Distributions of Principal Assets.
17530           .cmdWord06.SetFocus
17540           .cmdWord06_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17550           DoEvents
17560         End If
17570         If blnAllCancel = False Then
                ' ** Principal Remaining on Hand.
17580           .cmdWord07.SetFocus
17590           .cmdWord07_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17600           DoEvents
17610         End If
17620         If blnAllCancel = False Then
                ' ** Income Received.
17630           .cmdWord08.SetFocus
17640           .cmdWord08_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17650           DoEvents
17660         End If
17670         If blnAllCancel = False Then
                ' ** All Income Collected.
17680           .cmdWord09.SetFocus
17690           .cmdWord09_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17700           DoEvents
17710         End If
17720         If blnAllCancel = False Then
                ' ** Administration Expenses Chargeable to Income.
17730           .cmdWord10.SetFocus
17740           .cmdWord10_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17750           DoEvents
17760         End If
17770         If blnAllCancel = False Then
                ' ** Distributions of Income.
17780           .cmdWord11.SetFocus
17790           .cmdWord11_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17800           DoEvents
17810         End If
17820         If blnAllCancel = False Then
                ' ** Income Remaining on Hand.
17830           .cmdWord12.SetFocus
17840           .cmdWord12_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
17850           DoEvents
17860         End If

17870         DoCmd.Hourglass True
17880         DoEvents

17890         .cmdWordAll.SetFocus

17900         gblnPrintAll = False
17910         Beep

17920         If lngFiles > 0& Then

17930           DoCmd.Hourglass False

17940           strTmp01 = CStr(lngFiles) & " documents were created."
17950           If .chkOpenExcel = True Then
17960             strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
17970             msgResponse = MsgBox(strTmp01, vbInformation + vbOKCancel, "Reports Exported")
17980           Else
17990             msgResponse = MsgBox(strTmp01, vbInformation + vbOKOnly, "Reports Exported")
18000           End If

18010           .cmdWordAll_box01.Visible = False
18020           .cmdWordAll_box02.Visible = False
18030           .cmdWordAll_box03.Visible = False
18040           .cmdWordAll_box04.Visible = False

18050           If .chkOpenWord = True And msgResponse = vbOK Then
18060             DoCmd.Hourglass True
18070             DoEvents
18080             For lngX = 0& To (lngFiles - 1&)
18090               strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
18100               OpenExe strDocName  ' ** Module Function: modShellFuncs.
18110               DoEvents
18120               If lngX < (lngFiles - 1&) Then
18130                 ForcePause 2  ' ** Module Function: modCodeUtilities.
18140               End If
18150             Next
18160             Beep
18170           End If

                'CourtReport_NY_Summary_11_150101_To_151231
'1  rptCourtRptNY_00A  CourtReport_NY_Summary_11_150101_To_151231.rtf
                'Y  CourtReport_NY_Property_on_Hand_at_Ending_of_Account_Period_11_150101_To_151231
                'Y  2  rptCourtRptNY_00BA  CourtReport_NY_Property_on_Hand_at_Ending_of_Account_Period_11_150101_To_151231.rtf
                'CourtReport_NY_Principal_Received_11_150101_To_151231
'3  rptCourtRptNY_01  CourtReport_NY_Principal_Received_11_150101_To_151231.rtf
                'CourtReport_NY_Increases_On_Sales_Liquidation_or_Distribution_11_150101_To_151231
'4  rptCourtRptNY_02  CourtReport_NY_Increases_On_Sales_Liquidation_or_Distribution_11_150101_To_151231.rtf
                'CourtReport_NY_Decreases_Due_to_Sales_Liquidation_Collection_Distribution_or_Uncollectability_11_150101_To_151231
'5  rptCourtRptNY_03  CourtReport_NY_Decreases_Due_to_Sales_Liquidation_Collection_Distribution_or_Uncollectability_11_150101_To_151231.rtf
                'CourtReport_NY_Administration_Expenses_Chargeable_to_Principal_11_150101_To_151231
'6  rptCourtRptNY_04A  CourtReport_NY_Administration_Expenses_Chargeable_to_Principal_11_150101_To_151231.rtf
                'CourtReport_NY_Distributions_of_Principal_11_150101_To_151231
'7  rptCourtRptNY_05  CourtReport_NY_Distributions_of_Principal_11_150101_To_151231.rtf
                'CourtReport_NY_New_Investments_Exchanges_and_Stock_Distributions_of_Principal_Assets_11_150101_To_151231
'8  rptCourtRptNY_06  CourtReport_NY_New_Investments_Exchanges_and_Stock_Distributions_of_Principal_Assets_11_150101_To_151231.rtf
                'CourtReport_NY_Principal_Remaining_on_Hand_11_150101_To_151231
'9  rptCourtRptNY_07  CourtReport_NY_Principal_Remaining_on_Hand_11_150101_To_151231.rtf
                'CourtReport_NY_Income_Received_11_150101_To_151231
'10  rptCourtRptNY_08  CourtReport_NY_Income_Received_11_150101_To_151231.rtf
                'CourtReport_NY_All_Income_Collected_William B. Johnson Trust_150101_To_151231
'11  rptCourtRptNY_09A  CourtReport_NY_All_Income_Collected_William B. Johnson Trust_150101_To_151231.rtf
                'CourtReport_NY_Administration_Expenses_Chargeable_to_Income_11_150101_To_151231
'12  rptCourtRptNY_10A  CourtReport_NY_Administration_Expenses_Chargeable_to_Income_11_150101_To_151231.rtf
                'CourtReport_NY_Distributions_of_Income_11_150101_To_151231
'13  rptCourtRptNY_11  CourtReport_NY_Distributions_of_Income_11_150101_To_151231.rtf
                'CourtReport_NY_Income_Remaining_on_Hand_11_150101_To_151231
'14  rptCourtRptNY_12  CourtReport_NY_Income_Remaining_on_Hand_11_150101_To_151231.rtf
                'FILES: 14
'1  rptCourtRptNY_00A
'2  rptCourtRptNY_00BA
'3  rptCourtRptNY_01
'4  rptCourtRptNY_02
'5  rptCourtRptNY_03
'6  rptCourtRptNY_04A
'7  rptCourtRptNY_05
'8  rptCourtRptNY_06
'9  rptCourtRptNY_07
'10  rptCourtRptNY_08
'11  rptCourtRptNY_09A
'12  rptCourtRptNY_10A
'13  rptCourtRptNY_11
'14  rptCourtRptNY_12

18180         Else
18190           DoCmd.Hourglass False
18200           MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
18210           .cmdWordAll_box01.Visible = False
18220           .cmdWordAll_box02.Visible = False
18230           .cmdWordAll_box03.Visible = False
18240           .cmdWordAll_box04.Visible = False
18250         End If  ' ** lngFiles.

18260       End If  ' ** msgResponse.

18270     End If  ' ** Validate.

18280     DoCmd.Hourglass False

18290   End With

EXITP:
18300   Exit Sub

ERRH:
18310   blnAllCancel = True
18320   frm.AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
18330   AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
18340   gblnPrintAll = False
18350   DoCmd.Hourglass False
18360   Select Case ERR.Number
        Case Else
18370     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18380   End Select
18390   Resume EXITP

End Sub

Public Sub ExcelAll_NY(frm As Access.Form)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "ExcelAll_NY"

        Dim strDocName As String
        Dim blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim strTmp01 As String
        Dim lngX As Long

        ' ** Access: 14272506  Very Light Red
        ' ** Access: 12295153  Medium Red
        ' ** Word:   16770233  Very Light Blue
        ' ** Word:   16434048  Medium Blue
        ' ** Excel:  14677736  Very Light Green
        ' ** Excel:  5952646   Medium Green

18410   With frm
18420     If .Validate = True Then  ' ** Function: Below.

18430       DoCmd.Hourglass True
18440       DoEvents

18450       .cmdExcelAll_box01.Visible = True
18460       .cmdExcelAll_box02.Visible = True
18470       If .chkAssetList = True Then
18480         .cmdExcelAll_box03.Visible = True
18490       End If
18500       .cmdExcelAll_box04.Visible = True
18510       DoEvents

18520       blnExcel = True
18530       blnAllCancel = False
18540       .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
18550       AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
18560       blnAutoStart = .chkOpenExcel

18570       Beep
18580       DoCmd.Hourglass False
18590       msgResponse = MsgBox("This will send all highlighted reports to Microsoft Excel" & _
              IIf(blnAutoStart = True, ", " & vbCrLf & "then open them at the end of the process.", ".") & _
              vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
              "Send All Reports To Microsoft " & IIf(blnExcel = True, "Excel.", "Word."))

18600       If msgResponse = vbOK Then

18610         DoCmd.Hourglass True
18620         DoEvents

18630         blnAllCancel = False
18640         .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
18650         AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
18660         gblnPrintAll = True
18670         blnAutoStart = False  ' ** They'll open only after all have been exported.
18680         strThisProc = "cmdExcelAll_Click"

18690         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                ' ** It seems like it's not quite closed when it gets here,
                ' ** because if I stop the code and run the function again,
                ' ** it always comes up False.
18700           ForcePause 2  ' ** Module Function: modCodeUtilities.
18710           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18720             DoCmd.Hourglass False
18730             msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                    "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                    "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                    "You may close Excel before proceding, then click Retry." & vbCrLf & _
                    "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
                  ' ** ... Otherwise Trust Accountant will do it for you.
18740             If msgResponse <> vbRetry Then
18750               blnAllCancel = True
18760               .AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
18770               AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
18780             End If
18790           End If
18800         End If

18810         If blnAllCancel = False Then

18820           DoCmd.Hourglass True
18830           DoEvents

18840           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18850             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
18860           End If
18870           DoEvents

18880           lngFiles = 0&
18890           ReDim arr_varFile(F_ELEMS, 0)
18900           FileArrayInit  ' ** Module Procedure: modCourtReportsNY2.

18910           If blnAllCancel = False Then
                  ' ** Summary Statement.
18920             .cmdExcel00.SetFocus
18930             .cmdExcel00_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
18940             DoEvents
18950           End If
18960           If blnAllCancel = False Then
                  ' ** Principal Received.
18970             .cmdExcel01.SetFocus
18980             .cmdExcel01_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
18990             DoEvents
19000           End If
19010           If blnAllCancel = False Then
                  ' ** Increases On Sales, Liquidation or Distribution.
19020             .cmdExcel02.SetFocus
19030             .cmdExcel02_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19040             DoEvents
19050           End If
19060           If blnAllCancel = False Then
                  ' ** Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability.
19070             .cmdExcel03.SetFocus
19080             .cmdExcel03_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19090             DoEvents
19100           End If
19110           If blnAllCancel = False Then
                  ' ** Administration Expenses Chargeable to Principal.
19120             .cmdExcel04.SetFocus
19130             .cmdExcel04_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19140             DoEvents
19150           End If
19160           If blnAllCancel = False Then
                  ' ** Distributions of Principal.
19170             .cmdExcel05.SetFocus
19180             .cmdExcel05_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19190             DoEvents
19200           End If
19210           If blnAllCancel = False Then
                  ' ** New Investments, Exchanges and Stock Distributions of Principal Assets.
19220             .cmdExcel06.SetFocus
19230             .cmdExcel06_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19240             DoEvents
19250           End If
19260           If blnAllCancel = False Then
                  ' ** Principal Remaining on Hand.
19270             .cmdExcel07.SetFocus
19280             .cmdExcel07_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19290             DoEvents
19300           End If
19310           If blnAllCancel = False Then
                  ' ** Income Received.
19320             .cmdExcel08.SetFocus
19330             .cmdExcel08_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19340             DoEvents
19350           End If
19360           If blnAllCancel = False Then
                  ' ** All Income Collected.
19370             .cmdExcel09.SetFocus
19380             .cmdExcel09_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19390             DoEvents
19400           End If
19410           If blnAllCancel = False Then
                  ' ** Administration Expenses Chargeable to Income.
19420             .cmdExcel10.SetFocus
19430             .cmdExcel10_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19440             DoEvents
19450           End If
19460           If blnAllCancel = False Then
                  ' ** Distributions of Income.
19470             .cmdExcel11.SetFocus
19480             .cmdExcel11_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19490             DoEvents
19500           End If
19510           If blnAllCancel = False Then
                  ' ** Income Remaining on Hand.
19520             .cmdExcel12.SetFocus
19530             .cmdExcel12_Click  ' ** Form Procedure: frmRpt_CourtReports_NY.
19540             DoEvents
19550           End If

19560           DoCmd.Hourglass True
19570           DoEvents

19580           .cmdExcelAll.SetFocus

19590           gblnPrintAll = False
19600           Beep

19610           If lngFiles > 0& And blnAllCancel = False Then

19620             DoCmd.Hourglass False

19630             strTmp01 = CStr(lngFiles) & " documents were created."
19640             If .chkOpenExcel = True Then
19650               strTmp01 = strTmp01 & vbCrLf & vbCrLf & "Documents will open when this message closes."
19660               msgResponse = MsgBox(strTmp01, vbInformation + vbOKCancel, "Reports Exported")
19670             Else
19680               msgResponse = MsgBox(strTmp01, vbInformation + vbOKOnly, "Reports Exported")
19690             End If

19700             .cmdExcelAll_box01.Visible = False
19710             .cmdExcelAll_box02.Visible = False
19720             .cmdExcelAll_box03.Visible = False
19730             .cmdExcelAll_box04.Visible = False

19740             If .chkOpenExcel = True And msgResponse = vbOK Then
19750               DoCmd.Hourglass True
19760               DoEvents
19770               For lngX = 0& To (lngFiles - 1&)
19780                 strDocName = arr_varFile(F_PATH, lngX) & LNK_SEP & arr_varFile(F_FILE, lngX)
19790                 OpenExe strDocName  ' ** Module Function: modShellFuncs.
19800                 DoEvents
19810                 If lngX < (lngFiles - 1&) Then
19820                   ForcePause 2  ' ** Module Function: modCodeUtilities.
19830                 End If
19840               Next
19850             End If

19860           Else
19870             DoCmd.Hourglass False
19880             MsgBox "No files were exported.", vbInformation + vbOKOnly, "Nothing To Do"
19890             .cmdExcelAll_box01.Visible = False
19900             .cmdExcelAll_box02.Visible = False
19910             .cmdExcelAll_box03.Visible = False
19920             .cmdExcelAll_box04.Visible = False
19930           End If

19940         End If  ' ** blnAllCancel.

19950       End If  ' ** msgResponse.

19960     End If  ' ** Validate.

19970     DoCmd.Hourglass False

19980   End With

EXITP:
19990   Exit Sub

ERRH:
20000   blnAllCancel = True
20010   frm.AllCancelSet1_NY blnAllCancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
20020   AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
20030   gblnPrintAll = False
20040   DoCmd.Hourglass False
20050   Select Case ERR.Number
        Case Else
20060     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20070   End Select
20080   Resume EXITP

End Sub

Public Sub FileArraySet_NY(arr_varTmp00 As Variant)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "FileArraySet_NY"

        Dim blnFound As Boolean
        Dim lngTmp01 As Long
        Dim lngX As Long, lngY As Long, lngE As Long

20110   lngTmp01 = UBound(arr_varTmp00, 2) + 1&
20120   For lngX = 0& To (lngTmp01 - 1&)
20130     blnFound = False
20140     For lngY = 0& To (lngFiles - 1&)
20150       If arr_varFile(F_RNAM, lngY) = arr_varTmp00(F_RNAM, lngX) Then
20160         blnFound = True
20170         Exit For
20180       End If
20190     Next  ' ** lngY.
20200     If blnFound = False Then
20210       lngFiles = lngFiles + 1&
20220       lngE = lngFiles - 1&
20230       ReDim Preserve arr_varFile(F_ELEMS, lngE)
20240       arr_varFile(F_RNAM, lngE) = arr_varTmp00(F_RNAM, lngX)
20250       arr_varFile(F_FILE, lngE) = arr_varTmp00(F_FILE, lngX)
20260       arr_varFile(F_PATH, lngE) = arr_varTmp00(F_PATH, lngX)
20270     End If
20280   Next  ' ** lngX.

EXITP:
20290   Exit Sub

ERRH:
20300   Select Case ERR.Number
        Case Else
20310     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20320   End Select
20330   Resume EXITP

End Sub

Public Sub AllCancelSet2_NY(blnCancel As Boolean)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet2_NY"

20410   blnAllCancel = blnCancel

EXITP:
20420   Exit Sub

ERRH:
20430   Select Case ERR.Number
        Case Else
20440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20450   End Select
20460   Resume EXITP

End Sub
