Attribute VB_Name = "modCourtReportsNY2"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCourtReportsNY2"

'VGC 09/08/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  39520
' #########################

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
Private strThisProc As String
' **

Public Sub NY_CmdPrev00_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' **
' ** NY_CmdPrev00_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

100   On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev00_Click"

        Dim strRpt As String, strDocName As String
        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer

110     With frm

120       DoCmd.Hourglass True
130       DoEvents

140       blnContinue = True
150       strThisProc = "cmdPreview00_Click"

160       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

170         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

180         DoCmd.Hourglass False
190         strDocName = "frmRpt_CourtReports_NY_Input"
200         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm

210         DoCmd.Hourglass True
220         DoEvents

            ' ** Run function to fill Schedule A Asset List data.
230         If .CashAssets_Beg <> vbNullString Then

240           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm) ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

250           If intRetVal_BuildAssetListInfo = -2 Then
260             gcurCrtRpt_NY_IncomeBeg = 0@
270           Else
280             gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
290           End If

              ' ** Run function to empty and fill the tmpCourtReportData table.
300           blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
              ' ** No need to communicate this further, since there is no 'PreviewAll'.

310           If blnContinue = True Then

320             gstrAccountNo = .cmbAccounts.Column(0)

330             Select Case gblnUseReveuneExpenseCodes
                Case True
340               PreviewOrPrint_NY "0A", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
350             Case False
360               PreviewOrPrint_NY "0", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
370             End Select

380             If .chkAssetList = True Then
390               AssetList_PreviewPrint_NY acViewPreview, strCallingForm, frm, strThisProc  ' ** Function: Above.
400             End If

410           End If  ' ** blnContinue.

420         End If  ' ** CashAssets_Beg.

430       End If  ' ** Validate.

440       DoCmd.Hourglass False

450     End With

EXITP:
460     Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint00_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Summary Statement.
' **
' ** NY_CmdPrint00_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

500   On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint00_Click"

        Dim strRpt As String, strDocName As String
        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer

510     With frm

520       DoCmd.Hourglass True
530       DoEvents

540       blnContinue = True
550       strThisProc = "cmdPrint00_Click"

560       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

570         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

580         DoCmd.Hourglass False
590         strDocName = "frmRpt_CourtReports_NY_Input"
600         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm

610         DoCmd.Hourglass True
620         DoEvents

            ' ** Run function to fill Schedule A Asset List data.
630         If .CashAssets_Beg <> vbNullString Then

640           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm) ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

650           If intRetVal_BuildAssetListInfo = -2 Then
660             gcurCrtRpt_NY_IncomeBeg = 0@
670           Else
680             gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
690           End If

              ' ** Run function to empty and fill the tmpCourtReportData table.
700           blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
              ' ** No need to communicate this further, since this proc isn't called by 'PrintAll'.

710           If blnContinue = True Then

720             gstrAccountNo = .cmbAccounts.Column(0)

730             Select Case gblnUseReveuneExpenseCodes
                Case True
                  '##GTR_Ref: rptCourtRptNY_00A
740               PreviewOrPrint_NY "0A", strThisProc, acViewNormal, blnRebuildTable, frm ' ** Function: Above.
750             Case False
                  '##GTR_Ref: rptCourtRptNY_00
760               PreviewOrPrint_NY "0", strThisProc, acViewNormal, blnRebuildTable, frm ' ** Function: Above.
770             End Select

780             If .chkAssetList = True Then
790               DoEvents
800               AssetList_PreviewPrint_NY acViewNormal, strCallingForm, frm, strThisProc  ' ** Function: Above.
810             End If

820           End If  ' ** blnContinue.

830         End If  ' ** CashAssets_Beg.
840       End If  ' ** Validate.

850       DoCmd.Hourglass False

860     End With

EXITP:
870     Exit Sub

ERRH:
100     DoCmd.Hourglass False
110     Select Case ERR.Number
        Case Else
120       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
130     End Select
140     Resume EXITP

End Sub

Public Sub NY_CmdWord00_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Summary Statement.
' **
' ** NY_CmdWord00_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

900   On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord00_Click"

        Dim strRpt As String, strDocName As String
        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim lngE As Long

910     With frm

920       DoCmd.Hourglass True
930       DoEvents

940       blnContinue = True
950       blnUseSavedPath = False
960       blnExcel = False
970       blnAllCancel = False
980       .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
990       AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
1000      blnAutoStart = .chkOpenWord
1010      strThisProc = "cmdWord00_Click"

1020      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

1030        .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

1040        DoCmd.Hourglass False
1050        strDocName = "frmRpt_CourtReports_NY_Input"
1060        DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm

1070        DoCmd.Hourglass True
1080        DoEvents

            ' ** Run function to fill Schedule A Asset List data.
1090        If .CashAssets_Beg <> vbNullString Then

1100          intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm) ' ** Module Function: modCourtReportsNY.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

1110          If intRetVal_BuildAssetListInfo = -2 Then
1120            gcurCrtRpt_NY_IncomeBeg = 0@
1130          Else
1140            gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
1150          End If

              ' ** Run function to empty and fill the tmpCourtReportData table.
1160          blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
              ' ** blnAllCancel set below.
1170          DoEvents

1180          If blnContinue = True Then

1190            gstrAccountNo = .cmbAccounts.Column(0)

1200            gblnMessage = False
1210            Select Case gblnUseReveuneExpenseCodes
                Case False
1220              strRptName = "rptCourtRptNY_00"
1230            Case True
1240              strRptName = "rptCourtRptNY_00A"
1250            End Select

1260            .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

1270            strRptCap = vbNullString
1280            strRptCap = "CourtReport_NY_Summary_" & gstrAccountNo & "_" & _
                  Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

1290            If IsNull(.UserReportPath) = False Then
1300              If .UserReportPath <> vbNullString Then
1310                If .UserReportPath_chk = True Then
1320                  If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
1330                    blnUseSavedPath = True
1340                  End If
1350                End If
1360              End If
1370            End If

1380            Select Case blnUseSavedPath
                Case True
1390              strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
1400            Case False
1410              DoCmd.Hourglass False
1420              strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
1430            End Select

1440            If strRptPathFile <> vbNullString Then
1450              DoCmd.Hourglass True
1460              DoEvents
1470              If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
1480              If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
1490                Kill strRptPathFile
1500              End If
1510              Select Case gblnPrintAll
                  Case True
1520                lngFiles = lngFiles + 1&
1530                lngE = lngFiles - 1&
1540                ReDim Preserve arr_varFile(F_ELEMS, lngE)
1550                arr_varFile(F_RNAM, lngE) = strRptName
1560                arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
1570                arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
1580                FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
1590                DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
1600              Case False
1610                DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
1620              End Select
                  'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
1630              strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
1640              If strRptPath <> .UserReportPath Then
1650                .UserReportPath = strRptPath
1660                SetUserReportPath_NY frm  ' ** Procedure: Above.
1670              End If
1680            Else
1690              blnAllCancel = True
1700              .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
1710              AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
1720            End If  ' ** strRptPathFile.

1730            If blnAllCancel = False Then

1740              gstrAccountNo = .cmbAccounts.Column(0)

1750              If .chkAssetList = True Then
1760                DoEvents
1770                AssetList_Word_NY strCallingForm, frm  ' ** Procedure: Above.
1780              End If

1790            End If  ' ** blnAllCancel.

1800          Else
1810            blnAllCancel = True
1820            .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
1830            AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
1840          End If  ' ** blnContinue.

1850        End If  ' ** CashAssets_Beg.
1860      End If  ' ** Validate.

1870      DoCmd.Hourglass False

1880    End With

EXITP:
1890    Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel00_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Summary Statement.
' **
' ** NY_CmdExcel00_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel00_Click"

        Dim strRpt As String, strDocName As String
        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

1910    With frm

1920      DoCmd.Hourglass True
1930      DoEvents

1940      blnContinue = True
1950      blnUseSavedPath = False
1960      blnExcel = True
1970      blnAllCancel = False
1980      .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
1990      AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
2000      blnAutoStart = .chkOpenExcel
2010      strThisProc = "cmdExcel00_Click"

2020      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
2030        ForcePause 2  ' ** Module Function: modCodeUtilities.
2040        If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
2050          DoCmd.Hourglass False
2060          msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
2070          If msgResponse <> vbRetry Then
2080            blnAllCancel = True
2090            .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
2100            AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
2110            blnContinue = False
2120          End If
2130        End If
2140      End If

2150      If blnContinue = True Then

2160        DoCmd.Hourglass True
2170        DoEvents

2180        If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

2190          .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

2200          DoCmd.Hourglass False
2210          strDocName = "frmRpt_CourtReports_NY_Input"
2220          DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm

2230          DoCmd.Hourglass True
2240          DoEvents

              ' ** Run function to fill Schedule A Asset List data.
2250          If .CashAssets_Beg <> vbNullString Then

2260            intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm) ' ** Function: Above.
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.

2270            DoEvents

2280            If intRetVal_BuildAssetListInfo = -2 Then
2290              gcurCrtRpt_NY_IncomeBeg = 0@
2300            Else
2310              gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
2320            End If

2330            DoEvents

                ' ** Run function to empty and fill the tmpCourtReportData table.
2340            blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
                ' ** blnAllCancel set below.
2350            DoEvents

2360            If blnContinue = True Then

2370              gstrAccountNo = .cmbAccounts.Column(0)
2380              gdatStartDate = .DateEnd
2390              gdatEndDate = .DateStart
2400              gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
                  ' ** gstrCrtRpt_Ordinal and gstrCrtRpt_Version should be populated from the input window.

2410              gblnMessage = False: blnNoData = False
2420              Select Case gblnUseReveuneExpenseCodes
                  Case True
2430                strTmp01 = "rptCourtRptNY_00A"
2440                strQry = "qryCourtReport_NY_00A_X_17"
2450                varTmp00 = DCount("*", strQry)
2460                If IsNull(varTmp00) = True Then
2470                  blnNoData = True
2480                  strQry = "qryCourtReport_NY_00A_X_22"
2490                Else
2500                  If varTmp00 = 0 Then
2510                    blnNoData = True
2520                    strQry = "qryCourtReport_NY_00A_X_22"
2530                  End If
2540                End If
2550              Case False
2560                strTmp01 = "rptCourtRptNY_00"
2570                strQry = "qryCourtReport_NY_00_X_15"
2580                varTmp00 = DCount("*", strQry)
2590                If IsNull(varTmp00) = True Then
2600                  blnNoData = True
2610                  strQry = "qryCourtReport_NY_00_X_20"
2620                Else
2630                  If varTmp00 = 0 Then
2640                    blnNoData = True
2650                    strQry = "qryCourtReport_NY_00_X_20"
2660                  End If
2670                End If
2680              End Select

2690              .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

2700              strRptCap = vbNullString: strRptPathFile = vbNullString
2710              strRptPath = .UserReportPath
2720              strRptName = strTmp01
2730              DoEvents

2740              .CapArray_Load  ' ** Form Procedure: frmRpt_CourtReports_NY.
2750              DoEvents
2760              arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
2770              lngCaps = UBound(arr_varCap, 2) + 1&

2780              For lngX = 0& To (lngCaps - 1&)
2790                If arr_varCap(C_RNAM, lngX) = strRptName Then
2800                  strRptCap = arr_varCap(C_CAPN, lngX)
2810                  Exit For
2820                End If
2830              Next
2840              DoEvents

2850              If IsNull(.UserReportPath) = False Then
2860                If .UserReportPath <> vbNullString Then
2870                  If .UserReportPath_chk = True Then
2880                    If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
2890                      blnUseSavedPath = True
2900                    End If
2910                  End If
2920                End If
2930              End If

2940              strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
2950              If blnNoData = True Then
2960                strMacro = strMacro & "_nd"
2970              End If

2980              Select Case blnUseSavedPath
                  Case True
2990                strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
3000              Case False
3010                DoCmd.Hourglass False
3020                strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
3030              End Select

3040              If gblnPrintAll = False And .chkAssetList = True And .chkOpenExcel = True Then
                    ' ** The Summary won't let the Asset List export.
3050                blnAutoStart = False
3060              End If

3070              If strRptPathFile <> vbNullString Then
3080                DoCmd.Hourglass True
3090                DoEvents
3100                If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
3110                If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
3120                  Kill strRptPathFile
3130                End If
3140                If strQry <> vbNullString Then
                      ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                      ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
3150                  DoCmd.RunMacro strMacro
                      ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                      ' ** So, it's exported to 'CourtReport_NY_xxx.xls', which is then renamed.
3160                  DoEvents
3170                  If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                          FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
3180                    If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
3190                      Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
3200                    Else
3210                      Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
3220                    End If
3230                    DoEvents
3240                    If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
3250                      DoEvents
3260                      Select Case gblnPrintAll
                          Case True
3270                        lngFiles = lngFiles + 1&
3280                        lngE = lngFiles - 1&
3290                        ReDim Preserve arr_varFile(F_ELEMS, lngE)
3300                        arr_varFile(F_RNAM, lngE) = strRptName
3310                        arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
3320                        arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
3330                        FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
3340                      Case False
3350                        If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
3360                          EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
3370                        End If
3380                        DoEvents
3390                        If blnAutoStart = True Then
3400                          OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
3410                        End If
3420                      End Select
                          'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                          '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                          'End If
                          'DoEvents
                          'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
3430                    End If
3440                  End If
3450                Else
3460                  DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
3470                End If  ' ** strQry.
3480                strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
3490                If strRptPath <> .UserReportPath Then
3500                  .UserReportPath = strRptPath
3510                  SetUserReportPath_NY frm  ' ** Procedure: Above.
3520                End If

3530                gstrAccountNo = .cmbAccounts.Column(0)

3540                If .chkAssetList = True Then
3550                  DoEvents
3560                  AssetList_Excel_NY frm, strCallingForm, True  ' ** Module Procedure: modCourtReportsNY1.
3570                  If .chkOpenExcel = True And gblnPrintAll = False Then
                        ' ** Now open them.
3580                    OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
3590                    ForcePause 2  ' ** Module Function: modCodeUtilities.
                        ' ** Borrowing this from the Journal.
3600                    OpenExe gstrSaleAccountNumber  ' ** Module Function: modShellFuncs.
3610                  End If
3620                End If

3630              Else
3640                blnAllCancel = True
3650                .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
3660                AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
3670              End If  ' ** strRptPathFile.

3680            End If  ' ** blnContinue.

3690          End If  ' ** CashAssets_Beg.

3700        End If  ' ** Validate.

3710      End If ' ** blnContinue.

          ' ** Borrowing this from the Journal.
3720      gstrSaleAccountNumber = vbNullString

3730      DoCmd.Hourglass False

3740    End With

      #End If

EXITP:
3750    Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev01_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Received.
' **
' ** NY_CmdPrev01_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev01_Click"

        Dim strRpt As String, strDocName As String
        Dim intRetVal_BuildAssetListInfo As Integer

3810    With frm

3820      DoCmd.Hourglass True
3830      DoEvents

3840      strThisProc = "cmdPreview01_Click"

3850      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

3860        strRpt = vbNullString
3870        If gblnCrtRpt_NY_InvIncChange = False Then
3880          gstrCrtRpt_NY_InputTitle = "Invested Income"
3890          DoCmd.Hourglass False
3900          strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
3910          DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
3920          DoCmd.Hourglass True
3930          DoEvents
3940        End If

3950        .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

3960        If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then

              ' ** Run function to fill Schedule A Asset List data.
3970          intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

              ' ** Let it go through, regardless.  'VGC 03/06/2013.
              'Select Case intRetVal_BuildAssetListInfo
              'Case 0
              'IT MAY INDEED FIND NO DATA FOR THE PRIOR PERIOD!
              'Case -2
              '  Beep
              '  MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
              'Case -3, -9
              '  ' ** Message shown below.
              'End Select  ' ** intRetVal_BuildAssetListInfo

3980          PreviewOrPrint_NY "1", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

3990        End If  ' ** CashAssets_Beg.
4000      End If  ' ** Validate.

4010      DoCmd.Hourglass False

4020    End With

EXITP:
4030    Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint01_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Received.
' **
' ** NY_CmdPrint01_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint01_Click"

        Dim strRpt As String, strDocName As String
        Dim intRetVal_BuildAssetListInfo As Integer

4110    With frm

4120      DoCmd.Hourglass True
4130      DoEvents

4140      strThisProc = "cmdPrint01_Click"

4150      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

4160        strRpt = vbNullString

4170        If gblnCrtRpt_NY_InvIncChange = False Then
4180          gstrCrtRpt_NY_InputTitle = "Invested Income"
4190          DoCmd.Hourglass False
4200          strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
4210          DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
4220          DoCmd.Hourglass True
4230          DoEvents
4240        End If

4250        .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

4260        If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then

              ' ** Run function to fill Schedule A Asset List data.
4270          intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm) ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

              '##GTR_Ref: rptCourtRptNY_01
4280          PreviewOrPrint_NY "1", strThisProc, acViewNormal, blnRebuildTable, frm ' ** Function: Above.

4290        End If  ' ** CashAssets_Beg.
4300      End If  ' ** Validate.

4310      DoCmd.Hourglass False

4320    End With

EXITP:
4330    Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord01_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Received.
' **
' ** NY_CmdWord01_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord01_Click"

        Dim strRpt As String, strDocName As String
        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim lngE As Long

4410    With frm

4420      DoCmd.Hourglass True
4430      DoEvents

4440      blnUseSavedPath = False
4450      blnExcel = False
4460      blnAllCancel = False
4470      .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
4480      AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
4490      blnAutoStart = .chkOpenWord
4500      strThisProc = "cmdWord01_Click"

4510      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

4520        strRpt = vbNullString
4530        strRptName = "rptCourtRptNY_01"

4540        If gblnCrtRpt_NY_InvIncChange = False Then
4550          gstrCrtRpt_NY_InputTitle = "Invested Income"
4560          DoCmd.Hourglass False
4570          strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
4580          DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
4590          DoCmd.Hourglass True
4600          DoEvents
4610        End If

4620        .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

4630        If .CashAssets_Beg <> vbNullString Then

              ' ** Run function to fill Schedule A Asset List data.
4640          intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm) ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

4650          strRptCap = vbNullString
4660          strRptCap = "CourtReport_NY_Principal_Received_" & gstrAccountNo & "_" & _
                Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

4670          If IsNull(.UserReportPath) = False Then
4680            If .UserReportPath <> vbNullString Then
4690              If .UserReportPath_chk = True Then
4700                If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
4710                  blnUseSavedPath = True
4720                End If
4730              End If
4740            End If
4750          End If

4760          Select Case blnUseSavedPath
              Case True
4770            strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
4780          Case False
4790            DoCmd.Hourglass False
4800            strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
4810          End Select

4820          If strRptPathFile <> vbNullString Then
4830            DoCmd.Hourglass True
4840            DoEvents
4850            If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
4860            If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
4870              Kill strRptPathFile
4880            End If
4890            Select Case gblnPrintAll
                Case True
4900              lngFiles = lngFiles + 1&
4910              lngE = lngFiles - 1&
4920              ReDim Preserve arr_varFile(F_ELEMS, lngE)
4930              arr_varFile(F_RNAM, lngE) = strRptName
4940              arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
4950              arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
4960              FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
4970              DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
4980            Case False
4990              DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
5000            End Select
                'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
5010            strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
5020            If strRptPath <> .UserReportPath Then
5030              .UserReportPath = strRptPath
5040              SetUserReportPath_NY frm ' ** Function: Above.
5050            End If
5060          Else
5070            blnAllCancel = True
5080            .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
5090            AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
5100          End If  ' ** strRptPathFile.

5110        End If  ' ** CashAssets_Beg.
5120      End If  ' ** Validate.

5130      DoCmd.Hourglass False

5140    End With

EXITP:
5150    Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel01_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Received.
' **
' ** NY_CmdExcel01_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel01_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngCaps As Long, arr_varCap As Variant
        Dim strRpt As String, strDocName As String
        Dim strQry As String, strMacro As String
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

5210    With frm

5220      DoCmd.Hourglass True
5230      DoEvents

5240      blnContinue = True
5250      blnUseSavedPath = False
5260      blnExcel = True
5270      blnAllCancel = False
5280      .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
5290      AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
5300      blnAutoStart = .chkOpenExcel
5310      strThisProc = "cmdExcel01_Click"

5320      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
5330        ForcePause 2  ' ** Module Function: modCodeUtilities.
5340        If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
5350          DoCmd.Hourglass False
5360          msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
5370          If msgResponse <> vbRetry Then
5380            blnAllCancel = True
5390            .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
5400            AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
5410            blnContinue = False
5420          End If
5430        End If
5440      End If

5450      If blnContinue = True Then

5460        DoCmd.Hourglass True
5470        DoEvents

5480        If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

5490          strRpt = vbNullString
5500          strRptName = "rptCourtRptNY_01"

              ' ** Report uses gcurCrtRpt_NY_InputNew, from frmRpt_CourtReports_NY_Input_InvestedIncome form.
5510          If gblnCrtRpt_NY_InvIncChange = False Then
5520            gstrCrtRpt_NY_InputTitle = "Invested Income"
5530            DoCmd.Hourglass False
5540            strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
5550            DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
5560            DoCmd.Hourglass True
5570            DoEvents
5580          End If

5590          .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

5600          If .CashAssets_Beg <> vbNullString Then

                ' ** Run function to fill Schedule A Asset List data.
5610            intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm) ' ** Function: Above.
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.

5620            DoEvents

5630            gstrAccountNo = .cmbAccounts.Column(0)
5640            gdatStartDate = .DateEnd
5650            gdatEndDate = .DateStart
5660            gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
                ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

5670            gblnMessage = False
5680            strTmp01 = "rptCourtRptNY_01"
5690            strQry = "qryCourtReport_NY_01_X_41"
5700            varTmp00 = DCount("*", strQry)
5710            If IsNull(varTmp00) = True Then
5720              blnNoData = True
5730              strQry = "qryCourtReport_NY_01_X_57"
5740            Else
5750              If varTmp00 = 0 Then
5760                blnNoData = True
5770                strQry = "qryCourtReport_NY_01_X_57"
5780              End If
5790            End If

5800            .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

5810            strRptCap = vbNullString: strRptPathFile = vbNullString
5820            strRptPath = .UserReportPath
5830            strRptName = strTmp01
5840            DoEvents

5850            .CapArray_Load  ' ** Form Procedure: frmRpt_CourtReports_NY.
5860            DoEvents
5870            arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
5880            lngCaps = UBound(arr_varCap, 2) + 1&

5890            For lngX = 0& To (lngCaps - 1&)
5900              If arr_varCap(C_RNAM, lngX) = strRptName Then
5910                strRptCap = arr_varCap(C_CAPN, lngX)
5920                Exit For
5930              End If
5940            Next
5950            DoEvents

5960            If IsNull(.UserReportPath) = False Then
5970              If .UserReportPath <> vbNullString Then
5980                If .UserReportPath_chk = True Then
5990                  If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
6000                    blnUseSavedPath = True
6010                  End If
6020                End If
6030              End If
6040            End If

6050            strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
6060            If blnNoData = True Then
6070              strMacro = strMacro & "_nd"
6080            End If

6090            Select Case blnUseSavedPath
                Case True
6100              strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
6110            Case False
6120              DoCmd.Hourglass False
6130              strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
6140            End Select

6150            If strRptPathFile <> vbNullString Then

6160              DoCmd.Hourglass True
6170              DoEvents

6180              Set dbs = CurrentDb
6190              With dbs
                    ' ** Empty tblCourtReports_NY_PrinRec1.
6200                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_50")
6210                qdf.Execute
6220                Set qdf = Nothing
6230                DoEvents
                    ' ** Empty tblCourtReports_NY_PrinRec2.
6240                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_51")
6250                qdf.Execute
6260                Set qdf = Nothing
6270                DoEvents
                    ' ** Empty tblCourtReports_NY_PrinRec3.
6280                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_52")
6290                qdf.Execute
6300                Set qdf = Nothing
6310                DoEvents
                    ' ** Append qryCourtReport_NY_01_X_07 (xx) to tblCourtReports_NY_PrinRec1.
6320                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_47")
6330                qdf.Execute
6340                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_08 (xx) to tblCourtReports_NY_PrinRec1.
6350                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_48")
6360                qdf.Execute
6370                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_09 (xx) to tblCourtReports_NY_PrinRec1.
6380                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_49")
6390                qdf.Execute
6400                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_14 (xx) to tblCourtReports_NY_PrinRec2.
                    ' ** It complained that .._19 was too complex!
6410                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_42")
6420                qdf.Execute
6430                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_15 (xx) to tblCourtReports_NY_PrinRec2.
6440                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_43")
6450                qdf.Execute
6460                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_16 (xx) to tblCourtReports_NY_PrinRec2.
6470                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_44")
6480                qdf.Execute
6490                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_17 (xx) to tblCourtReports_NY_PrinRec2.
6500                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_45")
6510                qdf.Execute
6520                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_18 (xx) to tblCourtReports_NY_PrinRec2.
6530                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_46")
6540                qdf.Execute
6550                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_19 (xx) to tblCourtReports_NY_PrinRec3.
                    ' ** It complained that .._22 was too complex!
6560                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_33")
6570                qdf.Execute
6580                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_20 (xx) to tblCourtReports_NY_PrinRec3.
6590                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_34")
6600                qdf.Execute
6610                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_21 (xx) to tblCourtReports_NY_PrinRec3.
6620                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_35")
6630                qdf.Execute
6640                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_31 (xx) to tblCourtReports_NY_PrinRec3.
6650                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_36")
6660                qdf.Execute
6670                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_37 (xx) to tblCourtReports_NY_PrinRec3.
6680                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_38")
6690                qdf.Execute
6700                Set qdf = Nothing
                    ' ** Append qryCourtReport_NY_01_X_39 (xx) to tblCourtReports_NY_PrinRec3.
6710                Set qdf = .QueryDefs("qryCourtReport_NY_01_X_40")
6720                qdf.Execute
6730                Set qdf = Nothing
6740                .Close
6750              End With
6760              Set qdf = Nothing
6770              Set dbs = Nothing
6780              DoEvents

6790              If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
6800              If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
6810                Kill strRptPathFile
6820              End If
6830              If strQry <> vbNullString Then
                    ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                    ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
6840                DoCmd.RunMacro strMacro
                    ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                    ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
6850                DoEvents
6860                If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                        FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
6870                  If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
6880                    Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
6890                  Else
6900                    Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
6910                  End If
6920                  DoEvents
6930                  If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
6940                    DoEvents
6950                    Select Case gblnPrintAll
                        Case True
6960                      lngFiles = lngFiles + 1&
6970                      lngE = lngFiles - 1&
6980                      ReDim Preserve arr_varFile(F_ELEMS, lngE)
6990                      arr_varFile(F_RNAM, lngE) = strRptName
7000                      arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
7010                      arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7020                      FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
7030                    Case False
7040                      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
7050                        EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
7060                      End If
7070                      DoEvents
7080                      If blnAutoStart = True Then
7090                        OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
7100                      End If
7110                    End Select
                        'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                        '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                        'End If
                        'DoEvents
                        'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
7120                  End If
7130                End If
7140              Else
7150                DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
7160              End If  ' ** strQry.
7170              strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7180              If strRptPath <> .UserReportPath Then
7190                .UserReportPath = strRptPath
7200                .SetUserReportPath_NY frm  ' ** Form Procedure: frmRpt_CourtReports_NY.
7210              End If

7220            Else
7230              blnAllCancel = True
7240              .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
7250              AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
7260            End If  ' ** strRptPathFile.

7270          End If  ' ** CashAssets_Beg.
7280        End If  ' ** Validate().
7290      End If  ' ** blnContinue.

7300      DoCmd.Hourglass False

7310    End With

      #End If

EXITP:
7320    Set qdf = Nothing
7330    Set dbs = Nothing
7340    Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev02_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Increases on Sales, Liquidation or Distribution.
' **
' ** NY_CmdPrev02_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev02_Click"

7410    With frm

7420      DoCmd.Hourglass True
7430      DoEvents

7440      strThisProc = "cmdPreview02"

7450      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

7460        .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

7470        PreviewOrPrint_NY "2", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
            'DOES THIS_PROC NEED TO BE THE FORM'S? NO!

7480      End If  ' ** Validate.

7490      DoCmd.Hourglass False

7500    End With

EXITP:
7510    Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint02_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Increases on Sales, Liquidation or Distribution.
' **
' ** NY_CmdPrint02_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint02_Click"

7610    With frm

7620      DoCmd.Hourglass True
7630      DoEvents

7640      strThisProc = "cmdPrint02_Click"

7650      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.
            '##GTR_Ref: rptCourtRptNY_02
7660        PreviewOrPrint_NY "2", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
7670      End If

7680      DoCmd.Hourglass False

7690    End With

EXITP:
7700    Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord02_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Increases on Sales, Liquidation or Distribution.
' **
' ** NY_CmdWord02_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord02_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

7810    With frm

7820      DoCmd.Hourglass True
7830      DoEvents

7840      blnUseSavedPath = False
7850      blnExcel = False
7860      blnAllCancel = False
7870      .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
7880      AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
7890      blnAutoStart = .chkOpenWord
7900      strThisProc = "cmdWord02_Click"

7910      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

7920        strRptName = "rptCourtRptNY_02"
7930        .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

7940        strRptCap = vbNullString
7950        strRptCap = "CourtReport_NY_Increases_On_Sales_Liquidation_or_Distribution_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

7960        If IsNull(.UserReportPath) = False Then
7970          If .UserReportPath <> vbNullString Then
7980            If .UserReportPath_chk = True Then
7990              If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
8000                blnUseSavedPath = True
8010              End If
8020            End If
8030          End If
8040        End If

8050        Select Case blnUseSavedPath
            Case True
8060          strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
8070        Case False
8080          DoCmd.Hourglass False
8090          strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
8100        End Select

8110        If strRptPathFile <> vbNullString Then
8120          DoCmd.Hourglass True
8130          DoEvents
8140          If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
8150          If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
8160            Kill strRptPathFile
8170          End If
8180          Select Case gblnPrintAll
              Case True
8190            lngFiles = lngFiles + 1&
8200            lngE = lngFiles - 1&
8210            ReDim Preserve arr_varFile(F_ELEMS, lngE)
8220            arr_varFile(F_RNAM, lngE) = strRptName
8230            arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
8240            arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
8250            FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
8260            DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
8270          Case False
8280            DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
8290          End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
8300          strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
8310          If strRptPath <> .UserReportPath Then
8320            .UserReportPath = strRptPath
8330            SetUserReportPath_NY frm  ' ** Procedure: Above.
8340          End If
8350        Else
8360          blnAllCancel = True
8370          .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
8380          AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
8390        End If  ' ** strRptPathFile.

8400      End If  ' ** Validate.

8410      DoCmd.Hourglass False

8420    End With

EXITP:
8430    Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel02_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Increases on Sales, Liquidation or Distribution.
' **
' ** NY_CmdExcel02_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel02_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

8510    With frm

8520      DoCmd.Hourglass True
8530      DoEvents

8540      blnContinue = True
8550      blnUseSavedPath = False
8560      blnExcel = True
8570      blnAllCancel = False
8580      .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
8590      AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
8600      blnAutoStart = .chkOpenExcel
8610      strThisProc = "cmdExcel02_Click"

8620      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
8630        ForcePause 2  ' ** Module Function: modCodeUtilities.
8640        If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
8650          DoCmd.Hourglass False
8660          msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
8670          If msgResponse <> vbRetry Then
8680            blnAllCancel = True
8690            .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
8700            AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
8710            blnContinue = False
8720          End If
8730        End If
8740      End If

8750      If blnContinue = True Then

8760        DoCmd.Hourglass True
8770        DoEvents

8780        If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

8790          .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

8800          DoEvents

8810          gstrAccountNo = .cmbAccounts.Column(0)
8820          gdatStartDate = .DateEnd
8830          gdatEndDate = .DateStart
8840          gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

8850          gblnMessage = False
8860          strTmp01 = "rptCourtRptNY_02"
8870          strQry = "qryCourtReport_NY_02_X_09"
8880          varTmp00 = DCount("*", strQry)
8890          If IsNull(varTmp00) = True Then
8900            blnNoData = True
8910            strQry = "qryCourtReport_NY_02_X_14"
8920          Else
8930            If varTmp00 = 0 Then
8940              blnNoData = True
8950              strQry = "qryCourtReport_NY_02_X_14"
8960            End If
8970          End If

8980          .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

8990          strRptCap = vbNullString: strRptPathFile = vbNullString
9000          strRptPath = .UserReportPath
9010          strRptName = strTmp01
9020          DoEvents

9030          .CapArray_Load  ' ** Form Procedure: frmRpt_CourtReports_NY.
9040          DoEvents
9050          arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
9060          lngCaps = UBound(arr_varCap, 2) + 1&

9070          For lngX = 0& To (lngCaps - 1&)
9080            If arr_varCap(C_RNAM, lngX) = strRptName Then
9090              strRptCap = arr_varCap(C_CAPN, lngX)
9100              Exit For
9110            End If
9120          Next
9130          DoEvents

9140          If IsNull(.UserReportPath) = False Then
9150            If .UserReportPath <> vbNullString Then
9160              If .UserReportPath_chk = True Then
9170                If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
9180                  blnUseSavedPath = True
9190                End If
9200              End If
9210            End If
9220          End If

9230          strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
9240          If blnNoData = True Then
9250            strMacro = strMacro & "_nd"
9260          End If

9270          Select Case blnUseSavedPath
              Case True
9280            strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
9290          Case False
9300            DoCmd.Hourglass False
9310            strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
9320          End Select

9330          If strRptPathFile <> vbNullString Then
9340            DoCmd.Hourglass True
9350            DoEvents
9360            If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
9370            If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
9380              Kill strRptPathFile
9390            End If
9400            If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
9410              DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
9420              DoEvents
9430              If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
9440                If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
9450                  Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
9460                Else
9470                  Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
9480                End If
9490                DoEvents
9500                If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
9510                  DoEvents
9520                  Select Case gblnPrintAll
                      Case True
9530                    lngFiles = lngFiles + 1&
9540                    lngE = lngFiles - 1&
9550                    ReDim Preserve arr_varFile(F_ELEMS, lngE)
9560                    arr_varFile(F_RNAM, lngE) = strRptName
9570                    arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
9580                    arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
9590                    FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
9600                  Case False
9610                    If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
9620                      EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
9630                    End If
9640                    DoEvents
9650                    If blnAutoStart = True Then
9660                      OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
9670                    End If
9680                  End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
9690                End If
9700              End If
9710            Else
9720              DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
9730            End If  ' ** strQry.
9740            strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
9750            If strRptPath <> .UserReportPath Then
9760              .UserReportPath = strRptPath
9770              SetUserReportPath_NY frm  ' ** Procedure: Above.
9780            End If
9790          Else
9800            blnAllCancel = True
9810            .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
9820            AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
9830          End If  ' ** strRptPathFile.

9840        End If  ' ** Validate().
9850      End If ' ** blnContinue.

9860      DoCmd.Hourglass False

9870    End With

      #End If

EXITP:
9880    Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev03_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability.
' **
' ** NY_CmdPrev03_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev03_Click"

9910    With frm

9920      DoCmd.Hourglass True
9930      DoEvents

9940      strThisProc = "cmdPreview02_Click"

9950      If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

9960        .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

9970        PreviewOrPrint_NY "3", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

9980      End If  ' ** Validate.

9990      DoCmd.Hourglass False

10000   End With

EXITP:
10010   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint03_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability.
' **
' ** NY_CmdPrint03_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint03_Click"

10110   With frm

10120     DoCmd.Hourglass True
10130     DoEvents

10140     strThisProc = "cmdPrint02_Click"

10150     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

            '##GTR_Ref: rptCourtRptNY_03
10160       PreviewOrPrint_NY "3", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.

10170     End If

10180     DoCmd.Hourglass False

10190   End With

EXITP:
10200   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord03_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability.
' **
' ** NY_CmdWord03_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord03_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

10310   With frm

10320     DoCmd.Hourglass True
10330     DoEvents

10340     blnUseSavedPath = False
10350     blnExcel = False
10360     blnAllCancel = False
10370     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
10380     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
10390     blnAutoStart = .chkOpenWord
10400     strThisProc = "cmdWord03_Click"

10410     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

10420       strRptName = "rptCourtRptNY_03"
10430       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

10440       strRptCap = vbNullString
10450       strRptCap = "CourtReport_NY_Decreases_Due_to_Sales_Liquidation_Collection_Distribution_or_Uncollectability_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

10460       If IsNull(.UserReportPath) = False Then
10470         If .UserReportPath <> vbNullString Then
10480           If .UserReportPath_chk = True Then
10490             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
10500               blnUseSavedPath = True
10510             End If
10520           End If
10530         End If
10540       End If

10550       Select Case blnUseSavedPath
            Case True
10560         strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
10570       Case False
10580         DoCmd.Hourglass False
10590         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
10600       End Select

10610       If strRptPathFile <> vbNullString Then
10620         DoCmd.Hourglass True
10630         DoEvents
10640         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
10650         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
10660           Kill strRptPathFile
10670         End If
10680         Select Case gblnPrintAll
              Case True
10690           lngFiles = lngFiles + 1&
10700           lngE = lngFiles - 1&
10710           ReDim Preserve arr_varFile(F_ELEMS, lngE)
10720           arr_varFile(F_RNAM, lngE) = strRptName
10730           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
10740           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
10750           FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
10760           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
10770         Case False
10780           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
10790         End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
10800         strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
10810         If strRptPath <> .UserReportPath Then
10820           .UserReportPath = strRptPath
10830           SetUserReportPath_NY frm  ' ** Procedure: Above.
10840         End If
10850       Else
10860         blnAllCancel = True
10870         .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
10880         AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
10890       End If  ' ** strRptPathFile.

10900     End If  ' ** Validate.

10910     DoCmd.Hourglass False

10920   End With

EXITP:
10930   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel03_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability.
' **
' ** NY_CmdExcel03_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel03_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

11010   With frm

11020     DoCmd.Hourglass True
11030     DoEvents

11040     blnContinue = True
11050     blnUseSavedPath = False
11060     blnExcel = True
11070     blnAllCancel = False
11080     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
11090     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
11100     blnAutoStart = .chkOpenExcel
11110     strThisProc = "cmdExcel03_Click"

11120     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
11130       ForcePause 2  ' ** Module Function: modCodeUtilities.
11140       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
11150         DoCmd.Hourglass False
11160         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
11170         If msgResponse <> vbRetry Then
11180           blnAllCancel = True
11190           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
11200           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
11210           blnContinue = False
11220         End If
11230       End If
11240     End If

11250     If blnContinue = True Then

11260       DoCmd.Hourglass True
11270       DoEvents

11280       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

11290         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

11300         DoEvents

11310         gstrAccountNo = .cmbAccounts.Column(0)
11320         gdatStartDate = .DateEnd
11330         gdatEndDate = .DateStart
11340         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

11350         gblnMessage = False
11360         strTmp01 = "rptCourtRptNY_03"
11370         strQry = "qryCourtReport_NY_03_X_08"
11380         varTmp00 = DCount("*", strQry)
11390         If IsNull(varTmp00) = True Then
11400           blnNoData = True
11410           strQry = "qryCourtReport_NY_03_X_13"
11420         Else
11430           If varTmp00 = 0 Then
11440             blnNoData = True
11450             strQry = "qryCourtReport_NY_03_X_13"
11460           End If
11470         End If

11480         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

11490         strRptCap = vbNullString: strRptPathFile = vbNullString
11500         strRptPath = .UserReportPath
11510         strRptName = strTmp01
11520         DoEvents

11530         .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
11540         DoEvents
11550         arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
11560         lngCaps = UBound(arr_varCap, 2) + 1&

11570         For lngX = 0& To (lngCaps - 1&)
11580           If arr_varCap(C_RNAM, lngX) = strRptName Then
11590             strRptCap = arr_varCap(C_CAPN, lngX)
11600             Exit For
11610           End If
11620         Next
11630         DoEvents

11640         If IsNull(.UserReportPath) = False Then
11650           If .UserReportPath <> vbNullString Then
11660             If .UserReportPath_chk = True Then
11670               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
11680                 blnUseSavedPath = True
11690               End If
11700             End If
11710           End If
11720         End If

11730         strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
11740         If blnNoData = True Then
11750           strMacro = strMacro & "_nd"
11760         End If

11770         Select Case blnUseSavedPath
              Case True
11780           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
11790         Case False
11800           DoCmd.Hourglass False
11810           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
11820         End Select

11830         If strRptPathFile <> vbNullString Then
11840           DoCmd.Hourglass True
11850           DoEvents
11860           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
11870           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
11880             Kill strRptPathFile
11890           End If
11900           If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
11910             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
11920             DoEvents
11930             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
11940               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
11950                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
11960               Else
11970                 Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
11980               End If
11990               DoEvents
12000               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
12010                 DoEvents
12020                 Select Case gblnPrintAll
                      Case True
12030                   lngFiles = lngFiles + 1&
12040                   lngE = lngFiles - 1&
12050                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
12060                   arr_varFile(F_RNAM, lngE) = strRptName
12070                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
12080                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
12090                   FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
12100                 Case False
12110                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12120                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
12130                   End If
12140                   DoEvents
12150                   If blnAutoStart = True Then
12160                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
12170                   End If
12180                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
12190               End If
12200             End If
12210           Else
12220             DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
12230           End If  ' ** strQry.
12240           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
12250           If strRptPath <> .UserReportPath Then
12260             .UserReportPath = strRptPath
12270             SetUserReportPath_NY frm  ' ** Procedure: Above.
12280           End If
12290         Else
12300           blnAllCancel = True
12310           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
12320           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
12330         End If  ' ** strRptPathFile.

12340       End If  ' ** Validate().
12350     End If ' ** blnContinue.

12360     DoCmd.Hourglass False

12370   End With

      #End If

EXITP:
12380   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev04_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Principal.
' **
' ** NY_CmdPrev04_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev04_Click"

12410   With frm

12420     DoCmd.Hourglass True
12430     DoEvents

12440     strThisProc = "cmdPreview04_Click"

12450     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

12460       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

12470       Select Case gblnUseReveuneExpenseCodes
            Case True
12480         PreviewOrPrint_NY "4A", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
12490       Case False
12500         PreviewOrPrint_NY "4", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
12510       End Select

12520     End If  ' ** Validate.

12530     DoCmd.Hourglass False

12540   End With

EXITP:
12550   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint04_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Principal.
' **
' ** NY_CmdPrint04_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint04_Click"

12610   With frm

12620     DoCmd.Hourglass True
12630     DoEvents

12640     strThisProc = "cmdPrint04_Click"

12650     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

12660       Select Case gblnUseReveuneExpenseCodes
            Case True
              '##GTR_Ref: rptCourtRptNY_04A
12670         PreviewOrPrint_NY "4A", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
12680       Case False
              '##GTR_Ref: rptCourtRptNY_04
12690         PreviewOrPrint_NY "4", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
12700       End Select

12710     End If  ' ** Validate.

12720     DoCmd.Hourglass False

12730   End With

EXITP:
12740   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord04_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Principal.
' **
' ** NY_CmdWord04_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

12800 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord04_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

12810   With frm

12820     DoCmd.Hourglass True
12830     DoEvents

12840     blnUseSavedPath = False
12850     blnExcel = False
12860     blnAllCancel = False
12870     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
12880     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
12890     blnAutoStart = .chkOpenWord
12900     strThisProc = "cmdWord04_Click"

12910     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

12920       Select Case gblnUseReveuneExpenseCodes
            Case True
12930         strRptName = "rptCourtRptNY_04A"
12940       Case False
12950         strRptName = "rptCourtRptNY_04"
12960       End Select
12970       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

12980       strRptCap = vbNullString
12990       strRptCap = "CourtReport_NY_Administration_Expenses_Chargeable_to_Principal_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

13000       If IsNull(.UserReportPath) = False Then
13010         If .UserReportPath <> vbNullString Then
13020           If .UserReportPath_chk = True Then
13030             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
13040               blnUseSavedPath = True
13050             End If
13060           End If
13070         End If
13080       End If

13090       Select Case blnUseSavedPath
            Case True
13100         strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
13110       Case False
13120         DoCmd.Hourglass False
13130         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
13140       End Select

13150       If strRptPathFile <> vbNullString Then
13160         DoCmd.Hourglass True
13170         DoEvents
13180         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
13190         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
13200           Kill strRptPathFile
13210         End If
13220         Select Case gblnPrintAll
              Case True
13230           lngFiles = lngFiles + 1&
13240           lngE = lngFiles - 1&
13250           ReDim Preserve arr_varFile(F_ELEMS, lngE)
13260           arr_varFile(F_RNAM, lngE) = strRptName
13270           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
13280           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
13290           FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
13300           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
13310         Case False
13320           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
13330         End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
13340         strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
13350         If strRptPath <> .UserReportPath Then
13360           .UserReportPath = strRptPath
13370           SetUserReportPath_NY frm  ' ** Procedure: Above.
13380         End If
13390       Else
13400         blnAllCancel = True
13410         .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
13420         AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
13430       End If  ' ** strRptPathFile.

13440     End If  ' ** Validate.

13450     DoCmd.Hourglass False

13460   End With

EXITP:
13470   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel04_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Principal.
' **
' ** NY_CmdExcel04_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel04_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

13510   With frm

13520     DoCmd.Hourglass True
13530     DoEvents

13540     blnContinue = True
13550     blnUseSavedPath = False
13560     blnExcel = True
13570     blnAllCancel = False
13580     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
13590     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
13600     blnAutoStart = .chkOpenExcel
13610     strThisProc = "cmdExcel04_Click"

13620     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
13630       ForcePause 2  ' ** Module Function: modCodeUtilities.
13640       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13650         DoCmd.Hourglass False
13660         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
13670         If msgResponse <> vbRetry Then
13680           blnAllCancel = True
13690           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
13700           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
13710           blnContinue = False
13720         End If
13730       End If
13740     End If

13750     If blnContinue = True Then

13760       DoCmd.Hourglass True
13770       DoEvents

13780       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

13790         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

13800         DoEvents

13810         gstrAccountNo = .cmbAccounts.Column(0)
13820         gdatStartDate = .DateEnd
13830         gdatEndDate = .DateStart
13840         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

13850         gblnMessage = False: blnNoData = False
13860         Select Case gblnUseReveuneExpenseCodes
              Case True
13870           strTmp01 = "rptCourtRptNY_04A"
13880           strQry = "qryCourtReport_NY_04_X_20"
13890           varTmp00 = DCount("*", strQry)
13900           If IsNull(varTmp00) = True Then
13910             blnNoData = True
13920             strQry = "qryCourtReport_NY_04_X_30"
13930           Else
13940             If varTmp00 = 0 Then
13950               blnNoData = True
13960               strQry = "qryCourtReport_NY_04_X_30"
13970             End If
13980           End If
13990         Case False
14000           strTmp01 = "rptCourtRptNY_04"
14010           strQry = "qryCourtReport_NY_04_X_08"
14020           varTmp00 = DCount("*", strQry)
14030           If IsNull(varTmp00) = True Then
14040             blnNoData = True
14050             strQry = "qryCourtReport_NY_04_X_26"
14060           Else
14070             If varTmp00 = 0 Then
14080               blnNoData = True
14090               strQry = "qryCourtReport_NY_04_X_26"
14100             End If
14110           End If
14120         End Select

14130         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

14140         strRptCap = vbNullString: strRptPathFile = vbNullString
14150         strRptPath = .UserReportPath
14160         strRptName = strTmp01
14170         DoEvents

14180         .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
14190         DoEvents
14200         arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
14210         lngCaps = UBound(arr_varCap, 2) + 1&

14220         For lngX = 0& To (lngCaps - 1&)
14230           If arr_varCap(C_RNAM, lngX) = strRptName Then
14240             strRptCap = arr_varCap(C_CAPN, lngX)
14250             Exit For
14260           End If
14270         Next
14280         DoEvents

14290         If IsNull(.UserReportPath) = False Then
14300           If .UserReportPath <> vbNullString Then
14310             If .UserReportPath_chk = True Then
14320               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
14330                 blnUseSavedPath = True
14340               End If
14350             End If
14360           End If
14370         End If

14380         strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
14390         If blnNoData = True Then
14400           strMacro = strMacro & "_nd"
14410         End If

14420         Select Case blnUseSavedPath
              Case True
14430           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
14440         Case False
14450           DoCmd.Hourglass False
14460           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
14470         End Select

14480         If strRptPathFile <> vbNullString Then
14490           DoCmd.Hourglass True
14500           DoEvents
14510           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
14520           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
14530             Kill strRptPathFile
14540           End If
14550           If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
14560             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_NY_xxx.xls', which is then renamed.
14570             DoEvents
14580             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
14590               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
14600                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
14610               Else
14620                 Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
14630               End If
14640               DoEvents
14650               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
14660                 DoEvents
14670                 Select Case gblnPrintAll
                      Case True
14680                   lngFiles = lngFiles + 1&
14690                   lngE = lngFiles - 1&
14700                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
14710                   arr_varFile(F_RNAM, lngE) = strRptName
14720                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
14730                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
14740                   FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
14750                 Case False
14760                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
14770                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
14780                   End If
14790                   DoEvents
14800                   If blnAutoStart = True Then
14810                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
14820                   End If
14830                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
14840               End If
14850             End If
14860           Else
14870             DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
14880           End If  ' ** strQry.
14890           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
14900           If strRptPath <> .UserReportPath Then
14910             .UserReportPath = strRptPath
14920             SetUserReportPath_NY frm  ' ** Procedure: Above.
14930           End If
14940         Else
14950           blnAllCancel = True
14960           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
14970           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
14980         End If  ' ** strRptPathFile.

14990       End If  ' ** Validate().
15000     End If ' ** blnContinue.

15010     DoCmd.Hourglass False

15020   End With

      #End If

EXITP:
15030   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev05_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Principal.
' **
' ** NY_CmdPrev05_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev05_Click"

15110   With frm

15120     DoCmd.Hourglass True
15130     DoEvents

15140     strThisProc = "cmdPreview05_Click"

15150     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

15160       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

15170       PreviewOrPrint_NY "5", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

15180     End If  ' ** Validate.

15190     DoCmd.Hourglass False

15200   End With

EXITP:
15210   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint05_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Principal.
' **
' ** NY_CmdPrint05_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint05_Click"

15310   With frm

15320     DoCmd.Hourglass True
15330     DoEvents

15340     strThisProc = "cmdPrint05_Click"

15350     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

15360       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

            '##GTR_Ref: rptCourtRptNY_05
15370       PreviewOrPrint_NY "5", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.

15380     End If  ' ** Validate.

15390     DoCmd.Hourglass False

15400   End With

EXITP:
15410   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord05_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Principal.
' **
' ** NY_CmdWord05_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord05_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

15510   With frm

15520     DoCmd.Hourglass True
15530     DoEvents

15540     blnUseSavedPath = False
15550     blnExcel = False
15560     blnAllCancel = False
15570     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
15580     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
15590     blnAutoStart = .chkOpenWord
15600     strThisProc = "cmdWord05_Click"

15610     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

15620       strRptName = "rptCourtRptNY_05"
15630       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

15640       strRptCap = vbNullString
15650       strRptCap = "CourtReport_NY_Distributions_of_Principal_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

15660       If IsNull(.UserReportPath) = False Then
15670         If .UserReportPath <> vbNullString Then
15680           If .UserReportPath_chk = True Then
15690             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
15700               blnUseSavedPath = True
15710             End If
15720           End If
15730         End If
15740       End If

15750       Select Case blnUseSavedPath
            Case True
15760         strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
15770       Case False
15780         DoCmd.Hourglass False
15790         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
15800       End Select

15810       If strRptPathFile <> vbNullString Then
15820         DoCmd.Hourglass True
15830         DoEvents
15840         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
15850         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
15860           Kill strRptPathFile
15870         End If
15880         Select Case gblnPrintAll
              Case True
15890           lngFiles = lngFiles + 1&
15900           lngE = lngFiles - 1&
15910           ReDim Preserve arr_varFile(F_ELEMS, lngE)
15920           arr_varFile(F_RNAM, lngE) = strRptName
15930           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
15940           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
15950           FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
15960           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
15970         Case False
15980           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
15990         End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
16000         strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
16010         If strRptPath <> .UserReportPath Then
16020           .UserReportPath = strRptPath
16030           SetUserReportPath_NY frm  ' ** Procedure: Above.
16040         End If
16050       Else
16060         blnAllCancel = True
16070         .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
16080         AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
16090       End If  ' ** strRptPathFile.

16100     End If  ' ** Validate.

16110     DoCmd.Hourglass False

16120   End With

EXITP:
16130   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel05_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Principal.
' **
' ** NY_CmdExcel05_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel05_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

16210   With frm

16220     DoCmd.Hourglass True
16230     DoEvents

16240     blnContinue = True
16250     blnUseSavedPath = False
16260     blnExcel = True
16270     blnAllCancel = False
16280     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
16290     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
16300     blnAutoStart = .chkOpenExcel
16310     strThisProc = "cmdExcel05_Click"

16320     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
16330       ForcePause 2  ' ** Module Function: modCodeUtilities.
16340       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
16350         DoCmd.Hourglass False
16360         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
16370         If msgResponse <> vbRetry Then
16380           blnAllCancel = True
16390           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
16400           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
16410           blnContinue = False
16420         End If
16430       End If
16440     End If

16450     If blnContinue = True Then

16460       DoCmd.Hourglass True
16470       DoEvents

16480       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

16490         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

16500         DoEvents

16510         gstrAccountNo = .cmbAccounts.Column(0)
16520         gdatStartDate = .DateEnd
16530         gdatEndDate = .DateStart
16540         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

16550         gblnMessage = False
16560         strTmp01 = "rptCourtRptNY_05"
16570         strQry = "qryCourtReport_NY_05_X_08"
16580         varTmp00 = DCount("*", strQry)
16590         If IsNull(varTmp00) = True Then
16600           blnNoData = True
16610           strQry = "qryCourtReport_NY_05_X_13"
16620         Else
16630           If varTmp00 = 0 Then
16640             blnNoData = True
16650             strQry = "qryCourtReport_NY_05_X_13"
16660           End If
16670         End If

16680         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

16690         strRptCap = vbNullString: strRptPathFile = vbNullString
16700         strRptPath = .UserReportPath
16710         strRptName = strTmp01
16720         DoEvents

16730         .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
16740         DoEvents
16750         arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
16760         lngCaps = UBound(arr_varCap, 2) + 1&

16770         For lngX = 0& To (lngCaps - 1&)
16780           If arr_varCap(C_RNAM, lngX) = strRptName Then
16790             strRptCap = arr_varCap(C_CAPN, lngX)
16800             Exit For
16810           End If
16820         Next
16830         DoEvents

16840         If IsNull(.UserReportPath) = False Then
16850           If .UserReportPath <> vbNullString Then
16860             If .UserReportPath_chk = True Then
16870               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
16880                 blnUseSavedPath = True
16890               End If
16900             End If
16910           End If
16920         End If

16930         strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
16940         If blnNoData = True Then
16950           strMacro = strMacro & "_nd"
16960         End If

16970         Select Case blnUseSavedPath
              Case True
16980           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
16990         Case False
17000           DoCmd.Hourglass False
17010           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
17020         End Select

17030         If strRptPathFile <> vbNullString Then
17040           DoCmd.Hourglass True
17050           DoEvents
17060           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
17070           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
17080             Kill strRptPathFile
17090           End If
17100           If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
17110             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
17120             DoEvents
17130             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
17140               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
17150                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
17160               Else
17170                 Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
17180               End If
17190               DoEvents
17200               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
17210                 DoEvents
17220                 Select Case gblnPrintAll
                      Case True
17230                   lngFiles = lngFiles + 1&
17240                   lngE = lngFiles - 1&
17250                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
17260                   arr_varFile(F_RNAM, lngE) = strRptName
17270                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
17280                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
17290                   FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
17300                 Case False
17310                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
17320                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
17330                   End If
17340                   DoEvents
17350                   If blnAutoStart = True Then
17360                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
17370                   End If
17380                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
17390               End If
17400             End If
17410           Else
17420             DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
17430           End If  ' ** strQry.
17440           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
17450           If strRptPath <> .UserReportPath Then
17460             .UserReportPath = strRptPath
17470             SetUserReportPath_NY frm  ' ** Procedure: Above.
17480           End If
17490         Else
17500           blnAllCancel = True
17510           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
17520           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
17530         End If  ' ** strRptPathFile.

17540       End If  ' ** Validate().
17550     End If ' ** blnContinue.

17560     DoCmd.Hourglass False

17570   End With

      #End If

EXITP:
17580   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev06_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of New Investments, Exchanges and Stock Distributions of Principal Assets.
' **
' ** NY_CmdPrev06_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev06_Click"

17610   With frm

17620     DoCmd.Hourglass True
17630     DoEvents

17640     strThisProc = "cmdPreview06_Click"

17650     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

17660       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

17670       PreviewOrPrint_NY "6", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

17680     End If  ' ** Validate.

17690     DoCmd.Hourglass False

17700   End With

EXITP:
17710   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint06_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of New Investments, Exchanges and Stock Distributions of Principal Assets.
' **
' ** NY_CmdPrint06_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint06_Click"

17810   With frm

17820     DoCmd.Hourglass True
17830     DoEvents

17840     strThisProc = "cmdPrint06_Click"

17850     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

            '##GTR_Ref: rptCourtRptNY_06
17860       PreviewOrPrint_NY "6", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.

17870     End If

17880     DoCmd.Hourglass False

17890   End With

EXITP:
17900   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord06_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of New Investments, Exchanges and Stock Distributions of Principal Assets.
' **
' ** NY_CmdWord06_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord06_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

18010   With frm

18020     DoCmd.Hourglass True
18030     DoEvents

18040     blnUseSavedPath = False
18050     blnExcel = False
18060     blnAllCancel = False
18070     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
18080     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
18090     blnAutoStart = .chkOpenWord
18100     strThisProc = "cmdWord06_Click"

18110     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

18120       strRptName = "rptCourtRptNY_06"
18130       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

18140       strRptCap = vbNullString
18150       strRptCap = "CourtReport_NY_New_Investments_Exchanges_and_Stock_Distributions_of_Principal_Assets_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

18160       If IsNull(.UserReportPath) = False Then
18170         If .UserReportPath <> vbNullString Then
18180           If .UserReportPath_chk = True Then
18190             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
18200               blnUseSavedPath = True
18210             End If
18220           End If
18230         End If
18240       End If

18250       Select Case blnUseSavedPath
            Case True
18260         strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
18270       Case False
18280         DoCmd.Hourglass False
18290         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
18300       End Select

18310       If strRptPathFile <> vbNullString Then
18320         DoCmd.Hourglass True
18330         DoEvents
18340         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
18350         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
18360           Kill strRptPathFile
18370         End If
18380         Select Case gblnPrintAll
              Case True
18390           lngFiles = lngFiles + 1&
18400           lngE = lngFiles - 1&
18410           ReDim Preserve arr_varFile(F_ELEMS, lngE)
18420           arr_varFile(F_RNAM, lngE) = strRptName
18430           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
18440           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
18450           FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
18460           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
18470         Case False
18480           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
18490         End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
18500         strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
18510         If strRptPath <> .UserReportPath Then
18520           .UserReportPath = strRptPath
18530           SetUserReportPath_NY frm  ' ** Procedure: Above.
18540         End If
18550       Else
18560         blnAllCancel = True
18570         .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
18580         AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
18590       End If  ' ** strRptPathFile.

18600     End If  ' ** Validate.

18610     DoCmd.Hourglass False

18620   End With

EXITP:
18630   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel06_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of New Investments, Exchanges and Stock Distributions of Principal Assets.
' **
' ** NY_CmdExcel06_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel06_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

18710   With frm

18720     DoCmd.Hourglass True
18730     DoEvents

18740     blnContinue = True
18750     blnUseSavedPath = False
18760     blnExcel = True
18770     blnAllCancel = False
18780     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
18790     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
18800     blnAutoStart = .chkOpenExcel
18810     strThisProc = "cmdExcel06_Click"

18820     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
18830       ForcePause 2  ' ** Module Function: modCodeUtilities.
18840       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18850         DoCmd.Hourglass False
18860         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
18870         If msgResponse <> vbRetry Then
18880           blnAllCancel = True
18890           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
18900           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
18910           blnContinue = False
18920         End If
18930       End If
18940     End If

18950     If blnContinue = True Then

18960       DoCmd.Hourglass True
18970       DoEvents

18980       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

18990         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

19000         DoEvents

19010         gstrAccountNo = .cmbAccounts.Column(0)
19020         gdatStartDate = .DateEnd
19030         gdatEndDate = .DateStart
19040         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

19050         gblnMessage = False
19060         strTmp01 = "rptCourtRptNY_06"
19070         strQry = "qryCourtReport_NY_06_X_18"
19080         varTmp00 = DCount("*", strQry)
19090         If IsNull(varTmp00) = True Then
19100           blnNoData = True
19110           strQry = "qryCourtReport_NY_06_X_23"
19120         Else
19130           If varTmp00 = 0 Then
19140             blnNoData = True
19150             strQry = "qryCourtReport_NY_06_X_23"
19160           End If
19170         End If

19180         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

19190         strRptCap = vbNullString: strRptPathFile = vbNullString
19200         strRptPath = .UserReportPath
19210         strRptName = strTmp01
19220         DoEvents

19230         .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
19240         DoEvents
19250         arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
19260         lngCaps = UBound(arr_varCap, 2) + 1&

19270         For lngX = 0& To (lngCaps - 1&)
19280           If arr_varCap(C_RNAM, lngX) = strRptName Then
19290             strRptCap = arr_varCap(C_CAPN, lngX)
19300             Exit For
19310           End If
19320         Next
19330         DoEvents

19340         If IsNull(.UserReportPath) = False Then
19350           If .UserReportPath <> vbNullString Then
19360             If .UserReportPath_chk = True Then
19370               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
19380                 blnUseSavedPath = True
19390               End If
19400             End If
19410           End If
19420         End If

19430         strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
19440         If blnNoData = True Then
19450           strMacro = strMacro & "_nd"
19460         End If

19470         Select Case blnUseSavedPath
              Case True
19480           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
19490         Case False
19500           DoCmd.Hourglass False
19510           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
19520         End Select

19530         If strRptPathFile <> vbNullString Then
19540           DoCmd.Hourglass True
19550           DoEvents
19560           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
19570           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
19580             Kill strRptPathFile
19590           End If
19600           If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
19610             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
19620             DoEvents
19630             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
19640               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
19650                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
19660               Else
19670                 Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
19680               End If
19690               DoEvents
19700               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
19710                 DoEvents
19720                 Select Case gblnPrintAll
                      Case True
19730                   lngFiles = lngFiles + 1&
19740                   lngE = lngFiles - 1&
19750                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
19760                   arr_varFile(F_RNAM, lngE) = strRptName
19770                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
19780                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
19790                   FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
19800                 Case False
19810                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
19820                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
19830                   End If
19840                   DoEvents
19850                   If blnAutoStart = True Then
19860                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
19870                   End If
19880                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
19890               End If
19900             End If
19910           Else
19920             DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
19930           End If  ' ** strQry.
19940           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
19950           If strRptPath <> .UserReportPath Then
19960             .UserReportPath = strRptPath
19970             SetUserReportPath_NY frm  ' ** Procedure: Above.
19980           End If
19990         Else
20000           blnAllCancel = True
20010           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
20020           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
20030         End If  ' ** strRptPathFile.

20040       End If  ' ** Validate().
20050     End If ' ** blnContinue.

20060     DoCmd.Hourglass False

20070   End With

      #End If

EXITP:
20080   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev07_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Remaining on Hand.
' **
' ** NY_CmdPrev07_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev07_Click"

        Dim strRpt As String, strDocName As String
        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer

20110   With frm

20120     DoCmd.Hourglass True
20130     DoEvents

20140     blnContinue = True
20150     strThisProc = "cmdPreview07_Click"

20160     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

20170       strRpt = vbNullString
20180       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

20190       If gblnCrtRpt_NY_InvIncChange = False Then
20200         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.
20210         DoCmd.Hourglass False
20220         gstrCrtRpt_NY_InputTitle = "Invested Income"
20230         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
20240         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
20250         DoCmd.Hourglass True
20260         DoEvents
20270       End If

20280       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then

20290         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

20300         gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
20310         blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
              ' ** No need to communicate this further, since there is no 'PreviewAll'.

20320         If blnContinue = True Then

                ' ** Run function to fill Asset List data for Schedule F end of account period.
20330           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm) ' ** Function: Above.
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.

20340           PreviewOrPrint_NY "7", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

20350         End If  ' ** blnContinue.

20360       End If  ' ** CashAssets_Beg.
20370     End If  ' ** Validate.

20380     DoCmd.Hourglass False

20390   End With

EXITP:
20400   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint07_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Remaining on Hand.
' **
' ** NY_CmdPrint07_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint07_Click"

        Dim strRpt As String, strDocName As String
        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer

20510   With frm

20520     DoCmd.Hourglass True
20530     DoEvents

20540     blnContinue = True
20550     strThisProc = "cmdPrint07_Click"

20560     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

20570       strRpt = vbNullString
20580       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

20590       If gblnCrtRpt_NY_InvIncChange = False Then
20600         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.
20610         DoCmd.Hourglass False
20620         gstrCrtRpt_NY_InputTitle = "Invested Income"
20630         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
20640         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
20650         DoCmd.Hourglass True
20660         DoEvents
20670       End If

20680       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then
20690         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

20700         gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
20710         blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
              ' ** No need to communicate this further, since this proc isn't called by 'PrintAll'.

20720         If blnContinue = True Then

                ' ** Run function to fill Asset List data for Schedule F end of account period.
20730           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm)  ' ** Function: Above.
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.

                '##GTR_Ref: rptCourtRptNY_07
20740           PreviewOrPrint_NY "7", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.

20750         End If  ' ** blnContinue.

20760       End If  ' ** CashAssets_Beg.
20770     End If  ' ** Validate.

20780     DoCmd.Hourglass False

20790   End With

EXITP:
20800   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord07_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Remaining on Hand.
' **
' ** NY_CmdWord07_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord07_Click"

        Dim strRpt As String, strDocName As String
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim lngE As Long

20910   With frm

20920     DoCmd.Hourglass True
20930     DoEvents

20940     blnContinue = True
20950     blnUseSavedPath = False
20960     blnExcel = False
20970     blnAllCancel = False
20980     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
20990     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
21000     blnAutoStart = .chkOpenWord
21010     strThisProc = "cmdWord07_Click"

21020     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

21030       strRpt = vbNullString
21040       strRptName = "rptCourtRptNY_07"
21050       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

21060       If gblnCrtRpt_NY_InvIncChange = False Then
21070         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.
21080         DoCmd.Hourglass False
21090         gstrCrtRpt_NY_InputTitle = "Invested Income"
21100         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
21110         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
21120         DoCmd.Hourglass True
21130         DoEvents
21140       End If

21150       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then
21160         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

21170         gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
21180         blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
              ' ** blnAllCancel set below.

21190         If blnContinue = True Then

                ' ** Run function to fill Asset List data for Schedule F end of account period.
21200           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm)  ' ** Function: Above.
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.

21210         End If  ' ** blnContinue.

21220       End If  ' ** CashAssets_Beg.

21230       If blnContinue = True Then

21240         strRptCap = vbNullString
21250         strRptCap = "CourtReport_NY_Principal_Remaining_on_Hand_" & gstrAccountNo & "_" & _
                Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

21260         If IsNull(.UserReportPath) = False Then
21270           If .UserReportPath <> vbNullString Then
21280             If .UserReportPath_chk = True Then
21290               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
21300                 blnUseSavedPath = True
21310               End If
21320             End If
21330           End If
21340         End If

21350         Select Case blnUseSavedPath
              Case True
21360           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
21370         Case False
21380           DoCmd.Hourglass False
21390           strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
21400         End Select

21410         If strRptPathFile <> vbNullString Then
21420           DoCmd.Hourglass True
21430           DoEvents
21440           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
21450           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
21460             Kill strRptPathFile
21470           End If
21480           Select Case gblnPrintAll
                Case True
21490             lngFiles = lngFiles + 1&
21500             lngE = lngFiles - 1&
21510             ReDim Preserve arr_varFile(F_ELEMS, lngE)
21520             arr_varFile(F_RNAM, lngE) = strRptName
21530             arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
21540             arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
21550             FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
21560             DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
21570           Case False
21580             DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
21590           End Select
                'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
21600           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
21610           If strRptPath <> .UserReportPath Then
21620             .UserReportPath = strRptPath
21630             SetUserReportPath_NY frm  ' ** Procedure: Above.
21640           End If
21650         Else
21660           blnAllCancel = True
21670           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
21680           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
21690         End If  ' ** strRptPathFile.

21700       End If  ' ** blnContinue.
21710     End If  ' ** Validate.

21720     DoCmd.Hourglass False

21730   End With

EXITP:
21740   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel07_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Principal Remaining on Hand.
' **
' ** NY_CmdExcel07_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel07_Click"

        Dim strRpt As String, strDocName As String
        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

21810   With frm

21820     DoCmd.Hourglass True
21830     DoEvents

21840     blnContinue = True
21850     blnUseSavedPath = False
21860     blnExcel = True
21870     blnAllCancel = False
21880     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
21890     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
21900     blnAutoStart = .chkOpenExcel
21910     strThisProc = "cmdExcel07_Click"

21920     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
21930       ForcePause 2  ' ** Module Function: modCodeUtilities.
21940       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
21950         DoCmd.Hourglass False
21960         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
21970         If msgResponse <> vbRetry Then
21980           blnAllCancel = True
21990           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
22000           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
22010           blnContinue = False
22020         End If
22030       End If
22040     End If

22050     If blnContinue = True Then

22060       DoCmd.Hourglass True
22070       DoEvents

22080       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

22090         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

22100         DoEvents

22110         gstrAccountNo = .cmbAccounts.Column(0)
22120         gdatStartDate = .DateEnd
22130         gdatEndDate = .DateStart
22140         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

22150         gblnMessage = False
22160         strTmp01 = "rptCourtRptNY_07"
22170         strQry = "qryCourtReport_NY_07_X_30"
22180         varTmp00 = DCount("*", strQry)
22190         If IsNull(varTmp00) = True Then
22200           blnNoData = True
22210           strQry = "qryCourtReport_NY_07_X_35"
22220         Else
22230           If varTmp00 = 0 Then
22240             blnNoData = True
22250             strQry = "qryCourtReport_NY_07_X_35"
22260           End If
22270         End If

22280         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

22290         strRptCap = vbNullString: strRptPathFile = vbNullString
22300         strRptPath = .UserReportPath
22310         strRptName = strTmp01
22320         DoEvents

22330         If gblnCrtRpt_NY_InvIncChange = False Then
22340           .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.
22350           DoCmd.Hourglass False
22360           gstrCrtRpt_NY_InputTitle = "Invested Income"
                'THEY NEED TO BE ABLE TO CANCEL THIS!
22370           strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
22380           DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
22390           DoCmd.Hourglass True
22400           DoEvents
22410         End If  ' ** gblnCrtRpt_NY_InvIncChange.

22420         If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then
22430           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.

22440           DoCmd.Hourglass True
22450           DoEvents

22460           gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
22470           blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
                ' ** blnAllCancel set below.

22480           If blnContinue = True Then

22490             DoCmd.Hourglass True
22500             DoEvents
                  ' ** Run function to fill Asset List data for Schedule F end of account period.
22510             intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm)  ' ** Function: Above.
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -3  Missing entry, e.g., date.
                  ' **   -9  Error.

22520             DoEvents

22530             .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
22540             DoEvents
22550             arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
22560             lngCaps = UBound(arr_varCap, 2) + 1&

22570             For lngX = 0& To (lngCaps - 1&)
22580               If arr_varCap(C_RNAM, lngX) = strRptName Then
22590                 strRptCap = arr_varCap(C_CAPN, lngX)
22600                 Exit For
22610               End If
22620             Next
22630             DoEvents

22640             If IsNull(.UserReportPath) = False Then
22650               If .UserReportPath <> vbNullString Then
22660                 If .UserReportPath_chk = True Then
22670                   If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
22680                     blnUseSavedPath = True
22690                   End If
22700                 End If
22710               End If
22720             End If

22730             strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
22740             If blnNoData = True Then
22750               strMacro = strMacro & "_nd"
22760             End If

22770             Select Case blnUseSavedPath
                  Case True
22780               strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
22790             Case False
22800               DoCmd.Hourglass False
22810               strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
22820             End Select

22830             If strRptPathFile <> vbNullString Then
22840               DoCmd.Hourglass True
22850               DoEvents
22860               If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
22870               If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
22880                 Kill strRptPathFile
22890               End If
22900               If strQry <> vbNullString Then
                      ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                      ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
22910                 DoCmd.RunMacro strMacro
                      ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                      ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
22920                 DoEvents
22930                 If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                          FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
22940                   If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
22950                     Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
22960                   Else
22970                     Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
22980                   End If
22990                   DoEvents
23000                   If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
23010                     DoEvents
23020                     Select Case gblnPrintAll
                          Case True
23030                       lngFiles = lngFiles + 1&
23040                       lngE = lngFiles - 1&
23050                       ReDim Preserve arr_varFile(F_ELEMS, lngE)
23060                       arr_varFile(F_RNAM, lngE) = strRptName
23070                       arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
23080                       arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23090                       FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
23100                     Case False
23110                       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
23120                         EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
23130                       End If
23140                       DoEvents
23150                       If blnAutoStart = True Then
23160                         OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
23170                       End If
23180                     End Select
                          'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                          '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                          'End If
                          'DoEvents
                          'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
23190                   End If
23200                 End If
23210               Else
23220                 DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
23230               End If  ' ** strQry.
23240               strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23250               If strRptPath <> .UserReportPath Then
23260                 .UserReportPath = strRptPath
23270                 SetUserReportPath_NY frm  ' ** Procedure: Above.
23280               End If
23290             Else
23300               blnAllCancel = True
23310               .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
23320               AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
23330             End If  ' ** strRptPathFile.

23340           Else
23350             blnAllCancel = True
23360             .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
23370             AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
23380           End If  ' ** blnContinue.

23390         End If  ' ** CashAssets_Beg.

23400       End If  ' ** Validate().
23410     End If  ' ** blnContinue.

23420     DoCmd.Hourglass False

23430   End With

      #End If

EXITP:
23440   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev08_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Received.
' **
' ** NY_CmdPrev08_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev08_Click"

        Dim strRpt As String, strDocName As String
        Dim intRetVal_BuildAssetListInfo As Integer

23510   With frm

23520     DoCmd.Hourglass True
23530     DoEvents

23540     strThisProc = "cmdPreview08_Click"

23550     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

23560       strRpt = vbNullString
23570       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

23580       intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry, e.g., date.
            ' **   -9  Error.

            ' ** Let it go through, regardless.  'VGC 03/06/2013.
            'Select Case intRetVal_BuildAssetListInfo
            'Case 0
            'Case -2
            '  Beep
            '  MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
            'Case -3, -9
            '  ' ** Message shown below.
            'End Select  ' ** intRetVal_BuildAssetListInfo

23590       If gblnCrtRpt_NY_InvIncChange = False Then
23600         DoCmd.Hourglass False
23610         gstrCrtRpt_NY_InputTitle = "Invested Income"
23620         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
23630         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
23640         DoCmd.Hourglass True
23650         DoEvents
23660       End If

23670       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then
23680         PreviewOrPrint_NY "8", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
23690       End If  ' ** CashAssets_Beg.

            ' ** Income On Hand:  .tAmount
            ' **   =Nz(FormRef('IncomeCash'),0)
            ' ** Invested Income: .tInvestedIncome
            ' **   =Nz(FormRef('NewInput'),0)
            ' **    NOW: =(Nz(FormRef('NewInput'),0)+Nz(FormRef('IncomeOnHand'),0))
            ' ** Total:           .tsumAmount
            ' **   =(Nz(FormRef('IncomeCash'),0)+Nz(FormRef('NewInput'),0))
            ' **    NOW: =((Nz(FormRef('IncomeCash'),0)+Nz(FormRef('NewInput'),0))+Nz(FormRef('IncomeOnHand'),0))

23700     End If  ' ** Validate.

23710     DoCmd.Hourglass False

23720   End With

EXITP:
23730   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint08_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Received.
' **
' ** NY_CmdPrint08_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint08_Click"

        Dim strRpt As String, strDocName As String
        Dim intRetVal_BuildAssetListInfo As Integer

23810   With frm

23820     DoCmd.Hourglass True
23830     DoEvents

23840     strThisProc = "cmdPrint08_Click"

23850     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

23860       strRpt = vbNullString
23870       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

23880       intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry, e.g., date.
            ' **   -9  Error.

            'Select Case intRetVal_BuildAssetListInfo
            'Case 0
            'Case -2
            '  Beep
            '  MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
            'Case -3, -9
            '  ' ** Message shown below.
            'End Select  ' ** intRetVal_BuildAssetListInfo

23890       If gblnCrtRpt_NY_InvIncChange = False Then
23900         gstrCrtRpt_NY_InputTitle = "Invested Income"
23910         DoCmd.Hourglass False
23920         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
23930         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
23940         DoCmd.Hourglass True
23950         DoEvents
23960       End If

23970       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then
              '##GTR_Ref: rptCourtRptNY_08
23980         PreviewOrPrint_NY "8", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
23990       End If  ' ** CashAssets_Beg.

24000     End If  ' ** Validate.

24010     DoCmd.Hourglass False

24020   End With

EXITP:
24030   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord08_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Received.
' **
' ** NY_CmdWord08_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord08_Click"

        Dim strRpt As String, strDocName As String
        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim lngE As Long

24110   With frm

24120     DoCmd.Hourglass True
24130     DoEvents

24140     blnUseSavedPath = False
24150     blnExcel = False
24160     blnAllCancel = False
24170     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
24180     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
24190     blnAutoStart = .chkOpenWord
24200     strThisProc = "cmdWord08_Click"

24210     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

24220       strRpt = vbNullString
24230       strRptName = "rptCourtRptNY_08"
24240       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

24250       intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Function: Above.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -3  Missing entry, e.g., date.
            ' **   -9  Error.

24260       If gblnCrtRpt_NY_InvIncChange = False Then
24270         gstrCrtRpt_NY_InputTitle = "Invested Income"
24280         DoCmd.Hourglass False
24290         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
24300         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
24310         DoCmd.Hourglass True
24320         DoEvents
24330       End If

24340       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then

24350         strRptCap = vbNullString
24360         strRptCap = "CourtReport_NY_Income_Received_" & gstrAccountNo & "_" & _
                Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

24370         If IsNull(.UserReportPath) = False Then
24380           If .UserReportPath <> vbNullString Then
24390             If .UserReportPath_chk = True Then
24400               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
24410                 blnUseSavedPath = True
24420               End If
24430             End If
24440           End If
24450         End If

24460         Select Case blnUseSavedPath
              Case True
24470           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
24480         Case False
24490           DoCmd.Hourglass False
24500           strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
24510         End Select

24520         If strRptPathFile <> vbNullString Then
24530           DoCmd.Hourglass True
24540           DoEvents
24550           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
24560           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
24570             Kill strRptPathFile
24580           End If
24590           Select Case gblnPrintAll
                Case True
24600             lngFiles = lngFiles + 1&
24610             lngE = lngFiles - 1&
24620             ReDim Preserve arr_varFile(F_ELEMS, lngE)
24630             arr_varFile(F_RNAM, lngE) = strRptName
24640             arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
24650             arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24660             FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
24670             DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
24680           Case False
24690             DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
24700           End Select
                'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
24710           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
24720           If strRptPath <> .UserReportPath Then
24730             .UserReportPath = strRptPath
24740             SetUserReportPath_NY frm  ' ** Procedure: Above.
24750           End If
24760         Else
24770           blnAllCancel = True
24780           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
24790           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
24800         End If  ' ** strRptPathFile.

24810       End If  ' ** CashAssets_Beg.

24820     End If  ' ** Validate.

24830     DoCmd.Hourglass False

24840   End With

EXITP:
24850   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel08_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Received.
' **
' ** NY_CmdExcel08_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel08_Click"

        Dim strRpt As String, strDocName As String
        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

24910   With frm

24920     DoCmd.Hourglass True
24930     DoEvents

24940     blnContinue = True
24950     blnUseSavedPath = False
24960     blnExcel = True
24970     blnAllCancel = False
24980     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
24990     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
25000     blnAutoStart = .chkOpenExcel
25010     strThisProc = "cmdExcel08_Click"

25020     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
25030       ForcePause 2  ' ** Module Function: modCodeUtilities.
25040       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
25050         DoCmd.Hourglass False
25060         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
25070         If msgResponse <> vbRetry Then
25080           blnAllCancel = True
25090           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
25100           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
25110           blnContinue = False
25120         End If
25130       End If
25140     End If

25150     If blnContinue = True Then

25160       DoCmd.Hourglass True
25170       DoEvents

25180       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

25190         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

25200         DoEvents

25210         gstrAccountNo = .cmbAccounts.Column(0)
25220         gdatStartDate = .DateEnd
25230         gdatEndDate = .DateStart
25240         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
25250         gcurCrtRpt_NY_ICash = Nz(DLookup("[icash]", "qryCourtReport_NY_00_B_01"), 0)  'THIS LOOKS ODD!
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

25260         gblnMessage = False
25270         strTmp01 = "rptCourtRptNY_08"
25280         strQry = "qryCourtReport_NY_08_X_09"
25290         varTmp00 = DCount("*", strQry)
25300         If IsNull(varTmp00) = True Then
25310           blnNoData = True
25320           strQry = "qryCourtReport_NY_08_X_14"
25330         Else
25340           If varTmp00 = 0 Then
25350             blnNoData = True
25360             strQry = "qryCourtReport_NY_08_X_14"
25370           End If
25380         End If

25390         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

25400         strRptCap = vbNullString: strRptPathFile = vbNullString
25410         strRptPath = .UserReportPath
25420         strRptName = strTmp01
25430         DoEvents

25440         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, strThisProc, frm)  ' ** Funtion: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

25450         DoEvents

25460         If gblnCrtRpt_NY_InvIncChange = False Then
25470           gstrCrtRpt_NY_InputTitle = "Invested Income"
25480           DoCmd.Hourglass False
25490           strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
25500           DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
                'THEY NEED TO BE ABLE TO CANCEL THIS!
25510           DoCmd.Hourglass True
25520           DoEvents
25530         End If

25540         If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then

25550           gcurCrtRpt_NY_ICash = Nz(DLookup("[icash]", "qryCourtReport_NY_00_B_01"), 0)

25560           .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
25570           DoEvents
25580           arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
25590           lngCaps = UBound(arr_varCap, 2) + 1&

25600           For lngX = 0& To (lngCaps - 1&)
25610             If arr_varCap(C_RNAM, lngX) = strRptName Then
25620               strRptCap = arr_varCap(C_CAPN, lngX)
25630               Exit For
25640             End If
25650           Next
25660           DoEvents

25670           If IsNull(.UserReportPath) = False Then
25680             If .UserReportPath <> vbNullString Then
25690               If .UserReportPath_chk = True Then
25700                 If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
25710                   blnUseSavedPath = True
25720                 End If
25730               End If
25740             End If
25750           End If

25760           strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
25770           If blnNoData = True Then
25780             strMacro = strMacro & "_nd"
25790           End If

25800           Select Case blnUseSavedPath
                Case True
25810             strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
25820           Case False
25830             DoCmd.Hourglass False
25840             strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
25850           End Select

25860           If strRptPathFile <> vbNullString Then
25870             DoCmd.Hourglass True  ' ** The hourglass doesn't seem to want to come on!
25880             DoEvents
25890             If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
25900             If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
25910               Kill strRptPathFile
25920             End If
25930             DoCmd.Hourglass True
25940             DoEvents
25950             If strQry <> vbNullString Then
                    ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                    ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
25960               DoCmd.RunMacro strMacro
25970               DoCmd.Hourglass True
25980               DoEvents
                    ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                    ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
25990               DoEvents
26000               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                        FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
26010                 If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
26020                   Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
26030                 Else
26040                   Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
26050                 End If
26060                 DoEvents
26070                 If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
26080                   DoEvents
26090                   Select Case gblnPrintAll
                        Case True
26100                     lngFiles = lngFiles + 1&
26110                     lngE = lngFiles - 1&
26120                     ReDim Preserve arr_varFile(F_ELEMS, lngE)
26130                     arr_varFile(F_RNAM, lngE) = strRptName
26140                     arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
26150                     arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
26160                     FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
26170                   Case False
26180                     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
26190                       EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
26200                     End If
26210                     DoEvents
26220                     If blnAutoStart = True Then
26230                       OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
26240                     End If
26250                   End Select
                        'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                        '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                        'End If
                        'DoEvents
                        'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
26260                 End If
26270               End If
26280             Else
26290               DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
26300             End If  ' ** strQry.
26310             strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
26320             If strRptPath <> .UserReportPath Then
26330               .UserReportPath = strRptPath
26340               SetUserReportPath_NY frm  ' ** Procedure: Above.
26350             End If
26360           Else
26370             blnAllCancel = True
26380             .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
26390             AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
26400           End If  ' ** strRptPathFile.

26410         End If  ' ** CashAssets_Beg.

26420       End If  ' ** Validate().
26430     End If ' ** blnContinue.

26440     DoCmd.Hourglass False

26450   End With

      #End If

EXITP:
26460   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev09_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of All Income Collected.
' **
' ** NY_CmdPrev09_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev09_Click"

26510   With frm

26520     DoCmd.Hourglass True
26530     DoEvents

26540     strThisProc = "cmdPreview09_Click"

26550     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

26560       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

26570       Select Case gblnUseReveuneExpenseCodes
            Case True
26580         PreviewOrPrint_NY "9A", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
26590       Case False
26600         PreviewOrPrint_NY "9", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.
26610       End Select

26620     End If  ' ** Validate.

26630     DoCmd.Hourglass False

26640   End With

EXITP:
26650   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint09_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of All Income Collected.
' **
' ** NY_CmdPrint09_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint09_Click"

26710   With frm

26720     DoCmd.Hourglass True
26730     DoEvents

26740     strThisProc = "cmdPrint09_Click"

26750     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

26760       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

26770       Select Case gblnUseReveuneExpenseCodes
            Case True
              '##GTR_Ref: rptCourtRptNY_09A
26780         PreviewOrPrint_NY "9A", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
26790       Case False
              '##GTR_Ref: rptCourtRptNY_09
26800         PreviewOrPrint_NY "9", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
26810       End Select

26820     End If  ' ** Validate.

26830     DoCmd.Hourglass False

26840   End With

EXITP:
26850   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord09_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of All Income Collected.
' **
' ** NY_CmdWord09_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord09_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

26910   With frm

26920     DoCmd.Hourglass True
26930     DoEvents

26940     blnUseSavedPath = False
26950     blnExcel = False
26960     blnAllCancel = False
26970     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
26980     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
26990     blnAutoStart = .chkOpenWord
27000     strThisProc = "cmdWord09_Click"

27010     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

27020       Select Case gblnUseReveuneExpenseCodes
            Case True
27030         strRptName = "rptCourtRptNY_09A"
27040       Case False
27050         strRptName = "rptCourtRptNY_09"
27060       End Select
27070       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

27080       strRptCap = vbNullString
27090       strRptCap = "CourtReport_NY_All_Income_Collected_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

27100       If IsNull(.UserReportPath) = False Then
27110         If .UserReportPath <> vbNullString Then
27120           If .UserReportPath_chk = True Then
27130             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
27140               blnUseSavedPath = True
27150             End If
27160           End If
27170         End If
27180       End If

27190       Select Case blnUseSavedPath
            Case True
27200         strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
27210       Case False
27220         DoCmd.Hourglass False
27230         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
27240       End Select

27250       If strRptPathFile <> vbNullString Then
27260         DoCmd.Hourglass True
27270         DoEvents
27280         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
27290         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
27300           Kill strRptPathFile
27310         End If
27320         Select Case gblnPrintAll
              Case True
27330           lngFiles = lngFiles + 1&
27340           lngE = lngFiles - 1&
27350           ReDim Preserve arr_varFile(F_ELEMS, lngE)
27360           arr_varFile(F_RNAM, lngE) = strRptName
27370           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
27380           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
27390           FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
27400           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
27410         Case False
27420           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
27430         End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
27440         strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
27450         If strRptPath <> .UserReportPath Then
27460           .UserReportPath = strRptPath
27470           SetUserReportPath_NY frm  ' ** Procedure: Above.
27480         End If
27490       Else
27500         blnAllCancel = True
27510         .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
27520         AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
27530       End If  ' ** strRptPathFile.

27540     End If  ' ** Validate.

27550     DoCmd.Hourglass False

27560   End With

EXITP:
27570   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel09_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of All Income Collected.
' **
' ** NY_CmdExcel09_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel09_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

27610   With frm

27620     DoCmd.Hourglass True
27630     DoEvents

27640     blnContinue = True
27650     blnUseSavedPath = False
27660     blnExcel = True
27670     blnAllCancel = False
27680     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
27690     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
27700     blnAutoStart = .chkOpenExcel
27710     strThisProc = "cmdExcel09_Click"

27720     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
27730       ForcePause 2  ' ** Module Function: modCodeUtilities.
27740       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
27750         DoCmd.Hourglass False
27760         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
27770         If msgResponse <> vbRetry Then
27780           blnAllCancel = True
27790           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
27800           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
27810           blnContinue = False
27820         End If
27830       End If
27840     End If

27850     If blnContinue = True Then

27860       DoCmd.Hourglass True
27870       DoEvents

27880       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

27890         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

27900         DoEvents

27910         gstrAccountNo = .cmbAccounts.Column(0)
27920         gdatStartDate = .DateEnd
27930         gdatEndDate = .DateStart
27940         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

27950         gblnMessage = False
27960         Select Case gblnUseReveuneExpenseCodes
              Case True
27970           strTmp01 = "rptCourtRptNY_09A"
27980           strQry = "qryCourtReport_NY_09_X_24"
27990           varTmp00 = DCount("*", strQry)
28000           If IsNull(varTmp00) = True Then
28010             blnNoData = True
28020             strQry = "qryCourtReport_NY_09_X_34"
28030           Else
28040             If varTmp00 = 0 Then
28050               blnNoData = True
28060               strQry = "qryCourtReport_NY_09_X_34"
28070             End If
28080           End If
28090         Case False
28100           strTmp01 = "rptCourtRptNY_09"
28110           strQry = "qryCourtReport_NY_09_X_11"
28120           varTmp00 = DCount("*", strQry)
28130           If IsNull(varTmp00) = True Then
28140             blnNoData = True
28150             strQry = "qryCourtReport_NY_09_X_29"
28160           Else
28170             If varTmp00 = 0 Then
28180               blnNoData = True
28190               strQry = "qryCourtReport_NY_09_X_29"
28200             End If
28210           End If
28220         End Select

28230         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

28240         strRptCap = vbNullString: strRptPathFile = vbNullString
28250         strRptPath = .UserReportPath
28260         strRptName = strTmp01
28270         DoEvents

28280         .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
28290         DoEvents
28300         arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
28310         lngCaps = UBound(arr_varCap, 2) + 1&

28320         For lngX = 0& To (lngCaps - 1&)
28330           If arr_varCap(C_RNAM, lngX) = strRptName Then
28340             strRptCap = arr_varCap(C_CAPN, lngX)
28350             Exit For
28360           End If
28370         Next
28380         DoEvents

28390         If IsNull(.UserReportPath) = False Then
28400           If .UserReportPath <> vbNullString Then
28410             If .UserReportPath_chk = True Then
28420               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
28430                 blnUseSavedPath = True
28440               End If
28450             End If
28460           End If
28470         End If

28480         strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
28490         If blnNoData = True Then
28500           strMacro = strMacro & "_nd"
28510         End If

28520         Select Case blnUseSavedPath
              Case True
28530           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
28540         Case False
28550           DoCmd.Hourglass False
28560           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
28570         End Select

28580         If strRptPathFile <> vbNullString Then
28590           DoCmd.Hourglass True
28600           DoEvents
28610           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
28620           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
28630             Kill strRptPathFile
28640           End If
28650           If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
28660             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
28670             DoEvents
28680             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
28690               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
28700                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
28710               Else
28720                 Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
28730               End If
28740               DoEvents
28750               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
28760                 DoEvents
28770                 Select Case gblnPrintAll
                      Case True
28780                   lngFiles = lngFiles + 1&
28790                   lngE = lngFiles - 1&
28800                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
28810                   arr_varFile(F_RNAM, lngE) = strRptName
28820                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
28830                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
28840                   FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
28850                 Case False
28860                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
28870                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
28880                   End If
28890                   DoEvents
28900                   If blnAutoStart = True Then
28910                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
28920                   End If
28930                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
28940               End If
28950             End If
28960           Else
28970             DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
28980           End If  ' ** strQry.
28990           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
29000           If strRptPath <> .UserReportPath Then
29010             .UserReportPath = strRptPath
29020             SetUserReportPath_NY frm  ' ** Procedure: Above.
29030           End If
29040         Else
29050           blnAllCancel = True
29060           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
29070           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
29080         End If  ' ** strRptPathFile.

29090       End If  ' ** Validate().
29100     End If ' ** blnContinue.

29110     DoCmd.Hourglass False

29120   End With

      #End If

EXITP:
29130   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev10_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Income.
' **
' ** NY_CmdPrev10_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev10_Click"

29210   With frm

29220     DoCmd.Hourglass True
29230     DoEvents

29240     strThisProc = "cmdPreview10_Click"

29250     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

29260       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

29270       Select Case gblnUseReveuneExpenseCodes
            Case True
29280         PreviewOrPrint_NY "10A", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Module Function: modCourtReportsNY.
29290       Case False
29300         PreviewOrPrint_NY "10", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Module Function: modCourtReportsNY.
29310       End Select

29320     End If  ' ** Validate.

29330     DoCmd.Hourglass False

29340   End With

EXITP:
29350   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint10_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Income.
' **
' ** NY_CmdPrint10_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint10_Click"

29410   With frm

29420     DoCmd.Hourglass True
29430     DoEvents

29440     strThisProc = "cmdPrint10_Click"

29450     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

29460       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

29470       Select Case gblnUseReveuneExpenseCodes
            Case True
              '##GTR_Ref: rptCourtRptNY_10A
29480         PreviewOrPrint_NY "10A", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
29490       Case False
              '##GTR_Ref: rptCourtRptNY_10
29500         PreviewOrPrint_NY "10", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.
29510       End Select

29520     End If  ' ** Validate.

29530     DoCmd.Hourglass False

29540   End With

EXITP:
29550   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord10_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Income.
' **
' ** NY_CmdWord10_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

29600 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord10_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

29610   With frm

29620     DoCmd.Hourglass True
29630     DoEvents

29640     blnUseSavedPath = False
29650     blnExcel = False
29660     blnAllCancel = False
29670     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
29680     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
29690     blnAutoStart = .chkOpenWord
29700     strThisProc = "cmdWord10_Click"

29710     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

29720       Select Case gblnUseReveuneExpenseCodes
            Case True
29730         strRptName = "rptCourtRptNY_10A"
29740       Case False
29750         strRptName = "rptCourtRptNY_10"
29760       End Select
29770       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

29780       strRptCap = vbNullString
29790       strRptCap = "CourtReport_NY_Administration_Expenses_Chargeable_to_Income_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

29800       If IsNull(.UserReportPath) = False Then
29810         If .UserReportPath <> vbNullString Then
29820           If .UserReportPath_chk = True Then
29830             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
29840               blnUseSavedPath = True
29850             End If
29860           End If
29870         End If
29880       End If

29890       Select Case blnUseSavedPath
            Case True
29900         strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
29910       Case False
29920         DoCmd.Hourglass False
29930         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
29940       End Select

29950       If strRptPathFile <> vbNullString Then
29960         DoCmd.Hourglass True
29970         DoEvents
29980         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
29990         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
30000           Kill strRptPathFile
30010         End If
30020         Select Case gblnPrintAll
              Case True
30030           lngFiles = lngFiles + 1&
30040           lngE = lngFiles - 1&
30050           ReDim Preserve arr_varFile(F_ELEMS, lngE)
30060           arr_varFile(F_RNAM, lngE) = strRptName
30070           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
30080           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
30090           FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
30100           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
30110         Case False
30120           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
30130         End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
30140         strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
30150         If strRptPath <> .UserReportPath Then
30160           .UserReportPath = strRptPath
30170           SetUserReportPath_NY frm  ' ** Procedure: Above.
30180         End If
30190       Else
30200         blnAllCancel = True
30210         .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
30220         AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
30230       End If  ' ** strRptPathFile.

30240     End If  ' ** Validate.

30250     DoCmd.Hourglass False

30260   End With

EXITP:
30270   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel10_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Chargeable to Income.
' **
' ** NY_CmdExcel10_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel10_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

30310   With frm

30320     DoCmd.Hourglass True
30330     DoEvents

30340     blnContinue = True
30350     blnUseSavedPath = False
30360     blnExcel = True
30370     blnAllCancel = False
30380     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
30390     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
30400     blnAutoStart = .chkOpenExcel
30410     strThisProc = "cmdExcel10_Click"

30420     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
30430       ForcePause 2  ' ** Module Function: modCodeUtilities.
30440       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
30450         DoCmd.Hourglass False
30460         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
30470         If msgResponse <> vbRetry Then
30480           blnAllCancel = True
30490           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
30500           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
30510           blnContinue = False
30520         End If
30530       End If
30540     End If

30550     If blnContinue = True Then

30560       DoCmd.Hourglass True
30570       DoEvents

30580       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

30590         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

30600         DoEvents

30610         gstrAccountNo = .cmbAccounts.Column(0)
30620         gdatStartDate = .DateEnd
30630         gdatEndDate = .DateStart
30640         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

30650         gblnMessage = False
30660         Select Case gblnUseReveuneExpenseCodes
              Case True
30670           strTmp01 = "rptCourtRptNY_10A"
30680           strQry = "qryCourtReport_NY_10_X_18"
30690           varTmp00 = DCount("*", strQry)
30700           If IsNull(varTmp00) = True Then
30710             blnNoData = True
30720             strQry = "qryCourtReport_NY_10_X_28"
30730           Else
30740             If varTmp00 = 0 Then
30750               blnNoData = True
30760               strQry = "qryCourtReport_NY_10_X_28"
30770             End If
30780           End If
30790         Case False
30800           strTmp01 = "rptCourtRptNY_10"
30810           strQry = "qryCourtReport_NY_10_X_08"
30820           varTmp00 = DCount("*", strQry)
30830           If IsNull(varTmp00) = True Then
30840             blnNoData = True
30850             strQry = "qryCourtReport_NY_10_X_23"
30860           Else
30870             If varTmp00 = 0 Then
30880               blnNoData = True
30890               strQry = "qryCourtReport_NY_10_X_23"
30900             End If
30910           End If
30920         End Select

30930         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

30940         strRptCap = vbNullString: strRptPathFile = vbNullString
30950         strRptPath = .UserReportPath
30960         strRptName = strTmp01
30970         DoEvents

30980         .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
30990         DoEvents
31000         arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
31010         lngCaps = UBound(arr_varCap, 2) + 1&

31020         For lngX = 0& To (lngCaps - 1&)
31030           If arr_varCap(C_RNAM, lngX) = strRptName Then
31040             strRptCap = arr_varCap(C_CAPN, lngX)
31050             Exit For
31060           End If
31070         Next
31080         DoEvents

31090         If IsNull(.UserReportPath) = False Then
31100           If .UserReportPath <> vbNullString Then
31110             If .UserReportPath_chk = True Then
31120               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
31130                 blnUseSavedPath = True
31140               End If
31150             End If
31160           End If
31170         End If

31180         strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
31190         If blnNoData = True Then
31200           strMacro = strMacro & "_nd"
31210         End If

31220         Select Case blnUseSavedPath
              Case True
31230           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
31240         Case False
31250           DoCmd.Hourglass False
31260           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
31270         End Select

31280         If strRptPathFile <> vbNullString Then
31290           DoCmd.Hourglass True
31300           DoEvents
31310           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
31320           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
31330             Kill strRptPathFile
31340           End If
31350           If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
31360             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
31370             DoEvents
31380             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
31390               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
31400                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
31410               Else
31420                 Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
31430               End If
31440               DoEvents
31450               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
31460                 DoEvents
31470                 Select Case gblnPrintAll
                      Case True
31480                   lngFiles = lngFiles + 1&
31490                   lngE = lngFiles - 1&
31500                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
31510                   arr_varFile(F_RNAM, lngE) = strRptName
31520                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
31530                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
31540                   FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
31550                 Case False
31560                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
31570                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
31580                   End If
31590                   DoEvents
31600                   If blnAutoStart = True Then
31610                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
31620                   End If
31630                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
31640               End If
31650             End If
31660           Else
31670             DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
31680           End If  ' ** strQry.
31690           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
31700           If strRptPath <> .UserReportPath Then
31710             .UserReportPath = strRptPath
31720             SetUserReportPath_NY frm  ' ** Procedure: Above.
31730           End If
31740         Else
31750           blnAllCancel = True
31760           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
31770           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
31780         End If  ' ** strRptPathFile.

31790       End If  ' ** Validate().
31800     End If ' ** blnContinue.

31810     DoCmd.Hourglass False

31820   End With

      #End If

EXITP:
31830   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev11_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Income.
' **
' ** NY_CmdPrev11_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

31900 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev11_Click"

31910   With frm

31920     DoCmd.Hourglass True
31930     DoEvents

31940     strThisProc = "cmdPreview11_Click"

31950     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

31960       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

31970       PreviewOrPrint_NY "11", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

31980     End If  ' ** Validate.

31990     DoCmd.Hourglass False

32000   End With

EXITP:
32010   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint11_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Income.
' **
' ** NY_CmdPrint11_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

32100 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint11_Click"

32110   With frm

32120     DoCmd.Hourglass True
32130     DoEvents

32140     strThisProc = "cmdPrint11_Click"

32150     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

32160       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

            '##GTR_Ref: rptCourtRptNY_11
32170       PreviewOrPrint_NY "11", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.

32180     End If  ' ** Validate.

32190     DoCmd.Hourglass False

32200   End With

EXITP:
32210   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord11_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Income.
' **
' ** NY_CmdWord11_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord11_Click"

        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim lngE As Long

32310   With frm

32320     DoCmd.Hourglass True
32330     DoEvents

32340     blnUseSavedPath = False
32350     blnExcel = False
32360     blnAllCancel = False
32370     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
32380     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
32390     blnAutoStart = .chkOpenWord
32400     strThisProc = "cmdWord11_Click"

32410     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

32420       strRptName = "rptCourtRptNY_11"
32430       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

32440       strRptCap = vbNullString
32450       strRptCap = "CourtReport_NY_Distributions_of_Income_" & gstrAccountNo & "_" & _
              Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

32460       If IsNull(.UserReportPath) = False Then
32470         If .UserReportPath <> vbNullString Then
32480           If .UserReportPath_chk = True Then
32490             If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
32500               blnUseSavedPath = True
32510             End If
32520           End If
32530         End If
32540       End If

32550       Select Case blnUseSavedPath
            Case True
32560         strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
32570       Case False
32580         DoCmd.Hourglass False
32590         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.
32600       End Select

32610       If strRptPathFile <> vbNullString Then
32620         DoCmd.Hourglass True
32630         DoEvents
32640         If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
32650         If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
32660           Kill strRptPathFile
32670         End If
32680         Select Case gblnPrintAll
              Case True
32690           lngFiles = lngFiles + 1&
32700           lngE = lngFiles - 1&
32710           ReDim Preserve arr_varFile(F_ELEMS, lngE)
32720           arr_varFile(F_RNAM, lngE) = strRptName
32730           arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
32740           arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
32750           FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
32760           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
32770         Case False
32780           DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
32790         End Select
              'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
32800         strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
32810         If strRptPath <> .UserReportPath Then
32820           .UserReportPath = strRptPath
32830           SetUserReportPath_NY frm  ' ** Procedure: Above.
32840         End If
32850       Else
32860         blnAllCancel = True
32870         .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
32880         AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
32890       End If  ' ** strRptPathFile.

32900     End If  ' ** Validate.

32910     DoCmd.Hourglass False

32920   End With

EXITP:
32930   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel11_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Distributions of Income.
' **
' ** NY_CmdExcel11_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

33000 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel11_Click"

        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

33010   With frm

33020     DoCmd.Hourglass True
33030     DoEvents

33040     blnContinue = True
33050     blnUseSavedPath = False
33060     blnExcel = True
33070     blnAllCancel = False
33080     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
33090     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
33100     blnAutoStart = .chkOpenExcel
33110     strThisProc = "cmdExcel11_Click"

33120     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
33130       ForcePause 2  ' ** Module Function: modCodeUtilities.
33140       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
33150         DoCmd.Hourglass False
33160         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
33170         If msgResponse <> vbRetry Then
33180           blnAllCancel = True
33190           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
33200           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
33210           blnContinue = False
33220         End If
33230       End If
33240     End If

33250     If blnContinue = True Then

33260       DoCmd.Hourglass True
33270       DoEvents

33280       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

33290         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

33300         DoEvents

33310         gstrAccountNo = .cmbAccounts.Column(0)
33320         gdatStartDate = .DateEnd
33330         gdatEndDate = .DateStart
33340         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
33350         If glngTaxCode_Distribution = 0& Then
33360           glngTaxCode_Distribution = DLookup("[taxcode]", "TaxCode", "[taxcode_description] = 'Distribution'")
33370         End If
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

33380         gblnMessage = False
33390         strTmp01 = "rptCourtRptNY_11"
33400         strQry = "qryCourtReport_NY_11_X_08"
33410         varTmp00 = DCount("*", strQry)
33420         If IsNull(varTmp00) = True Then
33430           blnNoData = True
33440           strQry = "qryCourtReport_NY_11_X_13"
33450         Else
33460           If varTmp00 = 0 Then
33470             blnNoData = True
33480             strQry = "qryCourtReport_NY_11_X_13"
33490           End If
33500         End If

33510         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

33520         strRptCap = vbNullString: strRptPathFile = vbNullString
33530         strRptPath = .UserReportPath
33540         strRptName = strTmp01
33550         DoEvents

33560         .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
33570         DoEvents
33580         arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
33590         lngCaps = UBound(arr_varCap, 2) + 1&

33600         For lngX = 0& To (lngCaps - 1&)
33610           If arr_varCap(C_RNAM, lngX) = strRptName Then
33620             strRptCap = arr_varCap(C_CAPN, lngX)
33630             Exit For
33640           End If
33650         Next
33660         DoEvents

33670         If IsNull(.UserReportPath) = False Then
33680           If .UserReportPath <> vbNullString Then
33690             If .UserReportPath_chk = True Then
33700               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
33710                 blnUseSavedPath = True
33720               End If
33730             End If
33740           End If
33750         End If

33760         strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
33770         If blnNoData = True Then
33780           strMacro = strMacro & "_nd"
33790         End If

33800         Select Case blnUseSavedPath
              Case True
33810           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
33820         Case False
33830           DoCmd.Hourglass False
33840           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
33850         End Select

33860         If strRptPathFile <> vbNullString Then
33870           DoCmd.Hourglass True
33880           DoEvents
33890           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
33900           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
33910             Kill strRptPathFile
33920           End If
33930           If strQry <> vbNullString Then
                  ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                  ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
33940             DoCmd.RunMacro strMacro
                  ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                  ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
33950             DoEvents
33960             If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                      FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
33970               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
33980                 Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
33990               Else
34000                 Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                      ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
34010               End If
34020               DoEvents
34030               If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
34040                 DoEvents
34050                 Select Case gblnPrintAll
                      Case True
34060                   lngFiles = lngFiles + 1&
34070                   lngE = lngFiles - 1&
34080                   ReDim Preserve arr_varFile(F_ELEMS, lngE)
34090                   arr_varFile(F_RNAM, lngE) = strRptName
34100                   arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
34110                   arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
34120                   FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
34130                 Case False
34140                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
34150                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
34160                   End If
34170                   DoEvents
34180                   If blnAutoStart = True Then
34190                     OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
34200                   End If
34210                 End Select
                      'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                      '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                      'End If
                      'DoEvents
                      'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
34220               End If
34230             End If
34240           Else
34250             DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
34260           End If  ' ** strQry.
34270           strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
34280           If strRptPath <> .UserReportPath Then
34290             .UserReportPath = strRptPath
34300             SetUserReportPath_NY frm  ' ** Procedure: Above.
34310           End If
34320         Else
34330           blnAllCancel = True
34340           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
34350           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
34360         End If  ' ** strRptPathFile.

34370       End If  ' ** Validate().
34380     End If ' ** blnContinue.

34390     DoCmd.Hourglass False

34400   End With

      #End If

EXITP:
34410   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev12_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Remaining on Hand.
' **
' ** NY_CmdPrev12_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

34500 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev12_Click"

        Dim strRpt As String, strDocName As String
        Dim intRetVal_BuildAssetListInfo As Integer

34510   With frm

34520     DoCmd.Hourglass True
34530     DoEvents

34540     strThisProc = "cmdPreview12_Click"

34550     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

34560       strRpt = vbNullString
34570       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

34580       If gblnCrtRpt_NY_InvIncChange = False Then
34590         gstrCrtRpt_NY_InputTitle = "Invested Income"
34600         DoCmd.Hourglass False
34610         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
34620         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
34630         DoCmd.Hourglass True
34640         DoEvents
34650       End If

34660       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then
34670         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm)  ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

34680         Select Case intRetVal_BuildAssetListInfo
              Case 0

34690           PreviewOrPrint_NY "12", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

34700         Case -2
34710           Beep
34720           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
34730         Case -3, -9
                ' ** Message shown below.
34740         End Select  ' ** intRetVal_BuildAssetListInfo

34750       End If  ' ** CashAssets_Beg.
34760     End If  ' ** Validate.

34770     DoCmd.Hourglass False

34780   End With

EXITP:
34790   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint12_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Remaining on Hand.
' **
' ** NY_CmdPrint12_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

34800 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint12_Click"

        Dim strRpt As String, strDocName As String
        Dim intRetVal_BuildAssetListInfo As Integer

34810   With frm

34820     DoCmd.Hourglass True
34830     DoEvents

34840     strThisProc = "cmdPrint12_Click"

34850     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

34860       strRpt = vbNullString
34870       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

34880       If gblnCrtRpt_NY_InvIncChange = False Then
34890         gstrCrtRpt_NY_InputTitle = "Invested Income"
34900         DoCmd.Hourglass False
34910         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
34920         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
34930         DoCmd.Hourglass True
34940         DoEvents
34950       End If

34960       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then
34970         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm)  ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

34980         Select Case intRetVal_BuildAssetListInfo
              Case 0

                '##GTR_Ref: rptCourtRptNY_12
34990           PreviewOrPrint_NY "12", strThisProc, acNormal, blnRebuildTable, frm  ' ** Function: Above.

35000         Case -2
35010           Beep
35020           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
35030         Case -3, -9
                ' ** Message shown below.
35040         End Select  ' ** intRetVal_BuildAssetListInfo

35050       End If  ' ** CashAssets_Beg.
35060     End If  ' ** Validate.

35070     DoCmd.Hourglass False

35080   End With

EXITP:
35090   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord12_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Remaining on Hand.
' **
' ** NY_CmdWord12_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord12_Click"

        Dim strRpt As String, strDocName As String
        Dim blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim lngE As Long

35110   With frm

35120     DoCmd.Hourglass True
35130     DoEvents

35140     blnUseSavedPath = False
35150     blnExcel = False
35160     blnAllCancel = False
35170     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
35180     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
35190     blnAutoStart = .chkOpenWord
35200     strThisProc = "cmdWord12_Click"

35210     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

35220       strRpt = vbNullString
35230       strRptName = "rptCourtRptNY_12"
35240       .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

35250       If gblnCrtRpt_NY_InvIncChange = False Then
35260         gstrCrtRpt_NY_InputTitle = "Invested Income"
35270         DoCmd.Hourglass False
35280         strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
35290         DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
35300         DoCmd.Hourglass True
35310         DoEvents
35320       End If

35330       If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then

35340         strRptCap = vbNullString
35350         strRptCap = "CourtReport_NY_Income_Remaining_on_Hand_" & gstrAccountNo & "_" & _
                Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

35360         If IsNull(.UserReportPath) = False Then
35370           If .UserReportPath <> vbNullString Then
35380             If .UserReportPath_chk = True Then
35390               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
35400                 blnUseSavedPath = True
35410               End If
35420             End If
35430           End If
35440         End If

35450         Select Case blnUseSavedPath
              Case True
35460           strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
35470         Case False
35480           DoCmd.Hourglass False
35490           strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, strRptCap) ' ** Module Function: modBrowseFilesAndFolders.\
35500         End Select

35510         If strRptPathFile <> vbNullString Then
35520           DoCmd.Hourglass True
35530           DoEvents
35540           If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
35550           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
35560             Kill strRptPathFile
35570           End If
35580           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm)  ' ** Function: Above.
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.

35590           Select Case intRetVal_BuildAssetListInfo
                Case 0

35600             Select Case gblnPrintAll
                  Case True
35610               lngFiles = lngFiles + 1&
35620               lngE = lngFiles - 1&
35630               ReDim Preserve arr_varFile(F_ELEMS, lngE)
35640               arr_varFile(F_RNAM, lngE) = strRptName
35650               arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
35660               arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
35670               FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
35680               DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
35690             Case False
35700               DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
35710             End Select
                  'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
35720             strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
35730             If strRptPath <> .UserReportPath Then
35740               .UserReportPath = strRptPath
35750               SetUserReportPath_NY frm  ' ** Procedure: Above.
35760             End If

35770           Case -2
35780             Beep
35790             MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
35800           Case -3, -9
                  ' ** Message shown below.
35810           End Select  ' ** intRetVal_BuildAssetListInfo

35820         Else
35830           blnAllCancel = True
35840           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
35850           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
35860         End If  ' ** strRptPathFile.

35870       End If  ' ** CashAssets_Beg.
35880     End If  ' ** Validate.

35890     DoCmd.Hourglass False

35900   End With

EXITP:
35910   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel12_Click(strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String, strCallingForm As String, frm As Access.Form)
' ** Statement of Income Remaining on Hand.
' **
' ** NY_CmdExcel12_Click(
' **   strRptName As String, strRptCap As String, strRptPath As String,
' **   strRptPathFile As String, strCallingForm As String, frm As Access.Form
' ** )

36000 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel12_Click"

        Dim strRpt As String, strDocName As String
        Dim strQry As String, strMacro As String
        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

      #If Not NoExcel Then

36010   With frm

36020     DoCmd.Hourglass True
36030     DoEvents

36040     blnContinue = True
36050     blnUseSavedPath = False
36060     blnExcel = True
36070     blnAllCancel = False
36080     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
36090     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
36100     blnAutoStart = .chkOpenExcel
36110     strThisProc = "cmdExcel12_Click"

36120     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
            ' ** It seems like it's not quite closed when it gets here,
            ' ** because if I stop the code and run the function again,
            ' ** it always comes up False.
36130       ForcePause 2  ' ** Module Function: modCodeUtilities.
36140       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
36150         DoCmd.Hourglass False
36160         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                "You may close Excel before proceding, then click Retry." & vbCrLf & _
                "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
              ' ** ... Otherwise Trust Accountant will do it for you.
36170         If msgResponse <> vbRetry Then
36180           blnAllCancel = True
36190           .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
36200           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
36210           blnContinue = False
36220         End If
36230       End If
36240     End If

36250     If blnContinue = True Then

36260       DoCmd.Hourglass True
36270       DoEvents

36280       If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

36290         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

36300         DoEvents

36310         gstrAccountNo = .cmbAccounts.Column(0)
36320         gdatStartDate = .DateEnd
36330         gdatEndDate = .DateStart
36340         gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal, gstrCrtRpt_Version, and gcurCrtRpt_NY_InputNew should be populated from the input window.

36350         gblnMessage = False
36360         strTmp01 = "rptCourtRptNY_12"
36370         strQry = "qryCourtReport_NY_12_X_12"
36380         varTmp00 = DCount("*", strQry)
36390         If IsNull(varTmp00) = True Then
36400           blnNoData = True
36410           strQry = "qryCourtReport_NY_12_X_17"
36420         Else
36430           If varTmp00 = 0 Then
36440             blnNoData = True
36450             strQry = "qryCourtReport_NY_12_X_17"
36460           End If
36470         End If

36480         .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

36490         strRptCap = vbNullString: strRptPathFile = vbNullString
36500         strRptPath = .UserReportPath
36510         strRptName = strTmp01
36520         DoEvents

36530         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, strThisProc, frm) ' ** Function: Above.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

36540         Select Case intRetVal_BuildAssetListInfo
              Case 0

36550           DoEvents

36560           If gblnCrtRpt_NY_InvIncChange = False Then
36570             gstrCrtRpt_NY_InputTitle = "Invested Income"
36580             DoCmd.Hourglass False
36590             strDocName = "frmRpt_CourtReports_NY_Input_InvestedIncome"
36600             DoCmd.OpenForm strDocName, , , , , acDialog, strCallingForm
36610             DoCmd.Hourglass True
36620             DoEvents
36630           End If

36640           If .CashAssets_Beg <> vbNullString Or gblnCrtRpt_NY_InvIncChange = True Then

36650             .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
36660             DoEvents
36670             arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
36680             lngCaps = UBound(arr_varCap, 2) + 1&

36690             For lngX = 0& To (lngCaps - 1&)
36700               If arr_varCap(C_RNAM, lngX) = strRptName Then
36710                 strRptCap = arr_varCap(C_CAPN, lngX)
36720                 Exit For
36730               End If
36740             Next
36750             DoEvents

36760             If IsNull(.UserReportPath) = False Then
36770               If .UserReportPath <> vbNullString Then
36780                 If .UserReportPath_chk = True Then
36790                   If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
36800                     blnUseSavedPath = True
36810                   End If
36820                 End If
36830               End If
36840             End If

36850             strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
36860             If blnNoData = True Then
36870               strMacro = strMacro & "_nd"
36880             End If

36890             Select Case blnUseSavedPath
                  Case True
36900               strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
36910             Case False
36920               DoCmd.Hourglass False
36930               strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
36940             End Select

36950             If strRptPathFile <> vbNullString Then
36960               DoCmd.Hourglass True
36970               DoEvents
36980               If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
36990               If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
37000                 Kill strRptPathFile
37010               End If
37020               If strQry <> vbNullString Then
                      ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                      ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
37030                 DoCmd.RunMacro strMacro
                      ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                      ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
37040                 DoEvents
37050                 If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                          FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
37060                   If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
37070                     Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
37080                   Else
37090                     Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
37100                   End If
37110                   DoEvents
37120                   If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
37130                     DoEvents
37140                     Select Case gblnPrintAll
                          Case True
37150                       lngFiles = lngFiles + 1&
37160                       lngE = lngFiles - 1&
37170                       ReDim Preserve arr_varFile(F_ELEMS, lngE)
37180                       arr_varFile(F_RNAM, lngE) = strRptName
37190                       arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
37200                       arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
37210                       FileArraySet_NY arr_varFile  ' ** Module Procedure: modCourtReportsNY1.
37220                     Case False
37230                       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
37240                         EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
37250                       End If
37260                       DoEvents
37270                       If blnAutoStart = True Then
37280                         OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
37290                       End If
37300                     End Select
                          'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                          '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                          'End If
                          'DoEvents
                          'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
37310                   End If
37320                 End If
37330               Else
37340                 DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
37350               End If  ' ** strQry.
37360               strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
37370               If strRptPath <> .UserReportPath Then
37380                 .UserReportPath = strRptPath
37390                 SetUserReportPath_NY frm  ' ** Procedure: Above.
37400               End If
37410             Else
37420               blnAllCancel = True
37430               .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
37440               AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
37450             End If  ' ** strRptPathFile.

37460           End If  ' ** CashAssets_Beg.

37470         Case -2
37480           Beep
37490           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
37500         Case -3, -9
                ' ** Message shown below.
37510         End Select  ' ** intRetVal_BuildAssetListInfo

37520       End If  ' ** Validate().
37530     End If ' ** blnContinue.

37540     DoCmd.Hourglass False

37550   End With

      #End If

EXITP:
37560   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub NY_CmdPrev13_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Unpaid Chargeable to Principal.
' **
' ** NY_CmdPrev13_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

37600 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrev13_Click"

37610   With frm

37620     DoCmd.Hourglass True
37630     DoEvents

37640     strThisProc = "cmdPreview13_Click"

37650     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

37660       PreviewOrPrint_NY "13", strThisProc, acViewPreview, blnRebuildTable, frm  ' ** Function: Above.

37670     End If

37680     DoCmd.Hourglass False

37690   End With

EXITP:
37700   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdPrint13_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Unpaid Chargeable to Principal.
' **
' ** NY_CmdPrint13_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

37800 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdPrint13_Click"

37810   With frm

37820     DoCmd.Hourglass True
37830     DoEvents

37840     strThisProc = "cmdPrint13_Click"

37850     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

            '##GTR_Ref: rptCourtRptNY_13
37860       PreviewOrPrint_NY "13", strThisProc, acViewNormal, blnRebuildTable, frm  ' ** Function: Above.

37870     End If

37880     DoCmd.Hourglass False

37890   End With

EXITP:
37900   Exit Sub

ERRH:
470     DoCmd.Hourglass False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Sub

Public Sub NY_CmdWord13_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Unpaid Chargeable to Principal.
' **
' ** NY_CmdWord13_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

38000 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdWord13_Click"

        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnAutoStart As Boolean

38010   With frm

38020     DoCmd.Hourglass True
38030     DoEvents

38040     blnExcel = False
38050     blnAllCancel = False
38060     .AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
38070     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
38080     blnAutoStart = .chkOpenWord
38090     strThisProc = "cmdWord13_Click"

38100     If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

38110       .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
38120       DoEvents
38130       arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
38140       lngCaps = UBound(arr_varCap, 2) + 1&

38150       SendToFile_NY frm, "13", blnRebuildTable, strThisProc, lngCaps, arr_varCap, strCallingForm  ' ** Procedure: Above.

38160     End If

38170     DoCmd.Hourglass False

38180   End With

EXITP:
38190   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case Else
530       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
540     End Select
550     Resume EXITP

End Sub

Public Sub NY_CmdExcel13_Click(blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form)
' ** Statement of Administration Expenses Unpaid Chargeable to Principal.
' **
' ** NY_CmdExcel13_Click(
' **   blnRebuildTable As Boolean, strCallingForm As String, frm As Access.Form
' ** )

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_CmdExcel13_Click"

        Dim lngCaps As Long, arr_varCap As Variant
        Dim blnSkip As Boolean

      #If Not NoExcel Then

38210   With frm

38220     DoCmd.Hourglass True
38230     DoEvents

38240     blnSkip = True
38250     If blnSkip = False Then

38260       .CapArray_Load  ' ** Form Function: frmRpt_CourtReports_NY.
38270       DoEvents
38280       arr_varCap = .CapArray_Get  ' ** Form Function: frmRpt_CourtReports_NY.
38290       lngCaps = UBound(arr_varCap, 2) + 1&

38300       SendToFile_NY frm, "13", blnRebuildTable, THIS_PROC, lngCaps, arr_varCap, True  ' ** Procedure: Above.

38310     Else
38320       MsgBox "Temporarily unavailable.", vbInformation + vbOKOnly, "Export Unavailable"
38330     End If  ' ** blnSkip.

38340     DoCmd.Hourglass False

38350   End With

      #End If

EXITP:
38360   Exit Sub

ERRH:
470     blnAllCancel = True
480     frm.AllCancelSet1_NY blnAllCancel  ' ** Form Function: frmRpt_CourtReports_NY.
490     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
500     gblnPrintAll = False
510     DoCmd.Hourglass False
520     Select Case ERR.Number
        Case 70  ' ** Permission denied.
530       Beep
540       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub FileArrayInit()

38400 On Error GoTo ERRH

        Const THIS_PROC As String = "FileArrayInit"

38410   lngFiles = 0&
38420   ReDim arr_varfiles(F_ELEMS, 0)

EXITP:
38430   Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub

Public Function NY_ListRptTitlesPeriods() As Boolean

38500 On Error GoTo ERRH

        Const THIS_PROC As String = "NY_ListRptTitlesPeriods"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngQrys As Long, arr_varQry() As Variant
        Dim strDesc As String, strSQL As String
        Dim intMode As Integer
        Dim intPos01 As Integer, intPos02 As Integer, intPos03 As Integer, intLen As Integer
        Dim strTmp01 As String
        Dim lngX As Long, lngE As Long
        Dim blnRetVal As Boolean

        ' ** Array: arr_varQry().
        Const Q_ELEMS As Integer = 1  ' ** Array's first-element UBound().
        Const Q_QNAM  As Integer = 0
        Const Q_TITLE As Integer = 1

        Const QRY_BASE As String = "qryCourtReport_NY"

38510 On Error GoTo 0

38520   blnRetVal = True

38530   Win_Mod_Restore  ' ** Module Procedure: modWindowFunctions.
38540   DoEvents

38550   intLen = Len(QRY_BASE)
        'intMode = 1  ' ** Title.
38560   intMode = 2  ' ** Period.

38570   lngQrys = 0&
38580   ReDim arr_varQry(Q_ELEMS, 0)

38590   Set dbs = CurrentDb
38600   With dbs
38610     For Each qdf In .QueryDefs
38620       strDesc = vbNullString
38630       With qdf
38640         If Left(.Name, intLen) = QRY_BASE Then
38650 On Error Resume Next
38660           strDesc = .Properties("Description")
38670 On Error GoTo 0
38680           If strDesc <> vbNullString Then
38690             Select Case intMode
                  Case 1  ' ** Title.
38700               strTmp01 = "Report title"
38710             Case 2  ' ** Period.
38720               strTmp01 = "Report period"
38730             End Select
38740             If InStr(strDesc, strTmp01) > 0 Then
                    ' ** SELECT CInt(-2) AS ReportNumber, 'Title:' AS accountno,
                    ' **   'Property on Hand at Close of Accounting Period' AS ShortName,
                    ' **   '' AS CaseNum, '' AS assettype, '' AS assettype_description,
                    ' **   CDbl(0) AS TotalShareface, '' AS totdesc, CDbl(0) AS TotalMarket,
                    ' **   CCur(0) AS TotalCost, CInt(-2) AS Sort, '00000000' AS uniqueid,
                    ' **   '0000' AS uniqueidx, '00' AS uniqueidy
                    ' ** FROM tblYesNo
                    ' ** WHERE (((tblYesNo.yn_name)='Yes'));
                    ' ** SELECT 'Period:' AS [Account Num], GlobalVarGet("gstrCrtRpt_Period") AS Name,
                    ' **   "" AS [Case Number], "" AS Description, Null AS [Invenory Value], Null AS Sort1,
                    ' **   CLng(-1) AS Sort2, CLng(1) AS Sort3
                    ' ** FROM tblYesNo
                    ' ** WHERE (((tblYesNo.yn_boolean)=True));

                    ' ** 'Property on Hand at Close of Accounting Period' AS ShortName
38750               strSQL = .SQL
38760               intPos01 = InStr(strSQL, "AS accountno, ")
38770               intPos02 = InStr(strSQL, "AS [Account Num], ")
38780               intPos03 = InStr(strSQL, ".accountno, ")
38790               strTmp01 = vbNullString
38800               If intPos01 > 0 Then
38810                 strTmp01 = Trim(Mid(strSQL, (intPos01 + 13)))
38820                 intPos02 = InStr(strTmp01, "AS ShortName, ")
38830                 intPos03 = InStr(strTmp01, "AS Namex, ")
38840                 If intPos02 > 0 Then
38850                   strTmp01 = Trim(Left(strTmp01, (intPos02 + 13)))
38860                 ElseIf intPos03 > 0 Then
38870                   strTmp01 = Trim(Left(strTmp01, (intPos03 + 9)))
38880                 Else
38890                   Stop
38900                 End If
38910               ElseIf intPos02 > 0 Then
38920                 strTmp01 = Trim(Mid(strSQL, (intPos02 + 17)))
38930                 intPos01 = InStr(strTmp01, "AS Name, ")
38940                 If intPos01 > 0 Then
38950                   strTmp01 = Trim(Left(strTmp01, (intPos01 + 8)))
38960                 Else
38970                   Stop
38980                 End If
38990               ElseIf intPos03 > 0 Then
39000                 strTmp01 = Trim(Mid(strSQL, (intPos03 + 11)))
39010                 intPos02 = InStr(strTmp01, "AS ShortName, ")
39020                 intPos03 = InStr(strTmp01, "AS Namex, ")
39030                 If intPos02 > 0 Then
39040                   strTmp01 = Trim(Left(strTmp01, (intPos02 + 13)))
39050                 ElseIf intPos03 > 0 Then
39060                   strTmp01 = Trim(Left(strTmp01, (intPos03 + 9)))
39070                 Else
39080                   Stop
39090                 End If
39100               End If
39110               If strTmp01 <> vbNullString Then
39120                 intPos01 = InStr(strTmp01, " AS ")
39130                 If intPos01 > 0 Then
39140                   strTmp01 = Trim(Left(strTmp01, intPos01))
39150                 End If
39160                 lngQrys = lngQrys + 1&
39170                 lngE = lngQrys - 1&
39180                 ReDim Preserve arr_varQry(Q_ELEMS, lngE)
39190                 arr_varQry(Q_QNAM, lngE) = .Name
39200                 arr_varQry(Q_TITLE, lngE) = strTmp01
39210               End If
39220             End If
39230           End If
39240         End If
39250       End With  ' ** qdf.
39260     Next  ' ** qdf.
39270     Set qdf = Nothing
39280     .Close
39290   End With  ' ** dbs.
39300   Set dbs = Nothing

39310   Debug.Print "'QRYS: " & CStr(lngQrys)
39320   DoEvents

39330   If lngQrys > 0& Then
39340     For lngX = 0& To (lngQrys - 1&)
39350       Debug.Print "'" & Left(CStr(lngX + 1&) & "." & Space(5), 5) & Left(arr_varQry(Q_QNAM, lngX) & Space(28), 28) & _
              arr_varQry(Q_TITLE, lngX)
39360     Next  ' ** lngX.
39370   End If  ' ** lngQrys.

39380   Beep

        'QRYS: 52
'1. x qryCourtReport_NY_00_11b    'Summary of Account - ' & FormRef('OrdVer')
'2. x qryCourtReport_NY_00_A_08b  'Summary of Account - Grouped - ' & FormRef('OrdVer')
'3.   qryCourtReport_NY_00_X_12   'Summary Statement - ' & GlobalVarGet("gstrCrtRpt_Ordinal") & " and " & GlobalVarGet("gstrCrtRpt_Version") & " Account"
'4.   qryCourtReport_NY_00_X_17   'Summary Statement - ' & GlobalVarGet("gstrCrtRpt_Ordinal") & " and " & GlobalVarGet("gstrCrtRpt_Version") & " Account"
'5.   qryCourtReport_NY_00A_X_14  'Summary Statement - Grouped - ' & FormRef('OrdVer')
'6.   qryCourtReport_NY_00A_X_19  'Summary Statement - ' & GlobalVarGet("gstrCrtRpt_Ordinal") & " and " & GlobalVarGet("gstrCrtRpt_Version") & " Account"
'7.   qryCourtReport_NY_00_B_21a  'Property on Hand at Ending of Account Period'
'8.   qryCourtReport_NY_00_B_21b  'Property on Hand at Beginning of Accounting Period'  NOT USED!
'9.   qryCourtReport_NY_00_B_26   'Property on Hand at Ending of Account Period'
'10.  qryCourtReport_NY_00B_X_08  'Property on Hand at Ending of Account Period'
'11.  qryCourtReport_NY_00B_X_16  'Property on Hand at Ending of Account Period'
        '12.x qryCourtReport_NY_01_09     'Statement of Principal Received - Schedule A'
'13.  qryCourtReport_NY_01_X_20   'Statement of Principal Received - Schedule A'
'14.  qryCourtReport_NY_01_X_54   'Statement of Principal Received - Schedule A'
        '15.x qryCourtReport_NY_02_08     'Statement of Increases on Sales, Liquidation or Distribution - Schedule A-1'
'16.  qryCourtReport_NY_02_X_04   'Statement of Increases on Sales, Liquidation or Distribution - Schedule A-1'
'17.  qryCourtReport_NY_02_X_11   'Statement of Increases on Sales, Liquidation or Distribution - Schedule A-1'
        '18.x qryCourtReport_NY_03_12     'Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability - Schedule B'
        '19.x qryCourtReport_NY_03_17a    'Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability - Schedule B'
'20.  qryCourtReport_NY_03_X_03   'Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability - Schedule B'
'21.  qryCourtReport_NY_03_X_10   'Statement of Decreases Due to Sales, Liquidation, Collection, Distribution or Uncollectability - Schedule B'
        '22.x qryCourtReport_NY_04_08     'Statement of Administration Expenses Chargeable to Principal - Schedule C'
'23.  qryCourtReport_NY_04_X_03   'Statement of Administration Expenses Chargeable to Principal - Schedule C'
'24.  qryCourtReport_NY_04_X_13   'Statement of Administration Expenses Chargeable to Principal - Grouped - Schedule C'
'25.  qryCourtReport_NY_04_X_22   'Statement of Administration Expenses Chargeable to Principal - Schedule C'
'26.  qryCourtReport_NY_04_X_27   'Statement of Administration Expenses Chargeable to Principal - Grouped - Schedule C'
        '27.x qryCourtReport_NY_05_08     'Statement of Distributions of Principal - Schedule D'
'28.  qryCourtReport_NY_05_X_03   'Statement of Distributions of Principal - Schedule D'
'29.  qryCourtReport_NY_05_X_10   "Statement of Distributions of Principal - Schedule D"
        '30.x qryCourtReport_NY_06_09     'Statement of New Investments, Exchanges and Stock Distributions of Principal Assets - Schedule E'
'31.  qryCourtReport_NY_06_X_03   'Statement of New Investments, Exchanges and Stock Distributions of Principal Assets - Schedule E'
'32.  qryCourtReport_NY_06_X_20   'Statement of New Investments, Exchanges and Stock Distributions of Principal Assets - Schedule E'
        '33.x qryCourtReport_NY_07_12     'Statement of Principal Remaining on Hand - Schedule F'
'34.  qryCourtReport_NY_07_X_10   'Statement of Principal Remaining on Hand - Schedule F'
'35.  qryCourtReport_NY_07_X_32   'Statement of Principal Remaining on Hand - Schedule F'
        '36.x qryCourtReport_NY_08_12     'Statement of Income Received - Schedule AA-1'
        '37.x qryCourtReport_NY_08_17a    'Statement of Income Received - Schedule AA-1'
'38.  qryCourtReport_NY_08_X_05   'Statement of Income Received - Schedule AA-1'
'39.  qryCourtReport_NY_08_X_11   'Statement of Income Received - Schedule AA-1'
        '40.x qryCourtReport_NY_09_08     'Statement of All Income Collected - Schedule A-2'
'41.  qryCourtReport_NY_09_X_06   'Statement of All Income Collected - Schedule A-2'
'42.  qryCourtReport_NY_09_X_18   'Statement of All Income Collected - Grouped - Schedule A-2'
'43.  qryCourtReport_NY_09_X_26   'Statement of All Income Collected - Schedule A-2'
'44.  qryCourtReport_NY_09_X_31   'Statement of All Income Collected - Grouped - Schedule A-2'
'45.  qryCourtReport_NY_10_X_04   'Statement of Administration Expenses Chargeable to Income - Schedule C-2'
'46.  qryCourtReport_NY_10_X_13   'Statement of Administration Expenses Chargeable to Income - Grouped - Schedule C-2'
'47.  qryCourtReport_NY_10_X_20   'Statement of Administration Expenses Chargeable to Income - Schedule C-2'
'48.  qryCourtReport_NY_10_X_25   'Statement of Administration Expenses Chargeable to Income - Grouped - Schedule C-2'
'49.  qryCourtReport_NY_11_X_03   'Statement of Distributions of Income - Schedule D-1'
'50.  qryCourtReport_NY_11_X_10   'Statement of Distributions of Income - Schedule D-1'
'51.  qryCourtReport_NY_12_X_08   'Statement of Income Remaining on Hand - Schedule F-1'
'52.  qryCourtReport_NY_12_X_14   'Statement of Income Remaining on Hand - Schedule F-1'
        'DONE!

        'QRYS: 49
'1.   qryCourtReport_NY_00_11c    GlobalVarGet("gstrCrtRpt_Period")
'2.   qryCourtReport_NY_00_A_08c  GlobalVarGet("gstrCrtRpt_Period")
'3.   qryCourtReport_NY_00_B_22   GlobalVarGet("gstrCrtRpt_Period")
'4.   qryCourtReport_NY_00_B_27   GlobalVarGet("gstrCrtRpt_Period")
'5.   qryCourtReport_NY_00_X_13   GlobalVarGet("gstrCrtRpt_Period")
'6.   qryCourtReport_NY_00_X_18   GlobalVarGet("gstrCrtRpt_Period")
'7.   qryCourtReport_NY_00A_X_15  GlobalVarGet("gstrCrtRpt_Period")
'8.   qryCourtReport_NY_00A_X_20  GlobalVarGet("gstrCrtRpt_Period")
'9.   qryCourtReport_NY_00B_X_09  GlobalVarGet("gstrCrtRpt_Period")
'10.  qryCourtReport_NY_00B_X_17  GlobalVarGet("gstrCrtRpt_Period")
'11.  qryCourtReport_NY_01_10     GlobalVarGet("gstrCrtRpt_Period")
'12.  qryCourtReport_NY_01_X_21   GlobalVarGet("gstrCrtRpt_Period")
'13.  qryCourtReport_NY_01_X_55   GlobalVarGet("gstrCrtRpt_Period")
'14.  qryCourtReport_NY_02_09     GlobalVarGet("gstrCrtRpt_Period")
'15.  qryCourtReport_NY_02_X_05   GlobalVarGet("gstrCrtRpt_Period")
'16.  qryCourtReport_NY_02_X_12   GlobalVarGet("gstrCrtRpt_Period")
'17.  qryCourtReport_NY_03_13     GlobalVarGet("gstrCrtRpt_Period")
'18.  qryCourtReport_NY_03_X_04   GlobalVarGet("gstrCrtRpt_Period")
'19.  qryCourtReport_NY_03_X_11   GlobalVarGet("gstrCrtRpt_Period")
'20.  qryCourtReport_NY_04_09     GlobalVarGet("gstrCrtRpt_Period")
'21.  qryCourtReport_NY_04_X_04   GlobalVarGet("gstrCrtRpt_Period")
'22.  qryCourtReport_NY_04_X_14   GlobalVarGet("gstrCrtRpt_Period")
'23.  qryCourtReport_NY_04_X_23   GlobalVarGet("gstrCrtRpt_Period")
'24.  qryCourtReport_NY_04_X_28   GlobalVarGet("gstrCrtRpt_Period")
'25.  qryCourtReport_NY_05_09     GlobalVarGet("gstrCrtRpt_Period")
'26.  qryCourtReport_NY_05_X_04   GlobalVarGet("gstrCrtRpt_Period")
'27.  qryCourtReport_NY_05_X_11   GlobalVarGet("gstrCrtRpt_Period")
'28.  qryCourtReport_NY_06_10     GlobalVarGet("gstrCrtRpt_Period")
'29.  qryCourtReport_NY_06_X_04   GlobalVarGet("gstrCrtRpt_Period")
'30.  qryCourtReport_NY_06_X_21   GlobalVarGet("gstrCrtRpt_Period")
'31.  qryCourtReport_NY_07_13     GlobalVarGet("gstrCrtRpt_Period")
'32.  qryCourtReport_NY_07_X_11   GlobalVarGet("gstrCrtRpt_Period")
'33.  qryCourtReport_NY_07_X_33   GlobalVarGet("gstrCrtRpt_Period")
'34.  qryCourtReport_NY_08_13     GlobalVarGet("gstrCrtRpt_Period")
'35.  qryCourtReport_NY_08_X_06   GlobalVarGet("gstrCrtRpt_Period")
'36.  qryCourtReport_NY_08_X_12   GlobalVarGet("gstrCrtRpt_Period")
'37.  qryCourtReport_NY_09_09     GlobalVarGet("gstrCrtRpt_Period")
'38.  qryCourtReport_NY_09_X_07   GlobalVarGet("gstrCrtRpt_Period")
'39.  qryCourtReport_NY_09_X_19   GlobalVarGet("gstrCrtRpt_Period")
'40.  qryCourtReport_NY_09_X_27   GlobalVarGet("gstrCrtRpt_Period")
'41.  qryCourtReport_NY_09_X_32   GlobalVarGet("gstrCrtRpt_Period")
'42.  qryCourtReport_NY_10_X_05   GlobalVarGet("gstrCrtRpt_Period")
'43.  qryCourtReport_NY_10_X_14   GlobalVarGet("gstrCrtRpt_Period")
'44.  qryCourtReport_NY_10_X_21   GlobalVarGet("gstrCrtRpt_Period")
'45.  qryCourtReport_NY_10_X_26   GlobalVarGet("gstrCrtRpt_Period")
'46.  qryCourtReport_NY_11_X_04   GlobalVarGet("gstrCrtRpt_Period")
'47.  qryCourtReport_NY_11_X_11   GlobalVarGet("gstrCrtRpt_Period")
'48.  qryCourtReport_NY_12_X_09   GlobalVarGet("gstrCrtRpt_Period")
'49.  qryCourtReport_NY_12_X_15   GlobalVarGet("gstrCrtRpt_Period")
        'DONE!

39390   Debug.Print "'DONE!"
39400   DoEvents

EXITP:
39410   Set qdf = Nothing
39420   Set dbs = Nothing
39430   NY_ListRptTitlesPeriods = blnRetVal
39440   Exit Function

ERRH:
470     blnRetVal = False
480     Select Case ERR.Number
        Case Else
490       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
500     End Select
510     Resume EXITP

End Function

Public Sub AllCancelSet3_NY(blnCancel As Boolean)

39500 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet3_NY"

39510   blnAllCancel = blnCancel

EXITP:
39520   Exit Sub

ERRH:
470     Select Case ERR.Number
        Case Else
480       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
490     End Select
500     Resume EXITP

End Sub
