Attribute VB_Name = "modCourtReportsNY3"
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "modCourtReportsNY3"

'VGC 09/13/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

Private Const MEMO_MAX As Integer = 255

' ** Array: arr_varCap().
Private lngCaps As Long, arr_varCap As Variant
'Private Const C_RID   As Integer = 0
Private Const C_RNAM  As Integer = 1
'Private Const C_CAP   As Integer = 2
Private Const C_CAPN  As Integer = 3

Private blnAllCancel As Boolean, lngTpp As Long
' **

Public Sub Excel00B_Click_NY(strRptCap As String, strRptPathFile As String, strRptPath As String, strRptName As String, lngCaps As Long, arr_varCap As Variant, frm As Access.Form)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Excel00B_Click_NY"

        Dim strRpt As String
        Dim strQry As String
        Dim strMacro As String
        Dim blnUseSavedPath As Boolean, blnContinue As Boolean, blnNoData As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim intRetVal_BuildAssetListInfo As Integer
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

110     blnContinue = True
120     blnUseSavedPath = False

130     With frm

140       DoCmd.Hourglass True
150       DoEvents

160       blnAllCancel = False
170       AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
180       AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
190       blnNoData = False
200       blnAutoStart = .chkOpenExcel

210       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
220         DoCmd.Hourglass False
230         msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
240         If msgResponse <> vbRetry Then
250           blnAllCancel = True
260           AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
270           AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
280           blnContinue = False
290         End If
300       End If

310       If blnContinue = True Then

320         DoCmd.Hourglass True
330         DoEvents

340         If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.

350           .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

360           DoEvents

370           gstrAccountNo = .cmbAccounts.Column(0)
380           gdatStartDate = .DateStart
390           gdatEndDate = .DateEnd
400           gstrCrtRpt_Period = "From " & Format(gdatStartDate, "mm/dd/yyyy") & " To " & Format(gdatEndDate, "mm/dd/yyyy")
              ' ** gstrCrtRpt_Ordinal and gstrCrtRpt_Version should be populated from the input window.

410           gblnMessage = False
420           strTmp01 = "rptCourtRptNY_00B"
430           strQry = "qryCourtReport_NY_00B_X_14"
440           varTmp00 = DCount("*", "qryCourtReport_NY_00B_X_14")
450           If IsNull(varTmp00) = True Then
460             blnNoData = True
470             strQry = "qryCourtReport_NY_00B_X_19"
480           Else
490             If varTmp00 = 0 Then
500               blnNoData = True
510               strQry = "qryCourtReport_NY_00B_X_19"
520             End If
530           End If

540           .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

550           strRptCap = vbNullString: strRptPathFile = vbNullString
560           strRptPath = .UserReportPath
570           strRptName = strTmp01
580           strRpt = vbNullString
590           DoEvents

600           intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, THIS_PROC, frm) ' ** Module Function: modCourtReportsNY1.
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.

610           Select Case intRetVal_BuildAssetListInfo
              Case 0

                ' ** Though strRpt can return either "_00B" or "_00BA", we're not providing the grouped version in New York.
620             If strRpt <> vbNullString Then
630               DoEvents

640               .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.

650               .CapArray_Load  ' ** Form Procedure: frmRpt_CourtReports_NY.
660               DoEvents

670               For lngX = 0& To (lngCaps - 1&)
680                 If arr_varCap(C_RNAM, lngX) = strRptName Then
690                   strRptCap = arr_varCap(C_CAPN, lngX)
700                   Exit For
710                 End If
720               Next
730               DoEvents

740               If IsNull(.UserReportPath) = False Then
750                 If .UserReportPath <> vbNullString Then
760                   If .UserReportPath_chk = True Then
770                     If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
780                       blnUseSavedPath = True
790                     End If
800                   End If
810                 End If
820               End If

830               strMacro = "mcrExcelExport_CR_NY" & Mid(strRptName, InStr(strRptName, "_"))
840               If blnNoData = True Then
850                 strMacro = strMacro & "_nd"
860               End If

870               Select Case blnUseSavedPath
                  Case True
880                 strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
890               Case False
900                 DoCmd.Hourglass False
910                 strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
920               End Select

930               If strRptPathFile <> vbNullString Then
940                 DoCmd.Hourglass True
950                 DoEvents
960                 If gblnPrintAll = True Then blnAutoStart = False
970                 If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
980                   Kill strRptPathFile
990                 End If
1000                If strQry <> vbNullString Then
                      ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                      ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
1010                  DoCmd.RunMacro strMacro
                      ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                      ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
1020                  DoEvents
1030                  If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Or _
                          FileExists(strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
1040                    If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") = True Then
1050                      Name (CurrentAppPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
1060                    Else
1070                      Name (strRptPath & LNK_SEP & "CourtReport_NY_xxx.xls") As (strRptPathFile)
                          ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
1080                    End If
1090                    DoEvents
1100                    If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
1110                      DoEvents
1120                      If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
1130                        EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
1140                      End If
1150                      DoEvents
1160                      If blnAutoStart = True Then
1170                        OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
1180                      End If
1190                    End If
1200                  End If
1210                Else
1220                  DoCmd.OutputTo acOutputReport, strRptName, acFormatXLS, strRptPathFile, blnAutoStart
1230                End If  ' ** strQry.
1240                strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
1250                If strRptPath <> .UserReportPath Then
1260                  .UserReportPath = strRptPath
1270                  SetUserReportPath_NY frm  ' ** Module Procedure: modCourtReportsNY1.
1280                End If
1290              End If  ' ** strRptPathFile.

1300            Else
1310              DoCmd.Hourglass False
1320              Beep
1330              MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
1340            End If  ' ** strRpt.

1350          Case -2
1360            Beep
1370            MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
1380          Case -3, -9
                ' ** Message shown below.
1390          End Select  ' ** intRetVal_BuildAssetListInfo

1400        End If  ' ** Validate().
1410      End If  ' ** blnContinue.
1420    End With

1430    AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
1440    AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

1450    DoCmd.Hourglass False

      #End If

EXITP:
1460    Exit Sub

ERRH:
1470    DoCmd.Hourglass False
1480    Select Case ERR.Number
        Case 70  ' ** Permission denied.
1490      Beep
1500      MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
1510    Case Else
1520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1530    End Select
1540    Resume EXITP

End Sub

Public Sub UserRptPath_After_NY(frm As Access.Form)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "UserRptPath_After_NY"

1610    With frm
1620      Select Case .UserReportPath_chk
          Case True
1630        .UserReportPath_chk_lbl1.FontBold = True
1640        .UserReportPath_chk_lbl1_dim.FontBold = True
1650        .UserReportPath_chk_lbl1_dim_hi.FontBold = True
1660        .UserReportPath_chk_lbl2.FontBold = True
1670        .UserReportPath_chk_lbl2_dim.FontBold = True
1680        .UserReportPath_chk_lbl2_dim_hi.FontBold = True
1690      Case False
1700        .UserReportPath_chk_lbl1.FontBold = False
1710        .UserReportPath_chk_lbl1_dim.FontBold = False
1720        .UserReportPath_chk_lbl1_dim_hi.FontBold = False
1730        .UserReportPath_chk_lbl2.FontBold = False
1740        .UserReportPath_chk_lbl2_dim.FontBold = False
1750        .UserReportPath_chk_lbl2_dim_hi.FontBold = False
1760      End Select
1770    End With

EXITP:
1780    Exit Sub

ERRH:
1790    Select Case ERR.Number
        Case Else
1800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1810    End Select
1820    Resume EXITP

End Sub

Public Sub PrintWordExcelAll_Handler_NY(strProc As String, blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean, blnRebuildTable As Boolean, frm As Access.Form)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "PrintWordExcelAll_Handler_NY"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strRpt As String, strDocName As String
        Dim blnContinue As Boolean
        Dim strEvent As String, strCtlName As String
        Dim msgResponse As VbMsgBoxResult
        Dim intPos01 As Integer, lngCnt As Long
        Dim intRetVal_BuildAssetListInfo As Integer

1910    With frm

1920      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
1930      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
1940      strEvent = Mid(strProc, (intPos01 + 1))
1950      strCtlName = Left(strProc, (intPos01 - 1))

1960      Select Case strEvent
          Case "GotFocus"
1970        Select Case strCtlName
            Case "cmdPrintAll"
1980          blnPrintAll_Focus = True
1990          .cmdPrintAll_box01.Visible = True
2000          .cmdPrintAll_box02.Visible = True
2010          Select Case .chkAssetList
              Case True
2020            .cmdPrintAll_box03.Visible = True
2030          Case False
2040            .cmdPrintAll_box03.Visible = False
2050          End Select
2060          .cmdPrintAll_box04.Visible = True
2070        Case "cmdWordAll"
2080          blnWordAll_Focus = True
2090          .cmdWordAll_box01.Visible = True
2100          .cmdWordAll_box02.Visible = True
2110          Select Case .chkAssetList
              Case True
2120            .cmdWordAll_box03.Visible = True
2130          Case False
2140            .cmdWordAll_box03.Visible = False
2150          End Select
2160          .cmdWordAll_box04.Visible = True
2170        Case "cmdExcelAll"
2180          blnExcelAll_Focus = True
2190          .cmdExcelAll_box01.Visible = True
2200          .cmdExcelAll_box02.Visible = True
2210          Select Case .chkAssetList
              Case True
2220            .cmdExcelAll_box03.Visible = True
2230          Case False
2240            .cmdExcelAll_box03.Visible = False
2250          End Select
2260          .cmdExcelAll_box04.Visible = True
2270        End Select
2280      Case "MouseMove"
2290        Select Case strCtlName
            Case "cmdPrintAll"
2300          If gblnPrintAll = False Then
2310            .cmdPrintAll_box01.Visible = True
2320            .cmdPrintAll_box02.Visible = True
2330            Select Case .chkAssetList
                Case True
2340              .cmdPrintAll_box03.Visible = True
2350            Case False
2360              .cmdPrintAll_box03.Visible = False
2370            End Select
2380            .cmdPrintAll_box04.Visible = True
2390            If blnWordAll_Focus = False Then
2400              .cmdWordAll_box01.Visible = False
2410              .cmdWordAll_box02.Visible = False
2420              .cmdWordAll_box03.Visible = False
2430              .cmdWordAll_box04.Visible = False
2440            End If
2450            If blnExcelAll_Focus = False Then
2460              .cmdExcelAll_box01.Visible = False
2470              .cmdExcelAll_box02.Visible = False
2480              .cmdExcelAll_box03.Visible = False
2490              .cmdExcelAll_box04.Visible = False
2500            End If
2510          End If
2520        Case "cmdWordAll"
2530          .cmdWordAll_box01.Visible = True
2540          .cmdWordAll_box02.Visible = True
2550          Select Case .chkAssetList
              Case True
2560            .cmdWordAll_box03.Visible = True
2570          Case False
2580            .cmdWordAll_box03.Visible = False
2590          End Select
2600          .cmdWordAll_box04.Visible = True
2610          If blnPrintAll_Focus = False Then
2620            .cmdPrintAll_box01.Visible = False
2630            .cmdPrintAll_box02.Visible = False
2640            .cmdPrintAll_box03.Visible = False
2650            .cmdPrintAll_box04.Visible = False
2660          End If
2670          If blnExcelAll_Focus = False Then
2680            .cmdExcelAll_box01.Visible = False
2690            .cmdExcelAll_box02.Visible = False
2700            .cmdExcelAll_box03.Visible = False
2710            .cmdExcelAll_box04.Visible = False
2720          End If
2730        Case "cmdExcelAll"
2740          .cmdExcelAll_box01.Visible = True
2750          .cmdExcelAll_box02.Visible = True
2760          Select Case .chkAssetList
              Case True
2770            .cmdExcelAll_box03.Visible = True
2780          Case False
2790            .cmdExcelAll_box03.Visible = False
2800          End Select
2810          .cmdExcelAll_box04.Visible = True
2820          If blnPrintAll_Focus = False Then
2830            .cmdPrintAll_box01.Visible = False
2840            .cmdPrintAll_box02.Visible = False
2850            .cmdPrintAll_box03.Visible = False
2860            .cmdPrintAll_box04.Visible = False
2870          End If
2880          If blnWordAll_Focus = False Then
2890            .cmdWordAll_box01.Visible = False
2900            .cmdWordAll_box02.Visible = False
2910            .cmdWordAll_box03.Visible = False
2920            .cmdWordAll_box04.Visible = False
2930          End If
2940        End Select
2950      Case "LostFocus"
2960        Select Case strCtlName
            Case "cmdPrintAll"
2970          .cmdPrintAll_box01.Visible = False
2980          .cmdPrintAll_box02.Visible = False
2990          .cmdPrintAll_box03.Visible = False
3000          .cmdPrintAll_box04.Visible = False
3010          blnPrintAll_Focus = False
3020        Case "cmdWordAll"
3030          .cmdWordAll_box01.Visible = False
3040          .cmdWordAll_box02.Visible = False
3050          .cmdWordAll_box03.Visible = False
3060          .cmdWordAll_box04.Visible = False
3070          blnWordAll_Focus = False
3080        Case "cmdExcelAll"
3090          .cmdExcelAll_box01.Visible = False
3100          .cmdExcelAll_box02.Visible = False
3110          .cmdExcelAll_box03.Visible = False
3120          .cmdExcelAll_box04.Visible = False
3130          blnExcelAll_Focus = False
3140        End Select
3150      Case "Click"
3160        Select Case strCtlName
            Case "cmdPrintAll"
3170          DoCmd.Hourglass True
3180          DoEvents
3190          blnContinue = True
3200          If .Validate = True Then  ' ** Form Function: frmRpt_CourtReports_NY.
3210            Beep
3220            DoCmd.Hourglass False
3230            msgResponse = MsgBox("This will send all highlighted reports to your printer." & _
                  vbCrLf & vbCrLf & "Would you like to continue?", vbQuestion + vbOKCancel, _
                  "Send All Reports To Printer")
3240            If msgResponse = vbOK Then
3250              gblnMessage = True
3260              DoCmd.Hourglass False
3270              strDocName = "frmRpt_CourtReports_NY_Input"
3280              DoCmd.OpenForm strDocName, , , , , acDialog, frm.Name
                  ' ** If they cancel, go no further.
3290              If gblnMessage = True Then
3300                strRpt = vbNullString
3310                intRetVal_BuildAssetListInfo = 0
3320                ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
3330                If .chkAssetList = True Then
3340                  intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, THIS_PROC, frm) ' ** Module Function: modCourtReportsNY1.
                      ' ** Return codes:
                      ' **    0  Success.
                      ' **   -2  No data.
                      ' **   -3  Missing entry, e.g., date.
                      ' **   -9  Error.
3350                  Select Case intRetVal_BuildAssetListInfo
                      Case 0
                        ' ** strRpt should return either "_00B" or "_00BA".
3360                    If strRpt <> vbNullString Then
3370                      gdatStartDate = .DateStart
3380                      gdatEndDate = .DateEnd
3390                      gstrAccountNo = .cmbAccounts
3400                      gstrAccountName = .cmbAccounts.Column(3)
3410                      gblnMessage = False
                          ' ** 0B. Asset List.
3420                      DoCmd.OpenReport "rptCourtRptNY" & strRpt, acViewNormal
3430                    Else
3440                      blnContinue = False
3450                      DoCmd.Hourglass False
3460                      MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
3470                    End If
3480                  Case -2
3490                    Set dbs = CurrentDb
                        ' ** Empty tmpAssetList2.
3500                    Set qdf = dbs.QueryDefs("qryCourtReport_03")
3510                    qdf.Execute
3520                    Set qdf = Nothing
3530                    Set dbs = Nothing
3540                    gblnMessage = False
                        ' ** 0B. Asset List.
3550                    DoCmd.OpenReport "rptCourtRptNY_00B", acViewNormal
3560                  Case -3, -9
                        ' ** Message shown below.
3570                  End Select  ' ** intRetVal_BuildAssetListInfo.
3580                End If  ' ** chkAssetList.
3590                If blnContinue = True Then
                      ' ** Let the reports know none of these are in Preview mode.
3600                  gblnMessage = False
                      ' ** Run function to empty and fill the tmpCourtReportData table.
3610                  If .CashAssets_Beg <> vbNullString Then
3620                    .FillVar  ' ** Form Function: frmRpt_CourtReports_NY.
3630                    intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, THIS_PROC, frm) ' ** Module Function: modCourtReportsNY1.
                        ' ** Return codes:
                        ' **    0  Success.
                        ' **   -2  No data.
                        ' **   -3  Missing entry, e.g., date.
                        ' **   -9  Error.
3640                    Select Case intRetVal_BuildAssetListInfo
                        Case 0
                          ' ** 12.
3650                      blnContinue = PreviewOrPrint_NY("12", THIS_PROC & " - R11", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3660                      If blnContinue = True Then
3670                        intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, THIS_PROC, frm) ' ** Module Function: modCourtReportsNY1.
                            ' ** Return codes:
                            ' **    0  Success.
                            ' **   -2  No data.
                            ' **   -3  Missing entry, e.g., date.
                            ' **   -9  Error.
3680                        Select Case intRetVal_BuildAssetListInfo
                            Case 0
3690                          blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
                              ' ** This blnContinue is handled immediately.
3700                          If blnContinue = True Then
                                ' ** 11.
3710                            blnContinue = PreviewOrPrint_NY("11", THIS_PROC & " - R11", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3720                            If blnContinue = True Then
                                  ' ** 10.
3730                              Select Case gblnUseReveuneExpenseCodes
                                  Case True
3740                                blnContinue = PreviewOrPrint_NY("10A", THIS_PROC & " - R10A", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3750                              Case False
3760                                blnContinue = PreviewOrPrint_NY("10", THIS_PROC & " - R10", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3770                              End Select
3780                              If blnContinue = True Then
                                    ' ** 9.
3790                                Select Case gblnUseReveuneExpenseCodes
                                    Case True
3800                                  blnContinue = PreviewOrPrint_NY("9A", THIS_PROC & " - R9A", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3810                                Case False
3820                                  blnContinue = PreviewOrPrint_NY("9", THIS_PROC & " - R9", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3830                                End Select
3840                                If blnContinue = True Then
                                      ' ** 8.
3850                                  blnContinue = PreviewOrPrint_NY("8", THIS_PROC & " - R8", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3860                                  If blnContinue = True Then
3870                                    intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, THIS_PROC, frm) ' ** Module Function: modCourtReportsNY1.
                                        ' ** Return codes:
                                        ' **    0  Success.
                                        ' **   -2  No data.
                                        ' **   -3  Missing entry, e.g., date.
                                        ' **   -9  Error.
3880                                    Select Case intRetVal_BuildAssetListInfo
                                        Case 0
3890                                      gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
3900                                      blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
                                          ' ** This blnContinue is handled immediately.
3910                                      If blnContinue = True Then
3920                                        intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY(.DateStart, .DateEnd, "Ending", strRpt, THIS_PROC, frm) ' ** Module Function: modCourtReportsNY1.
                                            ' ** Return codes:
                                            ' **    0  Success.
                                            ' **   -2  No data.
                                            ' **   -3  Missing entry, e.g., date.
                                            ' **   -9  Error.
3930                                        Select Case intRetVal_BuildAssetListInfo
                                            Case 0
                                              ' ** 7.
3940                                          blnContinue = PreviewOrPrint_NY("7", THIS_PROC & " - R7", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3950                                          If blnContinue = True Then
                                                ' ** 6.
3960                                            blnContinue = PreviewOrPrint_NY("6", THIS_PROC & " - R6", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3970                                            If blnContinue = True Then
                                                  ' ** 5.
3980                                              blnContinue = PreviewOrPrint_NY("5", THIS_PROC & " - R5", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
3990                                              If blnContinue = True Then
                                                    ' ** 4.
4000                                                Select Case gblnUseReveuneExpenseCodes
                                                    Case True
4010                                                  blnContinue = PreviewOrPrint_NY("4A", THIS_PROC & " - R4A", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
4020                                                Case False
4030                                                  blnContinue = PreviewOrPrint_NY("4", THIS_PROC & " - R4", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
4040                                                End Select
4050                                                If blnContinue = True Then
                                                      ' ** 3.
4060                                                  blnContinue = PreviewOrPrint_NY("3", THIS_PROC & " - R3", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
4070                                                  If blnContinue = True Then
                                                        ' ** 2.
4080                                                    blnContinue = PreviewOrPrint_NY("2", THIS_PROC & " - R2", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
4090                                                    If blnContinue = True Then
4100                                                      intRetVal_BuildAssetListInfo = BuildAssetListInfo_NY("01/01/1900", (.DateStart - 1), "Beginning", strRpt, THIS_PROC, frm) ' ** Module Function: modCourtReportsNY1.
                                                          ' ** Return codes:
                                                          ' **    0  Success.
                                                          ' **   -2  No data.
                                                          ' **   -3  Missing entry, e.g., date.
                                                          ' **   -9  Error.
4110                                                      Select Case intRetVal_BuildAssetListInfo
                                                          Case 0
                                                            ' ** 1.
4120                                                        blnContinue = PreviewOrPrint_NY("1", THIS_PROC & " - R1", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
4130                                                        If blnContinue = True Then
4140                                                          gcurCrtRpt_NY_IncomeBeg = Nz(DLookup("icash", "qryCourtReport_NY_00_B_01"), 0)
4150                                                          blnContinue = BuildSummary_NY  ' ** Module Function: modCourtReportsNY1.
                                                              ' ** This blnContinue is handled immediately.
4160                                                          If blnContinue = True Then
                                                                ' ** 0.
4170                                                            Select Case gblnUseReveuneExpenseCodes
                                                                Case True
4180                                                              blnContinue = PreviewOrPrint_NY("0A", THIS_PROC & " - R0A", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
4190                                                            Case False
4200                                                              blnContinue = PreviewOrPrint_NY("0", THIS_PROC & " - R0", acViewNormal, blnRebuildTable, frm)  ' ** Module Function: modCourtReportsNY1.
4210                                                            End Select
4220                                                          End If  ' ** blnContinue.
4230                                                        End If
4240                                                      Case -2
                                                            'SHOULD THIS JUST SPIT OUT AN EMPTY?
4250                                                        Beep
4260                                                        MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
4270                                                      Case -3, -9
                                                            ' ** Message shown below.
4280                                                      End Select  ' ** intRetVal_BuildAssetListInfo
4290                                                    End If  ' ** blnContinue.
4300                                                  End If  ' ** blnContinue.
4310                                                End If  ' ** blnContinue.
4320                                              End If  ' ** blnContinue.
4330                                            End If  ' ** blnContinue.
4340                                          End If  ' ** blnContinue.
4350                                        Case -2
                                              'SHOULD THIS JUST SPIT OUT AN EMPTY?
4360                                          Beep
4370                                          MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
4380                                        Case -3, -9
                                              ' ** Message shown below.
4390                                        End Select  ' ** intRetVal_BuildAssetListInfo
4400                                      End If  ' ** blnContinue.
4410                                    Case -2
                                          'SHOULD THIS JUST SPIT OUT AN EMPTY?
4420                                      Beep
4430                                      MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
4440                                    Case -3, -9
                                          ' ** Message shown below.
4450                                    End Select  ' ** intRetVal_BuildAssetListInfo
4460                                  End If  ' ** blnContinue.
4470                                End If  ' ** blnContinue.
4480                              End If  ' ** blnContinue.
4490                            End If  ' ** blnContinue.
4500                          End If  ' ** blnContinue.
4510                        Case -2
                              'SHOULD THIS JUST SPIT OUT AN EMPTY?
4520                          Beep
4530                          MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
4540                        Case -3, -9
                              ' ** Message shown below.
4550                        End Select  ' ** intRetVal_BuildAssetListInfo
4560                      End If  ' ** blnContinue.
4570                    Case -2
                          'SHOULD THIS JUST SPIT OUT AN EMPTY?
4580                      Beep
4590                      MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
4600                    Case -3, -9
                          ' ** Message shown below.
4610                    End Select  ' ** intRetVal_BuildAssetListInfo
4620                  End If  ' ** CashAssets_Beg.
4630                End If  ' ** blnContinue.
4640              End If  ' ** gblnMessage.
4650            End If  ' ** msgResponse.
4660          End If  ' ** Validate.
4670          DoCmd.Hourglass False
4680        Case "cmdWordAll"
4690          blnAllCancel = False
4700          AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
4710          AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
4720          WordAll_NY frm  ' ** Module Procedure: modCourtReportsNY1.
4730        Case "cmdExcelAll"
4740          blnAllCancel = False
4750          AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
4760          AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.
4770          ExcelAll_NY frm  ' ** Module Procedure: modCourtReportsNY1.
4780        End Select
4790      End Select

4800    End With

EXITP:
4810    Exit Sub

ERRH:
4820    DoCmd.Hourglass False
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Public Sub OpenWordExcel_After_NY(strProc As String, frm As Access.Form)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "OpenWordExcel_After_NY"

        Dim strEvent As String, strCtlName As String
        Dim intPos01 As Integer, lngCnt As Long

4910    With frm

4920      lngCnt = CharCnt(strProc, "_")  ' ** Module Function: modStringFuncs.
4930      intPos01 = CharPos(strProc, lngCnt, "_")  ' ** Module Function: modStringFuncs.
4940      strEvent = Mid(strProc, (intPos01 + 1))
4950      strCtlName = Left(strProc, (intPos01 - 1))

4960      Select Case strCtlName
          Case "chkOpenWord"
4970        Select Case .chkOpenWord
            Case True
4980          .chkOpenWord_lbl.FontBold = True
4990          .chkOpenWord_lbl2.FontBold = True
5000          .chkOpenWord_lbl_dim_hi.FontBold = True
5010          .chkOpenWord_lbl2_dim_hi.FontBold = True
5020        Case False
5030          .chkOpenWord_lbl.FontBold = False
5040          .chkOpenWord_lbl2.FontBold = False
5050          .chkOpenWord_lbl_dim_hi.FontBold = False
5060          .chkOpenWord_lbl2_dim_hi.FontBold = False
5070        End Select
5080      Case "chkOpenExcel"
5090        Select Case .chkOpenExcel
            Case True
5100          .chkOpenExcel_lbl.FontBold = True
5110          .chkOpenExcel_lbl2.FontBold = True
5120          .chkOpenExcel_lbl_dim_hi.FontBold = True
5130          .chkOpenExcel_lbl2_dim_hi.FontBold = True
5140        Case False
5150          .chkOpenExcel_lbl.FontBold = False
5160          .chkOpenExcel_lbl2.FontBold = False
5170          .chkOpenExcel_lbl_dim_hi.FontBold = False
5180          .chkOpenExcel_lbl2_dim_hi.FontBold = False
5190        End Select
5200      End Select

5210    End With

EXITP:
5220    Exit Sub

ERRH:
5230    Select Case ERR.Number
        Case Else
5240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5250    End Select
5260    Resume EXITP

End Sub

Public Function OpenWordExcel_Key_NY(KeyCode As Integer, Shift As Integer, intMode As Integer, frm As Access.Form) As Integer

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "OpenWordExcel_Key_NY"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5320    intShiftDown = (Shift And acShiftMask) > 0
5330    intAltDown = (Shift And acAltMask) > 0
5340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5360      Select Case intRetVal
          Case vbKeyTab
5370        With frm
5380          intRetVal = 0
5390          Select Case intMode
              Case 1
5400            If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
5410              .cmdExcelAll.SetFocus
5420            Else
5430              .cmdClose.SetFocus
5440            End If
5450          Case 2
5460            .cmdClose.SetFocus
5470          End Select
5480        End With
5490      Case vbKeyUp
5500        With frm
5510          intRetVal = 0
5520          Select Case intMode
              Case 1
5530            .cmdWordAll.SetFocus
5540          Case 2
5550            .cmdExcelAll.SetFocus
5560          End Select
5570        End With
5580      Case vbKeyDown
5590        With frm
5600          intRetVal = 0
5610          Select Case intMode
              Case 1
5620            .cmdWord00.SetFocus
5630          Case 2
5640            .cmdExcel00.SetFocus
5650          End Select
5660        End With
5670      End Select
5680    End If

        ' ** Shift keys.
5690    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5700      Select Case intRetVal
          Case vbKeyTab
5710        With frm
5720          intRetVal = 0
5730          Select Case intMode
              Case 1
5740            .cmdWordAll.SetFocus
5750          Case 2
5760            .cmdExcelAll.SetFocus
5770          End Select
5780        End With
5790      End Select
5800    End If

        ' ** Ctrl keys.
5810    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5820      Select Case intRetVal
          Case vbKeyTab
5830        With frm
5840          intRetVal = 0
5850          Select Case intMode
              Case 1
5860            If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
5870              .chkOpenExcel.SetFocus
5880            ElseIf .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
5890              .UserReportPath_chk.SetFocus
5900            Else
5910              Beep
5920            End If
5930          Case 2
5940            If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
5950              .UserReportPath_chk.SetFocus
5960            Else
5970              Beep
5980            End If
5990          End Select
6000        End With
6010      End Select
6020    End If

EXITP:
6030    OpenWordExcel_Key_NY = intRetVal
6040    Exit Function

ERRH:
6050    intRetVal = 0
6060    Select Case ERR.Number
        Case Else
6070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6080    End Select
6090    Resume EXITP

End Function

Public Sub Detail_Mouse_NY(blnCalendar1_Focus As Boolean, blnCalendar2_Focus As Boolean, blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean, frm As Access.Form)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_Mouse_NY"

6110    With frm
6120      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
6130        Select Case blnCalendar1_Focus
            Case True
6140          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
6150          .cmdCalendar1_raised_img.Visible = False
6160        Case False
6170          .cmdCalendar1_raised_img.Visible = True
6180          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6190        End Select
6200        .cmdCalendar1_raised_focus_dots_img.Visible = False
6210        .cmdCalendar1_raised_focus_img.Visible = False
6220        .cmdCalendar1_sunken_focus_dots_img.Visible = False
6230        .cmdCalendar1_raised_img_dis.Visible = False
6240      End If
6250      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
6260        Select Case blnCalendar2_Focus
            Case True
6270          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
6280          .cmdCalendar2_raised_img.Visible = False
6290        Case False
6300          .cmdCalendar2_raised_img.Visible = True
6310          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
6320        End Select
6330        .cmdCalendar2_raised_focus_dots_img.Visible = False
6340        .cmdCalendar2_raised_focus_img.Visible = False
6350        .cmdCalendar2_sunken_focus_dots_img.Visible = False
6360        .cmdCalendar2_raised_img_dis.Visible = False
6370      End If
6380      If blnPrintAll_Focus = False And (.cmdPrintAll_box01.Visible = True Or .cmdPrintAll_box02.Visible = True Or _
              .cmdPrintAll_box03.Visible = True Or .cmdPrintAll_box01.Visible = True) Then
6390        .cmdPrintAll_box01.Visible = False
6400        .cmdPrintAll_box02.Visible = False
6410        .cmdPrintAll_box03.Visible = False
6420        .cmdPrintAll_box04.Visible = False
6430      End If
6440      If blnWordAll_Focus = False And (.cmdWordAll_box01.Visible = True Or .cmdWordAll_box02.Visible = True Or _
              .cmdWordAll_box03.Visible = True Or .cmdWordAll_box01.Visible = True) Then
6450        .cmdWordAll_box01.Visible = False
6460        .cmdWordAll_box02.Visible = False
6470        .cmdWordAll_box03.Visible = False
6480        .cmdWordAll_box04.Visible = False
6490      End If
6500      If blnExcelAll_Focus = False And (.cmdExcelAll_box01.Visible = True Or .cmdExcelAll_box02.Visible = True Or _
              .cmdExcelAll_box03.Visible = True Or .cmdExcelAll_box01.Visible = True) Then
6510        .cmdExcelAll_box01.Visible = False
6520        .cmdExcelAll_box02.Visible = False
6530        .cmdExcelAll_box03.Visible = False
6540        .cmdExcelAll_box04.Visible = False
6550      End If
6560    End With

EXITP:
6570    Exit Sub

ERRH:
6580    Select Case ERR.Number
        Case Else
6590      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6600    End Select
6610    Resume EXITP

End Sub

Public Sub IncExpGrp_After_NY(frm As Access.Form)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "IncExpGrp_After_NY"

6710    With frm
6720      If gblnRevenueExpenseTracking = True Then
6730        If .chkGroupBy_IncExpCode Then
6740          gblnUseReveuneExpenseCodes = True
6750          .chkGroupBy_IncExpCode_lbl.FontBold = True
6760          .chkGroupBy_IncExpCode_lbl4.FontSize = 20
6770          .chkGroupBy_IncExpCode_lbl5.Visible = True
6780          .cmdPreview00_lbl4.FontSize = 20
6790          .cmdPreview00_lbl5.Visible = True
6800          .cmdPreview04_lbl4.FontSize = 20
6810          .cmdPreview04_lbl5.Visible = True
6820          .cmdPreview09_lbl4.FontSize = 20
6830          .cmdPreview09_lbl5.Visible = True
6840          .cmdPreview10_lbl4.FontSize = 20
6850          .cmdPreview10_lbl5.Visible = True
6860        Else
6870          gblnUseReveuneExpenseCodes = False
6880          .chkGroupBy_IncExpCode_lbl.FontBold = False
6890          .chkGroupBy_IncExpCode_lbl4.FontSize = 14
6900          .chkGroupBy_IncExpCode_lbl5.Visible = False
6910          .cmdPreview00_lbl4.FontSize = 14
6920          .cmdPreview00_lbl5.Visible = False
6930          .cmdPreview04_lbl4.FontSize = 14
6940          .cmdPreview04_lbl5.Visible = False
6950          .cmdPreview10_lbl4.FontSize = 14
6960          .cmdPreview10_lbl5.Visible = False
6970          .cmdPreview09_lbl4.FontSize = 14
6980          .cmdPreview09_lbl5.Visible = False
6990        End If
7000      End If
7010    End With

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Public Sub Calendar_Handler_NY(strProc As String, blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean, blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean, clsMonthClass As clsMonthCal, frm As Access.Form)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "Calendar_Handler_NY"

        Dim strEvent As String, strCtlName As String
        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer, intNum As Integer
        Dim blnRetVal As Boolean

7110    With frm

7120      strEvent = Mid(strProc, (CharPos(strProc, 1, "_") + 1))  ' ** Module Function: modStringFuncs.
7130      strCtlName = Left(strProc, (CharPos(strProc, 1, "_") - 1))  ' ** Module Function: modStringFuncs.
7140      intNum = Val(Right(strCtlName, 1))

7150      Select Case strEvent
          Case "Click"
7160        Select Case intNum
            Case 1
7170          datStartDate = Date
7180          datEndDate = 0
7190          blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
7200          If blnRetVal = True Then
7210            .DateStart = datStartDate
7220          Else
7230            .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
7240          End If
7250          .DateStart.SetFocus
7260        Case 2
7270          datStartDate = Date
7280          datEndDate = 0
7290          blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
7300          If blnRetVal = True Then
7310            .DateEnd = datStartDate
7320          Else
7330            .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
7340          End If
7350          .DateEnd.SetFocus
7360          Cancel = 0
7370          .DateEnd_Exit Cancel  ' ** Form Procedure: frmRpt_CourtReports_CA.
7380          If Cancel = 0 Then
7390            .cmbAccounts.SetFocus
7400          End If
7410        End Select
7420      Case "GotFocus"
7430        Select Case intNum
            Case 1
7440          blnCalendar1_Focus = True
7450          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
7460          .cmdCalendar1_raised_img.Visible = False
7470          .cmdCalendar1_raised_focus_img.Visible = False
7480          .cmdCalendar1_raised_focus_dots_img.Visible = False
7490          .cmdCalendar1_sunken_focus_dots_img.Visible = False
7500          .cmdCalendar1_raised_img_dis.Visible = False
7510        Case 2
7520          blnCalendar2_Focus = True
7530          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
7540          .cmdCalendar2_raised_img.Visible = False
7550          .cmdCalendar2_raised_focus_img.Visible = False
7560          .cmdCalendar2_raised_focus_dots_img.Visible = False
7570          .cmdCalendar2_sunken_focus_dots_img.Visible = False
7580          .cmdCalendar2_raised_img_dis.Visible = False
7590        End Select
7600      Case "MouseDown"
7610        Select Case intNum
            Case 1
7620          blnCalendar1_MouseDown = True
7630          .cmdCalendar1_sunken_focus_dots_img.Visible = True
7640          .cmdCalendar1_raised_img.Visible = False
7650          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
7660          .cmdCalendar1_raised_focus_img.Visible = False
7670          .cmdCalendar1_raised_focus_dots_img.Visible = False
7680          .cmdCalendar1_raised_img_dis.Visible = False
7690        Case 2
7700          blnCalendar2_MouseDown = True
7710          .cmdCalendar2_sunken_focus_dots_img.Visible = True
7720          .cmdCalendar2_raised_img.Visible = False
7730          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
7740          .cmdCalendar2_raised_focus_img.Visible = False
7750          .cmdCalendar2_raised_focus_dots_img.Visible = False
7760          .cmdCalendar2_raised_img_dis.Visible = False
7770        End Select
7780      Case "MouseMove"
7790        Select Case intNum
            Case 1
7800          If blnCalendar1_MouseDown = False Then
7810            Select Case blnCalendar1_Focus
                Case True
7820              .cmdCalendar1_raised_focus_dots_img.Visible = True
7830              .cmdCalendar1_raised_focus_img.Visible = False
7840            Case False
7850              .cmdCalendar1_raised_focus_img.Visible = True
7860              .cmdCalendar1_raised_focus_dots_img.Visible = False
7870            End Select
7880            .cmdCalendar1_raised_img.Visible = False
7890            .cmdCalendar1_raised_semifocus_dots_img.Visible = False
7900            .cmdCalendar1_sunken_focus_dots_img.Visible = False
7910            .cmdCalendar1_raised_img_dis.Visible = False
7920          End If
7930        Case 2
7940          If blnCalendar2_MouseDown = False Then
7950            Select Case blnCalendar2_Focus
                Case True
7960              .cmdCalendar2_raised_focus_dots_img.Visible = True
7970              .cmdCalendar2_raised_focus_img.Visible = False
7980            Case False
7990              .cmdCalendar2_raised_focus_img.Visible = True
8000              .cmdCalendar2_raised_focus_dots_img.Visible = False
8010            End Select
8020            .cmdCalendar2_raised_img.Visible = False
8030            .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8040            .cmdCalendar2_sunken_focus_dots_img.Visible = False
8050            .cmdCalendar2_raised_img_dis.Visible = False
8060          End If
8070        End Select
8080      Case "MouseUp"
8090        Select Case intNum
            Case 1
8100          .cmdCalendar1_raised_focus_dots_img.Visible = True
8110          .cmdCalendar1_raised_img.Visible = False
8120          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
8130          .cmdCalendar1_raised_focus_img.Visible = False
8140          .cmdCalendar1_sunken_focus_dots_img.Visible = False
8150          .cmdCalendar1_raised_img_dis.Visible = False
8160          blnCalendar1_MouseDown = False
8170        Case 2
8180          .cmdCalendar2_raised_focus_dots_img.Visible = True
8190          .cmdCalendar2_raised_img.Visible = False
8200          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8210          .cmdCalendar2_raised_focus_img.Visible = False
8220          .cmdCalendar2_sunken_focus_dots_img.Visible = False
8230          .cmdCalendar2_raised_img_dis.Visible = False
8240          blnCalendar2_MouseDown = False
8250        End Select
8260      Case "LostFocus"
8270        Select Case intNum
            Case 1
8280          .cmdCalendar1_raised_img.Visible = True
8290          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
8300          .cmdCalendar1_raised_focus_img.Visible = False
8310          .cmdCalendar1_raised_focus_dots_img.Visible = False
8320          .cmdCalendar1_sunken_focus_dots_img.Visible = False
8330          .cmdCalendar1_raised_img_dis.Visible = False
8340          blnCalendar1_Focus = False
8350        Case 2
8360          .cmdCalendar2_raised_img.Visible = True
8370          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8380          .cmdCalendar2_raised_focus_img.Visible = False
8390          .cmdCalendar2_raised_focus_dots_img.Visible = False
8400          .cmdCalendar2_sunken_focus_dots_img.Visible = False
8410          .cmdCalendar2_raised_img_dis.Visible = False
8420          blnCalendar2_Focus = False
8430        End Select
8440      End Select

8450    End With

EXITP:
8460    Exit Sub

ERRH:
8470    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
8480    Case Else
8490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8500    End Select
8510    Resume EXITP

End Sub

Public Sub OpenProcs_NY(intMode As Integer, blnCancelAll As Boolean, blnRebuildTable As Boolean, strFootnoteDefault As String, Cancel As Integer, frm As Access.Form)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "OpenProcs_NY"

        Dim blnRetVal As Boolean

8610    With frm
8620      Select Case intMode
          Case 1
            ' ** Check some Public variables in case there's been an error.
8630        blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
8640        blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
8650      Case 2
8660        blnCancelAll = False
8670        blnAllCancel = blnCancelAll
8680        AllCancelSet2_NY blnCancelAll  ' ** Module Procedure: modCourtReportsNY1.
8690        AllCancelSet3_NY blnCancelAll  ' ** Module Procedure: modCourtReportsNY2.
8700        blnAllCancel = blnCancelAll
8710        blnRebuildTable = True
8720      Case 3
            ' ** EVENT CHECK: chkRememberMe!
8730        .cmbAccounts_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
            ' ** EVENT CHECK: chkRememberDates!
8740        If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
8750          .DateStart_Exit Cancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
8760          .DateEnd_Exit Cancel  ' ** Form Procedure: frmRpt_CourtReports_NY.
8770        End If
8780        .UserReportPath_chk_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8790        SetUserReportPath_NY frm  ' ** Module Procedure: modCourtReportsNY1.
8800        .chkGroupBy_IncExpCode_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8810        .chkLegalName_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8820        .opgAccountSource_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8830        .chkRememberMe_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8840        .chkRememberDates_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8850        .chkAssetList_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8860        .chkPageOf_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8870        .chkOpenWord_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8880        .chkOpenExcel_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8890        .chkIncludeFootnote_AfterUpdate  ' ** Form Procedure: frmRpt_CourtReports_NY.
8900        If IsNull(.CourtReports_Footnote) = True Then
8910          .CourtReports_Footnote = strFootnoteDefault
8920        End If
8930      End Select
8940    End With

EXITP:
8950    Exit Sub

ERRH:
8960    DoCmd.Hourglass False
8970    Select Case ERR.Number
        Case Else
8980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8990    End Select
9000    Resume EXITP

End Sub

Public Sub AcctSource_After_NY(frm As Access.Form)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "AcctSource_After_NY"

        Dim strAccountNo As String

9110    strAccountNo = vbNullString

9120    With frm
9130      If IsNull(.cmbAccounts) = False Then
9140        If Len(.cmbAccounts.Column(0)) > 0 Then
9150          strAccountNo = .cmbAccounts.Column(0)
9160        End If
9170      End If
9180      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
9190        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
9200        .opgAccountSource_optNumber_lbl.FontBold = True
9210        .opgAccountSource_optName_lbl.FontBold = False
9220      Case .opgAccountSource_optName.OptionValue
9230        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
9240        .opgAccountSource_optNumber_lbl.FontBold = False
9250        .opgAccountSource_optName_lbl.FontBold = True
9260      End Select
9270      DoEvents
9280      If strAccountNo <> vbNullString Then
9290        .cmbAccounts = strAccountNo
9300      End If
9310    End With

EXITP:
9320    Exit Sub

ERRH:
9330    Select Case ERR.Number
        Case Else
9340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9350    End Select
9360    Resume EXITP

End Sub

Public Sub Feet_After_NY(intMode As Integer, strFootnoteDefault As String, frm As Access.Form)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "Feet_After_NY"

        Dim varTmp00 As Variant

9410    With frm
9420      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
9430        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
9440      End If
9450      Select Case intMode
          Case 1
9460        .CourtReports_Footnote.Visible = True
9470        .CourtReports_Footnote_box.Visible = True
9480        .CourtReports_Footnote_lbl2.Visible = False
9490        Select Case .chkIncludeFootnote
            Case True
9500          .chkIncludeFootnote_lbl.FontBold = True
9510          .CourtReports_Footnote.Enabled = True
9520          .CourtReports_Footnote.BackStyle = acBackStyleNormal
9530          .CourtReports_Footnote_lbl.BackStyle = acBackStyleNormal
9540          .CourtReports_Footnote.BorderColor = CLR_LTBLU2
9550          If IsNull(.CourtReports_Footnote) = True Then
9560            .CourtReports_Footnote = strFootnoteDefault
9570          End If
9580          If gblnRevenueExpenseTracking = False Then
9590            .chkGroupBy_IncExpCode_lbl3.Visible = False
9600          End If
9610        Case False
9620          .chkIncludeFootnote_lbl.FontBold = False
9630          .CourtReports_Footnote.Enabled = False
9640          .CourtReports_Footnote.BackStyle = acBackStyleTransparent
9650          .CourtReports_Footnote_lbl.BackStyle = acBackStyleTransparent
9660          .CourtReports_Footnote.BorderColor = WIN_CLR_DISR
9670          If gblnRevenueExpenseTracking = False Then
9680            .chkGroupBy_IncExpCode_lbl3.Visible = True
9690          End If
9700        End Select
9710      Case 2
9720        varTmp00 = .CourtReports_Footnote
9730        If IsNull(varTmp00) = False Then
9740          If Trim(varTmp00) <> vbNullString Then
9750            varTmp00 = Trim(varTmp00)
9760            If Len(varTmp00) > MEMO_MAX Then
9770              Beep
9780              If .CourtReports_Footnote_lbl2.Visible = False Then
9790                .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
9800                .CourtReports_Footnote_lbl2.Visible = True
9810              End If
9820              varTmp00 = Left(varTmp00, MEMO_MAX)
9830              .CourtReports_Footnote = varTmp00
9840            End If
9850          Else
9860            .CourtReports_Footnote = Null
9870          End If
9880        End If
9890      End Select
9900    End With

EXITP:
9910    Exit Sub

ERRH:
9920    Select Case ERR.Number
        Case Else
9930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9940    End Select
9950    Resume EXITP

End Sub

Public Function Feet_Key_NY(KeyCode As Integer, Shift As Integer, frm As Access.Form) As Integer

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "Feet_Key_NY"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intLen As Integer
        Dim intRetVal As Integer

10010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10020   intShiftDown = (Shift And acShiftMask) > 0
10030   intAltDown = (Shift And acAltMask) > 0
10040   intCtrlDown = (Shift And acCtrlMask) > 0

10050   If lngTpp = 0& Then
          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
10060     lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10070   End If

        ' ** Plain keys.
10080   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10090     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10100       With frm
10110         intRetVal = 0
10120         .chkPageOf.SetFocus
10130       End With
10140     Case Else
10150       With frm
10160         If IsNull(.CourtReports_Footnote.text) = False Then
10170           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
10180             intLen = Len(Trim(.CourtReports_Footnote.text))
10190             If intLen > MEMO_MAX Then
10200               intRetVal = 0
10210               Beep
10220               If .CourtReports_Footnote_lbl2.Visible = False Then
10230                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
10240                 .CourtReports_Footnote_lbl2.Visible = True
10250               End If
10260               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
10270               .CourtReports_Footnote.SelStart = 999
10280             End If
10290           End If
10300         End If
10310       End With
10320     End Select
10330   End If

        ' ** Shift keys.
10340   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10350     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10360       With frm
10370         intRetVal = 0
10380         .chkIncludeFootnote.SetFocus
10390       End With
10400     Case Else
10410       With frm
10420         If IsNull(.CourtReports_Footnote.text) = False Then
10430           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
10440             intLen = Len(Trim(.CourtReports_Footnote.text))
10450             If intLen > MEMO_MAX Then
10460               intRetVal = 0
10470               Beep
10480               If .CourtReports_Footnote_lbl2.Visible = False Then
10490                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
10500                 .CourtReports_Footnote_lbl2.Visible = True
10510               End If
10520               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
10530               .CourtReports_Footnote.SelStart = 999
10540             End If
10550           End If
10560         End If
10570       End With
10580     End Select
10590   End If

EXITP:
10600   Feet_Key_NY = intRetVal
10610   Exit Function

ERRH:
10620   intRetVal = 0
10630   Select Case ERR.Number
        Case Else
10640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10650   End Select
10660   Resume EXITP

End Function
