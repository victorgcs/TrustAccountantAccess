VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_FL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_FL"

'VGC 10/09/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  39120
' #########################

' ** For GTR, sets options, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Guardian of Prop: D {opgType_optGuard}
' **   Excel All:        E {cmdExcelAll}
' **   From:             F {DateStart}
' **   Legal Name:       G {chkLegalName}
' **   Name:             M {opgAccountSource_optName}
' **   Footnote:         N {chkIncludeFootnote}
' **   Print All:        P {cmdPrintAll}
' **   Incl w/Summary:   S {chkAssetList}
' **   To:               T {DateEnd}
' **   Group Rev/Exp:    U {chkGroupBy_IncExpCode}
' **   Personal Rep:     V {opgType_optRep}
' **   Word All:         W {cmdWordAll}
' **   Exit:             X {cmdClose}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}
' **   Open Excel:       N {chkOpenExcel}
' **   Open Word:        O {chkOpenWord}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

' ** Array: arr_varCap().
Private lngCaps As Long, arr_varCap As Variant
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

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const MEMO_MAX As Integer = 255

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean
Private blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean

Private lngDisburse_Width As Long, lngDistribute_Top As Long, lngDistribute_Width As Long, lngDoubleOffset_Top As Long
Private lngDisburseStarOffset_Left As Long, lngDisburseStarOffset_Top As Long, lngScheduleOfset_Top As Long
Private lngButtonOffset_Top As Long, lngBox4_Height As Long

Private blnGoingToReport As Boolean, blnGoingToReport2 As Boolean, blnWasGTR As Boolean
Private lngBuildCnt As Long, intSpecSQLCnt As Integer
Private blnIsSummary As Boolean, blnAllCancel As Boolean, blnRebuildTable As Boolean, blnAccountChecked As Boolean
Private strFootnoteDefault As String, strLastArrow As String
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
160           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False

170           blnWasGTR = gblnGoToReport

              'lngTpp = GetTPP  ' ** Module Procedure: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           CLR_DISABLED_FG = CLR_DKGRY
200           CLR_DISABLED_BG = CLR_LTTEAL

210           blnPrintAll_Focus = False: blnWordAll_Focus = False: blnExcelAll_Focus = False

220           gblnLegalName = False
230           gblnUseReveuneExpenseCodes = False
240           gvarCrtRpt_FL_SpecData = Null
250           intSpecSQLCnt = 0: lngBuildCnt = 0&
260           strFootnoteDefault = .CourtReports_Footnote.DefaultValue
270           If Right(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Left(strFootnoteDefault, (Len(strFootnoteDefault) - 1))
280           If Left(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Mid(strFootnoteDefault, 2)

              ' ** Check some Public variables in case there's been an error.
290           blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
300           blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.

310           gstrFormQuerySpec = THIS_NAME
320           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
330           gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
340           gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString

350           blnAllCancel = False
360           AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
370           blnRebuildTable = True
380           blnAccountChecked = False

390           lngDisburse_Width = .cmdPreview02_lbl.Width
400           lngDisburseStarOffset_Left = (.cmdPreview02_lbl4.Left - (.cmdPreview02_lbl.Left + .cmdPreview02_lbl.Width))
410           lngDisburseStarOffset_Top = (.cmdPreview02_lbl.Top - .cmdPreview02_lbl4.Top)
420           lngDistribute_Top = .cmdPreview03_lbl.Top
430           lngDistribute_Width = .cmdPreview03_lbl.Width
440           lngScheduleOfset_Top = (.cmdPreview04_Sch_lbl.Top - .cmdPreview04_lbl.Top)
450           lngButtonOffset_Top = (.cmdPreview04.Top - .cmdPreview04_Sch_lbl.Top)
460           lngDoubleOffset_Top = (.cmdPreview05_lbl.Top - (.cmdPreview04_lbl_box.Top + .cmdPreview04_lbl_box.Height))
470           lngBox4_Height = .cmdPrintAll_box04.Height

480           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
490           DoEvents

              ' ** If revenue / expense tracking is checked on the options screen,
              ' ** let user select reports to be grouped by revenue/expense codes.
500           Select Case gblnRevenueExpenseTracking
              Case True
510             .chkGroupBy_IncExpCode.Enabled = True
520             .chkGroupBy_IncExpCode_lbl3.Visible = False     ' ** Off note.
530             .chkGroupBy_IncExpCode_lbl4_txt.Locked = True   ' ** Asterisk
540             .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
550             .cmdPreview00_lbl4_txt.Locked = True
560             .cmdPreview00_lbl4_cmd.Enabled = True
570             .cmdPreview01_lbl4_txt.Locked = True
580             .cmdPreview01_lbl4_cmd.Enabled = True
590             .cmdPreview02_lbl4_txt.Locked = True
600             .cmdPreview02_lbl4_cmd.Enabled = True
610           Case False
620             .chkGroupBy_IncExpCode.Enabled = False
630             .chkGroupBy_IncExpCode_lbl3.Visible = True      ' ** Off note.
640             .chkGroupBy_IncExpCode_lbl4_txt.Locked = False  ' ** Asterisk
650             .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
660             .cmdPreview00_lbl4_txt.Locked = False  ' ** So they look disabled.
670             .cmdPreview00_lbl4_cmd.Enabled = False
680             .cmdPreview01_lbl4_txt.Locked = False
690             .cmdPreview01_lbl4_cmd.Enabled = False
700             .cmdPreview02_lbl4_txt.Locked = False
710             .cmdPreview02_lbl4_cmd.Enabled = False
720           End Select

730           If gdatStartDate > 0 And gdatEndDate > 0 Then
740             .DateStart = gdatStartDate
750             .DateEnd = gdatEndDate
760           End If

770           Select Case .opgType
              Case .opgType_optRep.OptionValue
780             .opgType_optRep_lbl.FontBold = True
790             .opgType_optGuard_lbl.FontBold = False
800           Case opgType_optGuard.OptionValue
810             .opgType_optRep_lbl.FontBold = False
820             .opgType_optGuard_lbl.FontBold = True
830           End Select

      #If NoExcel Then
840           .cmdExcel00.Enabled = False
850           .cmdExcel00B.Enabled = False
860           .cmdExcel01.Enabled = False
870           .cmdExcel02.Enabled = False
880           .cmdExcel03.Enabled = False
890           .cmdExcel04.Enabled = False
900           .cmdExcel05.Enabled = False
910           .cmdExcelAll.Enabled = False
920           .chkOpenExcel.Enabled = False
930           .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
940           .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
950           .cmdExcel00.Enabled = True
960           .cmdExcel00B.Enabled = True
970           .cmdExcel01.Enabled = True
980           .cmdExcel02.Enabled = True
990           .cmdExcel03.Enabled = True
1000          .cmdExcel04.Enabled = True
1010          .cmdExcel05.Enabled = True
1020          .cmdExcelAll.Enabled = True
1030          .chkOpenExcel.Enabled = True
1040          .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
1050          .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

              ' ** EVENT CHECK: chkRememberMe!
1060          cmbAccounts_AfterUpdate  ' ** Procedure: Below.
              ' ** EVENT CHECK: chkRememberDates!
1070          If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
1080            DateStart_Exit Cancel  ' ** Procedure: Below.
1090            DateEnd_Exit Cancel  ' ** Procedure: Below.
1100          End If

1110          UserReportPath_chk_AfterUpdate  ' ** Procedure: Below.
1120          SetUserReportPath  ' ** Procedure: Below.
1130          opgType_AfterUpdate  ' ** Procedure: Below.
1140          chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Below.
1150          chkLegalName_AfterUpdate  ' ** Procedure: Below.
1160          opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1170          chkAssetList_AfterUpdate  ' ** Procedure: Below.
1180          chkRememberMe_AfterUpdate  ' ** Procedure: Below.
1190          chkRememberDates_AfterUpdate  ' ** Procedure: Below.
1200          chkPageOf_AfterUpdate  ' ** Procedure: Below.

1210          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
1220          chkOpenExcel_AfterUpdate  ' ** Procedure: Below.

1230          chkIncludeFootnote_AfterUpdate  ' ** Procedure: Below.
1240          If IsNull(.CourtReports_Footnote) = True Then
1250            .CourtReports_Footnote = strFootnoteDefault
1260          End If

1270          DoCmd.Hourglass False

1280          .DateStart.SetFocus

1290        Else
1300          Cancel = -1
1310        End If
1320      Else
1330        Cancel = -1
1340      End If
1350    End With

1360    If Cancel = -1 Then
1370      Beep
1380      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1390    End If

EXITP:
1400    Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Private Sub Form_Load()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1510    With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
1520      Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
1530      clsMonthClass.hWndForm = .hwnd
          ' ** Let's default to PositionAtCursor.
1540      clsMonthClass.PositionAtCursor = True

1550    End With

EXITP:
1560    Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strCtl As String, strNum As String
        Dim blnTmp01 As Boolean
        Dim intRetVal As Integer

1610    intRetVal = KeyCode
1620    blnTmp01 = False

        ' ** Use bit masks to determine which key was pressed.
1630    intShiftDown = (Shift And acShiftMask) > 0
1640    intAltDown = (Shift And acAltMask) > 0
1650    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Incl w/Summary:   S {chkAssetList}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}
        ' **   Open Excel:       N {chkOpenExcel}
        ' **   Open Word:        O {chkOpenWord}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}

1660  On Error Resume Next
1670    strCtl = Screen.ActiveControl.Name
1680  On Error GoTo ERRH

        ' ** Plain keys.
1690    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1700      Select Case intRetVal
          Case vbKeyF7
1710        intRetVal = 0
1720        cmdCalendar1_Click  ' ** Procedure: Below.
1730      Case vbKeyF8
1740        intRetVal = 0
1750        cmdCalendar2_Click  ' ** Procedure: Below.
1760      Case vbKeyUp
1770        With Me
1780          If strCtl <> vbNullString Then
1790            If Right(strCtl, 1) = "B" Then blnTmp01 = True
1800            strNum = Right(IIf(blnTmp01 = True, Left(strCtl, (Len(strCtl) - 1)), strCtl), 2)
1810            If IsNumeric(strNum) = True Then
                  ' ** It's one of the main report buttons.
1820              intRetVal = 0
1830              If strNum <> "00" Or blnTmp01 = True Then
1840                strLastArrow = vbNullString
1850                Select Case strNum
                    Case "00"
1860                  strCtl = Left(strCtl, (Len(strCtl) - 1))
1870                Case "01"
1880                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) - 1)), 2) & "B"
1890                Case Else
1900                  Select Case .opgType
                      Case .opgType_optRep.OptionValue
1910                    strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) - 1)), 2)
1920                  Case .opgType_optGuard.OptionValue
1930                    If strNum <> "04" Then
1940                      strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) - 1)), 2)
1950                    Else
1960                      strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) - 2)), 2)
1970                    End If
1980                  End Select
1990                End Select
2000              Else
2010                strLastArrow = strCtl
2020                strCtl = Mid(strCtl, 4, 4)
2030                Select Case strCtl
                    Case "Prev"
2040                  strCtl = "cmdPreview05"
2050                Case "Prin"
2060                  strCtl = "cmdPrintAll"
2070                Case "Word"
2080                  strCtl = "cmdWordAll"
2090                Case "Exce"
2100                  strCtl = "cmdExcelAll"
2110                End Select
2120              End If
2130              .Controls(strCtl).SetFocus
2140            End If
2150          End If
2160        End With
2170      Case vbKeyDown
2180        With Me
2190          If strCtl <> vbNullString Then
2200            If Right(strCtl, 1) = "B" Then blnTmp01 = True
2210            strNum = Right(IIf(blnTmp01 = True, Left(strCtl, (Len(strCtl) - 1)), strCtl), 2)
2220            If IsNumeric(strNum) = True Then
                  ' ** It's one of the main report buttons.
2230              intRetVal = 0
2240              If strNum <> "05" Then
2250                strLastArrow = vbNullString
2260                Select Case strNum
                    Case "00"
                      ' ** cmdPreview00B, cmdPrint00B, cmdWord00B, cmdExcel00B.
2270                  Select Case blnTmp01
                      Case True
2280                    strCtl = Left(strCtl, (Len(strCtl) - 3)) & Right("00" & CStr(Val("01")), 2)
2290                  Case False
2300                    strCtl = strCtl & "B"
2310                  End Select
2320                Case Else
2330                  Select Case .opgType
                      Case .opgType_optRep.OptionValue
2340                    strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) + 1)), 2)
2350                  Case .opgType_optGuard.OptionValue
2360                    If strNum <> "02" Then
2370                      strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) + 1)), 2)
2380                    Else
2390                      strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) + 2)), 2)
2400                    End If
2410                  End Select
2420                End Select
2430              Else
2440                strLastArrow = strCtl
2450                strCtl = Mid(strCtl, 4, 4)
2460                Select Case strCtl
                    Case "Prev"
2470                  strCtl = "cmdPreview00"
2480                Case "Prin"
2490                  strCtl = "cmdPrintAll"
2500                Case "Word"
2510                  strCtl = "cmdWordAll"
2520                Case "Exce"
2530                  strCtl = "cmdExcelAll"
2540                End Select
2550              End If
2560              .Controls(strCtl).SetFocus
2570            End If
2580          End If
2590        End With
2600      Case vbKeyEscape
2610        intRetVal = 0
2620        cmdClose_Click  ' ** Procedure: Below.
2630      End Select
2640    End If

        ' ** Alt keys.
2650    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyS
2670        With Me
2680          intRetVal = 0
2690          .chkAssetList.SetFocus
2700          .chkAssetList = (Not .chkAssetList)
2710          chkAssetList_AfterUpdate  ' ** Procedure: Below.
2720        End With
2730      Case vbKeyX
2740        intRetVal = 0
2750        cmdClose_Click  ' ** Procedure: Below.
2760      End Select
2770    End If

        ' ** Ctrl keys.
2780    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2790      Select Case intRetVal
          Case vbKeyG
2800        With Me
2810          intRetVal = 0
2820          If .chkPageOf.Enabled = True Then
2830            .chkPageOf.SetFocus
2840            .chkPageOf = (Not .chkPageOf)
2850            chkPageOf_AfterUpdate  ' ** Procedure: Below.
2860          End If
2870        End With
2880      Case vbKeyN
2890        With Me
2900          intRetVal = 0
2910          If .chkOpenExcel.Enabled = True Then
2920            .chkOpenExcel.SetFocus
2930            .chkOpenExcel = (Not .chkOpenExcel)
2940            chkOpenExcel_AfterUpdate  ' ** Procedure: Below.
2950          Else
2960            Beep
2970          End If
2980        End With
2990      Case vbKeyO
3000        With Me
3010          intRetVal = 0
3020          .chkOpenWord.SetFocus
3030          .chkOpenWord = (Not .chkOpenWord)
3040          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
3050        End With
3060      End Select
3070    End If

        ' ** Ctrl-Shift keys.
3080    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3090      Select Case intRetVal
          Case vbKeyT
3100        intRetVal = 0
3110        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3120      Case vbKeyF
3130        With Me
3140          intRetVal = 0
3150          .FocusHolder.SetFocus
3160        End With
3170      End Select
3180    End If

EXITP:
3190    KeyCode = intRetVal
3200    Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub Form_Timer()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngBtn_Top As Long, lngBtn_Left As Long

3310    With Me
3320      .TimerInterval = 0&
3330      If gblnSetFocus = True Then
3340        gblnSetFocus = False
3350        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
3360          Forms("frmLoadTimer").Modal = False
3370        End If
3380        strFormName = vbNullString
3390  On Error Resume Next
3400        Do Until strFormName = THIS_NAME
3410          DoCmd.SelectObject acForm, THIS_NAME, False
3420          strFormName = Screen.ActiveForm.Name
3430        Loop
3440  On Error GoTo ERRH
3450      End If
3460      If gblnGoToReport = True Then
3470        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Procedure: modWindowFunctions.
3480          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
3490        End If
3500        Select Case blnGoingToReport
            Case True
3510          Select Case blnGoingToReport2
              Case True
3520            DoCmd.Hourglass False  ' ** It should already be off.
3530            blnGoingToReport = False
3540            blnGoingToReport2 = False
3550            gblnGoToReport = False
3560            .GoToReport_arw_rpt_img.Visible = False
3570            DoEvents
3580            If garr_varGoToReport(GTR_CTL6) = "cmdPrint02" Then
3590              .cmdPreview02_lbl4_txt.Visible = True
3600              Select Case .chkGroupBy_IncExpCode
                  Case True
3610                .cmdPreview02_lbl5.Visible = True
3620              Case False
3630                .cmdPreview02_lbl5.Visible = False
3640              End Select
3650            End If
3660            DoEvents
3670          Case False
3680            Select Case garr_varGoToReport(GTR_CTL6)
                Case "cmdPrint00"
3690              DoCmd.Hourglass False  ' ** It should already be off.
3700              blnGoingToReport = False
3710              blnGoingToReport2 = False
3720              gblnGoToReport = False
3730              .GoToReport_arw_rpt_img.Visible = False
3740              DoEvents
3750              .cmdPreview00_lbl4_txt.Visible = True
3760              Select Case .chkGroupBy_IncExpCode
                  Case True
3770                .cmdPreview00_lbl5.Visible = True
3780              Case False
3790                .cmdPreview00_lbl5.Visible = False
3800              End Select
3810            Case "cmdPrint00B"
3820              DoCmd.Hourglass False
3830              blnGoingToReport = False
3840              blnGoingToReport2 = False
3850              gblnGoToReport = False
3860              .GoToReport_arw_rpt_img.Visible = False
3870              DoEvents
3880              .GoToReport_arw_alist_on_img.Visible = False
3890              .GoToReport_arw_alist_off_img.Visible = False
3900              .chkAssetList.Visible = True
3910              .chkAssetList_lbl2.Visible = True
3920              .chkAssetList_box.Visible = True
3930            Case "cmdPrint01"
3940              DoCmd.Hourglass False
3950              blnGoingToReport = False
3960              blnGoingToReport2 = False
3970              gblnGoToReport = False
3980              .GoToReport_arw_rpt_img.Visible = False
3990              DoEvents
4000              .cmdPreview01_lbl4_txt.Visible = True
4010              Select Case .chkGroupBy_IncExpCode
                  Case True
4020                .cmdPreview01_lbl5.Visible = True
4030              Case False
4040                .cmdPreview01_lbl5.Visible = False
4050              End Select
4060            Case "cmdPrint02"
4070              blnGoingToReport2 = True
4080              lngBtn_Top = (.cmdPreview02_lbl.Top + (lngTpp))
4090              lngBtn_Left = (.cmdPreview02_lbl.Left + .cmdPreview02_lbl.Width)
4100              .GoToReport_arw_rep_img.Visible = False
4110              .GoToReport_arw_grd_img.Visible = False
4120              .cmdPreview02_lbl4_txt.Visible = False
4130              .cmdPreview02_lbl5.Visible = False
4140              .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4150              .GoToReport_arw_rpt_img.Top = lngBtn_Top
4160              .GoToReport_arw_rpt_img.Left = lngBtn_Left
4170              .GoToReport_arw_rpt_img.Visible = True
4180              DoEvents
4190              Beep
4200              DoCmd.Hourglass False
4210              DoEvents
4220              .TimerInterval = (GTR_WAIT * 2&)
4230            Case "cmdPrint03"
4240              blnGoingToReport2 = True
4250              lngBtn_Top = (.cmdPreview03_lbl.Top + (lngTpp))
4260              lngBtn_Left = (.cmdPreview03_lbl.Left + .cmdPreview03_lbl.Width)
4270              .GoToReport_arw_rep_img.Visible = False
4280              .GoToReport_arw_grd_img.Visible = False
4290              .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4300              .GoToReport_arw_rpt_img.Top = lngBtn_Top
4310              .GoToReport_arw_rpt_img.Left = lngBtn_Left
4320              .GoToReport_arw_rpt_img.Visible = True
4330              DoEvents
4340              Beep
4350              DoCmd.Hourglass False
4360              DoEvents
4370              .TimerInterval = (GTR_WAIT * 2&)
4380            Case "cmdPrint05"
4390              DoCmd.Hourglass False
4400              blnGoingToReport = False
4410              blnGoingToReport2 = False
4420              gblnGoToReport = False
4430              .GoToReport_arw_rpt_img.Visible = False
4440              DoEvents
4450              .cmdPreview05_Sch_lbl.Visible = True
4460              .GoToReport_arw_alist_e_img.Visible = False
4470            End Select
4480            DoEvents
4490          End Select
4500        Case False
4510          DoCmd.Hourglass True  ' ** Make sure it's still running.
4520          DoEvents
4530          blnGoingToReport = True
4540          Select Case garr_varGoToReport(GTR_CTL6)
              Case "cmdPrint00"
4550            lngBtn_Top = (.cmdPreview00_lbl.Top + (lngTpp))
4560            lngBtn_Left = (.cmdPreview00_lbl.Left + .cmdPreview00_lbl.Width)
4570            .cmdPreview00_lbl4_txt.Visible = False
4580            .cmdPreview00_lbl5.Visible = False
4590            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4600            DoEvents
4610          Case "cmdPrint00B"
4620            lngBtn_Top = (.cmdPreview00B_lbl.Top - (0& * lngTpp))
4630            lngBtn_Left = ((.cmdPreview00B_lbl.Left + .cmdPreview00B_lbl.Width) - (lngTpp))
4640            .GoToReport_arw_alist_on_img.Top = lngBtn_Top
4650            .GoToReport_arw_alist_off_img.Top = lngBtn_Top
4660            .GoToReport_arw_alist_on_img.Left = lngBtn_Left
4670            .GoToReport_arw_alist_off_img.Left = lngBtn_Left
4680            Select Case .chkAssetList
                Case True
4690              .GoToReport_arw_alist_on_img.Visible = True
4700            Case False
4710              .GoToReport_arw_alist_off_img.Visible = True
4720            End Select
4730            .chkAssetList.Visible = False
4740            .chkAssetList_lbl2.Visible = False
4750            .chkAssetList_box.Visible = False
4760            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4770            DoEvents
4780          Case "cmdPrint01"
4790            lngBtn_Top = (.cmdPreview01_lbl.Top + (lngTpp))
4800            lngBtn_Left = (.cmdPreview01_lbl.Left + .cmdPreview01_lbl.Width)
4810            .cmdPreview01_lbl4_txt.Visible = False
4820            .cmdPreview01_lbl5.Visible = False
4830            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4840            DoEvents
4850          Case "cmdPrint02"
4860            Select Case garr_varGoToReport(GTR_RNAM)
                Case "rptCourtRptFL_04"
                  ' ** Representative.
4870              If .opgType <> .opgType_optRep.OptionValue Then
4880                .opgType = .opgType_optRep.OptionValue
4890                opgType_AfterUpdate  ' ** Procedure: Below.
4900              End If
4910              .GoToReport_arw_rep_img.Visible = True
4920            Case "rptCourtRptFL_04B"
                  ' ** Guardian.
4930              If .opgType <> .opgType_optGuard.OptionValue Then
4940                .opgType = .opgType_optGuard.OptionValue
4950                opgType_AfterUpdate  ' ** Procedure: Below.
4960              End If
4970              .GoToReport_arw_grd_img.Visible = True
4980            End Select
4990            DoEvents
5000            .TimerInterval = GTR_WAIT
5010          Case "cmdPrint03"
5020            Select Case garr_varGoToReport(GTR_RNAM)
                Case "rptCourtRptFL_06"
                  ' ** Representative.
5030              If .opgType <> .opgType_optRep.OptionValue Then
5040                .opgType = .opgType_optRep.OptionValue
5050                opgType_AfterUpdate  ' ** Procedure: Below.
5060              End If
5070              .GoToReport_arw_rep_img.Visible = True
5080            Case Else
                  ' ** Nothing else.
5090            End Select
5100            DoEvents
5110            .TimerInterval = GTR_WAIT
5120          Case "cmdPrint04"
5130            lngBtn_Top = (.cmdPreview04_lbl.Top + (lngTpp))
5140            lngBtn_Left = (.cmdPreview04_lbl.Left + .cmdPreview04_lbl.Width)
5150            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
5160            DoEvents
5170          Case "cmdPrint05"
5180            lngBtn_Top = (.cmdPreview05_lbl.Top - (0& * lngTpp))
5190            lngBtn_Left = ((.cmdPreview05_lbl.Left + .cmdPreview05_lbl.Width) - (0& * lngTpp))
5200            .GoToReport_arw_alist_e_img.Top = lngBtn_Top
5210            .GoToReport_arw_alist_e_img.Left = lngBtn_Left
5220            .GoToReport_arw_alist_e_img.Visible = True
5230            .cmdPreview05_Sch_lbl.Visible = False
5240            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
5250            DoEvents
5260          End Select
5270          If garr_varGoToReport(GTR_CTL6) <> "cmdPrint02" And garr_varGoToReport(GTR_CTL6) <> "cmdPrint03" Then
5280            If garr_varGoToReport(GTR_CTL6) <> "cmdPrint00B" And garr_varGoToReport(GTR_CTL6) <> "cmdPrint05" Then
5290              .GoToReport_arw_rpt_img.Top = lngBtn_Top
5300              .GoToReport_arw_rpt_img.Left = lngBtn_Left
5310              .GoToReport_arw_rpt_img.Visible = True
5320            End If
5330            DoEvents
5340            Beep
5350            DoCmd.Hourglass False
5360            DoEvents
5370            .TimerInterval = (GTR_WAIT * 2&)
5380          End If
5390        End Select
5400      End If
5410      If gblnTimer = True Then
5420        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
5430          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
5440          gblnTimer = False
5450          DoCmd.Hourglass False
5460        End If
5470      End If
5480    End With

'HERE! 2  Form_Load()  frmRpt_CourtReports_FL
'HERE! 1  Form_Open()  frmRpt_CourtReports_FL
'HERE! 3  Form_Timer()  frmRpt_CourtReports_FL

EXITP:
5490    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

5510    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
5520      Response = acDataErrContinue
5530    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
5540      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
5550        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
5560        Response = acDataErrContinue
5570      End If
5580    Case 2237  ' ** The text you entered isn't an item in the list.
5590      Response = acDataErrDisplay
5600    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
5610      Beep
5620      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
5630      Response = acDataErrContinue
5640    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
5650    Case Else
5660      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
5670      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
5680    End Select

EXITP:
5690    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

5710    With Me
5720      If .chkRememberMe = False Then
5730        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
5740      End If
5750      If .chkRememberDates = False Then
5760        .DateStart = Null
5770        .DateEnd = Null
5780      End If
5790    End With

5800    FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5810    DoEvents

5820    Select Case blnWasGTR
        Case True
5830      blnWasGTR = False
5840    Case False
          ' ** Don't save settings imposed by GoToReport!
5850      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
5860    End Select

5870    gblnLegalName = False

5880      ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
5890      ChkSpecLedgerEntryNull  ' ** Module Function: modUtilities.

5900    gdblCrtRpt_IncTot = 0#
5910    gdblCrtRpt_PrinTot = 0#
5920    gdblCrtRpt_CostTot = 0#

        ' ** Make sure all reports are closed.
5930    Do While Reports.Count > 0
5940      DoCmd.Close acReport, Reports(0).Name
5950    Loop

5960    strDocName = "frmRpt_CourtReports_FL_Input"
5970    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
5980      DoCmd.Close acForm, strDocName
5990    End If

6000    gstrFormQuerySpec = vbNullString
6010    gblnSetFocus = True
6020    gstrReturningForm = THIS_NAME
6030    Select Case strCallingForm
        Case "frmMenu_CourtReport"
6040      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Report"
6050    End Select

EXITP:
6060    Set qdf = Nothing
6070    Set dbs = Nothing
6080    Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

6110    With Me
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
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

6610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6620    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdClose_Click()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

6710    DoCmd.Close acForm, THIS_NAME

EXITP:
6720    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6820    intShiftDown = (Shift And acShiftMask) > 0
6830    intAltDown = (Shift And acAltMask) > 0
6840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6860      Select Case intRetVal
          Case vbKeyTab
6870        With Me
6880          intRetVal = 0
6890          .DateStart.SetFocus
6900        End With
6910      Case vbKeyLeft
6920        With Me
6930          intRetVal = 0
6940          If .cmdExcelAll.Enabled = True Then
6950            .cmdExcelAll.SetFocus
6960          Else
6970            .cmdWordAll.SetFocus
6980          End If
6990        End With
7000      Case vbKeyRight
7010        With Me
7020          intRetVal = 0
7030          .DateStart.SetFocus
7040        End With
7050      End Select
7060    End If

        ' ** Shift keys.
7070    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7080      Select Case intRetVal
          Case vbKeyTab
7090        With Me
7100          intRetVal = 0
7110          If .cmdExcelAll.Enabled = True Then
7120            .cmdExcelAll.SetFocus
7130          Else
7140            .cmdWordAll.SetFocus
7150          End If
7160        End With
7170      End Select
7180    End If

        ' ** Ctrl-Shift keys.
7190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7200      Select Case intRetVal
          Case vbKeyTab
7210        With Me
7220          intRetVal = 0
7230          If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
7240            .UserReportPath_chk.SetFocus
7250          End If
7260        End With
7270      End Select
7280    End If

EXITP:
7290    KeyCode = intRetVal
7300    Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7420    intShiftDown = (Shift And acShiftMask) > 0
7430    intAltDown = (Shift And acAltMask) > 0
7440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7470        With Me
7480          intRetVal = 0
7490          .DateEnd.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7550        With Me
7560          intRetVal = 0
7570          .cmdClose.SetFocus
7580        End With
7590      End Select
7600    End If

        ' ** Ctrl keys.
7610    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7630        With Me
7640          intRetVal = 0
7650          .cmdCalendar1.SetFocus
7660        End With
7670      End Select
7680    End If

EXITP:
7690    KeyCode = intRetVal
7700    Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Public Sub DateStart_Exit(Cancel As Integer)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

        Dim strTmp01 As String

7810    With Me
7820      blnRebuildTable = True
7830      If IsNull(.DateStart) = False Then
7840        gdatStartDate = .DateStart
7850        strTmp01 = CStr(CDbl(gdatStartDate))
7860        If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7870        glngStartDateLong = CLng(strTmp01)
7880      End If
7890    End With

EXITP:
7900    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case 94  ' ** Invalid use of Null.
          ' ** Ignore.
1420    Case Else
1430      THAT_PROC = THIS_PROC
1440      That_Erl = Erl
1450      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1460    End Select
1470    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

8010    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
8020    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

8110    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
8120    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

8210    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
8220    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

8310    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
8320    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

8410    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
8420    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

8510    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
8520    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8620    intShiftDown = (Shift And acShiftMask) > 0
8630    intAltDown = (Shift And acAltMask) > 0
8640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8670        With Me
8680          intRetVal = 0
8690          .cmbAccounts.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8750        With Me
8760          intRetVal = 0
8770          .DateStart.SetFocus
8780        End With
8790      End Select
8800    End If

        ' ** Ctrl keys.
8810    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8820      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8830        With Me
8840          intRetVal = 0
8850          .cmdCalendar2.SetFocus
8860        End With
8870      End Select
8880    End If

        ' ** Ctrl-Shift keys.
8890    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8900      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8910        With Me
8920          intRetVal = 0
8930          .cmdCalendar1.SetFocus
8940        End With
8950      End Select
8960    End If

EXITP:
8970    KeyCode = intRetVal
8980    Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Public Sub DateEnd_Exit(Cancel As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

        Dim strTmp01 As String

9010    With Me
9020      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
9030        .DateEnd = Date
9040      Else
9050        If .DateEnd = vbNullString Then
9060          .DateEnd = Date
9070        End If
9080      End If
9090      blnRebuildTable = True
9100      gdatEndDate = .DateEnd
9110      strTmp01 = CStr(CDbl(gdatEndDate))
9120      If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
9130      glngEndDateLong = CLng(strTmp01)
9140    End With

EXITP:
9150    Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

9210    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
9220    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

9310    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
9320    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

9410    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
9420    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9520    intShiftDown = (Shift And acShiftMask) > 0
9530    intAltDown = (Shift And acAltMask) > 0
9540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9570        With Me
9580          intRetVal = 0
9590          .chkRememberDates.SetFocus
9600        End With
9610      End Select
9620    End If

        ' ** Shift keys.
9630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9650        With Me
9660          intRetVal = 0
9670          .DateEnd.SetFocus
9680        End With
9690      End Select
9700    End If

EXITP:
9710    KeyCode = intRetVal
9720    Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

9810    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
9820    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

9910    Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
9920    Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

10010   Calendar_Handler_FL THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsFL.

EXITP:
10020   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

10110   With Me
10120     Select Case .chkRememberDates
          Case True
10130       .chkRememberDates_lbl.FontBold = True
10140     Case False
10150       .chkRememberDates_lbl.FontBold = False
10160     End Select
10170   End With

EXITP:
10180   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10220   intShiftDown = (Shift And acShiftMask) > 0
10230   intAltDown = (Shift And acAltMask) > 0
10240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10270       With Me
10280         intRetVal = 0
10290         .opgType.SetFocus
10300       End With
10310     End Select
10320   End If

        ' ** Shift keys.
10330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10350       With Me
10360         intRetVal = 0
10370         .cmdCalendar2.SetFocus
10380       End With
10390     End Select
10400   End If

EXITP:
10410   KeyCode = intRetVal
10420   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Public Sub opgType_AfterUpdate()
' ** Public for GoToReport.

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_AfterUpdate"

        Dim lngTmp01 As Long, lngTmp02 As Long

10510   With Me
10520     If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Procedure: modWindowFunctions.
10530       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
10540     End If
10550     Select Case .opgType
          Case .opgType_optRep.OptionValue
            ' ** Normal, Design view.
10560       .opgType_optRep_lbl.FontBold = True
10570       .opgType_optGuard_lbl.FontBold = False
10580       .cmdPreview02_lbl.Left = .cmdPreview01_lbl.Left
10590       .cmdPreview02_lbl.Width = lngDisburse_Width
10600       .cmdPreview02_lbl.BackStyle = acBackStyleNormal
10610       .cmdPreview02_lbl.TextAlign = acTextAlignCenter
10620       .cmdPreview02_lbl_box.Visible = False
10630       .cmdPreview02_lbl4.Top = (.cmdPreview02_lbl.Top - lngDisburseStarOffset_Top)
10640       .cmdPreview02_lbl4.Left = ((.cmdPreview02_lbl.Left + .cmdPreview02_lbl.Width) + lngDisburseStarOffset_Left)
10650       .cmdPreview02_lbl4_txt.Left = (.cmdPreview02_lbl4.Left + (.cmdPreview01_lbl4_txt.Left - .cmdPreview01_lbl4.Left))
10660       .cmdPreview02_lbl4_cmd.Top = .cmdPreview02_lbl4.Top
10670       .cmdPreview02_lbl4_cmd.Left = .cmdPreview02_lbl4.Left
10680       .cmdPreview02_lbl5.Top = (.cmdPreview02_lbl4.Top + lngTpp)
10690       .cmdPreview02_lbl5.Left = (.cmdPreview02_lbl4.Left + lngTpp)  ' ** Shadow.
10700       .cmdPreview02_Sch_lbl.Top = .cmdPreview02_lbl.Top
10710       .cmdPreview03_lbl.Caption = "Distributions"
10720       .cmdPreview03_lbl.Top = lngDistribute_Top
10730       .cmdPreview03_lbl.Left = .cmdPreview01_lbl.Left
10740       .cmdPreview03_lbl.Width = lngDistribute_Width
10750       .cmdPreview03_lbl.BackStyle = acBackStyleNormal
10760       .cmdPreview03_lbl.TextAlign = acTextAlignCenter
10770       .cmdPreview03_Sch_lbl.Visible = True
10780       .cmdPreview02.Top = (.cmdPreview02_Sch_lbl.Top + lngButtonOffset_Top)
10790       .cmdPrint02.Top = (.cmdPreview02_Sch_lbl.Top + lngButtonOffset_Top)
10800       .cmdWord02.Top = (.cmdPreview02_Sch_lbl.Top + lngButtonOffset_Top)
10810       .cmdExcel02.Top = (.cmdPreview02_Sch_lbl.Top + lngButtonOffset_Top)
10820       .cmdPreview03.Enabled = True
10830       .cmdPreview03.Visible = True
10840       .cmdPrint03.Enabled = True
10850       .cmdPrint03.Visible = True
10860       .cmdWord03.Enabled = True
10870       .cmdWord03.Visible = True
      #If NoExcel Then
10880       .cmdExcel03.Enabled = False
      #Else
10890       .cmdExcel03.Enabled = True
      #End If
10900       .cmdExcel03.Visible = True
10910       .cmdPreview04_lbl.Top = ((.cmdPreview03_lbl.Top + .cmdPreview03_lbl.Height) + lngDoubleOffset_Top)
10920       .cmdPreview04_lbl2.Top = (.cmdPreview04_lbl.Top + (.cmdPreview05_lbl2.Top - .cmdPreview05_lbl.Top))
10930       .cmdPreview04_lbl_box.Top = .cmdPreview04_lbl.Top
10940       .cmdPreview04_Sch_lbl.Top = (.cmdPreview04_lbl.Top + lngScheduleOfset_Top)
10950       .cmdPreview04_Sch_lbl.Caption = "Schedule D"
10960       .cmdPreview04.Top = (.cmdPreview04_Sch_lbl.Top + lngButtonOffset_Top)
10970       .cmdPrint04.Top = .cmdPreview04.Top
10980       .cmdWord04.Top = .cmdPreview04.Top
10990       .cmdExcel04.Top = .cmdPreview04.Top
11000       .cmdPreview05_lbl.Top = ((.cmdPreview04_lbl_box.Top + .cmdPreview04_lbl_box.Height) + lngDoubleOffset_Top)
11010       .cmdPreview05_lbl2.Top = (.cmdPreview05_lbl.Top + (.cmdPreview04_lbl2.Top - .cmdPreview04_lbl.Top))
11020       .cmdPreview05_lbl_box.Top = .cmdPreview05_lbl.Top
11030       .cmdPreview05_Sch_lbl.Top = (.cmdPreview05_lbl.Top + lngScheduleOfset_Top)
11040       .cmdPreview05_Sch_lbl.Caption = "Schedule E"
11050       .cmdPreview05.Top = (.cmdPreview05_Sch_lbl.Top + lngButtonOffset_Top)
11060       .cmdPrint05.Top = .cmdPreview05.Top
11070       .cmdWord05.Top = .cmdPreview05.Top
11080       .cmdExcel05.Top = .cmdPreview05.Top
11090       .cmdPrintAll_box04.Height = lngBox4_Height
11100       .cmdWordAll_box04.Height = lngBox4_Height
11110       .cmdExcelAll_box04.Height = lngBox4_Height
11120     Case .opgType_optGuard.OptionValue
            ' ** 2 & 3 combined, Disbursements & Distributions.
11130       .opgType_optRep_lbl.FontBold = False
11140       .opgType_optGuard_lbl.FontBold = True
11150       .cmdPreview02_lbl.Left = .cmdPreview04_lbl.Left
11160       .cmdPreview02_lbl.Width = lngDisburse_Width + (32& * lngTpp)
11170       .cmdPreview02_lbl.BackStyle = acBackStyleTransparent
11180       .cmdPreview02_lbl.TextAlign = acTextAlignLeft
11190       .cmdPreview02_lbl_box.Top = .cmdPreview02_lbl.Top
11200       .cmdPreview02_lbl_box.Left = .cmdPreview04_lbl_box.Left
11210       .cmdPreview02_lbl_box.Width = .cmdPreview02_lbl.Width
11220       .cmdPreview02_lbl_box.Height = .cmdPreview04_lbl_box.Height
11230       .cmdPreview02_lbl_box.Visible = True
11240       .cmdPreview02_lbl4.Top = (.cmdPreview02_lbl.Top - lngDisburseStarOffset_Top)
11250       .cmdPreview02_lbl4.Left = ((.cmdPreview02_lbl.Left + .cmdPreview02_lbl.Width) + lngDisburseStarOffset_Left)
11260       .cmdPreview02_lbl4_txt.Left = (.cmdPreview02_lbl4.Left + (.cmdPreview01_lbl4_txt.Left - .cmdPreview01_lbl4.Left))
11270       .cmdPreview02_lbl4_cmd.Top = .cmdPreview02_lbl4.Top
11280       .cmdPreview02_lbl4_cmd.Left = .cmdPreview02_lbl4.Left
11290       .cmdPreview02_lbl5.Top = (.cmdPreview02_lbl4.Top + lngTpp)
11300       .cmdPreview02_lbl5.Left = (.cmdPreview02_lbl4.Left + lngTpp)  ' ** Shadow.
11310       .cmdPreview02_Sch_lbl.Top = (.cmdPreview02_lbl.Top + lngScheduleOfset_Top)
11320       .cmdPreview03_lbl.Caption = "and Distributions"
11330       .cmdPreview03_lbl.Top = (.cmdPreview02_lbl.Top + (.cmdPreview04_lbl2.Top - .cmdPreview04_lbl.Top))
11340       .cmdPreview03_lbl.Left = .cmdPreview04_lbl.Left
11350       .cmdPreview03_lbl.Width = .cmdPreview02_lbl.Width
11360       .cmdPreview03_lbl.BackStyle = acBackStyleTransparent
11370       .cmdPreview03_lbl.TextAlign = acTextAlignLeft
11380       .cmdPreview03_Sch_lbl.Visible = False
11390       .cmdPreview02.Top = (.cmdPreview02_Sch_lbl.Top + lngButtonOffset_Top)
11400       .cmdPrint02.Top = .cmdPreview02.Top
11410       .cmdWord02.Top = .cmdPreview02.Top
11420       .cmdExcel02.Top = .cmdPreview02.Top
11430       .cmdPreview03.Enabled = False
11440       .cmdPreview03.Visible = False
11450       .cmdPrint03.Enabled = False
11460       .cmdPrint03.Visible = False
11470       .cmdWord03.Enabled = False
11480       .cmdWord03.Visible = False
      #If NoExcel Then
11490       .cmdExcel03.Enabled = False
      #Else
11500       .cmdExcel03.Enabled = False
      #End If
11510       .cmdExcel03.Visible = False
11520       .cmdPreview04_lbl.Top = ((.cmdPreview02_lbl_box.Top + .cmdPreview02_lbl_box.Height) + lngDoubleOffset_Top)
11530       .cmdPreview04_lbl2.Top = (.cmdPreview04_lbl.Top + (.cmdPreview05_lbl2.Top - .cmdPreview05_lbl.Top))
11540       .cmdPreview04_lbl_box.Top = .cmdPreview04_lbl.Top
11550       .cmdPreview04_Sch_lbl.Top = (.cmdPreview04_lbl.Top + lngScheduleOfset_Top)
11560       .cmdPreview04_Sch_lbl.Caption = "Schedule C"
11570       .cmdPreview04.Top = (.cmdPreview04_Sch_lbl.Top + lngButtonOffset_Top)
11580       .cmdPrint04.Top = .cmdPreview04.Top
11590       .cmdWord04.Top = .cmdPreview04.Top
11600       .cmdExcel04.Top = .cmdPreview04.Top
11610       .cmdPreview05_lbl.Top = ((.cmdPreview04_lbl_box.Top + .cmdPreview04_lbl_box.Height) + lngDoubleOffset_Top)
11620       .cmdPreview05_lbl2.Top = (.cmdPreview05_lbl.Top + (.cmdPreview04_lbl2.Top - .cmdPreview04_lbl.Top))
11630       .cmdPreview05_lbl_box.Top = .cmdPreview05_lbl.Top
11640       .cmdPreview05_Sch_lbl.Top = (.cmdPreview05_lbl.Top + lngScheduleOfset_Top)
11650       .cmdPreview05_Sch_lbl.Caption = "Schedule D"
11660       .cmdPreview05.Top = (.cmdPreview05_Sch_lbl.Top + lngButtonOffset_Top)
11670       .cmdPrint05.Top = .cmdPreview05.Top
11680       .cmdWord05.Top = .cmdPreview05.Top
11690       .cmdExcel05.Top = .cmdPreview05.Top
11700       lngTmp01 = ((.cmdPrint05.Top + .cmdPrint05.Height) + (2& * lngTpp))
11710       lngTmp02 = (.cmdPrintAll_box04.Top + lngBox4_Height)
11720       lngTmp01 = (lngTmp02 - lngTmp01)
11730       .cmdPrintAll_box04.Height = (lngBox4_Height - lngTmp01)
11740       .cmdWordAll_box04.Height = (lngBox4_Height - lngTmp01)
11750       .cmdExcelAll_box04.Height = (lngBox4_Height - lngTmp01)
11760     End Select
11770   End With

EXITP:
11780   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub opgType_optRep_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optRep_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11820   intShiftDown = (Shift And acShiftMask) > 0
11830   intAltDown = (Shift And acAltMask) > 0
11840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11870       With Me
11880         intRetVal = 0
11890         .opgAccountSource.SetFocus
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11950       With Me
11960         intRetVal = 0
11970         .chkRememberDates.SetFocus
11980       End With
11990     End Select
12000   End If

EXITP:
12010   KeyCode = intRetVal
12020   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub opgType_optGuard_KeyDown(KeyCode As Integer, Shift As Integer)

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgType_optGuard_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12120   intShiftDown = (Shift And acShiftMask) > 0
12130   intAltDown = (Shift And acAltMask) > 0
12140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12170       With Me
12180         intRetVal = 0
12190         .opgAccountSource.SetFocus
12200       End With
12210     End Select
12220   End If

        ' ** Shift keys.
12230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12250       With Me
12260         intRetVal = 0
12270         .chkRememberDates.SetFocus
12280       End With
12290     End Select
12300   End If

EXITP:
12310   KeyCode = intRetVal
12320   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

12410   strAccountNo = vbNullString

12420   With Me
12430     If IsNull(.cmbAccounts) = False Then
12440       If Len(.cmbAccounts.Column(0)) > 0 Then
12450         strAccountNo = .cmbAccounts.Column(0)
12460       End If
12470     End If
12480     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
12490       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
12500       .opgAccountSource_optNumber_lbl.FontBold = True
12510       .opgAccountSource_optName_lbl.FontBold = False
12520     Case .opgAccountSource_optName.OptionValue
12530       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
12540       .opgAccountSource_optNumber_lbl.FontBold = False
12550       .opgAccountSource_optName_lbl.FontBold = True
12560     End Select
12570     DoEvents
12580     If strAccountNo <> vbNullString Then
12590       .cmbAccounts = strAccountNo
12600     End If
12610   End With

EXITP:
12620   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12720   intShiftDown = (Shift And acShiftMask) > 0
12730   intAltDown = (Shift And acAltMask) > 0
12740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12770       With Me
12780         intRetVal = 0
12790         .chkRememberMe.SetFocus
12800       End With
12810     End Select
12820   End If

        ' ** Shift keys.
12830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12850       With Me
12860         intRetVal = 0
12870         .opgType.SetFocus
12880       End With
12890     End Select
12900   End If

EXITP:
12910   KeyCode = intRetVal
12920   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13020   intShiftDown = (Shift And acShiftMask) > 0
13030   intAltDown = (Shift And acAltMask) > 0
13040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13070       With Me
13080         intRetVal = 0
13090         .chkRememberMe.SetFocus
13100       End With
13110     End Select
13120   End If

        ' ** Shift keys.
13130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13150       With Me
13160         intRetVal = 0
13170         .opgType.SetFocus
13180       End With
13190     End Select
13200   End If

EXITP:
13210   KeyCode = intRetVal
13220   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

13310   With Me
13320     Select Case .chkRememberMe
          Case True
13330       .chkRememberMe_lbl.FontBold = True
13340     Case False
13350       .chkRememberMe_lbl.FontBold = False
13360     End Select
13370   End With

EXITP:
13380   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13420   intShiftDown = (Shift And acShiftMask) > 0
13430   intAltDown = (Shift And acAltMask) > 0
13440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13470       With Me
13480         intRetVal = 0
13490         .cmbAccounts.SetFocus
13500       End With
13510     End Select
13520   End If

        ' ** Shift keys.
13530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13550       With Me
13560         intRetVal = 0
13570         .opgAccountSource.SetFocus
13580       End With
13590     End Select
13600   End If

EXITP:
13610   KeyCode = intRetVal
13620   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

13710   With Me
13720     If .cmbAccounts.text = vbNullString Then
13730       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
13740       .cmbAccounts = .cmbAccounts.ItemData(0)
13750       blnRebuildTable = True
13760     End If
13770   End With

EXITP:
13780   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

13810   blnRebuildTable = True
13820   gstrCrtRpt_Account = Nz(Me.cmbAccounts, vbNullString)
13830   gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
13840   gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
13850   gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString
13860   gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False

EXITP:
13870   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13920   intShiftDown = (Shift And acShiftMask) > 0
13930   intAltDown = (Shift And acAltMask) > 0
13940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13970       With Me
13980         intRetVal = 0
13990         .cmdPreview00.SetFocus
14000       End With
14010     End Select
14020   End If

        ' ** Shift keys.
14030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14050       With Me
14060         intRetVal = 0
14070         .DateEnd.SetFocus
14080       End With
14090     End Select
14100   End If

        ' ** Ctrl keys.
14110   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
14120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14130       With Me
14140         intRetVal = 0
14150         If .chkGroupBy_IncExpCode.Enabled = True Then
14160           .chkGroupBy_IncExpCode.SetFocus
                'ElseIf .chkIncludeArchive.Enabled = True Then
                '  .chkIncludeArchive.SetFocus
14170         Else
14180           .chkLegalName.SetFocus
14190         End If
14200       End With
14210     End Select
14220   End If

        ' ** Ctrl-Shift keys.
14230   If intCtrlDown And (Not intAltDown) And intShiftDown Then
14240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14250       With Me
14260         intRetVal = 0
14270         .chkRememberMe.SetFocus
14280       End With
14290     End Select
14300   End If

EXITP:
14310   KeyCode = intRetVal
14320   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Public Sub chkGroupBy_IncExpCode_AfterUpdate()
' ** Public for GoToReport.
' ** Available on reports 1 and 2.

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_AfterUpdate"

14410   With Me
14420     Select Case .chkGroupBy_IncExpCode
          Case True
14430       gblnUseReveuneExpenseCodes = True
14440       .chkGroupBy_IncExpCode_lbl.FontBold = True
14450       .chkGroupBy_IncExpCode_lbl4.FontSize = 20
14460       .chkGroupBy_IncExpCode_lbl5.Visible = True
14470       .cmdPreview00_lbl4.FontSize = 20
14480       .cmdPreview00_lbl5.Visible = True
14490       .cmdPreview01_lbl4.FontSize = 20
14500       .cmdPreview01_lbl5.Visible = True
14510       .cmdPreview02_lbl4.FontSize = 20
14520       .cmdPreview02_lbl5.Visible = True
14530     Case False
14540       gblnUseReveuneExpenseCodes = False
14550       .chkGroupBy_IncExpCode_lbl.FontBold = False
14560       .chkGroupBy_IncExpCode_lbl4.FontSize = 14
14570       .chkGroupBy_IncExpCode_lbl5.Visible = False
14580       .cmdPreview00_lbl4.FontSize = 14
14590       .cmdPreview00_lbl5.Visible = False
14600       .cmdPreview01_lbl4.FontSize = 14
14610       .cmdPreview01_lbl5.Visible = False
14620       .cmdPreview02_lbl4.FontSize = 14
14630       .cmdPreview02_lbl5.Visible = False
14640     End Select
14650   End With

EXITP:
14660   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub chkGroupBy_IncExpCode_KeyDown(KeyCode As Integer, Shift As Integer)

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14720   intShiftDown = (Shift And acShiftMask) > 0
14730   intAltDown = (Shift And acAltMask) > 0
14740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14770       With Me
14780         intRetVal = 0
14790         .chkLegalName.SetFocus
14800       End With
14810     End Select
14820   End If

        ' ** Shift keys.
14830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14850       With Me
14860         intRetVal = 0
14870         .cmbAccounts.SetFocus
14880       End With
14890     End Select
14900   End If

EXITP:
14910   KeyCode = intRetVal
14920   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub chkLegalName_AfterUpdate()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_AfterUpdate"

15010   With Me
15020     Select Case .chkLegalName
          Case True
15030       .chkLegalName_lbl.FontBold = True
15040       gblnLegalName = True
15050     Case False
15060       .chkLegalName_lbl.FontBold = False
15070       gblnLegalName = False
15080     End Select
15090   End With

EXITP:
15100   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub chkLegalName_KeyDown(KeyCode As Integer, Shift As Integer)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15220   intShiftDown = (Shift And acShiftMask) > 0
15230   intAltDown = (Shift And acAltMask) > 0
15240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15270       With Me
15280         intRetVal = 0
15290         .chkIncludeFootnote.SetFocus
15300       End With
15310     End Select
15320   End If

        ' ** Shift keys.
15330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15350       With Me
15360         intRetVal = 0
15370         If .chkGroupBy_IncExpCode.Enabled = True Then
15380           .chkGroupBy_IncExpCode.SetFocus
15390         Else
15400           .cmbAccounts.SetFocus
15410         End If
15420       End With
15430     End Select
15440   End If

EXITP:
15450   KeyCode = intRetVal
15460   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub chkIncludeFootnote_AfterUpdate()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_AfterUpdate"

15510   With Me
15520     .CourtReports_Footnote.Visible = True
15530     .CourtReports_Footnote_box.Visible = True
15540     .CourtReports_Footnote_lbl2.Visible = False
15550     Select Case .chkIncludeFootnote
          Case True
15560       .chkIncludeFootnote_lbl.FontBold = True
15570       .CourtReports_Footnote.Enabled = True
15580       .CourtReports_Footnote.BackStyle = acBackStyleNormal
15590       .CourtReports_Footnote_lbl.BackStyle = acBackStyleNormal
15600       .CourtReports_Footnote.BorderColor = CLR_LTBLU2
15610       If IsNull(.CourtReports_Footnote) = True Then
15620         .CourtReports_Footnote = strFootnoteDefault
15630       End If
15640       If gblnRevenueExpenseTracking = False Then
15650         .chkGroupBy_IncExpCode_lbl3.Visible = False
15660       End If
15670     Case False
15680       .chkIncludeFootnote_lbl.FontBold = False
15690       .CourtReports_Footnote.Enabled = False
15700       .CourtReports_Footnote.BackStyle = acBackStyleTransparent
15710       .CourtReports_Footnote_lbl.BackStyle = acBackStyleTransparent
15720       .CourtReports_Footnote.BorderColor = WIN_CLR_DISR
15730       If gblnRevenueExpenseTracking = False Then
15740         .chkGroupBy_IncExpCode_lbl3.Visible = True
15750       End If
15760     End Select
15770   End With

EXITP:
15780   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub chkIncludeFootnote_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15820   intShiftDown = (Shift And acShiftMask) > 0
15830   intAltDown = (Shift And acAltMask) > 0
15840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15870       With Me
15880         intRetVal = 0
15890         Select Case .chkIncludeFootnote
              Case True
15900           .CourtReports_Footnote.SetFocus
15910         Case False
15920           .chkPageOf.SetFocus
15930         End Select
15940       End With
15950     End Select
15960   End If

        ' ** Shift keys.
15970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15990       With Me
16000         intRetVal = 0
16010         .chkLegalName.SetFocus
16020       End With
16030     End Select
16040   End If

EXITP:
16050   KeyCode = intRetVal
16060   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub CourtReports_Footnote_AfterUpdate()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_AfterUpdate"

        Dim varTmp00 As Variant

16110   With Me
16120     varTmp00 = .CourtReports_Footnote
16130     If IsNull(varTmp00) = False Then
16140       If Trim(varTmp00) <> vbNullString Then
16150         varTmp00 = Trim(varTmp00)
16160         If Len(varTmp00) > MEMO_MAX Then
16170           Beep
16180           If .CourtReports_Footnote_lbl2.Visible = False Then
16190             .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
16200             .CourtReports_Footnote_lbl2.Visible = True
16210           End If
16220           varTmp00 = Left(varTmp00, MEMO_MAX)
16230           .CourtReports_Footnote = varTmp00
16240         End If
16250       Else
16260         .CourtReports_Footnote = Null
16270       End If
16280     End If
16290   End With

EXITP:
16300   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub CourtReports_Footnote_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intLen As Integer
        Dim intRetVal As Integer

16410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16420   intShiftDown = (Shift And acShiftMask) > 0
16430   intAltDown = (Shift And acAltMask) > 0
16440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16470       With Me
16480         intRetVal = 0
16490         .chkPageOf.SetFocus
16500       End With
16510     Case Else
16520       With Me
16530         If IsNull(.CourtReports_Footnote.text) = False Then
16540           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
16550             intLen = Len(Trim(.CourtReports_Footnote.text))
16560             If intLen > MEMO_MAX Then
16570               intRetVal = 0
16580               Beep
16590               If .CourtReports_Footnote_lbl2.Visible = False Then
16600                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
16610                 .CourtReports_Footnote_lbl2.Visible = True
16620               End If
16630               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
16640               .CourtReports_Footnote.SelStart = 999
16650             End If
16660           End If
16670         End If
16680       End With
16690     End Select
16700   End If

        ' ** Shift keys.
16710   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16720     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16730       With Me
16740         intRetVal = 0
16750         .chkIncludeFootnote.SetFocus
16760       End With
16770     Case Else
16780       With Me
16790         If IsNull(.CourtReports_Footnote.text) = False Then
16800           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
16810             intLen = Len(Trim(.CourtReports_Footnote.text))
16820             If intLen > MEMO_MAX Then
16830               intRetVal = 0
16840               Beep
16850               If .CourtReports_Footnote_lbl2.Visible = False Then
16860                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
16870                 .CourtReports_Footnote_lbl2.Visible = True
16880               End If
16890               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
16900               .CourtReports_Footnote.SelStart = 999
16910             End If
16920           End If
16930         End If
16940       End With
16950     End Select
16960   End If

EXITP:
16970   KeyCode = intRetVal
16980   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

17010   With Me
17020     Select Case .chkPageOf
          Case True
17030       .chkPageOf_lbl.FontBold = True
17040     Case False
17050       .chkPageOf_lbl.FontBold = False
17060     End Select
17070   End With

EXITP:
17080   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17120   intShiftDown = (Shift And acShiftMask) > 0
17130   intAltDown = (Shift And acAltMask) > 0
17140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17170       With Me
17180         intRetVal = 0
17190         .cmdPreview00.SetFocus
17200       End With
17210     End Select
17220   End If

        ' ** Shift keys.
17230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17250       With Me
17260         intRetVal = 0
17270         .chkIncludeFootnote.SetFocus
17280       End With
17290     End Select
17300   End If

EXITP:
17310   KeyCode = intRetVal
17320   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub cmdPreview00_lbl4_cmd_DblClick(Cancel As Integer)

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_lbl4_cmd_DblClick"

17410   With Me
17420     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
17430     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
17440   End With

EXITP:
17450   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview00_Click()

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_Click"

17510   blnIsSummary = True

17520   blnAllCancel = False
17530   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
17540   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
17550   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

17560   intSpecSQLCnt = 0: lngBuildCnt = 0&
17570   PreviewOrPrint "0", THIS_PROC, acViewPreview

17580   blnIsSummary = False

EXITP:
17590   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview00_KeyDown(KeyCode As Integer, Shift As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17620   intShiftDown = (Shift And acShiftMask) > 0
17630   intAltDown = (Shift And acAltMask) > 0
17640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17660     Select Case intRetVal
          Case vbKeyTab
17670       With Me
17680         intRetVal = 0
17690         .cmdPrint00.SetFocus
17700       End With
17710     End Select
17720   End If

        ' ** Shift keys.
17730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17740     Select Case intRetVal
          Case vbKeyTab
17750       With Me
17760         intRetVal = 0
17770         .cmbAccounts.SetFocus
17780       End With
17790     End Select
17800   End If

        ' ** Ctrl-Shift keys.
17810   If intCtrlDown And (Not intAltDown) And intShiftDown Then
17820     Select Case intRetVal
          Case vbKeyTab
17830       With Me
17840         intRetVal = 0
17850         .chkPageOf.SetFocus
17860       End With
17870     End Select
17880   End If

EXITP:
17890   KeyCode = intRetVal
17900   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Public Sub cmdPrint00_Click()
' ** Public, because it's called from the TAReports CommandBar as well.

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00_Click"

        ' ** When called from the TAReports CommandBar, these variables should already be filled.
        ' **   gstrCrtRpt_Ordinal
        ' **   gstrCrtRpt_Version
        ' **   gstrCrtRpt_CashAssets_Beg
        ' **   gstrCrtRpt_NetIncome
        ' **   gstrCrtRpt_NetLoss
        ' **   gstrCrtRpt_CashAssets_End

18010   blnIsSummary = True

18020   blnAllCancel = False
18030   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
18040   blnAccountChecked = False

18050   If gblnCrtRpt_Zero = False Then
          ' ** Set up the TAReports Command Bar (visible with Print Preview).
18060     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
18070   Else
          ' ** cmdPrint00_Click() is being invoked by TAReports_Print1() or TAReports_Print2()
          ' ** in modReportFunctions from the TAReports CommandBar via the mcrPrint_CA_0 or
          ' ** mcrPrint_CA_0_Dialog macros.
18080   End If

18090   intSpecSQLCnt = 0: lngBuildCnt = 0&
        '##GTR_Ref: rptCourtRptFL_00
        '##GTR_Ref: rptCourtRptFL_00A
18100   PreviewOrPrint "0", THIS_PROC, acViewNormal

18110   blnIsSummary = False

EXITP:
18120   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdWord00_Click()

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_Click"

18210   blnIsSummary = True

18220   blnAllCancel = False
18230   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
18240   blnAccountChecked = False

18250   intSpecSQLCnt = 0: lngBuildCnt = 0&
18260   SendToFile_FL "0", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.

18270   blnIsSummary = False

EXITP:
18280   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdWord00_KeyDown(KeyCode As Integer, Shift As Integer)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18320   intShiftDown = (Shift And acShiftMask) > 0
18330   intAltDown = (Shift And acAltMask) > 0
18340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18360     Select Case intRetVal
          Case vbKeyTab
18370       With Me
18380         intRetVal = 0
18390         If .cmdExcel00.Enabled = True Then
18400           .cmdExcel00.SetFocus
18410         Else
18420           .cmdPreview00B.SetFocus
18430         End If
18440       End With
18450     Case vbKeyRight
18460       With Me
18470         intRetVal = 0
18480         If .cmdExcel00.Enabled = True Then
18490           .cmdExcel00.SetFocus
18500         Else
18510           .cmdPreview00B.SetFocus
18520         End If
18530       End With
18540     End Select
18550   End If

        ' ** Shift keys.
18560   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18570     Select Case intRetVal
          Case vbKeyTab
18580       With Me
18590         intRetVal = 0
18600         .cmdPrint00.SetFocus
18610       End With
18620     End Select
18630   End If

EXITP:
18640   KeyCode = intRetVal
18650   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Public Sub cmdExcel00_Click()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_Click"

18710   blnIsSummary = True

18720   blnAllCancel = False
18730   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
18740   blnAccountChecked = False

18750   intSpecSQLCnt = 0: lngBuildCnt = 0&
18760   SendToFile_FL "0", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.

18770   blnIsSummary = False

EXITP:
18780   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdExcel00_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18820   intShiftDown = (Shift And acShiftMask) > 0
18830   intAltDown = (Shift And acAltMask) > 0
18840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18860     Select Case intRetVal
          Case vbKeyTab
18870       With Me
18880         intRetVal = 0
18890         .cmdPreview00B.SetFocus
18900       End With
18910     Case vbKeyRight
18920       With Me
18930         intRetVal = 0
18940         .cmdPreview00B.SetFocus
18950       End With
18960     End Select
18970   End If

        ' ** Shift keys.
18980   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18990     Select Case intRetVal
          Case vbKeyTab
19000       With Me
19010         intRetVal = 0
19020         .cmdWord00.SetFocus
19030       End With
19040     End Select
19050   End If

EXITP:
19060   KeyCode = intRetVal
19070   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Public Sub chkAssetList_AfterUpdate()
' ** Public for GoToReport.

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_AfterUpdate"

19110   With Me
19120     Select Case .chkAssetList
          Case True
19130       .chkAssetList_lbl.FontBold = True
19140       .chkAssetList_lbl2.FontBold = True
19150     Case False
19160       .chkAssetList_lbl.FontBold = False
19170       .chkAssetList_lbl2.FontBold = False
19180     End Select
19190   End With

EXITP:
19200   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Private Sub chkAssetList_KeyDown(KeyCode As Integer, Shift As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19320   intShiftDown = (Shift And acShiftMask) > 0
19330   intAltDown = (Shift And acAltMask) > 0
19340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19370       With Me
19380         intRetVal = 0
19390         .cmdPreview00B.SetFocus
19400       End With
19410     End Select
19420   End If

        ' ** Shift keys.
19430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
19450       With Me
19460         intRetVal = 0
19470         If .cmdExcel00.Enabled = True Then
19480           .cmdExcel00.SetFocus
19490         Else
19500           .cmdWord00.SetFocus
19510         End If
19520       End With
19530     End Select
19540   End If

EXITP:
19550   KeyCode = intRetVal
19560   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub cmdPreview00B_Click()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_Click"

19610   If Validate = True Then  ' ** Function: Below.
19620     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
19630     PreviewOrPrint "0B", THIS_PROC, acViewPreview
19640   End If  ' ** Validate().

EXITP:
19650   Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Private Sub cmdPreview00B_KeyDown(KeyCode As Integer, Shift As Integer)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19720   intShiftDown = (Shift And acShiftMask) > 0
19730   intAltDown = (Shift And acAltMask) > 0
19740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19760     Select Case intRetVal
          Case vbKeyTab
19770       With Me
19780         intRetVal = 0
19790         .cmdPrint00B.SetFocus
19800       End With
19810     Case vbKeyLeft
19820       With Me
19830         intRetVal = 0
19840         If .cmdExcel00.Enabled = True Then
19850           .cmdExcel00.SetFocus
19860         Else
19870           .cmdWord00.SetFocus
19880         End If
19890       End With
19900     End Select
19910   End If

        ' ** Shift keys.
19920   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19930     Select Case intRetVal
          Case vbKeyTab
19940       With Me
19950         intRetVal = 0
19960         If .cmdExcel00.Enabled = True Then
19970           .cmdExcel00.SetFocus
19980         Else
19990           .cmdWord00.SetFocus
20000         End If
20010       End With
20020     End Select
20030   End If

EXITP:
20040   KeyCode = intRetVal
20050   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub cmdPrint00B_Click()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00B_Click"

20110   If Validate = True Then  ' ** Function: Below.
20120     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptFL_00B
20130     PreviewOrPrint "0B", THIS_PROC, acViewNormal
20140   End If  ' ** Validate().

EXITP:
20150   Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Public Sub cmdWord00B_Click()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00B_Click"

20210   If Validate = True Then  ' ** Function: Below.
20220     SendToFile_FL "0B", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.
20230   End If  ' ** Validate().

EXITP:
20240   Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Public Sub cmdExcel00B_Click()

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00B_Click"

20310   If Validate = True Then  ' ** Function: Below.
20320     SendToFile_FL "0B", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.
20330   End If  ' ** Validate().

EXITP:
20340   Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Private Sub cmdPreview01_lbl4_cmd_DblClick(Cancel As Integer)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview01_lbl4_cmd_DblClick"

20410   With Me
20420     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
20430     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
20440   End With

EXITP:
20450   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview01_Click()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview01_Click"

20510   blnAllCancel = False
20520   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
20530   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
20540   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

20550   If gblnUseReveuneExpenseCodes = True Then
20560     PreviewOrPrint "2A", THIS_PROC, acViewPreview
20570   Else
20580     PreviewOrPrint "2", THIS_PROC, acViewPreview
20590   End If

EXITP:
20600   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrint01_Click()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint01_Click"

20710   blnAllCancel = False
20720   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
20730   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
20740   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

20750   If gblnUseReveuneExpenseCodes = True Then
          '##GTR_Ref: rptCourtRptFL_02A
20760     PreviewOrPrint "2A", THIS_PROC, acViewNormal
20770   Else
          '##GTR_Ref: rptCourtRptFL_02
20780     PreviewOrPrint "2", THIS_PROC, acViewNormal
20790   End If

EXITP:
20800   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdWord01_Click()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord01_Click"

20910   blnAllCancel = False
20920   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
20930   blnAccountChecked = False

20940   If gblnUseReveuneExpenseCodes = True Then
20950     SendToFile_FL "2A", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.
20960   Else
20970     SendToFile_FL "2", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.
20980   End If

EXITP:
20990   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdExcel01_Click()

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel01_Click"

21010   blnAllCancel = False
21020   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
21030   blnAccountChecked = False

21040   If gblnUseReveuneExpenseCodes = True Then
21050     SendToFile_FL "2A", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.
21060   Else
21070     SendToFile_FL "2", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.
21080   End If

EXITP:
21090   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview02_lbl4_cmd_DblClick(Cancel As Integer)

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview02_lbl4_cmd_DblClick"

21110   With Me
21120     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
21130     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
21140   End With

EXITP:
21150   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview02_Click()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview02_Click"

21210   blnAllCancel = False
21220   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
21230   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
21240   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

21250   If gblnUseReveuneExpenseCodes = True Then
21260     PreviewOrPrint "4A", THIS_PROC, acViewPreview
21270   Else
21280     PreviewOrPrint "4", THIS_PROC, acViewPreview
21290   End If

EXITP:
21300   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrint02_Click()

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint02_Click"

21410   blnAllCancel = False
21420   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
21430   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
21440   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

21450   If gblnUseReveuneExpenseCodes = True Then
          '##GTR_Ref: rptCourtRptFL_04A
          '##GTR_Ref: rptCourtRptFL_04C
21460     PreviewOrPrint "4A", THIS_PROC, acViewNormal
21470   Else
          '##GTR_Ref: rptCourtRptFL_04
          '##GTR_Ref: rptCourtRptFL_04B
21480     PreviewOrPrint "4", THIS_PROC, acViewNormal
21490   End If

EXITP:
21500   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdWord02_Click()

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord02_Click"

21610   blnAllCancel = False
21620   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
21630   blnAccountChecked = False

21640   If gblnUseReveuneExpenseCodes = True Then
21650     SendToFile_FL "4A", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.
21660   Else
21670     SendToFile_FL "4", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.
21680   End If

EXITP:
21690   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdExcel02_Click()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel02_Click"

21710   blnAllCancel = False
21720   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
21730   blnAccountChecked = False

21740   If gblnUseReveuneExpenseCodes = True Then
21750     SendToFile_FL "4A", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.
21760   Else
21770     SendToFile_FL "4", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.
21780   End If

EXITP:
21790   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview03_Click()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview03_Click"

21810   blnAllCancel = False
21820   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
21830   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
21840   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

21850   PreviewOrPrint "6", THIS_PROC, acViewPreview

EXITP:
21860   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrint03_Click()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint03_Click"

21910   blnAllCancel = False
21920   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
21930   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
21940   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

        '##GTR_Ref: rptCourtRptFL_06
21950   PreviewOrPrint "6", THIS_PROC, acViewNormal

EXITP:
21960   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdWord03_Click()

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord03_Click"

22010   blnAllCancel = False
22020   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22030   blnAccountChecked = False

22040   SendToFile_FL "6", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.

EXITP:
22050   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdExcel03_Click()

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel03_Click"

22110   blnAllCancel = False
22120   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22130   blnAccountChecked = False

22140   SendToFile_FL "6", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.

EXITP:
22150   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview04_Click()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_Click"

22210   blnAllCancel = False
22220   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22230   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
22240   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

22250   PreviewOrPrint "3", THIS_PROC, acViewPreview

EXITP:
22260   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrint04_Click()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint04_Click"

22310   blnAllCancel = False
22320   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22330   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
22340   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

        '##GTR_Ref: rptCourtRptFL_03
22350   PreviewOrPrint "3", THIS_PROC, acViewNormal

EXITP:
22360   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdWord04_Click()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord04_Click"

22410   blnAllCancel = False
22420   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22430   blnAccountChecked = False

22440   SendToFile_FL "3", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.

EXITP:
22450   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdExcel04_Click()

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel04_Click"

22510   blnAllCancel = False
22520   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22530   blnAccountChecked = False

22540   SendToFile_FL "3", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.

EXITP:
22550   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPreview05_Click()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview05_Click"

22610   blnAllCancel = False
22620   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22630   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
22640   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

22650   PreviewOrPrint "7", THIS_PROC, acViewPreview

EXITP:
22660   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrint05_Click()

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint05_Click"

22710   blnAllCancel = False
22720   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22730   blnAccountChecked = False

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
22740   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

        '##GTR_Ref: rptCourtRptFL_07
22750   PreviewOrPrint "7", THIS_PROC, acViewNormal

EXITP:
22760   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdWord05_Click()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord05_Click"

22810   blnAllCancel = False
22820   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22830   blnAccountChecked = False

22840   SendToFile_FL "7", THIS_PROC, blnRebuildTable, blnIsSummary, Me, False  ' ** Module Procedure: modCourtReportsFL.

EXITP:
22850   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub cmdExcel05_Click()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel05_Click"

22910   blnAllCancel = False
22920   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
22930   blnAccountChecked = False

22940   SendToFile_FL "7", THIS_PROC, blnRebuildTable, blnIsSummary, Me, True  ' ** Module Procedure: modCourtReportsFL.

EXITP:
22950   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdExcel05_KeyDown(KeyCode As Integer, Shift As Integer)

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel05_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23020   intShiftDown = (Shift And acShiftMask) > 0
23030   intAltDown = (Shift And acAltMask) > 0
23040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23060     Select Case intRetVal
          Case vbKeyTab
23070       With Me
23080         intRetVal = 0
23090         .cmdPrintAll.SetFocus
23100       End With
23110     End Select
23120   End If

        ' ** Shift keys.
23130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23140     Select Case intRetVal
          Case vbKeyTab
23150       With Me
23160         intRetVal = 0
23170         .cmdWord05.SetFocus
23180       End With
23190     End Select
23200   End If

EXITP:
23210   KeyCode = intRetVal
23220   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

23310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23320   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

23410   With Me
23420     blnPrintAll_Focus = True
23430     .cmdPrintAll_box01.Visible = True
23440     .cmdPrintAll_box02.Visible = True
23450     Select Case .chkAssetList
          Case True
23460       .cmdPrintAll_box03.Visible = True
23470     Case False
23480       .cmdPrintAll_box03.Visible = False
23490     End Select
23500     .cmdPrintAll_box04.Visible = True
23510   End With

EXITP:
23520   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

        Dim strOrd As String, strVer As String, strCashBeg As String, dblCashBeg As Double
        Dim intRetVal As Integer

23610   blnAllCancel = False
23620   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
23630   blnAccountChecked = False
23640   gblnPrintAll = True

        ' ** Set up the TAReports Command Bar (visible with Print Preview).
23650   TAReports_SetZero False  ' ** Module Function: modReportFunctions.

23660   ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

        ' ** Get the Summary inputs first.
23670   intRetVal = FLGetCourtReportData(THIS_PROC & "^" & "cmdPrintAll")  ' ** Module Function: modCourtReportFL.
23680   If intRetVal <> 0 Then
23690     blnAllCancel = True
23700     AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
23710   Else

          ' ** Save these for later.
23720     strOrd = gstrCrtRpt_Ordinal
23730     strVer = gstrCrtRpt_Version
23740     strCashBeg = gstrCrtRpt_CashAssets_Beg
23750     dblCashBeg = Nz(Forms("frmRpt_CourtReports_FL").CashAssets_Beg, 0)

23760   End If

        'ASSET LIST GOES OUT 1ST,
        'THEN SUMMARY 2ND,
        'THEN E,
        'D,
        'C,
        'B,
        'A.

23770   If blnAllCancel = False Then  ' ** cmdPrint05_Click().
          ' ** Schedule E: Assets on Hand at Close of Accounting Period.
23780     PreviewOrPrint "7", THIS_PROC & " - R7", acViewNormal
23790   End If

23800   If blnAllCancel = False Then  ' ** cmdPrint04_Click().
          ' ** Schedule D: Capital Transactions and Adjustments.
23810     PreviewOrPrint "3", THIS_PROC & " - R3", acViewNormal
23820   End If

23830   If blnAllCancel = False And Me.opgType <> Me.opgType_optGuard.OptionValue Then  ' ** cmdPrint03_Click().
          ' ** Schedule C: Distributions.; Representative.
23840     PreviewOrPrint "6", THIS_PROC & " - R6", acViewNormal
23850   End If

23860   If blnAllCancel = False Then  ' ** cmdPrint02_Click().
23870     Select Case gblnUseReveuneExpenseCodes
          Case True
            ' ** Schedule B: Disbursements, grouped by Income/Expense codes; Representative.
23880       PreviewOrPrint "4A", THIS_PROC & " - R4A", acViewNormal
            ' ** "4C" is within PreviewOrPrint(); Schedule B: Disbursements and Distributions, grouped by Income/Expense codes; Guardian.
23890     Case False
            ' ** Schedule B: Disbursements; Representative.
23900       PreviewOrPrint "4", THIS_PROC & " - R4", acViewNormal
            ' ** "4B" is within PreviewOrPrint(); Schedule B: Disbursements and Distributions; Guardian.
23910     End Select
23920   End If

23930   If blnAllCancel = False Then  ' ** cmdPrint01_Click().
          ' ** Schedule A: Receipts.
23940     PreviewOrPrint "2", THIS_PROC & " - R2", acViewNormal
23950   End If

23960   If blnAllCancel = False Then

23970     gstrCrtRpt_Ordinal = strOrd
23980     gstrCrtRpt_Version = strVer
23990     gstrCrtRpt_CashAssets_Beg = strCashBeg
24000     Forms("frmRpt_CourtReports_FL").CashAssets_Beg = dblCashBeg

          ' ** This will print both the Summary and AssetList (if checked).
24010     PreviewOrPrint "0", THIS_PROC & " - R0", acViewNormal  ' ** cmdPrint00_Click().

24020   End If

24030   gblnPrintAll = False

EXITP:
24040   Exit Sub

ERRH:
1410    gblnPrintAll = False
1420    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
1430      blnAllCancel = True
1440      AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
1450    Case Else
1460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1470    End Select
1480    Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24120   intShiftDown = (Shift And acShiftMask) > 0
24130   intAltDown = (Shift And acAltMask) > 0
24140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24160     Select Case intRetVal
          Case vbKeyTab
24170       With Me
24180         intRetVal = 0
24190         .cmdWordAll.SetFocus
24200       End With
24210     Case vbKeyUp
24220       With Me
24230         intRetVal = 0
24240         If strLastArrow = vbNullString Then
24250           .cmdPrint05.SetFocus
24260         Else
24270           If Right(strLastArrow, 2) = "05" Or Right(strLastArrow, 2) = "00" Then
24280             .Controls(Left(strLastArrow, (InStr(strLastArrow, "0") - 1)) & "05").SetFocus
24290           Else
24300             .cmdPrint05.SetFocus
24310           End If
24320           strLastArrow = vbNullString
24330         End If
24340       End With
24350     Case vbKeyDown
24360       With Me
24370         intRetVal = 0
24380         If strLastArrow = vbNullString Then
24390           .cmdPrint00.SetFocus
24400         Else
24410           If Right(strLastArrow, 2) = "05" Or Right(strLastArrow, 2) = "00" Then
24420             .Controls(Left(strLastArrow, (InStr(strLastArrow, "0") - 1)) & "00").SetFocus
24430           Else
24440             .cmdPrint00.SetFocus
24450           End If
24460           strLastArrow = vbNullString
24470         End If
24480       End With
24490     Case vbKeyLeft
24500       With Me
24510         intRetVal = 0
24520         If .cmdExcel05.Enabled = True Then
24530           .cmdExcel05.SetFocus
24540         Else
24550           .cmdWord05.SetFocus
24560         End If
24570       End With
24580     Case vbKeyRight
24590       With Me
24600         intRetVal = 0
24610         .cmdWordAll.SetFocus
24620       End With
24630     End Select
24640   End If

        ' ** Shift keys.
24650   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24660     Select Case intRetVal
          Case vbKeyTab
24670       With Me
24680         intRetVal = 0
24690         If .cmdExcel05.Enabled = True Then
24700           .cmdExcel05.SetFocus
24710         Else
24720           .cmdWord05.SetFocus
24730         End If
24740       End With
24750     End Select
24760   End If

EXITP:
24770   KeyCode = intRetVal
24780   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1450    Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

24810   With Me
24820     If gblnPrintAll = False Then
24830       .cmdPrintAll_box01.Visible = True
24840       .cmdPrintAll_box02.Visible = True
24850       Select Case .chkAssetList
            Case True
24860         .cmdPrintAll_box03.Visible = True
24870       Case False
24880         .cmdPrintAll_box03.Visible = False
24890       End Select
24900       .cmdPrintAll_box04.Visible = True
24910       If blnWordAll_Focus = False Then
24920         .cmdWordAll_box01.Visible = False
24930         .cmdWordAll_box02.Visible = False
24940         .cmdWordAll_box03.Visible = False
24950         .cmdWordAll_box04.Visible = False
24960       End If
24970       If blnExcelAll_Focus = False Then
24980         .cmdExcelAll_box01.Visible = False
24990         .cmdExcelAll_box02.Visible = False
25000         .cmdExcelAll_box03.Visible = False
25010         .cmdExcelAll_box04.Visible = False
25020       End If
25030     End If
25040   End With

EXITP:
25050   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

25110   With Me
25120     .cmdPrintAll_box01.Visible = False
25130     .cmdPrintAll_box02.Visible = False
25140     .cmdPrintAll_box03.Visible = False
25150     .cmdPrintAll_box04.Visible = False
25160     blnPrintAll_Focus = False
25170   End With

EXITP:
25180   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

25210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
25220   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

25310   With Me
25320     blnWordAll_Focus = True
25330     .cmdWordAll_box01.Visible = True
25340     .cmdWordAll_box02.Visible = True
25350     Select Case .chkAssetList
          Case True
25360       .cmdWordAll_box03.Visible = True
25370     Case False
25380       .cmdWordAll_box03.Visible = False
25390     End Select
25400     .cmdWordAll_box04.Visible = True
25410   End With

EXITP:
25420   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

25510   With Me

25520     lngFiles = 0&
25530     ReDim arr_varFile(F_ELEMS, 0)

25540     WordAll_FL blnAccountChecked, Me  ' ** Module Procedure: modCourtReportsFL.

25550   End With

EXITP:
25560   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25620   intShiftDown = (Shift And acShiftMask) > 0
25630   intAltDown = (Shift And acAltMask) > 0
25640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25660     Select Case intRetVal
          Case vbKeyTab
25670       With Me
25680         intRetVal = 0
25690         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
25700           .cmdExcelAll.SetFocus
25710         Else
25720           .cmdClose.SetFocus
25730         End If
25740       End With
25750     Case vbKeyUp
25760       With Me
25770         intRetVal = 0
25780         .cmdWord05.SetFocus
25790         strLastArrow = vbNullString
25800       End With
25810     Case vbKeyDown
25820       With Me
25830         intRetVal = 0
25840         .cmdWord00.SetFocus
25850         strLastArrow = vbNullString
25860       End With
25870     Case vbKeyRight
25880       With Me
25890         intRetVal = 0
25900         If .cmdExcelAll.Enabled = True Then
25910           .cmdExcelAll.SetFocus
25920         Else
25930           .cmdClose.SetFocus
25940         End If
25950       End With
25960     End Select
25970   End If

        ' ** Shift keys.
25980   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25990     Select Case intRetVal
          Case vbKeyTab
26000       With Me
26010         intRetVal = 0
26020         .cmdPrintAll.SetFocus
26030       End With
26040     End Select
26050   End If

        ' ** Ctrl keys.
26060   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
26070     Select Case intRetVal
          Case vbKeyTab
26080       With Me
26090         intRetVal = 0
26100         .chkOpenWord.SetFocus
26110       End With
26120     End Select
26130   End If

EXITP:
26140   KeyCode = intRetVal
26150   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

26210   With Me
26220     If gblnPrintAll = False Then
26230       .cmdWordAll_box01.Visible = True
26240       .cmdWordAll_box02.Visible = True
26250       Select Case .chkAssetList
            Case True
26260         .cmdWordAll_box03.Visible = True
26270       Case False
26280         .cmdWordAll_box03.Visible = False
26290       End Select
26300       .cmdWordAll_box04.Visible = True
26310       If blnPrintAll_Focus = False Then
26320         .cmdPrintAll_box01.Visible = False
26330         .cmdPrintAll_box02.Visible = False
26340         .cmdPrintAll_box03.Visible = False
26350         .cmdPrintAll_box04.Visible = False
26360       End If
26370       If blnExcelAll_Focus = False Then
26380         .cmdExcelAll_box01.Visible = False
26390         .cmdExcelAll_box02.Visible = False
26400         .cmdExcelAll_box03.Visible = False
26410         .cmdExcelAll_box04.Visible = False
26420       End If
26430     End If
26440   End With

EXITP:
26450   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

26510   With Me
26520     .cmdWordAll_box01.Visible = False
26530     .cmdWordAll_box02.Visible = False
26540     .cmdWordAll_box03.Visible = False
26550     .cmdWordAll_box04.Visible = False
26560     blnWordAll_Focus = False
26570   End With

EXITP:
26580   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub chkOpenWord_AfterUpdate()

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

26610   With Me
26620     Select Case .chkOpenWord
          Case True
26630       .chkOpenWord_lbl.FontBold = True
26640       .chkOpenWord_lbl2.FontBold = True
26650       .chkOpenWord_lbl_dim_hi.FontBold = True
26660       .chkOpenWord_lbl2_dim_hi.FontBold = True
26670     Case False
26680       .chkOpenWord_lbl.FontBold = False
26690       .chkOpenWord_lbl2.FontBold = False
26700       .chkOpenWord_lbl_dim_hi.FontBold = False
26710       .chkOpenWord_lbl2_dim_hi.FontBold = False
26720     End Select
26730   End With

EXITP:
26740   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26820   intShiftDown = (Shift And acShiftMask) > 0
26830   intAltDown = (Shift And acAltMask) > 0
26840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26860     Select Case intRetVal
          Case vbKeyTab
26870       With Me
26880         intRetVal = 0
26890         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
26900           .cmdExcelAll.SetFocus
26910         Else
26920           .cmdClose.SetFocus
26930         End If
26940       End With
26950     Case vbKeyUp
26960       With Me
26970         intRetVal = 0
26980         .cmdWordAll.SetFocus
26990       End With
27000     Case vbKeyDown
27010       With Me
27020         intRetVal = 0
27030         .cmdWord00.SetFocus
27040       End With
27050     End Select
27060   End If

        ' ** Shift keys.
27070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27080     Select Case intRetVal
          Case vbKeyTab
27090       With Me
27100         intRetVal = 0
27110         .cmdWordAll.SetFocus
27120       End With
27130     End Select
27140   End If

        ' ** Ctrl keys.
27150   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
27160     Select Case intRetVal
          Case vbKeyTab
27170       With Me
27180         intRetVal = 0
27190         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
27200           .chkOpenExcel.SetFocus
27210         ElseIf .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
27220           .UserReportPath_chk.SetFocus
27230         Else
27240           Beep
27250         End If
27260       End With
27270     End Select
27280   End If

EXITP:
27290   KeyCode = intRetVal
27300   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

27410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27420   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

27510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27520   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

27610   With Me
27620     blnExcelAll_Focus = True
27630     .cmdExcelAll_box01.Visible = True
27640     .cmdExcelAll_box02.Visible = True
27650     Select Case .chkAssetList
          Case True
27660       .cmdExcelAll_box03.Visible = True
27670     Case False
27680       .cmdExcelAll_box03.Visible = False
27690     End Select
27700     .cmdExcelAll_box04.Visible = True
27710   End With

EXITP:
27720   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

27810   With Me

27820     lngFiles = 0&
27830     ReDim arr_varFile(F_ELEMS, 0)

27840     ExcelAll_FL blnAccountChecked, Me  ' ** Module Procedure: modCourtReportsFL.

27850   End With

EXITP:
27860   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27920   intShiftDown = (Shift And acShiftMask) > 0
27930   intAltDown = (Shift And acAltMask) > 0
27940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27960     Select Case intRetVal
          Case vbKeyTab
27970       With Me
27980         intRetVal = 0
27990         .cmdClose.SetFocus
28000       End With
28010     Case vbKeyUp
28020       With Me
28030         intRetVal = 0
28040         .cmdExcel05.SetFocus
28050         strLastArrow = vbNullString
28060       End With
28070     Case vbKeyDown
28080       With Me
28090         intRetVal = 0
28100         .cmdExcel00.SetFocus
28110         strLastArrow = vbNullString
28120       End With
28130     Case vbKeyRight
28140       With Me
28150         intRetVal = 0
28160         .cmdClose.SetFocus
28170       End With
28180     End Select
28190   End If

        ' ** Shift keys.
28200   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28210     Select Case intRetVal
          Case vbKeyTab
28220       With Me
28230         intRetVal = 0
28240         .cmdWordAll.SetFocus
28250       End With
28260     End Select
28270   End If

        ' ** Ctrl keys.
28280   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
28290     Select Case intRetVal
          Case vbKeyTab
28300       With Me
28310         intRetVal = 0
28320         .chkOpenExcel.SetFocus
28330       End With
28340     End Select
28350   End If

EXITP:
28360   KeyCode = intRetVal
28370   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

28410   With Me
28420     If gblnPrintAll = False Then
28430       .cmdExcelAll_box01.Visible = True
28440       .cmdExcelAll_box02.Visible = True
28450       Select Case .chkAssetList
            Case True
28460         .cmdExcelAll_box03.Visible = True
28470       Case False
28480         .cmdExcelAll_box03.Visible = False
28490       End Select
28500       .cmdExcelAll_box04.Visible = True
28510       If blnPrintAll_Focus = False Then
28520         .cmdPrintAll_box01.Visible = False
28530         .cmdPrintAll_box02.Visible = False
28540         .cmdPrintAll_box03.Visible = False
28550         .cmdPrintAll_box04.Visible = False
28560       End If
28570       If blnWordAll_Focus = False Then
28580         .cmdWordAll_box01.Visible = False
28590         .cmdWordAll_box02.Visible = False
28600         .cmdWordAll_box03.Visible = False
28610         .cmdWordAll_box04.Visible = False
28620       End If
28630     End If
28640   End With

EXITP:
28650   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

28700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

28710   With Me
28720     .cmdExcelAll_box01.Visible = False
28730     .cmdExcelAll_box02.Visible = False
28740     .cmdExcelAll_box03.Visible = False
28750     .cmdExcelAll_box04.Visible = False
28760     blnExcelAll_Focus = False
28770   End With

EXITP:
28780   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub chkOpenExcel_AfterUpdate()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

28810   With Me
28820     Select Case .chkOpenExcel
          Case True
28830       .chkOpenExcel_lbl.FontBold = True
28840       .chkOpenExcel_lbl2.FontBold = True
28850       .chkOpenExcel_lbl_dim_hi.FontBold = True
28860       .chkOpenExcel_lbl2_dim_hi.FontBold = True
28870     Case False
28880       .chkOpenExcel_lbl.FontBold = False
28890       .chkOpenExcel_lbl2.FontBold = False
28900       .chkOpenExcel_lbl_dim_hi.FontBold = False
28910       .chkOpenExcel_lbl2_dim_hi.FontBold = False
28920     End Select
28930   End With

EXITP:
28940   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub chkOpenExcel_KeyDown(KeyCode As Integer, Shift As Integer)

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29020   intShiftDown = (Shift And acShiftMask) > 0
29030   intAltDown = (Shift And acAltMask) > 0
29040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29060     Select Case intRetVal
          Case vbKeyTab
29070       With Me
29080         intRetVal = 0
29090         .cmdClose.SetFocus
29100       End With
29110     Case vbKeyUp
29120       With Me
29130         intRetVal = 0
29140         .cmdExcelAll.SetFocus
29150       End With
29160     Case vbKeyDown
29170       With Me
29180         intRetVal = 0
29190         .cmdExcel00.SetFocus
29200       End With
29210     End Select
29220   End If

        ' ** Shift keys.
29230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29240     Select Case intRetVal
          Case vbKeyTab
29250       With Me
29260         intRetVal = 0
29270         .cmdExcelAll.SetFocus
29280       End With
29290     End Select
29300   End If

        ' ** Ctrl keys.
29310   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
29320     Select Case intRetVal
          Case vbKeyTab
29330       With Me
29340         intRetVal = 0
29350         If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
29360           .UserReportPath_chk.SetFocus
29370         Else
29380           Beep
29390         End If
29400       End With
29410     End Select
29420   End If

EXITP:
29430   KeyCode = intRetVal
29440   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

29510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
29520   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub UserReportPath_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29600 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_box_MouseMove"

29610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
29620   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub UserReportPath_chk_AfterUpdate()

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_AfterUpdate"

29710   With Me
29720     Select Case .UserReportPath_chk
          Case True
29730       .UserReportPath_chk_lbl1.FontBold = True
29740       .UserReportPath_chk_lbl1_dim.FontBold = True
29750       .UserReportPath_chk_lbl1_dim_hi.FontBold = True
29760       .UserReportPath_chk_lbl2.FontBold = True
29770       .UserReportPath_chk_lbl2_dim.FontBold = True
29780       .UserReportPath_chk_lbl2_dim_hi.FontBold = True
29790     Case False
29800       .UserReportPath_chk_lbl1.FontBold = False
29810       .UserReportPath_chk_lbl1_dim.FontBold = False
29820       .UserReportPath_chk_lbl1_dim_hi.FontBold = False
29830       .UserReportPath_chk_lbl2.FontBold = False
29840       .UserReportPath_chk_lbl2_dim.FontBold = False
29850       .UserReportPath_chk_lbl2_dim_hi.FontBold = False
29860     End Select
29870   End With

EXITP:
29880   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub UserReportPath_chk_KeyDown(KeyCode As Integer, Shift As Integer)

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29920   intShiftDown = (Shift And acShiftMask) > 0
29930   intAltDown = (Shift And acAltMask) > 0
29940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
29970       With Me
29980         intRetVal = 0
29990         .cmdClose.SetFocus
30000       End With
30010     End Select
30020   End If

        ' ** Shift keys.
30030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30040     Select Case intRetVal
          Case vbKeyTab
30050       With Me
30060         intRetVal = 0
30070         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
30080           .cmdExcelAll.SetFocus
30090         Else
30100           .cmdWordAll.SetFocus
30110         End If
30120       End With
30130     End Select
30140   End If

        ' ** Ctrl-Shift keys.
30150   If intCtrlDown And (Not intAltDown) And intShiftDown Then
30160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30170       With Me
30180         intRetVal = 0
30190         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
30200           .chkOpenExcel.SetFocus
30210         Else
30220           .chkOpenWord.SetFocus
30230         End If
30240       End With
30250     End Select
30260   End If

EXITP:
30270   KeyCode = intRetVal
30280   Exit Sub

ERRH:
1410    intRetVal = 0
1420    THAT_PROC = THIS_PROC
1430    That_Erl = Erl
1440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1450    Resume EXITP

End Sub

Private Sub UserReportPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_MouseMove"

30310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30320   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Public Sub PreviewOrPrint(strReportNumber As String, strControlName As String, intView As Integer)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "PreviewOrPrint"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String, strDocName2 As String
        Dim curTmp01 As Currency, curTmp02 As Currency, curTmp03 As Currency, curTmp04 As Currency, curTmp05 As Currency, curTmp06 As Currency
        Dim blnTmp07 As Boolean
        Dim blnRetVal As Boolean, blnContinue As Boolean
        Dim intRetVal_SummaryNew As Integer
        Dim intRetVal_BuildCourtReportData As Integer, intRetVal_BuildAssetListInfo As Integer

30410   blnContinue = True
30420   intRetVal_BuildCourtReportData = 0: intRetVal_BuildAssetListInfo = 0: intRetVal_SummaryNew = 0: blnTmp07 = False

30430   DoCmd.Hourglass True
30440   DoEvents

30450   With Me

          ' ** Validate the dates.
30460     If Validate = True Then  ' ** Function: Below.

30470       ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

            ' ** Set global variables for report headers.
30480       gdatStartDate = .DateStart
30490       gdatEndDate = .DateEnd
30500       gstrAccountNo = .cmbAccounts.Column(0)
30510       gstrAccountName = .cmbAccounts.Column(3)

30520       Select Case .chkGroupBy_IncExpCode
            Case True
30530         gblnUseReveuneExpenseCodes = True
30540       Case False
30550         gblnUseReveuneExpenseCodes = False
30560       End Select

30570       Set dbs = CurrentDb
            ' ** Empty tmpCourtReportData3.
30580       Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01b")
30590       qdf.Execute
            ' ** Empty tmpCourtReportData4.
30600       Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01c")
30610       qdf.Execute
            ' ** Empty tmpCourtReportData5.
30620       Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01d")
30630       qdf.Execute
            ' ** Empty tmpCourtReportData6.
30640       Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01e")
30650       qdf.Execute
30660       dbs.Close

            ' ** When called from the TAReports CommandBar, these variables should already be filled.
            ' **   gstrCrtRpt_Ordinal
            ' **   gstrCrtRpt_Version
            ' **   gstrCrtRpt_CashAssets_Beg
            ' **   gstrCrtRpt_NetIncome
            ' **   gstrCrtRpt_NetLoss
            ' **   gstrCrtRpt_CashAssets_End

30670       If strReportNumber <> "7" Then
              ' ** Build a new summary report table for everything
              ' ** except report 7, the property on hand report.
30680         If strReportNumber = "0" Then
30690           intRetVal_BuildAssetListInfo = BuildAssetListInfo_FL(Me.DateStart, .DateEnd, "Ending", strControlName)  ' ** Function: Below.
30700           Set dbs = CurrentDb
                ' ** qryCourtReport_FL_00_New_10_50 (tmpAssetList2, grouped and summed by assetno), grouped and summed.
30710           Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_10_51")
30720           Set rst = qdf.OpenRecordset
30730           With rst
30740             .MoveFirst
30750             curTmp04 = ![ICash]
30760             curTmp05 = ![PCash]
30770             curTmp06 = ![TotalCost]
30780             .Close
30790           End With
30800           dbs.Close
30810           intRetVal_BuildAssetListInfo = BuildAssetListInfo_FL(Me.DateStart, .DateEnd, "Ending", strControlName, "SummaryNew0")  ' ** Function: Below.
30820         Else
30830           intRetVal_BuildAssetListInfo = BuildAssetListInfo_FL(Me.DateStart, .DateEnd, "Ending", strControlName)  ' ** Function: Below.
30840         End If
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry.
              ' **   -9  Error.

30850         If intRetVal_BuildAssetListInfo = 0 Then
                ' ** I'm going to try calling BuildAssetListInfo_FL() before FLBuildCourtReportData()
                ' ** to assure qryMaxBalDates and/or qryCourtReport_05 get updated first.

                ' ** Disbursements.
30860           If Left(strReportNumber, 1) = "4" Then
30870             Set dbs = CurrentDb
30880             Select Case .opgType
                  Case .opgType_optRep.OptionValue
30890               intRetVal_BuildCourtReportData = FLBuildCourtReportData(strReportNumber, strControlName)  ' ** Module Function: modCourtReportsFL.
30900               If intRetVal_BuildCourtReportData <> 0 Then
                      ' ** Return Codes:
                      ' **   0  Success.
                      ' **  -1  Canceled.
                      ' **  -9  Error.
30910                 blnContinue = False
30920                 blnAllCancel = True
30930                 AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
30940               End If
30950             Case .opgType_optGuard.OptionValue
30960               intRetVal_BuildCourtReportData = FLBuildCourtReportData("4", strControlName)  ' ** Module Function: modCourtReportsFL.
30970               If intRetVal_BuildCourtReportData = 0 Then
                      ' ** Empty tmpCourtReportData2.
30980                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01a")
30990                 qdf.Execute
                      ' ** Append qryCourtReport_FL_04a to tmpCourtReportData2 (for rptCourtRptFL_04).
31000                 Set qdf = dbs.QueryDefs("qryCourtReport_FL_04f")
31010                 qdf.Execute
31020                 intRetVal_BuildCourtReportData = FLBuildCourtReportData("6", strControlName)  ' ** Module Function: modCourtReportsFL.

                      'Rep Report: qryCourtReport_FL_06a
                      'Trim(IIf(IsNull([RecurringItem]),"",IIf([journaltype]="Received",[RecurringItem],IIf([journaltype]="Paid",[RecurringItem],[RecurringItem]))) &
                      '  IIf([assetno] Is Not Null,IIf([assetdate] Is Not Null,Format([assetdate],"mm/dd/yyyy") & " ") &
                      '  IIf([shareface]-CLng([shareface])=0,Format([shareface],"#,##0"),Format([shareface],"#,##0.0000")) & " " &
                      '  CStr([Description]) &
                      '  IIf([rate]>0," " & Format([rate],"#,##0.000%")),"") &
                      '  IIf([due] Is Not Null,"  Due " & Format([due],"mm/dd/yyyy")) & " " &
                      '  [Jcomment])

                      'Guard Report: qryCourtReport_FL_06m
                      'totdesc:
                      'Trim(IIf(IsNull([Description])=True,'',Trim([Description])) & IIf(IsNull([Description])=True,'',' ') &
                      '  IIf(IsNull([jcomment])=True,'',Trim([jcomment])) &
                      '    IIf(IsNull([Description])=True And IsNull([jcomment])=True,'',IIf(IsNull([jcomment])=True,'',' ')) &
                      '  IIf(IsNull([RecurringItem])=True,'',Trim([RecurringItem])))

                      'new:
                      'Trim(
                      '  IIf(IsNull([RecurringItem])=True,'',Trim([RecurringItem]) & ' ') &
                      '  IIf(IsNull([assetno])=True,'',IIf(IsNull([assetdate])=True,'',Format([assetdate],'mm/dd/yyyy') & ' ')) &
                      '  IIf(IsNull([shareface])=True,'',IIf([Shareface]=0,'',IIf([shareface]-CLng([shareface])=0,Format([shareface],'#,##0'),Format([shareface],'#,##0.0000')) & ' ')) &
                      '  IIf(IsNull([Description])=True,'',Trim([Description]) & ' ') &
                      '  IIf(IsNull([rate])=True,'',IIf([rate]=0,'',Format([rate],'#,##0.000%') & ' ')) &
                      '  IIf(IsNull([due])=True,'',' Due ' & Format([due],'mm/dd/yyyy') & ' ') &
                      '  IIf(IsNull([jcomment])=True,'',[jcomment])
                      ')

31030                 If intRetVal_BuildCourtReportData = 0 Then
                        ' ** Append qryCourtReport_FL_06m to tmpCourtReportData2 (for rptCourtRptFL_06).
31040                   Set qdf = dbs.QueryDefs("qryCourtReport_FL_04g")
31050                   qdf.Execute
31060                 Else
                        ' ** Return Codes:
                        ' **   0  Success.
                        ' **  -1  Canceled.
                        ' **  -9  Error.
31070                   blnContinue = False
31080                   blnAllCancel = True
31090                   AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
31100                 End If
31110               Else
                      ' ** Return Codes:
                      ' **   0  Success.
                      ' **  -1  Canceled.
                      ' **  -9  Error.
31120                 blnContinue = False
31130                 blnAllCancel = True
31140                 AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
31150               End If
31160             End Select
31170           End If

31180           If intRetVal_BuildCourtReportData = 0 Then
31190             blnRebuildTable = False
                  ' ** Summary.
31200             If strReportNumber = "0" Then  'Or strReportNumber = "0A" Then
                    'THE 2ND HIT IS THROUGH HERE!

31210               intRetVal_SummaryNew = SummaryNew_FL(strControlName, blnRebuildTable, blnIsSummary, Me)  ' ** Module Function: modCourtReportsFL.

31220               If intRetVal_SummaryNew <> 0 Then
                      ' ** Return Codes:
                      ' **    0  Success.
                      ' **   -1  Canceled.
                      ' **   -2  No data.
                      ' **   -3  Missing entry.
                      ' **   -9  Error.
31230                 blnContinue = False
31240                 blnAllCancel = True
31250                 AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
31260               Else
31270                 If IsNull(gvarCrtRpt_FL_SpecData) = False Then
31280                   Set dbs = CurrentDb

                        ' ** VGC 11/27/2009: I HAVE BEEN UNABLE TO DIVINE WHAT MY ORIGINAL INTENT WAS.
                        ' ** THOUGH I KNOW, AND RICH CONCURS, THAT THIS FLORIDA REPORT DID GIVE
                        ' ** CORRECT INFORMATION AT ONE TIME, I HAVE NO IDEA NOW WHAT HAPPENED TO IT.
                        ' ** SINCE I CAN'T MAKE HEADS-NOR-TAILS OUT OF THE DATA BEING COLLECTED, AND
                        ' ** SECTION I. AND SECTION VI. JUST AREN'T SHOWING GOOD DATA, I'M JUST GOING
                        ' ** TO HAVE TO GO GRAB IT SOMEHOW, HERE, AND PUT IT INTO SECTIONS I. AND VI.
                        ' ** WE NEED A GOOD OPENING BALANCE SET, AND
                        ' ** WE NEED A GOOD CLOSING BALANCE SET.

31290                   Select Case .opgType
                        Case .opgType_optRep.OptionValue
                          ' ** tmpCourtReportData3, grouped and summed, for Sections II, III, IV, V. (SKIPS INC/EXP DETAILS!)
31300                     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_10_52")
31310                   Case .opgType_optGuard.OptionValue
                          ' ** tmpCourtReportData3, grouped and summed, for Sections II, III, IV. (SKIPS INC/EXP DETAILS!)
31320                     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_10_53")
31330                   End Select
31340                   Set rst = qdf.OpenRecordset
31350                   With rst
31360                     .MoveFirst
31370                     curTmp01 = ![ICash]      '  $3,426.26
31380                     curTmp02 = ![PCash]      '   ($953.82)
31390                     curTmp03 = ![Total]      '  $2,472.44
31400                     .Close
31410                   End With

                        ' ** Update tmpCourtReportData3, Section I, by specified [inctot], [printot], [costot].
31420                   Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_10_54")
31430                   With qdf.Parameters
31440                     ![inctot] = 0@ + Val(gstrCrtRpt_CashAssets_Beg)
31450                     ![printot] = (((curTmp04 + curTmp05 + curTmp06) - curTmp03) - Val(gstrCrtRpt_CashAssets_Beg))
31460                     ![costot] = 0@
31470                   End With
31480                   qdf.Execute

31490                   Select Case .opgType
                        Case .opgType_optRep.OptionValue
                          ' ** Update tmpCourtReportData3, Section VI, by specified [inctot], [printot], [costot].
31500                     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_10_55")
31510                   Case .opgType_optGuard.OptionValue
                          ' ** Update tmpCourtReportData3, Section V, by specified [inctot], [printot], [costot].
31520                     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_10_56")
31530                   End Select
31540                   With qdf.Parameters
31550                     ![inctot] = (curTmp01 + Val(gstrCrtRpt_CashAssets_Beg))
31560                     ![printot] = ((curTmp04 + curTmp05 + curTmp06) - (curTmp01 + Val(gstrCrtRpt_CashAssets_Beg)))
                          ' **                     $120,433.68           - $3,426.26 = ($117,007.42 - {entered cash})
31570                     ![costot] = curTmp06
                          ' ** 3rd column on report is:  Total_Tot: ([Income_Tot]+[Principal_Tot]); Cost_Tot isn't used on report.
31580                   End With
31590                   qdf.Execute

31600                   dbs.Close
31610                 End If
31620               End If
31630             End If
31640           Else
                  ' ** Return Codes:
                  ' **   0  Success.
                  ' **  -1  Canceled.
                  ' **  -9  Error.
31650             If intRetVal_BuildAssetListInfo = -2 Then
31660               If strReportNumber = "0B" Then  ' ** Asset List.
31670                 blnTmp07 = True
31680               Else
31690                 MsgBox "There is no data for this report.", vbInformation + vbOKOnly, "Nothing To Do"
31700               End If
31710             Else
                    'Debug.Print "'intRetVal_BuildAssetListInfo: " & intRetVal_BuildAssetListInfo
31720             End If
31730             blnContinue = False
31740             blnAllCancel = True
31750             AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
31760           End If
31770         Else
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.
31780           If intRetVal_BuildAssetListInfo = -2 Then
31790             If strReportNumber = "0B" Then  ' ** Asset List.
31800               blnTmp07 = True
31810             Else
31820               MsgBox "There is no data for this report", vbInformation + vbOKOnly, "Nothing To Do"
31830             End If
31840           Else
                  'Debug.Print "'intRetVal_BuildAssetListInfo: " & intRetVal_BuildAssetListInfo
31850           End If
31860           blnContinue = False
31870           blnAllCancel = True
31880           AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
31890         End If
31900       Else
              ' ** rptCourtRptFL_07: Assets on Hand at Close of Accounting Period.
31910         FLCourtReportLoad  ' ** Module Function: modCourReportsFL.
31920         intRetVal_BuildAssetListInfo = BuildAssetListInfo_FL(Me.DateStart, .DateEnd, "Ending", strControlName, THIS_PROC)  ' ** Function: Below.
31930         If intRetVal_BuildAssetListInfo = 0 Then
31940           Set dbs = CurrentDb
31950           Set qdf = dbs.QueryDefs("qryCourtReport_FL_07c")
                'OK .._CA_07c
31960           Set rst = qdf.OpenRecordset
31970           If rst.BOF = True And rst.EOF = True Then
                  ' ** No roll backs were required; current end date.
31980             strReportNumber = "7A"
31990           Else
                  ' ** Non-current end date.
32000           End If
32010           rst.Close
32020           dbs.Close
32030         Else
                ' ** Return codes:
                ' **    0  Success.
                ' **   -2  No data.
                ' **   -3  Missing entry, e.g., date.
                ' **   -9  Error.
32040           blnContinue = False
32050           blnAllCancel = True
32060           AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
32070         End If
32080       End If

32090       If blnContinue = True Then
              ' ** Set up sources for the NEW FL REPORT's.
32100         Select Case strReportNumber
              Case "0"
32110           GetBal_Beg gstrAccountNo  ' ** Procedure: Below.
32120           GetBal_End gstrAccountNo  ' ** Procedure: Below.
32130         Case "2", "2A"
32140           Set dbs = CurrentDb
                ' ** Empty tmpCourtReportData2.
32150           Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_New_07_01a")
32160           qdf.Execute
                ' ** Append qryCourtReport_FL_02m to tmpCourtReportData2.
32170           Set qdf = dbs.QueryDefs("qryCourtReport_FL_02mb")
32180           qdf.Execute
                ' ** Append qryCourtReport_FL_02oa to tmpCourtReportData2.
32190           Set qdf = dbs.QueryDefs("qryCourtReport_FL_02ob")
32200           qdf.Execute
                ' ** Append qryCourtReport_FL_02on to tmpCourtReportData2.
32210           Set qdf = dbs.QueryDefs("qryCourtReport_FL_02pb")
32220           qdf.Execute
32230           dbs.Close
32240         End Select
32250         If Right(strReportNumber, 1) = "A" Or Right(strReportNumber, 1) = "B" Then
32260           If strReportNumber <> "4A" Then
32270             strDocName = "rptCourtRptFL_" & Right("00" & strReportNumber, 3)
32280           Else
                  ' ** RecordSource =
32290             Select Case .opgType
                  Case .opgType_optRep.OptionValue
32300               strDocName = "rptCourtRptFL_04A"
32310             Case .opgType_optGuard.OptionValue
32320               strDocName = "rptCourtRptFL_04C"
32330             End Select
32340           End If
32350         Else
32360           If strReportNumber = "4" Then
32370             Select Case .opgType
                  Case .opgType_optRep.OptionValue
                    ' ** Schedule B: Disbursements; Representative.
                    ' ** RecordSource = qryCourtReport_FL_04a.
32380               strDocName = "rptCourtRptFL_04"
32390             Case .opgType_optGuard.OptionValue
                    ' ** Schedule B: Disbursements and Distributions; Guardian.
                    ' ** RecordSource = qryCourtReport_FL_04h.
32400               strDocName = "rptCourtRptFL_04B"
32410             End Select
32420           ElseIf strReportNumber = "0" Then
32430             Select Case .chkGroupBy_IncExpCode
                  Case True
32440               strDocName = "rptCourtRptFL_00A"
32450             Case False
32460               strDocName = "rptCourtRptFL_00"
32470             End Select
32480           Else
32490             strDocName = "rptCourtRptFL_" & Right("00" & strReportNumber, 2)
32500           End If
32510         End If

              ' ** When moving Asset List to print before Summary, save Summary's tmpAssetList2, then restore it.
32520         If ((strReportNumber = "0" And .chkAssetList = True) Or (strReportNumber = "0B")) Then
32530           Set dbs = CurrentDb
                ' ** Empty tmpAssetList3.
32540           Set qdf = dbs.QueryDefs("qryCourtReport_FL_07ca")
32550           qdf.Execute
                ' ** Append tmpAssetList2 to tmpAssetList3.
32560           Set qdf = dbs.QueryDefs("qryCourtReport_FL_07cb")
32570           qdf.Execute
                ' ** Asset List (Property on Hand at Beginning of Account Period).
32580           .chkAssetList_Start = True
32590           intRetVal_BuildAssetListInfo = BuildAssetListInfo_FL(#1/1/1900#, (CDate(Me.DateStart) - 1), "Beginning", strControlName, THIS_PROC)  ' ** Function: Below.
32600           If intRetVal_BuildAssetListInfo = 0 Then
                  ' ** rptCourtRptFL_00B: Asset List - Assets on Hand at Beginning of Account Period.
32610             strDocName = "rptCourtRptFL_00B"
32620             Select Case intView
                  Case acViewNormal
32630               If gblnCrtRpt_Zero = False Then
32640                 DoCmd.Hourglass False
32650                 DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
32660                 DoCmd.OpenReport strDocName, intView
32670               Else
32680                 If gblnCrtRpt_ZeroDialog = False Then
32690                   DoCmd.Hourglass False
32700                   DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
32710                   DoCmd.OpenReport strDocName, intView
32720                 Else
32730                   DoCmd.Hourglass False
32740                   DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
32750                   DoCmd.OpenReport strDocName, intView
                        ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                        ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                        'DoCmd.RunCommand acCmdPrint
32760                 End If
32770               End If
32780               If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
32790                 DoCmd.Close acReport, strDocName
32800               End If
32810             Case acViewPreview
32820               DoCmd.Hourglass False
32830               DoCmd.OpenReport strDocName, intView
32840               DoCmd.Maximize
32850               DoCmd.RunCommand acCmdFitToWindow
32860             End Select
32870             Select Case .chkGroupBy_IncExpCode
                  Case True
32880               strDocName = "rptCourtRptFL_00A"
32890             Case False
32900               strDocName = "rptCourtRptFL_00"
32910             End Select
32920           Else
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -2  No data.
                  ' **   -3  Missing entry, e.g., date.
                  ' **   -9  Error.
32930             If intRetVal_BuildAssetListInfo = -2 Then
32940               If strReportNumber = "0B" Then  ' ** Asset List.
32950                 blnTmp07 = True
32960               Else
32970                 MsgBox "There is no data for this report", vbInformation + vbOKOnly, "Nothing To Do"
32980               End If
32990             Else
                    'Debug.Print "'intRetVal_BuildAssetListInfo: " & intRetVal_BuildAssetListInfo
33000             End If
33010             blnContinue = False
33020           End If
33030           .chkAssetList_Start = False
                ' ** Empty tmpAssetList2.
33040           Set qdf = dbs.QueryDefs("qryCourtReport_03")
33050           qdf.Execute
33060           If strReportNumber <> "0B" Then  ' ** Skip if not printing Summary.
                  ' ** Append tmpAssetList3 to tmpAssetList2.
33070             Set qdf = dbs.QueryDefs("qryCourtReport_FL_07cc")
33080             qdf.Execute
33090           End If  ' ** 0B.
33100         End If  ' ** strReportNumber = '0', chkAssetList = True.

33110         If strReportNumber <> "0B" Then  ' ** Skip if not printing Summary.
                ' ** Now print the report.
                ' ** rptCourtRptFL_00: Summary of Account.
33120           Select Case intView
                Case acViewNormal
33130             If gblnCrtRpt_Zero = False Then
33140               DoCmd.Hourglass False
33150               DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
33160               DoCmd.OpenReport strDocName, intView
33170             Else
33180               If gblnCrtRpt_ZeroDialog = False Then
33190                 DoCmd.Hourglass False
33200                 DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
33210                 DoCmd.OpenReport strDocName, intView
33220               Else
33230                 DoCmd.Hourglass False
33240                 DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
33250                 DoCmd.OpenReport strDocName, intView
                      ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                      ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                      'DoCmd.RunCommand acCmdPrint
33260               End If
33270             End If
33280             If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
33290               DoCmd.Close acReport, strDocName
33300             End If
33310           Case acViewPreview
33320             DoCmd.Hourglass False
33330             DoCmd.OpenReport strDocName, intView
33340             DoCmd.Maximize
33350             DoCmd.RunCommand acCmdFitToWindow
33360           End Select
33370         End If  ' ** 0B.

33380       End If  ' ** blnContinue.

33390       If strReportNumber <> "0B" Then  ' ** Skip if not printing Summary.
33400         If blnContinue = False And intRetVal_BuildAssetListInfo = -2 Then
                'FALSE FOR ANY NUMBER OF REASONS, -2 ALSO FOR MULTIPLE REASONS!
                'YET IT CONTINUES!?

33410           If Right(strReportNumber, 1) = "A" Or Right(strReportNumber, 1) = "B" Then
33420             If strReportNumber <> "4A" Then
33430               strDocName = "rptCourtRptFL_" & Right("00" & strReportNumber, 3)
33440             Else
33450               Select Case .opgType
                    Case .opgType_optRep.OptionValue
                      ' ** Schedule B: Disbursements, grouped by Income/Expense codes; Representative.
33460                 strDocName = "rptCourtRptFL_04A"
33470               Case .opgType_optGuard.OptionValue
                      ' ** Schedule B: Disbursements and Distributions, grouped by Income/Expense codes; Guardian.
33480                 strDocName = "rptCourtRptFL_04C"
33490               End Select
33500             End If
33510           Else
33520             If strReportNumber = "4" Then
33530               Select Case .opgType
                    Case .opgType_optRep.OptionValue
33540                 strDocName = "rptCourtRptFL_04"
33550               Case .opgType_optGuard.OptionValue
33560                 strDocName = "rptCourtRptFL_04B"
33570               End Select
33580             ElseIf strReportNumber = "0" Then
33590               Select Case .chkGroupBy_IncExpCode
                    Case True
33600                 strDocName = "rptCourtRptFL_00A"
33610               Case False
33620                 strDocName = "rptCourtRptFL_00"
33630               End Select
33640             Else
33650               strDocName = "rptCourtRptFL_" & Right("00" & strReportNumber, 2)
33660             End If
33670           End If

33680           If strDocName = "rptCourtRptFL_00" Or strDocName = "rptCourtRptFL_00A" Or strDocName = "rptCourtRptFL_07" Then
                  ' ** Empty tmpAssetList2.
33690             Set dbs = CurrentDb
33700             Set qdf = dbs.QueryDefs("qryCourtReport_03")
33710             qdf.Execute
33720             dbs.Close
33730           End If

                'WHY HASN'T THIS ALREADY BEEN PRINTED ABOVE?
                'BECAUSE IF IT HAD, IT WOULDN'T BE blnContinue = FALSE?
                ' ** Now print the report.
33740           If .chkAssetList = True And (strDocName = "rptCourtRptFL_00" Or strDocName = "rptCourtRptFL_00A") Then
                  ' ** rptCourtRptFL_00B: Asset List - Property on Hand at Beginning of Account Period.
33750             strDocName2 = "rptCourtRptFL_00B"
33760             Select Case intView
                  Case acViewNormal
33770               If gblnCrtRpt_Zero = False Then
33780                 DoCmd.Hourglass False
33790                 DoCmd.OpenReport strDocName2, acViewPreview  ' ** This will get the caption changed!
33800                 DoCmd.OpenReport strDocName2, intView
33810               Else
33820                 If gblnCrtRpt_ZeroDialog = False Then
33830                   DoCmd.Hourglass False
33840                   DoCmd.OpenReport strDocName2, acViewPreview  ' ** This will get the caption changed!
33850                   DoCmd.OpenReport strDocName2, intView
33860                 Else
33870                   DoCmd.Hourglass False
33880                   DoCmd.OpenReport strDocName2, acViewPreview  ' ** This will get the caption changed!
33890                   DoCmd.OpenReport strDocName2, intView
                        ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                        ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                        'DoCmd.RunCommand acCmdPrint
33900                 End If
33910               End If
33920               If IsLoaded(strDocName2, acReport) = True Then  ' ** Module Function: modFileUtilities.
33930                 DoCmd.Close acReport, strDocName2
33940               End If
33950             Case acViewPreview
33960               DoCmd.Hourglass False
33970               DoCmd.OpenReport strDocName2, intView
33980               DoCmd.Maximize
33990               DoCmd.RunCommand acCmdFitToWindow
34000             End Select
34010           End If

                ' ** rptCourtRptFL_..: All the rest.
34020           Select Case intView
                Case acViewNormal
34030             If gblnCrtRpt_Zero = False Then
34040               DoCmd.Hourglass False
34050               DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
34060               DoCmd.OpenReport strDocName, intView
34070             Else
34080               If gblnCrtRpt_ZeroDialog = False Then
34090                 DoCmd.Hourglass False
34100                 DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
34110                 DoCmd.OpenReport strDocName, intView
34120               Else
34130                 DoCmd.Hourglass False
34140                 DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
34150                 DoCmd.OpenReport strDocName, intView
                      ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                      ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                      'DoCmd.RunCommand acCmdPrint
34160               End If
34170             End If
34180             If IsLoaded(strDocName, acReport) = True Then  ' ** Module Function: modFileUtilities.
34190               DoCmd.Close acReport, strDocName
34200             End If
34210           Case acViewPreview
34220             DoCmd.Hourglass False
34230             DoCmd.OpenReport strDocName, intView
34240             DoCmd.Maximize
34250             DoCmd.RunCommand acCmdFitToWindow
34260           End Select

34270         End If

34280       Else
34290         If blnTmp07 = True Then
                ' ** Asset List, no data. rptCourtRptFL_00B
34300           blnTmp07 = False
34310           Select Case intView
                Case acViewNormal
34320             DoCmd.Hourglass False
34330             DoCmd.OpenReport strDocName, acViewPreview  ' ** This will get the caption changed!
34340             DoCmd.OpenReport strDocName, intView
34350           Case acViewPreview
34360             DoCmd.Hourglass False
34370             DoCmd.OpenReport strDocName, intView
34380             DoCmd.Maximize
34390             DoCmd.RunCommand acCmdFitToWindow
34400           End Select
34410         End If
34420       End If  ' ** 0B.

34430     Else
34440       blnAllCancel = True
34450       AllCancelSet2_FL blnAllCancel  ' ** Module Procedure: modCourtReportsFL.
34460     End If

34470   End With  ' ** Me.

34480   DoCmd.Hourglass False

EXITP:
        ' ** Check some Public variables in case there's been an error.
34490   blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
34500   blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
34510   Set rst = Nothing
34520   Set qdf = Nothing
34530   Set dbs = Nothing
34540   Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case 2202  ' ** You must set a default printer before you design, print, or preview.
1430      MsgBox "There does not appear to be a default printer defined." & vbCrLf & vbCrLf & _
            "Please define a default printer before trying to print this report.", vbCritical + vbOKOnly, "Printer Not Found"
1440    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
1450    Case Else
1460      THAT_PROC = THIS_PROC
1470      That_Erl = Erl
1480      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1490    End Select
1500    Resume EXITP

End Sub

Public Function Validate() As Boolean

34600 On Error GoTo ERRH

        Const THIS_PROC As String = "Validate"

        Dim blnRetVal As Boolean

34610   blnRetVal = False

        ' ** Validate main three data values.
34620   With Me
34630     If IsNull(.DateStart.Value) Then
34640       DoCmd.Hourglass False
34650       MsgBox "You must enter a From date to continue.", vbInformation + vbOKOnly, "Entry Required"
34660       .DateStart.SetFocus
34670     Else
34680       If IsNull(.DateEnd.Value) Then
34690         DoCmd.Hourglass False
34700         MsgBox "You must enter a To date to continue.", vbInformation + vbOKOnly, "Entry Required"
34710         .DateEnd.SetFocus
34720       Else
34730         If .DateStart.Value >= .DateEnd.Value Then
34740           DoCmd.Hourglass False
34750           MsgBox "You must enter a From date that is less than the To date to continue.", vbInformation + vbOKOnly, "Invalid date"
34760           .DateEnd.SetFocus
34770         Else
34780           If IsNull(.cmbAccounts.Column(0)) Or .cmbAccounts.Column(0) = vbNullString Then
34790             DoCmd.Hourglass False
34800             MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
34810             .cmbAccounts.SetFocus
34820           Else
34830             blnRetVal = True
34840           End If
34850         End If
34860       End If
34870     End If
34880   End With

EXITP:
34890   Validate = blnRetVal
34900   Exit Function

ERRH:
1410    DoCmd.Hourglass False
1420    blnRetVal = False
1430    THAT_PROC = THIS_PROC
1440    That_Erl = Erl
1450    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1460    Resume EXITP

End Function

Public Function BuildAssetListInfo_FL(varStartDate As Variant, varEndDate As Variant, strWhen As String, strProc As String, Optional varProc2 As Variant) As Integer
' ** Copied from frmStatementParameters to match rollback for Asset List.
' ** All code that referenced 'All Accounts' has been removed.
' ** Return Codes:
' **    0  Success.
' **   -2  No data.
' **   -3  Missing entry.
' **   -9  Error.

35000 On Error GoTo ERRH

        Const THIS_PROC As String = "BuildAssetListInfo_FL"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim rstWork As DAO.Recordset
        Dim rstTempAssetList As DAO.Recordset
        Dim rstTmpAccountInfo As DAO.Recordset
        Dim rstAccount As DAO.Recordset
        Dim rstMasterAsset As DAO.Recordset
        Dim strSQL As String, strWorkSQL As String
        Dim strButton As String
        Dim blnContinue As Boolean
        Dim blnNoMAsset As Boolean  ' ** TRUE if no MasterAsset involved.
        Dim intPos01 As Integer
        Dim intTmp01 As Integer
        Dim intRetVal_SetDateSpecificSQL As Integer
        Dim intRetVal As Integer

35010   intRetVal = 0   ' ** Unless proven otherwise.
35020   blnContinue = True

35030   lngBuildCnt = lngBuildCnt + 1&
        'Debug.Print "'CNT BuildAssetListInfo_FL: " & CStr(lngBuildCnt) & "  " & strProc & "  " & _
        '  IIf(IsMissing(varProc2) = True, vbNullString, varProc2) & "  intSpecSQLCnt: " & CStr(intSpecSQLCnt)

35040   strButton = strProc
35050   intPos01 = InStr(strButton, " - ")
35060   If intPos01 > 0 Then
35070     strButton = "cmdPrintAll"
35080   Else
35090     intPos01 = InStr(strButton, "_Click")
35100     strButton = Left(strButton, (intPos01 - 1))
35110   End If

35120   If Me.cmbAccounts.Visible = True And blnAccountChecked = False Then
35130     blnAccountChecked = True
35140     Me.cmbAccounts.SetFocus
35150     If Me.cmbAccounts.text = vbNullString Then
35160       DoCmd.Hourglass False
35170       MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
35180       intRetVal = -3  ' ** Missing entry.
35190     End If
35200   End If

35210   If intRetVal = 0 Then
35220     If IsNull(varEndDate) Then
35230       DoCmd.Hourglass False
35240       MsgBox "Must enter Period " & strWhen & " date to continue.", vbInformation + vbOKOnly, "Entry Required"
35250       Me.DateEnd.SetFocus
35260       intRetVal = -3  ' ** Missing entry.
35270     End If
35280   End If

35290   If intRetVal = 0 Then
          ' ** This code will update the qryMaxBalDates query to give
          ' ** us the balance numbers from the previous statement.
          '1st Time:
35300     intRetVal_SetDateSpecificSQL = SetDateSpecificSQL(Me.cmbAccounts, "Statements", THIS_NAME, varStartDate, varEndDate)  ' ** Module Function: modUtilities.
          ' ** Return codes:
          ' **    0  Success.
          ' **   -2  No data.
          ' **   -4  Date criteria not met.
          ' **   -9  Error.
35310     If IsMissing(varProc2) = False Then
35320       If InStr(varProc2, "SummaryNew") > 0 Then
35330         If InStr(strProc, "Word") > 0 Or InStr(strProc, "Excel") > 0 Then
                ' ** Unlike Preview and Print, the Word sequence only hits this the '2nd' time,
                ' ** which has the effect of only allowing Word through every other attempt!
35340           intSpecSQLCnt = 1
35350         End If
35360         If intSpecSQLCnt = 0 Or intSpecSQLCnt >= 2 Then
35370           intSpecSQLCnt = 1
35380         ElseIf intSpecSQLCnt = 1 Then
                ' ** The 2nd hit from SetDateSpecificSQL().
35390           intSpecSQLCnt = intSpecSQLCnt + 1
35400           If intRetVal_SetDateSpecificSQL = -4 Then

                  ' ** The actual 'no data' hit occurs in CheckFirstAcctBal(), in modUtilities.
35410             blnContinue = False
35420             intRetVal = -2  ' ** No data.

                  ' ** The Summary only messes up the 2nd time it hits this.
                  ' ** The 1st one is fine, but on a new installation, the
                  ' ** 2nd one finds no previous data whatsoever, and just
                  ' ** doesn't like that.

                  ' ** It's got to first finish the rest of this, then finish
                  ' ** SummaryNew_FL(), and finally return to PreviewOrPrint() finish that!
                  ' ** Somewhere in all of it I've got to get the right data to print!

35430           Else
35440             If intRetVal_SetDateSpecificSQL <> 0 Then
35450               blnContinue = False
35460               intRetVal = -2  ' ** No data.
35470             End If
35480           End If
35490         End If
35500       Else
35510         If intRetVal_SetDateSpecificSQL <> 0 Then
35520           blnContinue = False
35530           intRetVal = -2  ' ** No data.
35540         End If
35550       End If
35560     Else
35570       If intRetVal_SetDateSpecificSQL <> 0 Then
35580         blnContinue = False
35590         intRetVal = -2  ' ** No data.
35600       End If
35610     End If
35620   End If

35630   If intRetVal = 0 And blnContinue = True Then

35640     Me.Controls(strButton).SetFocus

35650     Set dbs = CurrentDb

          ' ** Empty tmpAssetList2.
35660     If TableExists("tmpAssetList2") = True Then  ' ** Module Function: modUtilities.
35670       Set qdf = dbs.QueryDefs("qryCourtReport_03")
35680       qdf.Execute
35690     End If
          ' ** Empty tmpAccountInfo.
35700     If TableExists("tmpAccountInfo") = True Then  ' ** Module Function: modUtilities.
35710       Set qdf = dbs.QueryDefs("qryCourtReport_04")
35720       qdf.Execute
35730     End If

          ' ** qryCourtReport_FL_07z.
          ' ** VGC 11/25/2009: «='90',-1,1)» to «='90',1,1)».
35740     strSQL = "SELECT ActiveAssets.assetno, masterasset.description AS MasterAssetDescription, masterasset.due, " & _
            "masterasset.rate, Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, " & _
            "Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface]))*" & _
            "IIf([assettype].[assettype]='90',1,1) AS TotalShareface, account.accountno, account.shortname, " & _
            "account.legalname, assettype.assettype, assettype_description, " & _
            "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
            "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
            "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc, "
35750     strSQL = strSQL & "account.icash, account.pcash, masterasset.currentDate, " & _
            "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX, " & _
            "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX, " & _
            "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX, " & CoInfo & " "
35760     strSQL = strSQL & "FROM (masterasset LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) " & _
            "RIGHT JOIN (account LEFT JOIN ActiveAssets ON account.accountno = ActiveAssets.accountno) " & _
            "ON masterasset.assetno = ActiveAssets.assetno " & _
            "GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, " & _
            "account.accountno, account.shortname, account.legalname, assettype.assettype, assettype_description, " & _
            "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & "
35770     strSQL = strSQL & "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
            "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), " & _
            "account.icash, account.pcash, masterasset.currentDate, " & _
            "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), " & _
            "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]), " & _
            "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) " & _
            "HAVING (((account.accountno)='" & Me.cmbAccounts & "'));"
35780     strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.

35790     Set rst = dbs.OpenRecordset(strSQL)  ' ** This is what's used for appending to the tmpAssetList2 table and rstTempAssetList.

          ' ** Ledger records to roll back if after specified 'To' date. (VGC 08/24/2010: CHANGES!)
35800     Set qdf = dbs.QueryDefs("qryCourtReport_FL_07d")
35810     With qdf.Parameters
35820       ![actno] = Me.cmbAccounts
35830       ![datEnd] = CDate(varEndDate)
35840     End With
35850     Set rstWork = qdf.OpenRecordset  'THIS SUPPLIES THE BASIS FOR ROLLING BACK!
35860     If rstWork.BOF = True And rstWork.EOF = True Then
            ' ** In the absense of newer data to roll back, take the total cost as it's normally computed for report 0.
35870       blnContinue = False  ' ** Leave intRetVal = 0 so that the procedure calling this continues.
35880       rstWork.Close
35890       If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
35900         dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
35910       End If
35920       rst.Close
35930       dbs.Close
35940     End If

35950     If blnContinue Then  ' ** There are transactions after the specified To date.
            ' ** Rollbacks only.

35960       Set rstAccount = dbs.OpenRecordset("SELECT account.* FROM account ORDER BY accountno", dbOpenSnapshot)
35970       Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_01")  'NO CRITERIA.
35980       Set rstMasterAsset = qdf.OpenRecordset()

35990       If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
36000         dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
36010         Set rstTempAssetList = dbs.OpenRecordset("SELECT tmpAssetList2.* FROM tmpAssetList2", dbOpenDynaset)
36020       Else
              ' ** ERROR.
36030         intRetVal = -9  ' ** Error.
              'MsgBox "Unable to create temporary table for reporting.", vbCritical + vbOKOnly, "Error"
36040         rstAccount.Close
36050         rstMasterAsset.Close
36060         rstWork.Close
36070         dbs.Close
36080       End If

36090       If intRetVal = 0 Then
              'At this point, tmpAssetList2 shows TotalShareFace = 20

              ' ** Create table to contain temporary account master records to track "global" account info.
36100         If MakeTempTable(rst, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
36110           dbs.Execute "INSERT INTO tmpAccountInfo " & strSQL  ' ** Copy data from query.
                ' ** Now, make records here distinct for each account.
36120           dbs.Execute "UPDATE tmpAccountInfo SET assetno = NULL, MasterAssetDescription = NULL, " & _
                  "due = NULL, rate = NULL, TotalCost = NULL, TotalShareFace = Null,  " & _
                  "assettype = NULL, assettype_description = NULL, TotDesc = NULL,  " & _
                  "currentDate = NULL, MarketValueX = NULL, MarketValueCurrentX = NULL, " & _
                  "YieldX = NULL  " & _
                  "WHERE True"
36130           Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT DISTINCTROW tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenSnapshot)
                ' ** Empty tmpAccountInfo.
36140           Set qdf = dbs.QueryDefs("qryCourtReport_04")
36150           qdf.Execute
36160           If rstTmpAccountInfo.RecordCount > 0 Then
36170             rstTmpAccountInfo.MoveLast
36180             rstTmpAccountInfo.MoveFirst
36190           End If
36200           If Not CopyToTempTable(rstTmpAccountInfo, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
                  ' ** ERROR.
36210             DoCmd.Hourglass False
36220             intRetVal = -9  ' ** Error.
36230             MsgBox "Unable to copy data to temporary table.", vbCritical + vbOKOnly, "Error"
36240           Else
36250             rstTmpAccountInfo.Close  ' ** Close old snapshot, then open as dynaset.
36260             Set rstTmpAccountInfo = Nothing
36270             Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenDynaset)
36280           End If
36290         Else
                ' ** ERROR.
36300           DoCmd.Hourglass False
36310           intRetVal = -9  ' ** Error.
36320           MsgBox "Unable to create temporary account table for reporting.", vbCritical + vbOKOnly, "Error"
36330           rstAccount.Close
36340           rstMasterAsset.Close
36350           rstWork.Close
36360           rstTempAssetList.Close
36370           dbs.Close
36380         End If

36390       End If

36400       If intRetVal = 0 Then

36410         rstTmpAccountInfo.MoveLast
36420         rstTmpAccountInfo.MoveFirst

              ' ** Get total number of records in recordset of subsequent ledger entries.
36430         rstWork.MoveLast
36440         rstWork.MoveFirst

              ' ** MORE KLUGE!
              ' ** The intSpecSQLCnt variable was created to cut down on the number of times certain sequences get repeated.
              ' ** It turns out, that with all the juggling going on with the Guardian summary, a particular sequence gets
              ' ** hit an extra time, bringing intSpecSQLCnt up to 2 by the time it gets here to total the beginning balance.
              ' ** So, ... I'm manually resetting it to 1 in order to get the correct beginning balance.
36450         If blnIsSummary = True Then
36460           If Me.opgType = Me.opgType_optRep.OptionValue Then
36470             intTmp01 = intSpecSQLCnt
36480           ElseIf Me.opgType = Me.opgType_optGuard.OptionValue Then
36490             intTmp01 = 1
36500           End If
36510         End If

36520         If intTmp01 <> 2 Then

36530           Do While Not rstWork.EOF  ' ** Move through each record, tracking changes.
                  'This is where those 20 shares get subtracted!

36540             rstTempAssetList.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "' AND [assetno] = " & CStr(rstWork![assetno])
36550             If rstTempAssetList.NoMatch Then
                    ' ** We need a new record because this asset was not in the reporting query after subsequent changes.
36560               rstMasterAsset.FindFirst "[assetno] = " & CStr(rstWork![assetno])
36570               blnNoMAsset = rstMasterAsset.NoMatch
36580               If blnNoMAsset Then
36590                 Select Case Trim(rstWork![journaltype])
                      Case "Misc.", "Paid", "Received"
                        ' ** OK to continue.
36600                 Case Else
                        ' ** ERROR.
36610                   DoCmd.Hourglass False
36620                   intRetVal = -9  ' ** Error.
36630                   MsgBox "Missing Master Asset record.", vbCritical + vbOKOnly, "Error"
36640                   rstAccount.Close
36650                   rstMasterAsset.Close
36660                   rstWork.Close
36670                   rstTempAssetList.Close
36680                   rstTmpAccountInfo.Close
36690                   dbs.Close
36700                 End Select
36710               Else
36720                 rstTempAssetList.AddNew
36730                 rstTempAssetList![accountno] = rstWork![accountno]
36740                 rstTempAssetList![assetno] = rstMasterAsset![assetno]
36750                 rstTempAssetList![MasterAssetDescription] = rstMasterAsset![description]
36760                 rstTempAssetList![assettype] = rstMasterAsset![assettype]
36770                 rstTempAssetList![assettype_description] = rstMasterAsset![assettype_description]
36780                 rstTempAssetList![due] = rstMasterAsset![due]
36790                 rstTempAssetList![rate] = rstMasterAsset![rate]
36800                 rstTempAssetList![totdesc] = CStr(rstMasterAsset![description]) & _
                        IIf(rstMasterAsset![rate] > 0, " " & Format(rstMasterAsset![rate], "0.000%"), "") & _
                        IIf(Not IsNull(rstMasterAsset![due]), "  Due " & Format(rstMasterAsset![due], "mm/dd/yyyy"), "")
36810                 rstTempAssetList![TotalCost] = 0
36820                 rstTempAssetList![TotalShareFace] = 0
36830                 rstTempAssetList![ICash] = 0
36840                 rstTempAssetList![PCash] = 0
                      ' ** LEAVE in editing mode.
36850               End If
36860             Else
36870               blnNoMAsset = False
36880               rstTempAssetList.Edit  ' ** Edit existing temp record.
36890             End If

36900             If intRetVal = 0 Then

36910               rstTmpAccountInfo.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"

36920               If rstTmpAccountInfo.NoMatch Then
                      ' ** We need a new account record.

36930                 rstAccount.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"
36940                 If rstAccount.NoMatch Then
                        ' ** ERROR.
36950                   DoCmd.Hourglass False
36960                   intRetVal = -9  ' ** Error.
36970                   MsgBox "Missing Account master record.", vbCritical + vbOKOnly, "Error"
36980                   rstAccount.Close
36990                   rstMasterAsset.Close
37000                   rstWork.Close
37010                   rstTempAssetList.Close
37020                   rstTmpAccountInfo.Close
37030                   dbs.Close
37040                 Else
37050                   rstTmpAccountInfo.AddNew
                        ' ** Copy info from Account table.
37060                   rstTmpAccountInfo![accountno] = rstAccount![accountno]
37070                   rstTmpAccountInfo![shortname] = rstAccount![shortname]
37080                   rstTmpAccountInfo![legalname] = rstAccount![legalname]
37090                   rstTmpAccountInfo![ICash] = rstAccount![ICash]
37100                   rstTmpAccountInfo![PCash] = rstAccount![PCash]
                        ' ** LEAVE in editing mode.
37110                 End If
37120               Else
37130                 rstTmpAccountInfo.Edit  ' ** Edit existing temp record.
37140               End If

37150             End If  ' ** intRetVal.

37160             If intRetVal = 0 Then

37170               If (Not blnNoMAsset) And (Trim(rstWork![journaltype]) <> "Received") Then
                      ' ** Add/Subtract info from ASSET record. NOTE: These are SUBTRACTIONS
                      ' ** because the query returns the changes since the date.
                      'VGC 08/20/2010: CHANGES!
37180                 rstTempAssetList![TotalShareFace] = _
                        rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, _
                        rstWork![shareface] * IIf(rstTempAssetList![assettype] = 90, 1, 1)))
                      'rstTempAssetList![TotalShareFace] = _
                      '  rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, rstWork![shareface] * _
                      '  IIf(rstTempAssetList![assettype] = 90, -1, 1)))
37190                 rstTempAssetList![TotalCost] = rstTempAssetList![TotalCost] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost])
                      ' ** Save ASSET temp table record.
37200                 rstTempAssetList.Update
37210               End If

                    ' ** Add/Subtract info from ACCOUNT record. NOTE: These are SUBTRACTIONS
                    ' ** because the query returns the changes since the date.
37220               rstTmpAccountInfo![ICash] = rstTmpAccountInfo![ICash] - rstWork![ICash]
37230               rstTmpAccountInfo![PCash] = rstTmpAccountInfo![PCash] - rstWork![PCash]
                    ' ** Save ACCOUNT temp table record.
37240               rstTmpAccountInfo.Update

37250               rstWork.MoveNext

37260             End If  ' ** intRetVal.

37270             If intRetVal <> 0 Then Exit Do

37280           Loop  ' ** Move through each record, tracking changes.

37290         End If

37300       End If  ' ** intRetVal.

37310       If intRetVal = 0 Then

              ' ** Remove any asset records which now have a ZERO totalshareface AND totalcost.
37320         strWorkSQL = "DELETE tmpAssetList2.*, tmpAssetList2.TotalShareface, tmpAssetList2.TotalCost " & _
                "FROM tmpAssetList2 " & _
                "WHERE ((Round((tmpAssetList2.TotalShareface),4)=0) AND (Round((tmpAssetList2.TotalCost),2)=0));"
              ' ** 07/21/2008: Added rounding.
37330         dbs.Execute strWorkSQL

              ' ** Now, update tmpAssetList2 to reflect the needed account information which is
              ' ** global to the report.
37340         rstTmpAccountInfo.MoveFirst
37350         While Not rstTmpAccountInfo.EOF
37360           strWorkSQL = "UPDATE tmpAssetList2 SET " & _
                  "shortname = " & SQLFormatStr(rstTmpAccountInfo![shortname], dbText) & ", " & _
                  "legalname = " & SQLFormatStr(rstTmpAccountInfo![legalname], dbText) & ", " & _
                  "icash = " & SQLFormatStr(rstTmpAccountInfo![ICash], dbCurrency) & ", " & _
                  "pcash = " & SQLFormatStr(rstTmpAccountInfo![PCash], dbCurrency) & " " & _
                  "WHERE accountno = " & SQLFormatStr(rstTmpAccountInfo![accountno], dbText) & ";"
37370           dbs.Execute strWorkSQL
37380           rstTmpAccountInfo.MoveNext
37390         Wend

              ' ** Close temp & working recordsets.
37400         rstAccount.Close
37410         rstMasterAsset.Close
37420         rstWork.Close
37430         rstTempAssetList.Close
37440         rstTmpAccountInfo.Close

              ' ** Finally, base report on this instead of on qryAssetList.
37450         Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_Ba")
37460         Set rst = qdf.OpenRecordset()            ' ** USED DIRECTLY BELOW.

37470       End If  ' ** intRetVal.

37480     End If  ' ** blnContinue.

37490     If intRetVal = 0 And blnContinue = True Then
37500       If rst.BOF = True And rst.EOF = True Then  ' ** USED HERE!
              ' ** On the 2nd go through from SummaryNew_FL(), tmpAssetList2 table is empty!
              ' ** If this is a new account, 'No Data' will be expected, so don't stop!
              ' ** If this is a PrintAll, continue and print empty report!
37510         If Left(strProc, 11) <> "cmdPrintAll" Then
37520           DoCmd.Hourglass False
37530           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
37540         End If
37550         intRetVal = -2  ' ** No data.
37560       End If
37570       rst.Close
37580       dbs.Close
37590     End If

37600   End If  ' ** intRetVal.

EXITP:
37610   Set rst = Nothing
37620   Set dbs = Nothing
37630   BuildAssetListInfo_FL = intRetVal
37640   Exit Function

ERRH:
1410    DoCmd.Hourglass False
1420    intRetVal = -9  ' ** Error.
1430    THAT_PROC = THIS_PROC
1440    That_Erl = Erl
1450    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1460    Resume EXITP

End Function

Public Sub GetBal_Beg(strAccountNo As String)

37700 On Error GoTo ERRH

        Const THIS_PROC As String = "GetBal_Beg"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef

37710   With Me
37720     gstrAccountNo = strAccountNo
37730     Set dbs = CurrentDb
          ' ** Get a beginning balance.
37740     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_24")  ' ** Ledger, by accountno.
37750     Set rst = qdf.OpenRecordset
37760     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in Ledger.
37770       .PropOnHand_Beg_Inc = 0@
37780       .PropOnHand_Beg_Prin = 0@
37790       .PropOnHand_Beg_Cost = 0@
37800     Else
37810       rst.MoveFirst
37820       If IsNull(rst![Sumshareface_dp]) = False Then
37830         .PropOnHand_Beg_Inc = rst![SumIcash]
37840         .PropOnHand_Beg_Prin = rst![SumPcash]
37850         .PropOnHand_Beg_Cost = rst![SumCost]
37860       Else
37870         .PropOnHand_Beg_Inc = 0@
37880         .PropOnHand_Beg_Prin = 0@
37890         .PropOnHand_Beg_Cost = 0@
37900       End If
37910     End If
37920     rst.Close
37930     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_28")  ' ** Archive, by accountno.
37940     Set rst = qdf.OpenRecordset
37950     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in LedgerArchive.
37960     Else
37970       rst.MoveFirst
37980       If IsNull(rst![Sumshareface_dp]) = False Then
37990         .PropOnHand_Beg_Inc = (.PropOnHand_Beg_Inc + rst![SumIcash])
38000         .PropOnHand_Beg_Prin = (.PropOnHand_Beg_Prin + rst![SumPcash])
38010         .PropOnHand_Beg_Cost = (.PropOnHand_Beg_Cost + rst![SumCost])
38020       End If
38030     End If
38040     rst.Close
38050     dbs.Close
38060   End With

EXITP:
38070   Set rst = Nothing
38080   Set qdf = Nothing
38090   Set dbs = Nothing
38100   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Public Sub GetBal_End(strAccountNo As String)

38200 On Error GoTo ERRH

        Const THIS_PROC As String = "GetBal_End"

        Dim dbs As DAO.Database, rst As DAO.Recordset, qdf As DAO.QueryDef

38210   With Me
38220     gstrAccountNo = strAccountNo
38230     Set dbs = CurrentDb
          ' ** Get an ending balance.
38240     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_44")  ' ** Ledger, by accountno.
38250     Set rst = qdf.OpenRecordset
38260     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in Ledger.
38270       .PropOnHand_End_Inc = 0@
38280       .PropOnHand_End_Prin = 0@
38290       .PropOnHand_End_Cost = 0@
38300     Else
38310       rst.MoveFirst
38320       If IsNull(rst![Sumshareface_dp]) = False Then
38330         .PropOnHand_End_Inc = rst![SumIcash]
38340         .PropOnHand_End_Prin = rst![SumPcash]
38350         .PropOnHand_End_Cost = rst![SumCost]
38360       Else
38370         .PropOnHand_End_Inc = 0@
38380         .PropOnHand_End_Prin = 0@
38390         .PropOnHand_End_Cost = 0@
38400       End If
38410     End If
38420     rst.Close
38430     Set qdf = dbs.QueryDefs("qryCourtReport_FL_00_48")  ' ** Archive, by accountno.
38440     Set rst = qdf.OpenRecordset
38450     If rst.BOF = True And rst.EOF = True Then
            ' ** No transactions listed in LedgerArchive.
38460     Else
38470       rst.MoveFirst
38480       If IsNull(rst![Sumshareface_dp]) = False Then
38490         .PropOnHand_End_Inc = (.PropOnHand_End_Inc + rst![SumIcash])
38500         .PropOnHand_End_Prin = (.PropOnHand_End_Prin + rst![SumPcash])
38510         .PropOnHand_End_Cost = (.PropOnHand_End_Cost + rst![SumCost])
38520       End If
38530     End If
38540     rst.Close
38550     dbs.Close
38560   End With

EXITP:
38570   Set rst = Nothing
38580   Set qdf = Nothing
38590   Set dbs = Nothing
38600   Exit Sub

ERRH:
1410    THAT_PROC = THIS_PROC
1420    That_Erl = Erl
1430    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1440    Resume EXITP

End Sub

Public Sub SetUserReportPath()

38700 On Error GoTo ERRH

        Const THIS_PROC As String = "SetUserReportPath"

        Dim blnEnable As Boolean

38710   With Me
38720     blnEnable = True
38730     Select Case IsNull(.UserReportPath)
          Case True
38740       blnEnable = False
38750     Case False
38760       If Trim(.UserReportPath) = vbNullString Then
38770         blnEnable = False
38780       End If
38790     End Select
38800     Select Case blnEnable
          Case True
38810       .UserReportPath.BorderColor = CLR_LTBLU2
38820       .UserReportPath.BackStyle = acBackStyleNormal
38830       .UserReportPath.Enabled = True  ' ** It remains locked.
38840       .UserReportPath_chk.Enabled = True
38850       .UserReportPath_chk.Locked = False
38860       .UserReportPath_chk_lbl1.Visible = True
38870       .UserReportPath_chk_lbl1_dim.Visible = False
38880       .UserReportPath_chk_lbl1_dim_hi.Visible = False
38890       .UserReportPath_chk_lbl2.Visible = True
38900       .UserReportPath_chk_lbl2_dim.Visible = False
38910       .UserReportPath_chk_lbl2_dim_hi.Visible = False
38920     Case False
38930       .UserReportPath = vbNullString
38940       .UserReportPath.BorderColor = WIN_CLR_DISR
38950       .UserReportPath.BackStyle = acBackStyleTransparent
38960       .UserReportPath.Enabled = False
38970       .UserReportPath_chk.Enabled = False
38980       .UserReportPath_chk.Locked = False
38990       .UserReportPath_chk_lbl1.Visible = False
39000       .UserReportPath_chk_lbl1_dim.Visible = True
39010       .UserReportPath_chk_lbl1_dim_hi.Visible = True
39020       .UserReportPath_chk_lbl2.Visible = False
39030       .UserReportPath_chk_lbl2_dim.Visible = True
39040       .UserReportPath_chk_lbl2_dim_hi.Visible = True
39050     End Select
39060   End With

EXITP:
39070   Exit Sub

ERRH:
1410    DoCmd.Hourglass False
1420    Select Case ERR.Number
        Case Else
1430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1440    End Select
1450    Resume EXITP

End Sub

Public Sub AllCancelSet1_FL(blnCancel As Boolean)

39100 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet1_FL"

39110   blnAllCancel = blnCancel

EXITP:
39120   Exit Sub

ERRH:
1410    Select Case ERR.Number
        Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub
