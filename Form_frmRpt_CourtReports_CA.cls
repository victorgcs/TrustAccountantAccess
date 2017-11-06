VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_CA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_CA"

'VGC 10/09/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  39320
' #########################

' ** For GTR, sets options, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Excel All:        E {cmdExcelAll}
' **   From:             F {DateStart}
' **   Legal Name:       G {chkLegalName}
' **   Include Archive:  H {chkIncludeArchive}
' **   Include CheckNum: K {chkIncludeCheckNum}
' **   Name:             M {opgAccountSource_optName}
' **   Footnote:         N {chkIncludeFootnote}
' **   Print All:        P {cmdPrintAll}
' **   Incl w/Summary:   S {chkAssetList}
' **   To:               T {DateEnd}
' **   Group Rev/Exp:    U {chkGroupBy_IncExpCode}
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

'gdblCrtRpt_CA_POHBeg As Double       '= Nz(Forms("frmRpt_CourtReports_CA").PropOnHand_Beg, 0)
'gdblCrtRpt_CA_POHEnd As Double       '= Nz(Forms("frmRpt_CourtReports_CA").PropOnHand_End, 0)
'gdblCrtRpt_CA_COHBeg As Double        '= Nz(Forms("frmRpt_CourtReports_CA").CashAssets_Beg, 0)
'gdblCrtRpt_CA_COHEnd As Double        '= Nz(Forms("frmRpt_CourtReports_CA").CashAssets_End, 0)
'gdblCrtRpt_CA_InvestInfo As Double    '= Nz(Forms("frmRpt_CourtReports_CA").InvestInfo, 0)
'gdblCrtRpt_CA_InvestChange As Double  '= Nz(Forms("frmRpt_CourtReports_CA").InvestChange, 0)

' ** VGC 02/06/2010: Added query-switching to Cap Gains/Losses reports; didn't bother combining code.

' ** See TAReports_..() functions in modReportFunctions for TAReports CommandBar control functions.
' ** TAReports_SetZero(), used below, sets which Print buttons will be visible on the TAReports CommandBar.
' ** One button, Button 2 (&PrintZero), calls mcrPrint_CA_0, which in turn calls TAReports_Print1().

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

' ** cmbAccounts combo box constants:
Private Const CBX_A_ACTNO  As Integer = 0  ' ** accountno
'Private Const CBX_A_DESC   As Integer = 1  ' ** Desc
'Private Const CBX_A_PREDAT As Integer = 2  ' ** predate
Private Const CBX_A_SHORT  As Integer = 3  ' ** shortname
'Private Const CBX_A_LEGAL  As Integer = 4  ' ** legalname  {Constant is used in reports!}
'Private Const CBX_A_BALDAT As Integer = 5  ' ** BalanceDate (earliest [balance date])
'Private Const CBX_A_HASREL As Integer = 6  ' ** HasRelated
'Private Const CBX_A_CASNUM As Integer = 7  ' ** CaseNum
'Private Const CBX_A_TRXDAT As Integer = 8  ' ** TransDate (earliest [transdate])

' ** Array: arr_varCap().
Private lngCaps As Long, arr_varCap As Variant
'Private Const C_RID  As Integer = 0
Private Const C_RNAM As Integer = 1
'Private Const C_CAP  As Integer = 2
Private Const C_CAPN As Integer = 3

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

Private strLastArrow As String, strFootnoteDefault As String
Private blnRebuildTable As Boolean, blnGoingToReport As Boolean, blnWasGTR As Boolean
Private blnAllCancel As Boolean, blnAccountChecked As Boolean
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim ctl As Access.Control
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
160           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False

170           blnWasGTR = gblnGoToReport

180           DoEvents
190           SetArchiveOption_CA True, Me  ' ** Module Procedure: modCourtReportsCA.
200           DoEvents

              'lngTpp = GetTPP  ' ** Module Procedure: modWindowFunctions.
210           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

220           CLR_DISABLED_FG = CLR_DKGRY
230           CLR_DISABLED_BG = CLR_LTTEAL

240           blnPrintAll_Focus = False: blnWordAll_Focus = False: blnExcelAll_Focus = False

250           strLastArrow = vbNullString
260           gblnLegalName = False
270           gblnUseReveuneExpenseCodes = False: blnWasGTR = False
280           gvarCrtRpt_FL_SpecData = Null
290           strFootnoteDefault = .CourtReports_Footnote.DefaultValue
300           If Right(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Left(strFootnoteDefault, (Len(strFootnoteDefault) - 1))
310           If Left(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Mid(strFootnoteDefault, 2)
320           DoEvents

              ' ** Check some Public variables in case there's been an error.
330           blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
340           blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
350           DoEvents

360           gstrFormQuerySpec = THIS_NAME
370           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
380           gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
390           gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString

400           blnAllCancel = False
410           AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
420           blnRebuildTable = True
430           blnAccountChecked = False

440           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
450           DoEvents

              ' ** If revenue / expense tracking is checked on the options screen,
              ' ** let user select reports to be grouped by revenue/expense codes.
460           Select Case gblnRevenueExpenseTracking
              Case True
470             .chkGroupBy_IncExpCode.Enabled = True
480             .chkGroupBy_IncExpCode_lbl3.Visible = False     ' ** Off note.
490             .chkGroupBy_IncExpCode_lbl4_txt.Locked = True   ' ** Asterisk
500             .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
510             .cmdPreview00_lbl4_txt.Locked = True
520             .cmdPreview00_lbl4_cmd.Enabled = True
530             .cmdPreview02_lbl4_txt.Locked = True
540             .cmdPreview02_lbl4_cmd.Enabled = True
550             .cmdPreview04_lbl4_txt.Locked = True
560             .cmdPreview04_lbl4_cmd.Enabled = True
570           Case False
580             .chkGroupBy_IncExpCode.Enabled = False
590             .chkGroupBy_IncExpCode_lbl3.Visible = True      ' ** Off note.
600             .chkGroupBy_IncExpCode_lbl4_txt.Locked = False  ' ** Asterisk
610             .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
620             .cmdPreview00_lbl4_txt.Locked = False  ' ** So they'll look disabled.
630             .cmdPreview00_lbl4_cmd.Enabled = False
640             .cmdPreview02_lbl4_txt.Locked = False
650             .cmdPreview02_lbl4_cmd.Enabled = False
660             .cmdPreview04_lbl4_txt.Locked = False
670             .cmdPreview04_lbl4_cmd.Enabled = False
680           End Select

690           If gdatStartDate > 0 And gdatEndDate > 0 Then
700             .DateStart = gdatStartDate
710             .DateEnd = gdatEndDate
720           End If
730           DoEvents

      #If NoExcel Then
740           .cmdExcel00.Enabled = False
750           .cmdExcel00B.Enabled = False
760           .cmdExcel01.Enabled = False
770           .cmdExcel02.Enabled = False
780           .cmdExcel03.Enabled = False
790           .cmdExcel04.Enabled = False
800           .cmdExcel05.Enabled = False
810           .cmdExcel06.Enabled = False
820           .cmdExcel07.Enabled = False
830           .cmdExcel08.Enabled = False
840           .cmdExcel09.Enabled = False
850           .cmdExcel10.Enabled = False
860           .cmdExcel11.Enabled = False
870           .cmdExcelAll.Enabled = False
880           .chkOpenExcel.Enabled = False
890           .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
900           .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
910           .cmdExcel00.Enabled = True
920           .cmdExcel00B.Enabled = True
930           .cmdExcel01.Enabled = True
940           .cmdExcel02.Enabled = True
950           .cmdExcel03.Enabled = True
960           .cmdExcel04.Enabled = True
970           .cmdExcel05.Enabled = True
980           .cmdExcel06.Enabled = True
990           .cmdExcel07.Enabled = True
1000          .cmdExcel08.Enabled = True
1010          .cmdExcel09.Enabled = True
1020          .cmdExcel10.Enabled = True
1030          .cmdExcel11.Enabled = True
1040          .cmdExcelAll.Enabled = True
1050          .chkOpenExcel.Enabled = True
1060          .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
1070          .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

              ' ** EVENT CHECK: chkRememberMe!
1080          cmbAccounts_AfterUpdate  ' ** Procedure: Below.
              ' ** EVENT CHECK: chkRememberDates!
1090          If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
1100            DateStart_Exit Cancel  ' ** Procedure: Below.
1110            DateEnd_Exit Cancel  ' ** Procedure: Below.
1120          End If

1130          UserReportPath_chk_AfterUpdate  ' ** Procedure: Below.
1140          SetUserReportPath_CA Me  ' ** Module Procedure: modCourtReportsCA.
1150          DoEvents
1160          chkAssetList_AfterUpdate  ' ** Procedure: Below.
1170          chkLegalName_AfterUpdate  ' ** Procedure: Below.
1180          DoEvents
1190          chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Below.
1200          opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1210          DoEvents
1220          chkRememberMe_AfterUpdate  ' ** Procedure: Below.
1230          chkRememberDates_AfterUpdate  ' ** Procedure: Below.
1240          DoEvents
1250          chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
1260          chkIncludeCheckNum_AfterUpdate  ' ** Procedure: Below.
1270          chkPageOf_AfterUpdate  ' ** Procedure: Below.

1280          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
1290          chkOpenExcel_AfterUpdate  ' ** Procedure: Below.

1300          chkIncludeFootnote_AfterUpdate  ' ** Procedure: Below.
1310          If IsNull(.CourtReports_Footnote) = True Then
1320            .CourtReports_Footnote = strFootnoteDefault
1330          End If

1340          DoCmd.Hourglass False

1350          .DateStart.SetFocus

1360        Else
1370          Cancel = -1
1380        End If
1390      Else
1400        Cancel = -1
1410      End If
1420    End With

1430    If Cancel = -1 Then
1440      Beep
1450      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1460    End If

EXITP:
1470    Set ctl = Nothing
1480    Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

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
100     DoCmd.Hourglass False
110     Select Case ERR.Number
        Case Else
120       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
130     End Select
140     Resume EXITP

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
1880                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("00")), 2) & "B"
1890                Case "04"
1900                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("10")), 2)
1910                Case "07"
1920                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("11")), 2)
1930                Case "10"
1940                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("03")), 2)
1950                Case "11"
1960                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("06")), 2)
1970                Case Else
1980                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) - 1)), 2)
1990                End Select
2000              Else
2010                strLastArrow = strCtl
2020                strCtl = Mid(strCtl, 4, 4)
2030                Select Case strCtl
                    Case "Prev"
2040                  strCtl = "cmdPreview09"
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
2240              If strNum <> "09" Then
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
2320                Case "03"
2330                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("10")), 2)
2340                Case "06"
2350                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("11")), 2)
2360                Case "10"
2370                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("04")), 2)
2380                Case "11"
2390                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("07")), 2)
2400                Case Else
2410                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) + 1)), 2)
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
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1530    Resume EXITP

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
3510          DoCmd.Hourglass False  ' ** It should already be off.
3520          blnGoingToReport = False
3530          gblnGoToReport = False
3540          .GoToReport_arw_rpt_img.Visible = False
3550          DoEvents
3560          Select Case garr_varGoToReport(GTR_CTL6)
              Case "cmdPrint00"
3570            .cmdPreview00_lbl4_txt.Visible = True
3580            Select Case .chkGroupBy_IncExpCode
                Case True
3590              .cmdPreview00_lbl5.Visible = True
3600            Case False
3610              .cmdPreview00_lbl5.Visible = False
3620            End Select
3630          Case "cmdPrint00B"
3640            .GoToReport_arw_alist_on_img.Visible = False
3650            .GoToReport_arw_alist_off_img.Visible = False
3660            .chkAssetList.Visible = True
3670            .chkAssetList_lbl2.Visible = True
3680            .chkAssetList_box.Visible = True
3690          Case "cmdPrint02"
3700            .cmdPreview02_lbl4_txt.Visible = True
3710            Select Case .chkGroupBy_IncExpCode
                Case True
3720              .cmdPreview02_lbl5.Visible = True
3730            Case False
3740              .cmdPreview02_lbl5.Visible = False
3750            End Select
3760          Case "cmdPrint04"
3770            .cmdPreview04_lbl4_txt.Visible = True
3780            .cmdPreview04_lbl6_txt.Visible = True
3790            Select Case .chkGroupBy_IncExpCode
                Case True
3800              .cmdPreview04_lbl5.Visible = True
3810            Case False
3820              .cmdPreview04_lbl5.Visible = False
3830            End Select
3840            Select Case .chkIncludeCheckNum
                Case True
3850              .cmdPreview04_lbl7.Visible = True
3860            Case False
3870              .cmdPreview04_lbl7.Visible = False
3880            End Select
3890          Case "cmdPrint06"
3900            .cmdPreview06_lbl6_txt.Visible = True
3910            Select Case .chkIncludeCheckNum
                Case True
3920              .cmdPreview06_lbl7.Visible = True
3930            Case False
3940              .cmdPreview06_lbl7.Visible = False
3950            End Select
3960          End Select
3970          DoEvents
3980        Case False
3990          DoCmd.Hourglass True  ' ** Make sure it's still running.
4000          DoEvents
4010          blnGoingToReport = True
4020          Select Case garr_varGoToReport(GTR_CTL6)
              Case "cmdPrint00"
4030            lngBtn_Top = (.cmdPreview00_lbl.Top + (lngTpp))
4040            lngBtn_Left = (.cmdPreview00_lbl.Left + .cmdPreview00_lbl.Width)
4050            .cmdPreview00_lbl4_txt.Visible = False
4060            .cmdPreview00_lbl5.Visible = False
4070            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4080            DoEvents
4090          Case "cmdPrint00B"
4100            lngBtn_Top = (.cmdPreview00B_lbl.Top - (0& * lngTpp))
4110            lngBtn_Left = ((.cmdPreview00B_lbl.Left + .cmdPreview00B_lbl.Width) - (lngTpp))
4120            .chkAssetList.Visible = False
4130            .chkAssetList_lbl2.Visible = False
4140            .chkAssetList_box.Visible = False
4150            .GoToReport_arw_alist_on_img.Top = lngBtn_Top
4160            .GoToReport_arw_alist_off_img.Top = lngBtn_Top
4170            .GoToReport_arw_alist_on_img.Left = lngBtn_Left
4180            .GoToReport_arw_alist_off_img.Left = lngBtn_Left
4190            Select Case .chkAssetList
                Case True
4200              .GoToReport_arw_alist_on_img.Visible = True
4210            Case False
4220              .GoToReport_arw_alist_off_img.Visible = True
4230            End Select
4240            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4250            DoEvents
4260          Case "cmdPrint01"
4270            lngBtn_Top = (.cmdPreview01_lbl.Top + (lngTpp))
4280            lngBtn_Left = (.cmdPreview01_lbl.Left + .cmdPreview01_lbl.Width)
4290            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4300            DoEvents
4310          Case "cmdPrint02"
4320            lngBtn_Top = (.cmdPreview02_lbl.Top + (lngTpp))
4330            lngBtn_Left = (.cmdPreview02_lbl.Left + .cmdPreview02_lbl.Width)
4340            .cmdPreview02_lbl4_txt.Visible = False
4350            .cmdPreview02_lbl5.Visible = False
4360            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4370            DoEvents
4380          Case "cmdPrint03"
4390            lngBtn_Top = (.cmdPreview03_lbl.Top + (lngTpp))
4400            lngBtn_Left = (.cmdPreview03_lbl.Left + .cmdPreview03_lbl.Width)
4410            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4420            DoEvents
4430          Case "cmdPrint04"
4440            lngBtn_Top = (.cmdPreview04_lbl.Top + (lngTpp))
4450            lngBtn_Left = (.cmdPreview04_lbl.Left + .cmdPreview04_lbl.Width)
4460            .cmdPreview04_lbl4_txt.Visible = False
4470            .cmdPreview04_lbl5.Visible = False
4480            .cmdPreview04_lbl6_txt.Visible = False
4490            .cmdPreview04_lbl7.Visible = False
4500            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4510            DoEvents
4520          Case "cmdPrint05"
4530            lngBtn_Top = (.cmdPreview05_lbl.Top + (lngTpp))
4540            lngBtn_Left = (.cmdPreview05_lbl.Left + .cmdPreview05_lbl.Width)
4550            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4560            DoEvents
4570          Case "cmdPrint06"
4580            lngBtn_Top = (.cmdPreview06_lbl.Top + (lngTpp))
4590            lngBtn_Left = (.cmdPreview06_lbl.Left + .cmdPreview06_lbl.Width)
4600            .cmdPreview06_lbl6_txt.Visible = False
4610            .cmdPreview06_lbl7.Visible = False
4620            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4630            DoEvents
4640          Case "cmdPrint07"
4650            lngBtn_Top = (.cmdPreview07_lbl.Top + (lngTpp))
4660            lngBtn_Left = (.cmdPreview07_lbl.Left + .cmdPreview07_lbl.Width)
4670            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4680            DoEvents
4690          Case "cmdPrint08"
4700            lngBtn_Top = (.cmdPreview08_lbl.Top + (lngTpp))
4710            lngBtn_Left = (.cmdPreview08_lbl.Left + .cmdPreview08_lbl.Width)
4720            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4730            DoEvents
4740          Case "cmdPrint09"
4750            lngBtn_Top = (.cmdPreview09_lbl.Top + (lngTpp))
4760            lngBtn_Left = (.cmdPreview09_lbl.Left + .cmdPreview09_lbl.Width)
4770            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4780            DoEvents
4790          Case "cmdPrint10"
4800            lngBtn_Top = (.cmdPreview10_lbl.Top + (lngTpp))
4810            lngBtn_Left = (.cmdPreview10_lbl.Left + .cmdPreview10_lbl.Width)
4820            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4830            DoEvents
4840          Case "cmdPrint11"
4850            lngBtn_Top = (.cmdPreview11_lbl.Top + (lngTpp))
4860            lngBtn_Left = (.cmdPreview11_lbl.Left + .cmdPreview11_lbl.Width)
4870            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4880            DoEvents
4890          End Select
4900          If garr_varGoToReport(GTR_CTL6) <> "cmdPrint00B" Then
4910            .GoToReport_arw_rpt_img.Top = lngBtn_Top
4920            .GoToReport_arw_rpt_img.Left = lngBtn_Left
4930            .GoToReport_arw_rpt_img.Visible = True
4940          End If
4950          DoEvents
4960          Beep
4970          DoCmd.Hourglass False
4980          DoEvents
4990          .TimerInterval = (GTR_WAIT * 2&)
5000        End Select
5010      End If
5020      If gblnTimer = True Then
5030        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
5040          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
5050          gblnTimer = False
5060          DoCmd.Hourglass False
5070        End If
5080      End If
5090    End With

'HERE! 2  Form_Load()  frmRpt_CourtReports_CA
'HERE! 1  Form_Open()  frmRpt_CourtReports_CA
'HERE! 3  Form_Timer()  frmRpt_CourtReports_CA

EXITP:
5100    Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

5210    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
5220      Response = acDataErrContinue
5230    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
5240      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
5250        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
5260        Response = acDataErrContinue
5270      End If
5280    Case 2237  ' ** The text you entered isn't an item in the list.
5290      Response = acDataErrDisplay
5300    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
5310      Beep
5320      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
5330      Response = acDataErrContinue
5340    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
5350    Case Else
5360      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
5370      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
5380    End Select

EXITP:
5390    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String

5410    With Me
5420      If .chkRememberMe = False Then
5430        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
5440      End If
5450      If .chkRememberDates = False Then
5460        .DateStart = Null
5470        .DateEnd = Null
5480      End If
5490    End With

5500    FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5510    DoEvents

5520    Select Case blnWasGTR
        Case True
5530      blnWasGTR = False
5540    Case False
          ' ** Don't save settings imposed by GoToReport!
5550      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
5560    End Select

5570    gblnLegalName = False

5580      ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
5590      ChkSpecLedgerEntryNull  ' ** Module Function: modUtilities.

        ' ** Make sure all reports are closed.
5600    Do While Reports.Count > 0
5610      DoCmd.Close acReport, Reports(0).Name
5620    Loop

5630    strDocName = "frmRpt_CourtReports_CA_Input"
5640    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
5650      DoCmd.Close acForm, strDocName
5660    End If

5670    gstrFormQuerySpec = vbNullString
5680    gblnSetFocus = True
5690    gstrReturningForm = THIS_NAME
5700    Select Case strCallingForm
        Case "frmMenu_CourtReport"
5710      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Report"
5720    End Select

EXITP:
5730    Set qdf = Nothing
5740    Set dbs = Nothing
5750    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

5810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5820    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

5910    Detail_Mouse_CA blnCalendar1_Focus, blnCalendar2_Focus, _
          blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
5920    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

6010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6020    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdClose_Click()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

6110    DoCmd.Close acForm, THIS_NAME

EXITP:
6120    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab
6270        With Me
6280          intRetVal = 0
6290          .DateStart.SetFocus
6300        End With
6310      Case vbKeyLeft
6320        With Me
6330          intRetVal = 0
6340          If .cmdExcelAll.Enabled = True Then
6350            .cmdExcelAll.SetFocus
6360          Else
6370            .cmdWordAll.SetFocus
6380          End If
6390        End With
6400      Case vbKeyRight
6410        With Me
6420          intRetVal = 0
6430          .DateStart.SetFocus
6440        End With
6450      End Select
6460    End If

        ' ** Shift keys.
6470    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6480      Select Case intRetVal
          Case vbKeyTab
6490        With Me
6500          intRetVal = 0
6510          If .cmdExcelAll.Enabled = True Then
6520            .cmdExcelAll.SetFocus
6530          Else
6540            .cmdWordAll.SetFocus
6550          End If
6560        End With
6570      End Select
6580    End If

        ' ** Ctrl-Shift keys.
6590    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6600      Select Case intRetVal
          Case vbKeyTab
6610        With Me
6620          intRetVal = 0
6630          If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
6640            .UserReportPath_chk.SetFocus
6650          End If
6660        End With
6670      End Select
6680    End If

EXITP:
6690    KeyCode = intRetVal
6700    Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

6810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
6820    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6920    intShiftDown = (Shift And acShiftMask) > 0
6930    intAltDown = (Shift And acAltMask) > 0
6940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6970        With Me
6980          intRetVal = 0
6990          .DateEnd.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7050        With Me
7060          intRetVal = 0
7070          .cmdClose.SetFocus
7080        End With
7090      End Select
7100    End If

        ' ** Ctrl keys.
7110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7120      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7130        With Me
7140          intRetVal = 0
7150          .cmdCalendar1.SetFocus
7160        End With
7170      End Select
7180    End If

EXITP:
7190    KeyCode = intRetVal
7200    Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

7310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
7320    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub DateStart_Exit(Cancel As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

        Dim strTmp01 As String

7410    With Me
7420      blnRebuildTable = True
7430      If IsNull(.DateStart) = False Then
7440        gdatStartDate = .DateStart
7450        strTmp01 = CStr(CDbl(gdatStartDate))
7460        If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7470        glngStartDateLong = CLng(strTmp01)
7480      End If
7490    End With

EXITP:
7500    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case 94  ' ** Invalid use of Null.
          ' ** Ignore.
1500    Case Else
1510      THAT_PROC = THIS_PROC
1520      That_Erl = Erl
1530      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1540    End Select
1550    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

7610    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
7620    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

7710    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
7720    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

7810    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
7820    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

7910    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
7920    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

8010    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
8020    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

8110    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
8120    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8220    intShiftDown = (Shift And acShiftMask) > 0
8230    intAltDown = (Shift And acAltMask) > 0
8240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8270        With Me
8280          intRetVal = 0
8290          .cmbAccounts.SetFocus
8300        End With
8310      End Select
8320    End If

        ' ** Shift keys.
8330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8350        With Me
8360          intRetVal = 0
8370          .DateStart.SetFocus
8380        End With
8390      End Select
8400    End If

        ' ** Ctrl keys.
8410    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8420      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8430        With Me
8440          intRetVal = 0
8450          .cmdCalendar2.SetFocus
8460        End With
8470      End Select
8480    End If

        ' ** Ctrl-Shift keys.
8490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8500      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8510        With Me
8520          intRetVal = 0
8530          .cmdCalendar1.SetFocus
8540        End With
8550      End Select
8560    End If

EXITP:
8570    KeyCode = intRetVal
8580    Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

8610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
8620    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub DateEnd_Exit(Cancel As Integer)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

        Dim strTmp01 As String

8710    With Me
8720      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
8730        .DateEnd = Date
8740      Else
8750        If .DateEnd = vbNullString Then
8760          .DateEnd = Date
8770        End If
8780      End If
8790      blnRebuildTable = True
8800      gdatEndDate = .DateEnd
8810      strTmp01 = CStr(CDbl(gdatEndDate))
8820      If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
8830      glngEndDateLong = CLng(strTmp01)
8840    End With

EXITP:
8850    Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

8910    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
8920    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

9010    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
9020    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

9110    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
9120    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9220    intShiftDown = (Shift And acShiftMask) > 0
9230    intAltDown = (Shift And acAltMask) > 0
9240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9270        With Me
9280          intRetVal = 0
9290          .chkRememberDates.SetFocus
9300        End With
9310      End Select
9320    End If

        ' ** Shift keys.
9330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9350        With Me
9360          intRetVal = 0
9370          .DateEnd.SetFocus
9380        End With
9390      End Select
9400    End If

EXITP:
9410    KeyCode = intRetVal
9420    Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

9510    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
9520    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

9610    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
9620    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

9710    Calendar_Handler_CA THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
9720    Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

9810    With Me
9820      Select Case .chkRememberDates
          Case True
9830        .chkRememberDates_lbl.FontBold = True
9840      Case False
9850        .chkRememberDates_lbl.FontBold = False
9860      End Select
9870    End With

EXITP:
9880    Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9920    intShiftDown = (Shift And acShiftMask) > 0
9930    intAltDown = (Shift And acAltMask) > 0
9940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9970        With Me
9980          intRetVal = 0
9990          .opgAccountSource.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10050       With Me
10060         intRetVal = 0
10070         .cmdCalendar2.SetFocus
10080       End With
10090     End Select
10100   End If

EXITP:
10110   KeyCode = intRetVal
10120   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

10210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
10220   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

10310   strAccountNo = vbNullString

10320   With Me
10330     If IsNull(.cmbAccounts) = False Then
10340       If Len(.cmbAccounts.Column(CBX_A_ACTNO)) > 0 Then
10350         strAccountNo = .cmbAccounts.Column(CBX_A_ACTNO)
10360       End If
10370     End If
10380     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
10390       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
10400       .opgAccountSource_optNumber_lbl.FontBold = True
10410       .opgAccountSource_optName_lbl.FontBold = False
10420     Case .opgAccountSource_optName.OptionValue
10430       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
10440       .opgAccountSource_optNumber_lbl.FontBold = False
10450       .opgAccountSource_optName_lbl.FontBold = True
10460     End Select
10470     DoEvents
10480     If strAccountNo <> vbNullString Then
10490       .cmbAccounts = strAccountNo
10500     End If
10510   End With

EXITP:
10520   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10620   intShiftDown = (Shift And acShiftMask) > 0
10630   intAltDown = (Shift And acAltMask) > 0
10640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10670       With Me
10680         intRetVal = 0
10690         .chkRememberMe.SetFocus
10700       End With
10710     End Select
10720   End If

        ' ** Shift keys.
10730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10750       With Me
10760         intRetVal = 0
10770         .chkRememberDates.SetFocus
10780       End With
10790     End Select
10800   End If

EXITP:
10810   KeyCode = intRetVal
10820   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10920   intShiftDown = (Shift And acShiftMask) > 0
10930   intAltDown = (Shift And acAltMask) > 0
10940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10970       With Me
10980         intRetVal = 0
10990         .chkRememberMe.SetFocus
11000       End With
11010     End Select
11020   End If

        ' ** Shift keys.
11030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11050       With Me
11060         intRetVal = 0
11070         .chkRememberDates.SetFocus
11080       End With
11090     End Select
11100   End If

EXITP:
11110   KeyCode = intRetVal
11120   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

11210   With Me
11220     Select Case .chkRememberMe
          Case True
11230       .chkRememberMe_lbl.FontBold = True
11240     Case False
11250       .chkRememberMe_lbl.FontBold = False
11260     End Select
11270   End With

EXITP:
11280   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11320   intShiftDown = (Shift And acShiftMask) > 0
11330   intAltDown = (Shift And acAltMask) > 0
11340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11370       With Me
11380         intRetVal = 0
11390         .cmbAccounts.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11450       With Me
11460         intRetVal = 0
11470         .opgAccountSource.SetFocus
11480       End With
11490     End Select
11500   End If

EXITP:
11510   KeyCode = intRetVal
11520   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

11610   With Me
11620     If .cmbAccounts.text = vbNullString Then
11630       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
11640       .cmbAccounts = .cmbAccounts.ItemData(0)
11650       blnRebuildTable = True
11660     End If
11670   End With

EXITP:
11680   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

11710   With Me
11720     blnRebuildTable = True
11730     If IsNull(.cmbAccounts) = False Then
11740       If .cmbAccounts <> vbNullString Then
11750         gstrCrtRpt_Account = Me.cmbAccounts
11760         gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
11770         gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
11780         gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString
11790         gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False
11800         If .HasArchive = True And .chkIncludeArchive = True Then
11810           SetArchiveOption_CA False, Me  ' ** Module Procedure: modCourtReportsCA.
11820         End If
11830       End If
11840     End If
11850   End With

EXITP:
11860   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11920   intShiftDown = (Shift And acShiftMask) > 0
11930   intAltDown = (Shift And acAltMask) > 0
11940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11970       With Me
11980         intRetVal = 0
11990         .cmdPreview00.SetFocus
12000       End With
12010     End Select
12020   End If

        ' ** Shift keys.
12030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12050       With Me
12060         intRetVal = 0
12070         .DateEnd.SetFocus
12080       End With
12090     End Select
12100   End If

        ' ** Ctrl keys.
12110   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
12120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12130       With Me
12140         intRetVal = 0
12150         If .chkGroupBy_IncExpCode.Enabled = True Then
12160           .chkGroupBy_IncExpCode.SetFocus
12170         ElseIf .chkIncludeArchive.Enabled = True Then
12180           .chkIncludeArchive.SetFocus
12190         Else
12200           .chkLegalName.SetFocus
12210         End If
12220       End With
12230     End Select
12240   End If

        ' ** Ctrl-Shift keys.
12250   If intCtrlDown And (Not intAltDown) And intShiftDown Then
12260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12270       With Me
12280         intRetVal = 0
12290         .chkRememberMe.SetFocus
12300       End With
12310     End Select
12320   End If

EXITP:
12330   KeyCode = intRetVal
12340   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Public Sub chkGroupBy_IncExpCode_AfterUpdate()
' ** Public for GoToReport.
' ** Available on reports 0, 2, and 4.

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_AfterUpdate"

12410   With Me
12420     Select Case .chkGroupBy_IncExpCode
          Case True
12430       gblnUseReveuneExpenseCodes = True
12440       .chkGroupBy_IncExpCode_lbl.FontBold = True
12450       .chkGroupBy_IncExpCode_lbl4.FontSize = 20
12460       .chkGroupBy_IncExpCode_lbl5.Visible = True
12470       .cmdPreview00_lbl4.FontSize = 20
12480       .cmdPreview00_lbl5.Visible = True
12490       .cmdPreview02_lbl4.FontSize = 20
12500       .cmdPreview02_lbl5.Visible = True
12510       .cmdPreview04_lbl4.FontSize = 20
12520       .cmdPreview04_lbl5.Visible = True
12530     Case False
12540       gblnUseReveuneExpenseCodes = False
12550       .chkGroupBy_IncExpCode_lbl.FontBold = False
12560       .chkGroupBy_IncExpCode_lbl4.FontSize = 14
12570       .chkGroupBy_IncExpCode_lbl5.Visible = False
12580       .cmdPreview00_lbl4.FontSize = 14
12590       .cmdPreview00_lbl5.Visible = False
12600       .cmdPreview02_lbl4.FontSize = 14
12610       .cmdPreview02_lbl5.Visible = False
12620       .cmdPreview04_lbl4.FontSize = 14
12630       .cmdPreview04_lbl5.Visible = False
12640     End Select
12650   End With

EXITP:
12660   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkGroupBy_IncExpCode_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_KeyDown"

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
12790         If .chkIncludeArchive.Enabled = True Then
12800           .chkIncludeArchive.SetFocus
12810         Else
12820           .chkLegalName.SetFocus
12830         End If
12840       End With
12850     End Select
12860   End If

        ' ** Shift keys.
12870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12880     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12890       With Me
12900         intRetVal = 0
12910         .cmbAccounts.SetFocus
12920       End With
12930     End Select
12940   End If

EXITP:
12950   KeyCode = intRetVal
12960   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

13010   With Me
13020     Select Case .chkIncludeArchive
          Case True
13030       .chkIncludeArchive_lbl.FontBold = True
13040       If IsNull(.cmbAccounts) = False Then
13050         SetArchiveOption_CA False, Me  ' ** Module Procedure: modCourtReportsCA.
13060       End If
13070     Case False
13080       .chkIncludeArchive_lbl.FontBold = False
13090     End Select
13100   End With

EXITP:
13110   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

13200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13220   intShiftDown = (Shift And acShiftMask) > 0
13230   intAltDown = (Shift And acAltMask) > 0
13240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13270       With Me
13280         intRetVal = 0
13290         .chkLegalName.SetFocus
13300       End With
13310     End Select
13320   End If

        ' ** Shift keys.
13330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13350       With Me
13360         intRetVal = 0
13370         If .chkGroupBy_IncExpCode.Enabled = True Then
13380           .chkGroupBy_IncExpCode.SetFocus
13390         Else
13400           .cmbAccounts.SetFocus
13410         End If
13420       End With
13430     End Select
13440   End If

EXITP:
13450   KeyCode = intRetVal
13460   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkLegalName_AfterUpdate()

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_AfterUpdate"

13510   With Me
13520     Select Case .chkLegalName
          Case True
13530       .chkLegalName_lbl.FontBold = True
13540       gblnLegalName = True
13550     Case False
13560       .chkLegalName_lbl.FontBold = False
13570       gblnLegalName = False
13580     End Select
13590   End With

EXITP:
13600   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkLegalName_KeyDown(KeyCode As Integer, Shift As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13720   intShiftDown = (Shift And acShiftMask) > 0
13730   intAltDown = (Shift And acAltMask) > 0
13740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13770       With Me
13780         intRetVal = 0
13790         .chkIncludeFootnote.SetFocus
13800       End With
13810     End Select
13820   End If

        ' ** Shift keys.
13830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13850       With Me
13860         intRetVal = 0
13870         If .chkIncludeArchive.Enabled = True Then
13880           .chkIncludeArchive.SetFocus
13890         ElseIf .chkGroupBy_IncExpCode.Enabled = True Then
13900           .chkGroupBy_IncExpCode.SetFocus
13910         Else
13920           .cmbAccounts.SetFocus
13930         End If
13940       End With
13950     End Select
13960   End If

EXITP:
13970   KeyCode = intRetVal
13980   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkIncludeFootnote_AfterUpdate()

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_AfterUpdate"

14010   With Me
14020     .CourtReports_Footnote.Visible = True
14030     .CourtReports_Footnote_box.Visible = True
14040     .CourtReports_Footnote_lbl2.Visible = False
14050     Select Case .chkIncludeFootnote
          Case True
14060       .chkIncludeFootnote_lbl.FontBold = True
14070       .CourtReports_Footnote.Enabled = True
14080       .CourtReports_Footnote.BackStyle = acBackStyleNormal
14090       .CourtReports_Footnote_lbl.BackStyle = acBackStyleNormal
14100       .CourtReports_Footnote.BorderColor = CLR_LTBLU2
14110       If IsNull(.CourtReports_Footnote) = True Then
14120         .CourtReports_Footnote = strFootnoteDefault
14130       End If
14140       If gblnRevenueExpenseTracking = False Then
14150         .chkGroupBy_IncExpCode_lbl3.Visible = False
14160       End If
14170     Case False
14180       .chkIncludeFootnote_lbl.FontBold = False
14190       .CourtReports_Footnote.Enabled = False
14200       .CourtReports_Footnote.BackStyle = acBackStyleTransparent
14210       .CourtReports_Footnote_lbl.BackStyle = acBackStyleTransparent
14220       .CourtReports_Footnote.BorderColor = WIN_CLR_DISR
14230       If gblnRevenueExpenseTracking = False Then
14240         .chkGroupBy_IncExpCode_lbl3.Visible = True
14250       End If
14260     End Select
14270   End With

EXITP:
14280   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkIncludeFootnote_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14320   intShiftDown = (Shift And acShiftMask) > 0
14330   intAltDown = (Shift And acAltMask) > 0
14340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14370       With Me
14380         intRetVal = 0
14390         Select Case .chkIncludeFootnote
              Case True
14400           .CourtReports_Footnote.SetFocus
14410         Case False
14420           .chkIncludeCheckNum.SetFocus
14430         End Select
14440       End With
14450     End Select
14460   End If

        ' ** Shift keys.
14470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14490       With Me
14500         intRetVal = 0
14510         .chkLegalName.SetFocus
14520       End With
14530     End Select
14540   End If

EXITP:
14550   KeyCode = intRetVal
14560   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub CourtReports_Footnote_AfterUpdate()

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_AfterUpdate"

        Dim varTmp00 As Variant

14610   With Me
14620     varTmp00 = .CourtReports_Footnote
14630     If IsNull(varTmp00) = False Then
14640       If Trim(varTmp00) <> vbNullString Then
14650         varTmp00 = Trim(varTmp00)
14660         If Len(varTmp00) > MEMO_MAX Then
14670           Beep
14680           If .CourtReports_Footnote_lbl2.Visible = False Then
14690             .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
14700             .CourtReports_Footnote_lbl2.Visible = True
14710           End If
14720           varTmp00 = Left(varTmp00, MEMO_MAX)
14730           .CourtReports_Footnote = varTmp00
14740         End If
14750       Else
14760         .CourtReports_Footnote = Null
14770       End If
14780     End If
14790   End With

EXITP:
14800   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub CourtReports_Footnote_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intLen As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14970       With Me
14980         intRetVal = 0
14990         .chkIncludeCheckNum.SetFocus
15000       End With
15010     Case Else
15020       With Me
15030         If IsNull(.CourtReports_Footnote.text) = False Then
15040           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
15050             intLen = Len(Trim(.CourtReports_Footnote.text))
15060             If intLen > MEMO_MAX Then
15070               intRetVal = 0
15080               Beep
15090               If .CourtReports_Footnote_lbl2.Visible = False Then
15100                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
15110                 .CourtReports_Footnote_lbl2.Visible = True
15120               End If
15130               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
15140               .CourtReports_Footnote.SelStart = 999
15150             End If
15160           End If
15170         End If
15180       End With
15190     End Select
15200   End If

        ' ** Shift keys.
15210   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15220     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15230       With Me
15240         intRetVal = 0
15250         .chkIncludeFootnote.SetFocus
15260       End With
15270     Case Else
15280       With Me
15290         If IsNull(.CourtReports_Footnote.text) = False Then
15300           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
15310             intLen = Len(Trim(.CourtReports_Footnote.text))
15320             If intLen > MEMO_MAX Then
15330               intRetVal = 0
15340               Beep
15350               If .CourtReports_Footnote_lbl2.Visible = False Then
15360                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
15370                 .CourtReports_Footnote_lbl2.Visible = True
15380               End If
15390               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
15400               .CourtReports_Footnote.SelStart = 999
15410             End If
15420           End If
15430         End If
15440       End With
15450     End Select
15460   End If

EXITP:
15470   KeyCode = intRetVal
15480   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkIncludeCheckNum_AfterUpdate()

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCheckNum_AfterUpdate"

15510   With Me
15520     Select Case .chkIncludeCheckNum
          Case True
15530       .chkIncludeCheckNum_lbl.FontBold = True
15540       .chkIncludeCheckNum_lbl4.FontSize = 20
15550       .chkIncludeCheckNum_lbl5.Visible = True
15560       .cmdPreview04_lbl6.FontSize = 20
15570       .cmdPreview04_lbl7.Visible = True
15580       .cmdPreview06_lbl6.FontSize = 20
15590       .cmdPreview06_lbl7.Visible = True
15600     Case False
15610       .chkIncludeCheckNum_lbl.FontBold = False
15620       .chkIncludeCheckNum_lbl4.FontSize = 14
15630       .chkIncludeCheckNum_lbl5.Visible = False
15640       .cmdPreview04_lbl6.FontSize = 14
15650       .cmdPreview04_lbl7.Visible = False
15660       .cmdPreview06_lbl6.FontSize = 14
15670       .cmdPreview06_lbl7.Visible = False
15680     End Select
15690   End With

EXITP:
15700   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkIncludeCheckNum_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCheckNum_KeyDown"

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
15890         .chkPageOf.SetFocus
15900       End With
15910     End Select
15920   End If

        ' ** Shift keys.
15930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15950       With Me
15960         intRetVal = 0
15970         .chkIncludeFootnote.SetFocus
15980       End With
15990     End Select
16000   End If

EXITP:
16010   KeyCode = intRetVal
16020   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

16110   With Me
16120     Select Case .chkPageOf
          Case True
16130       .chkPageOf_lbl.FontBold = True
16140     Case False
16150       .chkPageOf_lbl.FontBold = False
16160     End Select
16170   End With

EXITP:
16180   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

16200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16220   intShiftDown = (Shift And acShiftMask) > 0
16230   intAltDown = (Shift And acAltMask) > 0
16240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16270       With Me
16280         intRetVal = 0
16290         .cmdPreview00.SetFocus
16300       End With
16310     End Select
16320   End If

        ' ** Shift keys.
16330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16350       With Me
16360         intRetVal = 0
16370         .chkIncludeCheckNum.SetFocus
16380       End With
16390     End Select
16400   End If

EXITP:
16410   KeyCode = intRetVal
16420   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdPreview00_lbl4_cmd_DblClick(Cancel As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_lbl4_cmd_DblClick"

16510   With Me
16520     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
16530     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
16540   End With

EXITP:
16550   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview00_Click()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_Click"

16610   Preview00_CA Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
16620   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdPreview00_KeyDown(KeyCode As Integer, Shift As Integer)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16720   intShiftDown = (Shift And acShiftMask) > 0
16730   intAltDown = (Shift And acAltMask) > 0
16740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16760     Select Case intRetVal
          Case vbKeyTab
16770       With Me
16780         intRetVal = 0
16790         .cmdPrint00.SetFocus
16800       End With
16810     End Select
16820   End If

        ' ** Shift keys.
16830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16840     Select Case intRetVal
          Case vbKeyTab
16850       With Me
16860         intRetVal = 0
16870         .cmbAccounts.SetFocus
16880       End With
16890     End Select
16900   End If

        ' ** Ctrl-Shift keys.
16910   If intCtrlDown And (Not intAltDown) And intShiftDown Then
16920     Select Case intRetVal
          Case vbKeyTab
16930       With Me
16940         intRetVal = 0
16950         .chkPageOf.SetFocus
16960       End With
16970     End Select
16980   End If

EXITP:
16990   KeyCode = intRetVal
17000   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Public Sub cmdPrint00_Click()
' ** Public, because it's called from the TAReports CommandBar as well.

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00_Click"

17110   Print00_CA Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
17120   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Public Sub cmdWord00_Click()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_Click"

17210   Word00_CA blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
17220   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdWord00_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17320   intShiftDown = (Shift And acShiftMask) > 0
17330   intAltDown = (Shift And acAltMask) > 0
17340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17360     Select Case intRetVal
          Case vbKeyTab
17370       With Me
17380         intRetVal = 0
17390         If .cmdExcel00.Enabled = True Then
17400           .cmdExcel00.SetFocus
17410         Else
17420           .cmdPreview00B.SetFocus
17430         End If
17440       End With
17450     Case vbKeyRight
17460       With Me
17470         intRetVal = 0
17480         If .cmdExcel00.Enabled = True Then
17490           .cmdExcel00.SetFocus
17500         Else
17510           .cmdPreview00B.SetFocus
17520         End If
17530       End With
17540     End Select
17550   End If

        ' ** Shift keys.
17560   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17570     Select Case intRetVal
          Case vbKeyTab
17580       With Me
17590         intRetVal = 0
17600         .cmdPrint00.SetFocus
17610       End With
17620     End Select
17630   End If

EXITP:
17640   KeyCode = intRetVal
17650   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Public Sub cmdExcel00_Click()

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_Click"

17710   Excel00_CA blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
17720   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdExcel00_KeyDown(KeyCode As Integer, Shift As Integer)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17820   intShiftDown = (Shift And acShiftMask) > 0
17830   intAltDown = (Shift And acAltMask) > 0
17840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17860     Select Case intRetVal
          Case vbKeyTab
17870       With Me
17880         intRetVal = 0
17890         .cmdPreview00B.SetFocus
17900       End With
17910     Case vbKeyRight
17920       With Me
17930         intRetVal = 0
17940         .cmdPreview00B.SetFocus
17950       End With
17960     End Select
17970   End If

        ' ** Shift keys.
17980   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17990     Select Case intRetVal
          Case vbKeyTab
18000       With Me
18010         intRetVal = 0
18020         .cmdWord00.SetFocus
18030       End With
18040     End Select
18050   End If

EXITP:
18060   KeyCode = intRetVal
18070   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Public Sub chkAssetList_AfterUpdate()
' ** Public for GoToReport.

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_AfterUpdate"

18110   With Me
18120     Select Case .chkAssetList
          Case True
18130       .chkAssetList_lbl.FontBold = True
18140       .chkAssetList_lbl2.FontBold = True
18150     Case False
18160       .chkAssetList_lbl.FontBold = False
18170       .chkAssetList_lbl2.FontBold = False
18180     End Select
18190   End With

EXITP:
18200   Exit Sub

ERRH:
1490    THAT_PROC = THIS_PROC
1500    That_Erl = Erl
1510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1520    Resume EXITP

End Sub

Private Sub chkAssetList_KeyDown(KeyCode As Integer, Shift As Integer)

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
18370       With Me
18380         intRetVal = 0
18390         .cmdPreview00B.SetFocus
18400       End With
18410     End Select
18420   End If

        ' ** Shift keys.
18430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18450       With Me
18460         intRetVal = 0
18470         If .cmdExcel00.Enabled = True Then
18480           .cmdExcel00.SetFocus
18490         Else
18500           .cmdWord00.SetFocus
18510         End If
18520       End With
18530     End Select
18540   End If

EXITP:
18550   KeyCode = intRetVal
18560   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdPreview00B_Click()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_Click"

18610   If Validate = True Then  ' ** Function: Below.
18620     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
18630     PreviewOrPrint "0B", THIS_PROC, acViewPreview  ' ** Procedure: Below.
18640   End If  ' ** Validate().

EXITP:
18650   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdPreview00B_KeyDown(KeyCode As Integer, Shift As Integer)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18720   intShiftDown = (Shift And acShiftMask) > 0
18730   intAltDown = (Shift And acAltMask) > 0
18740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18760     Select Case intRetVal
          Case vbKeyTab
18770       With Me
18780         intRetVal = 0
18790         .cmdPrint00B.SetFocus
18800       End With
18810     Case vbKeyLeft
18820       With Me
18830         intRetVal = 0
18840         If .cmdExcel00.Enabled = True Then
18850           .cmdExcel00.SetFocus
18860         Else
18870           .cmdWord00.SetFocus
18880         End If
18890       End With
18900     End Select
18910   End If

        ' ** Shift keys.
18920   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18930     Select Case intRetVal
          Case vbKeyTab
18940       With Me
18950         intRetVal = 0
18960         .cmdWord00.SetFocus
18970       End With
18980     End Select
18990   End If

EXITP:
19000   KeyCode = intRetVal
19010   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdPrint00B_Click()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00B_Click"

19110   If Validate = True Then  ' ** Function: Below.
19120     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_00B
19130     PreviewOrPrint "0B", THIS_PROC, acViewNormal  ' ** Procedure: Below.
19140   End If  ' ** Validate().

EXITP:
19150   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Public Sub cmdWord00B_Click()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00B_Click"

19210   If Validate = True Then  ' ** Function: Below.
19220     SendToFile_CA "0B", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
19230   End If  ' ** Validate().

EXITP:
19240   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Public Sub cmdExcel00B_Click()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00B_Click"

19310   If Validate = True Then  ' ** Function: Below.
19320     SendToFile_CA "0B", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
19330   End If  ' ** Validate().

EXITP:
19340   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdPreview01_Click()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview01_Click"

19410   If Validate = True Then  ' ** Function: Below.
19420     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
19430     PreviewOrPrint "1", THIS_PROC, acViewPreview  ' ** Procedure: Below.
19440   End If  ' ** Validate

EXITP:
19450   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint01_Click()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint01_Click"

19510   If Validate = True Then  ' ** Function: Below.
19520     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_01
19530     PreviewOrPrint "1", THIS_PROC, acViewNormal  ' ** Procedure: Below.
19540   End If  ' ** Validate

EXITP:
19550   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord01_Click()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord01_Click"

19610   If Validate = True Then  ' ** Function: Below.
19620     SendToFile_CA "1", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
19630   End If  ' ** Validate

EXITP:
19640   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel01_Click()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel01_Click"

19710   If Validate = True Then  ' ** Function: Below.
19720     SendToFile_CA "1", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
19730   End If  ' ** Validate

EXITP:
19740   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview02_lbl4_cmd_DblClick(Cancel As Integer)

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview02_lbl4_cmd_DblClick"

19810   With Me
19820     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
19830     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
19840   End With

EXITP:
19850   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview02_Click()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview02_Click"

19910   If Validate = True Then  ' ** Function: Below.
19920     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
19930     If gblnUseReveuneExpenseCodes = True Then
19940       PreviewOrPrint "2A", THIS_PROC, acViewPreview  ' ** Procedure: Below.
19950     Else
19960       PreviewOrPrint "2", THIS_PROC, acViewPreview  ' ** Procedure: Below.
19970     End If
19980   End If  ' ** Validate

EXITP:
19990   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint02_Click()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint02_Click"

20010   If Validate = True Then  ' ** Function: Below.
20020     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
20030     If gblnUseReveuneExpenseCodes = True Then
            '##GTR_Ref: rptCourtRptCA_02A
20040       PreviewOrPrint "2A", THIS_PROC, acViewNormal  ' ** Procedure: Below.
20050     Else
            '##GTR_Ref: rptCourtRptCA_02
20060       PreviewOrPrint "2", THIS_PROC, acViewNormal  ' ** Procedure: Below.
20070     End If
20080   End If  ' ** Validate

EXITP:
20090   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord02_Click()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord02_Click"

20110   If Validate = True Then  ' ** Function: Below.
20120     If gblnUseReveuneExpenseCodes = True Then
20130       SendToFile_CA "2A", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
20140     Else
20150       SendToFile_CA "2", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
20160     End If
20170   End If  ' ** Validate

EXITP:
20180   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel02_Click()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel02_Click"

20210   If Validate = True Then  ' ** Function: Below.
20220     If gblnUseReveuneExpenseCodes = True Then
20230       SendToFile_CA "2A", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
20240     Else
20250       SendToFile_CA "2", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
20260     End If
20270   End If  ' ** Validate

EXITP:
20280   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview03_Click()
' ** Gains on Sales.

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview03_Click"

20310   If Validate = True Then  ' ** Function: Below.

20320     DoCmd.Hourglass True

20330     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

20340     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
20350       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
20360         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
20370           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
20380             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
20390             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
20400             CurrentDb.QueryDefs.Refresh
20410             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
20420             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
20430             CurrentDb.QueryDefs.Refresh
20440             CurrentDb.QueryDefs.Refresh
20450           End If
20460         End If
20470       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_03!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
20480       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
20490         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Then
20500           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
20510             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
20520             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
20530             CurrentDb.QueryDefs.Refresh
20540             CurrentDb.QueryDefs.Refresh
20550           End If
20560         End If
20570       End If

20580     End With

20590     PreviewOrPrint "3", THIS_PROC, acViewPreview  ' ** Procedure: Below.

20600   End If  ' ** Validate

EXITP:
20610   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdPrint03_Click()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint03_Click"

20710   If Validate = True Then  ' ** Function: Below.

20720     DoCmd.Hourglass True

20730     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

20740     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
20750       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") = 0 Then
              '    If QueryExists("qryCourtReport_CA_00s_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
              '      DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
              '      DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
20760         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
20770           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
20780             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
20790             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
20800             CurrentDb.QueryDefs.Refresh
20810             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
20820             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
20830             CurrentDb.QueryDefs.Refresh
20840             CurrentDb.QueryDefs.Refresh
20850           End If
20860         End If
              'End Select
20870       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_03!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
20880       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") = 0 Then
              '    If QueryExists("qryCourtReport-B_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport-B"
              '      DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue
20890         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Then
20900           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
20910             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
20920             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
20930             CurrentDb.QueryDefs.Refresh
20940             CurrentDb.QueryDefs.Refresh
20950           End If
20960         End If
              'End Select
20970       End If

20980     End With

          '##GTR_Ref: rptCourtRptCA_03
20990     PreviewOrPrint "3", THIS_PROC, acViewNormal  ' ** Procedure: Below.

21000   End If  ' ** Validate

EXITP:
21010   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Public Sub cmdWord03_Click()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord03_Click"

21110   If Validate = True Then  ' ** Function: Below.

21120     DoCmd.Hourglass True

21130     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
21140       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") = 0 Then
              '    If QueryExists("qryCourtReport_CA_00s_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
              '      DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
              '      DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
21150         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
21160           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
21170             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
21180             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
21190             CurrentDb.QueryDefs.Refresh
21200             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
21210             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
21220             CurrentDb.QueryDefs.Refresh
21230             CurrentDb.QueryDefs.Refresh
21240           End If
21250         End If
              'End Select
21260       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_03!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
21270       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") = 0 Then
              '    If QueryExists("qryCourtReport-B_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport-B"
              '      DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue
21280         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Then
21290           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
21300             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
21310             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
21320             CurrentDb.QueryDefs.Refresh
21330             CurrentDb.QueryDefs.Refresh
21340           End If
21350         End If
              'End Select
21360       End If

21370     End With

21380     SendToFile_CA "3", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.

21390   End If  ' ** Validate

EXITP:
21400   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Public Sub cmdExcel03_Click()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel03_Click"

21510   If Validate = True Then  ' ** Function: Below.

21520     DoCmd.Hourglass True

21530     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
21540       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") = 0 Then
              '    If QueryExists("qryCourtReport_CA_00s_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
              '      DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
21550         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
21560           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
21570             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
21580             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
21590             CurrentDb.QueryDefs.Refresh
21600             CurrentDb.QueryDefs.Refresh
21610           End If
21620         End If
              'End Select
21630       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_03!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
21640       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") = 0 Then
              '    If QueryExists("qryCourtReport-B_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport-B"
              '      DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue
21650         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_03a_04") > 0 Then
21660           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
21670             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
21680             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
21690             CurrentDb.QueryDefs.Refresh
21700             CurrentDb.QueryDefs.Refresh
21710           End If
21720         End If
              'End Select
21730       End If

            ' ###############################################################################################################
            ' ## qryCourtReport_CA_03a IS THE BASIS FOR THE QUERY THAT GETS EXPORTED TO EXCEL, qryCourtReport_CA_03i!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport_CA_03a_bak1  ' ** Gains: Ledger, Account, MasterAsset, grouped and summed.    ####  Individual Gains.  ####
            ' ** qryCourtReport_CA_03a_bak2  ' ** Gains: .._CA_03a_04.                                        ####  Combined Gains.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
21740       If QueryExists("qryCourtReport_CA_03a") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_03a").SQL, "_03a_04") = 0 Then
              '    If QueryExists("qryCourtReport_CA_03a_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
              '      DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
21750         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_03a").SQL, "_03a_04") > 0 Then
21760           If QueryExists("qryCourtReport_CA_03a_bak1") = True Then  ' ** Module Function: modFileUtilities.
21770             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_03a"
21780             DoCmd.CopyObject , "qryCourtReport_CA_03a", acQuery, "qryCourtReport_CA_03a_bak1"
21790             CurrentDb.QueryDefs.Refresh
21800             CurrentDb.QueryDefs.Refresh
21810           End If
21820         End If
              'End Select
21830       End If

21840     End With

21850     SendToFile_CA "3", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.

21860   End If  ' ** Validate

EXITP:
21870   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdPreview04_lbl4_cmd_DblClick(Cancel As Integer)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_lbl4_cmd_DblClick"

21910   With Me
21920     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
21930     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
21940   End With

EXITP:
21950   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview04_lbl6_cmd_DblClick(Cancel As Integer)

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_lbl6_cmd_DblClick"

22010   With Me
22020     .chkIncludeCheckNum = Not .chkIncludeCheckNum
22030     chkIncludeCheckNum_AfterUpdate  ' ** Procedure: Above.
22040   End With

EXITP:
22050   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview04_Click()

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_Click"

22110   If Validate = True Then  ' ** Function: Below.
22120     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
22130     If gblnUseReveuneExpenseCodes = True Then
22140       PreviewOrPrint "4A", THIS_PROC, acViewPreview  ' ** Procedure: Below.
22150     Else
22160       PreviewOrPrint "4", THIS_PROC, acViewPreview  ' ** Procedure: Below.
22170     End If
22180   End If  ' ** Validate

EXITP:
22190   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint04_Click()

22200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint04_Click"

22210   If Validate = True Then  ' ** Function: Below.
22220     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
22230     If gblnUseReveuneExpenseCodes = True Then
            '##GTR_Ref: rptCourtRptCA_04A
            '##GTR_Ref: rptCourtRptCA_04AH
22240       PreviewOrPrint "4A", THIS_PROC, acViewNormal  ' ** Procedure: Below.
22250     Else
            '##GTR_Ref: rptCourtRptCA_04
            '##GTR_Ref: rptCourtRptCA_04H
22260       PreviewOrPrint "4", THIS_PROC, acViewNormal  ' ** Procedure: Below.
22270     End If
22280   End If  ' ** Validate

EXITP:
22290   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord04_Click()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord04_Click"

22310   If Validate = True Then  ' ** Function: Below.
22320     If gblnUseReveuneExpenseCodes = True Then
22330       SendToFile_CA "4A", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
22340     Else
22350       SendToFile_CA "4", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
22360     End If
22370   End If  ' ** Validate

EXITP:
22380   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel04_Click()

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel04_Click"

22410   If Validate = True Then  ' ** Function: Below.
22420     If gblnUseReveuneExpenseCodes = True Then
22430       SendToFile_CA "4A", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
22440     Else
22450       SendToFile_CA "4", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
22460     End If
22470   End If  ' ** Validate

EXITP:
22480   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview05_Click()
' ** Losses on Sales.

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview05_Click"

22510   If Validate = True Then  ' ** Function: Below.

22520     DoCmd.Hourglass True

22530     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

22540     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
22550       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") = 0 Then
              '    If QueryExists("qryCourtReport_CA_00s_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
              '      DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
              '      DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
22560         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
22570           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
22580             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
22590             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
22600             CurrentDb.QueryDefs.Refresh
22610             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
22620             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
22630             CurrentDb.QueryDefs.Refresh
22640             CurrentDb.QueryDefs.Refresh
22650           End If
22660         End If
              'End Select
22670       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
22680       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") = 0 Then
              '    If QueryExists("qryCourtReport-B_bak3") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport-B"
              '      DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak3"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue
22690         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
22700           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
22710             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
22720             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
22730             CurrentDb.QueryDefs.Refresh
22740             CurrentDb.QueryDefs.Refresh
22750           End If
22760         End If
              'End Select
22770       End If

22780     End With

22790     PreviewOrPrint "5", THIS_PROC, acViewPreview  ' ** Procedure: Below.

22800   End If  ' ** Validate

EXITP:
22810   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdPrint05_Click()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint05_Click"

22910   If Validate = True Then  ' ** Function: Below.

22920     DoCmd.Hourglass True

22930     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

22940     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
22950       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") = 0 Then
              '    If QueryExists("qryCourtReport_CA_00s_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
              '      DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
              '      DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
22960         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
22970           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
22980             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
22990             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
23000             CurrentDb.QueryDefs.Refresh
23010             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
23020             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
23030             CurrentDb.QueryDefs.Refresh
23040             CurrentDb.QueryDefs.Refresh
23050           End If
23060         End If
              'End Select
23070       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
23080       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") = 0 Then
              '    If QueryExists("qryCourtReport-B_bak3") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport-B"
              '      DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak3"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue
23090         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
23100           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
23110             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
23120             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
23130             CurrentDb.QueryDefs.Refresh
23140             CurrentDb.QueryDefs.Refresh
23150           End If
23160         End If
              'End Select
23170       End If

23180     End With

          '##GTR_Ref: rptCourtRptCA_05
23190     PreviewOrPrint "5", THIS_PROC, acViewNormal  ' ** Procedure: Below.

23200   End If  ' ** Validate

EXITP:
23210   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Public Sub cmdWord05_Click()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord05_Click"

23310   If Validate = True Then  ' ** Function: Below.

23320     DoCmd.Hourglass True

23330     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
23340       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") = 0 Then
              '    If QueryExists("qryCourtReport_CA_00s_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
              '      DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
              '      DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
23350         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
23360           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
23370             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
23380             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
23390             CurrentDb.QueryDefs.Refresh
23400             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
23410             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
23420             CurrentDb.QueryDefs.Refresh
23430             CurrentDb.QueryDefs.Refresh
23440           End If
23450         End If
              'End Select
23460       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
23470       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") = 0 Then
              '    If QueryExists("qryCourtReport-B_bak3") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport-B"
              '      DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak3"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue
23480         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
23490           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
23500             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
23510             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
23520             CurrentDb.QueryDefs.Refresh
23530             CurrentDb.QueryDefs.Refresh
23540           End If
23550         End If
              'End Select
23560       End If

23570     End With

23580     SendToFile_CA "5", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.

23590   End If  ' ** Validate

EXITP:
23600   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Public Sub cmdExcel05_Click()

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel05_Click"

23710   If Validate = True Then  ' ** Function: Below.

23720     DoCmd.Hourglass True

23730     With Me

            ' ** Switch queries:
            ' ** qryCourtReport_CA_00s_bak1  ' ** RecordSource; rptCourtRptCA-0; tmpCourtReportData.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport_CA_00s_bak2  ' ** RecordSource; rptCourtRptCA-0; .._CA_00s_09.          ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
23740       If QueryExists("qryCourtReport_CA_00s") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") = 0 Then
              '    If QueryExists("qryCourtReport_CA_00s_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
              '      DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
23750         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_00s").SQL, "_00s_09") > 0 Then
23760           If QueryExists("qryCourtReport_CA_00s_bak1") = True Then  ' ** Module Function: modFileUtilities.
23770             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_00s"
23780             DoCmd.CopyObject , "qryCourtReport_CA_00s", acQuery, "qryCourtReport_CA_00s_bak1"
23790             CurrentDb.QueryDefs.Refresh
23800             CurrentDb.QueryDefs.Refresh
23810           End If
23820         End If
              'End Select
23830       End If

            ' ###############################################################################################################
            ' ## qryCourtReport-B IS THE RECORD SOURCE FOR rptCourtRptCA_05!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport-B_bak1  ' ** Same as qryCapitalGainsAndLoss1.    ####  Individual Gains and Losses.  ####
            ' ** qryCourtReport-B_bak2  ' ** .._CA_03a_04.                       ####  Combined Gains and Losses.    ####
            ' ** qryCourtReport-B_bak3  ' ** .._CA_05a_04.                       ####  Combined Gains and Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
23840       If QueryExists("qryCourtReport-B") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") = 0 Then
              '    If QueryExists("qryCourtReport-B_bak3") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport-B"
              '      DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak3"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue
23850         If InStr(CurrentDb.QueryDefs("qryCourtReport-B").SQL, "_05a_04") > 0 Then
23860           If QueryExists("qryCourtReport-B_bak1") = True Then  ' ** Module Function: modFileUtilities.
23870             DoCmd.DeleteObject acQuery, "qryCourtReport-B"
23880             DoCmd.CopyObject , "qryCourtReport-B", acQuery, "qryCourtReport-B_bak1"
23890             CurrentDb.QueryDefs.Refresh
23900             CurrentDb.QueryDefs.Refresh
23910           End If
23920         End If
              'End Select
23930       End If

            ' ###############################################################################################################
            ' ## qryCourtReport_CA_05a IS THE BASIS FOR THE QUERY THAT GETS EXPORTED TO EXCEL, qryCourtReport_CA_05i!
            ' ###############################################################################################################

            ' ** Switch queries:
            ' ** qryCourtReport_CA_05a_bak1  ' ** Losses: Ledger, Account, MasterAsset, grouped and summed.    ####  Individual Losses.  ####
            ' ** qryCourtReport_CA_05a_bak2  ' ** Losses: .._CA_05a_04.                                        ####  Combined Losses.    ####
            'DEFAULT VALUE: 2 : .opgGainLoss_optAllSplit.OptionValue
23940       If QueryExists("qryCourtReport_CA_05a") = True Then  ' ** Module Function: modFileUtilities.
              'Select Case .opgGainLoss
              'Case .opgGainLoss_optAllCombined.OptionValue
              '  If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_05a").SQL, "_05a_04") = 0 Then
              '    If QueryExists("qryCourtReport_CA_05a_bak2") = True Then  ' ** Module Function: modFileUtilities.
              '      DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
              '      DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak2"
              '      CurrentDb.QueryDefs.Refresh
              '      CurrentDb.QueryDefs.Refresh
              '    End If
              '  End If
              'Case .opgGainLoss_optAllSplit.OptionValue, .opgGainLoss_optGLSplitAssetsCombined.OptionValue
23950         If InStr(CurrentDb.QueryDefs("qryCourtReport_CA_05a").SQL, "_05a_04") > 0 Then
23960           If QueryExists("qryCourtReport_CA_05a_bak1") = True Then  ' ** Module Function: modFileUtilities.
23970             DoCmd.DeleteObject acQuery, "qryCourtReport_CA_05a"
23980             DoCmd.CopyObject , "qryCourtReport_CA_05a", acQuery, "qryCourtReport_CA_05a_bak1"
23990             CurrentDb.QueryDefs.Refresh
24000             CurrentDb.QueryDefs.Refresh
24010           End If
24020         End If
              'End Select
24030       End If

24040     End With

24050     SendToFile_CA "5", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.

24060   End If  ' ** Validate

EXITP:
24070   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case Else
1510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1520    End Select
1530    Resume EXITP

End Sub

Private Sub cmdPreview06_lbl6_cmd_DblClick(Cancel As Integer)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview06_lbl6_cmd_DblClick"

24110   With Me
24120     .chkIncludeCheckNum = Not .chkIncludeCheckNum
24130     chkIncludeCheckNum_AfterUpdate  ' ** Procedure: Above.
24140   End With

EXITP:
24150   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview06_Click()

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview06_Click"

24210   If Validate = True Then  ' ** Function: Below.
24220     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
24230     PreviewOrPrint "6", THIS_PROC, acViewPreview  ' ** Procedure: Below.
24240   End If  ' ** Validate

EXITP:
24250   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint06_Click()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint06_Click"

24310   If Validate = True Then  ' ** Function: Below.
24320     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_06
24330     PreviewOrPrint "6", THIS_PROC, acViewNormal  ' ** Procedure: Below.
24340   End If  ' ** Validate

EXITP:
24350   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord06_Click()

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord06_Click"

24410   If Validate = True Then  ' ** Function: Below.
24420     SendToFile_CA "6", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
24430   End If  ' ** Validate

EXITP:
24440   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel06_Click()

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel06_Click"

24510   If Validate = True Then  ' ** Function: Below.
24520     SendToFile_CA "6", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
24530   End If  ' ** Validate

EXITP:
24540   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview07_Click()

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview07_Click"

24610   If Validate = True Then  ' ** Function: Below.
24620     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
24630     PreviewOrPrint "7", THIS_PROC, acViewPreview  ' ** Procedure: Below.
24640   End If  ' ** Validate

EXITP:
24650   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint07_Click()

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint07_Click"

24710   If Validate = True Then  ' ** Function: Below.
24720     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_07
24730     PreviewOrPrint "7", THIS_PROC, acViewNormal  ' ** Procedure: Below.
24740   End If  ' ** Validate

EXITP:
24750   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord07_Click()

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord07_Click"

24810   If Validate = True Then  ' ** Function: Below.
24820     SendToFile_CA "7", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
24830   End If  ' ** Validate

EXITP:
24840   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel07_Click()

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel07_Click"

24910   If Validate = True Then  ' ** Function: Below.
24920     SendToFile_CA "7", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
24930   End If  ' ** Validate

EXITP:
24940   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview08_Click()

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview08_Click"

25010   If Validate = True Then  ' ** Function: Below.
25020     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
25030     PreviewOrPrint "8", THIS_PROC, acViewPreview  ' ** Procedure: Below.
25040   End If  ' ** Validate

EXITP:
25050   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint08_Click()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint08_Click"

25110   If Validate = True Then  ' ** Function: Below.
25120     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_08
25130     PreviewOrPrint "8", THIS_PROC, acViewNormal  ' ** Procedure: Below.
25140   End If  ' ** Validate

EXITP:
25150   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord08_Click()

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord08_Click"

25210   If Validate = True Then  ' ** Function: Below.
25220     SendToFile_CA "8", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
25230   End If  ' ** Validate

EXITP:
25240   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel08_Click()

25300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel08_Click"

25310   If Validate = True Then  ' ** Function: Below.
25320     SendToFile_CA "8", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
25330   End If  ' ** Validate

EXITP:
25340   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview09_Click()

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview09_Click"

25410   If Validate = True Then  ' ** Function: Below.
25420     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
25430     PreviewOrPrint "9", THIS_PROC, acViewPreview  ' ** Procedure: Below.
25440   End If  ' ** Validate

EXITP:
25450   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint09_Click()

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint09_Click"

25510   If Validate = True Then  ' ** Function: Below.
25520     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_09
25530     PreviewOrPrint "9", THIS_PROC, acViewNormal  ' ** Procedure: Below.
25540   End If  ' ** Validate

EXITP:
25550   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord09_Click()

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord09_Click"

25610   If Validate = True Then  ' ** Function: Below.
25620     SendToFile_CA "9", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
25630   End If  ' ** Validate

EXITP:
25640   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel09_Click()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel09_Click"

25710   If Validate = True Then  ' ** Function: Below.
25720     SendToFile_CA "9", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
25730   End If  ' ** Validate

EXITP:
25740   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel09_KeyDown(KeyCode As Integer, Shift As Integer)

25800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel09_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25820   intShiftDown = (Shift And acShiftMask) > 0
25830   intAltDown = (Shift And acAltMask) > 0
25840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25860     Select Case intRetVal
          Case vbKeyTab
25870       With Me
25880         intRetVal = 0
25890         .cmdPrintAll.SetFocus
25900       End With
25910     Case vbKeyLeft
25920       With Me
25930         intRetVal = 0
25940         .cmdWord09.SetFocus
25950       End With
25960     Case vbKeyRight
25970       With Me
25980         intRetVal = 0
25990         .cmdPrintAll.SetFocus
26000       End With
26010     End Select
26020   End If

        ' ** Shift keys.
26030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26040     Select Case intRetVal
          Case vbKeyTab
26050       With Me
26060         intRetVal = 0
26070         .cmdWord09.SetFocus
26080       End With
26090     End Select
26100   End If

EXITP:
26110   KeyCode = intRetVal
26120   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdPreview10_lbl6_cmd_DblClick(Cancel As Integer)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview10_lbl6_cmd_DblClick"

26210   With Me
26220     Select Case .cmdPreview10_lbl6_box.Visible
          Case True
26230       .cmdPreview10_lbl6c.Visible = False
26240       .cmdPreview10_lbl6b.Visible = False
26250       .cmdPreview10_lbl6a.Visible = False
26260       .cmdPreview10_lbl6_box.Visible = False
26270     Case False
26280       .cmdPreview10_lbl6_box.Visible = True
26290       .cmdPreview10_lbl6a.Visible = True
26300       .cmdPreview10_lbl6b.Visible = True
26310       .cmdPreview10_lbl6c.Visible = True
26320     End Select
26330   End With

EXITP:
26340   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview10_lbl6a_DblClick(Cancel As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview10_lbl6a_DblClick"

26410   With Me
26420     .cmdPreview10_lbl6c.Visible = False
26430     .cmdPreview10_lbl6b.Visible = False
26440     .cmdPreview10_lbl6a.Visible = False
26450     .cmdPreview10_lbl6_box.Visible = False
26460   End With

EXITP:
26470   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview10_Click()

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview10_Click"

26510   If Validate = True Then  ' ** Function: Below.
26520     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
26530     PreviewOrPrint "10", THIS_PROC, acViewPreview  ' ** Procedure: Below.
26540   End If  ' ** Validate

EXITP:
26550   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint10_Click()

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint10_Click"

26610   If Validate = True Then  ' ** Function: Below.
26620     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_10
26630     PreviewOrPrint "10", THIS_PROC, acViewNormal  ' ** Procedure: Below.
26640   End If  ' ** Validate

EXITP:
26650   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord10_Click()

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord10_Click"

26710   If Validate = True Then  ' ** Function: Below.
26720     SendToFile_CA "10", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
26730   End If  ' ** Validate

EXITP:
26740   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel10_Click()

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel10_Click"

26810   If Validate = True Then  ' ** Function: Below.
26820     SendToFile_CA "10", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
26830   End If  ' ** Validate

EXITP:
26840   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview11_lbl6_cmd_DblClick(Cancel As Integer)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview11_lbl6_cmd_DblClick"

26910   With Me
26920     Select Case .cmdPreview11_lbl6_box.Visible
          Case True
26930       .cmdPreview11_lbl6c.Visible = False
26940       .cmdPreview11_lbl6b.Visible = False
26950       .cmdPreview11_lbl6a.Visible = False
26960       .cmdPreview11_lbl6_box.Visible = False
26970     Case False
26980       .cmdPreview11_lbl6_box.Visible = True
26990       .cmdPreview11_lbl6a.Visible = True
27000       .cmdPreview11_lbl6b.Visible = True
27010       .cmdPreview11_lbl6c.Visible = True
27020     End Select
27030   End With

EXITP:
27040   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview11_lbl6a_DblClick(Cancel As Integer)

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview11_lbl6a_DblClick"

27110   With Me
27120     .cmdPreview11_lbl6c.Visible = False
27130     .cmdPreview11_lbl6b.Visible = False
27140     .cmdPreview11_lbl6a.Visible = False
27150     .cmdPreview11_lbl6_box.Visible = False
27160   End With

EXITP:
27170   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPreview11_Click()

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview11_Click"

27210   If Validate = True Then  ' ** Function: Below.
27220     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
27230     PreviewOrPrint "11", THIS_PROC, acViewPreview  ' ** Procedure: Below.
27240   End If  ' ** Validate

EXITP:
27250   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrint11_Click()

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint11_Click"

27310   If Validate = True Then  ' ** Function: Below.
27320     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
          '##GTR_Ref: rptCourtRptCA_11
27330     PreviewOrPrint "11", THIS_PROC, acViewNormal  ' ** Procedure: Below.
27340   End If  ' ** Validate

EXITP:
27350   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdWord11_Click()

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord11_Click"

27410   If Validate = True Then  ' ** Function: Below.
27420     SendToFile_CA "11", THIS_PROC, blnRebuildTable, Me, False  ' ** Form Procedure: frmRpt_CourtReports_CA.
27430   End If  ' ** Validate

EXITP:
27440   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub cmdExcel11_Click()

27500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel11_Click"

27510   If Validate = True Then  ' ** Function: Below.
27520     SendToFile_CA "11", THIS_PROC, blnRebuildTable, Me, True  ' ** Form Procedure: frmRpt_CourtReports_CA.
27530   End If  ' ** Validate

EXITP:
27540   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

27610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27620   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

27710   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
27720   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

        Dim strOrd As String, strVer As String, strCashBeg As String, dblCashBeg As Double, strCashEnd As String, dblCashEnd As Double
        Dim intRetVal As Integer

27810   If Validate = True Then  ' ** Function: Below.

27820     blnAllCancel = False
27830     AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
27840     blnAccountChecked = False
27850     gblnPrintAll = True

          ' ** Set up the TAReports Command Bar (visible with Print Preview).
27860     TAReports_SetZero False  ' ** Module Function: modReportFunctions.
27870     ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

          ' ** Get the Summary inputs first.
27880     intRetVal = CAGetCourtReportData(THIS_PROC & "^" & "cmdPrintAll")  ' ** Module Function: modCourtReportsCA.
27890     If intRetVal <> 0 Then
27900       blnAllCancel = True
27910       AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
27920     Else
            ' ** Save these for later.
27930       strOrd = gstrCrtRpt_Ordinal
27940       strVer = gstrCrtRpt_Version
27950       If Trim(gstrCrtRpt_CashAssets_Beg) = vbNullString Then gstrCrtRpt_CashAssets_Beg = "0"
27960       If IsNumeric(Trim(gstrCrtRpt_CashAssets_Beg)) = False Then gstrCrtRpt_CashAssets_Beg = "0"
27970       If Trim(gstrCrtRpt_CashAssets_End) = vbNullString Then gstrCrtRpt_CashAssets_End = "0"
27980       If IsNumeric(Trim(gstrCrtRpt_CashAssets_End)) = False Then gstrCrtRpt_CashAssets_End = "0"
27990       strCashBeg = gstrCrtRpt_CashAssets_Beg
28000       dblCashBeg = Nz(Me.CashAssets_Beg, 0)
28010       strCashEnd = gstrCrtRpt_CashAssets_End
28020       dblCashEnd = Nz(Me.CashAssets_End, 0)
28030     End If

28040     If blnAllCancel = False Then
28050       PreviewOrPrint "11", THIS_PROC & " - R11", acViewNormal  ' ** Procedure: Below.
28060     End If
28070     If blnAllCancel = False Then
28080       PreviewOrPrint "10", THIS_PROC & " - R10", acViewNormal  ' ** Procedure: Below.
28090     End If
28100     If blnAllCancel = False Then
28110       PreviewOrPrint "9", THIS_PROC & " - R9", acViewNormal  ' ** Procedure: Below.
28120     End If
28130     If blnAllCancel = False Then
28140       PreviewOrPrint "8", THIS_PROC & " - R8", acViewNormal  ' ** Procedure: Below.
28150     End If
28160     If blnAllCancel = False Then
28170       PreviewOrPrint "7", THIS_PROC & " - R7", acViewNormal  ' ** Procedure: Below.
28180     End If
28190     If blnAllCancel = False Then
28200       PreviewOrPrint "6", THIS_PROC & " - R6", acViewNormal  ' ** Procedure: Below.
28210     End If
28220     If blnAllCancel = False Then
28230       PreviewOrPrint "5", THIS_PROC & " - R5", acViewNormal  ' ** Procedure: Below.
28240     End If
28250     If blnAllCancel = False Then
28260       Select Case gblnUseReveuneExpenseCodes
            Case True
28270         PreviewOrPrint "4A", THIS_PROC & " - R4A", acViewNormal  ' ** Procedure: Below.
28280       Case False
28290         PreviewOrPrint "4", THIS_PROC & " - R4", acViewNormal  ' ** Procedure: Below.
28300       End Select
28310     End If
28320     If blnAllCancel = False Then
28330       PreviewOrPrint "3", THIS_PROC & " - R3", acViewNormal  ' ** Procedure: Below.
28340     End If
28350     If blnAllCancel = False Then
28360       PreviewOrPrint "2", THIS_PROC & " - R2", acViewNormal  ' ** Procedure: Below.
28370     End If
28380     If blnAllCancel = False Then
28390       PreviewOrPrint "1", THIS_PROC & " - R1", acViewNormal  ' ** Procedure: Below.
28400     End If

          ' ** Set up the TAReports Command Bar (visible with Print Preview).
28410     TAReports_SetZero True  ' ** Module Function: modReportFunctions.

28420     If blnAllCancel = False Then
28430       gstrCrtRpt_Ordinal = strOrd
28440       gstrCrtRpt_Version = strVer
28450       gstrCrtRpt_CashAssets_Beg = strCashBeg
28460       gdblCrtRpt_CA_COHBeg = dblCashBeg
28470       Me.CashAssets_Beg = dblCashBeg
28480       gstrCrtRpt_CashAssets_End = strCashEnd
28490       gdblCrtRpt_CA_COHEnd = dblCashEnd
28500       Me.CashAssets_End = dblCashEnd
28510       Select Case gblnUseReveuneExpenseCodes
            Case True
28520         PreviewOrPrint "0A", THIS_PROC & " - R0A", acViewNormal  ' ** Procedure: Below.
28530       Case False
28540         PreviewOrPrint "0", THIS_PROC & " - R0", acViewNormal  ' ** Procedure: Below.
28550       End Select
28560     End If

28570     gblnPrintAll = False

28580   End If  ' ** Validate

EXITP:
28590   Exit Sub

ERRH:
1490    gblnPrintAll = False
1500    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
1510      blnAllCancel = True
1520      AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
1530    Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28620   intShiftDown = (Shift And acShiftMask) > 0
28630   intAltDown = (Shift And acAltMask) > 0
28640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28660     Select Case intRetVal
          Case vbKeyTab
28670       With Me
28680         intRetVal = 0
28690         .cmdWordAll.SetFocus
28700       End With
28710     Case vbKeyUp
28720       With Me
28730         intRetVal = 0
28740         If strLastArrow = vbNullString Then
28750           .cmdPrint09.SetFocus
28760         Else
28770           If Right(strLastArrow, 2) = "09" Or Right(strLastArrow, 2) = "00" Then
28780             .Controls(Left(strLastArrow, (InStr(strLastArrow, "0") - 1)) & "09").SetFocus
28790           Else
28800             .cmdPrint09.SetFocus
28810           End If
28820           strLastArrow = vbNullString
28830         End If
28840       End With
28850     Case vbKeyDown
28860       With Me
28870         intRetVal = 0
28880         If strLastArrow = vbNullString Then
28890           .cmdPrint00.SetFocus
28900         Else
28910           If Right(strLastArrow, 2) = "09" Or Right(strLastArrow, 2) = "00" Then
28920             .Controls(Left(strLastArrow, (InStr(strLastArrow, "0") - 1)) & "00").SetFocus
28930           Else
28940             .cmdPrint00.SetFocus
28950           End If
28960           strLastArrow = vbNullString
28970         End If
28980       End With
28990     Case vbKeyLeft
29000       With Me
29010         intRetVal = 0
29020         If .cmdExcel09.Enabled = True Then
29030           .cmdExcel09.SetFocus
29040         Else
29050           .cmdWord09.SetFocus
29060         End If
29070       End With
29080     Case vbKeyRight
29090       With Me
29100         intRetVal = 0
29110         .cmdWordAll.SetFocus
29120       End With
29130     End Select
29140   End If

        ' ** Shift keys.
29150   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29160     Select Case intRetVal
          Case vbKeyTab
29170       With Me
29180         intRetVal = 0
29190         If .cmdExcel09.Enabled = True Then
29200           .cmdExcel09.SetFocus
29210         Else
29220           .cmdWord09.SetFocus
29230         End If
29240       End With
29250     End Select
29260   End If

EXITP:
29270   KeyCode = intRetVal
29280   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

29310   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
29320   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

29410   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
29420   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

29510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
29520   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

29600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

29610   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
29620   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

29710   With Me

29720     lngFiles = 0&
29730     ReDim arr_varFile(F_ELEMS, 0)

29740     WordAll_CA Me  ' ** Module Procedure: modCourtReportsCA.

29750   End With

EXITP:
29760   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29820   intShiftDown = (Shift And acShiftMask) > 0
29830   intAltDown = (Shift And acAltMask) > 0
29840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29860     Select Case intRetVal
          Case vbKeyTab
29870       With Me
29880         intRetVal = 0
29890         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
29900           .cmdExcelAll.SetFocus
29910         Else
29920           .cmdClose.SetFocus
29930         End If
29940       End With
29950     Case vbKeyUp
29960       With Me
29970         intRetVal = 0
29980         .cmdWord09.SetFocus
29990         strLastArrow = vbNullString
30000       End With
30010     Case vbKeyDown
30020       With Me
30030         intRetVal = 0
30040         .cmdWord00.SetFocus
30050         strLastArrow = vbNullString
30060       End With
30070     Case vbKeyRight
30080       With Me
30090         intRetVal = 0
30100         If .cmdExcelAll.Enabled = True Then
30110           .cmdExcelAll.SetFocus
30120         Else
30130           .cmdClose.SetFocus
30140         End If
30150       End With
30160     End Select
30170   End If

        ' ** Shift keys.
30180   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30190     Select Case intRetVal
          Case vbKeyTab
30200       With Me
30210         intRetVal = 0
30220         .cmdPrintAll.SetFocus
30230       End With
30240     End Select
30250   End If

        ' ** Ctrl keys.
30260   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
30270     Select Case intRetVal
          Case vbKeyTab
30280       With Me
30290         intRetVal = 0
30300         .chkOpenWord.SetFocus
30310       End With
30320     End Select
30330   End If

EXITP:
30340   KeyCode = intRetVal
30350   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

30410   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
30420   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

30510   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
30520   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub chkOpenWord_AfterUpdate()

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

30610   With Me
30620     Select Case .chkOpenWord
          Case True
30630       .chkOpenWord_lbl.FontBold = True
30640       .chkOpenWord_lbl2.FontBold = True
30650       .chkOpenWord_lbl_dim_hi.FontBold = True
30660       .chkOpenWord_lbl2_dim_hi.FontBold = True
30670     Case False
30680       .chkOpenWord_lbl.FontBold = False
30690       .chkOpenWord_lbl2.FontBold = False
30700       .chkOpenWord_lbl_dim_hi.FontBold = False
30710       .chkOpenWord_lbl2_dim_hi.FontBold = False
30720     End Select
30730   End With

EXITP:
30740   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30820   intShiftDown = (Shift And acShiftMask) > 0
30830   intAltDown = (Shift And acAltMask) > 0
30840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30860     Select Case intRetVal
          Case vbKeyTab
30870       With Me
30880         intRetVal = 0
30890         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
30900           .cmdExcelAll.SetFocus
30910         Else
30920           .cmdClose.SetFocus
30930         End If
30940       End With
30950     Case vbKeyUp
30960       With Me
30970         intRetVal = 0
30980         .cmdWordAll.SetFocus
30990       End With
31000     Case vbKeyDown
31010       With Me
31020         intRetVal = 0
31030         .cmdWord00.SetFocus
31040       End With
31050     End Select
31060   End If

        ' ** Shift keys.
31070   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
31080     Select Case intRetVal
          Case vbKeyTab
31090       With Me
31100         intRetVal = 0
31110         .cmdWordAll.SetFocus
31120       End With
31130     End Select
31140   End If

        ' ** Ctrl keys.
31150   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
31160     Select Case intRetVal
          Case vbKeyTab
31170       With Me
31180         intRetVal = 0
31190         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
31200           .chkOpenExcel.SetFocus
31210         ElseIf .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
31220           .UserReportPath_chk.SetFocus
31230         Else
31240           Beep
31250         End If
31260       End With
31270     End Select
31280   End If

EXITP:
31290   KeyCode = intRetVal
31300   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

31410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
31420   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

31500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

31510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
31520   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

31600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

31610   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
31620   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

31710   With Me

31720     lngFiles = 0&
31730     ReDim arr_varFile(F_ELEMS, 0)

31740     ExcelAll_CA Me  ' ** Module Procedure: modCourtReportsCA.

31750   End With

EXITP:
31760   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

31800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31820   intShiftDown = (Shift And acShiftMask) > 0
31830   intAltDown = (Shift And acAltMask) > 0
31840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31860     Select Case intRetVal
          Case vbKeyTab
31870       With Me
31880         intRetVal = 0
31890         .cmdClose.SetFocus
31900       End With
31910     Case vbKeyUp
31920       With Me
31930         intRetVal = 0
31940         .cmdExcel09.SetFocus
31950         strLastArrow = vbNullString
31960       End With
31970     Case vbKeyDown
31980       With Me
31990         intRetVal = 0
32000         .cmdExcel00.SetFocus
32010         strLastArrow = vbNullString
32020       End With
32030     Case vbKeyRight
32040       With Me
32050         intRetVal = 0
32060         .cmdClose.SetFocus
32070       End With
32080     End Select
32090   End If

        ' ** Shift keys.
32100   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32110     Select Case intRetVal
          Case vbKeyTab
32120       With Me
32130         intRetVal = 0
32140         .cmdWordAll.SetFocus
32150       End With
32160     End Select
32170   End If

        ' ** Ctrl keys.
32180   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
32190     Select Case intRetVal
          Case vbKeyTab
32200       With Me
32210         intRetVal = 0
32220         .chkOpenExcel.SetFocus
32230       End With
32240     End Select
32250   End If

EXITP:
32260   KeyCode = intRetVal
32270   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

32310   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
32320   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

32400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

32410   PrintWordExcel_Handler_CA THIS_PROC, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
32420   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub chkOpenExcel_AfterUpdate()

32500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

32510   With Me
32520     Select Case .chkOpenExcel
          Case True
32530       .chkOpenExcel_lbl.FontBold = True
32540       .chkOpenExcel_lbl2.FontBold = True
32550       .chkOpenExcel_lbl_dim_hi.FontBold = True
32560       .chkOpenExcel_lbl2_dim_hi.FontBold = True
32570     Case False
32580       .chkOpenExcel_lbl.FontBold = False
32590       .chkOpenExcel_lbl2.FontBold = False
32600       .chkOpenExcel_lbl_dim_hi.FontBold = False
32610       .chkOpenExcel_lbl2_dim_hi.FontBold = False
32620     End Select
32630   End With

EXITP:
32640   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub chkOpenExcel_KeyDown(KeyCode As Integer, Shift As Integer)

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
32720   intShiftDown = (Shift And acShiftMask) > 0
32730   intAltDown = (Shift And acAltMask) > 0
32740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32760     Select Case intRetVal
          Case vbKeyTab
32770       With Me
32780         intRetVal = 0
32790         .cmdClose.SetFocus
32800       End With
32810     Case vbKeyUp
32820       With Me
32830         intRetVal = 0
32840         .cmdExcelAll.SetFocus
32850       End With
32860     Case vbKeyDown
32870       With Me
32880         intRetVal = 0
32890         .cmdExcel00.SetFocus
32900       End With
32910     End Select
32920   End If

        ' ** Shift keys.
32930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32940     Select Case intRetVal
          Case vbKeyTab
32950       With Me
32960         intRetVal = 0
32970         .cmdExcelAll.SetFocus
32980       End With
32990     End Select
33000   End If

        ' ** Ctrl keys.
33010   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
33020     Select Case intRetVal
          Case vbKeyTab
33030       With Me
33040         intRetVal = 0
33050         If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
33060           .UserReportPath_chk.SetFocus
33070         Else
33080           Beep
33090         End If
33100       End With
33110     End Select
33120   End If

EXITP:
33130   KeyCode = intRetVal
33140   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

33210   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
33220   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub UserReportPath_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

33300 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_box_MouseMove"

33310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
33320   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub UserReportPath_chk_AfterUpdate()

33400 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_AfterUpdate"

33410   UserPath_After_CA Me  ' ** Module Procedure: modCourtReportsCA.

EXITP:
33420   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub UserReportPath_chk_KeyDown(KeyCode As Integer, Shift As Integer)

33500 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
33520   intShiftDown = (Shift And acShiftMask) > 0
33530   intAltDown = (Shift And acAltMask) > 0
33540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33570       With Me
33580         intRetVal = 0
33590         .cmdClose.SetFocus
33600       End With
33610     End Select
33620   End If

        ' ** Shift keys.
33630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33640     Select Case intRetVal
          Case vbKeyTab
33650       With Me
33660         intRetVal = 0
33670         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
33680           .cmdExcelAll.SetFocus
33690         Else
33700           .cmdWordAll.SetFocus
33710         End If
33720       End With
33730     End Select
33740   End If

        ' ** Ctrl-Shift keys.
33750   If intCtrlDown And (Not intAltDown) And intShiftDown Then
33760     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
33770       With Me
33780         intRetVal = 0
33790         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
33800           .chkOpenExcel.SetFocus
33810         Else
33820           .chkOpenWord.SetFocus
33830         End If
33840       End With
33850     End Select
33860   End If

EXITP:
33870   KeyCode = intRetVal
33880   Exit Sub

ERRH:
1490    intRetVal = 0
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Sub

Private Sub UserReportPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

33900 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_MouseMove"

33910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
33920   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Public Sub PreviewOrPrint(strReportNumber As String, strProc As String, intView As Integer)

34000 On Error GoTo ERRH

        Const THIS_PROC As String = "PreviewOrPrint"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strDocName As String
        Dim blnRetVal As Boolean, blnContinue As Boolean
        Dim intRetVal_BuildCourtReportData As Integer, intRetVal_BuildAssetListInfo As Integer

34010   blnContinue = True
34020   intRetVal_BuildCourtReportData = 0: intRetVal_BuildAssetListInfo = 0

34030   DoCmd.Hourglass True
34040   DoEvents

34050   With Me

          ' ** Validate the dates.
34060     If Validate = True Then  ' ** Function: Below.

34070       ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

            ' ** Set global variables for report headers.
34080       gdatStartDate = .DateStart
34090       gdatEndDate = .DateEnd
34100       gstrAccountNo = .cmbAccounts.Column(CBX_A_ACTNO)
34110       gstrAccountName = .cmbAccounts.Column(CBX_A_SHORT)

34120       Select Case .chkGroupBy_IncExpCode
            Case True
34130         gblnUseReveuneExpenseCodes = True
34140       Case False
34150         gblnUseReveuneExpenseCodes = False
34160       End Select

            ' ** When called from the TAReports CommandBar, these variables should already be filled.
            ' **   gstrCrtRpt_Ordinal
            ' **   gstrCrtRpt_Version
            ' **   gstrCrtRpt_CashAssets_Beg
            ' **   gstrCrtRpt_NetIncome
            ' **   gstrCrtRpt_NetLoss
            ' **   gstrCrtRpt_CashAssets_End

34170       If strReportNumber <> "7" Then
              ' ** Build a new summary report table for everything
              ' ** except report 7, the property on hand report.
              ' **** ARCHIVE ****
34180         intRetVal_BuildCourtReportData = CABuildCourtReportData(strReportNumber, strProc, .chkIncludeArchive, .chkIncludeCheckNum)  ' ** Module Function: modCourtReportsCA.
              ' *****************

34190         If intRetVal_BuildCourtReportData = 0 Then
34200           blnRebuildTable = False
                ' ** Summary.
34210           If strReportNumber = "0" Or strReportNumber = "0A" Then
34220             intRetVal_BuildAssetListInfo = BuildAssetListInfo(Me.DateStart, .DateEnd, "Ending", strProc)  ' ** Function: Below.
34230             If intRetVal_BuildAssetListInfo = 0 Then
                    ' ** Property On Hand = Sum([TotalCost])+IIf(IsNull([icash]),0,[icash])+[pcash].
34240               Set dbs = CurrentDb

                    ' ** Property on Hand at End of Account Period.
34250               Set qdf = dbs.QueryDefs("qryCourtReport_CA_07g")  ' ** tmpAssetList2, summed.
34260               Set rst = qdf.OpenRecordset
34270               If rst.BOF = True And rst.EOF = True Then
                      ' ** In the absense of newer data to roll back, take the total cost out of tmpCourtReportData.
34280                 rst.Close
                      ' ** tmpCourtReportData, summed with PropertyOnHand.
34290                 Set qdf = dbs.QueryDefs("qryCourtReport_CA_07e")
34300                 With qdf.Parameters
34310                   ![actno] = Me.cmbAccounts
34320                 End With
34330                 Set rst = qdf.OpenRecordset
34340                 rst.MoveFirst
34350                 gdblCrtRpt_CA_POHEnd = Nz(rst![PropertyOnHand], 0)
34360                 .PropOnHand_End = Nz(rst![PropertyOnHand], 0)
34370                 rst.Close
34380               Else
34390                 rst.MoveFirst
34400                 gdblCrtRpt_CA_POHEnd = Nz(rst![PropertyOnHand], 0)
34410                 .PropOnHand_End = Nz(rst![PropertyOnHand], 0)
34420                 rst.Close
34430               End If

                    ' ** For Property On Hand at Beginning of Account Period, we need 1 day prior to DateStart.
34440               intRetVal_BuildAssetListInfo = BuildAssetListInfo("01/01/1900", (.DateStart - 1), "Beginning", strProc)  ' ** Function: Below.
34450               If intRetVal_BuildAssetListInfo = 0 Then
                      ' ** The above function re-fills tmpAssetList2.
34460                 Set qdf = dbs.QueryDefs("qryCourtReport_CA_07g")  ' ** tmpAssetList2, summed.
34470                 Set rst = qdf.OpenRecordset
34480                 If rst.BOF = True And rst.EOF = True Then
                        ' ** In the absense of newer data to roll back, take the total cost out of tmpCourtReportData.
34490                   rst.Close
34500                   gdblCrtRpt_CA_POHBeg = 0@
34510                   .PropOnHand_Beg = 0@
34520                 Else
34530                   rst.MoveFirst
34540                   gdblCrtRpt_CA_POHBeg = Nz(rst![PropertyOnHand], 0)
34550                   .PropOnHand_Beg = Nz(rst![PropertyOnHand], 0)
34560                   rst.Close
34570                 End If
34580               ElseIf intRetVal_BuildAssetListInfo = -2 Then
                      ' ** If this is a new account, this may be True.
34590                 gdblCrtRpt_CA_POHBeg = 0@
34600                 .PropOnHand_Beg = 0@
34610               Else
34620                 blnContinue = False
                      ' ** Return codes:
                      ' **    0  Success.
                      ' **   -3  Missing entry, e.g., date.
                      ' **   -2  No data.
                      ' **   -9  Error.
34630               End If

34640               If blnContinue = True Then

                      ' ** Information for Investments Made total, report page 2.
34650                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
34660                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A2_02_archive")
                        ' *****************
34670                 Case False
34680                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A2_02")
34690                 End Select
34700                 Set rst = qdf.OpenRecordset
34710                 With rst
34720                   If .BOF = True And .EOF = True Then
                          ' ** Nothing to total.
34730                     gdblCrtRpt_CA_InvestInfo = 0@
34740                     Me.InvestInfo = 0&
34750                   Else
34760                     .MoveFirst
34770                     gdblCrtRpt_CA_InvestInfo = Nz(![Amount_Tot], 0)
34780                     Me.InvestInfo = Nz(![Amount_Tot], 0)
34790                   End If
34800                   .Close
34810                 End With

                      ' ** Change in Investment Holdings total, report page 2.
34820                 Select Case .chkIncludeArchive
                      Case True
                        ' **** ARCHIVE ****
34830                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A3_02_archive")
                        ' *****************
34840                 Case False
34850                   Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_A3_02")
34860                 End Select
34870                 Set rst = qdf.OpenRecordset
34880                 With rst
34890                   If .BOF = True And .EOF = True Then
                          ' ** Nothing to total.
34900                     gdblCrtRpt_CA_InvestChange = 0@
34910                     Me.InvestChange = 0&
34920                   Else
34930                     .MoveFirst
34940                     gdblCrtRpt_CA_InvestChange = Nz(![cost_tot], 0)
34950                     Me.InvestChange = Nz(![cost_tot], 0)
34960                   End If
34970                   .Close
34980                 End With

34990               End If

35000               dbs.Close
35010             Else
35020               blnContinue = False
                    ' ** Return codes:
                    ' **    0  Success.
                    ' **   -3  Missing entry, e.g., date.
                    ' **   -2  No data.
                    ' **   -9  Error.
35030             End If  ' ** BuildAssetListInfo.
35040           End If  ' ** strReportNumber.
35050         Else
35060           blnAllCancel = True
35070           AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
35080           blnContinue = False
35090         End If  ' ** CABuildCourtReportData(): intRetVal_BuildCourtReportData.
35100       Else
              ' ** Property on Hand at Close of Account Period.
35110         CACourtReportLoad  ' ** Module Function: modCourReportsCA.
35120         intRetVal_BuildAssetListInfo = BuildAssetListInfo(Me.DateStart, .DateEnd, "Ending", strProc)  ' ** Function: Below.
35130         If intRetVal_BuildAssetListInfo = 0 Then
35140           Set dbs = CurrentDb
35150           Set qdf = dbs.QueryDefs("qryCourtReport_CA_07c")  ' ** Uses tblCourtReportData.
35160           Set rst = qdf.OpenRecordset
35170           If rst.BOF = True And rst.EOF = True Then
                  ' ** No roll backs were required, current end date.
35180             strReportNumber = "7A"
35190           Else
                  ' ** Non-current end date
35200           End If
35210           rst.Close
35220           dbs.Close
35230         Else
35240           blnContinue = False
                ' ** Return codes:
                ' **    0  Success.
                ' **   -3  Missing entry, e.g., date.
                ' **   -2  No data.
                ' **   -9  Error.
35250         End If
35260       End If

35270       If blnContinue = True Then

              ' ** When moving this to print before Summary, save Summary's tmpAssetList2, then restore it.
35280         If ((Left(strReportNumber, 1) = "0" And .chkAssetList = True) Or (strReportNumber = "0B")) Then
35290           Set dbs = CurrentDb
                ' ** Empty tmpAssetList3.
35300           Set qdf = dbs.QueryDefs("qryCourtReport_CA_07ca")
35310           qdf.Execute
                ' ** Append tmpAssetList2 to tmpAssetList3.
35320           Set qdf = dbs.QueryDefs("qryCourtReport_CA_07cb")
35330           qdf.Execute
                ' ** Asset List (Property on Hand at Beginning of Account).
35340           .chkAssetList_Start = True
35350           intRetVal_BuildAssetListInfo = BuildAssetListInfo(#1/1/1900#, (CDate(Me.DateStart) - 1), "Beginning", strProc)  ' ** Function: Below.
35360           If intRetVal_BuildAssetListInfo = 0 Then
35370             strDocName = "rptCourtRptCA_00B"
35380             Select Case intView
                  Case acViewNormal
35390               If gblnCrtRpt_Zero = False Then
35400                 DoCmd.OpenReport strDocName, intView
35410               Else
35420                 If gblnCrtRpt_ZeroDialog = False Then
35430                   DoCmd.OpenReport strDocName, intView
35440                 Else
35450                   DoCmd.OpenReport strDocName, intView
                        ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                        ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                        'DoCmd.RunCommand acCmdPrint
35460                 End If
35470               End If
35480             Case acViewPreview
35490               DoCmd.OpenReport strDocName, intView
35500               DoCmd.Maximize
35510               DoCmd.RunCommand acCmdFitToWindow
35520             End Select
35530           Else
35540             blnContinue = False
                  ' ** Return codes:
                  ' **    0  Success.
                  ' **   -3  Missing entry, e.g., date.
                  ' **   -2  No data.
                  ' **   -9  Error.
35550             If intRetVal_BuildAssetListInfo = -2 Then
                    'ALL RPTS EXCEPT '0B' GET HANDLED WITH 'NO DATA' RPTS!
                    'WE NEED SOMETHING FOR '0B'!
35560               DoCmd.Hourglass False
35570               MsgBox "There is no data for the report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
35580             End If
35590           End If
35600           .chkAssetList_Start = False
                ' ** Empty tmpAssetList2.
35610           Set qdf = dbs.QueryDefs("qryCourtReport_03")
35620           qdf.Execute
35630           If strReportNumber <> "0B" Then  ' ** Skip if not printing Summary.
                  ' ** Append tmpAssetList3 to tmpAssetList2.
35640             Set qdf = dbs.QueryDefs("qryCourtReport_CA_07cc")
35650             qdf.Execute
35660           End If  ' ** 0B.
35670           dbs.Close
35680         End If

35690         If strReportNumber <> "0B" Then  ' ** Skip if not printing Summary.
                ' ** Now print the report.
35700           If Right(strReportNumber, 1) = "A" Or Right(strReportNumber, 1) = "B" Then
35710             strDocName = "rptCourtRptCA_" & Right("00" & strReportNumber, 3)
35720           Else
35730             If strReportNumber = "1" Then  ' ** (VGC 08/26/2010)
35740               strDocName = "rptCourtRptCA_01"
35750             Else
35760               strDocName = "rptCourtRptCA_" & Right("00" & strReportNumber, 2)
35770             End If
35780           End If
35790           If .chkIncludeCheckNum = True Then
35800             If strDocName = "rptCourtRptCA_04" Then
35810               strDocName = "rptCourtRptCA_04H"
35820             ElseIf strDocName = "rptCourtRptCA_04A" Then
35830               strDocName = "rptCourtRptCA_04AH"
35840             ElseIf strDocName = "rptCourtRptCA_06" Then
35850               strDocName = "rptCourtRptCA_06H"
35860             End If
35870           End If
35880           Select Case intView
                Case acViewNormal
35890             If gblnCrtRpt_Zero = False Then
35900               DoCmd.OpenReport strDocName, intView
35910             Else
35920               If gblnCrtRpt_ZeroDialog = False Then
35930                 DoCmd.OpenReport strDocName, intView
35940               Else
35950                 DoCmd.OpenReport strDocName, intView
                      ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                      ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                      'DoCmd.RunCommand acCmdPrint
35960               End If
35970             End If
35980           Case acViewPreview
35990             DoCmd.OpenReport strDocName, intView
36000             DoCmd.Maximize
36010             DoCmd.RunCommand acCmdFitToWindow
36020           End Select
36030         End If  ' ** 0B.

36040       End If  ' ** blnContinue.

36050       If strReportNumber <> "0B" Then  ' ** Skip if not printing Summary.
36060         If blnContinue = False And intRetVal_BuildAssetListInfo = -2 Then
36070           If Right(strReportNumber, 1) = "A" Or Right(strReportNumber, 1) = "B" Then
36080             strDocName = "rptCourtRptCA_" & Right("00" & strReportNumber, 3)
36090           Else
36100             strDocName = "rptCourtRptCA_" & Right("00" & strReportNumber, 2)
36110           End If

36120           If .chkAssetList = True And strDocName = "rptCourtRptCA_00" Or strDocName = "rptCourtRptCA_00A" Then
36130             Set dbs = CurrentDb
                  ' ** Empty tmpAssetList2.
36140             Set qdf = dbs.QueryDefs("qryCourtReport_03")
36150             qdf.Execute
36160             dbs.Close
36170             Select Case intView
                  Case acViewNormal
36180               If gblnCrtRpt_Zero = False Then
36190                 DoCmd.OpenReport "rptCourtRptCA_00B", intView
36200               Else
36210                 If gblnCrtRpt_ZeroDialog = False Then
36220                   DoCmd.OpenReport "rptCourtRptCA_00B", intView
36230                 Else
36240                   DoCmd.OpenReport "rptCourtRptCA_00B", intView
                        ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                        ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                        'DoCmd.RunCommand acCmdPrint
36250                 End If
36260               End If
36270             Case acViewPreview
36280               DoCmd.OpenReport "rptCourtRptCA_00B", intView
36290               DoCmd.Maximize
36300               DoCmd.RunCommand acCmdFitToWindow
36310             End Select
36320           End If
                ' ** Now print the report.
36330           Select Case intView
                Case acViewNormal
36340             If gblnCrtRpt_Zero = False Then
36350               DoCmd.OpenReport strDocName, intView
36360             Else
36370               If gblnCrtRpt_ZeroDialog = False Then
36380                 DoCmd.OpenReport strDocName, intView
36390               Else
36400                 DoCmd.OpenReport strDocName, intView
                      ' ** Invoking 'RunCommand acCmdPrint' instead of 'OpenReport' brings up
                      ' ** the Print dialog box, allowing a choice of output, as well as Cancel.
                      'DoCmd.RunCommand acCmdPrint
36410               End If
36420             End If
36430           Case acViewPreview
36440             DoCmd.OpenReport strDocName, intView
36450             DoCmd.Maximize
36460             DoCmd.RunCommand acCmdFitToWindow
36470           End Select
36480         End If
36490       End If  ' ** 0B.

36500     Else
36510       blnAllCancel = True
36520       AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
36530     End If

36540   End With  ' ** Me.

36550   DoCmd.Hourglass False

EXITP:
        ' ** Check some Public variables in case there's been an error.
36560   blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
36570   blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
36580   Set rst = Nothing
36590   Set qdf = Nothing
36600   Set dbs = Nothing
36610   Exit Sub

ERRH:
1490    DoCmd.Hourglass False
1500    Select Case ERR.Number
        Case 2202  ' ** You must set a default printer before you design, print, or preview.
1510      MsgBox "There does not appear to be a default printer defined." & vbCrLf & vbCrLf & _
            "Please define a default printer before trying to print this report.", vbCritical + vbOKOnly, "Printer Not Found"
1520    Case 2501  ' ** The '|' action was Canceled.
1530      blnAllCancel = True
1540      AllCancelSet2_CA blnAllCancel  ' ** Module Procedure: modCourtReportsCA.
          ' ** Do nothing.
1550    Case Else
1560      THAT_PROC = THIS_PROC
1570      That_Erl = Erl
1580      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1590    End Select
1600    Resume EXITP

End Sub

Public Function Validate() As Boolean

36700 On Error GoTo ERRH

        Const THIS_PROC As String = "Validate"

        Dim blnRetVal As Boolean

36710   blnRetVal = False

        ' ** Validate main three data values.
36720   With Me
36730     If IsNull(.DateStart.Value) Then
36740       MsgBox "You must enter a From date to continue.", vbInformation + vbOKOnly, "Entry Required"
36750       .DateStart.SetFocus
36760     Else
36770       If IsNull(.DateEnd.Value) Then
36780         MsgBox "You must enter a To date to continue.", vbInformation + vbOKOnly, "Entry Required"
36790         .DateEnd.SetFocus
36800       Else
36810         If .DateStart.Value >= .DateEnd.Value Then
36820           MsgBox "You must enter a From date that is less than the To date to continue.", vbInformation + vbOKOnly, "Invalid Date"
36830           .DateEnd.SetFocus
36840         Else
36850           If IsNull(.cmbAccounts.Column(CBX_A_ACTNO)) Or .cmbAccounts.Column(CBX_A_ACTNO) = vbNullString Then
36860             MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
36870             .cmbAccounts.SetFocus
36880           Else
36890             blnRetVal = True
36900           End If
36910         End If
36920       End If
36930     End If
36940   End With

EXITP:
36950   Validate = blnRetVal
36960   Exit Function

ERRH:
1490    blnRetVal = False
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Function

Public Function BuildAssetListInfo(varStartDate As Variant, varEndDate As Variant, strWhen As String, strProc As String) As Integer
' ** Copied from frmStatementParameters to match rollback for Asset List.
' ** All code that referenced 'All Accounts' has been removed.
' ** Return codes:
' **    0  Success.
' **   -2  No data.
' **   -3  Missing entry, e.g., date.
' **   -4  Date criteria not met.
' **   -9  Error.

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "BuildAssetListInfo"

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
        Dim lngTmp01 As Long
        Dim intRetVal_SetDateSpecificSQL As Integer
        Dim intRetVal As Integer

        Static intBuildIter As Integer

37010   With Me

37020     intBuildIter = intBuildIter + 1
37030     intRetVal = 0  ' ** Unless proven otherwise.
37040     blnContinue = True

37050     If .cmbAccounts.Visible = True Then
37060       .cmbAccounts.SetFocus
37070       If .cmbAccounts.text = vbNullString Then
37080         MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
37090         intRetVal = -3  ' ** Missing entry.
37100       End If
            ' ** Return focus to the button that called this.
37110       If Right(strProc, 6) = "_Click" Then
              ' ** cmdPrint00_Click.
37120         .Controls(Left(strProc, (Len(strProc) - 6))).SetFocus
37130         DoEvents
37140       End If
37150     End If

37160     If intRetVal = 0 Then
37170       If IsNull(varEndDate) Then
37180         MsgBox "Must enter Period " & strWhen & " date to continue.", vbInformation + vbOKOnly, "Entry Required"
37190         .DateEnd.SetFocus
37200         intRetVal = -3  ' ** Missing entry.
37210       End If
37220     End If

37230     If intRetVal = 0 Then
            ' ** This code will update the qryMaxBalDates query to give
            ' ** us the balance numbers from the previous statement.
            ' **** ARCHIVE ****
37240       intRetVal_SetDateSpecificSQL = SetDateSpecificSQL(.cmbAccounts, "Statements", THIS_NAME, varStartDate, varEndDate, .chkIncludeArchive)    ' ** Module Function: modUtilities.
            ' *****************
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -4  Date criteria not met.
            ' **   -9  Error.
37250       If intRetVal_SetDateSpecificSQL <> 0 Then
37260         blnContinue = False
37270         intRetVal = -2  ' ** No data.
37280       End If
37290     End If

37300     Set dbs = CurrentDb

          ' ** Empty tmpAssetList2.
37310     If TableExists("tmpAssetList2") = True Then  ' ** Module Function: modUtilities.
37320       Set qdf = dbs.QueryDefs("qryCourtReport_03")
37330       qdf.Execute
37340     End If
          ' ** Empty tmpAccountInfo.
37350     If TableExists("tmpAccountInfo") = True Then  ' ** Module Function: modUtilities.
37360       Set qdf = dbs.QueryDefs("qryCourtReport_04")
37370       qdf.Execute
37380     End If

37390     dbs.Close

37400     If intRetVal = 0 And blnContinue = True Then

37410       Set dbs = CurrentDb

            ' ** qryCourtReport_CA_07z.
            ' ** VGC 11/25/2009: «='90',-1,1)» to «='90',1,1)».
37420       strSQL = "SELECT ActiveAssets.assetno, masterasset.description AS MasterAssetDescription, masterasset.due, " & _
              "masterasset.rate, Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, " & _
              "Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface]))*" & _
              "IIf([assettype].[assettype]='90',1,1) AS TotalShareface, account.accountno, account.shortname, " & _
              "account.legalname, assettype.assettype, assettype_description, " & _
              "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
              "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
              "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc, "
37430       strSQL = strSQL & "account.icash, account.pcash, masterasset.currentDate, " & _
              "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX, " & _
              "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX, " & _
              "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX, " & CoInfo & " "
37440       strSQL = strSQL & "FROM (masterasset LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) " & _
              "RIGHT JOIN (account LEFT JOIN ActiveAssets ON account.accountno = ActiveAssets.accountno) " & _
              "ON masterasset.assetno = ActiveAssets.assetno " & _
              "GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, " & _
              "account.accountno, account.shortname, account.legalname, assettype.assettype, assettype_description, " & _
              "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & "
37450       strSQL = strSQL & "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
              "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), " & _
              "account.icash, account.pcash, masterasset.currentDate, " & _
              "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), " & _
              "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]), " & _
              "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) " & _
              "HAVING (((account.accountno)='" & Me.cmbAccounts & "'));"
37460       strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.

37470       Set rst = dbs.OpenRecordset(strSQL)  ' ** This is what's used for appending to the tmpAssetList2 table and rstTempAssetList.

            ' ** Ledger records to roll back if after specified 'To' date. (VGC 08/24/2010: CHANGES!)
37480       Select Case .chkIncludeArchive
            Case True
              ' **** ARCHIVE ****
37490         Set qdf = dbs.QueryDefs("qryCourtReport_CA_07d_archive_03")
              ' *****************
37500       Case False
37510         Set qdf = dbs.QueryDefs("qryCourtReport_CA_07d")
37520       End Select
37530       With qdf.Parameters
37540         ![actno] = Me.cmbAccounts
37550         ![datEnd] = CDate(varEndDate)
37560       End With
37570       Set rstWork = qdf.OpenRecordset
37580       If rstWork.BOF = True And rstWork.EOF = True Then
              ' ** In the absense of newer data to roll back, take the total cost as it's normally computed for report 0.
37590         blnContinue = False  ' ** Leave intRetVal = 0 so that the procedure calling this continues.
37600         rstWork.Close
37610         If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
37620           dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
37630         End If
37640         rst.Close
37650         dbs.Close
37660       End If

37670       If blnContinue Then  ' ** There are transactions after the specified To date.

37680         Set rstAccount = dbs.OpenRecordset("SELECT account.* FROM account ORDER BY accountno", dbOpenSnapshot)
37690         Set qdf = dbs.QueryDefs("qryStatementParameters_AssetList_01")
37700         Set rstMasterAsset = qdf.OpenRecordset()

37710         If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
37720           dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
37730           Set rstTempAssetList = dbs.OpenRecordset("SELECT tmpAssetList2.* FROM tmpAssetList2", dbOpenDynaset)
37740         Else
                ' ** ERROR.
37750           intRetVal = -9  ' ** Error.
37760           MsgBox "Unable to create temporary table for reporting.", vbCritical + vbOKOnly, "Error"
37770           rstAccount.Close
37780           rstMasterAsset.Close
37790           rstWork.Close
37800           dbs.Close
37810         End If

37820         If intRetVal = 0 Then

                ' ** Create table to contain temporary account master records to track "global" account info.
37830           If MakeTempTable(rst, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
37840             dbs.Execute "INSERT INTO tmpAccountInfo " & strSQL  ' ** Copy data from query.
                  ' ** Now, make records here distinct for each account.
37850             dbs.Execute "UPDATE tmpAccountInfo SET assetno=NULL, MasterAssetDescription=NULL, " & _
                    "due=NULL, rate=NULL, TotalCost=NULL, TotalShareFace=Null,  " & _
                    "assettype=NULL, assettype_description=NULL, TotDesc=NULL,  " & _
                    "currentDate=NULL, MarketValueX=NULL, MarketValueCurrentX=NULL, " & _
                    "YieldX=NULL  " & _
                    "WHERE True"
37860             Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT DISTINCTROW tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenSnapshot)
                  ' ** Empty tmpAccountInfo.
37870             Set qdf = dbs.QueryDefs("qryCourtReport_04")
37880             qdf.Execute
37890             If rstTmpAccountInfo.RecordCount > 0 Then
37900               rstTmpAccountInfo.MoveLast
37910               rstTmpAccountInfo.MoveFirst
37920             End If
37930             If Not CopyToTempTable(rstTmpAccountInfo, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
                    ' ** ERROR.
37940               intRetVal = -9  ' ** Error.
37950               MsgBox "Unable to copy data to temporary table.", vbCritical + vbOKOnly, "Error"
37960             Else
37970               rstTmpAccountInfo.Close  ' ** Close old snapshot, then open as dynaset.
37980               Set rstTmpAccountInfo = Nothing
37990               Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenDynaset)
38000             End If
38010           Else
                  ' ** ERROR.
38020             intRetVal = -9  ' ** Error.
38030             MsgBox "Unable to create temporary account table for reporting.", vbCritical + vbOKOnly, "Error"
38040             rstAccount.Close
38050             rstMasterAsset.Close
38060             rstWork.Close
38070             rstTempAssetList.Close
38080             dbs.Close
38090           End If

38100         End If

38110         If intRetVal = 0 Then

38120           rstTmpAccountInfo.MoveLast
38130           rstTmpAccountInfo.MoveFirst

                ' ** Get total number of records in recordset of subsequent ledger entries.
38140           rstWork.MoveLast
38150           rstWork.MoveFirst

38160           lngTmp01 = 0&
38170           Do While Not rstWork.EOF  ' ** Move through each record, tracking changes.

38180             rstTempAssetList.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "' AND [assetno] = " & CStr(rstWork![assetno])
38190             If rstTempAssetList.NoMatch Then
                    ' ** We need a new record because this asset was not in the reporting query after subsequent changes.
38200               rstMasterAsset.FindFirst "[assetno] = " & CStr(rstWork![assetno])
38210               blnNoMAsset = rstMasterAsset.NoMatch
38220               If blnNoMAsset Then
38230                 Select Case Trim(rstWork![journaltype])
                      Case "Misc.", "Paid", "Received"
                        ' ** OK to continue.
38240                 Case Else
                        ' ** ERROR.
38250                   intRetVal = -9  ' ** Error.
38260                   MsgBox "Missing Master Asset record.", vbCritical + vbOKOnly, "Error"
38270                   rstAccount.Close
38280                   rstMasterAsset.Close
38290                   rstWork.Close
38300                   rstTempAssetList.Close
38310                   rstTmpAccountInfo.Close
38320                   dbs.Close
38330                 End Select
38340               Else
38350                 rstTempAssetList.AddNew
38360                 rstTempAssetList![accountno] = rstWork![accountno]
38370                 rstTempAssetList![assetno] = rstMasterAsset![assetno]
38380                 rstTempAssetList![MasterAssetDescription] = rstMasterAsset![description]
38390                 rstTempAssetList![assettype] = rstMasterAsset![assettype]
38400                 rstTempAssetList![assettype_description] = rstMasterAsset![assettype_description]
38410                 rstTempAssetList![due] = rstMasterAsset![due]
38420                 rstTempAssetList![rate] = rstMasterAsset![rate]
38430                 rstTempAssetList![totdesc] = CStr(rstMasterAsset![description]) & _
                        IIf(rstMasterAsset![rate] > 0, " " & Format(rstMasterAsset![rate], "0.000%"), "") & _
                        IIf(Not IsNull(rstMasterAsset![due]), "  Due " & Format(rstMasterAsset![due], "mm/dd/yyyy"), "")
38440                 rstTempAssetList![TotalCost] = 0
38450                 rstTempAssetList![TotalShareFace] = 0
38460                 rstTempAssetList![ICash] = 0
38470                 rstTempAssetList![PCash] = 0
                      ' ** LEAVE in editing mode.
38480               End If
38490             Else
38500               blnNoMAsset = False
38510               rstTempAssetList.Edit  ' ** Edit existing temp record.
38520             End If

38530             If intRetVal = 0 Then

38540               rstTmpAccountInfo.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"

38550               If rstTmpAccountInfo.NoMatch Then
                      ' ** We need a new account record.

38560                 rstAccount.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"
38570                 If rstAccount.NoMatch Then
                        ' ** ERROR.
38580                   intRetVal = -9  ' ** Error.
38590                   MsgBox "Missing Account master record.", vbCritical + vbOKOnly, "Error"
38600                   rstAccount.Close
38610                   rstMasterAsset.Close
38620                   rstWork.Close
38630                   rstTempAssetList.Close
38640                   rstTmpAccountInfo.Close
38650                   dbs.Close
38660                 Else
38670                   rstTmpAccountInfo.AddNew
                        ' ** Copy info from Account table.
38680                   rstTmpAccountInfo![accountno] = rstAccount![accountno]
38690                   rstTmpAccountInfo![shortname] = rstAccount![shortname]
38700                   rstTmpAccountInfo![legalname] = rstAccount![legalname]
38710                   rstTmpAccountInfo![ICash] = rstAccount![ICash]
38720                   rstTmpAccountInfo![PCash] = rstAccount![PCash]
                        ' ** LEAVE in editing mode.
38730                 End If
38740               Else
38750                 rstTmpAccountInfo.Edit  ' ** Edit existing temp record.
38760               End If

38770             End If  ' ** intRetVal.

38780             If intRetVal = 0 Then

38790               If (Not blnNoMAsset) And (Trim(rstWork![journaltype]) <> "Received") Then
                      ' ** Add/Subtract info from ASSET record. NOTE: These are SUBTRACTIONS
                      ' ** because the query returns the changes since the date.
                      'VGC 08/20/2010: CHANGES!
38800                 rstTempAssetList![TotalShareFace] = _
                        rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, _
                        rstWork![shareface] * IIf(rstTempAssetList![assettype] = 90, 1, 1)))
                      'rstTempAssetList![TotalShareFace] = _
                      '  rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, rstWork![shareface] * _
                      '  IIf(rstTempAssetList![assettype] = 90, -1, 1)))
38810                 rstTempAssetList![TotalCost] = rstTempAssetList![TotalCost] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost])
                      ' ** Save ASSET temp table record.
38820                 rstTempAssetList.Update
38830               End If

                    ' ** Add/Subtract info from ACCOUNT record. NOTE: These are SUBTRACTIONS
                    ' ** because the query returns the changes since the date.
38840               rstTmpAccountInfo![ICash] = rstTmpAccountInfo![ICash] - rstWork![ICash]
38850               rstTmpAccountInfo![PCash] = rstTmpAccountInfo![PCash] - rstWork![PCash]
                    ' ** Save ACCOUNT temp table record.
38860               rstTmpAccountInfo.Update

38870               rstWork.MoveNext

38880             End If  ' ** intRetVal.

38890             If intRetVal <> 0 Then Exit Do

38900           Loop  ' ** Move through each record, tracking changes.

38910         End If  ' ** intRetVal.

38920         If intRetVal = 0 Then

                ' ** Remove any asset records which now have a ZERO totalshareface AND totalcost.
38930           strWorkSQL = "DELETE tmpAssetList2.*, tmpAssetList2.TotalShareface, tmpAssetList2.TotalCost " & _
                  "FROM tmpAssetList2 " & _
                  "WHERE ((Round((tmpAssetList2.TotalShareface),4)=0) AND (Round((tmpAssetList2.TotalCost),2)=0));"
                ' ** 07/21/2008: Added rounding.
38940           dbs.Execute strWorkSQL

                ' ** Now, update tmpAssetList2 to reflect the needed account information which is
                ' ** global to the report.
38950           rstTmpAccountInfo.MoveFirst
38960           While Not rstTmpAccountInfo.EOF
38970             strWorkSQL = "UPDATE tmpAssetList2 SET " & _
                    "shortname = " & SQLFormatStr(rstTmpAccountInfo![shortname], dbText) & ", " & _
                    "legalname = " & SQLFormatStr(rstTmpAccountInfo![legalname], dbText) & ", " & _
                    "icash = " & SQLFormatStr(rstTmpAccountInfo![ICash], dbCurrency) & ", " & _
                    "pcash = " & SQLFormatStr(rstTmpAccountInfo![PCash], dbCurrency) & " " & _
                    "WHERE accountno = " & SQLFormatStr(rstTmpAccountInfo![accountno], dbText) & ";"
38980             dbs.Execute strWorkSQL
38990             rstTmpAccountInfo.MoveNext
39000           Wend

                ' ** Close temp & working recordsets.
39010           rstAccount.Close
39020           rstMasterAsset.Close
39030           rstWork.Close
39040           rstTempAssetList.Close
39050           rstTmpAccountInfo.Close

                ' ** Finally, base report on this instead of on qryAssetList.
39060           Set qdf = dbs.QueryDefs("qryCourtReport_CA_00_Ba")
39070           Set rst = qdf.OpenRecordset()

39080         End If  ' ** intRetVal.

39090       End If  ' ** blnContinue.

39100       If intRetVal = 0 And blnContinue = True Then
39110         If rst.BOF = True And rst.EOF = True Then
39120           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
39130           intRetVal = -2  ' ** No data.
39140         End If
39150         rst.Close
39160         dbs.Close
39170       End If

39180     End If  ' ** intRetVal.

39190   End With

EXITP:
39200   Set rst = Nothing
39210   Set dbs = Nothing
39220   BuildAssetListInfo = intRetVal
39230   Exit Function

ERRH:
1490    intRetVal = -9
1500    THAT_PROC = THIS_PROC
1510    That_Erl = Erl
1520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1530    Resume EXITP

End Function

Public Sub AllCancelSet1_CA(blnCancel As Boolean)

39300 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet1_CA"

39310   blnAllCancel = blnCancel

EXITP:
39320   Exit Sub

ERRH:
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub
