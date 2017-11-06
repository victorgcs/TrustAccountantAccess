VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_NS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_NS"

'VGC 10/09/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  39820
' #########################

' ** For GTR, sets options, then stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Excel All:        E {cmdExcelAll}
' **   From:             F {DateStart}
' **   Legal Name:       G {chkLegalName}
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

Private blnGoingToReport As Boolean, blnWasGTR As Boolean
Private blnAllCancel As Boolean, blnRebuildTable As Boolean
Private strLastArrow As String, strFootnoteDefault As String
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
160           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False

170           blnWasGTR = gblnGoToReport

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           CLR_DISABLED_FG = CLR_DKGRY
200           CLR_DISABLED_BG = CLR_LTTEAL

210           blnPrintAll_Focus = False: blnWordAll_Focus = False: blnExcelAll_Focus = False

220           gblnLegalName = False
230           gblnUseReveuneExpenseCodes = False
240           gvarCrtRpt_FL_SpecData = Null
250           strFootnoteDefault = .CourtReports_Footnote.DefaultValue
260           If Right(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Left(strFootnoteDefault, (Len(strFootnoteDefault) - 1))
270           If Left(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Mid(strFootnoteDefault, 2)

280           gstrFormQuerySpec = THIS_NAME
290           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
300           gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
310           gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString

320           blnRebuildTable = True
330           blnAllCancel = False
340           AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

350           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              ' ** If revenue / expense tracking is checked on the options screen,
              ' ** let user select reports to be grouped by revenue/expense codes.
360           Select Case gblnRevenueExpenseTracking
              Case True
370             .chkGroupBy_IncExpCode.Enabled = True
380             .chkGroupBy_IncExpCode_lbl3.Visible = False     ' ** Off note.
390             .chkGroupBy_IncExpCode_lbl4_txt.Locked = True   ' ** Asterisk
400             .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
410             .cmdPreview00_lbl4_txt.Locked = True
420             .cmdPreview00_lbl4_cmd.Enabled = True
430             .cmdPreview01_lbl4_txt.Locked = True
440             .cmdPreview01_lbl4_cmd.Enabled = True
450             .cmdPreview03_lbl4_txt.Locked = True
460             .cmdPreview03_lbl4_cmd.Enabled = True
470             .cmdPreview07_lbl4_txt.Locked = True
480             .cmdPreview07_lbl4_cmd.Enabled = True
490             .cmdPreview08_lbl4_txt.Locked = True
500             .cmdPreview08_lbl4_cmd.Enabled = True
510           Case False
520             .chkGroupBy_IncExpCode.Enabled = False
530             .chkGroupBy_IncExpCode_lbl3.Visible = True      ' ** Off note.
540             .chkGroupBy_IncExpCode_lbl4_txt.Locked = False  ' ** Asterisk
550             .chkGroupBy_IncExpCode_lbl5.Visible = False     ' ** Asterisk Shadow.
560             .cmdPreview00_lbl4_txt.Locked = False  ' ** So they'll look disabled.
570             .cmdPreview00_lbl4_cmd.Enabled = False
580             .cmdPreview01_lbl4_txt.Locked = False
590             .cmdPreview01_lbl4_cmd.Enabled = False
600             .cmdPreview03_lbl4_txt.Locked = False
610             .cmdPreview03_lbl4_cmd.Enabled = False
620             .cmdPreview07_lbl4_txt.Locked = False
630             .cmdPreview07_lbl4_cmd.Enabled = False
640             .cmdPreview08_lbl4_txt.Locked = False
650             .cmdPreview08_lbl4_cmd.Enabled = False
660           End Select

670           If gdatStartDate > 0 And gdatEndDate > 0 Then
680             .DateStart = gdatStartDate
690             .DateEnd = gdatEndDate
700           End If

710           NSCourtReportLoad  ' ** Module Procedure: modCourtReportsNS.

      #If NoExcel Then
720           .cmdExcel00.Enabled = False
730           .cmdExcel00B.Enabled = False
740           .cmdExcel01.Enabled = False
750           .cmdExcel02.Enabled = False
760           .cmdExcel03.Enabled = False
770           .cmdExcel04.Enabled = False
780           .cmdExcel05.Enabled = False
790           .cmdExcel06.Enabled = False
800           .cmdExcel07.Enabled = False
810           .cmdExcel08.Enabled = False
820           .cmdExcel09.Enabled = False
830           .cmdExcel10.Enabled = False
840           .cmdExcelAll.Enabled = False
850           .chkOpenExcel.Enabled = False
860           .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
870           .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
880           .cmdExcel00.Enabled = True
890           .cmdExcel00B.Enabled = True
900           .cmdExcel01.Enabled = True
910           .cmdExcel02.Enabled = True
920           .cmdExcel03.Enabled = True
930           .cmdExcel04.Enabled = True
940           .cmdExcel05.Enabled = True
950           .cmdExcel06.Enabled = True
960           .cmdExcel07.Enabled = True
970           .cmdExcel08.Enabled = True
980           .cmdExcel09.Enabled = True
990           .cmdExcel10.Enabled = True
1000          .cmdExcelAll.Enabled = True
1010          .chkOpenExcel.Enabled = True
1020          .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
1030          .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

              ' ** EVENT CHECK: chkRememberMe!
1040          cmbAccounts_AfterUpdate
              ' ** EVENT CHECK: chkRememberDates!
1050          If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
1060            DateStart_Exit Cancel  ' ** Procedure: Below.
1070            DateEnd_Exit Cancel  ' ** Procedure: Below.
1080          End If

1090          UserReportPath_chk_AfterUpdate  ' ** Procedure: Below.
1100          SetUserReportPath_NS Me  ' ** Module Procedure: modCourtReportsNS.
1110          chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Below.
1120          chkLegalName_AfterUpdate  ' ** Procedure: Below.
1130          opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1140          chkAssetList_AfterUpdate  ' ** Procedure: Below.
1150          chkRememberMe_AfterUpdate  ' ** Procedure: Below.
1160          chkRememberDates_AfterUpdate  ' ** Procedure: Below.
1170          chkPageOf_AfterUpdate  ' ** Procedure: Below.

1180          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
1190          chkOpenExcel_AfterUpdate  ' ** Procedure: Below.

1200          chkIncludeFootnote_AfterUpdate  ' ** Procedure: Below.
1210          If IsNull(.CourtReports_Footnote) = True Then
1220            .CourtReports_Footnote = strFootnoteDefault
1230          End If

1240          DoCmd.Hourglass False

1250          .DateStart.SetFocus

1260        Else
1270          Cancel = -1
1280        End If
1290      Else
1300        Cancel = -1
1310      End If
1320    End With

1330    If Cancel = -1 Then
1340      Beep
1350      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1360    End If

EXITP:
1370    Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub Form_Load()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1410    With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
1420      Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
1430      clsMonthClass.hWndForm = .hwnd
          ' ** Let's default to PositionAtCursor.
1440      clsMonthClass.PositionAtCursor = True

1450    End With

EXITP:
1460    Exit Sub

ERRH:
100     DoCmd.Hourglass False
110     Select Case ERR.Number
        Case Else
120       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
130     End Select
140     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim strCtl As String, strNum As String
        Dim blnTmp01 As Boolean
        Dim intRetVal As Integer

1510    intRetVal = KeyCode
1520    blnTmp01 = False

        ' ** Use bit masks to determine which key was pressed.
1530    intShiftDown = (Shift And acShiftMask) > 0
1540    intAltDown = (Shift And acAltMask) > 0
1550    intCtrlDown = (Shift And acCtrlMask) > 0

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

1560  On Error Resume Next
1570    strCtl = Screen.ActiveControl.Name
1580  On Error GoTo ERRH

        ' ** Plain keys.
1590    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1600      Select Case intRetVal
          Case vbKeyF7
1610        intRetVal = 0
1620        cmdCalendar1_Click  ' ** Procedure: Below.
1630      Case vbKeyF8
1640        intRetVal = 0
1650        cmdCalendar2_Click  ' ** Procedure: Below.
1660      Case vbKeyUp
1670        With Me
1680          If strCtl <> vbNullString Then
1690            If Right(strCtl, 1) = "B" Then blnTmp01 = True
1700            strNum = Right(IIf(blnTmp01 = True, Left(strCtl, (Len(strCtl) - 1)), strCtl), 2)
1710            If IsNumeric(strNum) = True Then
                  ' ** It's one of the main report buttons.
1720              intRetVal = 0
1730              If strNum <> "00" Or blnTmp01 = True Then
1740                strLastArrow = vbNullString
1750                Select Case strNum
                    Case "00"
1760                  strCtl = Left(strCtl, (Len(strCtl) - 1))
1770                Case "01"
1780                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr(Val("00")), 2) & "B"
1790                Case Else
1800                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) - 1)), 2)
1810                End Select
1820              Else
1830                strLastArrow = strCtl
1840                strCtl = Mid(strCtl, 4, 4)
1850                Select Case strCtl
                    Case "Prev"
1860                  strCtl = "cmdPreview10"
1870                Case "Prin"
1880                  strCtl = "cmdPrintAll"
1890                Case "Word"
1900                  strCtl = "cmdWordAll"
1910                Case "Exce"
1920                  strCtl = "cmdExcelAll"
1930                End Select
1940              End If
1950              .Controls(strCtl).SetFocus
1960            End If
1970          End If
1980        End With
1990      Case vbKeyDown
2000        With Me
2010          If strCtl <> vbNullString Then
2020            If Right(strCtl, 1) = "B" Then blnTmp01 = True
2030            strNum = Right(IIf(blnTmp01 = True, Left(strCtl, (Len(strCtl) - 1)), strCtl), 2)
2040            If IsNumeric(strNum) = True Then
                  ' ** It's one of the main report buttons.
2050              intRetVal = 0
2060              If strNum <> "10" Then
2070                strLastArrow = vbNullString
2080                Select Case strNum
                    Case "00"
                      ' ** cmdPreview00B, cmdPrint00B, cmdWord00B, cmdExcel00B.
2090                  Select Case blnTmp01
                      Case True
2100                    strCtl = Left(strCtl, (Len(strCtl) - 3)) & Right("00" & CStr(Val("01")), 2)
2110                  Case False
2120                    strCtl = strCtl & "B"
2130                  End Select
2140                Case Else
2150                  strCtl = Left(strCtl, (Len(strCtl) - 2)) & Right("00" & CStr((Val(strNum) + 1)), 2)
2160                End Select
2170              Else
2180                strLastArrow = strCtl
2190                strCtl = Mid(strCtl, 4, 4)

2200                Select Case strCtl
                    Case "Prev"
2210                  strCtl = "cmdPreview00"
2220                Case "Prin"
2230                  strCtl = "cmdPrintAll"
2240                Case "Word"
2250                  strCtl = "cmdWordAll"
2260                Case "Exce"
2270                  strCtl = "cmdExcelAll"
2280                End Select
2290              End If
2300              .Controls(strCtl).SetFocus
2310            End If
2320          End If
2330        End With
2340      Case vbKeyEscape
2350        intRetVal = 0
2360        cmdClose_Click  ' ** Procedure: Below.
2370      End Select
2380    End If

        ' ** Alt keys.
2390    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
2400      Select Case intRetVal
          Case vbKeyS
2410        With Me
2420          intRetVal = 0
2430          .chkAssetList.SetFocus
2440          .chkAssetList = (Not .chkAssetList)
2450          chkAssetList_AfterUpdate  ' ** Procedure: Below.
2460        End With
2470      Case vbKeyX
2480        intRetVal = 0
2490        cmdClose_Click  ' ** Procedure: Below.
2500      End Select
2510    End If

        ' ** Ctrl keys.
2520    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2530      Select Case intRetVal
          Case vbKeyG
2540        With Me
2550          intRetVal = 0
2560          If .chkPageOf.Enabled = True Then
2570            .chkPageOf.SetFocus
2580            .chkPageOf = (Not .chkPageOf)
2590            chkPageOf_AfterUpdate  ' ** Procedure: Below.
2600          End If
2610        End With
2620      Case vbKeyN
2630        With Me
2640          intRetVal = 0
2650          If .chkOpenExcel.Enabled = True Then
2660            .chkOpenExcel.SetFocus
2670            .chkOpenExcel = (Not .chkOpenExcel)
2680            chkOpenExcel_AfterUpdate  ' ** Procedure: Below.
2690          Else
2700            Beep
2710          End If
2720        End With
2730      Case vbKeyO
2740        With Me
2750          intRetVal = 0
2760          .chkOpenWord.SetFocus
2770          .chkOpenWord = (Not .chkOpenWord)
2780          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
2790        End With
2800      End Select
2810    End If

        ' ** Ctrl-Shift keys.
2820    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2830      Select Case intRetVal
          Case vbKeyT
2840        intRetVal = 0
2850        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
2860      Case vbKeyF
2870        With Me
2880          intRetVal = 0
2890          .FocusHolder.SetFocus
2900        End With
2910      End Select
2920    End If

EXITP:
2930    KeyCode = intRetVal
2940    Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1420    Resume EXITP

End Sub

Private Sub Form_Timer()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngBtn_Top As Long, lngBtn_Left As Long

3010    With Me
3020      .TimerInterval = 0&
3030      If gblnSetFocus = True Then
3040        gblnSetFocus = False
3050        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
3060          Forms("frmLoadTimer").Modal = False
3070        End If
3080        strFormName = vbNullString
3090  On Error Resume Next
3100        Do Until strFormName = THIS_NAME
3110          DoCmd.SelectObject acForm, THIS_NAME, False
3120          strFormName = Screen.ActiveForm.Name
3130        Loop
3140  On Error GoTo ERRH
3150      End If
3160      If gblnGoToReport = True Then
3170        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Procedure: modWindowFunctions.
3180          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
3190        End If
3200        Select Case blnGoingToReport
            Case True
3210          DoCmd.Hourglass False  ' ** It should already be off.
3220          blnGoingToReport = False
3230          gblnGoToReport = False
3240          .GoToReport_arw_rpt_img.Visible = False
3250          Select Case garr_varGoToReport(GTR_CTL6)
              Case "cmdPrint00"
3260            .cmdPreview00_lbl4_txt.Visible = True
3270            Select Case .chkGroupBy_IncExpCode
                Case True
3280              .cmdPreview00_lbl5.Visible = True
3290            Case False
3300              .cmdPreview00_lbl5.Visible = False
3310            End Select
3320          Case "cmdPrint00B"
3330            .GoToReport_arw_alist_on_img.Visible = False
3340            .GoToReport_arw_alist_off_img.Visible = False
3350            .chkAssetList.Visible = True
3360            .chkAssetList_lbl2.Visible = True
3370            .chkAssetList_box.Visible = True
3380          Case "cmdPrint01"
3390            .cmdPreview01_lbl4_txt.Visible = True
3400            Select Case .chkGroupBy_IncExpCode
                Case True
3410              .cmdPreview01_lbl5.Visible = True
3420            Case False
3430              .cmdPreview01_lbl5.Visible = False
3440            End Select
3450          Case "cmdPrint03"
3460            .cmdPreview03_lbl4_txt.Visible = True
3470            Select Case .chkGroupBy_IncExpCode
                Case True
3480              .cmdPreview03_lbl5.Visible = True
3490            Case False
3500              .cmdPreview03_lbl5.Visible = False
3510            End Select
3520          Case "cmdPrint07"
3530            .cmdPreview07_lbl4_txt.Visible = True
3540            Select Case .chkGroupBy_IncExpCode
                Case True
3550              .cmdPreview07_lbl5.Visible = True
3560            Case False
3570              .cmdPreview07_lbl5.Visible = False
3580            End Select
3590          Case "cmdPrint08"
3600            .cmdPreview08_lbl4_txt.Visible = True
3610            Select Case .chkGroupBy_IncExpCode
                Case True
3620              .cmdPreview08_lbl5.Visible = True
3630            Case False
3640              .cmdPreview08_lbl5.Visible = False
3650            End Select
3660          End Select
3670          DoEvents
3680        Case False
3690          DoCmd.Hourglass True  ' ** Make sure it's still running.
3700          DoEvents
3710          blnGoingToReport = True
3720          Select Case garr_varGoToReport(GTR_CTL6)
              Case "cmdPrint00"
3730            lngBtn_Top = (.cmdPreview00_lbl.Top + (lngTpp))
3740            lngBtn_Left = (.cmdPreview00_lbl.Left + .cmdPreview00_lbl.Width)
3750            .cmdPreview00_lbl4_txt.Visible = False
3760            .cmdPreview00_lbl5.Visible = False
3770            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3780            DoEvents
3790          Case "cmdPrint00B"
3800            lngBtn_Top = (.cmdPreview00B_lbl.Top - (0& * lngTpp))
3810            lngBtn_Left = ((.cmdPreview00B_lbl.Left + .cmdPreview00B_lbl.Width) - (lngTpp))
3820            .chkAssetList.Visible = False
3830            .chkAssetList_lbl2.Visible = False
3840            .chkAssetList_box.Visible = False
3850            .GoToReport_arw_alist_on_img.Top = lngBtn_Top
3860            .GoToReport_arw_alist_off_img.Top = lngBtn_Top
3870            .GoToReport_arw_alist_on_img.Left = lngBtn_Left
3880            .GoToReport_arw_alist_off_img.Left = lngBtn_Left
3890            Select Case .chkAssetList
                Case True
3900              .GoToReport_arw_alist_on_img.Visible = True
3910            Case False
3920              .GoToReport_arw_alist_off_img.Visible = True
3930            End Select
3940            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3950            DoEvents
3960          Case "cmdPrint01"
3970            lngBtn_Top = (.cmdPreview01_lbl.Top + (lngTpp))
3980            lngBtn_Left = (.cmdPreview01_lbl.Left + .cmdPreview01_lbl.Width)
3990            .cmdPreview01_lbl4_txt.Visible = False
4000            .cmdPreview01_lbl5.Visible = False
4010            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4020            DoEvents
4030          Case "cmdPrint02"
4040            lngBtn_Top = (.cmdPreview02_lbl.Top + (lngTpp))
4050            lngBtn_Left = (.cmdPreview02_lbl.Left + .cmdPreview02_lbl.Width)
4060            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4070            DoEvents
4080          Case "cmdPrint03"
4090            lngBtn_Top = (.cmdPreview03_lbl.Top + (lngTpp))
4100            lngBtn_Left = (.cmdPreview03_lbl.Left + .cmdPreview03_lbl.Width)
4110            .cmdPreview03_lbl4_txt.Visible = False
4120            .cmdPreview03_lbl5.Visible = False
4130            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4140            DoEvents
4150          Case "cmdPrint04"
4160            lngBtn_Top = (.cmdPreview04_lbl.Top + (lngTpp))
4170            lngBtn_Left = (.cmdPreview04_lbl.Left + .cmdPreview04_lbl.Width)
4180            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4190            DoEvents
4200          Case "cmdPrint05"
4210            lngBtn_Top = (.cmdPreview05_lbl.Top + (lngTpp))
4220            lngBtn_Left = (.cmdPreview05_lbl.Left + .cmdPreview05_lbl.Width)
4230            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4240            DoEvents
4250          Case "cmdPrint06"
4260            lngBtn_Top = (.cmdPreview06_lbl.Top + (lngTpp))
4270            lngBtn_Left = (.cmdPreview06_lbl.Left + .cmdPreview06_lbl.Width)
4280            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4290            DoEvents
4300          Case "cmdPrint07"
4310            lngBtn_Top = (.cmdPreview07_lbl.Top + (lngTpp))
4320            lngBtn_Left = (.cmdPreview07_lbl.Left + .cmdPreview07_lbl.Width)
4330            .cmdPreview07_lbl4_txt.Visible = False
4340            .cmdPreview07_lbl5.Visible = False
4350            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4360            DoEvents
4370          Case "cmdPrint08"
4380            lngBtn_Top = (.cmdPreview08_lbl.Top + (lngTpp))
4390            lngBtn_Left = (.cmdPreview08_lbl.Left + .cmdPreview08_lbl.Width)
4400            .cmdPreview08_lbl4_txt.Visible = False
4410            .cmdPreview08_lbl5.Visible = False
4420            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4430            DoEvents
4440          Case "cmdPrint09"
4450            lngBtn_Top = (.cmdPreview09_lbl.Top + (lngTpp))
4460            lngBtn_Left = (.cmdPreview09_lbl.Left + .cmdPreview09_lbl.Width)
4470            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4480            DoEvents
4490          Case "cmdPrint10"
4500            lngBtn_Top = (.cmdPreview10_lbl.Top + (lngTpp))
4510            lngBtn_Left = (.cmdPreview10_lbl.Left + .cmdPreview10_lbl.Width)
4520            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
4530            DoEvents
4540          End Select
4550          If garr_varGoToReport(GTR_CTL6) <> "cmdPrint00B" Then
4560            .GoToReport_arw_rpt_img.Top = lngBtn_Top
4570            .GoToReport_arw_rpt_img.Left = lngBtn_Left
4580            .GoToReport_arw_rpt_img.Visible = True
4590          End If
4600          DoEvents
4610          Beep
4620          DoCmd.Hourglass False
4630          DoEvents
4640          .TimerInterval = (GTR_WAIT * 2&)
4650        End Select
4660      End If
4670      If gblnTimer = True Then
4680        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
4690          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
4700          gblnTimer = False
4710          DoCmd.Hourglass False
4720        End If
4730      End If
4740    End With

'HERE! 2  Form_Load()  frmRpt_CourtReports_NS
'HERE! 1  Form_Open()  frmRpt_CourtReports_NS
'HERE! 3  Form_Timer()  frmRpt_CourtReports_NS

EXITP:
4750    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

4810    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
4820      Response = acDataErrContinue
4830    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
4840      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
4850        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
4860        Response = acDataErrContinue
4870      End If
4880    Case 2237  ' ** The text you entered isn't an item in the list.
4890      Response = acDataErrDisplay
4900    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
4910      Beep
4920      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
4930      Response = acDataErrContinue
4940    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
4950    Case Else
4960      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
4970      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
4980    End Select

EXITP:
4990    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

5010    With Me
5020      If .chkRememberMe = False Then
5030        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
5040      End If
5050      If .chkRememberDates = False Then
5060        .DateStart = Null
5070        .DateEnd = Null
5080      End If
5090    End With

5100    FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
5110    DoEvents

5120    Select Case blnWasGTR
        Case True
5130      blnWasGTR = False
5140    Case False
          ' ** Don't save settings imposed by GoToReport!
5150      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
5160    End Select

5170    gblnLegalName = False

5180    ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
5190    ChkSpecLedgerEntryNull  ' ** Module Function: modUtilities.

        ' ** Make sure all reports are closed.
5200    Do While Reports.Count > 0
5210      DoCmd.Close acReport, Reports(0).Name
5220    Loop

5230    gstrFormQuerySpec = vbNullString
5240    gblnSetFocus = True
5250    gstrReturningForm = THIS_NAME
5260    Select Case strCallingForm
        Case "frmMenu_CourtReport"
5270      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Report"
5280    End Select

EXITP:
5290    Set qdf = Nothing
5300    Set dbs = Nothing
5310    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

5410    With Me
5420      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
5430        Select Case blnCalendar1_Focus
            Case True
5440          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
5450          .cmdCalendar1_raised_img.Visible = False
5460        Case False
5470          .cmdCalendar1_raised_img.Visible = True
5480          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5490        End Select
5500        .cmdCalendar1_raised_focus_dots_img.Visible = False
5510        .cmdCalendar1_raised_focus_img.Visible = False
5520        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5530        .cmdCalendar1_raised_img_dis.Visible = False
5540      End If
5550      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
5560        Select Case blnCalendar2_Focus
            Case True
5570          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
5580          .cmdCalendar2_raised_img.Visible = False
5590        Case False
5600          .cmdCalendar2_raised_img.Visible = True
5610          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
5620        End Select
5630        .cmdCalendar2_raised_focus_dots_img.Visible = False
5640        .cmdCalendar2_raised_focus_img.Visible = False
5650        .cmdCalendar2_sunken_focus_dots_img.Visible = False
5660        .cmdCalendar2_raised_img_dis.Visible = False
5670      End If
5680      If blnPrintAll_Focus = False And (.cmdPrintAll_box01.Visible = True Or .cmdPrintAll_box02.Visible = True Or _
              .cmdPrintAll_box03.Visible = True Or .cmdPrintAll_box01.Visible = True) Then
5690        .cmdPrintAll_box01.Visible = False
5700        .cmdPrintAll_box02.Visible = False
5710        .cmdPrintAll_box03.Visible = False
5720        .cmdPrintAll_box04.Visible = False
5730      End If
5740      If blnWordAll_Focus = False And (.cmdWordAll_box01.Visible = True Or .cmdWordAll_box02.Visible = True Or _
              .cmdWordAll_box03.Visible = True Or .cmdWordAll_box01.Visible = True) Then
5750        .cmdWordAll_box01.Visible = False
5760        .cmdWordAll_box02.Visible = False
5770        .cmdWordAll_box03.Visible = False
5780        .cmdWordAll_box04.Visible = False
5790      End If
5800      If blnExcelAll_Focus = False And (.cmdExcelAll_box01.Visible = True Or .cmdExcelAll_box02.Visible = True Or _
              .cmdExcelAll_box03.Visible = True Or .cmdExcelAll_box01.Visible = True) Then
5810        .cmdExcelAll_box01.Visible = False
5820        .cmdExcelAll_box02.Visible = False
5830        .cmdExcelAll_box03.Visible = False
5840        .cmdExcelAll_box04.Visible = False
5850      End If
5860    End With

EXITP:
5870    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

5910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5920    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdClose_Click()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

6010    DoCmd.Close acForm, THIS_NAME

EXITP:
6020    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6120    intShiftDown = (Shift And acShiftMask) > 0
6130    intAltDown = (Shift And acAltMask) > 0
6140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6160      Select Case intRetVal
          Case vbKeyTab
6170        With Me
6180          intRetVal = 0
6190          .DateStart.SetFocus
6200        End With
6210      Case vbKeyLeft
6220        With Me
6230          intRetVal = 0
6240          If .cmdExcelAll.Enabled = True Then
6250            .cmdExcelAll.SetFocus
6260          Else
6270            .cmdWordAll.SetFocus
6280          End If
6290        End With
6300      Case vbKeyRight
6310        With Me
6320          intRetVal = 0
6330          .DateStart.SetFocus
6340        End With
6350      End Select
6360    End If

        ' ** Shift keys.
6370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6380      Select Case intRetVal
          Case vbKeyTab
6390        With Me
6400          intRetVal = 0
6410          If .cmdExcelAll.Enabled = True Then
6420            .cmdExcelAll.SetFocus
6430          Else
6440            .cmdWordAll.SetFocus
6450          End If
6460        End With
6470      End Select
6480    End If

        ' ** Ctrl-Shift keys.
6490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6500      Select Case intRetVal
          Case vbKeyTab
6510        With Me
6520          intRetVal = 0
6530          If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
6540            .UserReportPath_chk.SetFocus
6550          End If
6560        End With
6570      End Select
6580    End If

EXITP:
6590    KeyCode = intRetVal
6600    Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6720    intShiftDown = (Shift And acShiftMask) > 0
6730    intAltDown = (Shift And acAltMask) > 0
6740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6770        With Me
6780          intRetVal = 0
6790          .DateEnd.SetFocus
6800        End With
6810      End Select
6820    End If

        ' ** Shift keys.
6830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6850        With Me
6860          intRetVal = 0
6870          .cmdClose.SetFocus
6880        End With
6890      End Select
6900    End If

        ' ** Ctrl keys.
6910    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6920      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6930        With Me
6940          intRetVal = 0
6950          .cmdCalendar1.SetFocus
6960        End With
6970      End Select
6980    End If

EXITP:
6990    KeyCode = intRetVal
7000    Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Public Sub DateStart_Exit(Cancel As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

        Dim strTmp01 As String

7110    With Me
7120      blnRebuildTable = True
7130      If IsNull(.DateStart) = False Then
7140        gdatStartDate = .DateStart
7150        strTmp01 = CStr(CDbl(gdatStartDate))
7160        If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7170        glngStartDateLong = CLng(strTmp01)
7180      End If
7190    End With

EXITP:
7200    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case 94  ' ** Invalid use of Null.
          ' ** Ignore.
1390    Case Else
1400      THAT_PROC = THIS_PROC
1410      That_Erl = Erl
1420      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

7310    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
7320    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

7410    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
7420    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

7510    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
7520    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

7610    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
7620    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

7710    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
7720    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

7810    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
7820    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7920    intShiftDown = (Shift And acShiftMask) > 0
7930    intAltDown = (Shift And acAltMask) > 0
7940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7970        With Me
7980          intRetVal = 0
7990          .cmbAccounts.SetFocus
8000        End With
8010      End Select
8020    End If

        ' ** Shift keys.
8030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8050        With Me
8060          intRetVal = 0
8070          .DateStart.SetFocus
8080        End With
8090      End Select
8100    End If

        ' ** Ctrl keys.
8110    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8120      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8130        With Me
8140          intRetVal = 0
8150          .cmdCalendar2.SetFocus
8160        End With
8170      End Select
8180    End If

        ' ** Ctrl-Shift keys.
8190    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8200      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8210        With Me
8220          intRetVal = 0
8230          .cmdCalendar1.SetFocus
8240        End With
8250      End Select
8260    End If

EXITP:
8270    KeyCode = intRetVal
8280    Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Public Sub DateEnd_Exit(Cancel As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

        Dim strTmp01 As String

8310    With Me
8320      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
8330        .DateEnd = Date
8340      Else
8350        If .DateEnd = vbNullString Then
8360          .DateEnd = Date
8370        End If
8380      End If
8390      blnRebuildTable = True
8400      gdatEndDate = .DateEnd
8410      strTmp01 = CStr(CDbl(gdatEndDate))
8420      If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
8430      glngEndDateLong = CLng(strTmp01)
8440    End With

EXITP:
8450    Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

8510    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
8520    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

8610    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
8620    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

8710    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
8720    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8820    intShiftDown = (Shift And acShiftMask) > 0
8830    intAltDown = (Shift And acAltMask) > 0
8840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8870        With Me
8880          intRetVal = 0
8890          .chkRememberDates.SetFocus
8900        End With
8910      End Select
8920    End If

        ' ** Shift keys.
8930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8950        With Me
8960          intRetVal = 0
8970          .DateEnd.SetFocus
8980        End With
8990      End Select
9000    End If

EXITP:
9010    KeyCode = intRetVal
9020    Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

9110    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
9120    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

9210    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
9220    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

9310    Calendar_Handler_NS THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNS.

EXITP:
9320    Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

9410    With Me
9420      Select Case .chkRememberDates
          Case True
9430        .chkRememberDates_lbl.FontBold = True
9440      Case False
9450        .chkRememberDates_lbl.FontBold = False
9460      End Select
9470    End With

EXITP:
9480    Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

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
9590          .opgAccountSource.SetFocus
9600        End With
9610      End Select
9620    End If

        ' ** Shift keys.
9630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9650        With Me
9660          intRetVal = 0
9670          .cmdCalendar2.SetFocus
9680        End With
9690      End Select
9700    End If

EXITP:
9710    KeyCode = intRetVal
9720    Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

9810    strAccountNo = vbNullString

9820    With Me
9830      If IsNull(.cmbAccounts) = False Then
9840        If Len(.cmbAccounts.Column(0)) > 0 Then
9850          strAccountNo = .cmbAccounts.Column(0)
9860        End If
9870      End If
9880      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
9890        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
9900        .opgAccountSource_optNumber_lbl.FontBold = True
9910        .opgAccountSource_optName_lbl.FontBold = False
9920      Case .opgAccountSource_optName.OptionValue
9930        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
9940        .opgAccountSource_optNumber_lbl.FontBold = False
9950        .opgAccountSource_optName_lbl.FontBold = True
9960      End Select
9970      DoEvents
9980      If strAccountNo <> vbNullString Then
9990        .cmbAccounts = strAccountNo
10000     End If
10010   End With

EXITP:
10020   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10120   intShiftDown = (Shift And acShiftMask) > 0
10130   intAltDown = (Shift And acAltMask) > 0
10140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10170       With Me
10180         intRetVal = 0
10190         .chkRememberMe.SetFocus
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10250       With Me
10260         intRetVal = 0
10270         .chkRememberDates.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10420   intShiftDown = (Shift And acShiftMask) > 0
10430   intAltDown = (Shift And acAltMask) > 0
10440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10470       With Me
10480         intRetVal = 0
10490         .chkRememberMe.SetFocus
10500       End With
10510     End Select
10520   End If

        ' ** Shift keys.
10530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10550       With Me
10560         intRetVal = 0
10570         .chkRememberDates.SetFocus
10580       End With
10590     End Select
10600   End If

EXITP:
10610   KeyCode = intRetVal
10620   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

10710   With Me
10720     Select Case .chkRememberMe
          Case True
10730       .chkRememberMe_lbl.FontBold = True
10740     Case False
10750       .chkRememberMe_lbl.FontBold = False
10760     End Select
10770   End With

EXITP:
10780   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10820   intShiftDown = (Shift And acShiftMask) > 0
10830   intAltDown = (Shift And acAltMask) > 0
10840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         .cmbAccounts.SetFocus
10900       End With
10910     End Select
10920   End If

        ' ** Shift keys.
10930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10940     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10950       With Me
10960         intRetVal = 0
10970         .opgAccountSource.SetFocus
10980       End With
10990     End Select
11000   End If

EXITP:
11010   KeyCode = intRetVal
11020   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

11110   With Me
11120     If .cmbAccounts.text = vbNullString Then
11130       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
11140       .cmbAccounts = .cmbAccounts.ItemData(0)
11150       blnRebuildTable = True
11160     End If
11170   End With

EXITP:
11180   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub cmbAccounts_AfterUpdate()

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

11210   blnRebuildTable = True
11220   gstrCrtRpt_Account = Nz(Me.cmbAccounts, vbNullString)
11230   gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
11240   gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
11250   gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString
11260   gblnCrtRpt_Zero = False: gblnCrtRpt_ZeroDialog = False

EXITP:
11270   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1410    Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

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
11390         .cmdPreview00.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11450       With Me
11460         intRetVal = 0
11470         .DateEnd.SetFocus
11480       End With
11490     End Select
11500   End If

        ' ** Ctrl keys.
11510   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
11520     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11530       With Me
11540         intRetVal = 0
11550         If .chkGroupBy_IncExpCode.Enabled = True Then
11560           .chkGroupBy_IncExpCode.SetFocus
                'ElseIf .chkIncludeArchive.Enabled = True Then
                '  .chkIncludeArchive.SetFocus
11570         Else
11580           .chkLegalName.SetFocus
11590         End If
11600       End With
11610     End Select
11620   End If

        ' ** Ctrl-Shift keys.
11630   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11650       With Me
11660         intRetVal = 0
11670         .chkRememberMe.SetFocus
11680       End With
11690     End Select
11700   End If

EXITP:
11710   KeyCode = intRetVal
11720   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Public Sub chkGroupBy_IncExpCode_AfterUpdate()
' ** Public for GoToReport.
' ** Available on reports 0, 3, and 8.

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_AfterUpdate"

11810   With Me
11820     If .chkGroupBy_IncExpCode Then
11830       gblnUseReveuneExpenseCodes = True
11840       .chkGroupBy_IncExpCode_lbl.FontBold = True
11850       .chkGroupBy_IncExpCode_lbl4.FontSize = 20
11860       .chkGroupBy_IncExpCode_lbl5.Visible = True
11870       .cmdPreview00_lbl4.FontSize = 20
11880       .cmdPreview00_lbl5.Visible = True
11890       .cmdPreview01_lbl4.FontSize = 20
11900       .cmdPreview01_lbl5.Visible = True
11910       .cmdPreview03_lbl4.FontSize = 20
11920       .cmdPreview03_lbl5.Visible = True
11930       .cmdPreview07_lbl4.FontSize = 20
11940       .cmdPreview07_lbl5.Visible = True
11950       .cmdPreview08_lbl4.FontSize = 20
11960       .cmdPreview08_lbl5.Visible = True
11970     Else
11980       gblnUseReveuneExpenseCodes = False
11990       .chkGroupBy_IncExpCode_lbl.FontBold = False
12000       .chkGroupBy_IncExpCode_lbl4.FontSize = 14
12010       .chkGroupBy_IncExpCode_lbl5.Visible = False
12020       .cmdPreview00_lbl4.FontSize = 14
12030       .cmdPreview00_lbl5.Visible = False
12040       .cmdPreview01_lbl4.FontSize = 14
12050       .cmdPreview01_lbl5.Visible = False
12060       .cmdPreview03_lbl4.FontSize = 14
12070       .cmdPreview03_lbl5.Visible = False
12080       .cmdPreview07_lbl4.FontSize = 14
12090       .cmdPreview07_lbl5.Visible = False
12100       .cmdPreview08_lbl4.FontSize = 14
12110       .cmdPreview08_lbl5.Visible = False
12120     End If
12130   End With

EXITP:
12140   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub chkGroupBy_IncExpCode_KeyDown(KeyCode As Integer, Shift As Integer)

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12220   intShiftDown = (Shift And acShiftMask) > 0
12230   intAltDown = (Shift And acAltMask) > 0
12240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12270       With Me
12280         intRetVal = 0
12290         .chkLegalName.SetFocus
12300       End With
12310     End Select
12320   End If

        ' ** Shift keys.
12330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12350       With Me
12360         intRetVal = 0
12370         .cmbAccounts.SetFocus
12380       End With
12390     End Select
12400   End If

EXITP:
12410   KeyCode = intRetVal
12420   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub chkLegalName_AfterUpdate()

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_AfterUpdate"

12510   With Me
12520     Select Case .chkLegalName
          Case True
12530       .chkLegalName_lbl.FontBold = True
12540       gblnLegalName = True
12550     Case False
12560       .chkLegalName_lbl.FontBold = False
12570       gblnLegalName = False
12580     End Select
12590   End With

EXITP:
12600   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub chkLegalName_KeyDown(KeyCode As Integer, Shift As Integer)

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_KeyDown"

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
12790         .chkIncludeFootnote.SetFocus
12800       End With
12810     End Select
12820   End If

        ' ** Shift keys.
12830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12840     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12850       With Me
12860         intRetVal = 0
12870         If .chkGroupBy_IncExpCode.Enabled = True Then
12880           .chkGroupBy_IncExpCode.SetFocus
12890         Else
12900           .cmbAccounts.SetFocus
12910         End If
12920       End With
12930     End Select
12940   End If

EXITP:
12950   KeyCode = intRetVal
12960   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub chkIncludeFootnote_AfterUpdate()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_AfterUpdate"

13010   With Me
13020     .CourtReports_Footnote.Visible = True
13030     .CourtReports_Footnote_box.Visible = True
13040     .CourtReports_Footnote_lbl2.Visible = False
13050     Select Case .chkIncludeFootnote
          Case True
13060       .chkIncludeFootnote_lbl.FontBold = True
13070       .CourtReports_Footnote.Enabled = True
13080       .CourtReports_Footnote.BackStyle = acBackStyleNormal
13090       .CourtReports_Footnote_lbl.BackStyle = acBackStyleNormal
13100       .CourtReports_Footnote.BorderColor = CLR_LTBLU2
13110       If IsNull(.CourtReports_Footnote) = True Then
13120         .CourtReports_Footnote = strFootnoteDefault
13130       End If
13140       If gblnRevenueExpenseTracking = False Then
13150         .chkGroupBy_IncExpCode_lbl3.Visible = False
13160       End If
13170     Case False
13180       .chkIncludeFootnote_lbl.FontBold = False
13190       .CourtReports_Footnote.Enabled = False
13200       .CourtReports_Footnote.BackStyle = acBackStyleTransparent
13210       .CourtReports_Footnote_lbl.BackStyle = acBackStyleTransparent
13220       .CourtReports_Footnote.BorderColor = WIN_CLR_DISR
13230       If gblnRevenueExpenseTracking = False Then
13240         .chkGroupBy_IncExpCode_lbl3.Visible = True
13250       End If
13260     End Select
13270   End With

EXITP:
13280   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub chkIncludeFootnote_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13320   intShiftDown = (Shift And acShiftMask) > 0
13330   intAltDown = (Shift And acAltMask) > 0
13340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13370       With Me
13380         intRetVal = 0
13390         Select Case .chkIncludeFootnote
              Case True
13400           .CourtReports_Footnote.SetFocus
13410         Case False
13420           .chkPageOf.SetFocus
13430         End Select
13440       End With
13450     End Select
13460   End If

        ' ** Shift keys.
13470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13480     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13490       With Me
13500         intRetVal = 0
13510         .chkLegalName.SetFocus
13520       End With
13530     End Select
13540   End If

EXITP:
13550   KeyCode = intRetVal
13560   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub CourtReports_Footnote_AfterUpdate()

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_AfterUpdate"

        Dim varTmp00 As Variant

13610   With Me
13620     varTmp00 = .CourtReports_Footnote
13630     If IsNull(varTmp00) = False Then
13640       If Trim(varTmp00) <> vbNullString Then
13650         varTmp00 = Trim(varTmp00)
13660         If Len(varTmp00) > MEMO_MAX Then
13670           Beep
13680           If .CourtReports_Footnote_lbl2.Visible = False Then
13690             .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
13700             .CourtReports_Footnote_lbl2.Visible = True
13710           End If
13720           varTmp00 = Left(varTmp00, MEMO_MAX)
13730           .CourtReports_Footnote = varTmp00
13740         End If
13750       Else
13760         .CourtReports_Footnote = Null
13770       End If
13780     End If
13790   End With

EXITP:
13800   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub CourtReports_Footnote_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intLen As Integer
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
13990         .chkPageOf.SetFocus
14000       End With
14010     Case Else
14020       With Me
14030         If IsNull(.CourtReports_Footnote.text) = False Then
14040           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
14050             intLen = Len(Trim(.CourtReports_Footnote.text))
14060             If intLen > MEMO_MAX Then
14070               intRetVal = 0
14080               Beep
14090               If .CourtReports_Footnote_lbl2.Visible = False Then
14100                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
14110                 .CourtReports_Footnote_lbl2.Visible = True
14120               End If
14130               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
14140               .CourtReports_Footnote.SelStart = 999
14150             End If
14160           End If
14170         End If
14180       End With
14190     End Select
14200   End If

        ' ** Shift keys.
14210   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14220     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14230       With Me
14240         intRetVal = 0
14250         .chkIncludeFootnote.SetFocus
14260       End With
14270     Case Else
14280       With Me
14290         If IsNull(.CourtReports_Footnote.text) = False Then
14300           If Trim(.CourtReports_Footnote.text) <> vbNullString Then
14310             intLen = Len(Trim(.CourtReports_Footnote.text))
14320             If intLen > MEMO_MAX Then
14330               intRetVal = 0
14340               Beep
14350               If .CourtReports_Footnote_lbl2.Visible = False Then
14360                 .CourtReports_Footnote.Width = (.CourtReports_Footnote.Width - (46& * lngTpp))
14370                 .CourtReports_Footnote_lbl2.Visible = True
14380               End If
14390               .CourtReports_Footnote.text = Left(Trim(.CourtReports_Footnote.text), MEMO_MAX)
14400               .CourtReports_Footnote.SelStart = 999
14410             End If
14420           End If
14430         End If
14440       End With
14450     End Select
14460   End If

EXITP:
14470   KeyCode = intRetVal
14480   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

14510   With Me
14520     Select Case .chkPageOf
          Case True
14530       .chkPageOf_lbl.FontBold = True
14540     Case False
14550       .chkPageOf_lbl.FontBold = False
14560     End Select
14570   End With

EXITP:
14580   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14620   intShiftDown = (Shift And acShiftMask) > 0
14630   intAltDown = (Shift And acAltMask) > 0
14640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14670       With Me
14680         intRetVal = 0
14690         .cmdPreview00.SetFocus
14700       End With
14710     End Select
14720   End If

        ' ** Shift keys.
14730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14750       With Me
14760         intRetVal = 0
14770         .chkIncludeFootnote.SetFocus
14780       End With
14790     End Select
14800   End If

EXITP:
14810   KeyCode = intRetVal
14820   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdPreview00_lbl4_cmd_DblClick(Cancel As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_lbl4_cmd_DblClick"

14910   With Me
14920     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
14930     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
14940   End With

EXITP:
14950   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview00_Click()

15000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_Click"

15010   With Me

15020     blnAllCancel = False
15030     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.
15040     gstrAccountNo = .cmbAccounts.Column(0)

15050     Select Case gblnUseReveuneExpenseCodes
          Case True
15060       PreviewOrPrint_NS "0A", THIS_PROC, acViewPreview  ' ** Function: Below.
15070     Case False
15080       PreviewOrPrint_NS "0", THIS_PROC, acViewPreview  ' ** Function: Below.
15090     End Select

15100     If .chkAssetList = True Then
15110       AssetList_PreviewPrint "Beginning", acViewPreview  ' ** Procedure: Below.
15120     End If

15130   End With

EXITP:
15140   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub cmdPreview00_KeyDown(KeyCode As Integer, Shift As Integer)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_KeyDown"

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
          Case vbKeyTab
15270       With Me
15280         intRetVal = 0
15290         .cmdPrint00.SetFocus
15300       End With
15310     End Select
15320   End If

        ' ** Shift keys.
15330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15340     Select Case intRetVal
          Case vbKeyTab
15350       With Me
15360         intRetVal = 0
15370         .cmbAccounts.SetFocus
15380       End With
15390     End Select
15400   End If

        ' ** Ctrl-Shift keys.
15410   If intCtrlDown And (Not intAltDown) And intShiftDown Then
15420     Select Case intRetVal
          Case vbKeyTab
15430       With Me
15440         intRetVal = 0
15450         .chkPageOf.SetFocus
15460       End With
15470     End Select
15480   End If

EXITP:
15490   KeyCode = intRetVal
15500   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Public Sub cmdPrint00_Click()

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00_Click"

15610   With Me

15620     blnAllCancel = False
15630     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.
15640     gstrAccountNo = .cmbAccounts.Column(0)

15650     Select Case gblnUseReveuneExpenseCodes
          Case True
            '##GTR_Ref: rptCourtRptNS_00A
15660       PreviewOrPrint_NS "0A", THIS_PROC, acViewNormal  ' ** Function: Below.
15670     Case False
            '##GTR_Ref: rptCourtRptNS_00
15680       PreviewOrPrint_NS "0", THIS_PROC, acViewNormal  ' ** Function: Below.
15690     End Select

15700     If .chkAssetList = True Then
15710       AssetList_PreviewPrint "Beginning", acViewNormal  ' ** Procedure: Below.
15720     End If

15730   End With

EXITP:
15740   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord00_Click()

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_Click"

15810   With Me

15820     blnAllCancel = False
15830     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.
15840     gstrAccountNo = .cmbAccounts.Column(0)

15850     Select Case gblnUseReveuneExpenseCodes
          Case True
15860       SendToFile_NS "0A", THIS_PROC, False  ' ** Procedure: Below.
15870     Case False
15880       SendToFile_NS "0", THIS_PROC, False  ' ** Procedure: Below.
15890     End Select

15900     If .chkAssetList = True Then
15910       AssetList_Word_NS "Beginning", Me  ' ** Module Procedure: modCourtReportsNS.
15920     End If

15930   End With

EXITP:
15940   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdWord00_KeyDown(KeyCode As Integer, Shift As Integer)

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16020   intShiftDown = (Shift And acShiftMask) > 0
16030   intAltDown = (Shift And acAltMask) > 0
16040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16060     Select Case intRetVal
          Case vbKeyTab
16070       With Me
16080         intRetVal = 0
16090         If .cmdExcel00.Enabled = True Then
16100           .cmdExcel00.SetFocus
16110         Else
16120           .cmdPreview00B.SetFocus
16130         End If
16140       End With
16150     Case vbKeyRight
16160       With Me
16170         intRetVal = 0
16180         If .cmdExcel00.Enabled = True Then
16190           .cmdExcel00.SetFocus
16200         Else
16210           .cmdPreview00B.SetFocus
16220         End If
16230       End With
16240     End Select
16250   End If

        ' ** Shift keys.
16260   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16270     Select Case intRetVal
          Case vbKeyTab
16280       With Me
16290         intRetVal = 0
16300         .cmdPrint00.SetFocus
16310       End With
16320     End Select
16330   End If

EXITP:
16340   KeyCode = intRetVal
16350   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Public Sub cmdExcel00_Click()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_Click"

16410   With Me

16420     blnAllCancel = False
16430     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.
16440     gstrAccountNo = .cmbAccounts.Column(0)

16450     Select Case gblnUseReveuneExpenseCodes
          Case True
16460       SendToFile_NS "0A", THIS_PROC, True  ' ** Procedure: Below.
16470     Case False
16480       SendToFile_NS "0", THIS_PROC, True  ' ** Procedure: Below.
16490     End Select

16500     If .chkAssetList = True Then
16510       AssetList_Excel_NS "Beginning", lngCaps, arr_varCap, Me  ' ** Module Procedure: modCourtReportsNS.
16520     End If

16530   End With

EXITP:
16540   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdExcel00_KeyDown(KeyCode As Integer, Shift As Integer)

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16620   intShiftDown = (Shift And acShiftMask) > 0
16630   intAltDown = (Shift And acAltMask) > 0
16640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16660     Select Case intRetVal
          Case vbKeyTab
16670       With Me
16680         intRetVal = 0
16690         .cmdPreview00B.SetFocus
16700       End With
16710     Case vbKeyRight
16720       With Me
16730         intRetVal = 0
16740         .cmdPreview00B.SetFocus
16750       End With
16760     End Select
16770   End If

        ' ** Shift keys.
16780   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16790     Select Case intRetVal
          Case vbKeyTab
16800       With Me
16810         intRetVal = 0
16820         .cmdWord00.SetFocus
16830       End With
16840     End Select
16850   End If

EXITP:
16860   KeyCode = intRetVal
16870   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub chkAssetList_AfterUpdate()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_AfterUpdate"

16910   With Me
16920     Select Case .chkAssetList
          Case True
16930       .chkAssetList_lbl.FontBold = True
16940       .chkAssetList_lbl2.FontBold = True
16950     Case False
16960       .chkAssetList_lbl.FontBold = False
16970       .chkAssetList_lbl2.FontBold = False
16980     End Select
16990   End With

EXITP:
17000   Exit Sub

ERRH:
1380    THAT_PROC = THIS_PROC
1390    That_Erl = Erl
1400    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1410    Resume EXITP

End Sub

Private Sub chkAssetList_KeyDown(KeyCode As Integer, Shift As Integer)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_KeyDown"

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
17190         .cmdPreview00B.SetFocus
17200       End With
17210     End Select
17220   End If

        ' ** Shift keys.
17230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
17250       With Me
17260         intRetVal = 0
17270         If .cmdExcel00.Enabled = True Then
17280           .cmdExcel00.SetFocus
17290         Else
17300           .cmdWord00.SetFocus
17310         End If
17320       End With
17330     End Select
17340   End If

EXITP:
17350   KeyCode = intRetVal
17360   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdPreview00B_Click()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_Click"

17410   If Validate = True Then  ' ** Function: Below.

17420     blnAllCancel = False
17430     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

17440     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

17450     AssetList_PreviewPrint "Beginning", acViewPreview  ' ** Procedure: Below.

17460   End If  ' ** Validate().

EXITP:
17470   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub cmdPreview00B_KeyDown(KeyCode As Integer, Shift As Integer)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17520   intShiftDown = (Shift And acShiftMask) > 0
17530   intAltDown = (Shift And acAltMask) > 0
17540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17560     Select Case intRetVal
          Case vbKeyTab
17570       With Me
17580         intRetVal = 0
17590         .cmdPrint00B.SetFocus
17600       End With
17610     Case vbKeyLeft
17620       With Me
17630         intRetVal = 0
17640         If .cmdExcel00.Enabled = True Then
17650           .cmdExcel00.SetFocus
17660         Else
17670           .cmdWord00.SetFocus
17680         End If
17690       End With
17700     End Select
17710   End If

        ' ** Shift keys.
17720   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17730     Select Case intRetVal
          Case vbKeyTab
17740       With Me
17750         intRetVal = 0
17760         If .cmdExcel00.Enabled = True Then
17770           .cmdExcel00.SetFocus
17780         Else
17790           .cmdWord00.SetFocus
17800         End If
17810       End With
17820     End Select
17830   End If

EXITP:
17840   KeyCode = intRetVal
17850   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdPrint00B_Click()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00B_Click"

17910   If Validate = True Then  ' ** Function: Below.

17920     blnAllCancel = False
17930     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

17940     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

          '##GTR_Ref: rptCourtRptNS_00BA
          '##GTR_Ref: rptCourtRptNS_00B
17950     AssetList_PreviewPrint "Beginning", acViewNormal  ' ** Procedure: Below.

17960   End If  ' ** Validate().

EXITP:
17970   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Public Sub cmdWord00B_Click()

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00B_Click"

18010   If Validate = True Then  ' ** Function: Below.

18020     blnAllCancel = False
18030     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18040     AssetList_Word_NS "Beginning", Me  ' ** Module Procedure: modCourtReportsNS.

18050   End If  ' ** Validate().

EXITP:
18060   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Public Sub cmdExcel00B_Click()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00B_Click"

18110   If Validate = True Then  ' ** Function: Below.

18120     blnAllCancel = False
18130     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18140     AssetList_Excel_NS "Beginning", lngCaps, arr_varCap, Me  ' ** Module Procedure: modCourtReportsNS.

18150   End If  ' ** Validate().

EXITP:
18160   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub cmdPreview01_lbl4_cmd_DblClick(Cancel As Integer)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview01_lbl4_cmd_DblClick"

18210   With Me
18220     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
18230     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
18240   End With

EXITP:
18250   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview01_Click()

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview01_Click"

18310   blnAllCancel = False
18320   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18330   Select Case gblnUseReveuneExpenseCodes
        Case True
18340     PreviewOrPrint_NS "1A", THIS_PROC, acViewPreview  ' ** Function: Below.
18350   Case False
18360     PreviewOrPrint_NS "1", THIS_PROC, acViewPreview  ' ** Function: Below.
18370   End Select

EXITP:
18380   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint01_Click()

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint01_Click"

18410   blnAllCancel = False
18420   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18430   Select Case gblnUseReveuneExpenseCodes
        Case True
          '##GTR_Ref: rptCourtRptNS_01A
18440     PreviewOrPrint_NS "1A", THIS_PROC, acViewNormal  ' ** Function: Below.
18450   Case False
          '##GTR_Ref: rptCourtRptNS_01
18460     PreviewOrPrint_NS "1", THIS_PROC, acViewNormal  ' ** Function: Below.
18470   End Select

EXITP:
18480   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord01_Click()

18500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord01_Click"

18510   blnAllCancel = False
18520   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18530   Select Case gblnUseReveuneExpenseCodes
        Case True
18540     SendToFile_NS "1A", THIS_PROC, False  ' ** Procedure: Below.
18550   Case False
18560     SendToFile_NS "1", THIS_PROC, False  ' ** Procedure: Below.
18570   End Select

EXITP:
18580   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel01_Click()

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel01_Click"

18610   blnAllCancel = False
18620   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18630   Select Case gblnUseReveuneExpenseCodes
        Case True
18640     SendToFile_NS "1A", THIS_PROC, True  ' ** Procedure: Below.
18650   Case False
18660     SendToFile_NS "1", THIS_PROC, True  ' ** Procedure: Below.
18670   End Select

EXITP:
18680   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview02_Click()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview02_Click"

18710   blnAllCancel = False
18720   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18730   PreviewOrPrint_NS "2", THIS_PROC, acViewPreview  ' ** Function: Below.

EXITP:
18740   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint02_Click()

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint02_Click"

18810   blnAllCancel = False
18820   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

        '##GTR_Ref: rptCourtRptNS_02
18830   PreviewOrPrint_NS "2", THIS_PROC, acViewNormal  ' ** Function: Below.

EXITP:
18840   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord02_Click()

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord02_Click"

18910   blnAllCancel = False
18920   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

18930   SendToFile_NS "2", THIS_PROC, False  ' ** Procedure: Below.

EXITP:
18940   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel02_Click()

19000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel02_Click"

19010   blnAllCancel = False
19020   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19030   SendToFile_NS "2", THIS_PROC, True  ' ** Procedure: Below.

EXITP:
19040   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview03_lbl4_cmd_DblClick(Cancel As Integer)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview03_lbl4_cmd_DblClick"

19110   With Me

19120     .chkGroupBy_IncExpCode = (Not .chkGroupBy_IncExpCode)
19130     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.

19140   End With

EXITP:
19150   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview03_Click()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview03_Click"

19210   blnAllCancel = False
19220   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19230   Select Case gblnUseReveuneExpenseCodes
        Case True
19240     PreviewOrPrint_NS "3A", THIS_PROC, acViewPreview  ' ** Function: Below.
19250   Case False
19260     PreviewOrPrint_NS "3", THIS_PROC, acViewPreview  ' ** Function: Below.
19270   End Select

EXITP:
19280   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint03_Click()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint03_Click"

19310   blnAllCancel = False
19320   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19330   Select Case gblnUseReveuneExpenseCodes
        Case True
          '##GTR_Ref: rptCourtRptNS_03A
19340     PreviewOrPrint_NS "3A", THIS_PROC, acViewNormal  ' ** Function: Below.
19350   Case False
          '##GTR_Ref: rptCourtRptNS_03
19360     PreviewOrPrint_NS "3", THIS_PROC, acViewNormal  ' ** Function: Below.
19370   End Select

EXITP:
19380   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord03_Click()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord03_Click"

19410   blnAllCancel = False
19420   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19430   Select Case gblnUseReveuneExpenseCodes
        Case True
19440     SendToFile_NS "3A", THIS_PROC, False  ' ** Procedure: Below.
19450   Case False
19460     SendToFile_NS "3", THIS_PROC, False  ' ** Procedure: Below.
19470   End Select

EXITP:
19480   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel03_Click()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel03_Click"

19510   blnAllCancel = False
19520   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19530   Select Case gblnUseReveuneExpenseCodes
        Case True
19540     SendToFile_NS "3A", THIS_PROC, True  ' ** Procedure: Below.
19550   Case False
19560     SendToFile_NS "3", THIS_PROC, True  ' ** Procedure: Below.
19570   End Select

EXITP:
19580   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview04_Click()

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_Click"

19610   blnAllCancel = False
19620   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19630   PreviewOrPrint_NS "4", THIS_PROC, acViewPreview  ' ** Function: Below.

EXITP:
19640   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint04_Click()

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint04_Click"

19710   blnAllCancel = False
19720   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

        '##GTR_Ref: rptCourtRptNS_04
19730   PreviewOrPrint_NS "4", THIS_PROC, acViewNormal  ' ** Function: Below.

EXITP:
19740   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord04_Click()

19800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord04_Click"

19810   blnAllCancel = False
19820   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19830   SendToFile_NS "4", THIS_PROC, False  ' ** Procedure: Below.

EXITP:
19840   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel04_Click()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel04_Click"

19910   blnAllCancel = False
19920   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

19930   SendToFile_NS "4", THIS_PROC, True  ' ** Procedure: Below.

EXITP:
19940   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview05_Click()

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview05_Click"

20010   blnAllCancel = False
20020   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

20030   PreviewOrPrint_NS "5", THIS_PROC, acViewPreview  ' ** Function: Below.

EXITP:
20040   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint05_Click()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint05_Click"

20110   blnAllCancel = False
20120   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

        '##GTR_Ref: rptCourtRptNS_05
20130   PreviewOrPrint_NS "5", THIS_PROC, acViewNormal  ' ** Function: Below.

EXITP:
20140   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord05_Click()

20200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord05_Click"

20210   blnAllCancel = False
20220   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

20230   SendToFile_NS "5", THIS_PROC, False  ' ** Procedure: Below.

EXITP:
20240   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel05_Click()

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel05_Click"

20310   blnAllCancel = False
20320   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

20330   SendToFile_NS "5", THIS_PROC, True  ' ** Procedure: Below.

EXITP:
20340   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview06_Click()

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview06_Click"

20410   blnAllCancel = False
20420   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

20430   PreviewOrPrint_NS "6", THIS_PROC, acViewPreview  ' ** Function: Below.

EXITP:
20440   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint06_Click()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint06_Click"

20510   blnAllCancel = False
20520   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

        '##GTR_Ref: rptCourtRptNS_06
20530   PreviewOrPrint_NS "6", THIS_PROC, acViewNormal  ' ** Function: Below.

EXITP:
20540   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord06_Click()

20600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord06_Click"

20610   blnAllCancel = False
20620   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

20630   SendToFile_NS "6", THIS_PROC, False  ' ** Procedure: Below.

EXITP:
20640   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel06_Click()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel06_Click"

20710   blnAllCancel = False
20720   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

20730   SendToFile_NS "6", THIS_PROC, True  ' ** Procedure: Below.

EXITP:
20740   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview07_lbl4_cmd_DblClick(Cancel As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview07_lbl4_cmd_DblClick"

20810   With Me

20820     .chkGroupBy_IncExpCode = (Not .chkGroupBy_IncExpCode)
20830     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.

20840   End With

EXITP:
20850   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview07_Click()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview07_Click"

20910   blnAllCancel = False
20920   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

20930   Select Case gblnUseReveuneExpenseCodes
        Case True
20940     PreviewOrPrint_NS "7A", THIS_PROC, acViewPreview  ' ** Function: Below.
20950   Case False
20960     PreviewOrPrint_NS "7", THIS_PROC, acViewPreview  ' ** Function: Below.
20970   End Select

EXITP:
20980   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint07_Click()

21000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint07_Click"

21010   blnAllCancel = False
21020   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21030   Select Case gblnUseReveuneExpenseCodes
        Case True
          '##GTR_Ref: rptCourtRptNS_07A
21040     PreviewOrPrint_NS "7A", THIS_PROC, acViewNormal  ' ** Function: Below.
21050   Case False
          '##GTR_Ref: rptCourtRptNS_07
21060     PreviewOrPrint_NS "7", THIS_PROC, acViewNormal  ' ** Function: Below.
21070   End Select

EXITP:
21080   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord07_Click()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord07_Click"

21110   blnAllCancel = False
21120   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21130   Select Case gblnUseReveuneExpenseCodes
        Case True
21140     SendToFile_NS "7A", THIS_PROC, False  ' ** Procedure: Below.
21150   Case False
21160     SendToFile_NS "7", THIS_PROC, False  ' ** Procedure: Below.
21170   End Select

EXITP:
21180   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel07_Click()

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel07_Click"

21210   blnAllCancel = False
21220   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21230   Select Case gblnUseReveuneExpenseCodes
        Case True
21240     SendToFile_NS "7A", THIS_PROC, True  ' ** Procedure: Below.
21250   Case False
21260     SendToFile_NS "7", THIS_PROC, True  ' ** Procedure: Below.
21270   End Select

EXITP:
21280   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview08_lbl4_cmd_DblClick(Cancel As Integer)

21300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview08_lbl4_cmd_DblClick"

21310   With Me

21320     .chkGroupBy_IncExpCode = (Not .chkGroupBy_IncExpCode)
21330     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.

21340   End With

EXITP:
21350   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview08_Click()

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview08_Click"

21410   blnAllCancel = False
21420   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21430   Select Case gblnUseReveuneExpenseCodes
        Case True
21440     PreviewOrPrint_NS "8A", THIS_PROC, acViewPreview  ' ** Function: Below.
21450   Case False
21460     PreviewOrPrint_NS "8", THIS_PROC, acViewPreview  ' ** Function: Below.
21470   End Select

EXITP:
21480   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint08_Click()

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint08_Click"

21510   blnAllCancel = False
21520   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21530   Select Case gblnUseReveuneExpenseCodes
        Case True
          '##GTR_Ref: rptCourtRptNS_08A
21540     PreviewOrPrint_NS "8A", THIS_PROC, acViewNormal  ' ** Function: Below.
21550   Case False
          '##GTR_Ref: rptCourtRptNS_08
21560     PreviewOrPrint_NS "8", THIS_PROC, acViewNormal  ' ** Function: Below.
21570   End Select

EXITP:
21580   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord08_Click()

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord08_Click"

21610   blnAllCancel = False
21620   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21630   Select Case gblnUseReveuneExpenseCodes
        Case True
21640     SendToFile_NS "8A", THIS_PROC, False  ' ** Procedure: Below.
21650   Case False
21660     SendToFile_NS "8", THIS_PROC, False  ' ** Procedure: Below.
21670   End Select

EXITP:
21680   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdExcel08_Click()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel08_Click"

21710   blnAllCancel = False
21720   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21730   Select Case gblnUseReveuneExpenseCodes
        Case True
21740     SendToFile_NS "8A", THIS_PROC, True  ' ** Procedure: Below.
21750   Case False
21760     SendToFile_NS "8", THIS_PROC, True  ' ** Procedure: Below.
21770   End Select

EXITP:
21780   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPreview09_Click()

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview09_Click"

21810   blnAllCancel = False
21820   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

21830   PreviewOrPrint_NS "9", THIS_PROC, acViewPreview  ' ** Function: Below.

EXITP:
21840   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrint09_Click()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint09_Click"

21910   blnAllCancel = False
21920   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

        '##GTR_Ref: rptCourtRptNS_09
21930   PreviewOrPrint_NS "9", THIS_PROC, acViewNormal  ' ** Function: Below.

EXITP:
21940   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Public Sub cmdWord09_Click()

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord09_Click"

22010   blnAllCancel = False
22020   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

22030   SendToFile_NS "9", THIS_PROC, False  ' ** Procedure: Below.

EXITP:
22040   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdWord09_KeyDown(KeyCode As Integer, Shift As Integer)

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord09_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22120   intShiftDown = (Shift And acShiftMask) > 0
22130   intAltDown = (Shift And acAltMask) > 0
22140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22160     Select Case intRetVal
          Case vbKeyTab
22170       With Me
22180         intRetVal = 0
22190         If .cmdExcel09.Enabled = True Then
22200           .cmdExcel09.SetFocus
22210         Else
22220           .cmdPreview10.SetFocus
22230         End If
22240       End With
22250     Case vbKeyLeft
22260       With Me
22270         intRetVal = 0
22280         .cmdPrint09.SetFocus
22290       End With
22300     Case vbKeyRight
22310       With Me
22320         intRetVal = 0
22330         If .cmdExcel09.Enabled = True Then
22340           .cmdExcel09.SetFocus
22350         Else
22360           .cmdPreview10.SetFocus
22370         End If
22380       End With
22390     End Select
22400   End If

        ' ** Shift keys.
22410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22420     Select Case intRetVal
          Case vbKeyTab
22430       With Me
22440         intRetVal = 0
22450         .cmdWord09.SetFocus
22460       End With
22470     End Select
22480   End If

EXITP:
22490   KeyCode = intRetVal
22500   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Public Sub cmdExcel09_Click()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel09_Click"

22610   blnAllCancel = False
22620   AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

22630   SendToFile_NS "9", THIS_PROC, True  ' ** Procedure: Below.

EXITP:
22640   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdExcel09_KeyDown(KeyCode As Integer, Shift As Integer)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel09_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22720   intShiftDown = (Shift And acShiftMask) > 0
22730   intAltDown = (Shift And acAltMask) > 0
22740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22760     Select Case intRetVal
          Case vbKeyTab
22770       With Me
22780         intRetVal = 0
22790         .cmdPreview10.SetFocus
22800       End With
22810     Case vbKeyLeft
22820       With Me
22830         intRetVal = 0
22840         .cmdWord09.SetFocus
22850       End With
22860     Case vbKeyRight
22870       With Me
22880         intRetVal = 0
22890         .cmdPreview10.SetFocus
22900       End With
22910     End Select
22920   End If

        ' ** Shift keys.
22930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22940     Select Case intRetVal
          Case vbKeyTab
22950       With Me
22960         intRetVal = 0
22970         .cmdWord09.SetFocus
22980       End With
22990     End Select
23000   End If

EXITP:
23010   KeyCode = intRetVal
23020   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdPreview10_Click()

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview10_Click"

23110   If Validate = True Then  ' ** Function: Below.

23120     blnAllCancel = False
23130     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

23140     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

23150     AssetList_PreviewPrint "Ending", acViewPreview  ' ** Procedure: Below.

23160   End If  ' ** Validate().

EXITP:
23170   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub cmdPrint10_Click()

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint10_Click"

23210   If Validate = True Then  ' ** Function: Below.

23220     blnAllCancel = False
23230     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

23240     TAReports_SetZero False  ' ** Module Function: modReportFunctions.

          '##GTR_Ref: rptCourtRptNS_00D
          '##GTR_Ref: rptCourtRptNS_00DA
23250     AssetList_PreviewPrint "Ending", acViewNormal  ' ** Procedure: Below.

23260   End If  ' ** Validate().

EXITP:
23270   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Public Sub cmdWord10_Click()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord10_Click"

23310   If Validate = True Then  ' ** Function: Below.

23320     blnAllCancel = False
23330     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

23340     AssetList_Word_NS "Ending", Me  ' ** Module Procedure: modCourtReportsNS.

23350   End If  ' ** Validate().

EXITP:
23360   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Public Sub cmdExcel10_Click()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel10_Click"

23410   If Validate = True Then  ' ** Function: Below.

23420     blnAllCancel = False
23430     AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.

23440     AssetList_Excel_NS "Ending", lngCaps, arr_varCap, Me  ' ** Module Procedure: modCourtReportsNS.

23450   End If  ' ** Validate().

EXITP:
23460   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case Else
1400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1410    End Select
1420    Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

23510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
23520   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

23610   With Me
23620     blnPrintAll_Focus = True
23630     .cmdPrintAll_box01.Visible = True
23640     .cmdPrintAll_box02.Visible = True
23650     Select Case .chkAssetList
          Case True
23660       .cmdPrintAll_box03.Visible = True
23670     Case False
23680       .cmdPrintAll_box03.Visible = False
23690     End Select
23700     .cmdPrintAll_box04.Visible = True
23710   End With

EXITP:
23720   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strRpt As String
        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer

23810   With Me

23820     blnContinue = True
23830     strRpt = vbNullString
23840     intRetVal_BuildAssetListInfo = 0

23850     ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

23860     If .chkAssetList = True Then
23870       intRetVal_BuildAssetListInfo = BuildAssetListInfo_NS(.DateStart, .DateEnd, "Ending", strRpt, THIS_PROC)  ' ** Function: Below.
23880       If intRetVal_BuildAssetListInfo = 0 Then
              ' ** strRpt should return either "_00B" or "_00BA".
23890         If strRpt <> vbNullString Then
23900           gdatStartDate = .DateStart
23910           gdatEndDate = .DateEnd
23920           gstrAccountNo = .cmbAccounts.Column(0)
23930           gstrAccountName = .cmbAccounts.Column(3)
23940           If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
23950             DoCmd.OpenReport "rptCourtRptNS" & strRpt, acViewPreview
23960           Else
23970             DoCmd.OpenReport "rptCourtRptNS" & strRpt, acViewNormal
23980           End If
23990         Else
24000           blnContinue = False
24010           MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
24020         End If
24030       Else
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.
24040         If intRetVal_BuildAssetListInfo = -2 Then
24050           Set dbs = CurrentDb
                ' ** Empty tmpAssetList2.
24060           Set qdf = dbs.QueryDefs("qryCourtReport_03")
24070           qdf.Execute
24080           DoCmd.OpenReport "rptCourtRptNS_00D", acViewNormal
24090         End If
24100       End If
24110     End If

          'DOES THIS MEAN IT DOESN'T ASK FOR THE INPUTS TILL ALL THE OTHER REPORTS HAVE BEEN PRINTED?!
24120     If blnContinue = True Then
24130       blnContinue = PreviewOrPrint_NS("9", THIS_PROC & " - R9", acViewNormal)  ' ** Function: Below.
24140       If blnContinue = True Then
24150         If gblnUseReveuneExpenseCodes = True Then
24160           blnContinue = PreviewOrPrint_NS("8A", THIS_PROC & " - R8A", acViewNormal)  ' ** Function: Below.
24170         Else
24180           blnContinue = PreviewOrPrint_NS("8", THIS_PROC & " - R8", acViewNormal)  ' ** Function: Below.
24190         End If
24200         If blnContinue = True Then
24210           blnContinue = PreviewOrPrint_NS("7", THIS_PROC & " - R7", acViewNormal)  ' ** Function: Below.
24220           If blnContinue = True Then
24230             blnContinue = PreviewOrPrint_NS("6", THIS_PROC & " - R6", acViewNormal)  ' ** Function: Below.
24240             If blnContinue = True Then
24250               blnContinue = PreviewOrPrint_NS("5", THIS_PROC & " - R5", acViewNormal)  ' ** Function: Below.
24260               If blnContinue = True Then
24270                 blnContinue = PreviewOrPrint_NS("4", THIS_PROC & " - R4", acViewNormal)  ' ** Function: Below.
24280                 If blnContinue = True Then
24290                   If gblnUseReveuneExpenseCodes = True Then
24300                     blnContinue = PreviewOrPrint_NS("3A", THIS_PROC & " - R3A", acViewNormal)  ' ** Function: Below.
24310                   Else
24320                     blnContinue = PreviewOrPrint_NS("3", THIS_PROC & " - R3", acViewNormal)  ' ** Function: Below.
24330                   End If
24340                   If blnContinue = True Then
24350                     blnContinue = PreviewOrPrint_NS("2", THIS_PROC & " - R2", acViewNormal)  ' ** Function: Below.
24360                     If blnContinue = True Then
24370                       blnContinue = PreviewOrPrint_NS("1", THIS_PROC & " - R1", acViewNormal)  ' ** Function: Below.
24380                       If blnContinue = True Then
24390                         If gblnUseReveuneExpenseCodes = True Then
24400                           blnContinue = PreviewOrPrint_NS("0A", THIS_PROC & " - R0A", acViewNormal)  ' ** Function: Below.
24410                         Else
24420                           blnContinue = PreviewOrPrint_NS("0", THIS_PROC & " - R0", acViewNormal)  ' ** Function: Below.
24430                         End If
24440                       End If
24450                     End If
24460                   End If
24470                 End If
24480               End If
24490             End If
24500           End If
24510         End If
24520       End If
24530     End If

24540   End With

EXITP:
24550   Set qdf = Nothing
24560   Set dbs = Nothing
24570   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24620   intShiftDown = (Shift And acShiftMask) > 0
24630   intAltDown = (Shift And acAltMask) > 0
24640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24660     Select Case intRetVal
          Case vbKeyTab
24670       With Me
24680         intRetVal = 0
24690         .cmdWordAll.SetFocus
24700       End With
24710     Case vbKeyUp
24720       With Me
24730         intRetVal = 0
24740         .cmdPrint10.SetFocus
24750         strLastArrow = vbNullString
24760       End With
24770     Case vbKeyDown
24780       With Me
24790         intRetVal = 0
24800         .cmdPrint00.SetFocus
24810         strLastArrow = vbNullString
24820       End With
24830     Case vbKeyLeft
24840       With Me
24850         intRetVal = 0
24860         If .cmdExcel10.Enabled = True Then
24870           .cmdExcel10.SetFocus
24880         Else
24890           .cmdWord10.SetFocus
24900         End If
24910       End With
24920     Case vbKeyRight
24930       With Me
24940         intRetVal = 0
24950         .cmdWordAll.SetFocus
24960       End With
24970     End Select
24980   End If

        ' ** Shift keys.
24990   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25000     Select Case intRetVal
          Case vbKeyTab
25010       With Me
25020         intRetVal = 0
25030         If .cmdExcel10.Enabled = True Then
25040           .cmdExcel10.SetFocus
25050         Else
25060           .cmdWord10.SetFocus
25070         End If
25080       End With
25090     End Select
25100   End If

EXITP:
25110   KeyCode = intRetVal
25120   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

25210   With Me
25220     If gblnPrintAll = False Then
25230       .cmdPrintAll_box01.Visible = True
25240       .cmdPrintAll_box02.Visible = True
25250       Select Case .chkAssetList
            Case True
25260         .cmdPrintAll_box03.Visible = True
25270       Case False
25280         .cmdPrintAll_box03.Visible = False
25290       End Select
25300       .cmdPrintAll_box04.Visible = True
25310       If blnWordAll_Focus = False Then
25320         .cmdWordAll_box01.Visible = False
25330         .cmdWordAll_box02.Visible = False
25340         .cmdWordAll_box03.Visible = False
25350         .cmdWordAll_box04.Visible = False
25360       End If
25370       If blnExcelAll_Focus = False Then
25380         .cmdExcelAll_box01.Visible = False
25390         .cmdExcelAll_box02.Visible = False
25400         .cmdExcelAll_box03.Visible = False
25410         .cmdExcelAll_box04.Visible = False
25420       End If
25430     End If
25440   End With

EXITP:
25450   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

25510   With Me
25520     .cmdPrintAll_box01.Visible = False
25530     .cmdPrintAll_box02.Visible = False
25540     .cmdPrintAll_box03.Visible = False
25550     .cmdPrintAll_box04.Visible = False
25560     blnPrintAll_Focus = False
25570   End With

EXITP:
25580   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

25610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
25620   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

25710   With Me
25720     blnWordAll_Focus = True
25730     .cmdWordAll_box01.Visible = True
25740     .cmdWordAll_box02.Visible = True
25750     Select Case .chkAssetList
          Case True
25760       .cmdWordAll_box03.Visible = True
25770     Case False
25780       .cmdWordAll_box03.Visible = False
25790     End Select
25800     .cmdWordAll_box04.Visible = True
25810   End With

EXITP:
25820   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

25910   With Me

25920     lngFiles = 0&
25930     ReDim arr_varFile(F_ELEMS, 0)

25940     WordAll_NS Me  ' ** Module Procedure: modCourtReportsNS.

25950   End With

EXITP:
25960   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26020   intShiftDown = (Shift And acShiftMask) > 0
26030   intAltDown = (Shift And acAltMask) > 0
26040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26060     Select Case intRetVal
          Case vbKeyTab
26070       With Me
26080         intRetVal = 0
26090         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
26100           .cmdExcelAll.SetFocus
26110         Else
26120           .cmdClose.SetFocus
26130         End If
26140       End With
26150     Case vbKeyUp
26160       With Me
26170         intRetVal = 0
26180         .cmdWord10.SetFocus
26190         strLastArrow = vbNullString
26200       End With
26210     Case vbKeyDown
26220       With Me
26230         intRetVal = 0
26240         .cmdWord00.SetFocus
26250         strLastArrow = vbNullString
26260       End With
26270     Case vbKeyRight
26280       With Me
26290         intRetVal = 0
26300         If .cmdExcelAll.Enabled = True Then
26310           .cmdExcelAll.SetFocus
26320         Else
26330           .cmdClose.SetFocus
26340         End If
26350       End With
26360     End Select
26370   End If

        ' ** Shift keys.
26380   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26390     Select Case intRetVal
          Case vbKeyTab
26400       With Me
26410         intRetVal = 0
26420         .cmdPrintAll.SetFocus
26430       End With
26440     End Select
26450   End If

        ' ** Ctrl keys.
26460   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
26470     Select Case intRetVal
          Case vbKeyTab
26480       With Me
26490         intRetVal = 0
26500         .chkOpenWord.SetFocus
26510       End With
26520     End Select
26530   End If

EXITP:
26540   KeyCode = intRetVal
26550   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

26610   With Me
26620     .cmdWordAll_box01.Visible = True
26630     .cmdWordAll_box02.Visible = True
26640     Select Case .chkAssetList
          Case True
26650       .cmdWordAll_box03.Visible = True
26660     Case False
26670       .cmdWordAll_box03.Visible = False
26680     End Select
26690     .cmdWordAll_box04.Visible = True
26700     If blnPrintAll_Focus = False Then
26710       .cmdPrintAll_box01.Visible = False
26720       .cmdPrintAll_box02.Visible = False
26730       .cmdPrintAll_box03.Visible = False
26740       .cmdPrintAll_box04.Visible = False
26750     End If
26760     If blnExcelAll_Focus = False Then
26770       .cmdExcelAll_box01.Visible = False
26780       .cmdExcelAll_box02.Visible = False
26790       .cmdExcelAll_box03.Visible = False
26800       .cmdExcelAll_box04.Visible = False
26810     End If
26820   End With

EXITP:
26830   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

26910   With Me
26920     .cmdWordAll_box01.Visible = False
26930     .cmdWordAll_box02.Visible = False
26940     .cmdWordAll_box03.Visible = False
26950     .cmdWordAll_box04.Visible = False
26960     blnWordAll_Focus = False
26970   End With

EXITP:
26980   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub chkOpenWord_AfterUpdate()

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

27010   With Me
27020     Select Case .chkOpenWord
          Case True
27030       .chkOpenWord_lbl.FontBold = True
27040       .chkOpenWord_lbl2.FontBold = True
27050       .chkOpenWord_lbl_dim_hi.FontBold = True
27060       .chkOpenWord_lbl2_dim_hi.FontBold = True
27070     Case False
27080       .chkOpenWord_lbl.FontBold = False
27090       .chkOpenWord_lbl2.FontBold = False
27100       .chkOpenWord_lbl_dim_hi.FontBold = False
27110       .chkOpenWord_lbl2_dim_hi.FontBold = False
27120     End Select
27130   End With

EXITP:
27140   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

27200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27220   intShiftDown = (Shift And acShiftMask) > 0
27230   intAltDown = (Shift And acAltMask) > 0
27240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27260     Select Case intRetVal
          Case vbKeyTab
27270       With Me
27280         intRetVal = 0
27290         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
27300           .cmdExcelAll.SetFocus
27310         Else
27320           .cmdClose.SetFocus
27330         End If
27340       End With
27350     Case vbKeyUp
27360       With Me
27370         intRetVal = 0
27380         .cmdWordAll.SetFocus
27390       End With
27400     Case vbKeyDown
27410       With Me
27420         intRetVal = 0
27430         .cmdWord00.SetFocus
27440       End With
27450     End Select
27460   End If

        ' ** Shift keys.
27470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
27480     Select Case intRetVal
          Case vbKeyTab
27490       With Me
27500         intRetVal = 0
27510         .cmdWordAll.SetFocus
27520       End With
27530     End Select
27540   End If

        ' ** Ctrl keys.
27550   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
27560     Select Case intRetVal
          Case vbKeyTab
27570       With Me
27580         intRetVal = 0
27590         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
27600           .chkOpenExcel.SetFocus
27610         ElseIf .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
27620           .UserReportPath_chk.SetFocus
27630         Else
27640           Beep
27650         End If
27660       End With
27670     End Select
27680   End If

EXITP:
27690   KeyCode = intRetVal
27700   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

27810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27820   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

27910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
27920   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

28000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

28010   With Me
28020     blnExcelAll_Focus = True
28030     .cmdExcelAll_box01.Visible = True
28040     .cmdExcelAll_box02.Visible = True
28050     Select Case .chkAssetList
          Case True
28060       .cmdExcelAll_box03.Visible = True
28070     Case False
28080       .cmdExcelAll_box03.Visible = False
28090     End Select
28100     .cmdExcelAll_box04.Visible = True
28110   End With

EXITP:
28120   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

28210   With Me

28220     lngFiles = 0&
28230     ReDim arr_varFile(F_ELEMS, 0)

28240     ExcelAll_NS Me  ' ** Module Procedure: modCourtReportsNS.

28250   End With

EXITP:
28260   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

28300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
28320   intShiftDown = (Shift And acShiftMask) > 0
28330   intAltDown = (Shift And acAltMask) > 0
28340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28360     Select Case intRetVal
          Case vbKeyTab
28370       With Me
28380         intRetVal = 0
28390         .cmdClose.SetFocus
28400       End With
28410     Case vbKeyUp
28420       With Me
28430         intRetVal = 0
28440         .cmdExcel10.SetFocus
28450         strLastArrow = vbNullString
28460       End With
28470     Case vbKeyDown
28480       With Me
28490         intRetVal = 0
28500         .cmdExcel00.SetFocus
28510         strLastArrow = vbNullString
28520       End With
28530     Case vbKeyRight
28540       With Me
28550         intRetVal = 0
28560         .cmdClose.SetFocus
28570       End With
28580     End Select
28590   End If

        ' ** Shift keys.
28600   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28610     Select Case intRetVal
          Case vbKeyTab
28620       With Me
28630         intRetVal = 0
28640         .cmdWordAll.SetFocus
28650       End With
28660     End Select
28670   End If

        ' ** Ctrl keys.
28680   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
28690     Select Case intRetVal
          Case vbKeyTab
28700       With Me
28710         intRetVal = 0
28720         .chkOpenExcel.SetFocus
28730       End With
28740     End Select
28750   End If

EXITP:
28760   KeyCode = intRetVal
28770   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

28810   With Me
28820     .cmdExcelAll_box01.Visible = True
28830     .cmdExcelAll_box02.Visible = True
28840     Select Case .chkAssetList
          Case True
28850       .cmdExcelAll_box03.Visible = True
28860     Case False
28870       .cmdExcelAll_box03.Visible = False
28880     End Select
28890     .cmdExcelAll_box04.Visible = True
28900     If blnPrintAll_Focus = False Then
28910       .cmdPrintAll_box01.Visible = False
28920       .cmdPrintAll_box02.Visible = False
28930       .cmdPrintAll_box03.Visible = False
28940       .cmdPrintAll_box04.Visible = False
28950     End If
28960     If blnWordAll_Focus = False Then
28970       .cmdWordAll_box01.Visible = False
28980       .cmdWordAll_box02.Visible = False
28990       .cmdWordAll_box03.Visible = False
29000       .cmdWordAll_box04.Visible = False
29010     End If
29020   End With

EXITP:
29030   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

29100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

29110   With Me
29120     .cmdExcelAll_box01.Visible = False
29130     .cmdExcelAll_box02.Visible = False
29140     .cmdExcelAll_box03.Visible = False
29150     .cmdExcelAll_box04.Visible = False
29160     blnExcelAll_Focus = False
29170   End With

EXITP:
29180   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub chkOpenExcel_AfterUpdate()

29200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

29210   With Me
29220     Select Case .chkOpenExcel
          Case True
29230       .chkOpenExcel_lbl.FontBold = True
29240       .chkOpenExcel_lbl2.FontBold = True
29250       .chkOpenExcel_lbl_dim_hi.FontBold = True
29260       .chkOpenExcel_lbl2_dim_hi.FontBold = True
29270     Case False
29280       .chkOpenExcel_lbl.FontBold = False
29290       .chkOpenExcel_lbl2.FontBold = False
29300       .chkOpenExcel_lbl_dim_hi.FontBold = False
29310       .chkOpenExcel_lbl2_dim_hi.FontBold = False
29320     End Select
29330   End With

EXITP:
29340   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub chkOpenExcel_KeyDown(KeyCode As Integer, Shift As Integer)

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
29420   intShiftDown = (Shift And acShiftMask) > 0
29430   intAltDown = (Shift And acAltMask) > 0
29440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29460     Select Case intRetVal
          Case vbKeyTab
29470       With Me
29480         intRetVal = 0
29490         .cmdClose.SetFocus
29500       End With
29510     Case vbKeyUp
29520       With Me
29530         intRetVal = 0
29540         .cmdExcelAll.SetFocus
29550       End With
29560     Case vbKeyDown
29570       With Me
29580         intRetVal = 0
29590         .cmdExcel00.SetFocus
29600       End With
29610     End Select
29620   End If

        ' ** Shift keys.
29630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
29640     Select Case intRetVal
          Case vbKeyTab
29650       With Me
29660         intRetVal = 0
29670         .cmdExcelAll.SetFocus
29680       End With
29690     End Select
29700   End If

        ' ** Ctrl keys.
29710   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
29720     Select Case intRetVal
          Case vbKeyTab
29730       With Me
29740         intRetVal = 0
29750         If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
29760           .UserReportPath_chk.SetFocus
29770         Else
29780           Beep
29790         End If
29800       End With
29810     End Select
29820   End If

EXITP:
29830   KeyCode = intRetVal
29840   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

29910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
29920   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub UserReportPath_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_box_MouseMove"

30010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30020   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub UserReportPath_chk_AfterUpdate()

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_AfterUpdate"

30110   With Me
30120     Select Case .UserReportPath_chk
          Case True
30130       .UserReportPath_chk_lbl1.FontBold = True
30140       .UserReportPath_chk_lbl1_dim.FontBold = True
30150       .UserReportPath_chk_lbl1_dim_hi.FontBold = True
30160       .UserReportPath_chk_lbl2.FontBold = True
30170       .UserReportPath_chk_lbl2_dim.FontBold = True
30180       .UserReportPath_chk_lbl2_dim_hi.FontBold = True
30190     Case False
30200       .UserReportPath_chk_lbl1.FontBold = False
30210       .UserReportPath_chk_lbl1_dim.FontBold = False
30220       .UserReportPath_chk_lbl1_dim_hi.FontBold = False
30230       .UserReportPath_chk_lbl2.FontBold = False
30240       .UserReportPath_chk_lbl2_dim.FontBold = False
30250       .UserReportPath_chk_lbl2_dim_hi.FontBold = False
30260     End Select
30270   End With

EXITP:
30280   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub UserReportPath_chk_KeyDown(KeyCode As Integer, Shift As Integer)

30300 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30320   intShiftDown = (Shift And acShiftMask) > 0
30330   intAltDown = (Shift And acAltMask) > 0
30340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30370       With Me
30380         intRetVal = 0
30390         .cmdClose.SetFocus
30400       End With
30410     End Select
30420   End If

        ' ** Shift keys.
30430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
30440     Select Case intRetVal
          Case vbKeyTab
30450       With Me
30460         intRetVal = 0
30470         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
30480           .cmdExcelAll.SetFocus
30490         Else
30500           .cmdWordAll.SetFocus
30510         End If
30520       End With
30530     End Select
30540   End If

        ' ** Ctrl-Shift keys.
30550   If intCtrlDown And (Not intAltDown) And intShiftDown Then
30560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
30570       With Me
30580         intRetVal = 0
30590         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
30600           .chkOpenExcel.SetFocus
30610         Else
30620           .chkOpenWord.SetFocus
30630         End If
30640       End With
30650     End Select
30660   End If

EXITP:
30670   KeyCode = intRetVal
30680   Exit Sub

ERRH:
1380    intRetVal = 0
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Private Sub UserReportPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

30700 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_MouseMove"

30710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
30720   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Function PreviewOrPrint_NS(strReportNumber As String, strControlName As String, intView As Integer) As Boolean

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "PreviewOrPrint_NS"

        Dim strDocName As String
        Dim blnRetVal As Boolean
        Dim intRetVal_BuildCourtReportData As Integer

30810   blnRetVal = True

30820   DoCmd.Hourglass True
30830   DoEvents

30840   With Me

          '  ** Validate the dates.
30850     If Validate = True Then  ' ** Function: Below.

30860       ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

            ' ** Set global variables for report headers.
30870       gdatStartDate = .DateStart
30880       gdatEndDate = .DateEnd
30890       gstrAccountNo = .cmbAccounts.Column(0)
30900       gstrAccountName = .cmbAccounts.Column(3)

            ' ** Build a new summary report table.
30910       intRetVal_BuildCourtReportData = NSBuildCourtReportData(strReportNumber)  ' ** Module Function: modCourtReportsNS.
30920       If intRetVal_BuildCourtReportData = 0 Then
30930         blnRebuildTable = False
30940         strDocName = "rptCourtRptNS_0" & strReportNumber
30950         DoCmd.OpenReport strDocName, intView
30960         If intView = acViewPreview Then
30970           DoCmd.Maximize
30980           DoCmd.RunCommand acCmdFitToWindow
30990         End If
31000       Else
              ' ** Return Codes:
              ' **   0  Success.
              ' **  -1  Canceled.
              ' **  -9  Error.
31010         blnRetVal = False
31020       End If

31030     Else
31040       blnRetVal = False
31050     End If

31060   End With  ' ** Me.

31070   DoCmd.Hourglass False

EXITP:
        ' ** Check some Public variables in case there's been an error.
31080   blnRetVal = CoOptions_Read  ' ** Module Function: modStartupFuncs.
31090   blnRetVal = IniFile_GetDataLoc  ' ** Module Procedure: modStartupFuncs.
31100   PreviewOrPrint_NS = blnRetVal
31110   Exit Function

ERRH:
1380    blnRetVal = False
1390    DoCmd.Hourglass False
1400    Select Case ERR.Number
        Case 2202  ' ** You must set a default printer before you design, print, or preview.
1410      MsgBox "There does not appear to be a default printer defined." & vbCrLf & vbCrLf & _
            "Please define a default printer before trying to print this report.", vbCritical + vbOKOnly, "Printer Not Found"
1420    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
1430    Case Else
1440      THAT_PROC = THIS_PROC
1450      That_Erl = Erl
1460      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1470    End Select
1480    Resume EXITP

End Function

Public Sub SendToFile_NS(strReportNumber As String, strControlName As String, Optional varExcel As Variant)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "SendToFile_NS"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strQry As String, strMacro As String
        Dim strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String
        Dim blnExcel As Boolean, blnNoData As Boolean
        Dim blnContinue As Boolean, blnUseSavedPath As Boolean, blnAutoStart As Boolean
        Dim msgResponse As VbMsgBoxResult
        Dim lngRecs As Long
        Dim varTmp00 As Variant
        Dim lngX As Long, lngE As Long
        Dim intRetVal_BuildCourtReportData As Integer

31210   blnContinue = True
31220   blnUseSavedPath = False

31230   With Me

31240     DoCmd.Hourglass True
31250     DoEvents

31260     If IsMissing(varExcel) = True Then
31270       blnExcel = False
31280     Else
31290       blnExcel = varExcel
31300     End If

31310     If blnExcel = True Then
31320       If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
              ' ** It seems like it's not quite closed when it gets here,
              ' ** because if I stop the code and run the function again,
              ' ** it always comes up False.
31330         ForcePause 2  ' ** Module Function: modCodeUtilities.
31340         If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                ' ** It seems like it's not quite closed when it gets here,
                ' ** because if I stop the code and run the function again,
                ' ** it always comes up False.
31350           ForcePause 2  ' ** Module Function: modCodeUtilities.
31360           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
31370             DoCmd.Hourglass False
31380             msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
                    "In order for Trust Accountant to reliably export your report," & vbCrLf & _
                    "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
                    "You may close Excel before proceding, then click Retry." & vbCrLf & _
                    "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
                  ' ** ... Otherwise Trust Accountant will do it for you.
31390             If msgResponse <> vbRetry Then
31400               blnContinue = False
31410               blnAllCancel = True
31420               AllCancelSet2_NS blnAllCancel  ' ** Module Procedure: modCourtReportsNS.
31430             End If
31440           End If
31450         End If
31460       End If
31470     End If

31480     If blnContinue = True Then

31490       DoCmd.Hourglass True
31500       DoEvents

31510       intRetVal_BuildCourtReportData = 0

            ' ** Validate the dates.
31520       If Validate = True Then  ' ** Function: Below.

31530         ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

31540         gstrFormQuerySpec = THIS_NAME

              ' ** Set global variables for report headers.
31550         gdatStartDate = .DateStart.Value
31560         gdatEndDate = .DateEnd.Value
31570         gstrAccountNo = .cmbAccounts.Column(0)
31580         gstrAccountName = .cmbAccounts.Column(3)

31590         Set dbs = CurrentDb
31600         With dbs
                ' ** tblReport, captions of Court Reports, by specified [CrtTyp].
31610           Set qdf = .QueryDefs("qryCourtReport_15")
31620           With qdf.Parameters
31630             ![CrtTyp] = "NS"
31640           End With
31650           Set rst = qdf.OpenRecordset
31660           With rst
31670             .MoveLast
31680             lngCaps = .RecordCount
31690             .MoveFirst
31700             arr_varCap = .GetRows(lngCaps)
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
31710             .Close
31720           End With
31730           .Close
31740         End With

31750         If IsNull(.UserReportPath) = False Then
31760           If .UserReportPath <> vbNullString Then
31770             If .UserReportPath_chk = True Then
31780               If DirExists(.UserReportPath) = True Then  ' ** Module Function: modFileUtilities.
31790                 blnUseSavedPath = True
31800               End If
31810             End If
31820           End If
31830         End If

              ' ** Build a new summary report table.
31840         intRetVal_BuildCourtReportData = NSBuildCourtReportData(strReportNumber)
31850         If intRetVal_BuildCourtReportData = 0 Then
31860           blnRebuildTable = False
31870           If blnExcel = True Then

31880             Set dbs = CurrentDb

                  ' ** Empty tmpCourtReportData2.
31890             Set qdf = dbs.QueryDefs("qryCourtReport_NS_01_01")
31900             qdf.Execute

31910             Select Case strReportNumber

                  Case "0", "0A"
                    ' ** 0 Summary of Account.

                    ' ** tmpCourtReportData
                    ' **   "date >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And date < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#"
                    ' **     " and accountno = '" & gstrAccountNo & "'"

                    ' ** Empty tmpCourtReportData7.
31920               Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_09")
31930               qdf.Execute

                    ' ** Append qryCourtReport_NS_00_08 to tmpCourtReportData7.
31940               Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_10")
31950               With qdf.Parameters
31960                 ![actno] = gstrAccountNo
31970                 ![datbeg] = gdatStartDate
31980                 ![datEnd] = gdatEndDate
31990               End With
32000               qdf.Execute

32010               Set rst = dbs.OpenRecordset("tmpCourtReportData7", dbOpenDynaset, dbConsistent)
32020               With rst
32030                 .MoveLast
32040                 lngRecs = .RecordCount
32050                 .MoveFirst
                      ' ** Set the initial sorting order.
32060                 For lngX = 1& To lngRecs
32070                   .Edit
32080                   ![sort] = CSng(lngX)
32090                   .Update
32100                   If lngX < lngRecs Then .MoveNext
32110                 Next
32120                 .Close
32130               End With

                    ' ** line 18:
                    ' ** grand total    : Div 60.5 Grp 50.8 Rpt 90.8
                    ' ** Max(tmp_id) where
                    ' ** Acquisition Value is not Null
32140               Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_15")
32150               qdf.Execute

                    ' ** line 4:
                    ' ** Grp 10's total :          Grp 10.3 Rpt 20.3
                    ' ** 20  10  20 IsNull(Category) = True
                    ' ** line 7:
                    ' ** Grp 20's total :          Grp 20.3 Rpt 40.3
                    ' ** 20  20  40 IsNull(Category) = True
                    ' ** line 13:
                    ' ** Grp 40's total :          Grp 40.3 Rpt 70.3
                    ' ** 60  40  70 IsNull(Category) = True
                    ' ** line 16:
                    ' ** Grp 50's total :          Grp 50.3 Rpt 90.3
                    ' ** 60  50  90 IsNull(Category) = True
32160               Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_17")
32170               qdf.Execute

                    ' ** line 8:
                    ' ** Div 20's total : Div 20.3 Grp 20.5 Rpt 40.5
                    ' ** 20  20  40 Max(tmp_id) for Div = 20
                    ' ** line 17:
                    ' ** Div 60's total : Div 60.3 Grp 50.5 Rpt 90.5
                    ' ** 60  50  90 Max(tmp_id) for Div = 60 not including Grand Total
32180               Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_20")
32190               qdf.Execute

                    ' ** line 6  is Rpt 40
                    ' ** line 10 is Rpt 60

32200               Set rst = dbs.OpenRecordset("tmpCourtReportData7", dbOpenDynaset, dbConsistent)
32210               With rst
                      ' ** Now add blank lines.
32220                 .AddNew
32230                 ![sort] = 4.5!
32240                 ![Report Division] = 20!
32250                 ![Report Group] = 10.3!
32260                 ![Report Number] = 20.4!
32270                 .Update
32280                 .AddNew
32290                 ![sort] = 7.5!
32300                 ![Report Division] = 20!
32310                 ![Report Group] = 20.3!
32320                 ![Report Number] = 40.4!
32330                 .Update
32340                 .AddNew
32350                 ![sort] = 8.5!
32360                 ![Report Division] = 20.3!
32370                 ![Report Group] = 20.5!
32380                 ![Report Number] = 40.6!
32390                 .Update
32400                 .AddNew
32410                 ![sort] = 10.5!
32420                 ![Report Division] = 40!
32430                 ![Report Group] = 30!
32440                 ![Report Number] = 60.1!
32450                 .Update
32460                 .AddNew
32470                 ![sort] = 13.5!
32480                 ![Report Division] = 60!
32490                 ![Report Group] = 40.3!
32500                 ![Report Number] = 70.4!
32510                 .Update
32520                 .AddNew
32530                 ![sort] = 16.5!
32540                 ![Report Division] = 60!
32550                 ![Report Group] = 50.3!
32560                 ![Report Number] = 90.4!
32570                 .Update
32580                 .AddNew
32590                 ![sort] = 17.5!
32600                 ![Report Division] = 60.3!
32610                 ![Report Group] = 50.5!
32620                 ![Report Number] = 90.6!
32630                 .Update
32640                 .Close
32650               End With

32660               blnNoData = False
32670               Select Case strReportNumber
                    Case "0"
                      ' ** Summary: For export.
32680                 strQry = "qryCourtReport_NS_00_11"

32690               Case "0A"
                      'Disbursements of Principal, by revcode_DESC
                      'Div 20, Grp 20, Rpt 30
                      'Disbursements of Income, by revcode_DESC
                      'Div 60, Grp 50, Rpt 80
                      ' ** Append qryCourtReport_NS_00_A_04 (Principal/Income detail entries, grouped and summed) to tmpCourtReportData7.
32700                 Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_A_05")
32710                 With qdf.Parameters
32720                   ![actno] = gstrAccountNo
32730                   ![datbeg] = gdatStartDate
32740                   ![datEnd] = gdatEndDate
32750                 End With
32760                 qdf.Execute
                      ' ** Append Disbursements of Principal/Income total lines to tmpCourtReportData7.
32770                 Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_A_06")
32780                 qdf.Execute
                      ' ** Update tmpCourtReportData7, Disbursements of Principal/Income header line, [Acquisition Value] = Null.
32790                 Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_A_07")
32800                 qdf.Execute
                      ' ** Summary: For export; grouped by Inc/Exp Codes.
32810                 strQry = "qryCourtReport_NS_00_A_08"

32820               End Select

32830             Case "1", "1A"
                    ' ** 1 Receipts of Principal.

                    ' ** qryCourtReport-A2
                    ' **   "transdate >= #" & Format(gdatStartDate,"mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#"
                    ' **     " and accountno = '" & gstrAccountNo & "'"
                    ' **     " and ((journaltype = 'Received' and pcash > 0) "
                    ' **     " or (journaltype = 'Misc.' and pcash > 0) "
                    ' **     " or (journaltype = 'Cost Adj.' and cost > 0) "
                    ' **     " or (journaltype = 'Deposit' and not(jcomment like '*stock split*')))"

                    ' ** Append qryCourtReport_NS_01_03 to tmpCourtReportData2.
32840               Set qdf = dbs.QueryDefs("qryCourtReport_NS_01_04")
32850               With qdf.Parameters
32860                 ![actno] = gstrAccountNo
32870                 ![datbeg] = gdatStartDate
32880                 ![datEnd] = gdatEndDate
32890               End With
32900               qdf.Execute
32910               dbs.Close

32920               Select Case strReportNumber
                    Case "1"
                      ' ** Receipts of Principal: For export.
32930                 strQry = "qryCourtReport_NS_01_12"
32940                 varTmp00 = DCount("*", "qryCourtReport_NS_01_11a")
32950                 If IsNull(varTmp00) = True Then
32960                   blnNoData = True
32970                   strQry = "qryCourtReport_NS_01_15"
32980                 Else
32990                   If varTmp00 = 0 Then
33000                     blnNoData = True
33010                     strQry = "qryCourtReport_NS_01_15"
33020                   End If
33030                 End If
33040               Case "1A"
                      ' ** Receipts of Principal, Inc/Exp: For Export.
33050                 strQry = "qryCourtReport_NS_01_28"
33060                 varTmp00 = DCount("*", "qryCourtReport_NS_01_19")
33070                 If IsNull(varTmp00) = True Then
33080                   blnNoData = True
33090                   strQry = "qryCourtReport_NS_01_31"
33100                 Else
33110                   If varTmp00 = 0 Then
33120                     blnNoData = True
33130                     strQry = "qryCourtReport_NS_01_31"
33140                   End If
33150                 End If
33160               End Select

33170             Case "2"
                    ' ** 2 Gains (Losses) on Sale or Other Dispositions.

                    ' ** qryCourtReport-B
                    ' **   "transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#"
                    ' **     " and accountno = '" & gstrAccountNo & "'"
                    ' **     " and (journaltype = 'Sold' and GainLoss <> 0)"

                    ' ** Append qryCourtReport_NS_02_02 to tmpCourtReportData2.
33180               Set qdf = dbs.QueryDefs("qryCourtReport_NS_02_03")
33190               With qdf.Parameters
33200                 ![actno] = gstrAccountNo
33210                 ![datbeg] = gdatStartDate
33220                 ![datEnd] = gdatEndDate
33230               End With
33240               qdf.Execute
33250               dbs.Close

                    ' ** Gains/Losses on Sale: For export.
33260               strQry = "qryCourtReport_NS_02_11"
33270               varTmp00 = DCount("*", "qryCourtReport_NS_02_10a")
33280               If IsNull(varTmp00) = True Then
33290                 blnNoData = True
33300                 strQry = "qryCourtReport_NS_02_14"
33310               Else
33320                 If varTmp00 = 0 Then
33330                   blnNoData = True
33340                   strQry = "qryCourtReport_NS_02_14"
33350                 End If
33360               End If

33370             Case "3", "3A"
                    ' ** 3 Disbursements of Principal.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# AND transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#)"
                    ' **     " AND accountno = '" & gstrAccountNo & "'"
                    ' **     " AND ((journaltype = 'Paid' AND pcash <> 0 and taxcode <> 11)"  '<> "Distribution"
                    ' **     " OR (journaltype = 'Misc.' AND pcash < 0 )"
                    ' **     " OR (journaltype = 'Cost Adj.' AND cost < 0) "
                    ' **     " OR (journaltype = 'Withdrawn' AND taxcode <> 11))"  '<> "Distribution"

                    ' ** Append qryCourtReport_NS_03_02 to tmpCourtReportData2.
33380               Set qdf = dbs.QueryDefs("qryCourtReport_NS_03_03")    '####  TAXCODE  ####
33390               With qdf.Parameters
33400                 ![actno] = gstrAccountNo
33410                 ![datbeg] = gdatStartDate
33420                 ![datEnd] = gdatEndDate
33430               End With
33440               qdf.Execute
33450               dbs.Close

33460               Select Case strReportNumber
                    Case "3"
                      ' ** Disbursements of Principal: For export.
33470                 strQry = "qryCourtReport_NS_03_15"
33480                 varTmp00 = DCount("*", "qryCourtReport_NS_03_14a")
33490                 If IsNull(varTmp00) = True Then
33500                   blnNoData = True
33510                   strQry = "qryCourtReport_NS_03_15c"
33520                 Else
33530                   If varTmp00 = 0 Then
33540                     blnNoData = True
33550                     strQry = "qryCourtReport_NS_03_15c"
33560                   End If
33570                 End If

33580               Case "3A"
                      ' ** Disbursements of Principal: For export, grouped by Inc/Exp Codes.
33590                 strQry = "qryCourtReport_NS_03_17"
33600                 varTmp00 = DCount("*", "qryCourtReport_NS_03_16a")
33610                 If IsNull(varTmp00) = True Then
33620                   blnNoData = True
33630                   strQry = "qryCourtReport_NS_03_20"
33640                 Else
33650                   If varTmp00 = 0 Then
33660                     blnNoData = True
33670                     strQry = "qryCourtReport_NS_03_20"
33680                   End If
33690                 End If

33700               End Select

33710             Case "4"
                    ' ** 4 Distributions of Principal to Beneficiaries.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#)"
                    ' **     " and accountno = '" & gstrAccountNo & "'"
                    ' **     " and ((journaltype = 'Paid' and pcash <> 0 and taxcode = 11)"  '= "Distribution"
                    ' **     " or (journaltype = 'Withdrawn' and taxcode = 11))"  '= "Distribution"

                    ' ** Append qryCourtReport_NS_04_02 to tmpCourtReportData2.
33720               Set qdf = dbs.QueryDefs("qryCourtReport_NS_04_03")
33730               With qdf.Parameters
33740                 ![actno] = gstrAccountNo
33750                 ![datbeg] = gdatStartDate
33760                 ![datEnd] = gdatEndDate
33770               End With
33780               qdf.Execute
33790               dbs.Close

                    ' ** Distributions of Principal: For export.
33800               strQry = "qryCourtReport_NS_04_11"
33810               varTmp00 = DCount("*", "qryCourtReport_NS_04_10a")
33820               If IsNull(varTmp00) = True Then
33830                 blnNoData = True
33840                 strQry = "qryCourtReport_NS_04_14"
33850               Else
33860                 If varTmp00 = 0 Then
33870                   blnNoData = True
33880                   strQry = "qryCourtReport_NS_04_14"
33890                 End If
33900               End If

33910             Case "5"
                    ' ** 5 Information For Investments Made.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "(journaltype = 'Purchase')"

                    ' ** Append qryCourtReport_NS_05_02 to tmpCourtReportData2.
33920               Set qdf = dbs.QueryDefs("qryCourtReport_NS_05_03")
33930               With qdf.Parameters
33940                 ![actno] = gstrAccountNo
33950                 ![datbeg] = gdatStartDate
33960                 ![datEnd] = gdatEndDate
33970               End With
33980               qdf.Execute
33990               dbs.Close

                    ' ** Info For Investers: For export.
34000               strQry = "qryCourtReport_NS_05_11"
34010               varTmp00 = DCount("*", "qryCourtReport_NS_05_10a")
34020               If IsNull(varTmp00) = True Then
34030                 blnNoData = True
34040                 strQry = "qryCourtReport_NS_05_14"
34050               Else
34060                 If varTmp00 = 0 Then
34070                   blnNoData = True
34080                   strQry = "qryCourtReport_NS_05_14"
34090                 End If
34100               End If

34110             Case "6"
                    ' ** 6 Change in Investment Holdings.

                    ' ** qryCourtReport-A3
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((journaltype = 'Sold' And GainLoss = 0) Or "
                    ' **     "(journaltype = 'Deposit' And jcomment Like '*stock split*') Or "
                    ' **     "(journaltype = 'Liability'))"

                    ' ** Append qryCourtReport_NS_06_03 to tmpCourtReportData2.
34120               Set qdf = dbs.QueryDefs("qryCourtReport_NS_06_04")
34130               With qdf.Parameters
34140                 ![actno] = gstrAccountNo
34150                 ![datbeg] = gdatStartDate
34160                 ![datEnd] = gdatEndDate
34170               End With
34180               qdf.Execute
34190               dbs.Close

                    ' ** Change in Investments: For export.
34200               strQry = "qryCourtReport_NS_06_12"
34210               varTmp00 = DCount("*", "qryCourtReport_NS_06_11a")
34220               If IsNull(varTmp00) = True Then
34230                 blnNoData = True
34240                 strQry = "qryCourtReport_NS_06_15"
34250               Else
34260                 If varTmp00 = 0 Then
34270                   blnNoData = True
34280                   strQry = "qryCourtReport_NS_06_15"
34290                 End If
34300               End If

34310             Case "7", "7A"
                    ' ** 7 Receipts of Income.

                    ' ** qryCourtReport - Receipts of Income 1
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate < #" & Format((gdatEndDate + 1), "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((icash > 0) Or "
                    ' **     "(journaltype = 'Purchase' And icash < 0 And pcash = -cost))"

                    ' ** Append qryCourtReport_NS_07_02 to tmpCourtReportData2.
34320               Set qdf = dbs.QueryDefs("qryCourtReport_NS_07_03")
34330               With qdf.Parameters
34340                 ![actno] = gstrAccountNo
34350                 ![datbeg] = gdatStartDate
34360                 ![datEnd] = gdatEndDate
34370               End With
34380               qdf.Execute
34390               dbs.Close

34400               Select Case strReportNumber
                    Case "7"
                      ' ** Receipts of Income: For export.
34410                 strQry = "qryCourtReport_NS_07_16"
34420                 varTmp00 = DCount("*", "qryCourtReport_NS_07_15a")
34430                 If IsNull(varTmp00) = True Then
34440                   blnNoData = True
34450                   strQry = "qryCourtReport_NS_07_19"
34460                 Else
34470                   If varTmp00 = 0 Then
34480                     blnNoData = True
34490                     strQry = "qryCourtReport_NS_07_19"
34500                   End If
34510                 End If
34520               Case "7A"
                      ' ** Receipts of Income, Inc/Exp: For Export.
34530                 strQry = "qryCourtReport_NS_07_34"
34540                 varTmp00 = DCount("*", "qryCourtReport_NS_07_25")
34550                 If IsNull(varTmp00) = True Then
34560                   blnNoData = True
34570                   strQry = "qryCourtReport_NS_07_37"
34580                 Else
34590                   If varTmp00 = 0 Then
34600                     blnNoData = True
34610                     strQry = "qryCourtReport_NS_07_37"
34620                   End If
34630                 End If
34640               End Select

34650             Case "8", "8A"
                    ' ** 8 Disbursements of Income.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((journaltype = 'Paid' And icash <> 0 And taxcode <> 11) Or "  '<> "Distribution"
                    ' **     "(journaltype = 'Misc.' And icash < 0) Or "
                    ' **     "(journaltype = 'Liability' And icash < 0))"

                    ' ** Append qryCourtReport_NS_08_02 to tmpCourtReportData2.
34660               Set qdf = dbs.QueryDefs("qryCourtReport_NS_08_03")    '####  TAXCODE  ####
34670               With qdf.Parameters
34680                 ![actno] = gstrAccountNo
34690                 ![datbeg] = gdatStartDate
34700                 ![datEnd] = gdatEndDate
34710               End With
34720               qdf.Execute
34730               dbs.Close

34740               Select Case strReportNumber
                    Case "8"
                      ' ** Disbursements of Income: For export.
34750                 strQry = "qryCourtReport_NS_08_15"
34760                 varTmp00 = DCount("*", "qryCourtReport_NS_08_14a")
34770                 If IsNull(varTmp00) = True Then
34780                   blnNoData = True
34790                   strQry = "qryCourtReport_NS_08_15c"
34800                 Else
34810                   If varTmp00 = 0 Then
34820                     blnNoData = True
34830                     strQry = "qryCourtReport_NS_08_15c"
34840                   End If
34850                 End If

34860               Case "8A"
                      ' ** Disbursements of Income: For export; grouped by Inc/Exp Codes.
34870                 strQry = "qryCourtReport_NS_08_17"
34880                 varTmp00 = DCount("*", "qryCourtReport_NS_08_16a")
34890                 If IsNull(varTmp00) = True Then
34900                   blnNoData = True
34910                   strQry = "qryCourtReport_NS_08_20"
34920                 Else
34930                   If varTmp00 = 0 Then
34940                     blnNoData = True
34950                     strQry = "qryCourtReport_NS_08_20"
34960                   End If
34970                 End If

34980               End Select

34990             Case "9"
                    ' ** 9 Distributions of Income.

                    ' ** qryCourtReport-A2
                    ' **   "(transdate >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And transdate <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#) And "
                    ' **     "accountno = '" & gstrAccountNo & "' And "
                    ' **     "((journaltype = 'Paid' And icash <> 0 And taxcode = 11))"  '= "Distribution"

                    ' ** Append qryCourtReport_NS_09_02 to tmpCourtReportData2.
35000               Set qdf = dbs.QueryDefs("qryCourtReport_NS_09_03")    '####  TAXCODE  ####
35010               With qdf.Parameters
35020                 ![actno] = gstrAccountNo
35030                 ![datbeg] = gdatStartDate
35040                 ![datEnd] = gdatEndDate
35050               End With
35060               qdf.Execute
35070               dbs.Close

                    ' ** Distributions of Income: For export.
35080               strQry = "qryCourtReport_NS_09_11"
35090               varTmp00 = DCount("*", "qryCourtReport_NS_09_10a")
35100               If IsNull(varTmp00) = True Then
35110                 blnNoData = True
35120                 strQry = "qryCourtReport_NS_09_14"
35130               Else
35140                 If varTmp00 = 0 Then
35150                   blnNoData = True
35160                   strQry = "qryCourtReport_NS_09_14"
35170                 End If
35180               End If

35190             End Select

35200             strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
35210             strRptPath = .UserReportPath

35220             If Len(strReportNumber) = 1 Then
35230               strRptName = ("rptCourtRptNS_0" & strReportNumber)
35240             ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
35250               strRptName = ("rptCourtRptNS_" & strReportNumber)  ' ** (Though no NS reports go up to 10.)
35260             Else
35270               strRptName = ("rptCourtRptNS_0" & strReportNumber)
35280             End If

35290             strMacro = "mcrExcelExport_CR_NS" & Mid(strRptName, InStr(strRptName, "_"))
35300             If blnNoData = True Then
35310               strMacro = strMacro & "_nd"
35320             End If

35330             For lngX = 0& To (lngCaps - 1&)
35340               If arr_varCap(C_RNAM, lngX) = strRptName Then
35350                 strRptCap = arr_varCap(C_CAPN, lngX)
35360                 Exit For
35370               End If
35380             Next

35390             Select Case blnUseSavedPath
                  Case True
35400               strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".xls"
35410             Case False
35420               DoCmd.Hourglass False
35430               strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
35440             End Select

35450             If strRptPathFile <> vbNullString Then
35460               DoCmd.Hourglass True
35470               DoEvents
35480               Select Case blnExcel
                    Case True
35490                 blnAutoStart = .chkOpenExcel
35500               Case False
35510                 blnAutoStart = .chkOpenWord
35520               End Select
35530               If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
35540               If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
35550                 Kill strRptPathFile
35560               End If
                    ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
                    ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
35570               DoCmd.RunMacro strMacro
                    ' ** The macro specifies the query in strQry, but cannot be given a dynamic file name.
                    ' ** So, it's exported to 'CourtReport_CA_xxx.xls', which is then renamed.
35580               If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NS_xxx.xls") = True Or _
                        FileExists(strRptPath & LNK_SEP & "CourtReport_NS_xxx.xls") = True Then
35590                 If FileExists(CurrentAppPath & LNK_SEP & "CourtReport_NS_xxx.xls") = True Then
35600                   Name (CurrentAppPath & LNK_SEP & "CourtReport_NS_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
35610                 Else
35620                   Name (strRptPath & LNK_SEP & "CourtReport_NS_xxx.xls") As (strRptPathFile)
                        ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
35630                 End If
35640                 DoEvents
35650                 If Excel_Court(strRptPathFile) = True Then  ' ** Module Function: modExcelFuncs.
35660                   DoEvents
35670                   Select Case gblnPrintAll
                        Case True
35680                     lngFiles = lngFiles + 1&
35690                     lngE = lngFiles - 1&
35700                     ReDim Preserve arr_varFile(F_ELEMS, lngE)
35710                     arr_varFile(F_RNAM, lngE) = strRptName
35720                     arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
35730                     arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
35740                     FileArraySet_NS arr_varFile  ' ** Module Procedure: modCourtReportsNS.
35750                   Case False
35760                     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
35770                       EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
35780                     End If
35790                     DoEvents
35800                     If blnAutoStart = True Then
35810                       OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
35820                     End If
35830                   End Select
                        'If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
                        '  EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
                        'End If
                        'DoEvents
                        'OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
35840                 End If
35850               End If
35860               strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
35870               If strRptPath <> .UserReportPath Then
35880                 .UserReportPath = strRptPath
35890                 SetUserReportPath_NS Me  ' ** Module Procedure: modCourtReportsNS.
35900               End If
35910             Else
35920               blnContinue = False
35930             End If

35940           Else

35950             strRptName = vbNullString: strRptCap = vbNullString: strRptPathFile = vbNullString
35960             strRptPath = .UserReportPath

35970             If Len(strReportNumber) = 1 Then
35980               strRptName = ("rptCourtRptNS_0" & strReportNumber)
35990             ElseIf IsNumeric(Mid(strReportNumber, 2, 1)) Then
36000               strRptName = ("rptCourtRptNS_" & strReportNumber)  ' ** (Though no NS reports go up to 10.)
36010             Else
36020               strRptName = ("rptCourtRptNS_0" & strReportNumber)
36030             End If

36040             For lngX = 0& To (lngCaps - 1&)
36050               If arr_varCap(C_RNAM, lngX) = strRptName Then
36060                 strRptCap = arr_varCap(C_CAPN, lngX)
36070                 Exit For
36080               End If
36090             Next

36100             Select Case blnUseSavedPath
                  Case True
36110               strRptPathFile = .UserReportPath & LNK_SEP & strRptCap & ".rtf"
36120             Case False
36130               DoCmd.Hourglass False
36140               strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
36150             End Select

36160             If strRptPathFile <> vbNullString Then
36170               DoCmd.Hourglass True
36180               DoEvents
36190               Select Case blnExcel
                    Case True
36200                 blnAutoStart = .chkOpenExcel
36210               Case False
36220                 blnAutoStart = .chkOpenWord
36230               End Select
36240               If gblnPrintAll = True Then blnAutoStart = False  ' ** They'll open only after all have been exported.
36250               If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
36260                 Kill strRptPathFile
36270               End If
36280               Select Case gblnPrintAll
                    Case True
36290                 lngFiles = lngFiles + 1&
36300                 lngE = lngFiles - 1&
36310                 ReDim Preserve arr_varFile(F_ELEMS, lngE)
36320                 arr_varFile(F_RNAM, lngE) = strRptName
36330                 arr_varFile(F_FILE, lngE) = Parse_File(strRptPathFile)  ' ** Module Function: modFileUtilities.
36340                 arr_varFile(F_PATH, lngE) = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
36350                 FileArraySet_NS arr_varFile  ' ** Module Procedure: modCourtReportsNS.
36360                 DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, False
36370               Case False
36380                 DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, blnAutoStart
36390               End Select
                    'DoCmd.OutputTo acOutputReport, strRptName, acFormatRTF, strRptPathFile, True
36400               strRptPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
36410               If strRptPath <> .UserReportPath Then
36420                 .UserReportPath = strRptPath
36430                 SetUserReportPath_NS Me  ' ** Module Procedure: modCourtReportsNS.
36440               End If
36450             Else
36460               blnContinue = False
36470             End If

36480           End If
36490         Else
                ' ** Return Codes:
                ' **   0  Success.
                ' **  -1  Canceled.
                ' **  -9  Error.
36500           blnContinue = False
36510         End If  ' ** intRetVal_BuildCourtReportData.

36520       End If

36530     End If  ' ** blnContinue.

36540   End With

36550   DoCmd.Hourglass False

EXITP:
36560   Set rst = Nothing
36570   Set qdf = Nothing
36580   Set dbs = Nothing
36590   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    Select Case ERR.Number
        Case 70  ' ** Permission denied.
1400      Beep
1410      MsgBox "Trust Accountant is unable to save the file." & vbCrLf & vbCrLf & _
            "If the program to which you're exporting is open," & vbCrLf & _
            "please close it and try again.", vbInformation + vbOKOnly, "Save Failed"
1420    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
1430    Case Else
1440      THAT_PROC = THIS_PROC
1450      That_Erl = Erl
1460      Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1470    End Select
1480    Resume EXITP

End Sub

Private Sub AssetList_PreviewPrint(strRptPeriod As String, intView As Integer)

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "AssetList_PreviewPrint"

        Dim strRptSuffix As String
        Dim blnContinue As Boolean
        Dim intRetVal_BuildAssetListInfo As Integer

36610   With Me

36620     DoCmd.Hourglass True
36630     DoEvents

36640     If Validate = True Then  ' ** Function: Below.

36650       strRptSuffix = vbNullString
36660       intRetVal_BuildAssetListInfo = 0

36670       Select Case strRptPeriod
            Case "Beginning"
36680         gdatStartDate = .DateStart
36690         gdatEndDate = .DateEnd
36700         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NS("01/01/1900", (.DateStart - 1), strRptPeriod, strRptSuffix, THIS_PROC)  ' ** Function: Below.
36710       Case "Ending"
36720         gdatStartDate = .DateStart
36730         gdatEndDate = .DateEnd
36740         intRetVal_BuildAssetListInfo = BuildAssetListInfo_NS(.DateStart, .DateEnd, strRptPeriod, strRptSuffix, THIS_PROC)  ' ** Function: Below.
36750       End Select

36760       If intRetVal_BuildAssetListInfo = 0 Then
              ' ** strRptSuffix should return among "_00B", "_00BA", "_00D", or "_00DA".
36770         If strRptSuffix <> vbNullString Then
36780           gstrAccountNo = .cmbAccounts.Column(0)
36790           gstrAccountName = .cmbAccounts.Column(3)
36800           DoCmd.OpenReport "rptCourtRptNS" & strRptSuffix, intView
36810           If intView = acViewPreview Then
36820             DoCmd.Maximize
36830             DoCmd.RunCommand acCmdFitToWindow
36840           End If
36850         Else
36860           MsgBox "Problem assembling Asset List.", vbInformation + vbOKOnly, "Asset List Error"
36870         End If
36880       Else
              ' ** Return codes:
              ' **    0  Success.
              ' **   -2  No data.
              ' **   -3  Missing entry, e.g., date.
              ' **   -9  Error.
36890         blnContinue = False
36900         If intRetVal_BuildAssetListInfo = -2 Then
36910           Select Case strRptPeriod
                Case "Beginning"
36920             strRptSuffix = "rptCourtRptNS_00B"
36930           Case "Ending"
36940             strRptSuffix = "rptCourtRptNS_00D"
36950           End Select
36960           DoCmd.OpenReport strRptSuffix, intView
36970           If intView = acViewPreview Then
36980             DoCmd.Maximize
36990             DoCmd.RunCommand acCmdFitToWindow
37000           End If
37010         End If
37020       End If

37030     End If  ' ** Validate().

37040     DoCmd.Hourglass False

37050   End With

EXITP:
37060   Exit Sub

ERRH:
1380    DoCmd.Hourglass False
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Sub

Public Function Validate() As Boolean

37100 On Error GoTo ERRH

        Const THIS_PROC As String = "Validate"

        Dim blnRetVal As Boolean

37110   blnRetVal = False

        ' ** Validate main three data values.
37120   If IsNull(Me.DateStart.Value) Then
37130     MsgBox "You must enter a From date to continue.", vbInformation + vbOKOnly, "Entry Required"
37140     Me.DateStart.SetFocus
37150   Else
37160     If IsNull(Me.DateEnd.Value) Then
37170       MsgBox "You must enter a To date to continue.", vbInformation + vbOKOnly, "Entry Required"
37180       Me.DateEnd.SetFocus
37190     Else
37200       If Me.DateStart.Value >= Me.DateEnd.Value Then
37210         MsgBox "You must enter a From date that is less than the To date to continue.", vbInformation + vbOKOnly, "Invalid Date"
37220         Me.DateEnd.SetFocus
37230       Else
37240         If IsNull(Me.cmbAccounts.Column(0)) Or Me.cmbAccounts.Column(0) = vbNullString Then
37250           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
37260           Me.cmbAccounts.SetFocus
37270         Else
37280           blnRetVal = True
37290         End If
37300       End If
37310     End If
37320   End If

EXITP:
37330   Validate = blnRetVal
37340   Exit Function

ERRH:
1380    blnRetVal = False
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Function

Public Function BuildAssetListInfo_NS(varStartDate As Variant, varEndDate As Variant, strWhen As String, strRptSuffix As String, strControlName As String) As Integer
' ** Copied from frmRpt_CourtReports_CA.
' ** Return codes:
' **    0  Success.
' **   -2  No data.
' **   -3  Missing entry, e.g., date.
' **   -9  Error.

37400 On Error GoTo ERRH

        Const THIS_PROC As String = "BuildAssetListInfo_NS"

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

37410   With Me

37420     intRetVal_BuildAssetListInfo = 0   ' ** Unless proven otherwise.
37430     blnContinue = True

37440     If .cmbAccounts.Visible = True Then
37450       .cmbAccounts.SetFocus
37460       If .cmbAccounts.text = vbNullString Then
37470         MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
37480         intRetVal_BuildAssetListInfo = -3  ' ** Missing entry.
37490       End If
            ' ** Return focus to the button that called this.
37500       If Right(strControlName, 6) = "_Click" Then
              ' ** cmdPrint00_Click.
37510         .Controls(Left(strControlName, (Len(strControlName) - 6))).SetFocus
37520         DoEvents
37530       ElseIf strControlName = "AssetList_Word_NS" Then
37540         .cmdWord00B.SetFocus
37550         DoEvents
37560       ElseIf strControlName = "AssetList_Excel_NS" Then
37570         .cmdExcel00B.SetFocus
37580         DoEvents
37590       End If
37600     End If

37610     If intRetVal_BuildAssetListInfo = 0 Then
37620       If IsNull(varEndDate) Then
37630         MsgBox "Must enter Period " & strWhen & " date to continue.", vbInformation + vbOKOnly, "Entry Required"
37640         .DateEnd.SetFocus
37650         intRetVal_BuildAssetListInfo = -3  ' ** Missing entry.
37660       End If
37670     End If

37680     If intRetVal_BuildAssetListInfo = 0 Then
            ' ** This code will update the qryMaxBalDates query to give
            ' ** us the balance numbers from the previous statement.
37690       intRetVal_SetDateSpecificSQL = SetDateSpecificSQL(.cmbAccounts, "Statements", THIS_NAME, varStartDate, varEndDate)  ' ** Module Function: modUtilities.
            ' ** Return codes:
            ' **    0  Success.
            ' **   -2  No data.
            ' **   -4  Date criteria not met.
            ' **   -9  Error.
37700       If intRetVal_SetDateSpecificSQL <> 0 Then
37710         blnContinue = False
37720         intRetVal_BuildAssetListInfo = -2  ' ** No data.
37730       End If
37740     End If

37750     If intRetVal_BuildAssetListInfo = 0 And blnContinue = True Then

37760       Set dbs = CurrentDb

            ' ** Empty tmpAssetList2.
37770       If TableExists("tmpAssetList2") = True Then  ' ** Module Function: modUtilities.
37780         Set qdf = dbs.QueryDefs("qryCourtReport_03")
37790         qdf.Execute
37800       End If
            ' ** Empty tmpAccountInfo.
37810       If TableExists("tmpAccountInfo") = True Then  ' ** Module Function: modUtilities.
37820         Set qdf = dbs.QueryDefs("qryCourtReport_04")
37830         qdf.Execute
37840       End If

            ' ** qryCourtReport_NS_07z.
            ' ** VGC 11/25/2009: «='90',-1,1)» to «='90',1,1)».
37850       strSQL = "SELECT ActiveAssets.assetno, masterasset.description AS MasterAssetDescription, masterasset.due, " & _
              "masterasset.rate, Sum(IIf(IsNull([ActiveAssets].[cost]),0,[ActiveAssets].[cost])) AS TotalCost, " & _
              "Sum(IIf(IsNull([ActiveAssets].[shareface]),0,[ActiveAssets].[shareface]))*" & _
              "IIf([assettype].[assettype]='90',1,1) AS TotalShareface, account.accountno, account.shortname, " & _
              "account.legalname, assettype.assettype, assettype_description, " & _
              "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & " & _
              "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
              "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))) AS totdesc, "
37860       strSQL = strSQL & "account.icash, account.pcash, masterasset.currentDate, " & _
              "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]) AS MarketValueX, " & _
              "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]) AS MarketValueCurrentX, " & _
              "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) AS YieldX, " & CoInfo & " "
37870       strSQL = strSQL & "FROM (masterasset LEFT JOIN assettype ON masterasset.assettype = assettype.assettype) " & _
              "RIGHT JOIN (account LEFT JOIN ActiveAssets ON account.accountno = ActiveAssets.accountno) " & _
              "ON masterasset.assetno = ActiveAssets.assetno " & _
              "GROUP BY ActiveAssets.assetno, masterasset.description, masterasset.due, masterasset.rate, " & _
              "account.accountno, account.shortname, account.legalname, assettype.assettype, assettype_description, " & _
              "IIf(IsNull([ActiveAssets].[assetno]),'',CStr([masterasset].[Description]) & "
37880       strSQL = strSQL & "IIf([masterasset].[rate]>0,' ' & Format([masterasset].[rate],'0.000%')) & " & _
              "IIf([masterasset].[due] Is Not Null,'  Due ' & Format([masterasset].[due],'mm/dd/yyyy'))), " & _
              "account.icash, account.pcash, masterasset.currentDate, " & _
              "IIf(IsNull([masterasset].[marketvalue]),0,[masterasset].[marketvalue]), " & _
              "IIf(IsNull([masterasset].[marketvaluecurrent]),0,[masterasset].[marketvaluecurrent]), " & _
              "IIf(IsNull([masterasset].[yield]),0,[masterasset].[yield]) " & _
              "HAVING (((account.accountno)='" & .cmbAccounts & "'));"
37890       strSQL = StringReplace(strSQL, "'' As ", "Null As ")  ' ** Module Function: modStringFuncs.

37900       Set rst = dbs.OpenRecordset(strSQL)  ' ** This is what's used for appending to the tmpAssetList2 table and rstTempAssetList.

            ' ** Ledger records to roll back if after specified 'To' date. (VGC 08/24/2010: CHANGES!)
37910       Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_B_03")
37920       With qdf.Parameters
37930         ![actno] = Me.cmbAccounts
37940         ![datEnd] = CDate(varEndDate)
37950       End With
37960       Set rstWork = qdf.OpenRecordset
37970       If rstWork.BOF = True And rstWork.EOF = True Then
              ' ** In the absense of newer data to roll back, take the total cost as it's normally computed for report 0.
37980         blnContinue = False  ' ** Leave intRetVal_BuildAssetListInfo = 0 so that the procedure calling this continues.
37990         rstWork.Close
38000         If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
38010           dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
38020         End If
38030         rst.Close
38040         dbs.Close
38050         Select Case strWhen
              Case "Beginning"
38060           strRptSuffix = "_00BA"
38070         Case "Ending"
38080           strRptSuffix = "_00DA"
38090         End Select
38100       End If

38110       If blnContinue Then  ' ** There are transactions after the specified To date.

38120         Set rstAccount = dbs.OpenRecordset("SELECT account.* FROM account ORDER BY accountno", dbOpenSnapshot)
38130         Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_B_04")
38140         Set rstMasterAsset = qdf.OpenRecordset()

38150         If MakeTempTable(rst, "tmpAssetList2") Then  ' ** Module Function: modFileUtilities.
38160           dbs.Execute "INSERT INTO tmpAssetList2 " & strSQL  ' ** Copy data from query.
38170           Set rstTempAssetList = dbs.OpenRecordset("SELECT tmpAssetList2.* FROM tmpAssetList2", dbOpenDynaset)
38180         Else
                ' ** ERROR.
38190           intRetVal_BuildAssetListInfo = -9  ' ** Error.
38200           MsgBox "Unable to create temporary table for reporting.", vbCritical + vbOKOnly, "Error"
38210           rstAccount.Close
38220           rstMasterAsset.Close
38230           rstWork.Close
38240           dbs.Close
38250         End If

38260         If intRetVal_BuildAssetListInfo = 0 Then

                ' ** Create table to contain temporary account master records to track "global" account info.
38270           If MakeTempTable(rst, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
38280             dbs.Execute "INSERT INTO tmpAccountInfo " & strSQL  ' ** Copy data from query.
                  ' ** Now, make records here distinct for each account.
38290             dbs.Execute "UPDATE tmpAccountInfo SET assetno = Null, MasterAssetDescription = Null, " & _
                    "due = Null, rate = Null, TotalCost = Null, TotalShareFace = Null,  " & _
                    "assettype = Null, assettype_description = Null, TotDesc = Null,  " & _
                    "currentDate = Null, MarketValueX = Null, MarketValueCurrentX = Null, " & _
                    "YieldX = Null  " & _
                    "WHERE True"
38300             Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT DISTINCTROW tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenSnapshot)
                  ' ** Empty tmpAccountInfo.
38310             Set qdf = dbs.QueryDefs("qryCourtReport_04")
38320             qdf.Execute
38330             If rstTmpAccountInfo.RecordCount > 0 Then
38340               rstTmpAccountInfo.MoveLast
38350               rstTmpAccountInfo.MoveFirst
38360             End If
38370             If Not CopyToTempTable(rstTmpAccountInfo, "tmpAccountInfo") Then  ' ** Module Function: modFileUtilities.
                    ' ** ERROR.
38380               intRetVal_BuildAssetListInfo = -9  ' ** Error.
38390               MsgBox "Unable to copy data to temporary table.", vbCritical + vbOKOnly, "Error"
38400             Else
38410               rstTmpAccountInfo.Close  ' ** Close old snapshot, then open as dynaset.
38420               Set rstTmpAccountInfo = Nothing
38430               Set rstTmpAccountInfo = dbs.OpenRecordset("SELECT tmpAccountInfo.* FROM tmpAccountInfo;", dbOpenDynaset)
38440             End If
38450           Else
                  ' ** ERROR.
38460             intRetVal_BuildAssetListInfo = -9  ' ** Error.
38470             MsgBox "Unable to create temporary account table for reporting.", vbCritical + vbOKOnly, "Error"
38480             rstAccount.Close
38490             rstMasterAsset.Close
38500             rstWork.Close
38510             rstTempAssetList.Close
38520             dbs.Close
38530           End If

38540         End If

38550         If intRetVal_BuildAssetListInfo = 0 Then

38560           rstTmpAccountInfo.MoveLast
38570           rstTmpAccountInfo.MoveFirst

                ' ** Get total number of records in recordset of subsequent ledger entries.
38580           rstWork.MoveLast
38590           rstWork.MoveFirst

38600           Do While Not rstWork.EOF  ' ** Move through each record, tracking changes.

38610             rstTempAssetList.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "' AND [assetno] = " & CStr(rstWork![assetno])
38620             If rstTempAssetList.NoMatch Then
                    ' ** We need a new record because this asset was not in the reporting query after subsequent changes.
38630               rstMasterAsset.FindFirst "[assetno] = " & CStr(rstWork![assetno])
38640               blnNoMAsset = rstMasterAsset.NoMatch
38650               If blnNoMAsset Then
38660                 Select Case Trim(rstWork![journaltype])
                      Case "Misc.", "Paid", "Received"
                        ' ** OK to continue.
38670                 Case Else
                        ' ** ERROR.
38680                   intRetVal_BuildAssetListInfo = -9  ' ** Error.
38690                   MsgBox "Missing Master Asset record.", vbCritical + vbOKOnly, "Error"
38700                   rstAccount.Close
38710                   rstMasterAsset.Close
38720                   rstWork.Close
38730                   rstTempAssetList.Close
38740                   rstTmpAccountInfo.Close
38750                   dbs.Close
38760                 End Select
38770               Else
38780                 rstTempAssetList.AddNew
38790                 rstTempAssetList![accountno] = rstWork![accountno]
38800                 rstTempAssetList![assetno] = rstMasterAsset![assetno]
38810                 rstTempAssetList![MasterAssetDescription] = rstMasterAsset![description]
38820                 rstTempAssetList![assettype] = rstMasterAsset![assettype]
38830                 rstTempAssetList![assettype_description] = rstMasterAsset![assettype_description]
38840                 rstTempAssetList![due] = rstMasterAsset![due]
38850                 rstTempAssetList![rate] = rstMasterAsset![rate]
38860                 rstTempAssetList![totdesc] = CStr(rstMasterAsset![description]) & _
                        IIf(rstMasterAsset![rate] > 0, " " & Format(rstMasterAsset![rate], "0.000%"), "") & _
                        IIf(Not IsNull(rstMasterAsset![due]), "  Due " & Format(rstMasterAsset![due], "mm/dd/yyyy"), "")
38870                 rstTempAssetList![TotalCost] = 0
38880                 rstTempAssetList![TotalShareFace] = 0
38890                 rstTempAssetList![ICash] = 0
38900                 rstTempAssetList![PCash] = 0
                      ' ** LEAVE in editing mode.
38910               End If
38920             Else
38930               blnNoMAsset = False
38940               rstTempAssetList.Edit  ' ** Edit existing temp record.
38950             End If

38960             If intRetVal_BuildAssetListInfo = 0 Then

38970               rstTmpAccountInfo.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"

38980               If rstTmpAccountInfo.NoMatch Then
                      ' ** We need a new account record.

38990                 rstAccount.FindFirst "[accountno] = '" & Trim(rstWork![accountno]) & "'"
39000                 If rstAccount.NoMatch Then
                        ' ** ERROR.
39010                   intRetVal_BuildAssetListInfo = -9  ' ** Error.
39020                   MsgBox "Missing Account master record.", vbCritical + vbOKOnly, "Error"
39030                   rstAccount.Close
39040                   rstMasterAsset.Close
39050                   rstWork.Close
39060                   rstTempAssetList.Close
39070                   rstTmpAccountInfo.Close
39080                   dbs.Close
39090                 Else
39100                   rstTmpAccountInfo.AddNew
                        ' ** Copy info from Account table.
39110                   rstTmpAccountInfo![accountno] = rstAccount![accountno]
39120                   rstTmpAccountInfo![shortname] = rstAccount![shortname]
39130                   rstTmpAccountInfo![legalname] = rstAccount![legalname]
39140                   rstTmpAccountInfo![ICash] = rstAccount![ICash]
39150                   rstTmpAccountInfo![PCash] = rstAccount![PCash]
                        ' ** LEAVE in editing mode.
39160                 End If
39170               Else
39180                 rstTmpAccountInfo.Edit  ' ** Edit existing temp record.
39190               End If

39200             End If  ' ** intRetVal_BuildAssetListInfo.

39210             If intRetVal_BuildAssetListInfo = 0 Then

39220               If (Not blnNoMAsset) And (Trim(rstWork![journaltype]) <> "Received") Then
                      ' ** Add/Subtract info from ASSET record. NOTE: These are SUBTRACTIONS
                      ' ** because the query returns the changes since the date.
                      'VGC 08/20/2010: CHANGES!
39230                 rstTempAssetList![TotalShareFace] = _
                        rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, _
                        rstWork![shareface] * IIf(rstTempAssetList![assettype] = 90, 1, 1)))
                      'rstTempAssetList![TotalShareFace] = _
                      '  rstTempAssetList![TotalShareFace] - (IIf(IsNull(rstWork![shareface]), 0, rstWork![shareface] * _
                      '  IIf(rstTempAssetList![assettype] = 90, -1, 1)))
39240                 rstTempAssetList![TotalCost] = rstTempAssetList![TotalCost] - IIf(IsNull(rstWork![Cost]), 0, rstWork![Cost])
                      ' ** Save ASSET temp table record.
39250                 rstTempAssetList.Update
39260               End If

                    ' ** Add/Subtract info from ACCOUNT record. NOTE: These are SUBTRACTIONS
                    ' ** because the query returns the changes since the date.
39270               rstTmpAccountInfo![ICash] = rstTmpAccountInfo![ICash] - rstWork![ICash]
39280               rstTmpAccountInfo![PCash] = rstTmpAccountInfo![PCash] - rstWork![PCash]
                    ' ** Save ACCOUNT temp table record.
39290               rstTmpAccountInfo.Update

39300               rstWork.MoveNext

39310             End If  ' ** intRetVal_BuildAssetListInfo.

39320             If intRetVal_BuildAssetListInfo <> 0 Then Exit Do

39330           Loop  ' ** Move through each record, tracking changes.

39340         End If  ' ** intRetVal_BuildAssetListInfo.

39350         If intRetVal_BuildAssetListInfo = 0 Then

                ' ** Remove any asset records which now have a ZERO totalshareface AND totalcost.
                'strWorkSQL = "DELETE tmpAssetList2.*, tmpAssetList2.TotalShareface, tmpAssetList2.TotalCost " & _
                '  "FROM tmpAssetList2 " & _
                '  "WHERE ((Round((tmpAssetList2.TotalShareface),4)=0) AND (Round((tmpAssetList2.TotalCost),2)=0));"
                ' ** Delete tmpAssetList2, for zero entries.
                ' ** 07/21/2008: Added rounding.
39360           Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_B_05")
39370           qdf.Execute

                ' ** Now, update tmpAssetList2 to reflect the needed account information which is
                ' ** global to the report.
39380           rstTmpAccountInfo.MoveFirst
39390           While Not rstTmpAccountInfo.EOF
39400             strWorkSQL = "UPDATE tmpAssetList2 SET " & _
                    "shortname = " & SQLFormatStr(rstTmpAccountInfo![shortname], dbText) & ", " & _
                    "legalname = " & SQLFormatStr(rstTmpAccountInfo![legalname], dbText) & ", " & _
                    "icash = " & SQLFormatStr(rstTmpAccountInfo![ICash], dbCurrency) & ", " & _
                    "pcash = " & SQLFormatStr(rstTmpAccountInfo![PCash], dbCurrency) & " " & _
                    "WHERE accountno = " & SQLFormatStr(rstTmpAccountInfo![accountno], dbText) & ";"
39410             dbs.Execute strWorkSQL
39420             rstTmpAccountInfo.MoveNext
39430           Wend

                ' ** Close temp & working recordsets.
39440           rstAccount.Close
39450           rstMasterAsset.Close
39460           rstWork.Close
39470           rstTempAssetList.Close
39480           rstTmpAccountInfo.Close

                ' ** Finally, base report on this instead of on qryAssetList.
39490           Set qdf = dbs.QueryDefs("qryCourtReport_NS_00_B_01")
39500           Set rst = qdf.OpenRecordset()
39510           Select Case strWhen
                Case "Beginning"
39520             strRptSuffix = "_00B"
39530           Case "Ending"
39540             strRptSuffix = "_00D"
39550           End Select

39560         End If  ' ** intRetVal_BuildAssetListInfo.

39570       End If  ' ** blnContinue.

39580       If intRetVal_BuildAssetListInfo = 0 And blnContinue = True Then
39590         If rst.BOF = True And rst.EOF = True Then
39600           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
39610           intRetVal_BuildAssetListInfo = -2  ' ** No data.
39620         End If
39630         rst.Close
39640         dbs.Close
39650       End If

39660     End If  ' ** intRetVal_BuildAssetListInfo.

39670   End With

EXITP:
39680   Set rstWork = Nothing
39690   Set rstTempAssetList = Nothing
39700   Set rstTmpAccountInfo = Nothing
39710   Set rstAccount = Nothing
39720   Set rstMasterAsset = Nothing
39730   Set rst = Nothing
39740   Set qdf = Nothing
39750   Set dbs = Nothing
39760   BuildAssetListInfo_NS = intRetVal_BuildAssetListInfo
39770   Exit Function

ERRH:
1380    intRetVal_BuildAssetListInfo = -9  ' ** Error.
1390    THAT_PROC = THIS_PROC
1400    That_Erl = Erl
1410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1420    Resume EXITP

End Function

Public Sub AllCancelSet1_NS(blnCancel As Boolean)

39800 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet1_NS"

39810   blnAllCancel = blnCancel

EXITP:
39820   Exit Sub

ERRH:
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub
