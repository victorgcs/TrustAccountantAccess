VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_NY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_NY"

'VGC 10/09/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' #########################
' ## Use VBA_RenumErrh().  39620
' #########################

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
'Private lngFiles As Long ', arr_varFile() As Variant
'Private Const F_ELEMS As Integer = 2  ' ** Array's first-element UBound().
'Private Const F_RNAM As Integer = 0
'Private Const F_FILE As Integer = 1
'Private Const F_PATH As Integer = 2

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const MEMO_MAX As Integer = 255

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean
Private blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean

Private datFirstDate As Date, strFirstDateMsg As String
Private strLastArrow As String, strFootnoteDefault As String
Private blnAllCancel As Boolean, blnRebuildTable As Boolean
Private blnGoingToReport As Boolean, blnWasGTR As Boolean
Private strRptName As String, strRptCap As String, strRptPath As String, strRptPathFile As String
Private strCallingForm As String, lngTpp As Long
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

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
180           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190           CLR_DISABLED_FG = CLR_DKGRY
200           CLR_DISABLED_BG = CLR_LTTEAL

210           blnPrintAll_Focus = False: blnWordAll_Focus = False: blnExcelAll_Focus = False

220           strLastArrow = vbNullString
230           gblnLegalName = False
240           gblnUseReveuneExpenseCodes = False
250           gvarCrtRpt_FL_SpecData = Null
260           gblnCrtRpt_NY_InvIncChange = False
270           strFootnoteDefault = .CourtReports_Footnote.DefaultValue
280           If Right(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Left(strFootnoteDefault, (Len(strFootnoteDefault) - 1))
290           If Left(strFootnoteDefault, 1) = Chr(34) Then strFootnoteDefault = Mid(strFootnoteDefault, 2)
300           datFirstDate = DateAdd("y", 1, Date)  ' ** Set this to tomorrow.

310           OpenProcs_NY 1, blnAllCancel, blnRebuildTable, strFootnoteDefault, Cancel, Me  ' ** Module Procedure: modCourtReportsNY3.

320           gstrFormQuerySpec = THIS_NAME
330           gstrCrtRpt_Ordinal = vbNullString: gstrCrtRpt_Version = vbNullString
340           gstrCrtRpt_NetIncome = vbNullString: gstrCrtRpt_NetLoss = vbNullString
350           gstrCrtRpt_CashAssets_Beg = vbNullString: gstrCrtRpt_CashAssets_End = vbNullString

360           OpenProcs_NY 2, blnAllCancel, blnRebuildTable, strFootnoteDefault, Cancel, Me  ' ** Module Procedure: modCourtReportsNY3.

370           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

              ' ** Excel buttons are handled here.
380           BtnsEnable_Set_NY Me  ' ** Module Procedure: modCourtReportsNY1.

390           OpenProcs_NY 3, blnAllCancel, blnRebuildTable, strFootnoteDefault, Cancel, Me  ' ** Module Procedure: modCourtReportsNY3.

400           ShowRptNums False, Me  ' ** Module Procedure: modCourtReportsNY1.

410           DoCmd.Hourglass False

420           .DateStart.SetFocus

430         Else
440           Cancel = -1
450         End If
460       Else
470         Cancel = -1
480       End If
490     End With

500     If Cancel = -1 Then
510       Beep
520       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
530     End If

EXITP:
540     Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
620       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
630       clsMonthClass.hWndForm = .hwnd
          ' ** Let's default to PositionAtCursor.
640       clsMonthClass.PositionAtCursor = True

650     End With

EXITP:
660     Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

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

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyF7
770         intRetVal = 0
780         cmdCalendar1_Click  ' ** Procedure: Below.
790       Case vbKeyF8
800         intRetVal = 0
810         cmdCalendar2_Click  ' ** Procedure: Below.
820       Case vbKeyEscape
830         intRetVal = 0
840         cmdClose_Click  ' ** Procedure: Below.
850       End Select
860     End If

        ' ** Alt keys.
870     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
880       Select Case intRetVal
          Case vbKeyS
890         With Me
900           intRetVal = 0
910           .chkAssetList.SetFocus
920           .chkAssetList = (Not .chkAssetList)
930           chkAssetList_AfterUpdate  ' ** Procedure: Below.
940         End With
950       Case vbKeyX
960         intRetVal = 0
970         cmdClose_Click  ' ** Procedure: Below.
980       End Select
990     End If

        ' ** Ctrl-Shift keys.
1000    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1010      Select Case intRetVal
          Case vbKeyF
1020        With Me
1030          intRetVal = 0
1040          .FocusHolder.SetFocus
1050        End With
1060      Case vbKeyN
1070        With Me
1080          intRetVal = 0
1090          Select Case .cmdPrint00_lbl.Visible
              Case True
1100            ShowRptNums False, Me  ' ** Module Procedure: modCourtReportsNY1.
1110          Case False
1120            ShowRptNums True, Me  ' ** Module Procedure: modCourtReportsNY1.
1130          End Select
1140        End With
1150      End Select
1160    End If

        ' ** Ctrl keys.
1170    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1180      Select Case intRetVal
          Case vbKeyG
1190        With Me
1200          intRetVal = 0
1210          If .chkPageOf.Enabled = True Then
1220            .chkPageOf.SetFocus
1230            .chkPageOf = (Not .chkPageOf)
1240            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1250          End If
1260        End With
1270      Case vbKeyN
1280        With Me
1290          intRetVal = 0
1300          If .chkOpenExcel.Enabled = True Then
1310            .chkOpenExcel.SetFocus
1320            .chkOpenExcel = (Not .chkOpenExcel)
1330            chkOpenExcel_AfterUpdate  ' ** Procedure: Below.
1340          Else
1350            Beep
1360          End If
1370        End With
1380      Case vbKeyO
1390        With Me
1400          intRetVal = 0
1410          .chkOpenWord.SetFocus
1420          .chkOpenWord = (Not .chkOpenWord)
1430          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
1440        End With
1450      End Select
1460    End If

        ' ** Ctrl-Shift keys.
1470    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1480      Select Case intRetVal
          Case vbKeyT
1490        intRetVal = 0
1500        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1510      Case vbKeyF
1520        With Me
1530          intRetVal = 0
1540          .FocusHolder.SetFocus
1550        End With
1560      End Select
1570    End If

EXITP:
1580    KeyCode = intRetVal
1590    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub Form_Timer()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String
        Dim lngBtn_Top As Long, lngBtn_Left As Long

1610    With Me
1620      .TimerInterval = 0&
1630      If gblnSetFocus = True Then
1640        gblnSetFocus = False
1650        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
1660          Forms("frmLoadTimer").Modal = False
1670        End If
1680        strFormName = vbNullString
1690  On Error Resume Next
1700        Do Until strFormName = THIS_NAME
1710          DoCmd.SelectObject acForm, THIS_NAME, False
1720          strFormName = Screen.ActiveForm.Name
1730        Loop
1740  On Error GoTo ERRH
1750      End If
1760      If gblnGoToReport = True Then
1770        If lngTpp = 0& Then
              'lngTpp = GetTPP  ' ** Module Procedure: modWindowFunctions.
1780          lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1790        End If
1800        Select Case blnGoingToReport
            Case True
1810          DoCmd.Hourglass False  ' ** It should already be off.
1820          blnGoingToReport = False
1830          gblnGoToReport = False
1840          .GoToReport_arw_rpt_img.Visible = False
1850          DoEvents
1860          Select Case garr_varGoToReport(GTR_CTL6)
              Case "cmdPrint00"
1870            .cmdPreview00_lbl4_txt.Visible = True
1880            Select Case .chkGroupBy_IncExpCode
                Case True
1890              .cmdPreview00_lbl5.Visible = True
1900            Case False
1910              .cmdPreview00_lbl5.Visible = False
1920            End Select
1930          Case "cmdPrint00B"
1940            .GoToReport_arw_alist_on_img.Visible = False
1950            .GoToReport_arw_alist_off_img.Visible = False
1960            .chkAssetList.Visible = True
1970            .chkAssetList_lbl2.Visible = True
1980            .chkAssetList_box.Visible = True
1990          Case "cmdPrint02"
2000            .cmdPreview02_Sch_lbl.Visible = True
2010            .GoToReport_arw_increases_img.Visible = False
2020          Case "cmdPrint04"
2030            .cmdPreview04_Sch_lbl.Visible = True
2040            .GoToReport_arw_admin_img.Visible = False
2050            .cmdPreview04_lbl4_txt.Visible = True
2060            .cmdPreview04_lbl4_cmd.Visible = True
2070            Select Case .chkGroupBy_IncExpCode
                Case True
2080              .cmdPreview04_lbl5.Visible = True
2090            Case False
2100              .cmdPreview04_lbl5.Visible = False
2110            End Select
2120          Case "cmdPrint09"
2130            .cmdPreview09_lbl4_txt.Visible = True
2140            Select Case .chkGroupBy_IncExpCode
                Case True
2150              .cmdPreview09_lbl5.Visible = True
2160            Case False
2170              .cmdPreview09_lbl5.Visible = False
2180            End Select
2190          Case "cmdPrint10"
2200            .cmdPreview10_lbl4_txt.Visible = True
2210            Select Case .chkGroupBy_IncExpCode
                Case True
2220              .cmdPreview10_lbl5.Visible = True
2230            Case False
2240              .cmdPreview10_lbl5.Visible = False
2250            End Select
2260          End Select
2270          DoEvents
2280        Case False
2290          DoCmd.Hourglass True  ' ** Make sure it's still running.
2300          DoEvents
2310          blnGoingToReport = True
2320          Select Case garr_varGoToReport(GTR_CTL6)
              Case "cmdPrint00"
2330            lngBtn_Top = (.cmdPreview00_lbl.Top + (lngTpp))
2340            lngBtn_Left = (.cmdPreview00_lbl.Left + .cmdPreview00_lbl.Width)
2350            .cmdPreview00_lbl4_txt.Visible = False
2360            .cmdPreview00_lbl5.Visible = False
2370            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2380            DoEvents
2390          Case "cmdPrint00B"
2400            lngBtn_Top = (.cmdPreview00B_lbl.Top - (8& * lngTpp))
2410            lngBtn_Left = ((.cmdPreview00B_lbl.Left + .cmdPreview00B_lbl.Width) - (0))
2420            .chkAssetList.Visible = False
2430            .chkAssetList_lbl2.Visible = False
2440            .chkAssetList_box.Visible = False
2450            .GoToReport_arw_alist_on_img.Top = lngBtn_Top
2460            .GoToReport_arw_alist_off_img.Top = lngBtn_Top
2470            .GoToReport_arw_alist_on_img.Left = lngBtn_Left
2480            .GoToReport_arw_alist_off_img.Left = lngBtn_Left
2490            Select Case .chkAssetList
                Case True
2500              .GoToReport_arw_alist_on_img.Visible = True
2510            Case False
2520              .GoToReport_arw_alist_off_img.Visible = True
2530            End Select
2540            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2550            DoEvents
2560          Case "cmdPrint01"
2570            lngBtn_Top = (.cmdPreview01_lbl.Top + (lngTpp))
2580            lngBtn_Left = (.cmdPreview01_lbl.Left + .cmdPreview01_lbl.Width)
2590            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2600            DoEvents
2610          Case "cmdPrint02"
2620            .cmdPreview02_Sch_lbl.Visible = False
2630            .GoToReport_arw_increases_img.Visible = True
2640            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2650            DoEvents
2660          Case "cmdPrint03"
2670            lngBtn_Top = (.cmdPreview03_lbl.Top + (lngTpp))
2680            lngBtn_Left = (.cmdPreview03_lbl_box.Left + .cmdPreview03_lbl_box.Width)
2690            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2700            DoEvents
2710          Case "cmdPrint04"
                'lngBtn_Top = (.cmdPreview04_lbl.Top + (lngTpp))
                'lngBtn_Left = (.cmdPreview04_lbl.Left + .cmdPreview04_lbl.Width)
2720            .cmdPreview04_lbl4_txt.Visible = False
2730            .cmdPreview04_lbl4_cmd.Visible = False
2740            .cmdPreview04_lbl5.Visible = False
2750            .cmdPreview04_Sch_lbl.Visible = False
2760            .GoToReport_arw_admin_img.Visible = True
2770            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2780            DoEvents
2790          Case "cmdPrint05"
2800            lngBtn_Top = (.cmdPreview05_lbl.Top + (lngTpp))
2810            lngBtn_Left = (.cmdPreview05_lbl.Left + .cmdPreview05_lbl.Width)
2820            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2830            DoEvents
2840          Case "cmdPrint06"
2850            lngBtn_Top = (.cmdPreview06_lbl.Top + (lngTpp))
2860            lngBtn_Left = (.cmdPreview06_lbl_box.Left + .cmdPreview06_lbl_box.Width)
2870            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2880            DoEvents
2890          Case "cmdPrint07"
2900            lngBtn_Top = (.cmdPreview07_lbl.Top + (lngTpp))
2910            lngBtn_Left = (.cmdPreview07_lbl.Left + .cmdPreview07_lbl.Width)
2920            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2930            DoEvents
2940          Case "cmdPrint08"
2950            lngBtn_Top = (.cmdPreview08_lbl.Top + (lngTpp))
2960            lngBtn_Left = (.cmdPreview08_lbl.Left + .cmdPreview08_lbl.Width)
2970            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
2980            DoEvents
2990          Case "cmdPrint09"
3000            lngBtn_Top = (.cmdPreview09_lbl.Top + (lngTpp))
3010            lngBtn_Left = (.cmdPreview09_lbl.Left + .cmdPreview09_lbl.Width)
3020            .cmdPreview09_lbl4_txt.Visible = False
3030            .cmdPreview09_lbl5.Visible = False
3040            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3050            DoEvents
3060          Case "cmdPrint10"
3070            lngBtn_Top = (.cmdPreview10_lbl.Top + (lngTpp))
3080            lngBtn_Left = (.cmdPreview10_lbl.Left + .cmdPreview10_lbl.Width)
3090            .cmdPreview10_lbl4_txt.Visible = False
3100            .cmdPreview10_lbl5.Visible = False
3110            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3120            DoEvents
3130          Case "cmdPrint11"
3140            lngBtn_Top = (.cmdPreview11_lbl.Top + (lngTpp))
3150            lngBtn_Left = (.cmdPreview11_lbl.Left + .cmdPreview11_lbl.Width)
3160            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3170            DoEvents
3180          Case "cmdPrint12"
3190            lngBtn_Top = (.cmdPreview12_lbl.Top + (lngTpp))
3200            lngBtn_Left = (.cmdPreview12_lbl.Left + .cmdPreview12_lbl.Width)
3210            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3220            DoEvents
3230          Case "cmdPrint13"
3240            lngBtn_Top = (.cmdPreview13_lbl.Top + (lngTpp))
3250            lngBtn_Left = (.cmdPreview13_lbl_box.Left + .cmdPreview13_lbl_box.Width)
3260            .Controls(garr_varGoToReport(GTR_CTL6)).SetFocus
3270            DoEvents
3280          End Select
3290          If garr_varGoToReport(GTR_CTL6) <> "cmdPrint00B" And garr_varGoToReport(GTR_CTL6) <> "cmdPrint02" And _
                  garr_varGoToReport(GTR_CTL6) <> "cmdPrint04" Then
3300            .GoToReport_arw_rpt_img.Top = lngBtn_Top
3310            .GoToReport_arw_rpt_img.Left = lngBtn_Left
3320            .GoToReport_arw_rpt_img.Visible = True
3330          End If
3340          DoEvents
3350          Beep
3360          DoCmd.Hourglass False
3370          DoEvents
3380          .TimerInterval = (GTR_WAIT * 2&)
3390        End Select
3400      End If
3410      If gblnTimer = True Then
3420        If IsLoaded("frmLoadTimer", acForm) = True Then  ' ** Module Function: modFileUtilities.
3430          Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
3440          gblnTimer = False
3450          DoCmd.Hourglass False
3460        End If
3470      End If
3480    End With

'HERE! 2  Form_Load()  frmRpt_CourtReports_NY
'HERE! 1  Form_Open()  frmRpt_CourtReports_NY
'HERE! 3  Form_Timer()  frmRpt_CourtReports_NY

EXITP:
3490    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3510    Select Case DataErr
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
3520    Case 2113  ' ** The value you entered isn't valid for this field.
3530      Select Case Me.ActiveControl.Name
          Case "DateStart"
3540        If Me.DateStart.text <> "__/__/____" Then
3550          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
3560        Else
3570          Me.DateStart.text = vbNullString  '  ""
3580        End If
3590      End Select
3600      Response = acDataErrContinue
3610    Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3620      If Right(Me.ActiveControl.Name, 4) = "Date" Then
3630        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
3640        Response = acDataErrContinue
3650      End If
3660    Case Else
3670      zErrorHandler THIS_NAME, "Undefined Form Error", DataErr  ' ** Module Function: modErrorHandler.
3680    End Select

EXITP:
3690    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3710    With Me
3720      If .chkRememberDates = False Then
3730        .DateStart = Null  ' ** So they won't be there when Pref_Save() runs.
3740        .DateEnd = Null
3750      End If
3760      If .chkRememberMe = False Then
3770        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
3780      End If
3790    End With

3800    FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3810    DoEvents

3820    Select Case blnWasGTR
        Case True
3830      blnWasGTR = False
3840    Case False
          ' ** Don't save settings imposed by GoToReport!
3850      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
3860    End Select

3870    gblnLegalName = False

3880      ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
3890      ChkSpecLedgerEntryNull  ' ** Module Function: modUtilities.

3900    gblnCrtRpt_NY_InvIncChange = False
3910    gcurCrtRpt_NY_InputAmtForm = 0

        ' ** Make sure all reports are closed.
3920    Do While Reports.Count > 0
3930      DoCmd.Close acReport, Reports(0).Name
3940    Loop

3950    gstrFormQuerySpec = vbNullString
3960    gblnSetFocus = True
3970    gstrReturningForm = THIS_NAME
3980    Select Case strCallingForm
        Case "frmMenu_CourtReport"
3990      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Report"
4000    End Select

EXITP:
4010    Set qdf = Nothing
4020    Set dbs = Nothing
4030    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4110    Detail_Mouse_NY blnCalendar1_Focus, blnCalendar2_Focus, blnPrintAll_Focus, blnWordAll_Focus, blnExcelAll_Focus, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
4120    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

4210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4220    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdClose_Click()

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4310    DoCmd.Close acForm, THIS_NAME

EXITP:
4320    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4420    intShiftDown = (Shift And acShiftMask) > 0
4430    intAltDown = (Shift And acAltMask) > 0
4440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4460      Select Case intRetVal
          Case vbKeyTab
4470        With Me
4480          intRetVal = 0
4490          .DateStart.SetFocus
4500        End With
4510      Case vbKeyLeft
4520        With Me
4530          intRetVal = 0
4540          If .cmdExcelAll.Enabled = True Then
4550            .cmdExcelAll.SetFocus
4560          Else
4570            .cmdWordAll.SetFocus
4580          End If
4590        End With
4600      Case vbKeyRight
4610        With Me
4620          intRetVal = 0
4630          .DateStart.SetFocus
4640        End With
4650      End Select
4660    End If

        ' ** Shift keys.
4670    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4680      Select Case intRetVal
          Case vbKeyTab
4690        With Me
4700          intRetVal = 0
4710          If .cmdExcelAll.Enabled = True Then
4720            .cmdExcelAll.SetFocus
4730          Else
4740            .cmdWordAll.SetFocus
4750          End If
4760        End With
4770      End Select
4780    End If

        ' ** Ctrl-Shift keys.
4790    If intCtrlDown And (Not intAltDown) And intShiftDown Then
4800      Select Case intRetVal
          Case vbKeyTab
4810        With Me
4820          intRetVal = 0
4830          If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
4840            .UserReportPath_chk.SetFocus
4850          End If
4860        End With
4870      End Select
4880    End If

EXITP:
4890    KeyCode = intRetVal
4900    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub DateStart_AfterUpdate()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_AfterUpdate"

5010    With Me
5020      If IsNull(.DateStart) = False Then
5030        If IsDate(.DateStart) = True Then
5040          .DateEnd.SetFocus
5050        Else
5060          MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
5070        End If
5080      End If
5090    End With

EXITP:
5100    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5220    intShiftDown = (Shift And acShiftMask) > 0
5230    intAltDown = (Shift And acAltMask) > 0
5240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5270        With Me
5280          intRetVal = 0
5290  On Error Resume Next
5300          .DateEnd.SetFocus
5310  On Error GoTo ERRH
5320        End With
5330      End Select
5340    End If

        ' ** Shift keys.
5350    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5370        With Me
5380          intRetVal = 0
5390          .cmdClose.SetFocus
5400        End With
5410      End Select
5420    End If

        ' ** Ctrl keys.
5430    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5450        With Me
5460          intRetVal = 0
5470          .cmdCalendar1.SetFocus
5480        End With
5490      End Select
5500    End If

EXITP:
5510    KeyCode = intRetVal
5520    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Public Sub DateStart_Exit(Cancel As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

        Dim strTmp01 As String

5610    With Me
5620      blnRebuildTable = True
5630      If IsNull(.DateStart) = False Then
5640        gdatStartDate = .DateStart
5650        strTmp01 = CStr(CDbl(gdatStartDate))
5660        If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
5670        glngStartDateLong = CLng(strTmp01)
5680      End If
5690    End With

EXITP:
5700    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case 94  ' ** Invalid use of Null.
          ' ** Ignore.
560     Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

5810    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
5820    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

5910    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
5920    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

6010    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
6020    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

6110    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
6120    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

6210    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
6220    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

6310    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
6320    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub DateEnd_AfterUpdate()

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_AfterUpdate"

6410    With Me
6420      If IsNull(.DateEnd) Then .DateEnd = Date
6430      gdatCrtRpt_NY_DateEnd = .DateEnd
6440    End With

EXITP:
6450    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6520    intShiftDown = (Shift And acShiftMask) > 0
6530    intAltDown = (Shift And acAltMask) > 0
6540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6570        With Me
6580          intRetVal = 0
6590  On Error Resume Next
6600          .cmbAccounts.SetFocus
6610  On Error GoTo ERRH
6620        End With
6630      End Select
6640    End If

        ' ** Shift keys.
6650    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6670        With Me
6680          intRetVal = 0
6690          .DateStart.SetFocus
6700        End With
6710      End Select
6720    End If

        ' ** Ctrl keys.
6730    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6750        With Me
6760          intRetVal = 0
6770          .cmdCalendar2.SetFocus
6780        End With
6790      End Select
6800    End If

        ' ** Ctrl-Shift keys.
6810    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6820      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6830        With Me
6840          intRetVal = 0
6850          .cmdCalendar1.SetFocus
6860        End With
6870      End Select
6880    End If

EXITP:
6890    KeyCode = intRetVal
6900    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Public Sub DateEnd_Exit(Cancel As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

        Dim strTmp01 As String

7010    With Me
7020      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
7030        .DateEnd = Date
7040      Else
7050        If .DateEnd = vbNullString Then
              ' ** Populate it with today's date.
7060          .DateEnd = Date
7070        End If
7080      End If
7090      blnRebuildTable = True
7100      gdatEndDate = .DateEnd
7110      strTmp01 = CStr(CDbl(gdatEndDate))
7120      If InStr(strTmp01, ".") > 0 Then strTmp01 = Left(strTmp01, (InStr(strTmp01, ".") - 1))
7130      glngEndDateLong = CLng(strTmp01)
7140    End With

EXITP:
7150    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

7210    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
7220    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

7310    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
7320    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

7410    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
7420    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7520    intShiftDown = (Shift And acShiftMask) > 0
7530    intAltDown = (Shift And acAltMask) > 0
7540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7570        With Me
7580          intRetVal = 0
7590          .chkRememberDates.SetFocus
7600        End With
7610      End Select
7620    End If

        ' ** Shift keys.
7630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7650        With Me
7660          intRetVal = 0
7670          .DateEnd.SetFocus
7680        End With
7690      End Select
7700    End If

EXITP:
7710    KeyCode = intRetVal
7720    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

7810    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
7820    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

7910    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
7920    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

8010    Calendar_Handler_NY THIS_PROC, blnCalendar1_Focus, blnCalendar1_MouseDown, _
          blnCalendar2_Focus, blnCalendar2_MouseDown, clsMonthClass, Me  ' ** Module Procedure: modCourtReportsNY.

EXITP:
8020    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub chkRememberDates_AfterUpdate()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

8110    With Me
8120      Select Case .chkRememberDates
          Case True
8130        .chkRememberDates_lbl.FontBold = True
8140      Case False
8150        .chkRememberDates_lbl.FontBold = False
8160      End Select
8170    End With

EXITP:
8180    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

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
8290          .opgAccountSource.SetFocus
8300        End With
8310      End Select
8320    End If

        ' ** Shift keys.
8330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8350        With Me
8360          intRetVal = 0
8370          .cmdCalendar2.SetFocus
8380        End With
8390      End Select
8400    End If

EXITP:
8410    KeyCode = intRetVal
8420    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Public Sub opgAccountSource_AfterUpdate()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

8510    AcctSource_After_NY Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
8520    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

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
8690          .chkRememberMe.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8750        With Me
8760          intRetVal = 0
8770          .chkRememberDates.SetFocus
8780        End With
8790      End Select
8800    End If

EXITP:
8810    KeyCode = intRetVal
8820    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8920    intShiftDown = (Shift And acShiftMask) > 0
8930    intAltDown = (Shift And acAltMask) > 0
8940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8960      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8970        With Me
8980          intRetVal = 0
8990          .chkRememberMe.SetFocus
9000        End With
9010      End Select
9020    End If

        ' ** Shift keys.
9030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9040      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9050        With Me
9060          intRetVal = 0
9070          .chkRememberDates.SetFocus
9080        End With
9090      End Select
9100    End If

EXITP:
9110    KeyCode = intRetVal
9120    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Public Sub chkRememberMe_AfterUpdate()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

9210    With Me
9220      Select Case .chkRememberMe
          Case True
9230        .chkRememberMe_lbl.FontBold = True
9240      Case False
9250        .chkRememberMe_lbl.FontBold = False
9260      End Select
9270    End With

EXITP:
9280    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9320    intShiftDown = (Shift And acShiftMask) > 0
9330    intAltDown = (Shift And acAltMask) > 0
9340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9370        With Me
9380          intRetVal = 0
9390          .cmbAccounts.SetFocus
9400        End With
9410      End Select
9420    End If

        ' ** Shift keys.
9430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9450        With Me
9460          intRetVal = 0
9470          .opgAccountSource.SetFocus
9480        End With
9490      End Select
9500    End If

EXITP:
9510    KeyCode = intRetVal
9520    Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

9610    With Me
9620      If .cmbAccounts.text = vbNullString Then
9630        ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
9640        .cmbAccounts = .cmbAccounts.ItemData(0)
9650        blnRebuildTable = True
9660      End If
9670    End With

EXITP:
9680    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub cmbAccounts_AfterUpdate()

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_AfterUpdate"

9710    With Me
9720      If IsNull(.cmbAccounts) = False Then
9730        If Trim(.cmbAccounts) <> vbNullString Then

9740          blnRebuildTable = True
9750          gstrCrtRpt_Account = .cmbAccounts
9760          gstrCrtRpt_NY_AccountNo = .cmbAccounts

9770        End If
9780      End If
9790    End With

EXITP:
9800    Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

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
9990          .cmdPreview00.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10050       With Me
10060         intRetVal = 0
10070         .DateEnd.SetFocus
10080       End With
10090     End Select
10100   End If

        ' ** Ctrl keys.
10110   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10120     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10130       With Me
10140         intRetVal = 0
10150         If .chkGroupBy_IncExpCode.Enabled = True Then
10160           .chkGroupBy_IncExpCode.SetFocus
10170         Else
10180           .chkLegalName.SetFocus
10190         End If
10200       End With
10210     End Select
10220   End If

        ' ** Ctrl-Shift keys.
10230   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10250       With Me
10260         intRetVal = 0
10270         .chkRememberMe.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Ignore.
570     Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Sub

Public Sub chkGroupBy_IncExpCode_AfterUpdate()
' ** Public for GoToReport.
' ** Available on reports 0, 1, 3, 8, 11, 12.

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_AfterUpdate"

10410   IncExpGrp_After_NY Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
10420   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkGroupBy_IncExpCode_KeyDown(KeyCode As Integer, Shift As Integer)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkGroupBy_IncExpCode_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10520   intShiftDown = (Shift And acShiftMask) > 0
10530   intAltDown = (Shift And acAltMask) > 0
10540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10570       With Me
10580         intRetVal = 0
10590         .chkLegalName.SetFocus
10600       End With
10610     End Select
10620   End If

        ' ** Shift keys.
10630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10650       With Me
10660         intRetVal = 0
10670         .cmbAccounts.SetFocus
10680       End With
10690     End Select
10700   End If

EXITP:
10710   KeyCode = intRetVal
10720   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub chkLegalName_AfterUpdate()

10800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_AfterUpdate"

10810   With Me
10820     Select Case .chkLegalName
          Case True
10830       .chkLegalName_lbl.FontBold = True
10840       gblnLegalName = True
10850     Case False
10860       .chkLegalName_lbl.FontBold = False
10870       gblnLegalName = False
10880     End Select
10890   End With

EXITP:
10900   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkLegalName_KeyDown(KeyCode As Integer, Shift As Integer)

11000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkLegalName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11020   intShiftDown = (Shift And acShiftMask) > 0
11030   intAltDown = (Shift And acAltMask) > 0
11040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11070       With Me
11080         intRetVal = 0
11090         .chkIncludeFootnote.SetFocus
11100       End With
11110     End Select
11120   End If

        ' ** Shift keys.
11130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11150       With Me
11160         intRetVal = 0
11170         If .chkGroupBy_IncExpCode.Enabled = True Then
11180           .chkGroupBy_IncExpCode.SetFocus
11190         Else
11200           .cmbAccounts.SetFocus
11210         End If
11220       End With
11230     End Select
11240   End If

EXITP:
11250   KeyCode = intRetVal
11260   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub chkIncludeFootnote_AfterUpdate()

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_AfterUpdate"

11310   Feet_After_NY 1, strFootnoteDefault, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
11320   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkIncludeFootnote_KeyDown(KeyCode As Integer, Shift As Integer)

11400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeFootnote_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11420   intShiftDown = (Shift And acShiftMask) > 0
11430   intAltDown = (Shift And acAltMask) > 0
11440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11470       With Me
11480         intRetVal = 0
11490         Select Case .chkIncludeFootnote
              Case True
11500           .CourtReports_Footnote.SetFocus
11510         Case False
11520           .chkPageOf.SetFocus
11530         End Select
11540       End With
11550     End Select
11560   End If

        ' ** Shift keys.
11570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11590       With Me
11600         intRetVal = 0
11610         .chkLegalName.SetFocus
11620       End With
11630     End Select
11640   End If

EXITP:
11650   KeyCode = intRetVal
11660   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub CourtReports_Footnote_AfterUpdate()

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_AfterUpdate"

11710   Feet_After_NY 2, strFootnoteDefault, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
11720   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub CourtReports_Footnote_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "CourtReports_Footnote_KeyDown"

        Dim intRetVal As Integer

11810   intRetVal = Feet_Key_NY(KeyCode, Shift, Me)  ' ** Module Function: modCourtReportsNY3.

EXITP:
11820   KeyCode = intRetVal
11830   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub chkPageOf_AfterUpdate()

11900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

11910   With Me
11920     Select Case .chkPageOf
          Case True
11930       .chkPageOf_lbl.FontBold = True
11940     Case False
11950       .chkPageOf_lbl.FontBold = False
11960     End Select
11970   End With

EXITP:
11980   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12020   intShiftDown = (Shift And acShiftMask) > 0
12030   intAltDown = (Shift And acAltMask) > 0
12040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12070       With Me
12080         intRetVal = 0
12090         .cmdPreview00.SetFocus
12100       End With
12110     End Select
12120   End If

        ' ** Shift keys.
12130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12150       With Me
12160         intRetVal = 0
12170         .chkIncludeFootnote.SetFocus
12180       End With
12190     End Select
12200   End If

EXITP:
12210   KeyCode = intRetVal
12220   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview00_lbl4_cmd_DblClick(Cancel As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_lbl4_cmd_DblClick"

12310   With Me
12320     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
12330     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
12340   End With

EXITP:
12350   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPreview00_Click()

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_Click"

12410     NY_CmdPrev00_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
12420   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview00_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode
12520   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
12530   intShiftDown = (Shift And acShiftMask) > 0
12540   intAltDown = (Shift And acAltMask) > 0
12550   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12560   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12570     Select Case intRetVal
          Case vbKeyTab
12580       With Me
12590         intRetVal = 0
12600         .cmdPrint00.SetFocus
12610       End With
12620     Case vbKeyUp
12630       With Me
12640         intRetVal = 0
12650         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
12660         .cmdPreview00B.SetFocus
12670       End With
12680     Case vbKeyDown
12690       With Me
12700         intRetVal = 0
12710         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
12720         .cmdPreview01.SetFocus
12730       End With
12740     End Select
12750   End If

        ' ** Shift keys.
12760   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12770     Select Case intRetVal
          Case vbKeyTab
12780       With Me
12790         intRetVal = 0
12800         .cmbAccounts.SetFocus
12810       End With
12820     End Select
12830   End If

        ' ** Ctrl-Shift keys.
12840   If intCtrlDown And (Not intAltDown) And intShiftDown Then
12850     Select Case intRetVal
          Case vbKeyTab
12860       With Me
12870         intRetVal = 0
12880         .chkPageOf.SetFocus
12890       End With
12900     End Select
12910   End If

EXITP:
12920   KeyCode = intRetVal
12930   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdPrint00_Click()

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00_Click"

13010     NY_CmdPrint00_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
13020   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint00_KeyDown(KeyCode As Integer, Shift As Integer)

13100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13110   intRetVal = KeyCode
13120   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
13130   intShiftDown = (Shift And acShiftMask) > 0
13140   intAltDown = (Shift And acAltMask) > 0
13150   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13160   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13170     Select Case intRetVal
          Case vbKeyUp
13180       With Me
13190         intRetVal = 0
13200         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
13210         .cmdPrintAll.SetFocus
13220       End With
13230     Case vbKeyDown
13240       With Me
13250         intRetVal = 0
13260         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
13270         .cmdPrint01.SetFocus
13280       End With
13290     End Select
13300   End If

EXITP:
13310   KeyCode = intRetVal
13320   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord00_Click()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_Click"

13410     blnAllCancel = False
13420     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
13430     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

13440     NY_CmdWord00_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
13450   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord00_KeyDown(KeyCode As Integer, Shift As Integer)

13500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13510   intRetVal = KeyCode
13520   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
13530   intShiftDown = (Shift And acShiftMask) > 0
13540   intAltDown = (Shift And acAltMask) > 0
13550   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13560   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13570     Select Case intRetVal
          Case vbKeyUp
13580       With Me
13590         intRetVal = 0
13600         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
13610         .cmdWordAll.SetFocus
13620       End With
13630     Case vbKeyDown
13640       With Me
13650         intRetVal = 0
13660         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
13670         .cmdWord01.SetFocus
13680       End With
13690     End Select
13700   End If

EXITP:
13710   KeyCode = intRetVal
13720   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel00_Click()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

13810     blnAllCancel = False
13820     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
13830     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

13840     NY_CmdExcel00_Click strRptName, strRptCap, strRptPath, strRptPathFile, strCallingForm, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
13850   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel00_KeyDown(KeyCode As Integer, Shift As Integer)

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13910   intRetVal = KeyCode
13920   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
13930   intShiftDown = (Shift And acShiftMask) > 0
13940   intAltDown = (Shift And acAltMask) > 0
13950   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13960   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13970     Select Case intRetVal
          Case vbKeyUp
13980       With Me
13990         intRetVal = 0
14000         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
14010         .cmdExcelAll.SetFocus
14020       End With
14030     Case vbKeyDown
14040       With Me
14050         intRetVal = 0
14060         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
14070         .cmdExcel01.SetFocus
14080       End With
14090     End Select
14100   End If

EXITP:
14110   KeyCode = intRetVal
14120   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview01_Click()

14200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview01_Click"

14210     NY_CmdPrev01_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
14220   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview01_KeyDown(KeyCode As Integer, Shift As Integer)

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview01_KeyDown"

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
          Case vbKeyUp
14370       With Me
14380         intRetVal = 0
14390         .cmdPreview00.SetFocus
14400       End With
14410     Case vbKeyDown
14420       With Me
14430         intRetVal = 0
14440         .cmdPreview02.SetFocus
14450       End With
14460     End Select
14470   End If

EXITP:
14480   KeyCode = intRetVal
14490   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint01_Click()

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint01_Click"

14510     NY_CmdPrint01_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
14520   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint01_KeyDown(KeyCode As Integer, Shift As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint01_KeyDown"

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
          Case vbKeyUp
14670       With Me
14680         intRetVal = 0
14690         .cmdPrint00.SetFocus
14700       End With
14710     Case vbKeyDown
14720       With Me
14730         intRetVal = 0
14740         .cmdPrint02.SetFocus
14750       End With
14760     End Select
14770   End If

EXITP:
14780   KeyCode = intRetVal
14790   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord01_Click()

14800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord01_Click"

14810     blnAllCancel = False
14820     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
14830     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

14840     NY_CmdWord01_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
14850   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord01_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord01_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14920   intShiftDown = (Shift And acShiftMask) > 0
14930   intAltDown = (Shift And acAltMask) > 0
14940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14960     Select Case intRetVal
          Case vbKeyUp
14970       With Me
14980         intRetVal = 0
14990         .cmdWord00.SetFocus
15000       End With
15010     Case vbKeyDown
15020       With Me
15030         intRetVal = 0
15040         .cmdWord02.SetFocus
15050       End With
15060     End Select
15070   End If

EXITP:
15080   KeyCode = intRetVal
15090   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel01_Click()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel01_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

15110     blnAllCancel = False
15120     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
15130     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

15140     NY_CmdExcel01_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
15150   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel01_KeyDown(KeyCode As Integer, Shift As Integer)

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel01_KeyDown"

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
          Case vbKeyUp
15270       With Me
15280         intRetVal = 0
15290         .cmdExcel00.SetFocus
15300       End With
15310     Case vbKeyDown
15320       With Me
15330         intRetVal = 0
15340         .cmdExcel02.SetFocus
15350       End With
15360     End Select
15370   End If

EXITP:
15380   KeyCode = intRetVal
15390   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview02_Click()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview02_Click"

15410     NY_CmdPrev02_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
15420   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview02_KeyDown(KeyCode As Integer, Shift As Integer)

15500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview02_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15520   intShiftDown = (Shift And acShiftMask) > 0
15530   intAltDown = (Shift And acAltMask) > 0
15540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15560     Select Case intRetVal
          Case vbKeyUp
15570       With Me
15580         intRetVal = 0
15590         .cmdPreview01.SetFocus
15600       End With
15610     Case vbKeyDown
15620       With Me
15630         intRetVal = 0
15640         .cmdPreview03.SetFocus
15650       End With
15660     End Select
15670   End If

EXITP:
15680   KeyCode = intRetVal
15690   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint02_Click()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint02_Click"

15710     NY_CmdPrint02_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
15720   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint02_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint02_KeyDown"

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
          Case vbKeyUp
15870       With Me
15880         intRetVal = 0
15890         .cmdPrint01.SetFocus
15900       End With
15910     Case vbKeyDown
15920       With Me
15930         intRetVal = 0
15940         .cmdPrint03.SetFocus
15950       End With
15960     End Select
15970   End If

EXITP:
15980   KeyCode = intRetVal
15990   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord02_Click()

16000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord02_Click"

16010     blnAllCancel = False
16020     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
16030     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

16040     NY_CmdWord02_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
16050   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord02_KeyDown(KeyCode As Integer, Shift As Integer)

16100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord02_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16120   intShiftDown = (Shift And acShiftMask) > 0
16130   intAltDown = (Shift And acAltMask) > 0
16140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16160     Select Case intRetVal
          Case vbKeyUp
16170       With Me
16180         intRetVal = 0
16190         .cmdWord01.SetFocus
16200       End With
16210     Case vbKeyDown
16220       With Me
16230         intRetVal = 0
16240         .cmdWord03.SetFocus
16250       End With
16260     End Select
16270   End If

EXITP:
16280   KeyCode = intRetVal
16290   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel02_Click()

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel02_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

16310     blnAllCancel = False
16320     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
16330     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

16340     NY_CmdExcel02_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
16350   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel02_KeyDown(KeyCode As Integer, Shift As Integer)

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel02_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16420   intShiftDown = (Shift And acShiftMask) > 0
16430   intAltDown = (Shift And acAltMask) > 0
16440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16460     Select Case intRetVal
          Case vbKeyUp
16470       With Me
16480         intRetVal = 0
16490         .cmdExcel01.SetFocus
16500       End With
16510     Case vbKeyDown
16520       With Me
16530         intRetVal = 0
16540         .cmdExcel03.SetFocus
16550       End With
16560     End Select
16570   End If

EXITP:
16580   KeyCode = intRetVal
16590   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview03_Click()

16600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview03_Click"

16610     NY_CmdPrev03_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
16620   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview03_KeyDown(KeyCode As Integer, Shift As Integer)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview03_KeyDown"

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
          Case vbKeyUp
16770       With Me
16780         intRetVal = 0
16790         .cmdPreview02.SetFocus
16800       End With
16810     Case vbKeyDown
16820       With Me
16830         intRetVal = 0
16840         .cmdPreview04.SetFocus
16850       End With
16860     End Select
16870   End If

EXITP:
16880   KeyCode = intRetVal
16890   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint03_Click()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint03_Click"

16910     NY_CmdPrint03_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
16920   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint03_KeyDown(KeyCode As Integer, Shift As Integer)

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint03_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

17010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
17020   intShiftDown = (Shift And acShiftMask) > 0
17030   intAltDown = (Shift And acAltMask) > 0
17040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
17050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
17060     Select Case intRetVal
          Case vbKeyUp
17070       With Me
17080         intRetVal = 0
17090         .cmdPrint02.SetFocus
17100       End With
17110     Case vbKeyDown
17120       With Me
17130         intRetVal = 0
17140         .cmdPrint04.SetFocus
17150       End With
17160     End Select
17170   End If

EXITP:
17180   KeyCode = intRetVal
17190   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord03_Click()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord03_Click"

17210     blnAllCancel = False
17220     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
17230     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

17240     NY_CmdWord03_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
17250   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord03_KeyDown(KeyCode As Integer, Shift As Integer)

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord03_KeyDown"

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
          Case vbKeyUp
17370       With Me
17380         intRetVal = 0
17390         .cmdWord02.SetFocus
17400       End With
17410     Case vbKeyDown
17420       With Me
17430         intRetVal = 0
17440         .cmdWord04.SetFocus
17450       End With
17460     End Select
17470   End If

EXITP:
17480   KeyCode = intRetVal
17490   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel03_Click()

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel03_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

17510     blnAllCancel = False
17520     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
17530     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

17540     NY_CmdExcel03_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
17550   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel03_KeyDown(KeyCode As Integer, Shift As Integer)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel03_KeyDown"

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
          Case vbKeyUp
17670       With Me
17680         intRetVal = 0
17690         .cmdExcel02.SetFocus
17700       End With
17710     Case vbKeyDown
17720       With Me
17730         intRetVal = 0
17740         .cmdExcel04.SetFocus
17750       End With
17760     End Select
17770   End If

EXITP:
17780   KeyCode = intRetVal
17790   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview04_Click()

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_Click"

17810     NY_CmdPrev04_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
17820   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview04_lbl4_cmd_DblClick(Cancel As Integer)

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_lbl4_cmd_DblClick"

17910   With Me
17920     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
17930     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
17940   End With

EXITP:
17950   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPreview04_KeyDown(KeyCode As Integer, Shift As Integer)

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview04_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18020   intShiftDown = (Shift And acShiftMask) > 0
18030   intAltDown = (Shift And acAltMask) > 0
18040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18060     Select Case intRetVal
          Case vbKeyUp
18070       With Me
18080         intRetVal = 0
18090         .cmdPreview03.SetFocus
18100       End With
18110     Case vbKeyDown
18120       With Me
18130         intRetVal = 0
18140         If .cmdPreview13.Enabled = True Then
18150           .cmdPreview13.SetFocus
18160         Else
18170           .cmdPreview05.SetFocus
18180         End If
18190       End With
18200     End Select
18210   End If

EXITP:
18220   KeyCode = intRetVal
18230   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint04_Click()

18300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint04_Click"

18310     NY_CmdPrint04_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
18320   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint04_KeyDown(KeyCode As Integer, Shift As Integer)

18400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint04_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18420   intShiftDown = (Shift And acShiftMask) > 0
18430   intAltDown = (Shift And acAltMask) > 0
18440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18460     Select Case intRetVal
          Case vbKeyUp
18470       With Me
18480         intRetVal = 0
18490         .cmdPrint03.SetFocus
18500       End With
18510     Case vbKeyDown
18520       With Me
18530         intRetVal = 0
18540         If .cmdPrint13.Enabled = True Then
18550           .cmdPrint13.SetFocus
18560         Else
18570           .cmdPrint05.SetFocus
18580         End If
18590       End With
18600     End Select
18610   End If

EXITP:
18620   KeyCode = intRetVal
18630   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord04_Click()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord04_Click"

18710     blnAllCancel = False
18720     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
18730     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

18740     NY_CmdWord04_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
18750   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord04_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord04_KeyDown"

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
          Case vbKeyUp
18870       With Me
18880         intRetVal = 0
18890         .cmdWord03.SetFocus
18900       End With
18910     Case vbKeyDown
18920       With Me
18930         intRetVal = 0
18940         If .cmdWord13.Enabled = True Then
18950           .cmdWord13.SetFocus
18960         Else
18970           .cmdWord05.SetFocus
18980         End If
18990       End With
19000     End Select
19010   End If

EXITP:
19020   KeyCode = intRetVal
19030   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel04_Click()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel04_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

19110     blnAllCancel = False
19120     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
19130     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

19140     NY_CmdExcel04_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
19150   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel04_KeyDown(KeyCode As Integer, Shift As Integer)

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel04_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19220   intShiftDown = (Shift And acShiftMask) > 0
19230   intAltDown = (Shift And acAltMask) > 0
19240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19260     Select Case intRetVal
          Case vbKeyUp
19270       With Me
19280         intRetVal = 0
19290         .cmdExcel03.SetFocus
19300       End With
19310     Case vbKeyDown
19320       With Me
19330         intRetVal = 0
19340         If .cmdExcel13.Enabled = True Then
19350           .cmdExcel13.SetFocus
19360         Else
19370           .cmdExcel05.SetFocus
19380         End If
19390       End With
19400     End Select
19410   End If

EXITP:
19420   KeyCode = intRetVal
19430   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview05_Click()

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview05_Click"

19510     NY_CmdPrev05_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
19520   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview05_KeyDown(KeyCode As Integer, Shift As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview05_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19620   intShiftDown = (Shift And acShiftMask) > 0
19630   intAltDown = (Shift And acAltMask) > 0
19640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19660     Select Case intRetVal
          Case vbKeyUp
19670       With Me
19680         intRetVal = 0
19690         If .cmdPreview13.Enabled = True Then
19700           .cmdPreview13.SetFocus
19710         Else
19720           .cmdPreview04.SetFocus
19730         End If
19740       End With
19750     Case vbKeyDown
19760       With Me
19770         intRetVal = 0
19780         .cmdPreview06.SetFocus
19790       End With
19800     End Select
19810   End If

EXITP:
19820   KeyCode = intRetVal
19830   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint05_Click()

19900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint05_Click"

19910     NY_CmdPrint05_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
19920   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint05_KeyDown(KeyCode As Integer, Shift As Integer)

20000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint05_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20020   intShiftDown = (Shift And acShiftMask) > 0
20030   intAltDown = (Shift And acAltMask) > 0
20040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20060     Select Case intRetVal
          Case vbKeyUp
20070       With Me
20080         intRetVal = 0
20090         If .cmdPrint13.Enabled = True Then
20100           .cmdPrint13.SetFocus
20110         Else
20120           .cmdPrint04.SetFocus
20130         End If
20140       End With
20150     Case vbKeyDown
20160       With Me
20170         intRetVal = 0
20180         .cmdPrint06.SetFocus
20190       End With
20200     End Select
20210   End If

EXITP:
20220   KeyCode = intRetVal
20230   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord05_Click()

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord05_Click"

20310     blnAllCancel = False
20320     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
20330     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

20340     NY_CmdWord05_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
20350   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord05_KeyDown(KeyCode As Integer, Shift As Integer)

20400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord05_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20420   intShiftDown = (Shift And acShiftMask) > 0
20430   intAltDown = (Shift And acAltMask) > 0
20440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20460     Select Case intRetVal
          Case vbKeyUp
20470       With Me
20480         intRetVal = 0
20490         If .cmdWord13.Enabled = True Then
20500           .cmdWord13.SetFocus
20510         Else
20520           .cmdWord04.SetFocus
20530         End If
20540       End With
20550     Case vbKeyDown
20560       With Me
20570         intRetVal = 0
20580         .cmdWord06.SetFocus
20590       End With
20600     End Select
20610   End If

EXITP:
20620   KeyCode = intRetVal
20630   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel05_Click()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel05_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

20710     blnAllCancel = False
20720     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
20730     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

20740     NY_CmdExcel05_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
20750   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel05_KeyDown(KeyCode As Integer, Shift As Integer)

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel05_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20820   intShiftDown = (Shift And acShiftMask) > 0
20830   intAltDown = (Shift And acAltMask) > 0
20840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20860     Select Case intRetVal
          Case vbKeyUp
20870       With Me
20880         intRetVal = 0
20890         If .cmdExcel13.Enabled = True Then
20900           .cmdExcel13.SetFocus
20910         Else
20920           .cmdExcel04.SetFocus
20930         End If
20940       End With
20950     Case vbKeyDown
20960       With Me
20970         intRetVal = 0
20980         .cmdExcel06.SetFocus
20990       End With
21000     End Select
21010   End If

EXITP:
21020   KeyCode = intRetVal
21030   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview06_Click()

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview06_Click"

21110     NY_CmdPrev06_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
21120   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview06_KeyDown(KeyCode As Integer, Shift As Integer)

21200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview06_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21220   intShiftDown = (Shift And acShiftMask) > 0
21230   intAltDown = (Shift And acAltMask) > 0
21240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21260     Select Case intRetVal
          Case vbKeyUp
21270       With Me
21280         intRetVal = 0
21290         .cmdPreview05.SetFocus
21300       End With
21310     Case vbKeyDown
21320       With Me
21330         intRetVal = 0
21340         .cmdPreview07.SetFocus
21350       End With
21360     End Select
21370   End If

EXITP:
21380   KeyCode = intRetVal
21390   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint06_Click()

21400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint06_Click"

21410     NY_CmdPrint06_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
21420   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint06_KeyDown(KeyCode As Integer, Shift As Integer)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint06_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21520   intShiftDown = (Shift And acShiftMask) > 0
21530   intAltDown = (Shift And acAltMask) > 0
21540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21560     Select Case intRetVal
          Case vbKeyUp
21570       With Me
21580         intRetVal = 0
21590         .cmdPrint05.SetFocus
21600       End With
21610     Case vbKeyDown
21620       With Me
21630         intRetVal = 0
21640         .cmdPrint07.SetFocus
21650       End With
21660     End Select
21670   End If

EXITP:
21680   KeyCode = intRetVal
21690   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord06_Click()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord06_Click"

21710     blnAllCancel = False
21720     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
21730     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

21740     NY_CmdWord06_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
21750   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord06_KeyDown(KeyCode As Integer, Shift As Integer)

21800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord06_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21820   intShiftDown = (Shift And acShiftMask) > 0
21830   intAltDown = (Shift And acAltMask) > 0
21840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21860     Select Case intRetVal
          Case vbKeyUp
21870       With Me
21880         intRetVal = 0
21890         .cmdWord05.SetFocus
21900       End With
21910     Case vbKeyDown
21920       With Me
21930         intRetVal = 0
21940         .cmdWord07.SetFocus
21950       End With
21960     End Select
21970   End If

EXITP:
21980   KeyCode = intRetVal
21990   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel06_Click()

22000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel06_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

22010     blnAllCancel = False
22020     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
22030     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

22040     NY_CmdExcel06_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
22050   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel06_KeyDown(KeyCode As Integer, Shift As Integer)

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel06_KeyDown"

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
          Case vbKeyUp
22170       With Me
22180         intRetVal = 0
22190         .cmdExcel05.SetFocus
22200       End With
22210     Case vbKeyDown
22220       With Me
22230         intRetVal = 0
22240         .cmdExcel07.SetFocus
22250       End With
22260     End Select
22270   End If

EXITP:
22280   KeyCode = intRetVal
22290   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview07_Click()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview07_Click"

22310     NY_CmdPrev07_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
22320   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview07_KeyDown(KeyCode As Integer, Shift As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview07_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22420   intShiftDown = (Shift And acShiftMask) > 0
22430   intAltDown = (Shift And acAltMask) > 0
22440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22460     Select Case intRetVal
          Case vbKeyUp
22470       With Me
22480         intRetVal = 0
22490         .cmdPreview06.SetFocus
22500       End With
22510     Case vbKeyDown
22520       With Me
22530         intRetVal = 0
22540         .cmdPreview08.SetFocus
22550       End With
22560     End Select
22570   End If

EXITP:
22580   KeyCode = intRetVal
22590   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint07_Click()

22600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint07_Click"

22610     NY_CmdPrint07_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
22620   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint07_KeyDown(KeyCode As Integer, Shift As Integer)

22700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint07_KeyDown"

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
          Case vbKeyUp
22770       With Me
22780         intRetVal = 0
22790         .cmdPrint06.SetFocus
22800       End With
22810     Case vbKeyDown
22820       With Me
22830         intRetVal = 0
22840         .cmdPrint08.SetFocus
22850       End With
22860     End Select
22870   End If

EXITP:
22880   KeyCode = intRetVal
22890   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord07_Click()

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord07_Click"

22910     blnAllCancel = False
22920     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
22930     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

22940     NY_CmdWord07_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
22950   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord07_KeyDown(KeyCode As Integer, Shift As Integer)

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord07_KeyDown"

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
          Case vbKeyUp
23070       With Me
23080         intRetVal = 0
23090         .cmdWord06.SetFocus
23100       End With
23110     Case vbKeyDown
23120       With Me
23130         intRetVal = 0
23140         .cmdWord08.SetFocus
23150       End With
23160     End Select
23170   End If

EXITP:
23180   KeyCode = intRetVal
23190   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel07_Click()

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel07_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

23210     blnAllCancel = False
23220     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
23230     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

23240     NY_CmdExcel07_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
23250   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel07_KeyDown(KeyCode As Integer, Shift As Integer)

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel07_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23320   intShiftDown = (Shift And acShiftMask) > 0
23330   intAltDown = (Shift And acAltMask) > 0
23340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23360     Select Case intRetVal
          Case vbKeyUp
23370       With Me
23380         intRetVal = 0
23390         .cmdExcel06.SetFocus
23400       End With
23410     Case vbKeyDown
23420       With Me
23430         intRetVal = 0
23440         .cmdExcel08.SetFocus
23450       End With
23460     End Select
23470   End If

EXITP:
23480   KeyCode = intRetVal
23490   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview08_Click()

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview08_Click"

23510     NY_CmdPrev08_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
23520   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview08_KeyDown(KeyCode As Integer, Shift As Integer)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview08_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23620   intShiftDown = (Shift And acShiftMask) > 0
23630   intAltDown = (Shift And acAltMask) > 0
23640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23660     Select Case intRetVal
          Case vbKeyUp
23670       With Me
23680         intRetVal = 0
23690         .cmdPreview07.SetFocus
23700       End With
23710     Case vbKeyDown
23720       With Me
23730         intRetVal = 0
23740         .cmdPreview09.SetFocus
23750       End With
23760     End Select
23770   End If

EXITP:
23780   KeyCode = intRetVal
23790   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint08_Click()

23800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint08_Click"

23810     NY_CmdPrint08_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
23820   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint08_KeyDown(KeyCode As Integer, Shift As Integer)

23900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint08_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23920   intShiftDown = (Shift And acShiftMask) > 0
23930   intAltDown = (Shift And acAltMask) > 0
23940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23960     Select Case intRetVal
          Case vbKeyUp
23970       With Me
23980         intRetVal = 0
23990         .cmdPrint07.SetFocus
24000       End With
24010     Case vbKeyDown
24020       With Me
24030         intRetVal = 0
24040         .cmdPrint09.SetFocus
24050       End With
24060     End Select
24070   End If

EXITP:
24080   KeyCode = intRetVal
24090   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord08_Click()

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord08_Click"

24110     blnAllCancel = False
24120     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
24130     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

24140     NY_CmdWord08_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
24150   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord08_KeyDown(KeyCode As Integer, Shift As Integer)

24200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord08_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24220   intShiftDown = (Shift And acShiftMask) > 0
24230   intAltDown = (Shift And acAltMask) > 0
24240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24260     Select Case intRetVal
          Case vbKeyUp
24270       With Me
24280         intRetVal = 0
24290         .cmdWord07.SetFocus
24300       End With
24310     Case vbKeyDown
24320       With Me
24330         intRetVal = 0
24340         .cmdWord09.SetFocus
24350       End With
24360     End Select
24370   End If

EXITP:
24380   KeyCode = intRetVal
24390   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel08_Click()

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel08_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

24410     blnAllCancel = False
24420     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
24430     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

24440     NY_CmdExcel08_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
24450   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel08_KeyDown(KeyCode As Integer, Shift As Integer)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel08_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24520   intShiftDown = (Shift And acShiftMask) > 0
24530   intAltDown = (Shift And acAltMask) > 0
24540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24560     Select Case intRetVal
          Case vbKeyUp
24570       With Me
24580         intRetVal = 0
24590         .cmdExcel07.SetFocus
24600       End With
24610     Case vbKeyDown
24620       With Me
24630         intRetVal = 0
24640         .cmdExcel09.SetFocus
24650       End With
24660     End Select
24670   End If

EXITP:
24680   KeyCode = intRetVal
24690   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview09_lbl4_cmd_DblClick(Cancel As Integer)

24700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview09_lbl4_cmd_DblClick"

24710   With Me
24720     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
24730     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
24740   End With

EXITP:
24750   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPreview09_Click()

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview09_Click"

24810     NY_CmdPrev09_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
24820   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview09_KeyDown(KeyCode As Integer, Shift As Integer)

24900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview09_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24920   intShiftDown = (Shift And acShiftMask) > 0
24930   intAltDown = (Shift And acAltMask) > 0
24940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24960     Select Case intRetVal
          Case vbKeyUp
24970       With Me
24980         intRetVal = 0
24990         .cmdPreview08.SetFocus
25000       End With
25010     Case vbKeyDown
25020       With Me
25030         intRetVal = 0
25040         .cmdPreview10.SetFocus
25050       End With
25060     End Select
25070   End If

EXITP:
25080   KeyCode = intRetVal
25090   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint09_Click()

25100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint09_Click"

25110     NY_CmdPrint09_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
25120   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint09_KeyDown(KeyCode As Integer, Shift As Integer)

25200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint09_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25220   intShiftDown = (Shift And acShiftMask) > 0
25230   intAltDown = (Shift And acAltMask) > 0
25240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25260     Select Case intRetVal
          Case vbKeyUp
25270       With Me
25280         intRetVal = 0
25290         .cmdPrint08.SetFocus
25300       End With
25310     Case vbKeyDown
25320       With Me
25330         intRetVal = 0
25340         .cmdPrint10.SetFocus
25350       End With
25360     End Select
25370   End If

EXITP:
25380   KeyCode = intRetVal
25390   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord09_Click()

25400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord09_Click"

25410     blnAllCancel = False
25420     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
25430     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

25440     NY_CmdWord09_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
25450   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord09_KeyDown(KeyCode As Integer, Shift As Integer)

25500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord09_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25520   intShiftDown = (Shift And acShiftMask) > 0
25530   intAltDown = (Shift And acAltMask) > 0
25540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25560     Select Case intRetVal
          Case vbKeyUp
25570       With Me
25580         intRetVal = 0
25590         .cmdWord08.SetFocus
25600       End With
25610     Case vbKeyDown
25620       With Me
25630         intRetVal = 0
25640         .cmdWord10.SetFocus
25650       End With
25660     End Select
25670   End If

EXITP:
25680   KeyCode = intRetVal
25690   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel09_Click()

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel09_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

25710     blnAllCancel = False
25720     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
25730     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

25740     NY_CmdExcel09_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
25750   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel09_KeyDown(KeyCode As Integer, Shift As Integer)

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
          Case vbKeyUp
25870       With Me
25880         intRetVal = 0
25890         .cmdExcel08.SetFocus
25900       End With
25910     Case vbKeyDown
25920       With Me
25930         intRetVal = 0
25940         .cmdExcel10.SetFocus
25950       End With
25960     End Select
25970   End If

EXITP:
25980   KeyCode = intRetVal
25990   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview10_lbl4_cmd_DblClick(Cancel As Integer)

26000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview10_lbl4_cmd_DblClick"

26010   With Me
26020     .chkGroupBy_IncExpCode = Not .chkGroupBy_IncExpCode
26030     chkGroupBy_IncExpCode_AfterUpdate  ' ** Procedure: Above.
26040   End With

EXITP:
26050   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPreview10_Click()

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview10_Click"

26110     NY_CmdPrev10_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
26120   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview10_KeyDown(KeyCode As Integer, Shift As Integer)

26200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdpreview10_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26220   intShiftDown = (Shift And acShiftMask) > 0
26230   intAltDown = (Shift And acAltMask) > 0
26240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26260     Select Case intRetVal
          Case vbKeyUp
26270       With Me
26280         intRetVal = 0
26290         .cmdPreview09.SetFocus
26300       End With
26310     Case vbKeyDown
26320       With Me
26330         intRetVal = 0
26340         .cmdPreview11.SetFocus
26350       End With
26360     End Select
26370   End If

EXITP:
26380   KeyCode = intRetVal
26390   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint10_Click()

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint10_Click"

26410     NY_CmdPrint10_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
26420   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint10_KeyDown(KeyCode As Integer, Shift As Integer)

26500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint10_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26520   intShiftDown = (Shift And acShiftMask) > 0
26530   intAltDown = (Shift And acAltMask) > 0
26540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26560     Select Case intRetVal
          Case vbKeyUp
26570       With Me
26580         intRetVal = 0
26590         .cmdPrint09.SetFocus
26600       End With
26610     Case vbKeyDown
26620       With Me
26630         intRetVal = 0
26640         .cmdPrint11.SetFocus
26650       End With
26660     End Select
26670   End If

EXITP:
26680   KeyCode = intRetVal
26690   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord10_Click()

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord10_Click"

26710     blnAllCancel = False
26720     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
26730     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

26740     NY_CmdWord10_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
26750   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord10_KeyDown(KeyCode As Integer, Shift As Integer)

26800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord10_KeyDown"

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
          Case vbKeyUp
26870       With Me
26880         intRetVal = 0
26890         .cmdWord09.SetFocus
26900       End With
26910     Case vbKeyDown
26920       With Me
26930         intRetVal = 0
26940         .cmdWord11.SetFocus
26950       End With
26960     End Select
26970   End If

EXITP:
26980   KeyCode = intRetVal
26990   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel10_Click()

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel10_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

27010     blnAllCancel = False
27020     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
27030     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

27040     NY_CmdExcel10_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
27050   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel10_KeyDown(KeyCode As Integer, Shift As Integer)

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel10_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27120   intShiftDown = (Shift And acShiftMask) > 0
27130   intAltDown = (Shift And acAltMask) > 0
27140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27160     Select Case intRetVal
          Case vbKeyUp
27170       With Me
27180         intRetVal = 0
27190         .cmdExcel09.SetFocus
27200       End With
27210     Case vbKeyDown
27220       With Me
27230         intRetVal = 0
27240         .cmdExcel11.SetFocus
27250       End With
27260     End Select
27270   End If

EXITP:
27280   KeyCode = intRetVal
27290   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview11_Click()

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview11_Click"

27310     NY_CmdPrev11_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
27320   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview11_KeyDown(KeyCode As Integer, Shift As Integer)

27400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview11_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27410   intRetVal = KeyCode
27420   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
27430   intShiftDown = (Shift And acShiftMask) > 0
27440   intAltDown = (Shift And acAltMask) > 0
27450   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27460   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27470     Select Case intRetVal
          Case vbKeyUp
27480       With Me
27490         intRetVal = 0
27500         .cmdPreview10.SetFocus
27510       End With
27520     Case vbKeyDown
27530       With Me
27540         intRetVal = 0
27550         .cmdPreview12.SetFocus
27560       End With
27570     End Select
27580   End If

EXITP:
27590   KeyCode = intRetVal
27600   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint11_Click()

27700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint11_Click"

27710     NY_CmdPrint11_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
27720   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint11_KeyDown(KeyCode As Integer, Shift As Integer)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint11_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27810   intRetVal = KeyCode
27820   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
27830   intShiftDown = (Shift And acShiftMask) > 0
27840   intAltDown = (Shift And acAltMask) > 0
27850   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27860   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27870     Select Case intRetVal
          Case vbKeyUp
27880       With Me
27890         intRetVal = 0
27900         .cmdPrint10.SetFocus
27910       End With
27920     Case vbKeyDown
27930       With Me
27940         intRetVal = 0
27950         .cmdPrint12.SetFocus
27960       End With
27970     End Select
27980   End If

EXITP:
27990   KeyCode = intRetVal
28000   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord11_Click()

28100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord11_Click"

28110     blnAllCancel = False
28120     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
28130     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

28140     NY_CmdWord11_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
28150   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord11_KeyDown(KeyCode As Integer, Shift As Integer)

28200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord11_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28210   intRetVal = KeyCode
28220   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
28230   intShiftDown = (Shift And acShiftMask) > 0
28240   intAltDown = (Shift And acAltMask) > 0
28250   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28260   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28270     Select Case intRetVal
          Case vbKeyUp
28280       With Me
28290         intRetVal = 0
28300         .cmdWord10.SetFocus
28310       End With
28320     Case vbKeyDown
28330       With Me
28340         intRetVal = 0
28350         .cmdWord12.SetFocus
28360       End With
28370     End Select
28380   End If

EXITP:
28390   KeyCode = intRetVal
28400   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel11_Click()

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel11_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

28510     blnAllCancel = False
28520     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
28530     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

28540     NY_CmdExcel11_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
28550   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel11_KeyDown(KeyCode As Integer, Shift As Integer)

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel11_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

28610   intRetVal = KeyCode
28620   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
28630   intShiftDown = (Shift And acShiftMask) > 0
28640   intAltDown = (Shift And acAltMask) > 0
28650   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
28660   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
28670     Select Case intRetVal
          Case vbKeyUp
28680       With Me
28690         intRetVal = 0
28700         .cmdExcel10.SetFocus
28710       End With
28720     Case vbKeyDown
28730       With Me
28740         intRetVal = 0
28750         .cmdExcel12.SetFocus
28760       End With
28770     End Select
28780   End If

EXITP:
28790   KeyCode = intRetVal
28800   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview12_Click()

28900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview12_Click"

28910     NY_CmdPrev12_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
28920   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview12_KeyDown(KeyCode As Integer, Shift As Integer)

29000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview12_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29010   intRetVal = KeyCode
29020   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
29030   intShiftDown = (Shift And acShiftMask) > 0
29040   intAltDown = (Shift And acAltMask) > 0
29050   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29060   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29070     Select Case intRetVal
          Case vbKeyUp
29080       With Me
29090         intRetVal = 0
29100         .cmdPreview11.SetFocus
29110       End With
29120     Case vbKeyDown
29130       With Me
29140         intRetVal = 0
29150         .cmdPreview00B.SetFocus
29160       End With
29170     End Select
29180   End If

EXITP:
29190   KeyCode = intRetVal
29200   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint12_Click()

29300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint12_Click"

29310     NY_CmdPrint12_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
29320   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint12_KeyDown(KeyCode As Integer, Shift As Integer)

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint12_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29410   intRetVal = KeyCode
29420   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
29430   intShiftDown = (Shift And acShiftMask) > 0
29440   intAltDown = (Shift And acAltMask) > 0
29450   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29460   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29470     Select Case intRetVal
          Case vbKeyUp
29480       With Me
29490         intRetVal = 0
29500         .cmdPrint11.SetFocus
29510       End With
29520     Case vbKeyDown
29530       With Me
29540         intRetVal = 0
29550         .cmdPrint00B.SetFocus
29560       End With
29570     End Select
29580   End If

EXITP:
29590   KeyCode = intRetVal
29600   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord12_Click()

29700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord12_Click"

29710     blnAllCancel = False
29720     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
29730     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

29740     NY_CmdWord12_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
29750   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord12_KeyDown(KeyCode As Integer, Shift As Integer)

29800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord12_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

29810   intRetVal = KeyCode
29820   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
29830   intShiftDown = (Shift And acShiftMask) > 0
29840   intAltDown = (Shift And acAltMask) > 0
29850   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
29860   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
29870     Select Case intRetVal
          Case vbKeyUp
29880       With Me
29890         intRetVal = 0
29900         .cmdWord11.SetFocus
29910       End With
29920     Case vbKeyDown
29930       With Me
29940         intRetVal = 0
29950         .cmdWord00B.SetFocus
29960       End With
29970     Case vbKeyRight
29980       With Me
29990         intRetVal = 0
30000         If .cmdExcel12.Enabled = True Then
30010           .cmdExcel12.SetFocus
30020         Else
30030           .cmdPreview00B.SetFocus
30040         End If
30050       End With
30060     End Select
30070   End If

EXITP:
30080   KeyCode = intRetVal
30090   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel12_Click()

30100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel12_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

30110     blnAllCancel = False
30120     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
30130     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

30140     NY_CmdExcel12_Click strRptName, strRptCap, strRptPath, strRptPathFile, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
30150   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case 70  ' ** Permission denied.
570       Beep
580       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
590     Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub cmdExcel12_KeyDown(KeyCode As Integer, Shift As Integer)

30200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel12_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30210   intRetVal = KeyCode
30220   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
30230   intShiftDown = (Shift And acShiftMask) > 0
30240   intAltDown = (Shift And acAltMask) > 0
30250   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30260   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30270     Select Case intRetVal
          Case vbKeyUp
30280       With Me
30290         intRetVal = 0
30300         .cmdExcel11.SetFocus
30310       End With
30320     Case vbKeyDown
30330       With Me
30340         intRetVal = 0
30350         .cmdExcel00B.SetFocus
30360       End With
30370     Case vbKeyRight
30380       With Me
30390         intRetVal = 0
30400         .cmdPreview00B.SetFocus
30410       End With
30420     End Select
30430   End If

EXITP:
30440   KeyCode = intRetVal
30450   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview13_Click()

30500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview13_Click"

30510     NY_CmdPrev13_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
30520   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview13_KeyDown(KeyCode As Integer, Shift As Integer)

30600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview13_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30620   intShiftDown = (Shift And acShiftMask) > 0
30630   intAltDown = (Shift And acAltMask) > 0
30640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30660     Select Case intRetVal
          Case vbKeyUp
30670       With Me
30680         intRetVal = 0
30690         .cmdPreview04.SetFocus
30700       End With
30710     Case vbKeyDown
30720       With Me
30730         intRetVal = 0
30740         .cmdPreview05.SetFocus
30750       End With
30760     End Select
30770   End If

EXITP:
30780   KeyCode = intRetVal
30790   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint13_Click()

30800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint13_Click"

30810     NY_CmdPrint13_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
30820   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint13_KeyDown(KeyCode As Integer, Shift As Integer)

30900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint13_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

30910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
30920   intShiftDown = (Shift And acShiftMask) > 0
30930   intAltDown = (Shift And acAltMask) > 0
30940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
30950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
30960     Select Case intRetVal
          Case vbKeyUp
30970       With Me
30980         intRetVal = 0
30990         .cmdPrint04.SetFocus
31000       End With
31010     Case vbKeyDown
31020       With Me
31030         intRetVal = 0
31040         .cmdPrint05.SetFocus
31050       End With
31060     End Select
31070   End If

EXITP:
31080   KeyCode = intRetVal
31090   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord13_Click()

31100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord13_Click"

31110     blnAllCancel = False
31120     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
31130     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

31140     NY_CmdWord13_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

EXITP:
31150   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWord13_KeyDown(KeyCode As Integer, Shift As Integer)

31200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord13_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31220   intShiftDown = (Shift And acShiftMask) > 0
31230   intAltDown = (Shift And acAltMask) > 0
31240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31260     Select Case intRetVal
          Case vbKeyUp
31270       With Me
31280         intRetVal = 0
31290         .cmdWord04.SetFocus
31300       End With
31310     Case vbKeyDown
31320       With Me
31330         intRetVal = 0
31340         .cmdWord05.SetFocus
31350       End With
31360     End Select
31370   End If

EXITP:
31380   KeyCode = intRetVal
31390   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel13_Click()

31400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel13_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

31410     blnAllCancel = False
31420     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
31430     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

31440     NY_CmdExcel13_Click blnRebuildTable, THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY2.

      #End If

EXITP:
31450   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case 70  ' ** Permission denied.
560       Beep
570       MsgBox "The file Trust Accountant is trying to save is already open." & vbCrLf & vbCrLf & _
            "Please close it and try again.", vbInformation + vbOKOnly, "File Is Open"
580     Case Else
590       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
600     End Select
610     Resume EXITP

End Sub

Private Sub cmdExcel13_KeyDown(KeyCode As Integer, Shift As Integer)

31500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel13_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31520   intShiftDown = (Shift And acShiftMask) > 0
31530   intAltDown = (Shift And acAltMask) > 0
31540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31560     Select Case intRetVal
          Case vbKeyUp
31570       With Me
31580         intRetVal = 0
31590         .cmdExcel04.SetFocus
31600       End With
31610     Case vbKeyDown
31620       With Me
31630         intRetVal = 0
31640         .cmdExcel05.SetFocus
31650       End With
31660     End Select
31670   End If

EXITP:
31680   KeyCode = intRetVal
31690   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub chkAssetList_AfterUpdate()

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_AfterUpdate"

31710   With Me
31720     Select Case .chkAssetList
          Case True
31730       .chkAssetList_lbl.FontBold = True
31740       .chkAssetList_lbl2.FontBold = True
31750     Case False
31760       .chkAssetList_lbl.FontBold = False
31770       .chkAssetList_lbl2.FontBold = False
31780     End Select
31790   End With

EXITP:
31800   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkAssetList_KeyDown(KeyCode As Integer, Shift As Integer)

31900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkAssetList_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

31910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
31920   intShiftDown = (Shift And acShiftMask) > 0
31930   intAltDown = (Shift And acAltMask) > 0
31940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
31950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
31960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
31970       With Me
31980         intRetVal = 0
31990         .cmdPreview00B.SetFocus
32000       End With
32010     End Select
32020   End If

        ' ** Shift keys.
32030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
32040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
32050       With Me
32060         intRetVal = 0
32070         If .cmdExcel12.Enabled = True Then
32080           .cmdExcel12.SetFocus
32090         Else
32100           .cmdWord12.SetFocus
32110         End If
32120       End With
32130     End Select
32140   End If

EXITP:
32150   KeyCode = intRetVal
32160   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPreview00B_Click()

32200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_Click"

32210   AssetList_PreviewPrint_NY acViewPreview, THIS_NAME, Me, THIS_PROC  ' ** Module Function: modCourtReportsNY1.

EXITP:
32220   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPreview00B_KeyDown(KeyCode As Integer, Shift As Integer)

32300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview00B_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32310   intRetVal = KeyCode
32320   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
32330   intShiftDown = (Shift And acShiftMask) > 0
32340   intAltDown = (Shift And acAltMask) > 0
32350   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32360   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32370     Select Case intRetVal
          Case vbKeyUp
32380       With Me
32390         intRetVal = 0
32400         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
32410         .cmdPreview12.SetFocus
32420       End With
32430     Case vbKeyDown
32440       With Me
32450         intRetVal = 0
32460         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
32470         .cmdPreview00.SetFocus
32480       End With
32490     Case vbKeyLeft
32500       With Me
32510         intRetVal = 0
32520         If .cmdExcel12.Enabled = True Then
32530           .cmdExcel12.SetFocus
32540         Else
32550           .cmdWord12.SetFocus
32560         End If
32570       End With
32580     End Select
32590   End If

EXITP:
32600   KeyCode = intRetVal
32610   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrint00B_Click()

32700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00B_Click"

        '##GTR_Ref: rptCourtRptNY_00B
        '##GTR_Ref: rptCourtRptNY_00BA
32710   AssetList_PreviewPrint_NY acViewNormal, THIS_NAME, Me, THIS_PROC  ' ** Module Function: modCourtReportsNY1.

EXITP:
32720   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPrint00B_KeyDown(KeyCode As Integer, Shift As Integer)

32800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint00B_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

32810   intRetVal = KeyCode
32820   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
32830   intShiftDown = (Shift And acShiftMask) > 0
32840   intAltDown = (Shift And acAltMask) > 0
32850   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
32860   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
32870     Select Case intRetVal
          Case vbKeyUp
32880       With Me
32890         intRetVal = 0
32900         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
32910         .cmdPrint12.SetFocus
32920       End With
32930     Case vbKeyDown
32940       With Me
32950         intRetVal = 0
32960         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
32970         .cmdPrintAll.SetFocus
32980       End With
32990     End Select
33000   End If

EXITP:
33010   KeyCode = intRetVal
33020   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdWord00B_Click()

33100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00B_Click"

33110     blnAllCancel = False
33120     AllCancelSet2_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY1.
33130     AllCancelSet3_NY blnAllCancel  ' ** Module Procedure: modCourtReportsNY2.

33140     AssetList_Word_NY THIS_NAME, Me  ' ** Module Procedure: modCourtReportsNY1.

EXITP:
33150   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdWord00B_KeyDown(KeyCode As Integer, Shift As Integer)

33200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord00B_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33210   intRetVal = KeyCode
33220   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
33230   intShiftDown = (Shift And acShiftMask) > 0
33240   intAltDown = (Shift And acAltMask) > 0
33250   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33260   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33270     Select Case intRetVal
          Case vbKeyUp
33280       With Me
33290         intRetVal = 0
33300         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
33310         .cmdWord12.SetFocus
33320       End With
33330     Case vbKeyDown
33340       With Me
33350         intRetVal = 0
33360         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
33370         .cmdWordAll.SetFocus
33380       End With
33390     Case vbKeyTab
33400       With Me
33410         intRetVal = 0
33420         If .cmdExcel00B.Enabled = True Then
33430           .cmdExcel00B.SetFocus
33440         Else
33450           .cmdPrintAll.SetFocus
33460         End If
33470       End With
33480     End Select
33490   End If

EXITP:
33500   KeyCode = intRetVal
33510   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub cmdExcel00B_Click()

33600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00B_Click"

      #If Not NoExcel Then  ' ** The buttons should be disabled anyway.

33610   Excel00B_Click_NY strRptCap, strRptPathFile, strRptPath, strRptName, lngCaps, arr_varCap, Me  ' ** Module Procedure: modCourtReportsNY3.

      #End If

EXITP:
33620   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdExcel00B_KeyDown(KeyCode As Integer, Shift As Integer)

33700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel00B_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

33710   intRetVal = KeyCode
33720   strLastArrow = vbNullString

        ' ** Use bit masks to determine which key was pressed.
33730   intShiftDown = (Shift And acShiftMask) > 0
33740   intAltDown = (Shift And acAltMask) > 0
33750   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
33760   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
33770     Select Case intRetVal
          Case vbKeyUp
33780       With Me
33790         intRetVal = 0
33800         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
33810         .cmdExcel12.SetFocus
33820       End With
33830     Case vbKeyDown
33840       With Me
33850         intRetVal = 0
33860         strLastArrow = Left(THIS_PROC, (InStr(THIS_PROC, "_") - 1))
33870         .cmdExcelAll.SetFocus
33880       End With
33890     Case vbKeyTab
33900       With Me
33910         intRetVal = 0
33920         .cmdPrintAll.SetFocus
33930       End With
33940     End Select
33950   End If

        ' ** Shift keys.
33960   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
33970     Select Case intRetVal
          Case vbKeyTab
33980       With Me
33990         intRetVal = 0
34000         .cmdWord00B.SetFocus
34010       End With
34020     End Select
34030   End If

EXITP:
34040   KeyCode = intRetVal
34050   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

34100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

34110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
34120   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

34200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

34210   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
34220   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

34300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

34310   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
34320   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

34400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

34410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
34420   intShiftDown = (Shift And acShiftMask) > 0
34430   intAltDown = (Shift And acAltMask) > 0
34440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
34450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
34460     Select Case intRetVal
          Case vbKeyTab
34470       With Me
34480         intRetVal = 0
34490         .cmdClose.SetFocus
34500       End With
34510     Case vbKeyUp
34520       With Me
34530         intRetVal = 0
34540         .cmdPrint00B.SetFocus
34550         strLastArrow = vbNullString
34560       End With
34570     Case vbKeyDown
34580       With Me
34590         intRetVal = 0
34600         .cmdPrint00.SetFocus
34610         strLastArrow = vbNullString
34620       End With
34630     Case vbKeyLeft
34640       With Me
34650         intRetVal = 0
34660         If .cmdExcel00B.Enabled = True Then
34670           .cmdExcel00B.SetFocus
34680         Else
34690           .cmdWord00B.SetFocus
34700         End If
34710       End With
34720     Case vbKeyRight
34730       With Me
34740         intRetVal = 0
34750         .cmdWordAll.SetFocus
34760       End With
34770     End Select
34780   End If

        ' ** Shift keys.
34790   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
34800     Select Case intRetVal
          Case vbKeyTab
34810       With Me
34820         intRetVal = 0
34830         If .cmdExcel00B.Enabled = True Then
34840           .cmdExcel00B.SetFocus
34850         Else
34860           .cmdWord00B.SetFocus
34870         End If
34880       End With
34890     End Select
34900   End If

        ' ** Ctrl keys.
34910   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
34920     Select Case intRetVal
          Case vbKeyTab
34930       With Me
34940         intRetVal = 0
34950         If .UserReportPath_chk.Visible = True And .UserReportPath_chk.Enabled = True Then
34960           .UserReportPath_chk.SetFocus
34970         End If
34980       End With
34990     End Select
35000   End If

EXITP:
35010   KeyCode = intRetVal
35020   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

35110   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
35120   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

35200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

35210   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
35220   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

35300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

35310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
35320   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

35400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

35410   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
35420   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

35500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

35510   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
35520   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

35600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

35610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
35620   intShiftDown = (Shift And acShiftMask) > 0
35630   intAltDown = (Shift And acAltMask) > 0
35640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
35650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
35660     Select Case intRetVal
          Case vbKeyTab
35670       With Me
35680         intRetVal = 0
35690         If .cmdExcelAll.Enabled = True Then
35700           .cmdExcelAll.SetFocus
35710         Else
35720           .cmdClose.SetFocus
35730         End If
35740       End With
35750     Case vbKeyUp
35760       With Me
35770         intRetVal = 0
35780         .cmdWord00B.SetFocus
35790         strLastArrow = vbNullString
35800       End With
35810     Case vbKeyDown
35820       With Me
35830         intRetVal = 0
35840         .cmdWord00.SetFocus
35850       End With
35860     Case vbKeyRight
35870       With Me
35880         intRetVal = 0
35890         If .cmdExcelAll.Enabled = True Then
35900           .cmdExcelAll.SetFocus
35910         Else
35920           .cmdClose.SetFocus
35930         End If
35940       End With
35950     End Select
35960   End If

        ' ** Shift keys.
35970   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
35980     Select Case intRetVal
          Case vbKeyTab
35990       With Me
36000         intRetVal = 0
36010         .cmdPrintAll.SetFocus
36020       End With
36030     End Select
36040   End If

        ' ** Ctrl keys.
36050   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
36060     Select Case intRetVal
          Case vbKeyTab
36070       With Me
36080         intRetVal = 0
36090         .chkOpenWord.SetFocus
36100       End With
36110     End Select
36120   End If

EXITP:
36130   KeyCode = intRetVal
36140   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

36210   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
36220   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

36300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

36310   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
36320   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub chkOpenWord_AfterUpdate()

36400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

36410   OpenWordExcel_After_NY THIS_PROC, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
36420   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

36500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intRetVal As Integer

36510   intRetVal = OpenWordExcel_Key_NY(KeyCode, Shift, 1, Me)  ' ** Module Function: modCourtReportsNY3.

EXITP:
36520   KeyCode = intRetVal
36530   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

36610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
36620   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

36700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

36710   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
36720   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

36800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

36810   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
36820   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

36900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

36910   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
36920   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

37000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

37010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
37020   intShiftDown = (Shift And acShiftMask) > 0
37030   intAltDown = (Shift And acAltMask) > 0
37040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
37050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
37060     Select Case intRetVal
          Case vbKeyTab
37070       With Me
37080         intRetVal = 0
37090         .cmdClose.SetFocus
37100       End With
37110     Case vbKeyUp
37120       With Me
37130         intRetVal = 0
37140         .cmdExcel00B.SetFocus
37150         strLastArrow = vbNullString
37160       End With
37170     Case vbKeyDown
37180       With Me
37190         intRetVal = 0
37200         .cmdExcel00.SetFocus
37210         strLastArrow = vbNullString
37220       End With
37230     Case vbKeyRight
37240       With Me
37250         intRetVal = 0
37260         .cmdClose.SetFocus
37270       End With
37280     End Select
37290   End If

        ' ** Shift keys.
37300   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
37310     Select Case intRetVal
          Case vbKeyTab
37320       With Me
37330         intRetVal = 0
37340         .cmdWordAll.SetFocus
37350       End With
37360     End Select
37370   End If

        ' ** Ctrl keys.
37380   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
37390     Select Case intRetVal
          Case vbKeyTab
37400       With Me
37410         intRetVal = 0
37420         .chkOpenExcel.SetFocus
37430       End With
37440     End Select
37450   End If

EXITP:
37460   KeyCode = intRetVal
37470   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

37500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

37510   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
37520   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

37600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

37610   PrintWordExcelAll_Handler_NY THIS_PROC, blnPrintAll_Focus, _
          blnWordAll_Focus, blnExcelAll_Focus, blnRebuildTable, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
37620   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub chkOpenExcel_AfterUpdate()

37700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

37710   OpenWordExcel_After_NY THIS_PROC, Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
37720   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub chkOpenExcel_KeyDown(KeyCode As Integer, Shift As Integer)

37800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_KeyDown"

        Dim intRetVal As Integer

37810   intRetVal = OpenWordExcel_Key_NY(KeyCode, Shift, 2, Me)  ' ** Module Function: modCourtReportsNY3.

EXITP:
37820   KeyCode = intRetVal
37830   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

37900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

37910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
37920   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Sub UserReportPath_chk_AfterUpdate()

38000 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_AfterUpdate"

38010   UserRptPath_After_NY Me  ' ** Module Procedure: modCourtReportsNY3.

EXITP:
38020   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub UserReportPath_chk_KeyDown(KeyCode As Integer, Shift As Integer)

38100 On Error GoTo ERRH

        Const THIS_PROC As String = "UserReportPath_chk_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

38110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
38120   intShiftDown = (Shift And acShiftMask) > 0
38130   intAltDown = (Shift And acAltMask) > 0
38140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
38150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
38160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
38170       With Me
38180         intRetVal = 0
38190         .cmdClose.SetFocus
38200       End With
38210     End Select
38220   End If

        ' ** Shift keys.
38230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
38240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
38250       With Me
38260         intRetVal = 0
38270         .cmdPrintAll.SetFocus
38280       End With
38290     End Select
38300   End If

EXITP:
38310   KeyCode = intRetVal
38320   Exit Sub

ERRH:
550     intRetVal = 0
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Function Validate() As Boolean

38400 On Error GoTo ERRH

        Const THIS_PROC As String = "Validate"

        Dim blnRetVal As Boolean

38410   With Me

38420     blnRetVal = False

          ' ** Validate main three data values.
38430     If IsNull(.DateStart.Value) Then
38440       DoCmd.Hourglass False
38450       MsgBox "You must enter a From date to continue.", vbInformation + vbOKOnly, "Entry Required"
38460       .DateStart.SetFocus
38470     Else
38480       If IsNull(.DateEnd.Value) Then
38490         DoCmd.Hourglass False
38500         MsgBox "You must enter a To date to continue.", vbInformation + vbOKOnly, "Entry Required"
38510         .DateEnd.SetFocus
38520       Else
38530         If .DateStart.Value >= .DateEnd.Value Then
38540           DoCmd.Hourglass False
38550           MsgBox "You must enter a From date that is less than the To date to continue.", vbInformation + vbOKOnly, "Invalid Date"
38560           .DateEnd.SetFocus
38570         Else
38580           If IsNull(.cmbAccounts.Column(0)) Or .cmbAccounts.Column(0) = vbNullString Then
38590             DoCmd.Hourglass False
38600             MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
38610             .cmbAccounts.SetFocus
38620           Else
38630             If FirstDate_NY(Me) = False Then  ' ** Module Funtion: modCourtReportsNY1.
38640               MsgBox strFirstDateMsg, vbInformation + vbOKOnly, "Nothing To Do"
38650             Else
38660               blnRetVal = True
38670             End If
38680           End If
38690         End If
38700       End If
38710     End If

38720   End With

EXITP:
38730   Validate = blnRetVal
38740   Exit Function

ERRH:
550     blnRetVal = False
560     DoCmd.Hourglass False
570     Select Case ERR.Number
        Case Else
580       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
590     End Select
600     Resume EXITP

End Function

Public Function FillVar() As Boolean

38800 On Error GoTo ERRH

        Const THIS_PROC As String = "FillVar"

        Dim blnRetVal As Boolean

38810   blnRetVal = True

38820   With Me
38830     gdatStartDate = .DateStart
38840     gdatEndDate = .DateEnd
38850     gstrAccountNo = .cmbAccounts.Column(0)
38860     gstrCrtRpt_NY_AccountNo = .cmbAccounts.Column(0)
38870     gstrAccountName = .cmbAccounts.Column(3)
38880     gdatCrtRpt_NY_DateStart = .DateStart
38890     gdatCrtRpt_NY_DateEnd = .DateEnd
38900   End With

EXITP:
38910   FillVar = blnRetVal
38920   Exit Function

ERRH:
550     blnRetVal = False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Function

Public Sub CapArray_Load()

39000 On Error GoTo ERRH

        Const THIS_PROC As String = "CapArray_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

39010   If lngCaps = 0& Or IsEmpty(arr_varCap) = True Then

39020     Set dbs = CurrentDb
39030     With dbs
            ' ** qryCourtReport_15a (tblReport, captions of Court Reports, with rpt_caption_newx),
            ' ** by specified GlobalVarGet(), FormRef('TypeShort'), [CrtTyp].
39040       Set qdf = .QueryDefs("qryCourtReport_15")
39050       With qdf.Parameters
39060         ![CrtTyp] = "NY"
39070       End With
39080       Set rst = qdf.OpenRecordset
39090       With rst
39100         .MoveLast
39110         lngCaps = .RecordCount
39120         .MoveFirst
39130         arr_varCap = .GetRows(lngCaps)
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
39140         .Close
39150       End With
39160       .Close
39170     End With

39180   End If

EXITP:
39190   Set rst = Nothing
39200   Set qdf = Nothing
39210   Set dbs = Nothing
39220   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Public Function CapArray_Get() As Variant

39300 On Error GoTo ERRH

        Const THIS_PROC As String = "CapArray_Get"

        Dim arr_varRetVal As Variant

39310   arr_varRetVal = arr_varCap

EXITP:
39320   CapArray_Get = arr_varRetVal
39330   Exit Function

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Function

Public Sub FirstDateMsg_Set(strMsg As String)

39400 On Error GoTo ERRH

        Const THIS_PROC As String = "FirstDateMsg_Set"

39410   strFirstDateMsg = strMsg

EXITP:
39420   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub FirstDate_Set(datDate As Date)

39500 On Error GoTo ERRH

        Const THIS_PROC As String = "FirstDate_Set"

39510   datFirstDate = datDate

EXITP:
39520   Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

End Sub

Public Sub AllCancelSet1_NY(blnCancel As Boolean)

39600 On Error GoTo ERRH

        Const THIS_PROC As String = "AllCancelSet1_NY"

39610   blnAllCancel = blnCancel

EXITP:
39620   Exit Sub

ERRH:
550     Select Case ERR.Number
        Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub
