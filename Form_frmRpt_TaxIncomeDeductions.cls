VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_TaxIncomeDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_TaxIncomeDeductions"

'VGC 09/29/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it opens, shows arrows, and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Account:          A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Deduct Word:      D {cmdTaxIncDed_DeductionWord}
' **   Income Excel:     E {cmdTaxIncDed_IncomeExcel}
' **   From:             F {DateStart}
' **   Include Archive:  H {chkIncludeArchive}
' **   All:              L {opgAccountNumber_optAll}
' **   Name:             M {opgAccountSource_optName}
' **   Deduct Print:     N {cmdTaxIncDed_DeductionPrint}
' **   Income Word:      O {cmdTaxIncDed_IncomeWord}
' **   Income Print:     P {cmdTaxIncDed_IncomePrint}
' **   Show All Trans:   S {chkShowTrans}
' **   To:               T {DateEnd}
' **   Specified:        U {opgAccountNumber_optSpecified}
' **   Deduct Preview:   V {cmdTaxIncDed_DeductionPreview}
' **   Income Preview:   W {cmdTaxIncDed_IncomePreview}
' **   Exit:             X {cmdClose}
' **   Deduct Excel:     Z {cmdTaxIncDed_DeductionExcel}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Print All:        A {cmdPrintAll}
' **   Excel All:        E {cmdExcelAll}
' **   Page Count:       G {chkPageOf}
' **   Open Excel:       N {chkOpenExcel}
' **   Open Word:        O {chkOpenWord}
' **   Word All:         W {cmdWordAll}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean
Private blnPrintAll_Focus As Boolean, blnWordAll_Focus As Boolean, blnExcelAll_Focus As Boolean

Private strFile1 As String, strFile2 As String
Private blnPrintBoth As Boolean, blnGoingToReport As Boolean
Private strCallingForm As String, lngTpp As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
160           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False
170           blnPrintAll_Focus = False: blnWordAll_Focus = False: blnExcelAll_Focus = False

180           gstrFormQuerySpec = THIS_NAME
190           blnPrintBoth = False

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
200           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Reset the variables to ensure we get a new table everytime we come into the form.
210           gstrAccountNo = vbNullString
220           gstrAccountName = vbNullString
230           gdatStartDate = 0
240           gdatEndDate = 0

      #If NoExcel Then
250           .cmdTaxIncDed_IncomeExcel.Enabled = False
260           .cmdTaxIncDed_DeductionExcel.Enabled = False
270           .cmdExcelAll.Enabled = False
280           .chkOpenExcel.Enabled = False
290           .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
300           .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
310           .cmdTaxIncDed_IncomeExcel.Enabled = True
320           .cmdTaxIncDed_DeductionExcel.Enabled = True
330           .cmdExcelAll.Enabled = True
340           .chkOpenExcel.Enabled = True
350           .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
360           .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!
              ' ** EVENT CHECK: chkRememberDates!
370           If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
380             DateStart_Exit Cancel  ' ** Procedure: Below.
390             DateEnd_Exit Cancel  ' ** Procedure: Below.
400           End If

410           Set dbs = CurrentDb
420           With dbs
                ' ** Empty tmpTaxReports.
430             Set qdf = .QueryDefs("qryTaxReporting_11")
440             qdf.Execute
450             .Close
460           End With

470           chkOpenWord_AfterUpdate  ' ** Procedure: Below.
480           chkOpenExcel_AfterUpdate  ' ** Procedure: Below.

490         Else
500           Cancel = -1
510         End If
520       Else
530         Cancel = -1
540       End If
550     End With

560     If Cancel = -1 Then
570       Beep
580       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
590     End If

600     DoCmd.Hourglass False

EXITP:
610     Set qdf = Nothing
620     Set dbs = Nothing
630     Exit Sub

ERRH:
640     DoCmd.Hourglass False
650     Select Case ERR.Number
        Case Else
660       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
670     End Select
680     Resume EXITP

End Sub

Private Sub Form_Load()

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset

710     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
720       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
730       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
740       clsMonthClass.PositionAtCursor = True

750       ArchiveSet  ' ** Procedure: Below.

760       Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
770       DoEvents

780       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

790       opgAccountSource_AfterUpdate  ' ** Procedure: Below.
800       chkShowTrans_AfterUpdate  ' ** Procedure: Below.
810       chkRememberMe_AfterUpdate  ' ** Procedure: Below.
820       chkRememberDates_AfterUpdate  ' ** Procedure: Below.
830       chkPageOf_AfterUpdate  ' ** Procedure: Below.
840       opgAccountNumber_AfterUpdate  ' ** Procedure: Below.
850       chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.

860     End With

EXITP:
870     Set rst = Nothing
880     Set qdf = Nothing
890     Set dbs = Nothing
900     Exit Sub

ERRH:
910     Select Case ERR.Number
        Case Else
920       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
930     End Select
940     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1020    intShiftDown = (Shift And acShiftMask) > 0
1030    intAltDown = (Shift And acAltMask) > 0
1040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Dis/Ded Excel:    Z {cmdTaxIncDed_DeductionExcel}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Print All:        A {cmdPrintAll}
        ' **   Excel All:        E {cmdExcelAll}
        ' **   Page Count:       G {chkPageOf}
        ' **   Open Excel:       N {chkOpenExcel}
        ' **   Open Word:        O {chkOpenWord}
        ' **   Word All:         W {cmdWordAll}

        ' ** Plain keys.
1050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1060      Select Case intRetVal
          Case vbKeyEscape
1070        intRetVal = 0
1080        cmdClose_Click  ' ** Procedure: Below.
1090      Case vbKeyF7
1100        intRetVal = 0
1110        cmdCalendar1_Click  ' ** Procedure: Below.
1120      Case vbKeyF8
1130        intRetVal = 0
1140        cmdCalendar2_Click  ' ** Procedure: Below.
1150      End Select
1160    End If

        ' ** Alt keys.
1170    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1180      Select Case intRetVal
          Case vbKeyX
1190        intRetVal = 0
1200        cmdClose_Click  ' ** Procedure: Below.
1210      Case vbKeyZ
1220        With Me
1230          intRetVal = 0
1240          If .cmdTaxIncDed_DeductionExcel.Enabled = True Then
1250            .cmdTaxIncDed_DeductionExcel.SetFocus
1260            cmdTaxIncDed_DeductionExcel_Click  ' ** Procedure: Below
1270          End If
1280        End With
1290      End Select
1300    End If

        ' ** Ctrl keys.
1310    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1320      Select Case intRetVal
          Case vbKeyA
1330        With Me
1340          intRetVal = 0
1350          .cmdPrintAll.SetFocus
1360          cmdPrintAll_Click  ' ** Procedure: Below.
1370        End With
1380      Case vbKeyE
1390        With Me
1400          intRetVal = 0
1410          .cmdExcelAll.SetFocus
1420          cmdExcelAll_Click  ' ** Procedure: Below.
1430        End With
1440      Case vbKeyG
1450        With Me
1460          intRetVal = 0
1470          If .chkPageOf.Enabled = True Then
1480            .chkPageOf.SetFocus
1490            .chkPageOf = (Not .chkPageOf)
1500            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1510          End If
1520        End With
1530      Case vbKeyN
1540        With Me
1550          intRetVal = 0
1560          .chkOpenExcel.SetFocus
1570          .chkOpenExcel = (Not .chkOpenExcel)
1580          chkOpenExcel_AfterUpdate  ' ** Procedure: Below.
1590        End With
1600      Case vbKeyO
1610        With Me
1620          intRetVal = 0
1630          .chkOpenWord.SetFocus
1640          .chkOpenWord = (Not .chkOpenWord)
1650          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
1660        End With
1670      Case vbKeyW
1680        With Me
1690          intRetVal = 0
1700          .cmdWordAll.SetFocus
1710          cmdWordAll_Click  ' ** Procedure: Below.
1720        End With
1730      End Select
1740    End If

        ' ** Ctrl-Shift keys.
1750    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1760      Select Case intRetVal
          Case vbKeyF
1770        With Me
1780          intRetVal = 0
1790          .FocusHolder.SetFocus
1800          .DateStart = Null
1810        End With
1820      End Select
1830    End If

EXITP:
1840    KeyCode = intRetVal
1850    Exit Sub

ERRH:
1860    intRetVal = 0
1870    THAT_PROC = THIS_PROC
1880    That_Erl = Erl
1890    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1900    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2010    With Me
2020      .TimerInterval = 0&
2030      If gblnSetFocus = True Then
2040        gblnSetFocus = False
2050        strFormName = vbNullString
2060  On Error Resume Next
2070        Do Until strFormName = THIS_NAME
2080          DoCmd.SelectObject acForm, THIS_NAME, False
2090          strFormName = Screen.ActiveForm.Name
2100        Loop
2110  On Error GoTo ERRH
2120      End If
2130      If gblnGoToReport = True Then
2140        Select Case blnGoingToReport
            Case True
2150          DoCmd.Hourglass False  ' ** It should already be off.
2160          blnGoingToReport = False
2170          gblnGoToReport = False
2180          .GoToReport_arw_rptinc_img.Visible = False
2190          .GoToReport_arw_rptded_img.Visible = False
2200          .cmdTaxIncDed_IncomeWord.Visible = True
2210          .cmdTaxIncDed_DeductionWord.Visible = True
2220          DoEvents
2230        Case False
2240          DoCmd.Hourglass True  ' ** Make sure it's still running.
2250          DoEvents
2260          blnGoingToReport = True
2270          Select Case garr_varGoToReport(GTR_RNAM)
              Case "rptTax_ReceiptsIncome"
2280            .cmdTaxIncDed_IncomePrint.SetFocus
2290            .cmdTaxIncDed_IncomeWord.Visible = False
2300            .GoToReport_arw_rptinc_img.Visible = True
2310          Case "rptTax_DisbursementsDeductions"
2320            .cmdTaxIncDed_DeductionPrint.SetFocus
2330            .cmdTaxIncDed_DeductionWord.Visible = False
2340            .GoToReport_arw_rptded_img.Visible = True
2350          End Select
2360          DoEvents
2370          Beep
2380          DoCmd.Hourglass False
2390          DoEvents
2400          .TimerInterval = (GTR_WAIT * 2&)
2410        End Select
2420      Else
2430        DoCmd.Hourglass False
2440      End If
2450    End With

EXITP:
2460    Exit Sub

ERRH:
2470    DoCmd.Hourglass False
2480    Select Case ERR.Number
        Case Else
2490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2500    End Select
2510    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2610    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
2620      Response = acDataErrContinue
2630    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2640      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
2650        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
2660        Response = acDataErrContinue
2670      End If
2680    Case 2237  ' ** The text you entered isn't an item in the list.
2690      Response = acDataErrDisplay
2700    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
2710      Beep
2720      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
2730      Response = acDataErrContinue
2740    Case 2501  ' ** The '|' action was Canceled.
2750      Response = acDataErrContinue
2760    Case Else
2770      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2780      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2790    End Select

EXITP:
2800    Exit Sub

ERRH:
2810    Select Case ERR.Number
        Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2910    DoCmd.Hourglass True
2920    DoEvents

2930    With Me
2940      If .chkRememberDates = False Then
2950        .DateStart = Null  ' ** So it won't be there when Pref_Save() runs.
2960        .DateEnd = Null
2970      End If
2980      If .chkRememberMe = False Then
2990        .cmbAccounts = Null
3000      End If
3010    End With

3020    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
3030    Do While Reports.Count > 0
3040      DoCmd.Close acReport, Reports(0).Name
3050    Loop

3060    gstrAccountNo = vbNullString
3070    gstrAccountName = vbNullString
3080    gdatStartDate = 0
3090    gdatEndDate = 0

3100    gstrFormQuerySpec = vbNullString
3110    gblnSetFocus = True
3120    gstrReturningForm = THIS_NAME
3130    Select Case strCallingForm
        Case "frmMenu_Report"
3140      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
3150    End Select

EXITP:
3160    Exit Sub

ERRH:
3170    Select Case ERR.Number
        Case Else
3180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3190    End Select
3200    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

3310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
3320    Exit Sub

ERRH:
3330    Select Case ERR.Number
        Case Else
3340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3350    End Select
3360    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

3410    With Me
3420      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
3430        Select Case blnCalendar1_Focus
            Case True
3440          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
3450          .cmdCalendar1_raised_img.Visible = False
3460        Case False
3470          .cmdCalendar1_raised_img.Visible = True
3480          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
3490        End Select
3500        .cmdCalendar1_raised_focus_dots_img.Visible = False
3510        .cmdCalendar1_raised_focus_img.Visible = False
3520        .cmdCalendar1_sunken_focus_dots_img.Visible = False
3530        .cmdCalendar1_raised_img_dis.Visible = False
3540      End If
3550      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
3560        Select Case blnCalendar2_Focus
            Case True
3570          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
3580          .cmdCalendar2_raised_img.Visible = False
3590        Case False
3600          .cmdCalendar2_raised_img.Visible = True
3610          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
3620        End Select
3630        .cmdCalendar2_raised_focus_dots_img.Visible = False
3640        .cmdCalendar2_raised_focus_img.Visible = False
3650        .cmdCalendar2_sunken_focus_dots_img.Visible = False
3660        .cmdCalendar2_raised_img_dis.Visible = False
3670      End If
3680      If blnPrintAll_Focus = False And (.cmdPrintAll_box01.Visible = True Or .cmdPrintAll_box02.Visible = True) Then
3690        .cmdPrintAll_box01.Visible = False
3700        .cmdPrintAll_box02.Visible = False
3710      End If
3720      If blnWordAll_Focus = False And (.cmdWordAll_box01.Visible = True Or .cmdWordAll_box02.Visible = True) Then
3730        .cmdWordAll_box01.Visible = False
3740        .cmdWordAll_box02.Visible = False
3750      End If
3760      If blnExcelAll_Focus = False And (.cmdExcelAll_box01.Visible = True Or .cmdExcelAll_box02.Visible = True) Then
3770        .cmdExcelAll_box01.Visible = False
3780        .cmdExcelAll_box02.Visible = False
3790      End If
3800    End With

EXITP:
3810    Exit Sub

ERRH:
3820    Select Case ERR.Number
        Case Else
3830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3840    End Select
3850    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub cmdClose_Click()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

4010    DoCmd.Hourglass True
4020    DoEvents

4030    DoCmd.Close acForm, THIS_NAME

EXITP:
4040    Exit Sub

ERRH:
4050    Select Case ERR.Number
        Case Else
4060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4070    End Select
4080    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4120    intShiftDown = (Shift And acShiftMask) > 0
4130    intAltDown = (Shift And acAltMask) > 0
4140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4160      Select Case intRetVal
          Case vbKeyTab
4170        With Me
4180          intRetVal = 0
4190          .DateStart.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab
4250        With Me
4260          intRetVal = 0
4270          If .cmdExcelAll.Enabled = True Then
4280            .cmdExcelAll.SetFocus
4290          Else
4300            .cmdWordAll.SetFocus
4310          End If
4320        End With
4330      End Select
4340    End If

        ' ** Ctrl-Shift keys.
4350    If intCtrlDown And (Not intAltDown) And intShiftDown Then
4360      Select Case intRetVal
          Case vbKeyTab
4370        With Me
4380          intRetVal = 0
4390          If .chkOpenExcel.Enabled = True Then
4400            .chkOpenExcel.SetFocus
4410          Else
4420            .chkOpenWord.SetFocus
4430          End If
4440        End With
4450      End Select
4460    End If

EXITP:
4470    KeyCode = intRetVal
4480    Exit Sub

ERRH:
4490    intRetVal = 0
4500    THAT_PROC = THIS_PROC
4510    That_Erl = Erl
4520    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4530    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

4610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case Else
4640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4650    End Select
4660    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4720    intShiftDown = (Shift And acShiftMask) > 0
4730    intAltDown = (Shift And acAltMask) > 0
4740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4770        With Me
4780          intRetVal = 0
4790          .DateEnd.SetFocus
4800        End With
4810      End Select
4820    End If

        ' ** Shift keys.
4830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4850        With Me
4860          intRetVal = 0
4870          .cmdClose.SetFocus
4880        End With
4890      End Select
4900    End If

        ' ** Ctrl keys.
4910    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4920      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4930        With Me
4940          intRetVal = 0
4950          .cmdCalendar1.SetFocus
4960        End With
4970      End Select
4980    End If

EXITP:
4990    KeyCode = intRetVal
5000    Exit Sub

ERRH:
5010    intRetVal = 0
5020    THAT_PROC = THIS_PROC
5030    That_Erl = Erl
5040    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5050    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

5110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

5210    With Me
5220      If IsDate(.DateEnd) = True And IsDate(.DateStart) = True Then
5230        If CDate(.DateStart) > CDate(.DateEnd) Then
5240          MsgBox "The From date must be less than or equal to the To date.", vbInformation + vbOKOnly, "Invalid Date"
5250          DoCmd.CancelEvent
5260        End If
5270      End If
5280    End With

EXITP:
5290    Exit Sub

ERRH:
5300    THAT_PROC = THIS_PROC
5310    That_Erl = Erl
5320    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5330    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

5410    With Me
5420      blnCalendar1_Focus = True
5430      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
5440      .cmdCalendar1_raised_img.Visible = False
5450      .cmdCalendar1_raised_focus_img.Visible = False
5460      .cmdCalendar1_raised_focus_dots_img.Visible = False
5470      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5480      .cmdCalendar1_raised_img_dis.Visible = False
5490    End With

EXITP:
5500    Exit Sub

ERRH:
5510    Select Case ERR.Number
        Case Else
5520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5530    End Select
5540    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

5610    With Me
5620      blnCalendar1_MouseDown = True
5630      .cmdCalendar1_sunken_focus_dots_img.Visible = True
5640      .cmdCalendar1_raised_img.Visible = False
5650      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5660      .cmdCalendar1_raised_focus_img.Visible = False
5670      .cmdCalendar1_raised_focus_dots_img.Visible = False
5680      .cmdCalendar1_raised_img_dis.Visible = False
5690    End With

EXITP:
5700    Exit Sub

ERRH:
5710    Select Case ERR.Number
        Case Else
5720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5730    End Select
5740    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

5810    With Me
5820      datStartDate = Date
5830      datEndDate = 0
5840      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
5850      If blnRetVal = True Then
5860        .DateStart = datStartDate
5870      Else
5880        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
5890      End If
5900      .DateStart.SetFocus
5910    End With

EXITP:
5920    Exit Sub

ERRH:
5930    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
5940    Case Else
5950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5960    End Select
5970    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6020    intShiftDown = (Shift And acShiftMask) > 0
6030    intAltDown = (Shift And acAltMask) > 0
6040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6070        With Me
6080          intRetVal = 0
6090          .DateEnd.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6150        With Me
6160          intRetVal = 0
6170          .DateStart.SetFocus
6180        End With
6190      End Select
6200    End If

EXITP:
6210    KeyCode = intRetVal
6220    Exit Sub

ERRH:
6230    intRetVal = 0
6240    THAT_PROC = THIS_PROC
6250    That_Erl = Erl
6260    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6270    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

6310    With Me
6320      If blnCalendar1_MouseDown = False Then
6330        Select Case blnCalendar1_Focus
            Case True
6340          .cmdCalendar1_raised_focus_dots_img.Visible = True
6350          .cmdCalendar1_raised_focus_img.Visible = False
6360        Case False
6370          .cmdCalendar1_raised_focus_img.Visible = True
6380          .cmdCalendar1_raised_focus_dots_img.Visible = False
6390        End Select
6400        .cmdCalendar1_raised_img.Visible = False
6410        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6420        .cmdCalendar1_sunken_focus_dots_img.Visible = False
6430        .cmdCalendar1_raised_img_dis.Visible = False
6440      End If
6450    End With

EXITP:
6460    Exit Sub

ERRH:
6470    Select Case ERR.Number
        Case Else
6480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6490    End Select
6500    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

6610    With Me
6620      .cmdCalendar1_raised_focus_dots_img.Visible = True
6630      .cmdCalendar1_raised_img.Visible = False
6640      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6650      .cmdCalendar1_raised_focus_img.Visible = False
6660      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6670      .cmdCalendar1_raised_img_dis.Visible = False
6680      blnCalendar1_MouseDown = False
6690    End With

EXITP:
6700    Exit Sub

ERRH:
6710    Select Case ERR.Number
        Case Else
6720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6730    End Select
6740    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

6810    With Me
6820      .cmdCalendar1_raised_img.Visible = True
6830      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6840      .cmdCalendar1_raised_focus_img.Visible = False
6850      .cmdCalendar1_raised_focus_dots_img.Visible = False
6860      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6870      .cmdCalendar1_raised_img_dis.Visible = False
6880      blnCalendar1_Focus = False
6890    End With

EXITP:
6900    Exit Sub

ERRH:
6910    Select Case ERR.Number
        Case Else
6920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6930    End Select
6940    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7020    intShiftDown = (Shift And acShiftMask) > 0
7030    intAltDown = (Shift And acAltMask) > 0
7040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7070        With Me
7080          intRetVal = 0
7090          .opgAccountNumber.SetFocus
7100        End With
7110      End Select
7120    End If

        ' ** Shift keys.
7130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7150        With Me
7160          intRetVal = 0
7170          .DateStart.SetFocus
7180        End With
7190      End Select
7200    End If

        ' ** Ctrl keys.
7210    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
7220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7230        With Me
7240          intRetVal = 0
7250          .cmdCalendar2.SetFocus
7260        End With
7270      End Select
7280    End If

        ' ** Ctrl-Shift keys.
7290    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7300      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7310        With Me
7320          intRetVal = 0
7330          .cmdCalendar1.SetFocus
7340        End With
7350      End Select
7360    End If

EXITP:
7370    KeyCode = intRetVal
7380    Exit Sub

ERRH:
7390    intRetVal = 0
7400    THAT_PROC = THIS_PROC
7410    That_Erl = Erl
7420    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7430    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

7510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
7520    Exit Sub

ERRH:
7530    Select Case ERR.Number
        Case Else
7540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7550    End Select
7560    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

7610    With Me

7620      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
7630        .DateEnd = Date
7640      End If

7650      If IsNull(.DateStart) Then
7660        MsgBox "The From date must be prior to or the same as the To date.", vbInformation + vbOKOnly, "Entry Required"
7670        .DateStart.SetFocus
7680      Else
7690        If CDate(.DateStart) > CDate(.DateEnd) Then
7700          MsgBox "The From date must be prior to or the same as the To date.", vbInformation + vbOKOnly, "Invalid Date"
7710          .DateStart.SetFocus
7720        End If
7730      End If

7740    End With

EXITP:
7750    Exit Sub

ERRH:
7760    THAT_PROC = THIS_PROC
7770    That_Erl = Erl
7780    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7790    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

7810    With Me
7820      blnCalendar2_Focus = True
7830      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
7840      .cmdCalendar2_raised_img.Visible = False
7850      .cmdCalendar2_raised_focus_img.Visible = False
7860      .cmdCalendar2_raised_focus_dots_img.Visible = False
7870      .cmdCalendar2_sunken_focus_dots_img.Visible = False
7880      .cmdCalendar2_raised_img_dis.Visible = False
7890    End With

EXITP:
7900    Exit Sub

ERRH:
7910    Select Case ERR.Number
        Case Else
7920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7930    End Select
7940    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

8010    With Me
8020      blnCalendar2_MouseDown = True
8030      .cmdCalendar2_sunken_focus_dots_img.Visible = True
8040      .cmdCalendar2_raised_img.Visible = False
8050      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8060      .cmdCalendar2_raised_focus_img.Visible = False
8070      .cmdCalendar2_raised_focus_dots_img.Visible = False
8080      .cmdCalendar2_raised_img_dis.Visible = False
8090    End With

EXITP:
8100    Exit Sub

ERRH:
8110    Select Case ERR.Number
        Case Else
8120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8130    End Select
8140    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

8210    With Me
8220      datStartDate = Date
8230      datEndDate = 0
8240      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
8250      If blnRetVal = True Then
8260        .DateEnd = datStartDate
8270      Else
8280        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
8290      End If
8300      .DateEnd.SetFocus
8310      Cancel = 0
8320      DateEnd_Exit Cancel  ' ** Procedure: Above.
8330      If Cancel = 0 Then
8340        .opgAccountNumber.SetFocus
8350      End If
8360    End With

EXITP:
8370    Exit Sub

ERRH:
8380    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
8390    Case Else
8400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8410    End Select
8420    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8520    intShiftDown = (Shift And acShiftMask) > 0
8530    intAltDown = (Shift And acAltMask) > 0
8540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8570        With Me
8580          intRetVal = 0
8590          .chkRememberDates.SetFocus
8600        End With
8610      End Select
8620    End If

        ' ** Shift keys.
8630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8650        With Me
8660          intRetVal = 0
8670          .DateEnd.SetFocus
8680        End With
8690      End Select
8700    End If

EXITP:
8710    KeyCode = intRetVal
8720    Exit Sub

ERRH:
8730    intRetVal = 0
8740    THAT_PROC = THIS_PROC
8750    That_Erl = Erl
8760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8770    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

8810    With Me
8820      If blnCalendar2_MouseDown = False Then
8830        Select Case blnCalendar2_Focus
            Case True
8840          .cmdCalendar2_raised_focus_dots_img.Visible = True
8850          .cmdCalendar2_raised_focus_img.Visible = False
8860        Case False
8870          .cmdCalendar2_raised_focus_img.Visible = True
8880          .cmdCalendar2_raised_focus_dots_img.Visible = False
8890        End Select
8900        .cmdCalendar2_raised_img.Visible = False
8910        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8920        .cmdCalendar2_sunken_focus_dots_img.Visible = False
8930        .cmdCalendar2_raised_img_dis.Visible = False
8940      End If
8950    End With

EXITP:
8960    Exit Sub

ERRH:
8970    Select Case ERR.Number
        Case Else
8980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8990    End Select
9000    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

9110    With Me
9120      .cmdCalendar2_raised_focus_dots_img.Visible = True
9130      .cmdCalendar2_raised_img.Visible = False
9140      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
9150      .cmdCalendar2_raised_focus_img.Visible = False
9160      .cmdCalendar2_sunken_focus_dots_img.Visible = False
9170      .cmdCalendar2_raised_img_dis.Visible = False
9180      blnCalendar2_MouseDown = False
9190    End With

EXITP:
9200    Exit Sub

ERRH:
9210    Select Case ERR.Number
        Case Else
9220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9230    End Select
9240    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

9310    With Me
9320      .cmdCalendar2_raised_img.Visible = True
9330      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
9340      .cmdCalendar2_raised_focus_img.Visible = False
9350      .cmdCalendar2_raised_focus_dots_img.Visible = False
9360      .cmdCalendar2_sunken_focus_dots_img.Visible = False
9370      .cmdCalendar2_raised_img_dis.Visible = False
9380      blnCalendar2_Focus = False
9390    End With

EXITP:
9400    Exit Sub

ERRH:
9410    Select Case ERR.Number
        Case Else
9420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9430    End Select
9440    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

9510    With Me
9520      Select Case .chkRememberDates
          Case True
9530        .chkRememberDates_lbl.FontBold = True
9540      Case False
9550        .chkRememberDates_lbl.FontBold = False
9560      End Select
9570    End With

EXITP:
9580    Exit Sub

ERRH:
9590    THAT_PROC = THIS_PROC
9600    That_Erl = Erl
9610    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
9620    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

9700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9720    intShiftDown = (Shift And acShiftMask) > 0
9730    intAltDown = (Shift And acAltMask) > 0
9740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9770        With Me
9780          intRetVal = 0
9790          .opgAccountNumber.SetFocus
9800        End With
9810      End Select
9820    End If

        ' ** Shift keys.
9830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9850        With Me
9860          intRetVal = 0
9870          .cmdCalendar2.SetFocus
9880        End With
9890      End Select
9900    End If

EXITP:
9910    KeyCode = intRetVal
9920    Exit Sub

ERRH:
9930    intRetVal = 0
9940    THAT_PROC = THIS_PROC
9950    That_Erl = Erl
9960    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9970    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

10010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
10020   Exit Sub

ERRH:
10030   Select Case ERR.Number
        Case Else
10040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10050   End Select
10060   Resume EXITP

End Sub

Private Sub opgAccountNumber_AfterUpdate()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_AfterUpdate"

10110   With Me
10120     Select Case .opgAccountNumber
          Case .opgAccountNumber_optSpecified.OptionValue
10130       .opgAccountNumber_optSpecified_lbl.FontBold = True
10140       .opgAccountNumber_optAll_lbl.FontBold = False
10150       .opgAccountNumber_optSpecified_lbl_box.Visible = True
10160       .opgAccountNumber_optAll_lbl_box.Visible = False
10170       .cmbAccounts.Enabled = True
10180       .cmbAccounts.BorderColor = CLR_LTBLU2
10190       .cmbAccounts.BackStyle = acBackStyleNormal
10200       .chkRememberMe.Enabled = True
10210       .chkRememberMe_lbl.Visible = True
10220       .chkRememberMe_lbl2_dim.Visible = False
10230       .chkRememberMe_lbl2_dim_hi.Visible = False
10240       .opgAccountSource.Enabled = True
10250       .opgAccountSource_optNumber_lbl2.ForeColor = CLR_VDKGRY
10260       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = False
10270       .opgAccountSource_optName_lbl2.ForeColor = CLR_VDKGRY
10280       .opgAccountSource_optName_lbl2_dim_hi.Visible = False
10290     Case .opgAccountNumber_optAll.OptionValue
10300       .opgAccountNumber_optSpecified_lbl.FontBold = False
10310       .opgAccountNumber_optAll_lbl.FontBold = True
10320       .opgAccountNumber_optSpecified_lbl_box.Visible = False
10330       .opgAccountNumber_optAll_lbl_box.Visible = True
10340       .cmbAccounts.Enabled = False
10350       .cmbAccounts.BorderColor = WIN_CLR_DISR
10360       .cmbAccounts.BackStyle = acBackStyleTransparent
10370       .chkRememberMe.Enabled = False
10380       .chkRememberMe_lbl.Visible = False
10390       .chkRememberMe_lbl2_dim.Visible = True
10400       .chkRememberMe_lbl2_dim_hi.Visible = True
10410       .opgAccountSource.Enabled = False
10420       .opgAccountSource_optNumber_lbl2.ForeColor = WIN_CLR_DISF
10430       .opgAccountSource_optNumber_lbl2_dim_hi.Visible = True
10440       .opgAccountSource_optName_lbl2.ForeColor = WIN_CLR_DISF
10450       .opgAccountSource_optName_lbl2_dim_hi.Visible = True
10460     End Select
10470   End With

EXITP:
10480   Exit Sub

ERRH:
10490   THAT_PROC = THIS_PROC
10500   That_Erl = Erl
10510   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
10520   Resume EXITP

End Sub

Private Sub opgAccountNumber_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optSpecified_KeyDown"

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
10690         If .cmbAccounts.Enabled = True Then
10700           .cmbAccounts.SetFocus
10710         Else
10720           .cmdTaxIncDed_IncomePreview.SetFocus
10730         End If
10740       End With
10750     End Select
10760   End If

        ' ** Shift keys.
10770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10790       With Me
10800         intRetVal = 0
10810         .DateEnd.SetFocus
10820       End With
10830     End Select
10840   End If

        ' ** Ctrl keys.
10850   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10870       With Me
10880         intRetVal = 0
10890         If .chkIncludeArchive.Enabled = True Then
10900           .chkIncludeArchive.SetFocus
10910         Else
10920           .chkShowTrans.SetFocus
10930         End If
10940       End With
10950     End Select
10960   End If

        ' ** Ctrl-Shift keys.
10970   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10990       With Me
11000         intRetVal = 0
11010         .chkRememberDates.SetFocus
11020       End With
11030     End Select
11040   End If

EXITP:
11050   KeyCode = intRetVal
11060   Exit Sub

ERRH:
11070   intRetVal = 0
11080   THAT_PROC = THIS_PROC
11090   That_Erl = Erl
11100   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11110   Resume EXITP

End Sub

Private Sub opgAccountNumber_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

11200 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountNumber_optAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11220   intShiftDown = (Shift And acShiftMask) > 0
11230   intAltDown = (Shift And acAltMask) > 0
11240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11270       With Me
11280         intRetVal = 0
11290         If .cmbAccounts.Enabled = True Then
11300           .cmbAccounts.SetFocus
11310         Else
11320           .cmdTaxIncDed_IncomePreview.SetFocus
11330         End If
11340       End With
11350     End Select
11360   End If

        ' ** Shift keys.
11370   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11380     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11390       With Me
11400         intRetVal = 0
11410         .DateEnd.SetFocus
11420       End With
11430     End Select
11440   End If

        ' ** Ctrl keys.
11450   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
11460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11470       With Me
11480         intRetVal = 0
11490         If .chkIncludeArchive.Enabled = True Then
11500           .chkIncludeArchive.SetFocus
11510         Else
11520           .chkShowTrans.SetFocus
11530         End If
11540       End With
11550     End Select
11560   End If

        ' ** Ctrl-Shift keys.
11570   If intCtrlDown And (Not intAltDown) And intShiftDown Then
11580     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11590       With Me
11600         intRetVal = 0
11610         .chkRememberDates.SetFocus
11620       End With
11630     End Select
11640   End If

EXITP:
11650   KeyCode = intRetVal
11660   Exit Sub

ERRH:
11670   intRetVal = 0
11680   THAT_PROC = THIS_PROC
11690   That_Erl = Erl
11700   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11710   Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

11810   With Me
11820     Select Case .chkIncludeArchive
          Case True
11830       .chkIncludeArchive_lbl.FontBold = True
11840     Case False
11850       .chkIncludeArchive_lbl.FontBold = False
11860     End Select
11870   End With

EXITP:
11880   Exit Sub

ERRH:
11890   Select Case ERR.Number
        Case Else
11900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11910   End Select
11920   Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
12010   intShiftDown = (Shift And acShiftMask) > 0
12020   intAltDown = (Shift And acAltMask) > 0
12030   intCtrlDown = (Shift And acCtrlMask) > 0

12040   intRetVal = KeyCode

        ' ** Plain keys.
12050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12070       With Me
12080         intRetVal = 0
12090         .chkShowTrans.SetFocus
12100       End With
12110     End Select
12120   End If

        ' ** Shift keys.
12130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12150       With Me
12160         intRetVal = 0
12170         .opgAccountNumber.SetFocus
12180       End With
12190     End Select
12200   End If

EXITP:
12210   KeyCode = intRetVal
12220   Exit Sub

ERRH:
12230   intRetVal = 0
12240   THAT_PROC = THIS_PROC
12250   That_Erl = Erl
12260   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12270   Resume EXITP

End Sub

Private Sub chkShowTrans_AfterUpdate()

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowTrans_AfterUpdate"

12310   With Me
12320     Select Case .chkShowTrans
          Case True
12330       .chkShowTrans_lbl.FontBold = True
12340     Case False
12350       .chkShowTrans_lbl.FontBold = False
12360     End Select
12370   End With

EXITP:
12380   Exit Sub

ERRH:
12390   THAT_PROC = THIS_PROC
12400   That_Erl = Erl
12410   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
12420   Resume EXITP

End Sub

Private Sub chkShowTrans_KeyDown(KeyCode As Integer, Shift As Integer)

12500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkShowTrans_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12520   intShiftDown = (Shift And acShiftMask) > 0
12530   intAltDown = (Shift And acAltMask) > 0
12540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12570       With Me
12580         intRetVal = 0
12590         .opgAccountSource.SetFocus
12600       End With
12610     End Select
12620   End If

        ' ** Shift keys.
12630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12650       With Me
12660         intRetVal = 0
12670         If .chkIncludeArchive.Enabled = True Then
12680           .chkIncludeArchive.SetFocus
12690         Else
12700           .opgAccountNumber.SetFocus
12710         End If
12720       End With
12730     End Select
12740   End If

EXITP:
12750   KeyCode = intRetVal
12760   Exit Sub

ERRH:
12770   intRetVal = 0
12780   THAT_PROC = THIS_PROC
12790   That_Erl = Erl
12800   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12810   Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

12900 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

12910   strAccountNo = vbNullString

12920   With Me
12930     If IsNull(.cmbAccounts) = False Then
12940       If Len(.cmbAccounts.Column(0)) > 0 Then
12950         strAccountNo = .cmbAccounts.Column(0)
12960       End If
12970     End If
12980     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
12990       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
13000       .opgAccountSource_optNumber_lbl2.FontBold = True
13010       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = True
13020       .opgAccountSource_optName_lbl2.FontBold = False
13030       .opgAccountSource_optName_lbl2_dim_hi.FontBold = False
13040     Case .opgAccountSource_optName.OptionValue
13050       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
13060       .opgAccountSource_optNumber_lbl2.FontBold = False
13070       .opgAccountSource_optNumber_lbl2_dim_hi.FontBold = False
13080       .opgAccountSource_optName_lbl2.FontBold = True
13090       .opgAccountSource_optName_lbl2_dim_hi.FontBold = True
13100     End Select
13110     DoEvents
13120     If strAccountNo <> vbNullString Then
13130       .cmbAccounts = strAccountNo
13140     End If
13150   End With

EXITP:
13160   Exit Sub

ERRH:
13170   THAT_PROC = THIS_PROC
13180   That_Erl = Erl
13190   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
13200   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

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
13390         .chkRememberMe.SetFocus
13400       End With
13410     End Select
13420   End If

        ' ** Shift keys.
13430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13450       With Me
13460         intRetVal = 0
13470         .chkShowTrans.SetFocus
13480       End With
13490     End Select
13500   End If

EXITP:
13510   KeyCode = intRetVal
13520   Exit Sub

ERRH:
13530   intRetVal = 0
13540   THAT_PROC = THIS_PROC
13550   That_Erl = Erl
13560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13570   Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

13600 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

13610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
13620   intShiftDown = (Shift And acShiftMask) > 0
13630   intAltDown = (Shift And acAltMask) > 0
13640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
13650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
13660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13670       With Me
13680         intRetVal = 0
13690         .chkRememberMe.SetFocus
13700       End With
13710     End Select
13720   End If

        ' ** Shift keys.
13730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13750       With Me
13760         intRetVal = 0
13770         .chkShowTrans.SetFocus
13780       End With
13790     End Select
13800   End If

EXITP:
13810   KeyCode = intRetVal
13820   Exit Sub

ERRH:
13830   intRetVal = 0
13840   THAT_PROC = THIS_PROC
13850   That_Erl = Erl
13860   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13870   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

13910   With Me
13920     Select Case .chkRememberMe
          Case True
13930       .chkRememberMe_lbl.FontBold = True
13940       .chkRememberMe_lbl2_dim.FontBold = True
13950       .chkRememberMe_lbl2_dim_hi.FontBold = True
13960     Case False
13970       .chkRememberMe_lbl.FontBold = False
13980       .chkRememberMe_lbl2_dim.FontBold = False
13990       .chkRememberMe_lbl2_dim_hi.FontBold = False
14000     End Select
14010   End With

EXITP:
14020   Exit Sub

ERRH:
14030   THAT_PROC = THIS_PROC
14040   That_Erl = Erl
14050   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
14060   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14120   intShiftDown = (Shift And acShiftMask) > 0
14130   intAltDown = (Shift And acAltMask) > 0
14140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14160     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14170       With Me
14180         intRetVal = 0
14190         .cmbAccounts.SetFocus
14200       End With
14210     End Select
14220   End If

        ' ** Shift keys.
14230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14250       With Me
14260         intRetVal = 0
14270         .opgAccountSource.SetFocus
14280       End With
14290     End Select
14300   End If

EXITP:
14310   KeyCode = intRetVal
14320   Exit Sub

ERRH:
14330   intRetVal = 0
14340   THAT_PROC = THIS_PROC
14350   That_Erl = Erl
14360   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14370   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

14400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

14410   With Me
14420     If .cmbAccounts.text = vbNullString Then
14430       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
14440       .cmbAccounts = Me.cmbAccounts.ItemData(0)
14450     End If
14460   End With

EXITP:
14470   Exit Sub

ERRH:
14480   THAT_PROC = THIS_PROC
14490   That_Erl = Erl
14500   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
14510   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

14600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

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
14690         .cmdTaxIncDed_IncomePreview.SetFocus
14700       End With
14710     End Select
14720   End If

        ' ** Shift keys.
14730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14750       With Me
14760         intRetVal = 0
14770         .opgAccountNumber.SetFocus
14780       End With
14790     End Select
14800   End If

        ' ** Ctrl keys.
14810   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
14820     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14830       With Me
14840         intRetVal = 0
14850         .chkPageOf.SetFocus
14860       End With
14870     End Select
14880   End If

        ' ** Ctrl-Shift keys.
14890   If intCtrlDown And (Not intAltDown) And intShiftDown Then
14900     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14910       With Me
14920         intRetVal = 0
14930         .chkRememberMe.SetFocus
14940       End With
14950     End Select
14960   End If

EXITP:
14970   KeyCode = intRetVal
14980   Exit Sub

ERRH:
14990   intRetVal = 0
15000   THAT_PROC = THIS_PROC
15010   That_Erl = Erl
15020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15030   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

15100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

15110   With Me
15120     Select Case .chkPageOf
          Case True
15130       .chkPageOf_lbl.FontBold = True
15140     Case False
15150       .chkPageOf_lbl.FontBold = False
15160     End Select
15170   End With

EXITP:
15180   Exit Sub

ERRH:
15190   DoCmd.Hourglass False
15200   Select Case ERR.Number
        Case Else
15210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15220   End Select
15230   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

15300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
15310   intShiftDown = (Shift And acShiftMask) > 0
15320   intAltDown = (Shift And acAltMask) > 0
15330   intCtrlDown = (Shift And acCtrlMask) > 0

15340   intRetVal = KeyCode

        ' ** Plain keys.
15350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15370       With Me
15380         intRetVal = 0
15390         .cmdTaxIncDed_IncomePreview.SetFocus
15400       End With
15410     End Select
15420   End If

        ' ** Shift keys.
15430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15450       With Me
15460         intRetVal = 0
15470         .cmbAccounts.SetFocus
15480       End With
15490     End Select
15500   End If

EXITP:
15510   KeyCode = intRetVal
15520   Exit Sub

ERRH:
15530   intRetVal = 0
15540   THAT_PROC = THIS_PROC
15550   That_Erl = Erl
15560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15570   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomePreview_Click()

15600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomePreview_Click"

        Dim strDocName As String

15610   If DoReport1 = True Then  ' ** Function: Below.
15620     strDocName = "rptTax_ReceiptsIncome"
15630     DoCmd.OpenReport strDocName, acViewPreview
15640     DoCmd.Maximize
15650     DoCmd.RunCommand acCmdFitToWindow
15660   End If

EXITP:
15670   Exit Sub

ERRH:
15680   Select Case ERR.Number
        Case Else
15690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15700   End Select
15710   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomePreview_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomePreview_KeyDown"

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
          Case vbKeyTab
15870       With Me
15880         intRetVal = 0
15890         .cmdTaxIncDed_IncomePrint.SetFocus
15900       End With
15910     Case vbKeyUp
15920       With Me
15930         intRetVal = 0
15940         If .cmbAccounts.Enabled = True Then
15950           .cmbAccounts.SetFocus
15960         Else
15970           .opgAccountNumber.SetFocus
15980         End If
15990       End With
16000     Case vbKeyDown
16010       With Me
16020         intRetVal = 0
16030         .cmdTaxIncDed_DeductionPreview.SetFocus
16040       End With
16050     Case vbKeyLeft
16060       With Me
16070         intRetVal = 0
16080         If .cmdExcelAll.Enabled = True Then
16090           .cmdExcelAll.SetFocus
16100         Else
16110           .cmdWordAll.SetFocus
16120         End If
16130       End With
16140     Case vbKeyRight
16150       With Me
16160         intRetVal = 0
16170         .cmdTaxIncDed_IncomePrint.SetFocus
16180       End With
16190     End Select
16200   End If

        ' ** Shift keys.
16210   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16220     Select Case intRetVal
          Case vbKeyTab
16230       With Me
16240         intRetVal = 0
16250         .cmbAccounts.SetFocus
16260       End With
16270     End Select
16280   End If

        ' ** Ctrl-Shift keys.
16290   If intCtrlDown And (Not intAltDown) And intShiftDown Then
16300     Select Case intRetVal
          Case vbKeyTab
16310       With Me
16320         intRetVal = 0
16330         .chkPageOf.SetFocus
16340       End With
16350     End Select
16360   End If

EXITP:
16370   KeyCode = intRetVal
16380   Exit Sub

ERRH:
16390   intRetVal = 0
16400   THAT_PROC = THIS_PROC
16410   That_Erl = Erl
16420   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16430   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomePrint_Click()

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomePrint_Click"

        Dim strDocName As String

16510   If DoReport1 = True Then  ' ** Function: Below.
16520     strDocName = "rptTax_ReceiptsIncome"
16530     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
16540       DoCmd.OpenReport strDocName, acViewPreview
16550       DoCmd.Maximize
16560       DoCmd.RunCommand acCmdFitToWindow
16570     Else
            '##GTR_Ref: rptTax_ReceiptsIncome
16580       DoCmd.OpenReport strDocName, acViewNormal
16590     End If
16600   Else
16610     blnPrintBoth = False
16620   End If

EXITP:
16630   Exit Sub

ERRH:
16640   blnPrintBoth = False
16650   Select Case ERR.Number
        Case Else
16660     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16670   End Select
16680   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomePrint_KeyDown(KeyCode As Integer, Shift As Integer)

16700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomePrint_KeyDown"

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
16790         .cmdTaxIncDed_IncomeWord.SetFocus
16800       End With
16810     Case vbKeyUp
16820       With Me
16830         intRetVal = 0
16840         .cmdPrintAll.SetFocus
16850       End With
16860     Case vbKeyDown
16870       With Me
16880         intRetVal = 0
16890         .cmdTaxIncDed_DeductionPrint.SetFocus
16900       End With
16910     Case vbKeyLeft
16920       With Me
16930         intRetVal = 0
16940         .cmdTaxIncDed_IncomePreview.SetFocus
16950       End With
16960     Case vbKeyRight
16970       With Me
16980         intRetVal = 0
16990         .cmdTaxIncDed_IncomeWord.SetFocus
17000       End With
17010     End Select
17020   End If

        ' ** Shift keys.
17030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17040     Select Case intRetVal
          Case vbKeyTab
17050       With Me
17060         intRetVal = 0
17070         .cmdTaxIncDed_IncomePreview.SetFocus
17080       End With
17090     End Select
17100   End If

EXITP:
17110   KeyCode = intRetVal
17120   Exit Sub

ERRH:
17130   intRetVal = 0
17140   THAT_PROC = THIS_PROC
17150   That_Erl = Erl
17160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
17170   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomeWord_Click()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomeWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

17210   With Me
17220     If DoReport1 = True Then  ' ** Function: Below.

17230       If IsNull(.UserReportPath) = True Then
17240         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
17250       Else
17260         strRptPath = .UserReportPath
17270       End If
17280       strRptCap = "rptTax_ReceiptsIncome_" & Format(Date, "yyyymmdd")

17290       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

17300       If strRptPathFile <> vbNullString Then
17310         strRpt = "rptTax_ReceiptsIncome"
17320         strFile1 = strRptPathFile
17330         Select Case .chkOpenWord
              Case True
17340           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
17350         Case False
17360           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
17370         End Select
17380         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
17390       End If

17400     End If
17410   End With

EXITP:
17420   Exit Sub

ERRH:
17430   Select Case ERR.Number
        Case Else
17440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17450   End Select
17460   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomeWord_KeyDown(KeyCode As Integer, Shift As Integer)

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomeWord_KeyDown"

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
17590         If .cmdTaxIncDed_IncomeExcel.Enabled = True Then
17600           .cmdTaxIncDed_IncomeExcel.SetFocus
17610         Else
17620           .cmdTaxIncDed_DeductionPreview.SetFocus
17630         End If
17640       End With
17650     Case vbKeyUp
17660       With Me
17670         intRetVal = 0
17680         .cmdTaxIncDed_DeductionWord.SetFocus
17690       End With
17700     Case vbKeyDown
17710       With Me
17720         intRetVal = 0
17730         .cmdTaxIncDed_DeductionWord.SetFocus
17740       End With
17750     Case vbKeyLeft
17760       With Me
17770         intRetVal = 0
17780         .cmdTaxIncDed_IncomePrint.SetFocus
17790       End With
17800     Case vbKeyRight
17810       With Me
17820         intRetVal = 0
17830         If .cmdTaxIncDed_IncomeExcel.Enabled = True Then
17840           .cmdTaxIncDed_IncomeExcel.SetFocus
17850         Else
17860           .cmdTaxIncDed_DeductionPreview.SetFocus
17870         End If
17880       End With
17890     End Select
17900   End If

        ' ** Shift keys.
17910   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17920     Select Case intRetVal
          Case vbKeyTab
17930       With Me
17940         intRetVal = 0
17950         .cmdTaxIncDed_IncomePrint.SetFocus
17960       End With
17970     End Select
17980   End If

EXITP:
17990   KeyCode = intRetVal
18000   Exit Sub

ERRH:
18010   intRetVal = 0
18020   THAT_PROC = THIS_PROC
18030   That_Erl = Erl
18040   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18050   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomeExcel_Click()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomeExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

18110   With Me

18120     DoCmd.Hourglass True
18130     DoEvents

18140     blnContinue = True

18150     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18160       DoCmd.Hourglass False
18170       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
18180       If msgResponse <> vbRetry Then
18190         blnContinue = False
18200       End If
18210     End If

18220     If blnContinue = True Then

18230       DoCmd.Hourglass True
18240       DoEvents

18250       If DoReport1 = True Then  ' ** Function: Below.

18260         blnContinue = True

              ' ** Prepare qryTaxReporting for export, via tmpReceiptsIncome table.
18270         If Rpt_Tax_ReceiptIncome = True Then  ' ** Module Function: modReportFunctions.

18280           If IsNull(.UserReportPath) = True Then
18290             strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
18300           Else
18310             strRptPath = .UserReportPath
18320           End If

18330           strRptCap = "rptTax_ReceiptsIncome_" & Format(Date, "yyyymmdd")

                ' ** Income/Receipts: For Excel Export.
18340           strQry = "qryTaxReporting_44"  ' ** Archive handled by DoReport1().

18350           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

18360           If strRptPathFile <> vbNullString Then
18370             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18380               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
18390             End If
18400             DoEvents
18410             If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
18420               Kill strRptPathFile
18430             End If
18440             DoEvents
18450             gstrFormQuerySpec = THIS_NAME  ' ** May not be used.
18460             Select Case gblnPrintAll
                  Case True
18470               DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
18480               strFile1 = strRptPathFile
18490             Case False
18500               DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
18510               If Excel_NameOnly(strRptPathFile, "Income") = True Then  ' ** Module Function: modExcelFuncs.
18520                 Select Case .chkOpenExcel
                      Case True
18530                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
18540                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
18550                   End If
18560                   OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
18570                 Case False
                        ' ** Nothing, that's it.
18580                 End Select
18590               End If
18600             End Select
18610             .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
18620           End If  ' ** strRptPathFile.

18630         End If  ' ** Rpt_Tax_ReceiptIncome.

18640       End If  ' ** DoReport1.
18650     End If  ' ** blnContinue.

18660     DoCmd.Hourglass False

18670   End With

EXITP:
18680   Exit Sub

ERRH:
18690   DoCmd.Hourglass False
18700   Select Case ERR.Number
        Case Else
18710     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18720   End Select
18730   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_IncomeExcel_KeyDown(KeyCode As Integer, Shift As Integer)

18800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_IncomeExcel_KeyDown"

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
18890         .cmdTaxIncDed_DeductionPreview.SetFocus
18900       End With
18910     Case vbKeyUp
18920       With Me
18930         intRetVal = 0
18940         .cmdTaxIncDed_DeductionExcel.SetFocus
18950       End With
18960     Case vbKeyDown
18970       With Me
18980         intRetVal = 0
18990         .cmdTaxIncDed_DeductionExcel.SetFocus
19000       End With
19010     Case vbKeyLeft
19020       With Me
19030         intRetVal = 0
19040         .cmdTaxIncDed_IncomeWord.SetFocus
19050       End With
19060     Case vbKeyRight
19070       With Me
19080         intRetVal = 0
19090         .cmdTaxIncDed_DeductionPreview.SetFocus
19100       End With
19110     End Select
19120   End If

        ' ** Shift keys.
19130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19140     Select Case intRetVal
          Case vbKeyTab
19150       With Me
19160         intRetVal = 0
19170         .cmdTaxIncDed_IncomeWord.SetFocus
19180       End With
19190     End Select
19200   End If

EXITP:
19210   KeyCode = intRetVal
19220   Exit Sub

ERRH:
19230   intRetVal = 0
19240   THAT_PROC = THIS_PROC
19250   That_Erl = Erl
19260   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19270   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionPreview_Click()

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionPreview_Click"

        Dim strDocName As String

19310   If DoReport2 = True Then  ' ** Function: Below.
19320     strDocName = "rptTax_DisbursementsDeductions"
19330     DoCmd.OpenReport strDocName, acViewPreview
19340     DoCmd.Maximize
19350     DoCmd.RunCommand acCmdFitToWindow
19360   End If

EXITP:
19370   Exit Sub

ERRH:
19380   Select Case ERR.Number
        Case Else
19390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19400   End Select
19410   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionPreview_KeyDown(KeyCode As Integer, Shift As Integer)

19500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

19510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
19520   intShiftDown = (Shift And acShiftMask) > 0
19530   intAltDown = (Shift And acAltMask) > 0
19540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
19550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
19560     Select Case intRetVal
          Case vbKeyTab
19570       With Me
19580         intRetVal = 0
19590         .cmdTaxIncDed_DeductionPrint.SetFocus
19600       End With
19610     Case vbKeyUp
19620       With Me
19630         intRetVal = 0
19640         .cmdTaxIncDed_IncomePreview.SetFocus
19650       End With
19660     Case vbKeyDown
19670       With Me
19680         intRetVal = 0
19690         .cmdTaxIncDed_IncomePreview.SetFocus
19700       End With
19710     Case vbKeyLeft
19720       With Me
19730         intRetVal = 0
19740         If .cmdTaxIncDed_IncomeExcel.Enabled = True Then
19750           .cmdTaxIncDed_IncomeExcel.SetFocus
19760         Else
19770           .cmdTaxIncDed_IncomeWord.SetFocus
19780         End If
19790       End With
19800     Case vbKeyRight
19810       With Me
19820         intRetVal = 0
19830         .cmdTaxIncDed_DeductionPrint.SetFocus
19840       End With
19850     End Select
19860   End If

        ' ** Shift keys.
19870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19880     Select Case intRetVal
          Case vbKeyTab
19890       With Me
19900         intRetVal = 0
19910         If .cmdTaxIncDed_IncomeExcel.Enabled = True Then
19920           .cmdTaxIncDed_IncomeExcel.SetFocus
19930         Else
19940           .cmdTaxIncDed_IncomeWord.SetFocus
19950         End If
19960       End With
19970     End Select
19980   End If

EXITP:
19990   KeyCode = intRetVal
20000   Exit Sub

ERRH:
20010   intRetVal = 0
20020   THAT_PROC = THIS_PROC
20030   That_Erl = Erl
20040   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20050   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionPrint_Click()

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionPrint_Click"

        Dim strDocName As String

20110   If DoReport2 = True Then  ' ** Function: Below.
20120     strDocName = "rptTax_DisbursementsDeductions"
20130     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
20140       DoCmd.OpenReport strDocName, acViewPreview
20150       DoCmd.Maximize
20160       DoCmd.RunCommand acCmdFitToWindow
20170     Else
            '##GTR_Ref: rptTax_DisbursementsDeductions
20180       DoCmd.OpenReport strDocName, acViewNormal
20190     End If
20200   Else
20210     blnPrintBoth = False
20220   End If

EXITP:
20230   Exit Sub

ERRH:
20240   blnPrintBoth = False
20250   Select Case ERR.Number
        Case Else
20260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20270   End Select
20280   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionPrint_KeyDown(KeyCode As Integer, Shift As Integer)

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20320   intShiftDown = (Shift And acShiftMask) > 0
20330   intAltDown = (Shift And acAltMask) > 0
20340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20360     Select Case intRetVal
          Case vbKeyTab
20370       With Me
20380         intRetVal = 0
20390         .cmdTaxIncDed_DeductionWord.SetFocus
20400       End With
20410     Case vbKeyUp
20420       With Me
20430         intRetVal = 0
20440         .cmdTaxIncDed_IncomePrint.SetFocus
20450       End With
20460     Case vbKeyDown
20470       With Me
20480         intRetVal = 0
20490         .cmdPrintAll.SetFocus
20500       End With
20510     Case vbKeyLeft
20520       With Me
20530         intRetVal = 0
20540         .cmdTaxIncDed_DeductionPreview.SetFocus
20550       End With
20560     Case vbKeyRight
20570       With Me
20580         intRetVal = 0
20590         .cmdTaxIncDed_DeductionWord.SetFocus
20600       End With
20610     End Select
20620   End If

        ' ** Shift keys.
20630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20640     Select Case intRetVal
          Case vbKeyTab
20650       With Me
20660         intRetVal = 0
20670         .cmdTaxIncDed_DeductionPreview.SetFocus
20680       End With
20690     End Select
20700   End If

EXITP:
20710   KeyCode = intRetVal
20720   Exit Sub

ERRH:
20730   intRetVal = 0
20740   THAT_PROC = THIS_PROC
20750   That_Erl = Erl
20760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20770   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionWord_Click()

20800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

20810   With Me
20820     If DoReport2 = True Then  ' ** Function: Below.

20830       If IsNull(.UserReportPath) = True Then
20840         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
20850       Else
20860         strRptPath = .UserReportPath
20870       End If
20880       strRptCap = "rptTax_DisbursementsDeductions_" & Format(Date, "yyyymmdd")

20890       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

20900       If strRptPathFile <> vbNullString Then
20910         strRpt = "rptTax_DisbursementsDeductions"
20920         strFile2 = strRptPathFile
20930         Select Case .chkOpenWord
              Case True
20940           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
20950         Case False
20960           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
20970         End Select
20980         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
20990       End If

21000     End If
21010   End With

EXITP:
21020   Exit Sub

ERRH:
21030   Select Case ERR.Number
        Case Else
21040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21050   End Select
21060   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionWord_KeyDown(KeyCode As Integer, Shift As Integer)

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

21110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
21120   intShiftDown = (Shift And acShiftMask) > 0
21130   intAltDown = (Shift And acAltMask) > 0
21140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
21150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
21160     Select Case intRetVal
          Case vbKeyTab
21170       With Me
21180         intRetVal = 0
21190         If .cmdTaxIncDed_DeductionExcel.Enabled = True Then
21200           .cmdTaxIncDed_DeductionExcel.SetFocus
21210         Else
21220           .cmdPrintAll.SetFocus
21230         End If
21240       End With
21250     Case vbKeyUp
21260       With Me
21270         intRetVal = 0
21280         .cmdTaxIncDed_IncomeWord.SetFocus
21290       End With
21300     Case vbKeyDown
21310       With Me
21320         intRetVal = 0
21330         .cmdTaxIncDed_IncomeWord.SetFocus
21340       End With
21350     Case vbKeyLeft
21360       With Me
21370         intRetVal = 0
21380         .cmdTaxIncDed_DeductionPrint.SetFocus
21390       End With
21400     Case vbKeyRight
21410       With Me
21420         intRetVal = 0
21430         If .cmdTaxIncDed_DeductionExcel.Enabled = True Then
21440           .cmdTaxIncDed_DeductionExcel.SetFocus
21450         Else
21460           .cmdPrintAll.SetFocus
21470         End If
21480       End With
21490     End Select
21500   End If

        ' ** Shift keys.
21510   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21520     Select Case intRetVal
          Case vbKeyTab
21530       With Me
21540         intRetVal = 0
21550         .cmdTaxIncDed_DeductionPrint.SetFocus
21560       End With
21570     End Select
21580   End If

EXITP:
21590   KeyCode = intRetVal
21600   Exit Sub

ERRH:
21610   intRetVal = 0
21620   THAT_PROC = THIS_PROC
21630   That_Erl = Erl
21640   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21650   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionExcel_Click()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

21710   With Me

21720     DoCmd.Hourglass True
21730     DoEvents

21740     blnContinue = True

21750     If EXE_IsRunning("EXCEL.EXE") = True And gblnPrintAll = False Then  ' ** Module Function: modProcessFuncs.
21760       DoCmd.Hourglass False
21770       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
21780       If msgResponse <> vbRetry Then
21790         blnContinue = False
21800       End If
21810     End If

21820     If blnContinue = True Then

21830       DoCmd.Hourglass True
21840       DoEvents

21850       If DoReport2 = True Then  ' ** Function: Below.

21860         blnContinue = True

              ' ** Prepare qryTaxReporting for export, via tmpTaxDisbursementsDeductions table.
21870         If Rpt_Tax_DisbursementDeduction = True Then  ' ** Module Function: modReportFunctions.

21880           If IsNull(.UserReportPath) = True Then
21890             strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
21900           Else
21910             strRptPath = .UserReportPath
21920           End If

21930           strRptCap = "rptTax_DisbursementsDeductions_" & Format(Date, "yyyymmdd")

                ' ** Deductions/Disbursements: For Excel Export.
21940           strQry = "qryTaxReporting_49"  ' ** Archive handled by DoReport2().

21950           Select Case gblnPrintAll
                Case True
21960             If strFile1 <> vbNullString Then
21970               strRptPathFile = StringReplace(strFile1, "ReceiptsIncome_", "DisbursementsDeductions_")  ' ** Module Function: modStringFuncs.
21980             Else
21990               strRptPathFile = strRptPath & LNK_SEP & strRptCap & ".xls"
22000             End If
22010           Case False
22020             strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
22030           End Select

22040           If strRptPathFile <> vbNullString Then
22050             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22060               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22070             End If
22080             DoEvents
22090             If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
22100               Kill strRptPathFile
22110             End If
22120             DoEvents
22130             gstrFormQuerySpec = THIS_NAME  ' ** May not be used.
22140             Select Case gblnPrintAll
                  Case True
22150               DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
22160               strFile2 = strRptPathFile
22170             Case False
22180               DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
22190               If Excel_NameOnly(strRptPathFile, "Deduction") = True Then  ' ** Module Function: modExcelFuncs.
22200                 Select Case .chkOpenExcel
                      Case True
22210                   If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22220                     EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22230                   End If
22240                   OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
22250                 Case False
                        ' ** Nothing, that's it.
22260                 End Select
22270               End If
22280             End Select
22290             .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
22300           End If  ' ** strRptPathFile.

22310         End If  ' ** Rpt_Tax_DisbursementDeduction.

22320       End If  ' ** DoReport2.
22330     End If  ' ** blnContinue.

22340     DoCmd.Hourglass False

22350   End With

EXITP:
22360   Exit Sub

ERRH:
22370   DoCmd.Hourglass False
22380   Select Case ERR.Number
        Case Else
22390     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22400   End Select
22410   Resume EXITP

End Sub

Private Sub cmdTaxIncDed_DeductionExcel_KeyDown(KeyCode As Integer, Shift As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdTaxIncDed_DeductionExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22520   intShiftDown = (Shift And acShiftMask) > 0
22530   intAltDown = (Shift And acAltMask) > 0
22540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22560     Select Case intRetVal
          Case vbKeyTab
22570       With Me
22580         intRetVal = 0
22590         .cmdPrintAll.SetFocus
22600       End With
22610     Case vbKeyUp
22620       With Me
22630         intRetVal = 0
22640         .cmdTaxIncDed_IncomeExcel.SetFocus
22650       End With
22660     Case vbKeyDown
22670       With Me
22680         intRetVal = 0
22690         .cmdTaxIncDed_IncomeExcel.SetFocus
22700       End With
22710     Case vbKeyLeft
22720       With Me
22730         intRetVal = 0
22740         .cmdTaxIncDed_DeductionWord.SetFocus
22750       End With
22760     Case vbKeyRight
22770       With Me
22780         intRetVal = 0
22790         .cmdPrintAll.SetFocus
22800       End With
22810     End Select
22820   End If

        ' ** Shift keys.
22830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22840     Select Case intRetVal
          Case vbKeyTab
22850       With Me
22860         intRetVal = 0
22870         .cmdTaxIncDed_DeductionWord.SetFocus
22880       End With
22890     End Select
22900   End If

EXITP:
22910   KeyCode = intRetVal
22920   Exit Sub

ERRH:
22930   intRetVal = 0
22940   THAT_PROC = THIS_PROC
22950   That_Erl = Erl
22960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
22970   Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

23010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
23020   Exit Sub

ERRH:
23030   Select Case ERR.Number
        Case Else
23040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23050   End Select
23060   Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

23100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

23110   With Me
23120     blnPrintAll_Focus = True
23130     .cmdPrintAll_box01.Visible = True
23140     .cmdPrintAll_box02.Visible = True
23150   End With

EXITP:
23160   Exit Sub

ERRH:
23170   Select Case ERR.Number
        Case Else
23180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23190   End Select
23200   Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

23310   blnPrintBoth = True
23320   cmdTaxIncDed_IncomePrint_Click  ' ** Procedure: Above.
23330   DoEvents
23340   If blnPrintBoth = True Then
23350     cmdTaxIncDed_DeductionPrint_Click  ' ** Procedure: Above.
23360   End If
23370   blnPrintBoth = False

EXITP:
23380   Exit Sub

ERRH:
23390   blnPrintBoth = False
23400   Select Case ERR.Number
        Case Else
23410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23420   End Select
23430   Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

23500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23520   intShiftDown = (Shift And acShiftMask) > 0
23530   intAltDown = (Shift And acAltMask) > 0
23540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23560     Select Case intRetVal
          Case vbKeyTab
23570       With Me
23580         intRetVal = 0
23590         .cmdClose.SetFocus
23600       End With
23610     Case vbKeyUp
23620       With Me
23630         intRetVal = 0
23640         .cmdTaxIncDed_DeductionPrint.SetFocus
23650       End With
23660     Case vbKeyDown
23670       With Me
23680         intRetVal = 0
23690         .cmdTaxIncDed_IncomePrint.SetFocus
23700       End With
23710     Case vbKeyLeft
23720       With Me
23730         intRetVal = 0
23740         If .cmdTaxIncDed_DeductionExcel.Enabled = True Then
23750           .cmdTaxIncDed_DeductionExcel.SetFocus
23760         Else
23770           .cmdTaxIncDed_DeductionWord.SetFocus
23780         End If
23790       End With
23800     Case vbKeyRight
23810       With Me
23820         intRetVal = 0
23830         .cmdWordAll.SetFocus
23840       End With
23850     End Select
23860   End If

        ' ** Shift keys.
23870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23880     Select Case intRetVal
          Case vbKeyTab
23890       With Me
23900         intRetVal = 0
23910         If .cmdTaxIncDed_DeductionExcel.Enabled = True Then
23920           .cmdTaxIncDed_DeductionExcel.SetFocus
23930         Else
23940           .cmdTaxIncDed_DeductionWord.SetFocus
23950         End If
23960       End With
23970     End Select
23980   End If

EXITP:
23990   KeyCode = intRetVal
24000   Exit Sub

ERRH:
24010   intRetVal = 0
24020   THAT_PROC = THIS_PROC
24030   That_Erl = Erl
24040   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24050   Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

24110   With Me
24120     If gblnPrintAll = False Then
24130       .cmdPrintAll_box01.Visible = True
24140       .cmdPrintAll_box02.Visible = True
24150       If blnWordAll_Focus = False Then
24160         .cmdWordAll_box01.Visible = False
24170         .cmdWordAll_box02.Visible = False
24180       End If
24190       If blnExcelAll_Focus = False Then
24200         .cmdExcelAll_box01.Visible = False
24210         .cmdExcelAll_box02.Visible = False
24220       End If
24230     End If
24240   End With

EXITP:
24250   Exit Sub

ERRH:
24260   Select Case ERR.Number
        Case Else
24270     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24280   End Select
24290   Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

24300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

24310   With Me
24320     .cmdPrintAll_box01.Visible = False
24330     .cmdPrintAll_box02.Visible = False
24340     blnPrintAll_Focus = False
24350   End With

EXITP:
24360   Exit Sub

ERRH:
24370   Select Case ERR.Number
        Case Else
24380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24390   End Select
24400   Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

24510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
24520   Exit Sub

ERRH:
24530   Select Case ERR.Number
        Case Else
24540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24550   End Select
24560   Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

24600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

24610   With Me
24620     blnWordAll_Focus = True
24630     .cmdWordAll_box01.Visible = True
24640     .cmdWordAll_box02.Visible = True
24650   End With

EXITP:
24660   Exit Sub

ERRH:
24670   Select Case ERR.Number
        Case Else
24680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24690   End Select
24700   Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

24800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

24810   With Me

24820     gblnPrintAll = True
24830     strFile1 = vbNullString: strFile2 = vbNullString

24840     cmdTaxIncDed_IncomeWord_Click  ' ** Procedure: Above.
24850     DoEvents
24860     If gblnPrintAll = True Then
24870       cmdTaxIncDed_DeductionWord_Click  ' ** Procedure: Above.
24880     End If

24890     gblnPrintAll = False

24900   End With

EXITP:
24910   Exit Sub

ERRH:
24920   gblnPrintAll = False
24930   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
24940   Case Else
24950     Select Case ERR.Number
          Case Else
24960       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
24970     End Select
24980   End Select
24990   Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

25000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

25010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
25020   intShiftDown = (Shift And acShiftMask) > 0
25030   intAltDown = (Shift And acAltMask) > 0
25040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
25050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
25060     Select Case intRetVal
          Case vbKeyTab
25070       With Me
25080         intRetVal = 0
25090         If .cmdExcelAll.Enabled = True Then
25100           .cmdExcelAll.SetFocus
25110         Else
25120           .cmdClose.SetFocus
25130         End If
25140       End With
25150     Case vbKeyUp
25160       With Me
25170         intRetVal = 0
25180         .cmdTaxIncDed_DeductionWord.SetFocus
25190       End With
25200     Case vbKeyDown
25210       With Me
25220         intRetVal = 0
25230         .cmdTaxIncDed_IncomeWord.SetFocus
25240       End With
25250     Case vbKeyLeft
25260       With Me
25270         intRetVal = 0
25280         .cmdPrintAll.SetFocus
25290       End With
25300     Case vbKeyRight
25310       With Me
25320         intRetVal = 0
25330         If .cmdExcelAll.Enabled = True Then
25340           .cmdExcelAll.SetFocus
25350         Else
25360           .cmdTaxIncDed_IncomePreview.SetFocus
25370         End If
25380       End With
25390     End Select
25400   End If

        ' ** Shift keys.
25410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
25420     Select Case intRetVal
          Case vbKeyTab
25430       With Me
25440         intRetVal = 0
25450         .cmdPrintAll.SetFocus
25460       End With
25470     End Select
25480   End If

        ' ** Ctrl keys.
25490   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
25500     Select Case intRetVal
          Case vbKeyTab
25510       With Me
25520         intRetVal = 0
25530         .chkOpenWord.SetFocus
25540       End With
25550     End Select
25560   End If

EXITP:
25570   KeyCode = intRetVal
25580   Exit Sub

ERRH:
25590   intRetVal = 0
25600   THAT_PROC = THIS_PROC
25610   That_Erl = Erl
25620   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
25630   Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

25700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

25710   With Me
25720     If gblnPrintAll = False Then
25730       .cmdWordAll_box01.Visible = True
25740       .cmdWordAll_box02.Visible = True
25750       If blnPrintAll_Focus = False Then
25760         .cmdPrintAll_box01.Visible = False
25770         .cmdPrintAll_box02.Visible = False
25780       End If
25790       If blnExcelAll_Focus = False Then
25800         .cmdExcelAll_box01.Visible = False
25810         .cmdExcelAll_box02.Visible = False
25820       End If
25830     End If
25840   End With

EXITP:
25850   Exit Sub

ERRH:
25860   Select Case ERR.Number
        Case Else
25870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25880   End Select
25890   Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

25900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

25910   With Me
25920     .cmdWordAll_box01.Visible = False
25930     .cmdWordAll_box02.Visible = False
25940     blnWordAll_Focus = False
25950   End With

EXITP:
25960   Exit Sub

ERRH:
25970   Select Case ERR.Number
        Case Else
25980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
25990   End Select
26000   Resume EXITP

End Sub

Private Sub chkOpenWord_AfterUpdate()

26100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

26110   With Me
26120     Select Case .chkOpenWord
          Case True
26130       .chkOpenWord_lbl.FontBold = True
26140       .chkOpenWord_lbl_dim_hi.FontBold = True
26150       .chkOpenWord_lbl2.FontBold = True
26160       .chkOpenWord_lbl2_dim_hi.FontBold = True
26170     Case False
26180       .chkOpenWord_lbl.FontBold = False
26190       .chkOpenWord_lbl_dim_hi.FontBold = False
26200       .chkOpenWord_lbl2.FontBold = False
26210       .chkOpenWord_lbl2_dim_hi.FontBold = False
26220     End Select
26230   End With

EXITP:
26240   Exit Sub

ERRH:
26250   Select Case ERR.Number
        Case Else
26260     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26270   End Select
26280   Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

26300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26320   intShiftDown = (Shift And acShiftMask) > 0
26330   intAltDown = (Shift And acAltMask) > 0
26340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26360     Select Case intRetVal
          Case vbKeyTab
26370       With Me
26380         intRetVal = 0
26390         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
26400           .cmdExcelAll.SetFocus
26410         Else
26420           .cmdClose.SetFocus
26430         End If
26440       End With
26450     Case vbKeyUp
26460       With Me
26470         intRetVal = 0
26480         .cmdWordAll.SetFocus
26490       End With
26500     Case vbKeyDown
26510       With Me
26520         intRetVal = 0
26530         .cmdTaxIncDed_IncomeWord.SetFocus
26540       End With
26550     End Select
26560   End If

        ' ** Shift keys.
26570   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26580     Select Case intRetVal
          Case vbKeyTab
26590       With Me
26600         intRetVal = 0
26610         .cmdWordAll.SetFocus
26620       End With
26630     End Select
26640   End If

        ' ** Ctrl keys.
26650   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
26660     Select Case intRetVal
          Case vbKeyTab
26670       With Me
26680         intRetVal = 0
26690         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
26700           .chkOpenExcel.SetFocus
26710         Else
26720           Beep
26730         End If
26740       End With
26750     End Select
26760   End If

EXITP:
26770   KeyCode = intRetVal
26780   Exit Sub

ERRH:
26790   intRetVal = 0
26800   THAT_PROC = THIS_PROC
26810   That_Erl = Erl
26820   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26830   Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

26910   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
26920   Exit Sub

ERRH:
26930   Select Case ERR.Number
        Case Else
26940     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26950   End Select
26960   Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

27000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

27010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
27020   Exit Sub

ERRH:
27030   Select Case ERR.Number
        Case Else
27040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27050   End Select
27060   Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

27110   With Me
27120     blnExcelAll_Focus = True
27130     .cmdExcelAll_box01.Visible = True
27140     .cmdExcelAll_box02.Visible = True
27150   End With

EXITP:
27160   Exit Sub

ERRH:
27170   Select Case ERR.Number
        Case Else
27180     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
27190   End Select
27200   Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

27300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

27310   With Me

27320     DoCmd.Hourglass True
27330     DoEvents

27340     gblnPrintAll = True
27350     strFile1 = vbNullString: strFile2 = vbNullString

          ' ** Excel not opened after export.
27360     cmdTaxIncDed_IncomeExcel_Click  ' ** Procedure: Above.
27370     DoEvents
27380     If gblnPrintAll = True Then
            ' ** Excel not opened after export.
27390       cmdTaxIncDed_DeductionExcel_Click  ' ** Procedure: Above.
27400     End If

27410     If gblnPrintAll = True Then
27420       Select Case .chkOpenExcel
            Case True
27430         DoCmd.Hourglass True
27440         DoEvents
27450         If Excel_NameOnly(strFile1, "Income") = True Then  ' ** Module Function: modExcelFuncs.
27460           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
27470             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
27480           End If
27490           OpenExe strFile1  ' ** Module Function: modShellFuncs.
27500         End If
27510         DoEvents
27520         ForcePause 2  ' ** Module Function: modCodeUtilities.
27530         If Excel_NameOnly(strFile2, "Deduction") = True Then  ' ** Module Function: modExcelFuncs.
27540           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
27550             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
27560           End If
27570           OpenExe strFile2  ' ** Module Function: modShellFuncs.
27580         End If
27590       Case False
              ' ** Nothing, that's it.
27600       End Select
27610       DoEvents
27620     End If

27630     gblnPrintAll = False

27640     DoCmd.Hourglass False

27650   End With

EXITP:
27660   Exit Sub

ERRH:
27670   gblnPrintAll = False
27680   DoCmd.Hourglass False
27690   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
27700   Case Else
27710     Select Case ERR.Number
          Case Else
27720       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
27730     End Select
27740   End Select
27750   Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

27800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

27810   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
27820   intShiftDown = (Shift And acShiftMask) > 0
27830   intAltDown = (Shift And acAltMask) > 0
27840   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
27850   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
27860     Select Case intRetVal
          Case vbKeyTab
27870       With Me
27880         intRetVal = 0
27890         .cmdClose.SetFocus
27900       End With
27910     Case vbKeyUp
27920       With Me
27930         intRetVal = 0
27940         .cmdTaxIncDed_DeductionExcel.SetFocus
27950       End With
27960     Case vbKeyDown
27970       With Me
27980         intRetVal = 0
27990         .cmdTaxIncDed_IncomeExcel.SetFocus
28000       End With
28010     Case vbKeyLeft
28020       With Me
28030         intRetVal = 0
28040         .cmdWordAll.SetFocus
28050       End With
28060     Case vbKeyRight
28070       With Me
28080         intRetVal = 0
28090         .cmdTaxIncDed_IncomePreview.SetFocus
28100       End With
28110     End Select
28120   End If

        ' ** Shift keys.
28130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
28140     Select Case intRetVal
          Case vbKeyTab
28150       With Me
28160         intRetVal = 0
28170         .cmdWordAll.SetFocus
28180       End With
28190     End Select
28200   End If

        ' ** Ctrl keys.
28210   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
28220     Select Case intRetVal
          Case vbKeyTab
28230       With Me
28240         intRetVal = 0
28250         .chkOpenExcel.SetFocus
28260       End With
28270     End Select
28280   End If

EXITP:
28290   KeyCode = intRetVal
28300   Exit Sub

ERRH:
28310   intRetVal = 0
28320   THAT_PROC = THIS_PROC
28330   That_Erl = Erl
28340   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
28350   Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

28400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

28410   With Me
28420     If gblnPrintAll = False Then
28430       .cmdExcelAll_box01.Visible = True
28440       .cmdExcelAll_box02.Visible = True
28450       If blnPrintAll_Focus = False Then
28460         .cmdPrintAll_box01.Visible = False
28470         .cmdPrintAll_box02.Visible = False
28480       End If
28490       If blnWordAll_Focus = False Then
28500         .cmdWordAll_box01.Visible = False
28510         .cmdWordAll_box02.Visible = False
28520       End If
28530     End If
28540   End With

EXITP:
28550   Exit Sub

ERRH:
28560   Select Case ERR.Number
        Case Else
28570     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28580   End Select
28590   Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

28600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

28610   With Me
28620     .cmdExcelAll_box01.Visible = False
28630     .cmdExcelAll_box02.Visible = False
28640     blnExcelAll_Focus = False
28650   End With

EXITP:
28660   Exit Sub

ERRH:
28670   Select Case ERR.Number
        Case Else
28680     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28690   End Select
28700   Resume EXITP

End Sub

Private Sub chkOpenExcel_AfterUpdate()

28800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

28810   With Me
28820     Select Case .chkOpenExcel
          Case True
28830       .chkOpenExcel_lbl.FontBold = True
28840       .chkOpenExcel_lbl_dim_hi.FontBold = True
28850       .chkOpenExcel_lbl2.FontBold = True
28860       .chkOpenExcel_lbl2_dim_hi.FontBold = True
28870     Case False
28880       .chkOpenExcel_lbl.FontBold = False
28890       .chkOpenExcel_lbl_dim_hi.FontBold = False
28900       .chkOpenExcel_lbl2.FontBold = False
28910       .chkOpenExcel_lbl2_dim_hi.FontBold = False
28920     End Select
28930   End With

EXITP:
28940   Exit Sub

ERRH:
28950   Select Case ERR.Number
        Case Else
28960     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28970   End Select
28980   Resume EXITP

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
29190         .cmdTaxIncDed_IncomeExcel.SetFocus
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

EXITP:
29310   KeyCode = intRetVal
29320   Exit Sub

ERRH:
29330   intRetVal = 0
29340   THAT_PROC = THIS_PROC
29350   That_Erl = Erl
29360   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
29370   Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

29400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

29410   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
29420   Exit Sub

ERRH:
29430   Select Case ERR.Number
        Case Else
29440     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29450   End Select
29460   Resume EXITP

End Sub

Private Sub ArchiveSet()

29500 On Error GoTo ERRH

        Const THIS_PROC As String = "ArchiveSet"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnHasArchs As Boolean

29510   Set dbs = CurrentDb
29520   With dbs
29530     Set rst = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
29540     With rst
29550       If .BOF = True And .EOF = True Then
29560         blnHasArchs = False
29570       Else
29580         .MoveLast
29590         If .RecordCount = 1 Then
29600           blnHasArchs = False
29610         Else
29620           blnHasArchs = True
29630         End If
29640       End If
29650       .Close
29660     End With
29670     Set rst = Nothing
29680     .Close
29690   End With
29700   Set dbs = Nothing

29710   With Me
29720     Select Case blnHasArchs
          Case True
29730       .chkIncludeArchive_lbl.Visible = True
29740       .chkIncludeArchive_lbl2.Visible = False
29750       .chkIncludeArchive_lbl2_dim_hi.Visible = False
29760     Case False
29770       .chkIncludeArchive = False
29780       .chkIncludeArchive.Enabled = False
29790       .chkIncludeArchive_lbl.Visible = False
29800       .chkIncludeArchive_lbl2.Visible = True
29810       .chkIncludeArchive_lbl2_dim_hi.Visible = True
29820     End Select
29830   End With

EXITP:
29840   Set rst = Nothing
29850   Set dbs = Nothing
29860   Exit Sub

ERRH:
29870   Select Case ERR.Number
        Case Else
29880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
29890   End Select
29900   Resume EXITP

End Sub

Public Function DoReport1() As Boolean
' ** INCOME.
' ** No queries are rewritten anymore, all use parameters.

30000 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport1"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngOutOfSync As Long
        Dim blnRetVal As Boolean

30010   blnRetVal = True

30020   With Me

30030     If IsNull(.DateStart) = True Or IsNull(.DateEnd) = True Then
30040       blnRetVal = False
30050       MsgBox "You must enter appropriate dates to continue.", vbInformation + vbOKOnly, "Entry Required"
30060       .DateStart.SetFocus
30070     Else
30080       If .cmbAccounts.Enabled = True Then
30090         .cmbAccounts.SetFocus
30100         If .cmbAccounts.text = vbNullString Then
30110           blnRetVal = False
30120           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
30130         End If
30140       End If
30150     End If

30160     If blnRetVal = True Then

30170       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
30180         gstrAccountNo = .cmbAccounts
30190       Case .opgAccountNumber_optAll.OptionValue
30200         gstrAccountNo = "All"
30210       End Select

30220       gdatStartDate = CDate(.DateStart)
30230       gdatEndDate = CDate(.DateEnd)

30240       Set dbs = CurrentDb

            ' ** Empty tmpTaxReports.
30250       Set qdf = dbs.QueryDefs("qryTaxReporting_11")
30260       qdf.Execute

30270       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specified account.
30280         Select Case .chkShowTrans
              Case True
                ' ** Show all transactions.
30290           Select Case .chkIncludeArchive
                Case True
30300             Set qdf = dbs.QueryDefs("qryTaxReporting_21_02a")
30310           Case False
30320             Set qdf = dbs.QueryDefs("qryTaxReporting_21_02")
30330           End Select
30340           With qdf.Parameters
30350             ![actno] = gstrAccountNo
30360             ![datbeg] = gdatStartDate
30370             ![datEnd] = gdatEndDate
30380           End With
30390           qdf.Execute
30400           Select Case .chkIncludeArchive
                Case True
30410             Set qdf = dbs.QueryDefs("qryTaxReporting_27_02a")
30420           Case False
30430             Set qdf = dbs.QueryDefs("qryTaxReporting_27_02")
30440           End Select
30450           With qdf.Parameters
30460             ![actno] = gstrAccountNo
30470             ![datbeg] = gdatStartDate
30480             ![datEnd] = gdatEndDate
30490           End With
30500           qdf.Execute
30510         Case False
                ' ** Only those with Tax Codes.
30520           Select Case .chkIncludeArchive
                Case True
30530             Set qdf = dbs.QueryDefs("qryTaxReporting_21_01a")
30540           Case False
30550             Set qdf = dbs.QueryDefs("qryTaxReporting_21_01")
30560           End Select
30570           With qdf.Parameters
30580             ![actno] = gstrAccountNo
30590             ![datbeg] = gdatStartDate
30600             ![datEnd] = gdatEndDate
30610           End With
30620           qdf.Execute
30630           Select Case .chkIncludeArchive
                Case True
30640             Set qdf = dbs.QueryDefs("qryTaxReporting_27_01a")
30650           Case False
30660             Set qdf = dbs.QueryDefs("qryTaxReporting_27_01")
30670           End Select
30680           With qdf.Parameters
30690             ![actno] = gstrAccountNo
30700             ![datbeg] = gdatStartDate
30710             ![datEnd] = gdatEndDate
30720           End With
30730           qdf.Execute
30740         End Select  ' ** chkShowTrans.
30750       Case .opgAccountNumber_optAll.OptionValue
              ' ** All accounts.
30760         Select Case .chkShowTrans
              Case True
                ' ** Show all transactions.
30770           Select Case .chkIncludeArchive
                Case True
30780             Set qdf = dbs.QueryDefs("qryTaxReporting_21_04a")
30790           Case False
30800             Set qdf = dbs.QueryDefs("qryTaxReporting_21_04")
30810           End Select
30820           With qdf.Parameters
30830             ![datbeg] = gdatStartDate
30840             ![datEnd] = gdatEndDate
30850           End With
30860           qdf.Execute
30870           Select Case .chkIncludeArchive
                Case True
30880             Set qdf = dbs.QueryDefs("qryTaxReporting_27_04a")
30890           Case False
30900             Set qdf = dbs.QueryDefs("qryTaxReporting_27_04")
30910           End Select
30920           With qdf.Parameters
30930             ![datbeg] = gdatStartDate
30940             ![datEnd] = gdatEndDate
30950           End With
30960           qdf.Execute
30970         Case False
                ' ** Only those with Tax Codes.
30980           Select Case .chkIncludeArchive
                Case True
30990             Set qdf = dbs.QueryDefs("qryTaxReporting_21_03a")
31000           Case False
31010             Set qdf = dbs.QueryDefs("qryTaxReporting_21_03")
31020           End Select
31030           With qdf.Parameters
31040             ![datbeg] = gdatStartDate
31050             ![datEnd] = gdatEndDate
31060           End With
31070           qdf.Execute
31080           Select Case .chkIncludeArchive
                Case True
31090             Set qdf = dbs.QueryDefs("qryTaxReporting_27_03a")
31100           Case False
31110             Set qdf = dbs.QueryDefs("qryTaxReporting_27_03")
31120           End Select
31130           With qdf.Parameters
31140             ![datbeg] = gdatStartDate
31150             ![datEnd] = gdatEndDate
31160           End With
31170           qdf.Execute
31180         End Select  ' ** chkShowTrans.
31190       End Select  ' ** opgAccountNumber.

            ' ** Update tmpTaxReports, taxcode = Null or 0 to 99. (So 'Unspecified' sorts together.)
31200       Set qdf = dbs.QueryDefs("qryTaxReporting_10")
31210       qdf.Execute

            ' ** OutOfSync: CLng(IIf(IsNull([taxcode])=True Or IsNull([revcode_TYPE])=True,1,
            ' ** IIf(([taxcode]>0 And [taxcode]<10) And [revcode_TYPE]=2,1,
            ' ** IIf([taxcode]>=10 And [revcode_TYPE]=1,1,0))))

31220       lngOutOfSync = 0&

            ' ** qryTaxReporting_14 (tmpTaxReports, for OutOfSync = 1, with taxcode_new,
            ' ** revcode_ID_new, revcode_TYPE_new), grouped and summed, by OutOfSync.
31230       Set qdf = dbs.QueryDefs("qryTaxReporting_15")
31240       Set rst = qdf.OpenRecordset
31250       With rst
31260         If .BOF = True And .EOF = True Then
                ' ** None out-of-sync.
31270         Else
31280           .MoveFirst
31290           lngOutOfSync = Nz(![OutOfSync], 0&)
31300         End If
31310         .Close
31320       End With
31330       If lngOutOfSync > 0& And gblnLinkRevTaxCodes = True Then
              ' ** Update qryTaxReporting_14.
              ' ** This just updates tmpTaxReports, not the Ledger.
31340         Set qdf = dbs.QueryDefs("qryTaxReporting_16")
31350         qdf.Execute
              ' ** Update Ledger, with qryTaxReporting_14's OutOfSync's.
31360         Set qdf = dbs.QueryDefs("qryTaxReporting_17")
31370         qdf.Execute
31380       End If

            ' ** qryTaxReporting_01, grouped and summed, by accountno, taxcode, totdesc, transdate, journaltype, assetno.
31390       Set qdf = dbs.QueryDefs("qryTaxReporting_02")
31400       Set rst = qdf.OpenRecordset
31410       With rst
31420         If .BOF = True And .EOF = True Then
31430           blnRetVal = False
31440           MsgBox "There is no data for the Income report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
31450         End If
31460         .Close
31470       End With

31480       dbs.Close

31490     End If  ' ** blnRetVal.

31500   End With

EXITP:
31510   Set rst = Nothing
31520   Set qdf = Nothing
31530   Set dbs = Nothing
31540   DoReport1 = blnRetVal
31550   Exit Function

ERRH:
31560   blnRetVal = False
31570   THAT_PROC = THIS_PROC
31580   That_Erl = Erl
31590   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
31600   Resume EXITP

End Function

Public Function DoReport2() As Boolean
' ** DEDUCTIONS.
' ** No queries are rewritten anymore, all use parameters.

31700 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport2"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

31710   blnRetVal = True  ' ** Unless proven otherwise.

31720   With Me

31730     If IsNull(.DateStart) = True Or IsNull(.DateEnd) = True Then
31740       blnRetVal = False
31750       MsgBox "You must enter appropriate dates to continue.", vbInformation + vbOKOnly, "Entry Required"
31760       .DateStart.SetFocus
31770     Else
31780       If cmbAccounts.Enabled = True Then
31790         cmbAccounts.SetFocus
31800         If cmbAccounts.text = vbNullString Then
31810           blnRetVal = False
31820           MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
31830         End If
31840       End If
31850     End If

31860     If blnRetVal = True Then

31870       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
31880         gstrAccountNo = .cmbAccounts
31890       Case .opgAccountNumber_optAll.OptionValue
31900         gstrAccountNo = "All"
31910       End Select

31920       gdatStartDate = CDate(.DateStart)
31930       gdatEndDate = CDate(.DateEnd)

31940       Set dbs = CurrentDb

            ' ** Empty tmpTaxReports.
31950       Set qdf = dbs.QueryDefs("qryTaxReporting_11")
31960       qdf.Execute

31970       Select Case .opgAccountNumber
            Case .opgAccountNumber_optSpecified.OptionValue
              ' ** Specified account.
31980         Select Case .chkShowTrans
              Case True
                ' ** Show all transactions.
31990           Select Case .chkIncludeArchive
                Case True
32000             Set qdf = dbs.QueryDefs("qryTaxReporting_31_02a")
32010           Case False
32020             Set qdf = dbs.QueryDefs("qryTaxReporting_31_02")
32030           End Select
32040           With qdf.Parameters
32050             ![actno] = gstrAccountNo
32060             ![datbeg] = gdatStartDate
32070             ![datEnd] = gdatEndDate
32080           End With
32090           qdf.Execute
32100           Select Case .chkIncludeArchive
                Case True
32110             Set qdf = dbs.QueryDefs("qryTaxReporting_37_02a")
32120           Case False
32130             Set qdf = dbs.QueryDefs("qryTaxReporting_37_02")
32140           End Select
32150           With qdf.Parameters
32160             ![actno] = gstrAccountNo
32170             ![datbeg] = gdatStartDate
32180             ![datEnd] = gdatEndDate
32190           End With
32200           qdf.Execute
32210         Case False
                ' ** Only those with Tax Codes.
32220           Select Case .chkIncludeArchive
                Case True
32230             Set qdf = dbs.QueryDefs("qryTaxReporting_31_01a")
32240           Case False
32250             Set qdf = dbs.QueryDefs("qryTaxReporting_31_01")
32260           End Select
32270           With qdf.Parameters
32280             ![actno] = gstrAccountNo
32290             ![datbeg] = gdatStartDate
32300             ![datEnd] = gdatEndDate
32310           End With
32320           qdf.Execute
32330           Select Case .chkIncludeArchive
                Case True
32340             Set qdf = dbs.QueryDefs("qryTaxReporting_37_01a")
32350           Case False
32360             Set qdf = dbs.QueryDefs("qryTaxReporting_37_01")
32370           End Select
32380           With qdf.Parameters
32390             ![actno] = gstrAccountNo
32400             ![datbeg] = gdatStartDate
32410             ![datEnd] = gdatEndDate
32420           End With
32430           qdf.Execute
32440         End Select  ' ** chkShowTrans.
32450       Case .opgAccountNumber_optAll.OptionValue
              ' ** All accounts.
32460         Select Case .chkShowTrans
              Case True
                ' ** Show all transactions.
32470           Select Case .chkIncludeArchive
                Case True
32480             Set qdf = dbs.QueryDefs("qryTaxReporting_31_04a")
32490           Case False
32500             Set qdf = dbs.QueryDefs("qryTaxReporting_31_04")
32510           End Select
32520           With qdf.Parameters
32530             ![datbeg] = gdatStartDate
32540             ![datEnd] = gdatEndDate
32550           End With
32560           qdf.Execute
32570           Select Case .chkIncludeArchive
                Case True
32580             Set qdf = dbs.QueryDefs("qryTaxReporting_37_04a")
32590           Case False
32600             Set qdf = dbs.QueryDefs("qryTaxReporting_37_04")
32610           End Select
32620           With qdf.Parameters
32630             ![datbeg] = gdatStartDate
32640             ![datEnd] = gdatEndDate
32650           End With
32660           qdf.Execute
32670         Case False
                ' ** Only those with Tax Codes.
32680           Select Case .chkIncludeArchive
                Case True
32690             Set qdf = dbs.QueryDefs("qryTaxReporting_31_03a")
32700           Case False
32710             Set qdf = dbs.QueryDefs("qryTaxReporting_31_03")
32720           End Select
32730           With qdf.Parameters
32740             ![datbeg] = gdatStartDate
32750             ![datEnd] = gdatEndDate
32760           End With
32770           qdf.Execute
32780           Select Case .chkIncludeArchive
                Case True
32790             Set qdf = dbs.QueryDefs("qryTaxReporting_37_03a")
32800           Case False
32810             Set qdf = dbs.QueryDefs("qryTaxReporting_37_03")
32820           End Select
32830           With qdf.Parameters
32840             ![datbeg] = gdatStartDate
32850             ![datEnd] = gdatEndDate
32860           End With
32870           qdf.Execute
32880         End Select  ' ** chkShowTrans.
32890       End Select  ' ** opgAccountNumber.

            ' ** Update tmpTaxReports, taxcode = Null or 0 to 99. (So 'Unspecified' sorts together.)
32900       Set qdf = dbs.QueryDefs("qryTaxReporting_10")
32910       qdf.Execute

            ' ** qryTaxReporting_01, grouped and summed, by accountno, taxcode, totdesc, transdate, journaltype, assetno.
32920       Set qdf = dbs.QueryDefs("qryTaxReporting_02")
32930       Set rst = qdf.OpenRecordset
32940       With rst
32950         If .BOF = True And .EOF = True Then
32960           blnRetVal = False
32970           MsgBox "There is no data for the Deductions report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
32980         End If
32990         .Close
33000       End With

33010       dbs.Close

33020     End If  ' ** blnRetVal.

33030   End With

EXITP:
33040   Set rst = Nothing
33050   Set qdf = Nothing
33060   Set dbs = Nothing
33070   DoReport2 = blnRetVal
33080   Exit Function

ERRH:
33090   blnRetVal = False
33100   THAT_PROC = THIS_PROC
33110   That_Erl = Erl
33120   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
33130   Resume EXITP

End Function
