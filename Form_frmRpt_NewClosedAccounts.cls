VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_NewClosedAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_NewClosedAccounts"

'VGC 09/29/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Closed Word:      D {cmdAcctNewClosed_ClosedWord}
' **   New Excel:        E {cmdAcctNewClosed_NewExcel}
' **   From:             F {DateStart}
' **   Closed Print:     N {cmdAcctNewClosed_ClosedPrint}
' **   New Word:         O {cmdAcctNewClosed_NewWord}
' **   New Print:        P {cmdAcctNewClosed_NewPrint}
' **   To:               T {DateEnd}
' **   Closed Preview:   V {cmdAcctNewClosed_ClosedPreview}
' **   New Preview:      W {cmdAcctNewClosed_NewPreview}
' **   Exit:             X {cmdClose}
' **   Closed Excel:     Z {cmdAcctNewClosed_ClosedExcel}

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

210           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
220           .cmdAcctNewClosed_NewExcel.Enabled = False
230           .cmdAcctNewClosed_ClosedExcel.Enabled = False
240           .cmdExcelAll.Enabled = False
250           .chkOpenExcel.Enabled = False
260           .chkOpenExcel_lbl2.ForeColor = WIN_CLR_DISF
270           .chkOpenExcel_lbl2_dim_hi.Visible = True
      #Else
280           .cmdAcctNewClosed_NewExcel.Enabled = True
290           .cmdAcctNewClosed_ClosedExcel.Enabled = True
300           .cmdExcelAll.Enabled = True
310           .chkOpenExcel.Enabled = True
320           .chkOpenExcel_lbl2.ForeColor = CLR_DKGRY
330           .chkOpenExcel_lbl2_dim_hi.Visible = False
      #End If

              ' ** EVENT CHECK: chkRememberDates!
340           If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
350             DateStart_Exit Cancel  ' ** Procedure: Below.
360             DateEnd_Exit Cancel  ' ** Procedure: Below.
370           End If

380           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
390           chkPageOf_AfterUpdate  ' ** Procedure: Below.
400           chkOpenWord_AfterUpdate  ' ** Procedure: Below.
410           chkOpenExcel_AfterUpdate  ' ** Procedure: Below.

420         Else
430           Cancel = -1
440         End If
450       Else
460         Cancel = -1
470       End If
480     End With

490     If Cancel = -1 Then
500       Beep
510       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
520     End If

530     DoCmd.Hourglass False

EXITP:
540     Exit Sub

ERRH:
550     DoCmd.Hourglass False
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
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

650       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

660     End With

EXITP:
670     Exit Sub

ERRH:
680     DoCmd.Hourglass False
690     Select Case ERR.Number
        Case Else
700       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
710     End Select
720     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Print All:        A {cmdPrintAll}
        ' **   Excel All:        E {cmdExcelAll}
        ' **   Page Count:       G {chkPageOf}
        ' **   Open Excel:       N {chkOpenExcel}
        ' **   Open Word:        O {chkOpenWord}
        ' **   Word All:         W {cmdWordAll}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       Case vbKeyF7
900         intRetVal = 0
910         cmdCalendar1_Click  ' ** Procedure: Below.
920       Case vbKeyF8
930         intRetVal = 0
940         cmdCalendar2_Click  ' ** Procedure: Below.
950       End Select
960     End If

        ' ** Alt keys.
970     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
980       Select Case intRetVal
          Case vbKeyX
990         intRetVal = 0
1000        cmdClose_Click  ' ** Procedure: Below.
1010      End Select
1020    End If

        ' ** Ctrl keys.
1030    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1040      Select Case intRetVal
          Case vbKeyA
1050        With Me
1060          intRetVal = 0
1070          .cmdPrintAll.SetFocus
1080          cmdPrintAll_Click  ' ** Procedure: Below.
1090        End With
1100      Case vbKeyE
1110        With Me
1120          intRetVal = 0
1130          .cmdExcelAll.SetFocus
1140          cmdExcelAll_Click  ' ** Procedure: Below.
1150        End With
1160      Case vbKeyG
1170        With Me
1180          intRetVal = 0
1190          If .chkPageOf.Enabled = True Then
1200            .chkPageOf.SetFocus
1210            .chkPageOf = (Not .chkPageOf)
1220            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1230          End If
1240        End With
1250      Case vbKeyN
1260        With Me
1270          intRetVal = 0
1280          .chkOpenExcel.SetFocus
1290          .chkOpenExcel = (Not .chkOpenExcel)
1300          chkOpenExcel_AfterUpdate  ' ** Procedure: Below.
1310        End With
1320      Case vbKeyO
1330        With Me
1340          intRetVal = 0
1350          .chkOpenWord.SetFocus
1360          .chkOpenWord = (Not .chkOpenWord)
1370          chkOpenWord_AfterUpdate  ' ** Procedure: Below.
1380        End With
1390      Case vbKeyW
1400        With Me
1410          intRetVal = 0
1420          .cmdWordAll.SetFocus
1430          cmdWordAll_Click  ' ** Procedure: Below.
1440        End With
1450      End Select
1460    End If

        ' ** Ctrl-Shift keys.
1470    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1480      Select Case intRetVal
          Case vbKeyF
1490        With Me
1500          intRetVal = 0
1510          .FocusHolder.SetFocus
1520          .DateStart = Null
1530        End With
1540      End Select
1550    End If

EXITP:
1560    KeyCode = intRetVal
1570    Exit Sub

ERRH:
1580    intRetVal = 0
1590    THAT_PROC = THIS_PROC
1600    That_Erl = Erl
1610    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1620    Resume EXITP

End Sub

Private Sub Form_Timer()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1710    With Me
1720      .TimerInterval = 0&
1730      If gblnSetFocus = True Then
1740        gblnSetFocus = False
1750        strFormName = vbNullString
1760  On Error Resume Next
1770        Do Until strFormName = THIS_NAME
1780          DoCmd.SelectObject acForm, THIS_NAME, False
1790          strFormName = Screen.ActiveForm.Name
1800        Loop
1810  On Error GoTo ERRH
1820      End If
1830      If gblnGoToReport = True Then
1840        Select Case blnGoingToReport
            Case True
1850          DoCmd.Hourglass False  ' ** It should already be off.
1860          blnGoingToReport = False
1870          gblnGoToReport = False
1880          .GoToReport_arw_rptnew_img.Visible = False
1890          .GoToReport_arw_rptclosed_img.Visible = False
1900          .cmdAcctNewClosed_NewWord.Visible = True
1910          .cmdAcctNewClosed_ClosedWord.Visible = True
1920          DoEvents
1930        Case False
1940          DoCmd.Hourglass True  ' ** Make sure it's still running.
1950          DoEvents
1960          blnGoingToReport = True
1970          Select Case garr_varGoToReport(GTR_RNAM)
              Case "rptNewClosedAccounts_New"
1980            .cmdAcctNewClosed_NewPrint.SetFocus
1990            .cmdAcctNewClosed_NewWord.Visible = False
2000            .GoToReport_arw_rptnew_img.Visible = True
2010          Case "rptNewClosedAccounts_Closed"
2020            .cmdAcctNewClosed_ClosedPrint.SetFocus
2030            .cmdAcctNewClosed_ClosedWord.Visible = False
2040            .GoToReport_arw_rptclosed_img.Visible = True
2050          End Select
2060          Beep
2070          DoCmd.Hourglass False
2080          DoEvents
2090          .TimerInterval = (GTR_WAIT * 2&)
2100        End Select
2110      Else
2120        DoCmd.Hourglass False
2130      End If
2140    End With

EXITP:
2150    Exit Sub

ERRH:
2160    DoCmd.Hourglass False
2170    Select Case ERR.Number
        Case Else
2180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2190    End Select
2200    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2310    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
2320      Response = acDataErrContinue
2330    Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
2340      Response = acDataErrContinue
2350    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2360      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
2370        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Date"
2380        Response = acDataErrContinue
2390      End If
2400    Case Else
2410      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2420      zErrorHandler Me.Name, THAT_PROC, DataErr, That_Erl
2430    End Select

EXITP:
2440    Exit Sub

ERRH:
2450    Select Case ERR.Number
        Case Else
2460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2470    End Select
2480    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2510    DoCmd.Hourglass True
2520    DoEvents

2530    With Me
2540      If .chkRememberDates = False Then
2550        .DateStart = Null  ' ** So it won't be there when Pref_Save() runs.
2560        .DateEnd = Null
2570      End If
2580    End With

2590    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2600    Do While Reports.Count > 0
2610      DoCmd.Close acReport, Reports(0).Name
2620    Loop

2630    gstrFormQuerySpec = vbNullString
2640    gblnSetFocus = True
2650    gstrReturningForm = THIS_NAME
2660    Select Case strCallingForm
        Case "frmMenu_Report"
2670      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2680    End Select

EXITP:
2690    Exit Sub

ERRH:
2700    Select Case ERR.Number
        Case Else
2710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2720    End Select
2730    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2910    With Me
2920      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
2930        Select Case blnCalendar1_Focus
            Case True
2940          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
2950          .cmdCalendar1_raised_img.Visible = False
2960        Case False
2970          .cmdCalendar1_raised_img.Visible = True
2980          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
2990        End Select
3000        .cmdCalendar1_raised_focus_dots_img.Visible = False
3010        .cmdCalendar1_raised_focus_img.Visible = False
3020        .cmdCalendar1_sunken_focus_dots_img.Visible = False
3030        .cmdCalendar1_raised_img_dis.Visible = False
3040      End If
3050      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
3060        Select Case blnCalendar2_Focus
            Case True
3070          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
3080          .cmdCalendar2_raised_img.Visible = False
3090        Case False
3100          .cmdCalendar2_raised_img.Visible = True
3110          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
3120        End Select
3130        .cmdCalendar2_raised_focus_dots_img.Visible = False
3140        .cmdCalendar2_raised_focus_img.Visible = False
3150        .cmdCalendar2_sunken_focus_dots_img.Visible = False
3160        .cmdCalendar2_raised_img_dis.Visible = False
3170      End If
3180      If blnPrintAll_Focus = False And (.cmdPrintAll_box01.Visible = True Or .cmdPrintAll_box02.Visible = True) Then
3190        .cmdPrintAll_box01.Visible = False
3200        .cmdPrintAll_box02.Visible = False
3210      End If
3220      If blnWordAll_Focus = False And (.cmdWordAll_box01.Visible = True Or .cmdWordAll_box02.Visible = True) Then
3230        .cmdWordAll_box01.Visible = False
3240        .cmdWordAll_box02.Visible = False
3250      End If
3260      If blnExcelAll_Focus = False And (.cmdExcelAll_box01.Visible = True Or .cmdExcelAll_box02.Visible = True) Then
3270        .cmdExcelAll_box01.Visible = False
3280        .cmdExcelAll_box02.Visible = False
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

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

3410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub cmdClose_Click()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

3510    DoCmd.Hourglass True
3520    DoEvents

3530    DoCmd.Close acForm, THIS_NAME

EXITP:
3540    Exit Sub

ERRH:
3550    Select Case ERR.Number
        Case Else
3560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3570    End Select
3580    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3620    intShiftDown = (Shift And acShiftMask) > 0
3630    intAltDown = (Shift And acAltMask) > 0
3640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3660      Select Case intRetVal
          Case vbKeyTab
3670        With Me
3680          intRetVal = 0
3690          .DateStart.SetFocus
3700        End With
3710      End Select
3720    End If

        ' ** Shift keys.
3730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3740      Select Case intRetVal
          Case vbKeyTab
3750        With Me
3760          intRetVal = 0
3770          If .cmdExcelAll.Enabled = True Then
3780            .cmdExcelAll.SetFocus
3790          Else
3800            .cmdWordAll.SetFocus
3810          End If
3820        End With
3830      End Select
3840    End If

        ' ** Ctrl-Shift keys.
3850    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3860      Select Case intRetVal
          Case vbKeyTab
3870        With Me
3880          intRetVal = 0
3890          If .chkOpenExcel.Enabled = True Then
3900            .chkOpenExcel.SetFocus
3910          Else
3920            .chkOpenWord.SetFocus
3930          End If
3940        End With
3950      End Select
3960    End If

EXITP:
3970    KeyCode = intRetVal
3980    Exit Sub

ERRH:
3990    intRetVal = 0
4000    THAT_PROC = THIS_PROC
4010    That_Erl = Erl
4020    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4030    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4220    intShiftDown = (Shift And acShiftMask) > 0
4230    intAltDown = (Shift And acAltMask) > 0
4240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4270        With Me
4280          intRetVal = 0
4290          .DateEnd.SetFocus
4300        End With
4310      End Select
4320    End If

        ' ** Shift keys.
4330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4350        With Me
4360          intRetVal = 0
4370          .cmdClose.SetFocus
4380        End With
4390      End Select
4400    End If

        ' ** Ctrl keys.
4410    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
4420      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4430        With Me
4440          intRetVal = 0
4450          .cmdCalendar1.SetFocus
4460        End With
4470      End Select
4480    End If

EXITP:
4490    KeyCode = intRetVal
4500    Exit Sub

ERRH:
4510    intRetVal = 0
4520    THAT_PROC = THIS_PROC
4530    That_Erl = Erl
4540    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4550    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

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

Private Sub DateStart_Exit(Cancel As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

4710    With Me
4720      If IsDate(.DateEnd) = True And IsDate(.DateStart) = True Then
4730        If CDate(.DateStart) > CDate(.DateEnd) Then
4740          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
4750          .DateEnd = Null
4760          .DateEnd.SetFocus
4770        End If
4780      End If
4790    End With

EXITP:
4800    Exit Sub

ERRH:
4810    THAT_PROC = THIS_PROC
4820    That_Erl = Erl
4830    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4840    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

4910    With Me
4920      blnCalendar1_Focus = True
4930      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4940      .cmdCalendar1_raised_img.Visible = False
4950      .cmdCalendar1_raised_focus_img.Visible = False
4960      .cmdCalendar1_raised_focus_dots_img.Visible = False
4970      .cmdCalendar1_sunken_focus_dots_img.Visible = False
4980      .cmdCalendar1_raised_img_dis.Visible = False
4990    End With

EXITP:
5000    Exit Sub

ERRH:
5010    Select Case ERR.Number
        Case Else
5020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5030    End Select
5040    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

5110    With Me
5120      blnCalendar1_MouseDown = True
5130      .cmdCalendar1_sunken_focus_dots_img.Visible = True
5140      .cmdCalendar1_raised_img.Visible = False
5150      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5160      .cmdCalendar1_raised_focus_img.Visible = False
5170      .cmdCalendar1_raised_focus_dots_img.Visible = False
5180      .cmdCalendar1_raised_img_dis.Visible = False
5190    End With

EXITP:
5200    Exit Sub

ERRH:
5210    Select Case ERR.Number
        Case Else
5220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5230    End Select
5240    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

5310    With Me
5320      datStartDate = Date
5330      datEndDate = 0
5340      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
5350      If blnRetVal = True Then
5360        .DateStart = datStartDate
5370      Else
5380        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
5390      End If
5400      .DateStart.SetFocus
5410    End With

EXITP:
5420    Exit Sub

ERRH:
5430    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
5440    Case Else
5450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5460    End Select
5470    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5520    intShiftDown = (Shift And acShiftMask) > 0
5530    intAltDown = (Shift And acAltMask) > 0
5540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5570        With Me
5580          intRetVal = 0
5590          .DateEnd.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          .DateStart.SetFocus
5680        End With
5690      End Select
5700    End If

EXITP:
5710    KeyCode = intRetVal
5720    Exit Sub

ERRH:
5730    intRetVal = 0
5740    THAT_PROC = THIS_PROC
5750    That_Erl = Erl
5760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5770    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

5810    With Me
5820      If blnCalendar1_MouseDown = False Then
5830        Select Case blnCalendar1_Focus
            Case True
5840          .cmdCalendar1_raised_focus_dots_img.Visible = True
5850          .cmdCalendar1_raised_focus_img.Visible = False
5860        Case False
5870          .cmdCalendar1_raised_focus_img.Visible = True
5880          .cmdCalendar1_raised_focus_dots_img.Visible = False
5890        End Select
5900        .cmdCalendar1_raised_img.Visible = False
5910        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5920        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5930        .cmdCalendar1_raised_img_dis.Visible = False
5940      End If
5950    End With

EXITP:
5960    Exit Sub

ERRH:
5970    Select Case ERR.Number
        Case Else
5980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5990    End Select
6000    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

6110    With Me
6120      .cmdCalendar1_raised_focus_dots_img.Visible = True
6130      .cmdCalendar1_raised_img.Visible = False
6140      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6150      .cmdCalendar1_raised_focus_img.Visible = False
6160      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6170      .cmdCalendar1_raised_img_dis.Visible = False
6180      blnCalendar1_MouseDown = False
6190    End With

EXITP:
6200    Exit Sub

ERRH:
6210    Select Case ERR.Number
        Case Else
6220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6230    End Select
6240    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

6310    With Me
6320      .cmdCalendar1_raised_img.Visible = True
6330      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6340      .cmdCalendar1_raised_focus_img.Visible = False
6350      .cmdCalendar1_raised_focus_dots_img.Visible = False
6360      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6370      .cmdCalendar1_raised_img_dis.Visible = False
6380      blnCalendar1_Focus = False
6390    End With

EXITP:
6400    Exit Sub

ERRH:
6410    Select Case ERR.Number
        Case Else
6420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6430    End Select
6440    Resume EXITP

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
6590          .cmdAcctNewClosed_NewPreview.SetFocus
6600        End With
6610      End Select
6620    End If

        ' ** Shift keys.
6630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6650        With Me
6660          intRetVal = 0
6670          .DateStart.SetFocus
6680        End With
6690      End Select
6700    End If

        ' ** Ctrl keys.
6710    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6720      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6730        With Me
6740          intRetVal = 0
6750          .cmdCalendar2.SetFocus
6760        End With
6770      End Select
6780    End If

        ' ** Ctrl-Shift keys.
6790    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6800      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6810        With Me
6820          intRetVal = 0
6830          .cmdCalendar1.SetFocus
6840        End With
6850      End Select
6860    End If

EXITP:
6870    KeyCode = intRetVal
6880    Exit Sub

ERRH:
6890    intRetVal = 0
6900    THAT_PROC = THIS_PROC
6910    That_Erl = Erl
6920    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6930    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

7010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
7020    Exit Sub

ERRH:
7030    Select Case ERR.Number
        Case Else
7040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7050    End Select
7060    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

7110    With Me
7120      If IsNull(.DateEnd) = True Then
            ' ** Populate it with today's date.
7130        .DateEnd = Date
7140      Else
7150        If IsNull(.DateStart) = False Then
7160          If CDate(.DateStart) > CDate(.DateEnd) Then
7170            MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
7180            .DateStart = Null
7190            .DateStart.SetFocus
7200          End If
7210        End If
7220      End If
7230    End With

EXITP:
7240    Exit Sub

ERRH:
7250    THAT_PROC = THIS_PROC
7260    That_Erl = Erl
7270    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7280    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

7310    With Me
7320      blnCalendar2_Focus = True
7330      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
7340      .cmdCalendar2_raised_img.Visible = False
7350      .cmdCalendar2_raised_focus_img.Visible = False
7360      .cmdCalendar2_raised_focus_dots_img.Visible = False
7370      .cmdCalendar2_sunken_focus_dots_img.Visible = False
7380      .cmdCalendar2_raised_img_dis.Visible = False
7390    End With

EXITP:
7400    Exit Sub

ERRH:
7410    Select Case ERR.Number
        Case Else
7420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7430    End Select
7440    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

7510    With Me
7520      blnCalendar2_MouseDown = True
7530      .cmdCalendar2_sunken_focus_dots_img.Visible = True
7540      .cmdCalendar2_raised_img.Visible = False
7550      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
7560      .cmdCalendar2_raised_focus_img.Visible = False
7570      .cmdCalendar2_raised_focus_dots_img.Visible = False
7580      .cmdCalendar2_raised_img_dis.Visible = False
7590    End With

EXITP:
7600    Exit Sub

ERRH:
7610    Select Case ERR.Number
        Case Else
7620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7630    End Select
7640    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

7710    With Me
7720      datStartDate = Date
7730      datEndDate = 0
7740      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
7750      If blnRetVal = True Then
7760        .DateEnd = datStartDate
7770      Else
7780        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
7790      End If
7800      .DateEnd.SetFocus
7810      Cancel = 0
7820      DateEnd_Exit Cancel  ' ** Procedure: Above.
7830      If Cancel = 0 Then
7840        .cmdAcctNewClosed_NewPreview.SetFocus
7850      End If
7860    End With

EXITP:
7870    Exit Sub

ERRH:
7880    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7890    Case Else
7900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7910    End Select
7920    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8020    intShiftDown = (Shift And acShiftMask) > 0
8030    intAltDown = (Shift And acAltMask) > 0
8040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8070        With Me
8080          intRetVal = 0
8090          .chkRememberDates.SetFocus
8100        End With
8110      End Select
8120    End If

        ' ** Shift keys.
8130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8150        With Me
8160          intRetVal = 0
8170          .DateEnd.SetFocus
8180        End With
8190      End Select
8200    End If

EXITP:
8210    KeyCode = intRetVal
8220    Exit Sub

ERRH:
8230    intRetVal = 0
8240    THAT_PROC = THIS_PROC
8250    That_Erl = Erl
8260    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8270    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

8310    With Me
8320      If blnCalendar2_MouseDown = False Then
8330        Select Case blnCalendar2_Focus
            Case True
8340          .cmdCalendar2_raised_focus_dots_img.Visible = True
8350          .cmdCalendar2_raised_focus_img.Visible = False
8360        Case False
8370          .cmdCalendar2_raised_focus_img.Visible = True
8380          .cmdCalendar2_raised_focus_dots_img.Visible = False
8390        End Select
8400        .cmdCalendar2_raised_img.Visible = False
8410        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8420        .cmdCalendar2_sunken_focus_dots_img.Visible = False
8430        .cmdCalendar2_raised_img_dis.Visible = False
8440      End If
8450    End With

EXITP:
8460    Exit Sub

ERRH:
8470    Select Case ERR.Number
        Case Else
8480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8490    End Select
8500    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

8610    With Me
8620      .cmdCalendar2_raised_focus_dots_img.Visible = True
8630      .cmdCalendar2_raised_img.Visible = False
8640      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8650      .cmdCalendar2_raised_focus_img.Visible = False
8660      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8670      .cmdCalendar2_raised_img_dis.Visible = False
8680      blnCalendar2_MouseDown = False
8690    End With

EXITP:
8700    Exit Sub

ERRH:
8710    Select Case ERR.Number
        Case Else
8720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8730    End Select
8740    Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

8810    With Me
8820      .cmdCalendar2_raised_img.Visible = True
8830      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8840      .cmdCalendar2_raised_focus_img.Visible = False
8850      .cmdCalendar2_raised_focus_dots_img.Visible = False
8860      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8870      .cmdCalendar2_raised_img_dis.Visible = False
8880      blnCalendar2_Focus = False
8890    End With

EXITP:
8900    Exit Sub

ERRH:
8910    Select Case ERR.Number
        Case Else
8920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8930    End Select
8940    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

9010    With Me
9020      Select Case .chkRememberDates
          Case True
9030        .chkRememberDates_lbl.FontBold = True
9040      Case False
9050        .chkRememberDates_lbl.FontBold = False
9060      End Select
9070    End With

EXITP:
9080    Exit Sub

ERRH:
9090    Select Case ERR.Number
        Case Else
9100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9110    End Select
9120    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

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
9290          .chkPageOf.SetFocus
9300        End With
9310      End Select
9320    End If

        ' ** Shift keys.
9330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9350        With Me
9360          intRetVal = 0
9370          .cmdCalendar2.SetFocus
9380        End With
9390      End Select
9400    End If

EXITP:
9410    KeyCode = intRetVal
9420    Exit Sub

ERRH:
9430    intRetVal = 0
9440    THAT_PROC = THIS_PROC
9450    That_Erl = Erl
9460    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9470    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

9510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
9520    Exit Sub

ERRH:
9530    Select Case ERR.Number
        Case Else
9540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9550    End Select
9560    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

9610    With Me
9620      Select Case .chkPageOf
          Case True
9630        .chkPageOf_lbl.FontBold = True
9640      Case False
9650        .chkPageOf_lbl.FontBold = False
9660      End Select
9670    End With

EXITP:
9680    Exit Sub

ERRH:
9690    DoCmd.Hourglass False
9700    Select Case ERR.Number
        Case Else
9710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9720    End Select
9730    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

9800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9820    intShiftDown = (Shift And acShiftMask) > 0
9830    intAltDown = (Shift And acAltMask) > 0
9840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9870        With Me
9880          intRetVal = 0
9890          .cmdAcctNewClosed_NewPreview.SetFocus
9900        End With
9910      End Select
9920    End If

        ' ** Shift keys.
9930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9950        With Me
9960          intRetVal = 0
9970          .chkRememberDates.SetFocus
9980        End With
9990      End Select
10000   End If

EXITP:
10010   KeyCode = intRetVal
10020   Exit Sub

ERRH:
10030   intRetVal = 0
10040   THAT_PROC = THIS_PROC
10050   That_Erl = Erl
10060   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10070   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewPreview_Click()

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewPreview_Click"

        Dim strDocName As String

10110   If DoReport("New") = True Then  ' ** Function: Below.
10120     strDocName = "rptNewClosedAccounts_New"
10130     DoCmd.OpenReport strDocName, acViewPreview
10140     DoCmd.Maximize
10150     DoCmd.RunCommand acCmdFitToWindow
10160   End If

EXITP:
10170   Exit Sub

ERRH:
10180   Select Case ERR.Number
        Case Else
10190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10200   End Select
10210   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewPreview_KeyDown(KeyCode As Integer, Shift As Integer)

10300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10320   intShiftDown = (Shift And acShiftMask) > 0
10330   intAltDown = (Shift And acAltMask) > 0
10340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10360     Select Case intRetVal
          Case vbKeyTab
10370       With Me
10380         intRetVal = 0
10390         .cmdAcctNewClosed_NewPrint.SetFocus
10400       End With
10410     Case vbKeyUp
10420       With Me
10430         intRetVal = 0
10440         .cmdAcctNewClosed_ClosedPreview.SetFocus
10450       End With
10460     Case vbKeyDown
10470       With Me
10480         intRetVal = 0
10490         .cmdAcctNewClosed_ClosedPreview.SetFocus
10500       End With
10510     Case vbKeyLeft
10520       With Me
10530         intRetVal = 0
10540         If .cmdExcelAll.Enabled = True Then
10550           .cmdExcelAll.SetFocus
10560         Else
10570           .cmdWordAll.SetFocus
10580         End If
10590       End With
10600     Case vbKeyRight
10610       With Me
10620         intRetVal = 0
10630         .cmdAcctNewClosed_NewPrint.SetFocus
10640       End With
10650     End Select
10660   End If

        ' ** Shift keys.
10670   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10680     Select Case intRetVal
          Case vbKeyTab
10690       With Me
10700         intRetVal = 0
10710         .DateEnd.SetFocus
10720       End With
10730     End Select
10740   End If

        ' ** Ctrl-Shift keys.
10750   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10760     Select Case intRetVal
          Case vbKeyTab
10770       With Me
10780         intRetVal = 0
10790         .chkPageOf.SetFocus
10800       End With
10810     End Select
10820   End If

EXITP:
10830   KeyCode = intRetVal
10840   Exit Sub

ERRH:
10850   intRetVal = 0
10860   THAT_PROC = THIS_PROC
10870   That_Erl = Erl
10880   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10890   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewPrint_Click()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewPrint_Click"

        Dim strDocName As String

10910   If DoReport("New") = True Then  ' ** Function: Below.
10920     strDocName = "rptNewClosedAccounts_New"
10930     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
10940       DoCmd.OpenReport strDocName, acViewPreview
10950       DoCmd.Maximize
10960       DoCmd.RunCommand acCmdFitToWindow
10970     Else
            '##GTR_Ref: rptNewClosedAccounts_New
10980       DoCmd.OpenReport strDocName, acViewNormal
10990     End If
11000   Else
11010     blnPrintBoth = False
11020   End If

EXITP:
11030   Exit Sub

ERRH:
11040   blnPrintBoth = False
11050   THAT_PROC = THIS_PROC
11060   That_Erl = Erl
11070   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11080   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewPrint_KeyDown(KeyCode As Integer, Shift As Integer)

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11120   intShiftDown = (Shift And acShiftMask) > 0
11130   intAltDown = (Shift And acAltMask) > 0
11140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11160     Select Case intRetVal
          Case vbKeyTab
11170       With Me
11180         intRetVal = 0
11190         .cmdAcctNewClosed_NewWord.SetFocus
11200       End With
11210     Case vbKeyUp
11220       With Me
11230         intRetVal = 0
11240         .cmdPrintAll.SetFocus
11250       End With
11260     Case vbKeyDown
11270       With Me
11280         intRetVal = 0
11290         .cmdAcctNewClosed_ClosedPrint.SetFocus
11300       End With
11310     Case vbKeyLeft
11320       With Me
11330         intRetVal = 0
11340         .cmdAcctNewClosed_NewPreview.SetFocus
11350       End With
11360     Case vbKeyRight
11370       With Me
11380         intRetVal = 0
11390         .cmdAcctNewClosed_NewWord.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab
11450       With Me
11460         intRetVal = 0
11470         .cmdAcctNewClosed_NewPreview.SetFocus
11480       End With
11490     End Select
11500   End If

EXITP:
11510   KeyCode = intRetVal
11520   Exit Sub

ERRH:
11530   intRetVal = 0
11540   THAT_PROC = THIS_PROC
11550   That_Erl = Erl
11560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11570   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewWord_Click()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

11610   With Me
11620     If DoReport("New") = True Then  ' ** Function: Below.

11630       If IsNull(.UserReportPath) = True Then
11640         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
11650       Else
11660         strRptPath = .UserReportPath
11670       End If

11680       strRpt = "rptNewClosedAccounts_New"
11690       gdatStartDate = .DateStart
11700       gdatEndDate = .DateEnd
11710       strRptCap = "New_Accounts_" & _
              Format$(gdatStartDate, "yymmdd") & "_To_" & Format$(gdatEndDate, "yymmdd")

11720       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

11730       If strRptPathFile <> vbNullString Then
11740         strFile1 = strRptPathFile
11750         Select Case .chkOpenWord
              Case True
11760           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
11770         Case False
11780           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
11790         End Select
11800         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
11810       Else
11820         gblnPrintAll = False
11830       End If

11840     Else
11850       gblnPrintAll = False
11860     End If
11870     gdatStartDate = 0
11880     gdatEndDate = 0
11890   End With

EXITP:
11900   Exit Sub

ERRH:
11910   gblnPrintAll = False
11920   Select Case ERR.Number
        Case Else
11930     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
11940   End Select
11950   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewWord_KeyDown(KeyCode As Integer, Shift As Integer)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewWord_KeyDown"

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
          Case vbKeyTab
12070       With Me
12080         intRetVal = 0
12090         If .cmdAcctNewClosed_NewExcel.Enabled = True Then
12100           .cmdAcctNewClosed_NewExcel.SetFocus
12110         Else
12120           .cmdAcctNewClosed_ClosedPreview.SetFocus
12130         End If
12140       End With
12150     Case vbKeyUp
12160       With Me
12170         intRetVal = 0
12180         .cmdAcctNewClosed_ClosedWord.SetFocus
12190       End With
12200     Case vbKeyDown
12210       With Me
12220         intRetVal = 0
12230         .cmdAcctNewClosed_ClosedWord.SetFocus
12240       End With
12250     Case vbKeyLeft
12260       With Me
12270         intRetVal = 0
12280         .cmdAcctNewClosed_NewPrint.SetFocus
12290       End With
12300     Case vbKeyRight
12310       With Me
12320         intRetVal = 0
12330         If .cmdAcctNewClosed_NewExcel.Enabled = True Then
12340           .cmdAcctNewClosed_NewExcel.SetFocus
12350         Else
12360           .cmdAcctNewClosed_ClosedPreview.SetFocus
12370         End If
12380       End With
12390     End Select
12400   End If

        ' ** Shift keys.
12410   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12420     Select Case intRetVal
          Case vbKeyTab
12430       With Me
12440         intRetVal = 0
12450         .cmdAcctNewClosed_NewPrint.SetFocus
12460       End With
12470     End Select
12480   End If

EXITP:
12490   KeyCode = intRetVal
12500   Exit Sub

ERRH:
12510   intRetVal = 0
12520   Select Case ERR.Number
        Case Else
12530     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12540   End Select
12550   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewExcel_Click()

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

12610   With Me

12620     DoCmd.Hourglass True
12630     DoEvents

12640     blnContinue = True

12650     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12660       DoCmd.Hourglass False
12670       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
12680       If msgResponse <> vbRetry Then
12690         blnContinue = False
12700       End If
12710     End If

12720     If blnContinue = True Then

12730       DoCmd.Hourglass True
12740       DoEvents

12750       If DoReport("New") = True Then  ' ** Procedure: Below.

12760         blnContinue = True

12770         If IsNull(.UserReportPath) = True Then
12780           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
12790         Else
12800           strRptPath = .UserReportPath
12810         End If

12820         strRptCap = "New_Accounts_" & Format$(.DateStart, "yymmdd") & "_To_" & Format$(.DateEnd, "yymmdd")

              ' ** qryRpt_NewClosedAccounts_New_17 (xx), For Export.
12830         strQry = "qryRpt_NewClosedAccounts_New_18"

12840         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

12850         If strRptPathFile <> vbNullString Then
12860           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
12870             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
12880           End If
12890           DoEvents
12900           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
12910             Kill strRptPathFile
12920           End If
12930           DoEvents
12940           gstrFormQuerySpec = THIS_NAME
12950           Select Case gblnPrintAll
                Case True
12960             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
12970             strFile1 = strRptPathFile
12980           Case False
12990             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
13000             If Excel_NameOnly(strRptPathFile, "Accounts New") = True Then  ' ** Module Function: modExcelFuncs.
13010               Select Case .chkOpenExcel
                    Case True
13020                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
13030                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
13040                 End If
13050                 OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
13060               Case False
                      ' ** Nothing, that's it.
13070               End Select
13080             End If
13090           End Select
13100           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
13110         End If  ' ** strRptPathFile.

13120       End If  ' ** DoReport().
13130     End If  ' ** blnContinue.

13140     DoCmd.Hourglass False

13150   End With

EXITP:
13160   Exit Sub

ERRH:
13170   DoCmd.Hourglass False
13180   Select Case ERR.Number
        Case Else
13190     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
13200   End Select
13210   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_NewExcel_KeyDown(KeyCode As Integer, Shift As Integer)

13300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_NewExcel_KeyDown"

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
          Case vbKeyTab
13370       With Me
13380         intRetVal = 0
13390         .cmdAcctNewClosed_ClosedPreview.SetFocus
13400       End With
13410     Case vbKeyUp
13420       With Me
13430         intRetVal = 0
13440         .cmdAcctNewClosed_ClosedExcel.SetFocus
13450       End With
13460     Case vbKeyDown
13470       With Me
13480         intRetVal = 0
13490         .cmdAcctNewClosed_ClosedExcel.SetFocus
13500       End With
13510     Case vbKeyLeft
13520       With Me
13530         intRetVal = 0
13540         .cmdAcctNewClosed_NewWord.SetFocus
13550       End With
13560     Case vbKeyRight
13570       With Me
13580         intRetVal = 0
13590         .cmdAcctNewClosed_ClosedPreview.SetFocus
13600       End With
13610     End Select
13620   End If

        ' ** Shift keys.
13630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13640     Select Case intRetVal
          Case vbKeyTab
13650       With Me
13660         intRetVal = 0
13670         .cmdAcctNewClosed_NewWord.SetFocus
13680       End With
13690     End Select
13700   End If

EXITP:
13710   KeyCode = intRetVal
13720   Exit Sub

ERRH:
13730   intRetVal = 0
13740   Select Case ERR.Number
        Case Else
13750     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13760   End Select
13770   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_Closed_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_Closed_box_MouseMove"

13810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
13820   Exit Sub

ERRH:
13830   Select Case ERR.Number
        Case Else
13840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13850   End Select
13860   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedPreview_Click()

13900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedPreview_Click"

        Dim strDocName As String

13910   If DoReport("Closed") = True Then  ' ** Function: Below.
13920     strDocName = "rptNewClosedAccounts_Closed"
13930     DoCmd.OpenReport strDocName, acViewPreview
13940     DoCmd.Maximize
13950     DoCmd.RunCommand acCmdFitToWindow
13960   End If

EXITP:
13970   Exit Sub

ERRH:
13980   Select Case ERR.Number
        Case Else
13990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
14000   End Select
14010   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedPreview_KeyDown(KeyCode As Integer, Shift As Integer)

14100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedPreview_KeyDown"

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
          Case vbKeyTab
14170       With Me
14180         intRetVal = 0
14190         .cmdAcctNewClosed_ClosedPrint.SetFocus
14200       End With
14210     Case vbKeyUp
14220       With Me
14230         intRetVal = 0
14240         .cmdAcctNewClosed_NewPreview.SetFocus
14250       End With
14260     Case vbKeyDown
14270       With Me
14280         intRetVal = 0
14290         .cmdAcctNewClosed_NewPreview.SetFocus
14300       End With
14310     Case vbKeyLeft
14320       With Me
14330         intRetVal = 0
14340         If .cmdAcctNewClosed_NewExcel.Enabled = True Then
14350           .cmdAcctNewClosed_NewExcel.SetFocus
14360         Else
14370           .cmdAcctNewClosed_NewWord.SetFocus
14380         End If
14390       End With
14400     Case vbKeyRight
14410       With Me
14420         intRetVal = 0
14430         .cmdAcctNewClosed_ClosedPrint.SetFocus
14440       End With
14450     End Select
14460   End If

        ' ** Shift keys.
14470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14480     Select Case intRetVal
          Case vbKeyTab
14490       With Me
14500         intRetVal = 0
14510         If .cmdAcctNewClosed_NewExcel.Enabled = True Then
14520           .cmdAcctNewClosed_NewExcel.SetFocus
14530         Else
14540           .cmdAcctNewClosed_NewWord.SetFocus
14550         End If
14560       End With
14570     End Select
14580   End If

EXITP:
14590   KeyCode = intRetVal
14600   Exit Sub

ERRH:
14610   intRetVal = 0
14620   THAT_PROC = THIS_PROC
14630   That_Erl = Erl
14640   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14650   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedPrint_Click()

14700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedPrint_Click"

        Dim strDocName As String

14710   If DoReport("Closed") = True Then  ' ** Function: Below.
14720     strDocName = "rptNewClosedAccounts_Closed"
14730     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
14740       DoCmd.OpenReport strDocName, acViewPreview
14750       DoCmd.Maximize
14760       DoCmd.RunCommand acCmdFitToWindow
14770     Else
            '##GTR_Ref: rptNewClosedAccounts_Closed
14780       DoCmd.OpenReport strDocName, acViewNormal
14790     End If
14800   Else
14810     blnPrintBoth = False
14820   End If

EXITP:
14830   Exit Sub

ERRH:
14840   blnPrintBoth = False
14850   THAT_PROC = THIS_PROC
14860   That_Erl = Erl
14870   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14880   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedPrint_KeyDown(KeyCode As Integer, Shift As Integer)

14900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedPrint_KeyDown"

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
          Case vbKeyTab
14970       With Me
14980         intRetVal = 0
14990         .cmdAcctNewClosed_ClosedWord.SetFocus
15000       End With
15010     Case vbKeyUp
15020       With Me
15030         intRetVal = 0
15040         .cmdAcctNewClosed_NewPrint.SetFocus
15050       End With
15060     Case vbKeyDown
15070       With Me
15080         intRetVal = 0
15090         .cmdPrintAll.SetFocus
15100       End With
15110     Case vbKeyLeft
15120       With Me
15130         intRetVal = 0
15140         .cmdAcctNewClosed_ClosedPreview.SetFocus
15150       End With
15160     Case vbKeyRight
15170       With Me
15180         intRetVal = 0
15190         .cmdAcctNewClosed_ClosedWord.SetFocus
15200       End With
15210     End Select
15220   End If

        ' ** Shift keys.
15230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15240     Select Case intRetVal
          Case vbKeyTab
15250       With Me
15260         intRetVal = 0
15270         .cmdAcctNewClosed_ClosedPreview.SetFocus
15280       End With
15290     End Select
15300   End If

EXITP:
15310   KeyCode = intRetVal
15320   Exit Sub

ERRH:
15330   intRetVal = 0
15340   THAT_PROC = THIS_PROC
15350   That_Erl = Erl
15360   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15370   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedWord_Click()

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

15410   With Me
15420     If DoReport("Closed") = True Then  ' ** Function: Below.

15430       If IsNull(.UserReportPath) = True Then
15440         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
15450       Else
15460         strRptPath = .UserReportPath
15470       End If

15480       strRpt = "rptNewClosedAccounts_Closed"
15490       gdatStartDate = .DateStart
15500       gdatEndDate = .DateEnd
15510       strRptCap = "Closed_Accounts_" & _
              Format$(gdatStartDate, "yymmdd") & "_To_" & Format$(gdatEndDate, "yymmdd")

15520       Select Case gblnPrintAll
            Case True
15530         strRptPathFile = strRptPath & LNK_SEP & strRptCap & ".rtf"
15540       Case False
15550         strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
15560       End Select

15570       If strRptPathFile <> vbNullString Then
15580         strFile2 = strRptPathFile
15590         Select Case .chkOpenWord
              Case True
15600           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
15610         Case False
15620           DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, False
15630         End Select
15640         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
15650       Else
15660         gblnPrintAll = False
15670       End If

15680     Else
15690       gblnPrintAll = False
15700     End If
15710     gdatStartDate = 0
15720     gdatEndDate = 0
15730   End With

EXITP:
15740   Exit Sub

ERRH:
15750   gblnPrintAll = False
15760   Select Case ERR.Number
        Case Else
15770     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
15780   End Select
15790   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedWord_KeyDown(KeyCode As Integer, Shift As Integer)

15800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedWord_KeyDown"

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
15890         If .cmdAcctNewClosed_ClosedExcel.Enabled = True Then
15900           .cmdAcctNewClosed_ClosedExcel.SetFocus
15910         Else
15920           .cmdPrintAll.SetFocus
15930         End If
15940       End With
15950     Case vbKeyUp
15960       With Me
15970         intRetVal = 0
15980         .cmdAcctNewClosed_NewWord.SetFocus
15990       End With
16000     Case vbKeyDown
16010       With Me
16020         intRetVal = 0
16030         .cmdAcctNewClosed_NewWord.SetFocus
16040       End With
16050     Case vbKeyLeft
16060       With Me
16070         intRetVal = 0
16080         .cmdAcctNewClosed_ClosedPrint.SetFocus
16090       End With
16100     Case vbKeyRight
16110       With Me
16120         intRetVal = 0
16130         If .cmdAcctNewClosed_ClosedExcel.Enabled = True Then
16140           .cmdAcctNewClosed_ClosedExcel.SetFocus
16150         Else
16160           .cmdPrintAll.SetFocus
16170         End If
16180       End With
16190     End Select
16200   End If

        ' ** Shift keys.
16210   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16220     Select Case intRetVal
          Case vbKeyTab
16230       With Me
16240         intRetVal = 0
16250         .cmdAcctNewClosed_ClosedPrint.SetFocus
16260       End With
16270     End Select
16280   End If

EXITP:
16290   KeyCode = intRetVal
16300   Exit Sub

ERRH:
16310   intRetVal = 0
16320   Select Case ERR.Number
        Case Else
16330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16340   End Select
16350   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedExcel_Click()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

16410   With Me

16420     DoCmd.Hourglass True
16430     DoEvents

16440     blnContinue = True

16450     If EXE_IsRunning("EXCEL.EXE") = True And gblnPrintAll = False Then  ' ** Module Function: modProcessFuncs.
16460       DoCmd.Hourglass False
16470       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
16480       If msgResponse <> vbRetry Then
16490         blnContinue = False
16500       End If
16510     End If

16520     If blnContinue = True Then

16530       DoCmd.Hourglass True
16540       DoEvents

16550       If DoReport("Closed") = True Then  ' ** Procedure: Below.

16560         blnContinue = True

16570         If IsNull(.UserReportPath) = True Then
16580           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
16590         Else
16600           strRptPath = .UserReportPath
16610         End If

16620         strRptCap = "Closed_Accounts_" & Format$(.DateStart, "yymmdd") & "_To_" & Format$(.DateEnd, "yymmdd")

              ' ** qryRpt_NewClosedAccounts_Closed_17 (xx), For Export.
16630         strQry = "qryRpt_NewClosedAccounts_Closed_18"

16640         Select Case gblnPrintAll
              Case True
16650           If strFile1 <> vbNullString Then
16660             strRptPathFile = StringReplace(strFile1, "New_", "Closed_")  ' ** Module Function: modStringFuncs.
16670           Else
16680             strRptPathFile = strRptPath & LNK_SEP & strRptCap & ".xls"
16690           End If
16700         Case False
16710           strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.
16720         End Select

16730         If strRptPathFile <> vbNullString Then
16740           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
16750             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
16760           End If
16770           DoEvents
16780           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
16790             Kill strRptPathFile
16800           End If
16810           DoEvents
16820           gstrFormQuerySpec = THIS_NAME
16830           Select Case gblnPrintAll
                Case True
16840             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
16850             strFile2 = strRptPathFile
16860           Case False
16870             DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
16880             If Excel_NameOnly(strRptPathFile, "Accounts Closed") = True Then  ' ** Module Function: modExcelFuncs.
16890               Select Case .chkOpenExcel
                    Case True
16900                 If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
16910                   EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
16920                 End If
16930                 OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
16940               Case False
                      ' ** Nothing, that's it.
16950               End Select
16960             End If
16970           End Select
16980           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
16990         End If

17000       End If  ' ** DoReport().
17010     End If  ' ** blnContinue.

17020     DoCmd.Hourglass False

17030   End With

EXITP:
17040   Exit Sub

ERRH:
17050   DoCmd.Hourglass False
17060   Select Case ERR.Number
        Case Else
17070     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
17080   End Select
17090   Resume EXITP

End Sub

Private Sub cmdAcctNewClosed_ClosedExcel_KeyDown(KeyCode As Integer, Shift As Integer)

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAcctNewClosed_ClosedExcel_KeyDown"

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
          Case vbKeyTab
17170       With Me
17180         intRetVal = 0
17190         .cmdPrintAll.SetFocus
17200       End With
17210     Case vbKeyUp
17220       With Me
17230         intRetVal = 0
17240         .cmdAcctNewClosed_NewExcel.SetFocus
17250       End With
17260     Case vbKeyDown
17270       With Me
17280         intRetVal = 0
17290         .cmdAcctNewClosed_NewExcel.SetFocus
17300       End With
17310     Case vbKeyLeft
17320       With Me
17330         intRetVal = 0
17340         .cmdAcctNewClosed_ClosedWord.SetFocus
17350       End With
17360     Case vbKeyRight
17370       With Me
17380         intRetVal = 0
17390         .cmdPrintAll.SetFocus
17400       End With
17410     End Select
17420   End If

        ' ** Shift keys.
17430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
17440     Select Case intRetVal
          Case vbKeyTab
17450       With Me
17460         intRetVal = 0
17470         .cmdAcctNewClosed_ClosedWord.SetFocus
17480       End With
17490     End Select
17500   End If

EXITP:
17510   KeyCode = intRetVal
17520   Exit Sub

ERRH:
17530   intRetVal = 0
17540   Select Case ERR.Number
        Case Else
17550     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17560   End Select
17570   Resume EXITP

End Sub

Private Sub cmdPrintAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_box01_MouseMove"

17610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
17620   Exit Sub

ERRH:
17630   Select Case ERR.Number
        Case Else
17640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17650   End Select
17660   Resume EXITP

End Sub

Private Sub cmdPrintAll_GotFocus()

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_GotFocus"

17710   With Me
17720     blnPrintAll_Focus = True
17730     .cmdPrintAll_box01.Visible = True
17740     .cmdPrintAll_box02.Visible = True
17750   End With

EXITP:
17760   Exit Sub

ERRH:
17770   Select Case ERR.Number
        Case Else
17780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17790   End Select
17800   Resume EXITP

End Sub

Private Sub cmdPrintAll_Click()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_Click"

17910   blnPrintBoth = True
17920   cmdAcctNewClosed_NewPrint_Click  ' ** Procedure: Above.
17930   DoEvents
17940   If blnPrintBoth = True Then
17950     cmdAcctNewClosed_ClosedPrint_Click  ' ** Procedure: Above.
17960   End If
17970   blnPrintBoth = False

EXITP:
17980   Exit Sub

ERRH:
17990   blnPrintBoth = False
18000   THAT_PROC = THIS_PROC
18010   That_Erl = Erl
18020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18030   Resume EXITP

End Sub

Private Sub cmdPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18110   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18120   intShiftDown = (Shift And acShiftMask) > 0
18130   intAltDown = (Shift And acAltMask) > 0
18140   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18150   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18160     Select Case intRetVal
          Case vbKeyTab
18170       With Me
18180         intRetVal = 0
18190         .cmdWordAll.SetFocus
18200       End With
18210     Case vbKeyUp
18220       With Me
18230         intRetVal = 0
18240         .cmdAcctNewClosed_ClosedPrint.SetFocus
18250       End With
18260     Case vbKeyDown
18270       With Me
18280         intRetVal = 0
18290         .cmdAcctNewClosed_NewPrint.SetFocus
18300       End With
18310     Case vbKeyLeft
18320       With Me
18330         intRetVal = 0
18340         If .cmdAcctNewClosed_ClosedExcel.Enabled = True Then
18350           .cmdAcctNewClosed_ClosedExcel.SetFocus
18360         Else
18370           .cmdAcctNewClosed_ClosedWord.SetFocus
18380         End If
18390       End With
18400     Case vbKeyRight
18410       With Me
18420         intRetVal = 0
18430         .cmdWordAll.SetFocus
18440       End With
18450     End Select
18460   End If

        ' ** Shift keys.
18470   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18480     Select Case intRetVal
          Case vbKeyTab
18490       With Me
18500         intRetVal = 0
18510         If .cmdAcctNewClosed_ClosedExcel.Enabled = True Then
18520           .cmdAcctNewClosed_ClosedExcel.SetFocus
18530         Else
18540           .cmdAcctNewClosed_ClosedWord.SetFocus
18550         End If
18560       End With
18570     End Select
18580   End If

EXITP:
18590   KeyCode = intRetVal
18600   Exit Sub

ERRH:
18610   intRetVal = 0
18620   THAT_PROC = THIS_PROC
18630   That_Erl = Erl
18640   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18650   Resume EXITP

End Sub

Private Sub cmdPrintAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_MouseMove"

18710   With Me
18720     If gblnPrintAll = False Then
18730       .cmdPrintAll_box01.Visible = True
18740       .cmdPrintAll_box02.Visible = True
18750       If blnWordAll_Focus = False Then
18760         .cmdWordAll_box01.Visible = False
18770         .cmdWordAll_box02.Visible = False
18780       End If
18790       If blnExcelAll_Focus = False Then
18800         .cmdExcelAll_box01.Visible = False
18810         .cmdExcelAll_box02.Visible = False
18820       End If
18830     End If
18840   End With

EXITP:
18850   Exit Sub

ERRH:
18860   Select Case ERR.Number
        Case Else
18870     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18880   End Select
18890   Resume EXITP

End Sub

Private Sub cmdPrintAll_LostFocus()

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintAll_LostFocus"

18910   With Me
18920     .cmdPrintAll_box01.Visible = False
18930     .cmdPrintAll_box02.Visible = False
18940     blnPrintAll_Focus = False
18950   End With

EXITP:
18960   Exit Sub

ERRH:
18970   Select Case ERR.Number
        Case Else
18980     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18990   End Select
19000   Resume EXITP

End Sub

Private Sub cmdWordAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_box01_MouseMove"

19110   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
19120   Exit Sub

ERRH:
19130   Select Case ERR.Number
        Case Else
19140     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19150   End Select
19160   Resume EXITP

End Sub

Private Sub cmdWordAll_GotFocus()

19200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_GotFocus"

19210   With Me
19220     blnWordAll_Focus = True
19230     .cmdWordAll_box01.Visible = True
19240     .cmdWordAll_box02.Visible = True
19250   End With

EXITP:
19260   Exit Sub

ERRH:
19270   Select Case ERR.Number
        Case Else
19280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19290   End Select
19300   Resume EXITP

End Sub

Private Sub cmdWordAll_Click()

19400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_Click"

19410   With Me

19420     gblnPrintAll = True
19430     strFile1 = vbNullString: strFile2 = vbNullString

19440     cmdAcctNewClosed_NewWord_Click  ' ** Procedure: Above.
19450     DoEvents
19460     If gblnPrintAll = True Then
19470       cmdAcctNewClosed_ClosedWord_Click  ' ** Procedure: Above.
19480     End If

19490     gblnPrintAll = False

19500   End With

EXITP:
19510   Exit Sub

ERRH:
19520   gblnPrintAll = False
19530   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
19540   Case Else
19550     Select Case ERR.Number
          Case Else
19560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
19570     End Select
19580   End Select
19590   Resume EXITP

End Sub

Private Sub cmdWordAll_KeyDown(KeyCode As Integer, Shift As Integer)

19600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_KeyDown"

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
          Case vbKeyTab
19670       With Me
19680         intRetVal = 0
19690         If .cmdExcelAll.Enabled = True Then
19700           .cmdExcelAll.SetFocus
19710         Else
19720           .cmdClose.SetFocus
19730         End If
19740       End With
19750     Case vbKeyUp
19760       With Me
19770         intRetVal = 0
19780         .cmdAcctNewClosed_ClosedWord.SetFocus
19790       End With
19800     Case vbKeyDown
19810       With Me
19820         intRetVal = 0
19830         .cmdAcctNewClosed_NewWord.SetFocus
19840       End With
19850     Case vbKeyLeft
19860       With Me
19870         intRetVal = 0
19880         .cmdPrintAll.SetFocus
19890       End With
19900     Case vbKeyRight
19910       With Me
19920         intRetVal = 0
19930         If .cmdExcelAll.Enabled = True Then
19940           .cmdExcelAll.SetFocus
19950         Else
19960           .cmdAcctNewClosed_NewPreview.SetFocus
19970         End If
19980       End With
19990     End Select
20000   End If

        ' ** Shift keys.
20010   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
20020     Select Case intRetVal
          Case vbKeyTab
20030       With Me
20040         intRetVal = 0
20050         .cmdPrintAll.SetFocus
20060       End With
20070     End Select
20080   End If

        ' ** Ctrl keys.
20090   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
20100     Select Case intRetVal
          Case vbKeyTab
20110       With Me
20120         intRetVal = 0
20130         .chkOpenWord.SetFocus
20140       End With
20150     End Select
20160   End If

EXITP:
20170   KeyCode = intRetVal
20180   Exit Sub

ERRH:
20190   intRetVal = 0
20200   THAT_PROC = THIS_PROC
20210   That_Erl = Erl
20220   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
20230   Resume EXITP

End Sub

Private Sub cmdWordAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_MouseMove"

20310   With Me
20320     If gblnPrintAll = False Then
20330       .cmdWordAll_box01.Visible = True
20340       .cmdWordAll_box02.Visible = True
20350       If blnPrintAll_Focus = False Then
20360         .cmdPrintAll_box01.Visible = False
20370         .cmdPrintAll_box02.Visible = False
20380       End If
20390       If blnExcelAll_Focus = False Then
20400         .cmdExcelAll_box01.Visible = False
20410         .cmdExcelAll_box02.Visible = False
20420       End If
20430     End If
20440   End With

EXITP:
20450   Exit Sub

ERRH:
20460   Select Case ERR.Number
        Case Else
20470     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20480   End Select
20490   Resume EXITP

End Sub

Private Sub cmdWordAll_LostFocus()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWordAll_LostFocus"

20510   With Me
20520     .cmdWordAll_box01.Visible = False
20530     .cmdWordAll_box02.Visible = False
20540     blnWordAll_Focus = False
20550   End With

EXITP:
20560   Exit Sub

ERRH:
20570   Select Case ERR.Number
        Case Else
20580     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20590   End Select
20600   Resume EXITP

End Sub

Private Sub chkOpenWord_AfterUpdate()

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_AfterUpdate"

20710   With Me
20720     Select Case .chkOpenWord
          Case True
20730       .chkOpenWord_lbl.FontBold = True
20740       .chkOpenWord_lbl_dim_hi.FontBold = True
20750       .chkOpenWord_lbl2.FontBold = True
20760       .chkOpenWord_lbl2_dim_hi.FontBold = True
20770     Case False
20780       .chkOpenWord_lbl.FontBold = False
20790       .chkOpenWord_lbl_dim_hi.FontBold = False
20800       .chkOpenWord_lbl2.FontBold = False
20810       .chkOpenWord_lbl2_dim_hi.FontBold = False
20820     End Select
20830   End With

EXITP:
20840   Exit Sub

ERRH:
20850   Select Case ERR.Number
        Case Else
20860     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20870   End Select
20880   Resume EXITP

End Sub

Private Sub chkOpenWord_KeyDown(KeyCode As Integer, Shift As Integer)

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

20910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
20920   intShiftDown = (Shift And acShiftMask) > 0
20930   intAltDown = (Shift And acAltMask) > 0
20940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
20950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
20960     Select Case intRetVal
          Case vbKeyTab
20970       With Me
20980         intRetVal = 0
20990         If .cmdExcelAll.Visible = True And .cmdExcelAll.Enabled = True Then
21000           .cmdExcelAll.SetFocus
21010         Else
21020           .cmdClose.SetFocus
21030         End If
21040       End With
21050     Case vbKeyUp
21060       With Me
21070         intRetVal = 0
21080         .cmdWordAll.SetFocus
21090       End With
21100     Case vbKeyDown
21110       With Me
21120         intRetVal = 0
21130         .cmdAcctNewClosed_NewWord.SetFocus
21140       End With
21150     End Select
21160   End If

        ' ** Shift keys.
21170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21180     Select Case intRetVal
          Case vbKeyTab
21190       With Me
21200         intRetVal = 0
21210         .cmdWordAll.SetFocus
21220       End With
21230     End Select
21240   End If

        ' ** Ctrl keys.
21250   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
21260     Select Case intRetVal
          Case vbKeyTab
21270       With Me
21280         intRetVal = 0
21290         If .chkOpenExcel.Visible = True And .chkOpenExcel.Enabled = True Then
21300           .chkOpenExcel.SetFocus
21310         Else
21320           Beep
21330         End If
21340       End With
21350     End Select
21360   End If

EXITP:
21370   KeyCode = intRetVal
21380   Exit Sub

ERRH:
21390   intRetVal = 0
21400   THAT_PROC = THIS_PROC
21410   That_Erl = Erl
21420   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21430   Resume EXITP

End Sub

Private Sub chkOpenWord_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenWord_MouseMove"

21510   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
21520   Exit Sub

ERRH:
21530   Select Case ERR.Number
        Case Else
21540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21550   End Select
21560   Resume EXITP

End Sub

Private Sub cmdExcelAll_box01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_box01_MouseMove"

21610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
21620   Exit Sub

ERRH:
21630   Select Case ERR.Number
        Case Else
21640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21650   End Select
21660   Resume EXITP

End Sub

Private Sub cmdExcelAll_GotFocus()

21700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_GotFocus"

21710   With Me
21720     blnExcelAll_Focus = True
21730     .cmdExcelAll_box01.Visible = True
21740     .cmdExcelAll_box02.Visible = True
21750   End With

EXITP:
21760   Exit Sub

ERRH:
21770   Select Case ERR.Number
        Case Else
21780     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21790   End Select
21800   Resume EXITP

End Sub

Private Sub cmdExcelAll_Click()

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_Click"

21910   With Me

21920     DoCmd.Hourglass True
21930     DoEvents

21940     gblnPrintAll = True
21950     strFile1 = vbNullString: strFile2 = vbNullString

          ' ** Excel not opened after export.
21960     cmdAcctNewClosed_NewExcel_Click  ' ** Procedure: Above.
21970     DoEvents
21980     If gblnPrintAll = True Then
            ' ** Excel not opened after export.
21990       cmdAcctNewClosed_ClosedExcel_Click  ' ** Procedure: Above.
22000     End If

22010     If gblnPrintAll = True Then
22020       Select Case .chkOpenExcel
            Case True
22030         DoCmd.Hourglass True
22040         DoEvents
22050         If Excel_NameOnly(strFile1, "Accounts New") = True Then  ' ** Module Function: modExcelFuncs.
22060           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22070             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22080           End If
22090           OpenExe strFile1  ' ** Module Function: modShellFuncs.
22100         End If
22110         DoEvents
22120         ForcePause 2  ' ** Module Function: modCodeUtilities.
22130         If Excel_NameOnly(strFile2, "Accounts Closed") = True Then  ' ** Module Function: modExcelFuncs.
22140           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
22150             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
22160           End If
22170           OpenExe strFile2  ' ** Module Function: modShellFuncs.
22180         End If
22190       Case False
              ' ** Nothing, that's it.
22200       End Select
22210       DoEvents
22220     End If

22230     gblnPrintAll = False

22240     DoCmd.Hourglass False

22250   End With

EXITP:
22260   Exit Sub

ERRH:
22270   gblnPrintAll = False
22280   DoCmd.Hourglass False
22290   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
22300   Case Else
22310     Select Case ERR.Number
          Case Else
22320       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl
22330     End Select
22340   End Select
22350   Resume EXITP

End Sub

Private Sub cmdExcelAll_KeyDown(KeyCode As Integer, Shift As Integer)

22400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_KeyDown"

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
          Case vbKeyTab
22470       With Me
22480         intRetVal = 0
22490         .cmdClose.SetFocus
22500       End With
22510     Case vbKeyUp
22520       With Me
22530         intRetVal = 0
22540         .cmdAcctNewClosed_ClosedExcel.SetFocus
22550       End With
22560     Case vbKeyDown
22570       With Me
22580         intRetVal = 0
22590         .cmdAcctNewClosed_NewExcel.SetFocus
22600       End With
22610     Case vbKeyLeft
22620       With Me
22630         intRetVal = 0
22640         .cmdWordAll.SetFocus
22650       End With
22660     Case vbKeyRight
22670       With Me
22680         intRetVal = 0
22690         .cmdAcctNewClosed_NewPreview.SetFocus
22700       End With
22710     End Select
22720   End If

        ' ** Shift keys.
22730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22740     Select Case intRetVal
          Case vbKeyTab
22750       With Me
22760         intRetVal = 0
22770         .cmdWordAll.SetFocus
22780       End With
22790     End Select
22800   End If

        ' ** Ctrl keys.
22810   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
22820     Select Case intRetVal
          Case vbKeyTab
22830       With Me
22840         intRetVal = 0
22850         .chkOpenExcel.SetFocus
22860       End With
22870     End Select
22880   End If

EXITP:
22890   KeyCode = intRetVal
22900   Exit Sub

ERRH:
22910   intRetVal = 0
22920   THAT_PROC = THIS_PROC
22930   That_Erl = Erl
22940   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
22950   Resume EXITP

End Sub

Private Sub cmdExcelAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

23000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_MouseMove"

23010   With Me
23020     If gblnPrintAll = False Then
23030       .cmdExcelAll_box01.Visible = True
23040       .cmdExcelAll_box02.Visible = True
23050       If blnPrintAll_Focus = False Then
23060         .cmdPrintAll_box01.Visible = False
23070         .cmdPrintAll_box02.Visible = False
23080       End If
23090       If blnWordAll_Focus = False Then
23100         .cmdWordAll_box01.Visible = False
23110         .cmdWordAll_box02.Visible = False
23120       End If
23130     End If
23140   End With

EXITP:
23150   Exit Sub

ERRH:
23160   Select Case ERR.Number
        Case Else
23170     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23180   End Select
23190   Resume EXITP

End Sub

Private Sub cmdExcelAll_LostFocus()

23200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcelAll_LostFocus"

23210   With Me
23220     .cmdExcelAll_box01.Visible = False
23230     .cmdExcelAll_box02.Visible = False
23240     blnExcelAll_Focus = False
23250   End With

EXITP:
23260   Exit Sub

ERRH:
23270   Select Case ERR.Number
        Case Else
23280     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23290   End Select
23300   Resume EXITP

End Sub

Private Sub chkOpenExcel_AfterUpdate()

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_AfterUpdate"

23410   With Me
23420     Select Case .chkOpenExcel
          Case True
23430       .chkOpenExcel_lbl.FontBold = True
23440       .chkOpenExcel_lbl_dim_hi.FontBold = True
23450       .chkOpenExcel_lbl2.FontBold = True
23460       .chkOpenExcel_lbl2_dim_hi.FontBold = True
23470     Case False
23480       .chkOpenExcel_lbl.FontBold = False
23490       .chkOpenExcel_lbl_dim_hi.FontBold = False
23500       .chkOpenExcel_lbl2.FontBold = False
23510       .chkOpenExcel_lbl2_dim_hi.FontBold = False
23520     End Select
23530   End With

EXITP:
23540   Exit Sub

ERRH:
23550   Select Case ERR.Number
        Case Else
23560     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23570   End Select
23580   Resume EXITP

End Sub

Private Sub chkOpenExcel_KeyDown(KeyCode As Integer, Shift As Integer)

23600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_KeyDown"

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
          Case vbKeyTab
23670       With Me
23680         intRetVal = 0
23690         .cmdClose.SetFocus
23700       End With
23710     Case vbKeyUp
23720       With Me
23730         intRetVal = 0
23740         .cmdExcelAll.SetFocus
23750       End With
23760     Case vbKeyDown
23770       With Me
23780         intRetVal = 0
23790         .cmdAcctNewClosed_NewExcel.SetFocus
23800       End With
23810     End Select
23820   End If

        ' ** Shift keys.
23830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23840     Select Case intRetVal
          Case vbKeyTab
23850       With Me
23860         intRetVal = 0
23870         .cmdExcelAll.SetFocus
23880       End With
23890     End Select
23900   End If

EXITP:
23910   KeyCode = intRetVal
23920   Exit Sub

ERRH:
23930   intRetVal = 0
23940   THAT_PROC = THIS_PROC
23950   That_Erl = Erl
23960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23970   Resume EXITP

End Sub

Private Sub chkOpenExcel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkOpenExcel_MouseMove"

24010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
24020   Exit Sub

ERRH:
24030   Select Case ERR.Number
        Case Else
24040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24050   End Select
24060   Resume EXITP

End Sub

Private Function DoReport(strMode As String) As Boolean

24100 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

24110   blnRetVal = True

24120   With Me
24130     If IsNull(.DateEnd) = True Then
            ' ** Populate it with today's date.
24140       .DateEnd = Date
24150     End If
24160     If IsNull(.DateStart) = True Then
24170       blnRetVal = False
24180       MsgBox "The start date must be prior to or the same as the end date.", vbInformation + vbOKOnly, "Entry Required"
24190       .DateStart.SetFocus
24200     Else
24210       If CDate(.DateStart) > CDate(.DateEnd) Then
24220         blnRetVal = False
24230         MsgBox "The start date must be prior to or the same as the end date.", vbInformation + vbOKOnly, "Invalid Date"
24240         .DateStart.SetFocus
24250       End If
24260     End If
24270   End With

24280   If blnRetVal = True Then
24290     Set dbs = CurrentDb
24300     With dbs
24310       Select Case strMode
            Case "New"
              ' ** Account, for appointmentdate, by specified FormRef().
24320         Set qdf = .QueryDefs("qryRpt_NewClosedAccounts_New_01")
24330       Case "Closed"
              ' ** Account, for dateclosed, by specified FormRef().
24340         Set qdf = .QueryDefs("qryRpt_NewClosedAccounts_Closed_01")
24350       End Select
24360       Set rst = qdf.OpenRecordset
24370       With rst
24380         If .BOF = True And .EOF = True Then
24390           blnRetVal = False
24400           MsgBox "There is no data for the report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
24410         End If
24420         .Close
24430       End With
24440       .Close
24450     End With
24460   End If

EXITP:
24470   Set rst = Nothing
24480   Set qdf = Nothing
24490   Set dbs = Nothing
24500   DoReport = blnRetVal
24510   Exit Function

ERRH:
24520   blnRetVal = False
24530   Select Case ERR.Number
        Case Else
24540     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
24550   End Select
24560   Resume EXITP

End Function
