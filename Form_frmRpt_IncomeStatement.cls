VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_IncomeStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_IncomeStatement"

'VGC 07/16/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Select Account:   A {cmbAccounts}
' **   Number:           B {opgAccountSource_optNumber}
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   From Date:        F {DateStart}
' **   Name:             M {opgAccountSource_optName}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   To Date:          T {DateEnd}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Focus Holder:     F {FocusHolder}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean

Private strCallingForm As String
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

170           gstrFormQuerySpec = THIS_NAME

              ' ** Reset the variables to ensure we get a new table everytime we come into the form.
180           gdatStartDate = 0
190           gdatEndDate = 0
200           gstrAccountNo = 0
210           gstrAccountName = vbNullString

220           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
230           .cmdExcel.Enabled = False
      #Else
240           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!
              ' ** EVENT CHECK: chkRememberDates!
250           If IsNull(.DateStart) = False And IsNull(.DateEnd) = False Then
260             DateStart_Exit Cancel  ' ** Procedure: Below.
270             DateEnd_Exit Cancel  ' ** Procedure: Below.
280           End If

290           opgAccountSource_AfterUpdate  ' ** Procedure: Below.
300           chkRememberMe_AfterUpdate  ' ** Procedure: Below.
310           chkPageOf_AfterUpdate  ' ** Procedure: Below.

320         Else
330           Cancel = -1
340         End If
350       Else
360         Cancel = -1
370       End If
380     End With

390     If Cancel = -1 Then
400       Beep
410       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
420     End If

430     DoCmd.Hourglass False

EXITP:
440     Exit Sub

ERRH:
450     DoCmd.Hourglass False
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
520       Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
530       clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
540       clsMonthClass.PositionAtCursor = True

550       .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

560     End With

EXITP:
570     Exit Sub

ERRH:
580     DoCmd.Hourglass False
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Focus Holder:     F {FocusHolder}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdClose_Click  ' ** Procedure: Below.
790       Case vbKeyF7
800         intRetVal = 0
810         cmdCalendar1_Click  ' ** Procedure: Below.
820       Case vbKeyF8
830         intRetVal = 0
840         cmdCalendar2_Click  ' ** Procedure: Below.
850       End Select
860     End If

        ' ** Alt keys.
870     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
880       Select Case intRetVal
          Case vbKeyX
890         intRetVal = 0
900         cmdClose_Click  ' ** Procedure: Below.
910       End Select
920     End If

        ' ** Ctrl keys.
930     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
940       Select Case intRetVal
          Case vbKeyG
950         With Me
960           intRetVal = 0
970           If .chkPageOf.Enabled = True Then
980             .chkPageOf.SetFocus
990             .chkPageOf = (Not .chkPageOf)
1000            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1010          End If
1020        End With
1030      End Select
1040    End If

        ' ** Ctrl-Shift keys.
1050    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1060      Select Case intRetVal
          Case vbKeyF
1070        With Me
1080          intRetVal = 0
1090          .FocusHolder.SetFocus
1100          .DateStart = Null
1110        End With
1120      End Select
1130    End If

EXITP:
1140    KeyCode = intRetVal
1150    Exit Sub

ERRH:
1160    intRetVal = 0
1170    THAT_PROC = THIS_PROC
1180    That_Erl = Erl
1190    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1200    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1310    With Me
1320      .TimerInterval = 0&
1330      If gblnSetFocus = True Then
1340        gblnSetFocus = False
1350        strFormName = vbNullString
1360  On Error Resume Next
1370        Do Until strFormName = THIS_NAME
1380          DoCmd.SelectObject acForm, THIS_NAME, False
1390          strFormName = Screen.ActiveForm.Name
1400        Loop
1410  On Error GoTo ERRH
1420      End If
1430      If gblnGoToReport = True Then
1440        If .cmdPrint.Enabled = True Then
1450          .cmdPrint.SetFocus
1460        Else
1470          .cmdClose.SetFocus
1480        End If
1490        gblnGoToReport = False
1500        Beep
1510        DoCmd.Hourglass False
1520        DoEvents
1530      Else
1540        DoCmd.Hourglass False
1550      End If
1560    End With

EXITP:
1570    Exit Sub

ERRH:
1580    DoCmd.Hourglass False
1590    Select Case ERR.Number
        Case Else
1600      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1610    End Select
1620    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1710    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
1720      Response = acDataErrContinue
1730    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1740      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateStart, DateEnd.
1750        MsgBox "You must enter the date in the format MM/DD/YYYY.", vbInformation + vbOKOnly, "Invalid Entry"
1760        Response = acDataErrContinue
1770      End If
1780    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
1790      Beep
1800      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
1810      Response = acDataErrContinue
1820    Case 2501  ' ** The '|' action was Canceled.
1830      Response = acDataErrContinue
1840    Case Else
1850      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1860      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1870    End Select

EXITP:
1880    Exit Sub

ERRH:
1890    Select Case ERR.Number
        Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2010    DoCmd.Hourglass True
2020    DoEvents

2030    With Me
2040      If .chkRememberMe = False Then
2050        .cmbAccounts = Null  ' ** So it won't be there when Pref_Save() runs.
2060      End If
2070    End With

2080    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2090    Do While Reports.Count > 0
2100      DoCmd.Close acReport, Reports(0).Name
2110    Loop

2120    gstrFormQuerySpec = vbNullString
2130    gblnSetFocus = True
2140    gstrReturningForm = THIS_NAME
2150    Select Case strCallingForm
        Case "frmMenu_Report"
2160      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2170    End Select

EXITP:
2180    Exit Sub

ERRH:
2190    Select Case ERR.Number
        Case Else
2200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2210    End Select
2220    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2410    With Me
2420      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
2430        Select Case blnCalendar1_Focus
            Case True
2440          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
2450          .cmdCalendar1_raised_img.Visible = False
2460        Case False
2470          .cmdCalendar1_raised_img.Visible = True
2480          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
2490        End Select
2500        .cmdCalendar1_raised_focus_dots_img.Visible = False
2510        .cmdCalendar1_raised_focus_img.Visible = False
2520        .cmdCalendar1_sunken_focus_dots_img.Visible = False
2530        .cmdCalendar1_raised_img_dis.Visible = False
2540      End If
2550      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
2560        Select Case blnCalendar2_Focus
            Case True
2570          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
2580          .cmdCalendar2_raised_img.Visible = False
2590        Case False
2600          .cmdCalendar2_raised_img.Visible = True
2610          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
2620        End Select
2630        .cmdCalendar2_raised_focus_dots_img.Visible = False
2640        .cmdCalendar2_raised_focus_img.Visible = False
2650        .cmdCalendar2_sunken_focus_dots_img.Visible = False
2660        .cmdCalendar2_raised_img_dis.Visible = False
2670      End If
2680    End With

EXITP:
2690    Exit Sub

ERRH:
2700    Select Case ERR.Number
        Case Else
2710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2720    End Select
2730    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2910    DoCmd.Hourglass True
2920    DoEvents

2930    DoCmd.Close acForm, THIS_NAME

EXITP:
2940    Exit Sub

ERRH:
2950    Select Case ERR.Number
        Case Else
2960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2970    End Select
2980    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3020    intShiftDown = (Shift And acShiftMask) > 0
3030    intAltDown = (Shift And acAltMask) > 0
3040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3060      Select Case intRetVal
          Case vbKeyTab
3070        With Me
3080          intRetVal = 0
3090          .DateStart.SetFocus
3100        End With
3110      End Select
3120    End If

        ' ** Shift keys.
3130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3140      Select Case intRetVal
          Case vbKeyTab
3150        With Me
3160          intRetVal = 0
3170          If .cmdExcel.Enabled = True Then
3180            .cmdExcel.SetFocus
3190          Else
3200            .cmdWord.SetFocus
3210          End If
3220        End With
3230      End Select
3240    End If

EXITP:
3250    KeyCode = intRetVal
3260    Exit Sub

ERRH:
3270    intRetVal = 0
3280    THAT_PROC = THIS_PROC
3290    That_Erl = Erl
3300    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3310    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

3410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
3420    Exit Sub

ERRH:
3430    Select Case ERR.Number
        Case Else
3440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3450    End Select
3460    Resume EXITP

End Sub

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3520    intShiftDown = (Shift And acShiftMask) > 0
3530    intAltDown = (Shift And acAltMask) > 0
3540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3560      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3570        With Me
3580          intRetVal = 0
3590          .DateEnd.SetFocus
3600        End With
3610      End Select
3620    End If

        ' ** Shift keys.
3630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3650        With Me
3660          intRetVal = 0
3670          .cmdClose.SetFocus
3680        End With
3690      End Select
3700    End If

        ' ** Ctrl keys.
3710    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3720      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3730        With Me
3740          intRetVal = 0
3750          .cmdCalendar1.SetFocus
3760        End With
3770      End Select
3780    End If

EXITP:
3790    KeyCode = intRetVal
3800    Exit Sub

ERRH:
3810    intRetVal = 0
3820    THAT_PROC = THIS_PROC
3830    That_Erl = Erl
3840    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3850    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

3910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

4010    With Me
4020      If IsDate(.DateEnd) = True And IsDate(.DateStart) = True Then
4030        If CDate(.DateStart) > CDate(.DateEnd) Then
4040          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
4050          .DateEnd = Null
4060          .DateEnd.SetFocus
4070        End If
4080      End If
4090    End With

EXITP:
4100    Exit Sub

ERRH:
4110    Select Case ERR.Number
        Case Else
4120      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4130    End Select
4140    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

4210    With Me
4220      blnCalendar1_Focus = True
4230      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4240      .cmdCalendar1_raised_img.Visible = False
4250      .cmdCalendar1_raised_focus_img.Visible = False
4260      .cmdCalendar1_raised_focus_dots_img.Visible = False
4270      .cmdCalendar1_sunken_focus_dots_img.Visible = False
4280      .cmdCalendar1_raised_img_dis.Visible = False
4290    End With

EXITP:
4300    Exit Sub

ERRH:
4310    Select Case ERR.Number
        Case Else
4320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4330    End Select
4340    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

4410    With Me
4420      blnCalendar1_MouseDown = True
4430      .cmdCalendar1_sunken_focus_dots_img.Visible = True
4440      .cmdCalendar1_raised_img.Visible = False
4450      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
4460      .cmdCalendar1_raised_focus_img.Visible = False
4470      .cmdCalendar1_raised_focus_dots_img.Visible = False
4480      .cmdCalendar1_raised_img_dis.Visible = False
4490    End With

EXITP:
4500    Exit Sub

ERRH:
4510    Select Case ERR.Number
        Case Else
4520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4530    End Select
4540    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

4610    With Me
4620      datStartDate = Date
4630      datEndDate = 0
4640      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
4650      If blnRetVal = True Then
4660        .DateStart = datStartDate
4670      Else
4680        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
4690      End If
4700      .DateStart.SetFocus
4710    End With

EXITP:
4720    Exit Sub

ERRH:
4730    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
4740    Case Else
4750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4760    End Select
4770    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4820    intShiftDown = (Shift And acShiftMask) > 0
4830    intAltDown = (Shift And acAltMask) > 0
4840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4870        With Me
4880          intRetVal = 0
4890          .DateEnd.SetFocus
4900        End With
4910      End Select
4920    End If

        ' ** Shift keys.
4930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4950        With Me
4960          intRetVal = 0
4970          .DateStart.SetFocus
4980        End With
4990      End Select
5000    End If

EXITP:
5010    KeyCode = intRetVal
5020    Exit Sub

ERRH:
5030    intRetVal = 0
5040    THAT_PROC = THIS_PROC
5050    That_Erl = Erl
5060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5070    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

5110    With Me
5120      If blnCalendar1_MouseDown = False Then
5130        Select Case blnCalendar1_Focus
            Case True
5140          .cmdCalendar1_raised_focus_dots_img.Visible = True
5150          .cmdCalendar1_raised_focus_img.Visible = False
5160        Case False
5170          .cmdCalendar1_raised_focus_img.Visible = True
5180          .cmdCalendar1_raised_focus_dots_img.Visible = False
5190        End Select
5200        .cmdCalendar1_raised_img.Visible = False
5210        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5220        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5230        .cmdCalendar1_raised_img_dis.Visible = False
5240      End If
5250    End With

EXITP:
5260    Exit Sub

ERRH:
5270    Select Case ERR.Number
        Case Else
5280      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5290    End Select
5300    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

5410    With Me
5420      .cmdCalendar1_raised_focus_dots_img.Visible = True
5430      .cmdCalendar1_raised_img.Visible = False
5440      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5450      .cmdCalendar1_raised_focus_img.Visible = False
5460      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5470      .cmdCalendar1_raised_img_dis.Visible = False
5480      blnCalendar1_MouseDown = False
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

Private Sub cmdCalendar1_LostFocus()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

5610    With Me
5620      .cmdCalendar1_raised_img.Visible = True
5630      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5640      .cmdCalendar1_raised_focus_img.Visible = False
5650      .cmdCalendar1_raised_focus_dots_img.Visible = False
5660      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5670      .cmdCalendar1_raised_img_dis.Visible = False
5680      blnCalendar1_Focus = False
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

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5820    intShiftDown = (Shift And acShiftMask) > 0
5830    intAltDown = (Shift And acAltMask) > 0
5840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5870        With Me
5880          intRetVal = 0
5890          .cmbAccounts.SetFocus
5900        End With
5910      End Select
5920    End If

        ' ** Shift keys.
5930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5950        With Me
5960          intRetVal = 0
5970          .DateStart.SetFocus
5980        End With
5990      End Select
6000    End If

        ' ** Ctrl keys.
6010    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6020      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6030        With Me
6040          intRetVal = 0
6050          .cmdCalendar2.SetFocus
6060        End With
6070      End Select
6080    End If

        ' ** Ctrl-Shift keys.
6090    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6100      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6110        With Me
6120          intRetVal = 0
6130          .cmdCalendar1.SetFocus
6140        End With
6150      End Select
6160    End If

EXITP:
6170    KeyCode = intRetVal
6180    Exit Sub

ERRH:
6190    intRetVal = 0
6200    THAT_PROC = THIS_PROC
6210    That_Erl = Erl
6220    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6230    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

6310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
6320    Exit Sub

ERRH:
6330    Select Case ERR.Number
        Case Else
6340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6350    End Select
6360    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

6410    With Me
6420      If IsNull(.DateEnd) = True Then
            ' ** Populate it with today's date.
6430        .DateEnd = Date
6440      Else
6450        If .DateEnd = vbNullString Then
              ' ** Populate it with today's date.
6460          .DateEnd = Date
6470        Else
6480          If IsNull(.DateStart) = False Then
6490            If CDate(.DateStart) > CDate(.DateEnd) Then
6500              MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
6510              .DateStart = Null
6520              .DateStart.SetFocus
6530            End If
6540          End If
6550        End If
6560      End If
6570    End With

EXITP:
6580    Exit Sub

ERRH:
6590    THAT_PROC = THIS_PROC
6600    That_Erl = Erl
6610    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
6620    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

6710    With Me
6720      blnCalendar2_Focus = True
6730      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
6740      .cmdCalendar2_raised_img.Visible = False
6750      .cmdCalendar2_raised_focus_img.Visible = False
6760      .cmdCalendar2_raised_focus_dots_img.Visible = False
6770      .cmdCalendar2_sunken_focus_dots_img.Visible = False
6780      .cmdCalendar2_raised_img_dis.Visible = False
6790    End With

EXITP:
6800    Exit Sub

ERRH:
6810    Select Case ERR.Number
        Case Else
6820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6830    End Select
6840    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

6910    With Me
6920      blnCalendar2_MouseDown = True
6930      .cmdCalendar2_sunken_focus_dots_img.Visible = True
6940      .cmdCalendar2_raised_img.Visible = False
6950      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
6960      .cmdCalendar2_raised_focus_img.Visible = False
6970      .cmdCalendar2_raised_focus_dots_img.Visible = False
6980      .cmdCalendar2_raised_img_dis.Visible = False
6990    End With

EXITP:
7000    Exit Sub

ERRH:
7010    Select Case ERR.Number
        Case Else
7020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7030    End Select
7040    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

7110    With Me
7120      datStartDate = Date
7130      datEndDate = 0
7140      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
7150      If blnRetVal = True Then
7160        .DateEnd = datStartDate
7170      Else
7180        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
7190      End If
7200      .DateEnd.SetFocus
7210      Cancel = 0
7220      DateEnd_Exit Cancel  ' ** Procedure: Above.
7230      If Cancel = 0 Then
7240        .cmdPreview.SetFocus
7250      End If
7260    End With

EXITP:
7270    Exit Sub

ERRH:
7280    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7290    Case Else
7300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7310    End Select
7320    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

7400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

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
7490          .chkRememberDates.SetFocus
7500        End With
7510      End Select
7520    End If

        ' ** Shift keys.
7530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7550        With Me
7560          intRetVal = 0
7570          .DateEnd.SetFocus
7580        End With
7590      End Select
7600    End If

EXITP:
7610    KeyCode = intRetVal
7620    Exit Sub

ERRH:
7630    intRetVal = 0
7640    THAT_PROC = THIS_PROC
7650    That_Erl = Erl
7660    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7670    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

7710    With Me
7720      If blnCalendar2_MouseDown = False Then
7730        Select Case blnCalendar2_Focus
            Case True
7740          .cmdCalendar2_raised_focus_dots_img.Visible = True
7750          .cmdCalendar2_raised_focus_img.Visible = False
7760        Case False
7770          .cmdCalendar2_raised_focus_img.Visible = True
7780          .cmdCalendar2_raised_focus_dots_img.Visible = False
7790        End Select
7800        .cmdCalendar2_raised_img.Visible = False
7810        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
7820        .cmdCalendar2_sunken_focus_dots_img.Visible = False
7830        .cmdCalendar2_raised_img_dis.Visible = False
7840      End If
7850    End With

EXITP:
7860    Exit Sub

ERRH:
7870    Select Case ERR.Number
        Case Else
7880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7890    End Select
7900    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

8010    With Me
8020      .cmdCalendar2_raised_focus_dots_img.Visible = True
8030      .cmdCalendar2_raised_img.Visible = False
8040      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8050      .cmdCalendar2_raised_focus_img.Visible = False
8060      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8070      .cmdCalendar2_raised_img_dis.Visible = False
8080      blnCalendar2_MouseDown = False
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

Private Sub cmdCalendar2_LostFocus()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

8210    With Me
8220      .cmdCalendar2_raised_img.Visible = True
8230      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8240      .cmdCalendar2_raised_focus_img.Visible = False
8250      .cmdCalendar2_raised_focus_dots_img.Visible = False
8260      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8270      .cmdCalendar2_raised_img_dis.Visible = False
8280      blnCalendar2_Focus = False
8290    End With

EXITP:
8300    Exit Sub

ERRH:
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

8410    With Me
8420      Select Case .chkRememberDates
          Case True
8430        .chkRememberDates_lbl.FontBold = True
8440      Case False
8450        .chkRememberDates_lbl.FontBold = False
8460      End Select
8470    End With

EXITP:
8480    Exit Sub

ERRH:
8490    THAT_PROC = THIS_PROC
8500    That_Erl = Erl
8510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
8520    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

8600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

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
8690          .opgAccountSource.SetFocus
8700        End With
8710      End Select
8720    End If

        ' ** Shift keys.
8730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8750        With Me
8760          intRetVal = 0
8770          .cmdCalendar2.SetFocus
8780        End With
8790      End Select
8800    End If

EXITP:
8810    KeyCode = intRetVal
8820    Exit Sub

ERRH:
8830    intRetVal = 0
8840    THAT_PROC = THIS_PROC
8850    That_Erl = Erl
8860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8870    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

8910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
8920    Exit Sub

ERRH:
8930    Select Case ERR.Number
        Case Else
8940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8950    End Select
8960    Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

9010    strAccountNo = vbNullString

9020    With Me
9030      If IsNull(.cmbAccounts) = False Then
9040        If Len(.cmbAccounts.Column(0)) > 0 Then
9050          strAccountNo = .cmbAccounts.Column(0)
9060        End If
9070      End If
9080      Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
9090        .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
9100        .opgAccountSource_optNumber_lbl.FontBold = True
9110        .opgAccountSource_optName_lbl.FontBold = False
9120      Case .opgAccountSource_optName.OptionValue
9130        .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
9140        .opgAccountSource_optNumber_lbl.FontBold = False
9150        .opgAccountSource_optName_lbl.FontBold = True
9160      End Select
9170      DoEvents
9180      If strAccountNo <> vbNullString Then
9190        .cmbAccounts = strAccountNo
9200      End If
9210    End With

EXITP:
9220    Exit Sub

ERRH:
9230    THAT_PROC = THIS_PROC
9240    That_Erl = Erl
9250    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
9260    Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

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
9390          .chkRememberMe.SetFocus
9400        End With
9410      End Select
9420    End If

        ' ** Shift keys.
9430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9450        With Me
9460          intRetVal = 0
9470          .chkRememberDates.SetFocus
9480        End With
9490      End Select
9500    End If

EXITP:
9510    KeyCode = intRetVal
9520    Exit Sub

ERRH:
9530    intRetVal = 0
9540    THAT_PROC = THIS_PROC
9550    That_Erl = Erl
9560    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9570    Resume EXITP

End Sub

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9620    intShiftDown = (Shift And acShiftMask) > 0
9630    intAltDown = (Shift And acAltMask) > 0
9640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9670        With Me
9680          intRetVal = 0
9690          .chkRememberMe.SetFocus
9700        End With
9710      End Select
9720    End If

        ' ** Shift keys.
9730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9750        With Me
9760          intRetVal = 0
9770          .chkRememberDates.SetFocus
9780        End With
9790      End Select
9800    End If

EXITP:
9810    KeyCode = intRetVal
9820    Exit Sub

ERRH:
9830    intRetVal = 0
9840    THAT_PROC = THIS_PROC
9850    That_Erl = Erl
9860    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9870    Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

9910    With Me
9920      Select Case .chkRememberMe
          Case True
9930        .chkRememberMe_lbl.FontBold = True
9940      Case False
9950        .chkRememberMe_lbl.FontBold = False
9960      End Select
9970    End With

EXITP:
9980    Exit Sub

ERRH:
9990    THAT_PROC = THIS_PROC
10000   That_Erl = Erl
10010   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
10020   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

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
10190         .cmbAccounts.SetFocus
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10250       With Me
10260         intRetVal = 0
10270         .opgAccountSource.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
10330   intRetVal = 0
10340   THAT_PROC = THIS_PROC
10350   That_Erl = Erl
10360   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10370   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

10410   With Me
10420     If .cmbAccounts.text = vbNullString Then
10430       ExpandCombo .cmbAccounts  ' ** Module Procedure: modUtilities.
10440       .cmbAccounts = .cmbAccounts.ItemData(0)
10450     End If
10460   End With

EXITP:
10470   Exit Sub

ERRH:
10480   THAT_PROC = THIS_PROC
10490   That_Erl = Erl
10500   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
10510   Resume EXITP

End Sub

Private Sub cmbAccounts_KeyDown(KeyCode As Integer, Shift As Integer)

10600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_KeyDown"

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
10690         .cmdPreview.SetFocus
10700       End With
10710     End Select
10720   End If

        ' ** Shift keys.
10730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10750       With Me
10760         intRetVal = 0
10770         .DateEnd.SetFocus
10780       End With
10790     End Select
10800   End If

        ' ** Ctrl-Shift keys.
10810   If intCtrlDown And (Not intAltDown) And intShiftDown Then
10820     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10830       With Me
10840         intRetVal = 0
10850         .chkRememberMe.SetFocus
10860       End With
10870     End Select
10880   End If

        ' ** Ctrl keys.
10890   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
10900     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
10910       With Me
10920         intRetVal = 0
10930         .chkPageOf.SetFocus
10940       End With
10950     End Select
10960   End If

EXITP:
10970   KeyCode = intRetVal
10980   Exit Sub

ERRH:
10990   intRetVal = 0
11000   THAT_PROC = THIS_PROC
11010   That_Erl = Erl
11020   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11030   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

11110   With Me
11120     Select Case .chkPageOf
          Case True
11130       .chkPageOf_lbl.FontBold = True
11140     Case False
11150       .chkPageOf_lbl.FontBold = False
11160     End Select
11170   End With

EXITP:
11180   Exit Sub

ERRH:
11190   Select Case ERR.Number
        Case Else
11200     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11210   End Select
11220   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

11300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

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
11390         .cmdPreview.SetFocus
11400       End With
11410     End Select
11420   End If

        ' ** Shift keys.
11430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11450       With Me
11460         intRetVal = 0
11470         .cmbAccounts.SetFocus
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

Private Sub cmdPreview_Click()

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

11610   If DoReport = True Then ' ** Function: Below.
11620     strDocName = "rptIncomeStatement"
11630     DoCmd.OpenReport strDocName, acViewPreview
11640     DoCmd.Maximize
11650     DoCmd.RunCommand acCmdFitToWindow
11660   End If

EXITP:
11670   Exit Sub

ERRH:
11680   Select Case ERR.Number
        Case Else
11690     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11700   End Select
11710   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

11800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
          Case vbKeyTab
11870       With Me
11880         intRetVal = 0
11890         .cmdPrint.SetFocus
11900       End With
11910     End Select
11920   End If

        ' ** Shift keys.
11930   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11940     Select Case intRetVal
          Case vbKeyTab
11950       With Me
11960         intRetVal = 0
11970         .cmbAccounts.SetFocus
11980       End With
11990     End Select
12000   End If

        ' ** Ctrl-Shift keys.
12010   If intCtrlDown And (Not intAltDown) And intShiftDown Then
12020     Select Case intRetVal
          Case vbKeyTab
12030       With Me
12040         intRetVal = 0
12050         .chkPageOf.SetFocus
12060       End With
12070     End Select
12080   End If

EXITP:
12090   KeyCode = intRetVal
12100   Exit Sub

ERRH:
12110   intRetVal = 0
12120   THAT_PROC = THIS_PROC
12130   That_Erl = Erl
12140   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12150   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

12200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

12210   If DoReport = True Then  ' ** Function: Below.
12220     strDocName = "rptIncomeStatement"
12230     If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
12240       DoCmd.OpenReport strDocName, acViewPreview
12250       DoCmd.Maximize
12260       DoCmd.RunCommand acCmdFitToWindow
12270     Else
            '##GTR_Ref: rptIncomeStatement
12280       DoCmd.OpenReport strDocName, acViewNormal
12290     End If
12300   End If

EXITP:
12310   Exit Sub

ERRH:
12320   Select Case ERR.Number
        Case Else
12330     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12340   End Select
12350   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

12400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12420   intShiftDown = (Shift And acShiftMask) > 0
12430   intAltDown = (Shift And acAltMask) > 0
12440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12460     Select Case intRetVal
          Case vbKeyTab
12470       With Me
12480         intRetVal = 0
12490         .cmdWord.SetFocus
12500       End With
12510     End Select
12520   End If

        ' ** Shift keys.
12530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12540     Select Case intRetVal
          Case vbKeyTab
12550       With Me
12560         intRetVal = 0
12570         .cmdPreview.SetFocus
12580       End With
12590     End Select
12600   End If

EXITP:
12610   KeyCode = intRetVal
12620   Exit Sub

ERRH:
12630   intRetVal = 0
12640   THAT_PROC = THIS_PROC
12650   That_Erl = Erl
12660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12670   Resume EXITP

End Sub

Private Sub cmdWord_Click()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

12710   With Me
12720     If DoReport = True Then ' ** Function: Below.

12730       If IsNull(.UserReportPath) = True Then
12740         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
12750       Else
12760         strRptPath = .UserReportPath
12770       End If
12780       strRptCap = "rptIncomeStatement_" & Format(Date, "yyyymmdd")

12790       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

12800       If strRptPathFile <> vbNullString Then
12810         strRpt = "rptIncomeStatement"
12820         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
12830         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
12840       End If

12850     End If
12860   End With

EXITP:
12870   Exit Sub

ERRH:
12880   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
12890   Case Else
12900     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12910   End Select
12920   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

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
          Case vbKeyTab
13070       With Me
13080         intRetVal = 0
13090         If .cmdExcel.Enabled = True Then
13100           .cmdExcel.SetFocus
13110         Else
13120           .cmdClose.SetFocus
13130         End If
13140       End With
13150     End Select
13160   End If

        ' ** Shift keys.
13170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13180     Select Case intRetVal
          Case vbKeyTab
13190       With Me
13200         intRetVal = 0
13210         .cmdPrint.SetFocus
13220       End With
13230     End Select
13240   End If

EXITP:
13250   KeyCode = intRetVal
13260   Exit Sub

ERRH:
13270   intRetVal = 0
13280   THAT_PROC = THIS_PROC
13290   That_Erl = Erl
13300   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13310   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

13410   With Me
13420     If DoReport = True Then  ' ** Function: Below.

13430       If IsNull(.UserReportPath) = True Then
13440         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
13450       Else
13460         strRptPath = .UserReportPath
13470       End If
13480       strRptCap = "rptIncomeStatement_" & Format(Date, "yyyymmdd")

13490       strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

13500       If strRptPathFile <> vbNullString Then
13510         gstrFormQuerySpec = THIS_NAME
13520         strQry = "qryIncomeStatement_09"
13530         DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
13540         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
13550       End If

13560     End If
13570   End With

EXITP:
13580   Exit Sub

ERRH:
13590   Select Case ERR.Number
        Case Else
13600     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
13610   End Select
13620   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

13700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

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
          Case vbKeyTab
13770       With Me
13780         intRetVal = 0
13790         .cmdClose.SetFocus
13800       End With
13810     End Select
13820   End If

        ' ** Shift keys.
13830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13840     Select Case intRetVal
          Case vbKeyTab
13850       With Me
13860         intRetVal = 0
13870         .cmdWord.SetFocus
13880       End With
13890     End Select
13900   End If

EXITP:
13910   KeyCode = intRetVal
13920   Exit Sub

ERRH:
13930   intRetVal = 0
13940   THAT_PROC = THIS_PROC
13950   That_Erl = Erl
13960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13970   Resume EXITP

End Sub

Private Function DoReport() As Boolean

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim blnRetVal As Boolean

14010   blnRetVal = True  ' ** Unless proven otherwise.

14020   With Me
          ' ** Validate main three data values.
14030     If IsNull(.DateStart.Value) Then
14040       blnRetVal = False
14050       MsgBox "You must enter a From date to continue.", vbInformation + vbOKOnly, "Entry Required"
14060       .DateStart.SetFocus
14070     Else
14080       If IsNull(.DateEnd.Value) Then
14090         blnRetVal = False
14100         MsgBox "You must enter a To date to continue.", vbInformation + vbOKOnly, "Entry Required"
14110         .DateEnd.SetFocus
14120       Else
14130         If .DateStart.Value >= .DateEnd.Value Then
14140           blnRetVal = False
14150           MsgBox "You must enter a From date that is less than the To date to continue.", vbInformation + vbOKOnly, "Invalid Date"
14160           .DateEnd.SetFocus
14170         Else
14180           If IsNull(.cmbAccounts.Column(0)) Or .cmbAccounts.Column(0) = vbNullString Then
14190             blnRetVal = False
14200             MsgBox "You must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
14210             .cmbAccounts.SetFocus
14220           Else
                  ' ** Something has changed, rebuild the temp table.
14230             If gdatStartDate <> .DateStart.Value Or gdatEndDate <> .DateEnd.Value Or gstrAccountNo <> .cmbAccounts.Column(0) Then
                    ' ** Set global variables for report headers.
14240               gdatStartDate = .DateStart.Value
14250               gdatEndDate = .DateEnd.Value
14260               gstrAccountNo = .cmbAccounts.Column(0)
14270               gstrAccountName = .cmbAccounts.Column(3)
14280               If IncomeExpense_BuildTable(gdatStartDate, gdatEndDate, gstrAccountNo, True) = False Then  ' ** Module Function: modIncExpFuncs.
14290                 blnRetVal = False
14300               End If
14310             End If
14320           End If
14330         End If
14340       End If
14350     End If
14360   End With

EXITP:
14370   DoReport = blnRetVal
14380   Exit Function

ERRH:
14390   blnRetVal = False
14400   THAT_PROC = THIS_PROC
14410   That_Erl = Erl
14420   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
14430   Resume EXITP

End Function
