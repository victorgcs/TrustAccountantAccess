VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Maturity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Maturity"

'VGC 07/16/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   From:             F {DateStart}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   To:               T {DateEnd}
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

170           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
180           .cmdExcel.Enabled = False
      #Else
190           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberDates!
200           Select Case IsNull(.DateStart)
              Case True
210             .DateStart = Date
220           Case False
230             DateStart_Exit Cancel  ' ** Procedure: Below.
240             DateEnd_Exit Cancel  ' ** Procedure: Below.
250           End Select

260           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
270           chkPageOf_AfterUpdate  ' ** Procedure: Below.

280         Else
290           Cancel = -1
300         End If
310       Else
320         Cancel = -1
330       End If
340     End With

350     If Cancel = -1 Then
360       Beep
370       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
380     End If

        'DoCmd.Hourglass False

EXITP:
390     Exit Sub

ERRH:
400     DoCmd.Hourglass False
410     Select Case ERR.Number
        Case Else
420       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
430     End Select
440     Resume EXITP

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
1170    Select Case ERR.Number
        Case Else
1180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1190    End Select
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
        Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1720      If Left(Me.ActiveControl.Name, 4) = "Date" Then
1730        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1740        Response = acDataErrContinue
1750      End If
1760    Case Else
1770      zErrorHandler THIS_NAME, "Undefined Form Error", DataErr  ' ** Module Function: modErrorHandler.
1780    End Select

EXITP:
1790    Exit Sub

ERRH:
1800    Select Case ERR.Number
        Case Else
1810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1820    End Select
1830    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1910    DoCmd.Hourglass True
1920    DoEvents

1930    With Me
1940      If .chkRememberDates = False Then
1950        .DateStart = Null  ' ** So it won't be there when Pref_Save() runs.
1960        .DateEnd = Null
1970      End If
1980    End With

1990    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2000    Do While Reports.Count > 0
2010      DoCmd.Close acReport, Reports(0).Name
2020    Loop

2030    gstrFormQuerySpec = vbNullString
2040    gblnSetFocus = True
2050    gstrReturningForm = THIS_NAME
2060    Select Case strCallingForm
        Case "frmMenu_Report"
2070      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2080    End Select

EXITP:
2090    Exit Sub

ERRH:
2100    Select Case ERR.Number
        Case Else
2110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2120    End Select
2130    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

2210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2310    With Me
2320      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
2330        Select Case blnCalendar1_Focus
            Case True
2340          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
2350          .cmdCalendar1_raised_img.Visible = False
2360        Case False
2370          .cmdCalendar1_raised_img.Visible = True
2380          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
2390        End Select
2400        .cmdCalendar1_raised_focus_dots_img.Visible = False
2410        .cmdCalendar1_raised_focus_img.Visible = False
2420        .cmdCalendar1_sunken_focus_dots_img.Visible = False
2430        .cmdCalendar1_raised_img_dis.Visible = False
2440      End If
2450      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
2460        Select Case blnCalendar2_Focus
            Case True
2470          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
2480          .cmdCalendar2_raised_img.Visible = False
2490        Case False
2500          .cmdCalendar2_raised_img.Visible = True
2510          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
2520        End Select
2530        .cmdCalendar2_raised_focus_dots_img.Visible = False
2540        .cmdCalendar2_raised_focus_img.Visible = False
2550        .cmdCalendar2_sunken_focus_dots_img.Visible = False
2560        .cmdCalendar2_raised_img_dis.Visible = False
2570      End If
2580    End With

EXITP:
2590    Exit Sub

ERRH:
2600    Select Case ERR.Number
        Case Else
2610      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2620    End Select
2630    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2720    Exit Sub

ERRH:
2730    Select Case ERR.Number
        Case Else
2740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2750    End Select
2760    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2810    DoCmd.Hourglass True
2820    DoEvents

2830    DoCmd.Close acForm, THIS_NAME

EXITP:
2840    Exit Sub

ERRH:
2850    Select Case ERR.Number
        Case Else
2860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2870    End Select
2880    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2920    intShiftDown = (Shift And acShiftMask) > 0
2930    intAltDown = (Shift And acAltMask) > 0
2940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2960      Select Case intRetVal
          Case vbKeyTab
2970        With Me
2980          intRetVal = 0
2990          .DateStart.SetFocus
3000        End With
3010      End Select
3020    End If

        ' ** Shift keys.
3030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3040      Select Case intRetVal
          Case vbKeyTab
3050        With Me
3060          intRetVal = 0
3070          If .cmdExcel.Enabled = True Then
3080            .cmdExcel.SetFocus
3090          Else
3100            .cmdWord.SetFocus
3110          End If
3120        End With
3130      End Select
3140    End If

EXITP:
3150    KeyCode = intRetVal
3160    Exit Sub

ERRH:
3170    intRetVal = 0
3180    Select Case ERR.Number
        Case Else
3190      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3200    End Select
3210    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

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

Private Sub DateStart_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3420    intShiftDown = (Shift And acShiftMask) > 0
3430    intAltDown = (Shift And acAltMask) > 0
3440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3470        With Me
3480          intRetVal = 0
3490          .DateEnd.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3550        With Me
3560          intRetVal = 0
3570          .cmdClose.SetFocus
3580        End With
3590      End Select
3600    End If

        ' ** Ctrl keys.
3610    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3630        With Me
3640          intRetVal = 0
3650          .cmdCalendar1.SetFocus
3660        End With
3670      End Select
3680    End If

EXITP:
3690    KeyCode = intRetVal
3700    Exit Sub

ERRH:
3710    intRetVal = 0
3720    Select Case ERR.Number
        Case Else
3730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3740    End Select
3750    Resume EXITP

End Sub

Private Sub DateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_MouseMove"

3810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub DateStart_Exit(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "DateStart_Exit"

3910    With Me
3920      If IsDate(.DateEnd) = True And IsDate(.DateStart) = True Then
3930        If CDate(.DateStart) > CDate(.DateEnd) Then
3940          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
3950          .DateEnd = Null
3960          .DateEnd.SetFocus
3970        End If
3980      End If
3990    End With

EXITP:
4000    Exit Sub

ERRH:
4010    Select Case ERR.Number
        Case Else
4020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4030    End Select
4040    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

4110    With Me
4120      blnCalendar1_Focus = True
4130      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4140      .cmdCalendar1_raised_img.Visible = False
4150      .cmdCalendar1_raised_focus_img.Visible = False
4160      .cmdCalendar1_raised_focus_dots_img.Visible = False
4170      .cmdCalendar1_sunken_focus_dots_img.Visible = False
4180      .cmdCalendar1_raised_img_dis.Visible = False
4190    End With

EXITP:
4200    Exit Sub

ERRH:
4210    Select Case ERR.Number
        Case Else
4220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4230    End Select
4240    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

4310    With Me
4320      blnCalendar1_MouseDown = True
4330      .cmdCalendar1_sunken_focus_dots_img.Visible = True
4340      .cmdCalendar1_raised_img.Visible = False
4350      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
4360      .cmdCalendar1_raised_focus_img.Visible = False
4370      .cmdCalendar1_raised_focus_dots_img.Visible = False
4380      .cmdCalendar1_raised_img_dis.Visible = False
4390    End With

EXITP:
4400    Exit Sub

ERRH:
4410    Select Case ERR.Number
        Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Sub

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

4510    With Me
4520      datStartDate = Date
4530      datEndDate = 0
4540      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
4550      If blnRetVal = True Then
4560        .DateStart = datStartDate
4570      Else
4580        .DateStart = CDate(Format(Date, "mm/dd/yyyy"))
4590      End If
4600      .DateStart.SetFocus
4610    End With

EXITP:
4620    Exit Sub

ERRH:
4630    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
4640    Case Else
4650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4660    End Select
4670    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

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
4870          .DateStart.SetFocus
4880        End With
4890      End Select
4900    End If

EXITP:
4910    KeyCode = intRetVal
4920    Exit Sub

ERRH:
4930    intRetVal = 0
4940    Select Case ERR.Number
        Case Else
4950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4960    End Select
4970    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

5010    With Me
5020      If blnCalendar1_MouseDown = False Then
5030        Select Case blnCalendar1_Focus
            Case True
5040          .cmdCalendar1_raised_focus_dots_img.Visible = True
5050          .cmdCalendar1_raised_focus_img.Visible = False
5060        Case False
5070          .cmdCalendar1_raised_focus_img.Visible = True
5080          .cmdCalendar1_raised_focus_dots_img.Visible = False
5090        End Select
5100        .cmdCalendar1_raised_img.Visible = False
5110        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5120        .cmdCalendar1_sunken_focus_dots_img.Visible = False
5130        .cmdCalendar1_raised_img_dis.Visible = False
5140      End If
5150    End With

EXITP:
5160    Exit Sub

ERRH:
5170    Select Case ERR.Number
        Case Else
5180      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5190    End Select
5200    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

5310    With Me
5320      .cmdCalendar1_raised_focus_dots_img.Visible = True
5330      .cmdCalendar1_raised_img.Visible = False
5340      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5350      .cmdCalendar1_raised_focus_img.Visible = False
5360      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5370      .cmdCalendar1_raised_img_dis.Visible = False
5380      blnCalendar1_MouseDown = False
5390    End With

EXITP:
5400    Exit Sub

ERRH:
5410    Select Case ERR.Number
        Case Else
5420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5430    End Select
5440    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

5510    With Me
5520      .cmdCalendar1_raised_img.Visible = True
5530      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
5540      .cmdCalendar1_raised_focus_img.Visible = False
5550      .cmdCalendar1_raised_focus_dots_img.Visible = False
5560      .cmdCalendar1_sunken_focus_dots_img.Visible = False
5570      .cmdCalendar1_raised_img_dis.Visible = False
5580      blnCalendar1_Focus = False
5590    End With

EXITP:
5600    Exit Sub

ERRH:
5610    Select Case ERR.Number
        Case Else
5620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5630    End Select
5640    Resume EXITP

End Sub

Private Sub DateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5720    intShiftDown = (Shift And acShiftMask) > 0
5730    intAltDown = (Shift And acAltMask) > 0
5740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5770        With Me
5780          intRetVal = 0
5790          .cmdPreview.SetFocus
5800        End With
5810      End Select
5820    End If

        ' ** Shift keys.
5830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5840      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5850        With Me
5860          intRetVal = 0
5870          .DateStart.SetFocus
5880        End With
5890      End Select
5900    End If

        ' ** Ctrl keys.
5910    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
5920      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5930        With Me
5940          intRetVal = 0
5950          .cmdCalendar2.SetFocus
5960        End With
5970      End Select
5980    End If

        ' ** Ctrl-Shift keys.
5990    If intCtrlDown And (Not intAltDown) And intShiftDown Then
6000      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6010        With Me
6020          intRetVal = 0
6030          .cmdCalendar1.SetFocus
6040        End With
6050      End Select
6060    End If

EXITP:
6070    KeyCode = intRetVal
6080    Exit Sub

ERRH:
6090    intRetVal = 0
6100    Select Case ERR.Number
        Case Else
6110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6120    End Select
6130    Resume EXITP

End Sub

Private Sub DateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_MouseMove"

6210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
6220    Exit Sub

ERRH:
6230    Select Case ERR.Number
        Case Else
6240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6250    End Select
6260    Resume EXITP

End Sub

Private Sub DateEnd_Exit(Cancel As Integer)

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "DateEnd_Exit"

6310    With Me
6320      If IsNull(.DateEnd) Then
            ' ** Populate it with today's date.
6330        .DateEnd = Date
6340      Else
6350        If IsNull(.DateStart) = False Then
6360          If CDate(.DateStart) > CDate(.DateEnd) Then
6370            MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
6380            .DateStart = Null
6390            .DateStart.SetFocus
6400          End If
6410        End If
6420      End If
6430    End With

EXITP:
6440    Exit Sub

ERRH:
6450    Select Case ERR.Number
        Case Else
6460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6470    End Select
6480    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

6510    With Me
6520      blnCalendar2_Focus = True
6530      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
6540      .cmdCalendar2_raised_img.Visible = False
6550      .cmdCalendar2_raised_focus_img.Visible = False
6560      .cmdCalendar2_raised_focus_dots_img.Visible = False
6570      .cmdCalendar2_sunken_focus_dots_img.Visible = False
6580      .cmdCalendar2_raised_img_dis.Visible = False
6590    End With

EXITP:
6600    Exit Sub

ERRH:
6610    Select Case ERR.Number
        Case Else
6620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6630    End Select
6640    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

6710    With Me
6720      blnCalendar2_MouseDown = True
6730      .cmdCalendar2_sunken_focus_dots_img.Visible = True
6740      .cmdCalendar2_raised_img.Visible = False
6750      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
6760      .cmdCalendar2_raised_focus_img.Visible = False
6770      .cmdCalendar2_raised_focus_dots_img.Visible = False
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

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

6910    With Me
6920      datStartDate = Date
6930      datEndDate = 0
6940      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
6950      If blnRetVal = True Then
6960        .DateEnd = datStartDate
6970      Else
6980        .DateEnd = CDate(Format(Date, "mm/dd/yyyy"))
6990      End If
7000      .DateEnd.SetFocus
7010      Cancel = 0
7020      DateEnd_Exit Cancel  ' ** Procedure: Above.
7030      If Cancel = 0 Then
7040        .cmdPreview.SetFocus
7050      End If
7060    End With

EXITP:
7070    Exit Sub

ERRH:
7080    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7090    Case Else
7100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7110    End Select
7120    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7220    intShiftDown = (Shift And acShiftMask) > 0
7230    intAltDown = (Shift And acAltMask) > 0
7240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7270        With Me
7280          intRetVal = 0
7290          .chkRememberDates.SetFocus
7300        End With
7310      End Select
7320    End If

        ' ** Shift keys.
7330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
7350        With Me
7360          intRetVal = 0
7370          .DateEnd.SetFocus
7380        End With
7390      End Select
7400    End If

EXITP:
7410    KeyCode = intRetVal
7420    Exit Sub

ERRH:
7430    intRetVal = 0
7440    Select Case ERR.Number
        Case Else
7450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7460    End Select
7470    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

7510    With Me
7520      If blnCalendar2_MouseDown = False Then
7530        Select Case blnCalendar2_Focus
            Case True
7540          .cmdCalendar2_raised_focus_dots_img.Visible = True
7550          .cmdCalendar2_raised_focus_img.Visible = False
7560        Case False
7570          .cmdCalendar2_raised_focus_img.Visible = True
7580          .cmdCalendar2_raised_focus_dots_img.Visible = False
7590        End Select
7600        .cmdCalendar2_raised_img.Visible = False
7610        .cmdCalendar2_raised_semifocus_dots_img.Visible = False
7620        .cmdCalendar2_sunken_focus_dots_img.Visible = False
7630        .cmdCalendar2_raised_img_dis.Visible = False
7640      End If
7650    End With

EXITP:
7660    Exit Sub

ERRH:
7670    Select Case ERR.Number
        Case Else
7680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7690    End Select
7700    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

7810    With Me
7820      .cmdCalendar2_raised_focus_dots_img.Visible = True
7830      .cmdCalendar2_raised_img.Visible = False
7840      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
7850      .cmdCalendar2_raised_focus_img.Visible = False
7860      .cmdCalendar2_sunken_focus_dots_img.Visible = False
7870      .cmdCalendar2_raised_img_dis.Visible = False
7880      blnCalendar2_MouseDown = False
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

Private Sub cmdCalendar2_LostFocus()

8000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

8010    With Me
8020      .cmdCalendar2_raised_img.Visible = True
8030      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
8040      .cmdCalendar2_raised_focus_img.Visible = False
8050      .cmdCalendar2_raised_focus_dots_img.Visible = False
8060      .cmdCalendar2_sunken_focus_dots_img.Visible = False
8070      .cmdCalendar2_raised_img_dis.Visible = False
8080      blnCalendar2_Focus = False
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

Private Sub chkRememberDates_AfterUpdate()

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

8210    With Me
8220      Select Case .chkRememberDates
          Case True
8230        .chkRememberDates_lbl.FontBold = True
8240      Case False
8250        .chkRememberDates_lbl.FontBold = False
8260      End Select
8270    End With

EXITP:
8280    Exit Sub

ERRH:
8290    Select Case ERR.Number
        Case Else
8300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8310    End Select
8320    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

8400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8420    intShiftDown = (Shift And acShiftMask) > 0
8430    intAltDown = (Shift And acAltMask) > 0
8440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8470        With Me
8480          intRetVal = 0
8490          .chkPageOf.SetFocus
8500        End With
8510      End Select
8520    End If

        ' ** Shift keys.
8530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8550        With Me
8560          intRetVal = 0
8570          .cmdCalendar2.SetFocus
8580        End With
8590      End Select
8600    End If

EXITP:
8610    KeyCode = intRetVal
8620    Exit Sub

ERRH:
8630    intRetVal = 0
8640    Select Case ERR.Number
        Case Else
8650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8660    End Select
8670    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

8710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
8720    Exit Sub

ERRH:
8730    Select Case ERR.Number
        Case Else
8740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8750    End Select
8760    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

8810    With Me
8820      Select Case .chkPageOf
          Case True
8830        .chkPageOf_lbl.FontBold = True
8840      Case False
8850        .chkPageOf_lbl.FontBold = False
8860      End Select
8870    End With

EXITP:
8880    Exit Sub

ERRH:
8890    DoCmd.Hourglass False
8900    Select Case ERR.Number
        Case Else
8910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8920    End Select
8930    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

9000  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

9010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
9020    intShiftDown = (Shift And acShiftMask) > 0
9030    intAltDown = (Shift And acAltMask) > 0
9040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
9050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
9060      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9070        With Me
9080          intRetVal = 0
9090          .cmdPreview.SetFocus
9100        End With
9110      End Select
9120    End If

        ' ** Shift keys.
9130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
9150        With Me
9160          intRetVal = 0
9170          .chkRememberDates.SetFocus
9180        End With
9190      End Select
9200    End If

EXITP:
9210    KeyCode = intRetVal
9220    Exit Sub

ERRH:
9230    intRetVal = 0
9240    Select Case ERR.Number
        Case Else
9250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9260    End Select
9270    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

9300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strLinkCriteria As String
        Dim strDocName As String

9310    If DoReport = True Then  ' ** Function: Below.
9320      strDocName = "rptMaturity"
9330      strLinkCriteria = "(((masterasset.due) Is Not Null And " & _
            "(masterasset.due) Between #" & Format(Me.DateStart, "mm/yy/yyyy") & "# And #" & Format(Me.DateEnd, "mm/dd/yyyy") & "#))"
9340      DoCmd.OpenReport strDocName, acViewPreview, , strLinkCriteria
9350      DoCmd.Maximize
9360      DoCmd.RunCommand acCmdFitToWindow
9370    End If

EXITP:
9380    Exit Sub

ERRH:
9390    Select Case ERR.Number
        Case Else
9400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9410    End Select
9420    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

9500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
          Case vbKeyTab
9570        With Me
9580          intRetVal = 0
9590          .cmdPrint.SetFocus
9600        End With
9610      End Select
9620    End If

        ' ** Shift keys.
9630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
9640      Select Case intRetVal
          Case vbKeyTab
9650        With Me
9660          intRetVal = 0
9670          .DateEnd.SetFocus
9680        End With
9690      End Select
9700    End If

        ' ** Ctrl_Shift keys.
9710    If intCtrlDown And (Not intAltDown) And intShiftDown Then
9720      Select Case intRetVal
          Case vbKeyTab
9730        With Me
9740          intRetVal = 0
9750          .chkPageOf.SetFocus
9760        End With
9770      End Select
9780    End If

EXITP:
9790    KeyCode = intRetVal
9800    Exit Sub

ERRH:
9810    intRetVal = 0
9820    Select Case ERR.Number
        Case Else
9830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9840    End Select
9850    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strLinkCriteria As String
        Dim strDocName As String

9910    If DoReport = True Then  ' ** Function: Below.
9920      strDocName = "rptMaturity"
9930      strLinkCriteria = "(((masterasset.due) Is Not Null And " & _
            "(masterasset.due) Between #" & Format(Me.DateStart, "mm/yy/yyyy") & "# And #" & Format(Me.DateEnd, "mm/dd/yyyy") & "#))"
9940      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
9950        DoCmd.OpenReport strDocName, acViewPreview, , strLinkCriteria
9960        DoCmd.Maximize
9970        DoCmd.RunCommand acCmdFitToWindow
9980      Else
            '##GTR_Ref: rptMaturity
9990        DoCmd.OpenReport strDocName, acViewNormal, , strLinkCriteria
10000     End If
10010   End If

EXITP:
10020   Exit Sub

ERRH:
10030   Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
10040   Case Else
10050     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10060   End Select
10070   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

10100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

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
          Case vbKeyTab
10170       With Me
10180         intRetVal = 0
10190         .cmdWord.SetFocus
10200       End With
10210     End Select
10220   End If

        ' ** Shift keys.
10230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10240     Select Case intRetVal
          Case vbKeyTab
10250       With Me
10260         intRetVal = 0
10270         .cmdPreview.SetFocus
10280       End With
10290     End Select
10300   End If

EXITP:
10310   KeyCode = intRetVal
10320   Exit Sub

ERRH:
10330   intRetVal = 0
10340   Select Case ERR.Number
        Case Else
10350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10360   End Select
10370   Resume EXITP

End Sub

Private Sub cmdWord_Click()

10400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

10410   With Me
10420     If DoReport = True Then  ' ** Procedure: Below.

10430       If IsNull(.UserReportPath) = True Then
10440         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
10450       Else
10460         strRptPath = .UserReportPath
10470       End If

10480       strRpt = "rptMaturity"
10490       gdatStartDate = .DateStart
10500       gdatEndDate = .DateEnd
10510       strRptCap = "Security_Maturity_" & _
              Format$(gdatStartDate, "yymmdd") & "_To_" & Format$(gdatEndDate, "yymmdd")

10520       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

10530       If strRptPathFile <> vbNullString Then
10540         gblnMessage = True  ' ** Indicating is this going to Word.
10550         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
10560         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
10570       End If

10580     End If
10590     gdatStartDate = 0
10600     gdatEndDate = 0
10610     gblnMessage = False
10620   End With

EXITP:
10630   Exit Sub

ERRH:
10640   Select Case ERR.Number
        Case Else
10650     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10660   End Select
10670   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

10710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
10720   intShiftDown = (Shift And acShiftMask) > 0
10730   intAltDown = (Shift And acAltMask) > 0
10740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
10750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
10760     Select Case intRetVal
          Case vbKeyTab
10770       With Me
10780         intRetVal = 0
10790         If .cmdExcel.Enabled = True Then
10800           .cmdExcel.SetFocus
10810         Else
10820           .cmdClose.SetFocus
10830         End If
10840       End With
10850     End Select
10860   End If

        ' ** Shift keys.
10870   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10880     Select Case intRetVal
          Case vbKeyTab
10890       With Me
10900         intRetVal = 0
10910         .cmdPrint.SetFocus
10920       End With
10930     End Select
10940   End If

EXITP:
10950   KeyCode = intRetVal
10960   Exit Sub

ERRH:
10970   intRetVal = 0
10980   Select Case ERR.Number
        Case Else
10990     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11000   End Select
11010   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

11100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim msgResponse As VbMsgBoxResult
        Dim blnContinue As Boolean

11110   With Me

11120     DoCmd.Hourglass True
11130     DoEvents

11140     blnContinue = True

11150     If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
11160       DoCmd.Hourglass False
11170       msgResponse = MsgBox("Microsoft Excel is currently open." & vbCrLf & vbCrLf & _
              "In order for Trust Accountant to reliably export your report," & vbCrLf & _
              "Microsoft Excel must be closed." & vbCrLf & vbCrLf & _
              "You may close Excel before proceding, then click Retry." & vbCrLf & _
              "Click Cancel to export your report later.", vbExclamation + vbRetryCancel, "Excel Is Open")
            ' ** ... Otherwise Trust Accountant will do it for you.
11180       If msgResponse <> vbRetry Then
11190         blnContinue = False
11200       End If
11210     End If

11220     If blnContinue = True Then

11230       DoCmd.Hourglass True
11240       DoEvents

11250       If DoReport = True Then  ' ** Procedure: Below.

11260         blnContinue = True

11270         If IsNull(.UserReportPath) = True Then
11280           strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
11290         Else
11300           strRptPath = .UserReportPath
11310         End If

11320         strRptCap = "Security_Maturity_" & Format$(.DateStart, "yymmdd") & "_To_" & Format$(.DateEnd, "yymmdd")

              ' ** qryRpt_Maturity_16 (xx), For Export.
11330         strQry = "qryRpt_Maturity_17"

11340         strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

11350         If strRptPathFile <> vbNullString Then
11360           If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
11370             EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
11380           End If
11390           DoEvents
11400           If FileExists(strRptPathFile) = True Then  ' ** Module Function: modFileUtilities.
11410             Kill strRptPathFile
11420           End If
11430           DoEvents
11440           gstrFormQuerySpec = THIS_NAME
11450           DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, False
11460           If Excel_NameOnly(strRptPathFile, "Maturity") = True Then  ' ** Module Function: modExcelFuncs.
11470             If EXE_IsRunning("EXCEL.EXE") = True Then  ' ** Module Function: modProcessFuncs.
11480               EXE_Terminate "EXCEL.EXE"  ' ** Module Function: modProcessFuncs.
11490             End If
11500             OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
11510           End If
11520           .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
11530         End If

11540       End If  ' ** DoReport().
11550     End If  ' ** blnContinue.

11560     DoCmd.Hourglass False

11570   End With

EXITP:
11580   Exit Sub

ERRH:
11590   DoCmd.Hourglass False
11600   Select Case ERR.Number
        Case Else
11610     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11620   End Select
11630   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

11700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11710   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11720   intShiftDown = (Shift And acShiftMask) > 0
11730   intAltDown = (Shift And acAltMask) > 0
11740   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11750   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11760     Select Case intRetVal
          Case vbKeyTab
11770       With Me
11780         intRetVal = 0
11790         .cmdClose.SetFocus
11800       End With
11810     End Select
11820   End If

        ' ** Shift keys.
11830   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11840     Select Case intRetVal
          Case vbKeyTab
11850       With Me
11860         intRetVal = 0
11870         .cmdWord.SetFocus
11880       End With
11890     End Select
11900   End If

EXITP:
11910   KeyCode = intRetVal
11920   Exit Sub

ERRH:
11930   intRetVal = 0
11940   Select Case ERR.Number
        Case Else
11950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
11960   End Select
11970   Resume EXITP

End Sub

Private Function DoReport() As Boolean

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

12010   blnRetVal = True

12020   With Me
12030     If IsNull(.DateStart) = True Then
12040       blnRetVal = False
12050       MsgBox "You must enter a From date in the format MM/DD/YYYY", vbInformation + vbOKOnly, "Entry Required"
12060       .DateStart.SetFocus
12070     Else
12080       If .DateStart = vbNullString Then
12090         blnRetVal = False
12100         MsgBox "You must enter a From date in the format MM/DD/YYYY", vbInformation + vbOKOnly, "Entry Required"
12110         .DateStart.SetFocus
12120       Else
12130         If IsNull(.DateEnd) = True Then
12140           blnRetVal = False
12150           MsgBox "You must enter a To date in the format MM/DD/YYYY", vbInformation + vbOKOnly, "Entry Required"
12160           .DateEnd.SetFocus
12170         Else
12180           If .DateEnd = vbNullString Then
12190             blnRetVal = False
12200             MsgBox "You must enter a To date in the format MM/DD/YYYY", vbInformation + vbOKOnly, "Entry Required"
12210             .DateEnd.SetFocus
12220           Else
                  ' ** Check to verify the end date is later than the start date.
12230             If CDate(.DateStart) > CDate(.DateEnd) Then
12240               blnRetVal = False
12250               MsgBox "You must enter a From date that is less than the To date to continue.", vbInformation + vbOKOnly, "Invalid Date"
12260               .DateEnd.SetFocus
12270             End If
12280           End If
12290         End If
12300       End If
12310     End If
12320   End With

12330   If blnRetVal = True Then
12340     Set dbs = CurrentDb
12350     With dbs
            ' ** ActiveAssets, grouped and summed, just Due <> Null, by specified [datbeg], [datend].
12360       Set qdf = .QueryDefs("qryRpt_Maturity_02")
12370       With qdf.Parameters
12380         ![datbeg] = CDate(Me.DateStart)
12390         ![datEnd] = CDate(Me.DateEnd)
12400       End With
12410       Set rst = qdf.OpenRecordset
12420       With rst
12430         If .BOF = True And .EOF = True Then
12440           blnRetVal = False
12450           MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
12460         End If
12470         .Close
12480       End With
12490       .Close
12500     End With
12510   End If

EXITP:
12520   Set rst = Nothing
12530   Set qdf = Nothing
12540   Set dbs = Nothing
12550   DoReport = blnRetVal
12560   Exit Function

ERRH:
12570   blnRetVal = False
12580   Select Case ERR.Number
        Case Else
12590     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12600   End Select
12610   Resume EXITP

End Function
