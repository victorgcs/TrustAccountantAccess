VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CashControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CashControl"

'VGC 07/16/2017: CHANGES!

' ** Conditional Compiler Constants:
' ** NOTE: THESE CONSTANTS ARE NOT PUBLIC, ONLY PRIVATE!
#Const NoExcel = 0  ' ** 0 = Excel included; -1 = Excel excluded.
' ** Also in:

' ** For GTR, it just opens and stops.

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Excel:            E {cmdExcel}
' **   As Of:            F {DateAsOf}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Focus Holder:     F {FocusHolder}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

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

160           .DateAsOf = Null
170           DoEvents

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

      #If NoExcel Then
190           .cmdExcel.Enabled = False
      #Else
200           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberDates!
210           If IsNull(.DateAsOf) = False Then
220             DateAsOf_Exit Cancel  ' ** Procedure: Below.
230           End If

240           chkRememberDates_AfterUpdate  ' ** Procedure: Below.
250           chkPageOf_AfterUpdate  ' ** Procedure: Below.

260           If IsNull(.DateAsOf) = True Then
270             .DateAsOf = TodaysDate  ' ** Module Function: modUtilities.
280           End If

290         Else
300           Cancel = -1
310         End If
320       Else
330         Cancel = -1
340       End If
350     End With

360     If Cancel = -1 Then
370       Beep
380       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
390     End If

400     DoCmd.Hourglass False

EXITP:
410     Exit Sub

ERRH:
420     DoCmd.Hourglass False
430     Select Case ERR.Number
        Case Else
440       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
450     End Select
460     Resume EXITP

End Sub

Private Sub Form_Load()

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

510     With Me

          ' ** Create an instance of the Calendar class.
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

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Shortcut F-Keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Focus Holder:     F {FocusHolder}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdClose_Click  ' ** Procedure: Below.
790       Case vbKeyF7
800         With Me
810           intRetVal = 0
820           If .cmdCalendar.Enabled = True Then
830             .cmdCalendar.SetFocus
840             cmdCalendar_Click  ' ** Procedure: Below.
850           End If
860         End With
870       End Select
880     End If

        ' ** Alt keys.
890     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
900       Select Case intRetVal
          Case vbKeyX
910         intRetVal = 0
920         cmdClose_Click  ' ** Procedure: Below.
930       End Select
940     End If

        ' ** Ctrl keys.
950     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyG
970         With Me
980           intRetVal = 0
990           If .chkPageOf.Enabled = True Then
1000            .chkPageOf.SetFocus
1010            .chkPageOf = (Not .chkPageOf)
1020            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1030          End If
1040        End With
1050      End Select
1060    End If

        ' ** Ctrl-Shift keys.
1070    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1080      Select Case intRetVal
          Case vbKeyF
1090        With Me
1100          intRetVal = 0
1110          .FocusHolder.SetFocus
1120          .DateAsOf = Null
1130        End With
1140      End Select
1150    End If

EXITP:
1160    KeyCode = intRetVal
1170    Exit Sub

ERRH:
1180    intRetVal = 0
1190    THAT_PROC = THIS_PROC
1200    That_Erl = Erl
1210    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1220    Resume EXITP

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
1740      If Left(Me.ActiveControl.Name, 4) = "Date" Then  ' ** DateAsOf.
1750        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1760        Response = acDataErrContinue
1770      End If
1780    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
1790      Beep
1800      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
1810      Response = acDataErrContinue
1820    Case Else
1830      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1840      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select

EXITP:
1860    Exit Sub

ERRH:
1870    Select Case ERR.Number
        Case Else
1880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1890    End Select
1900    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2010    DoCmd.Hourglass True
2020    DoEvents

2030    With Me
2040      If .chkRememberDates = False Then
2050        .DateAsOf = Null  ' ** So it won't be there when Pref_Save() runs.
2060        DoEvents
2070      End If
2080    End With

2090    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
2100    Do While Reports.Count > 0
2110      DoCmd.Close acReport, Reports(0).Name
2120    Loop

2130    gstrFormQuerySpec = vbNullString
2140    gblnSetFocus = True
2150    gstrReturningForm = THIS_NAME
2160    Select Case strCallingForm
        Case "frmMenu_Report"
2170      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2180    End Select

EXITP:
2190    Exit Sub

ERRH:
2200    Select Case ERR.Number
        Case Else
2210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2220    End Select
2230    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2310    With Me
2320      If .cmdCalendar_raised_focus_dots_img.Visible = True Or .cmdCalendar_raised_focus_img.Visible = True Then
2330        Select Case blnCalendar1_Focus
            Case True
2340          .cmdCalendar_raised_semifocus_dots_img.Visible = True
2350          .cmdCalendar_raised_img.Visible = False
2360        Case False
2370          .cmdCalendar_raised_img.Visible = True
2380          .cmdCalendar_raised_semifocus_dots_img.Visible = False
2390        End Select
2400        .cmdCalendar_raised_focus_img.Visible = False
2410        .cmdCalendar_raised_focus_dots_img.Visible = False
2420        .cmdCalendar_sunken_focus_dots_img.Visible = False
2430        .cmdCalendar_raised_img_dis.Visible = False
2440      End If
2450    End With

EXITP:
2460    Exit Sub

ERRH:
2470    Select Case ERR.Number
        Case Else
2480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2490    End Select
2500    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2610    DoCmd.Hourglass True
2620    DoEvents

2630    DoCmd.Close acForm, THIS_NAME

EXITP:
2640    Exit Sub

ERRH:
2650    Select Case ERR.Number
        Case Else
2660      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2670    End Select
2680    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab
2770        With Me
2780          intRetVal = 0
2790          .DateAsOf.SetFocus
2800        End With
2810      End Select
2820    End If

        ' ** Shift keys.
2830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2840      Select Case intRetVal
          Case vbKeyTab
2850        With Me
2860          intRetVal = 0
2870          If .cmdExcel.Enabled = True Then
2880            .cmdExcel.SetFocus
2890          Else
2900            .cmdWord.SetFocus
2910          End If
2920        End With
2930      End Select
2940    End If

EXITP:
2950    KeyCode = intRetVal
2960    Exit Sub

ERRH:
2970    intRetVal = 0
2980    THAT_PROC = THIS_PROC
2990    That_Erl = Erl
3000    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3010    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

3110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub DateAsOf_KeyDown(KeyCode As Integer, Shift As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3220    intShiftDown = (Shift And acShiftMask) > 0
3230    intAltDown = (Shift And acAltMask) > 0
3240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3260      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3270        With Me
3280          intRetVal = 0
3290          .cmdPreview.SetFocus
3300        End With
3310      End Select
3320    End If

        ' ** Shift keys.
3330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3350        With Me
3360          intRetVal = 0
3370          .cmdClose.SetFocus
3380        End With
3390      End Select
3400    End If

        ' ** Ctrl keys.
3410    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
3420      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3430        With Me
3440          intRetVal = 0
3450          .cmdCalendar.SetFocus
3460        End With
3470      End Select
3480    End If

EXITP:
3490    KeyCode = intRetVal
3500    Exit Sub

ERRH:
3510    intRetVal = 0
3520    THAT_PROC = THIS_PROC
3530    That_Erl = Erl
3540    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3550    Resume EXITP

End Sub

Private Sub DateAsOf_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_MouseMove"

3610    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub DateAsOf_Exit(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "DateAsOf_Exit"

3710    With Me
3720      If IsNull(.DateAsOf) = True Then
            ' ** Populate it with today's date.
3730        .DateAsOf = Date
3740      Else
3750        If Trim(.DateAsOf) = vbNullString Then
              ' ** Populate it with today's date.
3760          .DateAsOf = Date
3770        Else
3780          If .DateAsOf < Date Then
3790            MsgBox "Data is not rolled back to the entered date." & vbCrLf & _
                  "The report reflects current information.", vbInformation + vbOKOnly, "Current Data Only"
3800          ElseIf .DateAsOf > Date Then
3810            MsgBox "Report will reflect current information.", vbInformation + vbOKOnly, "Current Data Only"
3820          End If
3830        End If
3840      End If
3850    End With

EXITP:
3860    Exit Sub

ERRH:
3870    THAT_PROC = THIS_PROC
3880    That_Erl = Erl
3890    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3900    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

4010    With Me
4020      blnCalendar1_Focus = True
4030      .cmdCalendar_raised_semifocus_dots_img.Visible = True
4040      .cmdCalendar_raised_img.Visible = False
4050      .cmdCalendar_raised_focus_img.Visible = False
4060      .cmdCalendar_raised_focus_dots_img.Visible = False
4070      .cmdCalendar_sunken_focus_dots_img.Visible = False
4080      .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

4210    With Me
4220      blnCalendar1_MouseDown = True
4230      .cmdCalendar_sunken_focus_dots_img.Visible = True
4240      .cmdCalendar_raised_img.Visible = False
4250      .cmdCalendar_raised_semifocus_dots_img.Visible = False
4260      .cmdCalendar_raised_focus_img.Visible = False
4270      .cmdCalendar_raised_focus_dots_img.Visible = False
4280      .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_Click()

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

4410    With Me
4420      datStartDate = Date
4430      datEndDate = 0
4440      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
4450      If blnRetVal = True Then
4460        .DateAsOf = datStartDate
4470      Else
4480        .DateAsOf = CDate(Format(Date, "mm/dd/yyyy"))
4490      End If
4500      .DateAsOf.SetFocus
4510    End With

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub cmdCalendar_KeyDown(KeyCode As Integer, Shift As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4620    intShiftDown = (Shift And acShiftMask) > 0
4630    intAltDown = (Shift And acAltMask) > 0
4640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4660      Select Case intRetVal
          Case vbKeyTab
4670        With Me
4680          intRetVal = 0
4690          .chkRememberDates.SetFocus
4700        End With
4710      End Select
4720    End If

        ' ** Shift keys.
4730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4740      Select Case intRetVal
          Case vbKeyTab
4750        With Me
4760          intRetVal = 0
4770          .DateAsOf.SetFocus
4780        End With
4790      End Select
4800    End If

EXITP:
4810    KeyCode = intRetVal
4820    Exit Sub

ERRH:
4830    intRetVal = 0
4840    Select Case ERR.Number
        Case Else
4850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4860    End Select
4870    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

4910    With Me
4920      If blnCalendar1_MouseDown = False Then
4930        Select Case blnCalendar1_Focus
            Case True
4940          .cmdCalendar_raised_focus_dots_img.Visible = True
4950          .cmdCalendar_raised_focus_img.Visible = False
4960        Case False
4970          .cmdCalendar_raised_focus_img.Visible = True
4980          .cmdCalendar_raised_focus_dots_img.Visible = False
4990        End Select
5000        .cmdCalendar_raised_img.Visible = False
5010        .cmdCalendar_raised_semifocus_dots_img.Visible = False
5020        .cmdCalendar_sunken_focus_dots_img.Visible = False
5030        .cmdCalendar_raised_img_dis.Visible = False
5040      End If
5050    End With

EXITP:
5060    Exit Sub

ERRH:
5070    Select Case ERR.Number
        Case Else
5080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5090    End Select
5100    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

5210    With Me
5220      .cmdCalendar_raised_focus_dots_img.Visible = True
5230      .cmdCalendar_raised_img.Visible = False
5240      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5250      .cmdCalendar_raised_focus_img.Visible = False
5260      .cmdCalendar_sunken_focus_dots_img.Visible = False
5270      .cmdCalendar_raised_img_dis.Visible = False
5280      blnCalendar1_MouseDown = False
5290    End With

EXITP:
5300    Exit Sub

ERRH:
5310    Select Case ERR.Number
        Case Else
5320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5330    End Select
5340    Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

5410    With Me
5420      .cmdCalendar_raised_img.Visible = True
5430      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5440      .cmdCalendar_raised_focus_img.Visible = False
5450      .cmdCalendar_raised_focus_dots_img.Visible = False
5460      .cmdCalendar_sunken_focus_dots_img.Visible = False
5470      .cmdCalendar_raised_img_dis.Visible = False
5480      blnCalendar1_Focus = False
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

Private Sub chkRememberDates_AfterUpdate()

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

5610    With Me
5620      Select Case .chkRememberDates
          Case True
5630        .chkRememberDates_lbl.FontBold = True
5640      Case False
5650        .chkRememberDates_lbl.FontBold = False
5660      End Select
5670    End With

EXITP:
5680    Exit Sub

ERRH:
5690    THAT_PROC = THIS_PROC
5700    That_Erl = Erl
5710    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5720    Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

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
          Case vbKeyTab
5870        With Me
5880          intRetVal = 0
5890          .chkPageOf.SetFocus
5900        End With
5910      End Select
5920    End If

        ' ** Shift keys.
5930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5940      Select Case intRetVal
          Case vbKeyTab
5950        With Me
5960          intRetVal = 0
5970          .cmdCalendar.SetFocus
5980        End With
5990      End Select
6000    End If

EXITP:
6010    KeyCode = intRetVal
6020    Exit Sub

ERRH:
6030    intRetVal = 0
6040    THAT_PROC = THIS_PROC
6050    That_Erl = Erl
6060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6070    Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

6110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
6120    Exit Sub

ERRH:
6130    Select Case ERR.Number
        Case Else
6140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6150    End Select
6160    Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

6210    With Me
6220      Select Case .chkPageOf
          Case True
6230        .chkPageOf_lbl.FontBold = True
6240      Case False
6250        .chkPageOf_lbl.FontBold = False
6260      End Select
6270    End With

EXITP:
6280    Exit Sub

ERRH:
6290    Select Case ERR.Number
        Case Else
6300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6310    End Select
6320    Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6420    intShiftDown = (Shift And acShiftMask) > 0
6430    intAltDown = (Shift And acAltMask) > 0
6440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6460      Select Case intRetVal
          Case vbKeyTab
6470        With Me
6480          intRetVal = 0
6490          .cmdPreview.SetFocus
6500        End With
6510      End Select
6520    End If

        ' ** Shift keys.
6530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6540      Select Case intRetVal
          Case vbKeyTab
6550        With Me
6560          intRetVal = 0
6570          .chkRememberDates.SetFocus
6580        End With
6590      End Select
6600    End If

EXITP:
6610    KeyCode = intRetVal
6620    Exit Sub

ERRH:
6630    intRetVal = 0
6640    THAT_PROC = THIS_PROC
6650    That_Erl = Erl
6660    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6670    Resume EXITP

End Sub

Private Sub cmdPreview_Click()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

        Dim strDocName As String

6710    If DoReport = True Then  ' ** Function: Below.
6720      strDocName = "rptCashControl"
6730      DoCmd.OpenReport strDocName, acViewPreview
6740      DoCmd.Maximize
6750      DoCmd.RunCommand acCmdFitToWindow
6760    End If

EXITP:
6770    Exit Sub

ERRH:
6780    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
6790    Case Else
6800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6810    End Select
6820    Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

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
          Case vbKeyTab
6970        With Me
6980          intRetVal = 0
6990          .cmdPrint.SetFocus
7000        End With
7010      End Select
7020    End If

        ' ** Shift keys.
7030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7040      Select Case intRetVal
          Case vbKeyTab
7050        With Me
7060          intRetVal = 0
7070          .DateAsOf.SetFocus
7080        End With
7090      End Select
7100    End If

        ' ** Ctrl-Shift keys.
7110    If intCtrlDown And (Not intAltDown) And intShiftDown Then
7120      Select Case intRetVal
          Case vbKeyTab
7130        With Me
7140          intRetVal = 0
7150          .chkPageOf.SetFocus
7160        End With
7170      End Select
7180    End If

EXITP:
7190    KeyCode = intRetVal
7200    Exit Sub

ERRH:
7210    intRetVal = 0
7220    THAT_PROC = THIS_PROC
7230    That_Erl = Erl
7240    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7250    Resume EXITP

End Sub

Private Sub cmdPrint_Click()

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

        Dim strDocName As String

7310    If DoReport = True Then  ' ** Function: Below
7320      strDocName = "rptCashControl"
7330      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
7340        DoCmd.OpenReport strDocName, acViewPreview
7350        DoCmd.Maximize
7360        DoCmd.RunCommand acCmdFitToWindow
7370      Else
            '##GTR_Ref: rptCashControl
7380        DoCmd.OpenReport strDocName, acViewNormal
7390      End If
7400    End If

EXITP:
7410    Exit Sub

ERRH:
7420    Select Case ERR.Number
        Case 2501  ' ** The '|' action was Canceled.
          ' ** Do nothing.
7430    Case Else
7440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7450    End Select
7460    Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

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
          Case vbKeyTab
7570        With Me
7580          intRetVal = 0
7590          .cmdWord.SetFocus
7600        End With
7610      End Select
7620    End If

        ' ** Shift keys.
7630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7640      Select Case intRetVal
          Case vbKeyTab
7650        With Me
7660          intRetVal = 0
7670          .cmdPreview.SetFocus
7680        End With
7690      End Select
7700    End If

EXITP:
7710    KeyCode = intRetVal
7720    Exit Sub

ERRH:
7730    intRetVal = 0
7740    Select Case ERR.Number
        Case Else
7750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7760    End Select
7770    Resume EXITP

End Sub

Private Sub cmdWord_Click()

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

7810    With Me
7820      If DoReport = True Then  ' ** Function: Below.

7830        If IsNull(.UserReportPath) = True Then
7840          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
7850        Else
7860          strRptPath = .UserReportPath
7870        End If
7880        strRptCap = "rptCashControl_" & Format(Date, "yyyymmdd")

7890        strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

7900        If strRptPathFile <> vbNullString Then
7910          strRpt = "rptCashControl"
7920          DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
7930          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
7940        End If

7950      End If
7960    End With

EXITP:

7970    Exit Sub

ERRH:
7980    Select Case ERR.Number
        Case Else
7990      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8000    End Select
8010    Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8120    intShiftDown = (Shift And acShiftMask) > 0
8130    intAltDown = (Shift And acAltMask) > 0
8140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8160      Select Case intRetVal
          Case vbKeyTab
8170        With Me
8180          intRetVal = 0
8190          If .cmdExcel.Enabled = True Then
8200            .cmdExcel.SetFocus
8210          Else
8220            .cmdClose.SetFocus
8230          End If
8240        End With
8250      End Select
8260    End If

        ' ** Shift keys.
8270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8280      Select Case intRetVal
          Case vbKeyTab
8290        With Me
8300          intRetVal = 0
8310          .cmdPrint.SetFocus
8320        End With
8330      End Select
8340    End If

EXITP:
8350    KeyCode = intRetVal
8360    Exit Sub

ERRH:
8370    intRetVal = 0
8380    Select Case ERR.Number
        Case Else
8390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8400    End Select
8410    Resume EXITP

End Sub

Private Sub cmdExcel_Click()

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim strQry As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String

8510    With Me
8520      If DoReport = True Then  ' ** Function: Below.

8530        If IsNull(.UserReportPath) = True Then
8540          strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
8550        Else
8560          strRptPath = .UserReportPath
8570        End If
8580        strRptCap = "rptCashControl_" & Format(Date, "yyyymmdd")

8590        strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

8600        If strRptPathFile <> vbNullString Then
8610          gstrFormQuerySpec = THIS_NAME
8620          strQry = "qryCashControl_04"
8630          DoCmd.OutputTo acOutputQuery, strQry, acFormatXLS, strRptPathFile, True
8640          .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
8650        End If

8660      End If
8670    End With

EXITP:
8680    Exit Sub

ERRH:
8690    Select Case ERR.Number
        Case Else
8700      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8710    End Select
8720    Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

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
          Case vbKeyTab
8870        With Me
8880          intRetVal = 0
8890          .cmdClose.SetFocus
8900        End With
8910      End Select
8920    End If

        ' ** Shift keys.
8930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8940      Select Case intRetVal
          Case vbKeyTab
8950        With Me
8960          intRetVal = 0
8970          .cmdWord.SetFocus
8980        End With
8990      End Select
9000    End If

EXITP:
9010    KeyCode = intRetVal
9020    Exit Sub

ERRH:
9030    intRetVal = 0
9040    THAT_PROC = THIS_PROC
9050    That_Erl = Erl
9060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9070    Resume EXITP

End Sub

Private Function DoReport() As Boolean

9100  On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnRetVal As Boolean

9110    blnRetVal = True

9120    With Me
9130      If IsNull(.DateAsOf) = True Then
9140        blnRetVal = False
9150        MsgBox "You must enter a date to continue.", vbInformation + vbOKOnly, "Entry Required"
9160      End If
9170    End With

9180    If blnRetVal = True Then

9190      Set dbs = CurrentDb
9200      With dbs
            ' ** Account, just needed fields.
9210        Set qdf = .QueryDefs("qryCashControl_01")
9220        Set rst = qdf.OpenRecordset
9230        With rst
9240          If .BOF = True And .EOF = True Then
9250            blnRetVal = False
9260            MsgBox "There is no data for this report.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
9270          End If
9280          .Close
9290        End With
9300        .Close
9310      End With

9320    End If

EXITP:
9330    Set rst = Nothing
9340    Set qdf = Nothing
9350    Set dbs = Nothing
9360    DoReport = blnRetVal
9370    Exit Function

ERRH:
9380    blnRetVal = False
9390    THAT_PROC = THIS_PROC
9400    That_Erl = Erl
9410    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9420    Resume EXITP

End Function
