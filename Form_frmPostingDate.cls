VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPostingDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmPostingDate"

'VGC 10/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Posting Date:     D {Posting_Date}
' **   Columnar:         M {opgInput_optColumnar}
' **   OK:               O {cmdOK}
' **   Classic:          S {opgInput_optClassic}
' **   Include Currency: U {chkIncludeCurrency}
' **   Exit:             X {cmdCancel}

' ** Shortcut F-Keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Timer:      T {Form_KeyDown}

' ** The RecordSource for this form is the PostingDate table.
'QRY: 'qryPostingDate_01' Posting Date  TYP: dbQSelect

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Our MonthCalendar Class.
Public clsMonthClass As clsMonthCal

Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean

Private blnFromCalendar As Boolean, blnKeyDown_Return As Boolean
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
150           blnFromCalendar = False: blnKeyDown_Return = False
160           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
170           gblnSwitchTo = False

180           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

190         Else
200           Cancel = -1
210         End If
220       Else
230         Cancel = -1
240       End If
250     End With

260     DoCmd.Hourglass False

270     If Cancel = -1 Then
280       DoCmd.Hourglass False
290       Beep
300       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
310     End If

EXITP:
320     Exit Sub

ERRH:
330     DoCmd.Hourglass False
340     THAT_PROC = THIS_PROC
350     That_Erl = Erl
360     Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
370     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnAdd As Boolean
        Dim strFilter As String

        ' ** Create an instance of the Calendar class.
410     Set clsMonthClass = New clsMonthCal
        ' ** You MUST SET the class hWndForm prop!!!
420     clsMonthClass.hWndForm = Me.hwnd
        ' ** Let's default to PositionAtCursor.
430     clsMonthClass.PositionAtCursor = True

440     If DCount("*", "PostingDate", "[Username] = '" & CurrentUser & "'") = 0 Then  ' ** Internal Access Function: Trust Accountant login.
450       blnAdd = False
460       Set dbs = CurrentDb
470       With dbs
            ' ** PostingDate table, Username = Null.
480         Set qdf = .QueryDefs("qryPostingDate_03")
490         Set rst = qdf.OpenRecordset
500         With rst
510           If .BOF = True And .EOF = True Then
                ' ** No Null record available.
520             blnAdd = True
530           Else
540             .MoveFirst
550             glngPostingDateID = ![PostingDate_ID]
560             .Edit
570             ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
580             .Update
590           End If
600           .Close
610         End With
620         Set rst = Nothing
630         Set qdf = Nothing
640         If blnAdd = True Then
              ' ** Append new record to PostingDate table, by specified [usr].
650           Set qdf = .QueryDefs("qryPostingDate_02")
660           With qdf.Parameters
670             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
680           End With
690           qdf.Execute dbFailOnError
700           Set qdf = Nothing
              ' ** PostingDate, by specified [usr].
710           Set qdf = .QueryDefs("qryPostingDate_07")
720           With qdf.Parameters
730             ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
740           End With
750           Set rst = qdf.OpenRecordset
760           With rst
770             .MoveFirst
780             glngPostingDateID = ![PostingDate_ID]
790             .Close
800           End With
810           Set rst = Nothing
820           Set qdf = Nothing
              ' ** Update tblCalendar_Staging, for unique_id, by specified [unqid].
830           Set qdf = .QueryDefs("qryCalendar_02")
840           With qdf.Parameters
850             ![unqid] = glngPostingDateID
860           End With
870           qdf.Execute
880           Set qdf = Nothing
890         End If
900         .Close
910       End With
920     End If

930     With Me

940       strFilter = "[Username] = '" & CurrentUser & "'"  ' ** Internal Access Function: Trust Accountant login.
950       .Filter = strFilter
960       .FilterOn = True

970       opgInput_AfterUpdate  ' ** Procedure: Below.
980       chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.

990       Select Case gblnHasForEx
          Case True
1000        .chkIncludeCurrency.Enabled = True
1010      Case False
1020        .chkIncludeCurrency.Enabled = False
1030      End Select

1040    End With

1050    DoCmd.Hourglass False

EXITP:
1060    Set rst = Nothing
1070    Set qdf = Nothing
1080    Set dbs = Nothing
1090    Exit Sub

ERRH:
1100    DoCmd.Hourglass False
1110    THAT_PROC = THIS_PROC
1120    That_Erl = Erl
1130    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1140    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Timer:      T {Form_KeyDown}

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyEscape
1270        With Me
1280          intRetVal = 0
1290          cmdCancel_Click  ' ** Procedure: Below.
1300        End With
1310      Case vbKeyF7
1320        With Me
1330          intRetVal = 0
1340          .cmdCalendar.SetFocus
1350          cmdCalendar_Click  ' ** Procedure: Below.
1360        End With
1370      End Select
1380    End If

        ' ** Alt keys.
1390    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1400      Select Case intRetVal
          Case vbKeyX
1410        intRetVal = 0
1420        cmdCancel_Click  ' ** Procedure: Below.
1430      End Select
1440    End If

        ' ** Ctrl keys.
1450    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1460      Select Case intRetVal
          Case vbKeyS
1470        intRetVal = 0
1480        cmdSave_Click  ' ** Procedure: Below.
1490      End Select
1500    End If

        ' ** Ctrl-Shift keys.
1510    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1520      Select Case intRetVal
          Case vbKeyF
1530        With Me
1540          intRetVal = 0
1550          .FocusHolder.SetFocus
1560        End With
1570      Case vbKeyT
1580        intRetVal = 0
1590        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
1600      End Select
1610    End If

EXITP:
1620    KeyCode = intRetVal
1630    Exit Sub

ERRH:
1640    intRetVal = 0
1650    THAT_PROC = THIS_PROC
1660    That_Erl = Erl
1670    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1680    Resume EXITP

End Sub

Private Sub Form_Timer()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1710    With Me
1720      .TimerInterval = 0&

1730    End With

EXITP:
1740    Exit Sub

ERRH:
1750    DoCmd.Hourglass False
1760    THAT_PROC = THIS_PROC
1770    That_Erl = Erl
1780    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
1790    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1810    If DataErr <> 0 Then
1820      Select Case DataErr
          Case 13  ' ** Type Mismatch (means nothing for the dates).
            ' ** Do nothing.
1830      Case 2110  ' ** Microsoft Access can't move the focus to the control |.
            ' ** Ignore.
1840        Response = acDataErrContinue
1850      Case 2113  ' ** The value you entered isn't valid for this field.
1860        Select Case Me.ActiveControl.Name
            Case "Posting_Date"
1870          If Me.Posting_Date.text <> "__/__/____" Then
1880            MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1890          Else
1900            Me.Posting_Date.text = vbNullString
1910          End If
1920        End Select
1930        Response = acDataErrContinue
1940      Case 2279  ' ** The value you entered isn't appropriate for the input mask '|' specified for this field.
1950        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Date"
1960        Response = acDataErrContinue
1970      Case Else
1980        If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1990        zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2000      End Select
2010    End If

EXITP:
2020    Exit Sub

ERRH:
2030    DoCmd.Hourglass False
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)
' ** This is required in case user Closes Form with the
' ** Calendar still open. It also handles when the
' ** user closes the application with the Calendar
' ** still open.
' ** VGC 04/13/08: Calendar window seems to be Modal, preventing the
' **               user from doing anything else till it's closed.

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2110    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

2120    If Not clsMonthClass Is Nothing Then
2130      If clsMonthClass.IsCalendar Then
2140        Cancel = -1
2150      Else
2160        Set clsMonthClass = Nothing
2170      End If
2180    End If

2190    DoCmd.Hourglass False

EXITP:
2200    Exit Sub

ERRH:
2210    DoCmd.Hourglass False
2220    Select Case ERR.Number
        Case Else
2230      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2240    End Select
2250    Resume EXITP

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
2400        .cmdCalendar_raised_focus_dots_img.Visible = False
2410        .cmdCalendar_raised_focus_img.Visible = False
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

Private Sub cmdOK_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim strDocName As String

2610    DoCmd.Hourglass True
2620    DoEvents

2630    With Me
2640      If .Dirty = True Then
2650        cmdSave_Click  ' ** Procedure: Below.
2660      End If
2670      If DateCheck_Post(.Posting_Date) = True Then  ' ** Module Function: modUtilities.
2680        Select Case .opgInput
            Case .opgInput_optClassic.OptionValue
2690          strDocName = "frmJournal"
2700        Case .opgInput_optColumnar.OptionValue
2710          JrnlCol_Load  ' ** Procedure: Below.
2720          strDocName = "frmJournal_Columns"
2730        End Select
2740        gblnSetFocus = True
2750        DoCmd.Close acForm, THIS_NAME
2760        DoCmd.OpenForm strDocName, , , , , , THIS_NAME
2770      Else
            ' ** I want to make sure they see that today's date is being used.
2780        If IsNull(.Posting_Date) = True Then
2790          .Posting_Date = CDate(Format(Date, "mm/dd/yyyy"))
2800          cmdSave_Click  ' ** Procedure: Below.
2810          DoEvents
2820          Select Case .opgInput
              Case .opgInput_optClassic.OptionValue
2830            strDocName = "frmJournal"
2840          Case .opgInput_optColumnar.OptionValue
2850            JrnlCol_Load  ' ** Procedure: Below.
2860            strDocName = "frmJournal_Columns"
2870          End Select
2880          gblnSetFocus = True
2890          DoCmd.Close acForm, THIS_NAME
2900          DoEvents
2910          DoCmd.OpenForm strDocName, , , , , , THIS_NAME
2920          DoEvents
2930        Else
2940          If Trim(.Posting_Date) = vbNullString Then
2950            .Posting_Date = CDate(Format(Date, "mm/dd/yyyy"))
2960            cmdSave_Click  ' ** Procedure: Below.
2970            DoEvents
2980            Select Case .opgInput
                Case .opgInput_optClassic.OptionValue
2990              strDocName = "frmJournal"
3000            Case .opgInput_optColumnar.OptionValue
3010              JrnlCol_Load  ' ** Procedure: Below.
3020              strDocName = "frmJournal_Columns"
3030            End Select
3040            gblnSetFocus = True
3050            DoCmd.Close acForm, THIS_NAME
3060            DoCmd.OpenForm strDocName, , , , , , THIS_NAME
3070          Else
3080            DoCmd.Hourglass False
3090          End If
3100        End If
3110      End If
3120    End With

EXITP:
3130    Exit Sub

ERRH:
3140    DoCmd.Hourglass False
3150    THAT_PROC = THIS_PROC
3160    That_Erl = Erl
3170    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3180    Resume EXITP

End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_MouseMove"

3210    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3220    Exit Sub

ERRH:
3230    Select Case ERR.Number
        Case Else
3240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3250    End Select
3260    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim strDocName As String

3310    With Me
3320      If gblnTimer = True Then
3330        Forms("frmLoadTimer").cmdStop_Click  ' ** Form Procedure: frmLoadTimer.
3340        gblnTimer = False
3350        FormClose "frmLoadTimer"  ' ** Module Function: modFileUtilities.
3360      End If
3370      glngJournalForm = 0&
3380      gstrFormQuerySpec = vbNullString
3390      gstrReturningForm = THIS_NAME
3400      .Posting_Date.ForeColor = CLR_WHT  ' ** Make it disappear.
3410      DoEvents
3420      strDocName = "frmMenu_Background"
3430      DoCmd.OpenForm strDocName, , , , , , THIS_NAME
3440      DoCmd.Close acForm, THIS_NAME
3450      DoCmd.Hourglass False
3460    End With

EXITP:
3470    Exit Sub

ERRH:
3480    DoCmd.Hourglass False
3490    THAT_PROC = THIS_PROC
3500    That_Erl = Erl
3510    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3520    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3610    DoCmd.RunCommand acCmdSelectRecord
3620    DoCmd.RunCommand acCmdSaveRecord

EXITP:
3630    Exit Sub

ERRH:
3640    Select Case ERR.Number
        Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
          ' ** Ignore! I don't know why it happens on some machines and not mine.
3650    Case Else
3660      DoCmd.Hourglass False
3670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3680    End Select
3690    Resume EXITP

End Sub

Private Sub Posting_Date_Enter()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Posting_Date_Enter"

3710    With Me
3720      If blnFromCalendar = False Then
3730        .Posting_Date = vbNullString
3740      Else
3750        blnFromCalendar = False
3760        .Posting_Date.SelLength = 0
3770        .Posting_Date.SelStart = 99
3780      End If
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    DoCmd.Hourglass False
3820    THAT_PROC = THIS_PROC
3830    That_Erl = Erl
3840    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3850    Resume EXITP

End Sub

Private Sub Posting_Date_BeforeUpdate(Cancel As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Posting_Date_BeforeUpdate"

3910    With Me
3920      If IsNull(.Posting_Date) = False Then
3930        If IsDate(.Posting_Date) = True Then
3940          If DateCheck_Post(.Posting_Date) = False Then  ' ** Module Function: modUtilities.
3950            DoCmd.Hourglass False
                ' ** MsgBox already displayed.
3960            Cancel = -1
3970          End If
3980        End If
3990      End If
4000    End With

EXITP:
4010    Exit Sub

ERRH:
4020    DoCmd.Hourglass False
4030    THAT_PROC = THIS_PROC
4040    That_Erl = Erl
4050    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4060    Resume EXITP

End Sub

Private Sub Posting_Date_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Posting_Date_KeyDown"

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
            ' ** Let it move normally.
4170      Case vbKeyReturn
4180        blnKeyDown_Return = True
4190      End Select
4200    End If

        ' ** Shift keys.
4210    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4220      Select Case intRetVal
          Case vbKeyTab
4230        With Me
4240          intRetVal = 0
4250          .cmdCancel.SetFocus
4260        End With
4270      Case vbKeyReturn
4280        blnKeyDown_Return = True
4290      End Select
4300    End If

EXITP:
4310    KeyCode = intRetVal
4320    Exit Sub

ERRH:
4330    intRetVal = 0
4340    THAT_PROC = THIS_PROC
4350    That_Erl = Erl
4360    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4370    Resume EXITP

End Sub

Private Sub Posting_Date_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "Posting_Date_DblClick"

4410    cmdOK_Click  ' ** Procedure: Above.

EXITP:
4420    Exit Sub

ERRH:
4430    DoCmd.Hourglass False
4440    THAT_PROC = THIS_PROC
4450    That_Erl = Erl
4460    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4470    Resume EXITP

End Sub

Private Sub Posting_Date_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Posting_Date_MouseMove"

4510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4520    Exit Sub

ERRH:
4530    Select Case ERR.Number
        Case Else
4540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4550    End Select
4560    Resume EXITP

End Sub

Private Sub Posting_Date_Exit(Cancel As Integer)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "Posting_Date_Exit"

4610    With Me
4620      If IsNull(.Posting_Date) = False Then
            ' ** Proceed.
4630      Else
4640        If blnKeyDown_Return = True Then
4650          blnKeyDown_Return = False
4660          .Posting_Date = Date
4670        Else
4680          If IsNull(.Posting_Date.text) = False Then
4690            If Trim(.Posting_Date.text) <> vbNullString Then
                  ' ** Proceed.
4700            Else
                  ' ** Empty, do nothing.
4710            End If
4720          Else
                ' ** Empty, do nothing.
4730          End If
4740        End If

4750      End If
4760    End With

EXITP:
4770    Exit Sub

ERRH:
4780    DoCmd.Hourglass False
4790    THAT_PROC = THIS_PROC
4800    That_Erl = Erl
4810    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4820    Resume EXITP

End Sub

Private Sub cmdCalendar_GotFocus()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_GotFocus"

4910    With Me
4920      blnCalendar1_Focus = True
4930      .cmdCalendar_raised_semifocus_dots_img.Visible = True
4940      .cmdCalendar_raised_img.Visible = False
4950      .cmdCalendar_raised_focus_img.Visible = False
4960      .cmdCalendar_raised_focus_dots_img.Visible = False
4970      .cmdCalendar_sunken_focus_dots_img.Visible = False
4980      .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseDown"

5110    With Me
5120      blnCalendar1_MouseDown = True
5130      .cmdCalendar_sunken_focus_dots_img.Visible = True
5140      .cmdCalendar_raised_img.Visible = False
5150      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5160      .cmdCalendar_raised_focus_img.Visible = False
5170      .cmdCalendar_raised_focus_dots_img.Visible = False
5180      .cmdCalendar_raised_img_dis.Visible = False
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

Private Sub cmdCalendar_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

5310    With Me
5320      datStartDate = Date
5330      datEndDate = 0
5340      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
5350      If blnRetVal = True Then
            ' ** Allow posting up to 1 month into the future.
5360        If datStartDate > DateAdd("m", 1, Date) Then
5370          MsgBox "Only future dates up to 1 month from today are allowed.", vbInformation + vbOKOnly, "Invalid Date"
5380          .Posting_Date = CDate(Format(Date, "mm/dd/yyyy"))
5390          cmdSave_Click  ' ** Procedure: Above.
5400        Else
5410          blnFromCalendar = True
5420          .Posting_Date = datStartDate
5430          cmdSave_Click  ' ** Procedure: Above.
5440        End If
5450      Else
5460        blnFromCalendar = True
5470        .Posting_Date = CDate(Format(Date, "mm/dd/yyyy"))
5480        cmdSave_Click  ' ** Procedure: Above.
5490      End If
5500      .Posting_Date.SetFocus
5510    End With

EXITP:
5520    Exit Sub

ERRH:
5530    DoCmd.Hourglass False
5540    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
5550    Case Else
5560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5570    End Select
5580    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseMove"

5610    With Me
5620      If blnCalendar1_MouseDown = False Then
5630        Select Case blnCalendar1_Focus
            Case True
5640          .cmdCalendar_raised_focus_dots_img.Visible = True
5650          .cmdCalendar_raised_focus_img.Visible = False
5660        Case False
5670          .cmdCalendar_raised_focus_img.Visible = True
5680          .cmdCalendar_raised_focus_dots_img.Visible = False
5690        End Select
5700        .cmdCalendar_raised_img.Visible = False
5710        .cmdCalendar_raised_semifocus_dots_img.Visible = False
5720        .cmdCalendar_sunken_focus_dots_img.Visible = False
5730        .cmdCalendar_raised_img_dis.Visible = False
5740      End If
5750    End With

EXITP:
5760    Exit Sub

ERRH:
5770    Select Case ERR.Number
        Case Else
5780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5790    End Select
5800    Resume EXITP

End Sub

Private Sub cmdCalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_MouseUp"

5910    With Me
5920      .cmdCalendar_raised_focus_dots_img.Visible = True
5930      .cmdCalendar_raised_img.Visible = False
5940      .cmdCalendar_raised_semifocus_dots_img.Visible = False
5950      .cmdCalendar_raised_focus_img.Visible = False
5960      .cmdCalendar_sunken_focus_dots_img.Visible = False
5970      .cmdCalendar_raised_img_dis.Visible = False
5980      blnCalendar1_MouseDown = False
5990    End With

EXITP:
6000    Exit Sub

ERRH:
6010    Select Case ERR.Number
        Case Else
6020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6030    End Select
6040    Resume EXITP

End Sub

Private Sub cmdCalendar_LostFocus()

6100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar_LostFocus"

6110    With Me
6120      .cmdCalendar_raised_img.Visible = True
6130      .cmdCalendar_raised_semifocus_dots_img.Visible = False
6140      .cmdCalendar_raised_focus_img.Visible = False
6150      .cmdCalendar_raised_focus_dots_img.Visible = False
6160      .cmdCalendar_sunken_focus_dots_img.Visible = False
6170      .cmdCalendar_raised_img_dis.Visible = False
6180      blnCalendar1_Focus = False
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

Private Sub opgInput_AfterUpdate()

6300  On Error GoTo ERRH

        Const THIS_PROC As String = "opgInput_AfterUpdate"

6310    With Me
6320      Select Case .opgInput
          Case .opgInput_optClassic.OptionValue
6330        .opgInput_optClassic_lbl.FontBold = True
6340        .opgInput_optColumnar_lbl.FontBold = False
6350      Case .opgInput_optColumnar.OptionValue
6360        .opgInput_optClassic_lbl.FontBold = False
6370        .opgInput_optColumnar_lbl.FontBold = True
6380      End Select
6390    End With

EXITP:
6400    Exit Sub

ERRH:
6410    DoCmd.Hourglass False
6420    THAT_PROC = THIS_PROC
6430    That_Erl = Erl
6440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6450    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

6510    With Me
6520      DoCmd.Hourglass True
6530      DoEvents
6540      Select Case .chkIncludeCurrency
          Case True
6550        .chkIncludeCurrency_lbl.FontBold = True
6560      Case False
6570        .chkIncludeCurrency_lbl.FontBold = False
6580      End Select
6590      DoEvents
          ' ** Save this immediately, so it's available when frmJournal opens.
6600      CurrencyPrefSave .chkIncludeCurrency  ' ** Procedure: Below.
6610      DoEvents
6620      DoCmd.Hourglass False
6630    End With

EXITP:
6640    Exit Sub

ERRH:
6650    DoCmd.Hourglass False
6660    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Ignore.
6670    Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6870        With Me
6880          intRetVal = 0
6890          .cmdOK.SetFocus
6900        End With
6910      End Select
6920    End If

        ' ** Shift keys.
6930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6950        With Me
6960          intRetVal = 0
6970          .opgInput.SetFocus
6980        End With
6990      End Select
7000    End If

EXITP:
7010    KeyCode = intRetVal
7020    Exit Sub

ERRH:
7030    intRetVal = 0
7040    THAT_PROC = THIS_PROC
7050    That_Erl = Erl
7060    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7070    Resume EXITP

End Sub

Private Sub JrnlCol_Load()

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "JrnlCol_Load"

        Dim blnRetVal As Boolean

7110    DoCmd.Hourglass True
7120    DoEvents

7130    blnRetVal = JrnlCol_Set  ' ** Module Function: modUtilities.

EXITP:
7140    Exit Sub

ERRH:
7150    DoCmd.Hourglass False
7160    THAT_PROC = THIS_PROC
7170    That_Erl = Erl
7180    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7190    Resume EXITP

End Sub

Private Sub CurrencyPrefSave(blnTmp01 As Boolean)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "CurrencyPrefSave"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

7210    With Me
7220      Set dbs = CurrentDb
7230      With dbs
            ' ** Update tblPreference_User, just 'chkIncludeCurrency' on 'frmPostingDate', by specified [usr], [pbln].
7240        Set qdf = .QueryDefs("qryPreferences_07_03")
7250        With qdf.Parameters
7260          ![usr] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
7270          ![pbln] = blnTmp01
7280        End With
7290        qdf.Execute
7300        .Close
7310      End With
7320    End With

EXITP:
7330    Set qdf = Nothing
7340    Set dbs = Nothing
7350    Exit Sub

ERRH:
7360    THAT_PROC = THIS_PROC
7370    That_Erl = Erl
7380    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
7390    Resume EXITP

End Sub
