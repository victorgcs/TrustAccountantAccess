VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_TransactionsByType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_TransactionsByType"

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
' **   From:             F {TransDateStart}
' **   Include Archive:  H {chkIncludeArchive}
' **   All:              L {chkTransDateAll}
' **   Name:             M {opgAccountSource_optName}
' **   Select None:      N {cmdSelectNone}
' **   Word:             O {cmdWord}
' **   Print:            P {cmdPrint}
' **   Select All:       S {cmdSelectAll}
' **   To:               T {TransDateEnd}
' **   Preview:          W {cmdPreview}
' **   Exit:             X {cmdClose}

' ** Shortcut F-keys responsive from this form:
' **   Date Picker:      F7 {cmdCalendar1}
' **   Date Picker:      F8 {cmdCalendar2}

' ** Shortcut Ctrl keys responsive from this form:
' **   Page Count:       G {chkPageOf}

' ** Shortcut Ctrl-Alt keys responsive from this form:
' **   Paid:             A {JType_08_chk}
' **   Liability (-):    B {JType_11_chk}
' **   Dividend:         D {JType_01_chk}
' **   Interest:         I {JType_02_chk}
' **   Cost Adj.:        J {JType_12_chk}
' **   Liability (+):    L {JType_10_chk}
' **   Misc.:            M {JType_07_chk}
' **   Deposit:          O {JType_04_chk}
' **   Purchase:         P {JType_03_chk}
' **   Received:         R {JType_09_chk}
' **   Sold:             S {JType_05_chk}
' **   Withdrawn:        W {JType_06_chk}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Focus Holder:     F {FocusHolder}

' ** Copyright Lebans Holdings 1999 Ltd.
' ** Create an instance of Calendar Class.
Public clsMonthClass As clsMonthCal

' ** Array: arr_varJType().
Private lngJTypes As Long, arr_varJType As Variant
Private Const JT_TYPE As Integer = 0
Private Const JT_DESC As Integer = 1
'Private Const JT_SORT As Integer = 2
Private Const JT_CTL  As Integer = 3
Private Const JT_LBL  As Integer = 4
Private Const JT_VAL  As Integer = 5
Private Const JT_LTR  As Integer = 6
Private Const JT_BOX  As Integer = 7

Private blnSelectAll_Focus As Boolean, blnSelectAll_MouseDown As Boolean
Private blnSelectNone_Focus As Boolean, blnSelectNone_MouseDown As Boolean
Private blnCalendar1_Focus As Boolean, blnCalendar1_MouseDown As Boolean
Private blnCalendar2_Focus As Boolean, blnCalendar2_MouseDown As Boolean

Private strCallingForm As String
Private strFileName As String, strDocName As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset, ctl As Access.Control
        Dim lngX As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           gdatStartDate = 0: gdatEndDate = 0
160           gvarCrtRpt_FL_SpecData = Null
170           gstrFormQuerySpec = THIS_NAME
180           gstrReportQuerySpec = vbNullString
190           blnSelectAll_Focus = False: blnSelectAll_MouseDown = False
200           blnSelectNone_Focus = False: blnSelectNone_MouseDown = False
210           blnCalendar1_Focus = False: blnCalendar1_MouseDown = False
220           blnCalendar2_Focus = False: blnCalendar2_MouseDown = False

230           ArchiveSet  ' ** Procedure: Below.

240           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
250           DoEvents

260           Set dbs = CurrentDb
270           With dbs
280             Set qdf = .QueryDefs("qryTransactions_ByType_05")
290             Set rst = qdf.OpenRecordset
300             With rst
310               .MoveLast
320               lngJTypes = .RecordCount
330               .MoveFirst
340               arr_varJType = .GetRows(lngJTypes)
                  ' ************************************************************
                  ' ** Array: arr_varJType()
                  ' **
                  ' **   Field  Element  Name                       Constant
                  ' **   =====  =======  =========================  ==========
                  ' **     1       0     JournalType                JT_TYPE
                  ' **     2       1     JournalType_Description    JT_DESC
                  ' **     3       2     JournalType_Order          JT_SORT
                  ' **     4       3     ctl_name                   JT_CTL
                  ' **     5       4     ctl_label                  JT_LBL
                  ' **     6       5     ctl_value                  JT_VAL
                  ' **     7       6     ctl_letter                 JT_LTR
                  ' **     8       7     ctl_box                    JT_BOX
                  ' **
                  ' ************************************************************
350               .Close
360             End With
370             .Close
380           End With
390           DoEvents

400           For lngX = 0& To (lngJTypes - 1&)
410             Select Case arr_varJType(JT_TYPE, lngX)
                Case "Paid"
420               arr_varJType(JT_LTR, lngX) = "A"
430             Case "Liability (-)"
440               arr_varJType(JT_LTR, lngX) = "B"
450             Case "Dividend"
460               arr_varJType(JT_LTR, lngX) = "D"
470             Case "Interest"
480               arr_varJType(JT_LTR, lngX) = "I"
490             Case "Cost Adj."
500               arr_varJType(JT_LTR, lngX) = "J"
510             Case "Liability (+)"
520               arr_varJType(JT_LTR, lngX) = "L"
530             Case "Misc."
540               arr_varJType(JT_LTR, lngX) = "M"
550             Case "Deposit"
560               arr_varJType(JT_LTR, lngX) = "O"
570             Case "Purchase"
580               arr_varJType(JT_LTR, lngX) = "P"
590             Case "Received"
600               arr_varJType(JT_LTR, lngX) = "R"
610             Case "Sold"
620               arr_varJType(JT_LTR, lngX) = "S"
630             Case "Withdrawn"
640               arr_varJType(JT_LTR, lngX) = "W"
650             End Select
660           Next

670           For lngX = 0& To (lngJTypes - 1&)
                ' ** JType_01_lbl, JType_01_chk
680             Set ctl = .Controls("JType_" & Right("00" & CStr(lngX + 1), 2) & "_chk")
690             ctl.Visible = True
700             arr_varJType(JT_CTL, lngX) = ctl.Name
710             ctl.Tag = arr_varJType(JT_TYPE, lngX)
720             ctl.StatusBarText = "Journal Type: " & arr_varJType(JT_TYPE, lngX) & " - Ctrl+Alt+" & arr_varJType(JT_LTR, lngX)
730             ctl.ControlTipText = arr_varJType(JT_TYPE, lngX) & vbCrLf & "- Ctrl+Alt+" & arr_varJType(JT_LTR, lngX)
740             Set ctl = .Controls("JType_" & Right("00" & CStr(lngX + 1), 2) & "_lbl")
750             ctl.Caption = arr_varJType(JT_TYPE, lngX)
760             arr_varJType(JT_LBL, lngX) = ctl.Name
770             arr_varJType(JT_BOX, lngX) = ctl.Name & "_box"
780             Set ctl = Nothing
790             JType_Chk (arr_varJType(JT_CTL, lngX) & "_AfterUpdate") ' ** Procedure: Below.
800           Next
810           DoEvents

      #If NoExcel Then
820           .cmdExcel.Enabled = False
      #Else
830           .cmdExcel.Enabled = True
      #End If

              ' ** EVENT CHECK: chkRememberMe!
              ' ** NONE!
              ' ** EVENT CHECK: chkRememberDates!
840           If IsNull(.TransDateStart) = False And IsNull(.TransDateEnd) = False Then
850             TransDateStart_Exit Cancel  ' ** Procedure: Below.
860             TransDateEnd_Exit Cancel  ' ** Procedure: Below.
870           End If
              ' ** EVENT CHECK: chkRememberTypes!
880           JType_01_chk_AfterUpdate  ' ** Procedure: Below.
890           JType_02_chk_AfterUpdate  ' ** Procedure: Below.
900           JType_03_chk_AfterUpdate  ' ** Procedure: Below.
910           JType_04_chk_AfterUpdate  ' ** Procedure: Below.
920           JType_05_chk_AfterUpdate  ' ** Procedure: Below.
930           JType_06_chk_AfterUpdate  ' ** Procedure: Below.
940           JType_07_chk_AfterUpdate  ' ** Procedure: Below.
950           JType_08_chk_AfterUpdate  ' ** Procedure: Below.
960           JType_09_chk_AfterUpdate  ' ** Procedure: Below.
970           JType_10_chk_AfterUpdate  ' ** Procedure: Below.
980           JType_11_chk_AfterUpdate  ' ** Procedure: Below.
990           JType_12_chk_AfterUpdate  ' ** Procedure: Below.

1000          chkTransDateAll_AfterUpdate  ' ** Procedure: Below.
1010          opgAccountSource_AfterUpdate  ' ** Procedure: Below.
1020          chkRememberDates_AfterUpdate  ' ** Procedure: Below.
1030          chkRememberMe_AfterUpdate  ' ** Procedure: Below.
1040          chkRememberTypes_AfterUpdate  ' ** Procedure: Below.
1050          chkIncludeArchive_AfterUpdate  ' ** Procedure: Below.
1060          chkPageOf_AfterUpdate  ' ** Procedure: Below.

1070          Select Case .TransDateStart.Enabled
              Case True
1080            .TransDateStart.SetFocus
1090          Case False
1100            .cmbAccounts.SetFocus
1110          End Select

1120        Else
1130          Cancel = -1
1140        End If
1150      Else
1160        Cancel = -1
1170      End If
1180    End With  ' ** Me.

1190    If Cancel = -1 Then
1200      Beep
1210      MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
1220    End If

1230    DoCmd.Hourglass False

EXITP:
1240    Set ctl = Nothing
1250    Set rst = Nothing
1260    Set qdf = Nothing
1270    Set dbs = Nothing
1280    Exit Sub

ERRH:
1290    DoCmd.Hourglass False
1300    Select Case ERR.Number
        Case Else
1310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1320    End Select
1330    Resume EXITP

End Sub

Private Sub Form_Load()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

1410    With Me

          ' ** Copyright Lebans Holdings 1999 Ltd.
          ' ** Create an instance of the Calendar class
1420      Set clsMonthClass = New clsMonthCal
          ' ** You MUST SET the class hWndForm prop!!!
1430      clsMonthClass.hWndForm = Me.hwnd
          ' ** Let's default to PositionAtCursor.
1440      clsMonthClass.PositionAtCursor = True

1450      .UserReportPath = Pref_ReportPath(.UserReportPath, THIS_NAME)  ' ** Module Function: modPreferenceFuncs.

1460    End With

EXITP:
1470    Exit Sub

ERRH:
1480    DoCmd.Hourglass False
1490    Select Case ERR.Number
        Case Else
1500      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1510    End Select
1520    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim lngX As Long
        Dim intRetVal As Integer

1610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1620    intShiftDown = (Shift And acShiftMask) > 0
1630    intAltDown = (Shift And acAltMask) > 0
1640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut F-keys to other forms and functionality:
        ' **   Date Picker:      F7 {cmdCalendar1}
        ' **   Date Picker:      F8 {cmdCalendar2}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Page Count:       G {chkPageOf}

        ' ** Shortcut Ctrl-Alt keys to other forms and functionality:
        ' **   Paid:             A {JType_08_chk}
        ' **   Liability (-):    B {JType_11_chk}
        ' **   Dividend:         D {JType_01_chk}
        ' **   Interest:         I {JType_02_chk}
        ' **   Cost Adj.:        J {JType_12_chk}
        ' **   Liability (+):    L {JType_10_chk}
        ' **   Misc.:            M {JType_07_chk}
        ' **   Deposit:          O {JType_04_chk}
        ' **   Purchase:         P {JType_03_chk}
        ' **   Received:         R {JType_09_chk}
        ' **   Sold:             S {JType_05_chk}
        ' **   Withdrawn:        W {JType_06_chk}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Focus Holder:     F {FocusHolder}

        ' ** Plain keys.
1650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyEscape
1670        intRetVal = 0
1680        cmdClose_Click  ' ** Procedure: Below.
1690      Case vbKeyF7
1700        intRetVal = 0
1710        cmdCalendar1_Click  ' ** Procedure: Below.
1720      Case vbKeyF8
1730        intRetVal = 0
1740        cmdCalendar2_Click  ' ** Procedure: Below.
1750      End Select
1760    End If

        ' ** Alt keys.
1770    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1780      Select Case intRetVal
          Case vbKeyX
1790        intRetVal = 0
1800        cmdClose_Click  ' ** Procedure: Below.
1810      End Select
1820    End If

        ' ** Ctrl keys.
1830    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1840      Select Case intRetVal
          Case vbKeyG
1850        With Me
1860          intRetVal = 0
1870          If .chkPageOf.Enabled = True Then
1880            .chkPageOf.SetFocus
1890            .chkPageOf = (Not .chkPageOf)
1900            chkPageOf_AfterUpdate  ' ** Procedure: Below.
1910          End If
1920        End With
1930      End Select
1940    End If

        ' ** Ctrl-Alt keys.
1950    If intCtrlDown And intAltDown And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyA  ' ** Paid.
1970        With Me
1980          intRetVal = 0
1990          For lngX = 0& To (lngJTypes - 1&)
2000            If arr_varJType(JT_LTR, lngX) = "A" Then
2010              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2020              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2030              Exit For
2040            End If
2050          Next
2060        End With
2070      Case vbKeyB  ' ** Liability (-).
2080        With Me
2090          intRetVal = 0
2100          For lngX = 0& To (lngJTypes - 1&)
2110            If arr_varJType(JT_LTR, lngX) = "B" Then
2120              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2130              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2140              Exit For
2150            End If
2160          Next
2170        End With
2180      Case vbKeyD  ' ** Dividend.
2190        With Me
2200          intRetVal = 0
2210          For lngX = 0& To (lngJTypes - 1&)
2220            If arr_varJType(JT_LTR, lngX) = "D" Then
2230              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2240              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2250              Exit For
2260            End If
2270          Next
2280        End With
2290      Case vbKeyI  ' ** Interest.
2300        With Me
2310          intRetVal = 0
2320          For lngX = 0& To (lngJTypes - 1&)
2330            If arr_varJType(JT_LTR, lngX) = "I" Then
2340              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2350              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2360              Exit For
2370            End If
2380          Next
2390        End With
2400      Case vbKeyJ  ' ** Cost Adj.
2410        With Me
2420          intRetVal = 0
2430          For lngX = 0& To (lngJTypes - 1&)
2440            If arr_varJType(JT_LTR, lngX) = "J" Then
2450              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2460              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2470              Exit For
2480            End If
2490          Next
2500        End With
2510      Case vbKeyL  ' ** Liability (+).
2520        With Me
2530          intRetVal = 0
2540          For lngX = 0& To (lngJTypes - 1&)
2550            If arr_varJType(JT_LTR, lngX) = "L" Then
2560              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2570              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2580              Exit For
2590            End If
2600          Next
2610        End With
2620      Case vbKeyM  ' ** Misc.
2630        With Me
2640          intRetVal = 0
2650          For lngX = 0& To (lngJTypes - 1&)
2660            If arr_varJType(JT_LTR, lngX) = "M" Then
2670              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2680              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2690              Exit For
2700            End If
2710          Next
2720        End With
2730      Case vbKeyO  ' ** Deposit.
2740        With Me
2750          intRetVal = 0
2760          For lngX = 0& To (lngJTypes - 1&)
2770            If arr_varJType(JT_LTR, lngX) = "O" Then
2780              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2790              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2800              Exit For
2810            End If
2820          Next
2830        End With
2840      Case vbKeyP  ' ** Purchase.
2850        With Me
2860          intRetVal = 0
2870          For lngX = 0& To (lngJTypes - 1&)
2880            If arr_varJType(JT_LTR, lngX) = "P" Then
2890              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
2900              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
2910              Exit For
2920            End If
2930          Next
2940        End With
2950      Case vbKeyR  ' ** Received.
2960        With Me
2970          intRetVal = 0
2980          For lngX = 0& To (lngJTypes - 1&)
2990            If arr_varJType(JT_LTR, lngX) = "R" Then
3000              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
3010              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
3020              Exit For
3030            End If
3040          Next
3050        End With
3060      Case vbKeyS  ' ** Sold.
3070        With Me
3080          intRetVal = 0
3090          For lngX = 0& To (lngJTypes - 1&)
3100            If arr_varJType(JT_LTR, lngX) = "S" Then
3110              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
3120              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
3130              Exit For
3140            End If
3150          Next
3160        End With
3170      Case vbKeyW  ' ** Withdrawn.
3180        With Me
3190          intRetVal = 0
3200          For lngX = 0& To (lngJTypes - 1&)
3210            If arr_varJType(JT_LTR, lngX) = "W" Then
3220              .Controls(arr_varJType(JT_CTL, lngX)) = (Not (.Controls(arr_varJType(JT_CTL, lngX))))
3230              .Controls(arr_varJType(JT_LBL, lngX)).FontBold = .Controls(arr_varJType(JT_CTL, lngX))
3240              Exit For
3250            End If
3260          Next
3270        End With
3280      End Select
3290    End If

        ' ** Ctrl-Shift keys.
3300    If intCtrlDown And (Not intAltDown) And intShiftDown Then
3310      Select Case intRetVal
          Case vbKeyF
3320        With Me
3330          intRetVal = 0
3340          .FocusHolder.SetFocus
3350          .TransDateStart = Null
3360        End With
3370      End Select
3380    End If

EXITP:
3390    KeyCode = intRetVal
3400    Exit Sub

ERRH:
3410    intRetVal = 0
3420    THAT_PROC = THIS_PROC
3430    That_Erl = Erl
3440    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
3450    Resume EXITP

End Sub

Private Sub Form_Timer()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

3510    With Me
3520      .TimerInterval = 0&
3530      CmdBars_Hide True  ' ** Module Procedure: modWindowFunctions.
3540      If gblnSetFocus = True Then
3550        gblnSetFocus = False
3560        strFormName = vbNullString
3570  On Error Resume Next
3580        Do Until strFormName = THIS_NAME
3590          DoCmd.SelectObject acForm, THIS_NAME, False
3600          strFormName = Screen.ActiveForm.Name
3610        Loop
3620  On Error GoTo ERRH
3630      End If
3640      If gblnGoToReport = True Then
3650        If .cmdPrint.Enabled = True Then
3660          .cmdPrint.SetFocus
3670        Else
3680          .cmdClose.SetFocus
3690        End If
3700        gblnGoToReport = False
3710        Beep
3720        DoCmd.Hourglass False
3730        DoEvents
3740      Else
3750        DoCmd.Hourglass False
3760      End If
3770    End With

EXITP:
3780    Exit Sub

ERRH:
3790    Select Case ERR.Number
        Case Else
3800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3810    End Select
3820    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3910    Select Case DataErr
        Case 2110  ' ** Access can't move the focus to the control '|'.
3920      Response = acDataErrContinue
3930    Case 2113, 2279
          ' ** 2113: The value you entered isn't valid for this field.
          ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3940      If Left(Me.ActiveControl.Name, 9) = "TransDate" Then  ' ** TransDateStart, TransDateEnd.
3950        MsgBox "Please enter a valid date (MM/DD/YYYY).", vbInformation + vbOKOnly, "Invalid Entry"
3960        Response = acDataErrContinue
3970      End If
3980    Case 2143  ' ** You didn't specify search criteria with a FindRecord action.
          ' ** I have no idea what this means!
3990      Response = acDataErrContinue
4000    Case 2237  ' ** The text you entered isn't an item in the list.
4010      Response = acDataErrDisplay
4020    Case 2302  ' ** Microsoft Access can't save the output data to the file you've selected.
4030      Beep
4040      MsgBox "Trust Accountant is unable to save the export to the file you've selected." & vbCrLf & vbCrLf & _
            "The file you're replacing may still be open.", vbInformation + vbOKOnly, "Unable To Save"
4050      Response = acDataErrContinue
4060    Case 2501  ' ** The '|' action was Canceled.
4070      Response = acDataErrContinue
4080    Case Else
4090      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
4100      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
4110    End Select

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

4200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim lngX As Long

4210    DoCmd.Hourglass True
4220    DoEvents

4230    With Me
4240      If .chkRememberDates = False Then
4250        .TransDateStart = Null  ' ** So it won't be there when Pref_Save() runs.
4260        .TransDateEnd = Null
4270      End If
4280      If .chkRememberMe = False Then
4290        .cmbAccounts = Null
4300      End If
4310      If .chkRememberTypes = False Then
4320        For lngX = 1& To 12&
4330          .Controls("JType_" & Right("00" & CStr(lngX), 2) & "_chk") = False
4340        Next
4350      End If
4360    End With

4370    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

        ' ** Make sure all reports are closed.
4380    Do While Reports.Count > 0
4390      DoCmd.Close acReport, Reports(0).Name
4400    Loop

4410    ChkSpecLedgerEntry  ' ** Module Function: modUtilities.
4420    gvarCrtRpt_FL_SpecData = Null
4430    gstrFormQuerySpec = vbNullString
4440    gblnSetFocus = True
4450    gstrReturningForm = THIS_NAME
4460    Select Case strCallingForm
        Case "frmMenu_Report"
4470      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
4480    End Select

EXITP:
4490    Exit Sub

ERRH:
4500    Select Case ERR.Number
        Case Else
4510      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4520    End Select
4530    Resume EXITP

End Sub

Private Sub FormHeader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "FormHeader_MouseMove"

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

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

4710    With Me
4720      If .cmdCalendar1_raised_focus_dots_img.Visible = True Or .cmdCalendar1_raised_focus_img.Visible = True Then
4730        Select Case blnCalendar1_Focus
            Case True
4740          .cmdCalendar1_raised_semifocus_dots_img.Visible = True
4750          .cmdCalendar1_raised_img.Visible = False
4760        Case False
4770          .cmdCalendar1_raised_img.Visible = True
4780          .cmdCalendar1_raised_semifocus_dots_img.Visible = False
4790        End Select
4800        .cmdCalendar1_raised_focus_dots_img.Visible = False
4810        .cmdCalendar1_raised_focus_img.Visible = False
4820        .cmdCalendar1_sunken_focus_dots_img.Visible = False
4830        .cmdCalendar1_raised_img_dis.Visible = False
4840      End If
4850      If .cmdCalendar2_raised_focus_dots_img.Visible = True Or .cmdCalendar2_raised_focus_img.Visible = True Then
4860        Select Case blnCalendar2_Focus
            Case True
4870          .cmdCalendar2_raised_semifocus_dots_img.Visible = True
4880          .cmdCalendar2_raised_img.Visible = False
4890        Case False
4900          .cmdCalendar2_raised_img.Visible = True
4910          .cmdCalendar2_raised_semifocus_dots_img.Visible = False
4920        End Select
4930        .cmdCalendar2_raised_focus_dots_img.Visible = False
4940        .cmdCalendar2_raised_focus_img.Visible = False
4950        .cmdCalendar2_sunken_focus_dots_img.Visible = False
4960        .cmdCalendar2_raised_img_dis.Visible = False
4970      End If
4980      If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
4990        Select Case blnSelectAll_Focus
            Case True
5000          .cmdSelectAll_raised_semifocus_dots_img.Visible = True
5010          .cmdSelectAll_raised_img.Visible = False
5020        Case False
5030          .cmdSelectAll_raised_img.Visible = True
5040          .cmdSelectAll_raised_semifocus_dots_img.Visible = False
5050        End Select
5060        .cmdSelectAll_raised_focus_img.Visible = False
5070        .cmdSelectAll_raised_focus_dots_img.Visible = False
5080        .cmdSelectAll_sunken_focus_dots_img.Visible = False
5090        .cmdSelectAll_raised_img_dis.Visible = False
5100      End If
5110      If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
5120        Select Case blnSelectNone_Focus
            Case True
5130          .cmdSelectNone_raised_semifocus_dots_img.Visible = True
5140          .cmdSelectNone_raised_img.Visible = False
5150        Case False
5160          .cmdSelectNone_raised_img.Visible = True
5170          .cmdSelectNone_raised_semifocus_dots_img.Visible = False
5180        End Select
5190        .cmdSelectNone_raised_focus_img.Visible = False
5200        .cmdSelectNone_raised_focus_dots_img.Visible = False
5210        .cmdSelectNone_sunken_focus_dots_img.Visible = False
5220        .cmdSelectNone_raised_img_dis.Visible = False
5230      End If
5240    End With

EXITP:
5250    Exit Sub

ERRH:
5260    Select Case ERR.Number
        Case Else
5270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5280    End Select
5290    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

5310    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub cmdClose_Click()

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

5410    DoCmd.Hourglass True
5420    DoEvents

5430    DoCmd.Close acForm, THIS_NAME

EXITP:
5440    Exit Sub

ERRH:
5450    Select Case ERR.Number
        Case Else
5460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5470    End Select
5480    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
          Case vbKeyTab
5570        With Me
5580          intRetVal = 0
5590          .TransDateStart.SetFocus
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab
5650        With Me
5660          intRetVal = 0
5670          If .cmdExcel.Enabled = True Then
5680            .cmdExcel.SetFocus
5690          Else
5700            .cmdWord.SetFocus
5710          End If
5720        End With
5730      End Select
5740    End If

EXITP:
5750    KeyCode = intRetVal
5760    Exit Sub

ERRH:
5770    intRetVal = 0
5780    THAT_PROC = THIS_PROC
5790    That_Erl = Erl
5800    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
5810    Resume EXITP

End Sub

Private Sub Date_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "Date_box_MouseMove"

5910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
5920    Exit Sub

ERRH:
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Sub

Private Sub TransDateStart_KeyDown(KeyCode As Integer, Shift As Integer)

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_KeyDown"

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
6090          .TransDateEnd.SetFocus
6100        End With
6110      End Select
6120    End If

        ' ** Shift keys.
6130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6140      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6150        With Me
6160          intRetVal = 0
6170          .cmdClose.SetFocus
6180        End With
6190      End Select
6200    End If

        ' ** Ctrl keys.
6210    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
6220      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6230        With Me
6240          intRetVal = 0
6250          .cmdCalendar1.SetFocus
6260        End With
6270      End Select
6280    End If

EXITP:
6290    KeyCode = intRetVal
6300    Exit Sub

ERRH:
6310    intRetVal = 0
6320    THAT_PROC = THIS_PROC
6330    That_Erl = Erl
6340    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6350    Resume EXITP

End Sub

Private Sub TransDateStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_MouseMove"

6410    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
6420    Exit Sub

ERRH:
6430    Select Case ERR.Number
        Case Else
6440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6450    End Select
6460    Resume EXITP

End Sub

Private Sub TransDateStart_Exit(Cancel As Integer)

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateStart_Exit"

6510    With Me
6520      If IsDate(.TransDateEnd) = True And IsDate(.TransDateStart) = True Then
6530        If CDate(.TransDateStart) > CDate(.TransDateEnd) Then
6540          MsgBox "The From date must be less than or equal to the To date.", vbInformation + vbOKOnly, "Invalid Date"
6550          .TransDateEnd = Null
6560          .TransDateEnd.SetFocus
6570        End If
6580      End If
6590    End With

EXITP:
6600    Exit Sub

ERRH:
6610    THAT_PROC = THIS_PROC
6620    That_Erl = Erl
6630    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6640    Resume EXITP

End Sub

Private Sub cmdCalendar1_GotFocus()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_GotFocus"

6710    With Me
6720      blnCalendar1_Focus = True
6730      .cmdCalendar1_raised_semifocus_dots_img.Visible = True
6740      .cmdCalendar1_raised_img.Visible = False
6750      .cmdCalendar1_raised_focus_img.Visible = False
6760      .cmdCalendar1_raised_focus_dots_img.Visible = False
6770      .cmdCalendar1_sunken_focus_dots_img.Visible = False
6780      .cmdCalendar1_raised_img_dis.Visible = False
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

Private Sub cmdCalendar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

6900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseDown"

6910    With Me
6920      blnCalendar1_MouseDown = True
6930      .cmdCalendar1_sunken_focus_dots_img.Visible = True
6940      .cmdCalendar1_raised_img.Visible = False
6950      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
6960      .cmdCalendar1_raised_focus_img.Visible = False
6970      .cmdCalendar1_raised_focus_dots_img.Visible = False
6980      .cmdCalendar1_raised_img_dis.Visible = False
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

Private Sub cmdCalendar1_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

7100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim blnRetVal As Boolean

7110    With Me
7120      datStartDate = Date
7130      datEndDate = 0
7140      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
7150      If blnRetVal = True Then
7160        .TransDateStart = datStartDate
7170      Else
7180        .TransDateStart = CDate(Format(Date, "mm/dd/yyyy"))
7190      End If
7200      .TransDateStart.SetFocus
7210    End With

EXITP:
7220    Exit Sub

ERRH:
7230    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
7240    Case Else
7250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7260    End Select
7270    Resume EXITP

End Sub

Private Sub cmdCalendar1_KeyDown(KeyCode As Integer, Shift As Integer)

7300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

7310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
7320    intShiftDown = (Shift And acShiftMask) > 0
7330    intAltDown = (Shift And acAltMask) > 0
7340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
7350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
7360      Select Case intRetVal
          Case vbKeyTab
7370        With Me
7380          intRetVal = 0
7390          .TransDateEnd.SetFocus
7400        End With
7410      End Select
7420    End If

        ' ** Shift keys.
7430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7440      Select Case intRetVal
          Case vbKeyTab
7450        With Me
7460          intRetVal = 0
7470          .TransDateStart.SetFocus
7480        End With
7490      End Select
7500    End If

EXITP:
7510    KeyCode = intRetVal
7520    Exit Sub

ERRH:
7530    intRetVal = 0
7540    THAT_PROC = THIS_PROC
7550    That_Erl = Erl
7560    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
7570    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseMove"

7610    With Me
7620      If blnCalendar1_MouseDown = False Then
7630        Select Case blnCalendar1_Focus
            Case True
7640          .cmdCalendar1_raised_focus_dots_img.Visible = True
7650          .cmdCalendar1_raised_focus_img.Visible = False
7660        Case False
7670          .cmdCalendar1_raised_focus_img.Visible = True
7680          .cmdCalendar1_raised_focus_dots_img.Visible = False
7690        End Select
7700        .cmdCalendar1_raised_img.Visible = False
7710        .cmdCalendar1_raised_semifocus_dots_img.Visible = False
7720        .cmdCalendar1_sunken_focus_dots_img.Visible = False
7730        .cmdCalendar1_raised_img_dis.Visible = False
7740      End If
7750    End With

EXITP:
7760    Exit Sub

ERRH:
7770    Select Case ERR.Number
        Case Else
7780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7790    End Select
7800    Resume EXITP

End Sub

Private Sub cmdCalendar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_MouseUp"

7910    With Me
7920      .cmdCalendar1_raised_focus_dots_img.Visible = True
7930      .cmdCalendar1_raised_img.Visible = False
7940      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
7950      .cmdCalendar1_raised_focus_img.Visible = False
7960      .cmdCalendar1_sunken_focus_dots_img.Visible = False
7970      .cmdCalendar1_raised_img_dis.Visible = False
7980      blnCalendar1_MouseDown = False
7990    End With

EXITP:
8000    Exit Sub

ERRH:
8010    Select Case ERR.Number
        Case Else
8020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8030    End Select
8040    Resume EXITP

End Sub

Private Sub cmdCalendar1_LostFocus()

8100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar1_LostFocus"

8110    With Me
8120      .cmdCalendar1_raised_img.Visible = True
8130      .cmdCalendar1_raised_semifocus_dots_img.Visible = False
8140      .cmdCalendar1_raised_focus_img.Visible = False
8150      .cmdCalendar1_raised_focus_dots_img.Visible = False
8160      .cmdCalendar1_sunken_focus_dots_img.Visible = False
8170      .cmdCalendar1_raised_img_dis.Visible = False
8180      blnCalendar1_Focus = False
8190    End With

EXITP:
8200    Exit Sub

ERRH:
8210    Select Case ERR.Number
        Case Else
8220      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8230    End Select
8240    Resume EXITP

End Sub

Private Sub TransDateEnd_KeyDown(KeyCode As Integer, Shift As Integer)

8300  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

8310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
8320    intShiftDown = (Shift And acShiftMask) > 0
8330    intAltDown = (Shift And acAltMask) > 0
8340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
8350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
8360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8370        With Me
8380          intRetVal = 0
8390          .cmbAccounts.SetFocus
8400        End With
8410      End Select
8420    End If

        ' ** Shift keys.
8430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
8440      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8450        With Me
8460          intRetVal = 0
8470          .TransDateStart.SetFocus
8480        End With
8490      End Select
8500    End If

        ' ** Ctrl keys.
8510    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
8520      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8530        With Me
8540          intRetVal = 0
8550          .cmdCalendar2.SetFocus
8560        End With
8570      End Select
8580    End If

        ' ** Ctrl-Shift keys.
8590    If intCtrlDown And (Not intAltDown) And intShiftDown Then
8600      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
8610        With Me
8620          intRetVal = 0
8630          .cmdCalendar1.SetFocus
8640        End With
8650      End Select
8660    End If

EXITP:
8670    KeyCode = intRetVal
8680    Exit Sub

ERRH:
8690    intRetVal = 0
8700    THAT_PROC = THIS_PROC
8710    That_Erl = Erl
8720    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
8730    Resume EXITP

End Sub

Private Sub TransDateEnd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

8800  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_MouseMove"

8810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
8820    Exit Sub

ERRH:
8830    Select Case ERR.Number
        Case Else
8840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8850    End Select
8860    Resume EXITP

End Sub

Private Sub TransDateEnd_Exit(Cancel As Integer)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "TransDateEnd_Exit"

8910    With Me
8920      If IsDate(.TransDateStart) = True And IsDate(.TransDateEnd) = True Then
8930        If CDate(.TransDateStart) > CDate(.TransDateEnd) Then
8940          MsgBox "The start date must be less than or equal to the end date.", vbInformation + vbOKOnly, "Invalid Date"
8950          .TransDateStart.SetFocus
8960        End If
8970      Else
8980        If IsNull(.TransDateStart) = True And IsNull(.TransDateEnd) = True Then
              ' ** Just let them tab through the empty fields.
8990        Else
9000          If IsDate(.TransDateStart) = True Then
                ' ** Populate with today's date.
9010            .TransDateEnd = Date
9020          ElseIf IsDate(.TransDateEnd) = True Then
9030            MsgBox "The From date must be less than or equal to the To date.", vbInformation + vbOKOnly, "Invalid Date"
9040            .TransDateStart.SetFocus
9050          End If
9060        End If
9070      End If
9080    End With

EXITP:
9090    Exit Sub

ERRH:
9100    THAT_PROC = THIS_PROC
9110    That_Erl = Erl
9120    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
9130    Resume EXITP

End Sub

Private Sub cmdCalendar2_GotFocus()

9200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_GotFocus"

9210    With Me
9220      blnCalendar2_Focus = True
9230      .cmdCalendar2_raised_semifocus_dots_img.Visible = True
9240      .cmdCalendar2_raised_img.Visible = False
9250      .cmdCalendar2_raised_focus_img.Visible = False
9260      .cmdCalendar2_raised_focus_dots_img.Visible = False
9270      .cmdCalendar2_sunken_focus_dots_img.Visible = False
9280      .cmdCalendar2_raised_img_dis.Visible = False
9290    End With

EXITP:
9300    Exit Sub

ERRH:
9310    Select Case ERR.Number
        Case Else
9320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9330    End Select
9340    Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

9400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseDown"

9410    With Me
9420      blnCalendar2_MouseDown = True
9430      .cmdCalendar2_sunken_focus_dots_img.Visible = True
9440      .cmdCalendar2_raised_img.Visible = False
9450      .cmdCalendar2_raised_semifocus_dots_img.Visible = False
9460      .cmdCalendar2_raised_focus_img.Visible = False
9470      .cmdCalendar2_raised_focus_dots_img.Visible = False
9480      .cmdCalendar2_raised_img_dis.Visible = False
9490    End With

EXITP:
9500    Exit Sub

ERRH:
9510    Select Case ERR.Number
        Case Else
9520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9530    End Select
9540    Resume EXITP

End Sub

Private Sub cmdCalendar2_Click()
' ** Retrieve the currently selected date(s).
' ** Call our Function to display the Calendar.
' ** Defaults to showing Today's Date.

9600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_Click"

        Dim datStartDate As Date, datEndDate As Date
        Dim Cancel As Integer
        Dim blnRetVal As Boolean

9610    With Me
9620      datStartDate = Date
9630      datEndDate = 0
9640      blnRetVal = ShowMonthCalendar(clsMonthClass, datStartDate, datEndDate)  ' ** Module Function: modCalendar.
9650      If blnRetVal = True Then
9660        .TransDateEnd = datStartDate
9670      Else
9680        .TransDateEnd = CDate(Format(Date, "mm/dd/yyyy"))
9690      End If
9700      .TransDateEnd.SetFocus
9710      Cancel = 0
9720      TransDateEnd_Exit Cancel  ' ** Procedure: Above.
9730      If Cancel = 0 Then
9740        .cmbAccounts.SetFocus
9750      End If
9760    End With

EXITP:
9770    Exit Sub

ERRH:
9780    Select Case ERR.Number
        Case 2110  ' ** Access can't move the focus to the control '|'.
          ' ** Do nothing.
9790    Case Else
9800      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9810    End Select
9820    Resume EXITP

End Sub

Private Sub cmdCalendar2_KeyDown(KeyCode As Integer, Shift As Integer)

9900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_KeyDown"

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
          Case vbKeyTab
9970        With Me
9980          intRetVal = 0
9990          .chkTransDateAll.SetFocus
10000       End With
10010     End Select
10020   End If

        ' ** Shift keys.
10030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
10040     Select Case intRetVal
          Case vbKeyTab
10050       With Me
10060         intRetVal = 0
10070         .TransDateEnd.SetFocus
10080       End With
10090     End Select
10100   End If

EXITP:
10110   KeyCode = intRetVal
10120   Exit Sub

ERRH:
10130   intRetVal = 0
10140   THAT_PROC = THIS_PROC
10150   That_Erl = Erl
10160   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
10170   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10200 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseMove"

10210   With Me
10220     If blnCalendar2_MouseDown = False Then
10230       Select Case blnCalendar2_Focus
            Case True
10240         .cmdCalendar2_raised_focus_dots_img.Visible = True
10250         .cmdCalendar2_raised_focus_img.Visible = False
10260       Case False
10270         .cmdCalendar2_raised_focus_img.Visible = True
10280         .cmdCalendar2_raised_focus_dots_img.Visible = False
10290       End Select
10300       .cmdCalendar2_raised_img.Visible = False
10310       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
10320       .cmdCalendar2_sunken_focus_dots_img.Visible = False
10330       .cmdCalendar2_raised_img_dis.Visible = False
10340     End If
10350   End With

EXITP:
10360   Exit Sub

ERRH:
10370   Select Case ERR.Number
        Case Else
10380     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10390   End Select
10400   Resume EXITP

End Sub

Private Sub cmdCalendar2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_MouseUp"

10510   With Me
10520     .cmdCalendar2_raised_focus_dots_img.Visible = True
10530     .cmdCalendar2_raised_img.Visible = False
10540     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
10550     .cmdCalendar2_raised_focus_img.Visible = False
10560     .cmdCalendar2_sunken_focus_dots_img.Visible = False
10570     .cmdCalendar2_raised_img_dis.Visible = False
10580     blnCalendar2_MouseDown = False
10590   End With

EXITP:
10600   Exit Sub

ERRH:
10610   Select Case ERR.Number
        Case Else
10620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10630   End Select
10640   Resume EXITP

End Sub

Private Sub cmdCalendar2_LostFocus()

10700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCalendar2_LostFocus"

10710   With Me
10720     .cmdCalendar2_raised_img.Visible = True
10730     .cmdCalendar2_raised_semifocus_dots_img.Visible = False
10740     .cmdCalendar2_raised_focus_img.Visible = False
10750     .cmdCalendar2_raised_focus_dots_img.Visible = False
10760     .cmdCalendar2_sunken_focus_dots_img.Visible = False
10770     .cmdCalendar2_raised_img_dis.Visible = False
10780     blnCalendar2_Focus = False
10790   End With

EXITP:
10800   Exit Sub

ERRH:
10810   Select Case ERR.Number
        Case Else
10820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
10830   End Select
10840   Resume EXITP

End Sub

Private Sub chkTransDateAll_AfterUpdate()

10900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkTransDateAll_AfterUpdate"

10910   With Me
10920     Select Case .chkTransDateAll
          Case True
10930       .chkTransDateAll_lbl.FontBold = True
10940       .TransDateStart.Enabled = False
10950       .TransDateStart.BorderColor = WIN_CLR_DISR
10960       .TransDateStart.BackStyle = acBackStyleTransparent
10970       .cmdCalendar1.Enabled = False
10980       .cmdCalendar1_raised_img_dis.Visible = True
10990       .cmdCalendar1_raised_img.Visible = False
11000       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
11010       .cmdCalendar1_raised_focus_img.Visible = False
11020       .cmdCalendar1_raised_focus_dots_img.Visible = False
11030       .cmdCalendar1_sunken_focus_dots_img.Visible = False
11040       .TransDateEnd.Enabled = False
11050       .TransDateEnd.BorderColor = WIN_CLR_DISR
11060       .TransDateEnd.BackStyle = acBackStyleTransparent
11070       .cmdCalendar2.Enabled = False
11080       .cmdCalendar2_raised_img_dis.Visible = True
11090       .cmdCalendar2_raised_img.Visible = False
11100       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
11110       .cmdCalendar2_raised_focus_img.Visible = False
11120       .cmdCalendar2_raised_focus_dots_img.Visible = False
11130       .cmdCalendar2_sunken_focus_dots_img.Visible = False
11140       .chkRememberDates.Enabled = False
11150       .chkRememberDates_lbl.Visible = False
11160       .chkRememberDates_lbl2_dim.Visible = True
11170       .chkRememberDates_lbl2_dim_hi.Visible = True
11180     Case False
11190       .chkTransDateAll_lbl.FontBold = False
11200       .TransDateStart.Enabled = True
11210       .TransDateStart.BorderColor = CLR_LTBLU2
11220       .TransDateStart.BackStyle = acBackStyleNormal
11230       .cmdCalendar1.Enabled = True
11240       .cmdCalendar1_raised_img.Visible = True
11250       .cmdCalendar1_raised_semifocus_dots_img.Visible = False
11260       .cmdCalendar1_raised_focus_img.Visible = False
11270       .cmdCalendar1_raised_focus_dots_img.Visible = False
11280       .cmdCalendar1_sunken_focus_dots_img.Visible = False
11290       .cmdCalendar1_raised_img_dis.Visible = False
11300       .TransDateEnd.Enabled = True
11310       .TransDateEnd.BorderColor = CLR_LTBLU2
11320       .TransDateEnd.BackStyle = acBackStyleNormal
11330       .TransDateStart.SetFocus
11340       .cmdCalendar2.Enabled = True
11350       .cmdCalendar2_raised_img.Visible = True
11360       .cmdCalendar2_raised_semifocus_dots_img.Visible = False
11370       .cmdCalendar2_raised_focus_img.Visible = False
11380       .cmdCalendar2_raised_focus_dots_img.Visible = False
11390       .cmdCalendar2_sunken_focus_dots_img.Visible = False
11400       .cmdCalendar2_raised_img_dis.Visible = False
11410       .chkRememberDates.Enabled = True
11420       .chkRememberDates_lbl.Visible = True
11430       .chkRememberDates_lbl2_dim.Visible = False
11440       .chkRememberDates_lbl2_dim_hi.Visible = False
11450     End Select
11460   End With

EXITP:
11470   Exit Sub

ERRH:
11480   THAT_PROC = THIS_PROC
11490   That_Erl = Erl
11500   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
11510   Resume EXITP

End Sub

Private Sub chkTransDateAll_KeyDown(KeyCode As Integer, Shift As Integer)

11600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkTransDateAll_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

11610   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
11620   intShiftDown = (Shift And acShiftMask) > 0
11630   intAltDown = (Shift And acAltMask) > 0
11640   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
11650   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
11660     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11670       With Me
11680         intRetVal = 0
11690         If .chkRememberDates.Enabled = True Then
11700           .chkRememberDates.SetFocus
11710         Else
11720           .opgAccountSource.SetFocus
11730         End If
11740       End With
11750     End Select
11760   End If

        ' ** Shift keys.
11770   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
11780     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
11790       With Me
11800         intRetVal = 0
11810         If .cmdCalendar2.Enabled = True Then
11820           .cmdCalendar2.SetFocus
11830         ElseIf .TransDateEnd.Enabled = True Then
11840           .TransDateEnd.SetFocus
11850         Else
11860           .cmdClose.SetFocus
11870         End If
11880       End With
11890     End Select
11900   End If

EXITP:
11910   KeyCode = intRetVal
11920   Exit Sub

ERRH:
11930   intRetVal = 0
11940   THAT_PROC = THIS_PROC
11950   That_Erl = Erl
11960   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
11970   Resume EXITP

End Sub

Private Sub chkTransDateAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkTransDateAll_MouseMove"

12010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
12020   Exit Sub

ERRH:
12030   Select Case ERR.Number
        Case Else
12040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12050   End Select
12060   Resume EXITP

End Sub

Private Sub chkRememberDates_AfterUpdate()

12100 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_AfterUpdate"

12110   With Me
12120     Select Case .chkRememberDates
          Case True
12130       .chkRememberDates_lbl.FontBold = True
12140       .chkRememberDates_lbl2_dim.FontBold = True
12150       .chkRememberDates_lbl2_dim_hi.FontBold = True
12160     Case False
12170       .chkRememberDates_lbl.FontBold = False
12180       .chkRememberDates_lbl2_dim.FontBold = False
12190       .chkRememberDates_lbl2_dim_hi.FontBold = False
12200     End Select
12210   End With

EXITP:
12220   Exit Sub

ERRH:
12230   THAT_PROC = THIS_PROC
12240   That_Erl = Erl
12250   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
12260   Resume EXITP

End Sub

Private Sub chkRememberDates_KeyDown(KeyCode As Integer, Shift As Integer)

12300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

12310   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
12320   intShiftDown = (Shift And acShiftMask) > 0
12330   intAltDown = (Shift And acAltMask) > 0
12340   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
12350   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
12360     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12370       With Me
12380         intRetVal = 0
12390         .opgAccountSource.SetFocus
12400       End With
12410     End Select
12420   End If

        ' ** Shift keys.
12430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
12440     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
12450       With Me
12460         intRetVal = 0
12470         .chkTransDateAll.SetFocus
12480       End With
12490     End Select
12500   End If

EXITP:
12510   KeyCode = intRetVal
12520   Exit Sub

ERRH:
12530   intRetVal = 0
12540   THAT_PROC = THIS_PROC
12550   That_Erl = Erl
12560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
12570   Resume EXITP

End Sub

Private Sub chkRememberDates_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

12600 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberDates_MouseMove"

12610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Below.

EXITP:
12620   Exit Sub

ERRH:
12630   Select Case ERR.Number
        Case Else
12640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
12650   End Select
12660   Resume EXITP

End Sub

Private Sub opgAccountSource_AfterUpdate()

12700 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_AfterUpdate"

        Dim strAccountNo As String

12710   strAccountNo = vbNullString

12720   With Me
12730     If IsNull(.cmbAccounts) = False Then
12740       If Len(.cmbAccounts.Column(0)) > 0 Then
12750         strAccountNo = .cmbAccounts.Column(0)
12760       End If
12770     End If
12780     Select Case .opgAccountSource
          Case .opgAccountSource_optNumber.OptionValue
12790       .cmbAccounts.RowSource = "qryAccountNoDropDown_03"
12800       .opgAccountSource_optNumber_lbl.FontBold = True
12810       .opgAccountSource_optName_lbl.FontBold = False
12820     Case .opgAccountSource_optName.OptionValue
12830       .cmbAccounts.RowSource = "qryAccountNoDropDown_04"
12840       .opgAccountSource_optNumber_lbl.FontBold = False
12850       .opgAccountSource_optName_lbl.FontBold = True
12860     End Select
12870     DoEvents
12880     If strAccountNo <> vbNullString Then
12890       .cmbAccounts = strAccountNo
12900     End If
12910   End With

EXITP:
12920   Exit Sub

ERRH:
12930   THAT_PROC = THIS_PROC
12940   That_Erl = Erl
12950   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
12960   Resume EXITP

End Sub

Private Sub opgAccountSource_optNumber_KeyDown(KeyCode As Integer, Shift As Integer)

13000 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optNumber_KeyDown"

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
13170         If .chkRememberDates.Enabled = True Then
13180           .chkRememberDates.SetFocus
13190         Else
13200           .chkTransDateAll.SetFocus
13210         End If
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

Private Sub opgAccountSource_optName_KeyDown(KeyCode As Integer, Shift As Integer)

13400 On Error GoTo ERRH

        Const THIS_PROC As String = "opgAccountSource_optName_KeyDown"

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
13490         .chkRememberMe.SetFocus
13500       End With
13510     End Select
13520   End If

        ' ** Shift keys.
13530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
13540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
13550       With Me
13560         intRetVal = 0
13570         If .chkRememberDates.Enabled = True Then
13580           .chkRememberDates.SetFocus
13590         Else
13600           .chkTransDateAll.SetFocus
13610         End If
13620       End With
13630     End Select
13640   End If

EXITP:
13650   KeyCode = intRetVal
13660   Exit Sub

ERRH:
13670   intRetVal = 0
13680   THAT_PROC = THIS_PROC
13690   That_Erl = Erl
13700   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
13710   Resume EXITP

End Sub

Private Sub chkRememberMe_AfterUpdate()

13800 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_AfterUpdate"

13810   With Me
13820     Select Case .chkRememberMe
          Case True
13830       .chkRememberMe_lbl.FontBold = True
13840     Case False
13850       .chkRememberMe_lbl.FontBold = False
13860     End Select
13870   End With

EXITP:
13880   Exit Sub

ERRH:
13890   THAT_PROC = THIS_PROC
13900   That_Erl = Erl
13910   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
13920   Resume EXITP

End Sub

Private Sub chkRememberMe_KeyDown(KeyCode As Integer, Shift As Integer)

14000 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberMe_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

14010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
14020   intShiftDown = (Shift And acShiftMask) > 0
14030   intAltDown = (Shift And acAltMask) > 0
14040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
14050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
14060     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14070       With Me
14080         intRetVal = 0
14090         .cmbAccounts.SetFocus
14100       End With
14110     End Select
14120   End If

        ' ** Shift keys.
14130   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14140     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14150       With Me
14160         intRetVal = 0
14170         .opgAccountSource.SetFocus
14180       End With
14190     End Select
14200   End If

EXITP:
14210   KeyCode = intRetVal
14220   Exit Sub

ERRH:
14230   intRetVal = 0
14240   THAT_PROC = THIS_PROC
14250   That_Erl = Erl
14260   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14270   Resume EXITP

End Sub

Private Sub cmbAccounts_Enter()

14300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_Enter"

14310   With Me
14320     If .cmbAccounts.text = vbNullString Then
14330       ExpandCombo  ' ** Module Procedure: modUtilities.
14340       .cmbAccounts = .cmbAccounts.ItemData(0)
14350     End If
14360   End With

EXITP:
14370   Exit Sub

ERRH:
14380   THAT_PROC = THIS_PROC
14390   That_Erl = Erl
14400   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
14410   Resume EXITP

End Sub

Private Sub cmbAccounts_NotInList(NewData As String, Response As Integer)

14500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmbAccounts_NotInList"

14510   MsgBox "You must select a valid account.", vbInformation + vbOKOnly, "Invalid Entry"
14520   Response = acDataErrContinue

EXITP:
14530   Exit Sub

ERRH:
14540   THAT_PROC = THIS_PROC
14550   That_Erl = Erl
14560   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
14570   Resume EXITP

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
14690         .JType_01_chk.SetFocus
14700       End With
14710     End Select
14720   End If

        ' ** Shift keys.
14730   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
14740     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14750       With Me
14760         intRetVal = 0
14770         If .TransDateEnd.Enabled = True Then
14780           .TransDateEnd.SetFocus
14790         Else
14800           .chkTransDateAll.SetFocus
14810         End If
14820       End With
14830     End Select
14840   End If

        ' ** Ctrl keys.
14850   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
14860     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14870       With Me
14880         intRetVal = 0
14890         If .chkIncludeArchive.Enabled = True Then
14900           .chkIncludeArchive.SetFocus
14910         Else
14920           .chkRememberTypes.SetFocus
14930         End If
14940       End With
14950     End Select
14960   End If

        ' ** Ctrl-Shift keys.
14970   If intCtrlDown And (Not intAltDown) And intShiftDown Then
14980     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
14990       With Me
15000         intRetVal = 0
15010         .chkRememberMe.SetFocus
15020       End With
15030     End Select
15040   End If

EXITP:
15050   KeyCode = intRetVal
15060   Exit Sub

ERRH:
15070   intRetVal = 0
15080   THAT_PROC = THIS_PROC
15090   That_Erl = Erl
15100   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15110   Resume EXITP

End Sub

Private Sub chkIncludeArchive_AfterUpdate()

15200 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_AfterUpdate"

15210   With Me
15220     Select Case .chkIncludeArchive
          Case True
15230       .chkIncludeArchive_lbl.FontBold = True
15240     Case False
15250       .chkIncludeArchive_lbl.FontBold = False
15260     End Select
15270   End With

EXITP:
15280   Exit Sub

ERRH:
15290   Select Case ERR.Number
        Case Else
15300     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15310   End Select
15320   Resume EXITP

End Sub

Private Sub chkIncludeArchive_KeyDown(KeyCode As Integer, Shift As Integer)

15400 On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeArchive_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15420   intShiftDown = (Shift And acShiftMask) > 0
15430   intAltDown = (Shift And acAltMask) > 0
15440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15460     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15470       With Me
15480         intRetVal = 0
15490         .chkRememberTypes.SetFocus
15500       End With
15510     End Select
15520   End If

        ' ** Shift keys.
15530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
15540     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15550       With Me
15560         intRetVal = 0
15570         .cmbAccounts.SetFocus
15580       End With
15590     End Select
15600   End If

EXITP:
15610   KeyCode = intRetVal
15620   Exit Sub

ERRH:
15630   intRetVal = 0
15640   THAT_PROC = THIS_PROC
15650   That_Erl = Erl
15660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
15670   Resume EXITP

End Sub

Private Sub chkRememberTypes_AfterUpdate()

15700 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberTypes_AfterUpdate"

15710   With Me
15720     Select Case .chkRememberTypes
          Case True
15730       .chkRememberTypes_lbl.FontBold = True
15740     Case False
15750       .chkRememberTypes_lbl.FontBold = False
15760     End Select
15770   End With

EXITP:
15780   Exit Sub

ERRH:
15790   DoCmd.Hourglass False
15800   Select Case ERR.Number
        Case Else
15810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
15820   End Select
15830   Resume EXITP

End Sub

Private Sub chkRememberTypes_KeyDown(KeyCode As Integer, Shift As Integer)

15900 On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberTypes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

15910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
15920   intShiftDown = (Shift And acShiftMask) > 0
15930   intAltDown = (Shift And acAltMask) > 0
15940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
15950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
15960     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
15970       With Me
15980         intRetVal = 0
15990         .JType_01_chk.SetFocus
16000       End With
16010     End Select
16020   End If

        ' ** Shift keys.
16030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16040     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16050       With Me
16060         intRetVal = 0
16070         If .chkIncludeArchive.Enabled = True Then
16080           .chkIncludeArchive.SetFocus
16090         Else
16100           .cmbAccounts.SetFocus
16110         End If
16120       End With
16130     End Select
16140   End If

EXITP:
16150   KeyCode = intRetVal
16160   Exit Sub

ERRH:
16170   intRetVal = 0
16180   THAT_PROC = THIS_PROC
16190   That_Erl = Erl
16200   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16210   Resume EXITP

End Sub

Private Sub JType_box_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

16300 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_box_MouseMove"

16310   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
16320   Exit Sub

ERRH:
16330   Select Case ERR.Number
        Case Else
16340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
16350   End Select
16360   Resume EXITP

End Sub

Private Sub JType_01_chk_AfterUpdate()

16400 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_01_chk_AfterUpdate"

16410   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
16420   Exit Sub

ERRH:
16430   THAT_PROC = THIS_PROC
16440   That_Erl = Erl
16450   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
16460   Resume EXITP

End Sub

Private Sub JType_01_chk_KeyDown(KeyCode As Integer, Shift As Integer)

16500 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_01_chk_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

16510   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
16520   intShiftDown = (Shift And acShiftMask) > 0
16530   intAltDown = (Shift And acAltMask) > 0
16540   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
16550   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
16560     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16570       With Me
16580         intRetVal = 0
16590         .JType_02_chk.SetFocus
16600       End With
16610     End Select
16620   End If

        ' ** Shift keys.
16630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
16640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16650       With Me
16660         intRetVal = 0
16670         .cmbAccounts.SetFocus
16680       End With
16690     End Select
16700   End If

        ' ** Ctrl-Shift keys.
16710   If intCtrlDown And (Not intAltDown) And intShiftDown Then
16720     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
16730       With Me
16740         intRetVal = 0
16750         .chkRememberTypes.SetFocus
16760       End With
16770     End Select
16780   End If

EXITP:
16790   KeyCode = intRetVal
16800   Exit Sub

ERRH:
16810   intRetVal = 0
16820   THAT_PROC = THIS_PROC
16830   That_Erl = Erl
16840   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
16850   Resume EXITP

End Sub

Private Sub JType_02_chk_AfterUpdate()

16900 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_02_chk_AfterUpdate"

16910   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
16920   Exit Sub

ERRH:
16930   THAT_PROC = THIS_PROC
16940   That_Erl = Erl
16950   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
16960   Resume EXITP

End Sub

Private Sub JType_03_chk_AfterUpdate()

17000 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_03_chk_AfterUpdate"

17010   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17020   Exit Sub

ERRH:
17030   THAT_PROC = THIS_PROC
17040   That_Erl = Erl
17050   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17060   Resume EXITP

End Sub

Private Sub JType_04_chk_AfterUpdate()

17100 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_04_chk_AfterUpdate"

17110   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17120   Exit Sub

ERRH:
17130   THAT_PROC = THIS_PROC
17140   That_Erl = Erl
17150   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17160   Resume EXITP

End Sub

Private Sub JType_05_chk_AfterUpdate()

17200 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_05_chk_AfterUpdate"

17210   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17220   Exit Sub

ERRH:
17230   THAT_PROC = THIS_PROC
17240   That_Erl = Erl
17250   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17260   Resume EXITP

End Sub

Private Sub JType_06_chk_AfterUpdate()

17300 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_06_chk_AfterUpdate"

17310   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17320   Exit Sub

ERRH:
17330   THAT_PROC = THIS_PROC
17340   That_Erl = Erl
17350   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17360   Resume EXITP

End Sub

Private Sub JType_07_chk_AfterUpdate()

17400 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_07_chk_AfterUpdate"

17410   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17420   Exit Sub

ERRH:
17430   THAT_PROC = THIS_PROC
17440   That_Erl = Erl
17450   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17460   Resume EXITP

End Sub

Private Sub JType_08_chk_AfterUpdate()

17500 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_08_chk_AfterUpdate"

17510   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17520   Exit Sub

ERRH:
17530   THAT_PROC = THIS_PROC
17540   That_Erl = Erl
17550   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17560   Resume EXITP

End Sub

Private Sub JType_09_chk_AfterUpdate()

17600 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_09_chk_AfterUpdate"

17610   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17620   Exit Sub

ERRH:
17630   THAT_PROC = THIS_PROC
17640   That_Erl = Erl
17650   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17660   Resume EXITP

End Sub

Private Sub JType_10_chk_AfterUpdate()

17700 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_10_chk_AfterUpdate"

17710   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17720   Exit Sub

ERRH:
17730   THAT_PROC = THIS_PROC
17740   That_Erl = Erl
17750   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17760   Resume EXITP

End Sub

Private Sub JType_10_chk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

17800 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_10_chk_MouseMove"

17810   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
17820   Exit Sub

ERRH:
17830   Select Case ERR.Number
        Case Else
17840     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
17850   End Select
17860   Resume EXITP

End Sub

Private Sub JType_11_chk_AfterUpdate()

17900 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_11_chk_AfterUpdate"

17910   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
17920   Exit Sub

ERRH:
17930   THAT_PROC = THIS_PROC
17940   That_Erl = Erl
17950   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
17960   Resume EXITP

End Sub

Private Sub JType_11_chk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18000 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_11_chk_MouseMove"

18010   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
18020   Exit Sub

ERRH:
18030   Select Case ERR.Number
        Case Else
18040     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18050   End Select
18060   Resume EXITP

End Sub

Private Sub JType_12_chk_AfterUpdate()

18100 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_12_chk_AfterUpdate"

18110   JType_Chk THIS_PROC  ' ** Procedure: Below.

EXITP:
18120   Exit Sub

ERRH:
18130   THAT_PROC = THIS_PROC
18140   That_Erl = Erl
18150   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
18160   Resume EXITP

End Sub

Private Sub JType_12_chk_KeyDown(KeyCode As Integer, Shift As Integer)

18200 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_12_chk_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

18210   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
18220   intShiftDown = (Shift And acShiftMask) > 0
18230   intAltDown = (Shift And acAltMask) > 0
18240   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
18250   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
18260     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18270       With Me
18280         intRetVal = 0
18290         .cmdPreview.SetFocus
18300       End With
18310     End Select
18320   End If

        ' ** Shift keys.
18330   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
18340     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18350       With Me
18360         intRetVal = 0
18370         .JType_11_chk.SetFocus
18380       End With
18390     End Select
18400   End If

        ' ** Ctrl keys.
18410   If intCtrlDown And (Not intAltDown) And intShiftDown Then
18420     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
18430       With Me
18440         intRetVal = 0
18450         .cmdSelectAll.SetFocus
18460       End With
18470     End Select
18480   End If

EXITP:
18490   KeyCode = intRetVal
18500   Exit Sub

ERRH:
18510   intRetVal = 0
18520   THAT_PROC = THIS_PROC
18530   That_Erl = Erl
18540   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
18550   Resume EXITP

End Sub

Private Sub JType_12_chk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

18600 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_12_chk_MouseMove"

18610   Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
18620   Exit Sub

ERRH:
18630   Select Case ERR.Number
        Case Else
18640     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18650   End Select
18660   Resume EXITP

End Sub

Private Sub cmdSelectAll_GotFocus()

18700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_GotFocus"

18710   With Me
18720     blnSelectAll_Focus = True
18730     .cmdSelectAll_raised_semifocus_dots_img.Visible = True
18740     .cmdSelectAll_raised_img.Visible = False
18750     .cmdSelectAll_raised_focus_img.Visible = False
18760     .cmdSelectAll_raised_focus_dots_img.Visible = False
18770     .cmdSelectAll_sunken_focus_dots_img.Visible = False
18780     .cmdSelectAll_raised_img_dis.Visible = False
18790   End With

EXITP:
18800   Exit Sub

ERRH:
18810   Select Case ERR.Number
        Case Else
18820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
18830   End Select
18840   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

18900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseDown"

18910   With Me
18920     blnSelectAll_MouseDown = True
18930     .cmdSelectAll_sunken_focus_dots_img.Visible = True
18940     .cmdSelectAll_raised_img.Visible = False
18950     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
18960     .cmdSelectAll_raised_focus_img.Visible = False
18970     .cmdSelectAll_raised_focus_dots_img.Visible = False
18980     .cmdSelectAll_raised_img_dis.Visible = False
18990   End With

EXITP:
19000   Exit Sub

ERRH:
19010   Select Case ERR.Number
        Case Else
19020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19030   End Select
19040   Resume EXITP

End Sub

Private Sub cmdSelectAll_Click()

19100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_Click"

        Dim strCtl As String
        Dim lngX As Long

19110   With Me
19120     For lngX = 0& To (lngJTypes - 1&)
19130       strCtl = arr_varJType(JT_CTL, lngX)
19140       .Controls(strCtl) = True
19150       JType_Chk (strCtl & "_AfterUpdate")  ' ** Procedure: Below.
19160       DoEvents
19170     Next
19180   End With

EXITP:
19190   Exit Sub

ERRH:
19200   Select Case ERR.Number
        Case Else
19210     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
19220   End Select
19230   Resume EXITP

End Sub

Private Sub cmdSelectAll_KeyDown(KeyCode As Integer, Shift As Integer)

19300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_KeyDown"

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
          Case vbKeyTab
19370       With Me
19380         intRetVal = 0
19390         .cmdSelectNone.SetFocus
19400       End With
19410     End Select
19420   End If

        ' ** Shift keys.
19430   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
19440     Select Case intRetVal
          Case vbKeyTab
19450       With Me
19460         intRetVal = 0
19470         If .JType_12_chk.Enabled = True And .JType_12_chk.Visible = True Then
19480           .JType_12_chk.SetFocus
19490         Else
19500           .JType_11_chk.SetFocus
19510         End If
19520       End With
19530     End Select
19540   End If

        ' ** Ctrl-Shift keys.
19550   If intCtrlDown And (Not intAltDown) And intShiftDown Then
19560     Select Case intRetVal
          Case vbKeyTab
19570       With Me
19580         intRetVal = 0
19590         .chkRememberTypes.SetFocus
19600       End With
19610     End Select
19620   End If

EXITP:
19630   KeyCode = intRetVal
19640   Exit Sub

ERRH:
19650   intRetVal = 0
19660   THAT_PROC = THIS_PROC
19670   That_Erl = Erl
19680   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
19690   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

19700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseMove"

19710   With Me
19720     If blnSelectAll_MouseDown = False Then
19730       Select Case blnSelectAll_Focus
            Case True
19740         .cmdSelectAll_raised_focus_dots_img.Visible = True
19750         .cmdSelectAll_raised_focus_img.Visible = False
19760       Case False
19770         .cmdSelectAll_raised_focus_img.Visible = True
19780         .cmdSelectAll_raised_focus_dots_img.Visible = False
19790       End Select
19800       .cmdSelectAll_raised_img.Visible = False
19810       .cmdSelectAll_raised_semifocus_dots_img.Visible = False
19820       .cmdSelectAll_sunken_focus_dots_img.Visible = False
19830       .cmdSelectAll_raised_img_dis.Visible = False
19840     End If
19850     If .cmdSelectNone_raised_focus_dots_img.Visible = True Or .cmdSelectNone_raised_focus_img.Visible = True Then
19860       Select Case blnSelectNone_Focus
            Case True
19870         .cmdSelectNone_raised_semifocus_dots_img.Visible = True
19880         .cmdSelectNone_raised_img.Visible = False
19890       Case False
19900         .cmdSelectNone_raised_img.Visible = True
19910         .cmdSelectNone_raised_semifocus_dots_img.Visible = False
19920       End Select
19930       .cmdSelectNone_raised_focus_img.Visible = False
19940       .cmdSelectNone_raised_focus_dots_img.Visible = False
19950       .cmdSelectNone_sunken_focus_dots_img.Visible = False
19960       .cmdSelectNone_raised_img_dis.Visible = False
19970     End If
19980   End With

EXITP:
19990   Exit Sub

ERRH:
20000   Select Case ERR.Number
        Case Else
20010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20020   End Select
20030   Resume EXITP

End Sub

Private Sub cmdSelectAll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

20100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_MouseUp"

20110   With Me
20120     .cmdSelectAll_raised_focus_dots_img.Visible = True
20130     .cmdSelectAll_raised_img.Visible = False
20140     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
20150     .cmdSelectAll_raised_focus_img.Visible = False
20160     .cmdSelectAll_sunken_focus_dots_img.Visible = False
20170     .cmdSelectAll_raised_img_dis.Visible = False
20180     blnSelectAll_MouseDown = False
20190   End With

EXITP:
20200   Exit Sub

ERRH:
20210   Select Case ERR.Number
        Case Else
20220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20230   End Select
20240   Resume EXITP

End Sub

Private Sub cmdSelectAll_LostFocus()

20300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectAll_LostFocus"

20310   With Me
20320     .cmdSelectAll_raised_img.Visible = True
20330     .cmdSelectAll_raised_semifocus_dots_img.Visible = False
20340     .cmdSelectAll_raised_focus_img.Visible = False
20350     .cmdSelectAll_raised_focus_dots_img.Visible = False
20360     .cmdSelectAll_sunken_focus_dots_img.Visible = False
20370     .cmdSelectAll_raised_img_dis.Visible = False
20380     blnSelectAll_Focus = False
20390   End With

EXITP:
20400   Exit Sub

ERRH:
20410   Select Case ERR.Number
        Case Else
20420     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20430   End Select
20440   Resume EXITP

End Sub

Private Sub cmdSelectNone_GotFocus()

20500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_GotFocus"

20510   With Me
20520     blnSelectNone_Focus = True
20530     .cmdSelectNone_raised_semifocus_dots_img.Visible = True
20540     .cmdSelectNone_raised_img.Visible = False
20550     .cmdSelectNone_raised_focus_img.Visible = False
20560     .cmdSelectNone_raised_focus_dots_img.Visible = False
20570     .cmdSelectNone_sunken_focus_dots_img.Visible = False
20580     .cmdSelectNone_raised_img_dis.Visible = False
20590   End With

EXITP:
20600   Exit Sub

ERRH:
20610   Select Case ERR.Number
        Case Else
20620     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20630   End Select
20640   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

20700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseDown"

20710   With Me
20720     blnSelectNone_MouseDown = True
20730     .cmdSelectNone_sunken_focus_dots_img.Visible = True
20740     .cmdSelectNone_raised_img.Visible = False
20750     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
20760     .cmdSelectNone_raised_focus_img.Visible = False
20770     .cmdSelectNone_raised_focus_dots_img.Visible = False
20780     .cmdSelectNone_raised_img_dis.Visible = False
20790   End With

EXITP:
20800   Exit Sub

ERRH:
20810   Select Case ERR.Number
        Case Else
20820     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
20830   End Select
20840   Resume EXITP

End Sub

Private Sub cmdSelectNone_Click()

20900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_Click"

        Dim strCtl As String
        Dim lngX As Long

20910   With Me
20920     For lngX = 0& To (lngJTypes - 1&)
20930       strCtl = arr_varJType(JT_CTL, lngX)
20940       .Controls(strCtl) = False
20950       JType_Chk (strCtl & "_AfterUpdate")  ' ** Procedure: Below.
20960       DoEvents
20970     Next
20980   End With

EXITP:
20990   Exit Sub

ERRH:
21000   Select Case ERR.Number
        Case Else
21010     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21020   End Select
21030   Resume EXITP

End Sub

Private Sub cmdSelectNone_KeyDown(KeyCode As Integer, Shift As Integer)

21100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_KeyDown"

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
21190         .cmdPreview.SetFocus
21200       End With
21210     End Select
21220   End If

        ' ** Shift keys.
21230   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
21240     Select Case intRetVal
          Case vbKeyTab
21250       With Me
21260         intRetVal = 0
21270         .cmdSelectAll.SetFocus
21280       End With
21290     End Select
21300   End If

        ' ** Ctrl keys.
21310   If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
21320     Select Case intRetVal
          Case vbKeyTab
21330       With Me
21340         intRetVal = 0
21350         .chkPageOf.SetFocus
21360       End With
21370     End Select
21380   End If

EXITP:
21390   KeyCode = intRetVal
21400   Exit Sub

ERRH:
21410   intRetVal = 0
21420   THAT_PROC = THIS_PROC
21430   That_Erl = Erl
21440   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
21450   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

21500 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseMove"

21510   With Me
21520     If blnSelectNone_MouseDown = False Then
21530       Select Case blnSelectNone_Focus
            Case True
21540         .cmdSelectNone_raised_focus_dots_img.Visible = True
21550         .cmdSelectNone_raised_focus_img.Visible = False
21560       Case False
21570         .cmdSelectNone_raised_focus_img.Visible = True
21580         .cmdSelectNone_raised_focus_dots_img.Visible = False
21590       End Select
21600       .cmdSelectNone_raised_img.Visible = False
21610       .cmdSelectNone_raised_semifocus_dots_img.Visible = False
21620       .cmdSelectNone_sunken_focus_dots_img.Visible = False
21630       .cmdSelectNone_raised_img_dis.Visible = False
21640     End If
21650     If .cmdSelectAll_raised_focus_dots_img.Visible = True Or .cmdSelectAll_raised_focus_img.Visible = True Then
21660       Select Case blnSelectAll_Focus
            Case True
21670         .cmdSelectAll_raised_semifocus_dots_img.Visible = True
21680         .cmdSelectAll_raised_img.Visible = False
21690       Case False
21700         .cmdSelectAll_raised_img.Visible = True
21710         .cmdSelectAll_raised_semifocus_dots_img.Visible = False
21720       End Select
21730       .cmdSelectAll_raised_focus_img.Visible = False
21740       .cmdSelectAll_raised_focus_dots_img.Visible = False
21750       .cmdSelectAll_sunken_focus_dots_img.Visible = False
21760       .cmdSelectAll_raised_img_dis.Visible = False
21770     End If
21780   End With

EXITP:
21790   Exit Sub

ERRH:
21800   Select Case ERR.Number
        Case Else
21810     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
21820   End Select
21830   Resume EXITP

End Sub

Private Sub cmdSelectNone_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

21900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_MouseUp"

21910   With Me
21920     .cmdSelectNone_raised_focus_dots_img.Visible = True
21930     .cmdSelectNone_raised_img.Visible = False
21940     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
21950     .cmdSelectNone_raised_focus_img.Visible = False
21960     .cmdSelectNone_sunken_focus_dots_img.Visible = False
21970     .cmdSelectNone_raised_img_dis.Visible = False
21980     blnSelectNone_MouseDown = False
21990   End With

EXITP:
22000   Exit Sub

ERRH:
22010   Select Case ERR.Number
        Case Else
22020     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22030   End Select
22040   Resume EXITP

End Sub

Private Sub cmdSelectNone_LostFocus()

22100 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSelectNone_LostFocus"

22110   With Me
22120     .cmdSelectNone_raised_img.Visible = True
22130     .cmdSelectNone_raised_semifocus_dots_img.Visible = False
22140     .cmdSelectNone_raised_focus_img.Visible = False
22150     .cmdSelectNone_raised_focus_dots_img.Visible = False
22160     .cmdSelectNone_sunken_focus_dots_img.Visible = False
22170     .cmdSelectNone_raised_img_dis.Visible = False
22180     blnSelectNone_Focus = False
22190   End With

EXITP:
22200   Exit Sub

ERRH:
22210   Select Case ERR.Number
        Case Else
22220     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22230   End Select
22240   Resume EXITP

End Sub

Private Sub chkPageOf_AfterUpdate()

22300 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_AfterUpdate"

22310   With Me
22320     Select Case .chkPageOf
          Case True
22330       .chkPageOf_lbl.FontBold = True
22340     Case False
22350       .chkPageOf_lbl.FontBold = False
22360     End Select
22370   End With

EXITP:
22380   Exit Sub

ERRH:
22390   DoCmd.Hourglass False
22400   Select Case ERR.Number
        Case Else
22410     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22420   End Select
22430   Resume EXITP

End Sub

Private Sub chkPageOf_KeyDown(KeyCode As Integer, Shift As Integer)

22500 On Error GoTo ERRH

        Const THIS_PROC As String = "chkPageOf_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
22570       With Me
22580         intRetVal = 0
22590         .cmdPreview.SetFocus
22600       End With
22610     End Select
22620   End If

        ' ** Shift keys.
22630   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
22640     Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
22650       With Me
22660         intRetVal = 0
22670         .cmdSelectNone.SetFocus
22680       End With
22690     End Select
22700   End If

EXITP:
22710   KeyCode = intRetVal
22720   Exit Sub

ERRH:
22730   intRetVal = 0
22740   THAT_PROC = THIS_PROC
22750   That_Erl = Erl
22760   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
22770   Resume EXITP

End Sub

Private Sub cmdPreview_Click()

22800 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_Click"

22810   ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

22820   DoReport THIS_PROC  ' ** Function: Below.

EXITP:
22830   Exit Sub

ERRH:
22840   Select Case ERR.Number
        Case Else
22850     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
22860   End Select
22870   Resume EXITP

End Sub

Private Sub cmdPreview_KeyDown(KeyCode As Integer, Shift As Integer)

22900 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPreview_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

22910   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
22920   intShiftDown = (Shift And acShiftMask) > 0
22930   intAltDown = (Shift And acAltMask) > 0
22940   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
22950   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
22960     Select Case intRetVal
          Case vbKeyTab
22970       With Me
22980         intRetVal = 0
22990         .cmdPrint.SetFocus
23000       End With
23010     End Select
23020   End If

        ' ** Shift keys.
23030   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23040     Select Case intRetVal
          Case vbKeyTab
23050       With Me
23060         intRetVal = 0
23070         .cmdSelectNone.SetFocus
23080       End With
23090     End Select
23100   End If

        ' ** Ctrl-Shift keys.
23110   If intCtrlDown And (Not intAltDown) And intShiftDown Then
23120     Select Case intRetVal
          Case vbKeyTab
23130       With Me
23140         intRetVal = 0
23150         .chkPageOf.SetFocus
23160       End With
23170     End Select
23180   End If

EXITP:
23190   KeyCode = intRetVal
23200   Exit Sub

ERRH:
23210   intRetVal = 0
23220   THAT_PROC = THIS_PROC
23230   That_Erl = Erl
23240   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23250   Resume EXITP

End Sub

Private Sub cmdPrint_Click()

23300 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_Click"

23310   ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

        '##GTR_Ref: rptTransactions_ByType
23320   DoReport THIS_PROC  ' ** Function: Below.

EXITP:
23330   Exit Sub

ERRH:
23340   Select Case ERR.Number
        Case Else
23350     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23360   End Select
23370   Resume EXITP

End Sub

Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)

23400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrint_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

23410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
23420   intShiftDown = (Shift And acShiftMask) > 0
23430   intAltDown = (Shift And acAltMask) > 0
23440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
23450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
23460     Select Case intRetVal
          Case vbKeyTab
23470       With Me
23480         intRetVal = 0
23490         .cmdWord.SetFocus
23500       End With
23510     End Select
23520   End If

        ' ** Shift keys.
23530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
23540     Select Case intRetVal
          Case vbKeyTab
23550       With Me
23560         intRetVal = 0
23570         .cmdPreview.SetFocus
23580       End With
23590     End Select
23600   End If

EXITP:
23610   KeyCode = intRetVal
23620   Exit Sub

ERRH:
23630   intRetVal = 0
23640   THAT_PROC = THIS_PROC
23650   That_Erl = Erl
23660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
23670   Resume EXITP

End Sub

Private Sub cmdWord_Click()

23700 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_Click"

        Dim strRpt As String, strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim varTmp00 As Variant

23710   ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

23720   With Me
23730     If DoReport(THIS_PROC) = True Then  ' ** Function: Below.

23740       If IsNull(.UserReportPath) = True Then
23750         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
23760       Else
23770         strRptPath = .UserReportPath
23780       End If
23790       strRptCap = "rptTransactions_ByType_"
23800       varTmp00 = .cmbAccounts
23810       varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
23820       varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
23830       strRptCap = strRptCap & varTmp00
23840       strRptCap = strRptCap & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

23850       strRptPathFile = FileSaveDialog("rtf", strRptCap & ".rtf", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

23860       If strRptPathFile <> vbNullString Then
23870         strRpt = "rptTransactions_ByType"
23880         DoCmd.OutputTo acOutputReport, strRpt, acFormatRTF, strRptPathFile, True
23890         .UserReportPath = Parse_Path(strRptPathFile)  ' ** Module Function: modFileUtilities.
23900       End If

23910     End If
23920   End With

EXITP:
23930   Exit Sub

ERRH:
23940   Select Case ERR.Number
        Case Else
23950     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
23960   End Select
23970   Resume EXITP

End Sub

Private Sub cmdWord_KeyDown(KeyCode As Integer, Shift As Integer)

24000 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdWord_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

24010   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
24020   intShiftDown = (Shift And acShiftMask) > 0
24030   intAltDown = (Shift And acAltMask) > 0
24040   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
24050   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
24060     Select Case intRetVal
          Case vbKeyTab
24070       With Me
24080         intRetVal = 0
24090         If .cmdExcel.Enabled = True Then
24100           .cmdExcel.SetFocus
24110         Else
24120           .cmdClose.SetFocus
24130         End If
24140       End With
24150     End Select
24160   End If

        ' ** Shift keys.
24170   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
24180     Select Case intRetVal
          Case vbKeyTab
24190       With Me
24200         intRetVal = 0
24210         .cmdPrint.SetFocus
24220       End With
24230     End Select
24240   End If

EXITP:
24250   KeyCode = intRetVal
24260   Exit Sub

ERRH:
24270   intRetVal = 0
24280   THAT_PROC = THIS_PROC
24290   That_Erl = Erl
24300   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
24310   Resume EXITP

End Sub

Private Sub cmdExcel_Click()

24400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strRptCap As String
        Dim strRptPath As String, strRptPathFile As String
        Dim strLastAcctNo As String, strLastJType As String, strPeriod As String
        Dim lngJTypes As Long, arr_varJType() As Variant
        Dim dblICash As Double, dblPCash As Double, dblCost As Double
        Dim dblICashTot As Double, dblPCashTot As Double, dblCostTot As Double
        Dim lngRecs As Long
        Dim blnFound As Boolean
        Dim varTmp00 As Variant, strTmp01 As String
        Dim lngX As Long, lngE As Long

        ' ** Array: arr_varJType().
        Const J_ELEMS As Integer = 4  ' ** Array's first-element UBound().
        Const J_TYP   As Integer = 0
        Const J_ORD   As Integer = 1
        Const J_ICASH As Integer = 2
        Const J_PCASH As Integer = 3
        Const J_COST  As Integer = 4

24410   DoCmd.Hourglass True
24420   DoEvents

24430   ChkSpecLedgerEntry  ' ** Module Function: modUtilities.

24440   With Me
24450     If DoReport(THIS_PROC) = True Then  ' ** Function: Below.

24460       If IsNull(.UserReportPath) = True Then
24470         strRptPath = CurrentAppPath  ' ** Module Function: modFileUtilities.
24480       Else
24490         strRptPath = .UserReportPath
24500       End If

24510       Set dbs = CurrentDb

            ' ** Empty tmpTransactionsByType.
24520       Set qdf = dbs.QueryDefs("qryTransactions_ByType_07")
24530       qdf.Execute

24540       Select Case .chkIncludeArchive
            Case True
              ' ** Append qryTransactions_ByType_06_02 (qryTransactions_ByType_04_02 (qryTransactions_ByType_03
              ' ** (Union of qryTransactions_ByType_01 (Ledger, just needed fields, with journaltypex, by
              ' ** specified FormRef('accountno')), qryTransactions_ByType_02 (LedgerArchive, just needed fields,
              ' ** with journaltypex, by specified FormRef('accountno'))), with add'l fields), by journaltype
              ' ** specified FormRef()) to tmpTransactionsByType.
24550         Set qdf = dbs.QueryDefs("qryTransactions_ByType_08_02")
24560       Case False
              ' ** Append qryTransactions_ByType_06_01 (qryTransactions_ByType_04_01 (qryTransactions_ByType_01
              ' ** (Ledger, just needed fields, with journaltypex, by specified FormRef('accountno')), with
              ' ** add'l fields), by journaltype specified FormRef()) to tmpTransactionsByType.
24570         Set qdf = dbs.QueryDefs("qryTransactions_ByType_08_01")
24580       End Select
24590       qdf.Execute

            ' ** tmpTransactionsByType, sorted.
24600       Set qdf = dbs.QueryDefs("qryTransactions_ByType_09")
24610       Set rst = qdf.OpenRecordset
24620       With rst
24630         .MoveLast
24640         lngRecs = .RecordCount
24650         .MoveFirst

24660         lngJTypes = 0&
24670         ReDim arr_varJType(J_ELEMS, 0)

24680         strLastAcctNo = vbNullString: strLastJType = vbNullString
24690         dblICash = 0#: dblPCash = 0#: dblCost = 0#
24700         dblICashTot = 0#: dblPCashTot = 0#: dblCostTot = 0#
24710         For lngX = 1& To lngRecs
24720           .Edit
24730           If strLastAcctNo = vbNullString Then
24740             ![first_rec] = True
24750           End If
24760           If ![accountno] <> strLastAcctNo Then  ' ** Though there should always only be 1.
24770             strLastAcctNo = ![accountno]
24780             ![sort2] = 1
24790           Else
24800             ![Account No] = vbNullString  ' ** [accountno] always retains value.
24810             ![Account Name] = vbNullString
24820             ![sort2] = 2
24830           End If
24840           If ![journaltype] <> strLastJType Then
24850             dblICashTot = (dblICashTot + dblICash)
24860             dblPCashTot = (dblPCashTot + dblPCash)
24870             dblCostTot = (dblCostTot + dblCost)
24880             If lngJTypes > 0& Then
24890               arr_varJType(J_ICASH, (lngJTypes - 1&)) = dblICash
24900               arr_varJType(J_PCASH, (lngJTypes - 1&)) = dblPCash
24910               arr_varJType(J_COST, (lngJTypes - 1&)) = dblCost
24920             End If
24930             dblICash = 0#: dblPCash = 0#: dblCost = 0#
24940             strLastJType = ![journaltype]
24950             ![first_grp] = True
24960             lngJTypes = lngJTypes + 1&
24970             lngE = lngJTypes - 1&
24980             ReDim Preserve arr_varJType(J_ELEMS, lngE)
24990             arr_varJType(J_TYP, lngE) = strLastJType
25000             arr_varJType(J_ORD, lngE) = ![JournalType_Order]
25010             arr_varJType(J_ICASH, lngE) = CDbl(0)
25020             arr_varJType(J_PCASH, lngE) = CDbl(0)
25030             arr_varJType(J_COST, lngE) = CDbl(0)
25040           Else
25050             ![Transaction Type] = vbNullString
25060           End If
25070           dblICash = (dblICash + ![Income Cash])
25080           dblPCash = (dblPCash + ![Principal Cash])
25090           dblCost = (dblCost + ![Cost])
25100           .Update
25110           If lngX = lngRecs Then
25120             dblICashTot = (dblICashTot + dblICash)
25130             dblPCashTot = (dblPCashTot + dblPCash)
25140             dblCostTot = (dblCostTot + dblCost)
25150             arr_varJType(J_ICASH, (lngJTypes - 1&)) = dblICash
25160             arr_varJType(J_PCASH, (lngJTypes - 1&)) = dblPCash
25170             arr_varJType(J_COST, (lngJTypes - 1&)) = dblCost
25180           End If
25190           If lngX < lngRecs Then .MoveNext
25200         Next

              ' ** Add the JournalType totals.
25210         For lngX = 0& To (lngJTypes - 1&)
25220           .AddNew
25230           ![Account No] = vbNullString
25240           ![Account Name] = vbNullString
25250           ![Transaction Type] = vbNullString
25260           ![description] = Space(8) & arr_varJType(J_TYP, lngX) & " Total"
25270           ![Income Cash] = arr_varJType(J_ICASH, lngX)
25280           ![Principal Cash] = arr_varJType(J_PCASH, lngX)
25290           ![Cost] = arr_varJType(J_COST, lngX)
25300           ![JournalType_Order] = arr_varJType(J_ORD, lngX)
25310           ![accountno] = strLastAcctNo
25320           ![journaltype] = arr_varJType(J_TYP, lngX)
25330           ![sort1] = 2
25340           ![sort2] = 3
25350           .Update
                ' ** Add a blank line.
25360           .AddNew
25370           ![Account No] = vbNullString
25380           ![Account Name] = vbNullString
25390           ![Transaction Type] = vbNullString
25400           ![description] = vbNullString
25410           ![Income Cash] = Null
25420           ![Principal Cash] = Null
25430           ![Cost] = Null
25440           ![JournalType_Order] = arr_varJType(J_ORD, lngX)
25450           ![accountno] = strLastAcctNo
25460           ![journaltype] = arr_varJType(J_TYP, lngX)
25470           ![sort1] = 3
25480           ![sort2] = 3
25490           .Update
25500         Next

              ' ** Add the grand total.
25510         .AddNew
25520         ![Account No] = vbNullString
25530         ![Account Name] = vbNullString
25540         ![Transaction Type] = vbNullString
25550         ![description] = Space(16) & "Total"
25560         ![Income Cash] = dblICashTot
25570         ![Principal Cash] = dblPCashTot
25580         ![Cost] = dblCostTot
25590         ![JournalType_Order] = 99
25600         ![accountno] = strLastAcctNo
25610         ![journaltype] = "zzzz"
25620         ![sort1] = 3
25630         ![sort2] = 4
25640         .Update

              ' ** Add the report title.
25650         .AddNew
25660         ![Account No] = "Title:"
25670         ![Account Name] = "Transactions By Type"
25680         ![Transaction Type] = vbNullString
25690         ![Trade Date] = Now()  ' ** To make sure Excel understands the column.
25700         ![description] = vbNullString
25710         ![Due Date] = Date  ' ** To make sure Excel understands the column.
25720         ![Income Cash] = Null
25730         ![Principal Cash] = Null
25740         ![Cost] = Null
25750         ![JournalType_Order] = 0
25760         ![accountno] = strLastAcctNo
25770         ![journaltype] = Null
25780         ![sort1] = -2
25790         ![sort2] = 0
25800         .Update

25810         Select Case Me.chkTransDateAll
              Case True
25820           strPeriod = "All"
25830         Case False
25840           strPeriod = Format(Me.TransDateStart, "mm/dd/yyyy") & " To " & Format(Me.TransDateEnd, "mm/dd/yyyy")
25850         End Select

              ' ** Add the report period.
25860         .AddNew
25870         ![Account No] = "Period:"
25880         ![Account Name] = strPeriod
25890         ![Transaction Type] = vbNullString
25900         ![description] = vbNullString
25910         ![Income Cash] = Null
25920         ![Principal Cash] = Null
25930         ![Cost] = Null
25940         ![JournalType_Order] = 0
25950         ![accountno] = strLastAcctNo
25960         ![journaltype] = Null
25970         ![sort1] = -1
25980         ![sort2] = 0
25990         .Update

26000         .Close
26010       End With

26020       dbs.Close

26030       strRptCap = "rptTransactions_ByType_"
26040       varTmp00 = .cmbAccounts
26050       varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
26060       varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
26070       strRptCap = strRptCap & varTmp00
26080       strRptCap = strRptCap & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")

26090       strRptPathFile = FileSaveDialog("xls", strRptCap & ".xls", strRptPath, "Save File")  ' ** Module Function: modBrowseFilesAndFolders.

26100       If strRptPathFile <> vbNullString Then

              ' ** This is the only way to get Microsoft Excel 2003 format via OutputTo method.
              ' ** (And OutputTo results in a much better looking spreadsheet than TransferSpreadsheet!)
26110         DoCmd.RunMacro "mcrExcelExport_TransByType"
              'DoCmd.OutputTo acOutputQuery, "qryTransactions_ByType_10", "MicrosoftExcelBiff8 (*.xls)", (CurrentAppPath & LNK_SEP & strFileName), True  ' ** Module Function: modFileUtilities."
              'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qryTransactions_ByType_10", (CurrentAppPath & LNK_SEP & strFileName), False, , True

              ' ** The macro specifies qryTransactions_ByType_10, but cannot be given a dynamic file name.
              ' ** So, it's exported to 'Transactions_By_Type_xxx.xls', which is then renamed.
26120         blnFound = False
26130         If FileExists(CurrentAppPath & LNK_SEP & "Transactions_By_Type_xxx.xls") = True Then  ' ** Module Function: modFileUtilities.
26140           blnFound = True
26150           strTmp01 = CurrentAppPath & LNK_SEP & "Transactions_By_Type_xxx.xls"  ' ** Module Function: modFileUtilities.
26160         ElseIf FileExists(Parse_Path(strRptPathFile) & LNK_SEP & "Transactions_By_Type_xxx.xls") = True Then   ' ** Module Function: modFileUtilities.
26170           blnFound = True
26180           strTmp01 = Parse_Path(strRptPathFile) & LNK_SEP & "Transactions_By_Type_xxx.xls"  ' ** Module Function: modFileUtilities.
26190         End If
26200         If blnFound = True Then
26210           Name (strTmp01) As (strRptPathFile)
                ' ** Because the file must be renamed, AutoStart is set to 'No' in the macro.
26220           OpenExe strRptPathFile  ' ** Module Function: modShellFuncs.
26230         End If

26240       End If

            ' ** Though the Argument in the macro says 'MicrosoftExcelBiff8(*.xls)',
            ' ** all of these variations for the OutputFormat parameter result in
            ' ** Excel saying 'Microsoft Excel 5.0/95 Workbook' when it's saved!
            ' **   MicrosoftExcelBiff8(*.xls)
            ' **   Microsoft Excel Biff8 (*.xls)
            ' **   Microsoft Excel 97-2003 (*.xls)
            ' **   MicrosoftExcel97-2003(*.xls)
            ' **   Microsoft Excel (*.xls)
            ' **   acFormatXLS
            ' **   Microsoft Excel 8-10 (*.xls)
            ' **   Microsoft Excel 97-2003
            ' ** The macro, however, saves without a message.
            ' ** BIFF: Binary Interchange File Format.

26250     End If  ' ** DoReport().
26260   End With

26270   DoCmd.Hourglass False

EXITP:
26280   Set rst = Nothing
26290   Set qdf = Nothing
26300   Set rst = Nothing
26310   Exit Sub

ERRH:
26320   DoCmd.Hourglass False
26330   Select Case ERR.Number
        Case Else
26340     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
26350   End Select
26360   Resume EXITP

End Sub

Private Sub cmdExcel_KeyDown(KeyCode As Integer, Shift As Integer)

26400 On Error GoTo ERRH

        Const THIS_PROC As String = "cmdExcel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

26410   intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
26420   intShiftDown = (Shift And acShiftMask) > 0
26430   intAltDown = (Shift And acAltMask) > 0
26440   intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
26450   If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
26460     Select Case intRetVal
          Case vbKeyTab
26470       With Me
26480         intRetVal = 0
26490         .cmdClose.SetFocus
26500       End With
26510     End Select
26520   End If

        ' ** Shift keys.
26530   If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
26540     Select Case intRetVal
          Case vbKeyTab
26550       With Me
26560         intRetVal = 0
26570         .cmdWord.SetFocus
26580       End With
26590     End Select
26600   End If

EXITP:
26610   KeyCode = intRetVal
26620   Exit Sub

ERRH:
26630   intRetVal = 0
26640   THAT_PROC = THIS_PROC
26650   That_Erl = Erl
26660   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
26670   Resume EXITP

End Sub

Public Function JType_Get(strJType As String) As String

26700 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_Get"

        Dim lngX As Long
        Dim strRetVal As String

26710   strRetVal = vbNullString

        ' ************************************************************
        ' ** Array: arr_varJType()
        ' **
        ' **   Field  Element  Name                       Constant
        ' **   =====  =======  =========================  ==========
        ' **     1       0     JournalType                JT_TYPE
        ' **     2       1     JournalType_Description    JT_DESC
        ' **     3       2     JournalType_Order          JT_SORT
        ' **     4       3     ctl_name                   JT_CTL
        ' **     5       4     ctl_label                  JT_LBL
        ' **     6       5     ctl_value                  JT_VAL
        ' **     7       6     ctl_letter                 JT_LTR
        ' **     8       7     ctl_box                    JT_BOX
        ' **
        ' ************************************************************

26720   If strJType <> vbNullString Then
26730     For lngX = 0& To (lngJTypes - 1&)
26740       If arr_varJType(JT_TYPE, lngX) = strJType Then
26750         strRetVal = arr_varJType(JT_DESC, lngX)
26760         Exit For
26770       End If
26780     Next
26790   End If

EXITP:
26800   JType_Get = strRetVal
26810   Exit Function

ERRH:
26820   THAT_PROC = THIS_PROC
26830   That_Erl = Erl
26840   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
26850   Resume EXITP

End Function

Private Sub JType_Chk(strProc As String)

26900 On Error GoTo ERRH

        Const THIS_PROC As String = "JType_Chk"

        Dim strControl As String

26910   With Me
26920     strControl = Left(strProc, (Len(strProc) - Len("_AfterUpdate")))  ' ** JType_01_chk_AfterUpdate.
26930     Select Case .Controls(strControl)
          Case True
            '.Controls(Left(strControl, (Len(strControl) - 4)) & "_lbl").BackStyle = acBackStyleNormal
26940       .Controls(Left(strControl, (Len(strControl) - 4)) & "_lbl_box").Visible = True
26950       .Controls(Left(strControl, (Len(strControl) - 4)) & "_lbl").FontBold = True
26960     Case False
            '.Controls(Left(strControl, (Len(strControl) - 4)) & "_lbl").BackStyle = acBackStyleTransparent
26970       .Controls(Left(strControl, (Len(strControl) - 4)) & "_lbl_box").Visible = False
26980       .Controls(Left(strControl, (Len(strControl) - 4)) & "_lbl").FontBold = False
26990     End Select
27000   End With

EXITP:
27010   Exit Sub

ERRH:
27020   THAT_PROC = THIS_PROC
27030   That_Erl = Erl
27040   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Below.
27050   Resume EXITP

End Sub

Private Function DoReport(strProc As String) As Boolean

27100 On Error GoTo ERRH

        Const THIS_PROC As String = "DoReport"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim strMode As String
        Dim strFilter As String
        Dim blnFound As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long
        Dim blnRetVal As Boolean

27110   blnRetVal = True

27120   With Me

27130     strMode = Mid(Left(strProc, (InStr(strProc, "_") - 1)), 4)
27140     .ReportMode = strMode
27150     DoEvents

27160     Select Case .chkTransDateAll
          Case True
            ' ** No dates required.
27170     Case False
27180       If IsNull(.TransDateStart) Then
27190         blnRetVal = False
27200         MsgBox "Must enter the start date to continue.", vbInformation + vbOKOnly, "Entry Required"
27210         .TransDateStart.SetFocus
27220       Else
27230         If IsNull(.TransDateEnd) Then
27240           blnRetVal = False
27250           MsgBox "Must enter the end date to continue.", vbInformation + vbOKOnly, "Entry Required"
27260           .TransDateStart.SetFocus
27270         Else
27280           gdatStartDate = .TransDateStart
27290           gdatEndDate = .TransDateEnd
27300         End If
27310       End If
27320     End Select

27330     If blnRetVal = True Then
27340       blnFound = False
27350       For lngX = 0& To (lngJTypes - 1&)
27360         arr_varJType(JT_VAL, lngX) = .Controls(arr_varJType(JT_CTL, lngX))
27370         If arr_varJType(JT_VAL, lngX) = True Then
27380           blnFound = True
27390         End If
27400       Next
27410       If blnFound = False Then
27420         blnRetVal = False
27430         MsgBox "Must choose at least 1 Journal Type to continue.", vbInformation + vbOKOnly, "Entry Required"
27440         .JType_01_chk.SetFocus
27450       End If
27460     End If  ' ** blnRetVal.

          'SHOULD THIS HAVE AN 'ALL'?
27470     If blnRetVal = True Then
27480       .cmbAccounts.SetFocus
27490       If .cmbAccounts.text = vbNullString Then
27500         blnRetVal = False
27510         MsgBox "Must select an account to continue.", vbInformation + vbOKOnly, "Entry Required"
27520       End If
27530     End If  ' ** blnRetVal.

27540     If blnRetVal = True Then
27550       Set dbs = CurrentDb
27560       Select Case .chkIncludeArchive
            Case True
              ' ** qryTransactions_ByType_03 (Union of qryTransactions_ByType_01 (Ledger,
              ' ** just needed fields, with journaltypex, by specified FormRef('accountno')),
              ' ** qryTransactions_ByType_02 (LedgerArchive, just needed fields, with
              ' ** journaltypex, by specified FormRef('accountno'))), with add'l fields.
27570         Set qdf = dbs.QueryDefs("qryTransactions_ByType_04_02")
27580       Case False
              ' ** qryTransactions_ByType_01 (Ledger, just needed fields, with journaltypex,
              ' ** by specified FormRef('accountno')), with add'l fields.
27590         Set qdf = dbs.QueryDefs("qryTransactions_ByType_04_01")
27600       End Select
27610       Set rst1 = qdf.OpenRecordset
27620       With rst1
27630         If .BOF = True And .EOF = True Then
27640           blnRetVal = False
27650           MsgBox "There are no transactions for this account.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
27660         Else
27670           strFilter = "[journaltypex] In ("
27680           For lngX = 0& To (lngJTypes - 1&)
27690             If arr_varJType(JT_VAL, lngX) = True Then
27700               strFilter = strFilter & "'" & arr_varJType(JT_TYPE, lngX) & "',"
27710             End If
27720           Next
27730           If Right(strFilter, 1) = "," Then strFilter = Left(strFilter, (Len(strFilter) - 1))
27740           strFilter = strFilter & ")"
27750           Select Case Me.chkTransDateAll
                Case True
27760             .Filter = strFilter
27770             Set rst2 = rst1.OpenRecordset
27780             With rst2
27790               If .BOF = True And .EOF = True Then
27800                 blnRetVal = False
27810                 MsgBox "There are no transactions for the Journal Types specified.", _
                        vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
27820               Else
27830                 .MoveFirst
27840                 gdatStartDate = ![transdate]
27850                 .MoveLast
27860                 gdatEndDate = ![transdate]
27870               End If
27880               .Close
27890             End With
27900           Case False
27910             strFilter = strFilter & " And ([transdate] >= #" & Format(gdatStartDate, "mm/dd/yyyy") & "# And " & _
                    "[transdate] <= #" & Format(gdatEndDate, "mm/dd/yyyy") & "#)"
27920             .Filter = strFilter
27930             Set rst2 = .OpenRecordset
27940             With rst2
27950               If .BOF = True And .EOF = True Then
27960                 blnRetVal = False
27970                 MsgBox "There are no transactions for the dates and/or Journal Types specified.", _
                        vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
27980               End If
27990               .Close
28000             End With
28010           End Select
28020         End If
28030         .Close
28040       End With
28050       dbs.Close
28060     End If  ' ** blnRetVal.

28070     If blnRetVal = True Then
28080       strDocName = "rptTransactions_ByType"
28090       varTmp00 = .cmbAccounts
28100       varTmp00 = SpaceToUnder(varTmp00)  ' ** Module Function: modStringFuncs.
28110       varTmp00 = StringReplace(CStr(varTmp00), "/", "_")  ' ** Module Function: modStringFuncs.
28120       strFileName = "Transactions_By_Type_" & varTmp00
28130       strFileName = strFileName & "_" & Format(gdatStartDate, "yymmdd") & "_To_" & Format(gdatEndDate, "yymmdd")
28140       Select Case strMode
            Case "Preview"
28150         DoCmd.OpenReport strDocName, acViewPreview, , strFilter
28160         DoCmd.Maximize
28170         DoCmd.RunCommand acCmdFitToWindow
28180       Case "Print"
28190         If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
28200           DoCmd.OpenReport strDocName, acViewPreview, , strFilter
28210           DoCmd.Maximize
28220           DoCmd.RunCommand acCmdFitToWindow
28230         Else
28240           DoCmd.OpenReport strDocName, acViewNormal, , strFilter
28250         End If
28260       Case "Word"
              ' ** Back to OnClick event.
28270       Case "Excel"
              ' ** Back to OnClick event.
28280       End Select
28290     End If  ' ** blnRetVal.

28300   End With  ' ** Me.

EXITP:
28310   Set rst1 = Nothing
28320   Set rst2 = Nothing
28330   Set qdf = Nothing
28340   Set dbs = Nothing
28350   DoReport = blnRetVal
28360   Exit Function

ERRH:
28370   blnRetVal = False
28380   THAT_PROC = THIS_PROC
28390   That_Erl = Erl
28400   Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
28410   Resume EXITP

End Function

Private Sub ArchiveSet()

28500 On Error GoTo ERRH

        Const THIS_PROC As String = "ArchiveSet"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim blnHasArchs As Boolean

28510   Set dbs = CurrentDb
28520   With dbs
28530     Set rst = .OpenRecordset("LedgerArchive", dbOpenDynaset, dbReadOnly)
28540     With rst
28550       If .BOF = True And .EOF = True Then
28560         blnHasArchs = False
28570       Else
28580         .MoveLast
28590         If .RecordCount = 1 Then
28600           blnHasArchs = False
28610         Else
28620           blnHasArchs = True
28630         End If
28640       End If
28650       .Close
28660     End With
28670     Set rst = Nothing
28680     .Close
28690   End With
28700   Set dbs = Nothing

28710   With Me
28720     Select Case blnHasArchs
          Case True
28730       .chkIncludeArchive_lbl.Visible = True
28740       .chkIncludeArchive_lbl2.Visible = False
28750       .chkIncludeArchive_lbl2_dim_hi.Visible = False
28760     Case False
28770       .chkIncludeArchive = False
28780       .chkIncludeArchive.Enabled = False
28790       .chkIncludeArchive_lbl.Visible = False
28800       .chkIncludeArchive_lbl2.Visible = True
28810       .chkIncludeArchive_lbl2_dim_hi.Visible = True
28820     End Select
28830   End With

EXITP:
28840   Set rst = Nothing
28850   Set dbs = Nothing
28860   Exit Sub

ERRH:
28870   Select Case ERR.Number
        Case Else
28880     zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
28890   End Select
28900   Resume EXITP

End Sub
