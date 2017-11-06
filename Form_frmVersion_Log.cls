VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmVersion_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmVersion_Log"

'VGC 12/01/2015: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Print Report:     P {cmdPrintReport}
' **   Summary:          S {Status3}
' **   Exit:             X {cmdClose}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean
Private lngVerCnvID As Long
Private strLogPathFile As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           intPos01 = InStr(.OpenArgs, "~")
150           If intPos01 > 0 Then

160             strLogPathFile = Mid(.OpenArgs, (intPos01 + 1))
                'strLogPathFile = "C:\VictorGCS_Clients\TrustAccountant\NewWorking\Convert_New\ConvLog.txt"
170             strCallingForm = Left(.OpenArgs, (intPos01 - 1))

180             blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

190           Else
200             Cancel = -1
210           End If
220         Else
230           Cancel = -1
240         End If
250       Else
260         Cancel = -1
270       End If
280     End With

290     If Cancel = -1 Then
300       Beep
310       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
320     End If

EXITP:
330     Exit Sub

ERRH:
340     Select Case ERR.Number
        Case Else
350       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
360     End Select
370     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim datVer_Date As Date, strVer_Old As String, strVer_New As String, strVer_CoName As String
        Dim strLog As String
        Dim strInput As String
        Dim lngRecs As Long
        Dim intPos01 As Integer
        Dim lngX As Long
        Dim blnFound As Boolean

410     blnFound = False
420     strLog = vbNullString
430     lngVerCnvID = 0&

440     Set dbs = CurrentDb
450     With dbs
460       Set rst = .OpenRecordset("tblVersion_Conversion", dbOpenDynaset, dbReadOnly)
470       With rst
480         If .BOF = True And .EOF = True Then
              ' ** Get the file.
490         Else
500           .MoveLast
510           lngRecs = .RecordCount
520           .MoveFirst
530           For lngX = 1& To lngRecs  ' ** This may only ever have 1 record!
540             If IsNull(![vercnv_name]) = False Then
550               If IsNull(![vercnv_note]) = False Then
560                 If Left(![vercnv_note], 27) = "Trust Accountant Conversion" Then
570                   blnFound = True
580                   strLog = ![vercnv_note]
590                   lngVerCnvID = ![vercnv_id]
600                   Exit For
610                 End If
620               End If
630             End If
640             If lngX < lngRecs Then .MoveNext
650           Next
660         End If
670         .Close
680       End With
690       Set rst = Nothing
700       .Close
710     End With
720     Set dbs = Nothing

730     If blnFound = False Then
740       If FileExists(strLogPathFile) = True Then  ' ** Module Function: modFileUtilities.
750         Open strLogPathFile For Input As #1
760         Do While Not EOF(1)
770           Line Input #1, strInput
780           If Left(strInput, 7) = "Version" Then
790             intPos01 = InStr(strInput, " to ")
800             If intPos01 > 0 Then
810               strVer_New = Trim(Mid(strInput, (intPos01 + 3)))
820               strVer_Old = Trim(Mid(Left(strInput, intPos01), 8))
830             End If
840           ElseIf Right(strInput, 3) = " AM" Or Right(strInput, 3) = " PM" And strVer_CoName = vbNullString Then
850             If IsDate(strInput) = True Then
860               datVer_Date = CDate(strInput)
870             End If
880           ElseIf Left(strInput, 8) = "Company:" Then
890             strVer_CoName = Trim(Mid(strInput, 9))
900           End If
910           strLog = strLog & strInput & vbCrLf
920         Loop
930         Close #1
940         DoEvents
950         If datVer_Date = 0 Then datVer_Date = Now()
960         If strVer_CoName = vbNullString Then strVer_CoName = gstrCo_Name
970         If strVer_Old = vbNullString Then strVer_Old = "Unkown"
980         If strVer_New = vbNullString Then strVer_New = "Unkown"
990         Set dbs = CurrentDb
1000        With dbs
1010          Set rst = .OpenRecordset("tblVersion_Conversion", dbOpenDynaset, dbConsistent)
1020          With rst
1030            If lngRecs > 0& Then
1040              .MoveFirst
1050              lngVerCnvID = ![vercnv_id]
1060              .Edit
1070              ![vercnv_note] = strLog
1080              ![vercnv_datemodified] = Now()
1090              .Update
1100            Else
1110              .AddNew
1120              ![vercnv_date] = datVer_Date
1130              ![vercnv_name] = strVer_CoName
1140              ![vercnv_verold] = strVer_Old
1150              ![vercnv_vernew] = strVer_New
1160              ![vercnv_step] = 23&  ' ** Seems to indicate a successful conversion.
1170              ![vercnv_cancel] = False  ' ** I'm not going to see if the note lists an error.
1180              ![vercnv_error] = False
1190              ![vercnv_note] = strLog
1200              ![Username] = CurrentUser  ' ** Internal Access Function: Trust Accountant login.
1210              ![vercnv_user] = GetUserName  ' ** Module Function: modFileUtilities.
1220              ![vercnv_datemodified] = Now()
1230              .Update
1240              .Bookmark = .LastModified
1250              lngVerCnvID = ![vercnv_id]
1260            End If
                ' ** Trust Accountant Conversion
                ' ** Version 2.2.10 to 2.2.20
                ' ** 07/20/2012 02:48:05 PM
                ' ** ========================================
                ' ** Company: Mound City Bank
                ' ** Accounts: 60
                ' ** Ledger Entries: 14786
                ' ** Archived Ledger Entries: 0
                ' ** Tax Lots: 1849
                ' ** Master Assets: 526
                ' ** Income/Expense Codes: 4
                ' ** Statement Dates:
                ' **   Earliest: 06/30/2002
                ' **   Latest:   06/30/2012
1270            .Close
1280          End With
1290          Set rst = Nothing
1300          .Close
1310        End With
1320        Set dbs = Nothing
1330      End If
1340    End If

1350    With Me
1360      .vercnv_id = lngVerCnvID
1370      .Status3 = strLog
1380      .Status3.SetFocus
1390      .Status3.SelLength = 0
1400      .Status3.SelStart = 0
1410    End With

EXITP:
1420    Set rst = Nothing
1430    Set dbs = Nothing
1440    Exit Sub

ERRH:
1450    Close #1
1460    Select Case ERR.Number
        Case Else
1470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1480    End Select
1490    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1520    intShiftDown = (Shift And acShiftMask) > 0
1530    intAltDown = (Shift And acAltMask) > 0
1540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyEscape
1570        intRetVal = 0
1580        cmdClose_Click  ' ** Procedure: Below.
1590      End Select
1600    End If

        ' ** Alt keys.
1610    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1620      Select Case intRetVal
          Case vbKeyX
1630        intRetVal = 0
1640        cmdClose_Click  ' ** Procedure: Below.
1650      End Select
1660    End If

EXITP:
1670    KeyCode = intRetVal
1680    Exit Sub

ERRH:
1690    intRetVal = 0
1700    Select Case ERR.Number
        Case Else
1710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1720    End Select
1730    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1810    With Me
1820      .TimerInterval = 0&
1830      If gblnSetFocus = True Then
1840        gblnSetFocus = False
1850        strFormName = vbNullString
1860  On Error Resume Next
1870        Do Until strFormName = THIS_NAME
1880          DoCmd.SelectObject acForm, THIS_NAME, False
1890          strFormName = Screen.ActiveForm.Name
1900        Loop
1910  On Error GoTo ERRH
1920      End If  ' ** gblnSetFocus
1930      If gblnGoToReport = True Then
1940        DoCmd.Hourglass True  ' ** Make sure it's still running.
1950        DoEvents
1960        Select Case blnGoingToReport
            Case True
1970          gblnGoToReport = False
1980          blnGoingToReport = False
1990          .GoToReport_arw_pr_img.Visible = False
2000          .cmdClose.SetFocus
2010          DoCmd.Hourglass False
2020        Case False
2030          blnGoingToReport = True
2040          .GoToReport_arw_pr_img.Visible = True
2050          .cmdPrintReport.SetFocus
2060          Beep
2070          DoEvents
2080          .TimerInterval = (GTR_WAIT * 2&)
2090        End Select
2100      End If
2110    End With

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

2210    If Reports.Count > 0 Then
2220      Do While Reports.Count > 0
2230        DoCmd.Close acReport, Reports(0).Name
2240      Loop
2250    End If

2260    If strCallingForm <> vbNullString Then
2270      gstrReturningForm = THIS_NAME
2280      DoCmd.SelectObject acForm, strCallingForm, False
2290    End If

EXITP:
2300    Exit Sub

ERRH:
2310    Select Case ERR.Number
        Case 2489  ' ** The object '|' isn't open.
          ' ** Ignore; may hit this when I close the window before the Timer stuff plays out.
2320    Case Else
2330      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2340    End Select
2350    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2410    With Me
2420      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2430        Select Case blnPrintReport_Focus
            Case True
2440          .cmdPrintReport_raised_focus_dots_img.Visible = True
2450          .cmdPrintReport_raised_img.Visible = False
2460        Case False
2470          .cmdPrintReport_raised_img.Visible = True
2480          .cmdPrintReport_raised_focus_dots_img.Visible = False
2490        End Select
2500        .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2510        .cmdPrintReport_raised_focus_img.Visible = False
2520        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2530        .cmdPrintReport_raised_img_dis.Visible = False
2540      End If
2550    End With

EXITP:
2560    Exit Sub

ERRH:
2570    Select Case ERR.Number
        Case Else
2580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2590    End Select
2600    Resume EXITP

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

2810    DoCmd.Close acForm, THIS_NAME

EXITP:
2820    Exit Sub

ERRH:
2830    Select Case ERR.Number
        Case Else
2840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2850    End Select
2860    Resume EXITP

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
2990          .Status3.SetFocus
3000        End With
3010      End Select
3020    End If

        ' ** Shift keys.
3030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3040      Select Case intRetVal
          Case vbKeyTab
3050        With Me
3060          intRetVal = 0
3070          If .cmdPrintReport.Enabled = True Then
3080            .cmdPrintReport.SetFocus
3090          Else
3100            .Status3.SetFocus
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

Private Sub cmdPrintReport_GotFocus()

3300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3310    With Me
3320      blnPrintReport_Focus = True
3330      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3340      .cmdPrintReport_raised_img.Visible = False
3350      .cmdPrintReport_raised_focus_img.Visible = False
3360      .cmdPrintReport_raised_focus_dots_img.Visible = False
3370      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3380      .cmdPrintReport_raised_img_dis.Visible = False
3390    End With

EXITP:
3400    Exit Sub

ERRH:
3410    Select Case ERR.Number
        Case Else
3420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3430    End Select
3440    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3510    With Me
3520      blnPrintReport_MouseDown = True
3530      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3540      .cmdPrintReport_raised_img.Visible = False
3550      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3560      .cmdPrintReport_raised_focus_img.Visible = False
3570      .cmdPrintReport_raised_focus_dots_img.Visible = False
3580      .cmdPrintReport_raised_img_dis.Visible = False
3590    End With

EXITP:
3600    Exit Sub

ERRH:
3610    Select Case ERR.Number
        Case Else
3620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3630    End Select
3640    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

3710    With Me

3720      strDocName = "rptVersion_Convert"
3730      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
3740        DoCmd.OpenReport strDocName, acViewPreview, , "[vercnv_id] = " & CStr(lngVerCnvID), , "Preview"
3750        DoCmd.Maximize
3760        DoCmd.RunCommand acCmdFitToWindow
3770      Else
            '##GTR_Ref: rptVersion_Convert
3780        DoCmd.OpenReport strDocName, acViewNormal, , "[vercnv_id] = " & CStr(lngVerCnvID), , "Normal"
3790      End If
3800    End With

EXITP:
3810    Exit Sub

ERRH:
3820    DoCmd.Restore
3830    strDocName = "frmMenu_Report"
3840    If IsLoaded(strDocName, acForm) = True Then  ' ** Module Function: modFileUtilities.
3850      If Forms(strDocName).Visible = False Then
3860        Forms(strDocName).Visible = True
3870      End If
3880    End If
3890    If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
3900      If Forms(strCallingForm).Visible = False Then
3910        Forms(strCallingForm).Visible = True
3920      End If
3930    End If
3940    Forms(THIS_NAME).Visible = True
3950    Select Case ERR.Number
        Case Else
3960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3970    End Select
3980    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4020    intShiftDown = (Shift And acShiftMask) > 0
4030    intAltDown = (Shift And acAltMask) > 0
4040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4060      Select Case intRetVal
          Case vbKeyTab
4070        With Me
4080          intRetVal = 0
4090          .cmdClose.SetFocus
4100        End With
4110      End Select
4120    End If

        ' ** Shift keys.
4130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4140      Select Case intRetVal
          Case vbKeyTab
4150        With Me
4160          intRetVal = 0
4170          .Status3.SetFocus
4180        End With
4190      End Select
4200    End If

EXITP:
4210    KeyCode = intRetVal
4220    Exit Sub

ERRH:
4230    intRetVal = 0
4240    Select Case ERR.Number
        Case Else
4250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4260    End Select
4270    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

4310    With Me
4320      If blnPrintReport_MouseDown = False Then
4330        Select Case blnPrintReport_Focus
            Case True
4340          .cmdPrintReport_raised_focus_dots_img.Visible = True
4350          .cmdPrintReport_raised_focus_img.Visible = False
4360        Case False
4370          .cmdPrintReport_raised_focus_img.Visible = True
4380          .cmdPrintReport_raised_focus_dots_img.Visible = False
4390        End Select
4400        .cmdPrintReport_raised_img.Visible = False
4410        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4420        .cmdPrintReport_sunken_focus_dots_img.Visible = False
4430        .cmdPrintReport_raised_img_dis.Visible = False
4440      End If
4450    End With

EXITP:
4460    Exit Sub

ERRH:
4470    Select Case ERR.Number
        Case Else
4480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4490    End Select
4500    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

4610    With Me
4620      .cmdPrintReport_raised_focus_dots_img.Visible = True
4630      .cmdPrintReport_raised_img.Visible = False
4640      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4650      .cmdPrintReport_raised_focus_img.Visible = False
4660      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4670      .cmdPrintReport_raised_img_dis.Visible = False
4680      blnPrintReport_MouseDown = False
4690    End With

EXITP:
4700    Exit Sub

ERRH:
4710    Select Case ERR.Number
        Case Else
4720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4730    End Select
4740    Resume EXITP

End Sub

Private Sub cmdPrintReport_LostFocus()

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

4810    With Me
4820      .cmdPrintReport_raised_img.Visible = True
4830      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
4840      .cmdPrintReport_raised_focus_img.Visible = False
4850      .cmdPrintReport_raised_focus_dots_img.Visible = False
4860      .cmdPrintReport_sunken_focus_dots_img.Visible = False
4870      .cmdPrintReport_raised_img_dis.Visible = False
4880      blnPrintReport_Focus = False
4890    End With

EXITP:
4900    Exit Sub

ERRH:
4910    Select Case ERR.Number
        Case Else
4920      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4930    End Select
4940    Resume EXITP

End Sub

Private Sub Status3_GotFocus()

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "Status3_GotFocus"

5010    With Me
5020      .Status3.SelLength = 0
5030      .Status3.SelStart = 0
5040    End With

EXITP:
5050    Exit Sub

ERRH:
5060    Select Case ERR.Number
        Case Else
5070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5080    End Select
5090    Resume EXITP

End Sub

Private Sub Status3_KeyDown(KeyCode As Integer, Shift As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "Status3_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5120    intShiftDown = (Shift And acShiftMask) > 0
5130    intAltDown = (Shift And acAltMask) > 0
5140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5160      Select Case intRetVal
          Case vbKeyTab
5170        With Me
5180          intRetVal = 0
5190          If .cmdPrintReport.Enabled = True Then
5200            .cmdPrintReport.SetFocus
5210          Else
5220            .cmdClose.SetFocus
5230          End If
5240        End With
5250      End Select
5260    End If

        ' ** Shift keys.
5270    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5280      Select Case intRetVal
          Case vbKeyTab
5290        With Me
5300          intRetVal = 0
5310          .cmdClose.SetFocus
5320        End With
5330      End Select
5340    End If

EXITP:
5350    KeyCode = intRetVal
5360    Exit Sub

ERRH:
5370    intRetVal = 0
5380    Select Case ERR.Number
        Case Else
5390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5400    End Select
5410    Resume EXITP

End Sub

Private Sub Status3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "Status3_MouseMove"

5510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
5520    Exit Sub

ERRH:
5530    Select Case ERR.Number
        Case Else
5540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5550    End Select
5560    Resume EXITP

End Sub
