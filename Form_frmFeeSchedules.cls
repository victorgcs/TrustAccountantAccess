VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFeeSchedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmFeeSchedules"

'VGC 10/02/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add Schedule:     A {cmdAdd}
' **   Base:             B {Schedule_Base on frmFeeSchedules_Sub}
' **   Close:            C {cmdClose}
' **   Delete Schedule:  D {cmdDelete}
' **   Add Entry:        E {cmdAddDetail on frmFeeSchedules_Detail_Sub}
' **   Minimum:          I {Schedule_Minimum on frmFeeSchedules_Sub}
' **   Name:             M {Schedule_Name on frmFeeSchedules_Sub}
' **   Delete Entry:     N {cmdDeleteDetail on frmFeeSchedules_Detail_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Order:            R {scheddets_order on frmFeeSchedules_Detail_Sub}
' **   Rate:             T {ScheduleDetail_Rate on frmFeeSchedules_Detail_Sub}
' **   Amount:           U {ScheduleDetail_Amount on frmFeeSchedules_Detail_Sub}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmFeeSchedules_Sub, frmFeeSchedules_Detail_Sub}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

Private blnGoingToReport As Boolean, lngPrintReportLbl_Left As Long
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           lngPrintReportLbl_Left = .cmdPrintReport_lbl2.Left

160           blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.
170           DoEvents
180           blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.
190           DoEvents

              ' ** modQueryFunctions1 has to refer to the chosen Schedule_ID in the subform.
200           gstrFormQuerySpec = "frmFeeSchedules_Detail_Sub"
210           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

220           DoCmd.Hourglass False

230         Else
240           Cancel = -1
250         End If
260       Else
270         Cancel = -1
280       End If
290     End With

300     If Cancel = -1 Then
310       Beep
320       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
330     End If

EXITP:
340     Exit Sub

ERRH:
350     DoCmd.Hourglass False
360     Select Case ERR.Number
        Case Else
370       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
380     End Select
390     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim lngRecsCur2 As Long
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Base:             B {Schedule_Base on frmFeeSchedules_Sub}
        ' **   Minimum:          I {Schedule_Minimum on frmFeeSchedules_Sub}
        ' **   Name:             M {Schedule_Name on frmFeeSchedules_Sub}
        ' **   Order:            R {scheddets_order on frmFeeSchedules_Detail_Sub}
        ' **   Rate:             T {ScheduleDetail_Rate on frmFeeSchedules_Detail_Sub}
        ' **   Amount:           U {ScheduleDetail_Amount on frmFeeSchedules_Detail_Sub}

        ' ** Shortcut Ctrl to other forms and functionality:
        ' **   Save:             S {cmdSave on frmFeeSchedules_Sub, frmFeeSchedules_Detail_Sub}

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         intRetVal = 0
480         DoCmd.Close acForm, THIS_NAME
490       End Select
500     End If

        ' ** Alt keys.
510     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
520       Select Case intRetVal
          Case vbKeyB
530         With Me
540           intRetVal = 0
550           .frmFeeSchedules_Sub.SetFocus
560           .frmFeeSchedules_Sub.Form.Schedule_Base.SetFocus
570         End With
580       Case vbKeyI
590         With Me
600           intRetVal = 0
610           .frmFeeSchedules_Sub.SetFocus
620           .frmFeeSchedules_Sub.Form.Schedule_Minimum.SetFocus
630         End With
640       Case vbKeyM
650         With Me
660           intRetVal = 0
670           .frmFeeSchedules_Sub.SetFocus
680           .frmFeeSchedules_Sub.Form.Schedule_Name.SetFocus
690         End With
700       Case vbKeyR
710         With Me
720           intRetVal = 0
730           .frmFeeSchedules_Detail_Sub.SetFocus
740           .frmFeeSchedules_Detail_Sub.Form.scheddets_order.SetFocus
750         End With
760       Case vbKeyT
770         With Me
780           intRetVal = 0
790           .frmFeeSchedules_Detail_Sub.SetFocus
800           .frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Rate.SetFocus
810         End With
820       Case vbKeyU
830         With Me
840           intRetVal = 0
850           .frmFeeSchedules_Detail_Sub.SetFocus
860           .frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Amount.SetFocus
870         End With
880       Case vbKeyX
890         intRetVal = 0
900         cmdClose_Click  ' ** Procedure: Below.
910       End Select
920     End If

        ' ** Ctrl keys.
930     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
940       Select Case intRetVal
          Case vbKeyS
950         With Me
960           intRetVal = 0
970           lngRecsCur = .frmFeeSchedules_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Sub.
980           lngRecsCur2 = .frmFeeSchedules_Detail_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Detail_Sub.
990           If lngRecsCur > 0& Or lngRecsCur2 > 0 Then
1000            If lngRecsCur > 0& Then
1010              .frmFeeSchedules_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmFeeSchedules_Sub.
1020            End If
1030            If lngRecsCur2 > 0 Then
1040              .frmFeeSchedules_Detail_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmFeeSchedules_Detail_Sub.
1050            End If
1060          Else
1070            Beep
1080          End If
1090        End With
1100      End Select
1110    End If

        ' ** Ctrl-Shift keys.
1120    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1130      Select Case intRetVal
          Case vbKeyF
1140        With Me
1150          intRetVal = 0
1160          .FocusHolder.SetFocus
1170        End With
1180      End Select
1190    End If

EXITP:
1200    KeyCode = intRetVal
1210    Exit Sub

ERRH:
1220    intRetVal = 0
1230    Select Case ERR.Number
        Case Else
1240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1250    End Select
1260    Resume EXITP

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
1440        Select Case blnGoingToReport
            Case True
1450          DoCmd.Hourglass False  ' ** It should already be off.
1460          blnGoingToReport = False
1470          gblnGoToReport = False
1480          .GoToReport_arw_pr_img.Visible = False
1490          .cmdPrintReport_lbl2.Left = lngPrintReportLbl_Left
1500          DoEvents
1510        Case False
1520          DoCmd.Hourglass True  ' ** Make sure it's still running.
1530          DoEvents
1540          blnGoingToReport = True
1550          .cmdPrintReport_lbl2.Left = (.cmdPrintReport_lbl2.Left + .GoToReport_arw_pr_img.Width)
1560          .GoToReport_arw_pr_img.Visible = True
1570          If .cmdPrintReport.Enabled = True Then
1580            .cmdPrintReport.SetFocus
1590          Else
1600            .cmdClose.SetFocus
1610          End If
1620          DoEvents
1630          Beep
1640          DoCmd.Hourglass False
1650          DoEvents
1660          .TimerInterval = (GTR_WAIT * 2&)
1670        End Select
1680      End If
1690    End With

EXITP:
1700    Exit Sub

ERRH:
1710    Select Case ERR.Number
        Case 2100  ' ** The control or subform control is too large for this location.
          ' ** Ignore.
1720    Case Else
1730      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1740    End Select
1750    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1810    If DataErr <> 0 Then
1820      Select Case DataErr
          Case 3075, 8530, 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
1830      Case Else
1840        zErrorHandler Me.Name, "Undefined Form Error", DataErr
1850      End Select
1860    End If

EXITP:
1870    Exit Sub

ERRH:
1880    Select Case ERR.Number
        Case Else
1890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1900    End Select
1910    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim strDocName As String
        Dim blnRetVal As Boolean

        ' ** Save SubForm prefs.
2010    Pref_Save "frmFeeSchedules_Sub"  ' ** Module Procedure: modPreferenceFuncs.
2020    Pref_Save "frmFeeSchedules_Detail_Sub"  ' ** Module Procedure: modPreferenceFuncs.

2030    gstrFormQuerySpec = vbNullString

2040    blnRetVal = MouseWheelOFF  ' ** Module Function: modMouseWheel.

2050    Select Case strCallingForm
        Case "frmMenu_Utility"
2060      gblnSetFocus = True
2070      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
2080    Case vbNullString
2090      gblnSetFocus = True
2100      strDocName = "frmMenu_Utility"
2110      DoCmd.OpenForm strDocName, , , , , , "frmMenu_Main"
2120    Case Else
2130      DoCmd.OpenForm strCallingForm, , , , , , THIS_NAME
2140    End Select

EXITP:
2150    Exit Sub

ERRH:
2160    Select Case ERR.Number
        Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

2210    With Me
2220      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
2230        Select Case blnPrintReport_Focus
            Case True
2240          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2250          .cmdPrintReport_raised_img.Visible = False
2260        Case False
2270          .cmdPrintReport_raised_img.Visible = True
2280          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2290        End Select
2300        .cmdPrintReport_raised_focus_img.Visible = False
2310        .cmdPrintReport_raised_focus_dots_img.Visible = False
2320        .cmdPrintReport_sunken_focus_dots_img.Visible = False
2330        .cmdPrintReport_raised_img_dis.Visible = False
2340      End If
2350    End With

EXITP:
2360    Exit Sub

ERRH:
2370    Select Case ERR.Number
        Case Else
2380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2390    End Select
2400    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

2510    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long

2610    DoCmd.Hourglass True

2620    Set dbs = CurrentDb
2630    With dbs
          ' ** ScheduleDetail, grouped, with cnt.
2640      Set qdf = .QueryDefs("qryFeeSchedule_02")
2650      Set rst = qdf.OpenRecordset
2660      With rst
2670        .MoveFirst
2680        lngRecs = ![NumRecs]
2690        .Close
2700      End With
2710      .Close
2720    End With

2730    If lngRecs > 0& Then  ' ** One or more records with blank fields.
2740      DoCmd.Hourglass False
2750      MsgBox "There are one or more detail records with a blank column." & vbCrLf & _
            "Please fill in the column(s) or delete the records to continue.", vbInformation + vbOKOnly, "Entry Required"
2760    Else
2770      If ValidateFees = True Then  ' ** Function: Below.

2780        DoCmd.Close acForm, THIS_NAME

2790      End If

2800      DoCmd.Hourglass False

2810    End If

EXITP:
2820    Set rst = Nothing
2830    Set qdf = Nothing
2840    Set dbs = Nothing
2850    Exit Sub

ERRH:
2860    DoCmd.Hourglass False
2870    Select Case ERR.Number
        Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

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
3090          lngRecsCur = .frmFeeSchedules_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Sub.
3100          If lngRecsCur > 0& Then
3110            .frmFeeSchedules_Sub.SetFocus
3120            .frmFeeSchedules_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmFeeSchedules_Sub.
3130            .frmFeeSchedules_Sub.Form.Schedule_Name.SetFocus
3140          Else
3150            If .cmdPrintReport.Enabled = True Then
3160              .cmdPrintReport.SetFocus
3170            Else
3180              .cmdAdd.SetFocus
3190            End If
3200          End If
3210        End With
3220      End Select
3230    End If

        ' ** Shift keys.
3240    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3250      Select Case intRetVal
          Case vbKeyTab
3260        With Me
3270          intRetVal = 0
3280          If .cmdDelete.Enabled = True Then
3290            .cmdDelete.SetFocus
3300          Else
3310            .cmdAdd.SetFocus
3320          End If
3330        End With
3340      End Select
3350    End If

EXITP:
3360    KeyCode = intRetVal
3370    Exit Sub

ERRH:
3380    intRetVal = 0
3390    Select Case ERR.Number
        Case Else
3400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3410    End Select
3420    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

3510    With Me
3520      blnPrintReport_Focus = True
3530      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
3540      .cmdPrintReport_raised_img.Visible = False
3550      .cmdPrintReport_raised_focus_img.Visible = False
3560      .cmdPrintReport_raised_focus_dots_img.Visible = False
3570      .cmdPrintReport_sunken_focus_dots_img.Visible = False
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

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

3710    With Me
3720      blnPrintReport_MouseDown = True
3730      .cmdPrintReport_sunken_focus_dots_img.Visible = True
3740      .cmdPrintReport_raised_img.Visible = False
3750      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3760      .cmdPrintReport_raised_focus_img.Visible = False
3770      .cmdPrintReport_raised_focus_dots_img.Visible = False
3780      .cmdPrintReport_raised_img_dis.Visible = False
3790    End With

EXITP:
3800    Exit Sub

ERRH:
3810    Select Case ERR.Number
        Case Else
3820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3830    End Select
3840    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String, strOrdBy1 As String, strOrdBy2 As String
        Dim blnDesc1 As Boolean, blnDesc2 As Boolean
        Dim strSfx As String
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String

3910    With Me
3920      strOrdBy1 = .frmFeeSchedules_Sub.Form.OrderBy
3930      If InStr(strOrdBy1, " DESC") > 0 Then blnDesc1 = True
3940      strOrdBy2 = .frmFeeSchedules_Detail_Sub.Form.OrderBy
3950      If InStr(strOrdBy2, " DESC") > 0 Then blnDesc2 = True
3960      intPos01 = InStr(strOrdBy1, "]")
3970      strTmp01 = Mid(Left(strOrdBy1, (intPos01 - 1)), 2)
3980      intPos01 = InStr(strOrdBy2, "]")
3990      strTmp02 = Mid(Left(strOrdBy2, (intPos01 - 1)), 2)
4000      strSfx = vbNullString
4010      Select Case strTmp01
          Case "Schedule_ID"
4020        strSfx = "_01"
4030      Case "Schedule_Name"
4040        strSfx = "_02"
4050      Case "Schedule_Base"
4060        strSfx = "_03"
4070      Case "Schedule_Minimum"
4080        strSfx = "_04"
4090      End Select
4100      Select Case blnDesc1
          Case True
4110        strSfx = strSfx & "d"
4120      Case False
4130        strSfx = strSfx & "a"
4140      End Select
4150      Select Case strTmp02
          Case "scheddets_order"
4160        strSfx = strSfx & "_01"
4170      Case "ScheduleDetail_Rate"
4180        strSfx = strSfx & "_02"
4190      Case "ScheduleDetail_Amount"
4200        strSfx = strSfx & "_03"
4210      End Select
4220      Select Case blnDesc2
          Case True
4230        strSfx = strSfx & "d"
4240      Case False
4250        strSfx = strSfx & "a"
4260      End Select
4270      strDocName = "rptListOfFeeSchedules" & strSfx
4280      If gblnDev_Debug = True Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
4290        DoCmd.OpenReport strDocName, acViewPreview
4300        DoCmd.Maximize
4310        DoCmd.RunCommand acCmdFitToWindow
4320        .Visible = False
4330      Else
            '##GTR_Ref: rptListOfFeeSchedules
4340        DoCmd.OpenReport strDocName, acViewNormal
4350      End If
4360    End With

EXITP:
4370    Exit Sub

ERRH:
4380    Select Case ERR.Number
        Case Else
4390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4400    End Select
4410    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4520    intShiftDown = (Shift And acShiftMask) > 0
4530    intAltDown = (Shift And acAltMask) > 0
4540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4560      Select Case intRetVal
          Case vbKeyTab
4570        With Me
4580          intRetVal = 0
4590          If .cmdAdd.Enabled = True Then
4600            .cmdAdd.SetFocus
4610          ElseIf .cmdDelete.Enabled = True Then
4620            .cmdDelete.SetFocus
4630          Else
4640            .cmdClose.SetFocus
4650          End If
4660        End With
4670      End Select
4680    End If

        ' ** Shift keys.
4690    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4700      Select Case intRetVal
          Case vbKeyTab
4710        With Me
4720          intRetVal = 0
4730          lngRecsCur = .frmFeeSchedules_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Sub.
4740          If lngRecsCur > 0& Then
4750            .frmFeeSchedules_Detail_Sub.SetFocus
4760            If .frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = True Then
4770              .frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.SetFocus
4780            Else
4790              .frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.SetFocus
4800            End If
4810          Else
4820            .cmdClose.SetFocus
4830          End If
4840        End With
4850      End Select
4860    End If

EXITP:
4870    KeyCode = intRetVal
4880    Exit Sub

ERRH:
4890    intRetVal = 0
4900    Select Case ERR.Number
        Case Else
4910      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4920    End Select
4930    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

5000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

5010    With Me
5020      If blnPrintReport_MouseDown = False Then
5030        Select Case blnPrintReport_Focus
            Case True
5040          .cmdPrintReport_raised_focus_dots_img.Visible = True
5050          .cmdPrintReport_raised_focus_img.Visible = False
5060        Case False
5070          .cmdPrintReport_raised_focus_img.Visible = True
5080          .cmdPrintReport_raised_focus_dots_img.Visible = False
5090        End Select
5100        .cmdPrintReport_raised_img.Visible = False
5110        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5120        .cmdPrintReport_sunken_focus_dots_img.Visible = False
5130        .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

5310    With Me
5320      .cmdPrintReport_raised_focus_dots_img.Visible = True
5330      .cmdPrintReport_raised_img.Visible = False
5340      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5350      .cmdPrintReport_raised_focus_img.Visible = False
5360      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5370      .cmdPrintReport_raised_img_dis.Visible = False
5380      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

5500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

5510    With Me
5520      .cmdPrintReport_raised_img.Visible = True
5530      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
5540      .cmdPrintReport_raised_focus_img.Visible = False
5550      .cmdPrintReport_raised_focus_dots_img.Visible = False
5560      .cmdPrintReport_sunken_focus_dots_img.Visible = False
5570      .cmdPrintReport_raised_img_dis.Visible = False
5580      blnPrintReport_Focus = False
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

Private Sub cmdAdd_Click()

5700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim lngScheduleID As Long

5710    With Me
5720      DoCmd.Hourglass True
5730      DoEvents
5740      .Schedule_ID_new = 0&
5750      strDocName = .frmFeeSchedules_Sub.Form.RecordSource
5760      .frmFeeSchedules_Sub.Form.RecordSource = vbNullString
5770      DoEvents
5780      ChangeSeed_Ext "Schedule"  ' ** Module Function: modAutonumberFieldFuncs.
5790      DoEvents
5800      .frmFeeSchedules_Sub.Form.RecordSource = strDocName
5810      DoEvents
5820      strDocName = "frmFeeSchedules_Add"
5830      DoCmd.Hourglass False
5840      DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, THIS_NAME

5850      lngScheduleID = .Schedule_ID_new
5860      If lngScheduleID > 0& Then
            ' ** ScheduleDetail and/or tblScheduleDetail_Staging seems to retain items from deleted schedule!
5870        DoCmd.Hourglass True
5880        DoEvents
5890        Set dbs = CurrentDb
            ' ** Delete ScheduleDetail, by specified [schdid].
5900        Set qdf = dbs.QueryDefs("qryFeeSchedule_16_01")
5910        With qdf.Parameters
5920          ![schdid] = lngScheduleID
5930        End With
5940        qdf.Execute
5950        Set qdf = Nothing
            ' ** Delete tblScheduleDetail_Staging, by specified [schdid].
5960        Set qdf = dbs.QueryDefs("qryFeeSchedule_16_02")
5970        With qdf.Parameters
5980          ![schdid] = lngScheduleID
5990        End With
6000        qdf.Execute
6010        DoCmd.Hourglass False
6020      End If

6030    End With

EXITP:
6040    Set qdf = Nothing
6050    Set dbs = Nothing
6060    Exit Sub

ERRH:
6070    DoCmd.Hourglass False
6080    Select Case ERR.Number
        Case Else
6090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6100    End Select
6110    Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

6200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

6210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
6220    intShiftDown = (Shift And acShiftMask) > 0
6230    intAltDown = (Shift And acAltMask) > 0
6240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
6250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
6260      Select Case intRetVal
          Case vbKeyTab
6270        With Me
6280          intRetVal = 0
6290          If .cmdDelete.Enabled = True Then
6300            .cmdDelete.SetFocus
6310          Else
6320            .cmdClose.SetFocus
6330          End If
6340        End With
6350      End Select
6360    End If

        ' ** Shift keys.
6370    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6380      Select Case intRetVal
          Case vbKeyTab
6390        With Me
6400          intRetVal = 0
6410          If .cmdPrintReport.Enabled = True Then
6420            .cmdPrintReport.SetFocus
6430          Else
6440            lngRecsCur = .frmFeeSchedules_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Sub.
6450            If lngRecsCur > 0& Then
6460              .frmFeeSchedules_Detail_Sub.SetFocus
6470              If .frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = True Then
6480                .frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.SetFocus
6490              Else
6500                .frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.SetFocus
6510              End If
6520            Else
6530              .cmdClose.SetFocus
6540            End If
6550          End If
6560        End With
6570      End Select
6580    End If

EXITP:
6590    KeyCode = intRetVal
6600    Exit Sub

ERRH:
6610    intRetVal = 0
6620    Select Case ERR.Number
        Case Else
6630      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6640    End Select
6650    Resume EXITP

End Sub

Private Sub cmdDelete_Click()

6700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim strAccounts As String
        Dim lngRecs As Long, blnNoRecs As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long

6710    strAccounts = vbNullString: blnNoRecs = False

6720    varTmp00 = DCount("*", "Schedule")
6730    If IsNull(varTmp00) = True Then
6740      blnNoRecs = True
6750    Else
6760      If varTmp00 = 0 Then
6770        blnNoRecs = True
6780      End If
6790    End If

6800    If blnNoRecs = False Then

6810      Me.frmFeeSchedules_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmFeeSchedules_Sub.

6820      Set dbs = CurrentDb
6830      With dbs
            ' ** Account, by specified [schdid].
6840        Set qdf = .QueryDefs("qryFeeSchedule_09")
6850        With qdf.Parameters
6860          ![schdid] = Me.frmFeeSchedules_Sub.Form.Schedule_ID
6870        End With
6880        Set rst = qdf.OpenRecordset
6890        With rst
6900          If .BOF = True And .EOF = True Then
                ' ** Great, no problem!
6910          Else
6920            .MoveLast
6930            lngRecs = .RecordCount
6940            .MoveFirst
6950            For lngX = 1& To lngRecs
6960              strAccounts = strAccounts & ![accountno] & ", "
6970              If lngX < lngRecs Then .MoveNext
6980            Next
6990            strAccounts = Mid(strAccounts, 1, (Len(strAccounts) - 2))
7000          End If
7010          .Close
7020        End With

7030        If strAccounts = vbNullString Then
7040          If MsgBox("Are you sure you want to delete the '" & Me.frmFeeSchedules_Sub.Form.Schedule_Name & "' fee schedule?", _
                  vbQuestion + vbYesNo, "Delete Fee Schedule") = vbYes Then

                ' ** Delete Schedule, by specified [schdid].
7050            Set qdf = .QueryDefs("qryFeeSchedule_10")
7060            With qdf.Parameters
7070              ![schdid] = Me.frmFeeSchedules_Sub.Form.Schedule_ID
7080            End With
7090            qdf.Execute dbFailOnError

7100            With Me
7110              .frmFeeSchedules_Sub.Form.Requery
7120              .frmFeeSchedules_Detail_Sub.Form.Requery
7130              varTmp00 = DCount("*", "Schedule")
7140              If IsNull(varTmp00) = True Then
7150                blnNoRecs = True
7160              Else
7170                If varTmp00 = 0 Then
7180                  blnNoRecs = True
7190                End If
7200              End If
7210              If blnNoRecs = False Then
7220                .frmFeeSchedules_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmFeeSchedules_Sub.
7230              Else
7240                .cmdAdd.SetFocus
7250                .cmdDelete.Enabled = False
7260                .frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.Enabled = False
7270                .frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = False
7280              End If
7290            End With  ' ** Me.

7300          End If  ' ** MsgBox().
7310        Else
7320          MsgBox "This schedule cannot be deleted." & vbCrLf & vbCrLf & _
                "The following accounts have been assigned to this schedule:" & vbCrLf & strAccounts & vbCrLf & vbCrLf & _
                "To delete this schedule you will need to re-assign these accounts to another schedule.", _
                vbInformation + vbOKOnly, "Invalid Action"
7330        End If  ' ** strAccounts.

7340      End With  ' ** dbs.

7350    Else
7360      MsgBox "There are no Fee Schedules to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
7370    End If  ' ** DCount().

EXITP:
7380    Set rst = Nothing
7390    Set qdf = Nothing
7400    Set dbs = Nothing
7410    Exit Sub

ERRH:
7420    Select Case ERR.Number
        Case Else
7430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7440    End Select
7450    Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

7500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

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
7590          .cmdClose.SetFocus
7600        End With
7610      End Select
7620    End If

        ' ** Shift keys.
7630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
7640      Select Case intRetVal
          Case vbKeyTab
7650        With Me
7660          intRetVal = 0
7670          .cmdAdd.SetFocus
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

Private Sub ShortcutMenu_lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

7800  On Error GoTo ERRH

        Const THIS_PROC As String = "ShortcutMenu_lbl_MouseMove"

7810    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
7820    Exit Sub

ERRH:
7830    Select Case ERR.Number
        Case Else
7840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7850    End Select
7860    Resume EXITP

End Sub

Private Function ValidateFees() As Boolean
' ** Validates fees: checking that there is
' ** AT LEAST ONE schedule detail record for
' ** each fee schedule entered.

7900  On Error GoTo ERRH

        Const THIS_PROC As String = "ValidateFees"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngRecs As Long
        Dim blnRetVal As Boolean

7910    blnRetVal = True  ' ** Unless proven otherwise.

7920    Application.SysCmd acSysCmdSetStatus, "Validating Fee Schedule entries . . ."

7930    Set dbs = CurrentDb
7940    With dbs

          ' ** Schedule, not in ScheduleDetail, grouped, with cnt.
7950      Set qdf = .QueryDefs("qryFeeSchedule_03")
7960      Set rst = qdf.OpenRecordset
7970      With rst
7980        .MoveFirst
7990        lngRecs = ![NumRecs]
8000        .Close
8010      End With

8020      If lngRecs > 0 Then  ' ** At least one problem.
8030        If lngRecs = 1& Then
              ' ** Schedule, not in ScheduleDetail.
8040          Set qdf = .QueryDefs("qryFeeSchedule_04")
8050          Set rst = qdf.OpenRecordset
8060          With rst
8070            .MoveFirst
8080            If ![Schedule_Name] = "Can Be User Defined" Or ![Schedule_Name] = "{Can Be User Defined}" Then
                  ' ** Demo record. Ignore it!
8090            Else
8100              blnRetVal = False
8110            End If
8120            .Close
8130          End With
8140        Else
8150          blnRetVal = False
8160        End If
8170        If blnRetVal = False Then
8180          DoCmd.Hourglass False
8190          MsgBox "You have at least one fee schedule with no detail information." & vbCrLf & _
                "Please enter at least one schedule detail record for each schedule to continue." & vbCrLf & vbCrLf & _
                "If you do not wish to use a laddered rate schedule, enter a '0' in the rate and '1' in the amount columns.", _
                vbExclamation + vbOKOnly, "Entry Required"
8200        End If
8210      End If

8220      .Close
8230    End With

EXITP:
8240    Application.SysCmd acSysCmdClearStatus
8250    Set rst = Nothing
8260    Set qdf = Nothing
8270    Set dbs = Nothing
8280    ValidateFees = blnRetVal
8290    Exit Function

ERRH:
8300    blnRetVal = False
8310    Select Case ERR.Number
        Case Else
8320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8330    End Select
8340    Resume EXITP

End Function
