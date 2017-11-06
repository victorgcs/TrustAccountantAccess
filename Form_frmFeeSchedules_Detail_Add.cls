VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFeeSchedules_Detail_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmFeeSchedules_Detail_Add"

'VGC 06/20/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}
' **   Order:            R {scheddets_order}
' **   Rate:             T {ScheduleDetail_Rate}
' **   Amount:           U {ScheduleDetail_Amount}

Private lngScheduleID As Long, lngMaxSortOrd As Long
Private strCallingForm As String
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If strTmp01 <> vbNullString Then
150           If CharCnt(strTmp01, "~") = 2 Then

160             intPos01 = InStr(strTmp01, "~")
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190             intPos01 = InStr(strTmp01, "~")
200             lngScheduleID = CLng(Left(strTmp01, (intPos01 - 1)))
210             lngMaxSortOrd = CLng(Mid(strTmp01, (intPos01 + 1)))
220             .Schedule_ID.DefaultValue = lngScheduleID
230             .scheddets_order = (lngMaxSortOrd + 1&)

240           Else
250             Cancel = -1
260             strCallingForm = vbNullString
270           End If
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

EXITP:
390     Exit Sub

ERRH:
400     Select Case ERR.Number
        Case Else
410       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
420     End Select
430     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

510     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
520     intShiftDown = (Shift And acShiftMask) > 0
530     intAltDown = (Shift And acAltMask) > 0
540     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         With Me
580           intRetVal = 0
590           If .Dirty = True Then
600             DoCmd.RunCommand acCmdUndo
610           End If
620           DoCmd.Close acForm, THIS_NAME
630         End With
640       End Select
650     End If

        ' ** Ctrl-Shift keys.
660     If intCtrlDown And (Not intAltDown) And intShiftDown Then
670       Select Case intRetVal
          Case vbKeyF
680         With Me
690           intRetVal = 0
700           .FocusHolder.SetFocus
710         End With
720       End Select
730     End If

EXITP:
740     KeyCode = intRetVal
750     Exit Sub

ERRH:
760     intRetVal = 0
770     Select Case ERR.Number
        Case Else
780       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
790     End Select
800     Resume EXITP

End Sub

Private Sub Form_Timer()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

910     With Me
920       .TimerInterval = 0&
930       If gblnSetFocus = True Then
940         gblnSetFocus = False
950         strFormName = vbNullString
960   On Error Resume Next
970         Do Until strFormName = THIS_NAME
980           DoCmd.SelectObject acForm, THIS_NAME, False
990           strFormName = Screen.ActiveForm.Name
1000        Loop
1010  On Error GoTo ERRH
1020      End If
1030    End With

EXITP:
1040    Exit Sub

ERRH:
1050    DoCmd.Hourglass False
1060    Select Case ERR.Number
        Case Else
1070      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1080    End Select
1090    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

1110    Select Case DataErr
        Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
1120      MsgBox "You must enter numeric data.", vbInformation + vbOKOnly, "Invalid Entry"
1130      Response = acDataErrContinue
1140    Case 2046  ' ** The command or action isn't available now (first or last record).
1150      Response = acDataErrContinue
1160    Case 3201  ' ** You cannot add or change a record because a related record is required in table '|'.
          ' ** It failed to bring over the Schedule_ID that
          ' ** is needed to complete the relationship between
          ' ** Schedule and Schedule Detail.
1170      Me.Schedule_ID = Forms("frmFeeSchedules").frmFeeSchedules_Sub.Form.Schedule_ID
1180      Response = acDataErrContinue
1190    Case Else
1200      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
1210      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
1220    End Select

EXITP:
1230    Exit Sub

ERRH:
1240    Select Case ERR.Number
        Case Else
1250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1260    End Select
1270    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim lngScheduleID As Long, lngScheduleDetailID As Long

1310    With Me
1320      .ScheduleDetail_DateModified = Now()
1330      If IsNull(.ScheduleDetail_Rate) = True Then
1340        MsgBox "Please enter a valid rate.", vbInformation + vbOKOnly, "Entry Required"
1350        .ScheduleDetail_Rate.SetFocus
1360      Else
1370        If IsNull(.ScheduleDetail_Amount) = True Then
1380          MsgBox "Please enter a valid amount.", vbInformation + vbOKOnly, "Entry Required"
1390          .ScheduleDetail_Amount.SetFocus
1400        Else
1410          If .ScheduleDetail_Amount = 0 Then
1420            MsgBox "Please enter a valid amount.", vbInformation + vbOKOnly, "Entry Required"
1430            .ScheduleDetail_Amount.SetFocus
1440          Else
1450            lngScheduleID = .Schedule_ID
1460            lngScheduleDetailID = .ScheduleDetail_ID
1470            lngMaxSortOrd = .scheddets_order
1480            Set dbs = CurrentDb
1490            With dbs
1500              Set rst = .OpenRecordset("tblScheduleDetail_Staging", dbOpenDynaset, dbConsistent)
1510              With rst
1520                .AddNew
1530                ![Schedule_ID] = lngScheduleID
1540                ![ScheduleDetail_ID] = lngScheduleDetailID
                    ' ** ![scheddets_id] : AutoNumber.
1550                ![scheddets_order] = lngMaxSortOrd
1560                ![scheddets_datemodified] = Now()
1570                .Update
1580                .Close
1590              End With
1600              Set rst = Nothing
1610              .Close
1620            End With
1630            Set dbs = Nothing
1640            Forms("frmFeeSchedules").frmFeeSchedules_Detail_Sub.Form.Requery
1650            Forms("frmFeeSchedules").frmFeeSchedules_Detail_Sub.Form.scheddets_order.Visible = True
1660            Forms("frmFeeSchedules").frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Rate.Visible = True
1670            Forms("frmFeeSchedules").frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Amount.Visible = True
1680            Forms("frmFeeSchedules").frmFeeSchedules_Detail_Sub.Form.RecalcSortOrd lngScheduleDetailID  ' ** Form Procedure: frmFeeSchedules_Detail_Sub.
1690            DoCmd.Close acForm, THIS_NAME
1700          End If
1710        End If
1720      End If
1730    End With

EXITP:
1740    Set rst = Nothing
1750    Set dbs = Nothing
1760    Exit Sub

ERRH:
1770    THAT_PROC = THIS_PROC
1780    That_Erl = Erl
1790    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
1800    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1910    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1920    intShiftDown = (Shift And acShiftMask) > 0
1930    intAltDown = (Shift And acAltMask) > 0
1940    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyTab
1970        With Me
1980          intRetVal = 0
1990          .cmdCancel.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab
2050        With Me
2060          intRetVal = 0
2070          .scheddets_order.SetFocus
2080        End With
2090      End Select
2100    End If

EXITP:
2110    KeyCode = intRetVal
2120    Exit Sub

ERRH:
2130    intRetVal = 0
2140    Select Case ERR.Number
        Case Else
2150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2160    End Select
2170    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2210    With Me
2220      If .Dirty = True Then
2230        DoCmd.RunCommand acCmdUndo
2240      End If
2250      Forms("frmFeeSchedules").frmFeeSchedules_Detail_Sub.Form.Refresh
2260      DoCmd.Close acForm, THIS_NAME
2270    End With

EXITP:
2280    Exit Sub

ERRH:
2290    THAT_PROC = THIS_PROC
2300    That_Erl = Erl
2310    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
2320    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          .ScheduleDetail_Rate.SetFocus
2500        End With
2510      End Select
2520    End If

        ' ** Shift keys.
2530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2540      Select Case intRetVal
          Case vbKeyTab
2550        With Me
2560          intRetVal = 0
2570          .cmdOK.SetFocus
2580        End With
2590      End Select
2600    End If

EXITP:
2610    KeyCode = intRetVal
2620    Exit Sub

ERRH:
2630    intRetVal = 0
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Private Sub ScheduleDetail_Rate_AfterUpdate()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Rate_AfterUpdate"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

2710    blnContinue = True

2720    With Me
2730      If IsNull(.ScheduleDetail_Rate) = True Then
2740        .ScheduleDetail_Rate = 0#
2750      End If
2760      If IsNull(.ScheduleDetail_Amount) = True Then
2770        .ScheduleDetail_Amount = 0@
2780      End If
2790      If .ScheduleDetail_Rate = 0 And .ScheduleDetail_Amount = 0 Then
            ' ** OK.
2800      Else
2810        intPos01 = InStr(1, .ScheduleDetail_Rate, ".", vbTextCompare)
2820        If intPos01 > 0 Then
2830          If Len(.ScheduleDetail_Rate) - intPos01 > 6 Then
2840            blnContinue = False
2850            MsgBox "You can only enter a rate with up to 6 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
2860            .ScheduleDetail_Rate = Null
2870            .ScheduleDetail_Rate.SetFocus
2880          End If
2890        End If
2900        If blnContinue = True Then
2910          If .ScheduleDetail_Rate > 1 Then
2920            blnContinue = False
2930            MsgBox "Rate must be <= 100%", vbInformation + vbOKOnly, "Invalid Entry"
2940            .ScheduleDetail_Rate = Null
2950            .ScheduleDetail_Rate.SetFocus
2960          End If
2970        End If
2980      End If
2990    End With  ' ** Me.

EXITP:
3000    Exit Sub

ERRH:
3010    Select Case ERR.Number
        Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub

Private Sub ScheduleDetail_Rate_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Rate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3120    intShiftDown = (Shift And acShiftMask) > 0
3130    intAltDown = (Shift And acAltMask) > 0
3140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3160      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3170        With Me
3180          intRetVal = 0
3190          .ScheduleDetail_Amount.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3250        With Me
3260          intRetVal = 0
3270          .cmdCancel.SetFocus
3280        End With
3290      End Select
3300    End If

EXITP:
3310    KeyCode = intRetVal
3320    Exit Sub

ERRH:
3330    intRetVal = 0
3340    Select Case ERR.Number
        Case Else
3350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3360    End Select
3370    Resume EXITP

End Sub

Private Sub ScheduleDetail_Amount_AfterUpdate()

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Amount_AfterUpdate"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

3410    blnContinue = True

3420    With Me
3430      If IsNull(.ScheduleDetail_Amount) = True Then
3440        .ScheduleDetail_Amount = 0@
3450      End If
3460      If IsNull(.ScheduleDetail_Rate) = True Then
3470        .ScheduleDetail_Rate = 0#
3480      End If
3490      If .ScheduleDetail_Amount = 0 And .ScheduleDetail_Rate = 0 Then
3500        blnContinue = False
3510      Else
3520        intPos01 = InStr(1, .ScheduleDetail_Amount, ".", vbTextCompare)
3530        If intPos01 > 0 Then
3540          If Len(.ScheduleDetail_Amount) - intPos01 > 2 Then
3550            blnContinue = False
3560            MsgBox "You can only enter an Amount with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
3570            .ScheduleDetail_Amount = Null
3580            .ScheduleDetail_Amount.SetFocus
3590          End If
3600        End If
3610        If blnContinue = True Then
3620          If .ScheduleDetail_Amount > 999999999.99 Then
3630            blnContinue = False
3640            MsgBox "Amount is too large.", vbInformation + vbOKOnly, "Invalid Entry"
3650            .ScheduleDetail_Amount = Null
3660            .ScheduleDetail_Amount.SetFocus
3670          Else
3680            DoCmd.RunCommand acCmdSaveRecord
3690          End If
3700        End If
3710      End If
3720    End With  ' ** Me.

EXITP:
3730    Exit Sub

ERRH:
3740    THAT_PROC = THIS_PROC
3750    That_Erl = Erl
3760    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
3770    Resume EXITP

End Sub

Private Sub ScheduleDetail_Amount_KeyDown(KeyCode As Integer, Shift As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Amount_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

3810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
3820    intShiftDown = (Shift And acShiftMask) > 0
3830    intAltDown = (Shift And acAltMask) > 0
3840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
3850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
3860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3870        With Me
3880          intRetVal = 0
3890          .scheddets_order.SetFocus
3900        End With
3910      End Select
3920    End If

        ' ** Shift keys.
3930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3950        With Me
3960          intRetVal = 0
3970          .ScheduleDetail_Rate.SetFocus
3980        End With
3990      End Select
4000    End If

EXITP:
4010    KeyCode = intRetVal
4020    Exit Sub

ERRH:
4030    intRetVal = 0
4040    Select Case ERR.Number
        Case Else
4050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4060    End Select
4070    Resume EXITP

End Sub

Private Sub scheddets_order_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "scheddets_order_AfterUpdate"

4110    With Me
4120      If IsNull(.scheddets_order) = False Then
4130        If IsNumeric(.scheddets_order) = True Then
4140          If .scheddets_order < 0 Then
4150            MsgBox "Number must be a positive integer.", vbInformation + vbOKOnly, "Invalid Entry"
4160            .scheddets_order = lngMaxSortOrd
4170          End If
4180        Else
4190          .scheddets_order = lngMaxSortOrd
4200        End If
4210      Else
4220        .scheddets_order = lngMaxSortOrd
4230      End If
4240    End With

EXITP:
4250    Exit Sub

ERRH:
4260    THAT_PROC = THIS_PROC
4270    That_Erl = Erl
4280    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
4290    Resume EXITP

End Sub

Private Sub scheddets_order_KeyDown(KeyCode As Integer, Shift As Integer)

4300  On Error GoTo ERRH

        Const THIS_PROC As String = "scheddets_order_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

4310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
4320    intShiftDown = (Shift And acShiftMask) > 0
4330    intAltDown = (Shift And acAltMask) > 0
4340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
4350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
4360      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4370        With Me
4380          intRetVal = 0
4390  On Error Resume Next
4400          .cmdOK.SetFocus
4410  On Error GoTo ERRH
4420        End With
4430      End Select
4440    End If

        ' ** Shift keys.
4450    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4470        With Me
4480          intRetVal = 0
4490          .ScheduleDetail_Amount.SetFocus
4500        End With
4510      End Select
4520    End If

EXITP:
4530    KeyCode = intRetVal
4540    Exit Sub

ERRH:
4550    intRetVal = 0
4560    Select Case ERR.Number
        Case Else
4570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4580    End Select
4590    Resume EXITP

End Sub
