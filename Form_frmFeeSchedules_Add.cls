VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFeeSchedules_Add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmFeeSchedules_Add"

'VGC 06/20/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Base:             B {Schedule_Base}
' **   Cancel:           C {cmdCancel}
' **   Minimum:          I {Schedule_Minimum}
' **   Name:             M {Schedule_Name}
' **   OK:               O {cmdOK}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150         Else
160           Cancel = -1
170         End If
180       Else
190         Cancel = -1
200       End If
210     End With

220     If Cancel = -1 Then
230       Beep
240       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
250     End If

EXITP:
260     Exit Sub

ERRH:
270     Select Case ERR.Number
        Case Else
280       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
290     End Select
300     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
450     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
460       Select Case intRetVal
          Case vbKeyEscape
470         With Me
480           intRetVal = 0
490           If .Dirty = True Then
500             DoCmd.RunCommand acCmdUndo
510           End If
520           DoCmd.Close acForm, THIS_NAME
530         End With
540       End Select
550     End If

        ' ** Ctrl-Shift keys.
560     If intCtrlDown And (Not intAltDown) And intShiftDown Then
570       Select Case intRetVal
          Case vbKeyF
580         With Me
590           intRetVal = 0
600           .FocusHolder.SetFocus
610         End With
620       End Select
630     End If

EXITP:
640     KeyCode = intRetVal
650     Exit Sub

ERRH:
660     intRetVal = 0
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub Form_Timer()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

810     With Me
820       .TimerInterval = 0&
830       If gblnSetFocus = True Then
840         gblnSetFocus = False
850         strFormName = vbNullString
860   On Error Resume Next
870         Do Until strFormName = THIS_NAME
880           DoCmd.SelectObject acForm, THIS_NAME, False
890           strFormName = Screen.ActiveForm.Name
900         Loop
910   On Error GoTo ERRH
920       End If
930     End With

EXITP:
940     Exit Sub

ERRH:
950     DoCmd.Hourglass False
960     Select Case ERR.Number
        Case Else
970       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
980     End Select
990     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1010    DoCmd.SelectObject acForm, strCallingForm, False

EXITP:
1020    Exit Sub

ERRH:
1030    Select Case ERR.Number
        Case Else
1040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1050    End Select
1060    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim lngScheduleID As Long
        Dim intPos01 As Integer
        Dim blnContinue As Boolean

1110    blnContinue = True

1120    With Me
1130      .Schedule_DateCreated = Now()
1140      .Schedule_DateModified = Now()
1150      If IsNull(.Schedule_Name) = True Then
1160        blnContinue = False
1170        MsgBox "A Schedule Name is required.", vbInformation + vbOKOnly, "Entry Required"
1180        .Schedule_Name.SetFocus
1190      Else
1200        If Trim(.Schedule_Name) = vbNullString Then
1210          blnContinue = False
1220          MsgBox "A Schedule Name is required.", vbInformation + vbOKOnly, "Entry Required"
1230          .Schedule_Name.SetFocus
1240        Else
1250          If IsNull(.Schedule_Base) = True Then
1260            .Schedule_Base = 0#
1270          End If
1280          If IsNull(.Schedule_Minimum) = True Then
1290            .Schedule_Minimum = 0@
1300          End If
1310          If IsNumeric(.Schedule_Base) Then
1320            If CDbl(.Schedule_Base) > 999999999.99 Then
1330              blnContinue = False
1340              MsgBox "Base Fee is too large.", vbInformation + vbOKOnly, "Invalid Entry"
1350              .Schedule_Base.SetFocus
1360            End If
1370          Else
1380            MsgBox "Please enter a valid Base Fee.", vbInformation + vbOKOnly, "Invalid Entry"
1390            blnContinue = False
1400            .Schedule_Base.SetFocus
1410          End If
1420          If blnContinue = True Then
1430            intPos01 = InStr(1, .Schedule_Base, ".", vbTextCompare)
1440            If intPos01 > 0 Then
1450              If Len(.Schedule_Base) - intPos01 > 2 Then
1460                blnContinue = False
1470                MsgBox "You can only enter Base Fee with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
1480                .Schedule_Base.SetFocus
1490              End If
1500            End If
1510          End If
1520          If IsNumeric(.Schedule_Minimum) Then
1530            If CDbl(.Schedule_Minimum) > 999999999.99 Then
1540              blnContinue = False
1550              MsgBox "Minimum Fee is too large.", vbInformation + vbOKOnly, "Invalid Entry"
1560              .Schedule_Minimum.SetFocus
1570            End If
1580          Else
1590            blnContinue = False
1600            MsgBox "Please enter a valid Minimum Fee.", vbInformation + vbOKOnly, "Invalid Entry"
1610            .Schedule_Minimum.SetFocus
1620          End If
1630          If blnContinue = True Then
1640            intPos01 = InStr(1, .Schedule_Minimum, ".", vbTextCompare)
1650            If intPos01 > 0 Then
1660              If Len(.Schedule_Minimum) - intPos01 > 2 Then
1670                blnContinue = False
1680                MsgBox "You can only enter Minimum Fee with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
1690                .Schedule_Minimum.SetFocus
1700              End If
1710            End If
1720          End If
1730        End If
1740      End If
1750      lngScheduleID = .Schedule_ID
1760    End With

1770    If blnContinue = True Then
1780      DoCmd.RunCommand acCmdSaveRecord
1790      DoEvents
1800      Forms("frmFeeSchedules").Schedule_ID_new = lngScheduleID
1810      Forms("frmFeeSchedules").frmFeeSchedules_Sub.Form.Requery
1820      DoEvents
1830      Forms("frmFeeSchedules").frmFeeSchedules_Sub.Form.MoveRec 0, lngScheduleID  ' ** Form Procedure: frmFeeSchedules_Sub.
1840      Forms("frmFeeSchedules").frmFeeSchedules_Sub.SetFocus
1850      Forms("frmFeeSchedules").frmFeeSchedules_Sub.Form.Schedule_Minimum.SetFocus
1860      DoCmd.Close acForm, THIS_NAME
1870    End If

EXITP:
1880    Exit Sub

ERRH:
1890    Select Case ERR.Number
        Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2010    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2020    intShiftDown = (Shift And acShiftMask) > 0
2030    intAltDown = (Shift And acAltMask) > 0
2040    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2050    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2060      Select Case intRetVal
          Case vbKeyTab
2070        With Me
2080          intRetVal = 0
2090          .cmdCancel.SetFocus
2100        End With
2110      End Select
2120    End If

        ' ** Shift keys.
2130    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2140      Select Case intRetVal
          Case vbKeyTab
2150        With Me
2160          intRetVal = 0
2170          .Schedule_Minimum.SetFocus
2180        End With
2190      End Select
2200    End If

EXITP:
2210    KeyCode = intRetVal
2220    Exit Sub

ERRH:
2230    intRetVal = 0
2240    Select Case ERR.Number
        Case Else
2250      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2260    End Select
2270    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2310    With Me
2320      If .Dirty = True Then
2330        DoCmd.RunCommand acCmdUndo
2340      End If
2350      DoCmd.Close acForm, THIS_NAME
2360    End With

EXITP:
2370    Exit Sub

ERRH:
2380    Select Case ERR.Number
        Case Else
2390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2400    End Select
2410    DoCmd.Close acForm, THIS_NAME
2420    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2510    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2520    intShiftDown = (Shift And acShiftMask) > 0
2530    intAltDown = (Shift And acAltMask) > 0
2540    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2560      Select Case intRetVal
          Case vbKeyTab
2570        With Me
2580          intRetVal = 0
2590          .Schedule_Name.SetFocus
2600        End With
2610      End Select
2620    End If

        ' ** Shift keys.
2630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2640      Select Case intRetVal
          Case vbKeyTab
2650        With Me
2660          intRetVal = 0
2670          .cmdOK.SetFocus
2680        End With
2690      End Select
2700    End If

EXITP:
2710    KeyCode = intRetVal
2720    Exit Sub

ERRH:
2730    intRetVal = 0
2740    Select Case ERR.Number
        Case Else
2750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2760    End Select
2770    Resume EXITP

End Sub

Private Sub Schedule_Name_AfterUpdate()

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Name_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim intPos01 As Integer
        Dim strTmp01 As String

2810    With Me
2820      If IsNull(.Schedule_Name) = False Then
2830        If Trim(.Schedule_Name) <> vbNullString Then
2840          DoCmd.Hourglass True
2850          Application.SysCmd acSysCmdSetStatus, "Validating Schedule Name . . ."
2860          strTmp01 = Trim(.Schedule_Name)
2870          If Len(strTmp01) > 50 Then
2880            DoCmd.Hourglass False
2890            MsgBox "The Schedule Name is limited to 50 characters.", vbInformation + vbOKCancel, "Invalid Entry"
2900            DoCmd.Hourglass True
2910            strTmp01 = Left(strTmp01, 50)
2920            .Schedule_Name = strTmp01
2930          End If
2940          intPos01 = InStr(strTmp01, Chr(34))  ' ** No double-quotes allowed, so just replace with single-quote, apostrophe.
2950          If intPos01 > 0 Then
2960            Do While intPos01 > 0
2970              strTmp01 = Left(strTmp01, (intPos01 - 1)) & "'" & Mid(strTmp01, (intPos01 + 1))
2980              strTmp01 = Trim(strTmp01)
2990              intPos01 = InStr(strTmp01, Chr(34))
3000            Loop
3010            .Schedule_Name = strTmp01
3020          End If
3030          Set dbs = CurrentDb
3040          With dbs
                ' ** Schedule, by specified [schdnam], [schdid].
3050            Set qdf = .QueryDefs("qryFeeSchedule_13")
3060            With qdf.Parameters
3070              ![schdnam] = strTmp01
3080              ![schdid] = Me.Schedule_ID  ' ** Not equal to this, so we don't return this one.
3090            End With
3100            Set rst = qdf.OpenRecordset
3110            With rst
3120              If .BOF = True And .EOF = True Then
                    ' ** All's well.
3130              Else
3140                DoCmd.Hourglass False
3150                MsgBox "This Schedule Name already exists.", vbInformation + vbOKOnly, "Duplicate Schedule Name"
3160                Me.Schedule_Name.Undo
3170                Me.Schedule_Name.SetFocus
3180              End If
3190              .Close
3200            End With
3210            .Close
3220          End With
3230          DoCmd.Hourglass False
3240        Else
3250          .Schedule_Name = Null
3260        End If
3270      End If
3280    End With

EXITP:
3290    Application.SysCmd acSysCmdClearStatus
3300    Set rst = Nothing
3310    Set qdf = Nothing
3320    Set rst = Nothing
3330    Exit Sub

ERRH:
3340    DoCmd.Hourglass False
3350    Select Case ERR.Number
        Case Else
3360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3370    End Select
3380    Resume EXITP

End Sub

Private Sub Schedule_Name_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Name_KeyDown"

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
3490          .Schedule_Base.SetFocus
3500        End With
3510      End Select
3520    End If

        ' ** Shift keys.
3530    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3540      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3550        With Me
3560          intRetVal = 0
3570          .cmdCancel.SetFocus
3580        End With
3590      End Select
3600    End If

EXITP:
3610    KeyCode = intRetVal
3620    Exit Sub

ERRH:
3630    intRetVal = 0
3640    Select Case ERR.Number
        Case Else
3650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3660    End Select
3670    Resume EXITP

End Sub

Private Sub Schedule_Base_AfterUpdate()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Base_AfterUpdate"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

3710    blnContinue = True

3720    With Me
3730      If IsNumeric(.Schedule_Base) Then
3740        If CDbl(.Schedule_Base) > 999999999.99 Then
3750          blnContinue = False
3760          MsgBox "Base Fee is too large.", vbInformation + vbOKOnly, "Invalid Entry"
3770          .Schedule_Base = Null
3780          .Schedule_Base.SetFocus
3790        End If
3800      Else
3810        MsgBox "Please enter a valid Base Fee.", vbInformation + vbOKOnly, "Invalid Entry"
3820        blnContinue = False
3830        .Schedule_Base = Null
3840        .Schedule_Base.SetFocus
3850      End If
3860      If blnContinue = True Then
3870        intPos01 = InStr(1, .Schedule_Base, ".", vbTextCompare)
3880        If intPos01 > 0 Then
3890          If Len(.Schedule_Base) - intPos01 > 2 Then
3900            blnContinue = False
3910            MsgBox "You can only enter Base Fee with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
3920            .Schedule_Base = Null
3930            .Schedule_Base.SetFocus
3940          End If
3950        End If
3960      End If
3970    End With  ' ** Me.

EXITP:
3980    Exit Sub

ERRH:
3990    Select Case ERR.Number
        Case Else
4000      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4010    End Select
4020    Resume EXITP

End Sub

Private Sub Schedule_Minimum_AfterUpdate()

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Minimum_AfterUpdate"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

4110    blnContinue = True

4120    With Me
4130      If IsNumeric(.Schedule_Minimum) Then
4140        If CDbl(.Schedule_Minimum) > 999999999.99 Then
4150          blnContinue = False
4160          MsgBox "Minimum Fee is too large.", vbInformation + vbOKOnly, "Invalid Entry"
4170          .Schedule_Minimum = Null
4180          .Schedule_Minimum.SetFocus
4190        End If
4200      Else
4210        blnContinue = False
4220        MsgBox "Please enter a valid Minimum Fee.", vbInformation + vbOKOnly, "Invalid Entry"
4230        .Schedule_Minimum = Null
4240        .Schedule_Minimum.SetFocus
4250      End If
4260      If blnContinue = True Then
4270        intPos01 = InStr(1, .Schedule_Minimum, ".", vbTextCompare)
4280        If intPos01 > 0 Then
4290          If Len(.Schedule_Minimum) - intPos01 > 2 Then
4300            blnContinue = False
4310            MsgBox "You can only enter Minimum Fee with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
4320            .Schedule_Minimum = Null
4330            .Schedule_Minimum.SetFocus
4340          End If
4350        End If
4360      End If
4370    End With  ' ** Me.

EXITP:
4380    Exit Sub

ERRH:
4390    Select Case ERR.Number
        Case Else
4400      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4410    End Select
4420    Resume EXITP

End Sub

Private Sub Schedule_Minimum_KeyDown(KeyCode As Integer, Shift As Integer)

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Minimum_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
4570        With Me
4580          intRetVal = 0
4590          .cmdOK.SetFocus
4600        End With
4610      End Select
4620    End If

        ' ** Shift keys.
4630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4650        With Me
4660          intRetVal = 0
4670          .Schedule_Base.SetFocus
4680        End With
4690      End Select
4700    End If

EXITP:
4710    KeyCode = intRetVal
4720    Exit Sub

ERRH:
4730    intRetVal = 0
4740    Select Case ERR.Number
        Case Else
4750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4760    End Select
4770    Resume EXITP

End Sub
