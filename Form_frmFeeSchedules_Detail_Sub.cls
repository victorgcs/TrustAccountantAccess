VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFeeSchedules_Detail_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmFeeSchedules_Detail_Sub"

'VGC 06/20/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add Schedule:     A {cmdAdd on frmFeeSchedules}
' **   Base:             B {Schedule_Base on frmFeeSchedules_Sub}
' **   Close:            C {cmdClose on frmFeeSchedules}
' **   Delete Schedule:  D {cmdDelete on frmFeeSchedules}
' **   Add Entry:        E {cmdAddDetail}
' **   Minimum:          I {Schedule_Minimum on frmFeeSchedules_Sub}
' **   Name:             M {Schedule_Name on frmFeeSchedules_Sub}
' **   Delete Entry:     N {cmdDeleteDetail}
' **   Print Report:     P {cmdPrintReport on frmFeeSchedules}
' **   Order:            R {scheddets_order}
' **   Rate:             T {ScheduleDetail_Rate}
' **   Amount:           U {ScheduleDetail_Amount}
' **   Exit:             X {cmdClose on frmFeeSchedules}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[scheddets_order]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine1 As String = "Sort_line1"
Private Const strSortLine2 As String = "Sort_line2"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strCallingForm As String, lngRecsCur As Long, lngTpp As Long
Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Left As Long, lngSortLine_Width As Long
Private THAT_PROC As String, That_Erl As Long
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120   On Error Resume Next
130       strCallingForm = .Parent.Name
140   On Error GoTo ERRH
150       If strCallingForm = "frmFeeSchedules" Then

160         CLR_DISABLED_FG = CLR_DKGRY
170         CLR_DISABLED_BG = CLR_LTTEAL

180         Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
190         lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

200         lngSortLbl_Width = .Sort_lbl.Width

210       Else
220         Cancel = -1
230       End If
240     End With

250     If Cancel = -1 Then
260       Beep
270       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
280     End If

EXITP:
290     Exit Sub

ERRH:
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

        Dim blnDescending As Boolean
        Dim intPos01 As Integer
        Dim strTmp01 As String

410     With Me

420       .Parent.Schedule_ID_display.ForeColor = CLR_DISABLED_FG
430       .Parent.Schedule_ID_display.BackColor = CLR_DISABLED_BG
440       .Detail_Cnt.ForeColor = CLR_DISABLED_FG
450       .Detail_Cnt.BackColor = CLR_DISABLED_BG

460       blnDescending = False

470       Select Case IsNull(.SortNow_OrderBy)
          Case True
480         SortNow THIS_PROC  ' ** Procedure: Below.
490       Case False
            ' ** Restore previous sort.
500         strTmp01 = .SortNow_OrderBy
510         intPos01 = InStr(strTmp01, ",")
520         If intPos01 > 0 Then
530           strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
540           If InStr(strTmp01, " DESC") > 0 Then
550             blnDescending = True
560           End If
570           intPos01 = InStr(strTmp01, " ")
580           If intPos01 > 0 Then
590             strTmp01 = Trim(Left(strTmp01, intPos01))
600           Else
                ' ** Good as-is.
610           End If
620         Else
630           intPos01 = InStr(strTmp01, " ")
640           If intPos01 > 0 Then
650             If InStr(strTmp01, " DESC") > 0 Then
660               blnDescending = True
670             End If
680             strTmp01 = Trim(Left(strTmp01, intPos01))
690           Else
                ' ** Good as-is.
700           End If
710         End If
720         If strTmp01 <> vbNullString Then
730           strTmp01 = Rem_Brackets(strTmp01)  ' ** Module Function: modStringFuncs.
740           Select Case strTmp01
              Case "scheddets_order"
750             SortNow THIS_PROC  ' ** Procedure: Below.
760             DoEvents
770             If blnDescending = True Then
                  ' ** Do it a 2nd time for descending.
780               SortNow "scheddets_order_lbl_DblClick"  ' ** Procedure: Below.
790               DoEvents
800             End If
810           Case "ScheduleDetail_Rate"
820             SortNow "ScheduleDetail_Rate_lbl_DblClick"  ' ** Procedure: Below.
830             DoEvents
840             If blnDescending = True Then
                  ' ** Do it a 2nd time for descending.
850               SortNow "ScheduleDetail_Rate_lbl_DblClick"  ' ** Procedure: Below.
860               DoEvents
870             End If
880           Case "ScheduleDetail_Amount"
890             SortNow "ScheduleDetail_Amount_lbl_DblClick"  ' ** Procedure: Below.
900             DoEvents
910             If blnDescending = True Then
                  ' ** Do it a 2nd time for descending.
920               SortNow "ScheduleDetail_Amount_lbl_DblClick"  ' ** Procedure: Below.
930               DoEvents
940             End If
950           End Select
960         End If
970       End Select

980     End With

EXITP:
990     Exit Sub

ERRH:
1000    Select Case ERR.Number
        Case Else
1010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1020    End Select
1030    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Base:             B {Schedule_Base on frmFeeSchedules_Sub}
        ' **   Minimum:          I {Schedule_Minimum on frmFeeSchedules_Sub}
        ' **   Name:             M {Schedule_Name on frmFeeSchedules_Sub}
        ' **   Order:            R {scheddets_order}
        ' **   Rate:             T {ScheduleDetail_Rate}
        ' **   Amount:           U {ScheduleDetail_Amount}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyEscape
1170        With Me
1180          intRetVal = 0
1190          .Parent.cmdClose_Click  ' ** Form Procedure: frmFeeSchedules.
1200        End With
1210      Case vbKeyUp
1220        intRetVal = 0
1230        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1240      Case vbKeyDown
1250        intRetVal = 0
1260        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1270      End Select
1280    End If

        ' ** Alt keys.
1290    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1300      Select Case intRetVal
          Case vbKeyB
1310        With Me
1320          intRetVal = 0
1330          DoCmd.SelectObject acForm, .Parent.Name, False
1340          .Parent.frmFeeSchedules_Sub.SetFocus
1350          .Parent.frmFeeSchedules_Sub.Form.Schedule_Base.SetFocus
1360        End With
1370      Case vbKeyI
1380        With Me
1390          intRetVal = 0
1400          DoCmd.SelectObject acForm, .Parent.Name, False
1410          .Parent.frmFeeSchedules_Sub.SetFocus
1420          .Parent.frmFeeSchedules_Sub.Form.Schedule_Minimum.SetFocus
1430        End With
1440      Case vbKeyM
1450        With Me
1460          intRetVal = 0
1470          DoCmd.SelectObject acForm, .Parent.Name, False
1480          .Parent.frmFeeSchedules_Sub.SetFocus
1490          .Parent.frmFeeSchedules_Sub.Form.Schedule_Name.SetFocus
1500        End With
1510      Case vbKeyR
1520        With Me
1530          intRetVal = 0
1540          .scheddets_order.SetFocus
1550        End With
1560      Case vbKeyT
1570        With Me
1580          intRetVal = 0
1590          .ScheduleDetail_Rate.SetFocus
1600        End With
1610      Case vbKeyU
1620        With Me
1630          intRetVal = 0
1640          .ScheduleDetail_Amount.SetFocus
1650        End With
1660      Case vbKeyX
1670        With Me
1680          intRetVal = 0
1690          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountAssets.
1700        End With
1710      End Select
1720    End If

        ' ** Ctrl keys.
1730    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
1740      Select Case intRetVal
          Case vbKeyS
1750        intRetVal = 0
1760        cmdSave_Click  ' ** Procedure: Below.
1770      Case vbKeyTab, vbKeyReturn
1780        With Me
1790          intRetVal = 0
1800          DoCmd.SelectObject acForm, .Parent.Name, False
1810          If .Parent.cmdPrintReport.Enabled = True Then
1820            .Parent.cmdPrintReport.SetFocus
1830          Else
1840            .Parent.cmdAdd.SetFocus
1850          End If
1860        End With
1870      Case vbKeyUp, vbKeyPageUp, vbKeyHome
1880        intRetVal = 0
1890        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
1900      Case vbKeyDown, vbKeyPageDown, vbKeyEnd
1910        intRetVal = 0
1920        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
1930      End Select
1940    End If

        ' ** Ctrl-Shift keys.
1950    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1960      Select Case intRetVal
          Case vbKeyF
1970        With Me
1980          intRetVal = 0
1990          DoCmd.SelectObject acForm, .Parent.Name, False
2000          Parent.FocusHolder.SetFocus
2010        End With
2020      Case vbKeyTab, vbKeyReturn
2030        With Me
2040          intRetVal = 0
2050          DoCmd.SelectObject acForm, .Parent.Name, False
2060          .Parent.frmFeeSchedules_Sub.SetFocus
2070          .Parent.frmFeeSchedules_Sub.Form.MoveRec 0, .Schedule_ID  ' ** Form Procedure: frmFeeSchedules_Sub.
2080          .Parent.frmFeeSchedules_Sub.Form.Schedule_Minimum.SetFocus
2090        End With
2100      End Select
2110    End If

EXITP:
2120    KeyCode = intRetVal
2130    Exit Sub

ERRH:
2140    intRetVal = 0
2150    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2160    Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2210    With Me
2220      .ScheduleDetail_DateModified = Now()
2230    End With

EXITP:
2240    Exit Sub

ERRH:
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

2310    Select Case DataErr
        Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
2320      MsgBox "You must enter numeric data.", vbInformation + vbOKOnly, "Invalid Entry"
2330      Response = acDataErrContinue
2340    Case 2046  ' ** The command or action isn't available now (first or last record).
2350      Response = acDataErrContinue
2360    Case 3201
          ' ** It failed to bring over the Schedule_ID that
          ' ** is needed to complete the relationship between
          ' ** Schedule and Schedule Detail.
2370      Me.Schedule_ID = Forms("frmFeeSchedules").frmFeeSchedules_Sub.Form.Schedule_ID
2380      Response = acDataErrContinue
2390    Case Else
2400      If THAT_PROC = vbNullString Then THAT_PROC = "Undefined Form Error"
2410      zErrorHandler THIS_NAME, THAT_PROC, DataErr, That_Erl, ERR.description  ' ** Module Function: modErrorHandler.
2420    End Select

EXITP:
2430    Exit Sub

ERRH:
2440    Select Case ERR.Number
        Case Else
2450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2460    End Select
2470    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Moved to parent Form_Unload().
        'Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
2510    Exit Sub

ERRH:
2520    Select Case ERR.Number
        Case Else
2530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2540    End Select
2550    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2610    DoCmd.RunCommand acCmdSelectRecord
2620    DoCmd.RunCommand acCmdSaveRecord

EXITP:
2630    Exit Sub

ERRH:
2640    Select Case ERR.Number
        Case Else
2650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2660    End Select
2670    Resume EXITP

End Sub

Public Sub cmdAddDetail_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddDetail_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim strDocName As String
        Dim lngScheduleID As Long
        Dim varTmp00 As Variant

2710    With Me

2720      DoCmd.Hourglass True
2730      DoEvents

2740      cmdSave_Click  ' ** Procedure: Above.

2750      lngRecsCur = .Parent.frmFeeSchedules_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Sub.
2760      If lngRecsCur > 0& Then

2770        Set dbs = CurrentDb
            ' ** Delete qryFeeSchedule_16_03 (tblScheduleDetail_Staging, not in tblScheduleDetail).
2780        Set qdf = dbs.QueryDefs("qryFeeSchedule_16_04")
2790        qdf.Execute
2800        Set qdf = Nothing
2810        dbs.Close
2820        Set dbs = Nothing

2830        strDocName = .RecordSource
2840        .RecordSource = vbNullString
2850        ChangeSeed_Ext "ScheduleDetail"  ' ** Module Function: modAutonumberFieldFuncs.
2860        DoEvents
2870        .RecordSource = strDocName
2880        DoEvents

2890        lngScheduleID = .Parent.frmFeeSchedules_Sub.Form.Schedule_ID
2900        varTmp00 = ZeroIfNull(DMax("[scheddets_order]", "tblScheduleDetail_Staging", "[Schedule_ID] = " & CStr(lngScheduleID)))  ' ** Module Function: modStringFuncs.
2910        strDocName = "frmFeeSchedules_Detail_Add"
2920        DoCmd.Hourglass False
2930        DoCmd.OpenForm strDocName, , , , acFormAdd, acDialog, strCallingForm & "~" & _
              CStr(lngScheduleID) & "~" & CStr(varTmp00)

2940      Else
2950        Beep
2960      End If

2970      DoCmd.Hourglass False

2980    End With

EXITP:
2990    Set qdf = Nothing
3000    Set dbs = Nothing
3010    Exit Sub

ERRH:
3020    DoCmd.Hourglass False
3030    Select Case ERR.Number
        Case Else
3040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3050    End Select
3060    Resume EXITP

End Sub

Private Sub cmdAddDetail_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAddDetail_KeyDown"

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
          Case vbKeyTab
3170        With Me
3180          intRetVal = 0
3190          If .cmdDeleteDetail.Enabled = True Then
3200            .cmdDeleteDetail.SetFocus
3210          Else
3220            DoCmd.SelectObject acForm, .Parent.Name, False
3230            .Parent.cmdAdd.SetFocus
3240          End If
3250        End With
3260      End Select
3270    End If

        ' ** Shift keys.
3280    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3290      Select Case intRetVal
          Case vbKeyTab
3300        With Me
3310          intRetVal = 0
3320          lngRecsCur = RecCnt  ' ** Function: Below.
3330          If lngRecsCur > 0& Then
3340            MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below
3350            .ScheduleDetail_Amount.SetFocus
3360          Else
3370            DoCmd.SelectObject acForm, .Parent.Name, False
3380            .Parent.frmFeeSchedules_Sub.SetFocus
3390            .Parent.frmFeeSchedules_Sub.Form.Schedule_Minimum.SetFocus
3400          End If
3410        End With
3420      End Select
3430    End If

EXITP:
3440    KeyCode = intRetVal
3450    Exit Sub

ERRH:
3460    intRetVal = 0
3470    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3480    Case Else
3490      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3500    End Select
3510    Resume EXITP

End Sub

Private Sub cmdDeleteDetail_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDeleteDetail_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

3610    With Me
3620      lngRecsCur = RecCnt  ' ** Function: Below.
3630      If lngRecsCur > 0& Then
3640        If IsNull(.ScheduleDetail_ID) = False Then
3650          If MsgBox("Are you sure you want to delete the '" & _
                  Format(Nz(.ScheduleDetail_Rate, 0), "#0.000#") & "/" & Format(Nz(.ScheduleDetail_Amount, 0), "Currency") & _
                  "' detail entry?", vbQuestion + vbYesNo, "Delete Detail Entry") = vbYes Then
3660            Set dbs = CurrentDb
3670            With dbs
                  ' ** Delete ScheduleDetail, by specified [schddetid].
3680              Set qdf = .QueryDefs("qryFeeSchedule_14")
3690              With qdf.Parameters
3700                ![schddetid] = Me.ScheduleDetail_ID
3710              End With
3720              qdf.Execute dbFailOnError
3730              .Close
3740            End With
3750            .Requery
3760            RecalcSortOrd  ' ** Procedure: Below.
3770            DoEvents
3780            lngRecsCur = RecCnt  ' ** Function: Below.
3790            If lngRecsCur = 0& Then
3800              .cmdDeleteDetail.Enabled = False
3810            End If
3820          End If
3830        Else
3840          Beep
3850        End If
3860      Else
3870        Beep
3880        MsgBox "There is no detail to delete.", vbInformation + vbOKOnly, "Nothing To Do"
3890      End If
3900    End With

EXITP:
3910    Set qdf = Nothing
3920    Set dbs = Nothing
3930    Exit Sub

ERRH:
3940    Select Case ERR.Number
        Case Else
3950      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3960    End Select
3970    Resume EXITP

End Sub

Private Sub cmdDeleteDetail_KeyDown(KeyCode As Integer, Shift As Integer)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDeleteDetail_KeyDown"

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
4090          DoCmd.SelectObject acForm, .Parent.Name, False
4100          If .Parent.cmdPrintReport.Enabled = True Then
4110            .Parent.cmdPrintReport.SetFocus
4120          Else
4130            .Parent.cmdAdd.SetFocus
4140          End If
4150        End With
4160      End Select
4170    End If

        ' ** Shift keys.
4180    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4190      Select Case intRetVal
          Case vbKeyTab
4200        With Me
4210          intRetVal = 0
4220          .cmdAddDetail.SetFocus
4230        End With
4240      End Select
4250    End If

EXITP:
4260    KeyCode = intRetVal
4270    Exit Sub

ERRH:
4280    intRetVal = 0
4290    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4300    Case Else
4310      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4320    End Select
4330    Resume EXITP

End Sub

Private Sub scheddets_order_lbl_DblClick(Cancel As Integer)

4400  On Error GoTo ERRH

        Const THIS_PROC As String = "scheddets_order_lbl_DblClick"

4410    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4420    Exit Sub

ERRH:
4430    Select Case ERR.Number
        Case Else
4440      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4450    End Select
4460    Resume EXITP

End Sub

Private Sub scheddets_order_GotFocus()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "scheddets_order_GotFocus"

4510    With Me
4520      .scheddets_order.SelLength = 0
4530      .scheddets_order.SelStart = 9
4540    End With

EXITP:
4550    Exit Sub

ERRH:
4560    Select Case ERR.Number
        Case Else
4570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4580    End Select
4590    Resume EXITP

End Sub

Private Sub scheddets_order_AfterUpdate()

4600  On Error GoTo ERRH

        Const THIS_PROC As String = "scheddets_order_AfterUpdate"

4610    With Me

4620      cmdSave_Click  ' ** Procedure: Above.

4630      RecalcSortOrd .ScheduleDetail_ID  ' ** Procedure: Below.

4640    End With

EXITP:
4650    Exit Sub

ERRH:
4660    Select Case ERR.Number
        Case Else
4670      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4680    End Select
4690    Resume EXITP

End Sub

Private Sub scheddets_order_KeyDown(KeyCode As Integer, Shift As Integer)

4700  On Error GoTo ERRH

        Const THIS_PROC As String = "scheddets_order_KeyDown"

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
4790  On Error Resume Next
4800          .ScheduleDetail_Rate.SetFocus
4810  On Error GoTo ERRH
4820        End With
4830      End Select
4840    End If

        ' ** Shift keys.
4850    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4860      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4870        With Me
4880          intRetVal = 0
4890          If .CurrentRecord > 1 Then
4900            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4910            .ScheduleDetail_Amount.SetFocus
4920          Else
4930            DoCmd.SelectObject acForm, .Parent.Name, False
4940            .Parent.frmFeeSchedules_Sub.SetFocus
4950            .Parent.frmFeeSchedules_Sub.Form.Schedule_Minimum.SetFocus
4960          End If
4970        End With
4980      End Select
4990    End If

EXITP:
5000    KeyCode = intRetVal
5010    Exit Sub

ERRH:
5020    intRetVal = 0
5030    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5040    Case Else
5050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5060    End Select
5070    Resume EXITP

End Sub

Private Sub ScheduleDetail_Rate_lbl_DblClick(Cancel As Integer)

5100  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Rate_lbl_DblClick"

5110    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5120    Exit Sub

ERRH:
5130    Select Case ERR.Number
        Case Else
5140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5150    End Select
5160    Resume EXITP

End Sub

Private Sub ScheduleDetail_Rate_AfterUpdate()

5200  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Rate_AfterUpdate"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

5210    blnContinue = True

5220    With Me
5230      If IsNull(.ScheduleDetail_Rate) = True Then
5240        .ScheduleDetail_Rate = 0#
5250      End If
5260      If IsNull(.ScheduleDetail_Amount) = True Then
5270        .ScheduleDetail_Amount = 0@
5280      End If
5290      If .ScheduleDetail_Rate = 0 And .ScheduleDetail_Amount = 0 Then
            ' ** OK.
5300      Else
5310        intPos01 = InStr(1, .ScheduleDetail_Rate, ".", vbTextCompare)
5320        If intPos01 > 0 Then
5330          If Len(.ScheduleDetail_Rate) - intPos01 > 6 Then
5340            blnContinue = False
5350            MsgBox "You can only enter a rate with up to 6 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
5360            .ScheduleDetail_Rate = Null
5370            .ScheduleDetail_Rate.SetFocus
5380          End If
5390        End If
5400        If blnContinue = True Then
5410          If .ScheduleDetail_Rate > 1 Then
5420            blnContinue = False
5430            MsgBox "Rate must be <= 100%", vbInformation + vbOKOnly, "Invalid Entry"
5440            .ScheduleDetail_Rate = Null
5450            .ScheduleDetail_Rate.SetFocus
5460          End If
5470        End If
5480      End If
5490    End With  ' ** Me.

EXITP:
5500    Exit Sub

ERRH:
5510    Select Case ERR.Number
        Case Else
5520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5530    End Select
5540    Resume EXITP

End Sub

Private Sub ScheduleDetail_Rate_KeyDown(KeyCode As Integer, Shift As Integer)

5600  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Rate_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5620    intShiftDown = (Shift And acShiftMask) > 0
5630    intAltDown = (Shift And acAltMask) > 0
5640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5660      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5670        With Me
5680          intRetVal = 0
5690          .ScheduleDetail_Amount.SetFocus
5700        End With
5710      End Select
5720    End If

        ' ** Shift keys.
5730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5740      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5750        With Me
5760          intRetVal = 0
5770          .scheddets_order.SetFocus
5780        End With
5790      End Select
5800    End If

EXITP:
5810    KeyCode = intRetVal
5820    Exit Sub

ERRH:
5830    intRetVal = 0
5840    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5850    Case Else
5860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5870    End Select
5880    Resume EXITP

End Sub

Private Sub ScheduleDetail_Amount_lbl_DblClick(Cancel As Integer)

5900  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Amount_lbl_DblClick"

5910    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5920    Exit Sub

ERRH:
5930    Select Case ERR.Number
        Case Else
5940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5950    End Select
5960    Resume EXITP

End Sub

Private Sub ScheduleDetail_Amount_AfterUpdate()

6000  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Amount_AfterUpdate"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

6010    blnContinue = True

6020    With Me
6030      If IsNull(.ScheduleDetail_Amount) = True Then
6040        .ScheduleDetail_Amount = 0@
6050      End If
6060      If IsNull(.ScheduleDetail_Rate) = True Then
6070        .ScheduleDetail_Rate = 0#
6080      End If
6090      If .ScheduleDetail_Amount = 0 And .ScheduleDetail_Rate = 0 Then
6100        blnContinue = False
6110      Else
6120        intPos01 = InStr(1, .ScheduleDetail_Amount, ".", vbTextCompare)
6130        If intPos01 > 0 Then
6140          If Len(.ScheduleDetail_Amount) - intPos01 > 2 Then
6150            blnContinue = False
6160            MsgBox "You can only enter an Amount with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
6170            .ScheduleDetail_Amount = Null
6180            .ScheduleDetail_Amount.SetFocus
6190          End If
6200        End If
6210        If blnContinue = True Then
6220          If .ScheduleDetail_Amount > 999999999.99 Then
6230            blnContinue = False
6240            MsgBox "Amount is too large.", vbInformation + vbOKOnly, "Invalid Entry"
6250            .ScheduleDetail_Amount = Null
6260            .ScheduleDetail_Amount.SetFocus
6270          Else
6280            cmdSave_Click  ' ** Procedure: Above.
6290          End If
6300        End If
6310      End If
6320    End With  ' ** Me.

EXITP:
6330    Exit Sub

ERRH:
6340    THAT_PROC = THIS_PROC
6350    That_Erl = Erl
6360    Form_Error ERR.Number, acDataErrDisplay  ' ** Procedure: Above.
6370    Resume EXITP

End Sub

Private Sub ScheduleDetail_Amount_KeyDown(KeyCode As Integer, Shift As Integer)

6400  On Error GoTo ERRH

        Const THIS_PROC As String = "ScheduleDetail_Amount_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
6470        With Me
6480          intRetVal = 0
6490          lngRecsCur = RecCnt  ' ** Function: Below.
6500          If .CurrentRecord < lngRecsCur Then
6510            MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
6520            .ScheduleDetail_Rate.SetFocus
6530          Else
6540  On Error Resume Next  ' ** I don't know why this won't go there!
6550            .cmdAddDetail.SetFocus
6560  On Error GoTo ERRH
6570          End If
6580        End With
6590      End Select
6600    End If

        ' ** Shift keys.
6610    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
6620      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
6630        With Me
6640          intRetVal = 0
6650          .ScheduleDetail_Rate.SetFocus
6660        End With
6670      End Select
6680    End If

EXITP:
6690    KeyCode = intRetVal
6700    Exit Sub

ERRH:
6710    intRetVal = 0
6720    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
6730    Case Else
6740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6750    End Select
6760    Resume EXITP

End Sub

Public Sub RecalcSortOrd(Optional varThisID As Variant)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "RecalcSortOrd"

        Dim rst1 As DAO.Recordset, rst2 As DAO.Recordset
        Dim lngScheds As Long, arr_varSched() As Variant
        Dim lngThisID As Long, lngLastOrd As Long
        Dim lngRecs As Long, lngScheduleDetailID As Long
        Dim blnFound As Boolean
        Dim varTmp00 As Variant
        Dim lngX As Long, lngY As Long, lngZ As Long, lngE As Long

        ' ** Array: arr_varSched().
        Const S_ELEMS As Integer = 5  ' ** Array's first-element UBound()
        Const S_SID  As Integer = 0
        Const S_SDID As Integer = 1
        Const S_SSID As Integer = 2
        Const S_ORD  As Integer = 3
        Const S_DUPE As Integer = 4
        Const S_NEW  As Integer = 5

6810    With Me
6820      If IsNull(.ScheduleDetail_ID) = False Then

6830  On Error Resume Next
6840        lngScheduleDetailID = .ScheduleDetail_ID
6850        If ERR.Number = 0 Then
6860  On Error GoTo ERRH

              ' ** If a scheddets_order is duplicated, lngThisID takes precedence.
6870          Select Case IsMissing(varThisID)
              Case True
6880            lngThisID = 0&
6890          Case False
6900            lngThisID = varThisID
6910          End Select

6920          lngScheds = 0&
6930          ReDim arr_varSched(S_ELEMS, 0)

6940          Set rst1 = .RecordsetClone
6950          rst1.sort = "[scheddets_order]"
6960          Set rst2 = rst1.OpenRecordset
6970          With rst2

6980            .MoveLast
6990            lngRecs = .RecordCount
7000            .MoveFirst

7010            For lngX = 1& To lngRecs
7020              lngScheds = lngScheds + 1&
7030              lngE = lngScheds - 1&
7040              ReDim Preserve arr_varSched(S_ELEMS, lngE)
                  ' ******************************************************
                  ' ** Array: arr_varSched()
                  ' **
                  ' **   Field  Element  Name                 Constant
                  ' **   =====  =======  ===================  ==========
                  ' **     1       0     Schedule_ID          S_SID
                  ' **     2       1     ScheduleDetail_ID    S_SDID
                  ' **     3       2     scheddets_id         S_SSID
                  ' **     4       3     scheddets_order      S_ORD
                  ' **     5       4     Dupe                 S_DUPE
                  ' **     6       5     New Order            S_NEW
                  ' **
                  ' ******************************************************
7050              arr_varSched(S_SID, lngE) = ![Schedule_ID]
7060              arr_varSched(S_SDID, lngE) = ![ScheduleDetail_ID]
7070              arr_varSched(S_SSID, lngE) = ![scheddets_id]
7080              arr_varSched(S_ORD, lngE) = ![scheddets_order]
7090              arr_varSched(S_DUPE, lngE) = CBool(False)
7100              arr_varSched(S_NEW, lngE) = Null
7110              If lngX < lngRecs Then .MoveNext
7120            Next

7130            .Close
7140          End With
7150          rst1.Close
7160          Set rst1 = Nothing
7170          Set rst2 = Nothing

7180          If lngThisID > 0& Then
                ' ** First, are there dupes?
7190            blnFound = False: lngLastOrd = 0&
7200            For lngX = 0& To (lngScheds - 1&)
7210              If arr_varSched(S_ORD, lngX) <> lngLastOrd Then
7220                lngLastOrd = arr_varSched(S_ORD, lngX)
7230              Else
                    ' ** Yes, there are dupes.
7240                blnFound = True
7250                Exit For
7260              End If
7270            Next
7280            If blnFound = True Then
                  ' ** Find the order that stays.
7290              For lngX = 0& To (lngScheds - 1&)
7300                If arr_varSched(S_SDID, lngX) = lngThisID Then
7310                  lngLastOrd = arr_varSched(S_ORD, lngX)
7320                  arr_varSched(S_NEW, lngX) = lngLastOrd
7330                  Exit For
7340                End If
7350              Next
                  ' ** Change the other.
7360              For lngX = 0& To (lngScheds - 1&)
7370                If arr_varSched(S_SDID, lngX) <> lngThisID Then
7380                  If arr_varSched(S_ORD, lngX) = lngLastOrd Then
7390                    arr_varSched(S_NEW, lngX) = (lngLastOrd + 1&)
7400                    lngLastOrd = arr_varSched(S_NEW, lngX)
7410                    lngThisID = arr_varSched(S_SDID, lngX)
7420                    Exit For
7430                  End If
7440                End If
7450              Next
                  ' ** Now check again.
7460              blnFound = True
7470              Do While blnFound = True
7480                blnFound = False
7490                For lngX = 0& To (lngScheds - 1&)
7500                  If arr_varSched(S_SDID, lngX) <> lngThisID Then
                        ' ** NEW's should be taken care of, and potential conflicts would only be higher numbers.
7510                    If IsNull(arr_varSched(S_NEW, lngX)) = True Then
7520                      If arr_varSched(S_ORD, lngX) = lngLastOrd Then
                            ' ** A new conflict.
7530                        blnFound = True
7540                        arr_varSched(S_NEW, lngX) = (lngLastOrd + 1&)
7550                        lngLastOrd = arr_varSched(S_NEW, lngX)
7560                        lngThisID = arr_varSched(S_SDID, lngX)
7570                        Exit For
7580                      End If
7590                    End If
7600                  End If
7610                Next
7620              Loop
                  ' ** Make sure all have NEW, and this shouldn't create any new conflicts.
7630              For lngX = 0& To (lngScheds - 1&)
7640                If IsNull(arr_varSched(S_NEW, lngX)) = True Then
7650                  arr_varSched(S_NEW, lngX) = arr_varSched(S_ORD, lngX)
7660                End If
7670              Next
7680            End If
7690          End If

              ' ** Binary Sort arr_varSched() array by new order.
7700          For lngX = UBound(arr_varSched, 2) To 1 Step -1
7710            For lngY = 0 To (lngX - 1)
7720              If arr_varSched(S_NEW, lngY) > arr_varSched(S_NEW, (lngY + 1)) Then
7730                For lngZ = 0& To S_ELEMS
7740                  varTmp00 = arr_varSched(lngZ, lngY)
7750                  arr_varSched(lngZ, lngY) = arr_varSched(lngZ, (lngY + 1&))
7760                  arr_varSched(lngZ, (lngY + 1&)) = varTmp00
7770                Next
7780              End If
7790            Next
7800          Next

              ' ** Finally, renumber everything.
              ' ** At this point, there may have been no dupes to begin with,
              ' ** or there were dupes without a preference,
              ' ** or dupes have been handled.
7810          lngLastOrd = 0&
7820          For lngX = 0& To (lngScheds - 1&)
7830            lngLastOrd = lngLastOrd + 1&
7840            arr_varSched(S_NEW, lngX) = lngLastOrd
7850          Next

7860          Set rst1 = .RecordsetClone
7870          With rst1

                ' ** Save the new order.
7880            .MoveFirst
7890            For lngX = 0& To (lngScheds - 1&)
7900              .FindFirst "[scheddets_id] = " & CStr(arr_varSched(S_SSID, lngX))
7910              If .NoMatch = False Then
7920                .Edit  ' ** 3188  Could not update; currently locked by another session on this machine.
7930                ![scheddets_order] = arr_varSched(S_NEW, lngX)
7940                .Update
7950              Else
                    ' ** Shouldn't happen.
7960              End If
7970            Next

7980            .Close
7990          End With
8000          Set rst1 = Nothing

8010          .Requery

8020          MoveRec 0, lngScheduleDetailID  ' ** Procedure: Below.
8030          DoEvents
8040          .scheddets_order.SetFocus

8050        Else
8060  On Error GoTo ERRH
8070        End If

8080      End If
8090    End With

EXITP:
8100    Set rst1 = Nothing
8110    Set rst2 = Nothing
8120    Exit Sub

ERRH:
8130    Select Case ERR.Number
        Case Else
8140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8150    End Select
8160    Resume EXITP

End Sub

Public Function RecCnt() As Long

8200  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

8210    lngRetVal = 0&

8220    With Me
8230      Set rst = .RecordsetClone
8240      With rst
8250        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
8260        Else
8270          .MoveLast
8280          lngRetVal = .RecordCount
8290        End If
8300        .Close
8310      End With
8320    End With

EXITP:
8330    Set rst = Nothing
8340    RecCnt = lngRetVal
8350    Exit Function

ERRH:
8360    lngRetVal = 0&
8370    Select Case ERR.Number
        Case Else
8380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8390    End Select
8400    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

8500  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

8510    Select Case intWhere
        Case acCmdRecordsGoToFirst
8520      DoCmd.RunCommand acCmdRecordsGoToFirst
8530    Case acCmdRecordsGoToLast
8540      DoCmd.RunCommand acCmdRecordsGoToLast
8550    Case acCmdRecordsGoToPrevious
8560      DoCmd.RunCommand acCmdRecordsGoToPrevious
8570    Case acCmdRecordsGoToNext
8580      DoCmd.RunCommand acCmdRecordsGoToNext
8590    Case Else
8600      If IsMissing(varID) = False Then
8610        With Me
8620          Set rst = .RecordsetClone
8630          With rst
8640            .FindFirst "[ScheduleDetail_ID] = " & CStr(varID)
8650            If .NoMatch = False Then
8660              Me.Bookmark = .Bookmark
8670            End If
8680            .Close
8690          End With
8700        End With
8710      End If
8720    End Select

EXITP:
8730    Set rst = Nothing
8740    Exit Sub

ERRH:
8750    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8760    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
8770    Case Else
8780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8790    End Select
8800    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

8900  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [scheddets_order]"

8910    With Me
8920      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
8930        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
8940      End If
8950      .Controls(strSortLbl).Visible = False
8960      .Controls(strSortLine1).Visible = False
8970      .Controls(strSortLine2).Visible = False
8980      .Controls(strSortLine1).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
8990      .Controls(strSortLine2).Width = lngTpp
9000      If strProc = "Form_Load" Then
9010        strCalled = "scheddets_order"
9020        strSortNow = strSortOrig
9030        lngSortLbl_Top = ((.Controls(strCalled & "_lbl").Top - lngTpp) + lngTpp)  ' ** Because of ArialNarrow
9040        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (8& * lngTpp))
9050        lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
9060        lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
9070        .Controls(strSortLbl).Top = lngSortLbl_Top
9080        .Controls(strSortLbl).Left = lngSortLbl_Left
9090        .Controls(strSortLine1).Left = lngSortLine_Left
9100        .Controls(strSortLine2).Left = lngSortLine_Left
9110        .Controls(strSortLine1).Width = lngSortLine_Width
9120        .Controls(strSortLine2).Width = lngSortLine_Width
9130        .Controls(strSortLbl).Caption = strArwUp
9140        .Controls(strSortLbl).ForeColor = CLR_DKBLU
9150      Else
9160        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
9170        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
9180        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
9190        lngSortLine_Left = .Controls(strCalled & "_lbl_line").Left
9200        lngSortLine_Width = .Controls(strCalled & "_lbl_line").Width
9210        .Controls(strSortLbl).Top = lngSortLbl_Top
9220        .Controls(strSortLbl).Left = lngSortLbl_Left
9230        .Controls(strSortLine1).Left = lngSortLine_Left
9240        .Controls(strSortLine2).Left = lngSortLine_Left
9250        .Controls(strSortLine1).Width = lngSortLine_Width
9260        .Controls(strSortLine2).Width = lngSortLine_Width
9270        If strCalled = "scheddets_order" Then
9280          lngSortLbl_Top = (lngSortLbl_Top + lngTpp)
9290          lngSortLbl_Left = (lngSortLbl_Left + (8& * lngTpp))
9300          .Controls(strSortLbl).Top = lngSortLbl_Top
9310          .Controls(strSortLbl).Left = lngSortLbl_Left
9320          strSortAsc = strSortOrig
9330          If strSortNow = strSortAsc Then
9340            strSortNow = "[scheddets_order] DESC"
9350            .Controls(strSortLbl).Caption = strArwDn
9360            .Controls(strSortLbl).ForeColor = CLR_DKRED
9370          Else
9380            strSortNow = strSortAsc
9390            .Controls(strSortLbl).Caption = strArwUp
9400            .Controls(strSortLbl).ForeColor = CLR_DKBLU
9410          End If
9420        Else
9430          .Controls(strSortLbl).Caption = strArwUp
9440          .Controls(strSortLbl).ForeColor = CLR_DKBLU
9450          Select Case strCalled
              Case "ScheduleDetail_Rate"
9460            strSortAsc = "[ScheduleDetail_Rate]" & strStdAsc
9470          Case "ScheduleDetail_Amount"
9480            strSortAsc = "[ScheduleDetail_Amount]" & strStdAsc
9490          End Select
9500          If strSortNow = strSortAsc Then
9510            intCnt = CharCnt(strSortAsc, ",") + 1  ' ** Module Function: modStringFuncs.
9520            Select Case intCnt
                Case 1
9530              strTmp01 = strSortAsc & " DESC"
9540            Case 2
9550              intPos01 = InStr(strSortAsc, ",")
9560              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9570              strTmp01 = strTmp01 & Mid(strSortAsc, intPos01) & " DESC"
9580            Case 3
9590              intPos01 = InStr(strSortAsc, ",")
9600              strTmp01 = Left(strSortAsc, (intPos01 - 1)) & " DESC"
9610              strTmp02 = Mid(strSortAsc, intPos01)
9620              intPos01 = InStr(2, strTmp02, ",")
9630              strTmp01 = strTmp01 & Left(strTmp02, (intPos01 - 1)) & " DESC"
9640              strTmp01 = strTmp01 & Mid(strTmp02, intPos01) & " DESC"
9650            End Select
9660            strSortNow = strTmp01
9670            .Controls(strSortLbl).Caption = strArwDn
9680            .Controls(strSortLbl).ForeColor = CLR_DKRED
9690          Else
9700            strSortNow = strSortAsc
9710          End If
9720        End If
9730      End If
9740      .Controls(strSortLbl).Visible = True
9750      .Controls(strSortLine1).Visible = True
9760      .Controls(strSortLine2).Visible = True
9770      .OrderBy = strSortNow
9780      .OrderByOn = True
9790      .SortNow_OrderBy = strSortNow
9800    End With

EXITP:
9810    Exit Sub

ERRH:
9820    Select Case ERR.Number
        Case Else
9830      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
9840    End Select
9850    Resume EXITP

End Sub
