VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFeeSchedules_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmFeeSchedules_Sub"

'VGC 10/26/2017: CHANGES!

' ** NOTE: This form's RecordSource, qryFeeSchedule_06, has a DLookup() to qryFeeSchedule_05.

' ** Shortcut Alt keys responsive from this form:
' **   Add Schedule:     A {cmdAdd on frmFeeSchedules}
' **   Base:             B {Schedule_Base}
' **   Close:            C {cmdClose on frmFeeSchedules}
' **   Delete Schedule:  D {cmdDelete on frmFeeSchedules}
' **   Add Entry:        E {cmdAddDetail on frmFeeSchedules_Detail_Sub}
' **   Minimum:          I {Schedule_Minimum}
' **   Name:             M {Schedule_Name}
' **   Delete Entry:     N {cmdDeleteDetail on frmFeeSchedules_Detail_Sub}
' **   Print Report:     P {cmdPrintReport on frmFeeSchedules}
' **   Order:            R {scheddets_order on frmFeeSchedules_Detail_Sub}
' **   Rate:             T {ScheduleDetail_Rate on frmFeeSchedules_Detail_Sub}
' **   Amount:           U {ScheduleDetail_Amount on frmFeeSchedules_Detail_Sub}
' **   Exit:             X {cmdClose on frmFeeSchedules}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private Const strSortOrig As String = "[Schedule_ID]"
Private Const strDblClick As String = "_lbl_DblClick"
Private Const strSortLine1 As String = "Sort_line1"
Private Const strSortLine2 As String = "Sort_line2"
Private Const strSortLbl As String = "Sort_lbl"
Private Const strArwUp As String = "­"  ' ** ASCII = 173, Font = Symbol.
Private Const strArwDn As String = "¯"  ' ** ASCII = 175, Font = Symbol.

Private strSortNow As String, lngSortLbl_Top As Long, lngSortLbl_Left As Long, lngSortLbl_Width As Long
Private lngSortLine_Top As Long, lngSortLine_Left As Long, lngSortLine_Width As Long
Private strCallingForm As String, lngRecsCur As Long, lngTpp As Long
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

420       .Schedule_ID.ForeColor = CLR_DISABLED_FG
430       .Schedule_ID.BackColor = CLR_DISABLED_BG

440       blnDescending = False

450       Select Case IsNull(.SortNow_OrderBy)
          Case True
460         SortNow THIS_PROC  ' ** Procedure: Below.
470       Case False
            ' ** Restore previous sort.
480         strTmp01 = .SortNow_OrderBy
490         intPos01 = InStr(strTmp01, ",")
500         If intPos01 > 0 Then
510           strTmp01 = Trim(Left(strTmp01, (intPos01 - 1)))
520           If InStr(strTmp01, " DESC") > 0 Then
530             blnDescending = True
540           End If
550           intPos01 = InStr(strTmp01, " ")
560           If intPos01 > 0 Then
570             strTmp01 = Trim(Left(strTmp01, intPos01))
580           Else
                ' ** Good as-is.
590           End If
600         Else
610           intPos01 = InStr(strTmp01, " ")
620           If intPos01 > 0 Then
630             If InStr(strTmp01, " DESC") > 0 Then
640               blnDescending = True
650             End If
660             strTmp01 = Trim(Left(strTmp01, intPos01))
670           Else

680           End If
690         End If
700         If strTmp01 <> vbNullString Then
710           strTmp01 = Rem_Brackets(strTmp01)  ' ** Module Function: modStringFuncs.
720           Select Case strTmp01
              Case "Schedule_ID"
730             SortNow THIS_PROC  ' ** Procedure: Below.
740             DoEvents
750             If blnDescending = True Then
                  ' ** Do it a 2nd time for descending.
760               SortNow "Schedule_ID_lbl_DblClick"  ' ** Procedure: Below.
770               DoEvents
780             End If
790           Case "Schedule_Name"
800             SortNow "Schedule_Name_lbl_DblClick"  ' ** Procedure: Below.
810             DoEvents
820             If blnDescending = True Then
830               SortNow "Schedule_Name_lbl_DblClick"  ' ** Procedure: Below.
840               DoEvents
850             End If
860           Case "Schedule_Base"
870             SortNow "Schedule_Base_lbl_DblClick"  ' ** Procedure: Below.
880             DoEvents
890             If blnDescending = True Then
900               SortNow "Schedule_Base_lbl_DblClick"  ' ** Procedure: Below.
910               DoEvents
920             End If
930           Case "Schedule_Minimum"
940             SortNow "Schedule_Minimum_lbl_DblClick"  ' ** Procedure: Below.
950             DoEvents
960             If blnDescending = True Then
970               SortNow "Schedule_Minimum_lbl_DblClick"  ' ** Procedure: Below.
980               DoEvents
990             End If
1000          End Select
1010        End If
1020      End Select

1030  On Error Resume Next
1040      .Schedule_Name.SetFocus
1050  On Error GoTo ERRH

1060    End With

EXITP:
1070    Exit Sub

ERRH:
1080    Select Case ERR.Number
        Case Else
1090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1100    End Select
1110    Resume EXITP

End Sub

Private Sub Form_Current()

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Current"

        Dim varTmp00 As Variant

1210    With Me
1220      .Parent.frmFeeSchedules_Detail_Sub.Form.Requery
1230      .Parent.frmFeeSchedules_Detail_Sub.Form.scheddets_order.Visible = True
1240      .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Rate.Visible = True
1250      .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Amount.Visible = True
1260      If IsNull(.Schedule_Name) = False And IsNull(.Schedule_ID) = False Then
1270        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.Enabled = True
1280      Else
1290        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.Enabled = False
1300        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = False
1310      End If
1320      If IsNull(.Detail_Cnt) = False Then
            ' ** Access gets all bolloxed-up here when adding a new schedule.
1330  On Error Resume Next
1340        varTmp00 = .Detail_Cnt  'Can't find 'Detail_Cnt'.
1350        lngRecsCur = .Parent.frmFeeSchedules_Detail_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Detail_Sub.
1360        If lngRecsCur > 0& Then
1370          .Parent.frmFeeSchedules_Detail_Sub.Form.Detail_Cnt = varTmp00  'The value you entered isn't valid for the field.
1380          If .Detail_Cnt > 0& Then
1390            .Parent.frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = True
1400          Else
1410            .Parent.frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = False
1420          End If
1430        End If
1440  On Error GoTo ERRH
1450      Else
1460        .Parent.frmFeeSchedules_Detail_Sub.Form.Detail_Cnt = 0&
1470        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = False
1480      End If
1490      DoEvents
1500  On Error Resume Next
1510      .Parent.Schedule_ID_display = .Schedule_ID  'The value you entered isn't valid for the field.
1520  On Error GoTo ERRH
1530      .Parent.frmFeeSchedules_Detail_Sub.Form.Refresh
1540    End With

EXITP:
1550    Exit Sub

ERRH:
1560    Select Case ERR.Number
        Case Else
1570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1580    End Select
1590    Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1620    intShiftDown = (Shift And acShiftMask) > 0
1630    intAltDown = (Shift And acAltMask) > 0
1640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Base:             B {Schedule_Base}
        ' **   Minimum:          I {Schedule_Minimum}
        ' **   Name:             M {Schedule_Name}
        ' **   Order:            R {scheddets_order on frmFeeSchedules_Detail_Sub}
        ' **   Rate:             T {ScheduleDetail_Rate on frmFeeSchedules_Detail_Sub}
        ' **   Amount:           U {ScheduleDetail_Amount on frmFeeSchedules_Detail_Sub}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
1650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyEscape
1670        With Me
1680          intRetVal = 0
1690          .Parent.cmdClose_Click  ' ** Form Procedure: frmFeeSchedules.
1700        End With
1710      Case vbKeyUp
1720        intRetVal = 0
1730        MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
1740      Case vbKeyDown
1750        intRetVal = 0
1760        MoveRec acCmdRecordsGoToNext  ' ** Procedure: Below.
1770      End Select
1780    End If

        ' ** Alt keys.
1790    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1800      Select Case intRetVal
          Case vbKeyB
1810        With Me
1820          intRetVal = 0
1830          .Schedule_Base.SetFocus
1840        End With
1850      Case vbKeyI
1860        With Me
1870          intRetVal = 0
1880          .Schedule_Minimum.SetFocus
1890        End With
1900      Case vbKeyM
1910        With Me
1920          intRetVal = 0
1930          .Schedule_Name.SetFocus
1940        End With
1950      Case vbKeyR
1960        With Me
1970          intRetVal = 0
1980          DoCmd.SelectObject acForm, .Parent.Name, False
1990          .Parent.frmFeeSchedules_Detail_Sub.SetFocus
2000          .Parent.frmFeeSchedules_Detail_Sub.Form.scheddets_order.SetFocus
2010        End With
2020      Case vbKeyT
2030        With Me
2040          intRetVal = 0
2050          DoCmd.SelectObject acForm, .Parent.Name, False
2060          .Parent.frmFeeSchedules_Detail_Sub.SetFocus
2070          .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Rate.SetFocus
2080        End With
2090      Case vbKeyU
2100        With Me
2110          intRetVal = 0
2120          DoCmd.SelectObject acForm, .Parent.Name, False
2130          .Parent.frmFeeSchedules_Detail_Sub.SetFocus
2140          .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Amount.SetFocus
2150        End With
2160      Case vbKeyX
2170        With Me
2180          intRetVal = 0
2190          .Parent.cmdClose_Click  ' ** Form Procedure: frmAccountAssets.
2200        End With
2210      End Select
2220    End If

        ' ** Ctrl keys.
2230    If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
2240      Select Case intRetVal
          Case vbKeyS
2250        intRetVal = 0
2260        cmdSave_Click  ' ** Procedure: Below.
2270      Case vbKeyTab, vbKeyReturn
2280        With Me
2290          intRetVal = 0
2300          DoCmd.SelectObject acForm, .Parent.Name, False
2310          .Parent.frmFeeSchedules_Detail_Sub.SetFocus
2320          lngRecsCur = .Parent.frmFeeSchedules_Detail_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Detail_Sub.
2330          If lngRecsCur > 0& Then
2340            .Parent.frmFeeSchedules_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmFeeSchedules_Detail_Sub.
2350            .Parent.frmFeeSchedules_Detail_Sub.Form.scheddets_order.SetFocus
2360          Else
2370            .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.SetFocus
2380            .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail_Click  ' ** Form Procedure: frmFeeSchedules_Detail_Sub.
2390          End If
2400        End With
2410      Case vbKeyUp, vbKeyPageUp, vbKeyHome
2420        intRetVal = 0
2430        MoveRec acCmdRecordsGoToFirst  ' ** Procedure: Below.
2440      Case vbKeyDown, vbKeyPageDown, vbKeyEnd
2450        intRetVal = 0
2460        MoveRec acCmdRecordsGoToLast  ' ** Procedure: Below.
2470      End Select
2480    End If

        ' ** Ctrl-Shift keys.
2490    If intCtrlDown And (Not intAltDown) And intShiftDown Then
2500      Select Case intRetVal
          Case vbKeyF
2510        With Me
2520          intRetVal = 0
2530          DoCmd.SelectObject acForm, .Parent.Name, False
2540          Parent.FocusHolder.SetFocus
2550        End With
2560      Case vbKeyTab, vbKeyReturn
2570        With Me
2580          intRetVal = 0
2590          DoCmd.SelectObject acForm, .Parent.Name, False
2600          .Parent.cmdClose.SetFocus
2610        End With
2620      End Select
2630    End If

EXITP:
2640    KeyCode = intRetVal
2650    Exit Sub

ERRH:
2660    intRetVal = 0
2670    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2680    Case Else
2690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2700    End Select
2710    Resume EXITP

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_BeforeUpdate"

2810    With Me
2820      If IsNull(.Schedule_DateCreated) = True Then
2830        .Schedule_DateCreated = Now()
2840      End If
2850      .Schedule_DateModified = Now()
2860    End With

EXITP:
2870    Exit Sub

ERRH:
2880    Select Case ERR.Number
        Case Else
2890      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2900    End Select
2910    Resume EXITP

End Sub

Private Sub Form_AfterUpdate()

3000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_AfterUpdate"

3010    With Me
3020      If IsNull(.Schedule_Name) = False Then
3030        If .Schedule_Name <> vbNullString Then
3040          If IsNull(.Schedule_ID) = False Then
3050            If .Schedule_ID > 0& Then
3060              .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.Enabled = True
3070            End If
3080          End If
3090        End If
3100      End If
3110    End With

EXITP:
3120    Exit Sub

ERRH:
3130    Select Case ERR.Number
        Case Else
3140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3150    End Select
3160    Resume EXITP

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Error"

3210    If DataErr <> 0 Then
3220      Select Case DataErr
          Case 3075, 8530, 2501, 3101, 8519, 2108, 2116, 3020, 2169, 7753, 3314, 2237
            ' ** 2501: The '|' action was Canceled.
            ' ** Do nothing.
3230      Case 2113, 2279  ' ** 2113: The value you entered isn't valid for this field.  ' ** 2279: The value you entered isn't appropriate for the input mask '|' specified for this field.
3240        Select Case Me.ActiveControl.Name
            Case "Schedule_Name"
3250          MsgBox "Please enter a valid Schedule Name.", vbInformation + vbOKOnly, "Invalid Entry"
3260        Case "Schedule_Base"
3270          MsgBox "Please enter a valid Base Fee.", vbInformation + vbOKOnly, "Invalid Entry"
3280        Case "Schedule_Minimum"
3290          MsgBox "Please enter a valid Minimum Fee.", vbInformation + vbOKOnly, "Invalid Entry"
3300        End Select
3310        Response = acDataErrContinue
3320      Case 2046  ' ** The command or action isn't available now (first or last record).
3330        MsgBox "There are no records to delete.", vbInformation + vbOKOnly, ("Nothing To Do" & Space(40))
3340        Response = acDataErrContinue
3350      Case 3162
3360        MsgBox "The " & Me.ActiveControl.Name & " field cannot be left blank.", vbInformation + vbOKOnly, "Entry Required"
3370        DoCmd.CancelEvent
3380      Case 3022  ' ** The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship.
3390        MsgBox "Schedule name already exists.", vbInformation + vbOKOnly, "Duplicate Schedule Name"
3400        DoCmd.CancelEvent
3410      Case Else
3420        zErrorHandler Me.Name, "Undefined Form Error", DataErr
3430      End Select
3440    End If

EXITP:
3450    Exit Sub

ERRH:
3460    Select Case ERR.Number
        Case Else
3470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3480    End Select
3490    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Moved to parent Form_Unload().
        'Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

EXITP:
3510    Exit Sub

ERRH:
3520    Select Case ERR.Number
        Case Else
3530      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3540    End Select
3550    Resume EXITP

End Sub

Public Sub cmdSave_Click()

3600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

3610    DoCmd.RunCommand acCmdSaveRecord

EXITP:
3620    Exit Sub

ERRH:
3630    Select Case ERR.Number
        Case Else
3640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3650    End Select
3660    Resume EXITP

End Sub

Private Sub Schedule_ID_lbl_DblClick(Cancel As Integer)

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_ID_lbl_DblClick"

3710    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3720    Exit Sub

ERRH:
3730    Select Case ERR.Number
        Case Else
3740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3750    End Select
3760    Resume EXITP

End Sub

Private Sub Schedule_Name_lbl_DblClick(Cancel As Integer)

3800  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Name_lbl_DblClick"

3810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
3820    Exit Sub

ERRH:
3830    Select Case ERR.Number
        Case Else
3840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3850    End Select
3860    Resume EXITP

End Sub

Private Sub Schedule_Name_GotFocus()

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Name_GotFocus"

3910    With Me
3920      .Schedule_Name.SelLength = 0
3930      .Schedule_Name.SelStart = 99
3940    End With

EXITP:
3950    Exit Sub

ERRH:
3960    Select Case ERR.Number
        Case 2185  ' ** You can't reference a property or method for a control unless the control has the focus.
          ' ** No schedules, ignore.
3970    Case Else
3980      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3990    End Select
4000    Resume EXITP

End Sub

Private Sub Schedule_Name_KeyDown(KeyCode As Integer, Shift As Integer)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Name_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
4170        With Me
4180          intRetVal = 0
4190          .Schedule_Base.SetFocus
4200        End With
4210      End Select
4220    End If

        ' ** Shift keys.
4230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
4240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
4250        With Me
4260          intRetVal = 0
4270          If .CurrentRecord > 1 Then
4280            MoveRec acCmdRecordsGoToPrevious  ' ** Procedure: Below.
4290            .Schedule_Minimum.SetFocus
4300          Else
4310            DoCmd.SelectObject acForm, .Parent.Name, False
4320            .Parent.cmdClose.SetFocus
4330          End If
4340        End With
4350      End Select
4360    End If

EXITP:
4370    KeyCode = intRetVal
4380    Exit Sub

ERRH:
4390    intRetVal = 0
4400    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
4410    Case Else
4420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4430    End Select
4440    Resume EXITP

End Sub

Private Sub Schedule_Name_Change()

4500  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Name_Change"

4510    With Me
4520      .Parent.frmFeeSchedules_Detail_Sub.Form.Requery
4530      DoEvents
4540      lngRecsCur = .Parent.frmFeeSchedules_Detail_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Detail_Sub.
4550      If lngRecsCur > 0& Then
4560        .Parent.frmFeeSchedules_Detail_Sub.Form.scheddets_order.Visible = True
4570        .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Rate.Visible = True
4580        .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Amount.Visible = True
4590        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.Enabled = True
4600        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = True
4610      Else
4620        .Parent.frmFeeSchedules_Detail_Sub.Form.scheddets_order.Visible = False
4630        .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Rate.Visible = False
4640        .Parent.frmFeeSchedules_Detail_Sub.Form.ScheduleDetail_Amount.Visible = False
4650        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.Enabled = True
4660        .Parent.frmFeeSchedules_Detail_Sub.Form.cmdDeleteDetail.Enabled = False
4670      End If
4680    End With

EXITP:
4690    Exit Sub

ERRH:
4700    Select Case ERR.Number
        Case Else
4710      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4720    End Select
4730    Resume EXITP

End Sub

Private Sub Schedule_Base_lbl_DblClick(Cancel As Integer)

4800  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Base_lbl_DblClick"

4810    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
4820    Exit Sub

ERRH:
4830    Select Case ERR.Number
        Case Else
4840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4850    End Select
4860    Resume EXITP

End Sub

Private Sub Schedule_Base_AfterUpdate()

4900  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Base_AfterUpdate"

        Dim intPos01 As Integer
        Dim blnContinue As Boolean

4910    blnContinue = True

4920    With Me
4930      If IsNumeric(.Schedule_Base) Then
4940        If CDbl(.Schedule_Base) > 999999999.99 Then  ' ** TO THE NINES! (Hopefully, not at sixes and sevens.)
4950          blnContinue = False
4960          MsgBox "Base Fee is too large.", vbInformation + vbOKOnly, "Invalid Entry"
4970          .Schedule_Base = Null
4980          .Schedule_Base.SetFocus
4990        End If
5000      Else
5010        MsgBox "Please enter a valid Base Fee.", vbInformation + vbOKOnly, "Invalid Entry"
5020        blnContinue = False
5030        .Schedule_Base = Null
5040        .Schedule_Base.SetFocus
5050      End If

5060      If blnContinue = True Then
5070        intPos01 = InStr(1, .Schedule_Base, ".", vbTextCompare)
5080        If intPos01 > 0 Then
5090          If Len(.Schedule_Base) - intPos01 > 2 Then
5100            blnContinue = False
5110            MsgBox "You can only enter Base Fee with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
5120            .Schedule_Base = Null
5130            .Schedule_Base.SetFocus
5140          End If
5150        End If
5160      End If

5170      If blnContinue = True Then
5180        If Not IsNumeric(.Schedule_Base) Then
5190          MsgBox "Please enter a valid Base Fee.", vbInformation + vbOKOnly, "Invalid Entry"
5200          .Schedule_Base = Null
5210          .Schedule_Base.SetFocus
5220        End If
5230      End If

5240    End With  ' ** Me.

EXITP:
5250    Exit Sub

ERRH:
5260    Select Case ERR.Number
        Case Else
5270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5280    End Select
5290    Resume EXITP

End Sub

Private Sub Schedule_Minimum_lbl_DblClick(Cancel As Integer)

5300  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Minimum_lbl_DblClick"

5310    SortNow THIS_PROC  ' ** Procedure: Below.

EXITP:
5320    Exit Sub

ERRH:
5330    Select Case ERR.Number
        Case Else
5340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5350    End Select
5360    Resume EXITP

End Sub

Private Sub Schedule_Minimum_KeyDown(KeyCode As Integer, Shift As Integer)

5400  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Minimum_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

5410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
5420    intShiftDown = (Shift And acShiftMask) > 0
5430    intAltDown = (Shift And acAltMask) > 0
5440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
5450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
5460      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5470        With Me
5480          intRetVal = 0
5490          lngRecsCur = RecCnt  ' ** Function: Below.
5500          DoCmd.SelectObject acForm, .Parent.Name, False
5510          .Parent.frmFeeSchedules_Detail_Sub.SetFocus
5520          lngRecsCur = .Parent.frmFeeSchedules_Detail_Sub.Form.RecCnt  ' ** Form Function: frmFeeSchedules_Detail_Sub.
5530          If lngRecsCur > 0& Then
5540            .Parent.frmFeeSchedules_Detail_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmFeeSchedules_Detail_Sub.
5550            .Parent.frmFeeSchedules_Detail_Sub.Form.scheddets_order.SetFocus
5560          Else
5570            .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail.SetFocus
5580            .Parent.frmFeeSchedules_Detail_Sub.Form.cmdAddDetail_Click  ' ** Form Procedure: frmFeeSchedules_Detail_Sub.
5590          End If
5600        End With
5610      End Select
5620    End If

        ' ** Shift keys.
5630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
5640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
5650        With Me
5660          intRetVal = 0
5670          .Schedule_Base.SetFocus
5680        End With
5690      End Select
5700    End If

EXITP:
5710    KeyCode = intRetVal
5720    Exit Sub

ERRH:
5730    intRetVal = 0
5740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
5750    Case Else
5760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
5770    End Select
5780    Resume EXITP

End Sub

Private Sub Schedule_Minimum_AfterUpdate()

5800  On Error GoTo ERRH

        Const THIS_PROC As String = "Schedule_Minimum_AfterUpdate"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngScheduleID As Long
        Dim strDocName As String
        Dim intPos01 As Integer
        Dim blnContinue As Boolean

5810    blnContinue = True

5820    With Me

5830      If IsNumeric(.Schedule_Minimum) Then
5840        If CDbl(.Schedule_Minimum) > 999999999.99 Then  ' ** TO THE NINES! (Hopefully, not at sixes and sevens.)
5850          blnContinue = False
5860          MsgBox "Minimum Fee is too large.", vbInformation + vbOKOnly, "Invalid Entry"
5870          .Schedule_Minimum = Null
5880          .Schedule_Minimum.SetFocus
5890        End If
5900      Else
5910        blnContinue = False
5920        MsgBox "Please enter a valid Minimum Fee.", vbInformation + vbOKOnly, "Invalid Entry"
5930        .Schedule_Minimum = Null
5940        .Schedule_Minimum.SetFocus
5950      End If

5960      If blnContinue = True Then
5970        intPos01 = InStr(1, .Schedule_Minimum, ".", vbTextCompare)
5980        If intPos01 > 0 Then
5990          If Len(.Schedule_Minimum) - intPos01 > 2 Then
6000            blnContinue = False
6010            MsgBox "You can only enter Minimum Fee with up to 2 decimals.", vbInformation + vbOKOnly, "Invalid Entry"
6020            .Schedule_Minimum = Null
6030            .Schedule_Minimum.SetFocus
6040          End If
6050        End If
6060      End If

6070      If blnContinue = True Then
6080        If IsNull(.Schedule_Name) = False Then
6090          If .Schedule_Name <> vbNullString Then
6100            cmdSave_Click  ' ** Procedure: Above.
6110            Set dbs = CurrentDb
6120            With dbs
                  ' ** ScheduleDetail, by specified [schdid].
6130              Set qdf = .QueryDefs("qryFeeSchedule_11")
6140              With qdf.Parameters
6150                ![schdid] = Me.Schedule_ID
6160              End With
6170              Set rst = qdf.OpenRecordset
6180              With rst
6190                If .BOF = True And .EOF = True Then
                      ' ** No records, new schedule.
6200                Else
                      ' ** Detail exists, so just an edit
6210                  blnContinue = False
6220                End If
6230                .Close
6240              End With
6250              .Close
6260            End With  ' ** dbs.
6270          End If
6280        End If
6290      End If  ' ** blnContinue.

6300      If blnContinue = True Then
6310        lngScheduleID = .Schedule_ID
6320        strDocName = "frmFeeSchedules_Detail_Add"
6330        DoCmd.OpenForm strDocName, , , , acFormAdd, , strCallingForm & "~" & CStr(lngScheduleID) & "~0"
6340      End If

6350    End With  ' ** Me.

EXITP:
6360    Set rst = Nothing
6370    Set qdf = Nothing
6380    Set dbs = Nothing
6390    Exit Sub

ERRH:
6400    Select Case ERR.Number
        Case Else
6410      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6420    End Select
6430    Resume EXITP

End Sub

Public Function RecCnt() As Long

6500  On Error GoTo ERRH

        Const THIS_PROC As String = "RecCnt"

        Dim rst As DAO.Recordset
        Dim lngRetVal As Long

6510    lngRetVal = 0&

6520    With Me
6530      Set rst = .RecordsetClone
6540      With rst
6550        If .BOF = True And .EOF = True Then
              ' ** Shouldn't really be here if that's the case.
6560        Else
6570          .MoveLast
6580          lngRetVal = .RecordCount
6590        End If
6600        .Close
6610      End With
6620    End With

EXITP:
6630    Set rst = Nothing
6640    RecCnt = lngRetVal
6650    Exit Function

ERRH:
6660    lngRetVal = 0&
6670    Select Case ERR.Number
        Case Else
6680      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
6690    End Select
6700    Resume EXITP

End Function

Public Sub MoveRec(intWhere As AcCommand, Optional varID As Variant)

6800  On Error GoTo ERRH

        Const THIS_PROC As String = "MoveRec"

        Dim rst As DAO.Recordset

6810    Select Case intWhere
        Case acCmdRecordsGoToFirst
6820      DoCmd.RunCommand acCmdRecordsGoToFirst
6830    Case acCmdRecordsGoToLast
6840      DoCmd.RunCommand acCmdRecordsGoToLast
6850    Case acCmdRecordsGoToPrevious
6860      DoCmd.RunCommand acCmdRecordsGoToPrevious
6870    Case acCmdRecordsGoToNext
6880      DoCmd.RunCommand acCmdRecordsGoToNext
6890    Case Else
6900      If IsMissing(varID) = False Then
6910        With Me
6920          Set rst = .RecordsetClone
6930          With rst
6940            .FindFirst "[Schedule_ID] = " & CStr(varID)
6950            If .NoMatch = False Then
6960              Me.Bookmark = .Bookmark
6970            End If
6980            .Close
6990          End With
7000        End With
7010      End If
7020    End Select

EXITP:
7030    Set rst = Nothing
7040    Exit Sub

ERRH:
7050    Select Case ERR.Number
        Case 3021  ' ** No current record.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7060    Case 3420  ' ** Object invalid or no longer set.
          ' ** Ignore. If this gets called before the subform is fully populated, it can error.
7070    Case Else
7080      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
7090    End Select
7100    Resume EXITP

End Sub

Public Sub SortNow(strProc As String)

7200  On Error GoTo ERRH

        Const THIS_PROC As String = "SortNow"

        Dim ctl As Access.Control
        Dim strCalled As String, strSortAsc As String
        Dim intPos01 As Integer
        Dim strTmp01 As String, strTmp02 As String

        Const strStdAsc As String = ", [Schedule_Name]"

7210    With Me
7220      .Controls(strSortLbl).Visible = False
7230      .Controls(strSortLine1).Visible = False
7240      .Controls(strSortLine2).Visible = False
7250      .Controls(strSortLine1).Width = lngTpp  ' ** So it doesn't push off the right side of the form.
7260      .Controls(strSortLine2).Width = lngTpp
7270      If strProc = "Form_Load" Then
7280        strCalled = "Schedule_ID"
7290        strSortNow = strSortOrig
7300        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7310        lngSortLbl_Left = (((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width) + (7& * lngTpp))
7320        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7330        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7340        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7350        .Controls(strSortLbl).Top = lngSortLbl_Top
7360        .Controls(strSortLbl).Left = lngSortLbl_Left
7370        .Controls(strSortLine1).Top = lngSortLine_Top
7380        .Controls(strSortLine2).Top = (lngSortLine_Top + lngTpp)
7390        .Controls(strSortLine1).Left = lngSortLine_Left
7400        .Controls(strSortLine2).Left = lngSortLine_Left
7410        .Controls(strSortLine1).Width = lngSortLine_Width
7420        .Controls(strSortLine2).Width = lngSortLine_Width
7430        .Controls(strSortLbl).Caption = strArwUp
7440        .Controls(strSortLbl).ForeColor = CLR_DKBLU
7450      Else
7460        strCalled = Left(strProc, (Len(strProc) - Len(strDblClick)))  ' ** For example: taxcode_lbl_DblClick
7470        lngSortLbl_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7480        lngSortLbl_Left = ((.Controls(strCalled & "_lbl").Left + .Controls(strCalled & "_lbl").Width) - lngSortLbl_Width)
7490        lngSortLine_Top = (.Controls(strCalled & "_lbl").Top - lngTpp)
7500        lngSortLine_Left = .Controls(strCalled & "_lbl").Left
7510        lngSortLine_Width = (.Controls(strCalled & "_lbl").Width + lngTpp)
7520        .Controls(strSortLbl).Top = lngSortLbl_Top
7530        .Controls(strSortLbl).Left = lngSortLbl_Left
7540        .Controls(strSortLine1).Top = lngSortLine_Top
7550        .Controls(strSortLine2).Top = (lngSortLine_Top + lngTpp)
7560        .Controls(strSortLine1).Left = lngSortLine_Left
7570        .Controls(strSortLine2).Left = lngSortLine_Left
7580        .Controls(strSortLine1).Width = lngSortLine_Width
7590        .Controls(strSortLine2).Width = lngSortLine_Width
7600        If strCalled = "Schedule_ID" Then
7610          strSortAsc = strSortOrig
7620          If strSortNow = strSortAsc Then
7630            strSortNow = "[Schedule_ID] DESC"
7640            .Controls(strSortLbl).Caption = strArwDn
7650            .Controls(strSortLbl).ForeColor = CLR_DKRED
7660          Else
7670            strSortNow = strSortAsc
7680            .Controls(strSortLbl).Caption = strArwUp
7690            .Controls(strSortLbl).ForeColor = CLR_DKBLU
7700          End If
7710          lngSortLbl_Left = (lngSortLbl_Left + (7& * lngTpp))
7720          .Controls(strSortLbl).Left = lngSortLbl_Left
7730        Else
7740          .Controls(strSortLbl).Caption = strArwUp
7750          .Controls(strSortLbl).ForeColor = CLR_DKBLU
7760          Select Case strCalled
              Case "Schedule_Name"
7770            strSortAsc = "[Schedule_Name]"
7780          Case "Schedule_Base"
7790            strSortAsc = "[Schedule_Base]" & strStdAsc
7800          Case "Schedule_Minimum"
7810            strSortAsc = "[Schedule_Minimum]" & strStdAsc
7820          End Select
7830          If strSortNow = strSortAsc Then
7840            intPos01 = InStr(strSortAsc, ",")
7850            If intPos01 = 0 Then
7860              strSortNow = strSortAsc & " DESC"
7870            Else
7880              strTmp01 = Left(strSortAsc, (intPos01 - 1))
7890              strTmp02 = Mid(strSortAsc, intPos01)
7900              strTmp01 = strTmp01 & " DESC"
7910              intPos01 = InStr((intPos01 + 1), strTmp02, ",")
7920              If intPos01 = 0 Then
7930                strTmp01 = strTmp01 & strTmp02 & " DESC"
7940              Else
                    ' ** There aren't any.
7950              End If
7960              strSortNow = strTmp01
7970            End If
7980            .Controls(strSortLbl).Caption = strArwDn
7990            .Controls(strSortLbl).ForeColor = CLR_DKRED
8000          Else
8010            strSortNow = strSortAsc
8020          End If
8030        End If
8040      End If
8050      .Controls(strSortLbl).Visible = True
8060      .Controls(strSortLine1).Visible = True
8070      .Controls(strSortLine2).Visible = True
8080      .OrderBy = strSortNow
8090      .OrderByOn = True
8100      .SortNow_OrderBy = strSortNow
8110    End With

EXITP:
8120    Set ctl = Nothing
8130    Exit Sub

ERRH:
8140    Select Case ERR.Number
        Case Else
8150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
8160    End Select
8170    Resume EXITP

End Sub
