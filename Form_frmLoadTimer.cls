VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLoadTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmLoadTimer"

'VGC 10/27/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Stop:             S {cmdStop}
' **   Save:             V {chkSave}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close This:       T {Form_KeyDown}

Private blnTimer As Boolean, blnSetFocus As Boolean
Private dblSecondIncr As Double
Private lngFrmID1 As Long, lngFrmID2 As Long
Private strCallingForm As String, strTargetForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer, intCnt As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
140         If Trim(strTmp01) <> vbNullString Then
150           intCnt = CharCnt(strTmp01, "~")
160           If intCnt = 2 Then

170             .Modal = True: blnSetFocus = True  ' ** Don't use gblnSetFocus! When it's turned on for the form, this will turn it off!

180             intPos01 = InStr(strTmp01, "~")
190             strCallingForm = Left(strTmp01, (intPos01 - 1))
200             strTmp01 = Mid(strTmp01, (intPos01 + 1))
210             intPos01 = InStr(strTmp01, "~")
220             strTargetForm = Left(strTmp01, (intPos01 - 1))  ' ** Target may have 2 forms listed!
230             blnTimer = CBool(Mid(strTmp01, (intPos01 + 1)))

240             If blnTimer = True Then
250               .Msg_Start = time
260             End If
270             DoEvents

280             .PictureData = .Form_FormLoad_img
290             DoEvents

300             If .frm_top <> 0 And .frm_left <> 0 And .frm_width <> 0 And .frm_height <> 0 Then
310               DoCmd.SelectObject acForm, THIS_NAME, False
320               DoEvents
330               DoCmd.MoveSize .frm_left, .frm_top, .frm_width, .frm_height
340             End If
350             DoEvents

360             dblSecondIncr = (1000 / 60)  ' ** 16.6666666666667.

370             chkSave_AfterUpdate  ' ** Procedure: below
380             DoEvents

390           Else
400             Cancel = -1
410           End If
420         Else
430           Cancel = -1
440         End If
450       Else
460         Cancel = -1
470       End If
480     End With

490     If Cancel = -1 Then
500       Beep
510       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
520     End If

EXITP:
530     Exit Sub

ERRH:
540     Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

610     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
620     intShiftDown = (Shift And acShiftMask) > 0
630     intAltDown = (Shift And acAltMask) > 0
640     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave}

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close This:       T {Form_KeyDown}

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdClose_Click  ' ** Procedure: Below.
690       End Select
700     End If

        ' ** Alt keys.
710     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
720       Select Case intRetVal
          Case vbKeyX
730         intRetVal = 0
740         cmdClose_Click  ' ** Procedure: Below.
750       End Select
760     End If

        ' ** Ctrl keys.
770     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
780       Select Case intRetVal
          Case vbKeyS
790         intRetVal = 0
800         cmdSave_Click  ' ** Procedure: Below.
810       End Select
820     End If

        ' ** Ctrl-Shift keys.
830     If intCtrlDown And (Not intAltDown) And intShiftDown Then
840       Select Case intRetVal
          Case vbKeyF
850         With Me
860           intRetVal = 0
870           .FocusHolder.SetFocus
880         End With
890       Case vbKeyT
900         intRetVal = 0
910         cmdClose_Click  ' ** Procedure: Below.
920       End Select
930     End If

EXITP:
940     KeyCode = intRetVal
950     Exit Sub

ERRH:
960     intRetVal = 0
970     Select Case ERR.Number
        Case Else
980       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
990     End Select
1000    Resume EXITP

End Sub

Private Sub Form_Timer()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

        Static lngTmp01 As Long

1110    With Me
1120      .TimerInterval = 0&
1130      If blnSetFocus = True Then
1140        blnSetFocus = False
1150        strFormName = vbNullString
1160  On Error Resume Next
1170        Do Until strFormName = THIS_NAME
1180          lngTmp01 = lngTmp01 + 1&
1190          DoCmd.SelectObject acForm, THIS_NAME, False
1200          strFormName = Screen.ActiveForm.Name
1210          If lngTmp01 > 15& Then
1220            .Modal = False
1230            Exit Do
1240          End If
1250        Loop
1260  On Error GoTo ERRH
1270        DoEvents
1280      End If
1290      If blnTimer = True Then
1300        If IsNull(.Msg_Start) = False Then
1310          .Msg_Finish = time
1320          .Msg_Elapsed = Format((.Msg_Finish - .Msg_Start), "hh:nn:ss")
1330          DoEvents
1340          .Repaint
1350          .TimerInterval = dblSecondIncr
1360        End If
1370      End If
1380    End With

EXITP:
1390    Exit Sub

ERRH:
1400    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
1410    Case Else
1420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1430    End Select
1440    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, rst As DAO.Recordset
        Dim datStartTime As Date, datEndTime As Date, datElapsedTime As Date
        Dim lngThisDbsID As Long
        Dim lngTop As Long, lngLeft As Long, lngWidth As Long, lngHeight As Long
        Dim intPos01 As Integer
        Dim varTmp00 As Variant

1510    With Me

1520      .TimerInterval = 0&

1530      Select Case .chkSave
          Case True

1540        datStartTime = .Msg_Start
1550        datEndTime = .Msg_Finish
1560        datElapsedTime = CDate(.Msg_Elapsed)

1570        intPos01 = InStr(strTargetForm, "^")
1580        If intPos01 > 0 Then
              ' ** frmJournal^frmJournal_Columns.
1590          If IsLoaded(Left(strTargetForm, (intPos01 - 1)), acForm) = True Then  ' ** Module Function: modFileUtilities.
1600            strTargetForm = Left(strTargetForm, (intPos01 - 1))
1610          Else
                ' ** Assuming it must be the other.
1620            strTargetForm = Mid(strTargetForm, (intPos01 + 1))
1630          End If
1640        End If

            ' ** Variables are fed empty, then populated ByRef.
1650        GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Function: modWindowFunctions.

1660        .frm_top = lngTop
1670        .frm_left = lngLeft
1680        .frm_width = lngWidth
1690        .frm_height = lngHeight
1700        cmdSave_Click  ' ** Procedure: Below.

1710        lngThisDbsID = CurrentAppID  ' ** Module Function: modFileUtilities.
1720        varTmp00 = DLookup("[frm_id]", "tblForm", "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_name] = '" & strCallingForm & "'")
1730        If IsNull(varTmp00) = False Then
1740          lngFrmID1 = varTmp00
1750          varTmp00 = DLookup("[frm_id]", "tblForm", "[dbs_id] = " & CStr(lngThisDbsID) & " And [frm_name] = '" & strTargetForm & "'")
1760          If IsNull(varTmp00) = False Then
1770            lngFrmID2 = varTmp00

1780            Set dbs = CurrentDb
1790            Set rst = dbs.OpenRecordset("tblXAdmin_Load", dbOpenDynaset, dbAppendOnly)
1800            rst.AddNew
                ' ** rst![xadload_id] : AutoNumber.
1810            rst![dbs_id1] = lngThisDbsID
1820            rst![frm_id1] = lngFrmID1
1830            rst![frm_name1] = strCallingForm
1840            rst![dbs_id2] = lngThisDbsID
1850            rst![frm_id2] = lngFrmID2
1860            rst![frm_name2] = strTargetForm
1870            rst![xadload_start] = datStartTime
1880            rst![xadload_end] = datEndTime
1890            rst![xadload_time] = datElapsedTime
1900            rst![xadload_access_version] = AccessVer(Application.Version)  ' ** Module Function: modVersionConvertFuncs3.
1910            rst![xadload_datemodified] = Now()
1920            rst.Update
1930            rst.Close
1940            dbs.Close

1950          End If
1960        End If

1970      Case False
1980        .frm_top = 0&
1990        .frm_left = 0&
2000        .frm_width = 0&
2010        .frm_height = 0&
2020        cmdSave_Click  ' ** Procedure: Below.
2030      End Select  ' ** chkSave.

2040      If IsLoaded(strTargetForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
2050        DoCmd.SelectObject acForm, strTargetForm, False
2060      End If

2070    End With

EXITP:
2080    Set rst = Nothing
2090    Set dbs = Nothing
2100    Exit Sub

ERRH:
2110    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
2120    Case Else
2130      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2140    End Select
2150    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2210    DoCmd.Close acForm, THIS_NAME

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Public Sub cmdSave_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdSave_Click"

2310    DoCmd.RunCommand acCmdSaveRecord
2320    DoEvents

EXITP:
2330    Exit Sub

ERRH:
2340    Select Case ERR.Number
        Case Else
2350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2360    End Select
2370    Resume EXITP

End Sub

Public Sub cmdStop_Click()

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStop_Click"

2410    With Me
2420      .TimerInterval = 0&
2430      blnTimer = False
2440      .FocusHolder.SetFocus
2450      DoEvents
2460      .cmdStop.Enabled = False
2470      .Modal = False
2480      Beep
2490    End With

EXITP:
2500    Exit Sub

ERRH:
2510    Select Case ERR.Number
        Case Else
2520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2530    End Select
2540    Resume EXITP

End Sub

Private Sub cmdStop_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdStop_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyTab
2670        With Me
2680          intRetVal = 0
2690          .FocusHolder.SetFocus
2700        End With
2710      End Select
2720    End If

        ' ** Shift keys.
2730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2740      Select Case intRetVal
          Case vbKeyTab
2750        With Me
2760          intRetVal = 0
2770          .FocusHolder.SetFocus
2780        End With
2790      End Select
2800    End If

EXITP:
2810    KeyCode = intRetVal
2820    Exit Sub

ERRH:
2830    intRetVal = 0
2840    Select Case ERR.Number
        Case Else
2850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2860    End Select
2870    Resume EXITP

End Sub

Private Sub chkSave_AfterUpdate()

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "chkSave_AfterUpdate"

2910    With Me
2920      Select Case .chkSave
          Case True
2930        .chkSave_lbl.FontBold = True
2940      Case False
2950        .chkSave_lbl.FontBold = False
2960      End Select
2970      .FocusHolder.SetFocus
2980    End With

EXITP:
2990    Exit Sub

ERRH:
3000    Select Case ERR.Number
        Case Else
3010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3020    End Select
3030    Resume EXITP

End Sub

Private Sub chkSave_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "chkSave_KeyDown"

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
3190          .FocusHolder.SetFocus
3200        End With
3210      End Select
3220    End If

        ' ** Shift keys.
3230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3240      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
3250        With Me
3260          intRetVal = 0
3270          .FocusHolder.SetFocus
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
