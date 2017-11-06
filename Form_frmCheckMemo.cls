VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCheckMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCheckMemo"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Add:              A {cmdAdd}
' **   Close:            C {cmdClose}
' **   Delete:           D {cmdDelete}
' **   Memo:             M {ChkMemo_Memo on frmCheckMemo_Sub}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl keys responsive from this form:
' **   Save:             S {cmdSave on frmCheckMemo_Sub}

Private blnMemoFocus As Boolean
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           blnMemoFocus = False

160         Else
170           Cancel = -1
180         End If
190       Else
200         Cancel = -1
210       End If
220     End With

230     If Cancel = -1 Then
240       Beep
250       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
260     End If

EXITP:
270     Exit Sub

ERRH:
280     DoCmd.Hourglass False
290     Select Case ERR.Number
        Case Else
300       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
310     End Select
320     Resume EXITP

End Sub

Private Sub Form_Load()

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

410     With Me
420       DoCmd.Hourglass False
430     End With

EXITP:
440     Exit Sub

ERRH:
450     DoCmd.Hourglass False
460     Select Case ERR.Number
        Case Else
470       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
480     End Select
490     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality.
        ' **   Memo:             M {ChkMemo_Memo on frmCheckMemo_Sub}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl to other forms and functionality.
        ' **   Save:             S {cmdSave}

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdClose_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Alt keys.
610     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
620       Select Case intRetVal
          Case vbKeyM
630         With Me
640           intRetVal = 0
650           lngRecsCur = .frmCheckMemo_Sub.Form.RecCnt  ' ** Form Function: frmCheckMemo_Sub.
660           If lngRecsCur > 0& Then
670             .frmCheckMemo_Sub.SetFocus
680             .frmCheckMemo_Sub.Form.ChkMemo_Memo.SetFocus
690           End If
700         End With
710       Case vbKeyX
720         intRetVal = 0
730         cmdClose_Click  ' ** Procedure: Below.
740       End Select
750     End If

        ' ** Ctrl keys.
760     If intCtrlDown And (Not intAltDown) And (Not intShiftDown) Then
770       Select Case intRetVal
          Case vbKeyS
780         With Me
790           intRetVal = 0
800           lngRecsCur = .frmCheckMemo_Sub.Form.RecCnt  ' ** Form Function: frmCheckMemo_Sub.
810           If lngRecsCur > 0& Then
820             .frmCheckMemo_Sub.SetFocus
830             .frmCheckMemo_Sub.Form.cmdSave_Click  ' ** Form Procedure: frmCheckMemo_Sub.
840           Else
850             Beep
860           End If
870         End With
880       End Select
890     End If

EXITP:
900     KeyCode = intRetVal
910     Exit Sub

ERRH:
920     intRetVal = 0
930     Select Case ERR.Number
        Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub

Private Sub Form_Timer()

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1010    With Me
1020      .TimerInterval = 0&
1030      If gblnSetFocus = True Then
1040        gblnSetFocus = False
1050        strFormName = vbNullString
1060  On Error Resume Next
1070        Do Until strFormName = THIS_NAME
1080          DoCmd.SelectObject acForm, THIS_NAME, False
1090          strFormName = Screen.ActiveForm.Name
1100        Loop
1110  On Error GoTo ERRH
1120        blnMemoFocus = True
1130        .TimerInterval = 500&
1140      ElseIf blnMemoFocus = True Then
1150        blnMemoFocus = False
1160        lngRecsCur = .frmCheckMemo_Sub.Form.RecCnt  ' ** Form Function: frmCheckMemo_Sub.
1170        If lngRecsCur > 0& Then
1180          .frmCheckMemo_Sub.SetFocus
1190          .frmCheckMemo_Sub.Form.ChkMemo_Memo.SetFocus
1200        Else
1210          .cmdAdd.SetFocus
1220        End If
1230      End If
1240    End With

EXITP:
1250    Exit Sub

ERRH:
1260    Select Case ERR.Number
        Case Else
1270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1280    End Select
1290    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1310    Select Case strCallingForm
        Case "frmMenu_Post"
1320      gblnSetFocus = True
1330      gstrReturningForm = THIS_NAME
1340      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1350    End Select

EXITP:
1360    Exit Sub

ERRH:
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub

Public Sub cmdClose_Click()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1510    DoCmd.Close acForm, THIS_NAME

EXITP:
1520    Exit Sub

ERRH:
1530    Select Case ERR.Number
        Case Else
1540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1550    End Select
1560    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1620    intShiftDown = (Shift And acShiftMask) > 0
1630    intAltDown = (Shift And acAltMask) > 0
1640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1660      Select Case intRetVal
          Case vbKeyTab
1670        With Me
1680          intRetVal = 0
1690          lngRecsCur = .frmCheckMemo_Sub.Form.RecCnt  ' ** Form Function: frmCheckMemo_Sub.
1700          If lngRecsCur > 0& Then
1710            .frmCheckMemo_Sub.SetFocus
1720            .frmCheckMemo_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmCheckMemo_Sub.
1730            blnMemoFocus = True
1740            .TimerInterval = 100&
1750          Else
1760            .cmdAdd.SetFocus
1770          End If
1780        End With
1790      End Select
1800    End If

        ' ** Shift keys.
1810    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1820      Select Case intRetVal
          Case vbKeyTab
1830        With Me
1840          intRetVal = 0
1850          .cmdDelete.SetFocus
1860        End With
1870      End Select
1880    End If

EXITP:
1890    KeyCode = intRetVal
1900    Exit Sub

ERRH:
1910    intRetVal = 0
1920    Select Case ERR.Number
        Case Else
1930      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1940    End Select
1950    Resume EXITP

End Sub

Private Sub cmdAdd_Click()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_Click"

        Dim strDocName As String

2010    strDocName = "frmCheckMemo_Add"
2020    DoCmd.OpenForm strDocName, , , , , acDialog, THIS_NAME

EXITP:
2030    Exit Sub

ERRH:
2040    Select Case ERR.Number
        Case Else
2050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2060    End Select
2070    Resume EXITP

End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdAdd_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2120    intShiftDown = (Shift And acShiftMask) > 0
2130    intAltDown = (Shift And acAltMask) > 0
2140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2160      Select Case intRetVal
          Case vbKeyTab
2170        With Me
2180          intRetVal = 0
2190          .cmdDelete.SetFocus
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab
2250        With Me
2260          intRetVal = 0
2270          lngRecsCur = .frmCheckMemo_Sub.Form.RecCnt  ' ** Form Function: frmCheckMemo_Sub.
2280          If lngRecsCur > 0& Then
2290            .frmCheckMemo_Sub.SetFocus
2300            .frmCheckMemo_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmCheckMemo_Sub.
2310            blnMemoFocus = True
2320            .TimerInterval = 100&
2330          Else
2340            .cmdClose.SetFocus
2350          End If
2360        End With
2370      End Select
2380    End If

EXITP:
2390    KeyCode = intRetVal
2400    Exit Sub

ERRH:
2410    intRetVal = 0
2420    Select Case ERR.Number
        Case Else
2430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2440    End Select
2450    Resume EXITP

End Sub

Private Sub cmdDelete_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef
        Dim lngChkMemoID As Long
        Dim msgResponse As VbMsgBoxResult
        Dim varTmp00 As Variant

2510    With Me
2520      If IsNull(.frmCheckMemo_Sub.Form.ChkMemo_ID) = False Then
2530        lngChkMemoID = .frmCheckMemo_Sub.Form.ChkMemo_ID
2540        varTmp00 = .frmCheckMemo_Sub.Form.ChkMemo_Memo
2550        If IsNull(varTmp00) = True Then varTmp00 = "{empty}"  ' ** Though it shouldn't ever be!
2560        Beep
2570        msgResponse = MsgBox("Are you sure you want to delete the memo:" & vbCrLf & vbCrLf & _
              "  " & varTmp00, vbQuestion + vbYesNo, "Delete Check Memo")
2580        If msgResponse = vbYes Then
2590          DoCmd.Hourglass True
2600          DoEvents
2610          Set dbs = CurrentDb
2620          With dbs
                ' ** Delete tblCheckMemo, by specified [chkmemid].
2630            Set qdf = .QueryDefs("qryCheckMemo_03")
2640            With qdf.Parameters
2650              ![chkmemid] = lngChkMemoID
2660            End With
2670            qdf.Execute
2680            .Close
2690          End With
2700          .frmCheckMemo_Sub.Form.Requery
2710          DoEvents
2720          DoCmd.Hourglass False
2730        End If
2740      Else
2750        Beep
2760      End If
2770    End With

EXITP:
2780    Set qdf = Nothing
2790    Set dbs = Nothing
2800    Exit Sub

ERRH:
2810    Select Case ERR.Number
        Case Else
2820      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2830    End Select
2840    Resume EXITP

End Sub

Private Sub cmdDelete_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdDelete_KeyDown"

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
2990          .cmdClose.SetFocus
3000        End With
3010      End Select
3020    End If

        ' ** Shift keys.
3030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
3040      Select Case intRetVal
          Case vbKeyTab
3050        With Me
3060          intRetVal = 0
3070          .cmdAdd.SetFocus
3080        End With
3090      End Select
3100    End If

EXITP:
3110    KeyCode = intRetVal
3120    Exit Sub

ERRH:
3130    intRetVal = 0
3140    Select Case ERR.Number
        Case Else
3150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3160    End Select
3170    Resume EXITP

End Sub
