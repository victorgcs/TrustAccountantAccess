VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStatementBalance_Date"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStatementBalance_Date"

'VGC 11/27/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   Reset:            E {cmdReset}
' **   OK:               O {cmdOK}

Private CLR_DISABLED_FG As Long
Private CLR_DISABLED_BG As Long

Private blnReset_Focus As Boolean, blnReset_MouseDown As Boolean
Private datLastStmtDate As Date, datPrevStmtDate As Date, datLastBalDate As Date
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim blnContinue As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           .PictureData = .Form_StatementBalance_Date_img

160           CLR_DISABLED_FG = CLR_DKGRY
170           CLR_DISABLED_BG = CLR_LTTEAL

180           blnContinue = True
190           blnReset_Focus = False: blnReset_MouseDown = False

200           Set dbs = CurrentDb
210           With dbs

                ' ** Statement Date, just Statement_Date.
220             Set qdf = .QueryDefs("qryStatementDate_01")
230             Set rst = qdf.OpenRecordset
240             With rst
250               .MoveFirst
260               datLastStmtDate = ![Statement_Date]
270               .Close
280             End With

                ' ** qryStatementDate_02 (Balance, with IsLastDay), grouped by Balance_Date, descending, with cnt.
290             Set qdf = .QueryDefs("qryStatementDate_03")
300             Set rst = qdf.OpenRecordset
310             With rst
320               If .BOF = True And .EOF = True Then
330                 blnContinue = False
340               Else
350                 .MoveLast
360                 If .RecordCount = 1 Then
370                   blnContinue = False
380                 Else
390                   .MoveFirst
400                   datLastBalDate = ![balance_date]  ' ** Would normally be same as datLastStmtDate.
410                   .MoveNext
420                   datPrevStmtDate = ![balance_date]
430                 End If
440               End If
450               .Close
460             End With

470             .Close
480           End With

490           Select Case blnContinue
              Case True

500             .Statement_Date_Last.ForeColor = CLR_DISABLED_FG
510             .Statement_Date_Last.BackColor = CLR_DISABLED_BG
520             .Statement_Date_Prev.ForeColor = CLR_DISABLED_FG
530             .Statement_Date_Prev.BackColor = CLR_DISABLED_BG

540             .Statement_Date_Last = datLastStmtDate
550             .Statement_Date_Prev = datPrevStmtDate
560             .cmdOK.Enabled = False
570             .FocusHolder.SetFocus

580             Beep

590           Case False
600             Cancel = -1
610             Beep
620             MsgBox "There are insufficient balance records to proceed.", _
                  vbInformation + vbOKOnly, "Procedure Unavailable"
630           End Select

640         Else
650           Cancel = -1
660         End If
670       Else
680         Cancel = -1
690       End If
700     End With

710     If Cancel = -1 Then
720       If blnContinue = True Then
730         Beep
740         MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
750       End If
760     End If

EXITP:
770     Set rst = Nothing
780     Set qdf = Nothing
790     Set dbs = Nothing
800     Exit Sub

ERRH:
810     Select Case ERR.Number
        Case Else
820       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
830     End Select
840     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

900   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

910     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
920     intShiftDown = (Shift And acShiftMask) > 0
930     intAltDown = (Shift And acAltMask) > 0
940     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
950     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
960       Select Case intRetVal
          Case vbKeyEscape
970         intRetVal = 0
980         cmdCancel_Click  ' ** Procedure: Below.
990       End Select
1000    End If

        ' ** Alt keys.
1010    If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
1020      Select Case intRetVal
          Case vbKeyE
1030        With Me
1040          intRetVal = 0
1050          .cmdReset.SetFocus
1060          cmdReset_Click  ' ** Procedure: Below.
1070        End With
1080      End Select
1090    End If

EXITP:
1100    KeyCode = intRetVal
1110    Exit Sub

ERRH:
1120    intRetVal = 0
1130    Select Case ERR.Number
        Case Else
1140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1150    End Select
1160    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1210    gblnSetFocus = True
1220    DoCmd.SelectObject acForm, strCallingForm, False
1230    Forms(strCallingForm).TimerInterval = 100&

EXITP:
1240    Exit Sub

ERRH:
1250    Select Case ERR.Number
        Case Else
1260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1270    End Select
1280    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1310    With Me
1320      If .cmdReset_raised_focus_dots_img.Visible = True Or .cmdReset_raised_focus_img.Visible = True Then
1330        Select Case blnReset_Focus
            Case True
1340          .cmdReset_raised_semifocus_dots_img.Visible = True
1350          .cmdReset_raised_img.Visible = False
1360        Case False
1370          .cmdReset_raised_img.Visible = True
1380          .cmdReset_raised_semifocus_dots_img.Visible = False
1390        End Select
1400        .cmdReset_raised_focus_img.Visible = False
1410        .cmdReset_raised_focus_dots_img.Visible = False
1420        .cmdReset_sunken_focus_dots_img.Visible = False
1430        .cmdReset_raised_img_dis.Visible = False
1440      End If
1450    End With

EXITP:
1460    Exit Sub

ERRH:
1470    Select Case ERR.Number
        Case Else
1480      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1490    End Select
1500    Resume EXITP

End Sub

Private Sub cmdOK_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, rst As DAO.Recordset

1610    With Me

1620      datLastStmtDate = .Statement_Date_Last

1630      Set dbs = CurrentDb
1640      With dbs
1650        Set rst = .OpenRecordset("Statement Date", dbOpenDynaset, dbConsistent)
1660        With rst
1670          .MoveFirst
1680          .Edit
1690          ![Statement_Date] = datLastStmtDate
1700          .Update
1710          .Close
1720        End With
1730        .Close
1740      End With

1750      DoCmd.Close acForm, THIS_NAME

1760    End With

EXITP:
1770    Set rst = Nothing
1780    Set dbs = Nothing
1790    Exit Sub

ERRH:
1800    Select Case ERR.Number
        Case Else
1810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1820    End Select
1830    Resume EXITP

End Sub

Private Sub cmdOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_MouseMove"

1910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1920    Exit Sub

ERRH:
1930    Select Case ERR.Number
        Case Else
1940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1950    End Select
1960    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2010    DoCmd.Close acForm, THIS_NAME

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_MouseMove"

2110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub cmdReset_GotFocus()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_GotFocus"

2210    With Me
2220      blnReset_Focus = True
2230      .cmdReset_raised_semifocus_dots_img.Visible = True
2240      .cmdReset_raised_img.Visible = False
2250      .cmdReset_raised_focus_img.Visible = False
2260      .cmdReset_raised_focus_dots_img.Visible = False
2270      .cmdReset_sunken_focus_dots_img.Visible = False
2280      .cmdReset_raised_img_dis.Visible = False
2290    End With

EXITP:
2300    Exit Sub

ERRH:
2310    Select Case ERR.Number
        Case Else
2320      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2330    End Select
2340    Resume EXITP

End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseDown"

2410    With Me
2420      blnReset_MouseDown = True
2430      .cmdReset_sunken_focus_dots_img.Visible = True
2440      .cmdReset_raised_img.Visible = False
2450      .cmdReset_raised_semifocus_dots_img.Visible = False
2460      .cmdReset_raised_focus_img.Visible = False
2470      .cmdReset_raised_focus_dots_img.Visible = False
2480      .cmdReset_raised_img_dis.Visible = False
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

Private Sub cmdReset_Click()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_Click"

2610    With Me
2620      .Statement_Date_Last = .Statement_Date_Prev
2630      .cmdOK.Enabled = True
2640      .FocusHolder.SetFocus
2650      .cmdReset.Enabled = False
2660      .cmdReset_raised_img_dis.Visible = True
2670      .cmdReset_raised_img.Visible = False
2680      .cmdReset_raised_semifocus_dots_img.Visible = False
2690      .cmdReset_raised_focus_img.Visible = False
2700      .cmdReset_raised_focus_dots_img.Visible = False
2710      .cmdReset_sunken_focus_dots_img.Visible = False
2720      blnReset_Focus = False
2730      blnReset_MouseDown = False
2740    End With

EXITP:
2750    Exit Sub

ERRH:
2760    Select Case ERR.Number
        Case Else
2770      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2780    End Select
2790    Resume EXITP

End Sub

Private Sub cmdReset_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2810    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2820    intShiftDown = (Shift And acShiftMask) > 0
2830    intAltDown = (Shift And acAltMask) > 0
2840    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          If .cmdOK.Enabled = True Then
2900            .cmdOK.SetFocus
2910          Else
2920            .cmdCancel.SetFocus
2930          End If
2940        End With
2950      End Select
2960    End If

        ' ** Shift keys.
2970    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2980      Select Case intRetVal
          Case vbKeyTab
2990        With Me
3000          intRetVal = 0
3010          .cmdCancel.SetFocus
3020        End With
3030      End Select
3040    End If

EXITP:
3050    KeyCode = intRetVal
3060    Exit Sub

ERRH:
3070    intRetVal = 0
3080    Select Case ERR.Number
        Case Else
3090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3100    End Select
3110    Resume EXITP

End Sub

Private Sub cmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseMove"

3210    With Me
3220      If blnReset_MouseDown = False Then
3230        Select Case blnReset_Focus
            Case True
3240          .cmdReset_raised_focus_dots_img.Visible = True
3250          .cmdReset_raised_focus_img.Visible = False
3260        Case False
3270          .cmdReset_raised_focus_img.Visible = True
3280          .cmdReset_raised_focus_dots_img.Visible = False
3290        End Select
3300        .cmdReset_raised_img.Visible = False
3310        .cmdReset_raised_semifocus_dots_img.Visible = False
3320        .cmdReset_sunken_focus_dots_img.Visible = False
3330        .cmdReset_raised_img_dis.Visible = False
3340      End If
3350    End With

EXITP:
3360    Exit Sub

ERRH:
3370    Select Case ERR.Number
        Case Else
3380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3390    End Select
3400    Resume EXITP

End Sub

Private Sub cmdReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_MouseUp"

3510    With Me
3520      .cmdReset_raised_focus_dots_img.Visible = True
3530      .cmdReset_raised_img.Visible = False
3540      .cmdReset_raised_semifocus_dots_img.Visible = False
3550      .cmdReset_raised_focus_img.Visible = False
3560      .cmdReset_sunken_focus_dots_img.Visible = False
3570      .cmdReset_raised_img_dis.Visible = False
3580      blnReset_MouseDown = False
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

Private Sub cmdReset_LostFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdReset_LostFocus"

3710    With Me
3720      .cmdReset_raised_img.Visible = True
3730      .cmdReset_raised_semifocus_dots_img.Visible = False
3740      .cmdReset_raised_focus_img.Visible = False
3750      .cmdReset_raised_focus_dots_img.Visible = False
3760      .cmdReset_sunken_focus_dots_img.Visible = False
3770      .cmdReset_raised_img_dis.Visible = False
3780      blnReset_Focus = False
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

Private Sub SubHeader_lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3900  On Error GoTo ERRH

        Const THIS_PROC As String = "SubHeader_lbl2_MouseMove"

3910    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
3920    Exit Sub

ERRH:
3930    Select Case ERR.Number
        Case Else
3940      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3950    End Select
3960    Resume EXITP

End Sub

Private Sub Statement_Date_Last_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4000  On Error GoTo ERRH

        Const THIS_PROC As String = "Statement_Date_Last_MouseMove"

4010    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4020    Exit Sub

ERRH:
4030    Select Case ERR.Number
        Case Else
4040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4050    End Select
4060    Resume EXITP

End Sub

Private Sub Statement_Date_Prev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

4100  On Error GoTo ERRH

        Const THIS_PROC As String = "Statement_Date_Prev_MouseMove"

4110    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
4120    Exit Sub

ERRH:
4130    Select Case ERR.Number
        Case Else
4140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
4150    End Select
4160    Resume EXITP

End Sub
