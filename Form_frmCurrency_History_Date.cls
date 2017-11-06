VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_History_Date"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_History_Date"

'VGC 03/22/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}

Private lngRecsCur As Long
'Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           .PictureData = .Form_Currency_History_Date_img

150           .frmCurrency_History_Date_Sub.SetFocus

160         Else
170           Cancel = -1
180         End If
190       Else
200         Cancel = -1
210       End If
220     End With

230     DoCmd.Hourglass False

240     If Cancel = -1 Then
250       Beep
260       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
270     End If

EXITP:
280     Exit Sub

ERRH:
290     DoCmd.Hourglass False
300     Select Case ERR.Number
        Case Else
310       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
320     End Select
330     Resume EXITP

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
470         intRetVal = 0
480         cmdCancel_Click  ' ** Procedure: Below.
490       End Select
500     End If

EXITP:
510     KeyCode = intRetVal
520     Exit Sub

ERRH:
530     intRetVal = 0
540     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
550     Case Else
560       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
570     End Select
580     Resume EXITP

End Sub

Private Sub cmdOK_Click()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef, rst As DAO.Recordset
        Dim lngCurrHistID As Long, dblRate1 As Double, dblRate2 As Double
        Dim lngX As Long

610     With Me

620       DoCmd.Hourglass True
630       DoEvents

640       lngRecsCur = .frmCurrency_History_Date_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Date_Sub.
650       If lngRecsCur > 0& Then
660         Set dbs = CurrentDb
670         Set rst = .frmCurrency_History_Date_Sub.Form.RecordsetClone
680         With rst
690           .MoveFirst
700           For lngX = 1& To lngRecsCur
710             lngCurrHistID = ![currhist_id]
720             dblRate1 = ![curr_rate1]
730             dblRate2 = ![curr_rate2]
740             If ![R1x] = "X" Then
                  ' ** Update qryCurrency_History_Date_02_01 (tblCurrency_History,
                  ' ** with curr_rate1_new, by specified [histid], [rat1]).
750               Set qdf = dbs.QueryDefs("qryCurrency_History_Date_02_02")
760               With qdf.Parameters
770                 ![histid] = lngCurrHistID
780                 ![rat1] = dblRate1
790               End With
800               qdf.Execute
810               Set qdf = Nothing
820             End If
830             If ![R2x] = "X" Then
                  ' ** Update qryCurrency_History_Date_03_01 (tblCurrency_History,
                  ' ** with curr_rate2_new, by specified [histid], [rat2]).
840               Set qdf = dbs.QueryDefs("qryCurrency_History_Date_03_02")
850               With qdf.Parameters
860                 ![histid] = lngCurrHistID
870                 ![rat2] = dblRate2
880               End With
890               qdf.Execute
900               Set qdf = Nothing
910             End If
920             If lngX < lngRecsCur Then .MoveNext
930           Next
940           .Close
950         End With
960         Set rst = Nothing
970         dbs.Close
980         Set dbs = Nothing
990       End If

1000      gblnMessage = True  ' ** Just making sure.
1010      DoCmd.Close acForm, THIS_NAME

1020      DoCmd.Hourglass False

1030    End With

EXITP:
1040    Set rst = Nothing
1050    Set qdf = Nothing
1060    Set dbs = Nothing
1070    Exit Sub

ERRH:
1080    DoCmd.Hourglass False
1090    Select Case ERR.Number
        Case Else
1100      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1110    End Select
1120    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1220    intShiftDown = (Shift And acShiftMask) > 0
1230    intAltDown = (Shift And acAltMask) > 0
1240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1260      Select Case intRetVal
          Case vbKeyTab
1270        With Me
1280          intRetVal = 0
1290          .cmdCancel.SetFocus
1300        End With
1310      End Select
1320    End If

        ' ** Shift keys.
1330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1340      Select Case intRetVal
          Case vbKeyTab
1350        With Me
1360          intRetVal = 0
1370          .frmCurrency_History_Date_Sub.SetFocus
1380          .frmCurrency_History_Date_Sub.Form.FocusHolder.SetFocus
1390        End With
1400      End Select
1410    End If

EXITP:
1420    KeyCode = intRetVal
1430    Exit Sub

ERRH:
1440    intRetVal = 0
1450    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1460    Case Else
1470      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1480    End Select
1490    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1510    gblnMessage = False
1520    DoCmd.Close acForm, THIS_NAME

EXITP:
1530    Exit Sub

ERRH:
1540    Select Case ERR.Number
        Case Else
1550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1560    End Select
1570    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
1690          .frmCurrency_History_Date_Sub.SetFocus
1700          .frmCurrency_History_Date_Sub.Form.FocusHolder.SetFocus
1710        End With
1720      End Select
1730    End If

        ' ** Shift keys.
1740    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1750      Select Case intRetVal
          Case vbKeyTab
1760        With Me
1770          intRetVal = 0
1780          .cmdOK.SetFocus
1790        End With
1800      End Select
1810    End If

EXITP:
1820    KeyCode = intRetVal
1830    Exit Sub

ERRH:
1840    intRetVal = 0
1850    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1860    Case Else
1870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1880    End Select
1890    Resume EXITP

End Sub
