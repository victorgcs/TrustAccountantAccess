VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCurrency_History_Restore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCurrency_History_Restore"

'VGC 11/12/2015: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   All:              A {opgRestore_optAll}
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}
' **   Specified:        S {opgRestore_optSpecified}

Private strMode As String, datCurrDate As Date, lngCurrID As Long, strCurrCode As String, strCurrName As String
Private lngRecsCur As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         strTmp01 = Trim(.OpenArgs)
            'strTmp01 = "frmCurrency_History~Rate~10/10/2015~47~ERN~Eritrean Nakfa"
140         If strTmp01 <> vbNullString Then
150           If CharCnt(strTmp01, "~") = 5 Then  ' ** Module Function: modStringFuncs.

160             .PictureData = .Form_Currency_History_Restore_img

170             intPos01 = InStr(strTmp01, "~")
180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strTmp01 = Mid(strTmp01, (intPos01 + 1))
200             intPos01 = InStr(strTmp01, "~")
210             strMode = Left(strTmp01, (intPos01 - 1))
220             strTmp01 = Mid(strTmp01, (intPos01 + 1))
230             intPos01 = InStr(strTmp01, "~")
240             datCurrDate = CDate(Left(strTmp01, (intPos01 - 1)))
250             strTmp01 = Mid(strTmp01, (intPos01 + 1))
260             intPos01 = InStr(strTmp01, "~")
270             lngCurrID = CLng(Left(strTmp01, (intPos01 - 1)))
280             strTmp01 = Mid(strTmp01, (intPos01 + 1))
290             intPos01 = InStr(strTmp01, "~")
300             strCurrCode = Left(strTmp01, (intPos01 - 1))
310             strCurrName = Mid(strTmp01, (intPos01 + 1))

320             .opgRestore = .opgRestore_optSpecified.OptionValue
330             Select Case strMode
                Case "Date"
340               .opgRestore_optAll.Enabled = True
350             Case "Rate"
360               .opgRestore_optAll.Enabled = False
370             End Select

380             .MsgSubhead_lbl2b.Caption = strCurrCode & "  " & strCurrName
390             .MsgSubhead_lbl3b.Caption = Format(datCurrDate, "mm/dd/yyyy")

400             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
410             DoEvents

420             opgRestore_AfterUpdate  ' ** Procedure: Below.

430           Else
440             Cancel = -1
450           End If
460         Else
470           Cancel = -1
480         End If
490       Else
500         Cancel = -1
510       End If
520     End With

530     DoCmd.Hourglass False

540     If Cancel = -1 Then
550       Beep
560       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
570     End If

EXITP:
580     Exit Sub

ERRH:
590     DoCmd.Hourglass False
600     Select Case ERR.Number
        Case Else
610       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
620     End Select
630     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

710     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
720     intShiftDown = (Shift And acShiftMask) > 0
730     intAltDown = (Shift And acAltMask) > 0
740     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdCancel_Click  ' ** Procedure: Below.
790       End Select
800     End If

        ' ** Ctrl-Shift keys.
810     If intCtrlDown And (Not intAltDown) And intShiftDown Then
820       Select Case intRetVal
          Case vbKeyF
830         With Me
840           intRetVal = 0
850           .FocusHolder.SetFocus
860         End With
870       End Select
880     End If

EXITP:
890     KeyCode = intRetVal
900     Exit Sub

ERRH:
910     intRetVal = 0
920     Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
930     Case Else
940       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
950     End Select
960     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1010    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

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

        Dim frm As Access.Form

1110    With Me

1120      DoCmd.Hourglass True
1130      DoEvents

1140      gblnMessage = True  ' ** Just to be sure.

1150      Set frm = Forms(strCallingForm)

1160      Select Case strMode
          Case "Date"
1170        Select Case .opgRestore
            Case .opgRestore_optSpecified.OptionValue
1180          frm.curr_id_one = lngCurrID
1190          frm.curr_date_one = datCurrDate
1200          frm.curr_date_all = Null
1210        Case .opgRestore_optAll.OptionValue
1220          frm.curr_date_all = datCurrDate
1230          frm.curr_id_one = Null
1240          frm.curr_date_one = Null
1250        End Select
1260      Case "Rate"
1270        frm.curr_id_one = lngCurrID
1280        frm.curr_date_one = datCurrDate
1290        frm.curr_date_all = Null
1300      End Select
1310      DoEvents

1320      DoCmd.Close acForm, THIS_NAME

1330    End With

EXITP:
1340    Set frm = Nothing
1350    Exit Sub

ERRH:
1360    DoCmd.Hourglass False
1370    gblnMessage = False
1380    Select Case ERR.Number
        Case Else
1390      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1400    End Select
1410    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

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
          Case vbKeyTab
1570        With Me
1580          intRetVal = 0
1590          .cmdCancel.SetFocus
1600        End With
1610      End Select
1620    End If

        ' ** Shift keys.
1630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1640      Select Case intRetVal
          Case vbKeyTab
1650        With Me
1660          intRetVal = 0
1670          .opgRestore.SetFocus
1680        End With
1690      End Select
1700    End If

EXITP:
1710    KeyCode = intRetVal
1720    Exit Sub

ERRH:
1730    intRetVal = 0
1740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1750    Case Else
1760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1770    End Select
1780    Resume EXITP

End Sub

Public Sub cmdCancel_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1810    gblnMessage = False
1820    DoCmd.Close acForm, THIS_NAME

EXITP:
1830    Exit Sub

ERRH:
1840    Select Case ERR.Number
        Case Else
1850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1860    End Select
1870    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

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
1990          .opgRestore.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab
2050        With Me
2060          intRetVal = 0
2070          .cmdOK.SetFocus
2080        End With
2090      End Select
2100    End If

EXITP:
2110    KeyCode = intRetVal
2120    Exit Sub

ERRH:
2130    intRetVal = 0
2140    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2150    Case Else
2160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2170    End Select
2180    Resume EXITP

End Sub

Private Sub opgRestore_AfterUpdate()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgRestore_AfterUpdate"

2210    With Me
2220      Select Case .opgRestore
          Case .opgRestore_optSpecified.OptionValue
2230        .opgRestore_optSpecified_lbl.FontBold = True
2240        .opgRestore_optAll_lbl.FontBold = False
2250        .MsgSubhead_lbl2b.Caption = strCurrCode & "  " & strCurrName
2260        .MsgSubhead_lbl3b.Caption = Format(datCurrDate, "mm/dd/yyyy")
2270      Case .opgRestore_optAll.OptionValue
2280        lngRecsCur = Forms(strCallingForm).frmCurrency_History_Sub.Form.RecCnt  ' ** Form Function: frmCurrency_History_Sub.
2290        .opgRestore_optAll_lbl.FontBold = True
2300        .opgRestore_optSpecified_lbl.FontBold = False
2310        .MsgSubhead_lbl2b.Caption = CStr(lngRecsCur) & " Rates Dated"
2320        .MsgSubhead_lbl3b.Caption = Format(datCurrDate, "mm/dd/yyyy")
2330      End Select
2340      .cmdOK.SetFocus
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

Private Sub opgRestore_optSpecified_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgRestore_optSpecified_KeyDown"

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
2590          .cmdOK.SetFocus
2600        End With
2610      End Select
2620    End If

        ' ** Shift keys.
2630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2640      Select Case intRetVal
          Case vbKeyTab
2650        With Me
2660          intRetVal = 0
2670          .cmdCancel.SetFocus
2680        End With
2690      End Select
2700    End If

EXITP:
2710    KeyCode = intRetVal
2720    Exit Sub

ERRH:
2730    intRetVal = 0
2740    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2750    Case Else
2760      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2770    End Select
2780    Resume EXITP

End Sub

Private Sub opgRestore_optAll_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgRestore_optAll_KeyDown"

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
2890          .cmdOK.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab
2950        With Me
2960          intRetVal = 0
2970          .cmdCancel.SetFocus
2980        End With
2990      End Select
3000    End If

EXITP:
3010    KeyCode = intRetVal
3020    Exit Sub

ERRH:
3030    intRetVal = 0
3040    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3050    Case Else
3060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3070    End Select
3080    Resume EXITP

End Sub
