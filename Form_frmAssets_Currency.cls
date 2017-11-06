VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssets_Currency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAssets_Currency"

'VGC 11/12/2015: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   No Convert:       A {opgChangeCurr_optNoConvert}
' **   Cancel:           C {cmdCancel}
' **   Convert:          N {opgChangeCurr_optConvert}
' **   OK:               O {cmdOK}

Private lngCurrID_Old As Long, lngCurrID_New As Long, strCurrCode_Old As String, strCurrCode_New As String
Private strCurrName_Old As String, strCurrName_New As String
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
140         If strTmp01 <> vbNullString Then
150           If CharCnt(strTmp01, "~") = 6 Then
                ' ** THIS_NAME  lngCurrID  lngCurrID_New  strCurrCode  strCurrCode_New  strCurrName  strCurrName_New

160             intPos01 = InStr(strTmp01, "~")
170             strCallingForm = Left(strTmp01, (intPos01 - 1))
180             strTmp01 = Mid(strTmp01, (intPos01 + 1))
190             intPos01 = InStr(strTmp01, "~")
200             lngCurrID_Old = CLng(Left(strTmp01, (intPos01 - 1)))
210             strTmp01 = Mid(strTmp01, (intPos01 + 1))
220             intPos01 = InStr(strTmp01, "~")
230             lngCurrID_New = CLng(Left(strTmp01, (intPos01 - 1)))
240             strTmp01 = Mid(strTmp01, (intPos01 + 1))
250             intPos01 = InStr(strTmp01, "~")
260             strCurrCode_Old = Left(strTmp01, (intPos01 - 1))
270             strTmp01 = Mid(strTmp01, (intPos01 + 1))
280             intPos01 = InStr(strTmp01, "~")
290             strCurrCode_New = Left(strTmp01, (intPos01 - 1))
300             strTmp01 = Mid(strTmp01, (intPos01 + 1))
310             intPos01 = InStr(strTmp01, "~")
320             strCurrName_Old = Left(strTmp01, (intPos01 - 1))
330             strCurrName_New = Mid(strTmp01, (intPos01 + 1))

340             .MsgSubhead_lbl2b.Caption = strCurrCode_Old
350             .MsgSubhead_lbl2c.Caption = strCurrName_Old
360             .MsgSubhead_lbl3b.Caption = strCurrCode_New
370             .MsgSubhead_lbl3c.Caption = strCurrName_New

380             .PictureData = .Form_Assets_Currency_img

390             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
400             DoEvents

410             opgChangeCurr_AfterUpdate  ' ** Procedure: Below.

                ' ** gblnMessage: True = OK; False = Cancel.
                ' ** gblnForeignCurrencies: True = Convert; False = As-Is.

420           Else
430             Cancel = -1
440           End If
450         Else
460           Cancel = -1
470         End If
480       Else
490         Cancel = -1
500       End If
510     End With

520     DoCmd.Hourglass False

530     If Cancel = -1 Then
540       Beep
550       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
560     End If

EXITP:
570     Exit Sub

ERRH:
580     DoCmd.Hourglass False
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

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

1110    With Me

1120      DoCmd.Hourglass True
1130      DoEvents

1140      gblnMessage = True  ' ** Just to be sure.

1150      DoCmd.Close acForm, THIS_NAME

1160    End With

EXITP:
1170    Exit Sub

ERRH:
1180    DoCmd.Hourglass False
1190    gblnMessage = False
1200    Select Case ERR.Number
        Case Else
1210      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1220    End Select
1230    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1310    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1320    intShiftDown = (Shift And acShiftMask) > 0
1330    intAltDown = (Shift And acAltMask) > 0
1340    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1360      Select Case intRetVal
          Case vbKeyTab
1370        With Me
1380          intRetVal = 0
1390          .cmdCancel.SetFocus
1400        End With
1410      End Select
1420    End If

        ' ** Shift keys.
1430    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1440      Select Case intRetVal
          Case vbKeyTab
1450        With Me
1460          intRetVal = 0
1470          .opgChangeCurr.SetFocus
1480        End With
1490      End Select
1500    End If

EXITP:
1510    KeyCode = intRetVal
1520    Exit Sub

ERRH:
1530    intRetVal = 0
1540    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1550    Case Else
1560      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1570    End Select
1580    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

1610    gblnMessage = False

1620    DoCmd.Close acForm, THIS_NAME

EXITP:
1630    Exit Sub

ERRH:
1640    Select Case ERR.Number
        Case Else
1650      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1660    End Select
1670    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1720    intShiftDown = (Shift And acShiftMask) > 0
1730    intAltDown = (Shift And acAltMask) > 0
1740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1760      Select Case intRetVal
          Case vbKeyTab
1770        With Me
1780          intRetVal = 0
1790          .opgChangeCurr.SetFocus
1800        End With
1810      End Select
1820    End If

        ' ** Shift keys.
1830    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1840      Select Case intRetVal
          Case vbKeyTab
1850        With Me
1860          intRetVal = 0
1870          .cmdOK.SetFocus
1880        End With
1890      End Select
1900    End If

EXITP:
1910    KeyCode = intRetVal
1920    Exit Sub

ERRH:
1930    intRetVal = 0
1940    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1950    Case Else
1960      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1970    End Select
1980    Resume EXITP

End Sub

Private Sub opgChangeCurr_AfterUpdate()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChangeCurr_AfterUpdate"

2010    With Me
2020      Select Case .opgChangeCurr
          Case .opgChangeCurr_optNoConvert.OptionValue
2030        .opgChangeCurr_optNoConvert_lbl.FontBold = True
2040        .opgChangeCurr_optConvert_lbl.FontBold = False
2050        gblnForeignCurrencies = False
2060      Case .opgChangeCurr_optConvert.OptionValue
2070        .opgChangeCurr_optNoConvert_lbl.FontBold = False
2080        .opgChangeCurr_optConvert_lbl.FontBold = True
2090        gblnForeignCurrencies = True
2100      End Select
2110    End With

EXITP:
2120    Exit Sub

ERRH:
2130    Select Case ERR.Number
        Case Else
2140      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2150    End Select
2160    Resume EXITP

End Sub

Private Sub opgChangeCurr_optNoConvert_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChangeCurr_optNoConvert_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2210    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2220    intShiftDown = (Shift And acShiftMask) > 0
2230    intAltDown = (Shift And acAltMask) > 0
2240    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2250    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2260      Select Case intRetVal
          Case vbKeyTab
2270        With Me
2280          intRetVal = 0
2290          .cmdOK.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab
2350        With Me
2360          intRetVal = 0
2370          .cmdCancel.SetFocus
2380        End With
2390      End Select
2400    End If

EXITP:
2410    KeyCode = intRetVal
2420    Exit Sub

ERRH:
2430    intRetVal = 0
2440    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2450    Case Else
2460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2470    End Select
2480    Resume EXITP

End Sub

Private Sub opgChangeCurr_optConvert_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgChangeCurr_optConvert_KeyDown"

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
