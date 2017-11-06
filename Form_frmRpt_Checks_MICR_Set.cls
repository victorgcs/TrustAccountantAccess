VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_MICR_Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_MICR_Set"

'VGC 11/23/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   MICRN:            I {opgFont_optMICRN}
' **   MICR:             M {opgFont_optMICR}
' **   MICRN2:           N {opgFont_optMICRN2}
' **   OK:               O {cmdOK}
' **   MICRW:            R {opgFont_optMICRW}
' **   MICRW2:           W {opgFont_optMICRW2}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim ctl As Access.Control
        Dim strFont As String
        Dim intPos01 As Integer

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           strFont = IniFile_Get("General", "MICRFont", RET_ERR, CurrentAppPath & LNK_SEP & gstrFile_INI)  ' ** Module Function: modStartupFuncs, modFileUtilities.
160           If strFont <> vbNullString Then
170             strFont = "opgFont_opt" & strFont
180             For Each ctl In .Controls
190               With ctl
200                 If .Name = strFont Then
210                   Me.opgFont = .OptionValue
220                   Exit For
230                 End If
240               End With
250             Next
260           End If

270           opgFont_AfterUpdate  ' ** Procedure: Below.

280           .Found_MICR_lbl2.Visible = False
290           strFont = vbNullString
300           strFont = FindMICRFont  ' ** Module Function: modFontPicker.
310           If strFont <> vbNullString And strFont <> RET_ERR Then
320             intPos01 = InStr(strFont, ";")
330             If intPos01 = 0 Then
340               .Found_MICR_lbl1.Caption = "The following MICR font was found installed on your computer:"
350             Else
360               .Found_MICR_lbl1.Caption = "The following MICR fonts were found installed on your computer:"
370             End If
380             .Found_MICR_lbl2.Caption = strFont
390             .Found_MICR_lbl2.Visible = True
400           Else
410             .Found_MICR_lbl1.Caption = "No MICR fonts were found installed on your computer."
420           End If

430         Else
440           Cancel = -1
450         End If
460       Else
470         Cancel = -1
480       End If
490     End With

500     If Cancel = -1 Then
510       Beep
520       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
530     End If

EXITP:
540     Set ctl = Nothing
550     Exit Sub

ERRH:
560     Select Case ERR.Number
        Case Else
570       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
580     End Select
590     Resume EXITP

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

        ' ** Plain keys.
650     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
660       Select Case intRetVal
          Case vbKeyEscape
670         intRetVal = 0
680         cmdCancel_Click  ' ** Procedure: Below.
690       End Select
700     End If

EXITP:
710     KeyCode = intRetVal
720     Exit Sub

ERRH:
730     intRetVal = 0
740     Select Case ERR.Number
        Case Else
750       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
760     End Select
770     Resume EXITP

End Sub

Private Sub Form_Close()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Close"

810     If strCallingForm <> vbNullString Then
820       gblnSetFocus = True
830       Forms(strCallingForm).TimerInterval = 250&
840     End If

EXITP:
850     Exit Sub

ERRH:
860     Select Case ERR.Number
        Case Else
870       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
880     End Select
890     Resume EXITP

End Sub

Private Sub cmdOK_Click()

900   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim strFont As String
        Dim blnContinue As Boolean

910     blnContinue = True

920     With Me
930       Select Case Nz(.opgFont, 0)
          Case 0
940         blnContinue = False
950         MsgBox "You must select a font.", vbExclamation + vbOKOnly, "No Font Selected"
960       Case .opgFont_optMICRN.OptionValue
970         strFont = "MICRN"
980       Case .opgFont_optMICRN2.OptionValue
990         strFont = "MICRN2"
1000      Case .opgFont_optMICR.OptionValue
1010        strFont = "MICR"
1020      Case .opgFont_optMICRW2.OptionValue
1030        strFont = "MICRW2"
1040      Case .opgFont_optMICRW.OptionValue
1050        strFont = "MICRW"
1060      End Select
1070    End With

1080    If blnContinue = True Then
1090      If IniFile_Set("General", "MICRFont", strFont, CurrentAppPath & LNK_SEP & gstrFile_INI) = True Then  ' ** Module Function: modStartupFuncs, modFileUtilities.
1100        DoCmd.Close
1110      End If
1120    End If

EXITP:
1130    Exit Sub

ERRH:
1140    Select Case ERR.Number
        Case Else
1150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1160    End Select
1170    Resume EXITP

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
1370          .opgFont.SetFocus
1380        End With
1390      End Select
1400    End If

EXITP:
1410    KeyCode = intRetVal
1420    Exit Sub

ERRH:
1430    intRetVal = 0
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

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
1690          .opgFont.SetFocus
1700        End With
1710      End Select
1720    End If

        ' ** Shift keys.
1730    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1740      Select Case intRetVal
          Case vbKeyTab
1750        With Me
1760          intRetVal = 0
1770          .cmdOK.SetFocus
1780        End With
1790      End Select
1800    End If

EXITP:
1810    KeyCode = intRetVal
1820    Exit Sub

ERRH:
1830    intRetVal = 0
1840    Select Case ERR.Number
        Case Else
1850      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1860    End Select
1870    Resume EXITP

End Sub

Private Sub opgFont_AfterUpdate()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFont_AfterUpdate"

1910    With Me
1920      .opgFont_optMICR_lbl.FontBold = False
1930      .opgFont_optMICRN_lbl.FontBold = False
1940      .opgFont_optMICRN2_lbl.FontBold = False
1950      .opgFont_optMICRW_lbl.FontBold = False
1960      .opgFont_optMICRW2_lbl.FontBold = False
1970      Select Case .opgFont
          Case .opgFont_optMICR.OptionValue
1980        .opgFont_optMICR_lbl.FontBold = True
1990      Case .opgFont_optMICRN.OptionValue
2000        .opgFont_optMICRN_lbl.FontBold = True
2010      Case .opgFont_optMICRN2.OptionValue
2020        .opgFont_optMICRN2_lbl.FontBold = True
2030      Case .opgFont_optMICRW.OptionValue
2040        .opgFont_optMICRW_lbl.FontBold = True
2050      Case .opgFont_optMICRW2.OptionValue
2060        .opgFont_optMICRW2_lbl.FontBold = True
2070      End Select
2080    End With

EXITP:
2090    Exit Sub

ERRH:
2100    Select Case ERR.Number
        Case Else
2110      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2120    End Select
2130    Resume EXITP

End Sub

Private Sub opgFont_optMICR_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFont_optMICR_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2270        With Me
2280          intRetVal = 0
2290          .cmdOK.SetFocus
2300        End With
2310      End Select
2320    End If

        ' ** Shift keys.
2330    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2340      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
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
        Case Else
2450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2460    End Select
2470    Resume EXITP

End Sub

Private Sub opgFont_optMICRN_KeyDown(KeyCode As Integer, Shift As Integer)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFont_optMICRN_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2570        With Me
2580          intRetVal = 0
2590          .cmdOK.SetFocus
2600        End With
2610      End Select
2620    End If

        ' ** Shift keys.
2630    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2640      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
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
        Case Else
2750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2760    End Select
2770    Resume EXITP

End Sub

Private Sub opgFont_optMICRN2_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFont_optMICRN2_KeyDown"

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
          Case vbKeyTab, vbKeyReturn
2870        With Me
2880          intRetVal = 0
2890          .cmdOK.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2940      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
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
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub

Private Sub opgFont_optMICRW_KeyDown(KeyCode As Integer, Shift As Integer)

3100  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFont_optMICRW_KeyDown"

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
3190          .cmdOK.SetFocus
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

Private Sub opgFont_optMICRW2_KeyDown(KeyCode As Integer, Shift As Integer)

3400  On Error GoTo ERRH

        Const THIS_PROC As String = "opgFont_optMICRW2_KeyDown"

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
3490          .cmdOK.SetFocus
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
