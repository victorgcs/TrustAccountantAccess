VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAccountAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAccountAssets"

'VGC 04/15/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Shorten:          S {cmdShorten on frmAccountAssets_Sub}
' **   Include Currency: U {chkIncludeCurrency}
' **   Widen:            W {cmdWiden on frmAccountAssets_Sub}
' **   Exit:             X {cmdClose}

Private lngRecsCur As Long
Private strAccountNo As String, strShortName As String
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim strFilter As String
        Dim intPos01 As Integer
        Dim strTmp01 As String
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = Trim(.OpenArgs)
150           If CharCnt(strTmp01, "~") = 2 Then  ' ** Module Function: modStringFuncs.
                ' ** strCallingForm ~ accountno ~ shortname

160             gstrFormQuerySpec = THIS_NAME
170             gblnHasForExThis = False

180             intPos01 = InStr(strTmp01, "~")
190             strCallingForm = Left(strTmp01, (intPos01 - 1))
200             strTmp01 = Mid(strTmp01, (intPos01 + 1))
210             intPos01 = InStr(strTmp01, "~")
220             strAccountNo = Left(strTmp01, (intPos01 - 1))
230             strShortName = Mid(strTmp01, (intPos01 + 1))

240             .accountno = strAccountNo
250             .shortname = strShortName

260             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

270             Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
280             DoEvents

290             chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.

300             strFilter = "[accountno] = '" & strAccountNo & "'"
310             .frmAccountAssets_Sub.Form.Filter = strFilter
320             .frmAccountAssets_Sub.Form.FilterOn = True

330             DoCmd.Hourglass False

340           Else
350             Cancel = -1
360           End If
370         Else
380           Cancel = -1
390         End If
400       Else
410         Cancel = -1
420       End If
430     End With

440     If Cancel = -1 Then
450       Beep
460       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
470     End If

EXITP:
480     Exit Sub

ERRH:
490     DoCmd.Hourglass False
500     Select Case ERR.Number
        Case Else
510       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
520     End Select
530     Resume EXITP

End Sub

Private Sub Form_Load()

600   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Load"

610     With Me
620       .frmAccountAssets_Sub.SetFocus
630       .frmAccountAssets_Sub.Form.FocusHolder.SetFocus
640     End With

EXITP:
650     Exit Sub

ERRH:
660     DoCmd.Hourglass False
670     Select Case ERR.Number
        Case Else
680       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
690     End Select
700     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

810     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
820     intShiftDown = (Shift And acShiftMask) > 0
830     intAltDown = (Shift And acAltMask) > 0
840     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Shorten:          S {cmdShorten on frmAccountAssets_Sub}
        ' **   Widen:            W {cmdWiden on frmAccountAssets_Sub}
        ' **   Exit:             X {cmdClose}

        ' ** Shortcut Ctrl keys to other forms and functionality:
        ' **   Save:             S {cmdSave on frmAccountAssets_Sub}

        ' ** Plain keys.
850     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
860       Select Case intRetVal
          Case vbKeyEscape
870         intRetVal = 0
880         cmdClose_Click  ' ** Procedure: Below.
890       End Select
900     End If

        ' ** Alt keys.
910     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
920       Select Case intRetVal
          Case vbKeyS
            ' ** Invoke the shorten-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
930         With Me
940           intRetVal = 0
950           If .frmAccountAssets_Sub.Form.cmdShorten.Enabled = True Then
960             .frmAccountAssets_Sub.Form.cmdShorten_Click  ' ** Form Procedure: frmAccountAssets_Sub.
970           Else
980             Beep
990           End If
1000        End With
1010      Case vbKeyW
            ' ** Invoke the widen-field command.
            ' ** Activated here, rather than an ampersand in the command button caption,
            ' ** so that the focus doesn't jump to the column header area.
1020        With Me
1030          intRetVal = 0
1040          If .frmAccountAssets_Sub.Form.cmdWiden.Enabled = True Then
1050            .frmAccountAssets_Sub.Form.cmdWiden_Click  ' ** Form Procedure: frmAccountAssets_Sub.
1060          Else
1070            Beep
1080          End If
1090        End With
1100      Case vbKeyX
1110        intRetVal = 0
1120        cmdClose_Click  ' ** Procedure: Below.
1130      End Select
1140    End If

        ' ** Ctrl-Shift keys.
1150    If intCtrlDown And (Not intAltDown) And intShiftDown Then
1160      Select Case intRetVal
          Case vbKeyF
1170        With Me
1180          intRetVal = 0
1190          .FocusHolder.SetFocus
1200        End With
1210      End Select
1220    End If

EXITP:
1230    KeyCode = intRetVal
1240    Exit Sub

ERRH:
1250    intRetVal = 0
1260    Select Case ERR.Number
        Case Else
1270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1280    End Select
1290    Resume EXITP

End Sub

Private Sub Form_Timer()

1300  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1310    With Me
1320      .TimerInterval = 0&
1330      gblnHasForExThis = HasForEx_Acct(strAccountNo, "A")  ' ** Module Function: modCurrencyFuncs.
1340      Select Case gblnHasForEx
          Case True
1350        Select Case gblnHasForExThis
            Case True
1360          If Pref_HasPref(THIS_NAME, "chkIncludeCurrency") = False Then  ' ** Module Procedure: modPreferenceFuncs.
1370            .chkIncludeCurrency = True
1380            chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1390          End If
1400          .chkIncludeCurrency.Enabled = True
1410        Case False
1420          .chkIncludeCurrency.Enabled = True
1430        End Select
1440      Case False
1450        .chkIncludeCurrency = False
1460        chkIncludeCurrency_AfterUpdate  ' ** Procedure: Below.
1470        .chkIncludeCurrency.Enabled = False
1480      End Select
1490    End With

EXITP:
1500    Exit Sub

ERRH:
1510    Select Case ERR.Number
        Case Else
1520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1530    End Select
1540    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1600  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim blnRetVal As Boolean

1610    DoCmd.Hourglass True
1620    DoEvents

        ' ** Make sure all reports are closed.
1630    If Reports.Count > 0 Then
1640      Do While Reports.Count > 0
1650        DoCmd.Close acReport, Reports(0).Name
1660      Loop
1670    End If

1680    Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1690    blnRetVal = MouseWheelOFF(False)  ' ** Module Function: modMouseWheel.

1700    gstrReturningForm = THIS_NAME
1710    gblnSetFocus = True
1720    If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
1730      Forms(strCallingForm).TimerInterval = 250&
1740      DoCmd.SelectObject acForm, strCallingForm, False
1750    Else
1760      DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Main"
1770    End If

EXITP:
1780    Exit Sub

ERRH:
1790    DoCmd.Hourglass False
1800    Select Case ERR.Number
        Case Else
1810      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1820    End Select
1830    Resume EXITP

End Sub

Public Sub cmdClose_Click()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1910    With Me

1920      .TimerInterval = 0&

1930      DoCmd.Close acForm, THIS_NAME

1940    End With

EXITP:
1950    Exit Sub

ERRH:
1960    Select Case ERR.Number
        Case Else
1970      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1980    End Select
1990    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
2090          lngRecsCur = .frmAccountAssets_Sub.Form.RecCnt  ' ** Form Function: frmAccountAssets_Sub.
2100          If lngRecsCur > 0& Then
2110            .frmAccountAssets_Sub.SetFocus
2120            .frmAccountAssets_Sub.Form.MoveRec acCmdRecordsGoToFirst  ' ** Form Procedure: frmAccountAssets_Sub.
2130            .frmAccountAssets_Sub.Form.TotalShareFace_display.SetFocus
2140          Else
2150            If .chkIncludeCurrency.Enabled = True Then
2160              .chkIncludeCurrency.SetFocus
2170            Else
                  ' ** Nowhere else to go!
2180            End If
2190          End If
2200        End With
2210      End Select
2220    End If

        ' ** Shift keys.
2230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2240      Select Case intRetVal
          Case vbKeyTab
2250        With Me
2260          intRetVal = 0
2270          If .chkIncludeCurrency.Enabled = True Then
2280            .chkIncludeCurrency.SetFocus
2290          Else
2300            lngRecsCur = .frmAccountAssets_Sub.Form.RecCnt  ' ** Form Function: frmAccountAssets_Sub.
2310            If lngRecsCur > 0& Then
2320              .frmAccountAssets_Sub.SetFocus
2330              .frmAccountAssets_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountAssets_Sub.
2340              .frmAccountAssets_Sub.Form.assettype_description.SetFocus
2350            Else
                  ' ** Nowhere else to go!
2360            End If
2370          End If
2380        End With
2390      End Select
2400    End If

EXITP:
2410    KeyCode = intRetVal
2420    Exit Sub

ERRH:
2430    intRetVal = 0
2440    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
2450    Case Else
2460      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2470    End Select
2480    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_AfterUpdate()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_AfterUpdate"

2510    With Me
2520      Select Case .chkIncludeCurrency
          Case True
2530        .chkIncludeCurrency_lbl.FontBold = True
2540        .frmAccountAssets_Sub.Form.IncludeCurrency_Sub True  ' ** Form Procedure: frmAccountAssets_Sub.
2550      Case False
2560        .chkIncludeCurrency_lbl.FontBold = False
2570        .frmAccountAssets_Sub.Form.IncludeCurrency_Sub False  ' ** Form Procedure: frmAccountAssets_Sub.
2580      End Select
2590      .frmAccountAssets_Sub.SetFocus
2600    End With

EXITP:
2610    Exit Sub

ERRH:
2620    Select Case ERR.Number
        Case 2467  ' ** The expression you entered refers to an object that is closed or doesn't exist.
          ' ** Form closing, ignore.
2630    Case Else
2640      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2650    End Select
2660    Resume EXITP

End Sub

Private Sub chkIncludeCurrency_KeyDown(KeyCode As Integer, Shift As Integer)

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "chkIncludeCurrency_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2710    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2720    intShiftDown = (Shift And acShiftMask) > 0
2730    intAltDown = (Shift And acAltMask) > 0
2740    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2750    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2760      Select Case intRetVal
          Case vbKeyTab, vbKeyReturn
2770        With Me
2780          .cmdClose.SetFocus
2790        End With
2800      End Select
2810    End If

        ' ** Shift keys.
2820    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2830      Select Case intRetVal
          Case vbKeyTab
2840        With Me
2850          intRetVal = 0
2860          lngRecsCur = .frmAccountAssets_Sub.Form.RecCnt  ' ** Form Function: frmAccountAssets_Sub.
2870          If lngRecsCur > 0& Then
2880            .frmAccountAssets_Sub.SetFocus
2890            .frmAccountAssets_Sub.Form.MoveRec acCmdRecordsGoToLast  ' ** Form Procedure: frmAccountAssets_Sub.
2900            .frmAccountAssets_Sub.Form.assettype_description.SetFocus
2910          Else
2920            .cmdClose.SetFocus
2930          End If
2940        End With
2950      End Select
2960    End If

EXITP:
2970    KeyCode = intRetVal
2980    Exit Sub

ERRH:
2990    intRetVal = 0
3000    Select Case ERR.Number
        Case 2105  ' ** You can't go to the specified record.
          ' ** Do nothing.
3010    Case Else
3020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3030    End Select
3040    Resume EXITP

End Sub
