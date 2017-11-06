VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStatementParameters_ForEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmStatementParameters_ForEx"

'VGC 09/18/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   No:               N {cmdNo}
' **   Yes:              Y {cmdYes}

Private blnNormalClose As Boolean
Private strAccountNo As String
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

150           intPos01 = InStr(strTmp01, "~")
160           strCallingForm = Left(strTmp01, (intPos01 - 1))
170           strAccountNo = Mid(strTmp01, (intPos01 + 1))

180           .PictureData = .Form_StatementParameters_ForEx_img

190           blnNormalClose = False

200           Pref_Load THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.
210           DoEvents

              ' ** EVENT CHECK: chkRememberForEx!
220           chkRememberForEx_AfterUpdate  ' ** Procedure: Below.

230           .cmdYes.SetFocus

240           DoCmd.Hourglass False

250         Else
260           Cancel = -1
270         End If
280       Else
290         Cancel = -1
300       End If
310     End With

320     If Cancel = -1 Then
330       Beep
340       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
350     End If

EXITP:
360     Exit Sub

ERRH:
370     DoCmd.Hourglass False
380     Select Case ERR.Number
        Case Else
390       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
400     End Select
410     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

500   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
510     intShiftDown = (Shift And acShiftMask) > 0
520     intAltDown = (Shift And acAltMask) > 0
530     intCtrlDown = (Shift And acCtrlMask) > 0

540     intRetVal = KeyCode

        ' ** Plain keys.
550     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
560       Select Case intRetVal
          Case vbKeyEscape
570         intRetVal = 0
580         cmdCancel_Click  ' ** Procedure: Below.
590       End Select
600     End If

        ' ** Ctrl-Shift keys.
610     If intCtrlDown And (Not intAltDown) And intShiftDown Then
620       Select Case intRetVal
          Case vbKeyF
630         With Me
640           intRetVal = 0
650           .FocusHolder.SetFocus
660         End With
670       End Select
680     End If

EXITP:
690     KeyCode = intRetVal
700     Exit Sub

ERRH:
710     intRetVal = 0
720     Select Case ERR.Number
        Case Else
730       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
740     End Select
750     Resume EXITP

End Sub

Private Sub Form_Timer()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

810     With Me
820       .TimerInterval = 0&
830       If gblnSetFocus = True Then
840         gblnSetFocus = False
850         strFormName = vbNullString
860   On Error Resume Next
870         Do Until strFormName = THIS_NAME
880           DoCmd.SelectObject acForm, THIS_NAME, False
890           strFormName = Screen.ActiveForm.Name
900         Loop
910   On Error GoTo ERRH
920       End If
930     End With

EXITP:
940     Exit Sub

ERRH:
950     DoCmd.Hourglass False
960     Select Case ERR.Number
        Case Else
970       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
980     End Select
990     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        Dim dbs As DAO.Database, qdf As DAO.QueryDef

1010    With Me

1020      Pref_Save THIS_NAME  ' ** Module Procedure: modPreferenceFuncs.

1030      Select Case blnNormalClose
          Case True
1040        Select Case gblnMessage
            Case True
1050          Select Case gblnSwitchTo
              Case True
1060            If .chkRememberForEx = True Then
1070              Set dbs = CurrentDb
1080              With dbs
                    ' ** Update tblCurrency_Account, for curracct_suppress = True, by specified [actno].
1090                Set qdf = .QueryDefs("qryStatementParameters_ForEx_02")
1100                With qdf.Parameters
1110                  ![actno] = strAccountNo
1120                End With
1130                qdf.Execute
1140                Set qdf = Nothing
1150                .Close
1160              End With
1170            End If
1180          Case False ' ** NO: gblnSwitchTo = False
                ' ** Do include foreign exchange columns.
1190          End Select
1200        Case False  ' ** CANCEL: gblnMessage = False
              ' ** Cancel, so don't update tblCurrency_Account.
1210        End Select
1220      Case False
1230        gblnMessage = False
1240      End Select

1250    End With

EXITP:
1260    Set qdf = Nothing
1270    Set dbs = Nothing
1280    Exit Sub

ERRH:
1290    Select Case ERR.Number
        Case Else
1300      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1310    End Select
1320    Resume EXITP

End Sub

Private Sub cmdYes_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdYes_Click"

1410    blnNormalClose = True

        ' ** Both glnMessage and gblnSwitchTo remain True.
1420    DoCmd.Close acForm, THIS_NAME

EXITP:
1430    Exit Sub

ERRH:
1440    Select Case ERR.Number
        Case Else
1450      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1460    End Select
1470    Resume EXITP

End Sub

Private Sub cmdYes_KeyDown(KeyCode As Integer, Shift As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdYes_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1510    intShiftDown = (Shift And acShiftMask) > 0
1520    intAltDown = (Shift And acAltMask) > 0
1530    intCtrlDown = (Shift And acCtrlMask) > 0

1540    intRetVal = KeyCode

        ' ** Plain keys.
1550    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1560      Select Case intRetVal
          Case vbKeyTab
1570        With Me
1580          intRetVal = 0
1590          .cmdNo.SetFocus
1600        End With
1610      End Select
1620    End If

        ' ** Shift keys.
1630    If Not (intCtrlDown) And (Not intAltDown) And intShiftDown Then
1640      Select Case intRetVal
          Case vbKeyTab
1650        With Me
1660          intRetVal = 0
1670          .cmdCancel.SetFocus
1680        End With
1690      End Select
1700    End If

EXITP:
1710    KeyCode = intRetVal
1720    Exit Sub

ERRH:
1730    intRetVal = 0
1740    Select Case ERR.Number
        Case Else
1750      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1760    End Select
1770    Resume EXITP

End Sub

Private Sub cmdNo_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNo_Click"

1810    blnNormalClose = True

        ' ** gblnMessage remains True.
1820    gblnSwitchTo = False
1830    DoCmd.Close acForm, THIS_NAME

EXITP:
1840    Exit Sub

ERRH:
1850    Select Case ERR.Number
        Case Else
1860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1870    End Select
1880    Resume EXITP

End Sub

Private Sub cmdNo_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdNo_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
1910    intShiftDown = (Shift And acShiftMask) > 0
1920    intAltDown = (Shift And acAltMask) > 0
1930    intCtrlDown = (Shift And acCtrlMask) > 0

1940    intRetVal = KeyCode

        ' ** Plain keys.
1950    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1960      Select Case intRetVal
          Case vbKeyTab
1970        With Me
1980          intRetVal = 0
1990          .cmdCancel.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If Not (intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab
2050        With Me
2060          intRetVal = 0
2070          .cmdYes.SetFocus
2080        End With
2090      End Select
2100    End If

EXITP:
2110    KeyCode = intRetVal
2120    Exit Sub

ERRH:
2130    intRetVal = 0
2140    Select Case ERR.Number
        Case Else
2150      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2160    End Select
2170    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

2210    blnNormalClose = True

        ' ** gblnSwitchTo doesn't matter.
2220    gblnMessage = False
2230    DoCmd.Close acForm, THIS_NAME

EXITP:
2240    Exit Sub

ERRH:
2250    Select Case ERR.Number
        Case Else
2260      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2270    End Select
2280    Resume EXITP

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2310    intShiftDown = (Shift And acShiftMask) > 0
2320    intAltDown = (Shift And acAltMask) > 0
2330    intCtrlDown = (Shift And acCtrlMask) > 0

2340    intRetVal = KeyCode

        ' ** Plain keys.
2350    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2360      Select Case intRetVal
          Case vbKeyTab
2370        With Me
2380          intRetVal = 0
2390          .cmdYes.SetFocus
2400        End With
2410      End Select
2420    End If

        ' ** Shift keys.
2430    If Not (intCtrlDown) And (Not intAltDown) And intShiftDown Then
2440      Select Case intRetVal
          Case vbKeyTab
2450        With Me
2460          intRetVal = 0
2470          .cmdNo.SetFocus
2480        End With
2490      End Select
2500    End If

EXITP:
2510    KeyCode = intRetVal
2520    Exit Sub

ERRH:
2530    intRetVal = 0
2540    Select Case ERR.Number
        Case Else
2550      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2560    End Select
2570    Resume EXITP

End Sub

Private Sub chkRememberForEx_AfterUpdate()

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberForEx_AfterUpdate"

2610    With Me
2620      Select Case .chkRememberForEx
          Case True
2630        .chkRememberForEx_lbl.FontBold = True
2640        .cmdNo.Enabled = False
2650      Case False
2660        .chkRememberForEx_lbl.FontBold = False
2670        .cmdNo.Enabled = True
2680      End Select
2690    End With

EXITP:
2700    Exit Sub

ERRH:
2710    Select Case ERR.Number
        Case Else
2720      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2730    End Select
2740    Resume EXITP

End Sub

Private Sub chkRememberForEx_KeyDown(KeyCode As Integer, Shift As Integer)

2800  On Error GoTo ERRH

        Const THIS_PROC As String = "chkRememberForEx_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

        ' ** Use bit masks to determine which key was pressed.
2810    intShiftDown = (Shift And acShiftMask) > 0
2820    intAltDown = (Shift And acAltMask) > 0
2830    intCtrlDown = (Shift And acCtrlMask) > 0

2840    intRetVal = KeyCode

        ' ** Plain keys.
2850    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2860      Select Case intRetVal
          Case vbKeyTab
2870        With Me
2880          intRetVal = 0
2890          .cmdYes.SetFocus
2900        End With
2910      End Select
2920    End If

        ' ** Shift keys.
2930    If Not (intCtrlDown) And (Not intAltDown) And intShiftDown Then
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
        Case Else
3050      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3060    End Select
3070    Resume EXITP

End Sub
