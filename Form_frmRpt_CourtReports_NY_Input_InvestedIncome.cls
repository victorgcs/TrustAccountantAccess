VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_CourtReports_NY_Input_InvestedIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_CourtReports_NY_Input_InvestedIncome"

'VGC 12/29/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Cancel:           C {cmdCancel}
' **   OK:               O {cmdOK}
' **   Invested Income:  1 {InvestedIncome}

Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           strCallingForm = Trim(.OpenArgs)

150           .PictureData = .Form_Rpt_CourtReports_NY_Input_InvestedIncome_img

160           gstrCrtRpt_NY_InputTitle = "Invested Income"
170           InvestedIncome_lbl.Caption = gstrCrtRpt_NY_InputTitle
180           .Caption = "Court Report  - Input Value " & gstrCrtRpt_NY_InputTitle
190           .InvestedIncome = 0

200         Else
210           Cancel = -1
220         End If
230       Else
240         Cancel = -1
250       End If
260     End With

270     If Cancel = -1 Then
280       Beep
290       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
300     End If

EXITP:
310     Exit Sub

ERRH:
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

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

        ' ** Alt keys.
510     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
520       Select Case intRetVal
          Case vbKey1
530         With Me
540           intRetVal = 0
550           .InvestedIncome.SetFocus
560         End With
570       End Select
580     End If

EXITP:
590     KeyCode = intRetVal
600     Exit Sub

ERRH:
610     intRetVal = 0
620     Select Case ERR.Number
        Case Else
630       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
640     End Select
650     Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

700   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

710     Select Case strCallingForm
        Case "frmRpt_CourtReports_NY"
720       DoCmd.SelectObject acForm, strCallingForm, False
730     End Select

EXITP:
740     Exit Sub

ERRH:
750     Select Case ERR.Number
        Case Else
760       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
770     End Select
780     Resume EXITP

End Sub

Private Sub cmdOK_Click()

800   On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_Click"

        Dim frm As Access.Form
        Dim blnContinue As Boolean

810     blnContinue = True

820     Set frm = Forms("frmRpt_CourtReports_NY")

830     With Me
840       If IsNull(.InvestedIncome) = True Then
850         blnContinue = False
860         Beep
870         MsgBox "Income Balance must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
880       Else
890         If IsNumeric(.InvestedIncome) = False Then
900           blnContinue = False
910           Beep
920           MsgBox "Income Balance must be a numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
930         End If
940       End If

950       If blnContinue = True Then

960         frm.CashAssets_Beg = .InvestedIncome

970         gstrCrtRpt_CashAssets_Beg = CStr(.InvestedIncome)
980         gblnCrtRpt_NY_InvIncChange = True
990         gcurCrtRpt_NY_InputNew = .InvestedIncome

1000        DoCmd.Close acForm, THIS_NAME

1010      End If
1020    End With

EXITP:
1030    Set frm = Nothing
1040    Exit Sub

ERRH:
1050    Select Case ERR.Number
        Case Else
1060      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1070    End Select
1080    Resume EXITP

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdOK_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

1110    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
1120    intShiftDown = (Shift And acShiftMask) > 0
1130    intAltDown = (Shift And acAltMask) > 0
1140    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
1150    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
1160      Select Case intRetVal
          Case vbKeyTab
1170        With Me
1180          intRetVal = 0
1190          .cmdCancel.SetFocus
1200        End With
1210      End Select
1220    End If

        ' ** Shift keys.
1230    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
1240      Select Case intRetVal
          Case vbKeyTab
1250        With Me
1260          intRetVal = 0
1270          .InvestedIncome.SetFocus
1280        End With
1290      End Select
1300    End If

EXITP:
1310    KeyCode = intRetVal
1320    Exit Sub

ERRH:
1330    intRetVal = 0
1340    Select Case ERR.Number
        Case Else
1350      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1360    End Select
1370    Resume EXITP

End Sub

Private Sub cmdCancel_Click()

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdCancel_Click"

        Dim frm As Access.Form

1410    Set frm = Forms("frmRpt_CourtReports_NY")

1420    With frm
1430      .Ordinal = vbNullString  ' ** But leave the Public variable.
1440      .Version = vbNullString  ' ** But leave the Public variable.
1450      .InvestedIncome = vbNullString  ' ** But leave the Public variable.
1460      gstrCrtRpt_CashAssets_Beg = vbNullString
1470    End With
'InvestedIncome
1480    DoCmd.Close acForm, THIS_NAME

EXITP:
1490    Set frm = Nothing
1500    Exit Sub

ERRH:
1510    Select Case ERR.Number
        Case Else
1520      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1530    End Select
1540    Resume EXITP

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
1690          .InvestedIncome.SetFocus
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

Private Sub InvestedIncome_AfterUpdate()

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "InvestedIncome_AfterUpdate"

1910    With Me
1920      If IsNull(.InvestedIncome) = False Then
1930        If IsNumeric(.InvestedIncome) = True Then
1940          gstrCrtRpt_CashAssets_Beg = CStr(.InvestedIncome)
1950          .InvestedIncome_display_lbl.Caption = Format(.InvestedIncome, "Currency")
1960        Else
1970          Beep
1980          MsgBox "Numeric entry only.", vbInformation + vbOKOnly, "Invalid Entry"
1990          .InvestedIncome.SetFocus
2000        End If
2010      Else
2020        .InvestedIncome = 0&
2030        gstrCrtRpt_CashAssets_Beg = "0"
2040        .InvestedIncome_display_lbl.Caption = Format(.InvestedIncome, "Currency")
2050      End If
2060    End With

EXITP:
2070    Exit Sub

ERRH:
2080    Select Case ERR.Number
        Case Else
2090      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2100    End Select
2110    Resume EXITP

End Sub

Private Sub InvestedIncome_KeyDown(KeyCode As Integer, Shift As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "InvestedIncome_KeyDown"

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
