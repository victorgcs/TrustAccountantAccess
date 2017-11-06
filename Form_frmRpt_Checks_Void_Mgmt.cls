VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRpt_Checks_Void_Mgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmRpt_Checks_Void_Mgmt"

'VGC 07/13/2017: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Bank Acct Number: A {chkbank_acctnum on frmRpt_Checks_Void_Mgmt_Sub}
' **   Account Number:   B {accountno on frmRpt_Checks_Void_Mgmt_Sub}
' **   Close:            C {cmdClose}
' **   Bank Name:        E {chkbank_name on frmRpt_Checks_Void_Mgmt_Sub}
' **   Check Number:     K {chkvoid_chknum on frmRpt_Checks_Void_Mgmt_Sub}
' **   Account Name:     M {shortname on frmRpt_Checks_Void_Mgmt_Sub}
' **   Print Report:     P {cmdPrintReport}
' **   Exit:             X {cmdClose}

Private blnPrintReport_Focus As Boolean, blnPrintReport_MouseDown As Boolean

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
140         If strTmp01 <> vbNullString Then

150           strCallingForm = Trim(strTmp01)

160           .PictureData = .Form_Rpt_Checks_Void_Mgmt_img

170           blnPrintReport_Focus = False: blnPrintReport_MouseDown = False

180         Else
190           Cancel = -1
200         End If
210       Else
220         Cancel = -1
230       End If
240     End With

250     DoCmd.Hourglass False

260     If Cancel = -1 Then
270       Beep
280       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
290     End If

EXITP:
300     Exit Sub

ERRH:
310     DoCmd.Hourglass False
320     Select Case ERR.Number
        Case Else
330       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
340     End Select
350     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

400   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_KeyDown"

        Dim frm As Access.Form
        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

410     intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
420     intShiftDown = (Shift And acShiftMask) > 0
430     intAltDown = (Shift And acAltMask) > 0
440     intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Bank Acct Number: A {chkbank_acctnum on frmRpt_Checks_Void_Mgmt_Sub}
        ' **   Account Number:   B {accountno on frmRpt_Checks_Void_Mgmt_Sub}
        ' **   Bank Name:        E {chkbank_name on frmRpt_Checks_Void_Mgmt_Sub}
        ' **   Check Number:     K {chkvoid_chknum on frmRpt_Checks_Void_Mgmt_Sub}
        ' **   Account Name:     M {shortname on frmRpt_Checks_Void_Mgmt_Sub}
        ' **   Exit:             X {cmdClose}

450     Set frm = Me.frmRpt_Checks_Void_Mgmt_Sub.Form

        ' ** Plain keys.
460     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
470       Select Case intRetVal
          Case vbKeyEscape
480         intRetVal = 0
490         cmdClose_Click  ' ** Procedure: Below.
500       End Select
510     End If

        ' ** Alt keys.
520     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
530       Select Case intRetVal
          Case vbKeyA
540         With Me
550           intRetVal = 0
560           .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
570           frm.chkbank_acctnum.SetFocus
580         End With
590       Case vbKeyB
600         With Me
610           intRetVal = 0
620           .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
630           frm.accountno.SetFocus
640         End With
650       Case vbKeyE
660         With Me
670           intRetVal = 0
680           .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
690           frm.chkbank_name.SetFocus
700         End With
710       Case vbKeyK
720         With Me
730           intRetVal = 0
740           .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
750           frm.chkvoid_chknum.SetFocus
760         End With
770       Case vbKeyM
780         With Me
790           intRetVal = 0
800           .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
810           frm.shortname.SetFocus
820         End With
830       Case vbKeyX
840         intRetVal = 0
850         cmdClose_Click  ' ** Procedure: Below.
860       End Select
870     End If

        ' ** Ctrl-Shift keys.
880     If intCtrlDown And (Not intAltDown) And intShiftDown Then
890       Select Case intRetVal
          Case vbKeyF
900         With Me
910           intRetVal = 0
920           .FocusHolder.SetFocus
930         End With
940       End Select
950     End If

EXITP:
960     Set frm = Nothing
970     KeyCode = intRetVal
980     Exit Sub

ERRH:
990     intRetVal = 0
1000    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
1010    Case Else
1020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1030    End Select
1040    Resume EXITP

End Sub

Private Sub Form_Timer()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

1110    With Me
1120      .TimerInterval = 0&

1130    End With

EXITP:
1140    Exit Sub

ERRH:
1150    Select Case ERR.Number
        Case Else
1160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1170    End Select
1180    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1210    DoCmd.Hourglass True
1220    DoEvents
1230    gstrReturningForm = THIS_NAME
1240    gblnSetFocus = True
1250    Select Case strCallingForm
        Case "frmRpt_Checks"
1260      If IsLoaded(strCallingForm, acForm) = True Then  ' ** Module Function: modFileUtilities.
1270        DoCmd.SelectObject acForm, strCallingForm, False
1280        Forms(strCallingForm).TimerInterval = 250&
1290      Else
1300        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Post"
1310      End If
1320    End Select

EXITP:
1330    Exit Sub

ERRH:
1340    DoCmd.Hourglass False
1350    Select Case ERR.Number
        Case Else
1360      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1370    End Select
1380    Resume EXITP

End Sub

Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1400  On Error GoTo ERRH

        Const THIS_PROC As String = "Detail_MouseMove"

1410    With Me
1420      If .cmdPrintReport_raised_focus_dots_img.Visible = True Or .cmdPrintReport_raised_focus_img.Visible = True Then
1430        Select Case blnPrintReport_Focus
            Case True
1440          .cmdPrintReport_raised_semifocus_dots_img.Visible = True
1450          .cmdPrintReport_raised_img.Visible = False
1460        Case False
1470          .cmdPrintReport_raised_img.Visible = True
1480          .cmdPrintReport_raised_semifocus_dots_img.Visible = False
1490        End Select
1500        .cmdPrintReport_raised_focus_img.Visible = False
1510        .cmdPrintReport_raised_focus_dots_img.Visible = False
1520        .cmdPrintReport_sunken_focus_dots_img.Visible = False
1530        .cmdPrintReport_raised_img_dis.Visible = False
1540      End If
1550    End With

EXITP:
1560    Exit Sub

ERRH:
1570    Select Case ERR.Number
        Case Else
1580      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1590    End Select
1600    Resume EXITP

End Sub

Private Sub FormFooter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "FormFooter_MouseMove"

1710    Detail_MouseMove Button, Shift, X, Y  ' ** Procedure: Above.

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Public Sub cmdClose_Click()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1810    DoCmd.Close acForm, THIS_NAME

EXITP:
1820    Exit Sub

ERRH:
1830    Select Case ERR.Number
        Case Else
1840      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1850    End Select
1860    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

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
1990          .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
2000        End With
2010      End Select
2020    End If

        ' ** Shift keys.
2030    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2040      Select Case intRetVal
          Case vbKeyTab
2050        With Me
2060          intRetVal = 0
2070          If .cmdPrintReport.Enabled = True Then
2080            .cmdPrintReport.SetFocus
2090          Else
2100            .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
2110          End If
2120        End With
2130      End Select
2140    End If

EXITP:
2150    KeyCode = intRetVal
2160    Exit Sub

ERRH:
2170    intRetVal = 0
2180    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2190    Case Else
2200      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2210    End Select
2220    Resume EXITP

End Sub

Private Sub cmdPrintReport_GotFocus()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_GotFocus"

2310    With Me
2320      blnPrintReport_Focus = True
2330      .cmdPrintReport_raised_semifocus_dots_img.Visible = True
2340      .cmdPrintReport_raised_img.Visible = False
2350      .cmdPrintReport_raised_focus_img.Visible = False
2360      .cmdPrintReport_raised_focus_dots_img.Visible = False
2370      .cmdPrintReport_sunken_focus_dots_img.Visible = False
2380      .cmdPrintReport_raised_img_dis.Visible = False
2390    End With

EXITP:
2400    Exit Sub

ERRH:
2410    Select Case ERR.Number
        Case Else
2420      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2430    End Select
2440    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseDown"

2510    With Me
2520      blnPrintReport_MouseDown = True
2530      .cmdPrintReport_sunken_focus_dots_img.Visible = True
2540      .cmdPrintReport_raised_img.Visible = False
2550      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
2560      .cmdPrintReport_raised_focus_img.Visible = False
2570      .cmdPrintReport_raised_focus_dots_img.Visible = False
2580      .cmdPrintReport_raised_img_dis.Visible = False
2590    End With

EXITP:
2600    Exit Sub

ERRH:
2610    Select Case ERR.Number
        Case Else
2620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2630    End Select
2640    Resume EXITP

End Sub

Private Sub cmdPrintReport_Click()

2700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_Click"

        Dim strDocName As String

2710    With Me

2720      strDocName = "rptListOfVoidedChecks"
2730      If gblnDev_Debug Or GetUserName = gstrDevUserName Then  ' ** Module Function: modFileUtilities.
2740        DoCmd.OpenReport strDocName, acViewPreview
2750        DoCmd.Maximize
2760        DoCmd.RunCommand acCmdFitToWindow
2770        Forms(strCallingForm).Visible = False
2780        .Visible = False
2790      Else
            '##GTR_Ref: rptListOfVoidedChecks
2800        DoCmd.OpenReport strDocName, acViewNormal
2810      End If

2820    End With

EXITP:
2830    Exit Sub

ERRH:
2840    Forms(strCallingForm).Visible = True
2850    Me.Visible = True
2860    Select Case ERR.Number
        Case Else
2870      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2880    End Select
2890    Resume EXITP

End Sub

Private Sub cmdPrintReport_KeyDown(KeyCode As Integer, Shift As Integer)

2900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_KeyDown"

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
3070          .frmRpt_Checks_Void_Mgmt_Sub.SetFocus
3080        End With
3090      End Select
3100    End If

EXITP:
3110    KeyCode = intRetVal
3120    Exit Sub

ERRH:
3130    intRetVal = 0
3140    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
3150    Case Else
3160      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
3170    End Select
3180    Resume EXITP

End Sub

Private Sub cmdPrintReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

3200  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseMove"

3210    With Me
3220      If blnPrintReport_MouseDown = False Then
3230        Select Case blnPrintReport_Focus
            Case True
3240          .cmdPrintReport_raised_focus_dots_img.Visible = True
3250          .cmdPrintReport_raised_focus_img.Visible = False
3260        Case False
3270          .cmdPrintReport_raised_focus_img.Visible = True
3280          .cmdPrintReport_raised_focus_dots_img.Visible = False
3290        End Select
3300        .cmdPrintReport_raised_img.Visible = False
3310        .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3320        .cmdPrintReport_sunken_focus_dots_img.Visible = False
3330        .cmdPrintReport_raised_img_dis.Visible = False
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

Private Sub cmdPrintReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

3500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_MouseUp"

3510    With Me
3520      .cmdPrintReport_raised_focus_dots_img.Visible = True
3530      .cmdPrintReport_raised_img.Visible = False
3540      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3550      .cmdPrintReport_raised_focus_img.Visible = False
3560      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3570      .cmdPrintReport_raised_img_dis.Visible = False
3580      blnPrintReport_MouseDown = False
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

Private Sub cmdPrintReport_LostFocus()

3700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdPrintReport_LostFocus"

3710    With Me
3720      .cmdPrintReport_raised_img.Visible = True
3730      .cmdPrintReport_raised_semifocus_dots_img.Visible = False
3740      .cmdPrintReport_raised_focus_img.Visible = False
3750      .cmdPrintReport_raised_focus_dots_img.Visible = False
3760      .cmdPrintReport_sunken_focus_dots_img.Visible = False
3770      .cmdPrintReport_raised_img_dis.Visible = False
3780      blnPrintReport_Focus = False
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
