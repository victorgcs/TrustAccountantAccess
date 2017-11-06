VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmAbout"

'VGC 07/23/2016: CHANGES!

' ** Shortcut Alt keys responsive from this form:
' **   Burst:            B {cmdGoldLogo}
' **   Close:            C {cmdClose}
' **   Show DateTime:    T {cmdShowDateTime}
' **   Exit:             X {cmdClose}

' ** Shortcut Ctrl-Shift keys responsive from this form:
' **   Close Form:       X {cmdCloseForm}

Private blnDevExit As Boolean, blnGoldLogo_Burst As Boolean
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim lngTpp As Long
        Dim lngTmp01 As Long

110     With Me

120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strCallingForm = Trim(.OpenArgs)
150         End If
160       End If

170       .PictureData = .Form_About_img

          'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
180       lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

190       blnGoldLogo_Burst = False

          ' ** Set Demo status.
200       If Len(TA_SEC) > Len(TA_SEC2) Then
210         lngTmp01 = (3& * lngTpp)
220         .Eval_lbl.Visible = True
230         .Version_Trust.Top = (.Version_Trust.Top - lngTmp01)
240         .Version_Trust_lbl.Top = (.Version_Trust_lbl.Top - lngTmp01)
250         .Date_Trust.Top = (.Date_Trust.Top - lngTmp01)
260         .Date_Trust2.Top = (.Date_Trust2.Top - lngTmp01)
270         .Date_Trust_lbl.Top = (.Date_Trust_lbl.Top - lngTmp01)
280         .License_Name.Top = (.License_Name.Top - lngTmp01)
290         .License_Name_lbl.Top = (.License_Name_lbl.Top - lngTmp01)
300         .VersionBox_demo_img.Visible = True
310         .VersionBox_img.Visible = False
320         .m_VP_Trust.Top = (.m_VP_Trust.Top - lngTmp01)
330         .m_VP_Trust_lbl.Top = (.m_VP_Trust_lbl.Top - lngTmp01)
340         .m_VP_Trust2.Top = .m_VP_Trust.Top
350         .m_VP_Trust2_lbl.Top = .m_VP_Trust_lbl.Top
360         .m_VD_TrustDta.Top = (.m_VD_TrustDta.Top - lngTmp01)
370         .m_VD_TrustDta_lbl.Top = (.m_VD_TrustDta_lbl.Top - lngTmp01)
380         .m_VD_TrustDta2.Top = .m_VD_TrustDta.Top
390         .m_VD_TrustDta2_lbl.Top = .m_VD_TrustDta_lbl.Top
400         .m_VA_TrstArch.Top = (.m_VA_TrstArch.Top - lngTmp01)
410         .m_VA_TrstArch_lbl.Top = (.m_VA_TrstArch_lbl.Top - lngTmp01)
420         .m_VA_TrstArch2.Top = .m_VA_TrstArch.Top
430         .m_VA_TrstArch2_lbl.Top = .m_VA_TrstArch_lbl.Top
440         .m_VX_TrustAux.Top = (.m_VX_TrustAux.Top - lngTmp01)
450         .m_VX_TrustAux_lbl.Top = (.m_VX_TrustAux_lbl.Top - lngTmp01)
460         .m_VX_TrustAux2.Top = .m_VX_TrustAux.Top
470         .m_VX_TrustAux2_lbl.Top = .m_VX_TrustAux_lbl.Top
480       Else
490         .Eval_lbl.Visible = False
500       End If

510       .Supplement_lbl.Visible = AppSupplement_Get  ' ** Module Function: modAppVersionFuncs.

          ' ** Set Excel status.
520       .NoExcel_lbl.Visible = AppNoExcel_Get  ' ** Module Function: modAppVersionFuncs.

530       blnDevExit = False

540     End With

550     If Cancel = -1 Then
          ' ** Just for the record.
560       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Allowed"
570     End If

EXITP:
580     Exit Sub

ERRH:
590     Select Case ERR.Number
        Case Else
600       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
610     End Select
620     Resume EXITP

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

700   On Error GoTo ERRH

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer
        Dim strUsername As String

        Const THIS_PROC As String = "Form_KeyDown"

        ' ** Use bit masks to determine which key was pressed.
710     intShiftDown = (Shift And acShiftMask) > 0
720     intAltDown = (Shift And acAltMask) > 0
730     intCtrlDown = (Shift And acCtrlMask) > 0

740     intRetVal = KeyCode

        ' ** Shortcut Ctrl-Shift keys to other forms and functionality:
        ' **   Close Form:       X {cmdCloseForm}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         intRetVal = 0
780         cmdClose_Click  ' ** Procedure: Below.
790       End Select
800     End If

        ' ** Alt keys.
810     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
820       Select Case intRetVal
          Case vbKeyX
830         intRetVal = 0
840         cmdClose_Click  ' ** Procedure: Below.
850       End Select
860     End If

        ' ** Ctrl-Shift keys.
870     If intCtrlDown And (Not intAltDown) And intShiftDown Then
880       Select Case intRetVal
          Case vbKeyX
            ' ** Close form without moving to another.
890         strUsername = GetUserName  ' ** Module Function: modFileUtilities.
900         If strUsername = gstrDevUserName Then
910           intRetVal = 0
920           blnDevExit = True
930           SetOption_Dev  ' ** Module Function: modUtilities.
940           DoCmd.Close acForm, THIS_NAME
950         End If
960       End Select
970     End If

EXITP:
980     KeyCode = intRetVal
990     Exit Sub

ERRH:
1000    intRetVal = 0
1010    Select Case ERR.Number
        Case Else
1020      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1030    End Select
1040    Resume EXITP

End Sub

Private Sub Form_Timer()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Static lngTmp01 As Long, lngTmp02 As Long

1110    With Me
1120      .TimerInterval = 0&
1130      If blnGoldLogo_Burst = True Then
1140        lngTmp01 = lngTmp01 + 1&
1150        .Gold_Logo_img.Visible = False
1160        .Controls("Gold_Logo_" & Right("00" & CStr(lngTmp01), 2) & "_anim_img").Visible = True
1170        If lngTmp01 > 1& Then
1180          .Controls("Gold_Logo_" & Right("00" & CStr(lngTmp01 - 1&), 2) & "_anim_img").Visible = False
1190        End If
1200        DoEvents
1210        If lngTmp01 = 21& Then
1220          lngTmp01 = 0&
1230          lngTmp02 = lngTmp02 + 1&
1240          blnGoldLogo_Burst = False
1250          .Gold_Logo_img.Visible = True
1260          .Gold_Logo_21_anim_img.Visible = False
1270        Else
1280          .TimerInterval = 50&
1290        End If
1300      Else
1310        blnGoldLogo_Burst = True
1320        lngTmp01 = 0&: lngTmp02 = 0&
1330        .TimerInterval = 500&
1340      End If
1350    End With

EXITP:
1360    Exit Sub

ERRH:
1370    Select Case ERR.Number
        Case Else
1380      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1390    End Select
1400    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

1500  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

1510    If blnDevExit = False Then

1520      Select Case strCallingForm
          Case "frmMain_Menu"
1530        gblnSetFocus = True
1540        gstrReturningForm = THIS_NAME
1550        DoCmd.OpenForm strCallingForm, , , , , , "frmMenu_Title"
1560        DoCmd.SelectObject acForm, strCallingForm, False
1570        Forms(strCallingForm).TimerInterval = 100&
1580      End Select

1590    End If

EXITP:
1600    Exit Sub

ERRH:
1610    Select Case ERR.Number
        Case Else
1620      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1630    End Select
1640    Resume EXITP

End Sub

Private Sub cmdClose_Click()

1700  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

1710    DoCmd.Close acForm, THIS_NAME

EXITP:
1720    Exit Sub

ERRH:
1730    Select Case ERR.Number
        Case Else
1740      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1750    End Select
1760    Resume EXITP

End Sub

Private Sub cmdGoldLogo_DblClick(Cancel As Integer)

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdGoldLogo_DblClick"

1810    With Me
1820      .TimerInterval = 10&
1830    End With

EXITP:
1840    Exit Sub

ERRH:
1850    Select Case ERR.Number
        Case Else
1860      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1870    End Select
1880    Resume EXITP

End Sub

Private Sub cmdShowDateTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' ** Shows release date timestamp, and AppVersions instead of the m_Vx tables.

1900  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShowDateTime_MouseDown"

1910    With Me
1920      .Date_Trust2.Visible = True
1930      .m_VP_Trust2.Visible = True
1940      .m_VP_Trust.Visible = False
1950      .m_VD_TrustDta2.Visible = True
1960      .m_VD_TrustDta.Visible = False
1970      .m_VA_TrstArch2.Visible = True
1980      .m_VA_TrstArch.Visible = False
1990      .m_VX_TrustAux2.Visible = True
2000      .m_VX_TrustAux.Visible = False
2010    End With

EXITP:
2020    Exit Sub

ERRH:
2030    Select Case ERR.Number
        Case Else
2040      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2050    End Select
2060    Resume EXITP

End Sub

Private Sub cmdShowDateTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdShowDateTime_MouseUp"

2110    With Me
2120      .Date_Trust2.Visible = False
2130      .m_VP_Trust2.Visible = False
2140      .m_VP_Trust.Visible = True
2150      .m_VD_TrustDta2.Visible = False
2160      .m_VD_TrustDta.Visible = True
2170      .m_VA_TrstArch2.Visible = False
2180      .m_VA_TrstArch.Visible = True
2190      .m_VX_TrustAux2.Visible = False
2200      .m_VX_TrustAux.Visible = True
2210    End With

EXITP:
2220    Exit Sub

ERRH:
2230    Select Case ERR.Number
        Case Else
2240      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2250    End Select
2260    Resume EXITP

End Sub

Private Sub Eval_lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "Eval_lbl_MouseDown"

        Dim strTmp01 As String, strTmp02 As String

2310    With Me
2320      strTmp01 = DemoLicense_Get  ' ** Module Function: modSecurityFunctions
2330      If Left(strTmp01, 1) <> "#" Then
2340        strTmp02 = Mid(strTmp01, (InStr(strTmp01, "~") + 1))
2350        strTmp01 = Left(strTmp01, (InStr(strTmp01, "~") - 1))
2360        strTmp02 = Mid(strTmp02, InStr(strTmp02, "A")) & "ed: "
2370        .Eval_lbl.Caption = strTmp02 & strTmp01
2380        .Eval_lbl.BackStyle = acBackStyleNormal
2390      End If
2400    End With

EXITP:
2410    Exit Sub

ERRH:
2420    Select Case ERR.Number
        Case Else
2430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2440    End Select
2450    Resume EXITP

End Sub

Private Sub Eval_lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "Eval_lbl_MouseUp"

2510    With Me
2520      .Eval_lbl.Caption = "Evaluation Version"
2530      .Eval_lbl.BackStyle = acBackStyleTransparent
2540    End With

EXITP:
2550    Exit Sub

ERRH:
2560    Select Case ERR.Number
        Case Else
2570      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2580    End Select
2590    Resume EXITP

End Sub
