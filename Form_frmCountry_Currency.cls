VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCountry_Currency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmCountry_Currency"

'VGC 03/08/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Currency:         N {curr_name on frmCountry_Currency_Sub}
' **   Exit:             X {cmdClose}

Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long
Private lngSub_Width As Long, lngSub_Height As Long, lngSubBox_Width As Long, lngSubBox_Height As Long
Private lngClose_Left As Long, lngDetail_Height As Long, lngNavOffset_Top As Long, lngSizable_Offset As Long

Private blnSetSub As Boolean
Private lngCountryID As Long
Private lngRecsCur As Long, lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngTmp01 As Long

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then

140           blnIsOpen = True

150           strCallingForm = Trim(.OpenArgs)

160           With Forms(strCallingForm)
170             lngCountryID = .Controls(strCallingForm & "_Sub").Form.country_id
180             lngLeft = .frm_left
190             lngTop = .frm_top
200             lngWidth = .frm_width
210             lngHeight = .frm_height
220           End With

              'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions.
230           lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

              ' ** Variables are fed empty, then populated ByRef.
240           GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

250           FrmDimSave THIS_NAME, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modMonitorFuncs.

260           lngClose_Left = .cmdClose.Left
270           lngSub_Width = .frmCountry_Currency_Sub.Width
280           lngSub_Height = .frmCountry_Currency_Sub.Height
290           lngSubBox_Width = .frmCountry_Currency_Sub_box.Width
300           lngSubBox_Height = .frmCountry_Currency_Sub_box.Height
310           lngDetail_Height = .Detail.Height
320           lngSizable_Offset = .Sizable_lbl1.Left - .cmdClose.Left
330           lngNavOffset_Top = (.frmCountry_Currency_Sub_box.Top + .frmCountry_Currency_Sub_box.Height) - .Nav_hline01.Top

340           lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
350           lngMonitorNum = 1&: lngTmp01 = 0&
360           EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
370           If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.

380           If lngTop > 0& Or lngLeft > 0& Then
390             If lngMonitorNum = 1& Then lngTmp01 = lngTop
400             DoCmd.MoveSize lngLeft, lngTmp01, lngWidth, lngHeight  'lngTop
410             If lngMonitorNum > 1& Then
420               LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
430             End If
440           End If

450           .cmbCountries = lngCountryID

460           DoEvents
470           lngRecsCur = .frmCountry_Currency_Sub.Form.RecCnt  ' ** Form Function: frmCountry_Currency_Sub.
480           If lngRecsCur > 0& Then
490             .frmCountry_Currency_Sub.SetFocus
500           Else
510             .cmdClose.SetFocus
520           End If

530           blnIsOpen = False

540         Else
550           Cancel = -1
560         End If
570       Else
580         Cancel = -1
590       End If
600     End With

610     If Cancel = -1 Then
620       Beep
630       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
640     End If

EXITP:
650     Exit Sub

ERRH:
660     Select Case ERR.Number
        Case Else
670       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
680     End Select
690     Resume EXITP

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

        ' ** Shortcut Alt keys to other forms and functionality:
        ' **   Country:          N {country_name on frmCountry_Currency_Sub}

        ' ** Plain keys.
750     If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
760       Select Case intRetVal
          Case vbKeyEscape
770         With Me
780           intRetVal = 0
790           cmdClose_Click  ' ** Procedure: Below.
800         End With
810       End Select
820     End If

        ' ** Alt keys.
830     If (Not intCtrlDown) And intAltDown And (Not intShiftDown) Then
840       Select Case intRetVal
          Case vbKeyN
850         With Me
860           intRetVal = 0
870           .frmCountry_Currency_Sub.SetFocus
880           .frmCountry_Currency_Sub.Form.country_name.SetFocus
890         End With
900       Case vbKeyX
910         With Me
920           intRetVal = 0
930           cmdClose_Click  ' ** Procedure: Below.
940         End With
950       End Select
960     End If

EXITP:
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

Private Sub Form_Resize()

1100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Resize"

        Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
        Dim lngSubHeight As Long, lngNavTop As Long
        Dim lngTmp01 As Long

1110    With Me

1120      GetFormDimensions Me, lngLeft, lngTop, lngWidth, lngHeight  ' ** Module Procedure: modWindowFunctions.

1130      If lngWidth <> lngFrm_Width Or lngHeight <> lngFrm_Height Then

1140        If lngWidth <= lngFrm_Width Then
              ' ** Don't let it shrink narrower than it's opening width.
1150          .cmdClose.Left = (lngClose_Left - (lngFrm_Width - lngWidth))
1160          .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
1170          .Sizable_lbl2.Left = (.cmdClose.Left + lngSizable_Offset)
1180        Else
1190          .cmdClose.Left = (lngClose_Left + (lngWidth - lngFrm_Width))
1200          .Sizable_lbl1.Left = (.cmdClose.Left + lngSizable_Offset)
1210          .Sizable_lbl2.Left = (.cmdClose.Left + lngSizable_Offset)
1220          .frmCountry_Currency_Sub.Width = (lngSub_Width + (lngWidth - lngFrm_Width))
1230          .frmCountry_Currency_Sub_box.Width = ((lngSub_Width + (lngWidth - lngFrm_Width)) + (2& * lngTpp))
1240          lngTmp01 = ((.cmdClose.Left + .cmdClose.Width) + (9& * lngTpp))
1250          .Header_hline01.Width = lngTmp01
1260          .Header_hline02.Width = lngTmp01
1270          .Header_vline01.Left = lngTmp01
1280          .Header_vline02.Left = lngTmp01
1290          .Footer_hline01.Width = lngTmp01
1300          .Footer_hline02.Width = lngTmp01
1310          .Footer_vline01.Left = lngTmp01
1320          .Footer_vline02.Left = lngTmp01
1330          .Nav_box01.Width = .frmCountry_Currency_Sub.Width
1340        End If

1350        lngSubHeight = .frmCountry_Currency_Sub.Height
1360        If lngHeight > lngFrm_Height Then
1370          .Detail.Height = lngDetail_Height + (lngHeight - lngFrm_Height)
1380          .frmCountry_Currency_Sub.Height = lngSub_Height + (lngHeight - lngFrm_Height)
1390          .frmCountry_Currency_Sub_box.Height = (lngSub_Height + (lngHeight - lngFrm_Height)) + (6& * lngTpp)
1400          .Detail.Height = (.frmCountry_Currency_Sub.Top + .frmCountry_Currency_Sub.Height) + (32& * lngTpp)
1410          lngNavTop = ((.frmCountry_Currency_Sub_box.Top + .frmCountry_Currency_Sub_box.Height) - lngNavOffset_Top)
1420          .Nav_hline01.Top = lngNavTop
1430          .Nav_hline02.Top = (lngNavTop + lngTpp)
1440          .Nav_vline01.Top = lngNavTop
1450          .Nav_vline02.Top = lngNavTop
1460          .Nav_vline03.Top = lngNavTop
1470          .Nav_vline04.Top = lngNavTop
1480          .detail_height_line.Top = (.Detail.Height - .detail_height_line.Height)
1490        Else
1500          If (lngSub_Height - (lngFrm_Height - lngHeight)) > (.frmCountry_Currency_Sub.Form.Detail.Height + (19& * lngTpp)) Then
1510            .frmCountry_Currency_Sub.Height = (lngSub_Height - (lngFrm_Height - lngHeight))
1520            .frmCountry_Currency_Sub_box.Height = (lngSub_Height - (lngFrm_Height - lngHeight)) + (6& * lngTpp)
1530            lngNavTop = ((.frmCountry_Currency_Sub_box.Top + .frmCountry_Currency_Sub_box.Height) - lngNavOffset_Top)
1540            .Nav_hline01.Top = lngNavTop
1550            .Nav_hline02.Top = (lngNavTop + lngTpp)
1560            .Nav_vline01.Top = lngNavTop
1570            .Nav_vline02.Top = lngNavTop
1580            .Nav_vline03.Top = lngNavTop
1590            .Nav_vline04.Top = lngNavTop
1600            .detail_height_line.Top = ((lngDetail_Height - (lngFrm_Height - lngHeight)) - (24& * lngTpp))
1610            .Detail.Height = (.frmCountry_Currency_Sub.Top + .frmCountry_Currency_Sub.Height) + (32& * lngTpp)
1620            .detail_height_line.Top = (.Detail.Height - .detail_height_line.Height)
1630          End If
1640        End If

1650      End If

1660    End With

EXITP:
1670    Exit Sub

ERRH:
1680    Select Case ERR.Number
        Case Else
1690      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1700    End Select
1710    Resume EXITP

End Sub

Private Sub Form_Timer()

1800  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

1810    With Me
1820      .TimerInterval = 0&
1830      If gblnSetFocus = True Then
1840        gblnSetFocus = False
1850        strFormName = vbNullString
1860  On Error Resume Next
1870        Do Until strFormName = THIS_NAME
1880          DoCmd.SelectObject acForm, THIS_NAME, False
1890          strFormName = Screen.ActiveForm.Name
1900        Loop
1910  On Error GoTo ERRH
1920        blnSetSub = True
1930        .TimerInterval = 500&
1940      ElseIf blnSetSub = True Then
1950        blnSetSub = False
1960        .frmCountry_Currency_Sub.SetFocus
1970      End If
1980    End With

EXITP:
1990    Exit Sub

ERRH:
2000    Select Case ERR.Number
        Case Else
2010      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2020    End Select
2030    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2100  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Variables are fed empty, then populated ByRef.
2110    GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

2120    Select Case strCallingForm
        Case "frmCountryCode"
2130      With Forms(strCallingForm)
2140        .frm_left = lngFrm_Left
2150        .frm_top = lngFrm_Top
2160        .frm_width = lngFrm_Width
2170        .frm_height = lngFrm_Height
2180      End With
2190      DoEvents
2200      gstrReturningForm = THIS_NAME
2210      gblnSetFocus = True
2220      DoCmd.SelectObject acForm, strCallingForm, False
2230      Forms(strCallingForm).TimerInterval = 250&
2240    End Select

EXITP:
2250    Exit Sub

ERRH:
2260    Select Case ERR.Number
        Case Else
2270      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2280    End Select
2290    Resume EXITP

End Sub

Private Sub cmdClose_Click()

2300  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2310    DoCmd.Close acForm, THIS_NAME

EXITP:
2320    Exit Sub

ERRH:
2330    Select Case ERR.Number
        Case Else
2340      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2350    End Select
2360    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2400  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2410    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2420    intShiftDown = (Shift And acShiftMask) > 0
2430    intAltDown = (Shift And acAltMask) > 0
2440    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2450    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2460      Select Case intRetVal
          Case vbKeyTab
2470        With Me
2480          intRetVal = 0
2490          lngRecsCur = .frmCountry_Currency_Sub.Form.RecCnt  ' ** Form Function: frmCountry_Currency_Sub.
2500          If lngRecsCur > 0& Then
2510            .frmCountry_Currency_Sub.SetFocus
2520            .frmCountry_Currency_Sub.Form.FocusHolder.SetFocus
2530          Else
2540            Beep
2550          End If
2560        End With
2570      End Select
2580    End If

        ' ** Shift keys.
2590    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2600      Select Case intRetVal
          Case vbKeyTab
2610        With Me
2620          intRetVal = 0
2630          lngRecsCur = .frmCountry_Currency_Sub.Form.RecCnt  ' ** Form Function: frmCountry_Currency_Sub.
2640          If lngRecsCur > 0& Then
2650            .frmCountry_Currency_Sub.SetFocus
2660            .frmCountry_Currency_Sub.Form.FocusHolder.SetFocus
2670          Else
2680            Beep
2690          End If
2700        End With
2710      End Select
2720    End If

EXITP:
2730    KeyCode = intRetVal
2740    Exit Sub

ERRH:
2750    intRetVal = 0
2760    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2770    Case Else
2780      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2790    End Select
2800    Resume EXITP

End Sub
