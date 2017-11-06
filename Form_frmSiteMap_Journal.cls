VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSiteMap_Journal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const THIS_NAME As String = "frmSiteMap_Journal"

'VGC 09/06/2017: CHANGES!

' #######################################
' ## Monitor Funcs:
' ##   Form_Open()
' #######################################

' ** Shortcut Alt keys responsive from this form:
' **   Close:            C {cmdClose}
' **   Exit:             X {cmdClose}

Private lngForm_Width As Long, lngDetail_Height As Long, lngSub_Width As Long, lngSub_Height As Long
Private lngClose_Left As Long, lngSizable_Offset As Long, lngMin_Width As Long, lngMin_Height As Long
Private lngFrm_Left_Orig As Long, lngFrm_Top_Orig As Long, lngFrm_Width_Orig As Long, lngFrm_Height_Orig As Long
Private lngFrm_Left As Long, lngFrm_Top As Long, lngFrm_Width As Long, lngFrm_Height As Long

Private frmSiteMap As Access.Form
Private strJForm As String
Private lngTpp As Long, blnIsOpen As Boolean, lngMonitorCnt As Long, lngMonitorNum As Long
Private strCallingForm As String
' **

Private Sub Form_Open(Cancel As Integer)

100   On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Open"

        Dim intPos01 As Integer
        Dim strTmp01 As String, lngTmp02 As Long
        Dim blnRetVal As Boolean

110     With Me
120       If IsNull(.OpenArgs) = False Then
130         If Trim(.OpenArgs) <> vbNullString Then
140           strTmp01 = .OpenArgs
150           intPos01 = InStr(strTmp01, "~")
160           If intPos01 > 0 Then

170             blnIsOpen = True

180             strCallingForm = Left(strTmp01, (intPos01 - 1))
190             strJForm = Mid(strTmp01, (intPos01 + 1))

200             blnRetVal = MouseWheelON  ' ** Module Function: modMouseWheel.

                'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
210             lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!

220             Set frmSiteMap = Forms(strCallingForm)

230             lngForm_Width = .Width
240             lngDetail_Height = .Detail.Height
250             lngSub_Width = .frmSiteMap_Journal_Sub.Width
260             lngSub_Height = .frmSiteMap_Journal_Sub.Height
270             lngClose_Left = .cmdClose.Left
280             lngSizable_Offset = (.Width - .Sizable_lbl1.Left)

290             lngMin_Width = (3& * 1440&)  ' ** 1440 Twips per inch is independent of Twips per pixel.
300             lngMin_Height = (2& * 1440&)  ' ** This is Detail height.

                ' ** Variables are fed empty, then populated ByRef.
310             GetFormDimensions Me, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modWindowFunctions.

320             FrmDimSave THIS_NAME, lngFrm_Left_Orig, lngFrm_Top_Orig, lngFrm_Width_Orig, lngFrm_Height_Orig  ' ** Module Function: modMonitorFuncs.

                ' ** Last saved position and size.
330             With frmSiteMap
340               If .frmj_height > 0 Then
350                 lngMonitorCnt = GetMonitorCount  ' ** Module Function: modMonitorFuncs.
360                 lngMonitorNum = 1&: lngTmp02 = 0&
370                 EnumMonitors Me  ' ** Module Function: modMonitorFuncs.
380                 If lngMonitorCnt > 1& Then lngMonitorNum = GetMonitorNum  ' ** Module Function: modMonitorFuncs.
390                 If lngMonitorNum = 1& Then lngTmp02 = .frmj_top
400                 DoCmd.SelectObject acForm, THIS_NAME, False
410                 DoCmd.MoveSize .frmj_left, lngTmp02, .frmj_width, .frmj_height  '.frmj_top
420                 If lngMonitorNum > 1& Then
430                   LoadPosition .hwnd, THIS_NAME  ' ** Module Function: modMonitorFuncs.
440                 End If
450               End If
460             End With

470             Select Case strJForm
                Case "Classic"
480               .frmSiteMap_Journal_Sub.SetFocus
490               .frmSiteMap_Journal_Sub.Form.FocusHolder01.SetFocus
500             Case "Columnar"
510               .frmSiteMap_Journal_Sub.SetFocus
520               .frmSiteMap_Journal_Sub.Form.FocusHolder12.SetFocus
530             End Select

540             blnIsOpen = False

550           Else
560             Cancel = -1
570           End If
580         Else
590           Cancel = -1
600         End If
610       Else
620         Cancel = -1
630       End If
640     End With

650     If Cancel = -1 Then
660       Beep
670       MsgBox "This form is not to be opened independently.", vbCritical + vbOKOnly, "Form Inaccessible"
680     End If

EXITP:
690     Exit Sub

ERRH:
700     Select Case ERR.Number
        Case Else
710       zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
720     End Select
730     Resume EXITP

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
          Case vbKeyX
930         intRetVal = 0
940         cmdClose_Click  ' ** Procedure: Below.
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

        Dim lngDiff_Width As Long, lngDiff_Height As Long

1110    With Me

          ' ** Variables are fed empty, then populated ByRef.
1120      GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

1130      If lngTpp = 0& Then
            'lngTpp = GetTPP  ' ** Module Function: modWindowFunctions
1140        lngTpp = 15&  ' ** Windows keeps saying 20, but it's 15 that works!
1150      End If

1160      If lngFrm_Width > lngFrm_Width_Orig Then
1170        lngDiff_Width = (lngFrm_Width - lngFrm_Width_Orig)
1180        .Width = (lngForm_Width + lngDiff_Width)
1190        .frmSiteMap_Journal_Sub.Width = (lngSub_Width + lngDiff_Width)
1200        .frmSiteMap_Journal_Sub_box.Width = (.frmSiteMap_Journal_Sub.Width + (2& * lngTpp))
1210        .cmdClose.Left = (lngClose_Left + lngDiff_Width)
1220        .Sizable_lbl1.Left = ((lngForm_Width + lngDiff_Width) - lngSizable_Offset)
1230        .Sizable_lbl2.Left = ((lngForm_Width + lngDiff_Width) - lngSizable_Offset)
1240        .Header_vline01.Left = (lngForm_Width + lngDiff_Width)
1250        .Header_vline02.Left = (lngForm_Width + lngDiff_Width)
1260        .Footer_vline01.Left = (lngForm_Width + lngDiff_Width)
1270        .Footer_vline02.Left = (lngForm_Width + lngDiff_Width)
1280        .Header_hline01.Width = (lngForm_Width + lngDiff_Width)
1290        .Header_hline02.Width = (lngForm_Width + lngDiff_Width)
1300        .Footer_hline01.Width = (lngForm_Width + lngDiff_Width)
1310        .Footer_hline02.Width = (lngForm_Width + lngDiff_Width)
1320      ElseIf lngFrm_Width < lngFrm_Width_Orig Then
1330        .form_width_line.Left = 0&  ' ** Just get it out of the way.
1340        lngDiff_Width = (lngFrm_Width_Orig - lngFrm_Width)
1350        If (lngForm_Width - lngDiff_Width) >= lngMin_Width Then
1360          .frmSiteMap_Journal_Sub.Width = (lngSub_Width - lngDiff_Width)
1370          .frmSiteMap_Journal_Sub_box.Width = (.frmSiteMap_Journal_Sub.Width + (2& * lngTpp))
1380          .cmdClose.Left = (lngClose_Left - lngDiff_Width)
1390          .Sizable_lbl1.Left = ((lngForm_Width - lngDiff_Width) - lngSizable_Offset)
1400          .Sizable_lbl2.Left = ((lngForm_Width - lngDiff_Width) - lngSizable_Offset)
1410          .Header_vline01.Left = (lngForm_Width - lngDiff_Width)
1420          .Header_vline02.Left = (lngForm_Width - lngDiff_Width)
1430          .Footer_vline01.Left = (lngForm_Width - lngDiff_Width)
1440          .Footer_vline02.Left = (lngForm_Width - lngDiff_Width)
1450          .Header_hline01.Width = (lngForm_Width - lngDiff_Width)
1460          .Header_hline02.Width = (lngForm_Width - lngDiff_Width)
1470          .Footer_hline01.Width = (lngForm_Width - lngDiff_Width)
1480          .Footer_hline02.Width = (lngForm_Width - lngDiff_Width)
1490          .Width = (lngForm_Width - lngDiff_Width)
1500        End If
1510      Else
1520        .Width = lngForm_Width
1530        .frmSiteMap_Journal_Sub.Width = lngSub_Width
1540        .frmSiteMap_Journal_Sub_box.Width = (.frmSiteMap_Journal_Sub.Width + (2& * lngTpp))
1550        .cmdClose.Left = lngClose_Left
1560        .Sizable_lbl1.Left = (lngForm_Width - lngSizable_Offset)
1570        .Sizable_lbl2.Left = (lngForm_Width - lngSizable_Offset)
1580        .Header_vline01.Left = lngForm_Width
1590        .Header_vline02.Left = lngForm_Width
1600        .Footer_vline01.Left = lngForm_Width
1610        .Footer_vline02.Left = lngForm_Width
1620        .Header_hline01.Width = lngForm_Width
1630        .Header_hline02.Width = lngForm_Width
1640        .Footer_hline01.Width = lngForm_Width
1650        .Footer_hline02.Width = lngForm_Width
1660        .Width = lngForm_Width
1670      End If

1680      If lngFrm_Height > lngFrm_Height_Orig Then
1690        lngDiff_Height = (lngFrm_Height - lngFrm_Height_Orig)
1700        .Detail.Height = (lngDetail_Height + lngDiff_Height)
1710        .frmSiteMap_Journal_Sub.Height = (lngSub_Height + lngDiff_Height)
1720        .frmSiteMap_Journal_Sub_box.Height = (.frmSiteMap_Journal_Sub.Height + (2& * lngTpp))
1730      ElseIf lngFrm_Height < lngFrm_Height_Orig Then
1740        .detail_height_line.Top = 0&  ' ** Just get it out of the way.
1750        lngDiff_Height = (lngFrm_Height_Orig - lngFrm_Height)
1760        If (lngDetail_Height - lngDiff_Height) >= lngMin_Height Then
1770          .frmSiteMap_Journal_Sub.Height = (lngSub_Height - lngDiff_Height)
1780          .frmSiteMap_Journal_Sub_box.Height = (.frmSiteMap_Journal_Sub.Height + (2& * lngTpp))
1790          .Detail.Height = (lngDetail_Height - lngDiff_Height)
1800        End If
1810      Else
1820        .Detail.Height = lngDetail_Height
1830        .frmSiteMap_Journal_Sub.Height = lngSub_Height
1840        .frmSiteMap_Journal_Sub_box.Height = (.frmSiteMap_Journal_Sub.Height + (2& * lngTpp))
1850        .Detail.Height = lngDetail_Height
1860      End If

1870    End With

EXITP:
1880    Exit Sub

ERRH:
1890    Select Case ERR.Number
        Case Else
1900      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
1910    End Select
1920    Resume EXITP

End Sub

Private Sub Form_Timer()

2000  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Timer"

        Dim strFormName As String

2010    With Me
2020      .TimerInterval = 0&
2030      If gblnSetFocus = True Then
2040        gblnSetFocus = False
2050        strFormName = vbNullString
2060  On Error Resume Next
2070        Do Until strFormName = THIS_NAME
2080          DoCmd.SelectObject acForm, THIS_NAME, False
2090          strFormName = Screen.ActiveForm.Name
2100        Loop
2110  On Error GoTo ERRH
2120        .frmSiteMap_Journal_Sub.SetFocus
2130      End If
2140    End With

EXITP:
2150    Exit Sub

ERRH:
2160    Select Case ERR.Number
        Case Else
2170      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2180    End Select
2190    Resume EXITP

End Sub

Private Sub Form_Unload(Cancel As Integer)

2200  On Error GoTo ERRH

        Const THIS_PROC As String = "Form_Unload"

        ' ** Variables are fed empty, then populated ByRef.
2210    GetFormDimensions Me, lngFrm_Left, lngFrm_Top, lngFrm_Width, lngFrm_Height  ' ** Module Function: modWindowFunctions.

2220    With Forms(strCallingForm)
2230      Select Case .chkSaveSizePos
          Case True
2240        .frmj_top = lngFrm_Top
2250        .frmj_left = lngFrm_Left
2260        .frmj_width = lngFrm_Width
2270        .frmj_height = lngFrm_Height
2280      Case False
2290        .frmj_top = 0&
2300        .frmj_left = 0&
2310        .frmj_width = 0&
2320        .frmj_height = 0&
2330      End Select
2340    End With

2350    Select Case strCallingForm
        Case "frmSiteMap"
2360      gstrReturningForm = THIS_NAME
2370      gblnSetFocus = True
2380      DoCmd.SelectObject acForm, strCallingForm, False
2390      Forms(strCallingForm).TimerInterval = 100&
2400    End Select

EXITP:
2410    Exit Sub

ERRH:
2420    Select Case ERR.Number
        Case Else
2430      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2440    End Select
2450    Resume EXITP

End Sub

Public Sub cmdClose_Click()

2500  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_Click"

2510    DoCmd.Close acForm, THIS_NAME

EXITP:
2520    Exit Sub

ERRH:
2530    Select Case ERR.Number
        Case Else
2540      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2550    End Select
2560    Resume EXITP

End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)

2600  On Error GoTo ERRH

        Const THIS_PROC As String = "cmdClose_KeyDown"

        Dim intShiftDown As Integer, intAltDown As Integer, intCtrlDown As Integer
        Dim intRetVal As Integer

2610    intRetVal = KeyCode

        ' ** Use bit masks to determine which key was pressed.
2620    intShiftDown = (Shift And acShiftMask) > 0
2630    intAltDown = (Shift And acAltMask) > 0
2640    intCtrlDown = (Shift And acCtrlMask) > 0

        ' ** Plain keys.
2650    If (Not intCtrlDown) And (Not intAltDown) And (Not intShiftDown) Then
2660      Select Case intRetVal
          Case vbKeyTab
2670        With Me
2680          intRetVal = 0
2690          .frmSiteMap_Journal_Sub.SetFocus
2700          .frmSiteMap_Journal_Sub.Form.FocusHolder01.SetFocus
2710        End With
2720      End Select
2730    End If

        ' ** Shift keys.
2740    If (Not intCtrlDown) And (Not intAltDown) And intShiftDown Then
2750      Select Case intRetVal
          Case vbKeyTab
2760        With Me
2770          intRetVal = 0
2780          .frmSiteMap_Journal_Sub.SetFocus
2790          .frmSiteMap_Journal_Sub.Form.FocusHolder12.SetFocus
2800        End With
2810      End Select
2820    End If

EXITP:
2830    KeyCode = intRetVal
2840    Exit Sub

ERRH:
2850    intRetVal = 0
2860    Select Case ERR.Number
        Case 2046  ' ** The command or action isn't available now (first or last record).
          ' ** Do nothing.
2870    Case Else
2880      zErrorHandler THIS_NAME, THIS_PROC, ERR.Number, Erl, ERR.description  ' ** Module Function: modErrorHandler.
2890    End Select
2900    Resume EXITP

End Sub
